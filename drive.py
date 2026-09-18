"""نشر المقاطع المكتملة إلى Google Drive.

كل مقطع يُرفع داخل مجلد خاص به تحت مجلد رئيسي واحد، ليبقى الفيديو ووصفه
معًا ولا يختلطا بغيرهما:

    مُصحَف/
      Q-002_سورة_البقرة_1-7/
        Q-002_سورة_البقرة_1-7.mp4
        Q-002_سورة_البقرة_1-7_social.txt

Implemented straight against the Drive REST API with `requests`, which the
project already depends on, rather than google-api-python-client: the official
client pulls in a large dependency tree that PyInstaller has to be told about
piece by piece, and the two calls actually needed here — create a folder, and
upload a file — are a few lines of HTTP each.

The scope requested is `drive.file`, which grants access only to files this
application itself created. It never sees the rest of the user's Drive, and
because of that Google does not require the app to pass verification before
the owner can use it on their own account.
"""
from __future__ import annotations

import json
import mimetypes
import secrets
import threading
import time
import urllib.parse
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

import requests

from generator import BASE_DIR, log

CONFIG_FILE = BASE_DIR / "drive_config.json"
TOKEN_FILE = BASE_DIR / "drive_token.json"

AUTH_ENDPOINT = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_ENDPOINT = "https://oauth2.googleapis.com/token"
FILES_ENDPOINT = "https://www.googleapis.com/drive/v3/files"
UPLOAD_ENDPOINT = "https://www.googleapis.com/upload/drive/v3/files"
USERINFO_ENDPOINT = "https://www.googleapis.com/oauth2/v2/userinfo"

SCOPES = "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.email"
FOLDER_MIME = "application/vnd.google-apps.folder"
DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file"

# Small enough that one dropped connection costs a few seconds of re-sending
# rather than minutes, and that a 150 MB clip reports progress ~40 times.
CHUNK_SIZE = 4 * 1024 * 1024
# A reset connection mid-upload is normal on a domestic line; the resumable
# protocol is built for exactly this, so a break is retried from the byte
# Drive says it already holds instead of failing the whole file.
MAX_RETRIES = 6
DEFAULT_ROOT = "مُصحَف"

_lock = threading.Lock()


# ---------------------------------------------------------------------------
# Stored settings and tokens
# ---------------------------------------------------------------------------

def _read_json(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def load_config() -> dict:
    config = _read_json(CONFIG_FILE)
    config.setdefault("client_id", "")
    config.setdefault("client_secret", "")
    config.setdefault("root_folder", DEFAULT_ROOT)
    config.setdefault("auto_upload", False)
    return config


def save_config(**changes) -> dict:
    config = load_config()
    for key in ("client_id", "client_secret", "root_folder"):
        if key in changes and changes[key] is not None:
            config[key] = str(changes[key]).strip()
    if "auto_upload" in changes and changes["auto_upload"] is not None:
        config["auto_upload"] = bool(changes["auto_upload"])
    if not config["root_folder"]:
        config["root_folder"] = DEFAULT_ROOT
    _write_json(CONFIG_FILE, config)
    return config


def load_tokens() -> dict:
    return _read_json(TOKEN_FILE)


def save_tokens(data: dict) -> None:
    _write_json(TOKEN_FILE, data)


def disconnect() -> None:
    """Forget the account. The uploaded files stay in the user's Drive."""
    try:
        TOKEN_FILE.unlink()
    except FileNotFoundError:
        pass


def is_connected() -> bool:
    return bool(load_tokens().get("refresh_token"))


def status() -> dict:
    config = load_config()
    tokens = load_tokens()
    return {
        "configured": bool(config["client_id"] and config["client_secret"]),
        "connected": bool(tokens.get("refresh_token")),
        "email": tokens.get("email", ""),
        "root_folder": config["root_folder"],
        "auto_upload": config["auto_upload"],
        "client_id": config["client_id"],
        # The secret is never sent back to the page; only whether one is set.
        "has_secret": bool(config["client_secret"]),
    }


# ---------------------------------------------------------------------------
# OAuth: the loopback flow for installed applications
# ---------------------------------------------------------------------------

class _CallbackHandler(BaseHTTPRequestHandler):
    """Catches the single redirect Google sends back with the auth code."""

    code: str | None = None
    error: str | None = None
    expected_state: str = ""

    def log_message(self, *args):
        pass

    def do_GET(self):
        query = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
        state = (query.get("state") or [""])[0]
        if state != _CallbackHandler.expected_state:
            # A request that did not come from the authorisation we started.
            self.send_error(400)
            return
        _CallbackHandler.code = (query.get("code") or [None])[0]
        _CallbackHandler.error = (query.get("error") or [None])[0]
        body = (
            "<!doctype html><html lang='ar' dir='rtl'><meta charset='utf-8'>"
            "<body style=\"font-family:system-ui;background:#0A0F1D;color:#E8EDF7;"
            "display:flex;align-items:center;justify-content:center;height:100vh;margin:0\">"
            "<div style='text-align:center'><h2>"
            + ("تم ربط حساب Google Drive بنجاح ✅" if _CallbackHandler.code else "تعذّر الربط ❌")
            + "</h2><p>يمكنك إغلاق هذه الصفحة والعودة إلى لوحة التحكم.</p></div></body></html>"
        ).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


def connect(callback=None, timeout: float = 300.0) -> dict:
    """Run the browser consent flow and store the refresh token.

    Blocks until the user finishes in the browser or the timeout expires."""
    config = load_config()
    if not config["client_id"] or not config["client_secret"]:
        raise RuntimeError("أدخل معرّف العميل والسر أولًا (client_id و client_secret).")

    _CallbackHandler.code = None
    _CallbackHandler.error = None
    _CallbackHandler.expected_state = secrets.token_urlsafe(24)
    # Port 0 lets the OS pick a free one, so this never collides with the panel.
    httpd = HTTPServer(("127.0.0.1", 0), _CallbackHandler)
    redirect_uri = f"http://127.0.0.1:{httpd.server_port}"
    params = {
        "client_id": config["client_id"],
        "redirect_uri": redirect_uri,
        "response_type": "code",
        "scope": SCOPES,
        # Google only returns a refresh token when it is asked for offline
        # access, and only on the first consent unless prompt=consent forces
        # the screen again - without both, a re-connect yields a token that
        # dies in an hour with no way to renew it.
        "access_type": "offline",
        "prompt": "consent",
        "state": _CallbackHandler.expected_state,
    }
    url = f"{AUTH_ENDPOINT}?{urllib.parse.urlencode(params)}"
    log("افتح المتصفح وأكمل تسجيل الدخول إلى Google…", callback)
    threading.Thread(target=webbrowser.open, args=(url,), daemon=True).start()

    deadline = time.time() + timeout
    httpd.timeout = 1.0
    try:
        while time.time() < deadline:
            if _CallbackHandler.code or _CallbackHandler.error:
                break
            httpd.handle_request()
    finally:
        httpd.server_close()

    if _CallbackHandler.error:
        raise RuntimeError(f"رفض Google الطلب: {_CallbackHandler.error}")
    if not _CallbackHandler.code:
        raise RuntimeError("انتهت المهلة قبل إتمام تسجيل الدخول.")

    response = requests.post(TOKEN_ENDPOINT, data={
        "code": _CallbackHandler.code,
        "client_id": config["client_id"],
        "client_secret": config["client_secret"],
        "redirect_uri": redirect_uri,
        "grant_type": "authorization_code",
    }, timeout=30)
    if response.status_code != 200:
        raise RuntimeError(f"فشل تبادل رمز الدخول: {response.text[:300]}")
    payload = response.json()
    if not payload.get("refresh_token"):
        raise RuntimeError("لم تُرجع Google رمز تحديث. افصل التطبيق من حسابك ثم أعد الربط.")

    granted = (payload.get("scope") or "").split()
    if DRIVE_SCOPE not in granted:
        raise RuntimeError(
            "تم تسجيل الدخول لكن لم تُمنح صلاحية Google Drive. "
            "أعد الربط، وفي شاشة الموافقة أشّر على مربع الوصول إلى Drive "
            "(«الاطّلاع على ملفات Drive التي تفتحها بهذا التطبيق وتعديلها») قبل المتابعة."
        )

    tokens = {
        "refresh_token": payload["refresh_token"],
        "access_token": payload.get("access_token", ""),
        "scope": payload.get("scope", ""),
        "expires_at": time.time() + float(payload.get("expires_in", 0)) - 60,
    }
    try:
        info = requests.get(USERINFO_ENDPOINT, timeout=20,
                            headers={"Authorization": f"Bearer {tokens['access_token']}"})
        if info.status_code == 200:
            tokens["email"] = info.json().get("email", "")
    except Exception:
        # The address is only shown as a label; failing to read it must not
        # invalidate a connection that otherwise succeeded.
        pass
    save_tokens(tokens)
    log(f"تم ربط الحساب {tokens.get('email', '')}", callback)
    return tokens


def access_token() -> str:
    """A valid access token, refreshed from the stored refresh token if the
    current one has expired."""
    with _lock:
        tokens = load_tokens()
        if not tokens.get("refresh_token"):
            raise RuntimeError("لم يتم ربط حساب Google Drive بعد.")
        if tokens.get("access_token") and time.time() < float(tokens.get("expires_at", 0)):
            return tokens["access_token"]
        config = load_config()
        response = requests.post(TOKEN_ENDPOINT, data={
            "refresh_token": tokens["refresh_token"],
            "client_id": config["client_id"],
            "client_secret": config["client_secret"],
            "grant_type": "refresh_token",
        }, timeout=30)
        if response.status_code != 200:
            raise RuntimeError(f"تعذّر تجديد صلاحية الدخول: {response.text[:300]}")
        payload = response.json()
        tokens["access_token"] = payload["access_token"]
        tokens["expires_at"] = time.time() + float(payload.get("expires_in", 0)) - 60
        save_tokens(tokens)
        return tokens["access_token"]


# ---------------------------------------------------------------------------
# Drive operations
# ---------------------------------------------------------------------------

def _headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}"}


def _explain(response) -> str:
    """Google's setup failures read as raw JSON; name the fix instead."""
    body = response.text
    if "SERVICE_DISABLED" in body or "has not been used in project" in body:
        return ("Google Drive API غير مُفعَّل في مشروعك. افتح Google Cloud Console → "
                "APIs & Services → Library → Google Drive API → Enable.")
    if "insufficientPermissions" in body or "insufficient authentication scopes" in body:
        return ("لم تُمنح صلاحية Drive. اضغط «إعادة الربط» وأشّر على مربع الوصول "
                "إلى Drive في شاشة الموافقة.")
    return body[:300]


def _escape(name: str) -> str:
    return name.replace("\\", "\\\\").replace("'", "\\'")


def find_folder(name: str, parent: str | None, token: str) -> str | None:
    """The id of a folder this app created with that name, if it exists.

    Looked up before every create so re-running a batch adds to the existing
    tree rather than piling up duplicates named the same thing."""
    query = [f"name = '{_escape(name)}'", f"mimeType = '{FOLDER_MIME}'", "trashed = false"]
    query.append(f"'{parent}' in parents" if parent else "'root' in parents")
    response = requests.get(FILES_ENDPOINT, headers=_headers(token), timeout=30, params={
        "q": " and ".join(query),
        "fields": "files(id,name)",
        "pageSize": 10,
        "spaces": "drive",
    })
    if response.status_code != 200:
        raise RuntimeError(f"تعذّر البحث عن المجلد «{name}»: {_explain(response)}")
    files = response.json().get("files", [])
    return files[0]["id"] if files else None


def ensure_folder(name: str, parent: str | None, token: str) -> str:
    existing = find_folder(name, parent, token)
    if existing:
        return existing
    metadata = {"name": name, "mimeType": FOLDER_MIME}
    if parent:
        metadata["parents"] = [parent]
    response = requests.post(FILES_ENDPOINT, headers=_headers(token), timeout=30,
                             params={"fields": "id"}, json=metadata)
    if response.status_code not in (200, 201):
        raise RuntimeError(f"تعذّر إنشاء المجلد «{name}»: {response.text[:300]}")
    return response.json()["id"]


def _committed_bytes(session_url: str, size: int) -> int | None:
    """Ask Drive how much of this upload session it already holds.

    An empty PUT with `bytes */total` is the protocol's own "where were we?"
    question: 308 answers with a Range header covering what arrived, and
    200/201 means the file was actually committed before the connection
    dropped. Returns None when the upload is already complete."""
    probe = requests.put(session_url, timeout=60, headers={
        "Content-Length": "0",
        "Content-Range": f"bytes */{size}",
    })
    if probe.status_code in (200, 201):
        return None
    if probe.status_code != 308:
        raise RuntimeError(f"تعذّر استئناف الرفع: {probe.status_code} {probe.text[:200]}")
    received = probe.headers.get("Range")
    # No Range header at all means Drive has nothing yet.
    return int(received.split("-")[-1]) + 1 if received else 0


def upload_file(path: Path, parent: str, token: str, callback=None, progress=None) -> dict:
    """Resumable upload of one file into one folder.

    A file of the same name already in that folder is replaced rather than
    duplicated, so re-uploading a regenerated clip keeps the folder tidy."""
    size = path.stat().st_size
    mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"

    response = requests.get(FILES_ENDPOINT, headers=_headers(token), timeout=30, params={
        "q": f"name = '{_escape(path.name)}' and '{parent}' in parents and trashed = false",
        "fields": "files(id)", "pageSize": 1, "spaces": "drive",
    })
    existing = (response.json().get("files") or [None])[0] if response.status_code == 200 else None

    if existing:
        start = requests.patch(
            f"{UPLOAD_ENDPOINT}/{existing['id']}", headers=_headers(token), timeout=30,
            params={"uploadType": "resumable"}, json={"name": path.name})
    else:
        start = requests.post(
            UPLOAD_ENDPOINT, headers=_headers(token), timeout=30,
            params={"uploadType": "resumable"},
            json={"name": path.name, "parents": [parent]})
    if start.status_code not in (200, 201):
        raise RuntimeError(f"تعذّر بدء رفع «{path.name}»: {start.text[:300]}")
    session_url = start.headers.get("Location")
    if not session_url:
        raise RuntimeError(f"لم تُرجع Google رابط رفع لـ «{path.name}».")

    def report(sent: int, state: str = "uploading") -> None:
        if progress:
            progress({"file": path.name, "sent": sent, "total": size,
                      "percent": round(sent / size * 100) if size else 100, "state": state})

    sent = 0
    attempts = 0
    report(0)
    with path.open("rb") as handle:
        while sent < size:
            handle.seek(sent)
            chunk = handle.read(CHUNK_SIZE)
            if not chunk:
                break
            last = sent + len(chunk) - 1
            try:
                put = requests.put(session_url, data=chunk, timeout=300, headers={
                    "Content-Length": str(len(chunk)),
                    "Content-Range": f"bytes {sent}-{last}/{size}",
                    "Content-Type": mime,
                })
            except requests.RequestException as exc:
                # The connection died mid-chunk. Ask Drive what it kept and
                # carry on from there rather than throwing the file away.
                attempts += 1
                if attempts > MAX_RETRIES:
                    raise RuntimeError(
                        f"انقطع الاتصال أثناء رفع «{path.name}» بعد {MAX_RETRIES} محاولات: {exc}")
                wait = min(2 ** attempts, 30)
                log(f"    ↑ {path.name} — انقطع الاتصال، إعادة المحاولة بعد {wait}s "
                    f"({attempts}/{MAX_RETRIES})", callback)
                time.sleep(wait)
                try:
                    resumed = _committed_bytes(session_url, size)
                except Exception:
                    continue
                if resumed is None:
                    log(f"    ↑ {path.name} — اكتمل", callback)
                    report(size, "done")
                    return {"resumed": True}
                sent = resumed
                report(sent)
                continue

            # 308 means Drive has the chunk and wants the next one; 200/201
            # means that was the last one and the file is committed.
            if put.status_code in (200, 201):
                log(f"    ↑ {path.name} — اكتمل", callback)
                report(size, "done")
                return put.json()
            if put.status_code == 308:
                attempts = 0
                received = put.headers.get("Range")
                sent = int(received.split("-")[-1]) + 1 if received else last + 1
                report(sent)
                log(f"    ↑ {path.name} — {sent / size * 100:.0f}%", callback)
                continue
            if put.status_code in (500, 502, 503, 504):
                attempts += 1
                if attempts > MAX_RETRIES:
                    raise RuntimeError(f"فشل رفع «{path.name}»: {put.status_code}")
                wait = min(2 ** attempts, 30)
                log(f"    ↑ {path.name} — خطأ مؤقت من Google، إعادة المحاولة بعد {wait}s", callback)
                time.sleep(wait)
                resumed = _committed_bytes(session_url, size)
                if resumed is None:
                    report(size, "done")
                    return {"resumed": True}
                sent = resumed
                continue
            raise RuntimeError(f"فشل رفع «{path.name}»: {put.status_code} {_explain(put)}")
    raise RuntimeError(f"انتهى رفع «{path.name}» دون تأكيد من Google.")


def publish_clip(video: Path, caption: Path | None = None, callback=None, progress=None) -> dict:
    """Upload one finished clip and its caption into a folder of their own.

    The folder is named after the clip, so the video and the text that belongs
    to it can never be told apart from another clip's pair."""
    if not video.exists():
        raise RuntimeError(f"الملف غير موجود: {video.name}")
    if caption is None:
        guess = video.with_name(video.stem + "_social.txt")
        caption = guess if guess.exists() else None

    config = load_config()
    token = access_token()
    scope = load_tokens().get("scope")
    if scope is not None and DRIVE_SCOPE not in scope.split():
        raise RuntimeError(
            "الحساب مرتبط بدون صلاحية Google Drive. اضغط «إعادة الربط» "
            "وأشّر على مربع الوصول إلى Drive في شاشة الموافقة."
        )
    log(f"رفع «{video.stem}» إلى Google Drive…", callback)
    root_id = ensure_folder(config["root_folder"], None, token)
    clip_id = ensure_folder(video.stem, root_id, token)

    def stage(payload: dict) -> None:
        if progress:
            progress(dict(payload, clip=video.stem))

    uploaded = [upload_file(video, clip_id, token, callback, progress=stage)]
    if caption and caption.exists():
        uploaded.append(upload_file(caption, clip_id, token, callback, progress=stage))
    else:
        log(f"    (لا يوجد ملف وصف لـ {video.stem})", callback)

    link = f"https://drive.google.com/drive/folders/{clip_id}"
    log(f"تم رفع «{video.stem}» — {len(uploaded)} ملف", callback)
    return {"folder_id": clip_id, "folder_url": link, "files": len(uploaded),
            "root_folder": config["root_folder"], "clip_folder": video.stem}
