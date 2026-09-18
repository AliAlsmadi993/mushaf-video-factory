"""Double-click entry point for the packaged EXE.

Starts the local control-panel server in the background and opens it in the
default browser automatically — no command line needed.
"""
from __future__ import annotations

import shutil
import sys
import threading
import time
import webbrowser
from pathlib import Path

from generator import BACKGROUND_DIR, BASE_DIR, BUNDLE_DIR, OUTPUT_DIR, TEMP_DIR, FONT_DIR, TASKS_XLSX


def bootstrap() -> None:
    """Lay bundled resources down next to the EXE."""
    # index.html belongs to the build, so it must be refreshed whenever the
    # bundled copy differs. Seeding it only when missing meant a rebuilt EXE
    # kept serving the interface extracted by an older version, and none of
    # the newer controls ever appeared.
    ui_source = BUNDLE_DIR / "index.html"
    ui_target = BASE_DIR / "index.html"
    if ui_source.exists() and ui_source != ui_target:
        if not ui_target.exists() or ui_source.read_bytes() != ui_target.read_bytes():
            shutil.copy2(ui_source, ui_target)
    # tasks.xlsx accumulates per-clip generation status, so it is only seeded
    # when absent — overwriting it would wipe the user's progress.
    xlsx_source = BUNDLE_DIR / TASKS_XLSX.name
    if xlsx_source.exists() and (not TASKS_XLSX.exists() or TASKS_XLSX.stat().st_size == 0) and xlsx_source != TASKS_XLSX:
        shutil.copy2(xlsx_source, TASKS_XLSX)
    fonts_source = BUNDLE_DIR / "fonts"
    if fonts_source.exists() and fonts_source != FONT_DIR:
        FONT_DIR.mkdir(parents=True, exist_ok=True)
        for item in fonts_source.iterdir():
            dest = FONT_DIR / item.name
            if not dest.exists():
                shutil.copy2(item, dest)
    for folder in (OUTPUT_DIR, TEMP_DIR, BACKGROUND_DIR):
        folder.mkdir(parents=True, exist_ok=True)


def main() -> None:
    bootstrap()
    import server

    thread = threading.Thread(target=server.main, daemon=True, name="mushaf-server")
    thread.start()
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{server.PORT}")
    print("مُصحَف يعمل الآن — أبقِ هذه النافذة مفتوحة، وأغلقها لإيقاف الخادم.", flush=True)
    try:
        while thread.is_alive():
            thread.join(1)
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
