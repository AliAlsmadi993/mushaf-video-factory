from __future__ import annotations

import argparse
import json
from pathlib import Path

from generator import TEMP_DIR, fetch_surah_text, load_tasks

BASE_DIR = Path(__file__).resolve().parent
RAW_JSON = BASE_DIR / "tasks_raw.json"


def export_workbook() -> list[dict]:
    tasks = load_tasks(BASE_DIR / "tasks.xlsx")
    RAW_JSON.write_text(json.dumps(tasks, ensure_ascii=False, indent=2), encoding="utf-8")
    return tasks


def refresh_text_cache(tasks: list[dict]) -> int:
    seen = set()
    completed = 0
    for task in tasks:
        surah = int(task["surah_number"])
        if surah in seen:
            continue
        seen.add(surah)
        try:
            verses = fetch_surah_text(surah)
            completed += len(verses)
            print(f"تم تحديث نصوص السورة {surah}: {len(verses)} آية")
        except Exception as exc:
            print(f"تعذر تحديث السورة {surah}: {exc}")
    return completed


def main() -> int:
    parser = argparse.ArgumentParser(description="بناء نسخة JSON من فهرس ركوعات القرآن")
    parser.add_argument("--refresh-text", action="store_true", help="تحديث مخزن نصوص الآيات من المصادر العامة")
    args = parser.parse_args()
    tasks = export_workbook()
    print(f"تم تصدير {len(tasks)} ركوعاً إلى {RAW_JSON.name}")
    if args.refresh_text:
        count = refresh_text_cache(tasks)
        print(f"إجمالي الآيات المخزنة مؤقتاً: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
