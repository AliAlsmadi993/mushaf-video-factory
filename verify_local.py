from __future__ import annotations

import shutil
from pathlib import Path

from generator import OUTPUT_DIR, TASKS_XLSX, executable, load_tasks


def main() -> int:
    print('Python: OK')
    print(f'Excel: {TASKS_XLSX.name} — {len(load_tasks())} مهمة')
    print(f'FFmpeg: {executable("ffmpeg") or "غير موجود"}')
    print(f'FFprobe: {shutil.which("ffprobe") or "غير موجود — سيُستخدم fallback"}')
    print(f'المجلدات: خلفيات={Path("خلفيات").exists()} output={OUTPUT_DIR.exists()} temp={Path("temp").exists()}')
    if not executable('ffmpeg'):
        print('تحذير: لا يمكن إنشاء الفيديو حتى تثبّت FFmpeg أو imageio-ffmpeg.')
        return 1
    print('البيئة الأساسية جاهزة. لتجربة الرندر: python generator.py --demo --output output/demo.mp4')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
