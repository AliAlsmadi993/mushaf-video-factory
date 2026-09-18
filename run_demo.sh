#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
python3 generator.py --demo --output output/demo.mp4
