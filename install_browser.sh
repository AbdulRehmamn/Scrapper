#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"
if [[ ! -x ".venv/bin/python" ]]; then
  echo "Run ./start_scraper.sh once before using this installer."
  exit 1
fi
.venv/bin/python -m playwright install chromium
echo "Chromium is ready for JavaScript-rendered pages."
