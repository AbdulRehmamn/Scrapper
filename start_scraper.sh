#!/usr/bin/env bash
set -euo pipefail
cd "$(dirname "$0")"

python_command="${PYTHON_COMMAND:-python3}"
if [[ ! -x ".venv/bin/python" ]]; then
  echo "Creating the local Python environment..."
  "$python_command" -m venv .venv
fi

echo "Installing or checking required packages..."
.venv/bin/python -m pip install --disable-pip-version-check -r requirements.txt

echo "Installing or checking the Chromium browser engine..."
if ! .venv/bin/python -m playwright install chromium; then
  echo "Warning: Chromium setup failed. Direct HTTP scraping will still work."
fi

echo "Starting ScrapeFlow Pro at http://127.0.0.1:8765"
exec .venv/bin/python app.py
