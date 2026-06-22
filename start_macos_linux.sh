#!/usr/bin/env sh
set -eu

cd "$(dirname "$0")"

if command -v python3 >/dev/null 2>&1; then
  PYTHON_CMD=python3
elif command -v python >/dev/null 2>&1; then
  PYTHON_CMD=python
else
  echo "Python is not installed. Install Python 3.11 or newer."
  exit 1
fi

if [ ! -x ".venv/bin/python" ]; then
  echo "Creating local Python environment..."
  "$PYTHON_CMD" -m venv .venv
fi

echo "Installing/updating dependencies..."
".venv/bin/python" -m pip install --upgrade pip
".venv/bin/python" -m pip install -r requirements.txt

echo "Starting Budget Review App at http://127.0.0.1:5057"
if command -v open >/dev/null 2>&1; then
  open http://127.0.0.1:5057 >/dev/null 2>&1 || true
elif command -v xdg-open >/dev/null 2>&1; then
  xdg-open http://127.0.0.1:5057 >/dev/null 2>&1 || true
fi

".venv/bin/python" app.py

