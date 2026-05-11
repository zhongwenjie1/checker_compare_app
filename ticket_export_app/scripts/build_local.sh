#!/usr/bin/env bash
set -e

cd "$(dirname "$0")/.."

echo "== ticket_export_app build environment check =="
echo "Project dir: $(pwd)"

if [ ! -f "version.json" ]; then
  echo "ERROR: version.json not found."
  exit 1
fi

echo ""
echo "== version.json =="
cat version.json
echo ""

if [ ! -f "main.py" ]; then
  echo "ERROR: main.py not found."
  exit 1
fi

PYTHON_BIN="${PYTHON:-python3}"

if ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  echo "ERROR: Python executable not found: $PYTHON_BIN"
  echo "You can set PYTHON=/path/to/python before running this script."
  exit 1
fi

echo ""
echo "== Python =="
"$PYTHON_BIN" --version

echo ""
echo "== PyInstaller =="
if "$PYTHON_BIN" -m PyInstaller --version >/tmp/ticket_export_pyinstaller_version.txt 2>/tmp/ticket_export_pyinstaller_error.txt; then
  echo "PyInstaller available: $(cat /tmp/ticket_export_pyinstaller_version.txt)"
else
  echo "PyInstaller is not installed or not available in this Python environment."
  echo "Install later with:"
  echo "  $PYTHON_BIN -m pip install pyinstaller"
fi

echo ""
echo "Build environment check completed."
