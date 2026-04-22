#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN="${PYTHON_BIN:-python3}"

cd "$ROOT_DIR"

if [[ ! -f "$ROOT_DIR/LOGO.icns" ]]; then
  bash "$ROOT_DIR/build_macos_icon.sh"
fi

"$PYTHON_BIN" -m pip install --upgrade pip
"$PYTHON_BIN" -m pip install -r "$ROOT_DIR/requirements-mac.txt"
"$PYTHON_BIN" -m PyInstaller --clean "$ROOT_DIR/cost_calc_v95_mac.spec"

APP_PATH="$(find "$ROOT_DIR/dist" -maxdepth 2 -type d -name '*.app' | head -n 1)"
if [[ -z "${APP_PATH:-}" ]]; then
  echo "Build finished but no .app bundle was found under dist/" >&2
  exit 1
fi

echo "App build complete: $APP_PATH"
