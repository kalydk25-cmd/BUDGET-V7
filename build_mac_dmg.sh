#!/bin/bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
STAGING_DIR="$ROOT_DIR/dist/dmg-staging"
DMG_PATH="$ROOT_DIR/dist/RETEC-CostCalc-V95-Mac.dmg"
VOL_NAME="RETEC CostCalc V95 Mac"

if ! command -v hdiutil >/dev/null 2>&1; then
  echo "Missing required macOS tool: hdiutil" >&2
  exit 1
fi

APP_PATH="$(find "$ROOT_DIR/dist" -maxdepth 2 -type d -name '*.app' | head -n 1)"

if [[ -z "${APP_PATH:-}" || ! -d "$APP_PATH" ]]; then
  echo "Missing app bundle under: $ROOT_DIR/dist" >&2
  echo "Run build_mac_app.sh first." >&2
  exit 1
fi

rm -rf "$STAGING_DIR"
mkdir -p "$STAGING_DIR"
cp -R "$APP_PATH" "$STAGING_DIR/"
ln -s /Applications "$STAGING_DIR/Applications"
rm -f "$DMG_PATH"

hdiutil create \
  -volname "$VOL_NAME" \
  -srcfolder "$STAGING_DIR" \
  -ov \
  -format UDZO \
  "$DMG_PATH"

echo "DMG build complete: $DMG_PATH"
