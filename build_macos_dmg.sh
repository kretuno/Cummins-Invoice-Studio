#!/bin/zsh
set -euo pipefail

ROOT="/Users/eduard/Desktop/Total"
DIST_DIR="$ROOT/dist"
APP_NAME="CumminsInvoiceStudio.app"
APP_PATH="$DIST_DIR/$APP_NAME"
STAGING_DIR="$DIST_DIR/dmg_staging"
DMG_PATH="$DIST_DIR/CumminsInvoiceStudio-macOS.dmg"

if [ ! -d "$APP_PATH" ]; then
  echo "App bundle not found: $APP_PATH"
  echo "Build the app first with:"
  echo "  source .venv/bin/activate && PYINSTALLER_CONFIG_DIR=$ROOT/.pyinstaller pyinstaller --noconfirm CumminsInvoiceStudio.spec"
  exit 1
fi

rm -rf "$STAGING_DIR"
mkdir -p "$STAGING_DIR"
cp -R "$APP_PATH" "$STAGING_DIR/"
ln -s /Applications "$STAGING_DIR/Applications"
rm -f "$DMG_PATH"

hdiutil create \
  -volname "Cummins Invoice Studio" \
  -srcfolder "$STAGING_DIR" \
  -ov \
  -format UDZO \
  "$DMG_PATH"

rm -rf "$STAGING_DIR"
echo "Created DMG: $DMG_PATH"
