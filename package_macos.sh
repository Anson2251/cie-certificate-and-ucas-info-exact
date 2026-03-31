#!/bin/zsh

set -euo pipefail

APP_NAME="CIE Statement & UCAS Extractor"
DIST_DIR="dist"
BUILD_DIR="build/macos"
SPEC_PATH="build/macos-gui.spec"
ZIP_NAME="cie-statement-and-ucas-extractor-macos.zip"

mkdir -p "$BUILD_DIR"

python -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "$APP_NAME" \
  --specpath "$BUILD_DIR" \
  main.py

APP_PATH="$DIST_DIR/$APP_NAME.app"
ZIP_PATH="$DIST_DIR/$ZIP_NAME"

if [ ! -d "$APP_PATH" ]; then
  echo "Expected app bundle not found: $APP_PATH" >&2
  exit 1
fi

rm -f "$ZIP_PATH"
ditto -c -k --sequesterRsrc --keepParent "$APP_PATH" "$ZIP_PATH"

echo "Built app: $APP_PATH"
echo "Built zip: $ZIP_PATH"
