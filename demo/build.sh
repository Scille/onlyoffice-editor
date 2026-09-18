#!/usr/bin/env bash
# Assemble a vendored OnlyOffice `editor/` directory (the tree the parsec-cloud
# project consumes at client/vendors/onlyoffice/editor/) from the *source* in
# this repo (the CryptPad fork of sdkjs + web-apps + the onlyoffice-editor
# wrapper).
#
# This mirrors the Dockerfile `files-build` / `zip-build` stages, but runs
# locally (no Docker) so you can tweak the OnlyOffice source and rebuild fast.
#
# Usage:
#   ./build.sh [DEST]
#     DEST defaults to ./build/editor
#
# What it does:
#   1. Build sdkjs & web-apps                   -> REPO/sdkjs/deploy
#   2. Build onlyoffice-editor                  -> REPO/onlyoffice-editor/dist
#   3. Copies the built sdkjs deploy          -> DEST/sdkjs
#   4. Copies the built web-apps deploy      -> DEST/web-apps
#   5. Overlays repo vendor/                -> DEST/web-apps/vendor   (requirejs)
#   6. Copies repo fonts/*.ttf|*.otf         -> DEST/fonts/fonts
#   7. Copies repo dictionaries/             -> DEST/dictionaries
#   8. Swaps the API entry: built-vanilla api.js -> api-orig.js,
#      wrapper dist/api.js                   -> DEST/web-apps/apps/api/documents/api.js
#   9. (optional) Brotli-compresses js/css/html/wasm/svg/aff/dic -> *.br
#
# Then point parsec-cloud at DEST (see "Wiring it into parsec-cloud" below).
set -euo pipefail

REPO="$(cd "$(dirname "$0")" && cd .. && pwd)"

# Prerequisites: build sdkjs&web-apps and cryptpad wrapper

cd "$REPO/sdkjs"
make
cd "$REPO"

cd "$REPO/onlyoffice-editor"
npm run build
cd "$REPO"

# Now assemble

DEST="${1:-$REPO/build/editor}"

SDKJS_DEPLOY="$REPO/sdkjs/deploy"
WRAPPER_DIST="$REPO/onlyoffice-editor/dist"

need() { [ -e "$1" ] || { echo "missing: $1" >&2; exit 1; }; }
need "$SDKJS_DEPLOY/sdkjs/word/sdk-all-min.js"
need "$SDKJS_DEPLOY/web-apps/apps/documenteditor/main/app.js"
need "$WRAPPER_DIST/api.js"

echo "assembling $DEST"
rm -rf "$DEST"
mkdir -p "$DEST"

# 1+2: built sdk + web-apps
cp -r "$SDKJS_DEPLOY/sdkjs" "$DEST/sdkjs"
cp -r "$SDKJS_DEPLOY/web-apps" "$DEST/web-apps"

# 3: repo vendor overlay (requirejs etc.) on top of the web-apps build's vendor
mkdir -p "$DEST/web-apps/vendor"
cp -r "$REPO/vendor/." "$DEST/web-apps/vendor/"

# 4: fonts
mkdir -p "$DEST/fonts/fonts"
cp "$REPO"/fonts/*.ttf "$REPO"/fonts/*.otf "$DEST/fonts/fonts/" 2>/dev/null || true

# 5: dictionaries
cp -r "$REPO/dictionaries" "$DEST/dictionaries"

# 6: API entry swap (vanilla built api.js -> api-orig.js, wrapper -> api.js)
API_DIR="$DEST/web-apps/apps/api/documents"
mv "$API_DIR/api.js" "$API_DIR/api-orig.js"
cp "$WRAPPER_DIST/api.js" "$API_DIR/api.js"

# (optional) standalone polyfill the editor HTML references; the local grunt
# build doesn't emit a closure-compiled one, so copy the source verbatim if you
# want the file present. Non-essential — the editor runs without it.
if [ ! -e "$DEST/sdkjs/vendor/polyfill.js" ] && [ -f "$REPO/sdkjs/vendor/polyfill.js" ]; then
  mkdir -p "$DEST/sdkjs/vendor"
  cp "$REPO/sdkjs/vendor/polyfill.js" "$DEST/sdkjs/vendor/polyfill.js"
fi

echo "assembled $DEST ($(find "$DEST" -type f | wc -l) files)"

# 7: brotli (opt-in; the release zip ships .br sidecars and parsec-cloud's
# vite build also re-compresses, so this is only needed if you serve the
# folder directly with a static server that prefers pre-compressed assets)
if [ "${BROTLI:-0}" = "1" ]; then
  echo "brotli-compressing..."
  find "$DEST" \( -name '*.js' -o -name '*.css' -o -name '*.html' \
                -o -name '*.wasm' -o -name '*.svg' -o -name '*.aff' -o -name '*.dic' \) \
    -print0 | xargs -0 -P 8 -n 16 -- brotli -q 11 -k
fi
