#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
VERSION="${VERSION:-1.0.0}"
ARCH="${ARCH:-amd64}"
cd "$ROOT_DIR"

python3 -m venv .venv-build
source .venv-build/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
pyinstaller --clean --noconfirm combina-word.spec

PKG="${ROOT_DIR}/build/deb/combina-word_${VERSION}_${ARCH}"
rm -rf "$PKG"
mkdir -p "$PKG/DEBIAN" "$PKG/opt/combina-word" "$PKG/usr/bin" "$ROOT_DIR/dist"
cp dist/CombinaWord "$PKG/opt/combina-word/CombinaWord"
chmod 0755 "$PKG/opt/combina-word/CombinaWord"
cat > "$PKG/usr/bin/combina-word" <<'LAUNCHER'
#!/bin/sh
exec /opt/combina-word/CombinaWord "$@"
LAUNCHER
chmod 0755 "$PKG/usr/bin/combina-word"
cat > "$PKG/DEBIAN/control" <<CONTROL
Package: combina-word
Version: ${VERSION}
Section: utils
Priority: optional
Architecture: ${ARCH}
Maintainer: Combina Word Team
Description: Combina una plantilla Word con filas de Excel
 Aplicacion grafica Tkinter para generar un unico documento Word.
CONTROL
dpkg-deb --build "$PKG" "$ROOT_DIR/dist/combina-word_${VERSION}_${ARCH}.deb"
echo "Paquete creado: $ROOT_DIR/dist/combina-word_${VERSION}_${ARCH}.deb"
