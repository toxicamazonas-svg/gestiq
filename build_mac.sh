#!/bin/bash
# Compila Gestiq.app (Mac) y lo comprime en zip para distribuir.
set -e
cd "$(dirname "$0")"

echo "[1/5] Creando entorno virtual..."
python3 -m venv venv
source venv/bin/activate

echo "[2/5] Instalando dependencias..."
pip install --upgrade pip
pip install -r requirements.txt pyinstaller

echo "[3/5] Descargando Chromium..."
playwright install chromium

echo "[4/5] Compilando..."
pyinstaller --clean -y gestiq_mac.spec
ditto "$HOME/Library/Caches/ms-playwright" "dist/Gestiq.app/Contents/Resources/ms-playwright"
codesign --force --deep -s - "dist/Gestiq.app"

echo "[5/5] Comprimiendo..."
mkdir -p dist/paquete
cp -R "dist/Gestiq.app" NOTA_MAC.txt dist/paquete/
ditto -c -k dist/paquete "dist/Gestiq-Mac.zip"

echo ""
echo "LISTO: dist/Gestiq.app  y  dist/Gestiq-Mac.zip"
