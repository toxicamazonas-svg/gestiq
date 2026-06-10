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

echo "[3/5] Descargando Chromium (quedara dentro del .app)..."
export PLAYWRIGHT_BROWSERS_PATH=0
playwright install chromium

echo "[4/5] Compilando..."
pyinstaller --clean -y gestiq_mac.spec

echo "[5/5] Comprimiendo..."
cp NOTA_MAC.txt dist/ 2>/dev/null || true
cd dist && ditto -c -k --keepParent "Gestiq.app" "Gestiq-Mac.zip" && cd ..

echo ""
echo "LISTO: dist/Gestiq.app  y  dist/Gestiq-Mac.zip"
