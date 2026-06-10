@echo off
REM Compila Gestiq.exe (Windows). Ejecutar con doble clic en un PC Windows.
cd /d "%~dp0"

echo [1/4] Creando entorno virtual...
python -m venv venv || (echo ERROR: instala Python 3.10+ desde python.org & pause & exit /b 1)
call venv\Scripts\activate.bat

echo [2/4] Instalando dependencias...
pip install --upgrade pip
pip install -r requirements.txt pyinstaller

echo [3/4] Descargando Chromium (quedara dentro del .exe)...
set PLAYWRIGHT_BROWSERS_PATH=0
playwright install chromium

echo [4/4] Compilando...
pyinstaller --clean -y gestiq_win.spec

echo.
echo LISTO: dist\Gestiq.exe
pause
