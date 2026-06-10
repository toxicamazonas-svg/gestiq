# Runtime hook de PyInstaller: hace que Playwright busque Chromium
# dentro del paquete empaquetado (instalado con PLAYWRIGHT_BROWSERS_PATH=0).
import os
os.environ["PLAYWRIGHT_BROWSERS_PATH"] = "0"
