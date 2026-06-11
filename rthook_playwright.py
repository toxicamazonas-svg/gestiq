# Runtime hook de PyInstaller: hace que Playwright use el Chromium
# que viaja junto al ejecutable (carpeta ms-playwright).
import os
import sys

if getattr(sys, "frozen", False):
    _base = os.path.dirname(sys.executable)
    _candidatos = [
        os.path.join(_base, "ms-playwright"),  # Windows: carpeta junto al exe
        os.path.normpath(os.path.join(_base, "..", "Resources", "ms-playwright")),  # Mac: dentro del .app
    ]
    for _c in _candidatos:
        if os.path.isdir(_c):
            os.environ["PLAYWRIGHT_BROWSERS_PATH"] = _c
            break
