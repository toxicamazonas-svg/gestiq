# -*- mode: python ; coding: utf-8 -*-
# Spec Windows: Gestiq.exe (onefile, sin consola, Chromium embebido)
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], []
for pkg in ("playwright", "customtkinter", "keyring", "webview"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h
hiddenimports += ["keyring.backends.Windows", "win32ctypes.pywin32",
                  "webview.platforms.edgechromium", "webview.platforms.winforms",
                  "clr", "clr_loader", "pythonnet"]

import os as _os
for _f in ("logo_imagine.png", "logo_guardian.png", "gestiq_ui.html"):
    if _os.path.exists(_f):
        datas.append((_f, "."))

a = Analysis(
    ["gestiq_web.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    runtime_hooks=["rthook_playwright.py"],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Gestiq",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
)
