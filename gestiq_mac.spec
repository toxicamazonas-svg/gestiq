# -*- mode: python ; coding: utf-8 -*-
# Spec Mac: Gestiq.app (windowed, Chromium embebido)
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], []
for pkg in ("playwright", "customtkinter", "keyring"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h
hiddenimports += ["keyring.backends.macOS"]

import os as _os
for _logo in ("logo_imagine.png", "logo_guardian.png"):
    if _os.path.exists(_logo):
        datas.append((_logo, "."))

a = Analysis(
    ["gestiq.py"],
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
    [],
    exclude_binaries=True,
    name="Gestiq",
    debug=False,
    strip=False,
    upx=False,
    console=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="Gestiq",
)
app = BUNDLE(
    coll,
    name="Gestiq.app",
    icon=None,
    bundle_identifier="com.gestiq.app",
    info_plist={
        "NSHighResolutionCapable": True,
        "LSMinimumSystemVersion": "11.0",
    },
)
