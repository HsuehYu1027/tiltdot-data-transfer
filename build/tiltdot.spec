# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec：onedir + console，並把 Playwright 的 Chromium 一起打包。

打包方式（在 repo 根目錄）：
    set PLAYWRIGHT_BROWSERS_PATH=0
    python -m playwright install chromium
    python -m PyInstaller --noconfirm --clean build/tiltdot.spec
"""
import os
from PyInstaller.utils.hooks import collect_all

ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))

datas, binaries, hiddenimports = [], [], []
for _pkg in ("playwright",):
    _d, _b, _h = collect_all(_pkg)
    datas += _d
    binaries += _b
    hiddenimports += _h

a = Analysis(
    [os.path.join(ROOT, "tiltdotdatatransfer.py")],
    pathex=[ROOT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + ["openpyxl", "PIL.Image"],
    hookspath=[],
    runtime_hooks=[os.path.join(SPECPATH, "rthook_playwright.py")],
    excludes=[
        "tkinter", "PyQt5", "PyQt6", "PySide2", "PySide6",
        "IPython", "jupyter", "notebook", "pytest", "sphinx",
        "scipy", "matplotlib.backends._backend_tk",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="tiltdot",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,          # 程式靠 input() 互動，絕對不能用 --windowed
    disable_windowed_traceback=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name="tiltdot",
)
