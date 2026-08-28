# -*- coding: utf-8 -*-
"""PyInstaller runtime hook：讓打包後的程式找到內嵌的 Chromium。

打包時用 PLAYWRIGHT_BROWSERS_PATH=0 安裝，瀏覽器會被放進 playwright 套件目錄，
這裡在啟動時把該路徑指回去，避免程式跑去找使用者電腦上不存在的 ms-playwright。
"""
import os
import sys

if getattr(sys, "frozen", False):
    _base = getattr(sys, "_MEIPASS", None) or os.path.dirname(os.path.abspath(sys.executable))
    for _candidate in (
        os.path.join(_base, "playwright", "driver", "package", ".local-browsers"),
        os.path.join(_base, "_internal", "playwright", "driver", "package", ".local-browsers"),
    ):
        if os.path.isdir(_candidate):
            os.environ.setdefault("PLAYWRIGHT_BROWSERS_PATH", _candidate)
            break
