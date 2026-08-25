# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller onedir spec — ITGC SAP HANA APP (Windows)."""
from pathlib import Path

SPECPATH = Path(SPECPATH).resolve()
ROOT = SPECPATH

datas: list = []
binaries: list = []
hiddenimports: list = [
    "PySide6.QtCore",
    "PySide6.QtGui",
    "PySide6.QtWidgets",
    "win32com",
    "win32com.client",
    "pythoncom",
    "pywintypes",
    "openpyxl",
    "PIL",
    "PIL.Image",
]

kb_dir = ROOT / "data" / "knowledge_base"
if kb_dir.is_dir():
    datas.append((str(kb_dir), "data/knowledge_base"))

assets_dir = ROOT / "src" / "ui" / "assets"
if assets_dir.is_dir():
    datas.append((str(assets_dir), "src/ui/assets"))

a = Analysis(
    [str(ROOT / "src" / "main.py")],
    pathex=[str(ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "PySide6.QtWebEngineCore",
        "PySide6.QtWebEngineWidgets",
        "PySide6.QtWebEngineQuick",
        "PySide6.Qt3DCore",
        "PySide6.Qt3DRender",
        "PySide6.Qt3DInput",
        "PySide6.Qt3DLogic",
        "PySide6.Qt3DAnimation",
        "PySide6.Qt3DExtras",
        "PySide6.QtMultimedia",
        "PySide6.QtMultimediaWidgets",
        "PySide6.QtBluetooth",
        "PySide6.QtNfc",
        "PySide6.QtPositioning",
        "PySide6.QtSensors",
        "PySide6.QtPdf",
        "PySide6.QtPdfWidgets",
        "tkinter",
        "matplotlib",
        "numpy",
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="ITGC_SAP_APP",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="ITGC_SAP_APP",
)
