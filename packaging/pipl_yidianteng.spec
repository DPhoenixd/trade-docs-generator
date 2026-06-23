# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


ROOT = Path(SPECPATH).parent


datas = [
    (str(ROOT / "pipl-frontend" / "dist"), "pipl-frontend/dist"),
    (str(ROOT / "fabric_master_en.csv"), "."),
    (str(ROOT / "fabric_database_en.xlsx"), "."),
    (str(ROOT / "fabric_price_rules.csv"), "."),
] + collect_data_files("rapidocr_onnxruntime")

hiddenimports = (
    collect_submodules("uvicorn")
    + collect_submodules("rapidocr_onnxruntime")
    + collect_submodules("cv2")
)


a = Analysis(
    [str(ROOT / "desktop_launcher.py")],
    pathex=[str(ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["streamlit"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="PIPL-YiDianTeng",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
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
    upx=True,
    upx_exclude=[],
    name="PIPL-YiDianTeng",
)
