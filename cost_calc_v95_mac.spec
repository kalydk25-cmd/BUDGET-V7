# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = []
hiddenimports += collect_submodules("openpyxl")


a = Analysis(
    ["cost_calc_v95_mac.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("LOGO.png", "."),
        ("LOGO.ico", "."),
        ("task_config_overrides.json", "."),
        ("fixed_staff_wage_data.py", "."),
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pandas", "matplotlib", "numpy", "PySide6", "shiboken6", "sqlalchemy"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="RETEC-CostCalc-V95-Mac",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=True,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

app = BUNDLE(
    exe,
    name="RETEC-CostCalc-V95-Mac.app",
    icon="LOGO.icns",
    bundle_identifier="com.retec.costcalc.v95.mac",
    info_plist={
        "NSPrincipalClass": "NSApplication",
    },
)

coll = COLLECT(
    app,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="RETEC-CostCalc-V95-Mac",
)
