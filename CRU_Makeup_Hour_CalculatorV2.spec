# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['/Users/a016680753/CRU Python Scripts/CRU_Makeup_Hour_CalculatorV2.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='CRU_Makeup_Hour_CalculatorV2',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['CRU Logo.png'],
)
app = BUNDLE(
    exe,
    name='CRU_Makeup_Hour_CalculatorV2.app',
    icon='CRU Logo.png',
    bundle_identifier=None,
)
