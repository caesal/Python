# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['I2CFW_GUI_Ver0.4.py'],
    pathex=[],
    binaries=[],
    datas=[('jaguar-b0_mca_i2c_isp_driver_payload_Disable_WDT.bin', '.')],
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
    name='I2CFW_GUI_Ver0.4',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['I2C_application_icon.ico'],
)
