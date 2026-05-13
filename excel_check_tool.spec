# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['excel_check_tool.py'],
    pathex=[],
    binaries=[],
    datas=[('Nothing_You_Could_Do/NothingYouCouldDo-Regular.ttf', 'Nothing_You_Could_Do')],
    hiddenimports=['generate_table_bs', 'PIL', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont'],
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
    name='表格核对工具',
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
)
