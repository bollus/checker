# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

block_cipher = None
root = Path.cwd()

datas = []
for candidate in [
    root / "check_templates.json",
    root / "position_aliases.json",
]:
    if candidate.exists():
        datas.append((str(candidate), "."))

font_dir = root / "Nothing_You_Could_Do"
if font_dir.exists():
    datas.append((str(font_dir), "Nothing_You_Could_Do"))

a = Analysis(
    ["python_backend/backend_cli.py"],
    pathex=[str(root)],
    binaries=[],
    datas=datas,
    hiddenimports=["excel_check_tool", "generate_table_bs", "PIL", "PIL.Image", "PIL.ImageDraw", "PIL.ImageFont"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    [],
    name="excel-check-backend",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    exclude_binaries=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="excel-check-backend",
)
