# -*- mode: python ; coding: utf-8 -*-

# -*- mode: python ; coding: utf-8 -*-
import os
# Use a fixed extraction directory so embedded JSON persists across launches
runtime_tmpdir = os.path.expanduser(os.path.join("~", ".visual_tmp"))
os.makedirs(runtime_tmpdir, exist_ok=True)

a = Analysis(
    ['invoice_visual.py'],
    pathex=[],
    binaries=[],
    datas=[('program files/pipe_series_sdr.csv', 'program files'), ('program files/DIN_pivot.csv', 'program files'), ('program files/logo.png', 'program files'), ('program files/DejaVuSans.ttf', 'program files'), ('program files/PN_TO_SDR.CSV', 'program files'), ('program files/discount.csv', 'program files'), ('program files/invoice_counter.json', '.')],
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
    name='visual',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=runtime_tmpdir,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
app = BUNDLE(
    exe,
    name='Pipecal.app',
    icon=None,
    bundle_identifier=None,
)
