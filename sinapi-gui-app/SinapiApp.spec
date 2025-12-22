# -*- mode: python ; coding: utf-8 -*-
import os

# Diretório onde o .spec está
SPEC_DIR = os.path.dirname(os.path.abspath(__file__))
# Diretório do código fonte
SRC_DIR = os.path.join(SPEC_DIR, 'src')

a = Analysis(
    [os.path.join(SRC_DIR, 'main.py')],
    pathex=[SRC_DIR],
    binaries=[],
    datas=[(os.path.join(SRC_DIR, 'LINKS_SICRO.txt'), '.')],
    hiddenimports=['pandas._libs.tslibs'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='SinapiApp',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Equivalente a --windowed
    icon=None
)
