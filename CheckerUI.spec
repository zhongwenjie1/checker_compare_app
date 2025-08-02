# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import (
    collect_submodules,
    collect_data_files,
    collect_dynamic_libs,
)
import os

# Collect optional resources only if present
extra_datas = []
if os.path.isdir('themes'):
    extra_datas += [('themes/*.qss', 'themes')]
if os.path.isdir('assets'):
    extra_datas += [('assets/*', 'assets')]

# Build spec-time lists using stable collectors (avoid collect_all tuple confusions)
hiddenimports = (
    # Let official hooks bring in PySide6 & pandas; only ensure numpy submodules.
    collect_submodules('numpy')
)

datas = (
    # Do NOT collect PySide6 or numpy data to avoid duplicate Qt frameworks or numpy source trees.
    collect_data_files('pandas') +
    extra_datas
)

binaries = (
    # Avoid collecting PySide6 binaries explicitly; the hook already includes Qt frameworks.
    collect_dynamic_libs('pandas') +
    collect_dynamic_libs('numpy', destdir='numpy.libs')
)

a = Analysis(
    ['checker_ui/main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='CheckerUI',
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
    name='CheckerUI',
)
app = BUNDLE(
    coll,
    name='CheckerUI.app',
    icon=None,
    bundle_identifier=None,
)
