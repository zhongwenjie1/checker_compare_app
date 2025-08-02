# CheckerUI_win.spec（Windows 专用）
# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import (
    collect_submodules,
    collect_data_files,
    collect_dynamic_libs,
)
import os

# 可选静态资源（存在才打包）
extra_datas = []
if os.path.isdir('themes'):
    extra_datas += [('themes/*.qss', 'themes')]
if os.path.isdir('assets'):
    extra_datas += [('assets/*', 'assets')]

hiddenimports = collect_submodules('numpy')  # 确保 numpy 子模块

datas = (
    collect_data_files('pandas') +  # pandas 样式/配置等
    extra_datas
)

binaries = (
    collect_dynamic_libs('pandas') +
    collect_dynamic_libs('numpy', destdir='numpy.libs')  # 放到独立目录，避免 from-source 报错
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
    console=False,                 # 用户版 GUI；如要 Console 调试版，改 True
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

# Windows 下最终是 dist/CheckerUI/ 下的文件夹结构（包含 CheckerUI.exe）