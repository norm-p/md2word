# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for md2word.exe

Bundles pandoc.exe so the resulting binary is fully self-contained.
Build with:  pyinstaller md2word.spec
"""

import os
import shutil

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

# --- Collect magika model files (used by markitdown for file-type detection) ---
magika_datas = collect_data_files("magika", include_py_files=False)
markitdown_datas = collect_data_files("markitdown", include_py_files=False)

# --- Locate pandoc binary ---------------------------------------------------
pandoc_path = shutil.which("pandoc") or os.path.expandvars(
    r"%LOCALAPPDATA%\Pandoc\pandoc.exe"
)
if not os.path.isfile(pandoc_path):
    raise FileNotFoundError(
        f"Cannot find pandoc.exe (looked at: {pandoc_path}). "
        "Install pandoc or set it on PATH before building."
    )

# --- Analysis ----------------------------------------------------------------
a = Analysis(
    ["md2word_entry.py"],
    pathex=["src"],
    binaries=[(pandoc_path, ".")],
    datas=magika_datas + markitdown_datas,
    hiddenimports=[
        "magika",
        # lxml internals PyInstaller often misses
        "lxml._elementpath",
        "lxml.etree",
        "lxml.html",
        # docx / template libs
        "docx",
        "docxtpl",
        "openpyxl",
        # pypandoc
        "pypandoc",
        # click
        "click",
        # dotenv
        "dotenv",
        # AI providers (optional — include so the exe works with any provider)
        "anthropic",
        "openai",
        # markitdown
        "markitdown",
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="md2word",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    icon=None,
)
