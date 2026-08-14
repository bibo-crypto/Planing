# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec file for Planing
# Run via build.bat — do NOT run manually without activating the venv first.

from PyInstaller.utils.hooks import collect_all

block_cipher = None

# pdfplumber renders/reads PDFs through pypdfium2, which in turn loads a
# native pdfium.dll/.so from the separate `pypdfium2_raw` package at
# runtime via ctypes.CDLL(). PyInstaller's static import analysis cannot
# see that file — it only follows Python imports — so without explicitly
# collecting it here, the frozen .exe builds "successfully" but crashes
# the moment it tries to open a PDF on ANY machine that doesn't happen to
# have a matching pypdfium2 install on its PATH (including, sometimes,
# the build machine itself, purely by accident of leftover global installs).
datas = []
binaries = []
extra_hiddenimports = []
for _pkg in ("pypdfium2", "pypdfium2_raw", "pandas", "numpy"):
    _datas, _binaries, _hidden = collect_all(_pkg)
    datas += _datas
    binaries += _binaries
    extra_hiddenimports += _hidden

# Bundle icon.ico as a data file too (not just as the .exe's icon below) so
# gui.py can load it at runtime and set the window/taskbar icon while the
# app is running — the exe icon alone only affects the file's icon in
# Explorer, not the running window.
datas += [('icon.ico', '.')]

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=[
        # pdfplumber / pdfminer
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.converter',
        'pdfminer.pdfpage',
        'pdfminer.pdfinterp',
        'pdfminer.pdfdevice',
        'pdfminer.utils',
        'pdfminer.pdfdocument',
        'pdfminer.pdfparser',
        'pdfminer.cmapdb',
        'pdfminer.encodingdb',
        'pdfminer.fontmetrics',
        'pdfminer.glyphlist',
        'pdfminer.image',
        'pdfminer.jbig2',
        'pdfminer.lzw',
        'pdfminer.pdffont',
        'pdfminer.pdfcolor',
        'pdfminer.psparser',
        'pdfminer.ascii85',
        'pdfminer.ccitt',
        'pdfminer.runlength',
        # pypdfium2 (PDF rendering backend used by pdfplumber)
        'pypdfium2',
        'pypdfium2_raw',
        # Pillow (used by pdfplumber for image processing)
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFilter',
        # cryptography (pdfplumber dependency)
        'cryptography',
        'cryptography.hazmat',
        'cryptography.hazmat.primitives',
        'cryptography.hazmat.backends',
        # openpyxl
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.styles.fills',
        'openpyxl.styles.fonts',
        'openpyxl.styles.borders',
        'openpyxl.styles.alignment',
        'openpyxl.utils',
        'openpyxl.utils.cell',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.worksheet.worksheet',
        'et_xmlfile',
        # pandas C-extension modules are imported dynamically by pandas on
        # some Python/Windows combinations and need explicit registration in
        # the frozen application.
        'pandas._libs.pandas_parser',
        'pandas._libs.pandas_datetime',
        'pandas._libs.tslibs.base',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.timestamps',
        # tkinter
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        # stdlib extras
        'logging.handlers',
        'pathlib',
        'threading',
        'json',
        'dataclasses',
        'collections',
        'charset_normalizer',
        'sqlite3',
    ] + extra_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=['runtime_hook.py'],
    excludes=[
        'scipy',
        'IPython',
        'notebook',
        'pytest',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Planing',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,          # no black console window — GUI only
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Planing',
)
