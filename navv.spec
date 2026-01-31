# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

# -----------------------------
# Package data (NAVV assets)
# -----------------------------
# Bundle NAVV package data (e.g., data.pkl, mac-vendors.json)
datas = [
    ('src/navv/data', 'navv/data'),
]
binaries = []
hiddenimports = []

# -----------------------------
# Explicit excludes
# -----------------------------
EXCLUDES = [
    'pytest',
    '_pytest',
    'hypothesis',

    # pandas test trees
    'pandas.tests',
    'pandas.conftest',

    # numpy test trees
    'numpy.tests',
    'numpy.conftest',
    'numpy.testing.tests',
]

# -----------------------------
# Helper: filter bad hidden imports
# -----------------------------
def _filter_hiddenimports(his):
    bad_prefixes = (
        'pytest',
        '_pytest',
        'hypothesis',
        'pandas.tests',
        'pandas.conftest',
        'numpy.tests',
        'numpy.conftest',
        'numpy.testing',
    )

    def is_bad(m):
        if m.startswith(bad_prefixes):
            return True
        if m.startswith('numpy.') and '.tests' in m:
            return True
        if m.startswith('pandas.') and '.tests' in m:
            return True
        return False

    return [m for m in his if not is_bad(m)]

# -----------------------------
# Runtime dependencies
# -----------------------------

# click – CLI framework (minimal)
hiddenimports += [
    'click',
    'click.core',
    'click.decorators',
    'click.formatting',
    'click.parser',
    'click.termui',
    'click.types',
    'click.utils',
]

# tqdm – progress bars (minimal; avoid tensorflow/tkinter hooks)
hiddenimports += [
    'tqdm',
    'tqdm.auto',
    'tqdm.std',
    'tqdm.utils',
    'tqdm._utils',
    'tqdm._tqdm',
]

# openpyxl – Excel writer (targeted)
hiddenimports += [
    'openpyxl',
    'openpyxl.workbook',
    'openpyxl.worksheet',
    'openpyxl.writer',
    'openpyxl.writer.excel',
    'openpyxl.cell._writer',
    'openpyxl.styles.numbers',
    'openpyxl.worksheet._writer',
]

#
# pandas/numpy: rely on PyInstaller's built-in hooks (hook-pandas / hook-numpy)
# to collect the required binaries and data. Avoid duplicating package data here,
# because it bloats the onefile and slows startup (unpack time).

# Safety-net hidden imports PyInstaller sometimes misses
hiddenimports += [
    # pandas internal config modules that may be imported dynamically
    'pandas._config.config',
    'pandas._config.localization',

    # pandas excel/style paths that are sometimes imported dynamically
    'pandas.io.formats.excel',
    'pandas.io.formats.style',

    # pandas internal testing module
    'pandas.testing',

    # tqdm convenience wrapper
    'tqdm.auto',
]

hiddenimports = _filter_hiddenimports(hiddenimports)

# -----------------------------
# Analysis
# -----------------------------
a = Analysis(
    ['src/navv/__main__.py'],
    pathex=['src'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
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
    name='navv',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
