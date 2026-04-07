# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files, collect_submodules


SPEC_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
PROJECT_ROOT = SPEC_DIR.parent

asset_datas = [
    (str(PROJECT_ROOT / '00 dashboard.html'), '.'),
    (str(SPEC_DIR / '000 Launch_dashboard.bat'), '.'),
    (str(SPEC_DIR / '웹접속 주소.txt'), '.'),
    (str(PROJECT_ROOT / 'index.html'), '.'),
    (str(PROJECT_ROOT / 'manifest.webmanifest'), '.'),
    (str(PROJECT_ROOT / 'service-worker.js'), '.'),
    (str(PROJECT_ROOT / 'automated_scripts'), 'automated_scripts'),
    (str(PROJECT_ROOT / 'system_guides'), 'system_guides'),
]

hiddenimports = (
    collect_submodules('openpyxl')
    + collect_submodules('win32com')
    + collect_submodules('fitz')
    + collect_submodules('psutil')
    + [
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
        '_tkinter',
        'pythoncom',
        'pywintypes',
    ]
)

datas = (
    asset_datas
    + collect_data_files('openpyxl')
    + collect_data_files('fitz')
    + collect_data_files('win32com')
)


a = Analysis(
    [str(SPEC_DIR / 'dashboard_agent_launcher.py')],
    pathex=[str(SPEC_DIR), str(PROJECT_ROOT)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='WYGGKR02_Dashboard_Agent',
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
