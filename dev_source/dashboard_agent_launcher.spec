# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

SPEC_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
PROJECT_ROOT = SPEC_DIR.parent

datas = [
    (str(PROJECT_ROOT / '00 dashboard.html'), '.'),
    (str(SPEC_DIR / '000 Launch_dashboard.bat'), '.'),
    (str(SPEC_DIR / '웹접속 주소.txt'), '.'),
    (str(PROJECT_ROOT / 'index.html'), '.'),
    (str(PROJECT_ROOT / 'manifest.webmanifest'), '.'),
    (str(PROJECT_ROOT / 'service-worker.js'), '.'),
    (str(PROJECT_ROOT / 'system_guides'), 'system_guides'),
    (str(SPEC_DIR / 'run_dashboard.py'), 'dev_source'),
    (str(SPEC_DIR / 'runtime_store' / 'runtime_manifest.json'), 'runtime_store'),
]


a = Analysis(
    [str(SPEC_DIR / 'dashboard_agent_launcher.py')],
    pathex=[str(SPEC_DIR), str(PROJECT_ROOT)],
    binaries=[],
    datas=datas,
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
