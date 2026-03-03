# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('icon\\WavesS.ico', 'icon'),
        ('icon\\WaveS.ico', 'icon'),
        ('icon\\DHL.gif', 'icon'),
    ],
    hiddenimports=[
        # Módulos do app
        'app',
        'app.ui',
        'app.core',
        'app.utils',
        'app.outlook',
        'app.history',
        'app.dialogs',
        # Bibliotecas padrão
        'encodings', 
        'openpyxl', 
        'openpyxl.cell._writer',
        # Holidays e países
        'holidays',
        'holidays.countries',
        'holidays.countries.brazil',
        'holidays.countries.united_states',
        'holidays.countries.argentina',
        'holidays.countries.chile',
        'holidays.countries.peru',
        'holidays.countries.colombia',
        'holidays.countries.mexico',
        'holidays.countries.canada'
    ],
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
    [],
    exclude_binaries=True,
    name='Waves Scheduler',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['icon\\WavesS.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Waves Scheduler',
)
