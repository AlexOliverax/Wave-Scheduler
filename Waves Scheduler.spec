# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Coleta automaticamente TODOS os submódulos do pacote 'app'
app_hidden = collect_submodules('app')

a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('icon\\WavesS.ico', 'icon'),
        ('icon\\WaveS.ico',  'icon'),
        ('icon\\DHL.gif',    'icon'),
        ('VERSION',          '.'),
    ],
    hiddenimports=app_hidden + [
        # Bibliotecas padrão
        'encodings',
        # openpyxl
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
        'holidays.countries.canada',
        # PyQt5 extras (evitar erros de plataforma)
        'PyQt5',
        'PyQt5.QtWidgets',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        # win32com (Outlook)
        'win32com',
        'win32com.client',
        'win32timezone',
        'pywintypes',
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
