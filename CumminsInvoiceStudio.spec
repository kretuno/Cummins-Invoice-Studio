# -*- mode: python ; coding: utf-8 -*-

block_cipher = None
app_info_plist = {
    'CFBundleName': 'Cummins Invoice Studio',
    'CFBundleDisplayName': 'Cummins Invoice Studio',
    'CFBundleShortVersionString': '1.0.8',
    'CFBundleVersion': '1.0.8',
    'CFBundleIdentifier': 'com.eduard.cummins.invoicestudio',
    'NSHumanReadableCopyright': 'Eduard Osipov',
}


a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[('sample_data', 'sample_data'), ('assets', 'assets')],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='CumminsInvoiceStudio',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    icon='assets/app_icon.icns',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='CumminsInvoiceStudio',
)

app = BUNDLE(
    coll,
    name='CumminsInvoiceStudio.app',
    icon='assets/app_icon.icns',
    bundle_identifier='com.eduard.cummins.invoicestudio',
    info_plist=app_info_plist,
)
