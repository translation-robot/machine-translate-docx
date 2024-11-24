# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/newmm_tokenizer/words_th.txt', 'newmm_tokenizer'), ('/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/parsivar/resource/stemmer/*', 'parsivar/resource/stemmer'), ('/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/parsivar/resource/normalizer/*', 'parsivar/resource/normalizer'), ('/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/parsivar/resource/tokenizer/*', 'parsivar/resource/tokenizer'), ('/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/demoji/codes.json', 'demoji'), ('/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/usaddress/usaddr.crfsuite', 'usaddress'), ('/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/gensim/test/test_data/lee_background.cor', 'gensim/test/test_data'), ('/Library/Frameworks/Python.framework/Versions/3.11/lib/python3.11/site-packages/hazm/data/*', 'hazm/data')]
binaries = []
hiddenimports = ['sklearn', 'usaddress', 'hazm']
tmp_ret = collect_all('pycrfsuite')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('scipy')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('sklearn')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


block_cipher = None


a = Analysis(
    ['machine-translate-docx.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name='Machine Translator Term',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Machine Translator Term',
)
app = BUNDLE(coll,
             name='Machine Translator Term.app',
             icon='Machine Translator.icns',
             bundle_identifier=None) 
