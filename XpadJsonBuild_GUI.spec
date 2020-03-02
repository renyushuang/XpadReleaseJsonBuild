# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['XpadJsonBuild_GUI.py'],
             pathex=['XpadJsonBuild_1.py', 'CoinSDKJsonBuildV_3.py', 'XpadJsonBuild_2.py', 'XpadJsonBuild_data_pip.py', 'BaseXpadJsonBuild.py', '/Users/renyushuang/custom/PythonProject/XpadReleaseJsonBuild'],
             binaries=[],
             datas=[],
             hiddenimports=['XpadJsonBuild_1', 'XpadJsonBuild_2', 'XpadJsonBuild_data_pip', 'BaseXpadJsonBuild', 'CoinSDKJsonBuildV_3'],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='XpadJsonBuild_GUI',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=True )
coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='XpadJsonBuild_GUI')
