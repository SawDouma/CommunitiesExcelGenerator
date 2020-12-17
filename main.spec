# -*- mode: python ; coding: utf-8 -*-

block_cipher = None


a = Analysis(['main.py', 'CommunityExcelGenerator.py', 'dataclass\\Community.py'],
             pathex=['dataclass','venv\\lib\\site-packages'],
             binaries=[],
             datas=[],
             hiddenimports=['openpyxl'],
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
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='社区表格批量生成器',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False )
