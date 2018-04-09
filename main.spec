# -*- mode: python -*-

block_cipher = None


a = Analysis(['main.py'],
             pathex=['.'],
             binaries=[],
             datas=[],
             hiddenimports=[],
             hookspath=[],
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)
exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [('bg.jpg','bg.jpg','DATA')],
          [('rules.txt','rules.txt','DATA')],
          name='申报表自动审查辅助工具',
          debug=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False )
