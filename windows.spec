# -*- mode: python -*-

block_cipher = None


a = Analysis(['windows.py'],
             pathex=['/Users/ricky/OneDrive/projects/p3/dictionary_check'],
             binaries=[],
             datas=[],
             hiddenimports=[],
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
          name='windows',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          runtime_tmpdir=None,
          console=False , icon='Spell_Check_48px_1178702_easyicon.net.ico')
app = BUNDLE(exe,
             name='windows.app',
             icon='Spell_Check_48px_1178702_easyicon.net.ico',
             bundle_identifier=None)
