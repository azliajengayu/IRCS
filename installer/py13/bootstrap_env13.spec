from PyInstaller.utils.hooks import Tree

datas = [
    Tree('IRCS/modules13', prefix='modules13')
]

a = Analysis(
    ['IRCS/bootstrap_env.py'],
    pathex=['.'],              
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
