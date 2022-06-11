import PyInstaller.__main__

PyInstaller.__main__.run([
    'report_maker.py',
    '--onefile',
    '--distpath', './'
])