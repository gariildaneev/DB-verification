from setuptools import setup

APP = ['DB_script.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['pandas', 'openpyxl'],
    'iconfile': 'assets/icon.icns',
    'plist': {
        'CFBundleShortVersionString': '0.1.0',
        'LSUIElement': True,
    },
    'excludes': ['PyQt5.QtSql', 'PyQt5', 'PyInstaller.hooks.hook-PyQt5.QtSql']
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app', 'wheel'],
)
