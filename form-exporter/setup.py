from setuptools import setup

APP = ['src/main.py']
DATA_FILES = ['src/assets/logo_empresa.png']
OPTIONS = {
    'argv_emulation': True,
    'packages': ['tkinter', 'fpdf', 'openpyxl', 'PIL'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)