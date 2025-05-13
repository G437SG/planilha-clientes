"""
Script para criar o executável macOS do Formulário de Projeto Arquitetônico
"""

import sys
import os
from setuptools import setup

APP_NAME = "Formulário de Projeto Arquitetônico"
APP = ['main.py']
DATA_FILES = [
    ('', ['logo_empresa.png']),  # inclui a logo na raiz
]

OPTIONS = {
    'argv_emulation': True,
    'packages': ['tkinter', 'fpdf', 'openpyxl', 'PIL'],
    'iconfile': 'logo_empresa.png',  # substitua por um arquivo .icns se tiver
    'plist': {
        'CFBundleName': APP_NAME,
        'CFBundleDisplayName': APP_NAME,
        'CFBundleGetInfoString': "Formulário para coleta de dados de projetos arquitetônicos",
        'CFBundleIdentifier': "com.seunome.planilhaclientes",
        'CFBundleVersion': "1.0.0",
        'CFBundleShortVersionString': "1.0.0",
        'NSHumanReadableCopyright': u"Copyright © 2025, Sua Empresa, Todos os direitos reservados."
    }
}

setup(
    app=APP,
    name=APP_NAME,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)