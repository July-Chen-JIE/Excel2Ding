from distutils.core import setup
import py2exe
import sys
import os

# If run without args, build for windows
if len(sys.argv) == 1:
    sys.argv.append('py2exe')

# Additional data files to include
data_files = []

# Add the icon file if it exists
icon_path = '../Excel2Ding.ico'
if os.path.exists(icon_path):
    data_files.append(('.', [icon_path]))

# Add JSON files
if os.path.exists('../column_mapping.json'):
    data_files.append(('.', ['../column_mapping.json']))

# Setup py2exe options
setup(
    name='Excel2Ding',
    version='1.5',
    description='Excel数据处理工具',
    windows=[{
        'script': '../Excel2Ding_1.5.py',
        'icon_resources': [(1, icon_path)] if os.path.exists(icon_path) else [],
        'dest_base': 'Excel2Ding'
    }],
    data_files=data_files,
    options={
        'py2exe': {
            'bundle_files': 3,  # Don't bundle DLLs
            'compressed': True,
            'optimize': 2,
            'includes': [
                'tkinter',
                'pandas',
                'openpyxl',
                'tkcalendar',
                'json',
                'os',
                'sys',
                'datetime',
                'warnings',
                're',
                'traceback'
            ],
            'excludes': [
                'PyQt6',
                'PyQt6.QtCore',
                'PyQt6.QtGui',
                'PyQt6.QtWidgets',
                'PyQt6.QtWebEngineWidgets',
                'PyQt6.QtNetwork',
                'PyQt6.QtWebChannel',
                'PyQt6.QtPrintSupport',
                'scipy',
                'matplotlib',
                'requests',
                'bs4',
                'urllib3',
                'PIL'
            ],
            'dll_excludes': ['w9xpopen.exe']
        }
    }
)