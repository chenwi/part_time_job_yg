from cx_Freeze import setup, Executable  # py3.6
import sys
import os

PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')
# os.environ['TCL_LIBRARY'] = r'D:\Anaconda3\tcl\tcl8.6'
# os.environ['TK_LIBRARY'] = r'D:\Anaconda3\tcl\tk8.6'
base = 'WIN32GUI' if sys.platform == "win32" else None
# base=None
executables = [Executable("extractor.py", base=base, )]

packages = ['numpy', 'pandas', 'docx', 'tkinter']
include_files = [
    os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
    os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
]
options = {
    'build_exe': {
        'packages': packages,
        "includes": ["tkinter"],
        'include_files': include_files
    },

}

setup(
    name="LC96 报告生成工具",
    options=options,
    version="1.0",
    description='LC96 报告生成工具',
    executables=executables
)
