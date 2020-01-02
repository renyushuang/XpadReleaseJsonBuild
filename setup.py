"""
This is a setup.py script generated by py2applet

Usage:
    python setup.py py2app
"""

from setuptools import setup

APP = ['XpadJsonBuild_GUI.py']
DATA_FILES = ['XpadJsonBuild_data_pip.py','XpadJsonBuild_2.py','XpadJsonBuild_1.py']
DATA_FILES = ['XpadJsonBuild_data_pip.py','XpadJsonBuild_2.py','XpadJsonBuild_1.py']
OPTIONS = {
    'argv_emulation': True,
    'includes': ['openpyxl','Tkinter','defusedxml','xmlrpclib'],}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
