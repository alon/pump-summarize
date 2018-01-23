from setuptools import setup

setup(
    name='Pump-Summarize',
    description='Summary tool for pump xls files from postprocessing',
    version="0.1",
    install_requires=[
        'xlrd(==1.1.0)',
        'XlsxWriter(==1.0.2)',
        'PyInstaller(==3.3)',
        'colorama>=0.3.7',
        'pyqtgraph(==0.10.0)',
        'Qt.py(==1.0.0)',
        'PyQt5(==5.9)',
    ],
    packages=['summarize'],
    entry_points={
        'console_scripts': [
            'summarize = summarize:main [summarize]'
        ]
    }
)
