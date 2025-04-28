from setuptools import setup, find_packages
from pyPidBoMExtractor._version import __version__

setup(
    name='pyPidBoMExtractor',
    version=__version__,
    packages=find_packages(),
    install_requires=[
        'ezdxf',
        'numpy',
        'openpyxl',
    ],
    entry_points={
        'console_scripts': [
            'bom-extract = pyPidBoMExtractor.cli:main',
        ],
    },
)
