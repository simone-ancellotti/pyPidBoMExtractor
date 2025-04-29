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

# setup.py
from setuptools import setup
from Cython.Build import cythonize
import os

source_files = []
for root, dirs, files in os.walk("pyPidBoMExtractor"):
    for file in files:
        if file.endswith(".py") and not file.startswith("__"):
            source_files.append(os.path.join(root, file))

source_files.append("bom_extractor_ui.py")  # Add your main app if desired

setup(
    ext_modules=cythonize(source_files, compiler_directives={'language_level': "3"}),
)