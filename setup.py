from setuptools import setup, find_packages

setup(
    name='pyPidBoMExtractor',
    version='0.3',
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
