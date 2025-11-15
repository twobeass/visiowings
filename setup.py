"""Setup script for visiowings"""

from setuptools import setup, find_packages
from pathlib import Path

# Read README for long description
this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding='utf-8')

setup(
    name='visiowings',
    version='0.4.1',
    description='VBA Editor for Microsoft Visio with VS Code integration',
    long_description=long_description,
    long_description_content_type='text/markdown',
    author='twobeass',
    author_email='',
    url='https://github.com/twobeass/visiowings',
    packages=find_packages(),
    install_requires=[
        'pywin32>=305',
        'watchdog>=3.0.0',
    ],
    entry_points={
        'console_scripts': [
            'visiowings=visiowings.cli:main',
        ],
    },
    classifiers=[
        'Development Status :: 3 - Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Programming Language :: Python :: 3.12',
        'Operating System :: Microsoft :: Windows',
        'Topic :: Software Development :: Libraries :: Python Modules',
    ],
    python_requires='>=3.8',
    keywords='visio vba editor vscode automation',
)
