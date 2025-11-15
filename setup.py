"""Setup configuration for visiowings."""

from setuptools import setup, find_packages
import os

# Read README
with open('README.md', 'r', encoding='utf-8') as f:
    long_description = f.read()

# Read requirements
with open('requirements.txt', 'r') as f:
    requirements = [line.strip() for line in f if line.strip() and not line.startswith('#')]

# Debug requirements
debug_requirements = [
    'pywin32>=305',
    'asyncio>=3.4.3',
    'colorama>=0.4.6',
]

# Test requirements
test_requirements = [
    'pytest>=7.0.0',
    'pytest-asyncio>=0.21.0',
    'pytest-cov>=4.0.0',
]

setup(
    name='visiowings',
    version='0.3.0',
    author='visiowings Development Team',
    description='Git-friendly VBA code management for Microsoft Visio with remote debugging support',
    long_description=long_description,
    long_description_content_type='text/markdown',
    url='https://github.com/twobeass/visiowings',
    packages=find_packages(),
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Topic :: Software Development :: Version Control',
        'Topic :: Software Development :: Debuggers',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
        'Programming Language :: Python :: 3.10',
        'Programming Language :: Python :: 3.11',
        'Operating System :: Microsoft :: Windows',
    ],
    python_requires='>=3.8',
    install_requires=requirements,
    extras_require={
        'debug': debug_requirements,
        'test': test_requirements,
        'dev': debug_requirements + test_requirements,
    },
    entry_points={
        'console_scripts': [
            'visiowings=visiowings.cli:main',
            'visiowings-debug=visiowings.debug.cli:main',
        ],
    },
    include_package_data=True,
    zip_safe=False,
)
