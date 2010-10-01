#!/usr/bin/env python
"""Requires setuptools.

To build the setuptools egg use
    python setup.py bdist_egg
and either upload it to the PyPI with:
    python setup.py upload
or upload to your own server and register the release with PyPI:
    python setup.py register

A source distribution (.zip) can be built with
    python setup.py sdist --format=zip
That uses the manifest.in file for data files rather than searching for them here
"""

from setuptools import setup, Extension, find_packages
import sys
sys.path.append('src')#so we can import legacy_openpyxl
import legacy_openpyxl#to fetch __version__ etc

setup(name = "legacy_openpyxl",
    packages = find_packages('src'),
    include_package_data = True,
    package_dir = {'':'src'},
    package_data = {'': ['legacy_openpyxl/tests/*.xml', 'openpyxl/tests/*.xslx']}, #DOESN'T AFFECT zip DISTRIBUTION. MUST MODIFY MANIFEST.in TOO
    #metadata
    version = legacy_openpyxl.__version__,
    description = "A Python library to read/write Excel 2007 xlsx/xlsm files",
    long_description = "legacy_openpyxl is a pure python reader and writer of Excel OpenXML files." + \
        " It is ported from the PHPExcel project",
    author = legacy_openpyxl.__author__,
    author_email = legacy_openpyxl.__author_email__,
    url = legacy_openpyxl.__url__,
    license = legacy_openpyxl.__license__,
    download_url = legacy_openpyxl.__downloadUrl__,
    test_suite = 'nose.collector',
    classifiers = ['Development Status :: 4 - Beta',
          'Operating System :: MacOS :: MacOS X',
          'Operating System :: Microsoft :: Windows',
          'Operating System :: POSIX',
          'License :: OSI Approved :: MIT License',
          'Programming Language :: Python'],
    )
