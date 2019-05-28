#!/usr/bin/env python
from __future__ import absolute_import
from __future__ import print_function

import io
import re
from os.path import dirname
from os.path import join

from setuptools import find_packages
from setuptools import setup


def read(*names, **kwargs):
    return io.open(
        join(dirname(__file__), *names),
        encoding=kwargs.get('encoding', 'utf8')
    ).read()


setup(
    name='xlsx_handler',
    version='0.0.1',
    # license='BSD 2-Clause License',
    # description='Handle for XLSX/XLSM files',
    # long_description='%s\n' % (
    #     re.compile('^.. start-badges.*^.. end-badges', re.M | re.S).sub('', read('README.md'))
    # ),
    author='Sokolov Nikolay Valerevich',
    author_email='sokolov.nicholas@gmail.com',
    # url='https://github.com/nicholas-sokolov/xlsx_handler',
    # packages=find_packages('src'),
    packages=['src',],
    # zip_safe=False,
    # classifiers=[
    #     'Development Status :: 5 - Production/Stable',
    #     'Intended Audience :: Developers',
    #     'License :: OSI Approved :: BSD License',
    #     'Operating System :: Unix',
    #     'Operating System :: POSIX',
    #     'Operating System :: Microsoft :: Windows',
    #     'Programming Language :: Python :: 3.6',
    #     'Programming Language :: Python :: 3.7',
    # ],
)
