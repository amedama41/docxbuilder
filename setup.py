# -*- coding: utf-8 -*-
import os
from distutils.command import build
from setuptools import setup

class CustomBuild(build.build, object):
    def run(self):
        from create_style_file import create_style_file
        create_style_file()
        return super(CustomBuild, self).run()

BASEDIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASEDIR, 'README.rst'), 'r') as f:
    long_description = f.read()

setup(
    name='docxbuilder',
    version='1.1.0',
    description='Sphinx docx builder extension',
    long_description=long_description,
    url='https://github.com/amedama41/docxbuilder',
    author='amedama41',
    author_email='kamo.devel41@gmail.com',
    keywords=['sphinx', 'extension', 'docx', 'OpenXML'],
    packages=[
        'docxbuilder',
        'docxbuilder.docx',
    ],
    install_requires=[
        "Sphinx>=1.7.6",
        "lxml",
        "pillow",
        "six",
    ],
    python_requires='>=2.7,!=3.0.*,!=3.1.*,!=3.2.*,!=3.3.*,!=3.4.*',
    package_data={
        'docxbuilder.docx': ['style.docx'],
    },
    classifiers=[
        'Framework :: Sphinx :: Extension',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Topic :: Documentation :: Sphinx',
    ],
    cmdclass={
        'build': CustomBuild,
    }
)
