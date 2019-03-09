# -*- coding: utf-8 -*-
from distutils.command import build
from setuptools import setup

class CustomBuild(build.build, object):
    def run(self):
        from create_style_file import create_style_file
        create_style_file()
        return super(CustomBuild, self).run()

setup(
    name='sphinx-docxbuilder',
    version='0.0.1',
    description='Sphinx builder extension that generates docx files',
    license='MIT',
    keywords='sphinx',
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
    package_data={
        'docxbuilder.docx': ['style.docx'],
    },
    classifiers=[
        'Framework :: Sphinx :: Extension',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.5',
        'Topic :: Documentation :: Sphinx',
    ],
    cmdclass={
        'build': CustomBuild,
    }
)
