#!/usr/bin/python
# -*- coding: utf-8 -*-
from __future__ import print_function
import os
import zipfile

def create_style_file():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    style_file_path = os.path.normpath(
            os.path.join(base_dir, 'docxbuilder/docx/style.docx'))
    print('creating %s' % style_file_path)
    style_file = zipfile.ZipFile(
            style_file_path, mode='w', compression=zipfile.ZIP_DEFLATED)
    def addfile(dirpath, rootpath):
        for filename in os.listdir(dirpath):
            path = os.path.join(dirpath, filename)
            if os.path.isdir(path):
                addfile(path, rootpath + filename + '/')
            else:
                style_file.write(path, rootpath + filename)
    addfile(os.path.join(base_dir, 'style_file/docx'), '')
    style_file.close()

if __name__ == '__main__':
    create_style_file()

