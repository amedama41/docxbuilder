#! /usr/bin/python
# -*- coding: utf-8 -*-

import sys
import docx

if __name__ == '__main__' :
  flag = True
  if len(sys.argv) > 3:
     flag = False
  if len(sys.argv) > 2:
    doc = docx.DocxDocument( sys.argv[1] )
    doc.extract_files(sys.argv[2], flag)
  else:
    print sys.argv[0], " <docx file> <extract dir>"
