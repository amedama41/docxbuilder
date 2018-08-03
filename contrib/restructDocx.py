#! /usr/bin/python
# -*- coding: utf-8 -*-

import sys
import docx

if __name__ == '__main__' :
  if len(sys.argv) > 2:
    doc = docx.DocxDocument()
    doc.restruct_docx(sys.argv[1], sys.argv[2])
  else:
    print sys.argv[0], " <docx dir> <docx filename>"
