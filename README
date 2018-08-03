===============================================================
  sphinx-docxbuilder

  An extension for a docx file generation with Sphinx-1.1.2

  Copyright(c) 2011, haraisao at gmail dot com
===============================================================

About this
==========
This programs is an extension to generate a docx file with Sphinx-1.1.2.
This extension is developed by hacking both 'sphinxcontrib-docxbuilder' and 'python-docx'.

Special thanks to Mike MacCana and Takayuki SHIMIZUKAWA.

Information
=============
Auther: Isao Hara
Home Page: https://bitbucket.org/haraisao/sphinx-docxbuilder/
Keywords: sphinx,extension,builder,docx,OpenXML 
License: MIT 

Files
=====
This program is consist of following files.

* builder.py
  This file defined a builder class for Sphinx-1.1.2. This file is copyed from sphinx-docxbuilder's one.

* writer.py
  This file defined a writer class for Sphinx-1.1.2. This file is modified sphinx-docxbuilder's.

* highlight.py
  This file defined DocxFormatter and DocxPygmentsBridge to support highlighting in the literal block.

* docx/docx.py
  This file defined two classes to manipulate a docx file.

* docx/style.docx
  This is a default style file.
  If you customize a docx document when you generate sphinx-docxbuilder, please copy this file in a local directory and modify styles with MS Word.

* contrib/quickstart.py
  This is for 'sphinx-quickstart' command to add some definitions for 'sphinx-docxbuilder'.
  Please replace the original one.

* contrib/exportDocx.py contrib/restructDocx.py
  These are sample command to export/restruct docx file.
   
Requirements
=============
* Python2.6 or later
* lxml module
* PIL(Python Imaging Library) module 
* Optionally, you need MS Word 2007 or later to modify a style file. 

Features
==========
Currently it works followings:

* Headings
* Bullet List
* Enumerated Lists
* Definition Lists
* Field Lists (simple text only)
* Literal Blocks
* High Lighing in Literal Blocks
* Line Blocks
* Block Quotes
* Option Lists (simple text only)
* Simple table and csv table
* Images
* Admonitions
* Basic Inline Markups 

Known issue
============
Many sphinx syntaxes and directives aren't tested yet.


Quick Start
=============
Setup
-----
First of all, download or hg clone sphix-docxbuilder archive and extract all files into $(SPHINX_EGG_DIR)/sphinx-docxbuilder .
Usage

Set 'sphinx-docxbuilder' to 'extensions' line of target sphinx source
conf.py ::

  extensions = ['sphinx-docxbuilder']

Execute sphinx-build with below option ::

  $ sphinx-build -b docx [input-dir] [output-dir]

Customize Style file
---------------------
If you want to customize output file, only you have to do is to change named styles in 'sphinx-doc/docx/style.docx' .

Customize document properties
-----------------------------
sphinx-docxbuilder support to customize document properties. Currently, you can set 'title','subject','creator','company','category','descriptions','keyword' to append expressions to 'conf.py'.

For example, to set the creator and the keyword properties, you add follows ::

  docx_creator = 'Isao HARA'
  docx_keywords = ['Sphinx', 'OpenXML']

Futhermore, you can use your original style file to set following expressions to 'conf.py'. ::

  docx_style = 'MyStyle.docx'



