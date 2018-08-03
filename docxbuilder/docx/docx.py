 # -*- coding: utf-8 -*-
'''
  Microsoft Word 2007 Document Composer

  Copyright 2011 by haraisao at gmail dot com

  This software based on 'python-docx' which developed by Mike MacCana.

'''
'''
  Open and modify Microsoft Word 2007 docx files (called 'OpenXML' and 'Office OpenXML' by Microsoft)

  Part of Python's docx module - http://github.com/mikemaccana/python-docx
  See LICENSE for licensing information.
'''

from lxml import etree
import Image
import zipfile
import shutil
import re
import time
import os
from os.path import join
import tempfile
import sys


# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily. 
nsprefixes = {
    # Text Content
    'mv':'urn:schemas-microsoft-com:mac:vml',
    'mo':'http://schemas.microsoft.com/office/mac/office/2008/main',
    've':'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o':'urn:schemas-microsoft-com:office:office',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm':'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v':'urn:schemas-microsoft-com:vml',
    'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10':'urn:schemas-microsoft-com:office:word',
    'wne':'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp':'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp':"http://schemas.openxmlformats.org/package/2006/metadata/core-properties", 
    'dc':"http://purl.org/dc/elements/1.1/", 
    'dcterms':"http://purl.org/dc/terms/",
    'dcmitype':"http://purl.org/dc/dcmitype/",
    'xsi':"http://www.w3.org/2001/XMLSchema-instance",
    'ep':'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct':'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr':'http://schemas.openxmlformats.org/package/2006/relationships',
    # xml 
    'xml':'http://www.w3.org/XML/1998/namespace'
    }

Enum_Types = {
    'arabic':'decimal',
    'loweralpha':'lowerLetter',
    'upperalpha':'upperLetter',
    'lowerroman':'lowerRoman',
    'upperroman':'upperRoman'
    }

#####################
def norm_name(tagname, namespaces=nsprefixes):
    '''
       Convert the 'tagname' to a formal expression.
          'ns:tag' --> '{namespace}tag'
          'tag' --> 'tag'
    '''
    ns_name = tagname.split(':', 1)
    if len(ns_name) >1 :
      tagname = "{%s}%s" % (namespaces[ns_name[0]], ns_name[1])
    return tagname

def get_elements(xml, path, ns=nsprefixes):
    '''
       Get elements from a Element tree with 'path'.
    '''
    result = []
    try:
      result = xml.xpath(path, namespaces=ns)
    except:
      pass
    return result

def append_element(elem, xml, path=None, index=0, ns=nsprefixes):
    '''
       Append an Element
    '''
    try:
      dist = xml
      if path :
        dist = xml.xpath(path, namespaces=ns)
      dist[index].append(elem)
      return True
    except:
      print "Error  in append_element"

    return False

def find_file(filename, child_dir=None):
    '''
       Find file...
    '''
    fname = filename
    if not os.access( filename ,os.F_OK):
      for pth in sys.path:
        if child_dir :
          pth = join(pth, child_dir)
        fname = join(pth, filename)
        if os.access(fname, os.F_OK):
          break
        else:
          fname=None
    return fname

def get_enumerate_type(typ):
  '''
       
  '''
  try:
    typ=Enum_Types[typ]
  except:
    typ="decimal"
    pass
  return  typ

def parse_tag_list(tag):
  '''
       
  '''
  tagname = ''
  tagtext = ''
  attributes = {}

  if isinstance(tag,str) :
    tagname=tag
  elif isinstance(tag,list) :
    tagname=tag[0]
    taglen = len(tag)
    if taglen > 1 :
      if isinstance(tag[1],basestring) :
        tagtext = tag[1]
      else:
        attributes = tag[1]
    if taglen > 2:
      if isinstance(tag[2],basestring) :
        tagtext = tag[2]
      else:
        attributes = tag[2]
  else:
    print "Invalid tag:",tag

  return tagname,attributes,tagtext

def extract_nsmap(tag, attributes):
    '''
    '''
    result = {}
    ns_name = tag.split(':', 1)
    if len(ns_name) > 1 and nsprefixes.get(ns_name[0]) :
        result[ns_name[0]] = nsprefixes[ns_name[0]]

    for x in attributes:
      ns_name = x.split(':', 1)
      if len(ns_name) > 1 and nsprefixes.get(ns_name[0]) :
          result[ns_name[0]] = nsprefixes[ns_name[0]]

    return result

def make_element_tree(arg, _xmlns=None):
    '''
       
    '''
    tagname,attributes,tagtext = parse_tag_list(arg[0])
    children = arg[1:]

    nsmap = extract_nsmap(tagname, attributes)

    if _xmlns is None :
      newele = etree.Element(norm_name(tagname), nsmap=nsmap)
    else :
      newele = etree.Element(norm_name(tagname), xmlns=_xmlns, nsmap=nsmap)

    if tagtext :
      newele.text = tagtext

    for attr in attributes:
      newele.set(norm_name(attr), attributes[attr])

    for child in children:
      chld = make_element_tree(child)
      if chld is not None :
        newele.append(chld)

    return newele

def get_child_element(xml, p):
    '''
       
    '''
    elems = get_elements(xml, p)
    if elems == [] :
      ele = make_element_tree([p])
      xml.append(ele)
      return ele
    return elems[0]

def set_attributes(xml, path, attributes):
    '''
       
    '''
    elems = get_elements(xml, path)
    if elems == [] :
      pathes = path.split('/')
      elem=xml
      for p in pathes:
        elem = get_child_element(elem, p)
    else:
      elem = elems[0]

    for attr in attributes:
      elem.set(norm_name(attr), attributes[attr])
    return elem

def get_attribute(xml, path, name):
    '''
       
    '''
    elems = get_elements(xml, path)
    if elems == [] :
      return None
    return elems[0].attrib[norm_name(name)]

#
#  DocxDocument class
#   This class for analizing docx-file
#
class DocxDocument:
  def __init__(self, docxfile=None):
    '''
      Constructor
    '''
    self.title = ""
    self.subject = ""
    self.creator = "Python:DocDocument"
    self.company = ""
    self.category = ""
    self.descriptions = ""
    self.keywords = []
    self.stylenames = {}

    if docxfile :
      self.set_document(docxfile)
      self.docxfile = docxfile

  def set_document(self, fname):
    '''
      set docx document 
    '''
    if fname :
      self.docxfile = fname
      self.docx = zipfile.ZipFile(fname)

      self.document = self.get_xmltree('word/document.xml')
      self.docbody = get_elements(self.document, '/w:document/w:body')[0]

      self.numbering = self.get_xmltree('word/numbering.xml')
      self.styles = self.get_xmltree('word/styles.xml')
      self.extract_stylenames()
      self.paragraph_style_id = self.stylenames['Normal']
      self.character_style_id = self.stylenames['Default Paragraph Font']


    return self.document

  def get_xmltree(self, fname):
    '''
      Extract a document tree from the docx file
    '''
    try:
      return etree.fromstring(self.docx.read(fname))
    except:
      return None
    
  def extract_stylenames(self):
    '''
      Extract a stylenames from the docx file
    '''
    style_elems = get_elements(self.styles, 'w:style')

    for style_elem in style_elems:
        aliases_elems = get_elements(style_elem, 'w:aliases')
        if aliases_elems:
            name = aliases_elems[0].attrib[norm_name('w:val')]
        else:
            name_elem = get_elements(style_elem,'w:name')[0]
            name = name_elem.attrib[norm_name('w:val')]
        value = style_elem.attrib[norm_name('w:styleId')]
        self.stylenames[name] = value
    return self.stylenames

  def get_paper_info(self):
    self.paper_info = get_elements(self.document,'/w:document/w:body/w:sectPr')[0]
    self.paper_size = get_elements(self.document,'/w:document/w:body/w:sectPr/w:pgSz')[0]
    self.paper_margin = get_elements(self.document,'/w:document/w:body/w:sectPr/w:pgMar')[0]
    width = int(self.paper_size.get(norm_name('w:w'))) - int(self.paper_margin.get(norm_name('w:right'))) -int(self.paper_margin.get(norm_name('w:left'))) 
    height = int(self.paper_size.get(norm_name('w:h'))) - int(self.paper_margin.get(norm_name('w:top'))) -int(self.paper_margin.get(norm_name('w:bottom'))) 

    # paper info: unit ---> 2099 mm = 11900 paper_unit
    self.document_width = int(width * 2099 / 11900)  # mm
    self.document_height = int(height * 2970 / 16840)  # mm

    print self.document_width, self.document_height
    return self.paper_info
    
  def get_coverpage(self):
    coverInfo=get_attribute(self.docbody, 'w:sdt/w:sdtPr/w:docPartObj/w:docPartGallery', 'w:val')
    if coverInfo == "Cover Pages":
      self.coverpage=get_elements(self.docbody,'w:sdt')[0]
    else:
      self.coverpage=None

    return self.coverpage

  def extract_file(self,fname, outname=None, pprint=True):
    '''
      Extract file from docx 
    '''
    try:
      filelist = self.docx.namelist()

      if filelist.index(fname) >= 0 :
        xmlcontent = self.docx.read(fname)
        document = etree.fromstring(xmlcontent)
        xmlcontent = etree.tostring(document, encoding='UTF-8', pretty_print=pprint)
        if outname == None : outname = os.path.basename(fname)

        f = open(outname, 'w')
        f.write(xmlcontent)
        f.close()
    except:
        print "Error in extract_document: %s" % fname
        #print filelist
    return


  def extract_files(self,to_dir, pprint=False):
    '''
      Extract all files from docx 
    '''
    try:
      if not os.access(to_dir, os.F_OK) :
        os.mkdir(to_dir)

      filelist = self.docx.namelist()
      for fname in filelist :
        xmlcontent = self.docx.read(fname)
	fname_ext = os.path.splitext(fname)[1]
	if pprint and (fname_ext == '.xml'  or fname_ext == '.rels') :
          document = etree.fromstring(xmlcontent)
          xmlcontent = etree.tostring(document, encoding='UTF-8', pretty_print=True)
        file_name = join(to_dir,fname)
        if not os.path.exists(os.path.dirname(file_name)) :
          os.makedirs(os.path.dirname(file_name)) 
        f = open(file_name, 'w')
        f.write(xmlcontent)
        f.close()
    except:
      print "Error in extract_files ..."
      return False
    return True

  def restruct_docx(self, docx_dir, docx_filename, files_to_skip=[]):
    '''
       This function is copied and modified the 'savedocx' function contained 'python-docx' library
      Restruct docx file from files in 'doxc_dir'
    '''
    if not os.access( docx_dir ,os.F_OK):
      print "Can't found docx directory: %s" % docx_dir
      return

    docxfile = zipfile.ZipFile(docx_filename, mode='w', compression=zipfile.ZIP_DEFLATED)

    prev_dir = os.path.abspath('.')
    os.chdir(docx_dir)

    # Add & compress support files
    files_to_ignore = ['.DS_Store'] # nuisance from some os's
    for dirpath,dirnames,filenames in os.walk('.'):
        for filename in filenames:
            if filename in files_to_ignore:
                continue
            templatefile = join(dirpath,filename)
            archivename = os.path.normpath(templatefile)
            archivename = '/'.join(archivename.split(os.sep))
            if archivename in files_to_skip:
                continue
            #print 'Saving: '+archivename          
            docxfile.write(templatefile, archivename)

    os.chdir(prev_dir) # restore previous working dir
    return docxfile

  def get_filelist(self):
      '''
         Extract file names from docx file
      '''
      filelist = self.docx.namelist()
      return filelist

  def search(self, search):
    '''
      This function is copied from 'python-docx' library
      Search a document for a regex, return success / fail result
    '''
    result = False
    text_tag = norm_name('w:t')
    searchre = re.compile(search)
    for element in self.docbody.iter():
        if element.tag == text_tag :
            if element.text:
                if searchre.search(element.text):
                    result = True
    return result

  def replace(self, search,replace):
    '''
      This function copied from 'python-docx' library
      Replace all occurences of string with a different string, return updated document
    '''
    text_tag = norm_name('w:t')
    newdocument = self.docbody
    searchre = re.compile(search)
    for element in newdocument.iter():
        if element.tag == text_tag :
            if element.text:
                if searchre.search(element.text):
                    element.text = re.sub(search,replace,element.text)
    return newdocument

############
##  Numbering
  def get_numbering_style_id(self, style):
    '''
       
    '''
    try:
      style_elems = get_elements(self.styles, '/w:styles/w:style')
      for style_elem in style_elems:
        name_elem = get_elements(style_elem,'w:name')[0]
        name = name_elem.attrib[norm_name('w:val')]
	if name == style :
            numPr = get_elements(style_elem,'w:pPr/w:numPr/w:numId')[0]
            value = numPr.attrib[norm_name('w:val')]
            return value
    except: 
      pass
    return '0'

  def get_numbering_ids(self):
    '''
       
    '''
    num_elems = get_elements(self.numbering, '/w:numbering/w:num')
    result = []
    for num_elem in num_elems :
        nid = num_elem.attrib[norm_name('w:numId')]
        result.append( nid )
    return result

  def get_numbering_ids2(self):
    '''
       
    '''
    num_elems = get_elements(self.numbering, '/w:numbering/w:num')
    result = []
    for num_elem in num_elems :
        nid = num_elem.attrib[norm_name('w:numId')]
	elem = get_elements(num_elem, 'w:abstractNumId')[0]
	abstId = elem.attrib[norm_name('w:val')]
        result.append( [nid, abstId] )
    return result

  def get_abstractNum_ids(self):
    '''
       
    '''
    num_elems = get_elements(self.numbering, '/w:numbering/w:abstractNum')
    result = []
    for num_elem in num_elems :
        nid = num_elem.attrib[norm_name('w:abstractNumId')]
        result.append( nid )
    return result

  def get_max_numbering_id(self):
    '''
       
    '''
    max_id = 0
    num_ids = self.get_numbering_ids()
    for x in num_ids :
      if int(x) > max_id :  max_id = int(x)
    return max_id

  def get_numbering_left(self, style):
    '''
       get numbering indeces
    '''
    abstractNums=get_elements(self.numbering, 'w:abstractNum')

    indres=[0]

    for x in abstractNums :
      styles=get_elements(x, 'w:lvl/w:pStyle')
      if styles :
        pstyle_name = styles[0].get(norm_name('w:val') )
        if pstyle_name == style :
          ind=get_elements(x, 'w:lvl/w:pPr/w:ind')
	  if ind :
            indres=[]
	    for indx in ind :
              indres.append(int(indx.get(norm_name('w:left'))))
          return indres
    return indres


##########

  def getdocumenttext(self):
    '''
      This function copied from 'python-docx' library
      Return the raw text of a document, as a list of paragraphs.
    '''
    paragraph_tag == norm_nama('w:p')
    text_tag == norm_nama('w:text')
    paratextlist=[]   
    # Compile a list of all paragraph (p) elements
    paralist = []
    for element in self.document.iter():
        # Find p (paragraph) elements
        if element.tag == paragraph_tag:
            paralist.append(element)    
    # Since a single sentence might be spread over multiple text elements, iterate through each 
    # paragraph, appending all text (t) children to that paragraphs text.     
    for para in paralist:      
        paratext=u''  
        # Loop through each paragraph
        for element in para.iter():
            # Find t (text) elements
            if element.tag == text_tag:
                if element.text:
                    paratext = paratext+element.text
        # Add our completed paragraph text to the list of paragraph text    
        if not len(paratext) == 0:
            paratextlist.append(paratext)                    
    return paratextlist        

#
# DocxComposer Class
#
class DocxComposer:
  def __init__(self, stylefile=None):
    '''
       Constructor
    '''
    self._coreprops=None
    self._appprops=None
    self._contenttypes=None
    self._websettings=None
    self._wordrelationships=None
    self.breakbefore = False
    self.last_paragraph = None
    self.stylenames = {}
    self.title = ""
    self.subject = ""
    self.creator = "Python:DocDocument"
    self.company = ""
    self.category = ""
    self.descriptions = ""
    self.keywords = []
    self.max_table_width = 8000
    self.sizeof_field_list = [2000,5500]

    self.abstractNums = []
    self.numids = []

    self.images = 0
    self.nocoverpage = False

    if stylefile == None :
      self.template_dir = None
    else :
      self.new_document(stylefile)

  def set_style_file(self, stylefile):
    '''
       Set style file 
    '''
    fname = find_file(stylefile, 'sphinx-docxbuilder/docx')

    if fname == None:
      print "Error: style file( %s ) not found" % stylefile
      return None
      
    self.styleDocx = DocxDocument(fname)

    self.template_dir = tempfile.mkdtemp(prefix='docx-')
    result = self.styleDocx.extract_files(self.template_dir)

    if not result :
      print "Unexpected error in copy_docx_to_tempfile"
      shutil.rmtree(temp_dir, True)
      self.template_dir = None
      return 

    self.stylenames = self.styleDocx.extract_stylenames()
    self.paper_info = self.styleDocx.get_paper_info()
    self.bullet_list_indents = self.get_numbering_left('ListBullet')
    self.bullet_list_numId = self.styleDocx.get_numbering_style_id('ListBullet')
    self.number_list_indent = self.get_numbering_left('ListNumber')[0]
    self.number_list_numId = self.styleDocx.get_numbering_style_id('ListNumber')
    self.abstractNums = get_elements(self.styleDocx.numbering, 'w:abstractNum')
    self.numids = get_elements(self.styleDocx.numbering, 'w:num')
    self.numbering = make_element_tree(['w:numbering'])

    return

  def set_coverpage(self,flag=True):
    self.nocoverpage = not flag

  def get_numbering_ids(self):
    '''
       
    '''
    result = []
    for num_elem in self.numids :
        nid = num_elem.attrib[norm_name('w:numId')]
        result.append( nid )
    return result

  def get_max_numbering_id(self):
    '''
       
    '''
    max_id = 0
    num_ids = self.get_numbering_ids()
    for x in num_ids :
      if int(x) > max_id :  max_id = int(x)
    return max_id

  def delete_template(self):
    '''
       Delete the temporary directory which we use compose a new document. 
    '''
    shutil.rmtree(self.template_dir, True)

  def new_document(self, stylefile):
    '''
       Preparing a new document
    '''
    self.set_style_file(stylefile)
    self.document = make_element_tree([['w:document'],[['w:body']]])
    self.docbody = get_elements(self.document, '/w:document/w:body')[0]
    self.current_docbody = self.docbody

    self.relationships = self.relationshiplist()

    return self.document

  def set_props(self, title, subject, creator, company='', category='', descriptions='', keywords=[]):
    '''
      Set document's properties: title, subject, creatro, company, category, descriptions, keywrods.
    '''
    self.title = title
    self.subject = subject
    self.creator = creator
    self.company = company
    self.category = category
    self.descriptions = descriptions
    self.keywords = keywords

  def save(self, docxfilename):
    '''
      Save the composed document to the docx file 'docxfilename'.
    '''
    assert os.path.isdir(self.template_dir)

    self.coreproperties()
    self.appproperties()
    self.contenttypes()
    self.websettings()

    self.wordrelationships()

    for x in self.abstractNums :
      self.numbering.append(x)
    for x in self.numids :
      self.numbering.append(x)

    coverpage = self.styleDocx.get_coverpage()

    if not self.nocoverpage and coverpage is not None :
      print "output Coverpage"
      self.docbody.insert(0,coverpage)

    self.docbody.append(self.paper_info)


    # Serialize our trees into out zip file
    treesandfiles = {self.document:'word/document.xml',
                     self._coreprops:'docProps/core.xml',
                     self._appprops:'docProps/app.xml',
                     self._contenttypes:'[Content_Types].xml',
                     self.numbering:'word/numbering.xml',
                     self.styleDocx.styles:'word/styles.xml',
                     self._websettings:'word/webSettings.xml',
                     self._wordrelationships:'word/_rels/document.xml.rels'}

    docxfile = self.styleDocx.restruct_docx(self.template_dir, docxfilename, treesandfiles.values())

    for tree in treesandfiles:
        if tree != None:
            #print 'Saving: '+treesandfiles[tree]    
            treestring =  etree.tostring(tree, xml_declaration=True, encoding='UTF-8', standalone='yes')
            docxfile.writestr(treesandfiles[tree],treestring)
    
    print 'Saved new file to: '+docxfilename
    shutil.rmtree(self.template_dir)
    return
    
 ##################
  def set_docbody(self, body=None):
    '''
      Set docx body..
    '''
    if body is None:
      self.current_docbody = self.docbody
    else:
      self.current_docbody = body
    return self.current_docbody

  def append(self, para):
    '''
      Append paragraph to document
    '''
    self.current_docbody.append(para)
    self.last_paragraph = para
    return para

  def table_of_contents(self, toc_text='Contents:', maxlevel=3):
    '''
      Insert the Table of Content
    '''
    toc_tree = [['w:sdt'],
                  [['w:sdtPr'],
                       [['w:rPr'], [['w:long']] ],
                       [['w:docPartObj'], [['w:docPartGallery', {'w:val':'Table of Contents'}]], [['w:docPartUnique']] ]
                  ]
	       ]

    sdtContent_tree = [['w:sdtContent']]

    if toc_text :
      p_tree = [['w:p'], [['w:pPr'], [['w:pStyle', {'w:val':'TOC_Title'}]] ], [['w:r'], [['w:rPr'], [['w:long']] ], [['w:t',toc_text]] ] ]
      sdtContent_tree.append(p_tree)

    p_tree = [['w:p'],
		    [['w:pPr'],
			    [['w:pStyle', {'w:val':'TOC_Contents'}]],
			    [['w:tabs'], 
				    [['w:tab',{'w:val':'right', 'w:leader':'dot','w:pos':'8488'}] ]
		           ],
			    [['w:rPr'], [['w:b',{'w:val':'0'}]], [['w:noProof']] ]
	            ],
                    [['w:r'],[['w:fldChar', {'w:fldCharType':'begin'}]]],
                    [['w:r'],[['w:instrText', ' TOC \o "1-%d" \h \z \u ' % maxlevel , {'xml:space':'preserve'}]]],
                    [['w:r'],[['w:fldChar', {'w:fldCharType':'separare'}]]],
                    [['w:r'],[['w:fldChar', {'w:fldCharType':'end'}]]]
	    ]
    sdtContent_tree.append(p_tree)

    p_tree = [['w:p'], [ ['w:r'], [['w:fldChar',{'w:fldCharType':'end'}]] ] ]
    sdtContent_tree.append(p_tree)

    toc_tree.append(sdtContent_tree)
    sdt = make_element_tree(toc_tree)

    self.append(sdt)

#################
####       Output PageBreak
  def pagebreak(self,type='page', orient='portrait'):
    '''
      Insert a break, default 'page'.
      See http://openxmldeveloper.org/forums/thread/4075.aspx
      Return our page break element.

      This method is copied from 'python-docx' library
    '''
    # Need to enumerate different types of page breaks.
    validtypes = ['page', 'section']

    pagebreak_tree = [['w:p']]

    if type not in validtypes:
        raise ValueError('Page break style "%s" not implemented. Valid styles: %s.' % (type, validtypes))


    if type == 'page':
        run_tree = [['w:r'],[['w:br', {'w:type':'page'}]]]
    elif type == 'section':
        if orient == 'portrait':
            attrs = {'w:w':'12240','w:h':'15840'}
        elif orient == 'landscape':
            attrs={'w:h':'12240','w:w':'15840', 'w:orient':'landscape'}
        run_tree = [['w:pPr'],[['w:sectPr'], [['w:pgSz', attrs]] ] ]

    pagebreak_tree.append(run_tree)

    pagebreak = make_element_tree(pagebreak_tree)

    self.append(pagebreak)

    self.breakbrefore = True
    return pagebreak    

#################
####       Output Paragraph
  def make_paragraph(self, style='BodyText', block_level=0):
    '''
      Make a new paragraph element
    '''
    # if 'style' isn't defined, cretae new style.
    if style not in self.stylenames :
      self.new_paragraph_style(style)

    # calcurate indent
    ind = 0
    if block_level > 0 :
        ind = self.number_list_indent * block_level

    # set paragraph tree
    paragraph_tree = [['w:p'], 
		    	[['w:pPr'], 
	    			[['w:pStyle',{'w:val':style}]],
    				[['w:ind',{'w:leftChars':'0','w:left': str(ind)} ]]
	                ]
		     ]

    if self.breakbefore :
        paragraph_tree.append( [['w:r'], [['w:lastRenderedPageBreak']]] )

    # create paragraph
    paragraph = make_element_tree(paragraph_tree)
    return paragraph

#################
####       Output Paragraph
  def paragraph(self, paratext=None, style='BodyText', block_level=0, create_only=False):
    '''
      Make a new paragraph element, containing a run, and some text. 
      Return the paragraph element.
    '''
    isliteralblock=False
    if style == 'LiteralBlock' :
      paratext = paratext[0].splitlines()
      isliteralblock=True

    paragraph = self.make_paragraph(style, block_level)

    #  Insert a text run
    if paratext != None:
        self.make_runs(paragraph, paratext, isliteralblock)

    #  if the 'create_only' flag is True, append paragraph to the document
    if not create_only :
        self.append(paragraph)
        self.last_paragraph = paragraph

    return paragraph

  def insert_linespace(self):
    self.append(self.make_paragraph())

  def get_paragraph_text(self, paragraph=None):
    if paragraph is None: paragaph = self.last_paragraph
    txt_elem = get_elements(paragraph, 'w:r/w:t')
    result = ''
    for txt in txt_elem :
       result += txt.text
    return result

  def get_last_paragraph_style(self):
    result = get_attribute(self.last_paragraph,'w:pPr/w:pStyle', 'w:val')
    if result is None :
      result = 'BodyText'
    return result

  def insert_paragraph_property(self, paragraph, style='BodyText'):
    '''
       Insert paragraph property element with style.
    '''
    if style not in self.stylenames :
      self.new_paragraph_style(style)
    style = self.stylenames.get(style, 'BodyText')

    pPr = make_element_tree( [ ['w:pPr'], [['w:pStyle',{'w:val':style}]] ] )
    paragraph.append(pPr) 
    return paragraph

  def get_last_paragraph(self):
    paras = get_elements(self.current_docbody, 'w:p')
    if len(paras) > 1:
      return paras[-1]
    return None

  def trim_paragraph(self):
    paras = get_elements(self.current_docbody, 'w:p')
    if len(para) > 2:
      self.last_paragraph = paras[-2]
      self.current_docbody.remove(paras[-1])
    elif len(para) > 1:
      self.last_paragraph = None
      self.current_docbody.remove(paras[-1])
    return

  def get_paragraph_style(self, paragraph, force_create=False):
    '''
       Get stylename of the paragraph
    '''
    result = get_attribute(paragraph, 'w:pPr/w:pStyle', 'w:val')
    if result is None :
      if force_create :
        self.insert_paragraph_property(paragraph)
      result = 'BodyText'

    return result

  def set_indent(self, paragraph, lskip):
    '''
       Set indent of paragraph
    '''
    ind = set_attributes(paragraph, 'w:pPr/w:ind', {'w:leftChars':'0','w:left': str(lskip)} ) 

    return ind

  def make_runs(self, paragraph, targettext, literal_block=False):
    '''
      Make new runs with text.
    '''
    run = []
    if isinstance(targettext, (list)) :
        for i,x in enumerate(targettext) :
            if isinstance(x, (list)) :
                run.append(self.make_run(x[0], style=x[1]))
            else:
	        if literal_block :
                  run_list = self.make_run(x,rawXml=True)
                  run.extend(run_list)
                else:
                  run.append(self.make_run(x))
	    if literal_block and i+1 <  len(targettext) :
                run.append( self.make_run(':br') )
    else:
	if literal_block :
          run.extend(self.make_run(targettext,rawXml=True))
        else:
          run.append(self.make_run(targettext))

    for r in run:
        paragraph.append(r)    

    return paragraph

  def make_run(self, txt, style='Normal', rawXml=None):
    '''
      Make a new styled run from text.
    '''
    run_tree = [['w:r']]
    if txt == ":br" :
      run_tree.append([['w:br']])
    else:
      attr ={}
      if txt.find(' ') != -1 :
        attr ={'xml:space':'preserve'}

      if style != 'Normal' :
        if style not in self.stylenames :
          self.new_character_style(style)

	run_tree.append([['w:rPr'], [['w:rStyle',{'w:val':style}], [['w:t', txt, attr]] ]])
      else:
        run_tree.append([['w:t', txt, attr]])

    # Make run element
    if rawXml:
      xmltxt='<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'+txt+'</w:p>'
      p = etree.fromstring(xmltxt)
      run = get_elements(p, 'w:r')
      ## remove the last run, because it could be '<w:br>'
      run.pop()
    else:
      run = make_element_tree(run_tree)
                
    return run

  def add_br(self):
    '''
      append line break in current paragraph
    '''
    run = make_element_tree( [['w:r'],[['w:br']]] )

    if self.last_paragraph == None:
        self.paragraph(None)

    self.last_paragraph.append(run)    

    return run

  def add_space(self, style='Normal'):
    '''
      append a space in current paragraph
    '''
    if style != 'Normal' :
      style = self.stylenames.get(style, 'Normal')

    run_tree = [['w:r'],
                    [['w:rPr'], [['w:rStyle', {'w:val':style}]]],
		    [['w:t', ' ', {'xml:space':'preserve'}]]]

    # Make rum element
    run = make_element_tree(run_tree)
    if self.last_paragraph == None:
        self.paragraph(None)

    # append the run to last paragraph
    self.last_paragraph.append(run)    
    return run

########
##     Output Headinng
  def heading(self, headingtext, headinglevel):
    '''
      Make a heading
    '''
    # Make paragraph element
    paragraph = make_element_tree(['w:p'])
    self.insert_paragraph_property(paragraph, 'Heading'+str(headinglevel))

    self.make_runs(paragraph, headingtext)

    self.last_paragraph = paragraph
    self.append(paragraph)

    return paragraph   

########
##    Output ListItem
  def list_item(self, itemtext, style='ListBullet', lvl=1, nid=0, enum_prefix=None, enum_prefix_type=None, start=1):
    '''
      Make a new list paragraph
    '''
    # Make paragraph element
    paragraph = make_element_tree(['w:p'])

    self.insert_paragraph_property( paragraph, style)
    self.insert_numbering_property(paragraph, lvl-1, nid, start, enum_prefix, enum_prefix_type)
    self.make_runs(paragraph, itemtext)

    self.last_paragraph = paragraph
    self.append(paragraph)

    return paragraph   

########
##      Numbering Style

  def get_numbering_left(self, style):
    '''
       Get numbering indeces...
    '''
    return self.styleDocx.get_numbering_left(style)

  def get_numbering_indent(self, style='ListBullet', lvl=0, nId=0):
    '''
       Get indenent value
    '''
    result = 0

    if style == 'ListBullet' or nId == 0 :
      if len(self.bullet_list_indents) > lvl :
        result = self.bullet_list_indents[lvl]
      else:
        result = self.bullet_list_indents[-1]
    else:
      result = self.number_list_indent * (lvl+1)

    return result
     
  def find_numbering_paragraph(self, nId):
    '''
       
    '''
    result =[]
    for p in self.docbody :
      elem = get_elements(p, 'w:pPr/w:numPr/w:numId')
      for x in elem:
        if int(x.attrib[norm_name('w:val')]) == int(nId) :
          result.append(p)
    return result

  def set_numbering_id(self, paragraph, nId):
    '''
       
    '''
    elem = get_elements(paragraph, 'w:pPr/w:numPr/w:numId')
    if elem :
        elem[0].set(norm_name('w:val'), str(nId))

  def replace_numbering_id(self, oldId, newId):
    '''
       
    '''
    oldp = self.find_numbering_paragraph(oldId)
    for p  in oldp :
        self.set_numbering_id(p, newId)
     
  def insert_numbering_property(self, paragraph, lvl=0, nId=0, start=1, enum_prefix=None, enum_type=None):
    '''
       Insert paragraph property element with style.
    '''
    style=self.get_paragraph_style(paragraph, force_create=True)
    pPr = get_elements(paragraph, 'w:pPr')[0]

    ilvl = lvl 
    if style == 'ListNumber':
      ilvl = 0

    lvl_text='%1.'
    if nId <= 0 :
      if nId == 0 :
        num_id = '0'
      else :
        num_id = self.styleDocx.get_numbering_style_id(style)
    else :
      num_id = str(nId)
      if num_id not in self.get_numbering_ids() :

	if enum_prefix : lvl_text=enum_prefix
        newid = self.get_max_numbering_id()+1
	if newid < nId : newid = nId
        num_id = str(self.new_ListNumber_style(newid, start, lvl_text, enum_type))

    numPr_tree =[['w:numPr'], [['w:ilvl',{'w:val': str(ilvl)}]], [['w:numId',{'w:val': num_id}]] ]
    numPr = make_element_tree(numPr_tree)

    pPr.append(numPr)

    ind = self.get_numbering_indent(style, lvl, nId)
    self.set_indent(paragraph, ind)

    return pPr

  def get_ListNumber_style(self, nId):
    '''
       
    '''
    elem = get_elements(self.styleDocx.numbering, 'w:num')
    for x in elem :
      if x.get(norm_name('w:numId')) == str(nId) :
        return x
    return None

  def new_ListNumber_style_org(self, nId, start_val=1, lvl_txt='%1.', typ=None):
    '''
      create new List Number style 
    '''
    orig_numid = self.number_list_numId
    newid = nId
    typ =  get_enumerate_type(typ)

    num_tree = [['w:num', {'w:numId':str(newid)}],
                   [['w:abstractNumId', {'w:val':orig_numid}] ],
                   [['w:lvlOverride', {'w:ilvl':'0'}],
                       [['w:startOverride', {'w:val':str(start_val)}]] ,
                       [['w:lvl', {'w:ilvl':'0'}], [['w:lvlText', {'w:val': lvl_txt} ]],
                                                   [['w:numFmt', {'w:val': typ} ]]
                       ]
		  ]
               ]

    num = make_element_tree(num_tree)
    self.styleDocx.numbering.append(num)
    return  newid

  def create_dummy_nums(self, val):
    orig_numid = self.number_list_numId
    num_tree = [['w:num', {'w:numId':str(val)}],
                   [['w:abstractNumId', {'w:val': orig_numid}] ],
	  ]
    num = make_element_tree(num_tree)
    self.numids.append(num)
    return

  def new_ListNumber_style(self, nId, start_val=1, lvl_txt='%1.', typ=None):
    '''
      create new List Number style 
    '''
    newid = nId
    abstnewid = int(nId)

    cmaxid = self.get_max_numbering_id()

    if newid > cmaxid + 1 :
      for x in range(newid - cmaxid-1) :
        self.create_dummy_nums(cmaxid + x + 1)

    typ =  get_enumerate_type(typ)

    ind = self.number_list_indent
    abstnum_tree = [['w:abstractNum', {'w:abstractNumId':str(abstnewid)}],
                       [['w:multiLevelType', {'w:val':'singleLevel'}] ],
                       [['w:lvl', {'w:ilvl':'0'}],
                              [['w:start', {'w:val':str(start_val)}]] ,
			      [['w:lvlText', {'w:val': lvl_txt} ]],
                              [['w:lvlJc', {'w:val': 'left'} ]],
                              [['w:numFmt', {'w:val': typ} ]],
			      [['w:pPr'], [['w:ind',{'w:left':str(ind), 'w:hanging':str(ind)} ]]]
		       ]
                    ]

    num_tree = [['w:num', {'w:numId':str(newid)}],
                   [['w:abstractNumId', {'w:val':str(abstnewid)}] ],
               ]

    abstnum = make_element_tree(abstnum_tree)
    num = make_element_tree(num_tree)
    self.abstractNums.append(abstnum)
    self.numids.append(num)
    return  newid

########## 
##      Create New Style
  def new_character_style(self, styname):
    '''
       
    '''
    newstyle_tree = [['w:style', {'w:type':'character','w:customStye':'1', 'w:styleId': styname}],
                         [['w:name', {'w:val': styname}]],
                         [['w:basedOn', {'w:val': self.styleDocx.character_style_id}]],
                         [['w:rPr'], [['w:color', {'w:val': 'FF0000'}]] ]
                    ]

    newstyle = make_element_tree(newstyle_tree)
    self.styleDocx.styles.append(newstyle)
    self.stylenames[styname] = styname
    return styname

  def new_paragraph_style(self, styname):
    '''
       
    '''
    newstyle_tree = [['w:style', {'w:type':'paragraph','w:customStye':'1', 'w:styleId': styname}],
                         [['w:name', {'w:val': styname}]],
                         [['w:basedOn', {'w:val': self.styleDocx.paragraph_style_id}]],
                         [['w:qFormat'] ]
                    ]

    newstyle = make_element_tree(newstyle_tree)

    self.styleDocx.styles.append(newstyle)
    self.stylenames[styname] = styname
    return styname

############
## Table
  
  def get_table_cell(self, table, pos):
    '''
       
    '''
    try:
      rows = get_elements(table, 'w:tr')
      if len(rows) > pos[1] :
        return get_elements(rows[pos[1]], 'w:tc')[pos[0]]
      else :
        print  "Invalid position", pos
    except:
      print  "Error in get_table_cell", pos
    return None

  def append_paragrap_to_table_cell(self, table, paragraph, pos):
    '''
       
    '''
    cell = self.get_table_cell(table, pos)
    if len(cell) > 0 :
      cell.append(paragraph)
    return cell

  def create_table_row(self, n_cells, cellsize=None, contents=None, nline=0,firstCol=0):
    '''
      Create table row
    '''
    if nline < 0 :
      trPr_val = '100000000000'
    elif nline % 2  == 0 :
      trPr_val = '000000100000'
    else :
      trPr_val = '000000010000'

    tr_tree = [['w:tr'], [['w:trPr'], [['w:cnfStyle', {'w:val':trPr_val}]] ] ]

    row = make_element_tree(tr_tree)

    for i in range(n_cells):   
      i - firstCol
      if i < 0 :
        tcPr_val = '001000000000'
      elif i % 2  == 0 :
        tcPr_val = '000010000000'
      else :
        tcPr_val = '000001000000'

      tc_tree = [['w:tc'], [['w:tcPr'], [['w:cnfStyle', {'w:val':tcPr_val}]] ] ]
      cell = make_element_tree(tc_tree)
      row.append(cell)

      # Properties
      cellprops = get_elements(cell,'w:tcPr')[0]
      if cellsize > 0:
        cellwidth = make_element_tree([['w:tcW',{'w:w':str(cellsize[i]),'w:type':'dxa'}]])
        cellprops.append(cellwidth)

      if contents :
        cell.append(self.paragraph(contents[i], create_only=True))

      # Paragraph (Content)

    return row

  def create_table(self, colsize, tstyle='NormalTable'):
    '''
      Create table
    '''
    table_tree = [['w:tbl'],
                  [['w:tblPr'], [['w:tblStyle',{'w:val':tstyle}]], [['w:tblW',{'w:w':'0','w:type':'auto'}]] ],
		  ['w:tblGrid']
                 ]
    table = make_element_tree(table_tree)

    # Table Grid    
    tablegrid = get_elements(table,'w:tblGrid')[0]
    for csize in colsize:
        tablegrid.append(make_element_tree([['w:gridCol',{'w:w': str(csize)}]]))

    return table                 

##############
###### for reStructuredText (FieldList and Admonitions)
  def get_last_field_list_body(self, table):
    '''
       
    '''
    row = get_elements(table, 'w:tr')[-1]
    return get_elements(row, 'w:tc')[1]

  def set_field_list_item(self, table, contents, n=0):
    '''
       
    '''
    row = get_elements(table, 'w:tr')[-1]
    cell = get_elements(row, 'w:tc')[n]
    if isinstance(contents, str) :
      cell.append(self.paragraph(contents, create_only=True))
    elif isinstance(contents, list) :
      for x in contents: 
        cell.append(self.paragraph(x, create_only=True))
    else :
      print "Invalid parameter:", contents

  def insert_field_list_item(self, table, contents, n=0):
    '''
       
    '''
    row = self.create_table_row(2, self.sizeof_field_list,firstCol=1)
    table.append(row)
    self.set_field_list_item(table, contents, n)

  def insert_field_list_table(self):
    '''
       
    '''
    table = self.create_table(self.sizeof_field_list,tstyle='FieldList')
    self.append(table)
    return table

  def insert_option_list_item(self, table, contents, nrow=0):
    '''
       
    '''
    row = self.create_table_row(1, [self.max_table_width - 500], nline=nrow )
    table.append(row)
    cell = get_elements(row, 'w:tc')[0]
    if isinstance(contents, str) :
      paragraph = self.paragraph(contents, create_only=True)
      if nrow == 0:
        self.set_indent(paragraph, self.number_list_indent)
      cell.append(paragraph0r)
    elif isinstance(contents, list) :
      for x in contents: 
        paragraph = self.paragraph(x, create_only=True)
        if nrow == 0:
          self.set_indent(paragraph, self.number_list_indent)
        cell.append(paragraph)
    else :
      print "Invalid parameter:", contents

  def insert_option_list_table(self):
    '''
       
    '''
    table = self.create_table([self.max_table_width -500],tstyle='OptionList')
    self.append(table)
    return table

  def insert_admonition_table(self, contents, title='Note: ', tstyle='NoteAdmonition'):
    '''
       
    '''
    table = self.create_table([self.max_table_width-1000], tstyle=tstyle)
    for i in range(2) :
      row = self.create_table_row(1, nline=i-1)
      table.append(row)
    
    self.append_paragrap_to_table_cell(table, self.paragraph(title, create_only=True) , [0,0])

    self.append(table)
    self.insert_linespace()

    return self.get_table_cell(table, [0,1])

##############
######  Support a simple table only
  def table(self, contents, colsize=None, tstyle='rstTable'):
    '''
      Get a list of lists, return a table
      This function is copied from 'python-docx' library
    '''
    columns = len(contents[0])    

    if colsize is None : 
        for i in range(columns):
            colsize[i] = 2400
    sizeof_table = 0
    for n in colsize :
       sizeof_table += n

    colsize[-1] += self.max_table_width - sizeof_table

    table = self.create_table(colsize, tstyle=tstyle)

    for i,x in enumerate(contents) :
      row = self.create_table_row(columns, colsize, x, i-1)
      table.append(row)            

    self.append(table)
    return table                 

  def picture(self, picname, picdescription, pixelwidth=None,
            pixelheight=None, nochangeaspect=True, nochangearrowheads=True, align='center'):
    '''
      Take a relationshiplist, picture file name, and return a paragraph containing the image
      and an updated relationshiplist
      
      This function is copied from 'python-docx' library
    '''
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image. Return a paragraph containing the picture'''  
    # Copy the file into the media dir
    media_dir = join(self.template_dir,'word','media')
    if not os.path.isdir(media_dir):
        os.mkdir(media_dir)
#    picpath, picname = os.path.abspath(picname), os.path.basename(picname)

    picpath, picname = os.path.abspath(picname), os.path.basename(picname)
    picext = os.path.splitext(picname)
    self.images += 1
    if (picext[1] == '.jpg') :
      picname = 'image'+str(self.images)+'.jpeg'
    else:
      picname = 'image'+str(self.images)+picext[1]

    shutil.copyfile(picpath, join(media_dir,picname))
    relationshiplist = self.relationships

    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        # If not, get info from the picture itself
        pixelwidth,pixelheight = Image.open(picpath).size[0:2]

    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs            
    emuperpixel = 12667
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)   
    
    # Set relationship ID to the first available  
    picid = '2'    
    picrelid = 'rId'+str(len(relationshiplist)+1)
    relationshiplist.append([
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'media/'+picname])
    
    # There are 3 main elements inside a picture
    pic_tree = [['pic:pic'],
                   [['pic:nvPicPr'],  # The non visual picture properties 
                       [['pic:cNvPr', {'id':'0','name':'Picture 1','descr':picname}]],
                       [['pic:cNvPicPr'], [ ['a:picLocks', {'noChangeAspect':str(int(nochangeaspect)), 'noChangeArrowheads':str(int(nochangearrowheads))} ] ] ]
                   ],
                   [['pic:blipFill'],  # The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
                     [['a:blip',{'r:embed':picrelid}]],
                     [['a:srcRect']],
		     [['a:stretch'],[['a:fillRect']]]
		   ],
                   [['pic:spPr',{'bwMode':'auto'}],  #  The Shape properties
		     [['a:xfrm'],[['a:off',{'x':'0','y':'0'} ]], [['a:ext',{'cx':width,'cy':height}]]],
		     [['a:prstGeom',{'prst':'rect'}], ['a:avLst']],
		     [['a:noFill']]
		   ]
	       ]

    graphic_tree = [['a:graphic'],
                      [['a:graphicData', {'uri':'http://schemas.openxmlformats.org/drawingml/2006/picture'}], pic_tree ]

		      ]

    inline_tree = [['wp:inline',{'distT':"0",'distB':"0",'distL':"0",'distR':"0"}],
                       [['wp:extent',{'cx':width,'cy':height}]],
                       [['wp:effectExtent', {'l':'25400','t':'0','r':'0','b':'0'}]],
                       [['wp:docPr', {'id':picid,'name':'Picture 1','descr':picdescription}]], 
                       [['wp:cNvGraphicFramePr'], [['a:graphicFrameLocks',{'noChangeAspect':'1'} ]]],
		       graphic_tree
		       ]

    paragraph_tree = [['w:p'],
                         [['w:pPr'], [['w:jc', {'w:val':align}]]],
			 [['w:r'], [['w:rPr'], [['w:noProof']]], [['w:drawing'], inline_tree] ]
			 ]


    paragraph = make_element_tree(paragraph_tree)
    self.relationships = relationshiplist
    self.append(paragraph)

    self.last_paragraph = None
    return paragraph


  def contenttypes(self):
    '''
       create [Content_Types].xml 
       This function copied from 'python-docx' library
    '''
    prev_dir = os.getcwd() # save previous working dir
    os.chdir(self.template_dir)

    filename = '[Content_Types].xml'
    if not os.path.exists(filename):
        raise RuntimeError('You need %r file in template' % filename)

    parts = dict([
        (x.attrib['PartName'], x.attrib['ContentType'])
        for x in etree.fromstring(open(filename).read()).xpath('*')
        if 'PartName' in x.attrib
    ])

    # Add support for filetypes
    filetypes = {'rels':'application/vnd.openxmlformats-package.relationships+xml',
                 'xml':'application/xml',
                 'jpeg':'image/jpeg',
                 'jpg':'image/jpeg',
                 'gif':'image/gif',
                 'png':'image/png'}

    types_tree = [['Types']]

    for part in parts:
      types_tree.append([['Override',{'PartName':part,'ContentType':parts[part]}]])

    for extension in filetypes:
      types_tree.append([['Default',{'Extension':extension,'ContentType':filetypes[extension]}]])

    types = make_element_tree(types_tree, nsprefixes['ct'])
    os.chdir(prev_dir)
    self._contenttypes = types
    return types

  def coreproperties(self,lastmodifiedby=None):
    '''
      Create core properties (common document properties referred to in the 'Dublin Core' specification).
      See appproperties() for other stuff.
       This function copied from 'python-docx' library
    '''
    if not lastmodifiedby:
        lastmodifiedby = self.creator

    coreprops_tree = [['cp:coreProperties'],
                        [['dc:title',self.title]],
                        [['dc:subject',self.subject]],
			[['dc:creator',self.creator]],
                        [['cp:keywords',','.join(self.keywords)]],
                        [['cp:lastModifiedBy',lastmodifiedby]],
                        [['cp:revision','1']],
                        [['cp:category',self.category]],
                        [['dc:description',self.descriptions]]
		]

    currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')

    for doctime in ['created','modified']:
	coreprops_tree.append([['dcterms:'+doctime, {'xsi:type':'dcterms:W3CDTF'}, currenttime]])
        pass

    coreprops = make_element_tree(coreprops_tree)
    
    self._coreprops = coreprops
    return coreprops

  def appproperties(self):
    '''
       Create app-specific properties. See docproperties() for more common document properties.
       This function copied from 'python-docx' library
    '''
    appprops_tree = [['Properties'],
		    [['Template','Normal.dotm']],
		    [['TotalTime','6']],
		    [['Pages','1']],
		    [['Words','83']],
		    [['Characters','475']],
		    [['Application','Microsoft Word 12.0.0']],
		    [['DocSecurity','0']],
		    [['Lines','12']],
		    [['Paragraphs','8']],
		    [['ScaleCrop','false']],
		    [['LinksUpToDate','false']],
		    [['CharactersWithSpaces','583']],
		    [['SharedDoc','false']],
		    [['HyperlinksChanged','false']],
		    [['AppVersion','12.0000']],
		    [['Company',self.company]]
		    ]

    appprops=make_element_tree(appprops_tree, nsprefixes['ep'])
    self._appprops = appprops
    return appprops

  def websettings(self):
    '''
      Generate websettings
      This function copied from 'python-docx' library
    '''
    web_tree = [ ['w:webSettings'], [['w:allowPNG']], [['w:doNotSaveAsSingleFile']] ]
    web = make_element_tree(web_tree)
    self._websettings = web

    return web

  def relationshiplist(self):
    prev_dir = os.getcwd() # save previous working dir
    os.chdir(self.template_dir)

    filename = 'word/_rels/document.xml.rels'
    if not os.path.exists(filename):
        raise RuntimeError('You need %r file in template' % filename)

    relationships = etree.fromstring(open(filename).read())
    relationshiplist = [
            [x.attrib['Type'], x.attrib['Target']]
            for x in relationships.xpath('*')
    ]

    os.chdir(prev_dir)

    return relationshiplist

  def wordrelationships(self):
    '''
      Generate a Word relationships file
      This function copied from 'python-docx' library
    '''
    # Default list of relationships
    rel_tree=[['Relationships']]
    count = 0
    for relationship in self.relationships:
        # Relationship IDs (rId) start at 1.
	rel_tree.append([['Relationship',{'Id':'rId'+str(count+1), 'Type':relationship[0],'Target':relationship[1]}]])
        count += 1

    relationships = make_element_tree(rel_tree, nsprefixes['pr'])
    self._wordrelationships = relationships
    return relationships    

