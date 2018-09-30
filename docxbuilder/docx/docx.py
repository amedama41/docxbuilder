# -*- coding: utf-8 -*-
from __future__ import print_function
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
import zipfile
import shutil
import re
import six
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
    'mv': 'urn:schemas-microsoft-com:mac:vml',
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'o': 'urn:schemas-microsoft-com:office:office',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'v': 'urn:schemas-microsoft-com:vml',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    # Properties (core and extended)
    'cp': "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    'dc': "http://purl.org/dc/elements/1.1/",
    'dcterms': "http://purl.org/dc/terms/",
    'dcmitype': "http://purl.org/dc/dcmitype/",
    'xsi': "http://www.w3.org/2001/XMLSchema-instance",
    'ep': 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    # Content Types (we're just making up our own namespaces here to save time)
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships (we're just making up our own namespaces here to save time)
    'pr': 'http://schemas.openxmlformats.org/package/2006/relationships',
    # xml
    'xml': 'http://www.w3.org/XML/1998/namespace'
}

Enum_Types = {
    'arabic': 'decimal',
    'loweralpha': 'lowerLetter',
    'upperalpha': 'upperLetter',
    'lowerroman': 'lowerRoman',
    'upperroman': 'upperRoman'
}

#####################


def norm_name(tagname, namespaces=nsprefixes):
    '''
       Convert the 'tagname' to a formal expression.
          'ns:tag' --> '{namespace}tag'
          'tag' --> 'tag'
    '''
    ns_name = tagname.split(':', 1)
    if len(ns_name) > 1:
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
        if path:
            dist = xml.xpath(path, namespaces=ns)
        dist[index].append(elem)
        return True
    except:
        print("Error  in append_element")

    return False


def get_enumerate_type(typ):
    '''

    '''
    try:
        typ = Enum_Types[typ]
    except:
        typ = "decimal"
        pass
    return typ


def parse_tag_list(tag):
    '''

    '''
    tagname = ''
    tagtext = ''
    attributes = {}

    if isinstance(tag, str):
        tagname = tag
    elif isinstance(tag, list):
        tagname = tag[0]
        taglen = len(tag)
        if taglen > 1:
            if isinstance(tag[1], six.string_types):
                tagtext = tag[1]
            else:
                attributes = tag[1]
        if taglen > 2:
            if isinstance(tag[2], six.string_types):
                tagtext = tag[2]
            else:
                attributes = tag[2]
    else:
        print("Invalid tag:", tag)

    return tagname, attributes, tagtext


def extract_nsmap(tag, attributes):
    '''
    '''
    result = {}
    ns_name = tag.split(':', 1)
    if len(ns_name) > 1 and nsprefixes.get(ns_name[0]):
        result[ns_name[0]] = nsprefixes[ns_name[0]]

    for x in attributes:
        ns_name = x.split(':', 1)
        if len(ns_name) > 1 and nsprefixes.get(ns_name[0]):
            result[ns_name[0]] = nsprefixes[ns_name[0]]

    return result


def make_element_tree(arg, _xmlns=None):
    '''

    '''
    tagname, attributes, tagtext = parse_tag_list(arg[0])
    children = arg[1:]

    nsmap = extract_nsmap(tagname, attributes)

    if _xmlns is None:
        newele = etree.Element(norm_name(tagname), nsmap=nsmap)
    else:
        newele = etree.Element(norm_name(tagname), xmlns=_xmlns, nsmap=nsmap)

    if tagtext:
        newele.text = tagtext

    for attr in attributes:
        newele.set(norm_name(attr), attributes[attr])

    for child in children:
        chld = make_element_tree(child)
        if chld is not None:
            newele.append(chld)

    return newele


def get_child_element(xml, p):
    '''

    '''
    elems = get_elements(xml, p)
    if elems == []:
        ele = make_element_tree([p])
        xml.append(ele)
        return ele
    return elems[0]


def set_attributes(xml, path, attributes):
    '''

    '''
    elems = get_elements(xml, path)
    if elems == []:
        pathes = path.split('/')
        elem = xml
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
    if elems == []:
        return None
    return elems[0].attrib[norm_name(name)]

def get_special_footnotes(footnotes_xml):
    return get_elements(
            footnotes_xml,
            '/w:footnotes/w:footnote[@w:type and not(@w:type="normal")]')

#
#  DocxDocument class
#   This class for analizing docx-file
#


class DocxDocument:
    def __init__(self, docxfile):
        '''
          Constructor
        '''
        self.docx = zipfile.ZipFile(docxfile)

        self.document = self.get_xmltree('word/document.xml')
        self.docbody = get_elements(self.document, '/w:document/w:body')[0]
        self.numbering = self.get_xmltree('word/numbering.xml')
        self.styles = self.get_xmltree('word/styles.xml')
        stylenames = self.extract_stylenames()
        self.paragraph_style_id = stylenames['Normal']
        self.character_style_id = stylenames['Default Paragraph Font']
        width, height = self.get_contents_area_size()
        self.contents_width = width
        self.contents_height = height

    @property
    def footnotes(self):
        return self.get_xmltree('word/footnotes.xml')

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
        stylenames = {}
        style_elems = get_elements(self.styles, 'w:style')

        for style_elem in style_elems:
            aliases_elems = get_elements(style_elem, 'w:aliases')
            if aliases_elems:
                name = aliases_elems[0].attrib[norm_name('w:val')]
            else:
                name_elem = get_elements(style_elem, 'w:name')[0]
                name = name_elem.attrib[norm_name('w:val')]
            value = style_elem.attrib[norm_name('w:styleId')]
            stylenames[name] = value
        return stylenames

    def get_contents_area_size(self):
        paper_info = self.get_paper_info()
        paper_size = get_elements(
            self.document, '/w:document/w:body/w:sectPr/w:pgSz')[0]
        paper_margin = get_elements(
            self.document, '/w:document/w:body/w:sectPr/w:pgMar')[0]
        width = int(paper_size.get(norm_name('w:w'))) - int(paper_margin.get(
            norm_name('w:right'))) - int(paper_margin.get(norm_name('w:left')))
        height = int(paper_size.get(norm_name('w:h'))) - int(paper_margin.get(
            norm_name('w:top'))) - int(paper_margin.get(norm_name('w:bottom')))
        return width, height

    def get_paper_info(self):
        return get_elements(self.document, '/w:document/w:body/w:sectPr')[0]

    def get_coverpage(self):
        coverInfo = get_attribute(
            self.docbody, 'w:sdt/w:sdtPr/w:docPartObj/w:docPartGallery', 'w:val')
        if coverInfo == "Cover Pages":
            coverpage = get_elements(self.docbody, 'w:sdt')[0]
        else:
            coverpage = None

        return coverpage

    def get_number_of_medias(self):
        media_list = filter(lambda fname: fname.startswith('word/media/'),
                            self.docx.namelist())
        return len(list(media_list))

    def extract_files(self, to_dir, pprint=False):
        '''
          Extract all files from docx 
        '''
        try:
            if not os.access(to_dir, os.F_OK):
                os.mkdir(to_dir)

            filelist = self.docx.namelist()
            for fname in filelist:
                xmlcontent = self.docx.read(fname)
                fname_ext = os.path.splitext(fname)[1]
                if pprint and (fname_ext == '.xml' or fname_ext == '.rels'):
                    document = etree.fromstring(xmlcontent)
                    xmlcontent = etree.tostring(
                        document, encoding='UTF-8', pretty_print=True)
                file_name = join(to_dir, fname)
                if not os.path.exists(os.path.dirname(file_name)):
                    os.makedirs(os.path.dirname(file_name))
                with open(file_name, 'wb') as f:
                    f.write(xmlcontent)
        except Exception as ex:
            print("Error in extract_files ...", ex)
            return False
        return True

    def restruct_docx(self, docx_dir, docx_filename, files_to_skip=[]):
        '''
           This function is copied and modified the 'savedocx' function contained 'python-docx' library
          Restruct docx file from files in 'doxc_dir'
        '''
        if not os.access(docx_dir, os.F_OK):
            print("Can't found docx directory: %s" % docx_dir)
            return

        docxfile = zipfile.ZipFile(
            docx_filename, mode='w', compression=zipfile.ZIP_DEFLATED)

        prev_dir = os.path.abspath('.')
        os.chdir(docx_dir)

        # Add & compress support files
        files_to_ignore = ['.DS_Store']  # nuisance from some os's
        for dirpath, dirnames, filenames in os.walk('.'):
            for filename in filenames:
                if filename in files_to_ignore:
                    continue
                templatefile = join(dirpath, filename)
                archivename = os.path.normpath(templatefile)
                archivename = '/'.join(archivename.split(os.sep))
                if archivename in files_to_skip:
                    continue
                # print 'Saving: '+archivename
                docxfile.write(templatefile, archivename)

        os.chdir(prev_dir)  # restore previous working dir
        return docxfile

############
# Numbering
    def get_numbering_style_id(self, style):
        '''

        '''
        try:
            style_elems = get_elements(self.styles, '/w:styles/w:style')
            for style_elem in style_elems:
                name_elem = get_elements(style_elem, 'w:name')[0]
                name = name_elem.attrib[norm_name('w:val')]
                if name == style:
                    numPr = get_elements(
                        style_elem, 'w:pPr/w:numPr/w:numId')[0]
                    value = numPr.attrib[norm_name('w:val')]
                    return value
        except:
            pass
        return '0'

    def get_numbering_left(self, style):
        '''
           get numbering indeces
        '''
        abstractNums = get_elements(self.numbering, 'w:abstractNum')

        indres = [0]

        for x in abstractNums:
            styles = get_elements(x, 'w:lvl/w:pStyle')
            if styles:
                pstyle_name = styles[0].get(norm_name('w:val'))
                if pstyle_name == style:
                    ind = get_elements(x, 'w:lvl/w:pPr/w:ind')
                    if ind:
                        indres = []
                        for indx in ind:
                            indres.append(int(indx.get(norm_name('w:left'))))
                    return indres
        return indres

    def get_table_horizon_margin(self, style_name):
        misc_margin = 8 * 2 * 10 # Miscellaneous margin (e.g. border width)
        table_styles = get_elements(self.styles, '/w:styles/w:style')
        for style in table_styles:
            name_elem = style.find('w:name', nsprefixes)
            name = name_elem.get(norm_name('w:val'))
            if name == style_name:
                break
        else:
            return misc_margin

        cell_margin = style.find('w:tblPr/w:tblCellMar', nsprefixes)
        if cell_margin is None:
            return misc_margin # TODO: Check based style

        type_attr = norm_name('w:type')
        w_attr = norm_name('w:w')
        def get_margin(elem):
            if elem is None:
                return misc_margin
            if elem.get(type_attr) != 'dxa':
                return misc_margin
            return int(elem.get(w_attr))
        left = cell_margin.find('w:left', nsprefixes)
        right = cell_margin.find('w:right', nsprefixes)
        return get_margin(left) + get_margin(right) + misc_margin

##########

#
# DocxComposer Class
#


class DocxComposer:
    _picid = 100

    def __init__(self, stylefile=None):
        '''
           Constructor
        '''
        self._coreprops = None
        self._appprops = None
        self._contenttypes = None
        self._websettings = None
        self._wordrelationships = None
        self.stylenames = {}
        self.title = ""
        self.subject = ""
        self.creator = "Python:DocDocument"
        self.company = ""
        self.category = ""
        self.descriptions = ""
        self.keywords = []
        self.max_table_width = 8000
        self.table_margin_map = {}

        self.abstractNums = []
        self.numids = []

        self.images = 0
        self.nocoverpage = False

        self._hyperlink_rid_map = {} # target => relationship id
        self._image_rid_map = {} # imagepath => relationship id
        self._footnote_id_map = {} # docname#id => footnote id
        self._footnote_list = []
        self._max_footnote_id = 0

        if stylefile == None:
            self.template_dir = None
        else:
            self.new_document(stylefile)

    def set_style_file(self, stylefile):
        '''
           Set style file 
        '''
        self.styleDocx = DocxDocument(stylefile)

        self.template_dir = tempfile.mkdtemp(prefix='docx-')
        result = self.styleDocx.extract_files(self.template_dir)

        if not result:
            print("Unexpected error in copy_docx_to_tempfile")
            shutil.rmtree(self.template_dir, True)
            self.template_dir = None
            return

        self.stylenames = self.styleDocx.extract_stylenames()
        self.paper_info = self.styleDocx.get_paper_info()
        self.max_table_width = self.styleDocx.contents_width
        self.bullet_list_indents = self.get_numbering_left('ListBullet')
        self.bullet_list_numId = self.styleDocx.get_numbering_style_id(
            'ListBullet')
        self.number_list_indent = self.get_numbering_left('ListNumber')[0]
        self.number_list_numId = self.styleDocx.get_numbering_style_id(
            'ListNumber')
        self.abstractNums = get_elements(
            self.styleDocx.numbering, 'w:abstractNum')
        self.numids = get_elements(self.styleDocx.numbering, 'w:num')
        self.numbering = make_element_tree(['w:numbering'])
        self.images = self.styleDocx.get_number_of_medias()

        self._footnote_list.extend(
                get_special_footnotes(self.styleDocx.footnotes))
        norm_id = norm_name('w:id')
        self._max_footnote_id = max(
                map(lambda f: int(f.get(norm_id)), self._footnote_list),
                default=0)

        return

    def set_coverpage(self, flag=True):
        self.nocoverpage = not flag

    def get_numbering_ids(self):
        '''

        '''
        result = []
        for num_elem in self.numids:
            nid = num_elem.attrib[norm_name('w:numId')]
            result.append(nid)
        return result

    def get_max_numbering_id(self):
        '''

        '''
        max_id = 0
        num_ids = self.get_numbering_ids()
        for x in num_ids:
            if int(x) > max_id:
                max_id = int(x)
        return max_id

    def get_table_cell_margin(self, style_name):
        margin = self.table_margin_map.get(style_name)
        if margin is not None:
            return margin
        return self.table_margin_map.setdefault(
                style_name, self.styleDocx.get_table_horizon_margin(style_name))

    def new_document(self, stylefile):
        '''
           Preparing a new document
        '''
        self.set_style_file(stylefile)
        self.document = make_element_tree([['w:document'], [['w:body']]])
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

        for x in self.abstractNums:
            self.numbering.append(x)
        for x in self.numids:
            self.numbering.append(x)

        coverpage = self.styleDocx.get_coverpage()

        if not self.nocoverpage and coverpage is not None:
            print("output Coverpage")
            self.docbody.insert(0, coverpage)

        self.docbody.append(self.paper_info)

        footnotes = make_element_tree([['w:footnotes']])
        footnotes.extend(self._footnote_list)

        # Serialize our trees into out zip file
        treesandfiles = {self.document: 'word/document.xml',
                         self._coreprops: 'docProps/core.xml',
                         self._appprops: 'docProps/app.xml',
                         self._contenttypes: '[Content_Types].xml',
                         footnotes: 'word/footnotes.xml',
                         self.numbering: 'word/numbering.xml',
                         self.styleDocx.styles: 'word/styles.xml',
                         self._websettings: 'word/webSettings.xml',
                         self._wordrelationships: 'word/_rels/document.xml.rels'}

        docxfile = self.styleDocx.restruct_docx(
            self.template_dir, docxfilename, treesandfiles.values())

        for tree in treesandfiles:
            if tree != None:
                # print 'Saving: '+treesandfiles[tree]
                treestring = etree.tostring(
                    tree, xml_declaration=True, encoding='UTF-8', standalone='yes')
                docxfile.writestr(treesandfiles[tree], treestring)

        print('Saved new file to: '+docxfilename)
        shutil.rmtree(self.template_dir)
        return

 ##################
    @classmethod
    def make_table_of_contents(cls, toc_title, maxlevel, bookmark):
        '''
           Create the Table of Content
        '''
        toc_tree = [
                ['w:sdt'],
                [['w:sdtPr'],
                    [['w:rPr'], [['w:long']]],
                    [['w:docPartObj'],
                        [['w:docPartGallery', {'w:val': 'Table of Contents'}]],
                        [['w:docPartUnique']]
                    ]
                ]
        ]

        sdtContent_tree = [['w:sdtContent']]
        if toc_title is not None:
            sdtContent_tree.append([
                    ['w:p'],
                    [['w:pPr'], [['w:pStyle', {'w:val': 'TOC_Title'}]]],
                    [['w:r'], [['w:rPr'], [['w:long']]], [['w:t', toc_title]]]
            ])
        if maxlevel is not None:
            instr = r' TOC \o "1-%d" \b "%s" \h \z \u ' % (bookmark, maxlevel)
        else:
            instr = r' TOC \o \b "%s" \h \z \u ' % bookmark
        sdtContent_tree.append([
                ['w:p'],
                [['w:pPr'],
                    [['w:pStyle', {'w:val': 'TOC_Contents'}]],
                    [['w:tabs'],
                        [['w:tab', {
                            'w:val': 'right', 'w:leader': 'dot', 'w:pos': '8488'
                        }]]
                    ],
                    [['w:rPr'], [['w:b', {'w:val': '0'}]], [['w:noProof']]]
                ],
                [['w:r'], [['w:fldChar', {'w:fldCharType': 'begin'}]]],
                [['w:r'], [['w:instrText', instr, {'xml:space': 'preserve'}]]],
                [['w:r'], [['w:fldChar', {'w:fldCharType': 'end'}]]]
        ])

        toc_tree.append(sdtContent_tree)
        return make_element_tree(toc_tree)

#################
# Output PageBreak
    @classmethod
    def make_pagebreak(cls):
        return make_element_tree([
            ['w:p'],
            [['w:r'], [['w:br', {'w:type': 'page'}]]],
        ])

    @classmethod
    def make_sectionbreak(cls, orient='portrait'):
        if orient == 'portrait':
            attrs = {'w:w': '12240', 'w:h': '15840'}
        elif orient == 'landscape':
            attrs = {'w:h': '12240', 'w:w': '15840', 'w:orient': 'landscape'}
        return make_element_tree([
            ['w:p'],
            [['w:pPr'], [['w:sectPr'], [['w:pgSz', attrs]]]],
        ])

########
# Numbering Style

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

        if style == 'ListBullet' or nId == 0:
            if len(self.bullet_list_indents) > lvl:
                result = self.bullet_list_indents[lvl]
            else:
                result = self.bullet_list_indents[-1]
        else:
            result = self.number_list_indent * (lvl+1)

        return result

    def create_dummy_nums(self, val):
        orig_numid = self.number_list_numId
        num_tree = [['w:num', {'w:numId': str(val)}],
                    [['w:abstractNumId', {'w:val': orig_numid}]],
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

        if newid > cmaxid + 1:
            for x in range(newid - cmaxid-1):
                self.create_dummy_nums(cmaxid + x + 1)

        typ = get_enumerate_type(typ)

        ind = self.number_list_indent
        abstnum_tree = [['w:abstractNum', {'w:abstractNumId': str(abstnewid)}],
                        [['w:multiLevelType', {'w:val': 'singleLevel'}]],
                        [['w:lvl', {'w:ilvl': '0'}],
                            [['w:start', {'w:val': str(start_val)}]],
                            [['w:lvlText', {'w:val': lvl_txt}]],
                            [['w:lvlJc', {'w:val': 'left'}]],
                            [['w:numFmt', {'w:val': typ}]],
                         [['w:pPr'], [
                             ['w:ind', {'w:left': str(ind), 'w:hanging': str(ind)}]]]
                         ]
                        ]

        num_tree = [['w:num', {'w:numId': str(newid)}],
                    [['w:abstractNumId', {'w:val': str(abstnewid)}]],
                    ]

        abstnum = make_element_tree(abstnum_tree)
        num = make_element_tree(num_tree)
        self.abstractNums.append(abstnum)
        self.numids.append(num)
        return newid

##########
# Create New Style
    def new_character_style(self, styname):
        '''

        '''
        newstyle_tree = [['w:style', {'w:type': 'character', 'w:customStye': '1', 'w:styleId': styname}],
                         [['w:name', {'w:val': styname}]],
                         [['w:basedOn', {
                             'w:val': self.styleDocx.character_style_id}]],
                         [['w:rPr'], [['w:color', {'w:val': 'FF0000'}]]]
                         ]

        newstyle = make_element_tree(newstyle_tree)
        self.styleDocx.styles.append(newstyle)
        self.stylenames[styname] = styname
        return styname

    def new_paragraph_style(self, styname):
        '''

        '''
        newstyle_tree = [['w:style', {'w:type': 'paragraph', 'w:customStye': '1', 'w:styleId': styname}],
                         [['w:name', {'w:val': styname}]],
                         [['w:basedOn', {
                             'w:val': self.styleDocx.paragraph_style_id}]],
                         [['w:qFormat']]
                         ]

        newstyle = make_element_tree(newstyle_tree)

        self.styleDocx.styles.append(newstyle)
        self.stylenames[styname] = styname
        return styname

    def add_hyperlink_relationship(self, target):
        rid = self._hyperlink_rid_map.get(target)
        if rid is not None:
            return rid

        rid = 'rId%d' % (len(self.relationships) + 1)
        self.relationships.append({
            'Id': rid,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            'Target': target,
            'TargetMode': 'External'
        })
        self._hyperlink_rid_map[target] = rid
        return rid

    def add_image_relationship(self, imagepath):
        imagepath = os.path.abspath(imagepath)

        rid = self._image_rid_map.get(imagepath)
        if rid is not None:
            return rid

        # Copy the file into the media dir
        media_dir = os.path.join(self.template_dir, 'word', 'media')
        if not os.path.isdir(media_dir):
            os.mkdir(media_dir)
        picext = os.path.splitext(imagepath)
        if (picext[1] == '.jpg'):
            picext[1] = '.jpeg'
        self.images += 1
        picname = 'image%d%s' % (self.images, picext[1])
        shutil.copyfile(imagepath, os.path.join(media_dir, picname))

        # Calculate relationship ID to the first available
        rid = 'rId%d' % (len(self.relationships) + 1)
        self.relationships.append({
            'Id': rid,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'Target': 'media/' + picname
        })
        self._image_rid_map[imagepath] = rid
        return rid

    @classmethod
    def make_inline_picture_run(
            cls, rid, picname, cmwidth, cmheight, picdescription,
            nochangeaspect=True, nochangearrowheads=True):
        '''
          Take a relationship id, picture file name, and return a run element
          containing the image

          This function is based on 'python-docx' library
        '''
        # OpenXML measures on-screen objects in English Metric Units
        emupercm = 360000
        width = str(int(cmwidth * emupercm))
        height = str(int(cmheight * emupercm))

        cls._picid += 1
        picid = str(cls._picid)
        # There are 3 main elements inside a picture
        pic_tree = [['pic:pic'],
                    [['pic:nvPicPr'],  # The non visual picture properties
                     [['pic:cNvPr', {'id': picid,
                                     'name': picname, 'descr': picdescription}]],
                     [['pic:cNvPicPr'], [['a:picLocks', {
                         'noChangeAspect': str(int(nochangeaspect)),
                         'noChangeArrowheads': str(int(nochangearrowheads))}]]]
                     ],
                    # The Blipfill - specifies how the image fills the picture
                    # area (stretch, tile, etc.)
                    [['pic:blipFill'],
                     [['a:blip', {'r:embed': rid}]],
                     [['a:srcRect']],
                     [['a:stretch'], [['a:fillRect']]]
                     ],
                    [['pic:spPr', {'bwMode': 'auto'}],  # The Shape properties
                     [['a:xfrm'], [['a:off', {'x': '0', 'y': '0'}]], [
                         ['a:ext', {'cx': width, 'cy': height}]]],
                     [['a:prstGeom', {'prst': 'rect'}], ['a:avLst']],
                     [['a:noFill']]
                     ]
                    ]

        graphic_tree = [['a:graphic'],
                        [['a:graphicData', {
                            'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'}], pic_tree]

                        ]

        inline_tree = [['wp:inline', {'distT': "0", 'distB': "0", 'distL': "0", 'distR': "0"}],
                       [['wp:extent', {'cx': width, 'cy': height}]],
                       [['wp:effectExtent', {'l': '25400',
                                             't': '0', 'r': '0', 'b': '0'}]],
                       [['wp:docPr', {
                           'id': picid,
                           'name': picname, 'descr': picdescription}]],
                       [['wp:cNvGraphicFramePr'], [
                           ['a:graphicFrameLocks', {'noChangeAspect': '1'}]]],
                       graphic_tree
                       ]

        run_tree = [
                ['w:r'],
                [['w:rPr'], [['w:noProof']]],
                [['w:drawing'], inline_tree]
        ]
        return make_element_tree(run_tree)

    def set_default_footnote_id(self, key, default_fid=None):
        fid = self._footnote_id_map.get(key)
        if fid is not None:
            return fid
        if default_fid is None:
            self._max_footnote_id += 1
            default_fid = self._max_footnote_id
        self._footnote_id_map[key] = default_fid
        return default_fid

    def append_footnote(self, fid, contents):
        footnote = make_element_tree([['w:footnote', {'w:id': str(fid)}]])
        footnote.extend(contents)
        self._footnote_list.append(footnote)

    def contenttypes(self):
        '''
           create [Content_Types].xml 
           This function copied from 'python-docx' library
        '''
        filename = '[Content_Types].xml'
        filepath = os.path.join(self.template_dir, filename)
        if not os.path.exists(filepath):
            raise RuntimeError('You need %r file in template' % filename)

        with open(filepath, 'rb') as f:
            parts = dict([
                (x.attrib['PartName'], x.attrib['ContentType'])
                for x in etree.fromstring(f.read()).xpath('*')
                if 'PartName' in x.attrib
            ])

        # Add support for filetypes
        filetypes = {'rels': 'application/vnd.openxmlformats-package.relationships+xml',
                     'xml': 'application/xml',
                     'jpeg': 'image/jpeg',
                     'jpg': 'image/jpeg',
                     'gif': 'image/gif',
                     'png': 'image/png'}

        types_tree = [['Types']]

        for part in parts:
            types_tree.append(
                [['Override', {'PartName': part, 'ContentType': parts[part]}]])

        for extension in filetypes:
            types_tree.append(
                [['Default', {'Extension': extension, 'ContentType': filetypes[extension]}]])

        types = make_element_tree(types_tree, nsprefixes['ct'])
        self._contenttypes = types
        return types

    def coreproperties(self, lastmodifiedby=None):
        '''
          Create core properties (common document properties referred to in the 'Dublin Core' specification).
          See appproperties() for other stuff.
           This function copied from 'python-docx' library
        '''
        if not lastmodifiedby:
            lastmodifiedby = self.creator

        coreprops_tree = [['cp:coreProperties'],
                          [['dc:title', self.title]],
                          [['dc:subject', self.subject]],
                          [['dc:creator', self.creator]],
                          [['cp:keywords', ','.join(self.keywords)]],
                          [['cp:lastModifiedBy', lastmodifiedby]],
                          [['cp:revision', '1']],
                          [['cp:category', self.category]],
                          [['dc:description', self.descriptions]]
                          ]

        currenttime = time.strftime('%Y-%m-%dT%H:%M:%SZ')

        for doctime in ['created', 'modified']:
            coreprops_tree.append(
                [['dcterms:'+doctime, {'xsi:type': 'dcterms:W3CDTF'}, currenttime]])
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
                         [['Template', 'Normal.dotm']],
                         [['TotalTime', '6']],
                         [['Pages', '1']],
                         [['Words', '83']],
                         [['Characters', '475']],
                         [['Application', 'Microsoft Word 12.0.0']],
                         [['DocSecurity', '0']],
                         [['Lines', '12']],
                         [['Paragraphs', '8']],
                         [['ScaleCrop', 'false']],
                         [['LinksUpToDate', 'false']],
                         [['CharactersWithSpaces', '583']],
                         [['SharedDoc', 'false']],
                         [['HyperlinksChanged', 'false']],
                         [['AppVersion', '12.0000']],
                         [['Company', self.company]]
                         ]

        appprops = make_element_tree(appprops_tree, nsprefixes['ep'])
        self._appprops = appprops
        return appprops

    def websettings(self):
        '''
          Generate websettings
          This function copied from 'python-docx' library
        '''
        web_tree = [['w:webSettings'], [['w:allowPNG']],
                    [['w:doNotSaveAsSingleFile']]]
        web = make_element_tree(web_tree)
        self._websettings = web

        return web

    def relationshiplist(self):
        filename = 'word/_rels/document.xml.rels'
        filepath = os.path.join(self.template_dir, filename)
        if not os.path.exists(filepath):
            raise RuntimeError('You need %r file in template' % filename)

        with open(filepath, 'rb') as f:
            relationships = etree.fromstring(f.read())
        relationshiplist = [x.attrib for x in relationships.xpath('*')]

        return relationshiplist

    def wordrelationships(self):
        '''
          Generate a Word relationships file
          This function copied from 'python-docx' library
        '''
        # Default list of relationships
        rel_tree = [['Relationships']]
        for attributes in self.relationships:
            rel_tree.append([['Relationship', attributes]])

        relationships = make_element_tree(rel_tree, nsprefixes['pr'])
        self._wordrelationships = relationships
        return relationships
