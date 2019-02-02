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

import copy
import datetime
import os
import time
import six
import zipfile
from lxml import etree

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

#####################


def norm_name(tagname, namespaces=nsprefixes):
    '''
       Convert the 'tagname' to a formal expression.
          'ns:tag' --> '{namespace}tag'
          'tag' --> 'tag'
    '''
    if tagname.startswith('{'):
        return tagname
    ns_name = tagname.split(':', 1)
    if len(ns_name) > 1:
        tagname = "{%s}%s" % (namespaces[ns_name[0]], ns_name[1])
    return tagname


def get_elements(xml, path, ns=nsprefixes):
    '''
       Get elements from a Element tree with 'path'.
    '''
    return xml.xpath(path, namespaces=ns)


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
        raise RuntimeError("Invalid tag: %s" % tag)

    return tagname, attributes, tagtext


def extract_nsmap(tag, attributes):
    '''
    '''
    result = {}
    ns_name = tag.split(':', 1) if not tag.startswith('{') else []
    if len(ns_name) > 1 and nsprefixes.get(ns_name[0]):
        result[ns_name[0]] = nsprefixes[ns_name[0]]

    for x in attributes:
        ns_name = x.split(':', 1) if not x.startswith('{') else []
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


def get_attribute(xml, path, name):
    '''

    '''
    elems = get_elements(xml, path)
    if elems == []:
        return None
    return elems[0].attrib[norm_name(name)]

def get_special_footnotes(footnotes_xml):
    if footnotes_xml is None:
        def make_footnote(footnote_id, footnote_type):
            return make_element_tree([
                ['w:footnote', {
                    'w:type': footnote_type, 'w:id': str(footnote_id),
                }],
                [['w:p'],
                    [['w:pPr'], [['w:spacing', {'w:after': '0'}]]],
                    [['w:r'], [['w:' + footnote_type]]],
                ],
            ])
        return [
                make_footnote(-1, 'separate'),
                make_footnote(0, 'continuationSeparator')
        ]
    return get_elements(
            footnotes_xml,
            '/w:footnotes/w:footnote[@w:type and not(@w:type="normal")]')

def get_max_attribute(elems, attribute):
    '''
       Get the maximum integer attribute among the specified elems
    '''
    if not elems:
        return 0
    num_id = norm_name('w:numId')
    return max(map(lambda e: int(e.get(attribute)), elems))

def local_to_utc(value):
    utc = datetime.datetime.utcfromtimestamp(time.mktime(value.timetuple()))
    return utc.replace(microsecond=value.microsecond)

def convert_to_W3CDTF_string(value):
    if isinstance(value, datetime.datetime):
        if value.tzinfo is not None:
            offset = value.utcoffset()
            value = value.replace(tzinfo=None) - offset
        else:
            value = local_to_utc(value)
        return value.strftime('%Y-%m-%dT%H:%M:%SZ')
    if isinstance(value, datetime.date):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, six.string_types):
        for date_format in ('%Y', '%Y-%m', '%Y-%m-%d'):
            try:
                datetime.datetime.strptime(value, date_format)
            except ValueError:
                continue
            return value
        datetime_formats = [
                ('%Y-%m-%dT%H', '%Y-%m-%dT%H:%MZ'),
                ('%Y-%m-%dT%H:%M', '%Y-%m-%dT%H:%MZ'),
                ('%Y-%m-%dT%H:%M:%S', '%Y-%m-%dT%H:%M:%SZ'),
                ('%Y-%m-%dT%H:%M:%S.%f', '%Y-%m-%dT%H:%M:%S.%fZ'),
        ]
        for from_format, to_format in datetime_formats:
            try:
                d = local_to_utc(datetime.datetime.strptime(value, from_format))
            except ValueError:
                continue
            return d.strftime(to_format)
    return None

#
#  DocxDocument class
#   This class for analizing docx-file
#

def normalize_coreproperties(props):
    invalid_props = []

    last_printed = props.get('lastPrinted', None)
    if last_printed is not None:
        if isinstance(last_printed, datetime.datetime):
            props['lastPrinted'] = last_printed.strftime('%Y-%m-%dT%H:%M:%S')
        else:
            try:
                datetime.datetime.strptime(last_printed, '%Y-%m-%dT%H:%M:%S')
            except ValueError:
                invalid_props.append('lastPrinted')

    for doctime in ['created', 'modified']:
        value = props.get(doctime, None)
        if value is None:
            continue
        value = convert_to_W3CDTF_string(value)
        if value is None:
            invalid_props.append(doctime)
        else:
            props[doctime] = value

    for p in invalid_props:
        del props[p]
    return invalid_props

def get_orient(section_prop):
    page_size = get_elements(section_prop, 'w:pgSz')[0]
    return page_size.attrib.get(norm_name('w:orient'), 'portrait')

def rotate_orient(section_prop):
    page_size = get_elements(section_prop, 'w:pgSz')[0]
    orient_attr = norm_name('w:orient')
    current_orient = page_size.attrib.get(orient_attr, 'portrait')
    orient = 'landscape' if current_orient == 'portrait' else 'portrait'
    w_attr = norm_name('w:w')
    h_attr = norm_name('w:h')
    w = page_size.attrib.get(w_attr)
    h = page_size.attrib.get(h_attr)
    page_size.attrib[w_attr] = h
    page_size.attrib[h_attr] = w
    page_size.attrib[orient_attr] = orient
    return section_prop

def set_title_page(section_prop, is_title_page):
    value = 'true' if is_title_page else 'false'
    title_page = get_elements(section_prop, 'w:titlePg')
    if not title_page:
        section_prop.append(
                make_element_tree([['w:titlePg', {'w:val': value}]]))
        return
    title_page[0].attrib[norm_name('w:val')] = value

def get_contents_width(section_property):
    paper_size = get_elements(section_property, 'w:pgSz')[0]
    width = int(paper_size.get(norm_name('w:w')))
    paper_margin = get_elements(section_property, 'w:pgMar')[0]
    width_margin = (
            int(paper_margin.get(norm_name('w:right'))) +
            int(paper_margin.get(norm_name('w:left'))))
    return width - width_margin

# Paragraphs and Runs

def add_page_break_before_to_first_paragraph(xml):
    paragraphs = get_elements(xml, '//w:p')
    if not paragraphs:
        return
    p = paragraphs[0]
    p_props = get_elements(p, 'w:pPr')
    tree = [['w:pageBreakBefore', {'w:val': '1'}]]
    if p_props:
        p_props[0].append(make_element_tree(tree))
    else:
        p.append(make_element_tree([['w:pPr', tree]]))

def make_run_style_property(style_id):
    return {'w:rStyle': {'w:val': style_id}}

def make_paragraph(
        indent, right_indent, style, align, keep_lines, keep_next, list_info):
    style_tree = [
            ['w:pPr'],
            [['w:pStyle', {'w:val': style}]],
    ]
    ind_attrs = {}
    if list_info is not None:
        num_id, list_level = list_info
        style_tree.append([
            ['w:numPr'],
            [['w:ilvl', {'w:val': str(list_level)}]],
            [['w:numId', {'w:val': str(num_id)}]],
        ])
    if indent is not None:
        ind_attrs['w:leftChars'] = '0'
        ind_attrs['w:left'] = str(indent)
    if right_indent is not None:
        ind_attrs['w:right'] = str(right_indent)
    if ind_attrs:
        style_tree.append([['w:ind', ind_attrs]])
    if align is not None:
        style_tree.append([['w:jc', {'w:val': align}]])
    if keep_lines:
        style_tree.append([['w:keepLines']])
    if keep_next:
        style_tree.append([['w:keepNext']])

    paragraph_tree = [['w:p'], style_tree]
    return make_element_tree(paragraph_tree)

def make_section_prop_paragraph(section_prop):
    section_prop = copy.deepcopy(section_prop)
    p = make_element_tree([['w:p'], [['w:pPr']]])
    p[0].append(section_prop)
    return p

def make_run(text, style, preserve_space):
    run_tree = [['w:r']]
    run_prop = [['w:rPr']]
    for tagname, attrib in style.items():
        run_prop.append([[tagname, attrib]])
    if len(run_prop) != 1:
        run_tree.append(run_prop)
    if preserve_space:
        lines = text.split('\n')
        for index, line in enumerate(lines):
            run_tree.append([['w:t', line, {'xml:space': 'preserve'}]])
            if index != len(lines) - 1:
                run_tree.append([['w:br']])
    else:
        text = text.replace('\n', ' ')
        attrs = {}
        if text.startswith(' ') or text.endswith(' '):
            attrs['xml:space'] = 'preserve'
        run_tree.append([['w:t', text, attrs]])
    return make_element_tree(run_tree)

def make_break_run():
    return make_element_tree([['w:r'], [['w:br']]])

def make_inline_picture_run(
        rid, picid, picname, cmwidth, cmheight, picdescription,
        nochangeaspect=True, nochangearrowheads=True):
    '''
      Take a relationship id, picture file name, and return a run element
      containing the image

      This function is based on 'python-docx' library
    '''
    non_visual_pic_prop_attrs = {
            'id': str(picid), 'name': picname, 'descr': picdescription
    }
    # OpenXML measures on-screen objects in English Metric Units
    emupercm = 360000
    ext_attrs = {
            'cx': str(int(cmwidth * emupercm)),
            'cy': str(int(cmheight * emupercm))
    }

    # There are 3 main elements inside a picture
    pic_tree = [
            ['pic:pic'],
            [['pic:nvPicPr'],  # The non visual picture properties
                [['pic:cNvPr', non_visual_pic_prop_attrs]],
                [['pic:cNvPicPr'],
                    [['a:picLocks', {
                        'noChangeAspect': str(int(nochangeaspect)),
                        'noChangeArrowheads': str(int(nochangearrowheads))}]
                    ]
                ]
            ],
            # The Blipfill - specifies how the image fills the picture
            # area (stretch, tile, etc.)
            [['pic:blipFill'],
                [['a:blip', {'r:embed': rid}]],
                [['a:srcRect']],
                [['a:stretch'], [['a:fillRect']]]
            ],
            [['pic:spPr', {'bwMode': 'auto'}],  # The Shape properties
                [['a:xfrm'],
                    [['a:off', {'x': '0', 'y': '0'}]],
                    [['a:ext', ext_attrs]]
                ],
                [['a:prstGeom', {'prst': 'rect'}], ['a:avLst']],
                [['a:noFill']]
            ]
    ]

    graphic_tree = [
            ['a:graphic'],
            [['a:graphicData', {
                'uri': 'http://schemas.openxmlformats.org/drawingml/2006/picture'}],
                pic_tree
            ]
    ]

    inline_tree = [
            ['wp:inline', {'distT': "0", 'distB': "0", 'distL': "0", 'distR': "0"}],
            [['wp:extent', ext_attrs]],
            [['wp:effectExtent', {'l': '25400', 't': '0', 'r': '0', 'b': '0'}]],
            [['wp:docPr', non_visual_pic_prop_attrs]],
            [['wp:cNvGraphicFramePr'],
                [['a:graphicFrameLocks', {'noChangeAspect': '1'}]]
            ],
            graphic_tree
    ]

    run_tree = [
            ['w:r'],
            [['w:rPr'], [['w:noProof']]],
            [['w:drawing'], inline_tree]
    ]
    return make_element_tree(run_tree)


# Tables

def make_table(style, indent, align, grid_col_list, has_head, has_first_column):
    look_attrs = {
            'w:noHBand': 'false', 'w:noVBand': 'false',
            'w:lastRow': 'false', 'w:lastColumn': 'false'
    }
    look_attrs['w:firstRow'] = 'true' if has_head else 'false'
    look_attrs['w:firstColumn'] = 'true' if has_first_column else 'false'
    property_tree = [
            ['w:tblPr'],
            [['w:tblW', {'w:w': '0', 'w:type': 'auto'}]],
            [['w:tblInd', {'w:w': str(indent), 'w:type': 'dxa'}]],
            [['w:tblLook', look_attrs]],
    ]
    if style is not None:
        property_tree.insert(1, [['w:tblStyle', {'w:val': style}]])
    if align is not None:
        property_tree.append([['w:jc', {'w:val': align}]])

    table_grid_tree = [['w:tblGrid']]
    for grid_col in grid_col_list:
        table_grid_tree.append([['w:gridCol', {'w:w': str(grid_col)}]])

    table_tree = [
            ['w:tbl'],
            property_tree,
            table_grid_tree
    ]
    return make_element_tree(table_tree)

def make_row(index, is_head, cant_split, set_tbl_header):
    row_style_attrs = {
            'w:evenHBand': ('true' if index % 2 == 0 else 'false'),
            'w:oddHBand': ('true' if index % 2 != 0 else 'false'),
            'w:firstRow': ('true' if is_head else 'false'),
    }
    property_tree = [
            ['w:trPr'],
            [['w:cnfStyle', row_style_attrs]],
    ]
    if cant_split:
        property_tree.append([['w:cantSplit']])
    if set_tbl_header:
        property_tree.append([['w:tblHeader']])
    return make_element_tree([['w:tr'], property_tree])

def make_cell(index, is_first_column, cellsize, grid_span, vmerge):
    cell_style = {
            'w:evenVBand': ('true' if index % 2 == 0 else 'false'),
            'w:oddVBand': ('true' if index % 2 != 0 else 'false'),
            'w:firstColumn': ('true' if is_first_column else 'false'),
    }
    property_tree = [
            ['w:tcPr'],
            [['w:cnfStyle', cell_style]],
    ]
    if cellsize is not None:
        property_tree.append(
                [['w:tcW', {'w:w': str(cellsize), 'w:type': 'dxa'}]])
    if grid_span > 1:
        property_tree.append([['w:gridSpan', {'w:val': str(grid_span)}]])
    if vmerge is not None:
        property_tree.append([['w:vMerge', {'w:val': vmerge}]])
    return make_element_tree([['w:tc'], property_tree])

# Footnotes

def make_footnote_reference(footnote_id, style_id):
    return make_element_tree([
        ['w:r'],
        [['w:rPr'], [['w:rStyle', {'w:val': style_id}]]],
        [['w:footnoteReference', {'w:id': str(footnote_id)}]],
    ])

def make_footnote_ref(style_id):
    return make_element_tree([
        ['w:r'],
        [['w:rPr'], [['w:rStyle', {'w:val': style_id}]]],
        [['w:footnoteRef']],
    ])


# Annotations

def make_bookmark_start(id, name):
    return make_element_tree([
        ['w:bookmarkStart', {'w:id': str(id), 'w:name': name}]
    ])

def make_bookmark_end(id):
    return make_element_tree([['w:bookmarkEnd', {'w:id': str(id)}]])


# Hyperlinks

def make_hyperlink(relationship_id, anchor):
    attrs = {}
    if relationship_id is not None:
        attrs['r:id'] = relationship_id
    if anchor is not None:
        attrs['w:anchor'] = anchor
    hyperlink_tree = [['w:hyperlink', attrs]]
    return make_element_tree(hyperlink_tree)

# Structured Document Tags

def _make_toc_hyperlink(text, anchor):
    return [['w:hyperlink', {'w:anchor': anchor, 'w:history': '1'}],
            [['w:r'], [['w:t', text]]],
            [['w:r'], [['w:rPr'], [['w:webHidden']]], [['w:tab']]],
            [['w:r'], [['w:fldChar', {'w:fldCharType': 'begin'}]]],
            [['w:r'],
                [['w:instrText',
                    r' PAGEREF %s \h ' % anchor, {'xml:space': 'preserve'}
                ]]
            ],
            [['w:r'], [['w:fldChar', {'w:fldCharType': 'separate'}]]],
            [['w:r'], [['w:rPr'], [['w:webHidden']]], [['w:t', 'X']]],
            [['w:r'], [['w:fldChar', {'w:fldCharType': 'end'}]]],
    ]

def make_table_of_contents(toc_title, style_id, maxlevel, bookmark, outlines):
    '''
       Create the Table of Content
    '''
    sdtContent_tree = [['w:sdtContent']]
    if toc_title is not None:
        sdtContent_tree.append([
            ['w:p'],
            [['w:pPr'], [['w:pStyle', {'w:val': style_id}]]],
            [['w:r'], [['w:t', toc_title]]]
        ])
    if maxlevel is not None:
        instr = r' TOC \o "1-%d" \b "%s" \h \z \u ' % (maxlevel, bookmark)
    else:
        instr = r' TOC \o \b "%s" \h \z \u ' % bookmark
    tabs_tree = [
            ['w:tabs'],
            [['w:tab', {'w:val': 'right', 'w:leader': 'dot', 'w:pos': '8488'}]]
    ]
    run_prop_tree = [['w:rPr'], [['w:b', {'w:val': '0'}]], [['w:noProof']]]
    if outlines:
        sdtContent_tree.append([
            ['w:p'],
            [['w:pPr'],
                [['w:pStyle', {'w:val': outlines[0][1]}]],
                tabs_tree,
                run_prop_tree,
            ],
            [['w:r'], [['w:fldChar', {'w:fldCharType': 'begin'}]]],
            [['w:r'], [['w:instrText', instr, {'xml:space': 'preserve'}]]],
            [['w:r'], [['w:fldChar', {'w:fldCharType': 'separate'}]]],
            _make_toc_hyperlink(outlines[0][0], outlines[0][2]),
        ])
        for text, style_id, anchor in outlines[1:]:
            sdtContent_tree.append([
                ['w:p'],
                [['w:pPr'],
                    [['w:pStyle', {'w:val': style_id}]],
                    tabs_tree,
                    run_prop_tree,
                ],
                _make_toc_hyperlink(text, anchor),
            ])
        sdtContent_tree.append([
            ['w:p'],
            [['w:r'], [['w:fldChar', {'w:fldCharType': 'end'}]]]
        ])
    else:
        sdtContent_tree.append([
            ['w:p'],
            [['w:pPr'],
                tabs_tree,
                run_prop_tree,
            ],
            [['w:r'], [['w:fldChar', {'w:fldCharType': 'begin'}]]],
            [['w:r'], [['w:instrText', instr, {'xml:space': 'preserve'}]]],
            [['w:r'], [['w:fldChar', {'w:fldCharType': 'end'}]]],
        ])

    toc_tree = [
            ['w:sdt'],
            [['w:sdtPr'],
                [['w:docPartObj'],
                    [['w:docPartGallery', {'w:val': 'Table of Contents'}]],
                    [['w:docPartUnique']]
                ]
            ],
            sdtContent_tree
    ]
    return make_element_tree(toc_tree)

def make_vml_textbox(style, color, contents, wrap_style=None):
    rect_tree = [
            ['v:rect', { 'style': style, 'fillcolor': color }],
            [['v:textbox', {'style': 'mso-fit-shape-to-text:true'}],
                [['w:txbxContent']]
            ]
    ]
    if wrap_style is not None:
        rect_tree.append([['w10:wrap', wrap_style]])
    txbx = make_element_tree([['w:r'], [['w:pict'], rect_tree]])
    txbx[0][0][0][0].extend(contents)
    return txbx

def get_left(ind):
    left = ind.get(norm_name('w:left'), None)
    if left is not None:
        return left
    return ind.get(norm_name('w:start'), '0')

class DocxDocument:
    def __init__(self, docxfile):
        '''
          Constructor
        '''
        self.docx = zipfile.ZipFile(docxfile)

        self.document = self.get_xmltree('word/document.xml')
        self.numbering = self.get_xmltree('word/numbering.xml')
        self.styles = self.get_xmltree('word/styles.xml')

    @property
    def footnotes(self):
        return self.get_xmltree('word/footnotes.xml')

    def get_xmltree(self, fname):
        '''
          Extract a document tree from the docx file
        '''
        try:
            return etree.fromstring(self.docx.read(fname))
        except KeyError:
            return None

    def extract_style_info(self):
        '''
          Extract all style name/id/type from the docx file
        '''
        style_id_attr = norm_name('w:styleId')
        type_attr = norm_name('w:type')
        val_attr = norm_name('w:val')
        def get_info(style):
            style_id = style.attrib[style_id_attr]
            style_type = style.attrib[type_attr]
            names = get_elements(style, 'w:name')
            style_name = names[0].attrib[val_attr] if names else style_id
            return (style_name, (style_id, style_type))
        return dict(get_info(s) for s in get_elements(self.styles, 'w:style'))

    def get_default_style_name(self, style_type):
        '''
          Extract the last default style's id with style_type
        '''
        xpath = 'w:style[@w:type="%s" and (@w:default="1" or @w:default="true")]'
        styles = get_elements(self.styles, xpath % style_type)
        if not styles:
            return None
        name = get_attribute(styles[-1], 'w:name', 'w:val')
        if name is not None:
            return name
        else:
            return styles[-1].attrib[norm_name('w:styleId')]

    def get_section_properties(self):
        return get_elements(self.document, '//w:sectPr')

    def get_coverpage(self):
        coverpages = get_elements(
            self.document,
            '//w:sdt[w:sdtPr/w:docPartObj/w:docPartGallery[@w:val="Cover Pages"]]')
        return coverpages[0] if coverpages else None

    def get_number_of_medias(self):
        media_list = filter(lambda fname: fname.startswith('word/media/'),
                            self.docx.namelist())
        return len(list(media_list))

    def collect_items(self, zip_docxfile, files_to_skip=[]):
        # Add & compress support files
        filelist = self.docx.namelist()
        for fname in filter(lambda f: f not in files_to_skip, filelist):
            zip_docxfile.writestr(fname, self.docx.read(fname))

############
# Numbering
    def get_run_style_property(self, style_id):
        props = get_elements(
                self.styles,
                '/w:styles/w:style[@w:styleId="%s"]/w:rPr' % style_id)
        if not props:
            return {}
        return [(prop.tag, prop.attrib)
                for prop in props[0] if not prop.tag.endswith('rPrChange')]

    def get_numbering_style_id(self, style):
        '''

        '''
        style_elems = get_elements(self.styles, '/w:styles/w:style')
        for style_elem in style_elems:
            name_elem = get_elements(style_elem, 'w:name')[0]
            name = name_elem.attrib[norm_name('w:val')]
            if name == style:
                numPr = get_elements(
                    style_elem, 'w:pPr/w:numPr/w:numId')[0]
                value = numPr.attrib[norm_name('w:val')]
                return value
        return None

    def get_elems_from_numbering(self, elem_tag):
        if self.numbering is None:
            return []
        return get_elements(self.numbering, elem_tag)

    def get_indent(self, style_id):
        ind_elems = get_elements(
                self.styles,
                '/w:styles/w:style[@w:styleId="%s"]/w:pPr/w:ind' % style_id)
        if not ind_elems:
            return None
        return get_left(ind_elems[0])

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
    def __init__(self, stylefile):
        '''
           Constructor
        '''
        self._id = 100
        self.styleDocx = DocxDocument(stylefile)

        self._style_info = self.styleDocx.extract_style_info()
        self._abstract_nums = self.styleDocx.get_elems_from_numbering(
                'w:abstractNum')
        self._max_abstract_num_id = get_max_attribute(
                self._abstract_nums, norm_name('w:abstractNumId'))
        self._numids = self.styleDocx.get_elems_from_numbering('w:num')
        self._max_num_id = get_max_attribute(self._numids, norm_name('w:numId'))
        self.images = self.styleDocx.get_number_of_medias()

        self._hyperlink_rid_map = {} # target => relationship id
        self._image_info_map = {} # imagepath => (relationship id, imagename)

        self._footnote_list = get_special_footnotes(self.styleDocx.footnotes)
        self._footnote_id_map = {} # docname#id => footnote id
        norm_id = norm_name('w:id')
        self._max_footnote_id = get_max_attribute(
                self._footnote_list, norm_name('w:id'))

        self._run_style_property_cache = {}
        self.table_margin_map = {}

        self.document = make_element_tree([['w:document'], [['w:body']]])
        self.docbody = get_elements(self.document, '/w:document/w:body')[0]
        self.relationships = self.relationshiplist()

    def new_id(self):
        self._id += 1
        return self._id

    def get_each_orient_section_properties(self):
        section_props = self.styleDocx.get_section_properties()
        first = section_props[0]
        first_orient = get_orient(first)
        for sect_prop in section_props[1:]:
            if get_orient(sect_prop) != first_orient:
                return first, sect_prop
        return first, rotate_orient(copy.deepcopy(first))

    def get_style_info(self, style_name):
        style_info = self._style_info.get(style_name, None)
        if style_info is not None:
            return style_info
        style_info = self._style_info.get(style_name.lower(), None)
        if style_info is not None:
            return style_info
        return None

    def get_style_id(self, style_name):
        style_info = self.get_style_info(style_name)
        return style_info[0] if style_info is not None else style_name

    def get_indent(self, style_name, default):
        style_info = self.get_style_info(style_name)
        if style_info is None or style_info[1] != 'paragraph':
            return default
        indent = self.styleDocx.get_indent(style_info[0])
        if indent is None:
            return default
        return int(indent)

    def get_run_style_property(self, style_id):
        style = self._run_style_property_cache.get(style_id)
        if style is not None:
            return style
        return self._run_style_property_cache.setdefault(
                style_id, self.styleDocx.get_run_style_property(style_id))

    def get_bullet_list_num_id(self):
        return self.styleDocx.get_numbering_style_id('List Bullet')

    def get_table_cell_margin(self, style_name):
        margin = self.table_margin_map.get(style_name)
        if margin is not None:
            return margin
        return self.table_margin_map.setdefault(
                style_name, self.styleDocx.get_table_horizon_margin(style_name))

    def save(self, docxfilename, has_coverpage, title, creator, language, props):
        '''
          Save the composed document to the docx file 'docxfilename'.
        '''
        coreproperties = self.coreproperties(title, creator, language, props)
        appproperties = self.appproperties(props.get('company', ''))
        contenttypes = self.contenttypes()
        websettings = self.websettings()

        wordrelationships = self.wordrelationships()

        numbering = make_element_tree(['w:numbering'])
        numbering.extend(self._abstract_nums)
        numbering.extend(self._numids)

        coverpage = self.styleDocx.get_coverpage()

        if has_coverpage and coverpage is not None:
            self.docbody.insert(0, coverpage)

        footnotes = make_element_tree([['w:footnotes']])
        footnotes.extend(self._footnote_list)

        # Serialize our trees into out zip file
        treesandfiles = {self.document: 'word/document.xml',
                         coreproperties: 'docProps/core.xml',
                         appproperties: 'docProps/app.xml',
                         contenttypes: '[Content_Types].xml',
                         footnotes: 'word/footnotes.xml',
                         numbering: 'word/numbering.xml',
                         self.styleDocx.styles: 'word/styles.xml',
                         websettings: 'word/webSettings.xml',
                         wordrelationships: 'word/_rels/document.xml.rels'}

        docxfile = zipfile.ZipFile(
            docxfilename, mode='w', compression=zipfile.ZIP_DEFLATED)

        self.styleDocx.collect_items(docxfile, treesandfiles.values())

        for tree, xmlpath in treesandfiles.items():
            treestring = etree.tostring(
                tree, xml_declaration=True, encoding='UTF-8', standalone='yes')
            docxfile.writestr(xmlpath, treestring)

        for imgpath, (_, picname) in self._image_info_map.items():
            docxfile.write(imgpath, 'word/media/' + picname)

 ##################
########
# Numbering Style

    def get_numbering_left(self, style_name):
        '''
           Get numbering indeces...
        '''
        num_id = self.styleDocx.get_numbering_style_id(style_name)
        if num_id is None:
            return []

        def find_elem(elems, attr, value):
            for elem in elems:
                if elem.get(attr, None) == value:
                    return elem
            return None

        num = find_elem(self._numids, norm_name('w:numId'), num_id)
        if num is None:
            return []

        abst_num_id = get_attribute(num, 'w:abstractNumId', 'w:val')
        abstract_num = find_elem(
                self._abstract_nums, norm_name('w:abstractNumId'), abst_num_id)
        if abstract_num is None:
            return []

        indent_info = []
        ilvl_attr = norm_name('w:ilvl')
        for lvl in get_elements(abstract_num, 'w:lvl'):
            ind = get_elements(lvl, 'w:pPr/w:ind')
            if ind:
                indent_info.append(
                        (int(lvl.get(ilvl_attr)), int(get_left(ind[-1]))))
        indent_info.sort()

        indents = []
        for lvl, indent in indent_info:
            while len(indents) < lvl:
                indents.append(indents[-1] if indents else 0)
            indents.append(indent)
        return indents

    num_format_map = {
        'bullet': 'bullet',
        'arabic': 'decimal',
        'loweralpha': 'lowerLetter',
        'upperalpha': 'upperLetter',
        'lowerroman': 'lowerRoman',
        'upperroman': 'upperRoman',
    }

    def add_numbering_style(
            self, start_val, lvl_txt, typ, indent, style_id=None, font=None):
        '''
           Create a new numbering definition
        '''
        self._max_abstract_num_id += 1
        abstract_num_id = self._max_abstract_num_id
        typ = self.__class__.num_format_map.get(typ, 'decimal')
        lvl_tree = [
                ['w:lvl', {'w:ilvl': '0'}],
                [['w:start', {'w:val': str(start_val)}]],
                [['w:lvlText', {'w:val': lvl_txt}]],
                [['w:lvlJc', {'w:val': 'left'}]],
                [['w:numFmt', {'w:val': typ}]],
                [['w:pPr'], [['w:ind', {
                    'w:left': str(indent), 'w:hanging': str(int(indent * 0.75))
                }]]],
        ]
        if style_id is not None:
            lvl_tree.append([['w:pStyle', {'w:val': style_id}]])
        if font is not None:
            lvl_tree.append([
                ['w:rPr'], [['w:rFonts', {'w:ascii': font, 'w:hAnsi': font}]]
            ])
        abstnum = make_element_tree([
            ['w:abstractNum', {'w:abstractNumId': str(abstract_num_id)}],
            [['w:multiLevelType', {'w:val': 'singleLevel'}]],
            lvl_tree,
        ])
        self._abstract_nums.append(abstnum)

        self._max_num_id += 1
        num_id = self._max_num_id
        num_tree = [
                ['w:num', {'w:numId': str(num_id)}],
                [['w:abstractNumId', {'w:val': str(abstract_num_id)}]],
        ]
        num = make_element_tree(num_tree)
        self._numids.append(num)
        return num_id

    def get_default_style_names(self):
        '''
           Return default paragraph, character, table style ids
        '''
        paragraph_style_id = self.styleDocx.get_default_style_name('paragraph')
        character_style_id = self.styleDocx.get_default_style_name('character')
        table_style_id = self.styleDocx.get_default_style_name('table')
        return paragraph_style_id, character_style_id, table_style_id

    def create_style(self, style_type, new_style_name, based_style_name):
        '''
           Create a new style_stype style with new_style_id,
           which is based on based_style_id.
        '''
        if self.get_style_info(new_style_name) is not None:
            return False
        new_style_id = new_style_name
        style_tree = [
                ['w:style', {
                    'w:type': style_type,
                    'w:customStye': '1',
                    'w:styleId': new_style_id
                }],
                [['w:name', {'w:val': new_style_name}]],
                [['w:qFormat']],
        ]
        based_style_info = self.get_style_info(based_style_name)
        if based_style_info is not None and based_style_info[1] == style_type:
            style_tree.append([['w:basedOn', {'w:val': based_style_info[0]}]])
        newstyle = make_element_tree(style_tree)
        self.styleDocx.styles.append(newstyle)
        self._style_info[new_style_name] = (new_style_id, style_type)
        return True

    def create_list_style(
            self, new_style_name, format_type, lvl_text, font, indent):
        if self.get_style_info(new_style_name) is not None:
            return
        new_style_id = new_style_name
        num_id = self.add_numbering_style(
                1, lvl_text, format_type, indent, new_style_id, font)
        style_tree = [
                ['w:style', {'w:type': 'paragraph', 'w:styleId': new_style_id}],
                [['w:name', {'w:val': new_style_name}]],
                [['w:qFormat']],
                [['w:pPr'],
                    [['w:numPr'],
                        [['w:ilvl', {'w:val': '0'}]],
                        [['w:numId', {'w:val': str(num_id)}]],
                    ],
                ],
        ]
        newstyle = make_element_tree(style_tree)
        self.styleDocx.styles.append(newstyle)
        self._style_info[new_style_name] = (new_style_id, 'paragraph')
        return True

    def create_empty_paragraph_style(
            self, new_style_name, after_space, with_border):
        '''
           Create a new empty paragraph style
        '''
        if self.get_style_info(new_style_name) is not None:
            return
        new_style_id = new_style_name
        property_tree = [
                ['w:pPr'],
                [['w:spacing', {
                    'w:before': '0', 'w:beforeAutospacing': '0',
                    'w:after': str(after_space), 'w:afterAutospacing': '0',
                }]],
                [['w:rPr'], [['w:sz', {'w:val': '16'}]]],
        ]
        if with_border:
            property_tree.append([
                ['w:pBdr'],
                [['w:bottom', {'w:val': 'single', 'w:sz': '8', 'w:space': '1'}]]
            ])
        new_style = make_element_tree([
                ['w:style', {
                    'w:type': 'paragraph',
                    'w:customStye': '1',
                    'w:styleId': new_style_id
                }],
                [['w:name', {'w:val': new_style_name}]],
                [['w:qFormat']],
                property_tree,
        ])
        self.styleDocx.styles.append(new_style)
        self._style_info[new_style_name] = (new_style_id, 'paragraph')

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

        rid, _ = self._image_info_map.get(imagepath, (None, None))
        if rid is not None:
            return rid

        _, picext = os.path.splitext(imagepath)
        if picext == '.jpg':
            picext = '.jpeg'
        self.images += 1
        picname = 'image%d%s' % (self.images, picext)

        # Calculate relationship ID to the first available
        rid = 'rId%d' % (len(self.relationships) + 1)
        self.relationships.append({
            'Id': rid,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'Target': 'media/' + picname
        })
        self._image_info_map[imagepath] = (rid, picname)
        return rid

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
        content_types = self.styleDocx.get_xmltree(filename)

        parts = dict(
                (x.attrib['PartName'], x.attrib['ContentType'])
                for x in content_types.xpath('*') if 'PartName' in x.attrib)

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

        return make_element_tree(types_tree, nsprefixes['ct'])

    def coreproperties(self, title, creator, language, props):
        '''
           Create core properties (common document properties referred to in 
           the 'Dublin Core' specification).
           See appproperties() for other stuff.
        '''
        coreprops_tree = [
                ['cp:coreProperties'],
                [['dc:title', title]],
                [['dc:creator', creator]],
                [['dc:language', language]],
        ]
        properties = [
                ('cp', 'category'),
                ('cp', 'contentStatus'),
                ('dc', 'description'),
                ('dc', 'identifier'),
                ('cp', 'lastModifiedBy'),
                ('cp', 'lastPrinted'),
                ('cp', 'revision'),
                ('dc', 'subject'),
                ('cp', 'version'),
        ]
        for ns, prop in properties:
            value = props.get(prop, None)
            if value is None:
                continue
            coreprops_tree.append([['%s:%s' % (ns, prop), value]])

        keywords = props.get('keywords', None)
        if keywords is not None:
            if isinstance(keywords, (list, tuple)):
                keywords = ','.join(keywords)
            coreprops_tree.append([['cp:keywords', keywords]])

        for doctime in ['created', 'modified']:
            value = props.get(doctime, None)
            if value is None:
                continue
            coreprops_tree.append(
                [['dcterms:' + doctime, {'xsi:type': 'dcterms:W3CDTF'}, value]])

        return make_element_tree(coreprops_tree)

    def appproperties(self, company):
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
                         [['Company', company]]
                         ]

        return make_element_tree(appprops_tree, nsprefixes['ep'])

    def websettings(self):
        '''
          Generate websettings
          This function copied from 'python-docx' library
        '''
        web_tree = [['w:webSettings'], [['w:allowPNG']],
                    [['w:doNotSaveAsSingleFile']]]
        return make_element_tree(web_tree)

    def relationshiplist(self):
        filename = 'word/_rels/document.xml.rels'
        relationships = self.styleDocx.get_xmltree(filename)

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

        return make_element_tree(rel_tree, nsprefixes['pr'])
