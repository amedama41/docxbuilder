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
import io
import os
import posixpath
import re
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
    # Variant Types
    'vt': 'http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes',
    # xml
    'xml': 'http://www.w3.org/XML/1998/namespace'
}

REL_TYPE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
REL_TYPE_APP = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
REL_TYPE_CORE = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
REL_TYPE_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
REL_TYPE_NUMBERING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering'
REL_TYPE_FOOTNOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'
REL_TYPE_CUSTOM = 'http://purl.oclc.org/ooxml/officeDocument/relationships/customProperties'

REL_TYPE_COMMENTS = 'http://purl.oclc.org/ooxml/officeDocument/relationships/comments'
REL_TYPE_ENDNOTES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes'
REL_TYPE_FONT_TABLE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable'
REL_TYPE_GLOSSARY_DOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument'
REL_TYPE_SETTINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings'
REL_TYPE_STYLES_WITH_EFFECTS = 'http://schemas.microsoft.com/office/2007/relationships/stylesWithEffects'
REL_TYPE_THEME = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
REL_TYPE_WEB_SETTINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings'

REL_TYPE_CUSTOM_XML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
REL_TYPE_CUSTOM_XML_PROPS = 'http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps'
REL_TYPE_THUMBNAIL = 'http://purl.oclc.org/ooxml/officeDocument/relationships/metadata/thumbnail'


CONTENT_TYPE_DOC_MAIN = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml'
CONTENT_TYPE_STYLES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
CONTENT_TYPE_NUMBERING = 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
CONTENT_TYPE_FOOTNOTES = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'
CONTENT_TYPE_CORE_PROPERTIES = 'application/vnd.openxmlformats-package.core-properties+xml'
CONTENT_TYPE_EXTENDED_PROPERTIES = 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
CONTENT_TYPE_CUSTOM_PROPERTIES = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'

#####################

def xml_encode(value):
    value = re.sub(r'_(?=x[0-9a-fA-F]{4}_)', r'_x005f_', value)
    return re.sub(r'[\x00-\x1f]', lambda m: '_x%04x_' % ord(m.group(0)), value)

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

def get_max_attribute(elems, attribute, to_int=int):
    '''
       Get the maximum integer attribute among the specified elems
    '''
    if not elems:
        return 0
    return max(map(lambda e: to_int(e.get(attribute)), elems))

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

CORE_PROPERTY_KEYS = (
        ('dc', 'title', {}),
        ('dc', 'creator', {}),
        ('dc', 'language', {}),
        ('cp', 'category', {}),
        ('cp', 'contentStatus', {}),
        ('dc', 'description', {}),
        ('dc', 'identifier', {}),
        ('cp', 'lastModifiedBy', {}),
        ('cp', 'lastPrinted', {}),
        ('cp', 'revision', {}),
        ('dc', 'subject', {}),
        ('cp', 'version', {}),
        ('cp', 'keywords', {}),
        ('dcterms', 'created', {'xsi:type': 'dcterms:W3CDTF'}),
        ('dcterms', 'modified', {'xsi:type': 'dcterms:W3CDTF'}),
)

CUSTOM_PROPERTY_TYPES = (
        (bool, 'vt:bool', lambda v: str(v).lower()),
        (six.integer_types, 'vt:i8', str),
        (float, 'vt:r8', str),
        (six.string_types, 'vt:lpwstr', str),
        (datetime.datetime, 'vt:date', convert_to_W3CDTF_string),
)

def separate_core_and_custom_properties(props):
    core_props = {}
    custom_props = {}
    invalid_prop_keys = []

    core_prop_keys = set(key for _, key, _ in CORE_PROPERTY_KEYS)

    for key, value in props.items():
        if key in core_prop_keys:
            core_props[key] = value
            continue
        for prop_type, _, _ in CUSTOM_PROPERTY_TYPES:
            if isinstance(value, prop_type):
                custom_props[key] = value
                break
        else:
            invalid_prop_keys.append(key)

    time_fmt = '%Y-%m-%dT%H:%M:%S'
    last_printed = core_props.get('lastPrinted', None)
    if last_printed is not None:
        if isinstance(last_printed, datetime.datetime):
            core_props['lastPrinted'] = last_printed.strftime(time_fmt)
        else:
            try:
                datetime.datetime.strptime(last_printed, time_fmt)
            except ValueError:
                invalid_prop_keys.append('lastPrinted')
                del core_props['lastPrinted']

    for doctime in ['created', 'modified']:
        value = core_props.get(doctime, None)
        if value is None:
            continue
        value = convert_to_W3CDTF_string(value)
        if value is None:
            invalid_prop_keys.append(doctime)
            del core_props[doctime]
        else:
            core_props[doctime] = value

    return core_props, custom_props, invalid_prop_keys

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
    width = get_contents_size(section_property, 'w:w', ('w:left', 'w:right'))
    cols_elems = get_elements(section_property, 'w:cols')
    if not cols_elems:
        return width
    cols = cols_elems[-1]
    num = int(cols.attrib.get(norm_name('w:num'), '1'))
    space = int(cols.attrib.get(norm_name('w:space'), '720'))
    return (width - (space * (num - 1))) // num # TODO non equal col width

def get_contents_height(section_property):
    return get_contents_size(section_property, 'w:h', ('w:top', 'w:bottom'))

def get_contents_size(section_property, size_prop, margin_props):
    paper_size = get_elements(section_property, 'w:pgSz')[0]
    size = int(paper_size.get(norm_name(size_prop)))
    paper_margin = get_elements(section_property, 'w:pgMar')[0]
    margin = (
            int(paper_margin.get(norm_name(margin_props[0]))) +
            int(paper_margin.get(norm_name(margin_props[1]))))
    return size - margin

def make_default_page_size():
    return make_element_tree([['w:pgSz', {
        'w:w': '12240', 'w:h': '15840', 'w:orient': 'portrait',
    }]])

def make_default_page_margin():
    return make_element_tree([['w:pgMar', {
        'w:top': '1440', 'w:right': '1440',
        'w:bottom': '1440', 'w:left': '1440',
        'w:header': '720', 'w:footer': '720', 'w:gutter': '0',
    }]])

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
        indent, right_indent, style, align, keep_lines, keep_next, list_info,
        properties=None):
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
    if properties is not None:
        style_tree.extend(properties)

    paragraph_tree = [['w:p'], style_tree]
    return make_element_tree(paragraph_tree)

def make_paragraph_spacing_property(**kwargs):
    attr = {}
    for key in ['before', 'after', 'line']:
        value = kwargs.get(key)
        if value is None:
            continue
        attr['w:' + key] = str(value)
    return [['w:spacing', attr]]

def make_paragraph_shading_property(pattern, **kwargs):
    attr = {'w:val': pattern}
    if kwargs:
        for key in ['color', 'fill']:
            value = kwargs.get(key)
            if value is None:
                continue
            attr['w:' + key] = value
    return [['w:shd', attr]]

def make_paragraph_border_property(**kwargs):
    key_list = [
            ('size', 'w:sz'), ('space', 'w:space'), ('color', 'w:color'),
            ('shadow', 'w:shadow'), ('frame', 'w:frame')
    ]
    border_tree = [['w:pBdr']]
    for kind in ['top', 'left', 'bottom', 'right', 'between', 'bar']:
        if kind not in kwargs:
            continue
        value = kwargs[kind]
        if value is None:
            attr = {'w:val': 'nil'}
        else:
            pattern = value['pattern']
            attr = {'w:val': pattern if pattern is not None else 'nil'}
            for key, attr_key in key_list:
                v = value.get(key)
                if v is not None:
                    attr[attr_key] = str(v)
        border_tree.append([['w:' + kind, attr]])
    return border_tree

def make_border_info(border_attrs):
    identity = lambda x: x
    to_bool = lambda x: x == 'true' or x == '1'
    attr_list = [
            ('w:val', 'pattern', identity), ('w:color', 'color', identity),
            ('w:sz', 'size', int), ('w:space', 'space', int),
            ('w:shadow', 'shadow', to_bool), ('w:frame', 'frame', to_bool)
    ]
    border_info = {}
    for attr, key, convert in attr_list:
        val = border_attrs.get(norm_name(attr))
        if val is not None:
            border_info[key] = convert(val)
    return border_info

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

def make_table(
        style, width, indent, align, grid_col_list, has_head, has_first_column,
        properties=None):
    look_attrs = {
            'w:noHBand': 'false', 'w:noVBand': 'false',
            'w:lastRow': 'false', 'w:lastColumn': 'false'
    }
    look_attrs['w:firstRow'] = 'true' if has_head else 'false'
    look_attrs['w:firstColumn'] = 'true' if has_first_column else 'false'
    if width is not None:
        width_attr = {'w:w': '%f%%' % (width * 100), 'w:type': 'pct'}
    else:
        width_attr = {'w:w': '0', 'w:type': 'auto'}
    property_tree = [
            ['w:tblPr'],
            [['w:tblW', width_attr]],
            [['w:tblInd', {'w:w': str(indent), 'w:type': 'dxa'}]],
            [['w:tblLook', look_attrs]],
    ]
    if style is not None:
        property_tree.insert(1, [['w:tblStyle', {'w:val': style}]])
    if align is not None:
        property_tree.append([['w:jc', {'w:val': align}]])
    if properties is not None:
        property_tree.extend(properties)

    table_grid_tree = [['w:tblGrid']]
    for grid_col in grid_col_list:
        table_grid_tree.append([['w:gridCol', {'w:w': str(int(grid_col))}]])

    table_tree = [
            ['w:tbl'],
            property_tree,
            table_grid_tree
    ]
    return make_element_tree(table_tree)

def make_row(index, is_head, cant_split, set_tbl_header):
    row_style_attrs = {
            'w:evenHBand': ('true' if index % 2 != 0 else 'false'),
            'w:oddHBand': ('true' if index % 2 == 0 else 'false'),
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

def make_cell(index, is_first_column, cellsize, grid_span, vmerge, valign=None):
    cell_style = {
            'w:evenVBand': ('true' if index % 2 != 0 else 'false'),
            'w:oddVBand': ('true' if index % 2 == 0 else 'false'),
            'w:firstColumn': ('true' if is_first_column else 'false'),
    }
    property_tree = [
            ['w:tcPr'],
            [['w:cnfStyle', cell_style]],
    ]
    if cellsize is not None:
        property_tree.append(
                [['w:tcW', {'w:w': '%f%%' % (cellsize * 100), 'w:type': 'pct'}]])
    if grid_span > 1:
        property_tree.append([['w:gridSpan', {'w:val': str(grid_span)}]])
    if vmerge is not None:
        property_tree.append([['w:vMerge', {'w:val': vmerge}]])
    if valign is not None:
        property_tree.append([['w:vAlign', {'w:val': valign}]])
    return make_element_tree([['w:tc'], property_tree])

def make_table_cell_margin_property(**kwargs):
    margin_tree = [['w:tblCellMar']]
    for kind in ['top', 'left', 'bottom', 'right']:
        if kind in kwargs:
            margin_tree.append(
                    [['w:' + kind, make_table_width_attr(kwargs[kind])]])
    return margin_tree

def make_table_cell_spacing_property(val):
    return [['w:tblCellSpacing', make_table_width_attr(val)]]

def make_table_width_attr(val):
    if val is None:
        return {'w:type': 'nil', 'w:w': '0'}
    elif val == 'auto':
        return {'w:type': 'auto', 'w:w': '0'}
    elif isinstance(val, float) and val <= 1.0:
        return {'w:type': 'pct', 'w:w': '%f%%' % (val * 100)}
    else:
        return {'w:type': 'dxa', 'w:w': str(int(val))}

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

def make_table_of_contents(
        toc_title, style_id, maxlevel, bookmark, paragraph_width, outlines):
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
    tab_pos = str(paragraph_width - 10)
    tabs_tree = [
            ['w:tabs'],
            [['w:tab', {'w:val': 'right', 'w:leader': 'dot', 'w:pos': tab_pos}]]
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

def create_rels_path(path):
    return posixpath.join(
            posixpath.dirname(path), '_rels',
            posixpath.basename(path) + '.rels')

def make_relationships(relationships):
    '''Generate a relationships
    '''
    rel_tree = [['Relationships']]
    for attributes in relationships:
        rel_tree.append([['Relationship', attributes]])
    return make_element_tree(rel_tree, nsprefixes['pr'])


class StyleInfo(object):
    style_id_attr = norm_name('w:styleId')
    type_attr = norm_name('w:type')

    def __init__(self, style):
        self._style = style
        if get_elements(style, 'w:unhideWhenUsed'):
            self._semihidden_elems = get_elements(style, 'w:semiHidden')
        else:
            self._semihidden_elems = []

    @property
    def style_id(self):
        return self._style.attrib[type(self).style_id_attr]

    @property
    def style_type(self):
        return self._style.attrib[type(self).type_attr]

    def get_based_style_id(self):
        based_on_elems = get_elements(self._style, 'w:basedOn')
        if not based_on_elems:
            return None
        return based_on_elems[-1].attrib[norm_name('w:val')]

    def get_border_info(self, kind):
        border_elems = get_elements(self._style, 'w:pPr/w:pBdr/w:' + kind)
        if not border_elems:
            return None
        return border_elems[-1].attrib

    def get_run_style_property(self):
        props = get_elements(self._style, 'w:rPr')
        if not props:
            return {}
        return [(prop.tag, prop.attrib)
                for prop in props[0] if not prop.tag.endswith('rPrChange')]

    def get_table_horizon_margin(self):
        cell_margin_elems = get_elements(self._style, 'w:tblPr/w:tblCellMar')
        if not cell_margin_elems:
            return (None, None)

        cell_margin = cell_margin_elems[-1]
        type_attr = norm_name('w:type')
        w_attr = norm_name('w:w')
        def get_margin(elem):
            if elem is None or elem.get(type_attr) != 'dxa':
                return None
            return int(elem.get(w_attr))
        left = cell_margin.find('w:left', nsprefixes)
        right = cell_margin.find('w:right', nsprefixes)
        return (get_margin(left), get_margin(right))

    def used(self):
        for semihidden in self._semihidden_elems:
            self._style.remove(semihidden)
        self._semihidden_elems = []


class DocxDocument:
    def __init__(self, docxfile):
        '''
          Constructor
        '''
        self.docx = zipfile.ZipFile(docxfile)
        docpath = get_attribute(
                self.get_xmltree('_rels/.rels'),
                'pr:Relationship[@Type="%s"]' % REL_TYPE_DOC, 'Target')
        if docpath.startswith('/'):
            docpath = docpath[1:]

        self.docpath = docpath
        self.document = self.get_xmltree(docpath)
        self.relationships = self.get_xmltree(create_rels_path(docpath))
        self.numbering = self.get_xmltree('word/numbering.xml')
        self.styles = self.get_xmltree('word/styles.xml')

    @property
    def footnotes(self):
        return self.get_xmltree('word/footnotes.xml')

    @property
    def numbering_relationships(self):
        return self.get_xmltree(create_rels_path('word/numbering.xml'))

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
        val_attr = norm_name('w:val')
        def get_info(style):
            info = StyleInfo(style)
            names = get_elements(style, 'w:name')
            style_name = names[0].attrib[val_attr] if names else info.style_id
            return (style_name, info)
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

    def get_relationship_ids(self):
        rids = []
        id_attr = norm_name('Id')
        for rel in get_elements(self.relationships, 'pr:Relationship'):
            m = re.match(r'rId(\d+)', rel.get(id_attr))
            if m is not None:
                rids.append(int(m.group(1)))
        return rids

    def get_image_numbers(self):
        img_nums = []
        for path in self.docx.namelist():
            m = re.match(r'word/media/image(\d+)\.\w+', path)
            if m is not None:
                img_nums.append(int(m.group(1)))
        return img_nums

    def collect_items(self, zip_docxfile, collected_files):
        # Add & compress support files
        for fname in collected_files:
            zip_docxfile.writestr(fname, self.docx.read(fname))

    def collect_relation_files(self, rel_files, rel_attrs, basedir):
        for attr in rel_attrs:
            if attr.get('TargetMode', 'Internal') == 'External':
                continue
            filepath = posixpath.normpath(
                    posixpath.join(basedir, attr['Target']))
            if filepath.startswith('/'):
                filepath = filepath[1:]
            rel_files.add(filepath)
            rel_filepath = create_rels_path(filepath)
            rel_xml = self.get_xmltree(rel_filepath)
            if rel_xml is not None:
                rel_files.add(rel_filepath)
                self.collect_relation_files(
                        rel_files,
                        (r.attrib for r in get_elements(
                            rel_xml, '/pr:Relationships/pr:Relationship')),
                        posixpath.dirname(filepath))

    def collect_all_relation_files(self, rel_attrs):
        rel_files = set()
        self.collect_relation_files(
                rel_files, rel_attrs, posixpath.dirname(self.docpath))
        return rel_files

    def collect_num_ids(self, rel_attrs):
        """Collect num id used in files referenced by rel_attrs
        """
        num_ids = set()
        val_attr = norm_name('w:val')
        for attr in rel_attrs:
            if attr.get('TargetMode', 'Internal') == 'External':
                continue
            filepath = posixpath.normpath(
                    posixpath.join(self.docpath, attr['Target']))
            if filepath.startswith('/'):
                filepath = filepath[1:]
            xml = self.get_xmltree(filepath)
            if xml is None:
                continue
            num_id_elems = get_elements(xml, '//w:numId')
            num_ids.update(
                    [int(num_id.get(val_attr)) for num_id in num_id_elems])
        return num_ids

############
# Numbering
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

##########

class IdPool(object):
    def __init__(self, used_ids=[], init_id=1):
        self._used_ids = set(used_ids)
        self._next_id = init_id

    def next_id(self):
        while self._next_id in self._used_ids:
            self._next_id += 1
        next_id = self._next_id
        self._next_id += 1
        return next_id

class IdElements(object):
    def __init__(self, elems, attr, to_int=int, init_id=0):
        self._next_id = init_id
        self._elems = dict((to_int(elem.get(attr)), elem) for elem in elems)
        self._attr = attr
        self._to_int = to_int

    def next_id(self):
        while self._next_id in self._elems:
            self._next_id += 1
        next_id = self._next_id
        self._next_id += 1
        return next_id

    def append(self, elem):
        self._elems[self._to_int(elem.get(self._attr))] = elem

    def get(self, key, default=None):
        return self._elems.get(self._to_int(key), default)

    def __iter__(self):
        return iter(self._elems.items())

def collect_used_rel_attrs(relationships, xml, used_rel_types=set()):
    if relationships is None:
        return []
    used_rel_attrs = []
    for rel in get_elements(relationships, 'pr:Relationship'):
        if (rel.get('Type') in used_rel_types
                or get_elements(xml, '(.//*[@*="%s"])[1]' % rel.get('Id'))):
            used_rel_attrs.append(rel.attrib)
    return used_rel_attrs
#
# DocxComposer Class
#


class DocxComposer:
    def __init__(self, stylefile, has_coverpage):
        '''
           Constructor
        '''
        self._id = 100
        self.styleDocx = DocxDocument(stylefile)

        self._style_info = self.styleDocx.extract_style_info()
        self._abstract_nums = IdElements(
                self.styleDocx.get_elems_from_numbering('w:abstractNum'),
                norm_name('w:abstractNumId'))
        self._nums = IdElements(
                self.styleDocx.get_elems_from_numbering('w:num'),
                norm_name('w:numId'), init_id=1)

        # document part -> (relationships, relationship id pool)
        self._relationships_map = {
                'document': ([], IdPool(self.styleDocx.get_relationship_ids())),
                'footnotes': ([], IdPool([])),
        }
        self._add_required_relationships()
        self._hyperlink_rid_map = {} # target => relationship id
        self._image_info_map = {} # imagepath => (relationship id, imagename)
        self._img_num_pool = IdPool(self.styleDocx.get_image_numbers())

        self._footnotes = make_element_tree([['w:footnotes']])
        self._footnotes.extend(get_special_footnotes(self.styleDocx.footnotes))
        self._footnote_id_map = {} # docname#id => footnote id
        self._max_footnote_id = get_max_attribute(
                get_elements(self._footnotes, 'w:footnote'), norm_name('w:id'))

        self._run_style_property_cache = {}
        self._table_margin_cache = {}

        self.document = make_element_tree([['w:document'], [['w:body']]])
        self.docbody = get_elements(self.document, '/w:document/w:body')[0]

        coverpage = self.styleDocx.get_coverpage() if has_coverpage else None
        if coverpage is not None:
            self.docbody.append(coverpage)

    def new_id(self):
        self._id += 1
        return self._id

    def get_each_orient_section_properties(self):
        section_props = self.styleDocx.get_section_properties()
        if not section_props:
            section_props = [make_element_tree([['w:sectPr']])]
        first = section_props[0]
        if not get_elements(first, 'w:pgSz'):
            first.append(make_default_page_size())
        if not get_elements(first, 'w:pgMar'):
            first.append(make_default_page_margin())
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

    def get_style_info_from_id(self, style_id):
        for style_info in self._style_info.values():
            if style_info.style_id == style_id:
                return style_info
        return None

    def get_style_id(self, style_name):
        style_info = self.get_style_info(style_name)
        if style_info is None:
            return style_name
        style_info.used()
        return style_info.style_id

    def get_indent(self, style_name, default):
        style_info = self.get_style_info(style_name)
        if style_info is None or style_info.style_type != 'paragraph':
            return default
        indent = self.styleDocx.get_indent(style_info.style_id)
        if indent is None:
            return default
        return int(indent)

    def get_border_info(self, style_id, kind):
        if style_id is None:
            return None
        style_info = self.get_style_info_from_id(style_id)
        if style_info is None or style_info.style_type != 'paragraph':
            return None
        border_attrs = style_info.get_border_info(kind)
        if border_attrs is not None:
            return make_border_info(border_attrs)
        return self.get_border_info(style_info.get_based_style_id(), kind)

    def get_run_style_property(self, style_id):
        if style_id is None:
            return {}
        style_prop = self._run_style_property_cache.get(style_id)
        if style_prop is not None:
            return style_prop
        style_info = self.get_style_info_from_id(style_id)
        if style_info is None or style_info.style_type != 'character':
            return self._run_style_property_cache.setdefault(style_id, {})
        based_style_id = style_info.get_based_style_id()
        style_prop = {}
        if based_style_id is not None:
            style_prop.update(self.get_run_style_property(based_style_id))
        style_prop.update(style_info.get_run_style_property())
        return self._run_style_property_cache.setdefault(style_id, style_prop)

    def get_bullet_list_num_id(self, style_name):
        return self.styleDocx.get_numbering_style_id(style_name)

    def get_table_cell_margin(self, style_id):
        misc_margin = 8 * 2 * 10 # Miscellaneous margin (e.g. border width)
        left, right = self.get_table_horizon_margin(style_id)
        margin = left + right + misc_margin
        return margin

    def get_table_horizon_margin(self, style_id):
        default_margin = (115, 115)
        if style_id is None:
            return default_margin
        margin = self._table_margin_cache.get(style_id)
        if margin is not None:
            return margin

        style_info = self.get_style_info_from_id(style_id)
        if style_info is None or style_info.style_type != 'table':
            return self._table_margin_cache.setdefault(style_id, default_margin)
        left, right = style_info.get_table_horizon_margin()
        if left is None or right is None:
            based_left, based_right = self.get_table_horizon_margin(
                    style_info.get_based_style_id())
            left = left or based_left
            right = right or based_right
        return self._table_margin_cache.setdefault(style_id, (left, right))

    def asbytes(self, props, custom_props):
        '''Generate the composed document as docx binary.
        '''
        xml_files = [
                ('_rels/.rels', self.make_root_rels()),
                ('docProps/app.xml', self.make_app(custom_props)),
                ('docProps/core.xml', self.make_core(props)),
                ('docProps/custom.xml', self.make_custom(custom_props)),
        ]

        inherited_rel_attrs = self.collect_inherited_rel_attrs()
        numbering = self.make_numbering(inherited_rel_attrs)

        document_rels = self.make_document_rels(inherited_rel_attrs)
        xml_files.append(('word/_rels/document.xml.rels', document_rels))
        footnotes_rels = self.make_footnotes_rels()
        if footnotes_rels is not None:
            xml_files.append(('word/_rels/footnotes.xml.rels', footnotes_rels))
        numbering_rel_attrs = collect_used_rel_attrs(
                self.styleDocx.numbering_relationships, numbering)
        if numbering_rel_attrs:
            numbering_rels = self.make_numbering_rels(numbering_rel_attrs)
            xml_files.append(('word/_rels/numbering.xml.rels', numbering_rels))

        xml_files.append(('word/document.xml', self.document))
        xml_files.append(('word/footnotes.xml', self._footnotes))
        xml_files.append(('word/numbering.xml', numbering))
        xml_files.append(('word/styles.xml', self.styleDocx.styles))

        inherited_files = self.styleDocx.collect_all_relation_files(
                inherited_rel_attrs + numbering_rel_attrs)
        content_types = self.make_content_types(inherited_files)
        xml_files.append(('[Content_Types].xml', content_types))

        bytes_io = io.BytesIO()
        with zipfile.ZipFile(
                bytes_io, mode='w', compression=zipfile.ZIP_DEFLATED) as zip:
            self.styleDocx.collect_items(zip, inherited_files)
            for xmlpath, xml in xml_files:
                treestring = etree.tostring(
                    xml, xml_declaration=True,
                    encoding='UTF-8', standalone='yes')
                zip.writestr(xmlpath, treestring)
            for imgpath, (_, picname) in self._image_info_map.items():
                zip.write(imgpath, 'word/media/' + picname)

        return bytes_io.getvalue()


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

        num = self._nums.get(num_id)
        if num is None:
            return []

        abst_num_id = get_attribute(num, 'w:abstractNumId', 'w:val')
        abstract_num = self._abstract_nums.get(abst_num_id)
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
        abstract_num_id = self._abstract_nums.next_id()
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

        num_id = self._nums.next_id()
        num_tree = [
                ['w:num', {'w:numId': str(num_id)}],
                [['w:abstractNumId', {'w:val': str(abstract_num_id)}]],
        ]
        num = make_element_tree(num_tree)
        self._nums.append(num)
        return num_id

    def get_default_style_names(self):
        '''
           Return default paragraph, character, table style ids
        '''
        paragraph_style_id = self.styleDocx.get_default_style_name('paragraph')
        character_style_id = self.styleDocx.get_default_style_name('character')
        table_style_id = self.styleDocx.get_default_style_name('table')
        return paragraph_style_id, character_style_id, table_style_id

    def create_style(
            self, style_type, new_style_name, based_style_name, is_custom,
            is_hidden=False):
        '''
           Create a new style_stype style with new_style_id,
           which is based on based_style_id.
        '''
        return self._create_style(
                style_type, new_style_name, is_custom, is_hidden,
                based_style_name=based_style_name)

    def create_list_style(
            self, new_style_name, format_type, lvl_text, font, indent):
        def make_property_tree(new_style_id):
            num_id = self.add_numbering_style(
                    1, lvl_text, format_type, indent, new_style_id, font)
            return [
                    ['w:pPr'],
                    [['w:numPr'],
                        [['w:ilvl', {'w:val': '0'}]],
                        [['w:numId', {'w:val': str(num_id)}]],
                    ],
            ]
        is_custom = False
        is_hidden = False
        return self._create_style(
                'paragraph', new_style_name, is_custom, is_hidden,
                make_property_tree=make_property_tree)

    def create_empty_paragraph_style(
            self, new_style_name, after_space, with_border, is_hidden):
        '''
           Create a new empty paragraph style
        '''
        def make_property_tree(_):
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
                    [['w:bottom', {
                        'w:val': 'single', 'w:sz': '8', 'w:space': '1'
                    }]]
                ])
            return property_tree
        is_custom = True
        return self._create_style(
                'paragraph', new_style_name, is_custom, is_hidden,
                make_property_tree=make_property_tree)

    def _create_style(
            self, style_type, new_style_name, is_custom, is_hidden,
            based_style_name=None, make_property_tree=None):
        if self.get_style_info(new_style_name) is not None:
            return False
        new_style_id = new_style_name
        style_tree = [
                ['w:style', {
                    'w:type': style_type,
                    'w:customStye': '1' if is_custom else '0',
                    'w:styleId': new_style_id
                }],
                [['w:name', {'w:val': new_style_name}]],
                [['w:semiHidden']],
                [['w:qFormat']],
        ]
        if not is_hidden:
            style_tree.append([['w:unhideWhenUsed']])
        if based_style_name is not None:
            based_info = self.get_style_info(based_style_name)
            if based_info is not None and based_info.style_type == style_type:
                style_tree.append(
                        [['w:basedOn', {'w:val': based_info.style_id}]])
        if make_property_tree is not None:
            style_tree.append(make_property_tree(new_style_id))
        new_style = make_element_tree(style_tree)
        self.styleDocx.styles.append(new_style)
        self._style_info[new_style_name] = StyleInfo(new_style)
        return True

    def _add_required_relationships(self):
        relationships, id_pool = self._relationships_map['document']
        required_rel_types = (
                (REL_TYPE_STYLES, 'styles.xml'),
                (REL_TYPE_NUMBERING, 'numbering.xml'),
                (REL_TYPE_FOOTNOTES, 'footnotes.xml'),
        )
        for rel_type, target in required_rel_types:
            relationships.append({
                'Id': 'rId%d' % id_pool.next_id(),
                'Type': rel_type, 'Target': target
            })

    def add_hyperlink_relationship(self, target, part):
        rid_map = self._hyperlink_rid_map.get(target)
        if rid_map is not None:
            rid = rid_map.get(part, None)
            if rid is not None:
                return rid
        else:
            rid_map = {}

        relationships, id_pool = self._relationships_map[part]
        rid = 'rId%d' % id_pool.next_id()
        relationships.append({
            'Id': rid,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
            'Target': target,
            'TargetMode': 'External'
        })
        rid_map[part] = rid
        self._hyperlink_rid_map[target] = rid_map
        return rid

    def add_image_relationship(self, imagepath, part):
        imagepath = os.path.abspath(imagepath)

        rid_map, picname = self._image_info_map.get(imagepath, (None, None))
        if rid_map is not None:
            rid = rid_map.get(part, None)
            if rid is not None:
                return rid
        else:
            _, picext = os.path.splitext(imagepath)
            if picext == '.jpg':
                picext = '.jpeg'
            rid_map = {}
            picname = 'image%d%s' % (self._img_num_pool.next_id(), picext)

        relationships, id_pool = self._relationships_map[part]
        rid = 'rId%d' % id_pool.next_id()
        relationships.append({
            'Id': rid,
            'Type': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            'Target': 'media/' + picname
        })
        rid_map[part] = rid
        self._image_info_map[imagepath] = (rid_map, picname)
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
        self._footnotes.append(footnote)

    def collect_inherited_rel_attrs(self):
        """Collect relationships inherited from style file.
        """
        implicit_rel_types = {
                REL_TYPE_COMMENTS,
                REL_TYPE_ENDNOTES,
                REL_TYPE_FONT_TABLE,
                REL_TYPE_GLOSSARY_DOCUMENT,
                REL_TYPE_SETTINGS,
                REL_TYPE_STYLES_WITH_EFFECTS,
                REL_TYPE_THEME,
                REL_TYPE_WEB_SETTINGS,
                REL_TYPE_CUSTOM_XML,
                REL_TYPE_CUSTOM_XML_PROPS,
                REL_TYPE_THUMBNAIL,
        }
        return collect_used_rel_attrs(
                self.styleDocx.relationships, self.document, implicit_rel_types)

    def make_content_types(self, inherited_files):
        '''create [Content_Types].xml
        '''
        filename = '[Content_Types].xml'
        content_types = self.styleDocx.get_xmltree(filename)

        types_tree = [['Types']]
        # Add support for filetypes
        filetypes = {
                'rels': 'application/vnd.openxmlformats-package.relationships+xml',
                'xml': 'application/xml',
                'jpeg': 'image/jpeg',
                'jpg': 'image/jpeg',
                'gif': 'image/gif',
                'png': 'image/png',
                'emf': 'image/x-emf',
        }
        for ext, ctype in filetypes.items():
            types_tree.append(
                    [['Default', {'Extension': ext, 'ContentType': ctype}]])
        for elem in get_elements(content_types, '/ct:Types/ct:Default'):
            ext = elem.attrib['Extension']
            if ext in filetypes:
                continue
            types_tree.append([['Default', {
                'Extension': ext, 'ContentType': elem.attrib['ContentType']
            }]])

        required_content_types = [
                ('/docProps/core.xml', CONTENT_TYPE_CORE_PROPERTIES),
                ('/docProps/app.xml', CONTENT_TYPE_EXTENDED_PROPERTIES),
                ('/docProps/custom.xml', CONTENT_TYPE_CUSTOM_PROPERTIES),
                ('/word/document.xml', CONTENT_TYPE_DOC_MAIN),
                ('/word/styles.xml', CONTENT_TYPE_STYLES),
                ('/word/numbering.xml', CONTENT_TYPE_NUMBERING),
                ('/word/footnotes.xml', CONTENT_TYPE_FOOTNOTES),
        ]
        for name, ctype in required_content_types:
            types_tree.append([['Override', {
                'PartName': name, 'ContentType': ctype,
            }]])
        for elem in get_elements(content_types, '/ct:Types/ct:Override'):
            name = elem.attrib['PartName']
            if name[1:] not in inherited_files:
                continue
            types_tree.append([['Override', {
                'PartName': name, 'ContentType': elem.attrib['ContentType']
            }]])

        return make_element_tree(types_tree, nsprefixes['ct'])

    def make_core(self, props):
        '''Create core properties (common document properties referred to in
           the 'Dublin Core' specification).
        '''
        coreprops_tree = [['cp:coreProperties']]
        for ns, prop, attr in CORE_PROPERTY_KEYS:
            value = props.get(prop, None)
            if value is None:
                continue
            if isinstance(value, (list, tuple)):
                if prop == 'keywords':
                    value = ','.join(value)
                else:
                    value = '; '.join(value)
            value = xml_encode(value)
            coreprops_tree.append([['%s:%s' % (ns, prop), attr, value]])

        return make_element_tree(coreprops_tree)

    def make_app(self, custom_props):
        '''Create app-specific properties.
           This function is based on 'python-docx' library
        '''
        appprops_tree = [
                ['Properties'],
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
        ]
        for key in ['Company', 'Manager']:
            value = custom_props.get(key)
            if value is None:
                value = custom_props.get(key.lower())
                if value is None:
                    continue
            appprops_tree.append([[key, xml_encode(value)]])

        return make_element_tree(appprops_tree, nsprefixes['ep'])

    def make_custom(self, custom_props):
        props_tree = [['Properties']]
        # User defined pid must start from 2
        for pid, (name, value) in enumerate(custom_props.items(), 2):
            for type_, type_elem, to_str in CUSTOM_PROPERTY_TYPES:
                if not isinstance(value, type_):
                    continue
                # Fmtid of user defined properties
                fmtid = '{D5CDD505-2E9C-101B-9397-08002B2CF9AE}'
                attr = {'pid': str(pid), 'fmtid': fmtid, 'name': name}
                props_tree.append(
                        [['property', attr], [[type_elem, to_str(value)]]])
                break
        xmlns = 'http://purl.oclc.org/ooxml/officeDocument/customProperties'
        return make_element_tree(props_tree, xmlns)

    def make_document_rels(self, stylerels):
        rel_list = []
        docrel_list, _ = self._relationships_map['document']
        rel_list.extend(docrel_list)
        rel_list.extend(stylerels)
        return make_relationships(rel_list)

    def make_footnotes_rels(self):
        rel_list, _ = self._relationships_map['footnotes']
        if not rel_list:
            return None
        return make_relationships(rel_list)

    def make_numbering_rels(self, numbering_rel_attrs):
        return make_relationships(numbering_rel_attrs)

    def make_root_rels(self):
        rel_list = [
                (REL_TYPE_CORE, 'docProps/core.xml'),
                (REL_TYPE_APP, 'docProps/app.xml'),
                (REL_TYPE_CUSTOM, 'docProps/custom.xml'),
                (REL_TYPE_DOC, 'word/document.xml'),
        ]
        return make_relationships(
                {'Id': 'rId%d' % rid, 'Type': rtype, 'Target': target}
                for rid, (rtype, target) in enumerate(rel_list, 1))

    def make_numbering(self, inherited_rel_attrs):
        """Create numbering.xml from nums and abstract nums in use
        """
        used_num_ids = self.styleDocx.collect_num_ids(inherited_rel_attrs)
        val_attr = norm_name('w:val')
        def update_used_num_ids(xml):
            elems = get_elements(xml, '//w:numId')
            used_num_ids.update((int(num_id.get(val_attr)) for num_id in elems))
        update_used_num_ids(self.styleDocx.styles)
        update_used_num_ids(self.document)
        update_used_num_ids(self._footnotes)

        nums = [num for num_id, num in self._nums if num_id in used_num_ids]
        get_abst_num_id = lambda num: int(
                get_elements(num, 'w:abstractNumId')[-1].get(val_attr))
        abst_num_ids = set((get_abst_num_id(num) for num in nums))
        abstract_nums = [
                abst_num for abst_num_id, abst_num in self._abstract_nums
                if abst_num_id in abst_num_ids
        ]

        numbering = make_element_tree(['w:numbering'])
        numbering.extend(abstract_nums)
        numbering.extend(nums)

        for elem in get_elements(numbering, '//w:lvlPicBulletId'):
            bullet_id = elem.get(val_attr)
            num_pic_bullet_elems = get_elements(
                    self.styleDocx.numbering,
                    'w:numPicBullet[@w:numPicBulletId="%s"]' % bullet_id)
            if num_pic_bullet_elems:
                numbering.insert(0, num_pic_bullet_elems[-1])
        return numbering
