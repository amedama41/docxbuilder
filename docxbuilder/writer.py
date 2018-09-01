# -*- coding: utf-8 -*-
"""
    sphinx-docxwriter
    ~~~~~~~~~~~~~~~~~~~~~~~~~~

    Modified custom docutils writer for OpenXML (docx).
    Original code from 'sphinxcontrib-documentwriter'

    :copyright:
        Copyright 2011 by haraisao at gmail dot com 
    :license: MIT, see LICENSE for details.
"""
"""
    sphinxcontrib-docxwriter
    ~~~~~~~~~~~~~~~~~~~~~~~~~~

    Custom docutils writer for OpenXML (docx).

    :copyright:
        Copyright 2010 by shimizukawa at gmail dot com (Sphinx-users.jp).
    :license: BSD, see LICENSE for details.
"""

import itertools
import re

from docutils import nodes, writers

from sphinx import addnodes
from sphinx import highlighting
from sphinx.locale import admonitionlabels, versionlabels, _

from sphinx.ext import graphviz

import docxbuilder.docx as docx
import sys
import os
import six
from lxml import etree
from docxbuilder.highlight import DocxPygmentsBridge


#
# Is the PIL imaging library installed?
try:
    from PIL import Image
except ImportError as exp:
    Image = None

#
#  Logging for debugging
#
import logging
logging.basicConfig(filename='docx.log', filemode='w', level=logging.INFO,
                    format="%(asctime)-15s  %(message)s")
logger = logging.getLogger('docx')


def dprint(_func=None, **kw):
    f = sys._getframe(1)
    if kw:
        text = ', '.join('%s = %s' % (k, v) for k, v in kw.items())
    else:
        try:
            text = dict((k, repr(v)) for k, v in f.f_locals.items()
                        if k != 'self')
            text = six.text_type(text)
        except:
            text = ''

    if _func is None:
        _func = f.f_code.co_name

    logger.info(' '.join([_func, text]))

# Utility functions


def remove_items(src, target):
    for x in target:
        src.remove(x)


def get_items_list(src):
    result = []
    for x in src:
        if x and x != [[]]:
            result.append(x)
    return result


def findElement(elem, tag):
    if not isinstance(elem, nodes.Element):
        return None
    if elem.tagname == tag:
        return elem
    for child in elem:
        res = findElement(child, tag)
        if res is not None:
            return res
    return None

def get_toc_maxdepth(builder, docname):
    toc = findElement(builder.env.tocs[docname], 'toctree')
    if toc is not None:
        return toc.get('maxdepth', -1)
    return -1

def get_image_size(filename):
    if Image is None:
        raise RuntimeError(
            'image size not fully specified and PIL not installed')
    with Image.open(filename, 'r') as imageobj:
        dpi = imageobj.info.get('dpi', (72, 72))
        # dpi information can be (xdpi, ydpi) or xydpi
        try:
            iter(dpi)
        except:
            dpi = (dpi, dpi)
        width = imageobj.size[0]
        height = imageobj.size[1]
        cmperin = 2.54
        return (width * cmperin / dpi[0], height * cmperin / dpi[1])

def convert_to_cm_size(size_with_unit, max_width):
    if size_with_unit is None:
        return None
    if size_with_unit.endswith('%'):
        return max_width * float(size_with_unit[:-1]) / 100

    match = re.match(r'^(\d+(?:\.\d*)?)(\D*)$', size_with_unit)
    if not match:
        raise RuntimeError('Unexpected length unit: %s' % size_with_unit)
    size = float(match.group(1))
    unit = match.group(2)
    if not unit:
        unit = 'px'

    cmperin = 2.54
    ratio_map = {
            'em': 12 * cmperin / 144, # TODO: Use BodyText font size
            'ex': 12 * cmperin / 144,
            'mm': 0.1, 'cm': 1, 'in': cmperin,
            'px': cmperin / 96, 'pt': cmperin / 72, 'pc': cmperin / 6,
    }
    ratio = ratio_map.get(unit)
    if ratio is None:
        raise RuntimeError('Unknown length unit: %s' % size_with_unit)
    return size * ratio

#
#  DocxWriter class for sphinx
#


class DocxWriter(writers.Writer):
    supported = ('docx',)
    settings_spec = ('No options here.', '', ())
    settings_defaults = {}

    output = None

    def __init__(self, builder):
        writers.Writer.__init__(self)
        self.builder = builder
        self.docx = docx.DocxComposer()
        self.numsec_map = None
        self.numfig_map = None

        self.title = self.builder.config['docx_title']
        self.subject = self.builder.config['docx_subject']
        self.creator = self.builder.config['docx_creator']
        self.company = self.builder.config['docx_company']
        self.category = self.builder.config['docx_category']
        self.descriptions = self.builder.config['docx_descriptions']
        self.keywords = self.builder.config['docx_keywords']
        try:
            self.coverpage = self.builder.config['docx_coverpage']
        except:
            self.coverpage = True

        stylefile = self.builder.config['docx_style']
        if stylefile:
            self.docx.new_document(os.path.join(
                self.builder.confdir, os.path.join(stylefile)))
        else:
            default_style = os.path.join(
                    os.path.dirname(__file__), 'docx/style.docx')
            self.docx.new_document(default_style)

    def set_numsec_map(self, numsec_map):
        self.numsec_map = numsec_map

    def set_numfig_map(self, numfig_map):
        self.numfig_map = numfig_map

    def save(self, filename):
        self.docx.set_coverpage(self.coverpage)

        self.docx.set_props(title=self.title,
                            subject=self.subject,
                            creator=self.creator,
                            company=self.company,
                            category=self.category,
                            descriptions=self.descriptions,
                            keywords=self.keywords)

        self.docx.save(filename)

    def translate(self):
        visitor = DocxTranslator(
                self.document, self.builder, self.docx,
                self.numsec_map, self.numfig_map)
        self.document.walkabout(visitor)
        self.output = ''  # visitor.body

#
#  DocxTranslator class for sphinx
#

def make_run(text, style):
    run_tree = [['w:r']]
    if style:
        run_tree.append([['w:rPr'], [['w:rStyle', {'w:val': style}]]])
    attrs = {}
    if text.find(' ') != -1:
        attrs['xml:space'] = 'preserve'
    run_tree.append([['w:t', text, attrs]])
    return docx.make_element_tree(run_tree)

def make_break_run():
    return docx.make_element_tree([['w:r'], [['w:br']]])

def make_hyperlink(relationship_id, anchor):
    attrs = {}
    if relationship_id is not None:
        attrs['r:id'] = relationship_id
    if anchor is not None:
        attrs['w:anchor'] = anchor
    hyperlink_tree = [['w:hyperlink', attrs]]
    return docx.make_element_tree(hyperlink_tree)

def make_paragraph(indent_level, style, align):
    if style is None:
        style = 'BodyText'
    style_tree = [
            ['w:pPr'],
            [['w:pStyle', {'w:val': style}]],
    ]
    if indent_level is not None:
        indent = indent_level * 480 # TODO
        style_tree.append(
            [['w:ind', {'w:leftChars': '0', 'w:left': str(indent)}]])
    if align is not None:
        style_tree.append([['w:jc', {'w:val': align}]])

    paragraph_tree = [['w:p'], style_tree]
    return docx.make_element_tree(paragraph_tree)

def make_list_paragraph(num_id, list_indent_level, indent, style):
    paragraph_tree = [
            ['w:p'],
            [
                ['w:pPr'],
                [['w:pStyle', {'w:val': style}]],
                [
                    ['w:numPr'],
                    [['w:ilvl', {'w:val': str(list_indent_level)}]],
                    [['w:numId', {'w:val': str(num_id)}]],
                ],
                [['w:ind', {'w:leftChars': '0', 'w:left': str(indent)}]]
            ]
    ]
    return docx.make_element_tree(paragraph_tree)

def to_error_string(contents):
    from xml.etree.ElementTree import tostring
    func = lambda xml: tostring(xml, encoding='utf8').decode('utf8')
    xml_list = contents.to_xml()
    return type(contents).__name__ + '\n' + '\n'.join(map(func, xml_list))

class BookmarkStart(object):
    def __init__(self, id, name):
        self._id = id
        self._name = name

    def to_xml(self):
        return [docx.make_element_tree([
            ['w:bookmarkStart', {'w:id': str(self._id), 'w:name': self._name}]
        ])]

class BookmarkEnd(object):
    def __init__(self, id):
        self._id = id

    def to_xml(self):
        return [docx.make_element_tree([
            ['w:bookmarkEnd', {'w:id': str(self._id)}]
        ])]

class Paragraph(object):
    def __init__(self, indent_level=None, paragraph_style=None, align=None):
        self._indent_level = indent_level
        self._style = paragraph_style
        self._align = align
        self._run_list = []
        self._text_style_stack = [None]

    def add_text(self, text):
        text = re.sub(r'\n', ' ', text)
        self._run_list.append(make_run(text, self._text_style_stack[-1]))

    def add_break(self):
        self._run_list.append(make_break_run())

    def add_picture(self, rid, filename, width, height, alt):
        self._run_list.append(docx.DocxComposer.make_inline_picture_run(
            rid, filename, width, height, alt))

    def push_style(self, text_style):
        self._text_style_stack.append(text_style)

    def pop_style(self):
        self._text_style_stack.pop()

    def append(self, contents):
        if isinstance(contents, Paragraph): # for nested line_block
            self._run_list.extend(contents._run_list)
        elif isinstance(contents, (BookmarkStart, BookmarkEnd, HyperLink)):
            self._run_list.extend(contents.to_xml())
        else:
            raise RuntimeError('Can not append %s' % to_error_string(contents))

    def to_xml(self):
        p = make_paragraph(self._indent_level, self._style, self._align)
        p.extend(self._run_list)
        return [p]

class LiteralBlock(object):
    def __init__(self, indent_level, language, highlight_args, highlighter):
        self._indent_level = indent_level
        self._language = language
        self._highlight_args = highlight_args
        self._highlighter = highlighter
        self._text_list = []

    def add_text(self, text):
        self._text_list.append(text)

    def append(self, contents):
        raise RuntimeError('Can not append %s' % to_error_string(contents))

    def to_xml(self):
        p = make_paragraph(self._indent_level, 'LiteralBlock', None)
        for text in self._text_list:
            lineos = 1
            highlighted = self._highlighter.highlight_block(
                    text, self._language, lineos=lineos, **self._highlight_args)
            xml_text = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' + highlighted + '</w:p>'
            dummy_p = etree.fromstring(xml_text)
            run_list = docx.get_elements(dummy_p, 'w:r')
            # highlight_block may append a line break to the tail of the code
            if (not text.endswith('\n') and
                    run_list[-1][-1].tag == docx.norm_name('w:br')):
                run_list[-1].remove(run_list[-1][-1])
            for run in run_list:
                p.append(run)
        return [p]

class HyperLink(object):
    def __init__(self, rid, anchor):
        self._rid = rid
        self._anchor = anchor
        self._run_list = []
        self._text_style_stack = ['HyperLink']

    def add_text(self, text):
        text = re.sub(r'\n', ' ', text)
        self._run_list.append(make_run(text, self._text_style_stack[-1]))

    def add_break(self):
        self._run_list.append(make_break_run())

    def push_style(self, text_style):
        self._text_style_stack.append(text_style)

    def pop_style(self):
        self._text_style_stack.pop()

    def append(self, contents):
        if isinstance(contents, (BookmarkStart, BookmarkEnd)):
            self._run_list.extend(contents.to_xml())
        else:
            raise RuntimeError('Can not append %s' % to_error_string(contents))

    def to_xml(self):
        if self._rid is None and self._anchor is None:
            return self._run_list
        h = make_hyperlink(self._rid, self._anchor)
        h.extend(self._run_list)
        return [h]

class ListItem(object):
    def __init__(self, num_id, list_indent_level, indent, list_style):
        p = make_list_paragraph(num_id, list_indent_level, indent, list_style)
        self._xml_list = [p]
        self._available_list_paragraph = True

    def append(self, contents):
        if len(self._xml_list) == 1:
            if isinstance(contents, (BookmarkStart, BookmarkEnd)):
                self._xml_list[-1].extend(contents.to_xml())
                return
            if self._available_list_paragraph:
                if isinstance(contents, Paragraph) and contents._style is None:
                    self._xml_list[-1].extend(contents._run_list)
                    return
                else:
                    self._available_list_paragraph = False
        self._xml_list.extend(contents.to_xml())

    def to_xml(self):
        return self._xml_list

class DefinitionListItem(object):
    def __init__(self):
        self._contents_list = []
        self._last_term = None

    @property
    def last_term(self):
        return self._last_term

    def add_term(self, term_paragraph):
        self._contents_list.append(term_paragraph)
        self._last_term = term_paragraph

    def append(self, contents):
        self._contents_list.append(contents)

    def to_xml(self):
        return list(itertools.chain.from_iterable(
            map(lambda c: c.to_xml(), self._contents_list)))

class Table(object):
    def __init__(self, table_style, colsize_list):
        self._style = table_style
        self._colsize_list = colsize_list
        self._stub = 0
        self._head = []
        self._body = []
        self._current_target = self._body

    @property
    def style(self):
        return self._style

    def add_col_width(self, col_width):
        self._colsize_list.append(col_width)

    def add_stub(self):
        self._stub += 1

    def start_head(self):
        self._current_target = self._head

    def start_body(self):
        self._current_target = self._body

    def add_row(self):
        self._current_target.append([])

    def add_cell(self):
        self._current_target[-1].append([])

    def current_cell_width(self):
        num_cell = len(self._current_target[-1])
        if len(self._colsize_list) < num_cell:
            return None
        return self._colsize_list[num_cell - 1]

    def append(self, contents):
        self._current_target[-1][-1].extend(contents.to_xml())

    def to_xml(self):
        look_attrs = {
                'w:noHBand': 'false', 'w:noVBand': 'false',
                'w:lastRow': 'false', 'w:lastColumn': 'false'
        }
        look_attrs['w:firstRow'] = 'true' if self._head else 'false'
        look_attrs['w:firstColumn'] = 'true' if self._stub > 0 else 'false'
        table_tree = [
                ['w:tbl'],
                [
                    ['w:tblPr'],
                    [['w:tblStyle', {'w:val': self._style}]],
                    [['w:tblW', {'w:w': '0', 'w:type': 'auto'}]],
                    [['w:tblLook', look_attrs]],
                ],
        ]
        table_grid_tree = [['w:tblGrid']]
        for colsize in self._colsize_list:
            table_grid_tree.append([['w:gridCol', {'w:w': str(colsize)}]])
        table_tree.append(table_grid_tree)

        table = docx.make_element_tree(table_tree)
        for index, row in enumerate(self._head):
            table.append(self.make_row(index, row, True))
        for index, row in enumerate(self._body):
            table.append(self.make_row(index, row, False))
        return [table]

    def make_row(self, index, row, is_head):
        row_style = {
                'w:evenHBand': ('true' if index % 2 == 0 else 'false'),
                'w:oddHBand': ('true' if index % 2 != 0 else 'false'),
                'w:firstRow': ('true' if is_head else 'false'),
        }
        tr_tree = docx.make_element_tree([
                ['w:tr'],
                [
                    ['w:trPr'],
                    [['w:cnfStyle', row_style]]
                ]
        ])
        for index, cell in enumerate(row):
            tr_tree.append(self.make_cell(index, cell))
        return tr_tree

    def make_cell(self, index, cell):
        cell_style = {
                'w:evenVBand': ('true' if index % 2 == 0 else 'false'),
                'w:oddVBand': ('true' if index % 2 != 0 else 'false'),
                'w:firstColumn': ('true' if index < self._stub else 'false'),
        }
        cellsize = str(self._colsize_list[index])
        tc_tree = [
                ['w:tc'],
                [
                    ['w:tcPr'],
                    [['w:cnfStyle', cell_style]],
                    [['w:tcW', {'w:w': cellsize, 'w:type': 'dxa'}]]
                ]
        ]
        elem = docx.make_element_tree(tc_tree)
        elem.extend(cell)
        return elem

class Document(object):
    def __init__(self, body):
        self._body = body

    def append(self, contents):
        for xml in contents.to_xml():
            self._body.append(xml)

def admonition(table_style):
    def _visit_admonition(func):
        def visit_admonition(self, node):
            self._append_bookmark_start(node.get('ids', []))
            table_width = self._table_width_stack[-1]
            t = self._append_table(table_style, [table_width - 1000])
            t.start_head()
            t.add_row()
            self._add_table_cell()
            p = Paragraph(self._indent_level_stack[-1])
            p.add_text(admonitionlabels[node.tagname] + ':')
            t.append(p)
            t.start_body()
            t.add_row()
            self._add_table_cell()
        return visit_admonition
    return _visit_admonition

class DocxTranslator(nodes.NodeVisitor):
    def __init__(self, document, builder, docx, numsec_map, numfig_map):
        nodes.NodeVisitor.__init__(self, document)
        self._builder = builder
        self.builder = self._builder # Needs for graphviz.render_dot
        self._doc_stack = [Document(docx.docbody)]
        self._docname_stack = [builder.config.master_doc]
        self._section_level_stack = [0]
        self._indent_level_stack = [0]
        self._table_width_stack = [docx.max_table_width]
        self._line_block_level = 0
        self._docx = docx
        self._max_list_id = docx.get_max_numbering_id()
        self._list_id_stack = []
        self._bullet_list_id = (
                docx.styleDocx.get_numbering_style_id('ListBullet'))
        self._language = 'guess'
        self._highlighter = DocxPygmentsBridge(
                'html',
                builder.config.pygments_style,
                builder.config.trim_doctest_flags)
        self._numsec_map = numsec_map
        self._numfig_map = numfig_map
        self._bookmark_id = 0
        self._bookmark_id_map = {} # bookmark name => BookmarkStart id
        self._toc_out = False

    def _pop_and_append(self):
        contents = self._doc_stack.pop()
        self._doc_stack[-1].append(contents)

    def _append_bookmark_start(self, ids):
        docname = self._docname_stack[-1]
        for id in ids:
            name = '%s#%s' % (docname, id)
            self._bookmark_id += 1
            self._bookmark_id_map[name] = self._bookmark_id
            self._doc_stack[-1].append(BookmarkStart(self._bookmark_id, name))

    def _append_bookmark_end(self, ids):
        docname = self._docname_stack[-1]
        for id in ids:
            name = '%s#%s' % (docname, id)
            bookmark_id = self._bookmark_id_map.pop(name, None)
            if bookmark_id is None:
                continue
            self._doc_stack[-1].append(BookmarkEnd(bookmark_id))

    def _append_table(self, table_style, colsize_list):
        t = Table(table_style, colsize_list)
        self._doc_stack.append(t)
        self._indent_level_stack.append(0)
        self._table_width_stack.append(self._table_width_stack[-1])
        return t

    def _pop_and_append_table(self):
        self._indent_level_stack.pop()
        self._table_width_stack.pop()
        self._pop_and_append()

    def _add_table_cell(self):
        t = self._doc_stack[-1]
        t.add_cell()
        width = t.current_cell_width()
        if width is not None:
            margin = self._docx.get_table_cell_margin(t.style)
            self._table_width_stack[-1] = width - margin

    def _append_picture(self, filepath, width, height, alt):
        rid = self._docx.add_image_relationship(filepath)
        filename = os.path.basename(filepath)
        # TODO: Check whether _doc_stack already includes Paragraph
        p = Paragraph(align='center')
        p.add_picture(rid, filename, width, height, alt)
        self._doc_stack[-1].append(p)

    def _get_numsec(self, ids):
        for id in ids:
            num = self._numsec_map.get('%s/#%s' % (self._docname_stack[-1], id))
            if num:
                return '.'.join(map(str, num)) + ' '
        else:
            # First section of each file has no hash
            num = self._numsec_map.get('%s/' % self._docname_stack[-1], None)
            if num:
                return '.'.join(map(str, num)) + ' '
        return None

    def _get_numfig(self, figtype, ids):
        item = self._numfig_map.get(figtype)
        if item is None:
            return None
        prefix, num_map = item
        if prefix is None:
            return None
        for id in ids:
            num = num_map.get('%s/%s' % (self._docname_stack[-1], id))
            if num:
                return prefix % ('.'.join(map(str, num)) + ' ')
        return None

    def visit_start_of_file(self, node):
        self._docname_stack.append(node['docname'])
        self._append_bookmark_start([''])
        self._append_bookmark_start(node.get('ids', []))
        self._section_level_stack.append(0)

    def depart_start_of_file(self, node):
        self._section_level_stack.pop()
        self._append_bookmark_end(node.get('ids', []))
        self._append_bookmark_end([''])
        self._docname_stack.pop()

    def visit_Text(self, node):
        self._doc_stack[-1].add_text(node.astext())

    def depart_Text(self, node):
        pass

    def visit_document(self, node):
        self._toc_out = False

    def depart_document(self, node):
        pass

    def visit_title(self, node):
        self._append_bookmark_start(node.get('ids', []))
        if isinstance(node.parent, nodes.table):
            style = 'TableHeading'
            title_num = self._get_numfig('table', node.parent['ids'])
        elif isinstance(node.parent, nodes.section):
            style = 'Heading%d' % self._section_level_stack[-1]
            title_num = self._get_numsec(node.parent['ids'])
        else:
            style = None # TODO
            title_num = None
        self._doc_stack.append(Paragraph(paragraph_style=style))
        if title_num is not None:
            self._doc_stack[-1].add_text(title_num)

    def depart_title(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_subtitle(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack.append(Paragraph(self._indent_level_stack[-1])) # TODO

    def depart_subtitle(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_section(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._section_level_stack[-1] += 1

    def depart_section(self, node):
        self._section_level_stack[-1] -= 1
        self._append_bookmark_end(node.get('ids', []))

    def visit_topic(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_topic(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_sidebar(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_sidebar(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_transition(self, node):
        pass # TODO

    def depart_transition(self, node):
        pass

    def visit_paragraph(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack.append(Paragraph(self._indent_level_stack[-1]))

    def depart_paragraph(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_compound(self, node):
        self._append_bookmark_start(node.get('ids', []))
        if not self._toc_out: # TODO
            self._toc_out = True
            maxdepth = get_toc_maxdepth(
                    self._builder, self._builder.config.master_doc)
            if maxdepth < 1:
                maxdepth = 10
            self._docx.table_of_contents(toc_text='Contents', maxlevel=maxdepth)
            self._docx.pagebreak(type='page', orient='portrait')

    def depart_compound(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_container(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_container(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_literal_block(self, node):
        self._append_bookmark_start(node.get('ids', []))
        language = node.get('language', self._language)
        highlight_args = node.get('highlight_args', {})
        self._doc_stack.append(
                LiteralBlock(self._indent_level_stack[-1],
                             language, highlight_args, self._highlighter))

    def depart_literal_block(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_doctest_block(self, node):
        raise nodes.SkipNode # TODO

    def depart_doctest_block(self, node):
        pass

    def visit_math_block(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack.append(Paragraph(self._indent_level_stack[-1])) # TODO

    def depart_math_block(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_line_block(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack.append(Paragraph(self._indent_level_stack[-1]))
        self._line_block_level += 1

    def depart_line_block(self, node):
        self._line_block_level -= 1
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_line(self, node):
        self._append_bookmark_start(node.get('ids', []))
        indent = ''.join('    ' for _ in range(self._line_block_level - 1))
        self._doc_stack[-1].add_text(indent)

    def depart_line(self, node):
        self._doc_stack[-1].add_break()
        self._append_bookmark_end(node.get('ids', []))

    def visit_block_quote(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._indent_level_stack[-1] += 1

    def depart_block_quote(self, node):
        self._indent_level_stack[-1] -= 1
        self._append_bookmark_end(node.get('ids', []))

    def visit_attribution(self, node):
        self._append_bookmark_start(node.get('ids', []))
        p = Paragraph(self._indent_level_stack[-1])
        p.add_text('— ')
        self._doc_stack.append(p)

    def depart_attribution(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_table(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_table(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_tgroup(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._append_table('rstTable', [])

    def depart_tgroup(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    def visit_colspec(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table = self._doc_stack[-1]
        table_width = self._table_width_stack[-1]
        table.add_col_width(int(table_width * node['colwidth'] / 100))
        if node.get('stub', 0) == 1:
            table.add_stub()

    def depart_colspec(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_thead(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table = self._doc_stack[-1]
        table.start_head()

    def depart_thead(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_tbody(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table = self._doc_stack[-1]
        table.start_body()

    def depart_tbody(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_row(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table = self._doc_stack[-1]
        table.add_row()

    def depart_row(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_entry(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._add_table_cell()

    def depart_entry(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_figure(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_figure(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_caption(self, node):
        self._append_bookmark_start(node.get('ids', []))
        if isinstance(node.parent, nodes.figure):
            style = 'ImageCaption'
            figtype = 'figure'
        else:
            style = 'LiteralCaption'
            figtype = 'code-block'
        self._doc_stack.append(Paragraph(paragraph_style=style))
        caption_num = self._get_numfig(figtype, node.parent['ids'])
        if caption_num is not None:
            self._doc_stack[-1].add_text(caption_num)

    def depart_caption(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_legend(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_legend(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_footnote(self, node):
        raise nodes.SkipNode # TODO

    def depart_footnote(self, node):
        pass

    def visit_citation(self, node):
        raise nodes.SkipNode # TODO

    def depart_citation(self, node):
        pass

    def visit_label(self, node):
        pass # TODO

    def depart_label(self, node):
        pass

    def visit_rubric(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack.append(Paragraph(self._indent_level_stack[-1])) # TODO

    def depart_rubric(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_bullet_list(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._indent_level_stack[-1] += 1
        self._list_id_stack.append(self._bullet_list_id)

    def depart_bullet_list(self, node):
        self._indent_level_stack[-1] -= 1
        self._list_id_stack.pop()
        self._append_bookmark_end(node.get('ids', []))

    def visit_enumerated_list(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._indent_level_stack[-1] += 1
        self._max_list_id += 1
        self._list_id_stack.append(self._max_list_id)
        enumtype = node.get('enumtype', 'arabic')
        prefix = node.get('prefix', '')
        suffix = node.get('suffix', '')
        start = node.get('start', 1)
        self._docx.new_ListNumber_style(
                self._list_id_stack[-1], start,
                '{}%1{}'.format(prefix, suffix), enumtype)

    def depart_enumerated_list(self, node):
        self._indent_level_stack[-1] -= 1
        self._list_id_stack.pop()
        self._append_bookmark_end(node.get('ids', []))

    def visit_list_item(self, node):
        self._append_bookmark_start(node.get('ids', []))
        indent_level = self._indent_level_stack[-1] - 1
        if isinstance(node.parent, nodes.enumerated_list):
            style = 'ListNumber'
            list_indent_level = 0
        else:
            style = 'ListBullet'
            list_indent_level = indent_level
        list_id = self._list_id_stack[-1]
        indent = self._docx.get_numbering_indent(style, indent_level, list_id)
        self._doc_stack.append(
                ListItem(list_id, list_indent_level, indent, style))

    def depart_list_item(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_definition_list(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_definition_list(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_definition_list_item(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack.append(DefinitionListItem())

    def depart_definition_list_item(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_term(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack.append(
                Paragraph(self._indent_level_stack[-1], 'DefinitionItem'))

    def depart_term(self, node):
        term_paragraph = self._doc_stack.pop()
        self._doc_stack[-1].add_term(term_paragraph)
        self._append_bookmark_end(node.get('ids', []))

    def visit_classifier(self, node):
        self._append_bookmark_start(node.get('ids', []))
        term_paragraph = self._doc_stack[-1].last_term
        self._doc_stack.append(term_paragraph)
        term_paragraph.add_text(' : ')

    def depart_classifier(self, node):
        self._doc_stack.pop()
        self._append_bookmark_end(node.get('ids', []))

    def visit_definition(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._indent_level_stack[-1] += 1

    def depart_definition(self, node):
        self._indent_level_stack[-1] -= 1
        self._append_bookmark_end(node.get('ids', []))

    def visit_field_list(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table_width = self._table_width_stack[-1]
        colsize_list = [int(table_width * 1 / 4), int(table_width * 3 / 4)]
        self._append_table('FieldList', colsize_list)

    def depart_field_list(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    def visit_field(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table = self._doc_stack[-1]
        table.add_row()

    def depart_field(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_field_name(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._add_table_cell()
        self._doc_stack.append(Paragraph(self._indent_level_stack[-1]))

    def depart_field_name(self, node):
        self._doc_stack[-1].add_text(':')
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_field_body(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._add_table_cell()

    def depart_field_body(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_option_list(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table_width = self._table_width_stack[-1]
        self._append_table('OptionList', [table_width - 500])

    def depart_option_list(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    def visit_option_list_item(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_option_list_item(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_option_group(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table = self._doc_stack[-1]
        table.add_row()
        self._add_table_cell()
        self._doc_stack.append(Paragraph(self._indent_level_stack[-1]))

    def depart_option_group(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_option(self, node):
        self._append_bookmark_start(node.get('ids', []))
        parent = node.parent
        first_option_index = parent.first_child_matching_class(nodes.option)
        if parent[first_option_index] is not node:
            self._doc_stack[-1].add_text(', ')

    def depart_option(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_option_string(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_option_string(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_option_argument(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].add_text(node.get('delimiter', ' '))

    def depart_option_argument(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_description(self, node):
        self._append_bookmark_start(node.get('ids', []))
        table = self._doc_stack[-1]
        table.add_row()
        self._add_table_cell()
        self._indent_level_stack[-1] += 1

    def depart_description(self, node):
        self._indent_level_stack[-1] -= 1
        self._append_bookmark_end(node.get('ids', []))

    @admonition('AttentionAdmonition')
    def visit_attention(self, node):
        pass

    def depart_attention(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    @admonition('CautionAdmonition')
    def visit_caution(self, node):
        pass

    def depart_caution(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    @admonition('DangerAdmonition')
    def visit_danger(self, node):
        pass

    def depart_danger(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    @admonition('ErrorAdmonition')
    def visit_error(self, node):
        pass

    def depart_error(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    @admonition('HintAdmonition')
    def visit_hint(self, node):
        pass

    def depart_hint(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    @admonition('ImportantAdmonition')
    def visit_important(self, node):
        pass

    def depart_important(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    @admonition('NoteAdmonition')
    def visit_note(self, node):
        pass

    def depart_note(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    @admonition('TipAdmonition')
    def visit_tip(self, node):
        pass

    def depart_tip(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    @admonition('WarningAdmonition')
    def visit_warning(self, node):
        pass

    def depart_warning(self, node):
        self._pop_and_append_table()
        self._append_bookmark_end(node.get('ids', []))

    def visit_admonition(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_admonition(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_substitution_definition(self, node):
        raise nodes.SkipNode # TODO

    def visit_comment(self, node):
        raise nodes.SkipNode # TODO

    def visit_pending(self, node):
        raise nodes.SkipNode # TODO

    def visit_system_message(self, node):
        raise nodes.SkipNode # TODO


    def visit_emphasis(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('Emphasis')

    def depart_emphasis(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_strong(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('Strong')

    def depart_strong(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_literal(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('Literal')

    def depart_literal(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_math(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_math(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_reference(self, node):
        self._append_bookmark_start(node.get('ids', []))
        refuri = node.get('refuri', None)
        refid = node.get('refid')
        if refuri:
            if node.get('internal', False):
                rid = None
                anchor = self._get_bookmark_name(refuri)
            else:
                rid = self._docx.add_hyperlink_relationship(refuri)
                anchor = None
        else:
            rid = None
            anchor = '%s#%s' % (self._docname_stack[-1], refid)
        self._doc_stack.append(HyperLink(rid, anchor))

    def depart_reference(self, node):
        hyperlink = self._doc_stack.pop()
        if isinstance(self._doc_stack[-1], Paragraph):
            self._doc_stack[-1].append(hyperlink)
        else:
            p = Paragraph(self._indent_level_stack[-1])
            p.append(hyperlink)
            self._doc_stack[-1].append(p)
        self._append_bookmark_end(node.get('ids', []))

    def visit_footnote_reference(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_footnote_reference(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_citation_reference(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_citation_reference(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_substitution_reference(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_substitution_reference(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_title_reference(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('TitleReference')

    def depart_title_reference(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_abbreviation(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('Abbreviation') # TODO

    def depart_abbreviation(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_acronym(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_acronym(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_subscript(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('Subscript')

    def depart_subscript(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_superscript(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('Superscript')

    def depart_superscript(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_inline(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_inline(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_problematic(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('Problematic')

    def depart_problematic(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_generated(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_generated(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_target(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_target(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_image(self, node):
        self._append_bookmark_start(node.get('ids', []))
        uri = node.attributes['uri']
        file_path = os.path.join(self._builder.env.srcdir, uri)
        width, height = self._get_image_scaled_size(node, file_path)

        self._append_picture(file_path, width, height, node.get('alt', ''))

    def depart_image(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_raw(self, node):
        raise nodes.SkipNode # TODO


    def visit_compact_paragraph(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_compact_paragraph(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_literal_emphasis(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack[-1].push_style('LiteralEmphasis')

    def depart_literal_emphasis(self, node):
        self._doc_stack[-1].pop_style()
        self._append_bookmark_end(node.get('ids', []))

    def visit_highlightlang(self, node):
        self._language = node.get('lang', 'guess')
        raise nodes.SkipNode

    def visit_glossary(self, node):
        self._append_bookmark_start(node.get('ids', []))

    def depart_glossary(self, node):
        self._append_bookmark_end(node.get('ids', []))

    def visit_desc(self, node):
        raise nodes.SkipNode # TODO

    def depart_desc(self, node):
        pass

    def visit_desc_signature(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_desc_signature(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_name(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_name(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_addname(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_addname(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_type(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_type(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_returns(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_returns(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_parameterlist(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_parameterlist(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_parameter(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_parameter(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_optional(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_optional(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_annotation(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_annotation(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_desc_content(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass

    def depart_desc_content(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_productionlist(self, node):
        raise nodes.SkipNode # TODO

    def depart_productionlist(self, node):
        pass

    def visit_seealso(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_seealso(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_tabular_col_spec(self, node):
        raise nodes.SkipNode # Do nothing

    def visit_acks(self, node):
        raise nodes.SkipNode # TODO

    def depart_acks(self, node):
        pass

    def visit_centered(self, node):
        self._append_bookmark_start(node.get('ids', []))
        self._doc_stack.append(
                Paragraph(self._indent_level_stack[-1], align='center'))

    def depart_centered(self, node):
        self._pop_and_append()
        self._append_bookmark_end(node.get('ids', []))

    def visit_hlist(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO make table

    def depart_hlist(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_hlistcol(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_hlistcol(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_versionmodified(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_versionmodified(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_index(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_index(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_pending_xref(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_pending_xref(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_download_reference(self, node):
        self._append_bookmark_start(node.get('ids', []))
        pass # TODO

    def depart_download_reference(self, node):
        self._append_bookmark_end(node.get('ids', []))
        pass

    def visit_number_reference(self, node):
        self.visit_reference(node)

    def depart_number_reference(self, node):
        self.depart_reference(node)

    def visit_meta(self, node):
        raise nodes.SkipNode

    def visit_graphviz(self, node):
        self._append_bookmark_start(node.get('ids', []))
        fname, filename = graphviz.render_dot(
            self, node['code'], node['options'], 'png')
        width, height = self._get_image_scaled_size(node, filename)
        self._append_picture(filename, width, height, node.get('alt', ''))
        self._append_bookmark_end(node.get('ids', []))
        raise nodes.SkipNode

    def visit_refcount(self, node):
        raise nodes.SkipNode # TODO

    def depart_refcount(self, node):
        pass

    def unknown_visit(self, node):
        print(node.tagname)
        raise nodes.SkipNode

    def _get_bookmark_name(self, refuri):
        hashindex = refuri.find('#')
        if hashindex == 0:
            return self._docname_stack[-1] + refuri
        if hashindex < 0 and refuri in self._builder.env.all_docs:
            return refuri + '#'
        if refuri[:hashindex] in self._builder.env.all_docs:
            return refuri
        return None

    def _get_image_scaled_size(self, node, filename):
        twippercm = 567.0
        max_width = self._table_width_stack[-1] / twippercm

        width = self._get_cm_size(node, 'width', max_width)
        height = self._get_cm_size(node, 'height')

        if width is None and height is None:
            width, height = get_image_size(filename)
        elif width is None:
            img_width, img_height = get_image_size(filename)
            width = img_width * height / img_height
        elif height is None:
            img_width, img_height = get_image_size(filename)
            height = img_height * width / img_width

        scale = node.get('scale')
        if scale is not None:
            scale = float(scale) / 100
            width *= scale
            height *= scale

        if width > max_width:
            ratio = max_width / width
            width = max_width
            height *= ratio

        return width, height

    def _get_cm_size(self, node, attr, max_width=0):
        try:
            return convert_to_cm_size(node.get(attr), max_width)
        except Exception as e:
            self.document.reporter.warning(e)
            return None
