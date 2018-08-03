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

import re

from docutils import nodes, writers

from sphinx import addnodes
from sphinx import highlighting
from sphinx.locale import admonitionlabels, versionlabels, _

from sphinx.ext import graphviz

import docx
import sys
import os
import zipfile
import tempfile
from lxml import etree
from highlight import *


#
# Is the PIL imaging library installed?
try:
    import Image
except ImportError, exp:
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
            text = unicode(text)
        except:
            text = ''

    if _func is None:
        _func = f.f_code.co_name

    logger.info(' '.join([_func, text]))

###### Utility functions
def remove_items(src, target):
  for x in target:
    src.remove(x)

def get_items_list(src):
  result=[]
  for x in src:
    if x  and x != [[]]:
      result.append(x)
  return result

def findElement(elem, tag):
  res = None
  if not elem :
    return res

  for x in elem :
    try:
      if x.tagname == tag:
        return x
      else:
        res = findElement(x.children, tag)
    except:
      res = None
  return res

def get_toc_maxdepth(builder, docname):
  toc_maxdepth = 0
  try:
    toc = findElement(builder.env.tocs[docname], 'toctree')

    if toc :
      toc_maxdepth = toc['maxdepth']
  except:
    toc_maxdepth = 0
  return toc_maxdepth
  
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
        if stylefile :
            self.docx.new_document(stylefile)
        else:
            self.docx.new_document('style.docx')

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
        visitor = DocxTranslator(self.document, self.builder, self.docx)
        self.document.walkabout(visitor)
        self.output = ''  # visitor.body

#
#  DocxTranslator class for sphinx
#
class DocxTranslator(nodes.NodeVisitor):

    def __init__(self, document, builder, docx):
        self.builder = builder
        self.docx = docx
        nodes.NodeVisitor.__init__(self, document)

        self.states = [[]]
        self.list_style = []
        self.sectionlevel = 0
        self.table = None

        self.line_block_level = 0

        self.list_item_flushed = False

        self.current_block = []
	self.list_level = 0
	self.block_level = 0
	self.num_list_id = docx.get_max_numbering_id()  + 1
	self.max_num_list_id = self.num_list_id

	self.enum_prefix_style = []

        self.field_name = None

        self.admonition_body = None
        self.current_field_list = None
        self.current_option_list = None
	self.literal_block_lang = None

        self.highlighter = DocxPygmentsBridge('html', builder.config.pygments_style, builder.config.trim_doctest_flags)

        self.option = []

    def add_text(self, text):
        '''
	   Add text in states
        '''
        dprint()
	if not self.states :
	  if self.states[-1] is not [] :
	    self.states.append([])

        self.states[-1].append(text)

    def add_linebreak(self):
        '''
	   Add linebreak-text(:br) in states
        '''
        dprint()
	self.add_text(':br')

    def new_state(self):
        '''
	   create a new state
        '''
        dprint()

	if len(self.current_block) == 0 :
          self.ensure_state()
        if self.states[-1] is not [] :
          self.states.append([])

    def ensure_state(self):
        '''
	   ensure state and flush all states
        '''
        dprint()
        self.flush_state()

    def flush_state(self, _sty=None, typ = -1, enumprefix=None, enumprefixtype=None, start_num=1):
        '''
	   flush all states
        '''
        dprint()
        result=False

	if _sty is 'List_item' :
          self.flush_enum_list_item()
        else:
          if _sty is None :
	    if len(self.current_block) > 0 and self.current_block[-1] == 'List_item' :
              self.flush_enum_list_item()
              result=True

        self.flush_state_all(_sty)

        return result

    def flush_state_all(self, _sty=None, _create_only=False):
        '''
	   flush all states
        '''
        dprint()
	p=[]

        b_level = self.block_level + self.list_level

        for texts in  self.states:
          if texts :
            if _sty :
#                if _sty == 'LiteralBlock':
#                    print texts
#                    pass
                p.append( self.docx.paragraph(texts, style=_sty, block_level=b_level, create_only=_create_only))
            else:
                p.append( self.docx.paragraph(texts, block_level=b_level, create_only=_create_only))
        self.states = [[]]
	return p

    def end_state(self, first=None):
        '''
	   clear states
        '''
        dprint()
	try:
            result = self.states.pop()
            if first is not None and result:
                item = result[0]
                if item:
                    result.insert(0, [first + item[0]])
                    result[1] = item[1:]
	    if not self.states :
	        if self.states[-1] is not [] :
                    self.states.append([])
            self.states[-1].extend(result)
	except:
            self.states=[[]]
            pass

    def flush_enum_list_item(self):
        '''
	   flush an enum list item
        '''
        if len(self.current_block) > 1 and self.current_block[-2] == 'Number_list' :
	    num_style = self.enum_prefix_style[-1]
            # change numbering
            if num_style[0] <  self.max_num_list_id :
                self.max_num_list_id += 1
                paras = self.docx.find_numbering_paragraph(num_style[0])
		num_style[1][0] = len(paras)+1
                num_style[0] = self.max_num_list_id

            self.flush_list_item(num_style[0], start_num=num_style[1][0], 
                                 enumprefix=num_style[1][1],
                                 enumprefixtype=num_style[1][2] )
        else:
            self.flush_list_item()

    def flush_list_item(self, typ=-1, enumprefix=None, enumprefixtype=None, start_num=1):
        '''
	   flush a list item
        '''
        dprint()
	text_list = get_items_list(self.states)

        b_level = self.list_level+self.block_level

        for i,x in enumerate(text_list) :
            sty = self.list_style[-1]

            if i == 0 and not self.list_item_flushed :
              self.docx.list_item(x, sty, b_level, typ,
			      enum_prefix=enumprefix, 
			      enum_prefix_type=enumprefixtype, start=start_num)
              self.list_item_flushed=True
	    else:
              self.docx.list_item(x, sty, b_level, 0)

        del self.states
	self.states = [[]]

    def append_style(self, style):
        '''
	   append a list style...
        '''
        dprint()
        txt_list = self.states.pop()
	txt = txt_list.pop()
	txt_list.append([txt, style])
        self.states.append(txt_list)


    def visit_start_of_file(self, node):
        '''
	   start of a file
        '''
        dprint()
        self.new_state()
        self.sectionlevel = 0

#        self.docx.pagebreak(type='page', orient='portrait')

    def depart_start_of_file(self, node):
        '''
	   end of a file
        '''
        dprint()
        self.end_state()

    def visit_document(self, node):
        '''
	   start of a document
        '''
        dprint()
        self.toc_out=False
        self.new_state()

    def depart_document(self, node):
        '''
	   end of a document
        '''
        dprint()
        self.end_state()

    def visit_highlightlang(self, node):
        '''
	   start of a hight light
        '''
        dprint()
        raise nodes.SkipNode

    def visit_section(self, node):
        '''
	   start of a section
        '''
        dprint()
        self.sectionlevel += 1

    def depart_section(self, node):
        '''
	   end of a section
        '''
        dprint()
        self.ensure_state()
        if self.sectionlevel > 0:
            self.sectionlevel -= 1

    def visit_topic(self, node):
        '''
	   start of a topic  (ignore)
        '''
        dprint()
        #raise nodes.SkipNode
        #self.new_state()

    def depart_topic(self, node):
        '''
	   end of a topic (ignore)
        '''
        dprint()
        #raise nodes.SkipNode
        #self.end_state()

    visit_sidebar = visit_topic
    depart_sidebar = depart_topic

    def visit_rubric(self, node):
        '''
	   start of a rubric  (ignore)
        '''
        dprint()
        #raise nodes.SkipNode
        #self.new_state()
        #self.add_text('-[ ')

    def depart_rubric(self, node):
        '''
	   end of a rubric  (ignore)
        '''
        dprint()
        raise nodes.SkipNode
        #self.add_text(' ]-')
        #self.end_state()

    def visit_compound(self, node):
        '''
	   start of a compound (pass a text)
        '''
	if self.states[-1][0]  == 'Contents:' :
	   self.states.pop()
	   self.states.append(['  '])
	  
        if not self.toc_out :
           self.toc_out = True
           self.ensure_state()
           maxdepth = get_toc_maxdepth(self.builder, 'index')
           self.docx.table_of_contents(toc_text='Contents', maxlevel=maxdepth )
           self.docx.pagebreak(type='page', orient='portrait')
        dprint()
        pass

    def depart_compound(self, node):
        '''
	   end of a compound (pass a text)
        '''
        dprint()
        pass

    def visit_glossary(self, node):
        '''
	  start of a glossary (pass a text)
        '''
        dprint()
        pass

    def depart_glossary(self, node):
        '''
	  end of a glossary (pass a text)
        '''
        dprint()
        pass

    def visit_title(self, node):
        '''
	  start of a title
        '''
        dprint()
        self.new_state()

    def depart_title(self, node):
        '''
	  end of a title
        '''
        dprint()
        text = self.states.pop()
        dprint(_func='* heading', text=repr(text), level=self.sectionlevel)

        if self.table is not None :
            self.docx.paragraph(text, style='TableHeading')
        else :
            self.docx.heading(text, self.sectionlevel)

    def visit_subtitle(self, node):
        '''
	  start of a subtitle (pass a text)
        '''
        dprint()
        pass

    def depart_subtitle(self, node):
        '''
	  end of a subtitle (pass a text)
        '''
        dprint()
        pass

    def visit_attribution(self, node):
        '''
	  start of a attribution (ignore)
        '''
        dprint()
        #raise nodes.SkipNode
        #self.add_text('-- ')

    def depart_attribution(self, node):
        '''
	  end of a attribution (ignore)
        '''
        dprint()
        pass

    def visit_desc(self, node):
        '''
	  start of a desc (pass a text)
        '''
        dprint()
        pass

    def depart_desc(self, node):
        '''
	  start of a desc (pass a text)
        '''
        dprint()
        pass

    def visit_desc_signature(self, node):
        '''
	  start of a desc signature (ignore)
        '''
        dprint()
        #raise nodes.SkipNode
        #self.new_state()
        #if node.parent['objtype'] in ('class', 'exception'):
        #    self.add_text('%s ' % node.parent['objtype'])

    def depart_desc_signature(self, node):
        '''
	  end of a desc signature (ignore)
        '''
        dprint()
        #raise nodes.SkipNode
        ## XXX: wrap signatures in a way that makes sense
        #self.end_state()

    def visit_desc_name(self, node):
        '''
	  start of a desc name (pass a text)
        '''
        dprint()
        pass

    def depart_desc_name(self, node):
        '''
	  end of a desc name (pass a text)
        '''
        dprint()
        pass

    def visit_desc_addname(self, node):
        '''
	  start of a desc addname (pass a text)
        '''
        dprint()
        pass

    def depart_desc_addname(self, node):
        '''
	  end of a desc addname (pass a text)
        '''
        dprint()
        pass

    def visit_desc_type(self, node):
        '''
	  start of a desc type (pass a text)
        '''
        dprint()
        pass

    def depart_desc_type(self, node):
        '''
	  end of a desc type (pass a text)
        '''
        dprint()
        pass

    def visit_desc_returns(self, node):
        '''
	  start of a desc returns (ignore)
        '''
        dprint()
        #raise nodes.SkipNode
        #self.add_text(' -> ')

    def depart_desc_returns(self, node):
        '''
	  end of a desc returns (ignore)
        '''
        dprint()
        pass

    def visit_desc_parameterlist(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text('(')
        #self.first_param = 1

    def depart_desc_parameterlist(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text(')')

    def visit_desc_parameter(self, node):
        dprint()
        #raise nodes.SkipNode
        #if not self.first_param:
        #    self.add_text(', ')
        #else:
        #    self.first_param = 0
        #self.add_text(node.astext())
        ##raise nodes.SkipNode

    def visit_desc_optional(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text('[')

    def depart_desc_optional(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text(']')

    def visit_desc_annotation(self, node):
        dprint()
        pass

    def depart_desc_annotation(self, node):
        dprint()
        pass

    def visit_refcount(self, node):
        dprint()
        pass

    def depart_refcount(self, node):
        dprint()
        pass

    def visit_desc_content(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.new_state()
        #self.add_text('\n')

    def depart_desc_content(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.end_state()

    def visit_figure(self, node):
        # FIXME: figure text become normal paragraph instead of caption.
        dprint()
        self.new_state()

    def depart_figure(self, node):
        dprint()
        self.end_state()

    def visit_caption(self, node):
        dprint()
        pass

    def depart_caption(self, node):
        self.flush_state('ImageCaption')
        dprint()
        pass

    def visit_productionlist(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.new_state()
        #names = []
        #for production in node:
        #    names.append(production['tokenname'])
        #maxlen = max(len(name) for name in names)
        #for production in node:
        #    if production['tokenname']:
        #        self.add_text(production['tokenname'].ljust(maxlen) + ' ::=')
        #        lastname = production['tokenname']
        #    else:
        #        self.add_text('%s    ' % (' '*len(lastname)))
        #    self.add_text(production.astext() + '\n')
        #self.end_state()
        ##raise nodes.SkipNode

    def visit_seealso(self, node):
        dprint()
        self.new_state()

    def depart_seealso(self, node):
        dprint()
        self.end_state(first='')

    def visit_footnote(self, node):
        dprint()
        #raise nodes.SkipNode
        #self._footnote = node.children[0].astext().strip()
        #self.new_state()

    def depart_footnote(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.end_state(first='[%s] ' % self._footnote)

    def visit_citation(self, node):
        dprint()
        #raise nodes.SkipNode
        #if len(node) and isinstance(node[0], nodes.label):
        #    self._citlabel = node[0].astext()
        #else:
        #    self._citlabel = ''
        #self.new_state()

    def depart_citation(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.end_state(first='[%s] ' % self._citlabel)

    def visit_label(self, node):
        dprint()
        #raise nodes.SkipNode

    # XXX: option list could use some better styling

    def visit_option_list(self, node):
        dprint()
	self.flush_state()
	self.current_option_list = self.docx.insert_option_list_table()
        pass

    def depart_option_list(self, node):
        dprint()
	self.current_option_list = None
        pass

    def visit_option_list_item(self, node):
        dprint()

    def depart_option_list_item(self, node):
        dprint()
        self.docx.insert_option_list_item(self.current_option_list, get_items_list(self.states), 0)
	self.states=[[]]

    def visit_option_group(self, node):
        dprint()

    def depart_option_group(self, node):
        dprint()
	if self.states[-1][-1] == ', ' :
	  self.states[-1].pop()
        self.docx.insert_option_list_item(self.current_option_list, get_items_list(self.states), 1)
	self.states=[[]]

    def visit_option(self, node):
        dprint()

    def depart_option(self, node):
        self.add_text(', ')
        dprint()
        pass

    def visit_option_string(self, node):
        dprint()
        pass

    def depart_option_string(self, node):
        dprint()
        pass

    def visit_option_argument(self, node):
        dprint()
	if self.states[-1][-1][:2] == '--' :
          self.add_text('=')
        else :
          self.add_text(' ')

    def depart_option_argument(self, node):
        dprint()
        pass

    def visit_description(self, node):
        dprint()
        pass

    def depart_description(self, node):
        dprint()
        pass

    def visit_tabular_col_spec(self, node):
        dprint()
        #raise nodes.SkipNode

    def visit_colspec(self, node):
        dprint()
        self.table[0].append(node['colwidth'])

    def depart_colspec(self, node):
        dprint()

    def visit_tgroup(self, node):
        dprint()
        pass

    def depart_tgroup(self, node):
        dprint()
        pass

    def visit_thead(self, node):
        dprint()
        pass

    def depart_thead(self, node):
        dprint()
        pass

    def visit_tbody(self, node):
        dprint()
        self.table.append('sep')

    def depart_tbody(self, node):
        dprint()
        pass

    def visit_row(self, node):
        dprint()
        self.table.append([])

    def depart_row(self, node):
        dprint()
        pass

    def visit_entry(self, node):
        dprint()
        if 'morerows' in node or 'morecols' in node:
            raise NotImplementedError('Column or row spanning cells are '
                                      'not implemented.')
        self.new_state()

    def depart_entry(self, node):
        dprint()
	text = self.states.pop()
        #text = '\n'.join('\n'.join(x) for x in self.states.pop())
        self.table[-1].append(text)

    def visit_table(self, node):
        dprint()
        if self.table:
            raise NotImplementedError('Nested tables are not supported.')
        self.new_state()
        self.table = [[]]

    def depart_table(self, node):
        dprint()
        colsize_chars = self.table[0]
        colsize = []
	for i,x in enumerate(colsize_chars):
          colsize.append( int(x)*110 )

        lines = self.table[1:]
        fmted_rows = []

        # don't allow paragraphs in table cells for now
        for line in lines:
            if line == 'sep':
                pass
            else:
                fmted_rows.append(line)

        self.docx.table(fmted_rows, colsize)
        self.docx.paragraph("")
        self.table = None
        self.end_state()

    def visit_acks(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.new_state()
        #self.add_text(', '.join(n.astext() for n in node.children[0].children)
        #              + '.')
        #self.end_state()
        #raise nodes.SkipNode

    def visit_image(self, node):
        dprint()
        self.flush_state()
        dprint(_func=' image ', uri=node.attributes['uri'])
        uri = node.attributes['uri']
        file_path = os.path.join(self.builder.env.srcdir, uri)
        width, height = self.get_image_scaled_width_height(node, file_path)

        self.docx.picture(file_path, '',width, height)

    def depart_image(self, node):
        dprint()

    def get_image_width_height(self, node, attr):
        size = None
        if attr in node.attributes:
          size = node.attributes[attr]
          if size[-1] == '%' :
            size = float(size[:-1])
            size = [size, '%']
          else:
            unit = size[-2:]
            if unit.isalpha():
                size = size[:-2]
            else:
                unit = 'px'
            try:
                size = float(size)
            except ValueError, e:
                self.document.reporter.warning(
                    'Invalid %s for image: "%s"' % (
                        attr, node.attributes[attr]))
            size = [size, unit]
        return size

    def get_image_scale(self, node):
        if 'scale' in node.attributes:
            try:
                scale = int(node.attributes['scale'])
                if scale < 1: # or scale > 100:
                    self.document.reporter.warning(
                        'scale out of range (%s), using 1.' % (scale, ))
                    scale = 1
                scale = scale * 0.01
            except ValueError, e:
                self.document.reporter.warning(
                    'Invalid scale for image: "%s"' % (
                        node.attributes['scale'], ))
        else:
            scale = 1.0
        return scale

    def get_image_scaled_width_height(self, node, filename):
        dpi = (72, 72)

        if Image is not None :
            try:
              imageobj = Image.open(filename, 'r')
            except:
              raise RuntimeError('Fail to open image file: %s' % filename)

            dpi = imageobj.info.get('dpi', dpi)
            # dpi information can be (xdpi, ydpi) or xydpi
            try: iter(dpi)
            except: dpi = (dpi, dpi)
        else:
            imageobj = None
            raise RuntimeError('image size not fully specified and PIL not installed')

        scale = self.get_image_scale(node)
        width = self.get_image_width_height(node, 'width')
        height = self.get_image_width_height(node, 'height')

        if width is not None and width[1] == '%':
           width = [int(self.docx.styleDocx.document_width * width[0] * 0.00284 ), 'px']

        if height is not None and height[1] == '%':
           height = [int(self.docx.styleDocx.document_height * height[0] * 0.00284 ), 'px']

        if width is None or height is None:
            if imageobj is None:
                raise RuntimeError(
                    'image size not fully specified and PIL not installed')
            if width is None:
                if height is None:
                     width = [imageobj.size[0], 'px']
                     height = [imageobj.size[1], 'px']
                else:
                     scaled_width = imageobj.size[0] * height[0] /imageobj.size[1]
                     width = [scaled_width, 'px']
            else:
                if height is None:
                     scaled_height = imageobj.size[1] * width[0] / imageobj.size[0]
                     height = [scaled_height, 'px']
                else:
                     height = [imageobj.size[1], 'px']

        width[0] *= scale
        height[0] *= scale
        if width[1] == 'in': width = [width[0] * dip[0], 'px']
        if height[1] == 'in': height = [height[0] *dip[1], 'px']

        #  We shoule shulink image (multiply 72/96)
        width[0] *= 0.75
        height[0] *= 0.75

        # 
        maxwidth = int(self.docx.styleDocx.document_width  * 0.284 ) * 0.9

        if width[0] > maxwidth :
            ratio = height[0] / width[0]
            width[0] = maxwidth
            height[0] = width[0] * ratio

        maxheight = int(self.docx.styleDocx.document_width  * 0.284 ) * 0.9
        if height[0] > maxheight :
            ratio = width[0] / height[0]
            height[0] = maxheight
            width[0] = height[0] * ratio

        return int(width[0]), int(height[0])

    def visit_transition(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.new_state()
        #self.add_text('=' * 70)
        #self.end_state()

    def visit_bullet_list(self, node):
        dprint()
	self.new_state()
        self.flush_state()

        self.current_block.append('Bullet_list')
	self.list_level += 1
        self.list_style.append('ListBullet')

    def depart_bullet_list(self, node):
        dprint()
        self.current_block.pop()
        self.list_style.pop()
	self.list_level -= 1

    def visit_enumerated_list(self, node):
        dprint()
        
	if self.flush_state() : 
	    self.num_list_id += 1
        
        self.new_state()
        suffix=""
        prefix=""
        enumtype="arabic"
        start=1
	try:
          enumtype=node['enumtype']
          suffix=node['suffix']
          prefix=node['prefix']
	except:
          pass
	try:
          start=node['start']
	except:
          pass
	enumprefix =  "%s%%1%s" % (prefix,suffix)
	enumprefix_type = enumtype
	start_num = start

	if self.current_block.count('Number_list') == 0 :
	  self.num_list_id = self.max_num_list_id+1

	self.enum_prefix_style.append([self.num_list_id, [start_num, enumprefix, enumprefix_type]])
        self.current_block.append('Number_list')
        self.list_style.append('ListNumber')
	self.list_level += 1
	self.max_num_list_id += 1

    def depart_enumerated_list(self, node):
        dprint()
	if self.current_block :
          self.current_block.pop()
	self.enum_prefix_style.pop()
        self.list_style.pop()
	self.list_level -= 1
	self.num_list_id -= 1
        #print "depart_enum", self.list_level

    def visit_definition_list(self, node):
        self.flush_state()
        dprint()
        ##raise nodes.SkipNode
        #self.list_style.append(-2)

    def depart_definition_list(self, node):
        dprint()
        ##raise nodes.SkipNode
        #self.list_style.pop()

    def visit_list_item(self, node):
        dprint()
        self.list_item_flushed=False
	self.current_block.append('List_item')
        self.new_state()

    def depart_list_item(self, node):
        dprint()
        self.flush_state(_sty='List_item')
	if self.current_block :
	  self.current_block.pop()
       
    def visit_definition_list_item(self, node):
        dprint()
        self.flush_state()
        pass


    def depart_definition_list_item(self, node):
        dprint()
        self.flush_state()

    def visit_term(self, node):
        dprint()
        self.flush_state()
        self.new_state()

    def depart_term(self, node):
        dprint()
        if len(self.current_block) > 0 and self.current_block[-1] != 'List_item' :
          self.flush_state('DefinitionTerm')

    def visit_classifier(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text(' : ')

    def depart_classifier(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.end_state()

    def visit_definition(self, node):
        dprint()
        self.flush_state()
	self.block_level += 1

    def depart_definition(self, node):
        dprint()
        self.flush_state()
	self.block_level -= 1

    def visit_field_list(self, node):
        dprint()
        self.flush_state()
	self.current_field_list = self.docx.insert_field_list_table()
        pass

    def depart_field_list(self, node):
        self.current_field_list = None
        dprint()
        pass

    def visit_field(self, node):
        dprint()
        pass

    def depart_field(self, node):
        dprint()
        pass

    def visit_field_name(self, node):
        dprint()

    def depart_field_name(self, node):
        dprint()
	self.add_text(':')
	self.docx.insert_field_list_item(self.current_field_list,self.states)
        self.states=[[]]

    def visit_field_body(self, node):
        dprint()

    def depart_field_body(self, node):
        dprint()
	lbody = self.docx.set_field_list_item(self.current_field_list, get_items_list(self.states), 1)
        self.states=[[]]

    def visit_centered(self, node):
        dprint()
        pass

    def depart_centered(self, node):
        dprint()
        pass

    def visit_hlist(self, node):
        dprint()
        pass

    def depart_hlist(self, node):
        dprint()
        pass

    def visit_hlistcol(self, node):
        dprint()
        pass

    def depart_hlistcol(self, node):
        dprint()
        pass

    def _visit_admonition(name):
        def visit_admonition(self, node):
            dprint()
            self.flush_state()

            atitle = admonitionlabels[name.lower()] + ': '
	    self.admonition_body = self.docx.insert_admonition_table('', title=atitle,tstyle=name+'Admonition')
	    self.docx.set_docbody(self.admonition_body)
        return visit_admonition

    def _make_depart_admonition(name):
        def depart_admonition(self, node):
            dprint()
            self.flush_state()
	    self.docx.set_docbody()
        return depart_admonition

    visit_attention = _visit_admonition('Attention')
    depart_attention = _make_depart_admonition('Attention')
    visit_caution = _visit_admonition('Caution')
    depart_caution = _make_depart_admonition('Caution')
    visit_danger = _visit_admonition('Danger')
    depart_danger = _make_depart_admonition('Danger')
    visit_error = _visit_admonition('Error')
    depart_error = _make_depart_admonition('Error')
    visit_hint = _visit_admonition('Hint')
    depart_hint = _make_depart_admonition('Hint')
    visit_important = _visit_admonition('Important')
    depart_important = _make_depart_admonition('Important')
    visit_note = _visit_admonition('Note')
    depart_note = _make_depart_admonition('Note')
    visit_tip = _visit_admonition('Tip')
    depart_tip = _make_depart_admonition('Tip')
    visit_warning = _visit_admonition('Warning')
    depart_warning = _make_depart_admonition('Warning')

    def visit_versionmodified(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.new_state()
        #if node.children:
        #    self.add_text(
        #            versionlabels[node['type']] % node['version'] + ': ')
        #else:
        #    self.add_text(
        #            versionlabels[node['type']] % node['version'] + '.')

    def depart_versionmodified(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.end_state()

    def visit_literal_block(self, node):
        # FIXME: working but broken.
        dprint()
        self.flush_state()
        self.new_state()
	try:
	  self.literal_block_lang = node['language']
	except:
	  self.literal_block_lang = 'guess'

    def depart_literal_block(self, node):
        dprint()
	if self.docx.get_last_paragraph_style() == 'LiteralBlock' :
          self.docx.insert_linespace()
        # We should insert highlighter for docx....

        highlight_args = node.get('highlight_args', {})
        def warner(msg):
            self.builder.warn(msg, (self.builder.current_docname, node.line))
        result = []
        for  x in self.states:
          linenos = 1
          if x :
            highlighted = self.highlighter.highlight_block(
                     x[0], self.literal_block_lang, # warn=warner,
                    linenos=linenos, **highlight_args)
            result.append([highlighted])

        self.states=result
        self.flush_state(_sty='LiteralBlock')
        self.end_state()

    def visit_doctest_block(self, node):
        dprint()

    def depart_doctest_block(self, node):
        dprint()

    def visit_line_block(self, node):
        dprint()
        self.line_block_level += 1

    def depart_line_block(self, node):
        dprint()
        self.line_block_level -= 1

    def visit_line(self, node):
        dprint()
        if self.line_block_level > 0 :
          for n in range(0, self.line_block_level):
             self.add_text(' ')
        pass

    def depart_line(self, node):
        dprint()
        self.add_linebreak()
        pass

    def visit_block_quote(self, node):
        dprint()
        self.flush_state()
        self.block_level += 1
        self.new_state()

    def depart_block_quote(self, node):
        dprint()
        self.flush_state()
        self.block_level -= 1
        self.end_state()

    def visit_compact_paragraph(self, node):
        dprint()
        pass

    def depart_compact_paragraph(self, node):
        dprint()
        pass

    def visit_paragraph(self, node):
        dprint()
        self.new_state()
        #self.ensure_state()
        #if not isinstance(node.parent, nodes.Admonition) or \
        #       isinstance(node.parent, addnodes.seealso):
        #    self.new_state()

    def depart_paragraph(self, node):
        dprint()
        #self.ensure_state()
        #if not isinstance(node.parent, nodes.Admonition) or \
        #       isinstance(node.parent, addnodes.seealso):
        #    self.end_state()

    def visit_target(self, node):
        dprint()
        raise nodes.SkipNode

    def visit_index(self, node):
        dprint()
        #raise nodes.SkipNode

    def visit_substitution_definition(self, node):
        dprint()
        #raise nodes.SkipNode

    def visit_pending_xref(self, node):
        dprint()
        pass

    def depart_pending_xref(self, node):
        dprint()
        pass

    def visit_reference(self, node):
        dprint()
        pass

    def depart_reference(self, node):
        dprint()
        pass

    def visit_download_reference(self, node):
        dprint()
        pass

    def depart_download_reference(self, node):
        dprint()
        pass

    def visit_emphasis(self, node):
        dprint()

    def depart_emphasis(self, node):
        dprint()
        self.append_style('Emphasis')

    def visit_literal_emphasis(self, node):
        dprint()

    def depart_literal_emphasis(self, node):
        dprint()
        self.append_style('LiteralEmphasise')

    def visit_strong(self, node):
        dprint()

    def depart_strong(self, node):
        dprint()
        self.append_style('Strong')

    def visit_abbreviation(self, node):
        dprint()

    def depart_abbreviation(self, node):
        dprint()
        self.append_style('Abbreviation')

    def visit_title_reference(self, node):
        dprint()
        #self.add_text('*')

    def depart_title_reference(self, node):
        dprint()
        self.append_style('TitleReference')
        #self.add_text('*')

    def visit_literal(self, node):
        dprint()
        #self.add_text('``')

    def depart_literal(self, node):
        dprint()
        self.append_style('Literal')
        #self.add_text('``')

    def visit_subscript(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text('_')

    def depart_subscript(self, node):
        dprint()
        self.append_style('Subscript')
        pass

    def visit_superscript(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text('^')

    def depart_superscript(self, node):
        dprint()
        self.append_style('Superscript')
        pass

    def visit_footnote_reference(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text('[%s]' % node.astext())

    def visit_citation_reference(self, node):
        dprint()
        #raise nodes.SkipNode
        #self.add_text('[%s]' % node.astext())

    def visit_Text(self, node):
        dprint()
        self.add_text(node.astext())

    def depart_Text(self, node):
        dprint()
        pass

    def visit_generated(self, node):
        dprint()
        pass

    def depart_generated(self, node):
        dprint()
        pass

    def visit_inline(self, node):
        classes = node.get('classes', [])
        dprint()
        pass

    def depart_inline(self, node):
        dprint()
        pass

    def visit_problematic(self, node):
        dprint()

    def depart_problematic(self, node):
        dprint()
        self.append_style('Problematic')

    def visit_system_message(self, node):
        dprint()
        raise nodes.SkipNode
        #self.new_state()
        #self.add_text('<SYSTEM MESSAGE: %s>' % node.astext())
        #self.end_state()

    def visit_comment(self, node):
        dprint()
        raise nodes.SkipNode

    def visit_meta(self, node):
        dprint()
        raise nodes.SkipNode
        # only valid for HTML

    def visit_raw(self, node):
        dprint()
        raise nodes.SkipNode
        #if 'text' in node.get('format', '').split():
        #    self.body.append(node.astext())

    def visit_graphviz(self, node):
        dprint()
	fname, filename = graphviz.render_dot(self, node['code'], node['options'],'png')
        self.flush_state()
        width, height = self.get_image_scaled_width_height(node, filename)
        self.docx.picture(filename, '',width, height)
        raise nodes.SkipNode

    def unknown_visit(self, node):
        dprint()
        print node
        raise nodes.SkipNode
        #raise NotImplementedError('Unknown node: ' + node.__class__.__name__)
