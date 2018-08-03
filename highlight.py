
from sphinx import highlighting
from sphinx.highlighting import PygmentsBridge

from pygments.formatter import Formatter
from pygments.formatters import *

#--- Formatter
class DocxFormatter(RtfFormatter):
  def __init__(self, **options):
        RtfFormatter.__init__(self, **options)
        self.color_mapping = {}
        for _, style in self.style:
            for color in style['color'], style['bgcolor'], style['border']:
                if color and color not in self.color_mapping:
                    self.color_mapping[color] = r'%x%x%x' % (
                        int(color[0:2], 16),
                        int(color[2:4], 16),
                        int(color[4:6], 16)
                    )

  def format_unencoded(self, tokensource, outfile):
        for ttype, value in tokensource:
          if value == '\n':
            outfile.write(r'<w:r><w:br /></w:r>')
          else:
            outfile.write(r'<w:r>')
            while not self.style.styles_token(ttype) and ttype.parent:
                ttype = ttype.parent
            style = self.style.style_for_token(ttype)
            buf = []
            if style['bgcolor']:
                buf.append(r'<w:shd w:themeFill="%s" />' % self.color_mapping[style['bgcolor']])
            if style['color']:
		    buf.append(r'<w:color w:val="%s" />' % self.color_mapping[style['color']])
            if style['bold']:
                buf.append(r'<w:b />')
            if style['italic']:
                buf.append(r'<w:i />')
            if style['underline']:
                buf.append(r'<w:u />')
            if style['border']:
		    buf.append(r'<w:bdr w:val="single" w:space="0" w:color="%s" />' % self.color_mapping[style['border']])

            start = ''.join(buf)
            if start:
                outfile.write('<w:rPr>%s</w:rPr> ' % start)
            vals = value.split('\n')
            for i,txt in enumerate(vals) :
                if txt.find(' ') != -1 :
                    outfile.write(r'<w:t xml:space="preserve">')
                else: 
	            outfile.write(r'<w:t>')
                txt=txt.replace('<','&lt;')
                txt=txt.replace('>','&gt;')
                outfile.write(txt)
	        outfile.write(r'</w:t>')
                if i < len(vals) - 1 :
                    outfile.write(r'<w:br />')
	    outfile.write(r'</w:r>')



#--- PygmentsBridge
class DocxPygmentsBridge(PygmentsBridge) :
   def __init__(self, dest='docx', stylename='sphinx',
                 trim_doctest_flags=False):
    PygmentsBridge.__init__(self, dest, stylename, trim_doctest_flags)
    dest = "html"
    self.formatter = DocxFormatter

