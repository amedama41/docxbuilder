
from sphinx.highlighting import PygmentsBridge

from pygments.formatters import RtfFormatter

#--- Formatter


class DocxFormatter(RtfFormatter):
    def __init__(self, **options):
        RtfFormatter.__init__(self, **options)
        self.hl_lines = options.get('hl_lines', [])
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
        lines = [[]]
        for ttype, value in tokensource:
            if value == '\n':
                lines.append([])
            else:
                while not self.style.styles_token(ttype) and ttype.parent:
                    ttype = ttype.parent
                style = self.style.style_for_token(ttype)
                buf = []
                if style['bgcolor']:
                    buf.append(r'<w:shd w:themeFill="%s" />' %
                               self.color_mapping[style['bgcolor']])
                if style['color']:
                    buf.append(r'<w:color w:val="%s" />' %
                               self.color_mapping[style['color']])
                if style['bold']:
                    buf.append(r'<w:b />')
                if style['italic']:
                    buf.append(r'<w:i />')
                if style['underline']:
                    buf.append(r'<w:u />')
                if style['border']:
                    buf.append(r'<w:bdr w:val="single" w:space="0" w:color="%s" />' %
                               self.color_mapping[style['border']])

                style = ''.join(buf)
                value = value.replace('<', '&lt;')
                value = value.replace('>', '&gt;')
                index = 0
                while index < len(value):
                    idx = value.find('\n', index)
                    if idx == -1:
                        lines[-1].append((value[index:], style))
                        break
                    else:
                        lines[-1].append((value[index:idx], style))
                        lines.append([])
                        index = idx + 1

        for lineno, tokens in enumerate(lines, 1):
            for text, style in tokens:
                outfile.write(r'<w:r>')
                if lineno in self.hl_lines:
                    style += r'<w:highlight w:val="yellow" />' # TODO: color
                if style:
                    outfile.write(r'<w:rPr>%s</w:rPr>' % style)
                if text.find(' ') != -1:
                    outfile.write(r'<w:t xml:space="preserve">')
                else:
                    outfile.write(r'<w:t>')
                outfile.write(text)
                outfile.write(r'</w:t>')
                outfile.write(r'</w:r>')
            if lineno != len(lines):
                outfile.write(r'<w:r><w:br /></w:r>')



#--- PygmentsBridge
class DocxPygmentsBridge(PygmentsBridge):
    def __init__(self, dest='docx', stylename='sphinx',
                 trim_doctest_flags=False):
        PygmentsBridge.__init__(self, dest, stylename, trim_doctest_flags)
        self.formatter = DocxFormatter
