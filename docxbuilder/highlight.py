from xml.sax import saxutils
from pygments.formatter import Formatter
from sphinx.highlighting import PygmentsBridge

class DocxFormatter(Formatter):
    def __init__(self, **options):
        super(DocxFormatter, self).__init__(**options)
        self.linenos = options.get('linenos', False)
        self.hl_lines = options.get('hl_lines', [])
        self.linenostart = options.get('linenostart', 1)
        self.trim_last_line_break = options.get('trim_last_line_break', False)
        self.color_mapping = {}
        for _, style in self.style:
            for color in style['color'], style['bgcolor'], style['border']:
                if color and color not in self.color_mapping:
                    self.color_mapping[color] = r'%02x%02x%02x' % (
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
                value = saxutils.escape(value)
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

        if self.trim_last_line_break and lines[-1] == []:
            lines.pop()

        if self.linenos:
            self.output_as_table_with_linenos(outfile, lines)
        else:
            self.output_as_paragraph(outfile, lines)

    def output_as_paragraph(self, outfile, lines):
        outfile.write(
                '<w:p xmlns:w='
                '"http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
                '>')
        for lineno, tokens in enumerate(lines, 1):
            self.output_line(outfile, lineno, tokens)
            if lineno != len(lines):
                outfile.write(r'<w:r><w:br /></w:r>')
        outfile.write('</w:p>')

    def output_as_table_with_linenos(self, outfile, lines):
        outfile.write(
                '<w:tbl xmlns:w='
                '"http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
                '>')
        for lineno, tokens in enumerate(lines, 1):
            outfile.write('<w:tr>')
            outfile.write('<w:tc><w:p>')
            outfile.write(
                '<w:r><w:t>%d</w:t></w:r>' % (self.linenostart + lineno - 1))
            outfile.write('</w:p></w:tc>')
            outfile.write('<w:tc><w:p>')
            self.output_line(outfile, lineno, tokens)
            outfile.write('</w:p></w:tc>')
            outfile.write('</w:tr>')
        outfile.write('</w:tbl>')

    def output_line(self, outfile, lineno, tokens):
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

class DocxPygmentsBridge(PygmentsBridge):
    def __init__(self, dest, stylename, trim_doctest_flags):
        PygmentsBridge.__init__(self, dest, stylename, trim_doctest_flags)
        self.formatter = DocxFormatter

    def highlight_block(self, source, lang, *args, **kwargs):
        # highlight_block may append a line break to the tail of the code
        kwargs['trim_last_line_break'] = not source.endswith('\n')
        return super(DocxPygmentsBridge, self).highlight_block(
                source, lang, *args, **kwargs)
