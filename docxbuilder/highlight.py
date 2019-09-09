from xml.sax import saxutils
from pygments.formatter import Formatter
from sphinx.highlighting import PygmentsBridge

# Not use achromatic colors, which are unsuitable for highlight
HIGHLIGHT_COLOR_MAP = {
    # 'black': [0x00, 0x00, 0x00],
    'blue': [0x00, 0x00, 0xFF],
    'cyan': [0x00, 0xFF, 0xFF],
    'darkBlue': [0x00, 0x00, 0x8B],
    'darkCyan': [0x00, 0x8B, 0x8B],
    # 'darkGray': [0xA9, 0xA9, 0xA9],
    'darkGreen': [0x00, 0x64, 0x00],
    'darkMagenta': [0x80, 0x00, 0x80],
    'darkRed': [0x8B, 0x00, 0x00],
    'darkYellow': [0x80, 0x80, 0x00],
    'green': [0x00, 0xFF, 0x00],
    # 'lightGray': [0xD3, 0xD3, 0xD3],
    'magenta': [0xFF, 0x00, 0xFF],
    'red': [0xFF, 0x00, 0x00],
    # 'white': [0xFF, 0xFF, 0xFF],
    'yellow': [0xFF, 0xFF, 0x00],
}

def get_highlight_color_name(hex_highlight_color):
    """Get color name nearest from the argument.
    """
    highlight_rgb = [0, 0, 0]
    for idx in range(3):
        highlight_rgb[idx] = int(
            hex_highlight_color[idx * 2 + 1:idx * 2 + 3], 16)
    def dist(rgb):
        return sum((c1 - c2) ** 2 for c1, c2 in zip(rgb, highlight_rgb))
    color_name, _ = min(
        ((name, dist(rgb)) for name, rgb in HIGHLIGHT_COLOR_MAP.items()),
        key=lambda name_and_dist: name_and_dist[1])
    return color_name

class DocxFormatter(Formatter):
    def __init__(self, **options):
        super(DocxFormatter, self).__init__(**options)
        self.linenos = options.get('linenos', False)
        self.hl_lines = options.get('hl_lines', [])
        self.linenostart = options.get('linenostart', 1)
        self.trim_last_line_break = options.get('trim_last_line_break', False)
        self.highlight = get_highlight_color_name(self.style.highlight_color)

    def format_unencoded(self, tokensource, outfile):
        # pylint: disable=too-many-branches
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
                    buf.append(r'<w:shd w:themeFill="%s" />' % style['bgcolor'])
                if style['color']:
                    buf.append(r'<w:color w:val="%s" />' % style['color'])
                if style['bold']:
                    buf.append(r'<w:b />')
                if style['italic']:
                    buf.append(r'<w:i />')
                if style['underline']:
                    buf.append(r'<w:u />')
                if style['border']:
                    buf.append(r'<w:bdr w:val="single" w:space="0" w:color="%s" />' %
                               style['border'])

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
        outfile.write(
            '<w:pPr>'
            '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>'
            '</w:pPr>' % self.style.background_color[1:7])
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
        bgcolor = self.style.background_color[1:7]
        for lineno, tokens in enumerate(lines, 1):
            outfile.write('<w:tr>')
            outfile.write('<w:tc><w:p>')
            outfile.write('<w:pPr><w:shd w:val="clear"/></w:pPr>')
            outfile.write(
                '<w:r><w:t>%d</w:t></w:r>' % (self.linenostart + lineno - 1))
            outfile.write('</w:p></w:tc>')
            outfile.write('<w:tc><w:p>')
            outfile.write(
                '<w:pPr>'
                '<w:shd w:val="clear" w:color="auto" w:fill="%s"/>'
                '</w:pPr>' % bgcolor)
            self.output_line(outfile, lineno, tokens)
            outfile.write('</w:p></w:tc>')
            outfile.write('</w:tr>')
        outfile.write('</w:tbl>')

    def output_line(self, outfile, lineno, tokens):
        for text, style in tokens:
            outfile.write(r'<w:r>')
            if lineno in self.hl_lines:
                style += r'<w:highlight w:val="%s" />' % self.highlight
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
    def __init__(self, dest, stylename, trim_doctest_flags=None):
        if trim_doctest_flags is not None:
            PygmentsBridge.__init__(self, dest, stylename, trim_doctest_flags)
        else:
            PygmentsBridge.__init__(self, dest, stylename)
        self.formatter = DocxFormatter

    def highlight_block(self, source, lang, *args, **kwargs):
        # pylint: disable=arguments-differ
        # highlight_block may append a line break to the tail of the code
        kwargs['trim_last_line_break'] = not source.endswith('\n')
        return super(DocxPygmentsBridge, self).highlight_block(
            source, lang, *args, **kwargs)
