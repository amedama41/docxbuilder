from sphinx.util.osutil import make_filename
from docxbuilder.builder import DocxBuilder


def setup(app):
    app.add_builder(DocxBuilder)

    def default_docx_documents(conf):
        start_doc = conf.master_doc
        filename = '%s.docx' % make_filename(conf.project)
        title = conf.project
        # author configuration value is available from Sphinx 1.8
        author = getattr(conf, 'author', 'sphinx-docxbuilder')
        properties = {
                'title': title,
                'creator': author,
                'subject': '',
                'category': '',
                'description': 'This document generaged by sphix-docxbuilder',
                'keywords': ['python', 'Office Open XML', 'Word'],
        }
        toc_only = False
        return [(start_doc, filename, properties, toc_only)]

    app.add_config_value('docx_documents', default_docx_documents, 'env')
    app.add_config_value('docx_style', '', 'env')
    app.add_config_value('docx_pagebreak_before_section', 0, 'env')
    app.add_config_value('docx_pagebreak_after_table_of_contents', 0, 'env')
    app.add_config_value('docx_coverpage', True, 'env')
    app.add_config_value('docx_table_options', {
        'landscape_columns': 0,
        'in_single_page': False,
        'row_splittable': True,
        'header_in_all_page': False,
    }, 'env')
    app.add_config_value('docx_nested_character_style', True, 'env')
