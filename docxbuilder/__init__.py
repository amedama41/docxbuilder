from builder import DocxBuilder


def setup(app):
    app.add_builder(DocxBuilder)
    app.add_config_value('docx_style', 'style.docx', 'env')
    app.add_config_value('docx_title', 'SphinxDocx', 'env')
    app.add_config_value('docx_subject', 'Sphinx Document', 'env')
    app.add_config_value('docx_creator', 'sphinx-docxbuilder', 'env')
    app.add_config_value('docx_company', '', 'env')
    app.add_config_value('docx_category', 'sphinx document', 'env')
    app.add_config_value('docx_descriptions', 'This document generaged by sphix-docxbuilder', 'env')
    app.add_config_value('docx_keywords', ['python', 'Office Open XML', 'Word'] , 'env')
    app.add_config_value('docx_coverpage', True, 'env')

