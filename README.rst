###########
docxbuilder
###########

.. image:: https://readthedocs.org/projects/docxbuilder/badge/?version=latest
   :target: https://docxbuilder.readthedocs.io/en/latest/?badge=latest
   :alt: Documentation Status

Docxbuilder is a Sphinx extension to build docx formatted documents.

.. note::

   This extension is developed based on `sphinx-docxbuilder`_. Though,
   there is no compatibility between these extensions.

.. _`sphinx-docxbuilder`: https://bitbucket.org/haraisao/sphinx-docxbuilder/

************
Requirements
************

:Python: 2.7, 3.5 or latter
:Sphinx: 1.7.6 or later

*******
Install
*******

Use pip::

   pip install docxbuilder

*****
Usage
*****

Add 'docxbuilder' to ``extensions`` configuration of **conf.py**:

.. code:: python

   extensions = ['docxbuilder']

and build your documents::

   make docx

You can control the generated document by adding configurations into ``conf.py``:

.. code:: python

   docx_documents = [
       ('index', 'docxbuilder.docx', {
            'title': project,
            'creator': author,
            'subject': 'A manual of docxbuilder',
        }, True),
   ]
   docx_style = 'path/to/custom_style.docx'
   docx_pagebreak_before_section = 1

For more details, see `the documentation <https://docxbuilder.readthedocs.io/en/latest/>`_.

Style file
==========

Generated docx file's design is customized by a style file
(The default style is ``docxbuilder/docx/style.docx``).
The style file is a docx file, which defines some paragraph,
character, and table styles.

The below lists shows typical styles.

Character styles:

* Emphasis
* Strong
* Literal
* Hyperlink
* Footnote Reference

Paragraph styles:

* Body Text
* Footnote Text
* Definition Term
* Literal Block
* Image Caption, Table Caution, Literal Caption
* Heading 1, Heading 2, ..., Heading *N*
* TOC Heading
* toc 1, toc 2, ..., toc *N*
* List Bullet
* List Number

Table styles:

* Table
* Field List
* Admonition Note

****
TODO
****

- Support math role and directive.
- Support tabular_col_spec directive.
- Support URL path for images.

*******
Licence
*******

MIT Licence

