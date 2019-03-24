###########
docxbuilder
###########

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

Configuration
=============

.. list-table::
   :header-rows: 1

   * - variable
     - meaning
     - default
   * - docx_documents
     - This determines how to group the document tree by a list of tuples,
       like `latex_documents`_.
       The tuple is root document file, generated docx file name, document
       properties, and toctree_only flag.
     - The root file, docx name, and properties are generated based on other
       configurations. toctree_only is ``False``.
   * - docx_style
     - A path to a style file. If this value is an empty string, default
       style file is used.
     - ``''``. Use default style.
   * - docx_coverpage
     - If this value is true, the coverpage of the style file is inserted
       to generated documents.
     - ``True``
   * - docx_pagebreak_before_section
     - Specify a section level. Before each sections which level is larger
       than or equal to this option value, a page break is inserted.
     - ``0``. No page break is inserted.
   * - docx_pagebreak_after_table_of_contents
     - Specify a section level. After each table of contents which appears
       in section level larger than or equal to this option value,
       a page break is inserted.
     - ``0``. Page break is inserted only before first section.
   * - docx_table_options
     - Table arrangement option. Specify a dictionary with bellow keys.

       landscape_columns
         Tables with the number of columns equal to or larger than this option
         value, are arranged on landscape pages.
       in_single_page
         If this value is true, each table is arranged in single page as possible.
       row_splittable
         If this value is false, each table row shall not be arranged across
         multiple pages.
       header_in_all_page
         If this value is true and a table is arranged across multiple pages,
         the header is displayed in the each pages.
     - :landscape_columns: ``0``
       :in_single_page: ``False``
       :row_splittable: ``True``
       :header_in_all_page: ``False``

.. _`latex_documents`: http://www.sphinx-doc.org/en/master/usage/configuration.html#confval-latex_documents

The below code is a configuration example:

.. code:: python

   docx_documents = [
       (master_doc, 'docxbuilder.docx', {
            'title': project,
            'creator': author,
            'subject': 'A manual of docxbuilder',
        }, True),
   ]
   docx_style = 'path/to/custom_style.docx'
   docx_pagebreak_before_section = 1
   docx_pagebreak_after_table_of_contents = 0
   docx_table_options = {
           'landscape_columns': 6,
           'in_single_page': False,
           'row_splittable': True,
           'header_in_all_page': False,
   }

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

