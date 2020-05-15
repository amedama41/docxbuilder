Customization of the document
=============================

Docxbuilder provides two ways to customize generated documents.
The one is style file, and another is class based customization.

Style file
----------

The generated document inherits some properties from the style file,
which may be referenced by ``docx_style`` configuration.
The inherited properties are

* styles of the document contents,
* cover page, and,
* section settings (size, margins, borders, header, footer, etc.).

The contents in the style file are ignored.

If you want to change these properties, you must create a new style file.
The lists of styles used by Docxbuilder, see :numref:`document_style_section`.
About cover page, see :numref:`coverpage_section`.
About section settings, see :numref:`section_settings_section`.

.. _`document_style_section`:

Document style
^^^^^^^^^^^^^^

In OpenXML, there are multiple style types.
Docxbuilder uses character, paragraph, and table styles.
The character styles are described in :numref:`character_style_table`,
the paragraph styles are described in :numref:`paragraph_style_table`,
and the table styles are described in :numref:`table_style_table`.

.. list-table:: Character Style
   :header-rows: 1
   :stub-columns: 1
   :widths: auto
   :width: 100%
   :align: center
   :name: character_style_table

   * - Style name
     - Example
   * - Emphasis
     - This is *Emphasis*.
   * - Strong
     - This is **Strong**.
   * - Literal
     - This is ``Literal``.
   * - Hyperlink
     - .. _`hyper_link_example`:

       This is :ref:`Hyperlink <hyper_link_example>`.
   * - Superscript
     - This is :sup:`Superscript`.
   * - Subscript
     - This is :sub:`Subscript`.
   * - Problematic
     - This is |Problematic|\ [#Problematic]_.
   * - Title Reference
     - This is :title:`Title Reference`.
   * - Abbreviation
     - This is :abbr:`Abbr (Abbreviation)`.
   * - Footnote Reference
     - .. _`footnote_reference_example`:

       The label attached to this statement is Footnote Reference\ [#FootnoteExample]_.
   * - Option Argument
     - See :ref:`Option List <option_list_example>`.
       This style is same as Emphasis by default.
   * - Versionmodified
     - See :ref:`Admonition Versionadded <versionadded_example>`.
       This style is same as Emphasis by default.
   * - Desc Name
     - See :ref:`Function Descriptions <function_descriptions_example>`.
       This style is same as Strong by default.
   * - Desc Annotation
     - See :ref:`Function Descriptions <function_descriptions_example>`.
       This style is same as Emphasis by default.

.. list-table:: Paragraph Style
   :header-rows: 1
   :stub-columns: 1
   :widths: auto
   :width: 100%
   :align: center
   :name: paragraph_style_table

   * - Style name
     - Example
   * - Body Text
     - Style for default paragraph.
   * - Footnote Text
     - Style for footnote.
       See :ref:`Footnote of Footnote Reference <footnote_reference_example>`.
   * - Bibliography
     - .. [BIB] This is Bibliography.
   * - | Definition Term
       | Definition
     - This is Definition Term : classifier one : classifier two
           This is Definition.
       This is also Definition Term
           This is also Definition.
   * - | Literal Caption
       | Literal Block
     - .. code-block:: guess
          :caption: This is Literal Caption

          This is Literal Block
   * - Math Block
     - .. math::

          (a + b)^2 = a^2 + 2ab + b^2
   * - Image
     - .. image:: images/sample.png
   * - | Figure
       | Image Caption
       | Legend
     - .. figure:: images/sample.png
          :figwidth: 70%
          :align: center

          This is Image Caption

          This is Legend.
   * - Table Caption
     - .. list-table:: This is Table Caption
          :widths: auto
          :align: center

          * - \ 
   * - Heading 1, Heading 2, ..., Heading *N*
     - Styles for section heading.
   * - TOC Heading
     - Style for title for table of contents.
   * - Rubric Title Heading
     - .. rubric:: This is Rubric Title Heading.
   * - Topic Title Heading
     - Style for topic directive's title.
   * - Sidebar Title Heading
     - Style for sidebar directive's title.
   * - Sidebar Subtitle Heading
     - Style for sidebar directive's subtitle.
   * - toc 1, toc 2, ..., toc *N*
     - Style for table of contents.
   * - Transition
     - Style for transition.
   * - List Bullet
     - * item 1

         * nested-item 1
         * nested-item 2
       * item 2
   * - List Number
     - #. item 1

          (i) nested-item 1
          (#) nested-item 2
       #. item 2

.. tabularcolumns:: |C|C|

.. list-table:: Table Style
   :header-rows: 1
   :stub-columns: 1
   :widths: auto
   :width: 100%
   :align: center
   :name: table_style_table

   * - Style name
     - Example
   * - Table
     - Style for standard table.
   * - Field List
     - :field 1: description 1
       :field 2: description 2
   * - Option List
     - .. _`option_list_example`:

       --option1=arg    option description 1
       --option2, -o    option description 2
   * - Horizontal List
     - .. hlist::

          * item1
          * item2
          * item3
          * item4
   * - Admonition
     - .. admonition:: This is Admonition

          Contents of admonition
   * - Admonition Note
     - .. note:: This is Admonition Note
   * - Admonition Warning
     - .. warning:: This is Admonition Warning
   * - Admonition Caution
     - .. caution:: This is Admonition Caution
   * - Admonition Seealso
     - .. seealso:: This is Admonition Seealso
   * - Admonition Versionadded
     - .. _`versionadded_example`:

       .. versionadded:: 1.0
          This is Admonition Versionadded
   * - Function Descriptions
     - .. _`function_descriptions_example`:

       .. c:function:: int func(int param1, double param2)

          Descriptions of func

.. rubric:: Style automatic generation

If some styles are not defined in the style file,
Docxbuilder automatically generate the styles from other defined styles.
:numref:`based_paragraph_style_figure` and :numref:`based_table_style_figure`
represents which style is generated from which style.

.. graphviz::
   :caption: Generation relationship for paragraph styles
   :name: based_paragraph_style_figure
   :align: center

   digraph ParagraphStyleHierarchy {
      rankdir="RL";
      ratio=0.9;
      Normal [style=bold];
      BodyText [label="Body Text"];
      FootnoteText [label="Footnote Text"];
      Bibliography;
      DefinitionTerm [label="Definition Term"];
      Definition;
      LiteralBlock [label="Literal Block"];
      MathBlock [label="Math Block"];
      Figure;
      Legend;
      Caption [style=bold];
      Heading [style=bold];
      HeadingN [label=<Heading <I>N</I>>];
      TitleHeading [style=bold, label="Title Heading"];
      SubtitleHeading [style=bold, label="Subtitle Heading"];
      TOCHeading [label="TOC Heading"];
      RubricTitleHeading [label="Rubric Title Heading"];
      TopicTitleHeading [label="Topic Title Heading"];
      SidebarTitleHeading [label="Sidebar Title Heading"];
      SidebarSubtitleHeading [label="Sidebar Subtitle Heading"];
      TableCaption [label="Table Caption"];
      ImageCaption [label="Image Caption"];
      LiteralCaption [label="Literal Caption"];
      BodyText       -> Normal;
      FootnoteText   -> Normal;
      Bibliography   -> Normal;
      DefinitionTerm -> Normal;
      Definition     -> Normal;
      LiteralBlock   -> Normal;
      MathBlock      -> Normal;
      Figure         -> Normal;
      Legend         -> Normal;
      Caption   -> Normal;
      Heading   -> Normal;
      HeadingN        -> Heading;
      TitleHeading    -> Heading;
      SubtitleHeading -> Heading;
      TOCHeading           -> TitleHeading;
      RubricTitleHeading   -> TitleHeading;
      TopicTitleHeading    -> TitleHeading;
      SidebarTitleHeading  -> TitleHeading;
      SidebarSubtitleHeading -> SubtitleHeading;
      TableCaption   -> Caption;
      ImageCaption   -> Caption;
      LiteralCaption -> Caption;
   }

.. graphviz::
   :caption: Generation relationship for table styles
   :name: based_table_style_figure
   :align: center

   digraph TableStyleHierarchy {
      rankdir="RL";
      ratio=0.9;
      NormalTable [style=bold, label="Normal Table"];
      ListTable [style=bold, label="List Table"];
      Table;
      BasedAdmonition [style=bold];
      FieldList [label="Field List"];
      OptionList [label="Option List"];
      Admonition;
      AdmonitionDescriptions [style=bold, label="Admonition Descriptions"];
      AdmonitionVersionmodified [style=bold, label="Admonition Versionmodified"];
      AnyAdmonition [label=<Admonition <I>XXX</I>>];
      AnyDescriptions [label=<<I>XXX</I> Descriptions>];
      AnyVersionmodified [label=<Admonition <I>YYY</I>>];
      ListTable       -> NormalTable;
      Table           -> NormalTable;
      BasedAdmonition -> NormalTable;
      FieldList  -> ListTable;
      OptionList -> ListTable;
      Admonition                -> BasedAdmonition;
      AdmonitionDescriptions    -> BasedAdmonition;
      AdmonitionVersionmodified -> BasedAdmonition;
      AnyAdmonition             -> BasedAdmonition;
      AnyDescriptions   -> AdmonitionDescriptions;
      AnyVersionmodified    -> AdmonitionVersionmodified;
   }

.. rubric:: Footnotes

.. [#Problematic]
    The Problematic style is used only when some errors exists in documents
    (e.g. using non-exsistence cross reference, unknown rorles).
    Then it is almost unnecessary to define this style.
.. [#FootnoteExample] This is Footnote Text.

.. _`user_defined_styles_section`:

User defined styles
^^^^^^^^^^^^^^^^^^^

In addition to above styles, you can define your original styles.
These styles are applied to elements with the corresponding class name
The mapping from class name to original style are defined by `docx_style_names` configuration.

.. code-block:: python

   docx_style = 'path/to/custom-style.docx'
   docx_style_names = {
      'strike': 'Strike',
      'custom-table': 'Custom Table',
   }
   # And define Strike and Custom Table styles in the style file specified docx_style

The following reStructuredText show how to use the custom styles.

.. code-block:: rst

   .. Use role to specify character class name.
   .. role:: strike

   This :strike:`text` is striked.

   .. list-table:: Custom style table
      :class: custom-table

      * - Row1: Col1
        - Row1: Col2
      * - Row2: Col1
        - Row2: Col2

.. warning:: Currently, only table elements and character elements are enable to be applied user defined styles.

.. _`coverpage_section`:

Cover page
^^^^^^^^^^

If ``docx_coverpage`` is true, the cover page of the style file is inserted into the generated document.
Docxbuilder treat the first structured document tag with "Cover Pages" docPartGallery as the cover page.
If no tag is found, the contents far to the first section break are used as the cover page.
If no section break is found, the contents far to the first page break are used as the cover page.

.. topic:: How to create structured document tag with "Cover Pages" docPartGallery

   It seems that Office Word can not create only structured document tag.
   Therefore, if you want to create your original cover page, you must insert
   a pre designed cover page and then modify the cover page.

.. _`section_settings_section`:

Section settings
^^^^^^^^^^^^^^^^

The generated document inherits the section settings from the style file.
The settings includes header, footer, page size, page margins, page borders, and so on.

If the style file includes multiple sections, Docxbuilder apply the first section.
If you want to apply other section from the middle of the document,
use :ref:`Docxbuilder custom class <class_based_customization_section>`.
In the bellow example, section A and C use the first section settings,
and section B uses the second section settings.

.. code-block:: rst
   :caption: Example to specify section settings

   Section A
   =========

   contents

   .. Use 2nd section from the next section
   .. rst-class:: docx-section-portrait-1

   Section B
   =========

   contents

   .. Use 1st section from the next section
   .. rst-class:: docx-section-portrait-0

   Section C
   =========

   contents

.. _`class_based_customization_section`:

Class based customization
-------------------------

Docxbuilder provides class based customization.
Elements with special classes which has "docx-" prefix, are arranged based on the specified class by Docxbuilder.

In the bellow example, the table is arranged in landscape page.
This is useful for tables with many columns, or horizontally long figures.

.. code-block:: rst

   .. csv-table::
      :class: docx-landscape

      A,B,C,D,E,F,G,H,I,J,K,L,M,N,O,P,Q,R,S,T,U,V,W,X,Y,Z
      1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26

:numref:`custom_class_table` shows the list of custom classes.
For each class, kinds of elements enable to be applied the class are defined.

.. _`custom_class_table`:

.. list-table:: Speciall custom class list
   :header-rows: 1
   :stub-columns: 1
   :align: center
   :widths: auto

   * - Class
     - Target
     - Description
   * - | docx-section-portrait-*N*
       | docx-section-landscape-*N*
     - section
     - Use *N*\th portrait or landscape section in style file from the section.
       *N* of the first section is 0.
   * - docx-rotation-header-*N*
     - table
     - Rotate the table header and the height is *N*\% of the width.
   * - | docx-landscape
       | docx-no-landscape
     - figure, table
     - Arrange the figure or table in landscape page, or not.
   * - | docx-in-single-page
       | docx-no-in-single-page
     - table
     - Arrange the table in single page as much as possible, or not.
   * - | docx-row-splittable
       | docx-no-row-splittable
     - table
     - Allow to split the table row into multiple pages, or not.
   * - | docx-header-in-all-page
       | docx-no-header-in-all-page
     - table
     - Always show the table header when the table is arranged in multiple pages, or not.

