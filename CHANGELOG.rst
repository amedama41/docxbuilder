Unreleased
----------

Bug fix
*******

* Fix broken references.
* #5: Fix error when List Bullet / List Number styles have no numbering style.
* Fix error in operating section property without start number.

New features
************

* Add docx_pagebreak_before_table_of_contents option.
* Add Horizontal List style for hlist directive.
* Add docx_style_names to specify user defined styles.

Enhancement
***********

* Use OMML tags for equations.
* Support table width option.
* Change table margin.
* #6: Use the first section / page as cover page when missing cover page object.
* Support math role and directive. This needs math extras_require.

Release 1.1.5 (2019-09-30)
--------------------------

Bug fix
*******

* Fix document broken by multiple references to a same footnote.
* Fix documentation build error by broken bookmark in Windows.

Release 1.1.4 (2019-09-25)
--------------------------

Bug fix
*******

* Fix broken TOC by conflicts with bookmark IDs in cover page.

Enhancement
***********

* Extend top margin of description table cell in default style for appearance.

Release 1.1.3 (2019-09-24)
--------------------------

Bug fix
*******

* Fix an issue that unnecessary table bottom margins are unavailable to deleted.

Release 1.1.2 (2019-09-23)
--------------------------

Enhancement
***********

* Remove unnecessary table bottom margins.
* Change Japanese default fonts.
* Disable automatic layout coordination in default Literal Block style.

Release 1.1.1 (2019-09-16)
--------------------------

Enhancement
***********

* Support BMP/TIFF/ICO/WEBP image format
* Rename style id in order to enhance human readability
* Not to display only table margin paragraph in order to avoid empty page

Release 1.1.0 (2019-08-06)
--------------------------

Bug fix
*******

* Fix style name for topic and sidebar title

New features
************

* Add Definition and Legend styles
* Add docx_update_fields option
* Add docx_pagebreak_before_file option
* Add docx-section-portrait-*N* and docx-section-landscape-*N* custom classes.
* Add docx-rotation-header-*N* custom class.
* Support raw directive
* Enable to generate cover page properties

Enhancement
***********

* Not apply paragraph style to paragraphs in table
* Suppress field name wrap
* Not apply center alignment to cells of default field list style
* Change code highlight color according as pygments_style
* Enable to specify date object to lastPrinted property
* Refactor function to classify document properties
* Enhance style file information extraction
* Remove incorrect app properties

