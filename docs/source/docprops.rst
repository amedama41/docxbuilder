Document properties
===================

Docxbuilder is enable to embed document properties into the generated document.
The document properties can be referenced from the cover page (Use Quick Parts of Office Word).

The document properties is defined by **docx_documents's docproperties** configuration.
The docproperties is a dictionary from property name to the value.
Docxbuilder treats some names as the properties defined by OOXML.

Property names included in the below list are used as the Core Properties [ECMA376]_.

.. hlist::
   :columns: 3

   - title
   - creator
   - language
   - category
   - contentStatus
   - description
   - identifier
   - lastModifiedBy
   - lastPrinted
   - revision
   - subject
   - version
   - keywords
   - created
   - modified

The value of "created" or "modified" property must be a ``date`` or
``datetime`` object, or a string formatted by one of the following formats.

.. hlist::
   :columns: 3

   - ``YYYY``
   - ``YYYY-MM``
   - ``YYYY-MM-DD``
   - ``YYYY-MM-DDThh``
   - ``YYYY-MM-DDThh:mm``
   - ``YYYY-MM-DDThh:mm:ss``
   - ``YYYY-MM-DDThh:mm:ss.s``

The value of "lastPrinted" must be a ``date`` or ``datetime`` object,
or a string formatted by ``YYYY-MM-DDThh:mm:ss``.
All times expressed by string type are interpreted as system local time.

The value of "keywords" must be a string or a list of strings.
All value of other core properties must be a string.

Property names included in the below list are used as the Extended Properties [ECMA376]_.

.. hlist::
   :columns: 3

   * company
   * manager

Property names included in the below list are used as the Cover Page Properties [MSOE376]_.

.. hlist::
   :columns: 3

   * abstract
   * companyAddress
   * companyEmail
   * companyFax
   * companyPhone
   * publishDate

The value of "publishDate" property must be a ``date`` or ``datetime`` object,
or a string formatted by one of the above formats.

The other keys are used as custom properties.
The value of custom properties must be an integer, float, string, bool,
or ``datetime`` object.


