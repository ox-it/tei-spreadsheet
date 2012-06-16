Spreadsheet to TEI conversion
=============================

This repository currently contains stylesheets to convert Office Open XML
spreadsheets — i.e., the ``.xlsx`` spreadsheets created by Excel into TEI XML.

Using
-----

Simply use the convert script::

    $ ./convert.sh spreadsheet.xlsx > output.xml

Features
--------

* Converts all sheets within a workbook
* Extracts author and title information
* Pulls out string and numeric data within cells

Limitations
-----------

* Unable to convert styled text (e.g. italic, superscript)
* Dates are not converted from their serialization as a numeric value (approximately days since 1900-01-01)

