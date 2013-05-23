xlwt
===============

**A python library for creating Microsoft Excel spreadsheet files**

::

  from xlwt import Workbook

  wb = Workbook()
  
  sheet = wb.add_sheet('Hello xlwt')
  wb.add_sheet('Hello Excel')

  sheet.write(0, 0, 'Helo')
  sheet.write(0, 1, 'World')
  
  wb.save("hello-world.xls")
  
.. toctree::
    :maxdepth: 2

    quick_start
    examples
    Workbook
    Worksheet
    Row
    Column
    Cell
    Formatting
    Style
    ExcelFormula
    ExcelFormulaLexer
    ExcelFormulaParser
    ExcelMagic
    CompoundDoc
    Utils
    UnicodeUtils
    Bitmap
    antlr
    BIFFRecords

.. automodule:: xlwt
    :members:
    :undoc-members:
    :show-inheritance:






