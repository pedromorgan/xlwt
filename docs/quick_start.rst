Quick Start - writing an Excel file
===================================

Here is a simple example that creates a workbook with two sheets.
::

  from xlwt import Workbook

  ## Create a spreadsheet in memory
  wb = Workbook()
  
  ## add a tab and some stuff
  sheet1 = book.add_sheet('Sheet 1')
  book.add_sheet('Sheet 2')

  sheet1.write(0, 0, 'A1')
  sheet1.write(0, 1, 'B1')
  row1 = sheet1.row(1)
  row1.write(0, 'A2')
  row1.write(1, 'B2')
  sheet1.col(0).width = 10000

  ## create second tab
  sheet2 = book.get_sheet(1)
  sheet2.row(0).write(0,'Sheet 2 A1')
  sheet2.row(0).write(1,'Sheet 2 B1')
  sheet2.flush_row_data()
  sheet2.write(1,0,'Sheet 2 A3')
  sheet2.col(0).width = 5000
  sheet2.col(0).hidden = True

  # Write to file
  book.save('xlwt-example.xls')
  

Creating a Workbook 
-----------------------------------

Workbooks are created with `xlwt` by instantiating 
an :py:class:`~xlwt.Workbook.Workbook` object, manipulating 
it and then calling its :py:func:`~xlwt.Workbook.Workbook.save` method.

The :py:func:`~xlwt.Workbook.Workbook.save` method may be passed either a 
string containing the path to write to 
or a file-like object, opened for writing in binary mode
to which the binary Excel file data will be written.

The following objects can be created within a workbook:

Worksheets
~~~~~~~~~~~~~

Worksheets are created with the :py:class:`~xlwt.Workbook.Workbook.add_sheet()` method 
of the :py:class:`~xlwt.Workbook.Workbook` class.

To retrieve an existing sheet from a :py:class:`~xlwt.Workbook.Workbook`, use 
its :py:class:`~xlwt.Workbook.Workbook.get_sheet()` method. This method is particularly 
useful when the :py:class:`~xlwt.Workbook.Workbook` has been 
instantiated by :py:class:`xlutils.copy`.

Rows
~~~~~~~

Rows are created using the ``row`` method of the ``Worksheet`` class and contain all of the cells for a given row.

The ``row`` method is also used to retrieve existing rows from a ``Worksheet``.

If a large number of rows have been written to a ``Worksheet`` and memory usage is becoming a problem, the ``flush_row_data`` method may be called on the ``Worksheet``. Once called, any rows flushed cannot be accessed or modified.

It is recommended that ``flush_row_data`` is called for every 1000 or so rows of a normal size that are written to an ``xlwt.Workbook``. If the rows are huge, that number should be reduced.

Columns
~~~~~~~

Columns are created using the ``col`` method of the ``Worksheet`` class and contain display formatting information for a given column.

The ``col`` method is also used to retrieve existing columns from a ``Worksheet``.

Cells
~~~~~

Cells can be written using either the ``write`` method of either the ``Worksheet`` or ``Row`` class.

A more detailed discussion of different ways of writing cells and the different types of cell that may be written is covered later.

A Simple Example
~~~~~~~~~~~~~~~~

The following example shows how all of the above methods can be used to build and save a simple workbook:

::

  from tempfile import TemporaryFile
  from xlwt import Workbook

  book = Workbook()
  sheet1 = book.add_sheet('Sheet 1')
  book.add_sheet('Sheet 2')

  sheet1.write(0,0,'A1')
  sheet1.write(0,1,'B1')
  row1 = sheet1.row(1)
  row1.write(0,'A2')
  row1.write(1,'B2')
  sheet1.col(0).width = 10000

  sheet2 = book.get_sheet(1)
  sheet2.row(0).write(0,'Sheet 2 A1')
  sheet2.row(0).write(1,'Sheet 2 B1')
  sheet2.flush_row_data()
  sheet2.write(1,0,'Sheet 2 A3')
  sheet2.col(0).width = 5000
  sheet2.col(0).hidden = True

  book.save('simple.xls')
  book.save(TemporaryFile())
  
  simple.py

Unicode
--------

The best policy is to pass unicode objects to all ``xlwt``-related method calls.

If you absolutely have to use encoded strings then make sure that the encoding used is consistent across all calls to any ``xlwt``-related methods.

If encoded strings are used and the encoding is not ``'ascii'``, then any ``Workbook`` objects must be created with the appropriate encoding specified:

::

  from xlwt import Workbook
  book = Workbook(encoding='utf-8')
  
  