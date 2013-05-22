# -*- coding: utf-8 -*-

__version__ = '0.7.5'

__author__ = 'Landon Jurgens, Chris Withers'
__license__ = "== TODO =="

## Ssuggested by pete
__PYTHON_EXCEL_MODULE__ = "xlwt"
__PYTHON_EXCEL_TITLE__ = "xlwt - Excel writer"


__docformat__ = 'restructuredtext en'

__doc__ = """
:abstract: A python lib for generating Microsoft Excel spreadsheet files
:version: %s
:author: %s
:contact: http://www.python-excel.org
:date: 2013-05-17
:copyright: %s
""" % (__version__, __author__, __license__)



import sys
if sys.version_info[:2] < (2, 3):
    print >> sys.stderr, "Sorry, xlwt requires Python 2.3 or later"
    sys.exit(1)


from Workbook import Workbook
from Worksheet import Worksheet
from Row import Row
from Column import Column
from Formatting import Font, Alignment, Borders, Pattern, Protection
from Style import XFStyle, easyxf, easyfont, add_palette_colour
from ExcelFormula import *
