# -*- coding: utf-8 -*-

__VERSION__ = '0.7.5'


__AUTHOR__ = 'Landon Jurgens, Chris Withers'


## Ssuggested by pete
__PYTHON_EXCEL_MODULE__ = "xlwt"
__PYTHON_EXCEL_TITLE__ = "xlwt - Excel writer"


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
