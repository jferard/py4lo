# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   Py4LO is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   Py4LO is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import uno
import sys
import unohelper
import os

# py4lo: import lib py4lo_helper
# py4lo: import lib py4lo_commons
# py4lo: import example_lib
_ = py4lo_helper.Py4LO_helper.create(XSCRIPTCONTEXT)
c = py4lo_commons.Commons(XSCRIPTCONTEXT)
o = example_lib.O(XSCRIPTCONTEXT)

def message_example(*_args):
    from com.sun.star.awt.MessageBoxType import MESSAGEBOX
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    _.message_box(_.parent_win, "A message from main script example.py. Current dir is: "+os.path.abspath("."), "py4lo", MESSAGEBOX, BUTTONS_OK)

def xray_example(*_args):
    _.xray(_.doc)

def example_from_lib(*_args):
    o.lib_example()
