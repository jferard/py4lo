# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2017 J. Férard <https://github.com/jferard>

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
# py4lo: include license

# py4lo: import lib py4lo_helper
_ = py4lo_helper.Py4LO_helper()

import sys
import unohelper

class O(unohelper.Base):
    def lib_example(self):
        from com.sun.star.awt.MessageBoxType import MESSAGEBOX
        from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
        _.message_box(_.parent_win, "A message from imported lib example-lib.py", "py4lo", MESSAGEBOX, BUTTONS_OK)

def __export_o():
    return O()
