# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import unittest
from unittest.mock import *
import env
import sys, os

from py4lo_helper import *

class TestHelper(unittest.TestCase):
    def setUp(self):
        self.xsc = MagicMock()
        self.doc = MagicMock()
        self.ctxt = MagicMock()
        self.ctrl = MagicMock()
        self.frame = MagicMock()
        self.parent_win = MagicMock()
        self.sm = MagicMock()
        self.dsp = MagicMock()
        self.mspf = MagicMock()
        self.msp = MagicMock()
        self.reflect = MagicMock()
        self.dispatcher = MagicMock()
        self.loader = MagicMock()
        self.p = Py4LO_helper(self.doc, self.ctxt, self.ctrl, self.frame, self.parent_win, self.sm, self.dsp, self.mspf, self.msp, self.reflect, self.dispatcher, self.loader)

    def testXray(self):
        import uno
        self.p.use_xray()
        self.msp.getScript.assert_called_once_with('vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application')

    def testXray2(self):
        import uno
        self.p.xray(1)
        self.p.xray(2)
        self.msp.getScript.assert_called_once_with('vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application')
        self.msp.getScript.return_value.invoke.assert_has_calls([call((1,), (), ()), call((2,), (), ())])

if __name__ == '__main__':
    unittest.main()
