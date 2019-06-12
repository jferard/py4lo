# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>

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
import unittest
from unittest.mock import *
import env
from zip_updater import *

class TestZipUpdater(unittest.TestCase):
    @patch('zipfile.ZipFile', autospec=True)
    def test(self, zf):
        zu = ZipUpdater()
        cbs = []

        b1 = Mock()
        b2 = Mock()
        i1 = Mock()
        i2 = Mock()
        a1 = Mock()
        a2 = Mock()

        zout = Mock()
        p = PropertyMock()
        type(zout).comment = p
        zin = Mock(comment="a")
        zin.infolist.return_value = [1,2]

        rout = MagicMock()
        rout.__enter__.return_value = zout
        rout.__exit__.return_value = False
        rin = MagicMock()
        rin.__enter__.return_value = zin
        rin.__exit__.return_value = False

        zf.side_effect = [rout, rin]

        b1.call.return_value = True
        b2.call.return_value = False
        i1.call.return_value = [True, True]
        i2.call.return_value = [False, False]
        a1.call.return_value = True
        a2.call.return_value = False

        zu.before(b1).before(b2).item(i1).item(i2).after(a1).after(a2)
        zu.update("source", "dest")


        p.assert_called_once_with("a")
        b1.call.assert_called_once_with(zout)
        b2.call.assert_called_once_with(zout)
        self.assertEquals([call(zin, zout, 1), call(zin, zout, 2)], i1.call.mock_calls)
        self.assertEquals([call(zin, zout, 1), call(zin, zout, 2)], i2.call.mock_calls)
        a1.call.assert_called_once_with(zout)
        a2.call.assert_called_once_with(zout)
