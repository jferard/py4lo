#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
#
#     This file is part of Py4LO.
#
#     Py4LO is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     Py4LO is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
import unittest

import py4lo_helper
from py4lo_dialogs import message_box
from py4lo_io import (create_import_filter_options,
                      create_export_filter_options, Format)
from unittest import mock


class Py4LODialogsTestCase(unittest.TestCase):
    def setUp(self):
        self.xsc = mock.Mock()
        self.doc = mock.Mock()
        self.ctrl = mock.Mock()
        self.frame = mock.Mock()
        self.parent_win = mock.Mock()
        self.msp = mock.Mock()
        self.ctxt = mock.Mock()
        self.sm = mock.Mock()
        self.desktop = mock.Mock()
        py4lo_helper.provider = py4lo_helper._ObjectProvider(
            self.doc, self.ctrl, self.frame, self.parent_win, self.msp,
            self.ctxt, self.sm, self.desktop)
        py4lo_helper._inspect = py4lo_helper._Inspector(py4lo_helper.provider)

    def test_message_box(self):
        message_box("text", "title")
        self.sm.createInstanceWithContext.assert_called_once_with(
            "com.sun.star.awt.Toolkit", self.ctxt)
        sv = self.sm.createInstanceWithContext.return_value
        sv.createMessageBox.assert_called_once()
        mb = sv.createMessageBox.return_value
        mb.execute.assert_called_once()


if __name__ == '__main__':
    unittest.main()
