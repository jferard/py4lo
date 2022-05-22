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
from unittest import mock

from py4lo_dialogs import message_box, MessageBoxType


class Py4LODialogsTestCase(unittest.TestCase):
    @mock.patch("py4lo_dialogs.uno_service_ctxt")
    @mock.patch("py4lo_dialogs.provider")
    def test_message_box(self, provider, uno_service_ctxt):
        self.maxDiff = None
        sm = mock.Mock()

        # prepare
        toolkit = mock.Mock()
        uno_service_ctxt.return_value = toolkit

        # replay
        message_box("text", "title")

        self.assertEqual([
            mock.call('com.sun.star.awt.Toolkit'),
            mock.call().createMessageBox(provider.parent_win,
                                         MessageBoxType.MESSAGEBOX, 1,
                                         'title', 'text'),
            mock.call().createMessageBox().execute()
        ], uno_service_ctxt.mock_calls)


if __name__ == '__main__':
    unittest.main()
