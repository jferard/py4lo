#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>
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
from unittest.mock import Mock, call

from callbacks import AddScripts


class TestAddScripts(unittest.TestCase):
    def test_add_two_scripts(self):
        zout = Mock()
        t1 = Mock()
        # use configure_mock to set the name attribute
        t1.configure_mock(name="t1", script_content="c1")
        t2 = Mock()
        t2.configure_mock(name="t2", script_content="c2")
        a = AddScripts([t1, t2])
        a.call(zout)
        self.assertEqual([call.writestr('Scripts/python/t1', 'c1'),
                          call.writestr('Scripts/python/t2', 'c2')],
                         zout.mock_calls)


if __name__ == '__main__':
    unittest.main()
