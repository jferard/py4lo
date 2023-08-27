#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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
from logging import Logger
from unittest import mock
from zipfile import ZipFile

from callbacks import AddScripts
from core.script import DestinationScript


class TestAddScripts(unittest.TestCase):
    def test_add_two_scripts(self):
        logger: Logger = mock.Mock()
        zout: ZipFile = mock.Mock()
        t1: DestinationScript = mock.Mock(
            script_path="Scripts/python/t1", script_content="c1")
        t2: DestinationScript = mock.Mock(
            script_path="Scripts/python/t2", script_content="c2")

        a = AddScripts(logger, [t1, t2])
        a.call(zout)
        self.assertEqual([
            mock.call.writestr('Scripts/python/t1', 'c1'),
            mock.call.writestr('Scripts/python/t2', 'c2')
        ], zout.mock_calls)


if __name__ == '__main__':
    unittest.main()
