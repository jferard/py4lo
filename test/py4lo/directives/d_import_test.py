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
from pathlib import Path
from unittest.mock import Mock, call

from directives import Import


class TestImport(unittest.TestCase):
    def test(self):
        proc = Mock()
        d = Import(Path(""), Path(""))
        self.assertEqual(["import"], d.sig_elements())
        self.assertEqual(True, d.execute(proc, ["a"]))
        self.assertEqual([call.include('py4lo_import.py'),
                          call.append_script(Path('a.py')),
                          call.append('import a\n')], proc.mock_calls)


if __name__ == '__main__':
    unittest.main()
