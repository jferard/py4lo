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
from pathlib import Path
from unittest.mock import Mock, call

from core.script import SourceScript
from directives import EmbedLib


class TestEmbedLib(unittest.TestCase):
    def setUp(self):
        lib_path = Path(__file__).parent.parent.parent.parent / "lib"
        self._directive = EmbedLib(lib_path)

    def test_sig_elements(self):
        self.assertEqual(["embed", "lib"], self._directive.sig_elements())

    def test_execute(self):
        proc = Mock()
        self.assertEqual(True,
                         self._directive.execute(proc, None, ["py4lo_helper"]))
        self.assertEqual([call.append_script(
            SourceScript(
                Path('/home/jferard/prog/python/py4lo/lib/py4lo_helper.py'),
                Path('/home/jferard/prog/python/py4lo/lib'), False))],
            proc.mock_calls)

        if __name__ == '__main__':
            unittest.main()
