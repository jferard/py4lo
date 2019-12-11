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
from unittest.mock import Mock, call, MagicMock

from core.script import SourceScript, TempScript
from directives import EmbedScript


class TestEmbed(unittest.TestCase):
    def setUp(self):
        self._opt: Path = Mock()
        self._directive = EmbedScript(self._opt)

    def test_sig_elements(self):
        self.assertEqual(["embed", "script"], self._directive.sig_elements())

    def test_execute(self):
        proc = Mock()
        fpath: Path = Mock()
        bound = MagicMock()
        f = Mock()

        self._opt.joinpath.return_value = fpath
        fpath.is_dir.return_value = False
        fpath.open.return_value = bound
        bound.__enter__.return_value = f
        bound.__exit__.return_value = False
        f.read.return_value = b"content"
        self.assertEqual(True,
                         self._directive.execute(proc, None, ["a/b.py"]))
        self.assertEqual([call.add_script(
            TempScript(fpath, b"content", self._opt, [], None))],
            proc.mock_calls)

        if __name__ == '__main__':
            unittest.main()
