#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
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
import io
import unittest
from pathlib import Path
from unittest import mock

from core.script import TempScript
from directives import EmbedScript

from test.test_helper import file_path_mock, verify_open_path


class TestEmbed(unittest.TestCase):
    def setUp(self):
        self._opt: Path = mock.Mock()
        self._directive = EmbedScript(self._opt)

    def test_sig_elements(self):
        self.assertEqual(["embed", "script"], self._directive.sig_elements())

    def test_execute(self):
        # Arrange
        proc = mock.Mock()
        fpath = file_path_mock(io.BytesIO(b"content"), suffix = ".py")
        fpath.with_suffix.return_value = fpath

        self._opt.joinpath.return_value = fpath
        fpath.is_dir.return_value = False

        # Act
        execute = self._directive.execute(proc, None, ["a/b.py"])

        # Assert
        self.assertEqual(True, execute)
        self.assertEqual([mock.call.add_script(
            TempScript(fpath, b"content", self._opt, [], None))],
            proc.mock_calls)
        verify_open_path(self, fpath, 'rb')

    def test_execute_dir(self):
        # Arrange
        proc = mock.Mock()
        fpath = file_path_mock(io.BytesIO(b"content"), suffix=".py")
        fpath.with_suffix.return_value = fpath
        fpath.is_dir.return_value = False

        fdirpath = mock.Mock()
        fdirpath.is_dir.return_value = True
        fdirpath.glob.return_value = [fpath]

        self._opt.joinpath.return_value = fdirpath

        # Act
        execute = self._directive.execute(proc, None, ["a/b"])

        # Assert
        self.assertEqual(True, execute)
        self.assertEqual([mock.call.add_script(
            TempScript(fpath, b"content", self._opt, [], None))],
            proc.mock_calls)
        verify_open_path(self, fpath, 'rb')

    def test_execute_subdir(self):
        # Arrange
        proc = mock.Mock()

        fpath = file_path_mock(io.BytesIO(b"content"), suffix=".py")
        fpath.with_suffix.return_value = fpath
        fpath.is_dir.return_value = False

        fsubdirpath = mock.Mock()
        fsubdirpath.is_dir.return_value = True
        fsubdirpath.glob.return_value = [fpath]

        fdirpath = mock.Mock()
        fdirpath.is_dir.return_value = True
        fdirpath.glob.return_value = [fsubdirpath]

        self._opt.joinpath.return_value = fdirpath

        # Act
        execute = self._directive.execute(proc, None, ["a/b"])

        # Assert
        self.assertEqual(True, execute)
        self.assertEqual([mock.call.add_script(
            TempScript(fpath, b"content", self._opt, [], None))],
            proc.mock_calls)
        verify_open_path(self, fpath, 'rb')


if __name__ == '__main__':
    unittest.main()
