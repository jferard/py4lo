# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2021 J. Férard <https://github.com/jferard>

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

import io
import sys
import unittest
from unittest.mock import Mock

from toml_helper import TomlLoader, load_toml
from tst_env import file_path_mock, file_path_error_mock, verify_open_path


class TestTomlHelper(unittest.TestCase):
    def test_loader(self):
        default_toml = file_path_mock(io.StringIO("a=1"))
        local_toml = file_path_mock(io.StringIO("b=2"))

        tdata = TomlLoader(default_toml, local_toml, {}).load()
        self.assertTrue({
                            'a': 1,
                            'b': 2,
                            'log_level': 'INFO'}.items() <= tdata.items())
        self.assertEqual(
            {'a', 'b', 'log_level', 'python_exe', 'python_version'},
            set(tdata.keys()))
        verify_open_path(self, default_toml, 'r', encoding="utf-8")
        verify_open_path(self, local_toml, 'r', encoding="utf-8")

    def test_load_toml(self):
        default_toml = file_path_mock(io.StringIO("a=1"))
        local_toml = file_path_mock(io.StringIO("b=2"))

        tdata = load_toml(default_toml, local_toml, {})
        self.assertTrue({
                            'a': 1,
                            'b': 2,
                            'log_level': 'INFO'}.items() <= tdata.items())
        self.assertEqual(
            {'a', 'b', 'log_level', 'python_exe', 'python_version'},
            set(tdata.keys()))
        verify_open_path(self, default_toml, 'r', encoding="utf-8")
        verify_open_path(self, local_toml, 'r', encoding="utf-8")

    def test_load_toml_python_exe(self):
        default_toml = file_path_mock(io.StringIO("python_exe='man'"))
        local_toml = file_path_mock(io.StringIO("b=2"))

        tdata = load_toml(default_toml, local_toml, {})
        self.assertTrue({
                            'python_exe': 'man',
                            'b': 2,
                            'log_level': 'INFO'}.items() <= tdata.items())
        self.assertEqual(
            {'python_exe', 'b', 'log_level', 'python_version'},
            set(tdata.keys()))
        verify_open_path(self, default_toml, 'r', encoding="utf-8")
        verify_open_path(self, local_toml, 'r', encoding="utf-8")

    def test_exception(self):
        default_toml = file_path_error_mock()
        local_toml = file_path_mock(io.StringIO("b = 2"))

        err_bkp = sys.stderr
        sys.stderr = io.StringIO()
        tdata = load_toml(default_toml, local_toml, {})
        self.assertRegex(sys.stderr.getvalue(), "Error when loading toml file")
        sys.stderr = err_bkp

        self.assertEqual(
            {'b', 'log_level', 'python_exe', 'python_version'},
            set(tdata.keys()))
        verify_open_path(self, local_toml, 'r', encoding="utf-8")


if __name__ == '__main__':
    unittest.main()
