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

import io
import unittest
from pathlib import Path
from unittest.mock import *

from toml_helper import load_toml, TomlLoader
from env import file_path_mock, verify_open_path


class TestTomlHelper(unittest.TestCase):
    def test(self):
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


if __name__ == '__main__':
    unittest.main()
