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


class TestTomlHelper(unittest.TestCase):
    @patch('builtins.open', spec=open)
    def test(self, open_mock):
        path = Mock()
        def_toml = Mock()
        handle1 = MagicMock()
        handle1.__enter__.return_value = io.StringIO("a=1")
        handle1.__exit__.return_value = False
        path.joinpath.return_value = def_toml
        def_toml.open.return_value = handle1

        local_toml = Mock()
        handle2 = MagicMock()
        handle2.__enter__.return_value = io.StringIO("b=2")
        handle2.__exit__.return_value = False

        local_toml.open.return_value = handle2
        tdata = TomlLoader(path, local_toml).load()
        print(tdata)
        self.assertTrue({
            'a': 1,
            'b': 2,
            'log_level': 'INFO'}.items() <= tdata.items())
        self.assertEqual({'a', 'b', 'log_level', 'py4lo_path', 'python_exe', 'python_version'}, set(tdata.keys()))


if __name__ == '__main__':
    unittest.main()
