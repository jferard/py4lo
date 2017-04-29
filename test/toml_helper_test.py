# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""

import unittest
from unittest.mock import *
import env

from toml_helper import load_toml
import io

class TestTomlHelper(unittest.TestCase):
    def test0(self):
        tdata = load_toml("z")
        self.assertEqual({
            'add_readme': 'true',
            'calc_exe': 'C:/Program Files (x86)/LibreOffice 5/program/scalc.exe',
            'debug_file': 'py4lo-debug.ods',
            'default_suffix': 'updated',
            'init_file': 'new-project.ods',
            'log_level': 'DEBUG',
            'py4lo_path': '/home/jferard/prog/python/py4lo',
            'python_exe': '/usr/bin/python3',
            'python_version': 3,
            'readme_contact': 'add a contact here',
            'src_dir': 'src/main',
            'target_dir': 'target',
            'test_dir': 'src/test'
            }, tdata)

    @patch('builtins.open', spec=open)
    def test_empty(self, mock_open):
        handle1 = MagicMock()
        handle1.__enter__.return_value = io.StringIO("a=1")
        handle1.__exit__.return_value=False

        handle2 = MagicMock()
        handle2.__enter__.return_value = io.StringIO("b=2")
        handle2.__exit__.return_value=False

        mock_open.side_effect = [handle1, handle2]
        tdata = load_toml("a")
        self.assertTrue({
            'a': 1,
            'b': 2,
            'log_level': 'INFO'}.items() <= tdata.items())
        self.assertEqual(set(['a', 'b', 'log_level','py4lo_path', 'python_exe', 'python_version']), set(tdata.keys()))

if __name__ == '__main__':
    unittest.main()
