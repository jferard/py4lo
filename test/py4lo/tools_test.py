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
import unittest
from unittest.mock import patch, call

import env
from tools import *
from tools import _get_dest


class TestTools(unittest.TestCase):
    @patch('subprocess.call', spec=subprocess.call)
    def test_open(self, subprocess_call_mock):
        open_with_calc(Path("myfile.ods"), "mycalc.exe")
        self.assertEqual([call(['mycalc.exe', 'myfile.ods'])],
                         subprocess_call_mock.mock_calls)

    def test_get_dest(self):
        self.assertEqual(Path("dest.ods"), _get_dest({"dest_name": "dest.ods"}))
        self.assertEqual(Path("dest.ods"), _get_dest(
            {"dest_name": "dest.ods", "suffix": "up", "log_level": 0}))
        self.assertEqual(Path("source-up.ods"), _get_dest(
            {"source_file": "source.ods", "suffix": "up"}))
