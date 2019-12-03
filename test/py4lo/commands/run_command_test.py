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
from unittest.mock import Mock, patch, call

from commands.run_command import *


class TestRunCommand(unittest.TestCase):
    def setUp(self):
        self.provider = Mock()

    def test_with_empty_tdata(self):
        self.provider.get.return_value = {}
        self.assertRaises(KeyError,
                          lambda: RunCommand.create_executor([], self.provider))

    @patch("subprocess.call", autospec=True)
    def test_create(self, call_mock):
        self.provider.get.return_value = {"log_level": 0, "python_exe": "py",
                                          "python_version": 3.7,
                                          "test_dir": ".", "src_dir": ".",
                                          "base_path": ".", "calc_exe": "ca",
                                          "source_file": "ods",
                                          "suffix": ".ext", "src_ignore": "*",
                                          "assets_dir": ".",
                                          "target_dir": ".",
                                          "assets_dest_dir": ".",
                                          "assets_ignore": "*"}

    @patch("subprocess.call", autospec=True)
    def test(self, call_mock):
        h = RunCommand("calc")
        h.execute(0, Path(""))
        self.assertEqual([call(['calc', '.'])], call_mock.mock_calls)
