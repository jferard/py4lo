# -*- coding: utf-8 -*-
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
import unittest
from pathlib import Path
from unittest import mock

from commands.run_command import RunCommand
from core.properties import PropertiesProvider


class TestRunCommand(unittest.TestCase):
    def setUp(self):
        self.provider: PropertiesProvider = mock.Mock()

    @mock.patch("commands.run_command.secure_exe", autospec=True)
    @mock.patch("commands.test_command.secure_exe", autospec=True)
    def test_with_empty_tdata(self, secure_mock, secure_mock2):
        secure_mock.side_effect = lambda x, _y: x
        secure_mock2.side_effect = lambda x, _y: x
        self.provider.get.return_value = {}
        RunCommand.create_executor([], self.provider)

    @mock.patch("subprocess.call", autospec=True)
    def test_create(self, call_mock):
        self.provider.get.return_value = {"log_level": 0, "python_exe": "py",
                                          "python_version": 3.7,
                                          "test_dir": ".", "src_dir": ".",
                                          "base_path": ".", "lo_exe": "ca",
                                          "source_file": "ods",
                                          "suffix": ".ext", "src_ignore": "*",
                                          "assets_dir": ".",
                                          "target_dir": ".",
                                          "assets_dest_dir": ".",
                                          "assets_ignore": "*"}

    @mock.patch("tools.secure_exe", autospec=True)
    @mock.patch("subprocess.call", autospec=True)
    def test(self, call_mock, secure_mock):
        secure_mock.side_effect = lambda x, _y: x
        h = RunCommand("calc")
        h.execute(0, Path(""))
        self.assertEqual([mock.call(['calc', '.'])], call_mock.mock_calls)
