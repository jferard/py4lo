# -*- coding: utf-8 -*-
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
from unittest.mock import Mock, patch, call

from commands.help_command import *
from core.properties import PropertiesProvider


class TestHelpCommand(unittest.TestCase):
    def setUp(self):
        self.provider: PropertiesProvider = Mock()

    @patch("__main__.__builtins__.print", autospec=True)
    def test_without_command(self, print_mock):
        h = HelpCommand.create_executor([], self.provider)
        h.execute()
        self.assertEqual(
            [call('help [command]: Specific help message about command')],
            print_mock.mock_calls)

    @patch("__main__.__builtins__.print", autospec=True)
    def test_with_command(self, print_mock):
        h = HelpCommand.create_executor(["run"], self.provider)
        h.execute()
        self.assertEqual(
            [call('Update + open the created file')],
            print_mock.mock_calls)

    @patch("__main__.__builtins__.print", autospec=True)
    def test_with_grabage(self, print_mock):
        h = HelpCommand.create_executor(["a", "b", "c"], self.provider)
        h.execute()
        self.assertEqual(
            [call('help [command]: Specific help message about command')],
            print_mock.mock_calls)
