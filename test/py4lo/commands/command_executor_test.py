# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2020 J. Férard <https://github.com/jferard>

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
from logging import Logger
from unittest.mock import Mock, call

from commands import Command
from commands.command_executor import CommandExecutor


class TestCommandExecutor(unittest.TestCase):
    def setUp(self):
        self._logger: Logger = Mock()
        self._command: Command = Mock()

    def test_without_previous(self):
        ce = CommandExecutor(self._logger, self._command, None)
        ce.execute(["1", "2"])
        self.assertEqual([call.execute()], self._command.mock_calls)

    def test_with_previous(self):
        previous = Mock()

        previous.execute.return_value = ["3", "4"]

        ce = CommandExecutor(self._logger, self._command, previous)
        ce.execute(["1", "2"])

        self.assertEqual([call.execute(["1", "2"])], previous.mock_calls)
        self.assertEqual([call.execute("3", "4")], self._command.mock_calls)
