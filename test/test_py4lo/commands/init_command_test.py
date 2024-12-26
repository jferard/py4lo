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
from unittest import mock

from commands.init_command import InitCommand
from core.properties import PropertiesProvider


class TestInitCommand(unittest.TestCase):
    @mock.patch("commands.test_command.secure_exe", autospec=True)
    def test(self, secure_mock):
        secure_mock.side_effect = lambda x, _y: x
        provider: PropertiesProvider = mock.Mock()
        _ = InitCommand.create_executor([], provider)