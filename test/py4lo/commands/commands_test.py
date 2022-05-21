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
from unittest.mock import Mock

from commands import commands
from core.properties import PropertiesProvider


class TestCommands(unittest.TestCase):
    def setUp(self):
        self.provider: PropertiesProvider = Mock()

    def test(self):
        executor = commands.get("run", ["arg1", "arg2"], self.provider)
        # print(progress_executor.execute(["arg1", "arg2"]))

    def test_non_existing(self):
        executor = commands.get("foo", ["arg1", "arg2"], self.provider)
        # print(progress_executor.execute([["arg1", "arg2"]]))

    def test_help(self):
        self.assertEqual("""a command = debug | init | test | run | update | help
debug: Create a debug.ods file with button for each function
init: Create a new document from script
test: Do the test of the scripts to add to the spreadsheet
run: Update + open the created file
update: Update the file with all scripts
help: help [command]: Specific help message about command""",
                         commands.get_help_message())
