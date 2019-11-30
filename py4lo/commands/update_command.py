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
from commands import Command
from commands.command import PropertiesProvider
from commands.command_executor import CommandExecutor
from commands.test_command import TestCommand
from tools import update_ods


class UpdateCommand(Command):
    @staticmethod
    def create(args, provider: PropertiesProvider):
        test_executor = TestCommand.create(args, provider)
        update_command = UpdateCommand(provider)
        return CommandExecutor(update_command, test_executor)

    def __init__(self, provider: PropertiesProvider):
        self._tdata = provider.get()

    def execute(self, status):
        dest_name = update_ods(self._tdata)
        return status, dest_name

    @staticmethod
    def get_help():
        return "Update the file with all scripts"
