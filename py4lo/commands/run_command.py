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
import logging
from pathlib import Path

from commands.command import Command
from core.properties import PropertiesProvider
from commands.command_executor import CommandExecutor
from commands.update_command import UpdateCommand
from tools import open_with_calc


class RunCommand(Command):
    @staticmethod
    def create_executor(args, provider: PropertiesProvider):
        tdata = provider.get()
        update_executor = UpdateCommand.create_executor(args, provider)
        run_command = RunCommand(tdata["calc_exe"])
        return CommandExecutor(run_command, update_executor)

    def __init__(self, calc_exe: str):
        self._calc_exe = calc_exe

    def execute(self, status: int, dest_name: Path) -> None:
        if status == 0:
            print("All tests ok")
            logging.warning("%s %s", self._calc_exe, dest_name)
            open_with_calc(dest_name, self._calc_exe)
        else:
            print("Error: some tests failed")

    @staticmethod
    def get_help():
        return "Update + open the created file"
