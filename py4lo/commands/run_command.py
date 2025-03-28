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
import logging
from pathlib import Path
from typing import cast, Tuple, Any, List

from commands.command import Command
from commands.null_command import NullCommand
from core.properties import PropertiesProvider
from commands.command_executor import CommandExecutor
from commands.update_command import UpdateCommand
from tools import open_with_libreoffice, secure_exe


class RunCommand(Command):
    @staticmethod
    def create_executor(args: List[str], provider: PropertiesProvider) -> CommandExecutor:
        sec_lo_exe = secure_exe(provider.get("lo_exe"), "soffice")
        if sec_lo_exe is None:
            update_executor = None
            run_command = cast(Command, NullCommand("Can't find lo exe : {}".format(provider.get("lo_exe"))))
        else:
            update_executor = UpdateCommand.create_executor(args, provider)
            run_command = RunCommand(sec_lo_exe)
        return CommandExecutor(provider.get_logger(), run_command,
                               update_executor)

    def __init__(self, lo_exe: str):
        self._lo_exe = lo_exe

    def execute(self, status: int, dest_name: Path) -> Tuple[Any, ...]:
        if status == 0:
            print("All tests ok")
            logging.warning("%s %s", self._lo_exe, dest_name)
            open_with_libreoffice(dest_name, self._lo_exe)
        else:
            print("Error: some tests failed")
        return tuple()

    def get_help(self) -> str:
        return "Update + open the created file"
