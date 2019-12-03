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
from abc import ABC
from pathlib import Path

from commands import Command
from commands.command import PropertiesProvider
from commands.command_executor import CommandExecutor
from commands.debug_command import DebugCommand
from commands.test_command import TestCommand


class InitCommand(Command, ABC):
    @staticmethod
    def create_executor(args, provider: PropertiesProvider):
        tdata = provider.get()
        test_executor = TestCommand.create_executor(args, provider)
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        init_command = DebugCommand(
            logger,
            Path(tdata["base_path"]),
            Path(tdata["src_dir"]),
            tdata["src_ignore"],
            Path(tdata["assets_dir"]),
            tdata["assets_ignore"],
            Path(tdata["target_dir"]),
            Path(tdata["assets_dest_dir"]),
            tdata["python_version"],
            Path(tdata["init_file"])
        )
        return CommandExecutor(init_command, test_executor)

    @staticmethod
    def get_help():
        return "Create a new document from script"
