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
from typing import List, Tuple, Optional

from commands import Command
from commands.command_executor import CommandExecutor
from commands.debug_command import DebugCommand
from commands.ods_updater import OdsUpdaterHelper
from commands.test_command import TestCommand
from core.properties import PropertiesProvider


class InitCommand(Command):
    @staticmethod
    def create_executor(args, provider: PropertiesProvider) -> CommandExecutor:
        test_executor = TestCommand.create_executor(args, provider)
        logger = provider.get_logger()
        sources = provider.get_sources()
        destinations = provider.get_destinations()
        python_version = provider.get("python_version")
        helper = OdsUpdaterHelper(logger, sources,
                                  destinations,
                                  python_version)
        init_command = DebugCommand(logger, helper,
                                    sources, destinations,
                                    python_version)
        return CommandExecutor(logger, init_command, test_executor)

    def execute(self, *args: List[str]) -> Optional[Tuple]:
        pass

    @staticmethod
    def get_help():
        return "Create a new document from script"
