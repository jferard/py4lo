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
from typing import Dict, List

from commands.command import Command, PropertiesProvider
from commands.real_command_factory_by_name import real_command_factory_by_name
from commands.help_command import HelpCommand
import logging


class Commands:
    def __init__(self, command_factory_by_name: Dict[str, Command]):
        # assert "help" in command_factory_by_name
        self._command_factory_by_name = command_factory_by_name

    def get(self, command_name: str, args: List[str], provider: PropertiesProvider):
        try:
            return self._command_factory_by_name[command_name].create_executor(args,
                                                                               provider)
        except KeyError:
            logger = logging.getLogger("py4lo")
            logger.warning(
                "Command `{}` not found. Available commands are {}".format(
                    command_name, set(self._command_factory_by_name.keys())))
            return self._command_factory_by_name["help"].create_executor(args, provider)

    def get_help_message(self):
        lines = [
            "a command = {}".format(" | ".join(self._command_factory_by_name))]
        for name, cf in self._command_factory_by_name.items():
            lines.append("{}: {}".format(name, cf.get_help()))
        return "\n".join(lines)


cf_by_name = real_command_factory_by_name.copy()
# add help now
cf_by_name['help'] = HelpCommand

commands = Commands(cf_by_name)
