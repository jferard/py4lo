# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""

from commands.real_command_factory_by_name import real_command_factory_by_name
from commands.help_command import HelpCommand

class Commands():
    def __init__(self, command_factory_by_name):
        assert "help" in command_factory_by_name
        self.__command_factory_by_name = command_factory_by_name

    def get(self, command_name, args, tdata):
        try:
            return self.__command_factory_by_name[command_name].create(args, tdata)
        except KeyError:
            return self.__command_factory_by_name["help"].create(args, tdata)

command_factory_by_name = real_command_factory_by_name.copy()
# add help now
command_factory_by_name['help'] = HelpCommand

commands = Commands(command_factory_by_name)