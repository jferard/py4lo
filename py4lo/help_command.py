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

from genuine_commands import genuine_commands

DEFAULT_MSG = """usage: py4lo.py -h|--help|command [args]

Python for LibreOffice.

-h, --help  show this help message and exit
command     a command = debug|help|init|test|update
        debug:          creates a debug.ods file with button for each function
        help:           show this message
        help [command]: more specific help
        init:           create a standard file
        test:           test the scripts
        run:            update + open the created file
        update:         updates the file with all scripts"""

class HelpCommand():
    @staticmethod
    def create(args, tdata):
        if len(args) == 1:
            command_name = args[0]
        else:
            command_name = None
        return HelpCommand(genuine_commands, command_name)

    def __init__(self, genuine_commands, command_name = None):
        self.__genuine_commands = genuine_commands
        self.__command_name = command_name

    def execute(self):
        msg = DEFAULT_MSG
        if self.__command_name:
            try:
                msg = self.__genuine_commands[self.__command_name].create().get_help()
            except KeyError:
                if self.__command_name == 'help':
                    msg = self.get_help()
        print (msg)

    def get_help(self):
        return "You must be kitting!"
