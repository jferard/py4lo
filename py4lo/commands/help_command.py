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

from commands.real_command_factory_by_name import real_command_factory_by_name

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


class HelpCommand:
    @staticmethod
    def create(args, _tdata):
        if len(args) == 1:
            command_name = args[0]
        else:
            command_name = None
        return HelpCommand(real_command_factory_by_name, command_name)

    def __init__(self, command_factory_by_name, command_name=None):
        self._command_factory_by_name = command_factory_by_name
        self._command_name = command_name

    def execute(self):
        msg = DEFAULT_MSG
        if self._command_name:
            try:
                msg = self._command_factory_by_name[self._command_name].get_help()
            except KeyError as e:
                print (e)
                if self._command_name == 'help':
                    msg = self.get_help()
        print(msg)

    @staticmethod
    def get_help():
        return "You must be kitting!"
