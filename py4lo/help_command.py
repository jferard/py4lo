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
    def execute(self, args, tdata):
        msg = DEFAULT_MSG
        if len(args) == 1:
            try:
                msg = genuine_commands[args[0]].get_help()
            except KeyError:
                if args[0] == 'help':
                    msg = self.get_help()
        print (msg)

    def get_help(self):
        return "You must be kitting!"

help_command = HelpCommand()
