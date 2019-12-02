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
import argparse

from commands import commands
from commands.command import PropertiesProvider

parser = argparse.ArgumentParser(description="Python for LibreOffice",
                                 formatter_class=argparse.RawTextHelpFormatter)
parser.add_argument("-t", "--toml", help="the toml file",
                    default="py4lo.toml", type=str)
parser.add_argument("command",
                    help=commands.get_help_message(), type=str)
parser.add_argument("parameter", nargs="*",
                    help="command parameter")
args = parser.parse_args()

command = commands.get(args.command, args.parameter,
                       PropertiesProvider(args.toml))
command.execute()
