# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2021 J. Férard <https://github.com/jferard>

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

from commands.debug_command import DebugCommand
from commands.init_command import InitCommand
from commands.test_command import TestCommand
from commands.run_command import RunCommand
from commands.update_command import UpdateCommand


real_command_factory_by_name = {
    'debug': DebugCommand,
    'init': InitCommand,
    'test': TestCommand,
    'run': RunCommand,
    'update': UpdateCommand,
}
