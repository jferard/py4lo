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

from test_command import TestCommand
from debug_command import DebugCommand

class InitCommand():
    @staticmethod
    def create(args, tdata):
        test_executor = TestCommand.create(args, tdata)
        logger = logging.getLogger("py4lo")
        logger.setLevel(tdata["log_level"])
        init_command = DebugCommand(
            logger,
            tdata["src_dir"],
            tdata["target_dir"],
            tdata["python_version"],
            tdata["init_file"],
            "Creates a standard file"
        )
        return CommandExecutor(init_command, test_executor)
