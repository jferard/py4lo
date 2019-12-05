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
from typing import List, Tuple, Optional

from commands.command import Command


class CommandExecutor:
    def __init__(self, command: Command,
                 previous_executor: Optional["CommandExecutor"] = None):
        self._command = command
        self._previous_executor = previous_executor

    def execute(self, *args: List[str]) -> Optional[Tuple]:
        if self._previous_executor is None:
            cur_args = []
        else:
            cur_args = self._previous_executor.execute(*args)

        logging.warning("%s, args=%s", self._command, cur_args)
        ret = self._command.execute(*cur_args)
        logging.warning("%s, args=%s, ret=%s", self._command, cur_args, ret)
        return ret
