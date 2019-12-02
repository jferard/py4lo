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
import os
from pathlib import Path
from typing import List

from directives.directive import Directive


class Embed(Directive):
    """Embed a file. It's up to the user to import it if it is a module"""

    @staticmethod
    def sig_elements():
        return ["embed"]

    def __init__(self, _base_path: Path, _scripts_path: Path):
        pass

    def execute(self, processor, args: List[str]):
        processor.include("py4lo_import.py")
        fname = args[0]
        processor.append_script(fname)
        return True
