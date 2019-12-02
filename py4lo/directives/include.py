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


class Include(Directive):
    """Include the content of a file inside the script. The file should be in a inc directory"""

    @staticmethod
    def sig_elements():
        return ["include"]

    def __init__(self, _py4lo_path: Path, scripts_path: Path):
        self._scripts_path = scripts_path

    def execute(self, processor, args: List[str]):
        path = self._scripts_path.joinpath("inc", args[0] + ".py")

        s = "# begin py4lo include: {}\n".format(path)
        with path.open('r', encoding='utf-8') as f:
            for line in f:
                s += line
        s += "\n# end py4lo include\n"

        processor.append(s)
        return True
