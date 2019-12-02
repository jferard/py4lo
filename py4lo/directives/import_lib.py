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


class ImportLib(Directive):
    @staticmethod
    def sig_elements():
        return ["import", "lib"]

    def __init__(self, base_path: Path, _scripts_path: Path):
        self._base_path = base_path

    def execute(self, processor, args):
        processor.include("py4lo_import.py")
        script_ref = args[0]
        script_fname = self._base_path.joinpath("lib", script_ref + ".py")
        processor.append_script(script_fname)
        processor.append("import " + script_ref + "\n")
        return True
