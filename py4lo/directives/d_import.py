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
from pathlib import Path

from directives.directive import Directive


class Import(Directive):
    @staticmethod
    def sig_elements():
        return ["import"]

    def __init__(self, _base_path: Path, scripts_path: Path):
        self._scripts_path = scripts_path

    def execute(self, processor: "DirectiveProcessor", args):
        processor.include("py4lo_import.py")
        script_ref = args[0]
        script_path = self._scripts_path.joinpath(script_ref + ".py")
        processor.append_script(script_path)
        processor.append("import " + script_ref + "\n")
        return True
