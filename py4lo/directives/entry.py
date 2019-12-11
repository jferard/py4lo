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

from directives.include import Include
from directives.directive import Directive


class Entry(Directive):
    """
    This is an entry point. This will fix the python path.
    Use in the main file.
    """
    @staticmethod
    def sig_elements():
        return ["entry"]

    def __init__(self, inc_dir: Path):
        self._include_directive = Include(inc_dir)

    def execute(self, _processor: "DirectiveProcessor", line_processor: "DirectiveLineProcessor", args):
        return self._include_directive.execute(_processor, line_processor, ["py4lo_import.py", True])
