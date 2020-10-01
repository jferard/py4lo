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

from core.script import SourceScript
from directives.directive import Directive


class EmbedLib(Directive):
    """
    Embed a py4lo library
    """

    @staticmethod
    def sig_elements():
        return ["embed", "lib"]

    def __init__(self, lib_dir: Path):
        self._lib_dir = lib_dir

    def execute(self, processor: "DirectiveProcessor",
                _line_processor: "DirectiveLineProcessor", args):
        script_ref = args[0]
        # TODO : sript_ref might be a dir (script_ref/__init__.py)
        script_path = self._lib_dir.joinpath(script_ref + ".py")
        processor.append_script(SourceScript(script_path, self._lib_dir))
        return True
