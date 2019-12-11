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
from typing import List, Sequence

from core.script import TempScript
from directives.directive import Directive


class EmbedScript(Directive):
    """
    Embed a script in the archive and... do nothing else.
    It's up to the user to import it if it is a module
    """

    @staticmethod
    def sig_elements():
        return ["embed", "script"]

    def __init__(self, opt_dir: Path):
        self._opt_dir = opt_dir

    def execute(self, processor: "DirectiveProcessor",
                _line_processor: "DirectiveLineProcessor", args: List[str]):
        path = Path(args[0])
        fullpath = self._opt_dir.joinpath(path)
        temp_scripts = self._embed(fullpath)
        for temp_script in temp_scripts:
            processor.add_script(temp_script)
        return True

    def _embed(self, fullpath: Path) -> Sequence[TempScript]:
        if fullpath.is_dir():
            temp_scripts = []
            for p in fullpath.iterdir():
                temp_scripts.extend(self._embed(p))
        else:
            with fullpath.open('rb') as f:
                temp_script = TempScript(fullpath, f.read(), self._opt_dir, [],
                                         None)
            return [temp_script]
