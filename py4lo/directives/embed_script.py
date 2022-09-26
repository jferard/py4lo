# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
#
#     This file is part of Py4LO.
#
#     Py4LO is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     Py4LO is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
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
    def sig_elements() -> List[str]:
        return ["embed", "script"]

    def __init__(self, opt_dir: Path):
        self._opt_dir = opt_dir

    def execute(self, processor: "DirectiveProcessor",
                _line_processor: "DirectiveLineProcessor", args: List[str]):
        script_ref = args[0]
        script_path = self._opt_dir.joinpath(script_ref)
        temp_scripts = self._embed(script_path)
        for temp_script in temp_scripts:
            processor.add_script(temp_script)
        return True

    def _embed(self, script_path: Path) -> Sequence[TempScript]:
        stack = [script_path]
        ret = []

        while stack:
            path = stack.pop()
            if path.is_dir():
                stack.extend(script_path.glob("*.py"))
            elif path.suffix == "" or path.suffix == ".py":
                path = path.with_suffix(".py")
                with path.open('rb') as f:
                    temp_script = TempScript(
                        path, f.read(), self._opt_dir, [], None)
                ret.append(temp_script)

        return ret
