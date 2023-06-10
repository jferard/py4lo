# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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
from typing import Sequence

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
                line_processor: "DirectiveLineProcessor", args):
        lib_ref = args[0]
        lib_path = self._lib_dir.joinpath(lib_ref)
        source_scripts = self._embed(lib_path)
        for source_script in source_scripts:
            processor.append_script(source_script)

        if lib_ref == "py4lo_helper":
            line_processor.append(
                """# begin py4lo: init py4lo_helper
import py4lo_helper
try:
    py4lo_helper.init(XSCRIPTCONTEXT)
except NameError:
    pass
finally:
    del py4lo_helper  # does not wipe cache, but remove the access.
# end py4lo: init py4lo_helper""")
        return True

    def _embed(self, lib_path: Path) -> Sequence[SourceScript]:
        stack = [lib_path]
        ret = []

        while stack:
            path = stack.pop()
            if path.is_dir():
                stack.extend(lib_path.glob("*.py"))
            elif path.suffix == "" or path.suffix == ".py":
                path = path.with_suffix(".py")
                with path.open('rb') as f:
                    temp_script = SourceScript(
                        path, self._lib_dir, False)
                ret.append(temp_script)

        return ret
