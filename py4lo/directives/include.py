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
import os
from pathlib import Path
from typing import List

from directives.directive import Directive


class Include(Directive):
    """
    Include the content of a file inside the script.
    The file should be in a inc directory
    """

    @staticmethod
    def sig_elements() -> List[str]:
        return ["include"]

    def __init__(self, inc_dir: Path):
        self._inc_dir = inc_dir

    def execute(self, _processor: "DirectiveProcessor",
                line_processor: "DirectiveLineProcessor", args: List[str]):
        path = self._inc_dir.joinpath(args[0])
        if len(args) >= 2:
            strip = bool(args[1])
        else:
            strip = False

        s = ["# begin py4lo include: {}".format(path)]
        if strip:
            s.extend(IncludeStripper(path).process())
        else:
            with path.open('r', encoding='utf-8') as f:
                s.extend(f)
        s.append("# end py4lo include\n")

        line_processor.append("\n".join(s))
        return True


class IncludeStripper:
    """ A simple include stripper: remove docstrings and comments"""

    DOC_STRING_OPEN = '"""'
    DOC_STRING_CLOSE = '"""'
    DOC_STRING_SINGLE_OPEN = "'''"
    DOC_STRING_SINGLE_CLOSE = "'''"
    NORMAL = 0
    IN_DOC_STRING = 1
    IN_DOC_STRING_SINGLE = 2
    END = -1

    def __init__(self, path: Path):
        self._path = path
        self._inc_lines = []
        self._state = IncludeStripper.NORMAL

    def process(self) -> List[str]:
        """Return a list of lines"""
        if self._state != IncludeStripper.NORMAL:
            raise Exception("Create a new IncludeProcessor")

        with self._path.open('r', encoding="utf-8") as b:
            for line in b:
                self._process_line(line)

        self._state = IncludeStripper.END
        while self._inc_lines and not self._inc_lines[-1]:
            del self._inc_lines[-1]
        return self._inc_lines

    def _process_line(self, line: str):
        line = line.rstrip()
        stripped_line = line.strip()
        if self._state == IncludeStripper.NORMAL:
            if stripped_line.startswith('#'):
                pass
            elif stripped_line.startswith(IncludeStripper.DOC_STRING_OPEN):
                if self.is_open_doc_string(stripped_line,
                                           IncludeStripper.DOC_STRING_OPEN,
                                           IncludeStripper.DOC_STRING_CLOSE):
                    self._state = IncludeStripper.IN_DOC_STRING
            elif stripped_line.startswith(
                    IncludeStripper.DOC_STRING_SINGLE_OPEN):
                if self.is_open_doc_string(stripped_line,
                                           IncludeStripper.DOC_STRING_SINGLE_OPEN,
                                           IncludeStripper.DOC_STRING_SINGLE_CLOSE):
                    self._state = IncludeStripper.IN_DOC_STRING_SINGLE
            else:
                self._inc_lines.append(line)

        elif (self.is_close_doc_string(stripped_line,
                                       IncludeStripper.IN_DOC_STRING,
                                       IncludeStripper.DOC_STRING_CLOSE)
              or self.is_close_doc_string(stripped_line,
                                          IncludeStripper.IN_DOC_STRING_SINGLE,
                                          IncludeStripper.DOC_STRING_SINGLE_CLOSE)):
            self._state = IncludeStripper.NORMAL

    def is_open_doc_string(self, stripped_line: str, string_open: str,
                           string_close: str) -> bool:
        return (not stripped_line.endswith(string_close)
                or len(stripped_line) == len(string_open))

    def is_close_doc_string(self, stripped_line: str, doc: int,
                            string_close: str) -> bool:
        return self._state == doc and stripped_line.endswith(string_close)
