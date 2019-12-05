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
import shlex
from pathlib import Path
from typing import List, Optional

from branch_processor import BranchProcessor
from comparator import Comparator
from core.script import SourceScript
from directives import DirectiveProvider


class DirectiveProcessor:
    """A DirectiveProcessor processes directives, ie line that begins with #,
    in scripts. When a directive is parsed, the processor is passed to the
    directive, which may use some helpers: import2, ..."""

    @staticmethod
    def create(scripts_processor: "ScriptSetProcessor", directive_provider: DirectiveProvider, python_version: str,
               source_script: SourceScript):
        comparator = Comparator({'python_version': python_version})

        def local_is_true(args: List[str]):
            return comparator.check(args[0], args[1], args[2])

        branch_processor = BranchProcessor(local_is_true)

        return DirectiveProcessor(scripts_processor, branch_processor,
                                  directive_provider)

    def __init__(self, script_set_processor: "ScriptSetProcessor",
                 branch_processor: BranchProcessor,
                 directive_provider: DirectiveProvider):
        """
        Create a Directive processor. Scripts_path is the path to the scripts
        directory
        """
        self._script_set_processor = script_set_processor
        self._branch_processor = branch_processor
        self._directive_provider = directive_provider
        self._includes = set()

    def append_script(self, source_script: SourceScript):
        """Append a script to the script processor"""
        self._script_set_processor.append_script(source_script)

    def process_line(self, line: str) -> List[str]:
        """Process a line that starts with #"""
        return _DirectiveProcessorWorker(self, self._branch_processor,
                                         self._directive_provider,
                                         line).process_line()

    def include(self, fname: str) -> List[str]:
        s = [""]
        if fname not in self._includes:
            s = _IncludeProcessor(fname).process()
            self._includes.add(fname)

        return s

    def end(self):
        """Verify the end of the scripts"""
        self._branch_processor.end()

    def ignore_lines(self) -> bool:
        return self._branch_processor.skip()


class _DirectiveProcessorWorker:
    """A worker that processes the line"""

    def __init__(self, directive_processor: DirectiveProcessor,
                 branch_processor: BranchProcessor,
                 directive_provider: DirectiveProvider, line: str):
        self._directive_processor = directive_processor
        self._branch_processor = branch_processor
        self._directive_provider = directive_provider
        self._line = line
        self._target_lines = []

    def process_line(self) -> List[str]:
        """Return a list of lines"""
        if self._target_lines:
            return self._target_lines

        try:
            directive = self._get_directive()
            if directive is None: # that's maybe a simple comment
                self._comment_or_write()
            else:
                self._process_directive(self._line, directive)
        except ValueError:
            self._comment_or_write()

        return self._target_lines

    def _get_directive(self) -> Optional[List[str]]:
        mark, directive = self._line.split(":", 1)
        mark = mark.split()
        if mark == ['#py4lo'] or mark == ['#', 'py4lo']:
            return shlex.split(directive)
        else:
            return None

    def _process_directive(self, line: str, args: List[str]):
        is_branch_directive = self._branch_processor.handle_directive(args[0],
                                                                      args[1:])
        if is_branch_directive:
            return

        if self._branch_processor.skip():
            self.append("### " + line)
        else:
            try:
                directive, args = self._directive_provider.get(args)
                directive.execute(self, args)
            except KeyError:
                print("Wrong directive ({})".format(line.strip()))

    def _comment_or_write(self):
        if self._branch_processor.skip():
            self.append("### " + self._line)
        else:
            self.append(self._line)

    def include(self, fname: str):
        """Called by directives"""
        self._target_lines.extend(self._directive_processor.include(fname))

    def append(self, line: str):
        """Called by directives"""
        self._target_lines.append(line)

    def append_script(self, source_script: SourceScript):
        """Append a script to the script processor"""
        self._directive_processor.append_script(source_script)


class _IncludeProcessor:
    """ A simple include stripper: remove docstrings and comments"""

    DOC_STRING_OPEN = "\"\"\""
    DOC_STRING_CLOSE = "\"\"\""
    NORMAL = 0
    IN_DOC_STRING = 1
    END = -1

    def __init__(self, fname: str):
        self._fname = fname
        self._inc_lines = []
        self._state = _IncludeProcessor.NORMAL

    def process(self) -> List[str]:
        """Return a list of lines"""
        if self._state != _IncludeProcessor.NORMAL:
            raise Exception("Create a new IncludeProcessor")

        path = Path(__file__).parent
        fpath = path.joinpath("..", "inc", self._fname)
        with fpath.open('r', encoding="utf-8") as b:
            for line in b.readlines():
                self._process_line(line)

        self._state = _IncludeProcessor.END
        return self._inc_lines

    def _process_line(self, line: str):
        line = line.rstrip()
        stripped_line = line.strip()
        if self._state == _IncludeProcessor.NORMAL:
            if stripped_line.startswith('#'):
                pass
            elif stripped_line.startswith(_IncludeProcessor.DOC_STRING_OPEN):
                self._state = _IncludeProcessor.IN_DOC_STRING
            else:
                self._inc_lines.append(line)
        elif self._state == _IncludeProcessor.IN_DOC_STRING:
            if stripped_line.endswith(_IncludeProcessor.DOC_STRING_CLOSE):
                self._state = _IncludeProcessor.NORMAL
