# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

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
import os

from directives import DirectiveProvider
from branch_processor import BranchProcessor
from comparator import Comparator

class DirectiveProcessor():
    """A DirectiveProcessor processes directives, ie line that begins with #,
    in scripts. When a directive is parsed, the processor is passed to the
    directive, which may use some helpers: import2, ..."""

    @staticmethod
    def create(scripts_path, scripts_processor, python_version):
        comparator = Comparator({'python_version':python_version})
        def local_is_true(args):
            return comparator.check(args[0], args[1], args[2])

        branch_processor = BranchProcessor(local_is_true)

        py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")
        directive_provider = DirectiveProvider.create(py4lo_path, scripts_path)

        return DirectiveProcessor(scripts_path, scripts_processor, branch_processor, directive_provider)

    def __init__(self, scripts_path, scripts_processor, branch_processor, directive_provider):
        """Create a Directive processor. Scripts_path is the path to the scripts directory"""
        self.__scripts_processor = scripts_processor
        self.__scripts_path = scripts_path
        self.__branch_processor = branch_processor
        self.__directive_provider = directive_provider
        self.__includes = set()

    def append_script(self, script_fname):
        """Append a script to the script processor"""
        self.__scripts_processor.append_script(script_fname)

    def process_line(self, line):
        """Process a line that starts with #"""
        return _DirectiveProcessorWorker(self, self.__branch_processor, self.__directive_provider, line).process_line()

    def include(self, fname):
        s = ""
        if fname not in self.__includes:
            s = _IncludeProcessor(fname).process()
            self.__includes.add(fname)

        return s

    def end(self):
        """Verify the end of the scripts"""
        self.__branch_processor.end()

    def ignore_lines(self):
        return self.__branch_processor.skip();

class _DirectiveProcessorWorker():
    """A worker that processes the line"""

    def __init__(self, directive_processor, branch_processor, directive_provider, line):
        self.__directive_processor = directive_processor
        self.__branch_processor = branch_processor
        self.__directive_provider = directive_provider
        self.__line = line
        self.__target_lines = []

    def process_line(self):
        """Return a list of lines"""
        if self.__target_lines:
            return self.__target_lines

        try:
            ls = shlex.split(self.__line)
            if self.__is_directive(ls):
                self.__process_directive(ls[2:])
            else: # thats maybe a simple comment
                self.__comment_or_write()
        except ValueError:
            self.__comment_or_write()

        return self.__target_lines

    def __is_directive(self, ls):
        return len(ls) >= 2 and ls[0] == '#' and ls[1] == 'py4lo:'

    def __process_directive(self, args):
        is_branch_directive = self.__branch_processor.handle_directive(args[0], args[1:])
        if is_branch_directive:
            return

        if self.__branch_processor.skip():
            self.append("### "+line)
        else:
            try:
                directive, args = self.__directive_provider.get(args)
                directive.execute(self, args)
            except KeyError as e:
                print("Wrong directive ({})".format(line.strip()))

    def __comment_or_write(self):
        if self.__branch_processor.skip():
            self.append("### "+self.__line)
        else:
            self.append(self.__line)

    def include(self, fname):
        """Called by directives"""
        self.__target_lines.extend(self.__directive_processor.include(fname))

    def append(self, line):
        """Called by directives"""
        self.__target_lines.append(line)

    def append_script(self, script_fname):
        """Append a script to the script processor"""
        self.__directive_processor.append_script(script_fname)


class _IncludeProcessor():
    """ A simple include stripper: remove docstrings and comments"""

    DOC_STRING_OPEN = "\"\"\""
    DOC_STRING_CLOSE = "\"\"\""
    NORMAL = 0
    IN_DOC_STRING = 1
    END = -1

    def __init__(self, fname):
        self.__fname = fname
        self.__inc_lines = []
        self.__state = _IncludeProcessor.NORMAL

    def process(self):
        """Return a list of lines"""
        if self.__state != _IncludeProcessor.NORMAL:
            raise Exception("Create a new IncludeProcessor")

        path = os.path.dirname(os.path.realpath(__file__))
        fpath = os.path.join(path, "..", "inc", self.__fname)
        with open(fpath, 'r', encoding="utf-8") as b:
            for line in b.readlines():
                self.__process_line(line)

        self.__state = _IncludeProcessor.END
        return self.__inc_lines

    def __process_line(self, line):
        l = line.strip()
        if self.__state == _IncludeProcessor.NORMAL:
            if l.startswith('#'):
                pass
            elif l.startswith(_IncludeProcessor.DOC_STRING_OPEN):
                self.__state = _IncludeProcessor.IN_DOC_STRING
            else:
                self.__inc_lines.append(line)
        elif self.__state == _IncludeProcessor.IN_DOC_STRING:
            if l.endswith(_IncludeProcessor.DOC_STRING_CLOSE):
                self.__state = _IncludeProcessor.NORMAL
