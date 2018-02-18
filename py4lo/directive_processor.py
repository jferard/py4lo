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

from directives import UseLib, UseObject, Include, ImportLib, Import, Fail
from branch_processor import BranchProcessor
from comparator import Comparator

class DirectiveProcessor():
    def __init__(self, scripts_processor, python_version, scripts_path):
        self.__scripts_processor = scripts_processor
        self.__scripts_path = scripts_path
        self.__py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")
        comparator = Comparator({'python_version':python_version})
        def local_is_true(args):
            return comparator.check(args[0], args[1], args[2])

        self.__local_is_true = local_is_true
        self.__directives = [UseLib(self.__py4lo_path), UseObject(self.__scripts_path), Include(self.__scripts_path), ImportLib(self.__py4lo_path), Import(scripts_path), Fail()]

        self.__bootstrapped = False
        self.__imported = False
        self.__branch_processor = BranchProcessor(local_is_true)

    def end(self):
        self.__branch_processor.end()

    def process_line(self, line):
        """Process a line that starts with #"""
        self.__s = ""
        try:
            ls = shlex.split(line)
            if self.__is_directive(ls):
                directiveName = ls[2]
                args = ls[3:]
                is_branch_directive = self.__branch_processor.handle_directive(directiveName, args)
                if not is_branch_directive:
                    if self.__branch_processor.skip():
                        self.__s += "### "+line
                    else:
                        for directive in self.__directives:
                            if directive.execute(self, directiveName, args):
                                break
            else: # thats maybe a simple comment
                if self.__branch_processor.skip():
                    self.__s += "### "+line
                else:
                    self.__s += line
        except ValueError:
            if self.__branch_processor.skip():
                self.__s += "### "+line
            else:
                self.__s += line

        return self.__s

    def bootstrap(self):
        if not self.__bootstrapped:
            self.__s += self.__get_bootstrap()
            self.__bootstrapped = True

    def import2(self):
        if not self.__imported:
            self.__s += self.__get_import()
            self.__imported = True

    def append(self, s2):
        self.__s += s2

    def append_script(self, script_fname):
        """Append a script to the script processor"""
        self.__scripts_processor.append_script(script_fname)

    def __is_directive(self, ls):
        return len(ls) >= 2 and ls[0] == '#' and ls[1] == 'py4lo:'

    def __get_bootstrap(self):
        return self.__get_inc_file("py4lo_bootstrap.py")

    def __get_import(self):
        return self.__get_inc_file("py4lo_import.py")

    def __get_inc_file(self, fname):
        path = os.path.dirname(os.path.realpath(__file__))

        inc = ""
        state = 0
        with open(os.path.join(path, "..", "inc", fname), 'r', encoding="utf-8") as b:
            for line in b.readlines():
                l = line.strip()
                if state == 0:
                    if l.startswith('#'):
                        pass
                    elif l.startswith("\"\"\""):
                        state = 1
                    else:
                        inc += line
                elif state == 1:
                    if l.endswith("\"\"\""):
                        state = 0
        return inc + "\n"

    def ignore_lines(self):
        return self.__branch_processor.skip();
