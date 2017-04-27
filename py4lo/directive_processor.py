# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import shlex
import os

from directive import UseLib, UseObject, Include, Import, Fail

class DirectiveProcessor():
    def __init__(self, python_version, scripts_path, cur_script_fnames):
        self.__scripts_path = scripts_path
        self.__cur_script_fnames = cur_script_fnames
        self.__py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")
        def local_is_true(args):
            assert args[0] == "python_version"
            return is_true(python_version, args[1], args[2])
        
        self.__branch_processor = BranchProcessor(local_is_true)
        self.__directives = [UseLib(self.__py4lo_path), UseObject(self.__scripts_path), Include(self.__scripts_path), Import(self.__py4lo_path), Fail()]
        
    def new_script(self):
        self.__bootstrapped = False
        self.__imported = False
        self.__branch_processor.new_script()

    def process_line(self, line):
        """Process a line that starts with #"""
        self.__s = ""
        try:
            ls = shlex.split(line)
            if self.__is_directive(ls):
                directiveName = ls[2]
                args = ls[3:]
                isBranchDirective = self.__branch_processor.handle_directive(directiveName, args)
                if not isBranchDirective:
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

    def appendScript(self, script_fname, lib_ref_py):
        self.__cur_script_fnames.append((script_fname, lib_ref_py))

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

class BranchProcessor():
    def __init__(self, tester):
        self.__tester = tester
        self.__dont_skips = []
        
    def new_script(self):
        if len(self.__dont_skips):
            print ("Branch condition not closed!")
            
        self.__dont_skips = []

    def handle_directive(self, directive, args):
        if directive == 'if':
            self.__dont_skips.append(self.__tester(args))
        elif directive == 'elif':
            if self.__dont_skips[-1]:
                self.__dont_skips[-1] = False
            elif self.__tester(args):
                self.__dont_skips[-1] = True
            # else : self.__dont_skips stays False
        elif directive == 'else':
            self.__dont_skips[-1] = not self.__dont_skips[-1]
        elif directive == 'endif':
            self.__dont_skips.pop()
        else:
            return False
        
        return True
            
    def skip(self):
        for b in self.__dont_skips:
            if not b:
                return True
                
        return False

def is_true(str1, comparator, str2):
    if str1 < str2:
        cmp = -1
    elif str1 == str2:
        cmp = 0
    else:
        cmp = 1

    return cmp == -1 and comparator in ["<", "<="] or cmp == 0 and comparator in ["<=", "==", ">="] or cmp == 1 and comparator in [">", ">="]
