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
import logging

class BranchProcessor():
    def __init__(self, tester):
        self.__tester = tester
        self.__dont_skips = []
        
    def new_script(self):
        if len(self.__dont_skips):
            logging.error("Branch condition not closed!")
            
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
