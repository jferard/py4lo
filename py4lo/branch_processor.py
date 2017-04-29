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
        self.__assertion_is_true = tester
        self.__dont_skips = []

    def new_script(self):
        if len(self.__dont_skips):
            logging.error("Branch condition not closed!")
            raise ValueError("Branch condition not closed!")

        self.__dont_skips = []

    def handle_directive(self, directive, args):
        if directive == 'if':
            self.__begin_block_and_skip_if_not(self.__assertion_is_true(args))
        elif directive == 'elif':
            if self.__was_not_skipping():
                self.__start_skipping()
            elif self.__assertion_is_true(args):
                self.__stop_skipping()
            # else : continue to skip
        elif directive == 'else':
            self.__flip_skipping()
        elif directive == 'endif':
            self.__end_block()
        else:
            return False

        return True

    def __was_not_skipping(self):
        return self.__dont_skips[-1]

    def __start_skipping(self):
        self.__dont_skips[-1] = False

    def __stop_skipping(self):
        self.__dont_skips[-1] = True

    def __flip_skipping(self):
        self.__dont_skips[-1] = not self.__dont_skips[-1]

    def __begin_block_and_skip_if_not(self, b):
        self.__dont_skips.append(b)

    def __end_block(self):
        self.__dont_skips.pop()

    def skip(self):
        return False in self.__dont_skips

def is_true(str1, comparator, str2):
    if str1 < str2:
        cmp = -1
    elif str1 == str2:
        cmp = 0
    else:
        cmp = 1

    return cmp == -1 and comparator in ["<", "<="] or cmp == 0 and comparator in ["<=", "==", ">="] or cmp == 1 and comparator in [">", ">="]
