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

class Comparator():
    def __init__(self, accepted_locals = {}):
        self.__accepted_locals = accepted_locals

    def check(self, arg1, comparator, arg2):
        arg1 = self.__parse_expr(arg1)
        arg2 = self.__parse_expr(arg2)
        cmp_result = self.__cmp(arg1, arg2)
        return (   cmp_result == -1 and comparator in ["<", "<="]
                or cmp_result ==  0 and comparator in ["<=", "==", ">="]
                or cmp_result ==  1 and comparator in [">", ">="])

    def __parse_expr(self, expr):
        if expr[0] == '$':
            name = expr[1:]
            expr = eval(name, {'__builtin__':None}, self.__accepted_locals)
        try:
            return float(expr)
        except ValueError:
            return expr

    def __cmp(self, arg1, arg2):
        if arg1 < arg2:
            return -1
        elif arg1 == arg2:
            return 0
        else:
            return 1