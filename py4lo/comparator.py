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
from abc import ABC, abstractmethod
from typing import Any, Union, Dict


class Comparable(ABC):
    @abstractmethod
    def __lt__(self, other: Any) -> bool:
        pass


T = Union[int, float, str]


class Comparator:
    """A comparator will be used to evaluate expressions. Used by the branch
    processor"""

    def __init__(self, accepted_locals: Dict[str, T] = None):
        """accepted_locals are """
        if accepted_locals is None:
            self._accepted_locals = {}
        else:
            self._accepted_locals = accepted_locals

    def check(self, arg1: T, comparator: str, arg2: T):
        """Check arg1 vs arg2 using comparator. Args may be $var where var is
        a member of accepted_locals, a number 123456i, 123.456f, an expression
        or a litteral."""
        arg1 = self._parse_expr(arg1)
        arg2 = self._parse_expr(arg2)
        cmp_result = self._cmp(arg1, arg2)
        return (cmp_result == -1 and comparator in ["<", "<="]
                or cmp_result == 0 and comparator in ["<=", "==", ">="]
                or cmp_result == 1 and comparator in [">", ">="])

    def _parse_expr(self, expr: str) -> T:
        if expr[0] == '$':
            name = expr[1:]
            expr = eval(name, {'__builtin__': None}, self._accepted_locals)
        if not isinstance(expr, str):
            return expr
        try:
            if expr[-1] == 'f':
                return float(expr[0:-1])
            elif expr[-1] == 'i':
                return float(expr[0:-1])
        except ValueError:
            pass

        return expr

    @staticmethod
    def _cmp(arg1: T, arg2: T) -> int:
        if arg1 < arg2:
            return -1
        elif arg1 == arg2:
            return 0
        else:
            return 1
