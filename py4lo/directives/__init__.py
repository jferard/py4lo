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
from pathlib import Path
from typing import List, Type, Dict

from directives.d_import import Import
from directives.directive import Directive
from directives.embed import Embed
from directives.import_lib import ImportLib
from directives.include import Include


class _DirectiveProviderFactory:
    def __init__(self,
                 directive_classes: List[Type[Directive]]):
        self._directive_classes = directive_classes

    def create(self, base_path: Path, scripts_path: Path):
        self._directives_tree = {}
        for d in self._directive_classes:
            sig_elements = d.sig_elements()
            assert len(sig_elements)

            directive = d(base_path, scripts_path)

            self._put_directive_class(sig_elements, directive)

        return DirectiveProvider(self._directives_tree)

    def _put_directive_class(self, sig_elements: List[str],
                             directive: Directive):
        cur_directives_tree = self._directives_tree
        for fst in sig_elements:
            if fst not in cur_directives_tree:
                cur_directives_tree[fst] = {}
            cur_directives_tree = cur_directives_tree[fst]

        cur_directives_tree.update({"@": directive})


T = Dict[str, "T"]


class DirectiveProvider:
    @staticmethod
    def create(base_path: Path, scripts_path: Path):
        return _DirectiveProviderFactory(
            [Include, ImportLib, Import, Embed]).create(base_path,
                                                        scripts_path)

    def __init__(self, directives_tree: T):
        self._directives_tree = directives_tree

    def get(self, args: List[str]) -> (T, List[str]):
        """args are the shlex result"""
        cur_directives_tree = self._directives_tree

        for i in range(len(args)):
            arg = args[i]
            if arg in cur_directives_tree:
                cur_directives_tree = cur_directives_tree[arg]
            elif "@" in cur_directives_tree:
                return cur_directives_tree["@"], args[i:]
            else:
                raise KeyError(args)

        assert False, "no args"
