# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
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
import logging
from typing import List, Callable, Any, cast


class BranchProcessor:
    """
    The branch processor handles directives like 'if', 'elif', 'else', 'end'
    and acts as a preprocessor. Skipped block wont be included in the LO
    document
    """
    def __init__(self, tester: Callable[[List[str]], bool]):
        """The tester will evaluate the arguments of 'if' or 'elif'"""
        self._assertion_is_true = tester
        self._dont_skips = cast(List[bool], [])

    def end(self):
        """
        To call before the end, to verify if there are no unclosed if block
        """
        if self._dont_skips:
            logging.error("Branch condition not closed!")
            raise ValueError("Branch condition not closed!")

    def handle_directive(self, directive: str, args: List[Any]) -> bool:
        """Return True if the next block should be read, False otherwise"""

        if directive == 'if':
            self._begin_block_and_skip_if_not(self._assertion_is_true(args))
        elif directive == 'elif':
            if self._was_not_skipping():
                self._start_skipping()
            elif self._assertion_is_true(args):
                self._stop_skipping()
            # else : continue to skip
        elif directive == 'else':
            self._flip_skipping()
        elif directive == 'endif':
            self._end_block()
        else:
            return False

        return True

    def _was_not_skipping(self) -> bool:
        return self._dont_skips[-1]

    def _start_skipping(self):
        self._dont_skips[-1] = False

    def _stop_skipping(self):
        self._dont_skips[-1] = True

    def _flip_skipping(self):
        self._dont_skips[-1] = not self._dont_skips[-1]

    def _begin_block_and_skip_if_not(self, b: bool):
        self._dont_skips.append(b)

    def _end_block(self):
        self._dont_skips.pop()

    def skip(self) -> bool:
        """Return True if the current block is skipped"""
        return False in self._dont_skips
