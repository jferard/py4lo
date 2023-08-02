#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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
from abc import ABC, abstractmethod
from typing import List, Any


class Directive(ABC):
    @abstractmethod
    def execute(self, processor: Any,  # "DirectiveProcessor",
                line_processor: Any,  # "DirectiveLineProcessor",
                args: List[str]):
        """Execute the directive. May append a script to process"""
        pass

    @staticmethod
    @abstractmethod
    def sig_elements() -> List[str]:
        """List of tokens that identify the directive"""
        pass
