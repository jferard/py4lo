#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>
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
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

from toml_helper import load_toml


class PropertiesProvider:
    def __init__(self, toml_filename="py4lo.toml"):
        self._toml_filename = toml_filename
        self._tdata = None

    def get(self) -> Dict[str, Any]:
        if self._tdata is None:
            self._tdata = load_toml(Path(self._toml_filename))

        return self._tdata


class Command(ABC):
    @staticmethod
    def create_executor(args: List[str],
                        provider: PropertiesProvider) -> "CommandExecutor":
        pass

    @abstractmethod
    def execute(self, *args: List[str]) -> Optional[Tuple]:
        pass

    @staticmethod
    def get_help():
        pass
