#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2026 J. Férard <https://github.com/jferard>
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
from typing import Any

from core.properties import PropertiesProvider


class Command(ABC):
    @staticmethod
    def create_executor(args: list[str], provider: PropertiesProvider
                        ) -> Any:  # "CommandExecutor":
        pass

    @abstractmethod
    def execute(self, *args: list[str]) -> tuple[Any, ...]:
        pass

    @abstractmethod
    def get_help(self) -> str:
        pass
