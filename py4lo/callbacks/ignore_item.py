# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
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
from pathlib import Path
from zipfile import ZipFile, ZipInfo

from callbacks import ItemCallback


class IgnoreItem(ItemCallback):
    """
    Item callback. Ignore all existing scripts in source file.
    """

    def __init__(self, arc_scripts_path: Path):
        self._arc_scripts_path = arc_scripts_path

    def call(self, _zin: ZipFile, _zout: ZipFile, item: ZipInfo) -> bool:
        """
        @param _zin:
        @param _zout:
        @param item: the item to process
        @return: True if the item is in "Scripts/python"
        """
        return self._arc_scripts_path in Path(item.filename).parents
