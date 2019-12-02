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
from zipfile import ZipFile

from callbacks import AfterCallback


class AddReadmeWith(AfterCallback):
    """
    Add a readme in destination file
    """

    def __init__(self, inc_path: Path, contact: str):
        self._inc_path = inc_path
        self._contact = contact

    def call(self, zout: ZipFile) -> bool:
        zout.write(self._inc_path.joinpath("script-lc.xml"),
                   "Basic/script-lc.xml")
        zout.write(self._inc_path.joinpath("script-lb.xml"),
                   "Basic/Standard/script-lb.xml")
        with self._inc_path.joinpath("py4lo.xml.tpl").open('r',
                                                           encoding='utf-8') as f:
            tpl = f.read()
            xml = tpl.format(contact=self._contact)
            zout.writestr("Basic/Standard/py4lo.xml", xml)
        return True
