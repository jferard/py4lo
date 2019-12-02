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
import zipfile
from pathlib import Path
from zipfile import ZipFile

from callbacks.callback import BeforeAfterCallback, ItemCallback


class ZipUpdater:
    """A zip file updater. Applies callbacks before, after and to each item."""

    def __init__(self):
        self._before_callbacks = []
        self._item_callbacks = []
        self._after_callbacks = []

    def before(self, callback: BeforeAfterCallback):
        self._before_callbacks.append(callback)
        return self

    def item(self, callback: ItemCallback):
        self._item_callbacks.append(callback)
        return self

    def after(self, callback: BeforeAfterCallback):
        self._after_callbacks.append(callback)
        return self

    def update(self, zip_source: Path, zip_dest: Path):
        with ZipFile(zip_dest, 'w') as zout:
            self._do_before(zout)

            with ZipFile(zip_source, 'r') as zin:
                zout.comment = zin.comment  # preserve the comment
                self._do_items(zin, zout)

            self._do_after(zout)

    def _do_before(self, zout: ZipFile):
        for before_callback in self._before_callbacks:
            if not before_callback.call(zout):
                break

    def _do_items(self, zin: ZipFile, zout: ZipFile):
        for item in zin.infolist():
            self._do_item(zin, zout, item)

    def _do_item(self, zin: ZipFile, zout: ZipFile, item):
        for item_callback in self._item_callbacks:
            if not item_callback.call(zin, zout, item):
                break

    def _do_after(self, zout):
        for after_callback in self._after_callbacks:
            if not after_callback.call(zout):
                break
