# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2020 J. Férard <https://github.com/jferard>

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
from logging import Logger
from pathlib import Path
from typing import List
from zipfile import ZipFile, ZipInfo

from callbacks.callback import AfterCallback, BeforeCallback, ItemCallback


class ZipUpdaterBuilder:
    """
    A zip file updater. Applies callbacks before, after and to each item.
    """

    def __init__(self, logger: Logger):
        self._logger = logger
        self._before_callbacks = []
        self._item_callbacks = []
        self._after_callbacks = []

    def before(self, callback: BeforeCallback):
        self._before_callbacks.append(callback)
        return self

    def item(self, callback: ItemCallback):
        self._item_callbacks.append(callback)
        return self

    def after(self, callback: AfterCallback):
        self._after_callbacks.append(callback)
        return self

    def build(self):
        return ZipUpdater(self._logger, self._before_callbacks, self._item_callbacks,
                          self._after_callbacks)


class ZipUpdater:
    """
    A zip file updater. Applies callbacks before, after and to each item.
    """

    def __init__(self, logger: Logger, before_callbacks: List[BeforeCallback],
                 item_callbacks: List[ItemCallback], after_callbacks: List[
                AfterCallback]):
        self._logger = logger
        self._before_callbacks = before_callbacks
        self._item_callbacks = item_callbacks
        self._after_callbacks = after_callbacks

    def update(self, zip_source: Path, zip_dest: Path):
        """
        Update the given zip

        :param zip_source: a source zip archive
        :param zip_dest: a dest path
        """
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
            if not self._do_item(zin, zout, item):
                self._logger.debug("Copy %s to archive", item.filename)
                bs = zin.read(item.filename)
                zout.writestr(item.filename, bs)  # copy

    def _do_item(self, zin: ZipFile, zout: ZipFile, item: ZipInfo) -> bool:
        touched = False
        # don't use `any` because... don't use à LC
        for item_callback in self._item_callbacks:
            if item_callback.call(zin, zout, item):
                touched = True

        return touched

    def _do_after(self, zout):
        for after_callback in self._after_callbacks:
            if not after_callback.call(zout):
                break
