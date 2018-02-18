# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

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
import os
import zipfile

class ZipUpdater():
    def __init__(self):
        self.__before_callbacks = []
        self.__item_callbacks = []
        self.__after_callbacks = []

    def before(self, callback):
        self.__before_callbacks.append(callback)
        return self

    def item(self, callback):
        self.__item_callbacks.append(callback)
        return self

    def after(self, callback):
        self.__after_callbacks.append(callback)
        return self

    def update(self, zip_source_name, zip_dest_name):
        with zipfile.ZipFile(zip_dest_name, 'w') as zout:
            self.__do_before(zout)

            with zipfile.ZipFile(zip_source_name, 'r') as zin:
                zout.comment = zin.comment # preserve the comment
                self.__do_items(zin, zout)

            self.__do_after(zout)

    def __do_before(self, zout):
        for before_callback in self.__before_callbacks:
            if not before_callback.call(zout):
                break

    def __do_items(self, zin, zout):
        for item in zin.infolist():
            self.__do_item(zin, zout, item)

    def __do_item(self, zin, zout, item):
        for item_callback in self.__item_callbacks:
            if not item_callback.call(zin, zout, item):
                break

    def __do_after(self, zout):
        for after_callback in self.__after_callbacks:
            if not after_callback.call(zout):
                break
