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
from abc import ABC, abstractmethod
from zipfile import ZipFile, ZipInfo


class BeforeCallback(ABC):
    """
    Called before the item processing
    """

    @abstractmethod
    def call(self, zout: ZipFile) -> bool:
        """
        :param zout: write to this file
        :return: False to prevent the execution of nextcallbacks
        """
        pass


class AfterCallback(ABC):
    """
    Called after the item processing. Use an AfterCallback to add some files
    to the archive
    """

    @abstractmethod
    def call(self, zout: ZipFile) -> bool:
        """
        :param zout: write to this file
        :return: False to prevent the execution of nextcallbacks
        """
        pass


class ItemCallback(ABC):
    """
    Called on each item of the source. Use a ItemCallback to ignore a script.
    """
    @abstractmethod
    def call(self, zin: ZipFile, zout: ZipFile, item: ZipInfo) -> bool:
        """
        :param zin: read from this file
        :param zout: write to this file
        :param item: the item to process
        :return: True if the item was processed and copied. Else, return false.
        It's the responsibility of the updater to copy the file.
        """
        pass
