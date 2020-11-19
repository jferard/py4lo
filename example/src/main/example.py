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
# py4lo: entry
# py4lo: embed script alib.py
# py4lo: embed lib py4lo_helper
# py4lo: embed lib py4lo_commons
# py4lo: embed lib py4lo_io
from datetime import date, datetime

import uno
import sys
import unohelper
import os

import py4lo_helper
import py4lo_commons
import py4lo_io
import example_lib

try:
    _ = py4lo_helper.Py4LO_helper.create(XSCRIPTCONTEXT)
    c = py4lo_commons.Commons(XSCRIPTCONTEXT)
    o = example_lib.O(_)
except NameError:
    pass


def message_example(*_args):
    """
    A doc test
    >>> 1 == 1
    True

    """
    from com.sun.star.awt.MessageBoxType import MESSAGEBOX
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    _.message_box(_.parent_win,
                  "A message from main script example.py. Current dir is: " + os.path.abspath(
                      "."), "py4lo", MESSAGEBOX, BUTTONS_OK)


def xray_example(*_args):
    _.xray(_.doc)


def mri_example(*_args):
    _.mri(_.doc)


def example_from_lib(*_args):
    o.lib_example()


def reader_example(*_args):
    r = py4lo_io.dict_reader(_.doc.getCurrentController().getActiveSheet(),
                             restval="x", restkey="t",
                             type_cell=py4lo_io.TYPE_ALL)
    for row in r:
        _.xray(r.line_num)
        _.xray(str(row))


def writer_example(*_args):
    w = py4lo_io.dict_writer(_.doc.getCurrentController().getActiveSheet(),
                             ("a", "b", "c", "d", "e"),
                             type_cell=py4lo_io.TYPE_ALL)
    w.writeheader()
    # second row after header raises an exception
    for row in [{"a": "value", "b": 1, "c": True,
                 "d": datetime(2020, 11, 21, 12, 36, 50)},
                {"a": "other value", "b": 2, "c": False,
                 "f": datetime(2020, 11, 21, 12, 36, 50)}]:
        w.writerow(row)


def export_example(*_args):
    py4lo_io.export_to_csv(_.doc.getCurrentController().getActiveSheet(),
                           "./temp.csv")


def import_example(*_args):
    py4lo_io.import_from_csv(_.desktop, _.doc, "csv sheet", 0, "./temp.csv")
