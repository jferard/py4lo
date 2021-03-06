# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2021 J. Férard <https://github.com/jferard>

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
# py4lo: entry
# py4lo: embed script alib.py
# py4lo: embed lib py4lo_helper
# py4lo: embed lib py4lo_commons
# py4lo: embed lib py4lo_io
# py4lo: embed lib py4lo_dialogs
import time
from datetime import datetime

import example_lib
from py4lo_dialogs import ProgressExecutorBuilder, ConsoleExecutorBuilder
from py4lo_helper import provider as pr, xray, mri, message_box
from py4lo_io import (dict_reader, TYPE_ALL, dict_writer, export_to_csv,
                      import_from_csv)

o = example_lib.O(pr)


def message_example(*_args):
    """
    A doc test
    >>> 1 == 1
    True

    """
    from com.sun.star.awt.MessageBoxType import MESSAGEBOX
    from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
    message_box("A message from main script example.py. "
                "Current dir is: " + os.path.abspath(
                    "."), "py4lo", MESSAGEBOX, BUTTONS_OK)


def xray_example(*_args):
    xray(pr.doc)


def mri_example(*_args):
    mri(pr.doc)


def example_from_lib(*_args):
    o.lib_example()


def reader_example(*_args):
    r = dict_reader(pr.controller.getActiveSheet(),
                    restval="x", restkey="t",
                    type_cell=TYPE_ALL)
    for row in r:
        xray(r.line_num)
        xray(str(row))


def writer_example(*_args):
    w = dict_writer(pr.controller.getActiveSheet(),
                    ("a", "b", "c", "d", "e"),
                    type_cell=TYPE_ALL)
    w.writeheader()
    # second row after header raises an exception
    for row in [{"a": "value", "b": 1, "c": True,
                 "d": datetime(2020, 11, 21, 12, 36, 50)},
                {"a": "other value", "b": 2, "c": False,
                 "f": datetime(2020, 11, 21, 12, 36, 50)}]:
        w.writerow(row)


def export_example(*_args):
    export_to_csv(pr.controller.getActiveSheet(),
                  "./temp.csv")


def import_example(*_args):
    import_from_csv(pr.doc, "csv sheet", 0, "./temp.csv")


progress_executor = ProgressExecutorBuilder().build()

def progress_example(*_args):
    def test(progress_handler):
        """The test function"""
        progress_handler.message("a message")
        for i in range(1, 101):
            time.sleep(0.1)
            progress_handler.progress(1)
            if i > 50:
                progress_handler.message("another message")
        progress_handler.response = 5

    progress_executor.execute(test)


def after_progress(*_args):
    message_box("The return value was: {}".format(progress_executor.response), "Title")


console_executor = ConsoleExecutorBuilder().build()


def console_example(*_args):
    def test(console_handler):
        """The test function"""
        console_handler.message("a message")
        for i in range(1, 15):
            time.sleep(0.1)
            console_handler.message("next message: {}".format(i))

    console_executor.execute(test)
