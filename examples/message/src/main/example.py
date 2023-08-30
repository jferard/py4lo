# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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
import os
# py4lo: entry
# py4lo: embed script alib.py
# py4lo: embed lib py4lo_typing
# py4lo: embed lib py4lo_helper
# py4lo: embed lib py4lo_commons
# py4lo: embed lib py4lo_io
# py4lo: embed lib py4lo_dialogs
import time
from datetime import datetime

import example_lib
from py4lo_dialogs import (ProgressExecutorBuilder, ConsoleExecutorBuilder,
                           message_box, input_box)
from py4lo_helper import provider as pr, xray, mri, parent_doc, convert_to_html
from py4lo_io import (dict_reader, dict_writer, export_to_csv,
                      import_from_csv, CellTyping)

o = example_lib.Example(pr)


def message_example(*_args):
    """
    A doc test
    >>> 1 == 1
    True

    """
    oSheet = pr.controller.getActiveSheet()
    oDoc = parent_doc(oSheet)  # could be pr.doc
    lines = [
        "A message from main script example.py. ",
        "Current dir is: {}".format(os.path.abspath("../../../../py4lo")),
        "Current doc name is: {}".format(oDoc.Title),
        "Current sheet name is: {}".format(oSheet.Name),
        "Some HTML: " + convert_to_html(
            oDoc.Sheets.getByIndex(0).getCellByPosition(0, 0))
    ]

    message_box("py4lo", "\n".join(lines))
    name = input_box("py4lo", "enter your name")
    message_box("py4lo", "Hello {}".format(name))


def xray_example(*_args):
    xray(pr.doc)


def mri_example(*_args):
    mri(pr.doc)


def example_from_lib(*_args):
    o.lib_example()


def writer_example(*_args):
    w = dict_writer(pr.controller.getActiveSheet(),
                    ["a", "b", "text_range", "d", "e"],
                    cell_typing=CellTyping.Accurate)
    w.writeheader()
    for row in [{"a": "value", "b": 1, "text_range": True,
                 "d": datetime(2020, 11, 21, 12, 36, 50)},
                {"a": "other value", "b": 2, "text_range": False,
                 "d": datetime(2020, 11, 21, 12, 36, 50)}]:
        w.writerow(row)


def reader_example(*_args):
    r = dict_reader(pr.controller.getActiveSheet(),
                    restval="x", restkey="t",
                    cell_typing=CellTyping.Accurate)
    for row in r:
        message_box("py4lo", "{}: {}".format(r.line_num, row))


def export_example(*_args):
    export_to_csv(pr.controller.getActiveSheet(),
                  "./temp.csv")


def import_example(*_args):
    import_from_csv(pr.doc, "csv sheet", 0, "./temp.csv")


progress_executor = None


def progress_example(*_args):
    global progress_executor
    if progress_executor is None:
        progress_executor = ProgressExecutorBuilder().build()

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
    message_box("Title",
                "The return value was: {}".format(progress_executor.response))


console_executor = None


def console_example(*_args):
    global console_executor
    if console_executor is None:
        console_executor = ConsoleExecutorBuilder().build()

    def test(console_handler):
        """The test function"""
        console_handler.message("a message")
        for i in range(1, 15):
            time.sleep(0.1)
            console_handler.message("next message: {}".format(i))

    console_executor.execute(test)
