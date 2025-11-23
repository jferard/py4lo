#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2025 J. Férard <https://github.com/jferard>
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
# py4lo: entry
# py4lo: embed lib py4lo_typing
# py4lo: embed lib py4lo_helper
# py4lo: embed lib py4lo_commons
# py4lo: embed lib py4lo_io
# py4lo: embed lib py4lo_dialogs
import csv
from codecs import (BOM_UTF32_BE, BOM_UTF32_LE, BOM_UTF16_BE, BOM_UTF16_LE,
                    BOM_UTF8, getincrementaldecoder)

from py4lo_dialogs import FileFilter, file_dialog
from py4lo_helper import (make_pvs, provider, Target, FrameSearchFlag,
                          make_sort_field, sort_range, SheetFormatter)
from py4lo_commons import uno_url_to_path
from py4lo_io import create_import_filter_options, Filter


def open_csv(*_args):
    url = file_dialog("Choose a CSV file", [FileFilter("csv", "*.csv;*.txt")])
    if url is None:
        return

    path = uno_url_to_path(url)
    with path.open("rb") as s:
        data = s.read(4096)

    encoding = guess_encoding(data)
    text = getincrementaldecoder(encoding)().decode(data, final=False)
    sniffer = csv.Sniffer()
    dialect = sniffer.sniff(text)
    has_header = sniffer.has_header(text)
    filter_options = create_import_filter_options(dialect)
    pvs = make_pvs({"FilterName": Filter.CSV, "FilterOptions": filter_options})

    oSource = provider.desktop.loadComponentFromURL(url, Target.BLANK,
                                                    FrameSearchFlag.AUTO, pvs)

    oCSVSheet = oSource.CurrentController.ActiveSheet

    # just for testing purpose
    sort_range(oCSVSheet, (make_sort_field(0),), has_header)

    formatter = SheetFormatter(oCSVSheet)
    if has_header:
        formatter.first_row_as_header()
        formatter.fix_first_row()
        formatter.create_filter()
    else:
        formatter.set_print_area()

    formatter.set_format("# ##0,00", 0, 1, 2, 3)
    formatter.set_optimal_width(0, 1, 2, 3, min_width=4 * 100)


def guess_encoding(data: bytes, default: str = "iso8859_15") -> str:
    """
    A basic encoding detector.
    """
    for bom, encoding in [
        (BOM_UTF32_BE, "utf_32_be"),
        (BOM_UTF32_LE, "utf_32_le"),
        (BOM_UTF16_BE, "utf_16_be"),
        (BOM_UTF16_LE, "utf_16_le"),
        (BOM_UTF8, "utf_8_sig"),
    ]:
        if data.startswith(bom):
            return encoding

    decoder = getincrementaldecoder("utf_8")()
    try:
        decoder.decode(data, final=False)
    except UnicodeDecodeError:
        return default
    else:
        return "utf_8"
