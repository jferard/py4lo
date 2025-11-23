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
# py4lo: embed lib py4lo_dialogs
from py4lo_helper import (
    DatesHelper, provider, DateGroup, DateGroups,
    make_regular_sort_info,
    make_layout_info, make_auto_show_info, Diagram,
    DataPilotBuilder, DataPilotFieldGroupBy,
    DataPilotFieldSortMode, GeneralFunction2,
    DataPilotFieldLayoutMode, DataPilotFieldShowItemsMode,
    DataPilotFieldOrientation)


def data_pilot(*_args):
    oDoc = provider.doc
    oSheet = oDoc.Sheets.getByIndex(0)
    oRange = oSheet.getCellRangeByName("A1:D1001")
    oDestCell = oSheet.getCellByPosition(5, 5)
    dates_helper = DatesHelper.create(oDoc)
    builder = DataPilotBuilder("test", oRange, oDestCell, dates_helper)
    builder.add_page("Person")
    builder.add_row(
        "Category",
        sort=make_regular_sort_info(
            is_ascending=False,
            mode=DataPilotFieldSortMode.NAME),
        subtotals=(GeneralFunction2.AUTO,),
        layout=make_layout_info(mode=DataPilotFieldLayoutMode.TABULAR_LAYOUT),
        auto_show=make_auto_show_info(
            2, DataPilotFieldShowItemsMode.FROM_TOP),
        show_empty=True
    )
    builder.add_row(
        "Date",
        groups=DateGroups(
            DateGroup("years", DataPilotFieldGroupBy.YEARS),
            DateGroup("quarters", DataPilotFieldGroupBy.QUARTERS),
            DateGroup("months", DataPilotFieldGroupBy.MONTHS),
        )
    )
    builder.add_data("Value", GeneralFunction2.SUM)
    builder.add_data("Value", GeneralFunction2.COUNT)
    builder.set_data_orientation(DataPilotFieldOrientation.COLUMN)
    builder.build()
    builder.add_chart(oSheet, "First chart", 13000, 10000, 10000, 10000,
                      Diagram.Area)
