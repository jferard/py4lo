#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
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
import datetime as dt
import logging
from typing import Optional, Tuple, Union, Collection

from py4lo_helper import (DatesHelper, create_uno_struct, provider)
from py4lo_typing import UnoRange, UnoCell, UnoStruct, UnoService, UnoSheet

try:
    class DataPilotFieldOrientation:
        from com.sun.star.sheet.DataPilotFieldOrientation import (
            HIDDEN, ROW, COLUMN, DATA, PAGE)


    class DataPilotFieldGroupBy:
        from com.sun.star.sheet.DataPilotFieldGroupBy import (
            SECONDS, MINUTES, HOURS, DAYS, MONTHS, QUARTERS, YEARS)


    class DataPilotFieldSortMode:
        from com.sun.star.sheet.DataPilotFieldSortMode import (
            NONE, MANUAL, NAME, DATA
        )


    class GeneralFunction2:
        from com.sun.star.sheet.GeneralFunction2 import (
            NONE, AUTO, SUM, COUNT, AVERAGE, MAX, MIN, PRODUCT, COUNTNUMS,
            STDEV, STDEVP, VAR, VARP, MEDIAN
        )


    class DataPilotFieldLayoutMode:
        from com.sun.star.sheet.DataPilotFieldLayoutMode import (
            TABULAR_LAYOUT, OUTLINE_SUBTOTALS_TOP, OUTLINE_SUBTOTALS_BOTTOM,
            COMPACT_LAYOUT
        )


    class DataPilotFieldShowItemsMode:
        from com.sun.star.sheet.DataPilotFieldShowItemsMode import (FROM_TOP,
                                                                    FROM_BOTTOM)


except ImportError:
    class DataPilotFieldOrientation:
        pass


    class DataPilotFieldGroupBy:
        pass


    class DataPilotFieldSortMode:
        pass


    class DataPilotFieldLayoutMode:
        pass


class DateGroup:
    def __init__(
            self, name: str, group_by: DataPilotFieldGroupBy,
            start: Optional[dt.date] = None,
            end: Optional[dt.date] = None, day_step: int = 0):
        self.name = name
        self.group_by = group_by
        self.start = start
        self.end = end
        self.day_step = day_step

    def __repr__(self) -> str:
        return "{}({})".format(self.__class__.__name__, self.__dict__)


class DateGroups:
    def __init__(self, *groups: DateGroup):
        self.groups = groups

    def __repr__(self) -> str:
        return "{}({})".format(self.__class__.__name__, self.__dict__)


class NameGroup:
    def __init__(
            self, name: str, values: Collection[str]):
        self.name = name
        self.values = values

    def __repr__(self) -> str:
        return "{}({})".format(self.__class__.__name__, self.__dict__)


class NameGroups:
    def __init__(self, name: str, *groups: NameGroup):
        self.name = name
        self.groups = groups

    def __repr__(self) -> str:
        return "{}({})".format(self.__class__.__name__, self.__dict__)


def make_data_sort_info(
        field: str, is_ascending: bool = True) -> UnoStruct:
    return create_uno_struct("com.sun.star.sheet.DataPilotFieldSortInfo",
                             Field=field, IsAscending=is_ascending,
                             Mode=DataPilotFieldSortMode.DATA)

def make_regular_sort_info(
        is_ascending: bool = True,
        mode: DataPilotFieldSortMode = DataPilotFieldSortMode.NAME) -> UnoStruct:
    if mode == DataPilotFieldSortMode.DATA:
        raise ValueError()
    return create_uno_struct("com.sun.star.sheet.DataPilotFieldSortInfo",
                             IsAscending=is_ascending, Mode=mode)


def make_layout_info(
        mode: DataPilotFieldLayoutMode, add_empty_lines: bool = False):
    return create_uno_struct("com.sun.star.sheet.DataPilotFieldLayoutInfo",
                             LayoutMode=mode, AddEmptyLines=add_empty_lines)


def make_auto_show_info(
        item_count: int, show_items_mode: DataPilotFieldShowItemsMode,
        field_name: str = None, is_enabled: bool = True):
    struct = create_uno_struct(
        "com.sun.star.sheet.DataPilotFieldAutoShowInfo",
        IsEnabled=is_enabled, ShowItemsMode=show_items_mode,
        ItemCount=item_count,
    )
    if field_name:
        struct.DataField = field_name
    return struct


class DataPilotBuilder:
    _logger = logging.getLogger(__name__)

    def __init__(self, oSourceRange: UnoRange, oDestCell: UnoCell, name: str,
                 dates_helper: DatesHelper):
        self._oDestCell = oDestCell
        self._name = name
        self._dates_helper = dates_helper
        oSourceSheet = oSourceRange.Spreadsheet
        self._oTables = oSourceSheet.DataPilotTables
        self._oTableDescriptor = self._oTables.createDataPilotDescriptor()
        oRangeAddress = oSourceRange.RangeAddress
        self._oTableDescriptor.SourceRange = oRangeAddress
        self._oFields = self._oTableDescriptor.getDataPilotFields()
        self._index_by_name = {
            str(value): i
            for i, value in enumerate(oSourceRange.DataArray[0])
        }
        self._oChartDoc = None

    def build(self):
        self._oTables.insertNewByName(self._name, self._oDestCell.CellAddress,
                                      self._oTableDescriptor)

    def get_descriptor(self) -> UnoService:
        return self._oTableDescriptor

    def get_chart_doc(self) -> UnoService:
        return self._oChartDoc

    def add_row(
            self, field_name: str, groups: Union[NameGroups, DateGroups] = None,
            sort: UnoStruct = None,
            subtotals: Tuple[GeneralFunction2, ...] = None,
            layout: UnoStruct = None,
            auto_show=None,
            show_empty: bool = False,
    ):
        orientation = DataPilotFieldOrientation.ROW
        self._add_row_or_column(
            field_name, orientation, groups, sort, subtotals, layout, auto_show,
            show_empty)

    def add_column(
            self, field_name: str, groups: Union[NameGroups, DateGroups] = None,
            sort: UnoStruct = None,
            subtotals: Tuple[GeneralFunction2, ...] = None,
            layout: UnoStruct = None,
            auto_show=None,
            show_empty: bool = False,
    ):
        orientation = DataPilotFieldOrientation.COLUMN
        self._add_row_or_column(
            field_name, orientation, groups, sort, subtotals, layout, auto_show,
            show_empty)

    def _add_row_or_column(
            self, field_name: str, orientation: DataPilotFieldOrientation,
            groups: Union[NameGroups, DateGroups, None],
            sort: Optional[UnoStruct],
            subtotals: Optional[Tuple[GeneralFunction2, ...]],
            layout: UnoStruct, auto_show,
            show_empty: bool,
    ):
        index = self._index_by_name[field_name]
        oField = self._oFields.getByIndex(index)
        FieldHelper(field_name, oField).make_row_or_column(orientation, groups, sort,
                                               subtotals, layout, auto_show,
                                               show_empty)

    def add_page(self, name: str):
        index = self._index_by_name[name]
        self._oCurField = self._oFields.getByIndex(index)
        self._oCurField.Orientation = DataPilotFieldOrientation.PAGE

    def add_data(self, name: str, general_function: GeneralFunction2):
        index = self._index_by_name[name]
        self._oCurField = self._oFields.getByIndex(index)
        self._oCurField.Orientation = DataPilotFieldOrientation.DATA
        self._oCurField.Function2 = general_function

    def add_chart(self, oSheet: UnoSheet, name, x, y, width, height, diagram: str):
        rect = create_uno_struct(
            "com.sun.star.awt.Rectangle", X=x, Y=y, Width=width, Height=height)

        oPivotCharts = oSheet.PivotCharts
        oPivotCharts.addNewByName(name, rect, self._name)
        oPivotChart = oPivotCharts.getByName(name)
        oPivotChartDoc = oPivotChart.EmbeddedObject
        oPivotChartDoc.Diagram = oPivotChartDoc.createInstance(diagram)

    def set_table_options(self, **kwargs):
        for key, value in kwargs.items():
            setattr(self._oTableDescriptor, key, value)

    def set_data_orientation(self, orientation: DataPilotFieldOrientation):
        oExtraField = self._oFields.getByIndex(self._oFields.Count - 1)
        oExtraField.Orientation = orientation


class FieldHelper:
    def __init__(self, field_name: str, oField: UnoService):
        self._field_name = field_name
        self._oField = oField

    def make_row_or_column(
            self, orientation: DataPilotFieldOrientation,
            groups: Union[NameGroups, DateGroups, None],
            sort: Optional[UnoStruct],
            subtotals: Optional[Tuple[GeneralFunction2, ...]],
            layout: UnoStruct, auto_show: UnoStruct,
            show_empty: bool,
    ):
        if groups:
            self._add_group(orientation, groups)
        else:
            self._oField.Orientation = orientation
        if sort:
            self._oField.HasSortInfo = True
            self._oField.SortInfo = sort
        if subtotals:
            self._oField.Subtotals2 = subtotals
        self._oField.ShowEmpty = bool(show_empty)
        if layout:
            self._oField.LayoutInfo = layout
            self._oField.HasLayoutInfo = True
        if auto_show:
            if auto_show.DataField is None:
                auto_show.DataField = self._field_name
            self._oField.AutoShowInfo = auto_show

    def _add_group(
            self, orientation: DataPilotFieldOrientation,
            groups: Union[NameGroups, DateGroups]):
        if isinstance(groups, NameGroups):
            self._group_by_name(orientation, groups)
        elif isinstance(groups, DateGroups):
            self._group_by_date(orientation, groups)
        else:
            raise ValueError()

    def _group_by_date(self, orientation: DataPilotFieldOrientation,
                       groups: DateGroups):
        groups = sorted(groups.groups, key=lambda g: g.group_by)

        group_info = create_uno_struct(
            "com.sun.star.sheet.DataPilotFieldGroupInfo",
            SourceField=self._oField, HasDateValues=True)

        oCurField = self._oField
        fields = []

        for group in groups:
            group_info.GroupBy = group.group_by
            if group.start is None:
                group_info.HasAutoStart = True
            else:
                group_info.HasAutoStart = False
                group_info.Start = self._dates_helper.date_to_int(group.start)

            if group.end is None:
                group_info.HasAutoEnd = True
            else:
                group_info.HasAutoEnd = False
                group_info.End = self._dates_helper.date_to_int(group.end)
            if group.group_by == DataPilotFieldGroupBy.DAYS and group.day_step > 0:
                group_info.Step = group.day_step
            else:
                group_info.Step = 0

            oNextField = oCurField.createDateGroup(group_info)
            if oNextField is not None:
                oCurField = oNextField

            oCurField.Name = group.name
            fields.insert(0, oCurField)

        for oCurField in fields:
            oCurField.Orientation = orientation

    def _group_by_name(self, orientation: DataPilotFieldOrientation,
                       groups: NameGroups):
        self._oField.Orientation = orientation

        first_group = groups.groups[0]
        oGroupedField = self._oField.createNameGroup(first_group.values)
        oGroupedField.Name = groups.name

        for group in groups.groups[1:]:
            oGroupedField.GroupInfo  # side-effect: avoids LO crash
            self._oField.createNameGroup(group.values)


class Diagram:
    Area = "com.sun.star.chart.AreaDiagram"
    Bar = "com.sun.star.chart.BarDiagram"
    Bubble = "com.sun.star.chart.BubbleDiagram"
    Donut = "com.sun.star.chart.DonutDiagram"
    FilledNet = "com.sun.star.chart.FilledNetDiagram"
    Line = "com.sun.star.chart.LineDiagram"
    Net = "com.sun.star.chart.NetDiagram"
    Pie = "com.sun.star.chart.PieDiagram"
    Stock = "com.sun.star.chart.StockDiagram"
    XY = "com.sun.star.chart.XYDiagram"


def data_pilot(*_args):
    oDoc = provider.doc
    oSheet = oDoc.Sheets.getByIndex(0)
    oRange = oSheet.getCellRangeByName("A1:D1001")
    oDestCell = oSheet.getCellByPosition(5, 5)
    dates_helper = DatesHelper.create(oDoc)
    builder = DataPilotBuilder(oRange, oDestCell, "test", dates_helper)
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
    builder.add_chart(oSheet, "First chart", 13000, 10000, 10000, 10000, Diagram.Area)
