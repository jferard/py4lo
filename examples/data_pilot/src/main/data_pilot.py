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

from py4lo_helper import (DatesHelper, create_uno_struct, provider, parent_doc)
from py4lo_typing import UnoRange, UnoCell, UnoStruct, UnoService, UnoSheet

try:
    class DataPilotFieldOrientation:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.DataPilotFieldOrientation import (
            HIDDEN, ROW, COLUMN, DATA, PAGE)


    class DataPilotFieldGroupBy:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.DataPilotFieldGroupBy import (
            SECONDS, MINUTES, HOURS, DAYS, MONTHS, QUARTERS, YEARS)


    class DataPilotFieldSortMode:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.DataPilotFieldSortMode import (
            NONE, MANUAL, NAME, DATA
        )


    class GeneralFunction2:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.GeneralFunction2 import (
            NONE, AUTO, SUM, COUNT, AVERAGE, MAX, MIN, PRODUCT, COUNTNUMS,
            STDEV, STDEVP, VAR, VARP, MEDIAN
        )


    class DataPilotFieldLayoutMode:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.DataPilotFieldLayoutMode import (
            TABULAR_LAYOUT, OUTLINE_SUBTOTALS_TOP, OUTLINE_SUBTOTALS_BOTTOM,
            COMPACT_LAYOUT
        )


    class DataPilotFieldShowItemsMode:
        # noinspection PyUnresolvedReferences
        from com.sun.star.sheet.DataPilotFieldShowItemsMode import (FROM_TOP,
                                                                    FROM_BOTTOM)

except ImportError:
    from _mock_constants import (
        GeneralFunction2, DataPilotFieldOrientation, DataPilotFieldGroupBy,
        DataPilotFieldSortMode, DataPilotFieldLayoutMode,
        DataPilotFieldShowItemsMode)


class DateGroup:
    """
    A date group for grouping DataPilotFields
    """

    def __init__(
            self, name: str, group_by: DataPilotFieldGroupBy,
            start: Optional[dt.date] = None,
            end: Optional[dt.date] = None, day_step: int = 0):
        """
        :param name: the name of the group
        :param group_by: the DataPilotFieldGroupBy value (seconds to years)
        :param start: the start date or None if automatic
        :param end: the end date or None if automatic
        :param day_step: the step if group_by is set to DataPilotFieldGroupBy.DAYS
        """
        self.name = name
        self.group_by = group_by
        self.start = start
        self.end = end
        self.day_step = day_step

    def __repr__(self) -> str:
        return "{}({})".format(self.__class__.__name__, self.__dict__)


class DateGroups:
    """
    A list of DateGroups
    """

    def __init__(self, *groups: DateGroup):
        """
        :param groups: the groups
        """
        self.groups = groups

    def __repr__(self) -> str:
        return "{}({})".format(self.__class__.__name__, self.__dict__)


class NameGroup:
    """
    A name group for grouping DataPilotFields
    """

    def __init__(
            self, name: str, values: Collection[str]):
        """
        :param name: the name of the group (ignored)
        :param values: the values of the group
        """
        self.name = name
        self.values = values

    def __repr__(self) -> str:
        return "{}({})".format(self.__class__.__name__, self.__dict__)


class NameGroups:
    """
    A list of name groups
    """

    def __init__(self, name: str, *groups: NameGroup):
        """
        :param name: the name of the grouped field
        :param groups: the groups
        """
        self.name = name
        self.groups = groups

    def __repr__(self) -> str:
        return "{}({})".format(self.__class__.__name__, self.__dict__)


def make_data_sort_info(
        field_name: str, is_ascending: bool = True) -> UnoStruct:
    """
    Create a com.sun.star.sheet.DataPilotFieldSortInfo structure for data
    fields.
    :param field_name: the name the field that is the sort key
    :param is_ascending: True if ascending order, False otherwise
    :return: the struct
    """
    return create_uno_struct("com.sun.star.sheet.DataPilotFieldSortInfo",
                             Field=field_name, IsAscending=is_ascending,
                             Mode=DataPilotFieldSortMode.DATA)


def make_regular_sort_info(
        is_ascending: bool = True,
        mode: DataPilotFieldSortMode = DataPilotFieldSortMode.NAME) -> UnoStruct:
    """
    Create a com.sun.star.sheet.DataPilotFieldSortInfo structure for page, row
    or colum fields.
    :param is_ascending: ascending order if True, descending otherwise
    :param mode: the DataPilotFieldSortMode.
    :return:
    """
    if mode == DataPilotFieldSortMode.DATA:
        raise ValueError()
    return create_uno_struct("com.sun.star.sheet.DataPilotFieldSortInfo",
                             IsAscending=is_ascending, Mode=mode)


def make_layout_info(
        mode: DataPilotFieldLayoutMode, add_empty_lines: bool = False):
    """
    Create a com.sun.star.sheet.DataPilotFieldLayoutInfo structure
    :param mode: the DataPilotFieldLayoutMode.
    :param add_empty_lines: add empty lines between field values if True,
    don't add empty lines otherwise
    :return: the struct
    """
    return create_uno_struct("com.sun.star.sheet.DataPilotFieldLayoutInfo",
                             LayoutMode=mode, AddEmptyLines=add_empty_lines)


def make_auto_show_info(
        item_count: int, show_items_mode: DataPilotFieldShowItemsMode,
        field_name: str = None, is_enabled: bool = True):
    """
    Create a com.sun.star.sheet.DataPilotFieldAutoShowInfo structure

    :param item_count: the number of items to show
    :param show_items_mode: from top or from bottom
    :param field_name: ignore this field
    :param is_enabled: ignore this field
    :return: the struct
    """
    struct = create_uno_struct(
        "com.sun.star.sheet.DataPilotFieldAutoShowInfo",
        IsEnabled=is_enabled, ShowItemsMode=show_items_mode,
        ItemCount=item_count,
    )
    if field_name:
        struct.DataField = field_name
    return struct


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



class DataPilotBuilder:
    """
    A data pilot table builder. A simple example :

    ```
    oRange = oSheet.getCellRangeByName("A1:D1001")
    oDestCell = oSheet.getCellByPosition(5, 5)
    builder = DataPilotBuilder.create("test", oRange, oDestCell)
    builder.add_row("Category")
    builder.add_colmun("Person")
    builder.add_data("Value", GeneralFunction2.SUM)
    builder.build()
    ```
    """
    _logger = logging.getLogger(__name__)

    @staticmethod
    def create(name: str, oSourceRange: UnoRange,
               oDestCell: UnoCell) -> "DataPilotBuilder":
        """
        Create a new builder

        :param oSourceRange: the source of data
        :param oDestCell: the destination cell
        :param name: the name of the pilot table
        """
        dates_helper = DatesHelper.create(parent_doc(oSourceRange))
        return DataPilotBuilder(name, oSourceRange, oDestCell, dates_helper)

    def __init__(self, name: str, oSourceRange: UnoRange, oDestCell: UnoCell,
                 dates_helper: DatesHelper):
        """
        :param oSourceRange: the source of data
        :param oDestCell: the destination cell
        :param name: the name of the pilot table
        :param dates_helper: a DatesHelper
        """
        self._name = name
        self._oDestCell = oDestCell
        self._dates_helper = dates_helper

        self._index_by_name = {
            str(value): i
            for i, value in enumerate(oSourceRange.DataArray[0])
        }

        oSourceSheet = oSourceRange.Spreadsheet
        self._oTables = oSourceSheet.DataPilotTables
        self._oTableDescriptor = self._oTables.createDataPilotDescriptor()

        oRangeAddress = oSourceRange.RangeAddress
        self._oTableDescriptor.SourceRange = oRangeAddress
        self._oFields = self._oTableDescriptor.getDataPilotFields()
        self._oChartDoc = None

    def build(self):
        """
        Build the table and insert it at the destination cell
        """
        self._oTables.insertNewByName(
            self._name, self._oDestCell.CellAddress, self._oTableDescriptor)

    def get_descriptor(self) -> UnoService:
        """
        Return the underlying com.sun.star.sheet.DataPilotDescriptor. This
        method may be called at any time to acces the UNO API.

        :return: the descriptor
        """
        return self._oTableDescriptor

    def get_chart_doc(self) -> Optional[UnoService]:
        """
        Return the underlying com.sun.star.chart.ChartDocument if a chart
        was added (see `add_chart`). This ethod may be called after add_chart
        to acces the UNO API.

        :return: the chart document
        """
        return self._oChartDoc

    def add_row(
            self, field_name: str, groups: Union[NameGroups, DateGroups] = None,
            sort: UnoStruct = None,
            subtotals: Tuple[GeneralFunction2, ...] = None,
            layout: UnoStruct = None,
            auto_show=None,
            show_empty: bool = False,
    ):
        """
        Add a row field.

        :param field_name: the name of the field to add (in the source)
        :param groups: a NameGroups or DateGroups object to group the field
        or None
        :param sort: com.sun.star.sheet.DataPilotFieldSortInfo structure
        (see make_regular_sort_info and make_data_sort_info) or None
        :param subtotals: a tuple of GeneralFunction2 values or None
        :param layout: com.sun.star.sheet.DataPilotFieldLayoutInfo structure
        (see make_layout_info) or None
        :param auto_show: a com.sun.star.sheet.DataPilotFieldAutoShowInfo
        structure (see make_auto_show_info) or None
        :param show_empty: show empty lines if True. Default is False
        """
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
        """
        Add a column field.

        :param field_name: the name of the field to add (in the source)
        :param groups: a NameGroups or DateGroups object to group the field
        or None
        :param sort: com.sun.star.sheet.DataPilotFieldSortInfo structure
        (see make_regular_sort_info and make_data_sort_info) or None
        :param subtotals: a tuple of GeneralFunction2 values or None
        :param layout: com.sun.star.sheet.DataPilotFieldLayoutInfo structure
        (see make_layout_info) or None
        :param auto_show: a com.sun.star.sheet.DataPilotFieldAutoShowInfo
        structure (see make_auto_show_info) or None
        :param show_empty: show empty lines if True. Default is False
        """
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
        FieldHelper(field_name, oField).make_row_or_column(
            orientation, groups, sort, subtotals, layout, auto_show, show_empty)

    def add_page(self, field_name: str):
        """
        Add a page field.

        :param field_name: the name of the field to add (in the source)
        """
        index = self._index_by_name[field_name]
        oField = self._oFields.getByIndex(index)
        oField.Orientation = DataPilotFieldOrientation.PAGE

    def add_data(self, field_name: str, function: GeneralFunction2):
        """
        Add a page field.

        :param field_name: the name of the field to add (in the source)
        :param function: the GeneralFunction2 value
        """
        index = self._index_by_name[field_name]
        oField = self._oFields.getByIndex(index)
        oField.Orientation = DataPilotFieldOrientation.DATA
        oField.Function2 = function

    def add_chart(self, oSheet: UnoSheet, name, x, y, width, height,
                  diagram: Optional[str] = None):
        """
        Add a pivot chart

        :param oSheet: the sheet where to add the chart
        :param name: the name of the chart
        :param x: the X position of the chart
        :param y: the y position of the chart
        :param width: the width of the chart
        :param height: height of the chart
        :param diagram: the name of the type the chart (see Diagram class)
l        """
        rect = create_uno_struct(
            "com.sun.star.awt.Rectangle", X=x, Y=y, Width=width, Height=height)

        oPivotCharts = oSheet.PivotCharts
        oPivotCharts.addNewByName(name, rect, self._name)
        oPivotChart = oPivotCharts.getByName(name)
        oPivotChartDoc = oPivotChart.EmbeddedObject
        if diagram is not None:
            oPivotChartDoc.Diagram = oPivotChartDoc.createInstance(diagram)

    def set_table_options(self, **kwargs):
        for key, value in kwargs.items():
            setattr(self._oTableDescriptor, key, value)

    def set_data_orientation(self, orientation: DataPilotFieldOrientation):
        oExtraField = self._oFields.getByIndex(self._oFields.Count - 1)
        oExtraField.Orientation = orientation


class FieldHelper:
    """
    A field helper. Don't use directly.
    """
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
        """
        Create a row or a colmun field

        :param orientation: the field orientation
        :param groups: a NameGroups or DateGroups object to group the field
        or None
        :param sort: com.sun.star.sheet.DataPilotFieldSortInfo structure
        (see make_regular_sort_info and make_data_sort_info) or None
        :param subtotals: a tuple of GeneralFunction2 values or None
        :param layout: com.sun.star.sheet.DataPilotFieldLayoutInfo structure
        (see make_layout_info) or None
        :param auto_show: a com.sun.star.sheet.DataPilotFieldAutoShowInfo
        structure (see make_auto_show_info) or None
        :param show_empty: show empty lines if True. Default is False
        """
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
