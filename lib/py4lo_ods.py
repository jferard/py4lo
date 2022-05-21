# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
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
import itertools
import xml.etree.ElementTree as ET
import zipfile

from typing import (List, cast, Dict, Callable, IO, Iterator,
                    Optional, Union)

ACTIVE_TABLE_XPATH = "./office:settings/config:config-item-set[@config:name='ooo:view-settings']/config:config-item-map-indexed[@config:name='Views']/config:config-item-map-entry/config:config-item[@config:name='ActiveTable']"

OFFICE_NS_DICT = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
}


class NameSpace:
    def __init__(self, ns_dict: Dict[str, str] = OFFICE_NS_DICT):
        self._ns_dict = ns_dict

    def findall(self, element: ET.Element, path: str) -> List[ET.Element]:
        return element.findall(path, self._ns_dict)

    def attrib(self, attr: str) -> str:
        try:
            i = attr.index(":")
            attr = "{{{}}}{}".format(self._ns_dict.get(attr[:i], attr[:i]),
                                     attr[i + 1:])
        except ValueError:
            pass

        return attr


OFFICE_NS = NameSpace(OFFICE_NS_DICT)

SortFunc = Callable[[List[ET.Element]], List[ET.Element]]


class OdsTables:
    """An iterator over tables in an ods file"""

    @staticmethod
    def create(fullpath: str, ns: NameSpace = OFFICE_NS):
        return OdsTablesBuilder(fullpath).ns(ns).build()

    def __init__(self, root: ET.Element, sort_func: SortFunc,
                 ns: NameSpace = OFFICE_NS):
        self._root = root
        self._sort_func = sort_func
        self._ns = ns

    def __iter__(self) -> Iterator[ET.Element]:
        tables = self._ns.findall(self._root,
                                  "./office:body/office:spreadsheet/table:table")
        tables = self._sort_func(tables)
        for table in tables:
            yield table


class OdsTablesBuilder:
    def __init__(self, fullpath: str):
        self._fullpath = fullpath
        self._ns = OFFICE_NS
        self._sort_func_creator = dont_sort

    def ns(self, ns: NameSpace) -> "OdsTablesBuilder":
        self._ns = ns
        return self

    def sort_func_creator(self, sort_func_creator: Callable[
        [zipfile.ZipFile, NameSpace], SortFunc]):
        self._sort_func_creator = sort_func_creator

    def build(self) -> "OdsTables":
        with zipfile.ZipFile(self._fullpath) as z:
            sort_func = self._sort_func_creator(z, self._ns)
            with z.open('content.xml') as content:
                data = content.read().decode("utf-8")
                root = ET.fromstring(data)
                return OdsTables(root, sort_func, self._ns)


def dont_sort(*_args, **_kwargs) -> SortFunc:
    def sort_func(tables):
        return tables

    return sort_func


def put_active_first(z: zipfile.ZipFile, ns: NameSpace = OFFICE_NS) -> SortFunc:
    """
    :param z: the zip file
    :param ns: the name space
    :returns: a sort function for tables
    """
    active_table_name = _find_active_table_name_in_zip(z, ns)

    def sort_func(tables: List[ET.Element]) -> List[ET.Element]:
        for i, table in enumerate(tables):
            if table.get(ns.attrib("table:name")) == active_table_name:
                return [tables[i]] + tables[:i] + tables[i + 1:]

        return tables

    return sort_func


def _find_active_table_name_in_zip(z: zipfile.ZipFile,
                                   ns: NameSpace = OFFICE_NS) -> str:
    with z.open('settings.xml') as settings:
        active_table_name = _find_active_table_name(settings, ns)

    return active_table_name


def _find_active_table_name(settings: IO[bytes], ns=OFFICE_NS) -> str:
    data = settings.read().decode("utf-8")
    root = ET.fromstring(data)
    active_tables = ns.findall(root, ACTIVE_TABLE_XPATH)
    if active_tables:
        active_table_name = active_tables[0].text
    else:
        active_table_name = None
    return active_table_name


class OdsRows:
    """
    An iterator over rows in a table.
    The table must be simple (no spanned rows or repeated rows)

    * A cell spanned on multiple columns is duplicated on these columns
    * A cell spanned on multple rows is written on the first row, and empty
    below.
    """

    def __init__(self, table: ET.Element,
                 omit: Optional[Callable[[ET.Element, NameSpace], bool]] = None,
                 ns: NameSpace = OFFICE_NS):
        self._table = table
        self._ns = ns
        self._omit = omit

    def __iter__(self) -> Iterator[List[str]]:
        for row in self._ns.findall(self._table, ".//table:table-row"):
            if self._omit is not None and self._omit(row, self._ns):
                continue

            count = int(row.get(self._ns.attrib("table:number-rows-repeated"),
                                row.attrib.get(self._ns.attrib(
                                    "table:number-rows-spanned"), "1")))
            cells = self._get_cells(row)
            for _ in range(count):
                yield cells

    def _get_cells(self, row: ET.Element) -> List[str]:
        cell_tag = self._ns.attrib("table:table-cell")
        covered_cell_tag = self._ns.attrib("table:covered-table-cell")
        spanned_attr = self._ns.attrib("table:number-columns-spanned")
        repeated_attr = self._ns.attrib("table:number-columns-repeated")
        cells = cast(List[str], [])
        span_count = 0
        for e in row:
            if e.tag not in (cell_tag, covered_cell_tag):
                continue

            v1 = '\n'.join(
                p.text for p in self._ns.findall(e, "./text:p") if p.text)
            repeat_count = int(e.attrib.get(repeated_attr, "1"))
            if e.tag == cell_tag:
                span_count = int(e.attrib.get(spanned_attr, "1"))
                count = max(repeat_count, span_count)
                cells.extend([v1] * count)
            elif e.tag == covered_cell_tag:
                # expect span_count-1 covered cells
                # other cells are empty cells (row-span).
                if cells:
                    if repeat_count + 1 > span_count:
                        cells.extend([''] * (repeat_count + 1 - span_count))
                    else:  # span_count >= repeat_count + 1
                        span_count -= repeat_count  # span_count >= 1
                else:  # beginning of line
                    cells.extend([''] * repeat_count)

        cells = self._trim_list(cells)
        return cells

    def __getitem__(self, index: Union[int, slice]) -> Union[
        List[str], List[List[str]]]:
        if isinstance(index, int):
            if index >= 0:
                for i, row in enumerate(self):
                    if i == index:
                        return row
                raise IndexError("{} out of table".format(index))
            else:
                all_rows = list(self)
                return all_rows[index]
        elif isinstance(index, slice):
            # use itertools.islice to avoid a list creation
            return list(
                itertools.islice(self, index.start, index.stop, index.step))
        else:
            raise TypeError("index must be int or slice")

    def _trim_list(self, l: List[str]) -> List[str]:
        if not l:
            return l
        for i in range(len(l) - 1, -1, -1):
            if len(l[i]):
                return l[:i + 1]
        return []


def omit_filtered(row: ET.Element, ns: NameSpace) -> bool:
    return row.attrib.get(ns.attrib("table:visibility"), "") == "filter"
