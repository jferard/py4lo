# -*- coding: utf-8 -*-
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
"""
A weird module: read ODS file out of LibreOffice. The module py4lo_ods
provides functions to read the content.xml file.

It is useful to read ODS files having a simple structure really fast, but
needs some knowledge of the ODS file format.
"""
import itertools
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

from typing import (List, cast, Dict, Callable, IO, Iterator,
                    Optional, Union)

ACTIVE_TABLE_XPATH = (
    "./office:settings"
    "/config:config-item-set[@config:name='ooo:view-settings']"
    "/config:config-item-map-indexed[@config:name='Views']"
    "/config:config-item-map-entry"
    "/config:config-item[@config:name='ActiveTable']")
"""The XPath to the active table"""

OFFICE_NS_DICT = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
}
"""The LibreOffice namespace dictionary"""


class NameSpace:
    """
    An XML namespace. Wraps `xml.etree.ElementTree` functions.
    Example:

    ns = NameSpace(OFFICE_NS_DICT)*
    with ZipFile('file.ods', 'r') as z:
        xml_str = z.read('content.xml').decode("utf-8")
        ET.fromstring(xml_str)
        body = ns.find(root, "./office:document-content/office:body")
    """
    def __init__(self, ns_dict: Optional[Dict[str, str]] = None):
        """

        @param ns_dict:
        """
        if ns_dict is None:
            self._ns_dict = OFFICE_NS_DICT
        else:
            self._ns_dict = ns_dict

    def findall(self, element: ET.Element, path: str) -> List[ET.Element]:
        """
        See. `xml.etree.ElementTree.findall`
        @param element: an element
        @param path: the XPath
        @return: a list of elements
        """
        return element.findall(path, self._ns_dict)

    def attrib(self, attr: str) -> str:
        """
        Convert an attrib to the {namespace}tag notation.

        @param attr: the attribute
        @return: x.
        """
        try:
            i = attr.index(":")
            ns = attr[:i]
            tag = attr[i + 1:]
            attr = "{{{}}}{}".format(self._ns_dict.get(ns, ns), tag)
        except ValueError:
            pass

        return attr


OFFICE_NS = NameSpace(OFFICE_NS_DICT)
"""The LibreOffice Namespace"""

SortFunc = Callable[[List[ET.Element]], List[ET.Element]]


class OdsTables:
    """
    An iterator over tables in an ods file
    """
    @staticmethod
    def create(fullpath: Union[str, Path], ns: NameSpace = OFFICE_NS) -> "OdsTables":
        """
        Create a new OdsTables object.

        @param fullpath: the file path
        @param ns: the XML namespaces of the document
        @return: the OdsTables object.
        """
        return OdsTablesBuilder(fullpath).ns(ns).build()

    def __init__(self, root: ET.Element, sort_func: SortFunc,
                 ns: NameSpace = OFFICE_NS):
        """
        @param root: the root element (`office:document-content` tag)
        @param sort_func: the sort function to sort the tables
        @param ns: the XML namespaces of the document
        """
        self._root = root
        self._sort_func = sort_func
        self._ns = ns

    def __iter__(self) -> Iterator[ET.Element]:
        tables = self._ns.findall(
            self._root, "./office:body/office:spreadsheet/table:table")
        tables = self._sort_func(tables)
        for table in tables:
            yield table


class OdsTablesBuilder:
    """
    A builder for OdsTables object.
    """
    def __init__(self, fullpath: Union[str, Path]):
        """
        @param fullpath: the path of the file
        """
        if isinstance(fullpath, Path):
            self._fullpath = str(fullpath)  # py < 3.6.2
        else:
            self._fullpath = fullpath
        self._ns = OFFICE_NS
        self._sort_func_creator = dont_sort

    def ns(self, ns: NameSpace) -> "OdsTablesBuilder":
        """
        Set the namespaces
        @param ns: the namespaces
        @return: self
        """
        self._ns = ns
        return self

    def sort_func_creator(self, sort_func_creator: Callable[
        [zipfile.ZipFile, NameSpace], SortFunc
    ]) -> "OdsTablesBuilder":
        """
        Set the creator for sort function.

        @param sort_func_creator: a sort function factory
        @return: self
        """
        self._sort_func_creator = sort_func_creator
        return self

    def build(self) -> "OdsTables":
        """
        Build the OdsTables object
        @return: the OdsTables object
        """
        with zipfile.ZipFile(self._fullpath) as z:
            sort_func = self._sort_func_creator(z, self._ns)
            with z.open('content.xml') as content:
                data = content.read().decode("utf-8")
                root = ET.fromstring(data)
                return OdsTables(root, sort_func, self._ns)


def dont_sort(*_args, **_kwargs) -> SortFunc:
    """
    A "do not sort tables" factory, for OdsTablesBuilder.sort_func_creator.

    @param _args: ignored
    @param _kwargs: ignored
    @return: a SortFunc.
    """
    def sort_func(tables):
        return tables

    return sort_func


def put_active_first(z: zipfile.ZipFile, ns: NameSpace = OFFICE_NS
                     ) -> SortFunc:
    """
    A "put the active table first" factory, for OdsTablesBuilder.sort_func_creator.

    @param z: the zip file
    @param ns: the name space
    @return: a SortFunc
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
    An iterator over rows in a table, returns the values as a list of strings.
    The table must be simple (no spanned rows or repeated rows)

    * A cell spanned on multiple columns is duplicated on these columns
    * A cell spanned on multple rows is written on the first row, and empty
    below.

    Note: __getitem__ is slow because of the spanned/repeated rows.
    """
    def __init__(
            self, table: ET.Element,
            omit: Optional[Callable[[ET.Element, NameSpace], bool]] = None,
            ns: NameSpace = OFFICE_NS):
        """
        @param table: the table ET.Element
        @param omit: a function to omit specific rows
        @param ns: the namespace.
        """
        self._table = table
        self._ns = ns
        self._omit = omit

    def __iter__(self) -> Iterator[List[str]]:
        number_rows_repeated_attrib = self._ns.attrib("table:number-rows-repeated")
        number_rows_spanned_attrib = self._ns.attrib(
            "table:number-rows-spanned")
        for row in self._ns.findall(self._table, ".//table:table-row"):
            if self._omit is not None and self._omit(row, self._ns):
                continue

            count_str = row.get(number_rows_repeated_attrib)
            if count_str is None:
                count_str = row.get(number_rows_spanned_attrib)
            count = self._to_int_or_one(count_str)

            cells = self._get_cells(row)
            for _ in range(count):
                yield cells

    @staticmethod
    def _to_int_or_one(count_str: Optional[str]) -> int:
        return 1 if count_str is None else int(count_str)

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

            repeat_count_str = e.get(repeated_attr)
            repeat_count = self._to_int_or_one(repeat_count_str)

            if e.tag == cell_tag:

                span_count_str = e.get(spanned_attr)
                span_count = self._to_int_or_one(span_count_str)

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

    def __getitem__(self, index: Union[int, slice]
                    ) -> Union[List[str], List[List[str]]]:
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

    def _trim_list(self, lis: List[str]) -> List[str]:
        if not lis:
            return lis
        for i in range(len(lis) - 1, -1, -1):
            if len(lis[i]):
                return lis[:i + 1]
        return []


def omit_filtered(row: ET.Element, ns: NameSpace) -> bool:
    """
    Omit rows having "table:visibility" attribute set to "filter"

    @param row: the row element
    @param ns: the namespace
    @return: True if "table:visibility" attribute is set to "filter", False otherwise
    """
    return row.attrib.get(ns.attrib("table:visibility"), "") == "filter"
