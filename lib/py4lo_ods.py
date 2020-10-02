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
import zipfile
import xml.etree.ElementTree as ET
import itertools

ACTIVE_TABLE_XPATH = "./office:settings/config:config-item-set[@config:name='ooo:view-settings']/config:config-item-map-indexed[@config:name='Views']/config:config-item-map-entry/config:config-item[@config:name='ActiveTable']"

OFFICE_NS_DICT = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
}


class NameSpace:
    def __init__(self, ns_dict=OFFICE_NS_DICT):
        self._ns_dict = ns_dict

    def findall(self, element, match):
        return element.findall(match, self._ns_dict)

    def attrib(self, attr):
        try:
            i = attr.index(":")
            attr = "{{{}}}{}".format(self._ns_dict.get(attr[:i], attr[:i]),
                                     attr[i + 1:])
        except ValueError:
            pass

        return attr


OFFICE_NS = NameSpace(OFFICE_NS_DICT)


class OdsTables:
    """An iterator over tables in an ods file"""

    @staticmethod
    def create(fullpath, ns=OFFICE_NS):
        return OdsTablesBuilder(fullpath).ns(ns).build()

    def __init__(self, root, sort_func, ns=OFFICE_NS):
        self._root = root
        self._sort_func = sort_func
        self._ns = ns

    def __iter__(self):
        tables = self._ns.findall(self._root,
                                  "./office:body/office:spreadsheet/table:table")
        tables = self._sort_func(tables)
        for table in tables:
            yield table


class OdsTablesBuilder:
    def __init__(self, fullpath):
        self._fullpath = fullpath
        self._ns = OFFICE_NS
        self._sort_func_creator = dont_sort

    def ns(self, ns):
        self._ns = ns
        return self

    def sort_func_creator(self, sort_func_creator):
        self._sort_func_creator = sort_func_creator

    def build(self):
        with zipfile.ZipFile(self._fullpath) as z:
            sort_func = self._sort_func_creator(z, self._ns)
            with z.open('content.xml') as content:
                data = content.read().decode("utf-8")
                root = ET.fromstring(data)
                return OdsTables(root, sort_func, self._ns)


def dont_sort(*_args, **_kwargs):
    def sort_func(tables):
        return tables

    return sort_func


def put_active_first(z, ns=OFFICE_NS):
    """
    :param z: the zip file
    :param ns: the name space
    :returns: a sort function for tables
    """
    active_table_name = _find_active_table_name_in_zip(z, ns)

    def sort_func(tables):
        for i, table in enumerate(tables):
            if table.get(ns.attrib("table:name")) == active_table_name:
                return [tables[i]] + tables[:i] + tables[i + 1:]

        return tables

    return sort_func


def _find_active_table_name_in_zip(z, ns=OFFICE_NS):
    with z.open('settings.xml') as settings:
        active_table_name = _find_active_table_name(settings, ns)

    return active_table_name


def _find_active_table_name(settings, ns=OFFICE_NS):
    data = settings.read().decode("utf-8")
    root = ET.fromstring(data)
    active_tables = ns.findall(root, ACTIVE_TABLE_XPATH)
    if active_tables:
        active_table_name = active_tables[0].text
    else:
        active_table_name = None
    return active_table_name


class OdsRows:
    """An iterator over rows in a table.
    The table must be simple (no spanned rows or repeated rows)"""

    def __init__(self, table, omit=None, ns=OFFICE_NS):
        self._table = table
        self._ns = ns
        self._omit = omit

    def __iter__(self):
        for row in self._ns.findall(self._table, ".//table:table-row"):
            if self._omit is not None and self._omit(row, self._ns):
                continue

            count = int(row.get(self._ns.attrib("table:number-rows-repeated"),
                                row.get(self._ns.attrib(
                                    "table:number-rows-spanned"), "1")))
            cell_tags = (
                self._ns.attrib("table:table-cell"),
                self._ns.attrib("table:covered-table-cell")
            )
            cell_elements = [e for e in row if e.tag in cell_tags]
            cells = [v for c in cell_elements for v in self._values(c)]
            cells = self._trim_list(cells)
            for _ in range(count):
                yield cells

    def __getitem__(self, index):
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

    def _trim_list(self, l):
        if not l:
            return l
        for i in range(len(l) - 1, -1, -1):
            if len(l[i]):
                return l[:i + 1]
        return []

    def _values(self, c):
        spanned = self._ns.attrib("table:number-columns-spanned")
        repeated = self._ns.attrib("table:number-columns-repeated")
        count = int(c.get(repeated, c.get(spanned, "1")))
        v = '\n'.join(p.text for p in self._ns.findall(c, "./text:p") if p.text)
        return [v] * count


def omit_filtered(row, ns):
    return row.get(ns.attrib("table:visibility"), "") == "filter"
