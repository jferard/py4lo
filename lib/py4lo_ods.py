# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

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

OFFICE_NS_DICT = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}


class NameSpace:
    def __init__(self, ns_dict=OFFICE_NS_DICT):
        self._ns_dict = ns_dict

    def findall(self, element, match):
        return element.findall(match, self._ns_dict)

    def attrib(self, attr):
        try:
            i = attr.index(":")
            attr = "{{{}}}{}".format(self._ns_dict.get(attr[:i], attr[:i]), attr[i + 1:])
        except ValueError:
            pass

        return attr


OFFICE_NS = NameSpace(OFFICE_NS_DICT)


class OdsTables:
    """An iterator over tables in an ods file"""

    @staticmethod
    def create(fullpath, ns=OFFICE_NS):
        with zipfile.ZipFile(fullpath) as list_ods:
            with list_ods.open('content.xml') as content:
                data = content.read().decode("utf-8")
                root = ET.fromstring(data)
                return OdsTables(root, ns)

    def __init__(self, root, ns=OFFICE_NS):
        self._root = root
        self._ns = ns

    def __iter__(self):
        tables = self._ns.findall(self._root, "./office:body/office:spreadsheet/table:table")
        for table in tables:
            yield table


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
                                row.get(self._ns.attrib("table:number-rows-spanned"), "1")))
            cell_elements = self._ns.findall(row, "./table:table-cell")
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
            return list(itertools.islice(self, index.start, index.stop, index.step))
        else:
            raise TypeError("index must be int or slice")

    def _trim_list(self, l):
        if not l:
            return l
        for i in range(len(l) - 1, -1, -1):
            if len(l[i]):
                return l[:i + 1]

    def _values(self, c):
        spanned = self._ns.attrib("table:number-columns-spanned")
        repeated = self._ns.attrib("table:number-columns-repeated")
        count = int(c.get(repeated, c.get(spanned, "1")))
        v = '\n'.join(p.text for p in self._ns.findall(c, "./text:p") if p.text)
        return [v] * count


def omit_filtered(row, ns):
    return row.get(ns.attrib("table:visibility"), "") == "filter"
