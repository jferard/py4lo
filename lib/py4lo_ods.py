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

OFFICE_NS = {
    "office":"urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "table":"urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "text":"urn:oasis:names:tc:opendocument:xmlns:text:1.0",
}

class OdsTables():
    """An iterator over tables in an ods file"""
    @staticmethod
    def create(fullpath, ns=OFFICE_NS):
        with zipfile.ZipFile(fullpath) as list_ods:
            with list_ods.open('content.xml') as content:
                data = content.read().decode("utf-8")
                root = ET.fromstring(data)
                return OdsTablesLister(root, ns)

    def __init__(self, table, ns=OFFICE_NS):
        self.__root = root
        self.__ns = ns

    def __iter__():
        tables = self.__root.findall("./office:body/office:spreadsheet/table:table", self.__ns)
        for table in tables:
            yield table

class OdsRows():
    """An iterator over rows in a table.
    The table must be simple (no spanned rows or repeated rows)"""
    
    def __init__(self, table, ns=OFFICE_NS):
        self.__table = table
        self.__ns = ns

    def __iter__(self):
        for row in self.__table.findall("./table:table-row", self.__ns):
            cells = row.findall("./table:table-cell", self.__ns)
            l = [v for c in cells for v in self.__values(c)]
            yield self.__trim_list(l)

    def __trim_list(self, l):
        for i in range(len(l)-1, -1,-1):
            if len(l[i]):
                return l[:i]

    def __values(self, c):
        count = int(c.get(self.__attrib("table:number-columns-repeated"), c.get(self.__attrib("table:number-columns-spanned"), "1")))
        v = ''.join(c.itertext())
        return [v]*count

    def __attrib(self, attr):
        try:
            i = attr.index(":")
            attr = "{{{}}}{}".format(self.__ns.get(attr[:i], attr[:i]), attr[i+1:])
        except:
            pass

        return attr
