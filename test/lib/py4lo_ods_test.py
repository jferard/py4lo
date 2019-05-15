# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   FastODS is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   FastODS is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
import unittest
from unittest.mock import *
import env
import sys, os
import uno

from py4lo_ods import *
import xml.etree.ElementTree as ET

CONTENT_XML = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    office:version="1.2"><office:scripts/>
    <office:font-face-decls/>
    <office:automatic-styles/>
    <office:body>
        <office:spreadsheet>
            <table:table table:name="Test"><table:table-column  table:number-columns-repeated="4" table:default-cell-style-name="Default"/>
                <table:table-row table:style-name="ro1">
                    <table:table-cell office:value-type="string" table:number-columns-spanned="2" table:number-rows-spanned="1">
                        <text:p>A1:B1</text:p>
                    </table:table-cell>
                    <table:covered-table-cell/>
                    <table:table-cell office:value-type="string">
                        <text:p>C1</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>D1</text:p>
                    </table:table-cell>
                </table:table-row><table:table-row table:style-name="ro1">
                    <table:table-cell office:value-type="string" table:number-columns-spanned="1" table:number-rows-spanned="2">
                        <text:p>A2:A3</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string" table:number-columns-spanned="3" table:number-rows-spanned="2">
                        <text:p>B2:D3</text:p>
                    </table:table-cell><table:covered-table-cell table:number-columns-repeated="2"/>
                </table:table-row><table:table-row table:style-name="ro1">
                    <table:covered-table-cell table:number-columns-repeated="4"/>
                </table:table-row><table:table-row table:style-name="ro1">
                    <table:table-cell office:value-type="string">
                        <text:p>A4</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>B4</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>C4</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>D4</text:p>
                    </table:table-cell>
                </table:table-row>
                <table:table-row table:style-name="ro1">
                    <table:table-cell office:value-type="string">
                        <text:p>A5</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>B5</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>C5</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>D5</text:p>
                    </table:table-cell>
                </table:table-row>
                <table:table-row table:style-name="ro1">
                    <table:table-cell office:value-type="string">
                        <text:p>A6</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>B6</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>C6</text:p>
                    </table:table-cell>
                    <table:table-cell office:value-type="string">
                        <text:p>D6</text:p>
                    </table:table-cell>
                </table:table-row>
            </table:table><table:named-expressions/></office:spreadsheet>
    </office:body>
</office:document-content>"""


class TestOds(unittest.TestCase):
    def setUp(self):
        root = ET.fromstring(CONTENT_XML)
        table = root.find("./office:body/office:spreadsheet/table:table", OFFICE_NS)
        self.ods_rows = OdsRows(table)

    def test_list(self):
        self.assertEquals([
            ["A1:B1", "A1:B1", "C1", "D1"],
            ["A2:A3", "B2:D3", "B2:D3", "B2:D3"],
            [],
            ["A4", "B4", "C4", "D4"],
            ["A5", "B5", "C5", "D5"],
            ["A6", "B6", "C6", "D6"],
        ], list(self.ods_rows))

    def test_get_item1(self):
        self.assertEquals(["A2:A3", "B2:D3", "B2:D3", "B2:D3"], self.ods_rows[1])

    def test_get_item_neg1(self):
        self.assertEquals(["A6", "B6", "C6", "D6"], self.ods_rows[-1])

    def test_get_slice(self):
        self.assertEquals([
            ["A2:A3", "B2:D3", "B2:D3", "B2:D3"],
            ["A4", "B4", "C4", "D4"],
            ["A6", "B6", "C6", "D6"],
        ], self.ods_rows[1::2])

if __name__ == '__main__':
    unittest.main()
