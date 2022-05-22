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

from py4lo_helper import *
import py4lo_helper
from py4lo_helper import _ObjectProvider, _Inspector


class TestHelper(unittest.TestCase):
    def setUp(self):
        self.xsc = Mock()
        self.doc = Mock()
        self.ctrl = Mock()
        self.frame = Mock()
        self.parent_win = Mock()
        self.msp = Mock()
        self.ctxt = Mock()
        self.sm = Mock()
        self.desktop = Mock()
        py4lo_helper.provider = _ObjectProvider(self.doc, self.ctrl, self.frame,
                                                self.parent_win, self.msp,
                                                self.ctxt, self.sm,
                                                self.desktop)
        py4lo_helper._inspect = _Inspector(py4lo_helper.provider)

    def testXray(self):
        py4lo_helper._inspect.use_xray()
        self.msp.getScript.assert_called_once_with(
            'vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application')

    def testXray2(self):
        py4lo_helper._inspect.xray(1)
        py4lo_helper._inspect.xray(2)
        self.msp.getScript.assert_called_once_with(
            'vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application')
        self.msp.getScript.return_value.invoke.assert_has_calls(
            [call((1,), (), ()), call((2,), (), ())])

    def testPv(self):
        pv = make_pv("name", "value")
        self.assertTrue("uno.com.sun.star.beans.PropertyValue" in str(type(pv)))
        self.assertEqual("name", pv.Name)
        self.assertEqual("value", pv.Value)

    def testPvs(self):
        pvs = make_pvs({"name1": "value1", "name2": "value2"})
        pvs = sorted(pvs, key=lambda pv: pv.Name)
        self.assertEqual("name1", pvs[0].Name)
        self.assertEqual("value1", pvs[0].Value)
        self.assertEqual("name2", pvs[1].Name)
        self.assertEqual("value2", pvs[1].Value)


    def testUnoService(self):
        uno_service_ctxt("x")
        self.sm.createInstanceWithContext.assert_called_once_with("x",
                                                                  self.ctxt)

        uno_service_ctxt("y", [1, 2, 3])
        self.sm.createInstanceWithArgumentsAndContext.assert_called_once_with(
            "y", [1, 2, 3], self.ctxt)

        uno_service("z")
        self.sm.createInstance.assert_called_once_with("z")

    def test_read_options(self):
        oSheet = Mock()
        aAdress = Mock()
        aAdress.EndColumn = 10
        aAdress.StartColumn = 0
        aAdress.EndRow = 10
        aAdress.StartRow = 0
        read_options(oSheet, aAdress)

    def test_set_validation_list(self):
        oCell = Mock()
        try:
            py4lo_helper.provider.set_validation_list_by_cell(oCell,
                                                              ["a", "b", "c"])
        except Exception as e:
            pass

    def testDocBuilder(self):

        d = DocBuilder("calc")
        d.build()
        py4lo_helper.provider.desktop.loadComponentFromURL.assert_called_once_with(
            "private:factory/scalc", "_blank", 0, ())
        oDoc = py4lo_helper.provider.desktop.loadComponentFromURL.return_value
        oDoc.lockControllers.assert_called_once()
        oDoc.unlockControllers.assert_called_once()

    def testDocBuilderSheetNames(self):
        oDoc = py4lo_helper.provider.desktop.loadComponentFromURL.return_value
        oDoc.Sheets.getCount.side_effect = [3, 3, 3, 3, 3]
        s1, s2, s3 = MagicMock(), MagicMock(), MagicMock()
        oDoc.Sheets.getByIndex.side_effect = [s1, s2, s3]

        d = DocBuilder("calc")
        d.sheet_names("abcdef", expand_if_necessary=True)
        d.build()
        py4lo_helper.provider.desktop.loadComponentFromURL.assert_called_once_with(
            "private:factory/scalc", "_blank", 0, ())
        oDoc.lockControllers.assert_called_once()
        oDoc.Sheets.getCount.assert_called()

        oDoc.Sheets.getByIndex.assert_has_calls([call(0), call(1), call(2)])
        s1.setName.assert_called_once_with("a")
        s2.setName.assert_called_once_with("b")
        s3.setName.assert_called_once_with("c")
        oDoc.Sheets.insertNewByName.assert_has_calls(
            [call("d", 3), call("e", 4), call("f", 5)])

        oDoc.unlockControllers.assert_called_once()

    def testDocBuilderSheetNames2(self):
        oDoc = py4lo_helper.provider.desktop.loadComponentFromURL.return_value
        oDoc.Sheets.getCount.side_effect = [3, 3, 3, 3, 3, 2]
        s1, s2, s3 = MagicMock(), MagicMock(), MagicMock()
        oDoc.Sheets.getByIndex.side_effect = [s1, s2, s3, s3]

        d = DocBuilder("calc")
        d.sheet_names("ab", trunc_if_necessary=True)
        d.build()
        py4lo_helper.provider.desktop.loadComponentFromURL.assert_called_once_with(
            "private:factory/scalc", "_blank", 0, ())
        oDoc.lockControllers.assert_called_once()
        oDoc.Sheets.getCount.assert_called()

        oDoc.Sheets.getByIndex.assert_has_calls([call(0), call(1), call(2)])
        s1.setName.assert_called_once_with("a")
        s2.setName.assert_called_once_with("b")
        s3.getName.assert_called_once()
        oDoc.Sheets.removeByName.assert_called_once_with(
            s3.getName.return_value)

        oDoc.unlockControllers.assert_called_once()



if __name__ == '__main__':
    unittest.main()
