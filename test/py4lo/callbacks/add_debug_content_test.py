# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>

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
import unittest
import env
from callbacks import *
import io
import zipfile


class TestAddDebugContent(unittest.TestCase):
    BEFORE = """<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:dr3d="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:ooo="http://openoffice.org/2004/office" xmlns:ooow="http://openoffice.org/2004/writer" xmlns:oooc="http://openoffice.org/2004/calc" xmlns:dom="http://www.w3.org/2001/xml-events" xmlns:xforms="http://www.w3.org/2002/xforms" xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:rpt="http://openoffice.org/2005/report" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:xhtml="http://www.w3.org/1999/xhtml" xmlns:grddl="http://www.w3.org/2003/g/data-view#" xmlns:tableooo="http://openoffice.org/2009/table" xmlns:drawooo="http://openoffice.org/2010/draw" xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0" xmlns:field="urn:openoffice:names:experimental:ooo-ms-interop:xmlns:field:1.0" xmlns:formx="urn:openoffice:names:experimental:ooxml-odf-interop:xmlns:form:1.0" xmlns:css3t="http://www.w3.org/TR/css3-text/" office:version="1.2"><office:scripts/><office:font-face-decls><style:font-face style:name="Liberation Sans" svg:font-family="&apos;Liberation Sans&apos;" style:font-family-generic="swiss" style:font-pitch="variable"/><style:font-face style:name="Arial" svg:font-family="Arial" style:font-family-generic="system" style:font-pitch="variable"/><style:font-face style:name="Microsoft YaHei" svg:font-family="&apos;Microsoft YaHei&apos;" style:font-family-generic="system" style:font-pitch="variable"/><style:font-face style:name="Segoe UI" svg:font-family="&apos;Segoe UI&apos;" style:font-family-generic="system" style:font-pitch="variable"/><style:font-face style:name="Tahoma" svg:font-family="Tahoma" style:font-family-generic="system" style:font-pitch="variable"/></office:font-face-decls><office:automatic-styles><style:style style:name="co1" style:family="table-column"><style:table-column-properties fo:break-before="auto" style:column-width="22.58mm"/></style:style><style:style style:name="ro1" style:family="table-row"><style:table-row-properties style:row-height="4.52mm" fo:break-before="auto" style:use-optimal-row-height="true"/></style:style><style:style style:name="ta1" style:family="table" style:master-page-name="Default"><style:table-properties table:display="true" style:writing-mode="lr-tb"/></style:style><style:style style:name="P1" style:family="paragraph"><style:paragraph-properties fo:text-align="center"/></style:style></office:automatic-styles><office:body><office:spreadsheet><table:calculation-settings table:automatic-find-labels="false"/><table:table table:name="Feuille1" table:style-name="ta1">"""

    AFTER = """<table:table-column table:style-name="co1" table:default-cell-style-name="Default"/><table:table-row table:style-name="ro1"><table:table-cell/></table:table-row></table:table><table:named-expressions/></office:spreadsheet></office:body></office:document-content>"""

    def setUp(self):
        self.maxDiff = None

    def test_add_debug_content_empty(self):
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w')
        AddDebugContent({}).call(zout)
        self.assertEqual(TestAddDebugContent.BEFORE+
        """<office:forms form:automatic-focus="false" form:apply-design-mode="false"></office:forms><table:shapes></table:shapes>"""+TestAddDebugContent.AFTER, zout.read("content.xml").decode("utf-8"))

    def test_add_debug_content_one_empty_script(self):
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w')
        AddDebugContent({"s1":[]}).call(zout)
        self.assertEqual(TestAddDebugContent.BEFORE+
        """<office:forms form:automatic-focus="false" form:apply-design-mode="false"></office:forms><table:shapes></table:shapes>"""+TestAddDebugContent.AFTER, zout.read("content.xml").decode("utf-8"))

    def test_add_debug_content_two_empty_scripts(self):
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w')
        AddDebugContent({"s1":[], "s2":[]}).call(zout)
        self.assertEqual(TestAddDebugContent.BEFORE+
        """<office:forms form:automatic-focus="false" form:apply-design-mode="false"></office:forms><table:shapes></table:shapes>"""+TestAddDebugContent.AFTER, zout.read("content.xml").decode("utf-8"))

    def test_add_debug_content_one_scripts_one_func(self):
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w')
        AddDebugContent({"s1":["f1"]}).call(zout)
        self.assertEqual(TestAddDebugContent.BEFORE+
        """<office:forms form:automatic-focus="false" form:apply-design-mode="false"><form:form form:name="Formulaire" form:apply-filter="true" form:command-type="table" form:control-implementation="ooo:com.sun.star.form.component.Form" office:target-frame="" xlink:href="" xlink:type="simple"><form:properties><form:property form:property-name="PropertyChangeNotificationEnabled" office:value-type="boolean" office:boolean-value="true"/></form:properties><form:button form:name="name0" form:control-implementation="ooo:com.sun.star.form.component.CommandButton" xml:id="control0" form:id="control0" form:label="f1" office:target-frame="" xlink:href="" form:image-data="" form:delay-for-repeat="PT0.050000000S" form:image-position="center"><form:properties><form:property form:property-name="DefaultControl" office:value-type="string" office:string-value="com.sun.star.form.control.CommandButton"/></form:properties><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="form:performaction" xlink:href="vnd.sun.star.script:s1$f1?language=Python&amp;location=document" xlink:type="simple"/></office:event-listeners></form:button></form:form></office:forms><table:shapes><draw:control draw:z-index="0" draw:text-style-name="P1" svg:width="80mm" svg:height="10mm" svg:x="10mm" svg:y="10mm" draw:control="control0"/></table:shapes>"""+
        TestAddDebugContent.AFTER, zout.read("content.xml").decode("utf-8"))

    def test_add_debug_content_many_scripts(self):
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w')
        AddDebugContent({"s1":["f11","f12"], "s2":["f21"]}).call(zout)
        self.assertEqual(TestAddDebugContent.BEFORE+
        """<office:forms form:automatic-focus="false" form:apply-design-mode="false"><form:form form:name="Formulaire" form:apply-filter="true" form:command-type="table" form:control-implementation="ooo:com.sun.star.form.component.Form" office:target-frame="" xlink:href="" xlink:type="simple"><form:properties><form:property form:property-name="PropertyChangeNotificationEnabled" office:value-type="boolean" office:boolean-value="true"/></form:properties><form:button form:name="name0" form:control-implementation="ooo:com.sun.star.form.component.CommandButton" xml:id="control0" form:id="control0" form:label="f11" office:target-frame="" xlink:href="" form:image-data="" form:delay-for-repeat="PT0.050000000S" form:image-position="center"><form:properties><form:property form:property-name="DefaultControl" office:value-type="string" office:string-value="com.sun.star.form.control.CommandButton"/></form:properties><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="form:performaction" xlink:href="vnd.sun.star.script:s1$f11?language=Python&amp;location=document" xlink:type="simple"/></office:event-listeners></form:button></form:form><form:form form:name="Formulaire" form:apply-filter="true" form:command-type="table" form:control-implementation="ooo:com.sun.star.form.component.Form" office:target-frame="" xlink:href="" xlink:type="simple"><form:properties><form:property form:property-name="PropertyChangeNotificationEnabled" office:value-type="boolean" office:boolean-value="true"/></form:properties><form:button form:name="name1" form:control-implementation="ooo:com.sun.star.form.component.CommandButton" xml:id="control1" form:id="control1" form:label="f12" office:target-frame="" xlink:href="" form:image-data="" form:delay-for-repeat="PT0.050000000S" form:image-position="center"><form:properties><form:property form:property-name="DefaultControl" office:value-type="string" office:string-value="com.sun.star.form.control.CommandButton"/></form:properties><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="form:performaction" xlink:href="vnd.sun.star.script:s1$f12?language=Python&amp;location=document" xlink:type="simple"/></office:event-listeners></form:button></form:form><form:form form:name="Formulaire" form:apply-filter="true" form:command-type="table" form:control-implementation="ooo:com.sun.star.form.component.Form" office:target-frame="" xlink:href="" xlink:type="simple"><form:properties><form:property form:property-name="PropertyChangeNotificationEnabled" office:value-type="boolean" office:boolean-value="true"/></form:properties><form:button form:name="name2" form:control-implementation="ooo:com.sun.star.form.component.CommandButton" xml:id="control2" form:id="control2" form:label="f21" office:target-frame="" xlink:href="" form:image-data="" form:delay-for-repeat="PT0.050000000S" form:image-position="center"><form:properties><form:property form:property-name="DefaultControl" office:value-type="string" office:string-value="com.sun.star.form.control.CommandButton"/></form:properties><office:event-listeners><script:event-listener script:language="ooo:script" script:event-name="form:performaction" xlink:href="vnd.sun.star.script:s2$f21?language=Python&amp;location=document" xlink:type="simple"/></office:event-listeners></form:button></form:form></office:forms><table:shapes><draw:control draw:z-index="0" draw:text-style-name="P1" svg:width="80mm" svg:height="10mm" svg:x="10mm" svg:y="10mm" draw:control="control0"/><draw:control draw:z-index="0" draw:text-style-name="P1" svg:width="80mm" svg:height="10mm" svg:x="10mm" svg:y="25mm" draw:control="control1"/><draw:control draw:z-index="0" draw:text-style-name="P1" svg:width="80mm" svg:height="10mm" svg:x="10mm" svg:y="40mm" draw:control="control2"/></table:shapes>"""+TestAddDebugContent.AFTER, zout.read("content.xml").decode("utf-8"))

if __name__ == '__main__':
    unittest.main()
