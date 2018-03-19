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
import unittest
import env
from callbacks import *
import io
import zipfile
from scripts_processor import TargetScript

class TestRewriteManifest(unittest.TestCase):
    def setUp(self):
        self.maxDiff = None
        self.__temp = io.BytesIO()
        ztemp = zipfile.ZipFile(self.__temp, 'w')
        ztemp.writestr("x", "y")
        ztemp.writestr("META-INF/manifest.xml", """<?xml version="1.0" encoding="UTF-8"?>
        <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
        </manifest:manifest>""")
        ztemp.close()

    def test_rewrite_manifest_empty(self):
        zin = zipfile.ZipFile(self.__temp, 'r')
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w')

        RewriteManifest([]).call(zin, zout, zin.getinfo("x"))
        self.assertEquals(b"y", zout.read("x"))
        RewriteManifest([]).call(zin, zout, zin.getinfo("META-INF/manifest.xml"))
        self.assertEquals("""<?xml version="1.0" ?><manifest:manifest manifest:version="1.2" xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
<manifest:file-entry manifest:full-path="Basic/Standard/py4lo.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Scripts" manifest:media-type="application/binary"/>
    <manifest:file-entry manifest:full-path="Scripts/python" manifest:media-type="application/binary"/>
</manifest:manifest>""", zout.read("META-INF/manifest.xml").decode("utf-8"))

    def test_rewrite_manifest_one_script(self):
        zin = zipfile.ZipFile(self.__temp, 'r')
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w')

        RewriteManifest([TargetScript("script", None, None, None)]).call(zin, zout, zin.getinfo("META-INF/manifest.xml"))
        print (zout.read("META-INF/manifest.xml").decode("utf-8"))
        self.assertEquals("""<?xml version="1.0" ?><manifest:manifest manifest:version="1.2" xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
<manifest:file-entry manifest:full-path="Basic/Standard/py4lo.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Scripts" manifest:media-type="application/binary"/>
    <manifest:file-entry manifest:full-path="Scripts/python" manifest:media-type="application/binary"/>
    <manifest:file-entry manifest:full-path="Scripts/python/script" manifest:media-type=""/>
</manifest:manifest>""", zout.read("META-INF/manifest.xml").decode("utf-8"))

if __name__ == '__main__':
    unittest.main()
