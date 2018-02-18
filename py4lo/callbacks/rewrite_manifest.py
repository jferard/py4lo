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
import os
import xml.dom.minidom

MANIFEST_CLOSE_TAG = "</manifest:manifest>"
BASIC_AND_PYTHON_ENTRIES = """
    <manifest:file-entry manifest:full-path="Basic/Standard/py4lo.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Scripts" manifest:media-type="application/binary"/>
    <manifest:file-entry manifest:full-path="Scripts/python" manifest:media-type="application/binary"/>
"""
FILE_ENTRY_TPL = """    <manifest:file-entry manifest:full-path="Scripts/python/{0}" manifest:media-type=""/>
"""

class RewriteManifest():
    def __init__(self, scripts):
        self.__scripts = scripts

    def call(self, zin, zout, item):
        if item.filename == "META-INF/manifest.xml":
            data = self.__rewrite_manifest(zin, item)
        else:
            data = zin.read(item.filename) # copy

        zout.writestr(item.filename, data)
        return True

    def __rewrite_manifest(self, zin, manifest_item):
        pretty_manifest = self.__prettyfy_xml(zin, manifest_item.filename)
        s = self.__strip_close(pretty_manifest)
        s = self.__add_script_lines(s)
        return s.encode("utf-8")

    def __prettyfy_xml(self, zin, fname):
        data = zin.read(fname)
        xml_str = xml.dom.minidom.parseString(data.decode("utf-8"))
        return xml_str.toprettyxml(indent="   ", newl='')

    def __strip_close(self, pretty_manifest):
        s = ""
        for line in pretty_manifest.splitlines():
            if line.strip() == MANIFEST_CLOSE_TAG: # end of manifest
                return s
            s += line + "\n"

        raise Exception("no manifest closing tag in "+pretty_manifest)

    def __add_script_lines(self, s):
        s += BASIC_AND_PYTHON_ENTRIES
        for script in self.__scripts:
            s += FILE_ENTRY_TPL.format(script.get_name())
        s += MANIFEST_CLOSE_TAG
        return s
