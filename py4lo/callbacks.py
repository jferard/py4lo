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
import os
from debug import *

ARC_SCRIPTS_PATH = "Scripts/python"

def ignore_scripts(zin, zout, item):
    return not item.filename.startswith(ARC_SCRIPTS_PATH)

def rewrite_manifest(scripts):
    def callback(zin, zout, item):
        path, fname = os.path.split(item.filename)
        if item.filename == "META-INF/manifest.xml":
            temp = zin.read(item.filename)
            s = ""
            for line in temp.decode("utf-8").splitlines():
                if line.strip() == "</manifest:manifest>": # end of manifest
                    s += """<manifest:file-entry manifest:full-path="Basic/Standard/py4lo.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>
    <manifest:file-entry manifest:full-path="Scripts" manifest:media-type="application/binary"/>
    <manifest:file-entry manifest:full-path="Scripts/python" manifest:media-type="application/binary"/>
    """
                    for script in scripts:
                        s += " <manifest:file-entry manifest:full-path=\"Scripts/python/"+script.get_name()+"\" manifest:media-type=\"\"/>\n"

                s += line + "\n"
            zout.writestr(item, s.encode("utf-8"))
        else:
            zout.writestr(item.filename, zin.read(item.filename))
        return True
        
    return callback

def add_scripts(scripts):
    def callback(zout):
        py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")
        for script in scripts:
            zout.writestr(ARC_SCRIPTS_PATH+"/"+script.get_name(), script.get_data())
        return True
        
    return callback
        
def add_readme_cb(contact):
    def callback(zout):
        py4lo_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), "..")
        zout.write(os.path.join(py4lo_path, "inc", "script-lc.xml"), "Basic/script-lc.xml")
        zout.write(os.path.join(py4lo_path, "inc", "script-lb.xml"), "Basic/Standard/script-lb.xml")
        with open(os.path.join(py4lo_path, "inc", "py4lo.xml.tpl"), 'r', encoding='utf-8') as f:
            tpl = f.read()
            xml = tpl.format(contact = contact)
            zout.writestr("Basic/Standard/py4lo.xml", xml)
        return True
    
    return callback

    
def add_debug_content(funcs_by_script):
    def callback(zout):
        forms = begin_forms
        draw = begin_shapes

        i = 0
        for script in funcs_by_script:
            for func in funcs_by_script[script]:
                forms += form_tpl.format(name=func, id=i, file=script, func=func)
                draw += draw_control_tpl.format(x=10, y=15*i+10, id=i)
                i += 1

        forms += end_forms
        draw += end_shapes

        s = before + forms + draw + after
        zout.writestr("content.xml", s)
        return True
        
    return callback
