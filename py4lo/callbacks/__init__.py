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
from callbacks.debug import *
from callbacks.ignore_scripts import IgnoreScripts
from callbacks.rewrite_manifest import RewriteManifest

ARC_SCRIPTS_PATH = "Scripts/python"
py4lo_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.realpath(__file__))))

ignore_scripts_cb = IgnoreScripts(ARC_SCRIPTS_PATH)

def ignore_scripts(zin, zout, item):
    return ignore_scripts_cb.call(zin, zout, item)

def rewrite_manifest(scripts):
    rewrite_manifest_cb = RewriteManifest(scripts)
    def callback(zin, zout, item):
        rewrite_manifest_cb.call(zin, zout, item)

    return callback

def add_scripts(scripts):
    def callback(zout):
        for script in scripts:
            zout.writestr(ARC_SCRIPTS_PATH+"/"+script.get_name(), script.get_data())
        return True

    return callback

def add_readme_with(contact):
    def callback(zout):
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
