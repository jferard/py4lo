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
from callbacks.add_debug_content import AddDebugContent
from callbacks.add_readme_with import AddReadmeWith
from callbacks.ignore_scripts import IgnoreScripts
from callbacks.rewrite_manifest import RewriteManifest
from callbacks.add_scripts import AddScripts, ARC_SCRIPTS_PATH

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
    add_scripts_cb = AddScripts(scripts)
    def callback(zout):
        return add_scripts_cb.call(zout)

    return callback

def add_readme_with(contact):
    add_readme_with_cb = AddReadmeWith(os.path.join(py4lo_path, "inc"), contact)
    def callback(zout):
        return add_readme_with_cb.call(zout)

    return callback

def add_debug_content(funcs_by_script):
    add_debug_content_cb = AddDebugContent(funcs_by_script)
    def callback(zout):
        return add_debug_content_cb.call(zout)

    return callback
