# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2021 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   Py4LO is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   THIS FILE IS SUBJECT TO THE "CLASSPATH" EXCEPTION.

   Py4LO is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""
def use_local(object_ref):
    doc = XSCRIPTCONTEXT.getDocument()
    dsp = doc.getScriptProvider()
    (fname_wo_py, oname) = object_ref.split("::")
    uri = "vnd.sun.star.script:"+fname_wo_py+".py$__export_"+oname+"?language=Python&location=document"
    import_script = dsp.getScript(uri)
    return import_script.invoke((), (), ())[0]
