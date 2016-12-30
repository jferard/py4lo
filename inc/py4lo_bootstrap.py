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
def use_local(object_ref):
    try:
        doc = XSCRIPTCONTEXT.getDocument()
        dsp = doc.getScriptProvider()
        
        (fname_wo_py, oname) = object_ref.split("::")
        import_script = dsp.getScript("vnd.sun.star.script:"+fname_wo_py+".py$__export_"+oname+"?language=Python&location=document")
        return import_script.invoke((), (), ())[0]
    except:
        return None