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
import uno
# py4lo: if python_version >= 2.6
# py4lo: if python_version <= 3.0
import io
# py4lo: endif
# py4lo: endif
from com.sun.star.uno import RuntimeException
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
# py4lo: if python_version >= 3.0
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
# py4lo: else
MESSAGEBOX = 0
# py4lo: endif

class Py4LO_helper():
	def __init__(self):
		self.doc = XSCRIPTCONTEXT.getDocument()
		self.ctxt = uno.getComponentContext()

		self.ctrl = self.doc.CurrentController					
		self.frame = self.ctrl.Frame
		self.parent_win = self.frame.ContainerWindow
		self.sm = self.ctxt.getServiceManager()
		self.dsp = self.doc.getScriptProvider()
		
		mspf = self.uno_service_ctxt("com.sun.star.script.provider.MasterScriptProviderFactory")
		self.msp = mspf.createScriptProvider("")

		self.loader = self.uno_service( "com.sun.star.frame.Desktop" )
		self.reflect = self.uno_service( "com.sun.star.reflection.CoreReflection" )
		self.__xray_script = None
		
	def use_xray(self):
		try:
			self.__xray_script = self.msp.getScript("vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application")
		except:
			raise RuntimeException("\nBasic library Xray is not installed", self.ctxt)
	
	def xray(self, object):
		if self.__xray_script is None:
			self.use_xray()
		
		self.__xray_script.invoke((object,), (), ())
		
	def make_pv(self, name, value):
		pv = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
		pv.Name = name
		pv.Value = value
		return pv
		
	def make_pvs(self, d):
		l = []
		for k in d:
			l.append(self.make_pv(k, d[k]))
		return tuple(l)

	# from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603# (thanks Bernard !)
	def message_box(self, parent_win, msg_text, msg_title, msg_type=MESSAGEBOX, msg_buttons=BUTTONS_OK):
		sv = self.uno_service_ctxt("com.sun.star.awt.Toolkit")
		my_box = sv.createMessageBox(parent_win, msg_type, msg_buttons, msg_title, msg_text)
		return my_box.execute()

	def uno_service_ctxt(self, sname, args=None):
		if args is None:
			return self.sm.createInstanceWithContext(sname, self.ctxt)
		else:
			return self.sm.createInstanceWithArgumentsAndContext(sname, args, self.ctxt)
		
	def uno_service(self, sname, args=None, ctxt=None):
		if ctxt is None:
			return self.sm.createInstance(sname)
		else:
			if args is None:
				return self.sm.createInstanceWithContext(sname, ctxt)
			else:
				return self.sm.createInstanceWithArgumentsAndContext(sname, args, ctxt)

	def new_doc(self):
		return self.loader.loadComponentFromURL(
	                 "private:factory/scalc", "_blank", 0, () )
					 
	def get_last_used_row(self, oSheet):
		oCell = oSheet.GetCellByPosition(0, 0)
		oCursor = oSheet.createCursorByRange(oCell)
		oCursor.GotoEndOfUsedArea(True)
		aAddress = oCursor.RangeAddress
		return aAddress.EndRow
		
	def to_iter(self, o):
		for i in range(0, oXIndexAccess.getCount()):
			yield(oXIndexAccess.getByIndex(i))

	def to_dict(self, o):
		d = {}
		for name in o.getElementNames():
			d[name] = o.getByName(name)
		return d
			
# py4lo: if python_version >= 2.6
	def open_pmlog(self, fpath):
# py4lo: if python_version < 3.0
		self.__log_out = io.open(fpath, "w", encoding="utf-8")
# py4lo: else
		self.__log_out = open(fpath, "w", encoding="utf-8")
# py4lo: endif
			
	def pmlog(self, text):
		self.__log_out.write(text+"\n")
		self.__log_out.flush()

	def __del__(self):
		if self.__log_out is not None:
			self.__log_out.close()
# py4lo: endif
		
	def cur_dir(self):
		url = self.doc.getURL()
		path = uno.fileUrlToSystemPath( url )
		return os.path.dirname( path )
				
p = Py4LO_helper()		
def __export_Py4LO_helper():
	return p