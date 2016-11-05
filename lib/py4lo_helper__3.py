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
import uno
from com.sun.star.uno import RuntimeException
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK

class Py4LO_helper():
	def __init__(self):
		self.doc = XSCRIPTCONTEXT.getDocument()						#Document
		self.ctrl = self.doc.CurrentController						#Document
		self.frame = self.ctrl.Frame
		self.parent_win = self.frame.ContainerWindow
		self.ctxt = uno.getComponentContext()						#Contexte
		self.sm = self.ctxt.getServiceManager()							#ServiceManager
		mspf = self.sm.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", self.ctxt)
		self.msp = mspf.createScriptProvider("")
		self.dsp = self.doc.getScriptProvider()
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
			l.append(make_pv(k, d[k]))
		return tuple(l)

	# from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603# (thanks Bernard !)
	def message_box(self, parent_win, msg_text, msg_title, msg_type=MESSAGEBOX, msg_buttons=BUTTONS_OK):
		sv = self.sm.createInstanceWithContext("com.sun.star.awt.Toolkit", self.ctxt)
		my_box = sv.createMessageBox(parent_win, msg_type, msg_buttons, msg_title, msg_text)
		return my_box.execute()

def __export_Py4LO_helper():
	return Py4LO_helper()