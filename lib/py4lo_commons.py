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
import unicodedata
   
class Bus:
    """A minimal bus minimal to communicate with front end"""
    def __init__(self):
        self.__subscribers = []
        
    def subscribe(self, s):
        self.__subscribers.append(s)
        
    def post(self, event_type, event_data):
        for s in self.__subscribers:
            m_name = "_handle_"+event_type.lower()
            if hasattr(s, m_name):
                m = getattr(s, m_name)
                m(event_data)

# py4lo: if python_version >= 2.6
Class PMLogger:
	def __init__(self, fpath):
# py4lo: if python_version < 3.0
		import io
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
				
class Commons:
    def create_bus(self):
        return Bus()
        
	def cur_dir(self):
		url = self.doc.getURL()
		path = uno.fileUrlToSystemPath( url )
		return os.path.dirname( path )
            
    def sanitize(self, s):
		import unicodedata
        try:
            s = unicodedata.normalize('NFKD', s).encode('ascii','ignore')
        except Exception as e:
            pass
        return s

	def logger(self):
		return PMLogger()
        
# export        
c = Commons()
def __export_commons():
    return c