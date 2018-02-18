# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2018 J. Férard <https://github.com/jferard>

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
import unohelper
import uno
import os
import unicodedata
import logging

def init(xsc):
    Commons.xsc = xsc

class Bus(unohelper.Base):
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

class Commons(unohelper.Base):
    def __init__(self):
        self.doc = Commons.xsc.getDocument()
        self.__logger = None

    def __del__(self):
        if self.__logger is not None:
            for h in self.__logger.handlers:
                h.flush()
                h.close()

    def create_bus(self):
        return Bus()

    def cur_dir(self):
        """return the directory of the current document"""
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

    def init_logger(self, path="py4lo.log", mode="a", level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'):
        self.__logger = self.get_logger(path, mode, level)
        return self.__logger

    def get_logger(self, path="py4lo.log", mode="a", level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'):
        fh = logging.FileHandler(path, mode)
        formatter = logging.Formatter(format)
        fh.setFormatter(formatter)
        fh.setLevel(level)

        logger = logging.getLogger()
        logger.addHandler(fh)
        logger.setLevel(level)
        return logger

    def logger(self):
        if self.__logger is None:
            self.__logger = self.get_logger()
        return self.__logger
