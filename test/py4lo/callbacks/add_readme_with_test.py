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
import unittest
import env
from callbacks import *
import io
import zipfile

class TestAddDebugContent(unittest.TestCase):
    def test_add_debug_content_empty(self):
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, 'w')
        AddReadmeWith(env.inc_dir, "contact").call(zout)
        print (zout.read("Basic/Standard/py4lo.xml"))
        self.assertEquals(b'<?xml version="1.0" encoding="UTF-8"?>\r\n<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">\r\n<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">\r\n    <library:library library:name="Standard" library:link="false"/>\r\n</library:libraries>\r\n', zout.read("Basic/script-lc.xml"))
        self.assertEquals(b'<?xml version="1.0" encoding="UTF-8"?>\r\n<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\r\n<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">\r\n    <library:element library:name="py4lo"/>\r\n</library:library>\r\n', zout.read("Basic/Standard/script-lb.xml"))
        self.assertEquals(b'<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">REM  *****  BASIC  *****\nREM The scripts are written in Python with an external text editor, and included in this file with py4lo.\nREM Py4LO - Python Toolkit For LibreOffice Calc\nREM Copyright (C) 2016-2018 Julien F\xc3\xa9rard &lt;https://github.com/jferard&gt;\nREM See https://github.com/jferard/py4lo/README.md\n\nSub Readme\n\tMsgBox &quot;The scripts are written in Python with an external text editor, and included in this file with py4lo.&quot; &amp; chr(13) _\n\t&amp; &quot;Py4LO - Python Toolkit For LibreOffice Calc&quot; &amp; chr(13) _\n\t&amp; &quot;Copyright (C) 2016-2018 Julien F\xc3\xa9rard &lt;https://github.com/jferard&gt;&quot; &amp; chr(13) _\n\t&amp; &quot;See https://github.com/jferard/py4lo/README.md&quot; &amp; chr(13) _\n\t&amp; &quot;Contact : contact&quot;, IDOK, &quot;py4lo&quot;\nEnd Sub\n</script:module>\n', zout.read("Basic/Standard/py4lo.xml"))

if __name__ == '__main__':
    unittest.main()
