# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
#
#     This file is part of Py4LO.
#
#     Py4LO is free software: you can redistribute it and/or modify
#     it under the terms of the GNU General Public License as published by
#     the Free Software Foundation, either version 3 of the License, or
#     (at your option) any later version.
#
#     Py4LO is distributed in the hope that it will be useful,
#     but WITHOUT ANY WARRANTY; without even the implied warranty of
#     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#     GNU General Public License for more details.
#
#     You should have received a copy of the GNU General Public License
#     along with this program.  If not, see <http://www.gnu.org/licenses/>.
import unittest
import io
import zipfile

from callbacks import AddReadmeWith
from test import test_helper


class TestAddReadmeWith(unittest.TestCase):
    def setUp(self):
        self.maxDiff = None

    def test_add_readme_with(self):
        out = io.BytesIO()
        zout = zipfile.ZipFile(out, "w")
        print(test_helper.inc_dir, type(test_helper.inc_dir))
        AddReadmeWith(test_helper.inc_dir, "contact").call(zout)
        expected = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">
    <library:library library:name="Standard" library:link="false"/>
</library:libraries>
"""  # noqa: E501
        self.assertEqual(
            expected, zout.read("Basic/script-lc.xml").decode("utf-8")
        )
        expected = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">
    <library:element library:name="py4lo"/>
</library:library>
"""  # noqa: E501
        self.assertEqual(
            expected, zout.read("Basic/Standard/script-lb.xml").decode("utf-8")
        )
        expected = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">REM  *****  BASIC  *****
REM The scripts are written in Python with an external text editor, and included in this file with py4lo.\nREM Py4LO - Python Toolkit For LibreOffice Calc
REM Copyright (C) 2016-2018 Julien Férard &lt;https://github.com/jferard&gt;\nREM See https://github.com/jferard/py4lo/README.md

Sub Readme
    MsgBox &quot;The scripts are written in Python with an external text editor, and included in this file with py4lo.&quot; &amp; chr(13) _
    &amp; &quot;Py4LO - Python Toolkit For LibreOffice Calc&quot; &amp; chr(13) _
    &amp; &quot;Copyright (C) 2016-2018 Julien Férard &lt;https://github.com/jferard&gt;&quot; &amp; chr(13) _
    &amp; &quot;See https://github.com/jferard/py4lo/README.md&quot; &amp; chr(13) _
    &amp; &quot;Contact : contact&quot;, IDOK, &quot;py4lo&quot;\nEnd Sub\n</script:module>
"""  # noqa: E501
        self.assertEqual(
            expected, zout.read("Basic/Standard/py4lo.xml").decode("utf-8")
        )


if __name__ == "__main__":
    unittest.main()
