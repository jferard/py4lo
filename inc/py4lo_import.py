# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>

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
# mypy: disable-error-code="import-untyped"
# noinspection PyBroadException
try:
    # noinspection PyUnresolvedReferences
    XSCRIPTCONTEXT  # type: ignore[name-defined] # noqa: F821
except:  # noqa: E722
    pass
else:
    import uno
    import sys

    # add path/to/doc.os/Scripts/python to sys.path, to import Python
    # modules (*.py, *.py[co]) and packages from a ZIP-format archive.
    doc = XSCRIPTCONTEXT.getDocument()  # type: ignore[name-defined] # noqa: F821
    spath = uno.fileUrlToSystemPath(doc.URL + "/Scripts/python")
    if spath not in sys.path:
        sys.path.insert(0, spath)
