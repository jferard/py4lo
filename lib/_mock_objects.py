#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
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
import os
from pathlib import Path
from typing import Any
from unittest import mock
from urllib.parse import urlparse


##############################
# # DO NOT EMBED THIS FILE # #
##############################

class uno:
    @staticmethod
    def fileUrlToSystemPath(url: str) -> str:
        result = urlparse(url)
        if result.netloc:
            return os.path.join(result.netloc, result.path.lstrip("/"))
        else:
            return result.path

    @staticmethod
    def systemPathToFileUrl(path: str) -> str:
        return Path(path).as_uri()

    @staticmethod
    def createUnoStruct(struct_id: str) -> mock.Mock:
        m = mock.Mock(typeName=struct_id)
        return m

    @staticmethod
    def Any(name: str, value: Any) -> Any:
        return value


class unohelper:
    class Base:
        pass

    @staticmethod
    def ImplementationHelper():
        class C:
            @staticmethod
            def addImplementation(*args): return None

        return C


# Listeners

class XEventListener:
    pass


class XActionListener:
    pass


class XActivateListener:
    pass


class XAdjustmentListener:
    pass


class XDockableWindowListener:
    pass


class XFocusListener:
    pass


class XItemListener:
    pass


class XItemListListener:
    pass


class XKeyListener:
    pass


class XMenuListener:
    pass


class XMouseListener:
    pass


class XMouseMotionListener:
    pass


class XPaintListener:
    pass


class XSpinListener:
    pass


class XStyleChangeListener:
    pass


class XTabListener:
    pass


class XTextListener:
    pass


class XTopWindowListener:
    pass


class XVclContainerListener:
    pass


class XWindowListener:
    pass


class XWindowListener2:
    pass


# Exceptions

class UnoException(Exception):
    pass


class IOException(Exception):
    pass


class ScriptFrameworkErrorException(UnoException):
    pass


class UnoRuntimeException(UnoException):
    pass


Locale = object

class XTransferable:
    pass



