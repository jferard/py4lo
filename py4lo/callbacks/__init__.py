# -*- coding: utf-8 -*-
#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2024 J. Férard <https://github.com/jferard>
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

"""
These are callbacks for the ZipUpdater
"""
from callbacks.callback import (  # noqa
    BeforeCallback, ItemCallback, AfterCallback)
from callbacks.add_readme_with import AddReadmeWith  # noqa
from callbacks.add_assets import AddAssets  # noqa
from callbacks.add_debug_content import AddDebugContent  # noqa
from callbacks.add_scripts import AddScripts, ARC_SCRIPTS_PATH  # noqa
from callbacks.ignore_item import IgnoreItem  # noqa
from callbacks.rewrite_manifest import RewriteManifest  # noqa
