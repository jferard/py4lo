#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2026 J. Férard <https://github.com/jferard>
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
import subprocess  # nosec: B404
from collections.abc import Callable
from pathlib import Path
from typing import Any


def open_with_libreoffice(ods_path: Path, lo_exe: str):
    """Open a file with LibreOffice"""
    sec_lo_exe = secure_exe(lo_exe, "soffice")
    if sec_lo_exe is None:
        return
    subprocess.call([sec_lo_exe, str(ods_path)])  # nosec: B603


def nested_merge(d1: dict[str, Any], d2: dict[str, Any],
                 apply: Callable[[Any], Any]) -> dict[str, Any]:
    """
    Merge two dicts.
    >>> nested_merge({'a': 1}, {'b': 2}, lambda x: x*2)
    {'a': 1, 'b': 4}

    @param d1: the first dict.
    @param d2: the second dict.
    @param apply: a function to apply to the dict
    @return: a merged dict
    """
    for k, v in d2.items():
        if isinstance(v, dict):
            d1[k] = nested_merge(d1.get(k, {}), v, apply)
        elif isinstance(v, list):
            d1[k] = [apply(w) for w in d1.get(k, []) + v]
        else:
            d1[k] = apply(v)

    return d1

def secure_exe(exe: str, name: str) -> str | None:
    import shutil
    ret = shutil.which(exe)
    if ret is not None and name in ret:
        return ret
    else:
        return None
