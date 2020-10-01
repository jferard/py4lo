# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2020 J. Férard <https://github.com/jferard>

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
import subprocess
from pathlib import Path
from typing import List, Dict, Any, Set, Callable

def open_with_calc(ods_path: Path, calc_exe: str):
    """Open a file with calc"""
    _r = subprocess.call([calc_exe, str(ods_path)])


def nested_merge(d1: Dict[str, Any], d2: Dict[str, Any],
                 apply: Callable[[Any], Any]) -> Dict[str, Any]:
    for k, v in d2.items():
        if isinstance(v, Dict):
            d1[k] = nested_merge(d1.get(k, {}), v, apply)
        elif isinstance(v, List):
            d1[k] = [apply(w) for w in d1.get(k, []) + v]
        else:
            d1[k] = apply(v)

    return d1
