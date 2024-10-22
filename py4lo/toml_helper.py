import logging
import re
import subprocess   # nosec: B404
import sys
import traceback
from pathlib import Path
from typing import Dict, Any, Mapping, cast
import toml

from tools import nested_merge, secure_exe


def load_toml(default_py4lo_toml: Path, project_py4lo_toml: Path,
              kwargs: Mapping[str, Any]) -> Mapping[str, Any]:
    return TomlLoader(default_py4lo_toml, project_py4lo_toml, kwargs).load()


class TomlLoader:
    _logger = logging.getLogger(__name__)

    """Load a toml file and merge values with the default toml file"""

    def __init__(self, default_py4lo_toml: Path, project_py4lo_toml: Path,
                 kwargs: Mapping[str, Any]):
        self._kwargs = kwargs
        self._default_py4lo_toml = default_py4lo_toml
        self._project_py4lo_toml = project_py4lo_toml
        self._data = cast(Dict[str, Any], {})

    def load(self) -> Dict[str, Any]:
        TomlLoader._logger.debug(
            "Load TOML : %s (default=%s)", self._project_py4lo_toml,
            self._default_py4lo_toml)
        print("*** Load TOML : %s (default=%s)", self._project_py4lo_toml,
              self._default_py4lo_toml)
        self._load_toml(self._default_py4lo_toml)
        self._load_toml(self._project_py4lo_toml, skip_on_error=True)
        self._check_python_target_version()
        self._check_level()
        return self._data

    def _load_toml(self, path: Path, skip_on_error: bool = False):
        try:
            with path.open('r', encoding="utf-8") as s:
                content = s.read()
                data = toml.loads(content)
        except Exception as e:
            if not skip_on_error:
                print("Error when loading toml file {}: {}".format(path, e),
                      file=sys.stderr)
                traceback.print_exc(file=sys.stdout)
        else:
            self._data = nested_merge(self._data, data, self._apply)

    def _apply(self, v: Any) -> Any:
        try:
            return v.format(**self._kwargs)
        except (AttributeError, TypeError, ValueError):
            return v

    def _check_python_target_version(self):
        # get version from target executable
        if "python_exe" in self._data:
            python_exe = secure_exe(str(self._data["python_exe"]), "python")
            if python_exe is None:
                return
            status, out = subprocess.getstatusoutput(
                '"{}" -V'.format(python_exe)
            )
            if status == 0:
                m = re.match(r"^.* (3\.\d+)(\.\d+)?$", out)
                if m:
                    self._data["python_version"] = m.group(1)
                    return

        # if python_exe was not set, or did not return the expected result,
        # get from sys. It's the local python.
        if "python_version" not in self._data:
            self._data["python_exe"] = sys.executable
            self._data["python_version"] = "{}.{}".format(
                sys.version_info.major, sys.version_info.minor)

    def _check_level(self):
        if "log_level" not in self._data or self._data["log_level"] not in [
            "CRITICAL", "DEBUG", "ERROR", "FATAL",
            "INFO", "NOTSET", "WARN", "WARNING"
        ]:
            self._data["log_level"] = "INFO"
