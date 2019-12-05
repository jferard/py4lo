import subprocess
import sys
from pathlib import Path
from typing import Dict, Any, Mapping
import toml

from tools import nested_merge


def load_toml(default_py4lo_toml: Path, project_py4lo_toml: Path,
              kwargs: Mapping[str, Any]) -> Mapping[str, Any]:
    return TomlLoader(default_py4lo_toml, project_py4lo_toml, kwargs).load()


class TomlLoader:
    """Load a toml file and merge values with the default toml file"""

    def __init__(self, default_py4lo_toml: Path, project_py4lo_toml: Path,
                 kwargs: Mapping[str, Any]):
        self._kwargs = kwargs
        self._default_py4lo_toml = default_py4lo_toml
        self._project_py4lo_toml = project_py4lo_toml
        self._data: Dict[str, object] = {}

    def load(self) -> Dict[str, object]:
        self._load_toml(self._default_py4lo_toml)
        self._load_toml(self._project_py4lo_toml)
        self._check_python_target_version()
        self._check_level()
        return self._data

    def _load_toml(self, path: Path):
        try:
            with path.open('r', encoding="utf-8") as s:
                content = s.read()
                data = toml.loads(content)
        except Exception as e:
            print("Error when loading toml file {}: {}".format(path, e))
        else:
            self._data = nested_merge(self._data, data, self._apply)

    def _apply(self, v: Any) -> Any:
        try:
            return v.format(**self._kwargs)
        except:
            return v

    def _check_python_target_version(self):
        # get version from target executable
        if "python_exe" in self._data:
            status, version = subprocess.getstatusoutput(
                "\"" + str(self._data["python_exe"]) + "\" -V")
            if status == 0:
                self._data["python_version"] = \
                    ((version.split())[1].split("."))[0]
                return

        # if python_exe was not set, or did not return the expected result,
        # get from sys. It's the local python.
        if "python_version" not in self._data:
            self._data["python_exe"] = sys.executable
            self._data["python_version"] = str(
                sys.version_info.major) + "." + str(sys.version_info.minor)

    def _check_level(self):
        if "log_level" not in self._data or self._data["log_level"] not in [
            "CRITICAL", "DEBUG", "ERROR", "FATAL",
            "INFO", "NOTSET", "WARN", "WARNING"]:
            self._data["log_level"] = "INFO"
