from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional


@dataclass(eq=True, frozen=True)
class SourceScript:
    script_path: Path
    source_dir: Path
    export_funcs: bool

    @property
    def relative_path(self):
        return self.script_path.relative_to(self.source_dir)


@dataclass(eq=True, frozen=True)
class ParsedScriptContent:
    text: str
    exported_func_names: List[str]


@dataclass(eq=True, frozen=True)
class DestinationScript:
    """A target script has a file name, a content (the file was processed),
    some exported functions and the exceptions encountered by the
    interpreter"""

    script_path: Path
    script_content: bytes
    dest_dir: Path
    exported_func_names: List[str]
    exception: Optional[Exception]

    @property
    def relative_path(self):
        return self.script_path.relative_to(self.dest_dir)

    @property
    def dest_name(self) -> str:
        script_path = self.relative_path
        return ".".join(script_path.parts[:-1] + (script_path.stem,))


@dataclass(eq=True, frozen=True)
class TempScript:
    """A target script has a file name, a content (the file was processed),
    some exported functions and the exceptions encountered by the
    interpreter"""

    script_path: Path
    script_content: bytes
    temp_dir: Path
    exported_func_names: List[str]
    exception: Optional[Exception]

    @property
    def relative_path(self):
        return self.script_path.relative_to(self.temp_dir)

    def to_destination(self, dest_dir: Path) -> DestinationScript:
        script_path = dest_dir.joinpath(self.relative_path)
        return DestinationScript(
            script_path,
            self.script_content,
            dest_dir,
            self.exported_func_names,
            self.exception,
        )
