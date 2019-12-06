import logging
from typing import List

from core.asset import DestinationAsset
from core.properties import Destinations, Sources
from directives import DirectiveProvider
from script_set_processor import ScriptSetProcessor
from core.script import TempScript, SourceScript
from tools import get_paths


class OdsUpdaterHelper:
    def __init__(self, logger: logging.Logger, sources: Sources,
                 destinations: Destinations,
                 python_version: str):
        self._logger = logger
        self._sources = sources
        self._destinations = destinations
        self._python_version = python_version

    def get_assets(self) -> List[DestinationAsset]:
        assets_dest_dir = self._destinations.assets_dest_dir
        return [sa.to_dest(assets_dest_dir) for sa in
                self._sources.get_assets()]

    def get_temp_scripts(self) -> List[TempScript]:
        script_paths = get_paths(self._sources.src_dir,
                                 self._sources.src_ignore, "*.py")
        source_scripts = [SourceScript(sp, self._sources.src_dir) for sp in
                          script_paths]
        directive_provider = DirectiveProvider.create(self._logger,
                                                      self._sources)
        return ScriptSetProcessor(self._logger, self._destinations.temp_dir,
                                  self._python_version, directive_provider,
                                  source_scripts).process()
