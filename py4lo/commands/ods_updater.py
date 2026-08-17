import logging

from core.asset import DestinationAsset
from core.script import DestinationScript, TempScript
from core.source_dest import Destinations, Sources
from directives import DirectiveProvider
from script_set_processor import ScriptSetProcessor


class OdsUpdaterHelper:
    def __init__(self, logger: logging.Logger, sources: Sources,
                 destinations: Destinations,
                 python_version: str):
        self._logger = logger
        self._sources = sources
        self._destinations = destinations
        self._python_version = python_version

    def get_assets(self) -> list[DestinationAsset]:
        source_assets = self._sources.get_assets()
        return self._destinations.to_destination_assets(
            source_assets)

    def get_destination_scripts(self) -> list[DestinationScript]:
        temp_scripts = self.get_temp_scripts()
        return self._destinations.to_destination_scripts(temp_scripts)

    def get_temp_scripts(self) -> list[TempScript]:
        source_scripts = self._sources.get_src_scripts()
        directive_provider = DirectiveProvider.create(self._logger,
                                                      self._sources)
        return ScriptSetProcessor(self._logger, self._destinations.temp_dir,
                                  self._python_version, directive_provider,
                                  source_scripts).process()
