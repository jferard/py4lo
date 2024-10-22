#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2023 J. Férard <https://github.com/jferard>
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
from logging import Logger
from typing import List, Dict, Union, TypeVar, cast, Tuple

from core.source_dest import Sources
from directives.directive import Directive
from directives.embed_script import EmbedScript
from directives.embed_lib import EmbedLib
from directives.include import Include
from directives.entry import Entry

GET_DIRECTIVE = "@"
U = TypeVar('U', bound=Directive)
DirectiveTree = Dict[str, Union["DirectiveTree", Directive]]

class _DirectiveProviderFactory:
    def __init__(self, logger: Logger,
                 *directives: Directive):
        self._logger = logger
        self._directives = directives
        self._directives_tree = cast(DirectiveTree, {})

    def create_provider(self) -> "DirectiveProvider":
        """
        Create a new provider. Build the tree (embed lib & embed script ->
        embed -> (lib, script))
        @return: the provider
        """
        for directive in self._directives:
            sig_elements = directive.sig_elements()
            assert len(sig_elements)  # nosec: B101

            self._insert_directive_class(sig_elements, directive)

        self._logger.debug("Directives tree: %s", self._directives_tree)
        return DirectiveProvider(self._logger, self._directives_tree)

    def _insert_directive_class(self, sig_elements: List[str],
                                directive: Directive):
        cur_directives_tree = self._directives_tree
        for fst in sig_elements:
            if fst not in cur_directives_tree:
                cur_directives_tree[fst] = {}
            next_tree = cast(DirectiveTree, cur_directives_tree[fst])
            cur_directives_tree = next_tree

        cur_directives_tree.update({GET_DIRECTIVE: directive})


class DirectiveProvider:
    """
    Takes a full directive line (as shlex args) and splits between the
    directive id and the args. Then return the directive and the args.
    """

    @staticmethod
    def create(logger: Logger, sources: Sources):
        entry = Entry(sources.inc_dir, sources.get_module_names())
        return _DirectiveProviderFactory(
            logger, entry,
            Include(sources.inc_dir),
            EmbedLib(sources.lib_dir),
            EmbedScript(sources.opt_dir)).create_provider()

    def __init__(self, logger: Logger, directives_tree: DirectiveTree):
        self._logger = logger
        self._directives_tree = directives_tree

    def get(self, args: List[str]) -> Tuple[Directive, List[str]]:
        """
        args are the shlex result
        @return the directive + the list of args.
        """
        self._logger.debug("Lookup directive: %s", args)
        cur_directives_tree = self._directives_tree

        for i in range(len(args)):
            arg = args[i]
            if arg in cur_directives_tree:
                next_tree = cast(DirectiveTree, cur_directives_tree[arg])
                cur_directives_tree = next_tree
            elif GET_DIRECTIVE in cur_directives_tree:
                directive = cast(Directive, cur_directives_tree[GET_DIRECTIVE])
                return directive, args[i:]
            else:
                raise KeyError(args)

        if GET_DIRECTIVE in cur_directives_tree:
            return (
                cast(Directive, cur_directives_tree[GET_DIRECTIVE]),
                list(),
            )

        raise AssertionError("no args")

