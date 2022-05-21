# -*- coding: utf-8 -*-
# Py4LO - Python Toolkit For LibreOffice Calc
#       Copyright (C) 2016-2022 J. Férard <https://github.com/jferard>
#
#    This file is part of Py4LO.
#
#    Py4LO is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    THIS FILE IS SUBJECT TO THE "CLASSPATH" EXCEPTION.
#
#    Py4LO is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
"""py4lo_commons deals with ordinary Python objects (POPOs ?)."""
from pathlib import Path
import logging
import configparser
import datetime as dt
from typing import Union, Any, cast, List, Optional, TextIO, Iterable, Mapping, \
    Callable

from py4lo_typing import UnoXScriptContext

StrPath = Union[str, Path]

try:
    # noinspection PyUnresolvedReferences
    import uno
except ModuleNotFoundError:
    import os
    from urllib.parse import urlparse


    class uno:
        @staticmethod
        def fileUrlToSystemPath(url: str) -> str:
            result = urlparse(url)
            if result.netloc:
                return os.path.join(result.netloc, result.path.lstrip("/"))
            else:
                return result.path

ORIGIN = dt.datetime(1899, 12, 30)


def init(xsc):
    Commons.xsc = xsc


class Bus:
    """
    A minimal bus minimal to communicate with front end
    """

    def __init__(self):
        self._subscribers = cast(List[Any], [])

    def subscribe(self, s: Any):
        self._subscribers.append(s)

    def post(self, event_type: str, event_data: Any):
        for s in self._subscribers:
            m_name = "_handle_" + event_type.lower()
            if hasattr(s, m_name):
                m = getattr(s, m_name)
                m(event_data)


class Commons:
    @staticmethod
    def create(xsc: Optional[UnoXScriptContext] = None):
        if xsc is None:
            xsc = Commons.xsc
        doc = xsc.getDocument()
        return Commons(doc.URL)

    def __init__(self, url: str):
        self._url = url
        self._logger = None

    def __del__(self):
        if self._logger is not None:
            for h in self._logger.handlers:
                h.flush()
                h.close()

    def cur_dir(self) -> Path:
        """return the directory of the current document"""
        system_path = uno.fileUrlToSystemPath(self._url)
        path = Path(system_path)
        return path.parent

    def init_logger(
            self, file=None, mode="a", level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'):
        if self._logger is not None:
            raise Exception("use init_logger ONCE")

        self._logger = self.get_logger(file, mode, level, format)

    def get_logger(
            self, file: Optional[Union[StrPath, TextIO]] = None,
            mode: str = "a", level: int = logging.DEBUG,
            fmt: str = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    ) -> logging.Logger:
        if file is None:
            file = self.join_current_dir("py4lo.log")

        fh = Commons._get_handler(file, mode, level, fmt)

        logger = logging.getLogger()
        logger.addHandler(fh)
        logger.setLevel(level)
        return logger

    @staticmethod
    def _get_handler(file: Union[StrPath, TextIO], mode: str, level: int,
                     fmt: str):
        if isinstance(file, StrPath):
            fh = logging.FileHandler(file, mode)
        else:
            fh = logging.StreamHandler(file)
        formatter = logging.Formatter(fmt)
        fh.setFormatter(formatter)
        fh.setLevel(level)
        return fh

    def logger(self) -> logging.Logger:
        if self._logger is None:
            raise Exception("use init_logger")
        return self._logger

    def join_current_dir(self, filename: str) -> Path:
        return self.cur_dir() / filename

    def read_internal_config(self, filenames: Union[
        StrPath, Iterable[StrPath]], args=None,
                             apply=lambda config: None, encoding='utf-8'):
        """
        Read an internal config, in the assets directory of the document.
        See https://docs.python.org/3.7/library/configparser.html

        :param filenames: one filename or a list of filenames. Full path
                          *inside* the archive
        :param args: arguments to be passed to the ConfigParser
        :param apply: function to modify the config
        :param encoding: the encoding of the file

        Example: `config = commons.read_config("test.ini")`"""
        import configparser
        import zipfile
        import codecs

        if isinstance(filenames, StrPath):
            filenames = [filenames]

        config = _get_config(args)
        apply(config)
        reader = codecs.getreader(encoding)

        with zipfile.ZipFile(uno.fileUrlToSystemPath(self._url),
                             'r') as z:
            for filename in filenames:
                try:
                    file = z.open(filename)
                    config.read_file(reader(file))
                except KeyError:  # ignore non existing files
                    pass

        return config

    def get_asset(self, filename: str) -> bytes:
        """
        Read the content of an asset. To convert it to a reader, use :

            bs = commons.get_asset("foo.csv")
            reader = TextIOWrapper(BytesIO(bs), <encoding>))

        @param filename: name of the asset
        @return: file content as bytes
        """
        import zipfile
        with zipfile.ZipFile(uno.fileUrlToSystemPath(self._url),
                             'r') as z:
            with z.open(filename) as f:
                return f.read()


def create_bus() -> Bus:
    return Bus()


def read_config(filenames: Union[StrPath, Iterable[StrPath]],
                args: Optional[Mapping] = None,
                apply: Callable[[configparser.ConfigParser], None] = lambda
                        config: None,
                encoding: str='utf-8') -> configparser.ConfigParser:
    """
    Read a config. See https://docs.python.org/3.7/library/configparser.html

    :param filenames: one or many filenames
    :param args: arguments to be passed to the ConfigParser
    :param apply: function to modifiy the config
    :param encoding: the encoding of the file

    Example: `config = commons.read_config(commons.join_current_dir("pcrp.ini"))`
    """
    config = _get_config(args)
    apply(config)
    config.read(filenames, encoding)
    return config


def _get_config(args: Optional[Mapping]) -> configparser.ConfigParser:
    if args is None:
        config = configparser.ConfigParser()
    else:
        config = configparser.ConfigParser(**args)
    return config


def sanitize(s: str) -> str:
    """
    Remove accents and special chars from a string

    @param s: the unicode string
    @return: the ascii string
    """
    import unicodedata
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode(
        'ascii')
    return s


def date_to_int(a_date: Union[dt.date, dt.datetime]) -> int:
    if isinstance(a_date, dt.date):
        a_date = dt.datetime(a_date.year, a_date.month, a_date.day)
    return (a_date - ORIGIN).days


def date_to_float(a_date: Union[dt.date, dt.datetime, dt.time]) -> float:
    if not isinstance(a_date, dt.datetime):
        if isinstance(a_date, dt.date):
            a_date = dt.datetime(a_date.year, a_date.month, a_date.day)
        elif isinstance(a_date, dt.time):
            a_date = dt.datetime(1899, 12, 30, a_date.hour, a_date.minute,
                                 a_date.second, a_date.microsecond)

    return (a_date - ORIGIN).total_seconds() / 86400


def int_to_date(days: int) -> dt.datetime:
    return ORIGIN + dt.timedelta(days)


def float_to_date(days: float) -> dt.datetime:
    return ORIGIN + dt.timedelta(days)
