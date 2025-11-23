# -*- coding: utf-8 -*-
# Py4LO - Python Toolkit For LibreOffice Calc
#       Copyright (C) 2016-2025 J. Férard <https://github.com/jferard>
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
"""
The module py4lo_commons deals with ordinary Python objects (POPOs ?).

It allows to:
* use a Bus for events.
* get a logger
* get the current directory
"""
import configparser
import datetime as dt
import logging
import sys
from decimal import Decimal
# mypy: disable-error-code="import-untyped"
from pathlib import Path
from typing import (
    Union, Any, cast, List, Optional, TextIO, Iterable, Mapping, Callable,
    TypeVar)

from py4lo_typing import (
    UnoXScriptContext, StrPath, lazy)

try:
    # noinspection PyUnresolvedReferences
    import uno
except ImportError:
    from _mock_objects import uno

_xsc = None


def uno_url_to_path(url: str) -> Optional[Path]:
    """
    Wrapper for fileUrlToSystemPath.
    @param url: the url
    @return: the path or None if the url is empty
    """
    if url.strip():
        return Path(uno.fileUrlToSystemPath(url))
    else:
        return None


def uno_path_to_url(path: Union[str, Path]) -> str:
    """
    Wrapper for systemPathToFileUrl
    @param path: the path
    @return: the url
    """
    if isinstance(path, str):
        path = Path(path)
    try:
        path = path.resolve()
    except FileNotFoundError:  # 3.5 is strict
        pass

    return uno.systemPathToFileUrl(str(path))


def init(xsc: UnoXScriptContext):
    """
    Initialize the script context in main script.

    :param xsc: the value shoudl be XSCRIPTCONTEXT in main script.
    :deprecated: use `Commons.create(XSCRIPTCONTEXT)`
    """
    global _xsc
    _xsc = xsc


class Bus:
    """
    A minimal bus minimal to communicate with front end.

    Example:
    ```
    bus = Bus()

    class X:
        def _handle_func(x: Any):
            print(x)

    ...
    x = X()
    bus.subscribe(x)

    ...
    bus.post("func", "foo") # calls x._handle_func("foo")
    bus.post("other_func", "foo") # ignore
    """

    def __init__(self):
        self._subscribers = cast(List[Any], [])

    def subscribe(self, s: Any):
        """
        Subscribe to the bus.

        @param s: the object that subscribes
        """
        self._subscribers.append(s)

    def post(self, event_type: str, *event_data: Any):
        """
        Post an event. If a subscriber does not implement this event,
        the event is simply ignored.

        @param event_type: the name of the event
        @param event_data: the data
        """
        for s in self._subscribers:
            m_name = "_handle_" + event_type.lower()
            if hasattr(s, m_name):
                m = getattr(s, m_name)
                m(*event_data)


def create_bus() -> Bus:
    """
    Create a new bus
    @return: the bus
    """
    return Bus()


class Commons:
    """
    The Commons class implements some basic features.
    """
    _logger = logging.getLogger(__name__)

    @staticmethod
    def create(xsc: Optional[UnoXScriptContext] = None) -> "Commons":
        """
        Create a new Commons object.

        @param xsc: XSCRIPTCONTEXT
        @return:
        """
        if xsc is None:
            Commons._logger.warning(
                "Should call `Commons.create(XSCRIPTCONTEXT)`")
            # noinspection PyUnresolvedReferences
            xsc = _xsc
        if xsc is None:
            raise ValueError()
        doc = xsc.getDocument()
        return Commons(doc.URL)

    def __init__(self, url: str):
        """
        Create a new Commons object
        @param url: the url of the document.
        """
        self._url = url
        self._logger = lazy(logging.Logger)

    def __del__(self):
        """
        Flushes and closes all loggers.
        """
        if self._logger is not None:
            for h in self._logger.handlers:
                h.flush()
                h.close()

    def cur_dir(self) -> Optional[Path]:
        """
        @return: the directory of the current document or None (new document).
        """
        path = uno_url_to_path(self._url)
        if path is None:
            return None
        return path.parent

    def init_logger(
            self, file: Optional[Union[StrPath, TextIO]] = None, mode="a",
            level=logging.DEBUG,
            fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s'):
        """
        Initialize the root logger (but returns nothing)

        @param file: the log file
        @param mode: the log mode
        @param level: the log level
        @param fmt: the log entry fromat
        @raises: Exception if the logger is already initialized
        """
        if self._logger is not None:
            raise Exception("use init_logger ONCE")

        self._logger = self.get_logger(file, mode, level, fmt)

    def logger(self) -> logging.Logger:
        """
        @return: the logger.
        """
        if self._logger is None:
            raise Exception("use init_logger")
        return self._logger

    def get_logger(
            self, file: Optional[Union[StrPath, TextIO]] = None,
            mode: str = "a", level: int = logging.DEBUG,
            fmt: str = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    ) -> logging.Logger:
        """
        Initialize and return the root logger.

        ```
        logger = commons.get_logger(...)
        ```

        is the same as:
        ```
        commons.init_logger(...)
        logger = commons.logger()
        ```

        @param file: the log file
        @param mode: the log mode
        @param level: the log level
        @param fmt: the log entry fromat
        @return: the logger
        """
        if file is None:
            file = self.join_current_dir("py4lo.log")
        logger = logging.getLogger()
        init_logger(logger, file, mode, level, fmt)
        return logger

    def join_current_dir(self, filename: str) -> Path:
        """
        @param filename: a file name
        @return: the path of the file name in the doc current directory
        """
        cur_dir_path = self.cur_dir()
        if cur_dir_path is None:
            return Path(filename)
        else:
            return cur_dir_path / filename

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
        import zipfile
        import codecs

        if isinstance(filenames, (str, Path)):
            filenames = [filenames]

        config = _get_config(args)
        apply(config)
        reader = codecs.getreader(encoding)

        path = uno_url_to_path(self._url)
        str_path = str(path)  # py < 3.6.2
        with zipfile.ZipFile(str_path, 'r') as z:
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
        path = uno_url_to_path(self._url)
        str_path = str(path)  # py < 3.6.2
        with zipfile.ZipFile(str_path, 'r') as z:
            with z.open(filename) as f:
                return f.read()


def init_logger(
        logger: logging.Logger,
        file: Optional[Union[StrPath, TextIO]] = None,
        mode: str = "a", level: int = logging.DEBUG,
        fmt: str = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'):
    """
    Initialize a logger (but returns nothing)

    @param logger: the logger to initialize
    @param file: the log file
    @param mode: the log mode
    @param level: the log level
    @param fmt: the log entry fromat
    """
    if file is None:
        fh = _get_handler(sys.stdout, mode, level, fmt)
    else:
        fh = _get_handler(file, mode, level, fmt)
    logger.addHandler(fh)
    logger.setLevel(level)


def _get_handler(file: Union[StrPath, TextIO], mode: str, level: int,
                 fmt: str) -> logging.Handler:
    """
    Return a log handler.

    @param file: the log file
    @param mode: the log mode
    @param level: the log level
    @param fmt: the log entry fromat
    """
    if isinstance(file, Path):
        fh = logging.FileHandler(str(file), mode)
    elif isinstance(file, str):
        fh = logging.FileHandler(file, mode)
    else:
        fh = logging.StreamHandler(file)
    formatter = logging.Formatter(fmt)
    fh.setFormatter(formatter)
    fh.setLevel(level)
    return fh


def read_config(filenames: Union[StrPath, Iterable[StrPath]],
                args: Optional[Mapping] = None,
                apply: Callable[
                    [configparser.ConfigParser], None
                ] = lambda config: None,
                encoding: str = 'utf-8') -> configparser.ConfigParser:
    """
    Read a config. See https://docs.python.org/3.7/library/configparser.html

    :param filenames: one or many filenames
    :param args: arguments to be passed to the ConfigParser
    :param apply: function to modifiy the config
    :param encoding: the encoding of the file

    Example: `config = commons.read_config(
            commons.join_current_dir("pcrp.ini"))`
    """
    config = _get_config(args)
    apply(config)
    config.read(filenames, encoding)
    return config


def _get_config(args: Optional[Mapping]) -> configparser.ConfigParser:
    """
    @param args: arguments to be passed to the ConfigParser
    @return: the ConfigParser
    """
    if args is None:
        config = configparser.ConfigParser()
    else:
        config = configparser.ConfigParser(**args)
    return config


def sanitize(s: str) -> str:
    """
    Remove accents and special chars from a string.

    @param s: the unicode string
    @return: the ascii string
    """
    import unicodedata
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode(
        'ascii')
    return s


T = TypeVar("T")
U = TypeVar("U")


def secure_strip(v: T) -> Union[str, T]:
    return v.strip() if isinstance(v, str) else v


def create_parse_int_or(thousands_seps: str, default: T) -> Callable[
    [str], Union[int, T]]:
    trans_table = str.maketrans("", "", thousands_seps)

    def func(int_str: str) -> Union[int, T]:
        int_str = int_str.translate(trans_table)
        try:
            return int(int_str)
        except ValueError:
            return default

    return func


def create_parse_float_or(
        thousands_seps: str, decimal_seps: str, default: T
) -> Callable[[str], Union[float, T]]:
    return _create_parse_number_or(float, thousands_seps, decimal_seps, default)


def create_parse_decimal_or(
        thousands_seps: str, decimal_seps: str, default: T
) -> Callable[[str], Union[Decimal, T]]:
    return _create_parse_number_or(
        Decimal, thousands_seps, decimal_seps, default)


def _create_parse_number_or(
        target: Callable[[str], U], thousands_seps: str,
        decimal_seps: str, default: T
) -> Callable[[str], Union[U, T]]:
    th_trans_table = str.maketrans("", "", thousands_seps)
    dec_trans_table = str.maketrans(decimal_seps, "." * len(decimal_seps))

    def func(number_str: str) -> Union[U, T]:
        number_str = number_str.translate(th_trans_table)
        number_str = number_str.translate(dec_trans_table)
        try:
            return target(number_str)
        except (ValueError, ArithmeticError):
            return default

    return func


def create_parse_date_or(
        format_str: str, default: T) -> Callable[[str], Union[dt.date, T]]:
    def func(date_string: str) -> Union[dt.date, T]:
        try:
            d = dt.datetime.strptime(date_string, format_str)
        except ValueError:
            return default
        else:
            return d.date()

    return func


def create_parse_datetime_or(format_str: str, default: T) -> Union[dt.date, T]:
    def func(date_string: str) -> Union[dt.date, T]:
        try:
            return dt.datetime.strptime(date_string, format_str)
        except ValueError:
            return default

    return func


def create_parse_time_or(format_str: str, default: T) -> Union[dt.date, T]:
    def func(date_string: str) -> Union[dt.date, T]:
        try:
            d = dt.datetime.strptime(date_string, format_str)
        except ValueError:
            return default
        else:
            return d.time()

    return func


def create_format_int_or(thousands_sep: str, default: T) -> Callable[
    [Optional[int]], Union[str, T]]:
    if thousands_sep in (",", "_"):
        format_spec = thousands_sep + "d"

        def func(value: Optional[int]) -> Union[str, T]:
            if value is None:
                return default
            return format(value, format_spec)

    else:
        def func(value: Optional[int]) -> Union[str, T]:
            if value is None:
                return default
            return format(value, "_d").replace("_", thousands_sep)

    return func


def create_format_float_or(
        thousands_sep: str, decimal_sep: str,
        decimals: int, default: T
) -> Callable[[Optional[float]], Union[str, T]]:
    if thousands_sep == "":
        if decimal_sep == ".":
            if decimals < 0:
                apply = str
            else:
                format_spec = ".{}f".format(decimals)

                def apply(value: int) -> str:
                    return format(value, format_spec)
        else: # a decimal sep
            if decimals < 0:
                def apply(value: int) -> str:
                    return str(value).replace(".", decimal_sep)
            else:
                format_spec = ".{}f".format(decimals)

                def apply(value: int) -> str:
                    return format(value, format_spec).replace(".", decimal_sep)
    elif thousands_sep in (",", "_"):
        if decimal_sep == ".":
            if decimals < 0:
                format_spec = "{}f".format(thousands_sep)

                def apply(value: int) -> str:
                    return format(value, format_spec).rstrip("0")
            else:
                format_spec = "{}.{}f".format(thousands_sep, decimals)

                def apply(value: int) -> str:
                    return format(value, format_spec)
        else:
            if decimals < 0:
                format_spec = "{}f".format(thousands_sep)

                def apply(value: int) -> str:
                    return format(value, format_spec).rstrip("0").replace(".", decimal_sep)
            else:
                format_spec = "{}.{}f".format(thousands_sep, decimals)

                def apply(value: int) -> str:
                    return format(value, format_spec).replace(".", decimal_sep)
    else:
        if decimal_sep == ".":
            if decimals < 0:
                def apply(value: int) -> str:
                    return format(value, "_f").rstrip("0").replace("_", thousands_sep)
            else:
                format_spec = "_.{}f".format(decimals)

                def apply(value: int) -> str:
                    return format(value, format_spec).replace("_", thousands_sep)
        else:
            if decimals < 0:
                def apply(value: int) -> str:
                    return format(value, "_f").rstrip("0").replace("_", thousands_sep).replace(".", decimal_sep)
            else:
                format_spec = "_.{}f".format(decimals)

                def apply(value: int) -> str:
                    return format(value, format_spec).replace("_", thousands_sep).replace(".", decimal_sep)

    def func(value: Optional[int]) -> Union[str, T]:
        if value is None:
            return default
        return apply(value)

    return func


DECIMAL_ZERO = Decimal(0)


def create_format_decimal_or(
        thousands_sep: str, decimal_sep: str,
        decimals: int, default: T
) -> Callable[[Optional[Decimal]], Union[str, T]]:
    if thousands_sep in (",", "_"):
        format_spec = thousands_sep + "d"

        def format_int(value: int) -> str:
            return format(value, format_spec)

    else:
        def format_int(value: int) -> str:
            return format(value, "_d").replace("_", thousands_sep)

    if decimals >= 0:
        def func(value: Optional[Decimal]) -> Union[str, T]:
            if value is None:
                return default

            int_value = int(value // 1)
            dec_value = abs(value % 1)
            dec_value = round(dec_value * 10 ** decimals)
            return "{}{}{}".format(
                format_int(int_value), decimal_sep, dec_value)
    else:
        def func(value: Optional[Decimal]) -> Union[str, T]:
            if value is None:
                return default

            int_value = int(value // 1)
            dec_value = abs(value % 1)
            if dec_value == DECIMAL_ZERO:
                return "{}{}0".format(
                    format_int(int_value), decimal_sep)
            else:
                return "{}{}{}".format(
                    format_int(int_value), decimal_sep, str(dec_value)[2:])

    return func


def create_format_date_or(
        format_str: str, default: T
) -> Callable[[Optional[dt.date]], Union[str, T]]:
    def func(value: Optional[dt.date]) -> Union[str, T]:
        if value is None:
            return default
        return value.strftime(format_str)

    return func


def create_format_datetime_or(
        format_str: str, default: T
) -> Callable[[Optional[dt.datetime]], Union[str, T]]:
    def func(value: Optional[dt.datetime]) -> Union[str, T]:
        if value is None:
            return default
        return value.strftime(format_str)

    return func


def create_format_time_or(
        format_str: str, default: T
) -> Callable[[Optional[dt.time]], Union[str, T]]:
    def func(value: Optional[dt.time]) -> Union[str, T]:
        if value is None:
            return default
        return value.strftime(format_str)

    return func
