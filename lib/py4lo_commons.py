# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2020 J. Férard <https://github.com/jferard>

   This file is part of Py4LO.

   Py4LO is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   THIS FILE IS SUBJECT TO THE "CLASSPATH" EXCEPTION.

   Py4LO is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with this program.  If not, see <http://www.gnu.org/licenses/>."""

"""
py4lo_commons deals with ordinary Python objects (POPOs ?).

"""
try:
    import unohelper
    import uno
except ModuleNotFoundError:
    class unohelper:
        Base = object

import datetime
import logging
import os

import uno

ORIGIN = datetime.datetime(1899, 12, 30)


def init(xsc):
    Commons.xsc = xsc


class Bus(unohelper.Base):
    """
    A minimal bus minimal to communicate with front end
    """

    def __init__(self):
        self._subscribers = []

    def subscribe(self, s):
        self._subscribers.append(s)

    def post(self, event_type, event_data):
        for s in self._subscribers:
            m_name = "_handle_" + event_type.lower()
            if hasattr(s, m_name):
                m = getattr(s, m_name)
                m(event_data)


class Commons(unohelper.Base):
    @staticmethod
    def create(xsc=None):
        if xsc is None:
            xsc = Commons.xsc
        doc = xsc.getDocument()
        return Commons(doc.URL)

    def __init__(self, url):
        self._url = url
        self._logger = None

    def __del__(self):
        if self._logger is not None:
            for h in self._logger.handlers:
                h.flush()
                h.close()

    def cur_dir(self):
        """return the directory of the current document"""
        path = uno.fileUrlToSystemPath(self._url)
        return os.path.dirname(path)

    def init_logger(self, file=None, mode="a", level=logging.DEBUG,
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'):
        if self._logger is not None:
            raise Exception("use init_logger ONCE")

        self._logger = self.get_logger(file, mode, level, format)

    def get_logger(self, file=None, mode="a", level=logging.DEBUG,
                   format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'):
        if file is None:
            file = self.join_current_dir("py4lo.log")

        fh = Commons._get_handler(file, mode, level, format)

        logger = logging.getLogger()
        logger.addHandler(fh)
        logger.setLevel(level)
        return logger

    @staticmethod
    def _get_handler(file, mode, level, format):
        if type(file) == str:
            fh = logging.FileHandler(file, mode)
        else:
            fh = logging.StreamHandler(file)
        formatter = logging.Formatter(format)
        fh.setFormatter(formatter)
        fh.setLevel(level)
        return fh

    def logger(self):
        if self._logger is None:
            raise Exception("use init_logger")
        return self._logger

    def join_current_dir(self, filename):
        return os.path.join(unohelper.fileUrlToSystemPath(self._url), "..",
                            filename)

    def read_internal_config(self, filenames, args={},
                             apply=lambda config: None, encoding='utf-8'):
        """
        Read an internal config, in the assets directory of the document.
        See https://docs.python.org/3.7/library/configparser.html

        :param filenames: one filename or a list of filenames. Full path
                          *inside* the archive
        :param args: arguments to be passed to the ConfigParser
        :param apply: function to modify the config
        :param encoding: the encoding of the file
        :param assets_dir: dir of the assets in the zip file

        Example: `config = commons.read_config("pcrp.ini")`"""
        import configparser
        import zipfile
        import codecs

        if isinstance(filenames, (str, bytes)):
            filenames = [filenames]

        config = configparser.ConfigParser(**args)
        apply(config)
        reader = codecs.getreader(encoding)

        with zipfile.ZipFile(unohelper.fileUrlToSystemPath(self._url),
                             'r') as z:
            for filename in filenames:
                try:
                    file = z.open(filename)
                    config.read_file(reader(file))
                except KeyError:  # ignore non existing files
                    pass

        return config

    def get_asset(self, filename):
        """
        Read the content of an asset. To convert it to a reader, use :

            bs = commons.get_asset("foo.csv")
            reader = TextIOWrapper(BytesIO(bs), <encoding>))

        @param filename: name of the asset
        @return: file content as bytes
        """
        import zipfile
        with zipfile.ZipFile(unohelper.fileUrlToSystemPath(self._url),
                             'r') as z:
            with z.open(filename) as f:
                return f.read()


def create_bus():
    return Bus()


def read_config(filenames, args={}, apply=lambda config: None,
                encoding='utf-8'):
    """
    Read a config. See https://docs.python.org/3.7/library/configparser.html

    :param filenames: one or many filenames
    :param args: arguments to be passed to the ConfigParser
    :param apply: function to modifiy the config
    :param encoding: the encoding of the file

    Example: `config = commons.read_config(commons.join_current_dir("pcrp.ini"))`
    """
    import configparser
    config = configparser.ConfigParser(**args)
    apply(config)
    config.read(filenames, encoding)
    return config


def sanitize(s):
    """
    Remove accents and special chars from a string

    @param s: the unicode string
    @return: the ascii string
    """
    import unicodedata
    try:
        s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode(
            'ascii')
    except Exception:
        pass
    return s


def date_to_int(a_date):
    return (a_date - ORIGIN).days


def date_to_float(a_date):
    if not isinstance(a_date, datetime.datetime):
        if isinstance(a_date, datetime.date):
            a_date = datetime.datetime(a_date.year, a_date.month, a_date.day)
        elif isinstance(a_date, datetime.time):
            a_date = datetime.datetime(0, 0, 0, a_date.hour, a_date.minute,
                                       a_date.second, a_date.microsecond)

    return (a_date - ORIGIN).total_seconds() / 86400


def int_to_date(days):
    return ORIGIN + datetime.timedelta(days)


def float_to_date(days):
    return ORIGIN + datetime.timedelta(days)
