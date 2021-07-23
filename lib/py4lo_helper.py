# -*- coding: utf-8 -*-
"""Py4LO - Python Toolkit For LibreOffice Calc
      Copyright (C) 2016-2021 J. Férard <https://github.com/jferard>

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

"""py4lo_helper deals with LO objects."""

import os
import encodings
import uno

try:
    import unohelper
except ImportError:
    import unotools.unohelper

# py4lo: if $python_version >= 2.6
# py4lo: if $python_version <= 3.0
import io
# py4lo: endif
# py4lo: endif
from py4lo_commons import float_to_date
from com.sun.star.uno import RuntimeException, Exception as UnoException
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.sheet.ConditionOperator import FORMULA
from com.sun.star.script.provider import ScriptFrameworkErrorException
# py4lo: if $python_version >= 3.0
from com.sun.star.awt.MessageBoxType import MESSAGEBOX

# py4lo: else
MESSAGEBOX = 0
# py4lo: endif

from com.sun.star.lang import Locale

provider = None
_inspect = None
xray = None
mri = None


def init(xsc):
    """
    Mandatory call from entry with XSCRIPTCONTEXT as argument.
    @param xsc: XSCRIPTCONTEXT
    """
    global provider, _inspect, xray, mri
    provider = _ObjectProvider.create(xsc)
    _inspect = _Inspector(provider)
    xray = _inspect.xray
    mri = _inspect.mri


class _ObjectProvider:
    """
    Lazy object provider.
    """

    @staticmethod
    def create(xsc):
        doc = xsc.getDocument()
        controller = doc.CurrentController
        frame = controller.Frame
        parent_win = frame.ContainerWindow
        script_provider = doc.getScriptProvider()
        ctxt = xsc.getComponentContext()
        service_manager = ctxt.getServiceManager()
        desktop = xsc.getDesktop()
        return _ObjectProvider(doc, controller, frame, parent_win,
                               script_provider, ctxt, service_manager, desktop)

    def __init__(self, doc, controller, frame, parent_win, script_provider,
                 ctxt, service_manager, desktop):
        self.doc = doc
        self.controller = controller
        self.frame = frame
        self.parent_win = parent_win
        self.script_provider = script_provider
        self.ctxt = ctxt
        self.service_manager = service_manager
        self.desktop = desktop
        self._script_provider_factory = None
        self._script_provider = None
        self._reflect = None
        self._dispatcher = None

    def get_script_provider_factory(self):
        """
        > This service is used to create MasterScriptProviders
        @return:
        """
        if self._script_provider_factory is None:
            self._script_provider_factory = self.service_manager.createInstanceWithContext(
                "com.sun.star.script.provider.MasterScriptProviderFactory",
                self.ctxt)
        return self._script_provider_factory

    def get_script_provider(self):
        """
        > This interface provides a factory for obtaining objects implementing the XScript interface

        @return:
        """
        if self._script_provider is None:
            self._script_provider = self.get_script_provider_factory().createScriptProvider(
                "")
        return self._script_provider

    @property
    def reflect(self):
        """
        > This service is the implementation of the reflection API

        @return:
        """
        if self._reflect is None:
            self._reflect = self.service_manager.createInstance(
                "com.sun.star.reflection.CoreReflection")
        return self._reflect

    @property
    def dispatcher(self):
        """
        > provides an easy way to dispatch a URL using one call instead of multiple ones.
        @return:
        """
        if self._dispatcher is None:
            self._dispatcher = self.service_manager.createInstance(
                "com.sun.star.frame.DispatchHelper")
        return self._dispatcher


class _Inspector:
    def __init__(self, provider):
        self._provider = provider
        self._xray_script = None
        self._ignore_xray = False
        self._oMRI = None
        self._ignore_mri = False

    def use_xray(self, fail_on_error=False):
        """
        Try to load Xray lib.
        :param fail_on_error: Should this function fail on error
        :raises RuntimeException: if Xray is not avaliable and `fail_on_error` is True.
        """
        try:
            self._xray_script = self._provider.script_provider.getScript(
                "vnd.sun.star.script:XrayTool._Main.Xray?"
                "language=Basic&location=application")
        except ScriptFrameworkErrorException:
            if fail_on_error:
                raise RuntimeException("\nBasic library Xray is not installed",
                                       self._provider.ctxt)
            else:
                self._ignore_xray = True

    def xray(self, obj, fail_on_error=False):
        """
        Xray an obj. Loads dynamically the lib if possible.
        :@param obj: the obj
        :@param fail_on_error: Should this function fail on error
        :@raises RuntimeException: if Xray is not avaliable and `fail_on_error` is True.
        """
        if self._ignore_xray:
            return

        if self._xray_script is None:
            self.use_xray(fail_on_error)
            if self._ignore_xray:
                return

        self._xray_script.invoke((obj,), (), ())

    def mri(self, obj, fail_on_error=False):
        """
        MRI an object
        @param fail_on_error:
        @param obj: the object
        """
        if self._ignore_mri:
            return

        if self._oMRI is None:
            try:
                self._oMRI = uno_service("mytools.Mri")
            except UnoException:
                if fail_on_error:
                    raise RuntimeException("\nMRI is not installed",
                                           self._provider.ctxt)
                else:
                    self._ignore_mri = True
        self._oMRI.inspect(obj)


def uno_service(sname, args=None, ctxt=None):
    sm = provider.service_manager
    if ctxt is None:
        return sm.createInstance(sname)
    else:
        if args is None:
            return sm.createInstanceWithContext(sname, ctxt)
        else:
            return sm.createInstanceWithArgumentsAndContext(sname, args, ctxt)


def uno_service_ctxt(sname, args=None):
    return uno_service(sname, args, provider.ctxt)


def message_box(msg_text, msg_title, msg_type=MESSAGEBOX,
                msg_buttons=BUTTONS_OK, parent_win=None):
    # from https://forum.openoffice.org/fr/forum/viewtopic.php?f=15&t=47603#
    # (thanks Bernard !)
    sv = uno_service_ctxt("com.sun.star.awt.Toolkit")
    if parent_win is None:
        parent_win = provider.parent_win
    mb = sv.createMessageBox(parent_win, msg_type, msg_buttons, msg_title,
                             msg_text)
    return mb.execute()


###
# Open a document
###
NEW_CALC_DOCUMENT = "private:factory/scalc"
NEW_WRITER_DOCUMENT = "private:factory/swriter"

# special targets
TARGET_BLANK = "_blank"  # always creates a new frame
TARGET_DEFAULT = "_default"  # special UI functionality (e.g. detecting of already loaded documents, using of empty frames of creating of new top frames as fallback)
TARGET_SELF = "_self"  # means frame himself
TARGET_PARENT = "_parent"  # address direct parent of frame
TARGET_TOP = "_top"  # indicates top frame of current path in tree
TARGET_BEAMER = "_beamer"  # means special sub frame

FRAME_FLAG_AUTO = 0  # no longer supported
FRAME_FLAG_PARENT = 1  # allows search on the parent frames
FRAME_FLAG_SELF = 2  # includes the start frame himself
FRAME_FLAG_CHILDREN = 4  # includes all child frames of the start frame
FRAME_FLAG_CREATE = 8  # frame will be created if not found
FRAME_FLAG_SIBLINGS = 16  # includes the direct siblings of the start frame
FRAME_FLAG_TASKS = 32  # allow the search outside the current sub task tree of the whole possible frame tree
FRAME_FLAG_ALL = 23  # includes all frames except frames in other tasks sub trees but doesn't create any new frame
FRAME_FLAG_GLOBAL = 55  # searches in the whole hierarchy of frames but doesn't create any new frame


def open_in_calc(filename, target=TARGET_BLANK, frame_flags=FRAME_FLAG_AUTO,
                 **kwargs):
    """
    Open a document in calc
    :param filename: the name of the file to open
    :param target: "the name of the frame to view the document in" or a special target
    :param frame_flags: where to search the frame
    :param kwargs: les paramètres d'ouverture
    :return: a reference on the doc
    """
    url = uno.systemPathToFileUrl(os.path.realpath(filename))
    if kwargs:
        params = make_pvs(kwargs)
    else:
        params = ()
    return provider.desktop.loadComponentFromURL(url, target, frame_flags,
                                                 params)


###
# Filters
###
FILTER_XML = "StarOffice XML (Calc)"  # Standard XML filter
FILTER_XML_TEMPLATE = "calc_StarOffice_XML_Calc_Template"  # XML filter for templates
FILTER_STARCALC_5 = "StarCalc 5.0"  # The binary format of StarOffice Calc 5.x
FILTER_STARCALC_5_TEMPLATE = "StarCalc 5.0 Vorlage/Template"  # StarOffice Calc 5.x templates
FILTER_STARCALC_4 = "StarCalc 4.0"  # The binary format of StarCalc 4.x
FILTER_STARCALC_4_TEMPLATE = "StarCalc 4.0 Vorlage/Template"  # StarCalc 4.x templates
FILTER_STARCALC_3 = "StarCalc 3.0"  # The binary format of StarCalc 3.x
FILTER_STARCALC_3_TEMPLATE = "StarCalc 3.0 Vorlage/Template"  # StarCalc 3.x templates
FILTER_HTML = "HTML (StarCalc)"  # HTML filter
FILTER_HTML_WEBQUERY = "calc_HTML_WebQuery"  # HTML filter for external data queries
FILTER_EXCEL_97 = "MS Excel 97"  # Microsoft Excel 97/2000/XP
FILTER_EXCEL_97_TEMPLATE = "MS Excel 97 Vorlage/Template"  # Microsoft Excel 97/2000/XP templates
FILTER_EXCEL_95 = "MS Excel 95"  # Microsoft Excel 5.0/95
FILTER_EXCEL_95_TEMPLATE = "MS Excel 95 Vorlage/Template"  # Microsoft Excel 5.0/95 templates
FILTER_EXCEL_2_3_4 = "MS Excel 4.0"  # Microsoft Excel 2.1/3.0/4.0
FILTER_EXCEL_2_3_4_TEMPLATE = "MS Excel 4.0 Vorlage/Template"  # Microsoft Excel 2.1/3.0/4.0 templates
FILTER_LOTUS = "Lotus"  # Lotus 1-2-3
FILTER_CSV = "Text - txt - csv (StarCalc)"  # Comma separated values
FILTER_RTF = "Rich Text Format (StarCalc)"  #
FILTER_DBASE = "dBase"  # dBase
FILTER_SYLK = "SYLK"  # Symbolic Link
FILTER_DIF = "DIF"  # Data Interchange Format

INDEX_BY_ENCODING = {"unknown": 0,
                     "cp1252": 1,
                     "mac_roman": 2,
                     "cp437": 3,
                     "cp850": 4,
                     "cp860": 5,
                     "cp861": 6,
                     "cp863": 7,
                     "cp865": 8,
                     "default": 9,
                     "cp1038": 10,
                     "ascii": 11,
                     "latin_1": 12,
                     "iso8859_2": 13,
                     "iso8859_3": 14,
                     "iso8859_4": 15,
                     "iso8859_5": 16,
                     "iso8859_6": 17,
                     "iso8859_7": 18,
                     "iso8859_8": 19,
                     "iso8859_9": 20,
                     "iso8859_14": 21,
                     "iso8859_15": 22,
                     "cp737": 23,
                     "cp775": 24,
                     "cp852": 25,
                     "cp855": 26,
                     "cp857": 27,
                     "cp862": 28,
                     "cp864": 29,
                     "cp866": 30,
                     "cp869": 31,
                     "cp874": 32,
                     "cp1250": 33,
                     "cp1251": 34,
                     "cp1253": 35,
                     "cp1254": 36,
                     "cp1255": 37,
                     "cp1256": 38,
                     "cp1257": 39,
                     "cp1258": 40,
                     "mac_arabic": 41,
                     "mac-ce": 42,
                     "mac_croatian": 43,
                     "mac_cyrillic": 44,
                     "mac_devanagari": 45,
                     "mac_farsi": 46,
                     "mac_greek": 47,
                     "mac_gujarati": 48,
                     "mac_gurmukhi": 49,
                     "mac_hebrew": 50,
                     "mac_iceland": 51,
                     "mac_latin2": 52,
                     "mac_thai": 53,
                     "mac_turkish": 54,
                     "mac_ukrainian": 55,
                     "mac_chinesesimp": 56,
                     "mac_": 57,
                     "mac_japanese": 58,
                     "mac_korean": 59,
                     "cp932": 60,
                     "cp936": 61,
                     "cp949": 62,
                     "cp950": 63,
                     "shift_jis": 64,
                     "gb2312": 65,
                     "gb12345": 66,
                     "gbk": 67,
                     "big5": 68,
                     "euc_jp": 69,
                     "g2312": 70,
                     "euc_tw": 71,
                     "iso2022_jp": 72,
                     "is_2022": 73,
                     "koi8_r": 74,
                     "utf_7": 75,
                     "utf_8": 76,
                     "iso8859_10": 77,
                     "iso8859_13": 78,
                     "euc_kr": 79,
                     "iso2022_kr": 80,
                     "jis_x0201": 81,
                     "jis_x0208": 82,
                     "jis_x0212": 83,
                     "johab": 84,
                     "gb18030": 85,
                     "big5hkscs": 86,
                     "tis_620": 87,
                     "koi8_u": 88,
                     "iscii": 89,
                     "java_utf_8": 90,
                     "cp1276": 91,
                     "adobe_symbol": 92,
                     "ptcp154": 93,
                     "utf_32": 65534,
                     "utf_16": 65535}

FORMAT_STANDARD = 1  # Standard
FORMAT_TEXT = 2  # Text
FORMAT_MM_DD_YY = 3  # MM/DD/YY
FORMAT_DD_MM_YY = 4  # DD/MM/YY
FORMAT_YY_MM_DD = 5  # YY/MM/DD
FORMAT_IGNORE = 9  # IGNORE FIELD (do not import)
FORMAT_US = 10  # US-English


def import_filter_options(delimiter=",", quotechar='"', encoding="utf-8",
                          first_line=1, format_by_idx=None,
                          quoted_field_as_text=False,
                          detect_special_numbers=False):
    """
    See: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options
    @param delimiter: the delimiter
    @param quotechar: the quotechar
    @param encoding: the encoding
    @param first_line: the first line
    @param format_by_idx: a mapping field index (starting at 1) -> field format
    @param quoted_field_as_text: see checkbox
    @param detect_special_numbers: see checkbox
    @return: a CSV filter options string
    """
    options = _base_filter_options(
        delimiter, quotechar, encoding, first_line, format_by_idx
    ) + ["", str(quoted_field_as_text).lower(),
         str(detect_special_numbers).lower()]
    return ",".join(options)


def export_filter_options(delimiter=",", quotechar='"', encoding="utf-8",
                          first_line=1, format_by_idx=None,
                          quote_all_text_cells=False,
                          store_as_numbers=True,
                          save_cell_contents_as_shown=True):
    """
    See: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options
    @param delimiter: the delimiter
    @param quotechar: the quotechar
    @param encoding: the encoding
    @param first_line: the first line
    @param format_by_idx: a mapping field index (starting at 1) -> field format
    @param quote_all_text_cells: see checkbox
    @param store_as_numbers: if false, quote numbers (?)
    @param save_cell_contents_as_shown: see checkbox
    @return: a CSV filter options string
    """
    options = _base_filter_options(
        delimiter, quotechar, encoding, first_line, format_by_idx
    ) + ["", str(quote_all_text_cells).lower(), str(store_as_numbers).lower(),
         str(save_cell_contents_as_shown).lower()]
    return ",".join(options)


def _base_filter_options(delimiter, quotechar, encoding, first_line,
                         format_by_idx):
    """
    See: https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options
    @param delimiter: the delimiter
    @param quotechar: the quotechar
    @param encoding: the encoding
    @param first_line: the first line
    @param format_by_idx: a mapping field index (starting at 1) -> field format
    @return: a list of options
    """
    norm_encoding = encodings.normalize_encoding(encoding)
    norm_encoding = encodings.aliases.aliases.get(
        norm_encoding.lower(), norm_encoding)
    encoding_index = INDEX_BY_ENCODING.get(norm_encoding, 0)

    if format_by_idx is None:
        field_formats = ""
    else:
        field_formats = "/".join(["{}/{}".format(idx, format)
                                  for idx, format in format_by_idx.items()])

    return [str(ord(delimiter)), str(ord(quotechar)), str(encoding_index),
            str(first_line),
            field_formats]


# Create a document
###
def doc_builder(t="calc"):
    return DocBuilder(t)


def new_doc(t="calc"):
    """Create a blank new doc"""
    return doc_builder(t).build()


def add_filter(oDoc, oSheet, range_name):
    oController = oDoc.CurrentController
    oAll = oSheet.getCellRangeByName(range_name)
    oController.select(oAll)
    oFrame = oController.Frame
    provider.dispatcher.executeDispatch(oFrame, ".uno:DataFilterAutoFilter", "",
                                        0, ())


def read_options_from_sheet_name(oDoc, sheet_name, l=lambda s: s):
    oSheet = oDoc.Sheets.getByName(sheet_name)
    oRangeAddress = get_used_range_address(oSheet)
    return read_options(oSheet, oRangeAddress, l)


def get_named_cells(oDoc, name):
    return oDoc.NamedRanges.getByName(name).ReferredCells


def get_named_cell(oDoc, name):
    return get_named_cells(oDoc, name).getCellByPosition(0, 0)


def set_validation_list_by_name(oDoc, cell_name, fields,
                                default_string=None, allow_blank=False):
    oCell = get_named_cell(oDoc, cell_name)
    set_validation_list_by_cell(oCell, fields, default_string, allow_blank)


class DocBuilder:
    def __init__(self, t):
        """Create a blank new doc"""
        self._oDoc = provider.desktop.loadComponentFromURL(
            "private:factory/s" + t, "_blank", 0, ())
        self._oDoc.lockControllers()

    def build(self):
        self._oDoc.unlockControllers()
        return self._oDoc

    def sheet_names(self, sheet_names, expand_if_necessary=True,
                    trunc_if_necessary=True):
        oSheets = self._oDoc.Sheets
        it = iter(sheet_names)
        s = 0

        try:
            # rename
            while s < oSheets.getCount():
                oSheet = oSheets.getByIndex(s)
                oSheet.setName(next(it))  # may raise a StopIteration
                s += 1

            assert s == oSheets.getCount(), "s={} vs oSheets.getCount()={}".format(
                s, oSheets.getCount())

            if expand_if_necessary:
                # add
                for sheet_name in it:
                    oSheets.insertNewByName(sheet_name, s)
                    s += 1
        except StopIteration:  # it
            assert s <= oSheets.getCount(), "s={} vs oSheets.getCount()={}".format(
                s, oSheets.getCount())
            if trunc_if_necessary:
                self.trunc_to_count(s)

        return self

    def apply_func_to_sheets(self, func):
        oSheets = self._oDoc.Sheets
        for oSheet in oSheets:
            func(oSheet)
        return self

    def apply_func_list_to_sheets(self, funcs):
        oSheets = self._oDoc.Sheets
        for func, oSheet in zip(funcs, oSheets):
            func(oSheet)
        return self

    def duplicate_base_sheet(self, func, sheet_names, trunc=True):
        """Create a base sheet and duplicate it n-1 times"""
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        func(oBaseSheet)
        for s, sheet_name in enumerate(sheet_names, 1):
            oSheets.copyByName(oBaseSheet.Name, sheet_name, s)

        return self

    def make_base_sheet(self, func):
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        func(oBaseSheet)
        return self

    def duplicate_to(self, n):
        oSheets = self._oDoc.Sheets
        oBaseSheet = oSheets.getByIndex(0)
        for s in range(n + 1):
            oSheets.copyByName(oBaseSheet.Name, oBaseSheet.Name + str(s), s)

        return self

    def trunc_to_count(self, final_sheet_count):
        oSheets = self._oDoc.Sheets
        while final_sheet_count < oSheets.getCount():
            oSheet = oSheets.getByIndex(final_sheet_count)
            oSheets.removeByName(oSheet.getName())

        return self


def make_pv(name, value):
    pv = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
    pv.Name = name
    pv.Value = value
    return pv


def make_pvs(d={}):
    return tuple(make_pv(n, v) for n, v in d.items())


def get_last_used_row(oSheet):
    return get_used_range_address(oSheet).EndRow


def get_used_range_address(oSheet):
    oCell = oSheet.getCellByPosition(0, 0)
    oCursor = oSheet.createCursorByRange(oCell)
    oCursor.gotoEndOfUsedArea(True)
    return oCursor.RangeAddress


def get_used_range(oSheet):
    oRangeAddress = get_used_range_address(oSheet)
    return oSheet.getCellRangeByPosition(oRangeAddress.StartColumn,
                                         oRangeAddress.StartRow,
                                         oRangeAddress.EndColumn,
                                         oRangeAddress.EndRow)


def data_array(oSheet):
    return get_used_range(oSheet).DataArray


def to_iter(oXIndexAccess):
    for i in range(0, oXIndexAccess.getCount()):
        yield oXIndexAccess.getByIndex(i)


def to_dict(oXNameAccess):
    d = {}
    for name in oXNameAccess.getElementNames():
        d[name] = oXNameAccess.getByName(name)
    return d


# l is deprecated
def read_options(oSheet, aAddress, l=lambda s: s):
    options = {}
    width = aAddress.EndColumn - aAddress.StartColumn
    for r in range(aAddress.StartRow, aAddress.EndRow + 1):
        k = oSheet.getCellByPosition(aAddress.StartColumn, r).String.strip()
        if width <= 1:
            v = oSheet.getCellByPosition(aAddress.StartColumn + 1,
                                         r).String.strip()
        else:
            # from aAddress.StartColumn+1 to aAddress.EndColumn
            v = [oSheet.getCellByPosition(aAddress.StartColumn + 1 + c,
                                          r).String.strip() for c in
                 range(width)]
        if k and v:
            options[l(k)] = l(v)
    return options


def set_validation_list_by_cell(oCell, fields, default_string=None,
                                allow_blank=False):
    oValidation = oCell.Validation
    oValidation.Type = uno.getConstantByName(
        "com.sun.star.sheet.ValidationType.LIST")
    oValidation.ShowErrorMessage = True
    oValidation.ShowList = uno.getConstantByName(
        "com.sun.star.sheet.TableValidationVisibility.UNSORTED")
    formula = ";".join('"' + a + '"' for a in fields)
    oValidation.setFormula1(formula)
    oValidation.IgnoreBlankCells = allow_blank
    oCell.Validation = oValidation

    if default_string is not None:
        oCell.String = default_string


def clear_conditional_format(oSheet, range_name):
    oColoredColumns = oSheet.getCellRangeByName(range_name)
    oConditionalFormat = oColoredColumns.ConditionalFormat
    oConditionalFormat.clear()


def conditional_format_on_formulas(oSheet, range_name, style_by_formula,
                                   source_position=(0, 0)):
    oColoredColumns = oSheet.getCellRangeByName(range_name)
    oConditionalFormat = oColoredColumns.ConditionalFormat
    oSrc = oColoredColumns.getCellByPosition(*source_position).CellAddress

    for formula, style in style_by_formula.items():
        oConditionalEntry = get_formula_conditional_entry(formula, style, oSrc)
        oConditionalFormat.addNew(oConditionalEntry)

    oColoredColumns.ConditionalFormat = oConditionalFormat


def get_formula_conditional_entry(formula, style_name, oSrc):
    return get_conditional_entry(formula, "", FORMULA, style_name, oSrc)


def get_conditional_entry(formula1, formula2, operator, style_name, oSrc):
    return (
        make_pvs(
            {"Formula1": formula1, "Formula2": formula2, "Operator": operator,
             "StyleName": style_name,
             "SourcePosition": oSrc})
    )


# from Andrew Pitonyak 5.14 www.openoffice.org/documentation/HOW_TO/various_topics/AndrewMacro.odt
def find_or_create_number_format_style(oDoc, format):
    oFormats = oDoc.getNumberFormats()
    oLocale = Locale()
    formatNum = oFormats.queryKey(format, oLocale, True)
    if formatNum == -1:
        formatNum = oFormats.addNew(format, oLocale)

    return formatNum


def copy_row_at_index(oSheet, row, r):
    for c, value in enumerate(row):
        if type(value) == str:
            oSheet.getCellByPosition(c, r).String = value
        else:
            oSheet.getCellByPosition(c, r).Value = float(value)


def get_range_size(oRange):
    """
    Useful for `oRange.getRangeCellByPosition(...)`.

    :param oRange: a SheetCellRange obj
    :return: width and height of the range.
    """
    oAddress = oRange.RangeAddress
    width = oAddress.EndColumn - oAddress.StartColumn + 1
    height = oAddress.EndRow - oAddress.StartRow + 1
    return width, height


def get_doc(oCell):
    """
    Find the document that owns this cell.

    @param oCell: a cell
    @return: the document
    """
    return oCell.Spreadsheet.DrawPage.Forms.Parent


def get_cell_type(oCell):
    """
    @param oCell: the cell
    @return: 'EMPTY', 'TEXT', 'VALUE'
    """
    cell_type = oCell.getType().value
    if cell_type == 'FORMULA':
        cell_type = oCell.FormulaResultType.value

    return cell_type
