#  Py4LO - Python Toolkit For LibreOffice Calc
#     Copyright (C) 2016-2019 J. Férard <https://github.com/jferard>
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
import csv
import locale
import os
from datetime import (date, datetime, time)
import encodings
from encodings.aliases import aliases as enc_aliases
import itertools
import sys

import uno
from com.sun.star.util import NumberFormat
from com.sun.star.lang import Locale

# values of type_cell
from py4lo_commons import float_to_date, date_to_float
from py4lo_helper import (provider as pr, make_pvs, get_doc, get_cell_type,
                          get_used_range_address)

TYPE_NONE = 0
TYPE_MINIMAL = 1
TYPE_ALL = 2


##########
# Reader #
##########

def create_read_cell(type_cell=TYPE_MINIMAL, oFormats=None):
    """
    Create a function to read a cell
    @param type_cell: one of `TYPE_NONE` (return the String value),
                      `TYPE_MINIMAL` (String or Value), `TYPE_ALL` (the most
                      accurate type)
    @param oFormats: the container for NumberFormats.
    @return: a function to read the cell value
    """
    def read_cell_none(oCell):
        """
        Read a cell value
        @param oCell: the cell
        @return: the cell value as string
        """
        return oCell.String

    def read_cell_minimal(oCell):
        """
        Read a cell value
        @param oCell: the cell
        @return: the cell value as float or str
        """
        cell_type = get_cell_type(oCell)

        if cell_type == 'EMPTY':
            return None
        elif cell_type == 'TEXT':
            return oCell.String
        elif cell_type == 'VALUE':
            return oCell.Value

    def read_cell_all(oCell):
        """
        Read a cell value
        @param oCell: the cell
        @return: the cell value as float, bool, date or str
        """
        cell_type = get_cell_type(oCell)

        if cell_type == 'EMPTY':
            return None
        elif cell_type == 'TEXT':
            return oCell.String
        elif cell_type == 'VALUE':
            key = oCell.NumberFormat
            cell_data_type = oFormats.getByKey(key).Type
            if cell_data_type in {NumberFormat.DATE, NumberFormat.DATETIME,
                                  NumberFormat.TIME}:
                return float_to_date(oCell.Value)
            elif cell_data_type == NumberFormat.LOGICAL:
                return bool(oCell.Value)
            else:
                return oCell.Value

    if type_cell == TYPE_NONE:
        return read_cell_none
    elif type_cell == TYPE_MINIMAL:
        return read_cell_minimal
    elif type_cell == TYPE_ALL:
        if oFormats is None:
            raise ValueError("Need formats to type all values")
        return read_cell_all
    else:
        raise ValueError("type_cell must be one of TYPE_* values")


class reader:
    """
    A reader that returns rows as lists of values.
    """
    def __init__(self, oSheet, type_cell=TYPE_MINIMAL, oFormats=None,
                 read_cell=None):
        if read_cell is not None:
            self._read_cell = read_cell
        else:
            if type_cell == TYPE_ALL and oFormats is None:
                oFormats = oSheet.DrawPage.Forms.Parent.NumberFormats
            self._read_cell = create_read_cell(type_cell, oFormats)
        self._oSheet = oSheet
        self.line_num = 0
        self._oRangeAddress = get_used_range_address(oSheet)

    def __iter__(self):
        return self

    def __next__(self):
        i = self._oRangeAddress.StartRow + self.line_num
        if i > self._oRangeAddress.EndRow:
            raise StopIteration

        self.line_num += 1
        row = [self._read_cell(self._oSheet.getCellByPosition(j, i))
               for j in range(self._oRangeAddress.StartColumn,
                              self._oRangeAddress.EndColumn + 1)]

        # left strip the row
        i = len(row) - 1
        while row[i] is None and i > 0:
            i -= 1
        return row[:i + 1]


class dict_reader:
    """
    A reader that returns rows as dicts.
    """
    def __init__(self, oSheet, fieldnames=None, restkey=None, restval=None,
                 type_cell=TYPE_MINIMAL, oFormats=None, read_cell=None):
        self._reader = reader(oSheet, type_cell, oFormats, read_cell)
        if fieldnames is None:
            self.fieldnames = next(self._reader)
        else:
            self.fieldnames = fieldnames
        self._width = len(self.fieldnames)
        self.restkey = restkey
        self.restval = restval

    def __iter__(self):
        return self

    def __next__(self):
        row = next(self._reader)
        row_width = len(row)
        if row_width == self._width:
            return dict(zip(self.fieldnames, row))
        elif row_width < self._width:
            row += [self.restval] * (self._width - row_width)
            return dict(zip(self.fieldnames, row))
        elif self.restkey is None:
            return dict(zip(self.fieldnames, row))
        else:
            d = dict(zip(self.fieldnames, row))
            d[self.restkey] = row[self._width:]
            return d

    @property
    def line_num(self):
        return self._reader.line_num


##########
# Writer #
##########
def find_number_format_style(oFormats, format_id, oLocale=Locale()):
    """
    
    @param oFormats: the formats 
    @param format_id: a NumberFormat
    @param oLocale: the locale
    @return: the id of the format
    """
    return oFormats.getStandardFormat(format_id, oLocale)


def create_write_cell(type_cell=TYPE_MINIMAL, oFormats=None):
    """
    Create a cell writer
    @param type_cell: see `create_read_cell`
    @param oFormats: the NumberFormats
    @return: a function
    """
    def write_cell_none(oCell, value):
        """
        Write a cell value
        @param oCell: the cell
        @param value: the value
        """
        oCell.String = str(value)

    def write_cell_minimal(oCell, value):
        """
        Write a cell value
        @param oCell: the cell
        @param value: the value
        """
        if value is None:
            oCell.String = ""
        elif isinstance(value, str):
            oCell.String = value
        elif isinstance(value, (date, datetime, time)):
            oCell.Value = date_to_float(value)
        else:
            oCell.Value = value

    def create_write_cell_all(oFormats):
        date_id = find_number_format_style(oFormats, NumberFormat.DATE)
        datetime_id = find_number_format_style(oFormats, NumberFormat.DATETIME)
        boolean_id = find_number_format_style(oFormats, NumberFormat.LOGICAL)

        def write_cell_all(oCell, value):
            """
            Write a cell value
            @param oCell: the cell
            @param value: the value
            @return: the cell value as float or text
            """
            if value is None:
                oCell.String = ""
            elif isinstance(value, str):
                oCell.String = value
            elif isinstance(value, (datetime, time)):
                oCell.Value = date_to_float(value)
                oCell.NumberFormat = datetime_id
            elif isinstance(value, date):
                oCell.Value = date_to_float(value)
                oCell.NumberFormat = date_id
            elif isinstance(value, bool):
                oCell.Value = value
                oCell.NumberFormat = boolean_id
            else:
                oCell.Value = value

        return write_cell_all

    if type_cell == TYPE_NONE:
        return write_cell_none
    elif type_cell == TYPE_MINIMAL:
        return write_cell_minimal
    elif type_cell == TYPE_ALL:
        if oFormats is None:
            raise ValueError("Need formats to type all values")
        return create_write_cell_all(oFormats)
    else:
        raise ValueError("type_cell must be one of TYPE_* values")


class writer:
    """
    A writer that takes lists
    """

    def __init__(self, oSheet, type_cell=TYPE_MINIMAL, oFormats=None,
                 write_cell=None,
                 initial_pos=(0, 0)):
        self._oSheet = oSheet
        self._row, self._base_col = initial_pos
        if write_cell is not None:
            self._write_cell = write_cell
        else:
            if type_cell == TYPE_ALL and oFormats is None:
                oFormats = oSheet.DrawPage.Forms.Parent.NumberFormats
            self._write_cell = create_write_cell(type_cell, oFormats)

    def writerow(self, row):
        for i, value in enumerate(row, self._base_col):
            oCell = self._oSheet.getCellByPosition(i, self._row)
            self._write_cell(oCell, value)
        self._row += 1

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)


class dict_writer:
    """
    A writer that takes dicts
    """
    def __init__(self, oSheet, fieldnames, restval='', extrasaction='raise',
                 type_cell=TYPE_MINIMAL, oFormats=None, write_cell=None):
        self.writer = writer(oSheet, type_cell, oFormats, write_cell)
        self.fieldnames = fieldnames
        self._set_fieldnames = set(fieldnames)
        self.restval = restval
        self.extrasaction = extrasaction

    def writeheader(self):
        self.writer.writerow(self.fieldnames)

    def writerow(self, row):
        if self.extrasaction == 'raise' and set(row) - self._set_fieldnames:
            raise ValueError()
        flat_row = [row.get(name, self.restval) for name in self.fieldnames]
        self.writer.writerow(flat_row)

    def writerows(self, rows):
        for row in rows:
            self.writerow(row)


#####################
# Import/Export CSV #
#####################

# see https://api.libreoffice.org/docs/cpp/ref/a00391_source.html (rtl/textenc.h)
CHARSET_ID_BY_NAME = {
    'adobe_dingbats': 94,
    'adobe_standard': 91,
    'adobe_symbol': 92,
    'ascii': 11,
    'big5': 68,
    'big5hkscs': 86,
    'cp1250': 33,
    'cp1251': 34,
    'cp1252': 1,
    'cp1253': 35,
    'cp1254': 36,
    'cp1255': 37,
    'cp1256': 38,
    'cp1257': 39,
    'cp1258': 40,
    'cp1361': 84,
    'cp437': 3,
    'cp737': 23,
    'cp775': 24,
    'cp850': 4,
    'cp852': 25,
    'cp855': 26,
    'cp857': 27,
    'cp860': 5,
    'cp861': 6,
    'cp862': 28,
    'cp863': 7,
    'cp864': 29,
    'cp865': 8,
    'cp866': 30,
    'cp869': 31,
    'cp874': 32,
    'cp932': 60,
    'cp936': 61,
    'cp949': 62,
    'cp950': 63,
    'euc_cn': 70,
    'euc_jp': 69,
    'euc_kr': 79,
    'euc_tw': 71,
    'gb18030': 85,
    'gb2312': 65,
    'gbk': 67,
    'gbt12345': 66,
    'iscii_devanagari': 89,
    'iso2022_cn': 73,
    'iso2022_jp': 72,
    'iso2022_kr': 80,
    'iso8859_1': 12,
    'iso8859_10': 77,
    'iso8859_13': 78,
    'iso8859_14': 21,
    'iso8859_15': 22,
    'iso8859_2': 13,
    'iso8859_3': 14,
    'iso8859_4': 15,
    'iso8859_5': 16,
    'iso8859_6': 17,
    'iso8859_7': 18,
    'iso8859_8': 19,
    'iso8859_9': 20,
    'java_utf8': 90,
    'jis_x_0201': 81,
    'jis_x_0208': 82,
    'jis_x_0212': 83,
    'koi8_r': 74,
    'koi8_u': 88,
    'mac_arabic': 41,
    'mac_centeuro': 42,
    'mac_chinsimp': 56,
    'mac_chintrad': 57,
    'mac_croatian': 43,
    'mac_cyrillic': 44,
    'mac_devanagari': 45,
    'mac_farsi': 46,
    'mac_greek': 47,
    'mac_gujarati': 48,
    'mac_gurmukhi': 49,
    'mac_hebrew': 50,
    'mac_iceland': 51,
    'mac_japanese': 58,
    'mac_korean': 59,
    'mac_roman': 2,
    'mac_romanian': 52,
    'mac_thai': 53,
    'mac_turkish': 54,
    'mac_ukrainian': 55,
    'ptcp154': 93,
    'shift_jis': 64,
    'symbol': 10,
    'tis_620': 87,
    'utf_16': 65535,
    'utf_32': 65534,
    'user_end': 61439,
    'user_start': 32768,
    sys.getdefaultencoding(): 9,
    'utf_7': 75,
    'utf_8': 76
}

FORMAT_STANDARD = 1
FORMAT_TEXT = 2
FORMAT_MM_DD_YY = 3
FORMAT_DD_MM_YY = 4
FORMAT_YY_MM_DD = 5
FORMAT_IGNORE = 9
FORMAT_US_ENGLISH = 10

# see https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oe376/6c085406-a698-4e12-9d4d-c3b0ee3dbc4a
LANGUAGE_ID_BY_CODE = {
    "ar_SA": 1025,
    "bg_BG": 1026,
    "ca_ES": 1027,
    "zh_TW": 1028,
    "cs_CZ": 1029,
    "da_DK": 1030,
    "de_DE": 1031,
    "el_GR": 1032,
    "en_US": 1033,
    "es_ES": 1034,
    "fi_FI": 1035,
    "fr_FR": 1036,
    "he_IL": 1037,
    "hu_HU": 1038,
    "is_IS": 1039,
    "it_IT": 1040,
    "ja_JP": 1041,
    "ko_KR": 1042,
    "nl_NL": 1043,
    "nb_NO": 1044,
    "pl_PL": 1045,
    "pt_BR": 1046,
    "rm_CH": 1047,
    "ro_RO": 1048,
    "ru_RU": 1049,
    "hr_HR": 1050,
    "sk_SK": 1051,
    "sq_AL": 1052,
    "sv_SE": 1053,
    "th_TH": 1054,
    "tr_TR": 1055,
    "ur_PK": 1056,
    "id_ID": 1057,
    "uk_UA": 1058,
    "be_BY": 1059,
    "sl_SI": 1060,
    "et_EE": 1061,
    "lv_LV": 1062,
    "lt_LT": 1063,
    "tg_Cyrl_TJ": 1064,
    "fa_IR": 1065,
    "vi_VN": 1066,
    "hy_AM": 1067,
    "az_Latn_AZ": 1068,
    "eu_ES": 1069,
    "wen_DE": 1070,
    "mk_MK": 1071,
    "st_ZA": 1072,
    "ts_ZA": 1073,
    "tn_ZA": 1074,
    "ven_ZA": 1075,
    "xh_ZA": 1076,
    "zu_ZA": 1077,
    "af_ZA": 1078,
    "ka_GE": 1079,
    "fo_FO": 1080,
    "hi_IN": 1081,
    "mt_MT": 1082,
    "se_NO": 1083,
    "gd_GB": 1084,
    "yi": 1085,
    "ms_MY": 1086,
    "kk_KZ": 1087,
    "ky_KG": 1088,
    "sw_KE": 1089,
    "tk_TM": 1090,
    "uz_Latn_UZ": 1091,
    "tt_RU": 1092,
    "bn_IN": 1093,
    "pa_IN": 1094,
    "gu_IN": 1095,
    "or_IN": 1096,
    "ta_IN": 1097,
    "te_IN": 1098,
    "kn_IN": 1099,
    "ml_IN": 1100,
    "as_IN": 1101,
    "mr_IN": 1102,
    "sa_IN": 1103,
    "mn_MN": 1104,
    "bo_CN": 1105,
    "cy_GB": 1106,
    "km_KH": 1107,
    "lo_LA": 1108,
    "my_MM": 1109,
    "gl_ES": 1110,
    "kok_IN": 1111,
    "mni": 1112,
    "sd_IN": 1113,
    "syr_SY": 1114,
    "si_LK": 1115,
    "chr_US": 1116,
    "iu_Cans_CA": 1117,
    "am_ET": 1118,
    "tmz": 1119,
    "ks_Arab_IN": 1120,
    "ne_NP": 1121,
    "fy_NL": 1122,
    "ps_AF": 1123,
    "fil_PH": 1124,
    "dv_MV": 1125,
    "bin_NG": 1126,
    "fuv_NG": 1127,
    "ha_Latn_NG": 1128,
    "ibb_NG": 1129,
    "yo_NG": 1130,
    "quz_BO": 1131,
    "nso_ZA": 1132,
    "ig_NG": 1136,
    "kr_NG": 1137,
    "gaz_ET": 1138,
    "ti_ER": 1139,
    "gn_PY": 1140,
    "haw_US": 1141,
    "la": 1142,
    "so_SO": 1143,
    "ii_CN": 1144,
    "pap_AN": 1145,
    "ug_Arab_CN": 1152,
    "mi_NZ": 1153,
    "ar_IQ": 2049,
    "zh_CN": 2052,
    "de_CH": 2055,
    "en_GB": 2057,
    "es_MX": 2058,
    "fr_BE": 2060,
    "it_CH": 2064,
    "nl_BE": 2067,
    "nn_NO": 2068,
    "pt_PT": 2070,
    "ro_MD": 2072,
    "ru_MD": 2073,
    "sr_Latn_CS": 2074,
    "sv_FI": 2077,
    "ur_IN": 2080,
    "az_Cyrl_AZ": 2092,
    "ga_IE": 2108,
    "ms_BN": 2110,
    "uz_Cyrl_UZ": 2115,
    "bn_BD": 2117,
    "pa_PK": 2118,
    "mn_Mong_CN": 2128,
    "bo_BT": 2129,
    "sd_PK": 2137,
    "tzm_Latn_DZ": 2143,
    "ks_Deva_IN": 2144,
    "ne_IN": 2145,
    "quz_EC": 2155,
    "ti_ET": 2163,
    "ar_EG": 3073,
    "zh_HK": 3076,
    "de_AT": 3079,
    "en_AU": 3081,
    "fr_CA": 3084,
    "sr_Cyrl_CS": 3098,
    "quz_PE": 3179,
    "ar_LY": 4097,
    "zh_SG": 4100,
    "de_LU": 4103,
    "en_CA": 4105,
    "es_GT": 4106,
    "fr_CH": 4108,
    "hr_BA": 4122,
    "ar_DZ": 5121,
    "zh_MO": 5124,
    "de_LI": 5127,
    "en_NZ": 5129,
    "es_CR": 5130,
    "fr_LU": 5132,
    "bs_Latn_BA": 5146,
    "ar_MO": 6145,
    "en_IE": 6153,
    "es_PA": 6154,
    "fr_MC": 6156,
    "ar_TN": 7169,
    "en_ZA": 7177,
    "es_DO": 7178,
    "fr_029": 7180,
    "ar_OM": 8193,
    "en_JM": 8201,
    "es_VE": 8202,
    "fr_RE": 8204,
    "ar_YE": 9217,
    "en_029": 9225,
    "es_CO": 9226,
    "fr_CG": 9228,
    "ar_SY": 10241,
    "en_BZ": 10249,
    "es_PE": 10250,
    "fr_SN": 10252,
    "ar_JO": 11265,
    "en_TT": 11273,
    "es_AR": 11274,
    "fr_CM": 11276,
    "ar_LB": 12289,
    "en_ZW": 12297,
    "es_EC": 12298,
    "fr_CI": 12300,
    "ar_KW": 13313,
    "en_PH": 13321,
    "es_CL": 13322,
    "fr_ML": 13324,
    "ar_AE": 14337,
    "en_ID": 14345,
    "es_UY": 14346,
    "fr_MA": 14348,
    "ar_BH": 15361,
    "en_HK": 15369,
    "es_PY": 15370,
    "fr_HT": 15372,
    "ar_QA": 16385,
    "en_IN": 16393,
    "es_BO": 16394,
    "en_MY": 17417,
    "es_SV": 17418,
    "en_SG": 18441,
    "es_HN": 18442,
    "es_NI": 19466,
    "es_PR": 20490,
    "es_US": 21514,
    "es_419": 58378,
    "fr_015": 58380
}


def _get_charset_id(encoding):
    encoding = encodings.normalize_encoding(encoding)
    encoding = enc_aliases.get(encoding, encoding)
    return CHARSET_ID_BY_NAME.get(encoding, encoding)


def create_import_filter_options(dialect=csv.unix_dialect, encoding="utf-8",
                                 first_line=1, type_by_col=None,
                                 language_code=locale.getlocale()[0],
                                 detect_numbers=True):
    """
    # See https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options
    @param dialect: the Python csv dialect
    @param encoding: the source file encoding
    @param first_line: the first line
    @param type_by_col: a mappin col -> type (type is FORMAT_MM_DD_YY)
    @param language_code: en_US
    @param detect_numbers: if true, detect numbers
    @return: a filter options string
    """
    charset = _get_charset_id(encoding)
    if type_by_col is None:
        cols = ""
    else:
        cols = "/".join(itertools.chain(*sorted(type_by_col.items())))
    language_id = LANGUAGE_ID_BY_CODE.get(language_code, 0)
    quoting = "true" if dialect.quoting == csv.QUOTE_ALL else "false"
    detect = "true" if detect_numbers else "false"
    tokens = ",".join(
        [str(ord(dialect.delimiter)), str(ord(dialect.quotechar)), str(charset),
         str(first_line), cols, str(language_id), quoting, detect, ""])
    return tokens


def import_from_csv(oDoc, sheet_name, position, path,
                    dialect=csv.unix_dialect,
                    encoding="utf-8", first_line=1, type_by_col=None,
                    language_code=locale.getlocale()[0], detect_numbers=True):
    """
    @param sheet_name: the sheet name
    @param position: position
    @param path: path to the file
    @param dialect: the Python csv dialect
    @param encoding: the source file encoding
    @param first_line: the first line
    @param type_by_col: a mappin col -> type (type is FORMAT_MM_DD_YY)
    @param language_code: en_US
    @param detect_numbers: if true, detect numbers
    @return: the sheet
    """
    filter_options = create_import_filter_options(dialect, encoding, first_line,
                                                  type_by_col, language_code,
                                                  detect_numbers)
    pvs = make_pvs({"FilterName": "Text - txt - csv (StarCalc)",
                                 "FilterOptions": filter_options,
                                 "Hidden": "True"})

    oDoc.lockControllers()
    url = uno.systemPathToFileUrl(os.path.realpath(path))
    oSource = pr.desktop.loadComponentFromURL(url, "_blank", 0, pvs)
    oSource.Sheets.getByIndex(0).Name = sheet_name
    name = oSource.Sheets.getElementNames()[0]
    oDoc.Sheets.importSheet(oSource, name, position)
    oSource.close(True)
    oDoc.unlockControllers()


def create_export_filter_options(dialect=csv.unix_dialect, encoding="utf-8",
                                 first_line=1,
                                 language_code=locale.getlocale()[0]):
    """
    See https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options

    @param dialect: the Python csv dialect
    @param encoding: the source file encoding
    @param first_line: the first line
    @param language_code: en_US
    @return: tokens
    """
    charset = _get_charset_id(encoding)
    language_id = LANGUAGE_ID_BY_CODE.get(language_code, 0)
    quoting = "true" if dialect.quoting == csv.QUOTE_ALL else "false"
    tokens = ",".join(
        [str(ord(dialect.delimiter)), str(ord(dialect.quotechar)), str(charset),
         str(first_line), "", str(language_id), quoting, "true", "true"])
    return tokens


def export_to_csv(oSheet, path, dialect=csv.unix_dialect, encoding="utf-8",
                  first_line=1, language_code=locale.getlocale()[0],
                  overwrite=True):
    """
    save a sheet to a csv file

    @param oSheet:
    @param path:
    @param dialect: the Python csv dialect
    @param encoding: the source file encoding
    @param first_line: the first line
    @param language_code: en_US
    @param overwrite:
    @return: tokens
    """

    filter_options = create_export_filter_options(dialect, encoding, first_line,
                                                  language_code)
    pvs = make_pvs({"FilterName": "Text - txt - csv (StarCalc)",
                                 "FilterOptions": filter_options,
                                 "Overwrite": overwrite})
    oDoc = get_doc(oSheet)
    oActive = oDoc.CurrentController.getActiveSheet()
    oDoc.lockControllers()
    oDoc.CurrentController.setActiveSheet(oSheet)
    url = uno.systemPathToFileUrl(os.path.realpath(path))
    oDoc.storeToURL(url, pvs)
    oDoc.CurrentController.setActiveSheet(oActive)
    oDoc.unlockControllers()
