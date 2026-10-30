# -*- coding: utf-8 -*-
"""Excelの「詳細データ」シートをTSVファイルへ変換します。"""

import calendar
import csv
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation, ROUND_DOWN, ROUND_HALF_UP
from typing import Any, Dict, List, NoReturn, Optional, Tuple

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.styles import Border, Color, Side
from openpyxl.utils import get_column_letter


SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm"}
TARGET_SHEET_NAME = "詳細データ"
WHITESPACE_ONLY_PATTERN = re.compile(r"^[ \u3000\t\r\n]*$")
LINE_BREAK_PATTERN = re.compile(r"\r\n|\r|\n|\t")
COLOR_AND_CONDITION_PATTERN = re.compile(r"\[[^]]*\]")
QUOTED_TEXT_PATTERN = re.compile(r'"([^"]*)"')
GROUPS: List[Tuple[str, str]] = [("1", "契約"), ("2", "直営"), ("4", "フリー")]
VARIETIES: List[Tuple[str, str]] = [
    ("5", "阿波尾鶏"),
    ("6", "赤どり"),
    ("2", "地養鳥"),
    ("1", "ブロイラー"),
]
BROILER_BREAKDOWNS: List[Tuple[str, str]] = [
    ("1", "ブロイラー（貞光分）"),
    ("1", "ブロイラー(外部購入)"),
]
STEP0002_VARIETIES: List[Tuple[str, str]] = VARIETIES + BROILER_BREAKDOWNS
SPLIT_VARIETIES: List[Tuple[str, str]] = VARIETIES[:3] + BROILER_BREAKDOWNS
VARIETY_NAME_PATTERN: str = "(" + "|".join(
    re.escape(pszName) for _, pszName in SPLIT_VARIETIES
) + ")"
AVERAGE_WEIGHT_COLUMN_INDEX = 3
SHIPPING_COUNT_COLUMN_INDEX = 18
AGE_COLUMN_INDEX = 19
SHIPPING_WEIGHT_COLUMN_INDEX = 20
SIMPLE_AVERAGE_COLUMN_INDEXES = {AGE_COLUMN_INDEX}
SUM_COLUMN_INDEXES = {4, 5, 17, 18, *range(20, 64)}
NUMERIC_COLUMN_INDEXES = SIMPLE_AVERAGE_COLUMN_INDEXES | SUM_COLUMN_INDEXES
STEP0003_ITEMS: List[Tuple[str, int, str, bool]] = [
    ("入雛羽数", 17, "入雛羽数", False),
    ("出荷羽数", 18, "出荷羽数", False),
    ("正味重量", 20, "出荷重量", False),
    ("平均体重", 3, "平均体重", True),
    ("懸鳥前死鶏羽数", 24, "懸鳥前死鶏羽数", False),
    ("懸鳥前発育不良羽数", 25, "懸鳥前発育不良羽数", False),
    ("懸鳥前死鶏重量", 44, "懸鳥前死鶏重量", False),
    ("懸鳥前発育不良重量", 45, "懸鳥前発育不良重量", False),
    ("出荷重量", 20, "出荷重量", False),
    ("マレック羽数", 27, "マレック羽数", False),
    ("マレック重量", 47, "マレック重量", False),
    ("大腸菌症重量", 48, "大腸菌症重量", False),
    ("大腸菌症羽数", 28, "大腸菌症羽数", False),
    ("ブドウ球菌症重量", 49, "ブドウ球菌症重量", False),
    ("敗血症重量", 50, "敗血症重量", False),
    ("変性重量", 51, "変性重量", False),
    ("腹水症重量", 52, "腹水症重量", False),
    ("出血重量", 53, "出血重量", False),
    ("炎症重量", 54, "炎症重量", False),
    ("腫瘍重量", 55, "腫瘍重量", False),
    ("黄疸重量", 56, "黄疸重量", False),
    ("発育不良重量", 57, "発育不良重量", False),
    ("放血不良重量", 58, "放血不良重量", False),
    ("湯漬過度重量", 59, "湯漬過度重量", False),
    ("異常鶏重量", 60, "異常鶏重量", False),
    ("予備２重量", 61, "予備２重量", False),
    ("予備３重量", 62, "予備３重量", False),
    ("予備４重量", 63, "予備４重量", False),
    ("ブドウ球菌症羽数", 29, "ブドウ球菌症羽数", False),
    ("敗血症羽数", 30, "敗血症羽数", False),
    ("変性羽数", 31, "変性羽数", False),
    ("腹水症羽数", 32, "腹水症羽数", False),
    ("出血羽数", 33, "出血羽数", False),
    ("炎症羽数", 34, "炎症羽数", False),
    ("腫瘍羽数", 35, "腫瘍羽数", False),
    ("黄疸羽数", 36, "黄疸羽数", False),
    ("発育不良羽数", 37, "発育不良羽数", False),
    ("放血不良羽数", 38, "放血不良羽数", False),
    ("湯漬過度羽数", 39, "湯漬過度羽数", False),
    ("異常鶏羽数", 40, "異常鶏羽数", False),
    ("予備２羽数", 41, "予備２羽数", False),
    ("予備３羽数", 42, "予備３羽数", False),
    ("予備４羽数", 43, "予備４羽数", False),
]
STEP0003_VARIETY_ADDITIONAL_ITEMS: List[Tuple[str, int, str, bool]] = [
    ("農家廃棄羽数", 4, "農家廃棄羽数", False),
    ("農家廃棄重量", 5, "農家廃棄重量", False),
    ("日令", 19, "日令", True),
]
STEP0004_ITEMS: List[Tuple[str, bool]] = [
    ("入雛羽数", False),
    ("出荷羽数", False),
    ("正味重量", False),
    ("平均体重", True),
]
STEP0004_HEADER = "年\t月\tｸﾞﾙｰﾌﾟ№\tｸﾞﾙｰﾌﾟ名\t入雛羽数\t出荷羽数\t正味重量\t平均体重"
STEP0004_VARIETY_ITEMS: List[Tuple[str, str, bool]] = [
    ("入雛羽数", "入雛羽数", False),
    ("出荷羽数", "出荷羽数", False),
    ("正味重量", "正味重量", False),
    ("平均体重", "平均体重", True),
    ("飼育日数平均", "日令", True),
    ("廃棄控除前育成率", "育成率", False),
]
STEP0004_VARIETY_HEADER = (
    "年\t月\t品種No.\t品種名\t入雛羽数\t出荷羽数\t正味重量\t平均体重"
    "\t飼育日数平均\t廃棄控除前育成率"
)
STEP0004_VARIETY_DEDUCTION_ITEMS: List[Tuple[str, str, bool]] = [
    ("入雛羽数", "入雛羽数", False),
    ("処理羽数", "処理羽数", False),
    ("処理重量", "処理重量", False),
    ("控除後平均体重", "平均体重", False),
    ("飼育日数平均", "日令", True),
    ("廃棄控除後育成率", "育成率", False),
]
STEP0004_VARIETY_DEDUCTION_HEADER = (
    "年\t月\t品種No.\t品種名\t入雛羽数\t処理羽数\t処理重量"
    "\t控除後平均体重\t飼育日数平均\t廃棄控除後育成率"
)
STEP0008_VARIETY_ITEM_NAMES: List[str] = [
    pszName for pszName, _, _ in STEP0004_VARIETY_ITEMS[:-1]
] + ["育成率"]
STEP0008_VARIETY_DEDUCTION_ITEM_NAMES: List[str] = [
    pszName for pszName, _, _ in STEP0004_VARIETY_DEDUCTION_ITEMS[:-1]
] + ["育成率"]


STEP0010_GROUP_DEDUCTION_ITEMS: List[Tuple[str, bool]] = [
    ("入雛羽数", False),
    ("処理羽数", False),
    ("処理重量", False),
    ("平均体重", True),
]
STEP0010_GROUP_DEDUCTION_HEADER = (
    "年\t月\tｸﾞﾙｰﾌﾟ№\tｸﾞﾙｰﾌﾟ名\t入雛羽数\t処理羽数\t処理重量\t平均体重"
)
STEP0012_GROUP_DEDUCTION_ROW_STRUCTURE: List[Tuple[str, str]] = [
    ("", "年"),
    ("", "月"),
]
for _, pszGroupName in GROUPS:
    STEP0012_GROUP_DEDUCTION_ROW_STRUCTURE.extend(
        [
            (pszGroupName, "入雛羽数"),
            ("", "処理羽数"),
            ("", "処理重量"),
            ("", "平均体重"),
            ("", "育成率"),
        ]
    )


class ConversionError(Exception):
    def __init__(self, pszErrorCode: str, pszMessage: str) -> None:
        super().__init__(pszMessage)
        self.pszErrorCode = pszErrorCode


def raise_conversion_error(pszErrorCode: str, pszMessage: str) -> NoReturn:
    raise ConversionError(pszErrorCode, pszMessage)


def validate_paths(pszOutputTsvPath: str, pszInputExcelPath: str) -> None:
    if not os.path.isfile(pszInputExcelPath):
        raise_conversion_error(
            "INPUT_FILE_NOT_FOUND",
            "input Excel file was not found. Path = " + pszInputExcelPath,
        )
    if os.path.splitext(pszInputExcelPath)[1].lower() not in SUPPORTED_EXTENSIONS:
        raise_conversion_error(
            "UNSUPPORTED_INPUT_EXTENSION",
            "unsupported Excel extension. Path = " + pszInputExcelPath,
        )
    if os.path.splitext(pszOutputTsvPath)[1].lower() != ".tsv":
        raise_conversion_error(
            "INVALID_OUTPUT_EXTENSION",
            "output file extension must be .tsv. Path = " + pszOutputTsvPath,
        )

    pszOutputDirectory: str = os.path.dirname(pszOutputTsvPath)
    if not os.path.isdir(pszOutputDirectory):
        raise_conversion_error(
            "OUTPUT_DIRECTORY_NOT_FOUND",
            "output directory was not found. Path = " + pszOutputDirectory,
        )
    if os.path.isdir(pszOutputTsvPath):
        raise_conversion_error(
            "INVALID_OUTPUT_PATH",
            "output path points to a directory. Path = " + pszOutputTsvPath,
        )
    if os.path.normcase(pszInputExcelPath) == os.path.normcase(pszOutputTsvPath):
        raise_conversion_error(
            "INPUT_OUTPUT_PATH_CONFLICT",
            "input and output paths must be different. Path = " + pszInputExcelPath,
        )
    if os.path.exists(pszOutputTsvPath):
        raise_conversion_error(
            "OUTPUT_ALREADY_EXISTS",
            "output TSV already exists. Path = " + pszOutputTsvPath,
        )


def create_error_path(pszInputExcelPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputExcelPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszInputExcelPath))[0]
    pszStem: str = pszBaseName + "_error"
    pszErrorPath: str = os.path.join(pszDirectory, pszStem + ".txt")
    iSuffix: int = 2
    while os.path.exists(pszErrorPath):
        pszErrorPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.txt")
        iSuffix += 1
    return pszErrorPath


def write_error_file(
    pszInputExcelPath: str, pszErrorCode: str, pszErrorDetail: str
) -> str:
    while True:
        pszErrorPath: str = create_error_path(pszInputExcelPath)
        try:
            with open(pszErrorPath, "x", encoding="utf-8", newline="") as objErrorFile:
                objErrorFile.write(
                    "ExcelからTSVへの変換、年月別分割、グループ別集計、または"
                    "step0003抽出・step0004結合・step0005年間連結・"
                    "step0006年間行列入れ替え・step0007年間グループ連結・"
                    "step0007年間Excel作成・step0008育成率追加・"
                    "品種別控除廃棄step0008年間TSV作成・"
                    "品種別step0009・step0010年間TSV作成・step0010年間Excel作成・"
                    "控除廃棄品種別step0009・step0010年間TSV作成・"
                    "step0010年間Excel作成・品種別step0010年間TSV作成に"
                    "失敗しました。\r\n"
                )
                objErrorFile.write("エラー種別: " + pszErrorCode + "\r\n")
                objErrorFile.write("入力Excel: " + pszInputExcelPath + "\r\n")
                objErrorFile.write("対象シート: " + TARGET_SHEET_NAME + "\r\n")
                objErrorFile.write("エラー詳細: " + pszErrorDetail + "\r\n")
            return pszErrorPath
        except FileExistsError:
            continue


def normalize_cell_text(pszValue: str) -> str:
    if WHITESPACE_ONLY_PATTERN.fullmatch(pszValue):
        return ""
    return LINE_BREAK_PATTERN.sub(" ", pszValue)


def select_number_format_section(pszNumberFormat: str, objValue: Any) -> str:
    arrSections = pszNumberFormat.split(";")
    if not isinstance(objValue, (int, float, Decimal)) or isinstance(objValue, bool):
        return arrSections[3] if len(arrSections) >= 4 else arrSections[0]
    if objValue > 0:
        return arrSections[0]
    if objValue < 0:
        return arrSections[1] if len(arrSections) >= 2 else arrSections[0]
    return arrSections[2] if len(arrSections) >= 3 else arrSections[0]


def clean_number_format(pszNumberFormat: str) -> str:
    pszFormat: str = COLOR_AND_CONDITION_PATTERN.sub("", pszNumberFormat)
    pszFormat = pszFormat.replace("\\-", "-").replace("\\/", "/")
    pszFormat = pszFormat.replace("\\ ", " ").replace("_", "")
    pszFormat = re.sub(r"\*.", "", pszFormat)
    return pszFormat


def excel_datetime_format_to_strftime(pszNumberFormat: str) -> Tuple[str, bool, bool, bool, bool]:
    pszFormat: str = clean_number_format(pszNumberFormat)
    pszFormat = QUOTED_TEXT_PATTERN.sub(lambda objMatch: objMatch.group(1), pszFormat)
    pszFormat = re.sub(r"\[h\]", "%H", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"AM/PM", "%p", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"yyyy", "%Y", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"yy", "%y", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"mmmm", "%B", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"mmm", "%b", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"dddd", "%A", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"ddd", "%a", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"dd", "%d", pszFormat, flags=re.IGNORECASE)
    pszFormat, iSingleDayCount = re.subn(
        r"(?<!%)d", "{EXCEL_DAY}", pszFormat, flags=re.IGNORECASE
    )
    pszFormat = re.sub(r"hh", "%H", pszFormat, flags=re.IGNORECASE)
    pszFormat, iSingleHourCount = re.subn(
        r"(?<!%)h", "{EXCEL_HOUR}", pszFormat, flags=re.IGNORECASE
    )
    pszFormat = re.sub(r"ss", "%S", pszFormat, flags=re.IGNORECASE)
    pszFormat = re.sub(r"(?<!%)s", "%-S", pszFormat, flags=re.IGNORECASE)

    # Excelのmは、時刻部分では分、それ以外では月を表します。
    bTimeContext: bool = bool(re.search(r"%H|%S", pszFormat))
    pszMonthOrMinute: str = "%M" if bTimeContext else "%m"
    pszFormat = re.sub(r"mm", pszMonthOrMinute, pszFormat, flags=re.IGNORECASE)
    pszFormat, iSingleMonthMinuteCount = re.subn(
        r"(?<!%)m",
        "{EXCEL_MINUTE}" if bTimeContext else "{EXCEL_MONTH}",
        pszFormat,
        flags=re.IGNORECASE,
    )
    return (
        pszFormat,
        iSingleDayCount > 0,
        iSingleHourCount > 0,
        iSingleMonthMinuteCount > 0 and not bTimeContext,
        iSingleMonthMinuteCount > 0 and bTimeContext,
    )


def format_datetime_value(objValue: Any, pszNumberFormat: str) -> str:
    (
        pszFormat,
        bSingleDay,
        bSingleHour,
        bSingleMonth,
        bSingleMinute,
    ) = excel_datetime_format_to_strftime(pszNumberFormat)
    try:
        pszResult: str = objValue.strftime(pszFormat)
        if bSingleDay:
            pszResult = pszResult.replace("{EXCEL_DAY}", str(objValue.day))
        if bSingleHour:
            pszResult = pszResult.replace("{EXCEL_HOUR}", str(objValue.hour))
        if bSingleMonth:
            pszResult = pszResult.replace("{EXCEL_MONTH}", str(objValue.month))
        if bSingleMinute:
            pszResult = pszResult.replace("{EXCEL_MINUTE}", str(objValue.minute))
        return pszResult
    except (ValueError, TypeError):
        if isinstance(objValue, datetime):
            return objValue.isoformat(sep=" ")
        if isinstance(objValue, (date, time)):
            return objValue.isoformat()
        if isinstance(objValue, timedelta):
            return str(objValue)
        return str(objValue)


def decimal_places_from_format(pszNumberFormat: str) -> int:
    pszFormat: str = clean_number_format(pszNumberFormat)
    objMatch = re.search(r"[0#?]+\.([0#?]+)", pszFormat)
    return len(objMatch.group(1)) if objMatch else 0


def format_numeric_value(objValue: Any, pszNumberFormat: str) -> str:
    pszSection: str = clean_number_format(
        select_number_format_section(pszNumberFormat, objValue)
    )
    if pszSection.lower() == "general":
        return str(objValue)

    bPercent: bool = "%" in pszSection
    bThousands: bool = bool(re.search(r"[0#?],[0#?]", pszSection))
    iDecimalPlaces: int = decimal_places_from_format(pszSection)
    objDecimal = Decimal(str(objValue))
    if bPercent:
        objDecimal *= Decimal("100")

    pszIntegerPattern: str = pszSection.split(".", 1)[0]
    pszIntegerPattern = re.sub(r'"[^"]*"', "", pszIntegerPattern)
    pszIntegerPattern = re.sub(r"[^0#?]", "", pszIntegerPattern)
    iMinimumIntegerDigits: int = pszIntegerPattern.count("0")

    try:
        objQuantum = Decimal(1).scaleb(-iDecimalPlaces)
        objRounded = objDecimal.quantize(objQuantum, rounding=ROUND_HALF_UP)
    except InvalidOperation:
        return str(objValue)

    pszSpecification: str = ("," if bThousands else "") + "." + str(iDecimalPlaces) + "f"
    pszResult: str = format(abs(objRounded), pszSpecification)
    if iMinimumIntegerDigits > 1 and not bThousands:
        pszInteger, pszSeparator, pszFraction = pszResult.partition(".")
        pszInteger = pszInteger.zfill(iMinimumIntegerDigits)
        pszResult = pszInteger + (pszSeparator + pszFraction if pszSeparator else "")
    if objDecimal < 0:
        if "(" in pszSection and ")" in pszSection:
            pszResult = "(" + pszResult + ")"
        else:
            pszResult = "-" + pszResult
    if bPercent:
        pszResult += "%"

    arrLiteralTexts = QUOTED_TEXT_PATTERN.findall(pszSection)
    if arrLiteralTexts:
        pszResult = "".join(arrLiteralTexts[:1]) + pszResult + "".join(arrLiteralTexts[1:])
    return pszResult


def format_cell_value(objCell: Cell, objValue: Any) -> str:
    if objValue is None:
        return ""
    if isinstance(objValue, str):
        return normalize_cell_text(objValue)
    if isinstance(objValue, bool):
        return "TRUE" if objValue else "FALSE"
    if isinstance(objValue, (datetime, date, time, timedelta)):
        return normalize_cell_text(format_datetime_value(objValue, objCell.number_format))
    if isinstance(objValue, (int, float, Decimal)):
        return normalize_cell_text(format_numeric_value(objValue, objCell.number_format))
    return normalize_cell_text(str(objValue))


def has_data_value(objValue: Any) -> bool:
    if objValue is None:
        return False
    if isinstance(objValue, str):
        return normalize_cell_text(objValue) != ""
    return True


def find_data_bounds(
    objFormulaWorksheet: Any, objValueWorksheet: Any
) -> Optional[Tuple[int, int]]:
    iLastRow: int = 0
    for iRow in range(2, objFormulaWorksheet.max_row + 1):
        if all(
            has_data_value(objValueWorksheet.cell(iRow, iColumn).value)
            for iColumn in range(1, 4)
        ):
            iLastRow = iRow

    if iLastRow == 0:
        return None

    iLastColumn: int = 0
    for objRow in objFormulaWorksheet.iter_rows(
        min_row=1, max_row=iLastRow
    ):
        for objCell in objRow:
            if objCell.value is not None:
                iLastColumn = max(iLastColumn, objCell.column)
    if iLastColumn == 0:
        return None
    return iLastRow, iLastColumn


def reserve_output_path(pszOutputTsvPath: str) -> None:
    try:
        iFileDescriptor: int = os.open(
            pszOutputTsvPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
        )
        os.close(iFileDescriptor)
    except FileExistsError:
        raise_conversion_error(
            "OUTPUT_ALREADY_EXISTS",
            "output TSV already exists. Path = " + pszOutputTsvPath,
        )
    except OSError as objException:
        raise_conversion_error(
            "TSV_TEMP_CREATE_FAILED",
            "failed to reserve output TSV. Detail = " + str(objException),
        )


def strip_line_ending(pszLine: str) -> str:
    if pszLine.endswith("\r\n"):
        return pszLine[:-2]
    if pszLine.endswith("\r") or pszLine.endswith("\n"):
        return pszLine[:-1]
    return pszLine


def parse_year_month(pszYear: str, pszMonth: str, iLineNumber: int) -> Tuple[int, int]:
    pszNormalizedYear: str = pszYear.strip(" \u3000\t")
    pszNormalizedMonth: str = pszMonth.strip(" \u3000\t")
    try:
        objYear = Decimal(pszNormalizedYear)
        objMonth = Decimal(pszNormalizedMonth)
        if not objYear.is_finite() or objYear != objYear.to_integral_value():
            raise InvalidOperation
        if not objMonth.is_finite() or objMonth != objMonth.to_integral_value():
            raise InvalidOperation
        iYear: int = int(objYear)
        iMonth: int = int(objMonth)
    except (InvalidOperation, ValueError, OverflowError):
        raise_conversion_error(
            "INVALID_YEAR_MONTH",
            "invalid year or month in TSV. Line = "
            + str(iLineNumber)
            + ", Year = "
            + pszYear
            + ", Month = "
            + pszMonth,
        )

    if iYear < 1 or iYear > 9999 or iMonth < 1 or iMonth > 12:
        raise_conversion_error(
            "INVALID_YEAR_MONTH",
            "invalid year or month in TSV. Line = "
            + str(iLineNumber)
            + ", Year = "
            + pszYear
            + ", Month = "
            + pszMonth,
        )
    return iYear, iMonth


def read_tsv_by_month(
    pszSourceTsvPath: str,
) -> Tuple[str, Dict[Tuple[int, int], List[str]]]:
    try:
        with open(pszSourceTsvPath, "r", encoding="utf-8", newline="") as objTsvFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objTsvFile]
    except Exception as objException:
        raise_conversion_error(
            "SOURCE_TSV_READ_FAILED",
            "failed to read source TSV. Detail = " + str(objException),
        )

    if not arrLines or arrLines[0] == "":
        raise_conversion_error("SOURCE_TSV_HEADER_MISSING", "source TSV header was missing.")
    if len(arrLines[0].split("\t")) < 2:
        raise_conversion_error(
            "SOURCE_TSV_COLUMN_MISSING",
            "source TSV must contain year and month columns.",
        )

    dicRowsByMonth: Dict[Tuple[int, int], List[str]] = {}
    for iLineNumber, pszLine in enumerate(arrLines[1:], start=2):
        arrFields: List[str] = pszLine.split("\t")
        if len(arrFields) < 2:
            raise_conversion_error(
                "SOURCE_TSV_COLUMN_MISSING",
                "source TSV row must contain year and month columns. Line = "
                + str(iLineNumber),
            )
        objYearMonth: Tuple[int, int] = parse_year_month(
            arrFields[0], arrFields[1], iLineNumber
        )
        dicRowsByMonth.setdefault(objYearMonth, []).append(pszLine)

    if not dicRowsByMonth:
        raise_conversion_error(
            "INVALID_YEAR_MONTH", "source TSV did not contain any data rows."
        )
    return arrLines[0], dicRowsByMonth


def create_monthly_output_path(
    pszSourceTsvPath: str, iYear: int, iMonth: int
) -> str:
    pszDirectory: str = os.path.dirname(pszSourceTsvPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem: str = f"{pszBaseName}_step0001_{iYear:04d}_{iMonth:02d}"
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def write_monthly_tsv(
    pszOutputPath: str, pszHeader: str, arrRows: List[str]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "MONTHLY_OUTPUT_ALREADY_EXISTS",
                "monthly TSV already exists. Path = " + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "MONTHLY_TSV_CREATE_FAILED",
                "failed to reserve monthly TSV. Detail = " + str(objException),
            )

        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_monthly_",
                suffix=".tsv",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            with open(pszTemporaryPath, "w", encoding="utf-8", newline="") as objFile:
                objFile.write(pszHeader + "\r\n")
                for pszRow in arrRows:
                    objFile.write(pszRow + "\r\n")
        except Exception as objException:
            raise_conversion_error(
                "MONTHLY_TSV_WRITE_FAILED",
                "failed to write monthly TSV. Detail = " + str(objException),
            )

        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "MONTHLY_TSV_FINALIZE_FAILED",
                "failed to finalize monthly TSV. Detail = " + str(objException),
            )
    finally:
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def split_tsv_by_month(
    pszSourceTsvPath: str,
) -> Tuple[List[str], Tuple[int, int], Tuple[int, int]]:
    pszHeader, dicRowsByMonth = read_tsv_by_month(pszSourceTsvPath)
    arrYearMonths: List[Tuple[int, int]] = sorted(dicRowsByMonth)
    arrOutputPlans: List[Tuple[str, List[str]]] = [
        (
            create_monthly_output_path(pszSourceTsvPath, iYear, iMonth),
            dicRowsByMonth[(iYear, iMonth)],
        )
        for iYear, iMonth in arrYearMonths
    ]
    arrCreatedPaths: List[str] = []
    try:
        for pszOutputPath, arrRows in arrOutputPlans:
            write_monthly_tsv(pszOutputPath, pszHeader, arrRows)
            arrCreatedPaths.append(pszOutputPath)
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszCreatedPath in arrCreatedPaths:
            try:
                os.remove(pszCreatedPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(
                    pszCreatedPath + ": " + str(objCleanupException)
                )
        if arrCleanupErrors:
            raise_conversion_error(
                "MONTHLY_OUTPUT_CLEANUP_FAILED",
                "monthly TSV creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise
    return arrCreatedPaths, arrYearMonths[0], arrYearMonths[-1]


def create_daily_operation_files(arrMonthlyTsvPaths: List[str]) -> List[str]:
    """step0001の処理日から、年別の33行×13列の稼動判定表を作成します。

    最終年月より後の月は未提供として空欄にします。それ以前の不足月や
    入力不備がある年は出力せず、正常な年を出力した後にエラーを報告します。
    """
    dicInputs: Dict[Tuple[str, str, int], Dict[int, str]] = {}
    arrErrors: List[str] = []
    for pszPath in arrMonthlyTsvPaths:
        objMatch = re.fullmatch(
            r"(.+)_step0001_([0-9]{4})_([0-9]{2})(?:_[0-9]+)?\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error("DAILY_INVALID_FILE_NAME", "Path = " + pszPath)
        pszPrefix, pszYear, pszMonth = objMatch.groups()
        iYear, iMonth = int(pszYear), int(pszMonth)
        if not 1 <= iYear <= 9999 or not 1 <= iMonth <= 12:
            raise_conversion_error("DAILY_INVALID_YEAR_MONTH", "Path = " + pszPath)
        objKey = (os.path.dirname(os.path.abspath(pszPath)), pszPrefix, iYear)
        dicMonths = dicInputs.setdefault(objKey, {})
        if iMonth in dicMonths:
            raise_conversion_error("DAILY_DUPLICATE_MONTH", "Path = " + pszPath)
        dicMonths[iMonth] = pszPath

    dicLastMonths: Dict[Tuple[str, str], Tuple[int, int]] = {}
    for (pszDirectory, pszPrefix, iYear), dicMonths in dicInputs.items():
        objSeries = (pszDirectory, pszPrefix)
        dicLastMonths[objSeries] = max(
            dicLastMonths.get(objSeries, (0, 0)), (iYear, max(dicMonths))
        )

    # 年自体が抜けている場合も、不足として検出します。
    for objSeries, (iLastYear, _) in dicLastMonths.items():
        iFirstYear = min(k[2] for k in dicInputs if k[:2] == objSeries)
        for iYear in range(iFirstYear, iLastYear + 1):
            dicInputs.setdefault((*objSeries, iYear), {})

    arrCreatedPaths: List[str] = []
    for (pszDirectory, pszPrefix, iYear), dicMonths in sorted(dicInputs.items()):
        arrRows: List[List[str]] = [[""] * 13 for _ in range(33)]
        arrRows[0] = [f"{iYear}年"] + [f"{m}月" for m in range(1, 13)]
        arrRows[1][0] = "稼動日数"
        for iDay in range(1, 32):
            arrRows[1 + iDay][0] = str(iDay)
        iLastYear, iLastMonth = dicLastMonths[(pszDirectory, pszPrefix)]
        iExpectedLastMonth = iLastMonth if iYear == iLastYear else 12
        arrYearErrors: List[str] = []
        for iMonth in range(1, iExpectedLastMonth + 1):
            pszPath = dicMonths.get(iMonth)
            if pszPath is None:
                arrYearErrors.append(f"入力ファイル不足: {pszPrefix}_step0001_{iYear:04d}_{iMonth:02d}.tsv")
                continue
            try:
                setDays = set()
                with open(pszPath, "r", encoding="utf-8-sig", newline="") as objFile:
                    objReader = csv.reader(objFile, delimiter="\t")
                    arrHeader = next(objReader, [])
                    if len(arrHeader) < 15 or arrHeader[14] != "処理日":
                        raise ValueError("1行目: O列の見出しが「処理日」ではありません。")
                    for arrFields in objReader:
                        iLine = objReader.line_num
                        pszDate = arrFields[14] if len(arrFields) >= 15 else ""
                        if re.fullmatch(r"[0-9]{8}", pszDate) is None:
                            raise ValueError(f"{iLine}行目: 不正な処理日 = {pszDate!r}")
                        try:
                            objDate = date(int(pszDate[:4]), int(pszDate[4:6]), int(pszDate[6:]))
                        except ValueError:
                            raise ValueError(f"{iLine}行目: 不正な処理日 = {pszDate!r}")
                        if (objDate.year, objDate.month) != (iYear, iMonth):
                            raise ValueError(f"{iLine}行目: ファイル名と処理日の年月が不一致 = {pszDate}")
                        setDays.add(objDate.day)
                iMonthDays = calendar.monthrange(iYear, iMonth)[1]
                for iDay in range(1, 32):
                    arrRows[1 + iDay][iMonth] = str(int(iDay <= iMonthDays and iDay in setDays))
                arrRows[1][iMonth] = str(len(setDays))
            except (OSError, UnicodeError, csv.Error, ValueError) as objException:
                arrYearErrors.append(pszPath + ": " + str(objException))
        if arrYearErrors:
            arrErrors.extend(arrYearErrors)
            continue

        pszOutputPath = os.path.join(
            pszDirectory, f"{pszPrefix}_step0001_日別稼動判定_{iYear:04d}.tsv"
        )
        bReserved = False
        try:
            with open(pszOutputPath, "x", encoding="utf-8-sig", newline="") as objFile:
                bReserved = True
                csv.writer(objFile, delimiter="\t", lineterminator="\r\n").writerows(arrRows)
            arrCreatedPaths.append(pszOutputPath)
        except OSError as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszOutputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("DAILY_OPERATION_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def create_step0002_output_path(
    pszStep0001Path: str, bVariety: bool = True
) -> str:
    pszDirectory: str = os.path.dirname(pszStep0001Path)
    pszFileName: str = os.path.basename(pszStep0001Path)
    if "_step0001_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0001_FILE_NAME",
            "step0001 marker was not found. Path = " + pszStep0001Path,
        )
    pszCategory: str = "品種別" if bVariety else "グループ別"
    pszStem: str = os.path.splitext(
        pszFileName.replace("_step0001_", f"_step0002_{pszCategory}_", 1)
    )[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def create_step0002_total_output_path(pszStep0001Path: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0001Path)
    pszFileName: str = os.path.basename(pszStep0001Path)
    if "_step0001_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0001_FILE_NAME",
            "step0001 marker was not found. Path = " + pszStep0001Path,
        )
    pszStem: str = os.path.splitext(pszFileName.replace("_step0001_", "_step0002_総計_", 1))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def parse_decimal_value(
    pszValue: str, iLineNumber: int, iColumnIndex: int, pszColumnName: str
) -> Optional[Decimal]:
    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
    if pszNormalizedValue == "":
        return None
    try:
        objValue = Decimal(pszNormalizedValue)
    except InvalidOperation:
        objValue = Decimal("NaN")
    if not objValue.is_finite():
        raise_conversion_error(
            "NON_NUMERIC_AGGREGATION_VALUE",
            "non-numeric aggregation value. Line = "
            + str(iLineNumber)
            + ", Column = "
            + str(iColumnIndex + 1)
            + ", Name = "
            + pszColumnName
            + ", Value = "
            + pszValue,
        )
    return objValue


def format_sum_value(objValue: Decimal) -> str:
    pszValue: str = format(objValue, "f")
    if "." in pszValue:
        pszValue = pszValue.rstrip("0").rstrip(".")
    return pszValue if pszValue not in {"", "-0"} else "0"


def parse_process_month(pszValue: str, iLineNumber: int) -> Tuple[int, int, str]:
    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
    try:
        objValue = Decimal(pszNormalizedValue)
        if not objValue.is_finite() or objValue != objValue.to_integral_value():
            raise InvalidOperation
        pszDate: str = str(int(objValue))
        if len(pszDate) != 8:
            raise ValueError
        objDate = datetime.strptime(pszDate, "%Y%m%d")
    except (InvalidOperation, ValueError, OverflowError):
        raise_conversion_error(
            "INVALID_PROCESS_DATE",
            "invalid process date. Line = "
            + str(iLineNumber)
            + ", Value = "
            + pszValue,
        )
    return objDate.year, objDate.month, pszDate[:6]


def add_warning(
    arrWarnings: List[Dict[str, str]],
    pszType: str,
    pszInputPath: str,
    pszOutputPath: str,
    pszProcessMonth: str,
    pszGroupNumber: str,
    pszGroupName: str,
    pszColumns: str,
    pszReason: str,
    pszOutputValue: str,
) -> None:
    arrWarnings.append(
        {
            "type": pszType,
            "input": pszInputPath,
            "output": pszOutputPath,
            "month": pszProcessMonth,
            "group_number": pszGroupNumber,
            "group_name": pszGroupName,
            "columns": pszColumns,
            "reason": pszReason,
            "output_value": pszOutputValue,
        }
    )


def write_step0002_tsv(
    pszOutputPath: str, pszHeader: str, arrRows: List[str]
) -> None:
    try:
        write_monthly_tsv(pszOutputPath, pszHeader, arrRows)
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0002_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0002_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0002_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0002_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0002_WRITE_FAILED"),
            str(objException),
        )


def aggregate_step0001(
    pszInputPath: str,
    pszOutputPath: str,
    arrWarnings: List[Dict[str, str]],
    bVariety: bool = True,
) -> None:
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0001_READ_FAILED", "failed to read step0001 TSV. Detail = " + str(objException)
        )
    if not arrLines:
        raise_conversion_error("STEP0001_HEADER_MISSING", "step0001 header was missing.")
    arrHeader: List[str] = arrLines[0].split("\t")
    if len(arrHeader) != 64:
        raise_conversion_error(
            "INVALID_AGGREGATION_COLUMN_COUNT",
            "step0001 header must contain 64 columns. Actual = " + str(len(arrHeader)),
        )

    arrCategories = VARIETIES if bVariety else GROUPS
    iNumberColumnIndex, iNameColumnIndex = (15, 16) if bVariety else (10, 11)
    pszCategoryCode = "VARIETY" if bVariety else "GROUP"
    pszCategoryName = "variety" if bVariety else "group"
    dicRows: Dict[str, List[Tuple[int, List[str]]]] = {
        pszVarietyName: [] for _, pszVarietyName in arrCategories
    }
    objYearMonth: Optional[Tuple[int, int]] = None
    for iLineNumber, pszLine in enumerate(arrLines[1:], start=2):
        arrFields: List[str] = pszLine.split("\t")
        if len(arrFields) != 64:
            raise_conversion_error(
                "INVALID_AGGREGATION_COLUMN_COUNT",
                "step0001 row must contain 64 columns. Line = "
                + str(iLineNumber)
                + ", Actual = "
                + str(len(arrFields)),
            )
        iYear, iMonth = parse_year_month(arrFields[0], arrFields[1], iLineNumber)
        iProcessYear, iProcessMonth, _ = parse_process_month(arrFields[14], iLineNumber)
        if (iYear, iMonth) != (iProcessYear, iProcessMonth):
            raise_conversion_error(
                "YEAR_MONTH_MISMATCH",
                "year/month and process date do not match. Line = " + str(iLineNumber),
            )
        if objYearMonth is None:
            objYearMonth = (iYear, iMonth)
        elif objYearMonth != (iYear, iMonth):
            raise_conversion_error(
                "YEAR_MONTH_MISMATCH", "multiple year/month values in step0001 TSV."
            )

        pszVarietyNumber: str = arrFields[iNumberColumnIndex].strip(" \u3000\t")
        pszVarietyName: str = arrFields[iNameColumnIndex].strip(" \u3000\t")
        dicExpectedVarieties: Dict[str, str] = {
            pszName: pszNumber for pszNumber, pszName in arrCategories
        }
        if pszVarietyName not in dicExpectedVarieties:
            raise_conversion_error(
                "INVALID_" + pszCategoryCode + "_NAME",
                "unsupported " + pszCategoryName + " name. Line = "
                + str(iLineNumber)
                + ", Value = "
                + pszVarietyName,
            )
        if pszVarietyNumber != dicExpectedVarieties[pszVarietyName]:
            raise_conversion_error(
                "INVALID_" + pszCategoryCode + "_MAPPING",
                pszCategoryName + " number and name do not match. Line = " + str(iLineNumber),
            )
        for iColumnIndex in NUMERIC_COLUMN_INDEXES:
            parse_decimal_value(
                arrFields[iColumnIndex], iLineNumber, iColumnIndex, arrHeader[iColumnIndex]
            )
        dicRows[pszVarietyName].append((iLineNumber, arrFields))

    if objYearMonth is None:
        raise_conversion_error("STEP0001_DATA_MISSING", "step0001 did not contain data rows.")
    iYear, iMonth = objYearMonth
    pszProcessMonth: str = f"{iYear:04d}{iMonth:02d}"
    if bVariety:
        # Preserve the original four rows, then append the two broiler details.
        for _, pszName in BROILER_BREAKDOWNS:
            arrGroupNumbers = (
                {"1", "2"} if pszName == BROILER_BREAKDOWNS[0][1] else {"4"}
            )
            dicRows[pszName] = [
                (iLineNumber, arrFields)
                for iLineNumber, arrFields in dicRows["ブロイラー"]
                if arrFields[10].strip(" \u3000\t") in arrGroupNumbers
            ]
        arrCategories = STEP0002_VARIETIES
    arrOutputRows: List[str] = []
    for pszVarietyNumber, pszVarietyName in arrCategories:
        arrGroupRows = dicRows[pszVarietyName]
        arrOutput: List[str] = ["－"] * 64
        arrOutput[0] = f"{iYear:04d}"
        arrOutput[1] = f"{iMonth:02d}"
        arrOutput[10] = "－"
        arrOutput[11] = "－"
        arrOutput[14] = pszProcessMonth
        arrOutput[iNumberColumnIndex] = pszVarietyNumber
        arrOutput[iNameColumnIndex] = pszVarietyName

        if not arrGroupRows:
            arrOutput[3] = ""
            arrOutput[19] = ""
            for iColumnIndex in SUM_COLUMN_INDEXES:
                arrOutput[iColumnIndex] = "0"
            add_warning(
                arrWarnings,
                "MISSING_" + pszCategoryCode + "_DATA",
                pszInputPath,
                pszOutputPath,
                pszProcessMonth,
                pszVarietyNumber,
                pszVarietyName,
                "4(D), 5(E), 6(F), 18(R)-64(BL)",
                pszCategoryName + " contains no input rows.",
                "average columns are empty; sum columns are 0",
            )
        else:
            dicColumnValues: Dict[int, List[Decimal]] = {}
            for iColumnIndex in sorted(NUMERIC_COLUMN_INDEXES):
                arrValues: List[Decimal] = []
                for iLineNumber, arrFields in arrGroupRows:
                    objValue = parse_decimal_value(
                        arrFields[iColumnIndex],
                        iLineNumber,
                        iColumnIndex,
                        arrHeader[iColumnIndex],
                    )
                    if objValue is not None:
                        arrValues.append(objValue)
                dicColumnValues[iColumnIndex] = arrValues
                if not arrValues:
                    bAverage: bool = iColumnIndex in SIMPLE_AVERAGE_COLUMN_INDEXES
                    arrOutput[iColumnIndex] = "" if bAverage else "0"
                    add_warning(
                        arrWarnings,
                        "ALL_BLANK_AVERAGE_COLUMN" if bAverage else "ALL_BLANK_SUM_COLUMN",
                        pszInputPath,
                        pszOutputPath,
                        pszProcessMonth,
                        pszVarietyNumber,
                        pszVarietyName,
                        str(iColumnIndex + 1)
                        + "("
                        + get_column_letter(iColumnIndex + 1)
                        + ") "
                        + arrHeader[iColumnIndex],
                        "column contains no numeric values.",
                        "empty" if bAverage else "0",
                    )
                elif iColumnIndex in SIMPLE_AVERAGE_COLUMN_INDEXES:
                    objAverage = sum(arrValues, Decimal("0")) / Decimal(len(arrValues))
                    arrOutput[iColumnIndex] = format(
                        objAverage.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), ".2f"
                    )
                else:
                    arrOutput[iColumnIndex] = format_sum_value(
                        sum(arrValues, Decimal("0"))
                    )

            arrShippingCounts: List[Decimal] = dicColumnValues[
                SHIPPING_COUNT_COLUMN_INDEX
            ]
            objShippingCountTotal: Decimal = sum(
                arrShippingCounts, Decimal("0")
            )
            if not arrShippingCounts:
                arrOutput[AVERAGE_WEIGHT_COLUMN_INDEX] = ""
            elif objShippingCountTotal == 0:
                arrOutput[AVERAGE_WEIGHT_COLUMN_INDEX] = ""
                add_warning(
                    arrWarnings,
                    "AVERAGE_WEIGHT_DIVISION_BY_ZERO",
                    pszInputPath,
                    pszOutputPath,
                    pszProcessMonth,
                    pszVarietyNumber,
                    pszVarietyName,
                    "4(D) 平均体重, 19(S) 出荷羽数, 21(U) 出荷重量",
                    "shipping count total is zero; average weight cannot be calculated.",
                    "empty",
                )
            else:
                objShippingWeightTotal: Decimal = sum(
                    dicColumnValues[SHIPPING_WEIGHT_COLUMN_INDEX], Decimal("0")
                )
                objAverageWeight: Decimal = (
                    objShippingWeightTotal / objShippingCountTotal
                )
                arrOutput[AVERAGE_WEIGHT_COLUMN_INDEX] = format(
                    objAverageWeight.quantize(
                        Decimal("0.00001"), rounding=ROUND_HALF_UP
                    ),
                    ".5f",
                )
        arrOutputRows.append("\t".join(arrOutput))
    write_step0002_tsv(pszOutputPath, "\t".join(arrHeader), arrOutputRows)


def create_warning_path(pszInputExcelPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszInputExcelPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszInputExcelPath))[0]
    pszStem: str = pszBaseName + "_warning"
    pszPath: str = os.path.join(pszDirectory, pszStem + ".txt")
    iSuffix: int = 2
    while os.path.exists(pszPath):
        pszPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.txt")
        iSuffix += 1
    return pszPath


def write_warning_file(
    pszInputExcelPath: str, arrWarnings: List[Dict[str, str]]
) -> str:
    pszWarningPath: str = create_warning_path(pszInputExcelPath)
    try:
        with open(pszWarningPath, "x", encoding="utf-8", newline="") as objFile:
            objFile.write("月次・グループ別・品種別・総計集計が警告付きで完了しました。\r\n")
            objFile.write("入力Excel: " + pszInputExcelPath + "\r\n")
            objFile.write("警告件数: " + str(len(arrWarnings)) + "\r\n")
            for iIndex, dicWarning in enumerate(arrWarnings, start=1):
                objFile.write("\r\n[警告 " + str(iIndex) + "]\r\n")
                objFile.write("警告種別: " + dicWarning["type"] + "\r\n")
                objFile.write("入力step0001: " + dicWarning["input"] + "\r\n")
                objFile.write("出力step0002: " + dicWarning["output"] + "\r\n")
                objFile.write("処理年月: " + dicWarning["month"] + "\r\n")
                if dicWarning["group_name"] == "総計":
                    objFile.write("集計対象: 総計\r\n")
                elif dicWarning["group_name"] in {pszName for _, pszName in GROUPS}:
                    objFile.write("グループ№: " + dicWarning["group_number"] + "\r\n")
                    objFile.write("グループ名: " + dicWarning["group_name"] + "\r\n")
                else:
                    objFile.write("品種No.: " + dicWarning["group_number"] + "\r\n")
                    objFile.write("品種名: " + dicWarning["group_name"] + "\r\n")
                objFile.write("対象列: " + dicWarning["columns"] + "\r\n")
                objFile.write("警告理由: " + dicWarning["reason"] + "\r\n")
                objFile.write("出力値: " + dicWarning["output_value"] + "\r\n")
        return pszWarningPath
    except Exception as objException:
        if os.path.exists(pszWarningPath):
            os.remove(pszWarningPath)
        raise_conversion_error(
            "WARNING_FILE_CREATE_FAILED",
            "failed to create warning file. Detail = " + str(objException),
        )


def aggregate_step0001_files(
    arrStep0001Paths: List[str],
    pszInputExcelPath: str,
    bVariety: bool = True,
    pszExistingWarningPath: str = "",
    iExistingWarningCount: int = 0,
) -> Tuple[List[str], str, int]:
    arrStep0002Paths: List[str] = []
    arrWarnings: List[Dict[str, str]] = []
    try:
        for pszStep0001Path in arrStep0001Paths:
            pszStep0002Path: str = create_step0002_output_path(pszStep0001Path, bVariety)
            aggregate_step0001(pszStep0001Path, pszStep0002Path, arrWarnings, bVariety)
            arrStep0002Paths.append(pszStep0002Path)
        pszWarningPath: str = pszExistingWarningPath
        if not pszExistingWarningPath:
            if arrWarnings:
                pszWarningPath = write_warning_file(pszInputExcelPath, arrWarnings)
        else:
            pszWarningPath = write_step0002_total_warnings(
                pszInputExcelPath, pszExistingWarningPath, iExistingWarningCount, arrWarnings
            )
        return arrStep0002Paths, pszWarningPath, iExistingWarningCount + len(arrWarnings)
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrStep0002Paths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0002_CLEANUP_FAILED",
                "step0002 aggregation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def aggregate_step0001_total(
    pszInputPath: str, pszOutputPath: str, arrWarnings: List[Dict[str, str]]
) -> None:
    """step0001の全明細から、丸め前の値で月次総計を計算します。"""
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0001_READ_FAILED", "failed to read step0001 TSV. Detail = " + str(objException)
        )
    if not arrLines:
        raise_conversion_error("STEP0001_HEADER_MISSING", "step0001 header was missing.")
    arrHeader: List[str] = arrLines[0].split("\t")
    if len(arrHeader) != 64:
        raise_conversion_error(
            "INVALID_AGGREGATION_COLUMN_COUNT",
            "step0001 header must contain 64 columns. Actual = " + str(len(arrHeader)),
        )

    arrTotalRows: List[Tuple[int, List[str]]] = []
    objYearMonth: Optional[Tuple[int, int]] = None
    for iLineNumber, pszLine in enumerate(arrLines[1:], start=2):
        arrFields: List[str] = pszLine.split("\t")
        if len(arrFields) != 64:
            raise_conversion_error(
                "INVALID_AGGREGATION_COLUMN_COUNT",
                "step0001 row must contain 64 columns. Line = "
                + str(iLineNumber)
                + ", Actual = "
                + str(len(arrFields)),
            )
        iYear, iMonth = parse_year_month(arrFields[0], arrFields[1], iLineNumber)
        iProcessYear, iProcessMonth, _ = parse_process_month(arrFields[14], iLineNumber)
        if (iYear, iMonth) != (iProcessYear, iProcessMonth):
            raise_conversion_error(
                "YEAR_MONTH_MISMATCH",
                "year/month and process date do not match. Line = " + str(iLineNumber),
            )
        if objYearMonth is None:
            objYearMonth = (iYear, iMonth)
        elif objYearMonth != (iYear, iMonth):
            raise_conversion_error(
                "YEAR_MONTH_MISMATCH", "multiple year/month values in step0001 TSV."
            )

        for iColumnIndex in NUMERIC_COLUMN_INDEXES:
            parse_decimal_value(
                arrFields[iColumnIndex], iLineNumber, iColumnIndex, arrHeader[iColumnIndex]
            )
        arrTotalRows.append((iLineNumber, arrFields))

    if objYearMonth is None:
        raise_conversion_error("STEP0001_DATA_MISSING", "step0001 did not contain data rows.")
    iYear, iMonth = objYearMonth
    pszProcessMonth: str = f"{iYear:04d}{iMonth:02d}"
    pszGroupNumber: str = ""
    pszGroupName: str = "総計"
    arrGroupRows = arrTotalRows
    arrOutput: List[str] = ["－"] * 64
    arrOutput[0] = f"{iYear:04d}"
    arrOutput[1] = f"{iMonth:02d}"
    arrOutput[10] = pszGroupNumber
    arrOutput[11] = pszGroupName
    arrOutput[14] = pszProcessMonth

    dicColumnValues: Dict[int, List[Decimal]] = {}
    for iColumnIndex in sorted(NUMERIC_COLUMN_INDEXES):
        arrValues: List[Decimal] = []
        for iLineNumber, arrFields in arrGroupRows:
            objValue = parse_decimal_value(
                arrFields[iColumnIndex],
                iLineNumber,
                iColumnIndex,
                arrHeader[iColumnIndex],
            )
            if objValue is not None:
                arrValues.append(objValue)
        dicColumnValues[iColumnIndex] = arrValues
        if not arrValues:
            bAverage: bool = iColumnIndex in SIMPLE_AVERAGE_COLUMN_INDEXES
            arrOutput[iColumnIndex] = "" if bAverage else "0"
            add_warning(
                arrWarnings,
                "ALL_BLANK_AVERAGE_COLUMN" if bAverage else "ALL_BLANK_SUM_COLUMN",
                pszInputPath,
                pszOutputPath,
                pszProcessMonth,
                pszGroupNumber,
                pszGroupName,
                str(iColumnIndex + 1)
                + "("
                + get_column_letter(iColumnIndex + 1)
                + ") "
                + arrHeader[iColumnIndex],
                "column contains no numeric values.",
                "empty" if bAverage else "0",
            )
        elif iColumnIndex in SIMPLE_AVERAGE_COLUMN_INDEXES:
            objAverage = sum(arrValues, Decimal("0")) / Decimal(len(arrValues))
            arrOutput[iColumnIndex] = format(
                objAverage.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), ".2f"
            )
        else:
            arrOutput[iColumnIndex] = format_sum_value(
                sum(arrValues, Decimal("0"))
            )

    arrShippingCounts: List[Decimal] = dicColumnValues[
        SHIPPING_COUNT_COLUMN_INDEX
    ]
    objShippingCountTotal: Decimal = sum(
        arrShippingCounts, Decimal("0")
    )
    if not arrShippingCounts:
        arrOutput[AVERAGE_WEIGHT_COLUMN_INDEX] = ""
    elif objShippingCountTotal == 0:
        arrOutput[AVERAGE_WEIGHT_COLUMN_INDEX] = ""
        add_warning(
            arrWarnings,
            "AVERAGE_WEIGHT_DIVISION_BY_ZERO",
            pszInputPath,
            pszOutputPath,
            pszProcessMonth,
            pszGroupNumber,
            pszGroupName,
            "4(D) 平均体重, 19(S) 出荷羽数, 21(U) 出荷重量",
            "shipping count total is zero; average weight cannot be calculated.",
            "empty",
        )
    else:
        objShippingWeightTotal: Decimal = sum(
            dicColumnValues[SHIPPING_WEIGHT_COLUMN_INDEX], Decimal("0")
        )
        objAverageWeight: Decimal = (
            objShippingWeightTotal / objShippingCountTotal
        )
        arrOutput[AVERAGE_WEIGHT_COLUMN_INDEX] = format(
            objAverageWeight.quantize(
                Decimal("0.00001"), rounding=ROUND_HALF_UP
            ),
            ".5f",
        )
    write_step0002_tsv(pszOutputPath, "\t".join(arrHeader), ["\t".join(arrOutput)])


def format_step0002_total_warning(iIndex: int, dicWarning: Dict[str, str]) -> str:
    pszText: str = "\r\n[警告 " + str(iIndex) + "]\r\n"
    for pszLabel, pszKey in [
        ("警告種別", "type"), ("入力step0001", "input"),
        ("出力step0002", "output"), ("処理年月", "month"),
        ("グループ№", "group_number"), ("グループ名", "group_name"),
        ("対象列", "columns"), ("警告理由", "reason"), ("出力値", "output_value"),
    ]:
        pszText += pszLabel + ": " + dicWarning[pszKey] + "\r\n"
    return pszText


def write_step0002_total_warnings(
    pszInputExcelPath: str,
    pszExistingWarningPath: str,
    iExistingWarningCount: int,
    arrWarnings: List[Dict[str, str]],
) -> str:
    if not arrWarnings:
        return pszExistingWarningPath
    iTotalWarningCount: int = iExistingWarningCount + len(arrWarnings)
    if pszExistingWarningPath:
        pszTemporaryPath: str = ""
        try:
            with open(
                pszExistingWarningPath, "r", encoding="utf-8", newline=""
            ) as objFile:
                pszExistingText: str = objFile.read()
            objCountMatch = re.search(r"警告件数: (\d+)", pszExistingText)
            if objCountMatch is None:
                raise_conversion_error(
                    "WARNING_FILE_UPDATE_FAILED",
                    "warning count was not found in existing warning file. Path = "
                    + pszExistingWarningPath,
                )
            pszExistingText = (
                pszExistingText[: objCountMatch.start()]
                + "警告件数: "
                + str(iTotalWarningCount)
                + pszExistingText[objCountMatch.end() :]
            )
            for iIndex, dicWarning in enumerate(
                arrWarnings, start=iExistingWarningCount + 1
            ):
                pszExistingText += format_step0002_total_warning(iIndex, dicWarning)
            iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_warning_",
                suffix=".txt",
                dir=os.path.dirname(pszExistingWarningPath),
            )
            os.close(iFileDescriptor)
            with open(
                pszTemporaryPath, "w", encoding="utf-8", newline=""
            ) as objFile:
                objFile.write(pszExistingText)
            os.replace(pszTemporaryPath, pszExistingWarningPath)
            pszTemporaryPath = ""
            return pszExistingWarningPath
        except ConversionError:
            raise
        except Exception as objException:
            raise_conversion_error(
                "WARNING_FILE_UPDATE_FAILED",
                "failed to update warning file. Detail = " + str(objException),
            )
        finally:
            if pszTemporaryPath and os.path.exists(pszTemporaryPath):
                os.remove(pszTemporaryPath)

    pszWarningPath: str = create_warning_path(pszInputExcelPath)
    try:
        with open(pszWarningPath, "x", encoding="utf-8", newline="") as objFile:
            objFile.write("月次・総計集計が警告付きで完了しました。\r\n")
            objFile.write("入力Excel: " + pszInputExcelPath + "\r\n")
            objFile.write("警告件数: " + str(len(arrWarnings)) + "\r\n")
            for iIndex, dicWarning in enumerate(arrWarnings, start=1):
                objFile.write(format_step0002_total_warning(iIndex, dicWarning))
        return pszWarningPath
    except Exception as objException:
        if os.path.exists(pszWarningPath):
            os.remove(pszWarningPath)
        raise_conversion_error(
            "WARNING_FILE_CREATE_FAILED",
            "failed to create total warning file. Detail = "
            + str(objException),
        )


def create_step0002_total_files(
    arrStep0001Paths: List[str],
    pszInputExcelPath: str,
    pszExistingWarningPath: str,
    iExistingWarningCount: int,
) -> Tuple[List[str], str, int]:
    """グループ別出力と独立して、全対象月の総計TSVを作成します。"""
    arrCreatedPaths: List[str] = []
    arrWarnings: List[Dict[str, str]] = []
    try:
        for pszStep0001Path in arrStep0001Paths:
            pszOutputPath: str = create_step0002_total_output_path(pszStep0001Path)
            aggregate_step0001_total(pszStep0001Path, pszOutputPath, arrWarnings)
            arrCreatedPaths.append(pszOutputPath)
        pszWarningPath: str = write_step0002_total_warnings(
            pszInputExcelPath, pszExistingWarningPath, iExistingWarningCount, arrWarnings
        )
        return arrCreatedPaths, pszWarningPath, iExistingWarningCount + len(arrWarnings)
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0002_TOTAL_CLEANUP_FAILED",
                "step0002 total aggregation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0002_file_name(
    pszStep0002Path: str, bVariety: bool = False
) -> Tuple[str, int, int, str]:
    pszFileName: str = os.path.basename(pszStep0002Path)
    pszCategory = "品種別" if bVariety else "グループ別"
    objMatch = re.fullmatch(
        rf"(.+)_step0002_{pszCategory}_(\d{{4}})_(\d{{2}})(_\d+)?\.tsv", pszFileName
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0002_FILE_NAME",
            "invalid step0002 file name. Path = " + pszStep0002Path,
        )
    iYear: int = int(objMatch.group(2))
    iMonth: int = int(objMatch.group(3))
    if iMonth < 1 or iMonth > 12:
        raise_conversion_error(
            "INVALID_STEP0002_FILE_NAME",
            "invalid month in step0002 file name. Path = " + pszStep0002Path,
        )
    return objMatch.group(1), iYear, iMonth, objMatch.group(4) or ""


def create_step0003_output_path(
    pszStep0002Path: str, pszItemName: str, bVariety: bool = False
) -> str:
    pszBaseName, iYear, iMonth, pszExistingSuffix = parse_step0002_file_name(
        pszStep0002Path, bVariety
    )
    pszDirectory: str = os.path.dirname(pszStep0002Path)
    pszCategory = "品種別" if bVariety else "グループ別"
    pszStem: str = (
        f"{pszBaseName}_step0003_{pszCategory}_{pszItemName}_{iYear:04d}_{iMonth:02d}"
        + pszExistingSuffix
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def write_step0003_tsv(
    pszOutputPath: str, pszHeader: str, arrRows: List[str]
) -> None:
    try:
        write_monthly_tsv(pszOutputPath, pszHeader, arrRows)
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0003_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0003_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0003_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0003_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0003_WRITE_FAILED"),
            str(objException),
        )


def extract_step0003_item_tsv(
    pszInputPath: str,
    pszOutputPath: str,
    pszItemName: str,
    iSourceColumnIndex: int,
    pszSourceHeader: str,
    bAllowBlank: bool,
    bVariety: bool = False,
) -> Tuple[int, int]:
    _, iFileYear, iFileMonth, _ = parse_step0002_file_name(pszInputPath, bVariety)
    arrCategories = SPLIT_VARIETIES if bVariety else GROUPS
    iNumberColumnIndex, iNameColumnIndex = (15, 16) if bVariety else (10, 11)
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0002_READ_FAILED", "failed to read step0002 TSV. Detail = " + str(objException)
        )
    if not arrLines:
        raise_conversion_error("STEP0002_HEADER_MISSING", "step0002 header was missing.")
    arrHeader: List[str] = arrLines[0].split("\t")
    if len(arrHeader) != 64:
        raise_conversion_error(
            "INVALID_STEP0002_COLUMN_COUNT",
            "step0002 header must contain 64 columns. Actual = " + str(len(arrHeader)),
        )
    dicRequiredHeaders: Dict[int, str] = {
        0: "年",
        1: "月",
        iSourceColumnIndex: pszSourceHeader,
    }
    if not bVariety:
        dicRequiredHeaders.update({10: "ｸﾞﾙｰﾌﾟ№", 11: "ｸﾞﾙｰﾌﾟ名"})
    for iColumnIndex, pszExpectedHeader in dicRequiredHeaders.items():
        if arrHeader[iColumnIndex].strip(" \u3000\t") != pszExpectedHeader:
            raise_conversion_error(
                "INVALID_STEP0002_HEADER",
                "unexpected step0002 header. Column = "
                + str(iColumnIndex + 1)
                + ", Expected = "
                + pszExpectedHeader
                + ", Actual = "
                + arrHeader[iColumnIndex],
            )
    if len(arrLines) != (len(STEP0002_VARIETIES) if bVariety else len(GROUPS)) + 1:
        raise_conversion_error(
            "INVALID_STEP0002_ROW_COUNT",
            "step0002 must contain one header and "
            + ("six" if bVariety else "three") + " data rows. Actual rows = "
            + str(len(arrLines)),
        )

    arrOutputRows: List[str] = []
    arrDataRowIndexes = [1, 2, 3, 5, 6] if bVariety else [1, 2, 3]
    for iGroupIndex, (pszExpectedNumber, pszExpectedName) in zip(
        arrDataRowIndexes, arrCategories
    ):
        iLineNumber: int = iGroupIndex + 1
        arrFields: List[str] = arrLines[iGroupIndex].split("\t")
        if len(arrFields) != 64:
            raise_conversion_error(
                "INVALID_STEP0002_COLUMN_COUNT",
                "step0002 row must contain 64 columns. Line = "
                + str(iLineNumber)
                + ", Actual = "
                + str(len(arrFields)),
            )
        iYear, iMonth = parse_year_month(arrFields[0], arrFields[1], iLineNumber)
        if (iYear, iMonth) != (iFileYear, iFileMonth):
            raise_conversion_error(
                "STEP0002_YEAR_MONTH_MISMATCH",
                "step0002 file name and row year/month do not match. Line = "
                + str(iLineNumber),
            )
        pszGroupNumber: str = arrFields[iNumberColumnIndex].strip(" \u3000\t")
        pszGroupName: str = arrFields[iNameColumnIndex].strip(" \u3000\t")
        if (pszGroupNumber, pszGroupName) != (pszExpectedNumber, pszExpectedName):
            raise_conversion_error(
                "INVALID_STEP0002_VARIETY_ORDER" if bVariety else "INVALID_STEP0002_GROUP_ORDER",
                "unexpected " + ("variety" if bVariety else "group")
                + " order or mapping. Line = " + str(iLineNumber),
            )
        try:
            objItemValue = parse_decimal_value(
                arrFields[iSourceColumnIndex],
                iLineNumber,
                iSourceColumnIndex,
                arrHeader[iSourceColumnIndex],
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0003_ITEM_VALUE",
                "invalid step0003 item value. Item = "
                + pszItemName
                + ". Detail = "
                + str(objException),
            )
        if objItemValue is None and not bAllowBlank:
            raise_conversion_error(
                "INVALID_STEP0003_ITEM_VALUE",
                "step0003 item value was blank. Item = "
                + pszItemName
                + ", Line = "
                + str(iLineNumber),
            )
        arrOutputRows.append(
            "\t".join(
                [
                    arrFields[0],
                    arrFields[1],
                    arrFields[iNumberColumnIndex],
                    arrFields[iNameColumnIndex],
                    arrFields[iSourceColumnIndex],
                ]
            )
        )
    arrOutputHeader: List[str] = [
        arrHeader[0],
        arrHeader[1],
        "品種No." if bVariety else arrHeader[10],
        "品種名" if bVariety else arrHeader[11],
        pszItemName,
    ]
    write_step0003_tsv(pszOutputPath, "\t".join(arrOutputHeader), arrOutputRows)
    return iFileYear, iFileMonth


def create_step0003_files(
    arrStep0002Paths: List[str], iLatestYear: int, bVariety: bool = False
) -> Tuple[List[str], List[Tuple[int, int]]]:
    setTargetYears = set(range(iLatestYear - 4, iLatestYear + 1))
    arrTargets: List[Tuple[int, int, str]] = []
    for pszStep0002Path in arrStep0002Paths:
        _, iYear, iMonth, _ = parse_step0002_file_name(pszStep0002Path, bVariety)
        if iYear in setTargetYears:
            arrTargets.append((iYear, iMonth, pszStep0002Path))
    arrTargets.sort()

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for iExpectedYear, iExpectedMonth, pszStep0002Path in arrTargets:
            for (
                pszItemName,
                iSourceColumnIndex,
                pszSourceHeader,
                bAllowBlank,
            ) in STEP0003_ITEMS:
                pszStep0003Path: str = create_step0003_output_path(
                    pszStep0002Path, pszItemName, bVariety
                )
                try:
                    iYear, iMonth = extract_step0003_item_tsv(
                        pszStep0002Path,
                        pszStep0003Path,
                        pszItemName,
                        iSourceColumnIndex,
                        pszSourceHeader,
                        bAllowBlank,
                        bVariety,
                    )
                except ConversionError as objException:
                    raise_conversion_error(
                        objException.pszErrorCode,
                        "step0003 extraction failed. Input step0002 = "
                        + pszStep0002Path
                        + ", Target year/month = "
                        + f"{iExpectedYear:04d}_{iExpectedMonth:02d}"
                        + ", Item = "
                        + pszItemName
                        + ". Detail = "
                        + str(objException),
                    )
                if (iYear, iMonth) != (iExpectedYear, iExpectedMonth):
                    raise_conversion_error(
                        "STEP0002_YEAR_MONTH_MISMATCH",
                        "unexpected year/month returned for step0002. Path = "
                        + pszStep0002Path,
                    )
                arrCreatedPaths.append(pszStep0003Path)
            arrYearMonths.append((iExpectedYear, iExpectedMonth))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0003_CLEANUP_FAILED",
                "step0003 extraction failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0003_variety_additional_files(
    arrStep0002VarietyPaths: List[str], iLatestYear: int
) -> List[str]:
    """品種別step0002から農家廃棄羽数・重量と日令を追加抽出します。"""
    setTargetYears = set(range(iLatestYear - 4, iLatestYear + 1))
    arrTargets: List[Tuple[int, int, str]] = []
    for pszStep0002Path in arrStep0002VarietyPaths:
        _, iYear, iMonth, _ = parse_step0002_file_name(
            pszStep0002Path, True
        )
        if iYear in setTargetYears:
            arrTargets.append((iYear, iMonth, pszStep0002Path))
    arrTargets.sort()

    arrCreatedPaths: List[str] = []
    try:
        for iExpectedYear, iExpectedMonth, pszStep0002Path in arrTargets:
            for (
                pszItemName,
                iSourceColumnIndex,
                pszSourceHeader,
                bAllowBlank,
            ) in STEP0003_VARIETY_ADDITIONAL_ITEMS:
                pszStep0003Path = create_step0003_output_path(
                    pszStep0002Path, pszItemName, True
                )
                try:
                    iYear, iMonth = extract_step0003_item_tsv(
                        pszStep0002Path,
                        pszStep0003Path,
                        pszItemName,
                        iSourceColumnIndex,
                        pszSourceHeader,
                        bAllowBlank,
                        True,
                    )
                except ConversionError as objException:
                    raise_conversion_error(
                        objException.pszErrorCode,
                        "additional variety step0003 extraction failed. "
                        "Input step0002 = " + pszStep0002Path
                        + ", Target year/month = "
                        + f"{iExpectedYear:04d}_{iExpectedMonth:02d}"
                        + ", Item = " + pszItemName
                        + ". Detail = " + str(objException),
                    )
                if (iYear, iMonth) != (iExpectedYear, iExpectedMonth):
                    raise_conversion_error(
                        "STEP0002_YEAR_MONTH_MISMATCH",
                        "unexpected year/month returned for step0002. Path = "
                        + pszStep0002Path,
                    )
                arrCreatedPaths.append(pszStep0003Path)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0003_VARIETY_ADDITIONAL_CLEANUP_FAILED",
                "additional variety step0003 extraction failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0003_variety_husbandry_average_output_path(
    pszDayAgePath: str,
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0003_品種別_日令_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszDayAgePath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0003_VARIETY_DAY_AGE_FILE_NAME",
            "invalid variety day-age step0003 file name. Path = "
            + pszDayAgePath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
    pszStem = (
        f"{objMatch.group(1)}_step0003_品種別_飼育日数平均_"
        f"{iYear:04d}_{iMonth:02d}{objMatch.group(4)}"
    )
    pszDirectory = os.path.dirname(pszDayAgePath)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0003_variety_husbandry_average_files(
    arrDayAgePaths: List[str],
) -> List[str]:
    """日令TSVを内容そのままで複製し、飼育日数平均名で保存します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszDayAgePath in sorted(arrDayAgePaths):
            try:
                with open(pszDayAgePath, "rb") as objInputFile:
                    objContents = objInputFile.read()
            except Exception as objException:
                raise_conversion_error(
                    "STEP0003_VARIETY_DAY_AGE_COPY_READ_FAILED",
                    "failed to read variety day-age TSV. Path = "
                    + pszDayAgePath + "; Detail = " + str(objException),
                )

            pszOutputPath = create_step0003_variety_husbandry_average_output_path(
                pszDayAgePath
            )
            while True:
                try:
                    with open(pszOutputPath, "xb") as objOutputFile:
                        arrCreatedPaths.append(pszOutputPath)
                        objOutputFile.write(objContents)
                    break
                except FileExistsError:
                    pszOutputPath = (
                        create_step0003_variety_husbandry_average_output_path(
                            pszDayAgePath
                        )
                    )
                except Exception as objException:
                    raise_conversion_error(
                        "STEP0003_VARIETY_DAY_AGE_COPY_WRITE_FAILED",
                        "failed to copy variety day-age TSV. Path = "
                        + pszOutputPath + "; Detail = " + str(objException),
                    )
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0003_VARIETY_DAY_AGE_COPY_CLEANUP_FAILED",
                "variety husbandry average file creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0007_variety_output_path(
    pszStep0003Path: str, pszItemName: str
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0003_品種別_(.+)_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszStep0003Path),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0003_VARIETY_FILE_NAME",
            "invalid variety step0003 file name. Path = " + pszStep0003Path,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    pszFileItemName = (
        "控除後平均体重" if pszItemName == "平均体重" else pszItemName
    )
    pszStem = (
        f"{objMatch.group(1)}_step0003_品種別_{pszFileItemName}_"
        f"{iYear:04d}_{iMonth:02d}{objMatch.group(5)}"
    )
    pszDirectory = os.path.dirname(pszStep0003Path)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def read_variety_converted_discard_weights(
    pszPath: str, iYear: int, iMonth: int
) -> Dict[Tuple[str, str], Decimal]:
    """既存の品種区分で、明細ごとの0.72換算・切り捨て重量を合計します。"""
    dicTotals = {key: Decimal("0") for key in SPLIT_VARIETIES}
    arrWeightItems = STEP0004_TOTAL_WEIGHT_ITEMS + ["予備３重量", "予備４重量"]
    pszContext = f"Year/month = {iYear:04d}/{iMonth:02d}; Path = {pszPath}"
    with open(pszPath, "r", encoding="utf-8-sig", newline="") as objFile:
        objReader = csv.reader(objFile, delimiter="\t")
        arrHeader = next(objReader, [])
        if (len(arrHeader) != 64
                or [arrHeader[i] for i in (0, 1, 10, 11, 15, 16)]
                != ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "品種№", "品種名"]
                or arrHeader[47:64] != arrWeightItems):
            raise_conversion_error("INVALID_VARIETY_WEIGHT_HEADER", pszContext)
        for iLine, arrFields in enumerate(objReader, start=2):
            pszRowContext = pszContext + f"; Line = {iLine}"
            if len(arrFields) != 64:
                raise_conversion_error("INVALID_VARIETY_WEIGHT_COLUMNS", pszRowContext)
            objKey = (arrFields[15].strip(), arrFields[16].strip(" \u3000\t"))
            pszRowContext += "; Variety = " + str(objKey)
            try:
                if (parse_year_month(arrFields[0], arrFields[1], iLine) != (iYear, iMonth)
                        or objKey not in VARIETIES):
                    raise ValueError("year/month or variety did not match")
                if objKey == ("1", "ブロイラー"):
                    pszGroup = arrFields[10].strip()
                    if pszGroup not in {"1", "2", "4"}:
                        raise ValueError("invalid broiler group")
                    objKey = BROILER_BREAKDOWNS[0 if pszGroup in {"1", "2"} else 1]
                objSubtotal = Decimal("0")
                for iColumn, pszItem in enumerate(arrWeightItems, start=47):
                    objValue = parse_decimal_value(
                        arrFields[iColumn].replace(",", ""), iLine, iColumn, pszItem
                    )
                    if objValue is None:
                        raise ValueError("blank weight. Item = " + pszItem)
                    objSubtotal += objValue
            except (ConversionError, ValueError) as objException:
                raise_conversion_error(
                    "INVALID_VARIETY_WEIGHT_ROW",
                    pszRowContext + "; Detail = " + str(objException),
                )
            dicTotals[objKey] += (objSubtotal / Decimal("0.72")).quantize(
                Decimal("0.1"), rounding=ROUND_DOWN
            )
    return dicTotals


def create_step0007_variety_files(
    arrStep0003VarietyPaths: List[str],
    arrStep0003YearMonths: List[Tuple[int, int]],
    arrStep0001Paths: List[str],
) -> List[str]:
    """品種別step0003から廃棄合計と廃棄控除後の3項目を追加出力します。"""
    dicPaths = build_step0003_path_map(
        arrStep0003VarietyPaths, arrStep0003YearMonths
    )
    arrRequiredItems = [
        "出荷羽数", "出荷重量",
        *STEP0004_TOTAL_COUNT_ITEMS, *STEP0004_TOTAL_WEIGHT_ITEMS,
        "懸鳥前死鶏羽数", "懸鳥前発育不良羽数",
        "懸鳥前死鶏重量", "懸鳥前発育不良重量",
    ]
    arrOutputItems = [
        "全廃棄羽数", "全廃棄重量", "処理羽数", "処理重量", "平均体重",
    ]
    arrCreatedPaths: List[str] = []
    dicMonthlyPaths: Dict[Tuple[str, str, int, int], str] = {}
    for pszPath in arrStep0001Paths:
        objMatch = re.fullmatch(
            r"(.+)_step0001_(\d{4})_(\d{2})(?:_\d+)*\.tsv", os.path.basename(pszPath)
        )
        if objMatch is None:
            raise_conversion_error("INVALID_VARIETY_WEIGHT_FILENAME", pszPath)
        pszBase, pszYear, pszMonth = objMatch.groups()
        iYear, iMonth = parse_year_month(pszYear, pszMonth, 1)
        objMonthlyKey = (os.path.dirname(os.path.abspath(pszPath)), pszBase, iYear, iMonth)
        if objMonthlyKey in dicMonthlyPaths:
            raise_conversion_error(
                "DUPLICATE_VARIETY_WEIGHT_INPUT",
                "duplicate step0001 input. Key = " + str(objMonthlyKey) + "; Path = " + pszPath,
            )
        dicMonthlyPaths[objMonthlyKey] = pszPath
    try:
        for iYear, iMonth in arrStep0003YearMonths:
            pszSourcePath = dicPaths[(iYear, iMonth, "出荷羽数")]
            objSourceMatch = re.fullmatch(
                r"(.+)_step0003_品種別_出荷羽数_\d{4}_\d{2}(?:_\d+)*\.tsv",
                os.path.basename(pszSourcePath),
            )
            if objSourceMatch is None:
                raise_conversion_error("INVALID_VARIETY_SHIPPING_FILENAME", pszSourcePath)
            for pszPreItem in [
                "懸鳥前死鶏羽数", "懸鳥前発育不良羽数",
                "懸鳥前死鶏重量", "懸鳥前発育不良重量",
            ]:
                pszPrePath = dicPaths[(iYear, iMonth, pszPreItem)]
                pszPattern = (
                    re.escape(objSourceMatch.group(1)) + "_step0003_品種別_"
                    + re.escape(pszPreItem) + f"_{iYear:04d}_{iMonth:02d}"
                    + r"(?:_\d+)*\.tsv"
                )
                if (os.path.dirname(os.path.abspath(pszPrePath))
                        != os.path.dirname(os.path.abspath(pszSourcePath))
                        or re.fullmatch(pszPattern, os.path.basename(pszPrePath)) is None):
                    raise_conversion_error(
                        "VARIETY_PRE_SHACKLING_INPUT_MISMATCH",
                        f"Year/month = {iYear:04d}/{iMonth:02d}; Item = {pszPreItem}; "
                        f"Path = {pszPrePath}; Source = {pszSourcePath}",
                    )
            objMonthlyKey = (
                os.path.dirname(os.path.abspath(pszSourcePath)),
                objSourceMatch.group(1), iYear, iMonth,
            )
            if objMonthlyKey not in dicMonthlyPaths:
                raise_conversion_error(
                    "MISSING_VARIETY_WEIGHT_INPUT",
                    "missing step0001. Key = " + str(objMonthlyKey) + "; Source = " + pszSourcePath,
                )
            dicConvertedWeights = read_variety_converted_discard_weights(
                dicMonthlyPaths[objMonthlyKey], iYear, iMonth
            )
            dicItemRows = {
                pszItemName: read_step0003_for_step0004(
                    dicPaths[(iYear, iMonth, pszItemName)],
                    pszItemName, False, iYear, iMonth, True,
                )
                for pszItemName in arrRequiredItems
            }
            dicOutputRows: Dict[str, List[str]] = {
                pszItemName: [] for pszItemName in arrOutputItems
            }
            for objVarietyKey in SPLIT_VARIETIES:
                arrCommonFields = dicItemRows["出荷羽数"][objVarietyKey][:4]
                for pszItemName in arrRequiredItems:
                    if dicItemRows[pszItemName][objVarietyKey][:4] != arrCommonFields:
                        raise_conversion_error(
                            "STEP0003_VARIETY_KEY_MISMATCH",
                            "variety step0003 row keys do not match. Item = "
                            + pszItemName + "; Variety = " + str(objVarietyKey),
                        )
                objDiscardCount = sum(
                    (Decimal(dicItemRows[pszItemName][objVarietyKey][4])
                     for pszItemName in STEP0004_TOTAL_COUNT_ITEMS), Decimal("0")
                )
                # Match the existing group pipeline: round the combined weight once.
                objDiscardWeight = sum(
                    (Decimal(dicItemRows[pszItemName][objVarietyKey][4])
                     for pszItemName in STEP0004_TOTAL_WEIGHT_ITEMS), Decimal("0")
                ).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
                objProcessedCount = (
                    Decimal(dicItemRows["出荷羽数"][objVarietyKey][4]) - objDiscardCount
                    - Decimal(dicItemRows["懸鳥前死鶏羽数"][objVarietyKey][4])
                    - Decimal(dicItemRows["懸鳥前発育不良羽数"][objVarietyKey][4])
                )
                objProcessedWeight = (
                    Decimal(dicItemRows["出荷重量"][objVarietyKey][4])
                    - Decimal(dicItemRows["懸鳥前死鶏重量"][objVarietyKey][4])
                    - Decimal(dicItemRows["懸鳥前発育不良重量"][objVarietyKey][4])
                    - dicConvertedWeights[objVarietyKey]
                ).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
                objAverageWeight = (
                    (objProcessedWeight / objProcessedCount).quantize(
                        Decimal("0.00000001"), rounding=ROUND_HALF_UP
                    ) if objProcessedCount != 0 else Decimal("0")
                )
                arrValues = [
                    format(Decimal(format_sum_value(objDiscardCount)), ",f"),
                    format(objDiscardWeight, ",.1f"),
                    format_sum_value(objProcessedCount),
                    format(objProcessedWeight, ".1f"),
                    format_sum_value(objAverageWeight),
                ]
                for pszItemName, pszValue in zip(arrOutputItems, arrValues):
                    dicOutputRows[pszItemName].append(
                        "\t".join(arrCommonFields + [pszValue])
                    )
            for pszItemName in arrOutputItems:
                pszOutputPath = create_step0007_variety_output_path(
                    dicPaths[(iYear, iMonth, "出荷羽数")], pszItemName
                )
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(["年", "月", "品種No.", "品種名", pszItemName]),
                    dicOutputRows[pszItemName],
                )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0007_VARIETY_CLEANUP_FAILED",
                "variety step0007 creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0003_variety_deduction_growth_rate_output_path(
    pszPlacementPath: str,
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0003_品種別_入雛羽数_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszPlacementPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0003_VARIETY_PLACEMENT_FILE_NAME",
            "invalid variety placement step0003 file name. Path = "
            + pszPlacementPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
    pszStem = (
        f"{objMatch.group(1)}_step0003_品種別_廃棄控除後育成率_"
        f"{iYear:04d}_{iMonth:02d}{objMatch.group(4)}"
    )
    pszDirectory = os.path.dirname(pszPlacementPath)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0003_variety_deduction_growth_rate_files(
    arrStep0003VarietyPaths: List[str],
    arrStep0007VarietyPaths: List[str],
    arrStep0003YearMonths: List[Tuple[int, int]],
) -> List[str]:
    """入雛羽数と処理羽数から品種別の控除廃棄育成率を作成します。"""
    dicStep0003Paths = build_step0003_path_map(
        arrStep0003VarietyPaths, arrStep0003YearMonths
    )
    dicProcessedPaths: Dict[Tuple[int, int], str] = {}
    for pszPath in arrStep0007VarietyPaths:
        objMatch = re.fullmatch(
            r".+_step0003_品種別_処理羽数_(\d{4})_(\d{2})(?:_\d+)*\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            continue
        iYear, iMonth = parse_year_month(objMatch.group(1), objMatch.group(2), 1)
        objKey = (iYear, iMonth)
        if objKey in dicProcessedPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0003_VARIETY_PROCESSED_COUNT",
                "duplicate variety processed count input. Year/month = "
                + f"{iYear:04d}_{iMonth:02d}",
            )
        dicProcessedPaths[objKey] = pszPath

    setExpectedYearMonths = set(arrStep0003YearMonths)
    if set(dicProcessedPaths) != setExpectedYearMonths:
        raise_conversion_error(
            "STEP0003_VARIETY_PROCESSED_COUNT_INPUT_MISMATCH",
            "variety processed count inputs did not match step0003 year/months.",
        )

    arrCreatedPaths: List[str] = []
    try:
        for iYear, iMonth in arrStep0003YearMonths:
            dicPlacementRows = read_step0003_for_step0004(
                dicStep0003Paths[(iYear, iMonth, "入雛羽数")],
                "入雛羽数",
                False,
                iYear,
                iMonth,
                True,
            )
            dicProcessedRows = read_step0003_for_step0004(
                dicProcessedPaths[(iYear, iMonth)],
                "処理羽数",
                False,
                iYear,
                iMonth,
                True,
            )
            arrOutputRows: List[str] = []
            for pszVarietyNumber, pszVarietyName in SPLIT_VARIETIES:
                objVarietyKey = (pszVarietyNumber, pszVarietyName)
                arrPlacementFields = dicPlacementRows[objVarietyKey]
                arrProcessedFields = dicProcessedRows[objVarietyKey]
                if arrPlacementFields[:4] != arrProcessedFields[:4]:
                    raise_conversion_error(
                        "STEP0003_VARIETY_GROWTH_RATE_KEY_MISMATCH",
                        "placement and processed count keys did not match. "
                        "Year/month = " + f"{iYear:04d}_{iMonth:02d}"
                        + ", Variety = " + str(objVarietyKey),
                    )
                objPlacementCount = Decimal(arrPlacementFields[4])
                objProcessedCount = Decimal(arrProcessedFields[4])
                objGrowthRate = (
                    (objProcessedCount / objPlacementCount).quantize(
                        Decimal("0.0001"), rounding=ROUND_HALF_UP
                    )
                    if objPlacementCount != 0
                    else Decimal("0")
                )
                arrOutputRows.append(
                    "\t".join(
                        arrPlacementFields[:4] + [format_sum_value(objGrowthRate)]
                    )
                )

            pszOutputPath = create_step0003_variety_deduction_growth_rate_output_path(
                dicStep0003Paths[(iYear, iMonth, "入雛羽数")]
            )
            write_monthly_tsv(
                pszOutputPath,
                "\t".join(["年", "月", "品種No.", "品種名", "育成率"]),
                arrOutputRows,
            )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0003_VARIETY_GROWTH_RATE_CLEANUP_FAILED",
                "variety deduction growth rate creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0003_variety_pre_discard_growth_rate_output_path(
    pszPlacementPath: str,
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0003_品種別_入雛羽数_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszPlacementPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0003_VARIETY_PLACEMENT_FILE_NAME",
            "invalid variety placement step0003 file name. Path = "
            + pszPlacementPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
    pszStem = (
        f"{objMatch.group(1)}_step0003_品種別_廃棄控除前育成率_"
        f"{iYear:04d}_{iMonth:02d}{objMatch.group(4)}"
    )
    pszDirectory = os.path.dirname(pszPlacementPath)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0003_variety_pre_discard_growth_rate_files(
    arrStep0003VarietyPaths: List[str],
    arrStep0003YearMonths: List[Tuple[int, int]],
) -> List[str]:
    """出荷羽数と入雛羽数から品種別の廃棄控除前育成率を作成します。"""
    dicStep0003Paths = build_step0003_path_map(
        arrStep0003VarietyPaths, arrStep0003YearMonths
    )
    arrCreatedPaths: List[str] = []
    try:
        for iYear, iMonth in arrStep0003YearMonths:
            dicPlacementRows = read_step0003_for_step0004(
                dicStep0003Paths[(iYear, iMonth, "入雛羽数")],
                "入雛羽数",
                False,
                iYear,
                iMonth,
                True,
            )
            dicShipmentRows = read_step0003_for_step0004(
                dicStep0003Paths[(iYear, iMonth, "出荷羽数")],
                "出荷羽数",
                False,
                iYear,
                iMonth,
                True,
            )
            arrOutputRows: List[str] = []
            for pszVarietyNumber, pszVarietyName in SPLIT_VARIETIES:
                objVarietyKey = (pszVarietyNumber, pszVarietyName)
                arrPlacementFields = dicPlacementRows[objVarietyKey]
                arrShipmentFields = dicShipmentRows[objVarietyKey]
                if arrPlacementFields[:4] != arrShipmentFields[:4]:
                    raise_conversion_error(
                        "STEP0003_VARIETY_PRE_DISCARD_GROWTH_RATE_KEY_MISMATCH",
                        "placement and shipment count keys did not match. "
                        "Year/month = " + f"{iYear:04d}_{iMonth:02d}"
                        + ", Variety = " + str(objVarietyKey),
                    )
                objPlacementCount = Decimal(arrPlacementFields[4])
                objShipmentCount = Decimal(arrShipmentFields[4])
                objGrowthRate = (
                    (objShipmentCount / objPlacementCount).quantize(
                        Decimal("0.0001"), rounding=ROUND_HALF_UP
                    )
                    if objPlacementCount != 0
                    else Decimal("0")
                )
                arrOutputRows.append(
                    "\t".join(
                        arrPlacementFields[:4] + [format_sum_value(objGrowthRate)]
                    )
                )

            pszOutputPath = (
                create_step0003_variety_pre_discard_growth_rate_output_path(
                    dicStep0003Paths[(iYear, iMonth, "入雛羽数")]
                )
            )
            write_monthly_tsv(
                pszOutputPath,
                "\t".join(["年", "月", "品種No.", "品種名", "育成率"]),
                arrOutputRows,
            )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0003_VARIETY_PRE_DISCARD_GROWTH_RATE_CLEANUP_FAILED",
                "variety pre-discard growth rate creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0002_total_file_name(pszStep0002Path: str) -> Tuple[str, int, int, str]:
    pszFileName: str = os.path.basename(pszStep0002Path)
    objMatch = re.fullmatch(
        r"(.+)_step0002_総計_(\d{4})_(\d{2})(_\d+)?\.tsv", pszFileName
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0002_FILE_NAME",
            "invalid step0002 total file name. Path = " + pszStep0002Path,
        )
    iYear: int = int(objMatch.group(2))
    iMonth: int = int(objMatch.group(3))
    if iMonth < 1 or iMonth > 12:
        raise_conversion_error(
            "INVALID_STEP0002_FILE_NAME",
            "invalid month in step0002 file name. Path = " + pszStep0002Path,
        )
    return objMatch.group(1), iYear, iMonth, objMatch.group(4) or ""


def create_step0003_total_output_path(pszStep0002Path: str, pszItemName: str) -> str:
    pszBaseName, iYear, iMonth, pszExistingSuffix = parse_step0002_total_file_name(
        pszStep0002Path
    )
    pszDirectory: str = os.path.dirname(pszStep0002Path)
    pszStem: str = (
        f"{pszBaseName}_step0003_総計_{pszItemName}_{iYear:04d}_{iMonth:02d}"
        + pszExistingSuffix
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def extract_step0003_total_item_tsv(
    pszInputPath: str,
    pszOutputPath: str,
    pszItemName: str,
    iSourceColumnIndex: int,
    pszSourceHeader: str,
    bAllowBlank: bool,
) -> Tuple[int, int]:
    """step0002総計の値を再計算せず、1項目の総計TSVへ抽出します。"""
    _, iFileYear, iFileMonth, _ = parse_step0002_total_file_name(pszInputPath)
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0002_READ_FAILED", "failed to read step0002 TSV. Detail = " + str(objException)
        )
    if not arrLines:
        raise_conversion_error("STEP0002_HEADER_MISSING", "step0002 header was missing.")
    arrHeader: List[str] = arrLines[0].split("\t")
    if len(arrHeader) != 64:
        raise_conversion_error(
            "INVALID_STEP0002_COLUMN_COUNT",
            "step0002 header must contain 64 columns. Actual = " + str(len(arrHeader)),
        )
    dicRequiredHeaders: Dict[int, str] = {
        0: "年",
        1: "月",
        10: "ｸﾞﾙｰﾌﾟ№",
        11: "ｸﾞﾙｰﾌﾟ名",
        iSourceColumnIndex: pszSourceHeader,
    }
    for iColumnIndex, pszExpectedHeader in dicRequiredHeaders.items():
        if arrHeader[iColumnIndex].strip(" \u3000\t") != pszExpectedHeader:
            raise_conversion_error(
                "INVALID_STEP0002_HEADER",
                "unexpected step0002 header. Column = "
                + str(iColumnIndex + 1)
                + ", Expected = "
                + pszExpectedHeader
                + ", Actual = "
                + arrHeader[iColumnIndex],
            )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0002_ROW_COUNT",
            "step0002 must contain one header and one total data row. Actual rows = "
            + str(len(arrLines)),
        )

    arrOutputRows: List[str] = []
    for iGroupIndex, (pszExpectedNumber, pszExpectedName) in enumerate([("", "総計")], start=1):
        iLineNumber: int = iGroupIndex + 1
        arrFields: List[str] = arrLines[iGroupIndex].split("\t")
        if len(arrFields) != 64:
            raise_conversion_error(
                "INVALID_STEP0002_COLUMN_COUNT",
                "step0002 row must contain 64 columns. Line = "
                + str(iLineNumber)
                + ", Actual = "
                + str(len(arrFields)),
            )
        iYear, iMonth = parse_year_month(arrFields[0], arrFields[1], iLineNumber)
        if (iYear, iMonth) != (iFileYear, iFileMonth):
            raise_conversion_error(
                "STEP0002_YEAR_MONTH_MISMATCH",
                "step0002 file name and row year/month do not match. Line = "
                + str(iLineNumber),
            )
        pszGroupNumber: str = arrFields[10].strip(" \u3000\t")
        pszGroupName: str = arrFields[11].strip(" \u3000\t")
        if (pszGroupNumber, pszGroupName) != (pszExpectedNumber, pszExpectedName):
            raise_conversion_error(
                "INVALID_STEP0002_GROUP_ORDER",
                "expected an empty group number and total group name. Line = " + str(iLineNumber),
            )
        try:
            objItemValue = parse_decimal_value(
                arrFields[iSourceColumnIndex],
                iLineNumber,
                iSourceColumnIndex,
                arrHeader[iSourceColumnIndex],
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0003_ITEM_VALUE",
                "invalid step0003 item value. Item = "
                + pszItemName
                + ". Detail = "
                + str(objException),
            )
        if objItemValue is None and not bAllowBlank:
            raise_conversion_error(
                "INVALID_STEP0003_ITEM_VALUE",
                "step0003 item value was blank. Item = "
                + pszItemName
                + ", Line = "
                + str(iLineNumber),
            )
        arrOutputRows.append(
            "\t".join(
                [
                    arrFields[0],
                    arrFields[1],
                    arrFields[10],
                    arrFields[11],
                    arrFields[iSourceColumnIndex],
                ]
            )
        )
    arrOutputHeader: List[str] = [
        arrHeader[0],
        arrHeader[1],
        arrHeader[10],
        arrHeader[11],
        pszItemName,
    ]
    write_step0003_tsv(pszOutputPath, "\t".join(arrOutputHeader), arrOutputRows)
    return iFileYear, iFileMonth


def create_step0003_total_files(
    arrStep0002Paths: List[str], iLatestYear: int
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """直近5年間の総計43項目を、グループ別処理と独立して出力します。"""
    setTargetYears = set(range(iLatestYear - 4, iLatestYear + 1))
    arrTargets: List[Tuple[int, int, str]] = []
    for pszStep0002Path in arrStep0002Paths:
        _, iYear, iMonth, _ = parse_step0002_total_file_name(pszStep0002Path)
        if iYear in setTargetYears:
            arrTargets.append((iYear, iMonth, pszStep0002Path))
    arrTargets.sort()

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for iExpectedYear, iExpectedMonth, pszStep0002Path in arrTargets:
            for (
                pszItemName,
                iSourceColumnIndex,
                pszSourceHeader,
                bAllowBlank,
            ) in STEP0003_ITEMS:
                pszStep0003Path: str = create_step0003_total_output_path(
                    pszStep0002Path, pszItemName
                )
                try:
                    iYear, iMonth = extract_step0003_total_item_tsv(
                        pszStep0002Path,
                        pszStep0003Path,
                        pszItemName,
                        iSourceColumnIndex,
                        pszSourceHeader,
                        bAllowBlank,
                    )
                except ConversionError as objException:
                    raise_conversion_error(
                        objException.pszErrorCode,
                        "step0003 total extraction failed. Input step0002 = "
                        + pszStep0002Path
                        + ", Target year/month = "
                        + f"{iExpectedYear:04d}_{iExpectedMonth:02d}"
                        + ", Item = "
                        + pszItemName
                        + ". Detail = "
                        + str(objException),
                    )
                if (iYear, iMonth) != (iExpectedYear, iExpectedMonth):
                    raise_conversion_error(
                        "STEP0002_YEAR_MONTH_MISMATCH",
                        "unexpected year/month returned for step0002. Path = "
                        + pszStep0002Path,
                    )
                arrCreatedPaths.append(pszStep0003Path)
            arrYearMonths.append((iExpectedYear, iExpectedMonth))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0003_TOTAL_CLEANUP_FAILED",
                "step0003 total extraction failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


STEP0004_TOTAL_WEIGHT_ITEMS: List[str] = [
    "マレック重量", "大腸菌症重量", "ブドウ球菌症重量", "敗血症重量",
    "変性重量", "腹水症重量", "出血重量", "炎症重量", "腫瘍重量",
    "黄疸重量", "発育不良重量", "放血不良重量",
    "湯漬過度重量", "異常鶏重量", "予備２重量",
]


def parse_step0003_total_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0003_総計_(.+)_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_FILE_NAME",
            "invalid step0003 total file name. Path = " + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return objMatch.group(1), objMatch.group(2), iYear, iMonth, objMatch.group(5)


def create_step0004_total_weight_output_path(
    pszFirstInputPath: str,
) -> str:
    pszBaseName, _, iYear, iMonth, pszSuffix = parse_step0003_total_file_name(
        pszFirstInputPath
    )
    pszDirectory: str = os.path.dirname(pszFirstInputPath)
    pszStem: str = (
        f"{pszBaseName}_step0004_総計_廃棄データ別_重量_{iYear:04d}_{iMonth:02d}"
        + pszSuffix
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def read_step0003_total_weight_item(
    pszInputPath: str, pszExpectedItem: str,
) -> Tuple[List[str], List[str]]:
    _, pszItemName, iFileYear, iFileMonth, _ = parse_step0003_total_file_name(
        pszInputPath
    )
    if pszItemName != pszExpectedItem:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ITEM", "unexpected item. Path = " + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0003_TOTAL_READ_FAILED",
            "failed to read step0003 total TSV. Path = " + pszInputPath
            + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ROW_COUNT",
            "step0003 total TSV must contain one header and one data row. Path = "
            + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    if len(arrHeader) != 5 or len(arrFields) != 5:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_COLUMN_COUNT",
            "step0003 total TSV must contain five columns. Path = " + pszInputPath,
        )
    if arrHeader != ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", pszExpectedItem]:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_HEADER",
            "unexpected step0003 total header. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0003_TOTAL_YEAR_MONTH_MISMATCH",
            "file name and data year/month do not match. Path = " + pszInputPath,
        )
    if arrFields[2:4] != ["", "総計"]:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_GROUP",
            "expected an empty group number and total group name. Path = " + pszInputPath,
        )
    if parse_decimal_value(arrFields[4], 2, 4, pszExpectedItem) is None:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ITEM_VALUE",
            "total weight item was blank. Path = " + pszInputPath,
        )
    return arrHeader, arrFields


def create_step0004_total_weight_files(
    arrStep0003TotalPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """各年月の総計重量15項目を指定順に横結合し、数値表記を保持します。"""
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszInputPath in arrStep0003TotalPaths:
        pszBaseName, pszItemName, iYear, iMonth, pszSuffix = (
            parse_step0003_total_file_name(pszInputPath)
        )
        objKey = (os.path.dirname(os.path.abspath(pszInputPath)),
                  pszBaseName, iYear, iMonth, pszSuffix)
        dicItems = dicInputs.setdefault(objKey, {})
        if pszItemName not in STEP0004_TOTAL_WEIGHT_ITEMS:
            continue
        if pszItemName in dicItems:
            raise_conversion_error(
                "DUPLICATE_STEP0003_TOTAL_WEIGHT_ITEM",
                "duplicate total weight item. Path = " + pszInputPath,
            )
        dicItems[pszItemName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicItems = dicInputs[objKey]
            arrMissing = [pszItem for pszItem in STEP0004_TOTAL_WEIGHT_ITEMS
                          if pszItem not in dicItems]
            if arrMissing:
                raise_conversion_error(
                    "MISSING_STEP0003_TOTAL_WEIGHT_ITEMS",
                    "missing total weight items for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissing),
                )
            arrOutputHeader: List[str] = []
            arrOutputFields: List[str] = []
            for pszItem in STEP0004_TOTAL_WEIGHT_ITEMS:
                arrHeader, arrFields = read_step0003_total_weight_item(
                    dicItems[pszItem], pszItem
                )
                if not arrOutputFields:
                    arrOutputHeader = arrHeader[:4]
                    arrOutputFields = arrFields[:4]
                elif arrFields[:4] != arrOutputFields[:4]:
                    raise_conversion_error(
                        "STEP0003_TOTAL_COMMON_COLUMNS_MISMATCH",
                        "common columns do not match. Path = " + dicItems[pszItem],
                    )
                arrOutputHeader.append(arrHeader[4])
                arrOutputFields.append(arrFields[4])
            pszOutputPath = create_step0004_total_weight_output_path(
                dicItems[STEP0004_TOTAL_WEIGHT_ITEMS[0]]
            )
            write_step0004_tsv(
                pszOutputPath, "\t".join(arrOutputHeader), "\t".join(arrOutputFields)
            )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_TOTAL_WEIGHT_CLEANUP_FAILED",
                "total weight combination failed. Original error = " + str(objException)
                + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


STEP0004_TOTAL_COUNT_ITEMS: List[str] = [
    "マレック羽数", "大腸菌症羽数", "ブドウ球菌症羽数", "敗血症羽数",
    "変性羽数", "腹水症羽数", "出血羽数", "炎症羽数", "腫瘍羽数",
    "黄疸羽数", "発育不良羽数", "放血不良羽数",
    "湯漬過度羽数", "異常鶏羽数", "予備２羽数",
]


def create_step0004_total_count_output_path(
    pszFirstInputPath: str,
) -> str:
    pszBaseName, _, iYear, iMonth, pszSuffix = parse_step0003_total_file_name(
        pszFirstInputPath
    )
    pszDirectory: str = os.path.dirname(pszFirstInputPath)
    pszStem: str = (
        f"{pszBaseName}_step0004_総計_廃棄データ別_羽数_{iYear:04d}_{iMonth:02d}"
        + pszSuffix
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def read_step0003_total_count_item(
    pszInputPath: str, pszExpectedItem: str,
) -> Tuple[List[str], List[str]]:
    _, pszItemName, iFileYear, iFileMonth, _ = parse_step0003_total_file_name(
        pszInputPath
    )
    if pszItemName != pszExpectedItem:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ITEM", "unexpected item. Path = " + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0003_TOTAL_READ_FAILED",
            "failed to read step0003 total TSV. Path = " + pszInputPath
            + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ROW_COUNT",
            "step0003 total TSV must contain one header and one data row. Path = "
            + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    if len(arrHeader) != 5 or len(arrFields) != 5:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_COLUMN_COUNT",
            "step0003 total TSV must contain five columns. Path = " + pszInputPath,
        )
    if arrHeader != ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", pszExpectedItem]:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_HEADER",
            "unexpected step0003 total header. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0003_TOTAL_YEAR_MONTH_MISMATCH",
            "file name and data year/month do not match. Path = " + pszInputPath,
        )
    if arrFields[2:4] != ["", "総計"]:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_GROUP",
            "expected an empty group number and total group name. Path = " + pszInputPath,
        )
    if parse_decimal_value(arrFields[4], 2, 4, pszExpectedItem) is None:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ITEM_VALUE",
            "total count item was blank. Path = " + pszInputPath,
        )
    return arrHeader, arrFields


def create_step0004_total_count_files(
    arrStep0003TotalPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """各年月の総計羽数15項目を指定順に横結合し、数値表記を保持します。"""
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszInputPath in arrStep0003TotalPaths:
        pszBaseName, pszItemName, iYear, iMonth, pszSuffix = (
            parse_step0003_total_file_name(pszInputPath)
        )
        objKey = (os.path.dirname(os.path.abspath(pszInputPath)),
                  pszBaseName, iYear, iMonth, pszSuffix)
        dicItems = dicInputs.setdefault(objKey, {})
        if pszItemName not in STEP0004_TOTAL_COUNT_ITEMS:
            continue
        if pszItemName in dicItems:
            raise_conversion_error(
                "DUPLICATE_STEP0003_TOTAL_COUNT_ITEM",
                "duplicate total count item. Path = " + pszInputPath,
            )
        dicItems[pszItemName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicItems = dicInputs[objKey]
            arrMissing = [pszItem for pszItem in STEP0004_TOTAL_COUNT_ITEMS
                          if pszItem not in dicItems]
            if arrMissing:
                raise_conversion_error(
                    "MISSING_STEP0003_TOTAL_COUNT_ITEMS",
                    "missing total count items for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissing),
                )
            arrOutputHeader: List[str] = []
            arrOutputFields: List[str] = []
            for pszItem in STEP0004_TOTAL_COUNT_ITEMS:
                arrHeader, arrFields = read_step0003_total_count_item(
                    dicItems[pszItem], pszItem
                )
                if not arrOutputFields:
                    arrOutputHeader = arrHeader[:4]
                    arrOutputFields = arrFields[:4]
                elif arrFields[:4] != arrOutputFields[:4]:
                    raise_conversion_error(
                        "STEP0003_TOTAL_COMMON_COLUMNS_MISMATCH",
                        "common columns do not match. Path = " + dicItems[pszItem],
                    )
                arrOutputHeader.append(arrHeader[4])
                arrOutputFields.append(arrFields[4])
            pszOutputPath = create_step0004_total_count_output_path(
                dicItems[STEP0004_TOTAL_COUNT_ITEMS[0]]
            )
            write_step0004_tsv(
                pszOutputPath, "\t".join(arrOutputHeader), "\t".join(arrOutputFields)
            )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_TOTAL_COUNT_CLEANUP_FAILED",
                "total count combination failed. Original error = " + str(objException)
                + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0004_total_count_history_output_path(
    pszInputPath: str, pszItemName: str
) -> str:
    pszBaseName, _, _, _, _ = parse_step0003_total_file_name(pszInputPath)
    pszDirectory: str = os.path.dirname(pszInputPath)
    pszStem: str = f"{pszBaseName}_step0004_総計_{pszItemName}_過去05年"
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def create_step0004_total_count_history_files(
    arrStep0003TotalPaths: List[str], iLatestYear: int
) -> List[str]:
    """step0003総計の羽数15項目を、項目別に直近5年の月別表にまとめます。"""
    iFirstYear: int = iLatestYear - 4
    dicInputPaths: Dict[str, Dict[Tuple[int, int], str]] = {
        pszItemName: {} for pszItemName in STEP0004_TOTAL_COUNT_ITEMS
    }
    pszInputBaseName: Optional[str] = None
    pszInputDirectory: Optional[str] = None

    for pszInputPath in arrStep0003TotalPaths:
        pszBaseName, pszItemName, iYear, iMonth, _ = (
            parse_step0003_total_file_name(pszInputPath)
        )
        if (
            pszItemName not in STEP0004_TOTAL_COUNT_ITEMS
            or iYear < iFirstYear
            or iYear > iLatestYear
        ):
            continue
        pszDirectory: str = os.path.dirname(os.path.abspath(pszInputPath))
        if pszInputBaseName is None:
            pszInputBaseName = pszBaseName
            pszInputDirectory = pszDirectory
        elif (pszBaseName, pszDirectory) != (pszInputBaseName, pszInputDirectory):
            raise_conversion_error(
                "MISMATCH_STEP0003_TOTAL_HISTORY_SOURCE",
                "step0003 abdominal-disease inputs must use the same base name "
                "and directory. Path = " + pszInputPath,
            )
        objKey = (iYear, iMonth)
        dicItemPaths = dicInputPaths[pszItemName]
        if objKey in dicItemPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0003_TOTAL_HISTORY_MONTH",
                "duplicate step0003 count item year/month. Item = "
                + pszItemName
                + ", Year/month = "
                + f"{iYear:04d}_{iMonth:02d}",
            )
        dicItemPaths[objKey] = pszInputPath

    if pszInputBaseName is None:
        raise_conversion_error(
            "MISSING_STEP0003_TOTAL_HISTORY_INPUT",
            "no step0003 total count inputs were found for years "
            + f"{iFirstYear:04d}-{iLatestYear:04d}.",
        )

    arrCreatedPaths: List[str] = []
    try:
        for pszItemName in STEP0004_TOTAL_COUNT_ITEMS:
            dicMonthPaths = dicInputPaths[pszItemName]
            if not dicMonthPaths:
                raise_conversion_error(
                    "MISSING_STEP0003_TOTAL_HISTORY_ITEM",
                    "no step0003 total count input was found for item "
                    + pszItemName
                    + " in years "
                    + f"{iFirstYear:04d}-{iLatestYear:04d}.",
                )

            dicValues: Dict[Tuple[int, int], Decimal] = {}
            for (iYear, iMonth), pszInputPath in sorted(dicMonthPaths.items()):
                _, arrFields = read_step0003_total_count_item(
                    pszInputPath, pszItemName
                )
                objValue = parse_decimal_value(arrFields[4], 2, 4, pszItemName)
                if objValue is None:
                    raise_conversion_error(
                        "INVALID_STEP0003_TOTAL_HISTORY_ITEM_VALUE",
                        "step0003 total count item was blank. Item = "
                        + pszItemName
                        + ", Path = "
                        + pszInputPath,
                    )
                dicValues[(iYear, iMonth)] = objValue

            arrRows: List[str] = [
                "\t".join([""] + [f"{iMonth}月" for iMonth in range(1, 13)])
            ]
            for iYear in range(iLatestYear, iFirstYear - 1, -1):
                arrYearValues: List[str] = [f"{iYear}年"]
                for iMonth in range(1, 13):
                    objValue = dicValues.get((iYear, iMonth), Decimal("0"))
                    arrYearValues.append(format(objValue, ",f"))
                arrRows.append("\t".join(arrYearValues))

            pszOutputPath: str = create_step0004_total_count_history_output_path(
                next(iter(dicMonthPaths.values())), pszItemName
            )
            try:
                write_monthly_tsv(pszOutputPath, pszItemName, arrRows)
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0004_TOTAL_COUNT_HISTORY_CREATE_FAILED",
                    "failed to write step0004 total count history TSV. Item = "
                    + pszItemName
                    + ", Path = "
                    + pszOutputPath
                    + "; Detail = "
                    + str(objException),
                )
            arrCreatedPaths.append(pszOutputPath)

        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_TOTAL_COUNT_HISTORY_CLEANUP_FAILED",
                "step0004 total count history creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0004_total_count_history_file_name(
    pszInputPath: str,
) -> Tuple[str, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0004_総計_(.+)_過去05年((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0004_TOTAL_COUNT_HISTORY_FILE_NAME",
            "invalid step0004 total count history file name. Path = "
            + pszInputPath,
        )
    pszItemName = objMatch.group(2)
    if pszItemName not in STEP0004_TOTAL_COUNT_ITEMS:
        raise_conversion_error(
            "INVALID_STEP0004_TOTAL_COUNT_HISTORY_ITEM",
            "unexpected step0004 total count history item. Path = "
            + pszInputPath,
        )
    return objMatch.group(1), pszItemName


def create_step0005_total_count_history_output_path(
    pszInputPath: str,
) -> str:
    pszBaseName, _ = parse_step0004_total_count_history_file_name(pszInputPath)
    pszDirectory: str = os.path.dirname(pszInputPath)
    pszStem: str = (
        f"{pszBaseName}_step0005_総計_出荷廃棄_羽数_過去05年"
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def read_step0004_total_count_history_file(
    pszInputPath: str, pszExpectedItem: str
) -> List[str]:
    _, pszItemName = parse_step0004_total_count_history_file_name(pszInputPath)
    if pszItemName != pszExpectedItem:
        raise_conversion_error(
            "STEP0004_TOTAL_COUNT_HISTORY_ITEM_MISMATCH",
            "unexpected step0004 total count history item. Expected = "
            + pszExpectedItem
            + ", Path = "
            + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0004_TOTAL_COUNT_HISTORY_READ_FAILED",
            "failed to read step0004 total count history TSV. Path = "
            + pszInputPath
            + "; Detail = "
            + str(objException),
        )
    if len(arrLines) != 7:
        raise_conversion_error(
            "INVALID_STEP0004_TOTAL_COUNT_HISTORY_ROW_COUNT",
            "step0004 total count history TSV must contain seven rows. Path = "
            + pszInputPath,
        )
    if arrLines[0] != pszExpectedItem:
        raise_conversion_error(
            "INVALID_STEP0004_TOTAL_COUNT_HISTORY_HEADER",
            "step0004 total count history title does not match its item. Path = "
            + pszInputPath,
        )
    arrExpectedMonths: List[str] = [""] + [
        f"{iMonth}月" for iMonth in range(1, 13)
    ]
    if arrLines[1].split("\t") != arrExpectedMonths:
        raise_conversion_error(
            "INVALID_STEP0004_TOTAL_COUNT_HISTORY_MONTH_HEADER",
            "unexpected step0004 total count history month header. Path = "
            + pszInputPath,
        )
    for iLineIndex, pszLine in enumerate(arrLines[2:], start=3):
        arrFields = pszLine.split("\t")
        if len(arrFields) != 13 or not re.fullmatch(r"\d{4}年", arrFields[0]):
            raise_conversion_error(
                "INVALID_STEP0004_TOTAL_COUNT_HISTORY_DATA_ROW",
                "invalid step0004 total count history data row. Line = "
                + str(iLineIndex)
                + ", Path = "
                + pszInputPath,
            )
    return arrLines


def create_step0005_total_count_history_files(
    arrStep0004TotalCountHistoryPaths: List[str],
) -> List[str]:
    """15個のstep0004総計羽数過去5年表を項目順に縦結合します。"""
    dicInputPaths: Dict[str, str] = {}
    pszInputBaseName: Optional[str] = None
    pszInputDirectory: Optional[str] = None
    for pszInputPath in arrStep0004TotalCountHistoryPaths:
        pszBaseName, pszItemName = parse_step0004_total_count_history_file_name(
            pszInputPath
        )
        pszDirectory = os.path.dirname(os.path.abspath(pszInputPath))
        if pszInputBaseName is None:
            pszInputBaseName = pszBaseName
            pszInputDirectory = pszDirectory
        elif (pszBaseName, pszDirectory) != (pszInputBaseName, pszInputDirectory):
            raise_conversion_error(
                "MISMATCH_STEP0004_TOTAL_COUNT_HISTORY_SOURCE",
                "step0004 total count history inputs must use the same base name "
                "and directory. Path = "
                + pszInputPath,
            )
        if pszItemName in dicInputPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0004_TOTAL_COUNT_HISTORY_ITEM",
                "duplicate step0004 total count history item. Item = "
                + pszItemName,
            )
        dicInputPaths[pszItemName] = pszInputPath

    arrMissingItems = [
        pszItemName
        for pszItemName in STEP0004_TOTAL_COUNT_ITEMS
        if pszItemName not in dicInputPaths
    ]
    if arrMissingItems:
        raise_conversion_error(
            "MISSING_STEP0004_TOTAL_COUNT_HISTORY_ITEMS",
            "missing step0004 total count history inputs. Items = "
            + ", ".join(arrMissingItems),
        )

    arrOutputLines: List[str] = []
    arrExpectedYears: Optional[List[str]] = None
    for iItemIndex, pszItemName in enumerate(STEP0004_TOTAL_COUNT_ITEMS):
        pszInputPath = dicInputPaths[pszItemName]
        arrLines = read_step0004_total_count_history_file(
            pszInputPath, pszItemName
        )
        arrYears = [pszLine.split("\t", 1)[0] for pszLine in arrLines[2:]]
        if arrExpectedYears is None:
            arrExpectedYears = arrYears
        elif arrYears != arrExpectedYears:
            raise_conversion_error(
                "STEP0004_TOTAL_COUNT_HISTORY_YEAR_MISMATCH",
                "year rows do not match across step0004 total count history inputs. "
                "Path = "
                + pszInputPath,
            )
        if iItemIndex > 0:
            arrOutputLines.append("")
        arrOutputLines.extend(arrLines)

    pszOutputPath = create_step0005_total_count_history_output_path(
        dicInputPaths[STEP0004_TOTAL_COUNT_ITEMS[0]]
    )
    try:
        write_monthly_tsv(pszOutputPath, arrOutputLines[0], arrOutputLines[1:])
    except ConversionError as objException:
        raise_conversion_error(
            "STEP0005_TOTAL_COUNT_HISTORY_CREATE_FAILED",
            "failed to write step0005 total count history TSV. Path = "
            + pszOutputPath
            + "; Detail = "
            + str(objException),
        )
    return [pszOutputPath]


def create_step0005_total_count_history_excel_output_path(
    pszStep0005TsvPath: str,
) -> str:
    pszBaseName: str = os.path.basename(pszStep0005TsvPath)
    objMatch = re.fullmatch(
        r"(.+_step0005_総計_出荷廃棄_羽数_過去05年(?:_\d+)*)\.tsv",
        pszBaseName,
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0005_TOTAL_COUNT_HISTORY_FILE_NAME",
            "unexpected step0005 total count history file name. Path = "
            + pszStep0005TsvPath,
        )
    pszOutputPath: str = os.path.join(
        os.path.dirname(pszStep0005TsvPath), objMatch.group(1) + ".xlsx"
    )
    pszStem: str = os.path.splitext(pszOutputPath)[0]
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = f"{pszStem}_{iSuffix:02d}.xlsx"
        iSuffix += 1
    return pszOutputPath


def read_step0005_total_count_history_for_excel(
    pszInputPath: str,
) -> List[List[str]]:
    if re.fullmatch(
        r".+_step0005_総計_出荷廃棄_羽数_過去05年(?:_\d+)*\.tsv",
        os.path.basename(pszInputPath),
    ) is None:
        raise_conversion_error(
            "INVALID_STEP0005_TOTAL_COUNT_HISTORY_FILE_NAME",
            "unexpected step0005 total count history file name. Path = "
            + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0005_TOTAL_COUNT_HISTORY_EXCEL_READ_FAILED",
            "failed to read step0005 total count history TSV. Path = "
            + pszInputPath
            + "; Detail = "
            + str(objException),
        )

    iExpectedLineCount: int = len(STEP0004_TOTAL_COUNT_ITEMS) * 7 + len(
        STEP0004_TOTAL_COUNT_ITEMS
    ) - 1
    if len(arrLines) != iExpectedLineCount:
        raise_conversion_error(
            "INVALID_STEP0005_TOTAL_COUNT_HISTORY_ROW_COUNT",
            "step0005 total count history TSV has an unexpected row count. Path = "
            + pszInputPath
            + ", Actual = "
            + str(len(arrLines)),
        )

    arrRows: List[List[str]] = []
    iLineIndex: int = 0
    arrExpectedYears: Optional[List[str]] = None
    for iItemIndex, pszItemName in enumerate(STEP0004_TOTAL_COUNT_ITEMS):
        if iItemIndex > 0:
            if arrLines[iLineIndex] != "":
                raise_conversion_error(
                    "INVALID_STEP0005_TOTAL_COUNT_HISTORY_SEPARATOR",
                    "expected a blank separator row in step0005 total count history TSV. Line = "
                    + str(iLineIndex + 1)
                    + ", Path = "
                    + pszInputPath,
                )
            arrRows.append([""])
            iLineIndex += 1
        arrBlock: List[str] = arrLines[iLineIndex : iLineIndex + 7]
        if len(arrBlock) != 7 or arrBlock[0] != pszItemName:
            raise_conversion_error(
                "INVALID_STEP0005_TOTAL_COUNT_HISTORY_BLOCK",
                "unexpected item block in step0005 total count history TSV. Item = "
                + pszItemName
                + ", Path = "
                + pszInputPath,
            )
        arrMonthFields: List[str] = arrBlock[1].split("\t")
        if arrMonthFields != [""] + [f"{iMonth}月" for iMonth in range(1, 13)]:
            raise_conversion_error(
                "INVALID_STEP0005_TOTAL_COUNT_HISTORY_MONTH_HEADER",
                "unexpected month header in step0005 total count history TSV. Path = "
                + pszInputPath,
            )
        arrYearRows: List[str] = arrBlock[2:]
        if any(
            re.fullmatch(r"\d{4}年", pszLine.split("\t", 1)[0]) is None
            for pszLine in arrYearRows
        ):
            raise_conversion_error(
                "INVALID_STEP0005_TOTAL_COUNT_HISTORY_DATA_ROW",
                "invalid year label in step0005 total count history TSV. Path = "
                + pszInputPath,
            )
        arrYears: List[str] = [pszLine.split("\t", 1)[0] for pszLine in arrYearRows]
        if any(
            int(arrYears[iYearIndex][:-1]) - int(arrYears[iYearIndex + 1][:-1]) != 1
            for iYearIndex in range(len(arrYears) - 1)
        ):
            raise_conversion_error(
                "INVALID_STEP0005_TOTAL_COUNT_HISTORY_DATA_ROW",
                "step0005 total count history years must be consecutive and descending. Path = "
                + pszInputPath,
            )
        if arrExpectedYears is None:
            arrExpectedYears = arrYears
        elif arrYears != arrExpectedYears:
            raise_conversion_error(
                "STEP0005_TOTAL_COUNT_HISTORY_YEAR_MISMATCH",
                "year rows do not match across step0005 item blocks. Path = "
                + pszInputPath,
            )
        arrRows.append([pszItemName])
        arrRows.append(arrMonthFields)
        for pszLine in arrYearRows:
            arrFields: List[str] = pszLine.split("\t")
            if len(arrFields) != 13:
                raise_conversion_error(
                    "INVALID_STEP0005_TOTAL_COUNT_HISTORY_DATA_ROW",
                    "unexpected year row in step0005 total count history TSV. Path = "
                    + pszInputPath,
                )
            arrRows.append(arrFields)
        iLineIndex += 7
    return arrRows


def write_step0005_total_count_history_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0005_TOTAL_COUNT_HISTORY_EXCEL_CREATE_FAILED",
                "step0005 total count history Excel already exists. Path = "
                + pszOutputPath,
            )
        iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
            prefix="group_shipment_summary_step0005_history_",
            suffix=".xlsx",
            dir=os.path.dirname(pszOutputPath),
        )
        os.close(iTemporaryFileDescriptor)
        objWorkbook = Workbook()
        objWorksheet = objWorkbook.active
        objWorksheet.title = "総計羽数過去5年"
        for iRowIndex, arrFields in enumerate(arrRows, start=1):
            bYearDataRow: bool = bool(
                arrFields and re.fullmatch(r"\d{4}年", arrFields[0])
            )
            for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                if iColumnIndex == 1 and re.fullmatch(r"\d{4}年", pszValue):
                    objCell.value = int(pszValue[:-1])
                    objCell.number_format = '0"年"'
                elif bYearDataRow and iColumnIndex > 1:
                    pszNumericValue: str = pszValue.replace(",", "")
                    try:
                        objNumericValue = Decimal(pszNumericValue)
                    except InvalidOperation:
                        objNumericValue = Decimal("NaN")
                    if not objNumericValue.is_finite():
                        raise_conversion_error(
                            "INVALID_STEP0005_TOTAL_COUNT_HISTORY_VALUE",
                            "non-numeric value in step0005 total count history TSV. Value = "
                            + pszValue
                            + ", Row = "
                            + str(iRowIndex),
                        )
                    if objNumericValue == objNumericValue.to_integral_value():
                        objCell.value = int(objNumericValue)
                        objCell.number_format = "#,##0"
                    else:
                        objCell.value = float(objNumericValue)
                        objCell.number_format = "#,##0.##########"
                else:
                    objCell.value = pszValue if pszValue != "" else None
        objWorkbook.save(pszTemporaryPath)
        objWorkbook.close()
        objWorkbook = None
        os.replace(pszTemporaryPath, pszOutputPath)
        pszTemporaryPath = ""
        bOutputReserved = False
    except ConversionError:
        raise
    except Exception as objException:
        raise_conversion_error(
            "STEP0005_TOTAL_COUNT_HISTORY_EXCEL_WRITE_FAILED",
            "failed to write step0005 total count history Excel. Path = "
            + pszOutputPath
            + "; Detail = "
            + str(objException),
        )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0005_total_count_history_excel_files(
    arrStep0005TsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszInputPath in arrStep0005TsvPaths:
            arrRows = read_step0005_total_count_history_for_excel(pszInputPath)
            pszOutputPath = create_step0005_total_count_history_excel_output_path(
                pszInputPath
            )
            write_step0005_total_count_history_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0005_TOTAL_COUNT_HISTORY_EXCEL_CLEANUP_FAILED",
                "step0005 total count history Excel creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0004_group_discard_count_output_path(
    pszFirstInputPath: str, pszGroupName: str
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0003_グループ別_.+_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszFirstInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0003_GROUP_FILE_NAME",
            "invalid step0003 group file name. Path = " + pszFirstInputPath,
        )
    if pszGroupName not in {pszName for _, pszName in GROUPS}:
        raise_conversion_error(
            "INVALID_STEP0004_DISCARD_COUNT_GROUP",
            "unexpected group name. Group = " + pszGroupName,
        )
    pszBaseName, pszYear, pszMonth, pszExistingSuffix = objMatch.groups()
    pszDirectory = os.path.dirname(pszFirstInputPath)
    pszStem = (
        f"{pszBaseName}_step0004_グループ別_廃棄データ別_"
        f"{pszGroupName}_羽数_{pszYear}_{pszMonth}{pszExistingSuffix}"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0004_group_discard_count_files(
    arrStep0003Paths: List[str],
    arrStep0003YearMonths: List[Tuple[int, int]],
) -> List[str]:
    """各年月・グループの廃棄羽数15項目を横結合したTSVを作成します。"""
    dicStep0003Paths = build_step0003_path_map(
        arrStep0003Paths, arrStep0003YearMonths
    )
    arrCreatedPaths: List[str] = []
    try:
        for iYear, iMonth in arrStep0003YearMonths:
            dicItemRows: Dict[str, Dict[Tuple[str, str], List[str]]] = {}
            for pszItemName in STEP0004_TOTAL_COUNT_ITEMS:
                objPathKey = (iYear, iMonth, pszItemName)
                if objPathKey not in dicStep0003Paths:
                    raise_conversion_error(
                        "STEP0003_GROUP_INPUT_MISSING",
                        "required step0003 group input was missing. Key = "
                        + str(objPathKey),
                    )
                dicItemRows[pszItemName] = read_step0003_for_step0004(
                    dicStep0003Paths[objPathKey],
                    pszItemName,
                    False,
                    iYear,
                    iMonth,
                )

            for pszGroupNumber, pszGroupName in GROUPS:
                objGroupKey = (pszGroupNumber, pszGroupName)
                arrCommonFields = dicItemRows[STEP0004_TOTAL_COUNT_ITEMS[0]][
                    objGroupKey
                ][:4]
                for pszItemName in STEP0004_TOTAL_COUNT_ITEMS[1:]:
                    if dicItemRows[pszItemName][objGroupKey][:4] != arrCommonFields:
                        raise_conversion_error(
                            "STEP0003_GROUP_DISCARD_KEY_MISMATCH",
                            "step0003 group keys did not match. Year/month = "
                            + f"{iYear:04d}_{iMonth:02d}"
                            + ", Group = " + str(objGroupKey)
                            + ", Item = " + pszItemName,
                        )
                arrOutputHeader = [
                    "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
                ]
                arrOutputFields = list(arrCommonFields)
                for pszItemName in STEP0004_TOTAL_COUNT_ITEMS:
                    arrOutputHeader.append(pszItemName)
                    arrOutputFields.append(
                        dicItemRows[pszItemName][objGroupKey][4]
                    )
                pszOutputPath = create_step0004_group_discard_count_output_path(
                    dicStep0003Paths[
                        (iYear, iMonth, STEP0004_TOTAL_COUNT_ITEMS[0])
                    ],
                    pszGroupName,
                )
                write_step0004_tsv(
                    pszOutputPath,
                    "\t".join(arrOutputHeader),
                    "\t".join(arrOutputFields),
                )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_GROUP_DISCARD_COUNT_CLEANUP_FAILED",
                "step0004 group discard count creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0004_group_discard_weight_output_path(
    pszFirstInputPath: str, pszGroupName: str
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0003_グループ別_.+_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszFirstInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0003_GROUP_FILE_NAME",
            "invalid step0003 group file name. Path = " + pszFirstInputPath,
        )
    if pszGroupName not in {pszName for _, pszName in GROUPS}:
        raise_conversion_error(
            "INVALID_STEP0004_DISCARD_WEIGHT_GROUP",
            "unexpected group name. Group = " + pszGroupName,
        )
    pszBaseName, pszYear, pszMonth, pszExistingSuffix = objMatch.groups()
    pszDirectory = os.path.dirname(pszFirstInputPath)
    pszStem = (
        f"{pszBaseName}_step0004_グループ別_廃棄データ別_"
        f"{pszGroupName}_重量_{pszYear}_{pszMonth}{pszExistingSuffix}"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0004_group_discard_weight_files(
    arrStep0003Paths: List[str],
    arrStep0003YearMonths: List[Tuple[int, int]],
) -> List[str]:
    """各年月・グループの廃棄重量15項目を横結合したTSVを作成します。"""
    dicStep0003Paths = build_step0003_path_map(
        arrStep0003Paths, arrStep0003YearMonths
    )
    arrCreatedPaths: List[str] = []
    try:
        for iYear, iMonth in arrStep0003YearMonths:
            dicItemRows: Dict[str, Dict[Tuple[str, str], List[str]]] = {}
            for pszItemName in STEP0004_TOTAL_WEIGHT_ITEMS:
                objPathKey = (iYear, iMonth, pszItemName)
                if objPathKey not in dicStep0003Paths:
                    raise_conversion_error(
                        "STEP0003_GROUP_INPUT_MISSING",
                        "required step0003 group input was missing. Key = "
                        + str(objPathKey),
                    )
                dicItemRows[pszItemName] = read_step0003_for_step0004(
                    dicStep0003Paths[objPathKey],
                    pszItemName,
                    False,
                    iYear,
                    iMonth,
                )

            for pszGroupNumber, pszGroupName in GROUPS:
                objGroupKey = (pszGroupNumber, pszGroupName)
                arrCommonFields = dicItemRows[STEP0004_TOTAL_WEIGHT_ITEMS[0]][
                    objGroupKey
                ][:4]
                for pszItemName in STEP0004_TOTAL_WEIGHT_ITEMS[1:]:
                    if dicItemRows[pszItemName][objGroupKey][:4] != arrCommonFields:
                        raise_conversion_error(
                            "STEP0003_GROUP_DISCARD_WEIGHT_KEY_MISMATCH",
                            "step0003 group keys did not match. Year/month = "
                            + f"{iYear:04d}_{iMonth:02d}"
                            + ", Group = " + str(objGroupKey)
                            + ", Item = " + pszItemName,
                        )
                arrOutputHeader = [
                    "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
                ]
                arrOutputFields = list(arrCommonFields)
                for pszItemName in STEP0004_TOTAL_WEIGHT_ITEMS:
                    arrOutputHeader.append(pszItemName)
                    arrOutputFields.append(
                        dicItemRows[pszItemName][objGroupKey][4]
                    )
                pszOutputPath = create_step0004_group_discard_weight_output_path(
                    dicStep0003Paths[
                        (iYear, iMonth, STEP0004_TOTAL_WEIGHT_ITEMS[0])
                    ],
                    pszGroupName,
                )
                write_step0004_tsv(
                    pszOutputPath,
                    "\t".join(arrOutputHeader),
                    "\t".join(arrOutputFields),
                )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_GROUP_DISCARD_WEIGHT_CLEANUP_FAILED",
                "step0004 group discard weight creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS: List[str] = [
    "懸鳥前死鶏羽数", "懸鳥前発育不良羽数",
]


def create_step0004_total_pre_slaughter_count_output_path(
    pszFirstInputPath: str,
) -> str:
    pszBaseName, _, iYear, iMonth, pszSuffix = parse_step0003_total_file_name(
        pszFirstInputPath
    )
    pszDirectory: str = os.path.dirname(pszFirstInputPath)
    pszStem: str = (
        f"{pszBaseName}_step0004_総計_懸鳥前データ別_羽数_{iYear:04d}_{iMonth:02d}"
        + pszSuffix
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def create_step0004_total_pre_slaughter_count_files(
    arrStep0003TotalPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """各年月の懸鳥前総計羽数2項目を指定順に横結合し、数値表記を保持します。"""
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszInputPath in arrStep0003TotalPaths:
        pszBaseName, pszItemName, iYear, iMonth, pszSuffix = (
            parse_step0003_total_file_name(pszInputPath)
        )
        objKey = (os.path.dirname(os.path.abspath(pszInputPath)),
                  pszBaseName, iYear, iMonth, pszSuffix)
        dicItems = dicInputs.setdefault(objKey, {})
        if pszItemName not in STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS:
            continue
        if pszItemName in dicItems:
            raise_conversion_error(
                "DUPLICATE_STEP0003_TOTAL_PRE_SLAUGHTER_COUNT_ITEM",
                "duplicate pre-slaughter total count item. Path = " + pszInputPath,
            )
        dicItems[pszItemName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicItems = dicInputs[objKey]
            arrMissing = [pszItem for pszItem in STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS
                          if pszItem not in dicItems]
            if arrMissing:
                raise_conversion_error(
                    "MISSING_STEP0003_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS",
                    "missing pre-slaughter total count items for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissing),
                )
            arrOutputHeader: List[str] = []
            arrOutputFields: List[str] = []
            for pszItem in STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS:
                arrHeader, arrFields = read_step0003_total_count_item(
                    dicItems[pszItem], pszItem
                )
                if not arrOutputFields:
                    arrOutputHeader = arrHeader[:4]
                    arrOutputFields = arrFields[:4]
                elif arrFields[:4] != arrOutputFields[:4]:
                    raise_conversion_error(
                        "STEP0003_TOTAL_COMMON_COLUMNS_MISMATCH",
                        "common columns do not match. Path = " + dicItems[pszItem],
                    )
                arrOutputHeader.append(arrHeader[4])
                arrOutputFields.append(arrFields[4])
            pszOutputPath = create_step0004_total_pre_slaughter_count_output_path(
                dicItems[STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS[0]]
            )
            write_step0004_tsv(
                pszOutputPath, "\t".join(arrOutputHeader), "\t".join(arrOutputFields)
            )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_CLEANUP_FAILED",
                "pre-slaughter total count combination failed. Original error = " + str(objException)
                + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS: List[str] = [
    "懸鳥前死鶏重量", "懸鳥前発育不良重量",
]


def create_step0004_total_pre_slaughter_weight_output_path(
    pszFirstInputPath: str,
) -> str:
    pszBaseName, _, iYear, iMonth, pszSuffix = parse_step0003_total_file_name(
        pszFirstInputPath
    )
    pszDirectory: str = os.path.dirname(pszFirstInputPath)
    pszStem: str = (
        f"{pszBaseName}_step0004_総計_懸鳥前データ別_重量_{iYear:04d}_{iMonth:02d}"
        + pszSuffix
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def create_step0004_total_pre_slaughter_weight_files(
    arrStep0003TotalPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """各年月の懸鳥前総計重量2項目を指定順に横結合し、数値表記を保持します。"""
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszInputPath in arrStep0003TotalPaths:
        pszBaseName, pszItemName, iYear, iMonth, pszSuffix = (
            parse_step0003_total_file_name(pszInputPath)
        )
        objKey = (os.path.dirname(os.path.abspath(pszInputPath)),
                  pszBaseName, iYear, iMonth, pszSuffix)
        dicItems = dicInputs.setdefault(objKey, {})
        if pszItemName not in STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS:
            continue
        if pszItemName in dicItems:
            raise_conversion_error(
                "DUPLICATE_STEP0003_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEM",
                "duplicate pre-slaughter total weight item. Path = " + pszInputPath,
            )
        dicItems[pszItemName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicItems = dicInputs[objKey]
            arrMissing = [pszItem for pszItem in STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS
                          if pszItem not in dicItems]
            if arrMissing:
                raise_conversion_error(
                    "MISSING_STEP0003_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS",
                    "missing pre-slaughter total weight items for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissing),
                )
            arrOutputHeader: List[str] = []
            arrOutputFields: List[str] = []
            for pszItem in STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS:
                arrHeader, arrFields = read_step0003_total_weight_item(
                    dicItems[pszItem], pszItem
                )
                if not arrOutputFields:
                    arrOutputHeader = arrHeader[:4]
                    arrOutputFields = arrFields[:4]
                elif arrFields[:4] != arrOutputFields[:4]:
                    raise_conversion_error(
                        "STEP0003_TOTAL_COMMON_COLUMNS_MISMATCH",
                        "common columns do not match. Path = " + dicItems[pszItem],
                    )
                arrOutputHeader.append(arrHeader[4])
                arrOutputFields.append(arrFields[4])
            pszOutputPath = create_step0004_total_pre_slaughter_weight_output_path(
                dicItems[STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS[0]]
            )
            write_step0004_tsv(
                pszOutputPath, "\t".join(arrOutputHeader), "\t".join(arrOutputFields)
            )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_CLEANUP_FAILED",
                "pre-slaughter total weight combination failed. Original error = " + str(objException)
                + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0004_pre_slaughter_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0004_総計_懸鳥前データ別_(羽数|重量)_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0004_PRE_SLAUGHTER_FILE_NAME",
            "invalid pre-slaughter step0004 file name. Path = " + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return objMatch.group(1), objMatch.group(2), iYear, iMonth, objMatch.group(5)


def read_step0004_pre_slaughter_subtotal_input(
    pszInputPath: str,
) -> Tuple[List[str], List[str], Decimal]:
    _, pszKind, iFileYear, iFileMonth, _ = parse_step0004_pre_slaughter_file_name(
        pszInputPath
    )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0004_PRE_SLAUGHTER_READ_FAILED",
            "failed to read pre-slaughter step0004. Path = " + pszInputPath
            + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0004_PRE_SLAUGHTER_ROW_COUNT",
            "pre-slaughter step0004 must contain one header and one total row. Path = "
            + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    if len(arrHeader) != 6 or len(arrFields) != 6:
        raise_conversion_error(
            "INVALID_STEP0004_PRE_SLAUGHTER_COLUMN_COUNT",
            "pre-slaughter step0004 must contain six columns. Path = " + pszInputPath,
        )
    arrExpectedHeader = ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
                         "懸鳥前死鶏" + pszKind, "懸鳥前発育不良" + pszKind]
    if arrHeader != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0004_PRE_SLAUGHTER_HEADER",
            "unexpected pre-slaughter step0004 header. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0004_PRE_SLAUGHTER_YEAR_MONTH_MISMATCH",
            "file name and row year/month do not match. Path = " + pszInputPath,
        )
    if arrFields[2:4] != ["", "総計"]:
        raise_conversion_error(
            "INVALID_STEP0004_PRE_SLAUGHTER_GROUP",
            "expected an empty group number and total group name. Path = " + pszInputPath,
        )
    objSubtotal = Decimal("0")
    for iColumnIndex in (4, 5):
        objValue = parse_decimal_value(arrFields[iColumnIndex], 2, iColumnIndex,
                                       arrHeader[iColumnIndex])
        if objValue is None:
            raise_conversion_error(
                "INVALID_PRE_SLAUGHTER_SUBTOTAL_VALUE",
                "subtotal source value was blank. Path = " + pszInputPath,
            )
        objSubtotal += objValue
    return arrHeader, arrFields, objSubtotal


def create_step0005_pre_slaughter_subtotal_output_path(pszInputPath: str) -> str:
    pszBaseName, pszKind, iYear, iMonth, pszSuffix = (
        parse_step0004_pre_slaughter_file_name(pszInputPath)
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszStem = (f"{pszBaseName}_step0005_総計_懸鳥前データ別_{pszKind}_{iYear:04d}_{iMonth:02d}"
               + pszSuffix)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def create_step0005_pre_slaughter_subtotal_files(
    arrCountPaths: List[str], arrWeightPaths: List[str],
) -> Tuple[List[str], List[str], List[Tuple[int, int]]]:
    """懸鳥前の羽数・重量を別々に加算し、入力列を保持して小計列を追加します。"""
    dicPairs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszExpectedKind, arrPaths in [("羽数", arrCountPaths), ("重量", arrWeightPaths)]:
        for pszPath in arrPaths:
            pszBase, pszKind, iYear, iMonth, pszSuffix = (
                parse_step0004_pre_slaughter_file_name(pszPath)
            )
            if pszKind != pszExpectedKind:
                raise_conversion_error(
                    "INVALID_PRE_SLAUGHTER_SUBTOTAL_INPUT_KIND",
                    "unexpected subtotal input kind. Path = " + pszPath,
                )
            objKey = (os.path.dirname(os.path.abspath(pszPath)), pszBase,
                      iYear, iMonth, pszSuffix)
            dicPair = dicPairs.setdefault(objKey, {})
            if pszKind in dicPair:
                raise_conversion_error(
                    "DUPLICATE_PRE_SLAUGHTER_SUBTOTAL_INPUT",
                    "duplicate subtotal input. Path = " + pszPath,
                )
            dicPair[pszKind] = pszPath

    arrCreatedPaths: List[str] = []
    arrCountOutputPaths: List[str] = []
    arrWeightOutputPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicPairs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicPair = dicPairs[objKey]
            if set(dicPair) != {"羽数", "重量"}:
                raise_conversion_error(
                    "MISSING_PRE_SLAUGHTER_SUBTOTAL_INPUT",
                    "count and weight inputs must be paired for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}",
                )
            arrCountHeader, arrCountFields, objCountSubtotal = (
                read_step0004_pre_slaughter_subtotal_input(dicPair["羽数"])
            )
            arrWeightHeader, arrWeightFields, objWeightSubtotal = (
                read_step0004_pre_slaughter_subtotal_input(dicPair["重量"])
            )
            if arrCountFields[:4] != arrWeightFields[:4]:
                raise_conversion_error(
                    "PRE_SLAUGHTER_SUBTOTAL_COMMON_COLUMNS_MISMATCH",
                    "count and weight common columns do not match. Path = " + dicPair["重量"],
                )
            pszCountSubtotal = format(Decimal(format_sum_value(objCountSubtotal)), ",f")
            pszWeightSubtotal = format(
                objWeightSubtotal.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP), ",.1f"
            )
            for pszKind, arrHeader, arrFields, pszSubtotal, arrOutputPaths in [
                ("羽数", arrCountHeader, arrCountFields, pszCountSubtotal, arrCountOutputPaths),
                ("重量", arrWeightHeader, arrWeightFields, pszWeightSubtotal, arrWeightOutputPaths),
            ]:
                pszOutputPath = create_step0005_pre_slaughter_subtotal_output_path(
                    dicPair[pszKind]
                )
                write_step0005_tsv(pszOutputPath, "\t".join(arrHeader + ["小計"]),
                                   ["\t".join(arrFields + [pszSubtotal])])
                arrCreatedPaths.append(pszOutputPath)
                arrOutputPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCountOutputPaths, arrWeightOutputPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "PRE_SLAUGHTER_SUBTOTAL_CLEANUP_FAILED",
                "pre-slaughter subtotal failed. Original error = " + str(objException)
                + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0004_discard_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0004_総計_廃棄データ別_(羽数|重量)_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0004_DISCARD_FILE_NAME",
            "invalid discard step0004 file name. Path = " + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return objMatch.group(1), objMatch.group(2), iYear, iMonth, objMatch.group(5)


def read_step0004_discard_subtotal_input(
    pszInputPath: str,
) -> Tuple[List[str], List[str], Decimal]:
    _, pszKind, iFileYear, iFileMonth, _ = parse_step0004_discard_file_name(
        pszInputPath
    )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0004_DISCARD_READ_FAILED",
            "failed to read discard step0004. Path = " + pszInputPath
            + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0004_DISCARD_ROW_COUNT",
            "discard step0004 must contain one header and one total row. Path = "
            + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    arrExpectedItems = (STEP0004_TOTAL_COUNT_ITEMS if pszKind == "羽数"
                        else STEP0004_TOTAL_WEIGHT_ITEMS)
    iExpectedColumnCount = 4 + len(arrExpectedItems)
    if len(arrHeader) != iExpectedColumnCount or len(arrFields) != iExpectedColumnCount:
        raise_conversion_error(
            "INVALID_STEP0004_DISCARD_COLUMN_COUNT",
            "unexpected discard step0004 column count. Path = " + pszInputPath,
        )
    arrExpectedHeader = ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名"] + arrExpectedItems
    if arrHeader != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0004_DISCARD_HEADER",
            "unexpected discard step0004 header. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0004_DISCARD_YEAR_MONTH_MISMATCH",
            "file name and row year/month do not match. Path = " + pszInputPath,
        )
    if arrFields[2:4] != ["", "総計"]:
        raise_conversion_error(
            "INVALID_STEP0004_DISCARD_GROUP",
            "expected an empty group number and total group name. Path = " + pszInputPath,
        )
    objSubtotal = Decimal("0")
    for iColumnIndex in range(4, len(arrFields)):
        objValue = parse_decimal_value(arrFields[iColumnIndex], 2, iColumnIndex,
                                       arrHeader[iColumnIndex])
        if objValue is None:
            raise_conversion_error(
                "INVALID_DISCARD_SUBTOTAL_VALUE",
                "subtotal source value was blank. Path = " + pszInputPath,
            )
        objSubtotal += objValue
    return arrHeader, arrFields, objSubtotal


def create_step0005_discard_subtotal_output_path(pszInputPath: str) -> str:
    pszBaseName, pszKind, iYear, iMonth, pszSuffix = (
        parse_step0004_discard_file_name(pszInputPath)
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszStem = (f"{pszBaseName}_step0005_総計_廃棄データ別_{pszKind}_{iYear:04d}_{iMonth:02d}"
               + pszSuffix)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def create_step0005_discard_subtotal_files(
    arrCountPaths: List[str], arrWeightPaths: List[str],
) -> Tuple[List[str], List[str], List[Tuple[int, int]]]:
    """予備２を含む廃棄羽数・重量各15項目を加算し、小計列を追加します。"""
    dicPairs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszExpectedKind, arrPaths in [("羽数", arrCountPaths), ("重量", arrWeightPaths)]:
        for pszPath in arrPaths:
            pszBase, pszKind, iYear, iMonth, pszSuffix = (
                parse_step0004_discard_file_name(pszPath)
            )
            if pszKind != pszExpectedKind:
                raise_conversion_error(
                    "INVALID_DISCARD_SUBTOTAL_INPUT_KIND",
                    "unexpected subtotal input kind. Path = " + pszPath,
                )
            objKey = (os.path.dirname(os.path.abspath(pszPath)), pszBase,
                      iYear, iMonth, pszSuffix)
            dicPair = dicPairs.setdefault(objKey, {})
            if pszKind in dicPair:
                raise_conversion_error(
                    "DUPLICATE_DISCARD_SUBTOTAL_INPUT",
                    "duplicate subtotal input. Path = " + pszPath,
                )
            dicPair[pszKind] = pszPath

    arrCreatedPaths: List[str] = []
    arrCountOutputPaths: List[str] = []
    arrWeightOutputPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicPairs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicPair = dicPairs[objKey]
            if set(dicPair) != {"羽数", "重量"}:
                raise_conversion_error(
                    "MISSING_DISCARD_SUBTOTAL_INPUT",
                    "count and weight inputs must be paired for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}",
                )
            arrCountHeader, arrCountFields, objCountSubtotal = (
                read_step0004_discard_subtotal_input(dicPair["羽数"])
            )
            arrWeightHeader, arrWeightFields, objWeightSubtotal = (
                read_step0004_discard_subtotal_input(dicPair["重量"])
            )
            if arrCountFields[:4] != arrWeightFields[:4]:
                raise_conversion_error(
                    "DISCARD_SUBTOTAL_COMMON_COLUMNS_MISMATCH",
                    "count and weight common columns do not match. Path = " + dicPair["重量"],
                )
            pszCountSubtotal = format(Decimal(format_sum_value(objCountSubtotal)), ",f")
            pszWeightSubtotal = format(
                objWeightSubtotal.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP), ",.1f"
            )
            for pszKind, arrHeader, arrFields, pszSubtotal, arrOutputPaths in [
                ("羽数", arrCountHeader, arrCountFields, pszCountSubtotal, arrCountOutputPaths),
                ("重量", arrWeightHeader, arrWeightFields, pszWeightSubtotal, arrWeightOutputPaths),
            ]:
                pszOutputPath = create_step0005_discard_subtotal_output_path(
                    dicPair[pszKind]
                )
                write_step0005_tsv(pszOutputPath, "\t".join(arrHeader + ["小計"]),
                                   ["\t".join(arrFields + [pszSubtotal])])
                arrCreatedPaths.append(pszOutputPath)
                arrOutputPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCountOutputPaths, arrWeightOutputPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "DISCARD_SUBTOTAL_CLEANUP_FAILED",
                "discard subtotal failed. Original error = " + str(objException)
                + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0004_group_discard_count_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0004_グループ別_廃棄データ別_(契約|直営|フリー)_羽数_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_DISCARD_COUNT_FILE_NAME",
            "invalid group discard count step0004 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return (
        objMatch.group(1), objMatch.group(2), iYear, iMonth, objMatch.group(5)
    )


def read_step0004_group_discard_count_subtotal_input(
    pszInputPath: str,
) -> Tuple[List[str], List[str], Decimal]:
    _, pszGroupName, iFileYear, iFileMonth, _ = (
        parse_step0004_group_discard_count_file_name(pszInputPath)
    )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0004_GROUP_DISCARD_COUNT_READ_FAILED",
            "failed to read group discard count step0004 TSV. Path = "
            + pszInputPath + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_DISCARD_COUNT_ROW_COUNT",
            "group discard count step0004 TSV must contain one header and one "
            "data row. Path = " + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
    ] + STEP0004_TOTAL_COUNT_ITEMS
    if arrHeader != arrExpectedHeader or len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_DISCARD_COUNT_HEADER",
            "unexpected group discard count step0004 header or column count. "
            "Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0004_GROUP_DISCARD_COUNT_YEAR_MONTH_MISMATCH",
            "file name and row year/month do not match. Path = " + pszInputPath,
        )
    dicGroupNumbers = {pszName: pszNumber for pszNumber, pszName in GROUPS}
    if arrFields[2:4] != [dicGroupNumbers[pszGroupName], pszGroupName]:
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_DISCARD_COUNT_GROUP",
            "file name and row group do not match. Path = " + pszInputPath,
        )
    objSubtotal = Decimal("0")
    for iColumnIndex in range(4, len(arrFields)):
        objValue = parse_decimal_value(
            arrFields[iColumnIndex], 2, iColumnIndex, arrHeader[iColumnIndex]
        )
        if objValue is None:
            raise_conversion_error(
                "INVALID_GROUP_DISCARD_COUNT_SUBTOTAL_VALUE",
                "group discard count subtotal source value was blank. Path = "
                + pszInputPath,
            )
        objSubtotal += objValue
    return arrHeader, arrFields, objSubtotal


def create_step0005_group_discard_count_output_path(
    pszInputPath: str,
) -> str:
    pszBaseName, pszGroupName, iYear, iMonth, pszSuffix = (
        parse_step0004_group_discard_count_file_name(pszInputPath)
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszStem = (
        f"{pszBaseName}_step0005_グループ別_廃棄データ別_"
        f"{pszGroupName}_羽数_{iYear:04d}_{iMonth:02d}{pszSuffix}"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0005_group_discard_count_files(
    arrGroupDiscardCountPaths: List[str],
) -> List[str]:
    """グループ別の廃棄羽数15項目を加算し、小計列を追加します。"""
    dicInputs: Dict[
        Tuple[str, str, int, int, str], Dict[str, str]
    ] = {}
    for pszInputPath in arrGroupDiscardCountPaths:
        pszBaseName, pszGroupName, iYear, iMonth, pszSuffix = (
            parse_step0004_group_discard_count_file_name(pszInputPath)
        )
        objKey = (
            os.path.dirname(os.path.abspath(pszInputPath)),
            pszBaseName, iYear, iMonth, pszSuffix,
        )
        dicGroupPaths = dicInputs.setdefault(objKey, {})
        if pszGroupName in dicGroupPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0004_GROUP_DISCARD_COUNT_INPUT",
                "duplicate group discard count input. Path = " + pszInputPath,
            )
        dicGroupPaths[pszGroupName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrGroupOrder = ["直営", "契約", "フリー"]
    try:
        for objKey in sorted(dicInputs, key=lambda obj: (
            obj[2], obj[3], obj[0], obj[1], obj[4]
        )):
            dicGroupPaths = dicInputs[objKey]
            arrMissingGroups = [
                pszGroupName for pszGroupName in arrGroupOrder
                if pszGroupName not in dicGroupPaths
            ]
            if arrMissingGroups:
                raise_conversion_error(
                    "MISSING_STEP0004_GROUP_DISCARD_COUNT_INPUT",
                    "missing group discard count inputs for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissingGroups),
                )
            for pszGroupName in arrGroupOrder:
                pszInputPath = dicGroupPaths[pszGroupName]
                arrHeader, arrFields, objSubtotal = (
                    read_step0004_group_discard_count_subtotal_input(pszInputPath)
                )
                pszSubtotal = format(
                    Decimal(format_sum_value(objSubtotal)), ",f"
                )
                pszOutputPath = create_step0005_group_discard_count_output_path(
                    pszInputPath
                )
                write_step0005_tsv(
                    pszOutputPath,
                    "\t".join(arrHeader + ["小計"]),
                    ["\t".join(arrFields + [pszSubtotal])],
                )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0005_GROUP_DISCARD_COUNT_CLEANUP_FAILED",
                "group discard count subtotal failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0004_group_discard_weight_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0004_グループ別_廃棄データ別_(契約|直営|フリー)_重量_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_DISCARD_WEIGHT_FILE_NAME",
            "invalid group discard weight step0004 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return (
        objMatch.group(1), objMatch.group(2), iYear, iMonth, objMatch.group(5)
    )


def read_step0004_group_discard_weight_subtotal_input(
    pszInputPath: str,
) -> Tuple[List[str], List[str], Decimal]:
    _, pszGroupName, iFileYear, iFileMonth, _ = (
        parse_step0004_group_discard_weight_file_name(pszInputPath)
    )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0004_GROUP_DISCARD_WEIGHT_READ_FAILED",
            "failed to read group discard weight step0004 TSV. Path = "
            + pszInputPath + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_DISCARD_WEIGHT_ROW_COUNT",
            "group discard weight step0004 TSV must contain one header and one "
            "data row. Path = " + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
    ] + STEP0004_TOTAL_WEIGHT_ITEMS
    if arrHeader != arrExpectedHeader or len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_DISCARD_WEIGHT_HEADER",
            "unexpected group discard weight step0004 header or column count. "
            "Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0004_GROUP_DISCARD_WEIGHT_YEAR_MONTH_MISMATCH",
            "file name and row year/month do not match. Path = " + pszInputPath,
        )
    dicGroupNumbers = {pszName: pszNumber for pszNumber, pszName in GROUPS}
    if arrFields[2:4] != [dicGroupNumbers[pszGroupName], pszGroupName]:
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_DISCARD_WEIGHT_GROUP",
            "file name and row group do not match. Path = " + pszInputPath,
        )
    objSubtotal = Decimal("0")
    for iColumnIndex in range(4, len(arrFields)):
        objValue = parse_decimal_value(
            arrFields[iColumnIndex], 2, iColumnIndex, arrHeader[iColumnIndex]
        )
        if objValue is None:
            raise_conversion_error(
                "INVALID_GROUP_DISCARD_WEIGHT_SUBTOTAL_VALUE",
                "group discard weight subtotal source value was blank. Path = "
                + pszInputPath,
            )
        objSubtotal += objValue
    return arrHeader, arrFields, objSubtotal


def create_step0005_group_discard_weight_output_path(
    pszInputPath: str,
) -> str:
    pszBaseName, pszGroupName, iYear, iMonth, pszSuffix = (
        parse_step0004_group_discard_weight_file_name(pszInputPath)
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszStem = (
        f"{pszBaseName}_step0005_グループ別_廃棄データ別_"
        f"{pszGroupName}_重量_{iYear:04d}_{iMonth:02d}{pszSuffix}"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0005_group_discard_weight_files(
    arrGroupDiscardWeightPaths: List[str],
) -> List[str]:
    """グループ別の廃棄重量15項目を加算し、小計列を追加します。"""
    dicInputs: Dict[
        Tuple[str, str, int, int, str], Dict[str, str]
    ] = {}
    for pszInputPath in arrGroupDiscardWeightPaths:
        pszBaseName, pszGroupName, iYear, iMonth, pszSuffix = (
            parse_step0004_group_discard_weight_file_name(pszInputPath)
        )
        objKey = (
            os.path.dirname(os.path.abspath(pszInputPath)),
            pszBaseName, iYear, iMonth, pszSuffix,
        )
        dicGroupPaths = dicInputs.setdefault(objKey, {})
        if pszGroupName in dicGroupPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0004_GROUP_DISCARD_WEIGHT_INPUT",
                "duplicate group discard weight input. Path = " + pszInputPath,
            )
        dicGroupPaths[pszGroupName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrGroupOrder = ["直営", "契約", "フリー"]
    try:
        for objKey in sorted(dicInputs, key=lambda obj: (
            obj[2], obj[3], obj[0], obj[1], obj[4]
        )):
            dicGroupPaths = dicInputs[objKey]
            arrMissingGroups = [
                pszGroupName for pszGroupName in arrGroupOrder
                if pszGroupName not in dicGroupPaths
            ]
            if arrMissingGroups:
                raise_conversion_error(
                    "MISSING_STEP0004_GROUP_DISCARD_WEIGHT_INPUT",
                    "missing group discard weight inputs for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissingGroups),
                )
            for pszGroupName in arrGroupOrder:
                pszInputPath = dicGroupPaths[pszGroupName]
                arrHeader, arrFields, objSubtotal = (
                    read_step0004_group_discard_weight_subtotal_input(
                        pszInputPath
                    )
                )
                pszSubtotal = format(
                    objSubtotal.quantize(
                        Decimal("0.1"), rounding=ROUND_HALF_UP
                    ), ",.1f"
                )
                pszOutputPath = create_step0005_group_discard_weight_output_path(
                    pszInputPath
                )
                write_step0005_tsv(
                    pszOutputPath,
                    "\t".join(arrHeader + ["小計"]),
                    ["\t".join(arrFields + [pszSubtotal])],
                )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0005_GROUP_DISCARD_WEIGHT_CLEANUP_FAILED",
                "group discard weight subtotal failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0005_group_discard_weight_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0005_グループ別_廃棄データ別_(契約|直営|フリー)_重量_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_WEIGHT_FILE_NAME",
            "invalid group discard weight step0005 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return (
        objMatch.group(1), objMatch.group(2), iYear, iMonth, objMatch.group(5)
    )


def read_step0005_group_discard_weight_input(
    pszInputPath: str,
) -> Tuple[List[str], List[str]]:
    _, pszGroupName, iFileYear, iFileMonth, _ = (
        parse_step0005_group_discard_weight_file_name(pszInputPath)
    )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0005_GROUP_DISCARD_WEIGHT_READ_FAILED",
            "failed to read group discard weight step0005 TSV. Path = "
            + pszInputPath + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_WEIGHT_ROW_COUNT",
            "group discard weight step0005 TSV must contain one header and one "
            "data row. Path = " + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
    ] + STEP0004_TOTAL_WEIGHT_ITEMS + ["小計"]
    if arrHeader != arrExpectedHeader or len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_WEIGHT_HEADER",
            "unexpected group discard weight step0005 header or column count. "
            "Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0005_GROUP_DISCARD_WEIGHT_YEAR_MONTH_MISMATCH",
            "file name and row year/month do not match. Path = " + pszInputPath,
        )
    dicGroupNumbers = {pszName: pszNumber for pszNumber, pszName in GROUPS}
    if arrFields[2:4] != [dicGroupNumbers[pszGroupName], pszGroupName]:
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_WEIGHT_GROUP",
            "file name and row group do not match. Path = " + pszInputPath,
        )
    for iColumnIndex in range(4, len(arrFields) - 1):
        objValue = parse_decimal_value(
            arrFields[iColumnIndex], 2, iColumnIndex, arrHeader[iColumnIndex]
        )
        if objValue is None:
            raise_conversion_error(
                "INVALID_STEP0005_GROUP_DISCARD_WEIGHT_VALUE",
                "group discard weight step0005 value was blank. Path = "
                + pszInputPath,
            )
    pszSubtotalValue = arrFields[-1].replace(",", "")
    if parse_decimal_value(
        pszSubtotalValue, 2, len(arrFields) - 1, arrHeader[-1]
    ) is None:
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_WEIGHT_SUBTOTAL",
            "group discard weight subtotal was blank. Path = " + pszInputPath,
        )
    return arrHeader, arrFields


def create_step0006_group_discard_weight_output_path(
    pszInputPath: str,
) -> str:
    pszBaseName, _, iYear, iMonth, pszSuffix = (
        parse_step0005_group_discard_weight_file_name(pszInputPath)
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszStem = (
        f"{pszBaseName}_step0006_グループ別_廃棄データ別_重量_"
        f"{iYear:04d}_{iMonth:02d}{pszSuffix}"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0006_group_discard_weight_files(
    arrGroupDiscardWeightPaths: List[str],
) -> List[str]:
    """契約・直営・フリーのstep0005重量TSVを縦結合します。"""
    dicInputs: Dict[
        Tuple[str, str, int, int, str], Dict[str, str]
    ] = {}
    for pszInputPath in arrGroupDiscardWeightPaths:
        pszBaseName, pszGroupName, iYear, iMonth, pszSuffix = (
            parse_step0005_group_discard_weight_file_name(pszInputPath)
        )
        objKey = (
            os.path.dirname(os.path.abspath(pszInputPath)),
            pszBaseName, iYear, iMonth, pszSuffix,
        )
        dicGroupPaths = dicInputs.setdefault(objKey, {})
        if pszGroupName in dicGroupPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0005_GROUP_DISCARD_WEIGHT_INPUT",
                "duplicate group discard weight input. Path = " + pszInputPath,
            )
        dicGroupPaths[pszGroupName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrGroupOrder = ["契約", "直営", "フリー"]
    try:
        for objKey in sorted(dicInputs, key=lambda obj: (
            obj[2], obj[3], obj[0], obj[1], obj[4]
        )):
            dicGroupPaths = dicInputs[objKey]
            arrMissingGroups = [
                pszGroupName for pszGroupName in arrGroupOrder
                if pszGroupName not in dicGroupPaths
            ]
            if arrMissingGroups:
                raise_conversion_error(
                    "MISSING_STEP0005_GROUP_DISCARD_WEIGHT_INPUT",
                    "missing group discard weight inputs for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissingGroups),
                )
            arrOutputRows: List[List[str]] = []
            arrOutputHeader: List[str] = []
            for iGroupIndex, pszGroupName in enumerate(arrGroupOrder):
                arrHeader, arrFields = read_step0005_group_discard_weight_input(
                    dicGroupPaths[pszGroupName]
                )
                if iGroupIndex == 0:
                    arrOutputHeader = arrHeader
                elif arrHeader != arrOutputHeader:
                    raise_conversion_error(
                        "STEP0005_GROUP_DISCARD_WEIGHT_HEADER_MISMATCH",
                        "group discard weight headers did not match. Key = "
                        + str(objKey) + ", Group = " + pszGroupName,
                    )
                arrOutputRows.append(arrFields)
            pszOutputPath = create_step0006_group_discard_weight_output_path(
                dicGroupPaths[arrGroupOrder[0]]
            )
            write_step0006_tsv(
                pszOutputPath,
                [arrOutputHeader] + arrOutputRows,
            )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0006_GROUP_DISCARD_WEIGHT_CLEANUP_FAILED",
                "group discard weight merge failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0005_group_discard_count_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0005_グループ別_廃棄データ別_(契約|直営|フリー)_羽数_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_COUNT_FILE_NAME",
            "invalid group discard count step0005 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return (
        objMatch.group(1), objMatch.group(2), iYear, iMonth, objMatch.group(5)
    )


def read_step0005_group_discard_count_input(
    pszInputPath: str,
) -> Tuple[List[str], List[str]]:
    _, pszGroupName, iFileYear, iFileMonth, _ = (
        parse_step0005_group_discard_count_file_name(pszInputPath)
    )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0005_GROUP_DISCARD_COUNT_READ_FAILED",
            "failed to read group discard count step0005 TSV. Path = "
            + pszInputPath + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_COUNT_ROW_COUNT",
            "group discard count step0005 TSV must contain one header and one "
            "data row. Path = " + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
    ] + STEP0004_TOTAL_COUNT_ITEMS + ["小計"]
    if arrHeader != arrExpectedHeader or len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_COUNT_HEADER",
            "unexpected group discard count step0005 header or column count. "
            "Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0005_GROUP_DISCARD_COUNT_YEAR_MONTH_MISMATCH",
            "file name and row year/month do not match. Path = " + pszInputPath,
        )
    dicGroupNumbers = {pszName: pszNumber for pszNumber, pszName in GROUPS}
    if arrFields[2:4] != [dicGroupNumbers[pszGroupName], pszGroupName]:
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_COUNT_GROUP",
            "file name and row group do not match. Path = " + pszInputPath,
        )
    for iColumnIndex in range(4, len(arrFields) - 1):
        objValue = parse_decimal_value(
            arrFields[iColumnIndex], 2, iColumnIndex, arrHeader[iColumnIndex]
        )
        if objValue is None:
            raise_conversion_error(
                "INVALID_STEP0005_GROUP_DISCARD_COUNT_VALUE",
                "group discard count step0005 value was blank. Path = "
                + pszInputPath,
            )
    pszSubtotalValue = arrFields[-1].replace(",", "")
    if parse_decimal_value(
        pszSubtotalValue, 2, len(arrFields) - 1, arrHeader[-1]
    ) is None:
        raise_conversion_error(
            "INVALID_STEP0005_GROUP_DISCARD_COUNT_SUBTOTAL",
            "group discard count subtotal was blank. Path = " + pszInputPath,
        )
    return arrHeader, arrFields


def create_step0006_group_discard_count_output_path(
    pszInputPath: str,
) -> str:
    pszBaseName, _, iYear, iMonth, pszSuffix = (
        parse_step0005_group_discard_count_file_name(pszInputPath)
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszStem = (
        f"{pszBaseName}_step0006_グループ別_廃棄データ別_羽数_"
        f"{iYear:04d}_{iMonth:02d}{pszSuffix}"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0006_group_discard_count_files(
    arrGroupDiscardCountPaths: List[str],
) -> List[str]:
    """契約・直営・フリーのstep0005羽数TSVを縦結合します。"""
    dicInputs: Dict[
        Tuple[str, str, int, int, str], Dict[str, str]
    ] = {}
    for pszInputPath in arrGroupDiscardCountPaths:
        pszBaseName, pszGroupName, iYear, iMonth, pszSuffix = (
            parse_step0005_group_discard_count_file_name(pszInputPath)
        )
        objKey = (
            os.path.dirname(os.path.abspath(pszInputPath)),
            pszBaseName, iYear, iMonth, pszSuffix,
        )
        dicGroupPaths = dicInputs.setdefault(objKey, {})
        if pszGroupName in dicGroupPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0005_GROUP_DISCARD_COUNT_INPUT",
                "duplicate group discard count input. Path = " + pszInputPath,
            )
        dicGroupPaths[pszGroupName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrGroupOrder = ["契約", "直営", "フリー"]
    try:
        for objKey in sorted(dicInputs, key=lambda obj: (
            obj[2], obj[3], obj[0], obj[1], obj[4]
        )):
            dicGroupPaths = dicInputs[objKey]
            arrMissingGroups = [
                pszGroupName for pszGroupName in arrGroupOrder
                if pszGroupName not in dicGroupPaths
            ]
            if arrMissingGroups:
                raise_conversion_error(
                    "MISSING_STEP0005_GROUP_DISCARD_COUNT_INPUT",
                    "missing group discard count inputs for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissingGroups),
                )
            arrOutputRows: List[List[str]] = []
            arrOutputHeader: List[str] = []
            for iGroupIndex, pszGroupName in enumerate(arrGroupOrder):
                arrHeader, arrFields = read_step0005_group_discard_count_input(
                    dicGroupPaths[pszGroupName]
                )
                if iGroupIndex == 0:
                    arrOutputHeader = arrHeader
                elif arrHeader != arrOutputHeader:
                    raise_conversion_error(
                        "STEP0005_GROUP_DISCARD_COUNT_HEADER_MISMATCH",
                        "group discard count headers did not match. Key = "
                        + str(objKey) + ", Group = " + pszGroupName,
                    )
                arrOutputRows.append(arrFields)
            pszOutputPath = create_step0006_group_discard_count_output_path(
                dicGroupPaths[arrGroupOrder[0]]
            )
            write_step0006_tsv(
                pszOutputPath,
                [arrOutputHeader] + arrOutputRows,
            )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0006_GROUP_DISCARD_COUNT_CLEANUP_FAILED",
                "group discard count merge failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0006_group_discard_count_file_name(
    pszInputPath: str,
) -> Tuple[str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0006_グループ別_廃棄データ別_羽数_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0006_GROUP_DISCARD_COUNT_FILE_NAME",
            "invalid group discard count step0006 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
    return objMatch.group(1), iYear, iMonth, objMatch.group(4)


def create_step0007_group_total_discard_count_output_path(
    pszInputPath: str,
) -> str:
    pszBaseName, iYear, iMonth, pszSuffix = (
        parse_step0006_group_discard_count_file_name(pszInputPath)
    )
    pszStem = (
        f"{pszBaseName}_step0007_グループ別_全廃棄羽数_"
        f"{iYear:04d}_{iMonth:02d}{pszSuffix}"
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0007_group_total_discard_count_files(
    arrStep0006GroupDiscardCountPaths: List[str],
) -> List[str]:
    """step0006グループ別廃棄羽数から全廃棄羽数だけを出力します。"""
    arrCreatedPaths: List[str] = []
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
    ] + STEP0004_TOTAL_COUNT_ITEMS + ["小計"]
    dicExpectedGroupNumbers = {
        pszName: pszNumber for pszNumber, pszName in GROUPS
    }
    try:
        for pszInputPath in sorted(arrStep0006GroupDiscardCountPaths):
            _, iFileYear, iFileMonth, _ = (
                parse_step0006_group_discard_count_file_name(pszInputPath)
            )
            try:
                with open(
                    pszInputPath, "r", encoding="utf-8", newline=""
                ) as objFile:
                    arrLines = [
                        strip_line_ending(pszLine) for pszLine in objFile
                    ]
            except Exception as objException:
                raise_conversion_error(
                    "STEP0006_GROUP_DISCARD_COUNT_READ_FAILED",
                    "failed to read group discard count step0006 TSV. Path = "
                    + pszInputPath + "; Detail = " + str(objException),
                )
            if len(arrLines) < 2:
                raise_conversion_error(
                    "INVALID_STEP0006_GROUP_DISCARD_COUNT_ROW_COUNT",
                    "group discard count step0006 TSV must contain a header "
                    "and data rows. Path = " + pszInputPath,
                )
            arrHeader = arrLines[0].split("\t")
            if arrHeader != arrExpectedHeader:
                raise_conversion_error(
                    "INVALID_STEP0006_GROUP_DISCARD_COUNT_HEADER",
                    "unexpected group discard count step0006 header. Path = "
                    + pszInputPath,
                )
            arrOutputRows: List[List[str]] = [
                arrHeader[:4] + ["全廃棄羽数"]
            ]
            dicSeenGroups: Dict[str, bool] = {}
            for iRowNumber, pszLine in enumerate(arrLines[1:], start=2):
                arrFields = pszLine.split("\t")
                if len(arrFields) != len(arrExpectedHeader):
                    raise_conversion_error(
                        "INVALID_STEP0006_GROUP_DISCARD_COUNT_COLUMNS",
                        "unexpected group discard count step0006 column count. "
                        "Path = " + pszInputPath,
                    )
                if parse_year_month(
                    arrFields[0], arrFields[1], iRowNumber
                ) != (iFileYear, iFileMonth):
                    raise_conversion_error(
                        "STEP0006_GROUP_DISCARD_COUNT_YEAR_MONTH_MISMATCH",
                        "file name and row year/month do not match. Path = "
                        + pszInputPath,
                    )
                pszGroupNumber, pszGroupName = arrFields[2:4]
                if (
                    pszGroupName not in dicExpectedGroupNumbers
                    or dicExpectedGroupNumbers[pszGroupName] != pszGroupNumber
                    or pszGroupName in dicSeenGroups
                ):
                    raise_conversion_error(
                        "INVALID_STEP0006_GROUP_DISCARD_COUNT_GROUP",
                        "invalid or duplicate group in step0006 TSV. Path = "
                        + pszInputPath,
                    )
                dicSeenGroups[pszGroupName] = True
                if parse_decimal_value(
                    arrFields[-1].replace(",", ""),
                    iRowNumber,
                    len(arrFields) - 1,
                    "小計",
                ) is None:
                    raise_conversion_error(
                        "INVALID_STEP0006_GROUP_DISCARD_COUNT_SUBTOTAL",
                        "group discard count subtotal was blank. Path = "
                        + pszInputPath,
                    )
                arrOutputRows.append(arrFields[:4] + [arrFields[-1]])
            if set(dicSeenGroups) != set(dicExpectedGroupNumbers):
                raise_conversion_error(
                    "MISSING_STEP0006_GROUP_DISCARD_COUNT_GROUP",
                    "step0006 TSV must contain contract, direct, and free "
                    "groups. Path = " + pszInputPath,
                )
            pszOutputPath = (
                create_step0007_group_total_discard_count_output_path(
                    pszInputPath
                )
            )
            write_step0007_discard_summary_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0007_GROUP_TOTAL_DISCARD_COUNT_CLEANUP_FAILED",
                "group total discard count creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0006_group_discard_weight_file_name(
    pszInputPath: str,
) -> Tuple[str, int, int, str]:
    objMatch = re.fullmatch(
        r"(.+)_step0006_グループ別_廃棄データ別_重量_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0006_GROUP_DISCARD_WEIGHT_FILE_NAME",
            "invalid group discard weight step0006 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
    return objMatch.group(1), iYear, iMonth, objMatch.group(4)


def create_step0007_group_total_discard_weight_output_path(
    pszInputPath: str,
) -> str:
    pszBaseName, iYear, iMonth, pszSuffix = (
        parse_step0006_group_discard_weight_file_name(pszInputPath)
    )
    pszStem = (
        f"{pszBaseName}_step0007_グループ別_全廃棄重量_"
        f"{iYear:04d}_{iMonth:02d}{pszSuffix}"
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0007_group_total_discard_weight_files(
    arrStep0006GroupDiscardWeightPaths: List[str],
) -> List[str]:
    """step0006グループ別廃棄重量から全廃棄重量だけを出力します。"""
    arrCreatedPaths: List[str] = []
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名",
    ] + STEP0004_TOTAL_WEIGHT_ITEMS + ["小計"]
    dicExpectedGroupNumbers = {
        pszName: pszNumber for pszNumber, pszName in GROUPS
    }
    try:
        for pszInputPath in sorted(arrStep0006GroupDiscardWeightPaths):
            _, iFileYear, iFileMonth, _ = (
                parse_step0006_group_discard_weight_file_name(pszInputPath)
            )
            try:
                with open(
                    pszInputPath, "r", encoding="utf-8", newline=""
                ) as objFile:
                    arrLines = [
                        strip_line_ending(pszLine) for pszLine in objFile
                    ]
            except Exception as objException:
                raise_conversion_error(
                    "STEP0006_GROUP_DISCARD_WEIGHT_READ_FAILED",
                    "failed to read group discard weight step0006 TSV. Path = "
                    + pszInputPath + "; Detail = " + str(objException),
                )
            if len(arrLines) < 2:
                raise_conversion_error(
                    "INVALID_STEP0006_GROUP_DISCARD_WEIGHT_ROW_COUNT",
                    "group discard weight step0006 TSV must contain a header "
                    "and data rows. Path = " + pszInputPath,
                )
            arrHeader = arrLines[0].split("\t")
            if arrHeader != arrExpectedHeader:
                raise_conversion_error(
                    "INVALID_STEP0006_GROUP_DISCARD_WEIGHT_HEADER",
                    "unexpected group discard weight step0006 header. Path = "
                    + pszInputPath,
                )
            arrOutputRows: List[List[str]] = [
                arrHeader[:4] + ["全廃棄重量"]
            ]
            dicSeenGroups: Dict[str, bool] = {}
            for iRowNumber, pszLine in enumerate(arrLines[1:], start=2):
                arrFields = pszLine.split("\t")
                if len(arrFields) != len(arrExpectedHeader):
                    raise_conversion_error(
                        "INVALID_STEP0006_GROUP_DISCARD_WEIGHT_COLUMNS",
                        "unexpected group discard weight step0006 column count. "
                        "Path = " + pszInputPath,
                    )
                if parse_year_month(
                    arrFields[0], arrFields[1], iRowNumber
                ) != (iFileYear, iFileMonth):
                    raise_conversion_error(
                        "STEP0006_GROUP_DISCARD_WEIGHT_YEAR_MONTH_MISMATCH",
                        "file name and row year/month do not match. Path = "
                        + pszInputPath,
                    )
                pszGroupNumber, pszGroupName = arrFields[2:4]
                if (
                    pszGroupName not in dicExpectedGroupNumbers
                    or dicExpectedGroupNumbers[pszGroupName] != pszGroupNumber
                    or pszGroupName in dicSeenGroups
                ):
                    raise_conversion_error(
                        "INVALID_STEP0006_GROUP_DISCARD_WEIGHT_GROUP",
                        "invalid or duplicate group in step0006 TSV. Path = "
                        + pszInputPath,
                    )
                dicSeenGroups[pszGroupName] = True
                objSubtotal = parse_decimal_value(
                    arrFields[-1].replace(",", ""),
                    iRowNumber,
                    len(arrFields) - 1,
                    "小計",
                )
                if objSubtotal is None:
                    raise_conversion_error(
                        "INVALID_STEP0006_GROUP_DISCARD_WEIGHT_SUBTOTAL",
                        "group discard weight subtotal was blank. Path = "
                        + pszInputPath,
                    )
                pszFormattedSubtotal = format(
                    objSubtotal.quantize(
                        Decimal("0.1"), rounding=ROUND_HALF_UP
                    ),
                    ",.1f",
                )
                arrOutputRows.append(
                    [arrFields[0], str(iFileMonth), *arrFields[2:4],
                     pszFormattedSubtotal]
                )
            if set(dicSeenGroups) != set(dicExpectedGroupNumbers):
                raise_conversion_error(
                    "MISSING_STEP0006_GROUP_DISCARD_WEIGHT_GROUP",
                    "step0006 TSV must contain contract, direct, and free "
                    "groups. Path = " + pszInputPath,
                )
            pszOutputPath = (
                create_step0007_group_total_discard_weight_output_path(
                    pszInputPath
                )
            )
            write_step0007_discard_summary_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0007_GROUP_TOTAL_DISCARD_WEIGHT_CLEANUP_FAILED",
                "group total discard weight creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0004_group_month_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int]:
    objMatch = re.fullmatch(
        r"(.+)_step0004_グループ別_(契約|直営|フリー)_"
        r"(\d{4})_(\d{2})(?:_\d+)*\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0004_GROUP_MONTH_FILE_NAME",
            "invalid monthly group step0004 file name. Path = " + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return objMatch.group(1), objMatch.group(2), iYear, iMonth


def parse_step0007_group_total_file_name(
    pszInputPath: str, pszKind: str
) -> Tuple[str, int, int]:
    if pszKind not in {"羽数", "重量"}:
        raise_conversion_error(
            "INVALID_STEP0007_GROUP_TOTAL_KIND",
            "group total kind must be 羽数 or 重量.",
        )
    objMatch = re.fullmatch(
        rf"(.+)_step0007_グループ別_全廃棄{pszKind}_"
        r"(\d{4})_(\d{2})(?:_\d+)*\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0007_GROUP_TOTAL_FILE_NAME",
            "invalid group total discard step0007 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
    return objMatch.group(1), iYear, iMonth


def read_step0007_group_total_rows(
    pszInputPath: str, pszKind: str
) -> Tuple[Tuple[str, int, int], Dict[str, List[str]]]:
    pszBaseName, iFileYear, iFileMonth = parse_step0007_group_total_file_name(
        pszInputPath, pszKind
    )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0007_GROUP_TOTAL_READ_FAILED",
            "failed to read group total discard step0007 TSV. Path = "
            + pszInputPath + "; Detail = " + str(objException),
        )
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "全廃棄" + pszKind,
    ]
    if len(arrLines) != len(GROUPS) + 1:
        raise_conversion_error(
            "INVALID_STEP0007_GROUP_TOTAL_ROW_COUNT",
            "group total discard step0007 TSV must contain one header and "
            "one row per group. Path = " + pszInputPath,
        )
    if arrLines[0].split("\t") != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0007_GROUP_TOTAL_HEADER",
            "unexpected group total discard step0007 header. Path = "
            + pszInputPath,
        )
    dicExpectedGroupNumbers = {
        pszName: pszNumber for pszNumber, pszName in GROUPS
    }
    dicGroupRows: Dict[str, List[str]] = {}
    for iRowNumber, pszLine in enumerate(arrLines[1:], start=2):
        arrFields = pszLine.split("\t")
        if len(arrFields) != len(arrExpectedHeader):
            raise_conversion_error(
                "INVALID_STEP0007_GROUP_TOTAL_COLUMN_COUNT",
                "unexpected group total discard step0007 column count. Path = "
                + pszInputPath,
            )
        if parse_year_month(arrFields[0], arrFields[1], iRowNumber) != (
            iFileYear, iFileMonth
        ):
            raise_conversion_error(
                "STEP0007_GROUP_TOTAL_YEAR_MONTH_MISMATCH",
                "file name and row year/month do not match. Path = "
                + pszInputPath,
            )
        pszGroupNumber, pszGroupName = arrFields[2:4]
        if (
            pszGroupName not in dicExpectedGroupNumbers
            or dicExpectedGroupNumbers[pszGroupName] != pszGroupNumber
            or pszGroupName in dicGroupRows
        ):
            raise_conversion_error(
                "INVALID_STEP0007_GROUP_TOTAL_GROUP",
                "invalid or duplicate group in step0007 TSV. Path = "
                + pszInputPath,
            )
        if parse_decimal_value(
            arrFields[4].replace(",", ""),
            iRowNumber,
            4,
            arrExpectedHeader[4],
        ) is None:
            raise_conversion_error(
                "INVALID_STEP0007_GROUP_TOTAL_VALUE",
                "group total discard value was blank. Path = " + pszInputPath,
            )
        dicGroupRows[pszGroupName] = arrFields
    if set(dicGroupRows) != set(dicExpectedGroupNumbers):
        raise_conversion_error(
            "MISSING_STEP0007_GROUP_TOTAL_GROUP",
            "step0007 TSV must contain contract, direct, and free groups. "
            "Path = " + pszInputPath,
        )
    return (pszBaseName, iFileYear, iFileMonth), dicGroupRows


def create_step0008_group_discard_deduction_output_path(
    pszStep0004Path: str, pszGroupName: str, iYear: int, iMonth: int
) -> str:
    pszBaseName, _, _, _ = parse_step0004_group_month_file_name(pszStep0004Path)
    objMatch = re.fullmatch(
        r".+_step0004_グループ別_(?:契約|直営|フリー)_"
        r"\d{4}_\d{2}((?:_\d+)*)\.tsv",
        os.path.basename(pszStep0004Path),
    )
    pszInputSuffix = objMatch.group(1) if objMatch is not None else ""
    pszStem = (
        f"{pszBaseName}_step0008_グループ別_控除廃棄_{pszGroupName}_"
        f"{iYear:04d}_{iMonth:02d}{pszInputSuffix}"
    )
    pszDirectory = os.path.dirname(pszStep0004Path)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def read_group_pre_shackling_counts(
    pszPath: str, pszItemName: str, iYear: int, iMonth: int
) -> Dict[str, str]:
    """グループ別step0003から懸鳥前羽数・重量を検証して読み込みます。"""
    with open(pszPath, "r", encoding="utf-8-sig", newline="") as objFile:
        arrRows = [strip_line_ending(line).split("\t") for line in objFile]
    if not arrRows or arrRows[0] != [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", pszItemName
    ]:
        raise_conversion_error(
            "INVALID_PRE_SHACKLING_HEADER",
            "invalid header. Item = " + pszItemName + "; Path = " + pszPath,
        )
    dicNumbers = {name: number for number, name in GROUPS}
    dicValues: Dict[str, str] = {}
    for iLine, arrFields in enumerate(arrRows[1:], start=2):
        pszContext = (
            f"Year/month = {iYear:04d}/{iMonth:02d}; Item = {pszItemName}; "
            f"Line = {iLine}; Path = {pszPath}"
        )
        if len(arrFields) != 5:
            raise_conversion_error("INVALID_PRE_SHACKLING_COLUMNS", pszContext)
        pszGroup = arrFields[3].strip(" \u3000\t")
        pszContext += "; Group = " + pszGroup
        if (parse_year_month(arrFields[0], arrFields[1], iLine) != (iYear, iMonth)
                or pszGroup not in dicNumbers
                or arrFields[2].strip() != dicNumbers[pszGroup]):
            raise_conversion_error("INVALID_PRE_SHACKLING_KEY", pszContext)
        if pszGroup in dicValues:
            raise_conversion_error("DUPLICATE_PRE_SHACKLING_GROUP", pszContext)
        try:
            objValue = parse_decimal_value(
                arrFields[4].replace(",", ""), iLine, 4, pszItemName
            )
            if objValue is None:
                raise ValueError("value was blank")
        except (ConversionError, ValueError) as objException:
            raise_conversion_error(
                "INVALID_PRE_SHACKLING_VALUE", pszContext + "; " + str(objException)
            )
        dicValues[pszGroup] = arrFields[4]
    for pszGroup in dicNumbers:
        if pszGroup not in dicValues:
            raise_conversion_error(
                "MISSING_PRE_SHACKLING_GROUP",
                f"Year/month = {iYear:04d}/{iMonth:02d}; Group = {pszGroup}; "
                f"Item = {pszItemName}; Path = {pszPath}",
            )
    return dicValues


def read_group_converted_discard_weights(
    pszPath: str, iYear: int, iMonth: int
) -> Dict[str, Decimal]:
    """明細ごとに自主廃棄重量を0.72換算・切り捨てし、グループ別に合計します。"""
    dicNumbers = {name: number for number, name in GROUPS}
    dicTotals = {name: Decimal("0") for _, name in GROUPS}
    setSeenGroups = set()
    arrWeightItems = STEP0004_TOTAL_WEIGHT_ITEMS + ["予備３重量", "予備４重量"]
    pszContext = f"Year/month = {iYear:04d}/{iMonth:02d}; Path = {pszPath}"
    with open(pszPath, "r", encoding="utf-8-sig", newline="") as objFile:
        objReader = csv.reader(objFile, delimiter="\t")
        arrHeader = next(objReader, [])
        if (len(arrHeader) != 64
                or [arrHeader[i] for i in (0, 1, 10, 11)]
                != ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名"]
                or arrHeader[47:64] != arrWeightItems):
            raise_conversion_error("INVALID_CONVERTED_DISCARD_HEADER", pszContext)
        for iLine, arrFields in enumerate(objReader, start=2):
            pszRowContext = pszContext + f"; Line = {iLine}"
            if len(arrFields) != 64:
                raise_conversion_error("INVALID_CONVERTED_DISCARD_COLUMNS", pszRowContext)
            pszGroup = arrFields[11].strip(" \u3000\t")
            pszRowContext += "; Group = " + pszGroup
            try:
                if (parse_year_month(arrFields[0], arrFields[1], iLine) != (iYear, iMonth)
                        or pszGroup not in dicNumbers
                        or arrFields[10].strip() != dicNumbers[pszGroup]):
                    raise ValueError("year/month or group did not match")
                objSubtotal = Decimal("0")
                for iColumn, pszItem in enumerate(arrWeightItems, start=47):
                    objValue = parse_decimal_value(
                        arrFields[iColumn].replace(",", ""), iLine, iColumn, pszItem
                    )
                    if objValue is None:
                        raise ValueError("blank weight. Item = " + pszItem)
                    objSubtotal += objValue
            except (ConversionError, ValueError) as objException:
                raise_conversion_error(
                    "INVALID_CONVERTED_DISCARD_ROW",
                    pszRowContext + "; Detail = " + str(objException),
                )
            dicTotals[pszGroup] += (objSubtotal / Decimal("0.72")).quantize(
                Decimal("0.1"), rounding=ROUND_DOWN
            )
            setSeenGroups.add(pszGroup)
    for pszGroup in dicNumbers:
        if pszGroup not in setSeenGroups:
            raise_conversion_error(
                "MISSING_CONVERTED_DISCARD_GROUP", pszContext + "; Group = " + pszGroup
            )
    return dicTotals


def create_step0008_group_discard_deduction_files(
    arrStep0004GroupPaths: List[str],
    arrStep0007GroupTotalDiscardCountPaths: List[str],
    arrStep0007GroupTotalDiscardWeightPaths: List[str],
    arrStep0003GroupPaths: List[str],
    arrStep0001Paths: List[str],
) -> List[str]:
    """既存データへ懸鳥前羽数・重量と明細から換算した自主廃棄重量を追加します。"""
    arrPreItems = [
        "懸鳥前死鶏羽数", "懸鳥前発育不良羽数",
        "懸鳥前死鶏重量", "懸鳥前発育不良重量",
    ]
    dicPreCounts: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszPath in arrStep0003GroupPaths:
        objMatch = re.fullmatch(
            r"(.+)_step0003_グループ別_(懸鳥前死鶏羽数|懸鳥前発育不良羽数|"
            r"懸鳥前死鶏重量|懸鳥前発育不良重量)_"
            r"(\d{4})_(\d{2})(?:_\d+)*\.tsv", os.path.basename(pszPath)
        )
        if objMatch is None:
            continue
        pszBase, pszItem, pszYear, pszMonth = objMatch.groups()
        iYear, iMonth = parse_year_month(pszYear, pszMonth, 1)
        objPreKey = (
            os.path.dirname(os.path.abspath(pszPath)), pszBase, iYear, iMonth, pszItem
        )
        if objPreKey in dicPreCounts:
            raise_conversion_error(
                "DUPLICATE_PRE_SHACKLING_INPUT",
                "duplicate group pre-shackling input. Key = " + str(objPreKey)
                + "; Path = " + pszPath,
            )
        dicPreCounts[objPreKey] = read_group_pre_shackling_counts(
            pszPath, pszItem, iYear, iMonth
        )
    dicStep0004Groups: Dict[
        Tuple[str, str, int, int], Dict[str, Tuple[str, List[str]]]
    ] = {}
    for pszPath in arrStep0004GroupPaths:
        pszBaseName, pszGroupName, iYear, iMonth = (
            parse_step0004_group_month_file_name(pszPath)
        )
        objKey = (
            os.path.dirname(os.path.abspath(pszPath)),
            pszBaseName, iYear, iMonth,
        )
        dicGroups = dicStep0004Groups.setdefault(objKey, {})
        if pszGroupName in dicGroups:
            raise_conversion_error(
                "DUPLICATE_STEP0004_GROUP_MONTH_INPUT",
                "duplicate step0004 group/month input. Path = " + pszPath,
            )
        dicExpectedGroupNumbers = {
            pszName: pszNumber for pszNumber, pszName in GROUPS
        }
        pszHeader, pszRow = read_step0004_for_step0005(
            pszPath,
            iYear,
            iMonth,
            dicExpectedGroupNumbers[pszGroupName],
            pszGroupName,
        )
        dicGroups[pszGroupName] = (pszHeader, pszRow.split("\t"))

    def build_step0007_group_total_map(
        arrPaths: List[str], pszKind: str
    ) -> Dict[Tuple[str, str, int, int], Dict[str, List[str]]]:
        dicResult: Dict[
            Tuple[str, str, int, int], Dict[str, List[str]]
        ] = {}
        for pszPath in arrPaths:
            (pszBaseName, iYear, iMonth), dicRows = (
                read_step0007_group_total_rows(pszPath, pszKind)
            )
            objKey = (
                os.path.dirname(os.path.abspath(pszPath)),
                pszBaseName, iYear, iMonth,
            )
            if objKey in dicResult:
                raise_conversion_error(
                    "DUPLICATE_STEP0007_GROUP_TOTAL_INPUT",
                    "duplicate step0007 group total input. Kind = "
                    + pszKind + "; Path = " + pszPath,
                )
            dicResult[objKey] = dicRows
        return dicResult

    dicStep0007Counts = build_step0007_group_total_map(
        arrStep0007GroupTotalDiscardCountPaths, "羽数"
    )
    dicStep0007Weights = build_step0007_group_total_map(
        arrStep0007GroupTotalDiscardWeightPaths, "重量"
    )
    if set(dicStep0004Groups) != set(dicStep0007Counts) or set(
        dicStep0004Groups
    ) != set(dicStep0007Weights):
        raise_conversion_error(
            "STEP0008_GROUP_DISCARD_DEDUCTION_INPUT_MISMATCH",
            "step0004, step0007 count, and step0007 weight inputs did not "
            "cover the same base/year/month sets.",
        )

    arrCreatedPaths: List[str] = []
    arrExpectedGroups = [pszName for _, pszName in GROUPS]
    dicStep0001Paths: Dict[Tuple[str, str, int, int], str] = {}
    for pszPath in arrStep0001Paths:
        objMatch = re.fullmatch(
            r"(.+)_step0001_(\d{4})_(\d{2})(?:_\d+)*\.tsv", os.path.basename(pszPath)
        )
        if objMatch is None:
            raise_conversion_error("INVALID_CONVERTED_DISCARD_FILENAME", pszPath)
        pszBase, pszYear, pszMonth = objMatch.groups()
        iYear, iMonth = parse_year_month(pszYear, pszMonth, 1)
        objMonthlyKey = (
            os.path.dirname(os.path.abspath(pszPath)), pszBase, iYear, iMonth
        )
        if objMonthlyKey in dicStep0001Paths:
            raise_conversion_error(
                "DUPLICATE_CONVERTED_DISCARD_INPUT",
                "duplicate step0001 input. Key = " + str(objMonthlyKey) + "; Path = " + pszPath,
            )
        dicStep0001Paths[objMonthlyKey] = pszPath
    dicConvertedWeights: Dict[Tuple[str, str, int, int], Dict[str, Decimal]] = {}
    for objKey in dicStep0004Groups:
        for pszItem in arrPreItems:
            if (*objKey, pszItem) not in dicPreCounts:
                pszExpectedPath = os.path.join(
                    objKey[0], f"{objKey[1]}_step0003_グループ別_{pszItem}_"
                    f"{objKey[2]:04d}_{objKey[3]:02d}.tsv"
                )
                raise_conversion_error(
                    "MISSING_PRE_SHACKLING_INPUT",
                    "missing input for 契約・直営・フリー. Path = " + pszExpectedPath,
                )
        if objKey not in dicStep0001Paths:
            raise_conversion_error(
                "MISSING_CONVERTED_DISCARD_INPUT",
                "missing step0001 input for 契約・直営・フリー. Key = " + str(objKey)
                + "; Path = " + os.path.join(
                    objKey[0], f"{objKey[1]}_step0001_{objKey[2]:04d}_{objKey[3]:02d}.tsv"
                ),
            )
        dicConvertedWeights[objKey] = read_group_converted_discard_weights(
            dicStep0001Paths[objKey], objKey[2], objKey[3]
        )
    try:
        for objKey in sorted(dicStep0004Groups):
            dicGroups = dicStep0004Groups[objKey]
            arrMissingGroups = [
                pszGroupName for pszGroupName in arrExpectedGroups
                if pszGroupName not in dicGroups
            ]
            if arrMissingGroups:
                raise_conversion_error(
                    "MISSING_STEP0004_GROUP_MONTH_INPUT",
                    "missing step0004 groups for month. Key = " + str(objKey)
                    + "; Groups = " + ", ".join(arrMissingGroups),
                )
            for pszGroupName in arrExpectedGroups:
                pszStep0004Path = next(
                    pszPath for pszPath in arrStep0004GroupPaths
                    if os.path.dirname(os.path.abspath(pszPath)) == objKey[0]
                    and parse_step0004_group_month_file_name(pszPath)
                    == (objKey[1], pszGroupName, objKey[2], objKey[3])
                )
                _, arrStep0004Fields = dicGroups[pszGroupName]
                arrCountFields = dicStep0007Counts[objKey][pszGroupName]
                arrWeightFields = dicStep0007Weights[objKey][pszGroupName]
                arrOutputFields = [
                    *arrStep0004Fields[:6],
                    arrCountFields[4],
                    arrStep0004Fields[6],
                    arrWeightFields[4],
                    arrStep0004Fields[7],
                    *[dicPreCounts[(*objKey, pszItem)][pszGroupName]
                      for pszItem in arrPreItems],
                    format(dicConvertedWeights[objKey][pszGroupName], ".1f"),
                ]
                arrOutputRows = [[
                    "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数",
                    "出荷羽数", "全廃棄羽数", "出荷重量", "全廃棄重量",
                    "平均体重", *arrPreItems, "自主廃棄重量（0.72換算後）",
                ], arrOutputFields]
                pszOutputPath = (
                    create_step0008_group_discard_deduction_output_path(
                        pszStep0004Path, pszGroupName, objKey[2], objKey[3]
                    )
                )
                write_step0008_tsv(pszOutputPath, arrOutputRows)
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0008_GROUP_DISCARD_DEDUCTION_CLEANUP_FAILED",
                "group discard deduction creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0008_group_discard_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int]:
    objMatch = re.fullmatch(
        r"(.+)_step0008_グループ別_控除廃棄_(契約|直営|フリー)_"
        r"(\d{4})_(\d{2})(?:_\d+)*\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0008_GROUP_DISCARD_FILE_NAME",
            "invalid group discard deduction step0008 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return objMatch.group(1), objMatch.group(2), iYear, iMonth


def read_step0008_group_discard_row(
    pszInputPath: str,
    iExpectedYear: int,
    iExpectedMonth: int,
    pszExpectedGroupName: str,
) -> List[str]:
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数", "出荷羽数",
        "全廃棄羽数", "出荷重量", "全廃棄重量", "平均体重",
        "懸鳥前死鶏羽数", "懸鳥前発育不良羽数",
        "懸鳥前死鶏重量", "懸鳥前発育不良重量", "自主廃棄重量（0.72換算後）",
    ]
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0008_GROUP_DISCARD_READ_FAILED",
            "failed to read group discard deduction step0008 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    if len(arrRows) != 2 or arrRows[0] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0008_GROUP_DISCARD_ROWS",
            "step0008 group discard TSV must contain the expected header and "
            "one data row. 旧形式の10列・12列TSVはstep0008から再作成してください。Path = " + pszInputPath,
        )
    arrFields = arrRows[1]
    if len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0008_GROUP_DISCARD_COLUMNS",
            "step0008 must contain 15 columns. step0008から再作成してください。Path = "
            + pszInputPath,
        )
    try:
        iRowYear, iRowMonth = parse_year_month(arrFields[0], arrFields[1], 2)
        pszGroupNumber = str(int(arrFields[2].replace(",", "").strip()))
    except Exception as objException:
        raise_conversion_error(
            "INVALID_STEP0008_GROUP_DISCARD_KEY",
            "invalid year, month, or group number in step0008 row. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    dicGroupNumbers = {pszName: pszNumber for pszNumber, pszName in GROUPS}
    if (
        iRowYear != iExpectedYear
        or iRowMonth != iExpectedMonth
        or pszGroupNumber != dicGroupNumbers[pszExpectedGroupName]
        or arrFields[3].strip(" \u3000\t") != pszExpectedGroupName
    ):
        raise_conversion_error(
            "STEP0008_GROUP_DISCARD_KEY_MISMATCH",
            "step0008 row keys did not match the file name. Path = "
            + pszInputPath,
        )
    for iColumn, pszItem in enumerate(arrExpectedHeader[10:], start=10):
        try:
            parse_step0008_group_discard_numeric_value(arrFields[iColumn], pszItem, pszInputPath)
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0008_PRE_SHACKLING_VALUE",
                f"Year/month = {iExpectedYear:04d}/{iExpectedMonth:02d}; "
                f"Group = {pszExpectedGroupName}; Item = {pszItem}; "
                f"Path = {pszInputPath}; Detail = {objException}",
            )
    return arrFields


def parse_step0008_group_discard_numeric_value(
    pszValue: str, pszColumnName: str, pszInputPath: str
) -> Decimal:
    objValue = parse_decimal_value(
        pszValue.replace(",", ""), 2, 1, pszColumnName
    )
    if objValue is None:
        raise_conversion_error(
            "INVALID_STEP0008_GROUP_DISCARD_VALUE",
            "step0008 numeric value was blank. Column = " + pszColumnName
            + "; Path = " + pszInputPath,
        )
    return objValue


def create_step0009_group_discard_output_path(
    pszStep0008Path: str,
) -> str:
    pszDirectory = os.path.dirname(pszStep0008Path)
    pszFileName = os.path.basename(pszStep0008Path)
    if "_step0008_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0008_FOR_GROUP_STEP0009",
            "step0008 marker was not found. Path = " + pszStep0008Path,
        )
    pszFileName = pszFileName.replace("_step0008_", "_step0009_", 1)
    pszStem, pszExtension = os.path.splitext(pszFileName)
    pszOutputPath = os.path.join(pszDirectory, pszFileName)
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0009_group_discard_row(
    arrInputFields: List[str], pszInputPath: str
) -> List[str]:
    objProcessedCount = (
        parse_step0008_group_discard_numeric_value(
            arrInputFields[5], "出荷羽数", pszInputPath
        )
        - parse_step0008_group_discard_numeric_value(
            arrInputFields[6], "全廃棄羽数", pszInputPath
        )
        - parse_step0008_group_discard_numeric_value(
            arrInputFields[10], "懸鳥前死鶏羽数", pszInputPath
        )
        - parse_step0008_group_discard_numeric_value(
            arrInputFields[11], "懸鳥前発育不良羽数", pszInputPath
        )
    )
    objProcessedWeight = (
        parse_step0008_group_discard_numeric_value(
            arrInputFields[7], "出荷重量", pszInputPath
        )
        - parse_step0008_group_discard_numeric_value(
            arrInputFields[12], "懸鳥前死鶏重量", pszInputPath
        )
        - parse_step0008_group_discard_numeric_value(
            arrInputFields[13], "懸鳥前発育不良重量", pszInputPath
        )
        - parse_step0008_group_discard_numeric_value(
            arrInputFields[14], "自主廃棄重量（0.72換算後）", pszInputPath
        )
    ).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)
    objAverageWeight = (
        (objProcessedWeight / objProcessedCount).quantize(
            Decimal("0.00000001"), rounding=ROUND_HALF_UP
        )
        if objProcessedCount != 0 else Decimal("0")
    )
    return [
        arrInputFields[0].strip(),
        arrInputFields[1].strip(),
        arrInputFields[2].strip(),
        arrInputFields[3].strip(" \u3000\t"),
        format_sum_value(
            parse_step0008_group_discard_numeric_value(
                arrInputFields[4], "入雛羽数", pszInputPath
            )
        ),
        format_sum_value(objProcessedCount),
        format(objProcessedWeight, ".1f"),
        format_sum_value(objAverageWeight),
    ]


def create_step0009_group_discard_files(
    arrStep0008GroupDiscardPaths: List[str],
) -> List[str]:
    """step0008の廃棄数・重量を控除し、グループ別step0009を出力します。"""
    dicInputs: Dict[
        Tuple[str, str, int, int], Dict[str, Tuple[str, List[str]]]
    ] = {}
    for pszPath in arrStep0008GroupDiscardPaths:
        pszBaseName, pszGroupName, iYear, iMonth = (
            parse_step0008_group_discard_file_name(pszPath)
        )
        objKey = (
            os.path.dirname(os.path.abspath(pszPath)),
            pszBaseName, iYear, iMonth,
        )
        dicGroups = dicInputs.setdefault(objKey, {})
        if pszGroupName in dicGroups:
            raise_conversion_error(
                "DUPLICATE_STEP0008_GROUP_DISCARD_INPUT",
                "duplicate step0008 group/month input. Path = " + pszPath,
            )
        arrFields = read_step0008_group_discard_row(
            pszPath, iYear, iMonth, pszGroupName
        )
        dicGroups[pszGroupName] = (pszPath, arrFields)

    arrCreatedPaths: List[str] = []
    arrExpectedGroups = [pszName for _, pszName in GROUPS]
    arrOutputHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数", "処理羽数",
        "処理重量", "平均体重",
    ]
    try:
        for objKey in sorted(dicInputs):
            dicGroups = dicInputs[objKey]
            arrMissingGroups = [
                pszGroupName for pszGroupName in arrExpectedGroups
                if pszGroupName not in dicGroups
            ]
            if arrMissingGroups:
                raise_conversion_error(
                    "MISSING_STEP0008_GROUP_DISCARD_INPUT",
                    "step0008 inputs must contain contract, direct, and free "
                    "groups. Key = " + str(objKey) + "; Groups = "
                    + ", ".join(arrMissingGroups),
                )
            for pszGroupName in arrExpectedGroups:
                pszStep0008Path, arrInputFields = dicGroups[pszGroupName]
                arrOutputRow = create_step0009_group_discard_row(
                    arrInputFields, pszStep0008Path
                )
                pszOutputPath = create_step0009_group_discard_output_path(
                    pszStep0008Path
                )
                try:
                    write_monthly_tsv(
                        pszOutputPath,
                        "\t".join(arrOutputHeader),
                        ["\t".join(arrOutputRow)],
                    )
                except ConversionError as objException:
                    raise_conversion_error(
                        "STEP0009_GROUP_DISCARD_WRITE_FAILED",
                        str(objException),
                    )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0009_GROUP_DISCARD_CLEANUP_FAILED",
                "group step0009 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0009_group_discard_file_name(
    pszInputPath: str,
) -> Tuple[str, str, int, int]:
    objMatch = re.fullmatch(
        r"(.+)_step0009_グループ別_控除廃棄_(契約|直営|フリー)_"
        r"(\d{4})_(\d{2})(?:_\d+)*\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0009_GROUP_DISCARD_FILE_NAME",
            "invalid group discard deduction step0009 file name. Path = "
            + pszInputPath,
        )
    iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
    return objMatch.group(1), objMatch.group(2), iYear, iMonth


def read_step0009_group_discard_row(
    pszInputPath: str,
    iExpectedYear: int,
    iExpectedMonth: int,
    pszExpectedGroupName: str,
) -> Tuple[str, str]:
    arrExpectedHeader = [
        "年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数", "処理羽数",
        "処理重量", "平均体重",
    ]
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0009_GROUP_DISCARD_READ_FAILED",
            "failed to read group discard deduction step0009 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0009_GROUP_DISCARD_ROW_COUNT",
            "step0009 group discard TSV must contain one header and one data "
            "row. Path = " + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    if arrHeader != arrExpectedHeader or len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0009_GROUP_DISCARD_COLUMNS",
            "unexpected step0009 group discard header or column count. Path = "
            + pszInputPath,
        )
    try:
        iRowYear, iRowMonth = parse_year_month(arrFields[0], arrFields[1], 2)
        pszGroupNumber = str(int(arrFields[2].replace(",", "").strip()))
    except Exception as objException:
        raise_conversion_error(
            "INVALID_STEP0009_GROUP_DISCARD_KEY",
            "invalid year, month, or group number in step0009 row. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    dicGroupNumbers = {pszName: pszNumber for pszNumber, pszName in GROUPS}
    if (
        (iRowYear, iRowMonth) != (iExpectedYear, iExpectedMonth)
        or pszGroupNumber != dicGroupNumbers[pszExpectedGroupName]
        or arrFields[3].strip(" \u3000\t") != pszExpectedGroupName
    ):
        raise_conversion_error(
            "STEP0009_GROUP_DISCARD_KEY_MISMATCH",
            "step0009 row keys did not match the file name. Path = "
            + pszInputPath,
        )
    return arrLines[0], arrLines[1]


def create_step0010_group_discard_output_path(
    pszSourcePath: str, pszGroupName: str, iYear: int
) -> str:
    pszDirectory = os.path.dirname(pszSourcePath)
    pszBaseName, _, _, _ = parse_step0009_group_discard_file_name(
        pszSourcePath
    )
    pszStem = (
        f"{pszBaseName}_step0010_グループ別_控除廃棄_{pszGroupName}_"
        f"{iYear:04d}_年間"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0010_group_discard_annual_files(
    arrStep0009GroupDiscardPaths: List[str],
) -> List[str]:
    """step0009月別行をグループ・年ごとのstep0010 TSVへ連結します。"""
    dicRowsByYear: Dict[
        Tuple[str, str, str, int], Dict[int, Tuple[str, str, str]]
    ] = {}
    for pszPath in arrStep0009GroupDiscardPaths:
        pszBaseName, pszGroupName, iYear, iMonth = (
            parse_step0009_group_discard_file_name(pszPath)
        )
        pszHeader, pszRow = read_step0009_group_discard_row(
            pszPath, iYear, iMonth, pszGroupName
        )
        objKey = (
            os.path.dirname(os.path.abspath(pszPath)),
            pszBaseName,
            pszGroupName,
            iYear,
        )
        dicMonths = dicRowsByYear.setdefault(objKey, {})
        if iMonth in dicMonths:
            raise_conversion_error(
                "DUPLICATE_STEP0009_GROUP_DISCARD_MONTH",
                "duplicate group step0009 year/month input. Path = "
                + pszPath,
            )
        dicMonths[iMonth] = (pszPath, pszHeader, pszRow)

    arrCreatedPaths: List[str] = []
    try:
        for objKey in sorted(dicRowsByYear):
            dicMonths = dicRowsByYear[objKey]
            arrMonths = sorted(dicMonths)
            pszFirstPath, pszHeader, _ = dicMonths[arrMonths[0]]
            arrRows: List[str] = []
            for iMonth in arrMonths:
                pszPath, pszMonthHeader, pszRow = dicMonths[iMonth]
                if pszMonthHeader != pszHeader:
                    raise_conversion_error(
                        "STEP0009_GROUP_DISCARD_HEADER_MISMATCH",
                        "step0009 headers did not match within the annual "
                        "group. Path = " + pszPath,
                    )
                arrRows.append(pszRow)
            pszOutputPath = create_step0010_group_discard_output_path(
                pszFirstPath, objKey[2], objKey[3]
            )
            try:
                write_monthly_tsv(pszOutputPath, pszHeader, arrRows)
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0010_GROUP_DISCARD_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0010_GROUP_DISCARD_CLEANUP_FAILED",
                "annual group step0010 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0010_for_step0011(pszInputPath: str) -> List[List[str]]:
    objFileNameMatch = re.search(
        r"_step0010_グループ別_控除廃棄_(契約|直営|フリー)_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0010_FILENAME",
            "unexpected step0010 file name. Path = " + pszInputPath,
        )
    pszExpectedGroupName: str = objFileNameMatch.group(1)
    iExpectedYear: int = int(objFileNameMatch.group(2))
    dicExpectedGroups: Dict[str, str] = {
        pszGroupName: pszGroupNumber for pszGroupNumber, pszGroupName in GROUPS
    }
    pszExpectedGroupNumber: str = dicExpectedGroups[pszExpectedGroupName]

    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0010_READ_FAILED",
            "failed to read step0010 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    if len(arrRows) < 2:
        raise_conversion_error(
            "INVALID_STEP0010_ROW_COUNT",
            "step0010 must contain one header and at least one data row. Path = "
            + pszInputPath,
        )
    arrExpectedHeader: List[str] = STEP0010_GROUP_DEDUCTION_HEADER.split("\t")
    if [pszValue.strip(" \u3000\t") for pszValue in arrRows[0]] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0010_HEADER",
            "unexpected step0010 header. Path = " + pszInputPath,
        )

    iPreviousMonth: int = 0
    setMonths = set()
    for iLineNumber, arrFields in enumerate(arrRows[1:], start=2):
        if len(arrFields) != len(arrExpectedHeader):
            raise_conversion_error(
                "INVALID_STEP0010_COLUMN_COUNT",
                "step0010 row must contain 8 columns. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber)
                + ", Actual = "
                + str(len(arrFields)),
            )
        iYear, iMonth = parse_year_month(
            arrFields[0], arrFields[1], iLineNumber
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0010_YEAR",
                "step0010 year did not match its file name. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber),
            )
        if iMonth in setMonths:
            raise_conversion_error(
                "DUPLICATE_STEP0010_MONTH",
                "duplicate month in step0010. Path = "
                + pszInputPath
                + ", Month = "
                + str(iMonth),
            )
        if iMonth <= iPreviousMonth:
            raise_conversion_error(
                "INVALID_STEP0010_MONTH_ORDER",
                "step0010 months were not in ascending order. Path = "
                + pszInputPath,
            )
        setMonths.add(iMonth)
        iPreviousMonth = iMonth
        if (
            arrFields[2].strip(" \u3000\t") != pszExpectedGroupNumber
            or arrFields[3].strip(" \u3000\t") != pszExpectedGroupName
        ):
            raise_conversion_error(
                "INVALID_STEP0010_GROUP",
                "step0010 group did not match its file name. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber),
            )
        for iColumnIndex, (pszItemName, bAllowBlank) in enumerate(
            STEP0010_GROUP_DEDUCTION_ITEMS, start=4
        ):
            try:
                objValue = parse_decimal_value(
                    arrFields[iColumnIndex],
                    iLineNumber,
                    iColumnIndex,
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0010_ITEM_VALUE",
                    "invalid step0010 item value. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iLineNumber)
                    + ", Item = "
                    + pszItemName
                    + ". Detail = "
                    + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0010_ITEM_VALUE",
                    "step0010 item value was blank. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iLineNumber)
                    + ", Item = "
                    + pszItemName,
                )
    return arrRows



def create_step0011_group_deduction_output_path(pszInputPath: str) -> str:
    pszDirectory = os.path.dirname(pszInputPath)
    pszFileName = os.path.basename(pszInputPath)
    if "_step0010_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0010_FILENAME",
            "step0010 file name did not contain the expected marker. Path = "
            + pszInputPath,
        )
    pszStep0011FileName = pszFileName.replace("_step0010_", "_step0011_", 1)
    pszStem, pszExtension = os.path.splitext(pszStep0011FileName)
    pszOutputPath = os.path.join(pszDirectory, pszStep0011FileName)
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0011_group_deduction_files(
    arrStep0010GroupDeductionPaths: List[str],
) -> List[str]:
    """step0010の控除廃棄グループ年間表を転置してstep0011を作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszInputPath in arrStep0010GroupDeductionPaths:
            arrInputRows = read_step0010_for_step0011(pszInputPath)
            arrTransposedRows = transpose_tsv_rows(arrInputRows)
            pszOutputPath = create_step0011_group_deduction_output_path(
                pszInputPath
            )
            write_step0006_tsv(pszOutputPath, arrTransposedRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0011_GROUP_DEDUCTION_CLEANUP_FAILED",
                "step0011 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise



def create_step0012_group_deduction_output_path(pszSourceTsvPath: str, iYear: int) -> str:
    pszDirectory: str = os.path.dirname(pszSourceTsvPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem: str = f"{pszBaseName}_step0012_グループ別_控除廃棄_{iYear:04d}_年間"
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def parse_step0011_group_deduction_file_name(pszStep0011Path: str) -> Tuple[int, str]:
    objMatch = re.search(
        r"_step0011_グループ別_控除廃棄_(契約|直営|フリー)_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszStep0011Path),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0011_GROUP_DEDUCTION_FILENAME",
            "unexpected step0011 file name. Path = " + pszStep0011Path,
        )
    return int(objMatch.group(2)), objMatch.group(1)


def read_step0011_for_step0012(
    pszInputPath: str, iExpectedYear: int, pszExpectedGroupName: str
) -> Dict[str, List[str]]:
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0011_GROUP_DEDUCTION_READ_FAILED",
            "failed to read step0011 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    arrExpectedRowHeaders: List[str] = STEP0010_GROUP_DEDUCTION_HEADER.split("\t")
    if len(arrRows) != len(arrExpectedRowHeaders):
        raise_conversion_error(
            "INVALID_STEP0011_GROUP_DEDUCTION_ROW_COUNT",
            "step0011 must contain 8 rows. Path = "
            + pszInputPath
            + ", Actual = "
            + str(len(arrRows)),
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount < 2 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0011_GROUP_DEDUCTION_COLUMN_COUNT",
            "step0011 rows must have an equal column count and at least one data column. Path = "
            + pszInputPath,
        )
    arrActualRowHeaders: List[str] = [
        arrRow[0].strip(" \u3000\t") for arrRow in arrRows
    ]
    if arrActualRowHeaders != arrExpectedRowHeaders:
        raise_conversion_error(
            "INVALID_STEP0011_GROUP_DEDUCTION_ROW_HEADER",
            "unexpected step0011 row headers. Path = " + pszInputPath,
        )
    dicRows: Dict[str, List[str]] = {
        arrRow[0].strip(" \u3000\t"): arrRow[1:] for arrRow in arrRows
    }
    dicExpectedGroups: Dict[str, str] = {
        pszGroupName: pszGroupNumber for pszGroupNumber, pszGroupName in GROUPS
    }
    pszExpectedGroupNumber: str = dicExpectedGroups[pszExpectedGroupName]
    arrMonths: List[int] = []
    for iValueIndex in range(iColumnCount - 1):
        iLineNumber: int = iValueIndex + 2
        iYear, iMonth = parse_year_month(
            dicRows["年"][iValueIndex], dicRows["月"][iValueIndex], iLineNumber
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0011_GROUP_DEDUCTION_YEAR",
                "step0011 year did not match its file name. Path = "
                + pszInputPath
                + ", Column = "
                + str(iValueIndex + 2),
            )
        if iMonth in arrMonths:
            raise_conversion_error(
                "DUPLICATE_STEP0011_GROUP_DEDUCTION_MONTH",
                "duplicate month in step0011. Path = "
                + pszInputPath
                + ", Month = "
                + str(iMonth),
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0011_GROUP_DEDUCTION_MONTH",
                "step0011 months were not in ascending order. Path = "
                + pszInputPath,
            )
        arrMonths.append(iMonth)
        if (
            dicRows["ｸﾞﾙｰﾌﾟ№"][iValueIndex].strip(" \u3000\t")
            != pszExpectedGroupNumber
            or dicRows["ｸﾞﾙｰﾌﾟ名"][iValueIndex].strip(" \u3000\t")
            != pszExpectedGroupName
        ):
            raise_conversion_error(
                "INVALID_STEP0011_GROUP_DEDUCTION_GROUP",
                "step0011 group did not match its file name. Path = "
                + pszInputPath
                + ", Column = "
                + str(iValueIndex + 2),
            )
        for pszItemName, bAllowBlank in STEP0010_GROUP_DEDUCTION_ITEMS:
            try:
                objValue = parse_decimal_value(
                    dicRows[pszItemName][iValueIndex],
                    iLineNumber,
                    arrExpectedRowHeaders.index(pszItemName),
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0011_GROUP_DEDUCTION_ITEM_VALUE",
                    "invalid step0011 item value. Path = "
                    + pszInputPath
                    + ", Item = "
                    + pszItemName
                    + ". Detail = "
                    + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0011_GROUP_DEDUCTION_ITEM_VALUE",
                    "step0011 item value was blank. Path = "
                    + pszInputPath
                    + ", Item = "
                    + pszItemName
                    + ", Column = "
                    + str(iValueIndex + 2),
                )
    dicRows["月"] = [f"{iMonth:02d}" for iMonth in arrMonths]
    return dicRows


def create_step0012_group_deduction_rows(
    pszGroupName: str, dicStep0011Rows: Dict[str, List[str]]
) -> List[str]:
    """年別の控除廃棄5項目を1グループ分出力します。"""
    arrOutputRows: List[str] = []
    for iItemIndex, (pszItemName, _) in enumerate(STEP0010_GROUP_DEDUCTION_ITEMS):
        arrValues: List[str] = list(dicStep0011Rows[pszItemName])
        if pszItemName == "平均体重":
            arrValues = [
                format_step0007_average_weight(pszValue) for pszValue in arrValues
            ]
        arrOutputRows.append(
            "\t".join(
                ([pszGroupName if iItemIndex == 0 else "", pszItemName] + arrValues)
            )
        )
    arrGrowthValues: List[str] = []
    for pszPlacementCount, pszProcessedCount in zip(
        dicStep0011Rows["入雛羽数"], dicStep0011Rows["処理羽数"]
    ):
        pszGrowthRate, _ = calculate_growth_rate(
            pszPlacementCount, pszProcessedCount
        )
        arrGrowthValues.append(pszGrowthRate)
    arrOutputRows.append("\t".join(["", "育成率"] + arrGrowthValues))
    return arrOutputRows
def write_step0012_group_deduction_tsv(pszOutputPath: str, arrRows: List[str]) -> None:
    if not arrRows:
        raise_conversion_error(
            "INVALID_STEP0012_GROUP_DEDUCTION_ROW_COUNT", "step0012 did not contain any rows."
        )
    try:
        write_monthly_tsv(pszOutputPath, arrRows[0], arrRows[1:])
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0012_GROUP_DEDUCTION_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0012_GROUP_DEDUCTION_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0012_GROUP_DEDUCTION_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0012_GROUP_DEDUCTION_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0012_GROUP_DEDUCTION_WRITE_FAILED"),
            str(objException),
        )


def create_step0012_group_deduction_files(
    pszSourceTsvPath: str, arrStep0011Paths: List[str]
) -> Tuple[List[str], List[int]]:
    """控除廃棄版step0011を年ごとに契約・直営・フリー順で結合します。"""
    dicPaths: Dict[Tuple[int, str], str] = {}
    for pszPath in arrStep0011Paths:
        iYear, pszGroupName = parse_step0011_group_deduction_file_name(pszPath)
        objKey: Tuple[int, str] = (iYear, pszGroupName)
        if objKey in dicPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0011_GROUP_DEDUCTION_YEAR_GROUP",
                "duplicate step0011 year/group. Key = " + str(objKey),
            )
        dicPaths[objKey] = pszPath
    arrYears: List[int] = sorted({iYear for iYear, _ in dicPaths})
    arrCreatedPaths: List[str] = []
    try:
        for iYear in arrYears:
            dicGroupRows: Dict[str, Dict[str, List[str]]] = {}
            for _, pszGroupName in GROUPS:
                objKey = (iYear, pszGroupName)
                if objKey not in dicPaths:
                    raise_conversion_error(
                        "MISSING_STEP0011_GROUP_DEDUCTION_GROUP",
                        "required step0011 group was missing. Key = " + str(objKey),
                    )
                dicGroupRows[pszGroupName] = read_step0011_for_step0012(
                    dicPaths[objKey], iYear, pszGroupName
                )
            arrReferenceMonths: List[str] = dicGroupRows[GROUPS[0][1]]["月"]
            for _, pszGroupName in GROUPS[1:]:
                if dicGroupRows[pszGroupName]["月"] != arrReferenceMonths:
                    raise_conversion_error(
                        "STEP0011_GROUP_DEDUCTION_MONTH_MISMATCH",
                        "step0011 month columns did not match between groups. Year = "
                        + str(iYear)
                        + ", Group = "
                        + pszGroupName,
                    )
            iMonthCount: int = len(arrReferenceMonths)
            arrRows: List[str] = [
                "\t".join(["", "年"] + [f"{iYear:04d}"] * iMonthCount),
                "\t".join(["", "月"] + arrReferenceMonths),
            ]
            for _, pszGroupName in GROUPS:
                arrRows.extend(
                    create_step0012_group_deduction_rows(
                        pszGroupName, dicGroupRows[pszGroupName]
                    )
                )
            arrAdditionalMonths: List[str] = [
                f"{iMonth:02d}" for iMonth in range(int(arrReferenceMonths[-1]) + 1, 13)
            ]
            if arrAdditionalMonths:
                dicFillValues: Dict[str, str] = {
                    "入雛羽数": "0",
                    "処理羽数": "0",
                    "処理重量": "0.0",
                    "平均体重": "0.000",
                    "育成率": "",
                }
                for iRowIndex, pszRow in enumerate(arrRows):
                    pszItemName: str = pszRow.split("\t")[1]
                    if pszItemName == "年":
                        arrAdditionalValues = [f"{iYear:04d}"] * len(arrAdditionalMonths)
                    elif pszItemName == "月":
                        arrAdditionalValues = arrAdditionalMonths
                    else:
                        arrAdditionalValues = [dicFillValues[pszItemName]] * len(arrAdditionalMonths)
                    arrRows[iRowIndex] = pszRow + "\t" + "\t".join(arrAdditionalValues)
            pszOutputPath: str = create_step0012_group_deduction_output_path(
                pszSourceTsvPath, iYear
            )
            write_step0012_group_deduction_tsv(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, arrYears
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0012_GROUP_DEDUCTION_CLEANUP_FAILED",
                "step0012 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise



def parse_step0012_group_deduction_file_name(pszInputPath: str) -> int:
    objMatch = re.search(
        r"_step0012_グループ別_控除廃棄_(\d{4})_年間(?:_\d+)*\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0012_GROUP_DEDUCTION_FILENAME",
            "unexpected step0012 group deduction file name. Path = " + pszInputPath,
        )
    return int(objMatch.group(1))


def read_step0012_for_step0013(
    pszInputPath: str, iExpectedYear: int
) -> Tuple[List[str], Dict[str, Dict[str, List[str]]]]:
    iFileYear = parse_step0012_group_deduction_file_name(pszInputPath)
    if iFileYear != iExpectedYear:
        raise_conversion_error(
            "INVALID_STEP0012_GROUP_DEDUCTION_YEAR",
            "step0012 year did not match its file name. Path = " + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0012_GROUP_DEDUCTION_READ_FAILED",
            "failed to read step0012 group deduction TSV. Path = "
            + pszInputPath + "; Detail = " + str(objException),
        )
    if len(arrRows) != len(STEP0012_GROUP_DEDUCTION_ROW_STRUCTURE):
        raise_conversion_error(
            "INVALID_STEP0012_GROUP_DEDUCTION_ROW_COUNT",
            "step0012 group deduction TSV must contain 17 rows. Path = "
            + pszInputPath,
        )
    iColumnCount = len(arrRows[0])
    if iColumnCount < 3 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0012_GROUP_DEDUCTION_COLUMN_COUNT",
            "step0012 rows must have equal columns and monthly data. Path = "
            + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(
        STEP0012_GROUP_DEDUCTION_ROW_STRUCTURE
    ):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0012_GROUP_DEDUCTION_ROW_STRUCTURE",
                "unexpected step0012 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )

    dicExpectedGroupNumbers = {
        pszName: pszNumber for pszNumber, pszName in GROUPS
    }
    arrMonths: List[str] = []
    dicGroupRows: Dict[str, Dict[str, List[str]]] = {}
    arrItemNames = ["入雛羽数", "処理羽数", "処理重量", "平均体重", "育成率"]
    for iValueIndex in range(iColumnCount - 2):
        iLineNumber = iValueIndex + 2
        iYear, iMonth = parse_year_month(
            arrRows[0][iValueIndex + 2],
            arrRows[1][iValueIndex + 2],
            iLineNumber,
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0012_GROUP_DEDUCTION_YEAR",
                "step0012 row year did not match its file name. Path = "
                + pszInputPath + ", Column = " + str(iValueIndex + 3),
            )
        if arrMonths and iMonth <= int(arrMonths[-1]):
            raise_conversion_error(
                "INVALID_STEP0012_GROUP_DEDUCTION_MONTH",
                "step0012 months must be unique and ascending. Path = "
                + pszInputPath,
            )
        arrMonths.append(f"{iMonth:02d}")
        for iGroupIndex, (_, pszGroupName) in enumerate(GROUPS):
            iFirstRow = 2 + iGroupIndex * 5
            dicItemValues = dicGroupRows.setdefault(
                pszGroupName, {pszItemName: [] for pszItemName in arrItemNames}
            )
            pszGroupLabel = arrRows[iFirstRow][0].strip(" \u3000\t")
            if (
                pszGroupLabel != pszGroupName
                or any(
                    arrRows[iFirstRow + iItemIndex][0].strip(" \u3000\t")
                    != (pszGroupName if iItemIndex == 0 else "")
                    for iItemIndex in range(5)
                )
            ):
                raise_conversion_error(
                    "INVALID_STEP0012_GROUP_DEDUCTION_GROUP",
                    "step0012 group labels did not match the expected group order. "
                    "Path = " + pszInputPath,
                )
            for iItemIndex, pszItemName in enumerate(arrItemNames):
                pszValue = arrRows[iFirstRow + iItemIndex][iValueIndex + 2]
                objValue = parse_decimal_value(
                    pszValue,
                    iLineNumber,
                    iFirstRow + iItemIndex,
                    pszItemName,
                )
                if objValue is None and pszItemName not in {"平均体重", "育成率"}:
                    raise_conversion_error(
                        "INVALID_STEP0012_GROUP_DEDUCTION_ITEM_VALUE",
                        "step0012 item value was blank. Path = "
                        + pszInputPath + ", Item = " + pszItemName,
                    )
                if (
                    objValue is not None
                    and pszItemName in {"入雛羽数", "処理羽数"}
                    and objValue != objValue.to_integral_value()
                ):
                    raise_conversion_error(
                        "INVALID_STEP0012_GROUP_DEDUCTION_ITEM_VALUE",
                        "step0012 count was not an integer. Path = "
                        + pszInputPath + ", Item = " + pszItemName,
                    )
                dicItemValues[pszItemName].append(pszValue)
    return arrMonths, dicGroupRows


def create_step0013_group_deduction_output_path(pszInputPath: str) -> str:
    pszDirectory = os.path.dirname(pszInputPath)
    pszFileName = os.path.basename(pszInputPath)
    pszOutputName = pszFileName.replace("_step0012_", "_step0013_", 1)
    if pszOutputName == pszFileName:
        raise_conversion_error(
            "INVALID_STEP0012_GROUP_DEDUCTION_FILENAME",
            "step0012 marker was not found. Path = " + pszInputPath,
        )
    pszStem, pszExtension = os.path.splitext(pszOutputName)
    pszOutputPath = os.path.join(pszDirectory, pszOutputName)
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0013_group_deduction_rows(
    dicGroupRows: Dict[str, Dict[str, List[str]]],
    arrMonths: List[str],
) -> List[List[str]]:
    arrOutputRows: List[List[str]] = []
    arrItemNames = ["入雛羽数", "処理羽数", "処理重量", "平均体重", "育成率"]
    for _, pszGroupName in GROUPS:
        for pszItemName in arrItemNames:
            arrValues = dicGroupRows[pszGroupName][pszItemName]
            arrOutputRows.append(
                [
                    pszGroupName if pszItemName == "入雛羽数" else "",
                    pszItemName,
                ]
                + [
                    format_sum_value(Decimal(pszValue))
                    if pszValue.strip(" \u3000\t")
                    else ""
                    for pszValue in arrValues
                ]
            )

    arrTotals = [
        ["合計", "入雛羽数"],
        ["", "処理羽数"],
        ["", "処理重量"],
        ["", "平均体重"],
        ["", "育成率"],
    ]
    for iValueIndex in range(len(arrMonths)):
        objPlacement = sum(
            (
                Decimal(dicGroupRows[pszGroupName]["入雛羽数"][iValueIndex])
                for _, pszGroupName in GROUPS
            ),
            Decimal("0"),
        )
        objProcessed = sum(
            (
                Decimal(dicGroupRows[pszGroupName]["処理羽数"][iValueIndex])
                for _, pszGroupName in GROUPS
            ),
            Decimal("0"),
        )
        objProcessedWeight = sum(
            (
                Decimal(dicGroupRows[pszGroupName]["処理重量"][iValueIndex])
                for _, pszGroupName in GROUPS
            ),
            Decimal("0"),
        )
        objAverageWeight = (
            (objProcessedWeight / objProcessed).quantize(
                Decimal("0.001"), rounding=ROUND_HALF_UP
            )
            if objProcessed != 0 else Decimal("0")
        )
        objGrowthRate = (
            (objProcessed / objPlacement).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            )
            if objPlacement != 0 else Decimal("0")
        )
        for arrRow, objValue in zip(
            arrTotals,
            [
                objPlacement,
                objProcessed,
                objProcessedWeight,
                objAverageWeight,
                objGrowthRate,
            ],
        ):
            arrRow.append(format_sum_value(objValue))
    return arrOutputRows + arrTotals


def write_step0013_group_deduction_tsv(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    if not arrRows:
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_DEDUCTION_ROW_COUNT",
            "step0013 did not contain any rows.",
        )
    try:
        write_monthly_tsv(
            pszOutputPath,
            "\t".join(arrRows[0]),
            ["\t".join(arrRow) for arrRow in arrRows[1:]],
        )
    except ConversionError as objException:
        dicErrorCodes = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0013_GROUP_DEDUCTION_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0013_GROUP_DEDUCTION_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0013_GROUP_DEDUCTION_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0013_GROUP_DEDUCTION_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(
                objException.pszErrorCode,
                "STEP0013_GROUP_DEDUCTION_WRITE_FAILED",
            ),
            str(objException),
        )


def create_step0013_group_deduction_files(
    arrStep0012Paths: List[str],
) -> Tuple[List[str], List[int]]:
    dicPaths: Dict[int, str] = {}
    for pszPath in arrStep0012Paths:
        iYear = parse_step0012_group_deduction_file_name(pszPath)
        if iYear in dicPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0012_GROUP_DEDUCTION_YEAR",
                "duplicate step0012 group deduction year. Year = " + str(iYear),
            )
        dicPaths[iYear] = pszPath
    arrYears = sorted(dicPaths)
    arrCreatedPaths: List[str] = []
    try:
        for iYear in arrYears:
            arrMonths, dicGroupRows = read_step0012_for_step0013(
                dicPaths[iYear], iYear
            )
            arrRows = [
                ["", "年"] + [f"{iYear:04d}"] * len(arrMonths),
                ["", "月"] + arrMonths,
            ]
            arrRows.extend(
                create_step0013_group_deduction_rows(dicGroupRows, arrMonths)
            )
            pszOutputPath = create_step0013_group_deduction_output_path(
                dicPaths[iYear]
            )
            write_step0013_group_deduction_tsv(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, arrYears
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0013_GROUP_DEDUCTION_CLEANUP_FAILED",
                "step0013 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE: List[Tuple[str, str]] = [
    ("", "年"),
    ("", "月"),
]
for _, pszGroupName in GROUPS:
    STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE.extend(
        [
            (pszGroupName, "入雛羽数"),
            ("", "処理羽数"),
            ("", "処理重量"),
            ("", "平均体重"),
            ("", "育成率"),
        ]
    )
STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE.extend(
    [
        ("合計", "入雛羽数"),
        ("", "処理羽数"),
        ("", "処理重量"),
        ("", "平均体重"),
        ("", "育成率"),
    ]
)


def create_step0014_group_deduction_output_path(
    pszStep0013Path: str,
) -> str:
    pszDirectory = os.path.dirname(pszStep0013Path)
    pszFileName = os.path.basename(pszStep0013Path)
    objMatch = re.search(
        r"_step0013_グループ別_控除廃棄_(\d{4}_年間)(?:_(\d{2}))?\.tsv$",
        pszFileName,
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_DEDUCTION_FILENAME",
            "unexpected step0013 group deduction file name. Path = "
            + pszStep0013Path,
        )
    pszOutputName = re.sub(
        r"_step0013_グループ別_控除廃棄_",
        "_step0014_グループ別_控除廃棄_",
        pszFileName,
        count=1,
    )
    pszStem, pszExtension = os.path.splitext(pszOutputName)
    pszOutputPath = os.path.join(pszDirectory, pszOutputName)
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def read_step0013_for_step0014(
    pszInputPath: str,
) -> List[List[str]]:
    objMatch = re.search(
        r"_step0013_グループ別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_DEDUCTION_FILENAME",
            "unexpected step0013 group deduction file name. Path = "
            + pszInputPath,
        )
    iExpectedYear = int(objMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0014_GROUP_DEDUCTION_READ_FAILED",
            "failed to read step0013 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    if len(arrRows) != len(STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE):
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_DEDUCTION_ROW_COUNT",
            "step0013 TSV must contain 22 rows. Path = " + pszInputPath,
        )
    if any(len(arrRow) != 14 for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_DEDUCTION_COLUMN_COUNT",
            "step0013 TSV must contain 14 columns. Path = " + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(
        STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE
    ):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0013_GROUP_DEDUCTION_ROW_STRUCTURE",
                "unexpected step0013 row structure. Path = "
                + pszInputPath
                + ", Line = "
                + str(iRowIndex + 1),
            )
    for iMonthIndex in range(12):
        iColumnIndex = iMonthIndex + 2
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear or iMonth != iMonthIndex + 1:
            raise_conversion_error(
                "INVALID_STEP0013_GROUP_DEDUCTION_YEAR_MONTH",
                "step0013 year or month did not match its file name or column. Path = "
                + pszInputPath
                + ", Column = "
                + str(iColumnIndex + 1),
            )
        for iRowIndex in range(2, len(arrRows)):
            pszItemName = STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE[iRowIndex][1]
            objValue = parse_decimal_value(
                arrRows[iRowIndex][iColumnIndex],
                iRowIndex + 1,
                iColumnIndex,
                pszItemName,
            )
            if objValue is None and pszItemName not in {"平均体重", "育成率"}:
                raise_conversion_error(
                    "INVALID_STEP0013_GROUP_DEDUCTION_ITEM_VALUE",
                    "step0013 required value was blank. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iRowIndex + 1),
                )
            if (
                objValue is not None
                and pszItemName in {"入雛羽数", "処理羽数"}
                and objValue != objValue.to_integral_value()
            ):
                raise_conversion_error(
                    "INVALID_STEP0013_GROUP_DEDUCTION_ITEM_VALUE",
                    "step0013 count was not an integer. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iRowIndex + 1),
                )
    return arrRows


def create_step0014_group_deduction_rows(
    arrStep0013Rows: List[List[str]],
) -> List[List[str]]:
    arrOutputRows = [arrRow.copy() + [""] for arrRow in arrStep0013Rows]
    arrOutputRows[1][14] = "合計"
    dicAnnualValues: Dict[str, Dict[str, Decimal]] = {}
    for iGroupIndex, (_, pszGroupName) in enumerate(GROUPS):
        iFirstRow = 2 + iGroupIndex * 5
        dicValues: Dict[str, Decimal] = {}
        for iOffset, pszItemName in enumerate(
            ["入雛羽数", "処理羽数", "処理重量"]
        ):
            iRowIndex = iFirstRow + iOffset
            objTotal = sum(
                (
                    Decimal(arrStep0013Rows[iRowIndex][iColumnIndex])
                    for iColumnIndex in range(2, 14)
                ),
                Decimal("0"),
            )
            dicValues[pszItemName] = objTotal
            arrOutputRows[iRowIndex][14] = format_sum_value(objTotal)
        objProcessed = dicValues["処理羽数"]
        objPlacement = dicValues["入雛羽数"]
        objProcessedWeight = dicValues["処理重量"]
        objAverageWeight = (
            (objProcessedWeight / objProcessed).quantize(
                Decimal("0.001"), rounding=ROUND_HALF_UP
            )
            if objProcessed != 0
            else Decimal("0")
        )
        objGrowthRate = (
            (objProcessed / objPlacement).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            )
            if objPlacement != 0
            else Decimal("0")
        )
        arrOutputRows[iFirstRow + 3][14] = format_sum_value(objAverageWeight)
        arrOutputRows[iFirstRow + 4][14] = format_sum_value(objGrowthRate)
        dicAnnualValues[pszGroupName] = dicValues

    iTotalRow = 17
    dicTotalValues = {
        pszItemName: sum(
            (dicAnnualValues[pszGroupName][pszItemName] for _, pszGroupName in GROUPS),
            Decimal("0"),
        )
        for pszItemName in ["入雛羽数", "処理羽数", "処理重量"]
    }
    for iOffset, pszItemName in enumerate(
        ["入雛羽数", "処理羽数", "処理重量"]
    ):
        arrOutputRows[iTotalRow + iOffset][14] = format_sum_value(
            dicTotalValues[pszItemName]
        )
    objTotalProcessed = dicTotalValues["処理羽数"]
    objTotalPlacement = dicTotalValues["入雛羽数"]
    objTotalWeight = dicTotalValues["処理重量"]
    objTotalAverageWeight = (
        (objTotalWeight / objTotalProcessed).quantize(
            Decimal("0.001"), rounding=ROUND_HALF_UP
        )
        if objTotalProcessed != 0
        else Decimal("0")
    )
    objTotalGrowthRate = (
        (objTotalProcessed / objTotalPlacement).quantize(
            Decimal("0.0001"), rounding=ROUND_HALF_UP
        )
        if objTotalPlacement != 0
        else Decimal("0")
    )
    arrOutputRows[20][14] = format_sum_value(objTotalAverageWeight)
    arrOutputRows[21][14] = format_sum_value(objTotalGrowthRate)
    return arrOutputRows


def create_step0014_group_deduction_files(
    arrStep0013Paths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0013Path in arrStep0013Paths:
            arrInputRows = read_step0013_for_step0014(pszStep0013Path)
            arrOutputRows = create_step0014_group_deduction_rows(arrInputRows)
            pszOutputPath = create_step0014_group_deduction_output_path(
                pszStep0013Path
            )
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrOutputRows[0]),
                    ["\t".join(arrRow) for arrRow in arrOutputRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0014_GROUP_DEDUCTION_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0014_GROUP_DEDUCTION_CLEANUP_FAILED",
                "step0014 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise



def create_step0006_discard_summary_upper_output_path(
    pszPlacementAveragePath: str,
    pszSuffix: str,
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0004_総計_羽数データ別_入雛羽数・平均体重_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszPlacementAveragePath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0004_PLACEMENT_AVERAGE_FILE_NAME",
            "invalid placement/average step0004 file name. Path = "
            + pszPlacementAveragePath,
        )
    pszOutputStem = (
        f"{objMatch.group(1)}_step0006_総計_廃棄データ一覧表_上段_"
        + f"{objMatch.group(2)}_{objMatch.group(3)}"
        + pszSuffix
    )
    pszDirectory = os.path.dirname(pszPlacementAveragePath)
    pszOutputPath = os.path.join(pszDirectory, pszOutputStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszOutputStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def read_step0006_discard_summary_upper_input(
    pszInputPath: str,
    arrExpectedHeader: List[str],
    iExpectedYear: int,
    iExpectedMonth: int,
) -> Tuple[List[str], List[str]]:
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0006_DISCARD_SUMMARY_UPPER_READ_FAILED",
            "failed to read step0006 upper-section input. Path = " + pszInputPath
            + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0006_DISCARD_SUMMARY_UPPER_ROW_COUNT",
            "upper-section input must contain one header and one total row. Path = "
            + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    if arrHeader != arrExpectedHeader or len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0006_DISCARD_SUMMARY_UPPER_COLUMNS",
            "unexpected upper-section input columns. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (
        iExpectedYear, iExpectedMonth
    ):
        raise_conversion_error(
            "STEP0006_DISCARD_SUMMARY_UPPER_YEAR_MONTH_MISMATCH",
            "file name and row year/month do not match. Path = " + pszInputPath,
        )
    if arrFields[2:4] != ["", "総計"]:
        raise_conversion_error(
            "INVALID_STEP0006_DISCARD_SUMMARY_UPPER_GROUP",
            "expected an empty group number and total group name. Path = "
            + pszInputPath,
        )
    return arrHeader, arrFields


def create_step0006_discard_summary_upper_files(
    arrPlacementAveragePaths: List[str],
    arrPreSlaughterCountPaths: List[str],
    arrDiscardCountPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """入雛・平均体重、懸鳥前羽数、廃棄羽数の総計行を横結合します。"""
    objPatterns = {
        "placement": re.compile(
            r"(.+)_step0004_総計_羽数データ別_入雛羽数・平均体重_"
            r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv"
        ),
        "pre_slaughter": re.compile(
            r"(.+)_step0005_総計_懸鳥前データ別_羽数_"
            r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv"
        ),
        "discard": re.compile(
            r"(.+)_step0005_総計_廃棄データ別_羽数_"
            r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv"
        ),
    }
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszKind, arrPaths in [
        ("placement", arrPlacementAveragePaths),
        ("pre_slaughter", arrPreSlaughterCountPaths),
        ("discard", arrDiscardCountPaths),
    ]:
        for pszPath in arrPaths:
            objMatch = objPatterns[pszKind].fullmatch(os.path.basename(pszPath))
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_STEP0006_DISCARD_SUMMARY_UPPER_FILE_NAME",
                    "invalid upper-section input file name. Path = " + pszPath,
                )
            iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
            objKey = (
                os.path.dirname(os.path.abspath(pszPath)), objMatch.group(1),
                iYear, iMonth, objMatch.group(4),
            )
            dicPair = dicInputs.setdefault(objKey, {})
            if pszKind in dicPair:
                raise_conversion_error(
                    "DUPLICATE_STEP0006_DISCARD_SUMMARY_UPPER_INPUT",
                    "duplicate upper-section input. Path = " + pszPath,
                )
            dicPair[pszKind] = pszPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicPair = dicInputs[objKey]
            if set(dicPair) != {"placement", "pre_slaughter", "discard"}:
                raise_conversion_error(
                    "MISSING_STEP0006_DISCARD_SUMMARY_UPPER_INPUT",
                    "placement, pre-slaughter and discard count inputs must be paired for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}",
                )
            arrPlacementHeader, arrPlacementFields = (
                read_step0006_discard_summary_upper_input(
                    dicPair["placement"],
                    ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数", "平均体重"],
                    objKey[2], objKey[3],
                )
            )
            arrPreHeader, arrPreFields = read_step0006_discard_summary_upper_input(
                dicPair["pre_slaughter"],
                ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名"]
                + STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS + ["小計"],
                objKey[2], objKey[3],
            )
            arrDiscardHeader, arrDiscardFields = read_step0006_discard_summary_upper_input(
                dicPair["discard"],
                ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名"]
                + STEP0004_TOTAL_COUNT_ITEMS + ["小計"],
                objKey[2], objKey[3],
            )
            if not (
                arrPlacementFields[:4] == arrPreFields[:4]
                == arrDiscardFields[:4]
            ):
                raise_conversion_error(
                    "STEP0006_DISCARD_SUMMARY_UPPER_COMMON_COLUMNS_MISMATCH",
                    "input common columns do not match. Key = " + str(objKey),
                )
            arrOutputHeader = (
                arrPlacementHeader
                + arrPreHeader[4:]
                + arrDiscardHeader[4:]
            )
            arrOutputFields = (
                arrPlacementFields
                + arrPreFields[4:]
                + arrDiscardFields[4:]
            )
            pszOutputPath = create_step0006_discard_summary_upper_output_path(
                dicPair["placement"], objKey[4]
            )
            write_step0006_tsv(pszOutputPath, [arrOutputHeader, arrOutputFields])
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0006_DISCARD_SUMMARY_UPPER_CLEANUP_FAILED",
                "upper-section combination failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0006_discard_summary_lower_output_path(
    pszShipmentCountWeightPath: str,
    pszSuffix: str,
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0004_総計_羽数データ別_出荷羽数・出荷重量_(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszShipmentCountWeightPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0004_SHIPMENT_COUNT_WEIGHT_FILE_NAME",
            "invalid shipment count/weight step0004 file name. Path = "
            + pszShipmentCountWeightPath,
        )
    pszOutputStem = (
        f"{objMatch.group(1)}_step0006_総計_廃棄データ一覧表_下段_"
        + f"{objMatch.group(2)}_{objMatch.group(3)}"
        + pszSuffix
    )
    pszDirectory = os.path.dirname(pszShipmentCountWeightPath)
    pszOutputPath = os.path.join(pszDirectory, pszOutputStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszOutputStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def read_step0006_discard_summary_lower_input(
    pszInputPath: str,
    arrExpectedHeader: List[str],
    iExpectedYear: int,
    iExpectedMonth: int,
) -> Tuple[List[str], List[str]]:
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0006_DISCARD_SUMMARY_LOWER_READ_FAILED",
            "failed to read step0006 lower-section input. Path = " + pszInputPath
            + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0006_DISCARD_SUMMARY_LOWER_ROW_COUNT",
            "lower-section input must contain one header and one total row. Path = "
            + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    if arrHeader != arrExpectedHeader or len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0006_DISCARD_SUMMARY_LOWER_COLUMNS",
            "unexpected lower-section input columns. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (
        iExpectedYear, iExpectedMonth
    ):
        raise_conversion_error(
            "STEP0006_DISCARD_SUMMARY_LOWER_YEAR_MONTH_MISMATCH",
            "file name and row year/month do not match. Path = " + pszInputPath,
        )
    if arrFields[2:4] != ["", "総計"]:
        raise_conversion_error(
            "INVALID_STEP0006_DISCARD_SUMMARY_LOWER_GROUP",
            "expected an empty group number and total group name. Path = "
            + pszInputPath,
        )
    return arrHeader, arrFields


def create_step0006_discard_summary_lower_files(
    arrShipmentCountWeightPaths: List[str],
    arrPreSlaughterWeightPaths: List[str],
    arrDiscardWeightPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """出荷羽数・重量、懸鳥前重量、廃棄重量の総計行を横結合します。"""
    objPatterns = {
        "shipment": re.compile(
            r"(.+)_step0004_総計_羽数データ別_出荷羽数・出荷重量_"
            r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv"
        ),
        "pre_slaughter_weight": re.compile(
            r"(.+)_step0005_総計_懸鳥前データ別_重量_"
            r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv"
        ),
        "discard_weight": re.compile(
            r"(.+)_step0005_総計_廃棄データ別_重量_"
            r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv"
        ),
    }
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszKind, arrPaths in [
        ("shipment", arrShipmentCountWeightPaths),
        ("pre_slaughter_weight", arrPreSlaughterWeightPaths),
        ("discard_weight", arrDiscardWeightPaths),
    ]:
        for pszPath in arrPaths:
            objMatch = objPatterns[pszKind].fullmatch(os.path.basename(pszPath))
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_STEP0006_DISCARD_SUMMARY_LOWER_FILE_NAME",
                    "invalid lower-section input file name. Path = " + pszPath,
                )
            iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
            objKey = (
                os.path.dirname(os.path.abspath(pszPath)), objMatch.group(1),
                iYear, iMonth, objMatch.group(4),
            )
            dicPair = dicInputs.setdefault(objKey, {})
            if pszKind in dicPair:
                raise_conversion_error(
                    "DUPLICATE_STEP0006_DISCARD_SUMMARY_LOWER_INPUT",
                    "duplicate lower-section input. Path = " + pszPath,
                )
            dicPair[pszKind] = pszPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicPair = dicInputs[objKey]
            if set(dicPair) != {"shipment", "pre_slaughter_weight", "discard_weight"}:
                raise_conversion_error(
                    "MISSING_STEP0006_DISCARD_SUMMARY_LOWER_INPUT",
                    "shipment, pre-slaughter and discard weight inputs must be paired for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}",
                )
            arrShipmentHeader, arrShipmentFields = (
                read_step0006_discard_summary_lower_input(
                    dicPair["shipment"],
                    ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "出荷羽数", "出荷重量"],
                    objKey[2], objKey[3],
                )
            )
            arrPreHeader, arrPreFields = read_step0006_discard_summary_lower_input(
                dicPair["pre_slaughter_weight"],
                ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名"]
                + STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS + ["小計"],
                objKey[2], objKey[3],
            )
            arrDiscardHeader, arrDiscardFields = read_step0006_discard_summary_lower_input(
                dicPair["discard_weight"],
                ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名"]
                + STEP0004_TOTAL_WEIGHT_ITEMS + ["小計"],
                objKey[2], objKey[3],
            )
            if not (
                arrShipmentFields[:4] == arrPreFields[:4]
                == arrDiscardFields[:4]
            ):
                raise_conversion_error(
                    "STEP0006_DISCARD_SUMMARY_LOWER_COMMON_COLUMNS_MISMATCH",
                    "input common columns do not match. Key = " + str(objKey),
                )
            arrOutputHeader = (
                arrShipmentHeader
                + arrPreHeader[4:]
                + arrDiscardHeader[4:]
            )
            arrOutputFields = (
                arrShipmentFields
                + arrPreFields[4:]
                + arrDiscardFields[4:]
            )
            pszOutputPath = create_step0006_discard_summary_lower_output_path(
                dicPair["shipment"], objKey[4]
            )
            write_step0006_tsv(pszOutputPath, [arrOutputHeader, arrOutputFields])
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0006_DISCARD_SUMMARY_LOWER_CLEANUP_FAILED",
                "lower-section combination failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def write_step0007_discard_summary_tsv(
    pszOutputPath: str,
    arrRows: List[List[str]],
) -> None:
    if not arrRows:
        raise_conversion_error(
            "INVALID_STEP0007_DISCARD_SUMMARY_ROW_COUNT",
            "step0007 discard summary did not contain any rows.",
        )
    try:
        write_monthly_tsv(
            pszOutputPath,
            "\t".join(arrRows[0]),
            ["\t".join(arrRow) for arrRow in arrRows[1:]],
        )
    except ConversionError as objException:
        dicErrorCodes = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0007_DISCARD_SUMMARY_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0007_DISCARD_SUMMARY_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0007_DISCARD_SUMMARY_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0007_DISCARD_SUMMARY_WRITE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(
                objException.pszErrorCode, "STEP0007_DISCARD_SUMMARY_WRITE_FAILED"
            ),
            str(objException),
        )


def create_step0007_discard_summary_output_path(
    pszInputPath: str,
    pszSection: str,
) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0006_総計_廃棄データ一覧表_(上段|下段)_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None or objMatch.group(2) != pszSection:
        raise_conversion_error(
            "INVALID_STEP0006_DISCARD_SUMMARY_FILE_NAME",
            "invalid step0006 discard summary file name. Path = " + pszInputPath,
        )
    pszStem = (
        f"{objMatch.group(1)}_step0007_総計_廃棄データ一覧表_{pszSection}_"
        + f"{objMatch.group(3)}_{objMatch.group(4)}"
        + objMatch.group(5)
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def create_step0007_discard_summary_files(
    arrInputPaths: List[str],
    pszSection: str,
    arrExpectedHeader: List[str],
    bWeight: bool,
) -> Tuple[List[str], List[Tuple[int, int]]]:
    if pszSection not in {"上段", "下段"}:
        raise_conversion_error(
            "INVALID_STEP0007_DISCARD_SUMMARY_SECTION",
            "section must be 上段 or 下段.",
        )
    dicInputs: Dict[Tuple[str, str, int, int, str], str] = {}
    for pszInputPath in arrInputPaths:
        objMatch = re.fullmatch(
            r"(.+)_step0006_総計_廃棄データ一覧表_(上段|下段)_"
            r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
            os.path.basename(pszInputPath),
        )
        if objMatch is None or objMatch.group(2) != pszSection:
            raise_conversion_error(
                "INVALID_STEP0006_DISCARD_SUMMARY_FILE_NAME",
                "invalid step0006 discard summary file name. Path = " + pszInputPath,
            )
        iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
        objKey = (
            os.path.dirname(os.path.abspath(pszInputPath)), objMatch.group(1),
            iYear, iMonth, objMatch.group(5),
        )
        if objKey in dicInputs:
            raise_conversion_error(
                "DUPLICATE_STEP0006_DISCARD_SUMMARY_FILE",
                "duplicate step0006 discard summary file. Path = " + pszInputPath,
            )
        dicInputs[objKey] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            pszInputPath = dicInputs[objKey]
            try:
                with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
                    arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
            except Exception as objException:
                raise_conversion_error(
                    "STEP0006_DISCARD_SUMMARY_READ_FAILED",
                    "failed to read step0006 discard summary. Path = " + pszInputPath
                    + "; Detail = " + str(objException),
                )
            if len(arrLines) != 2:
                raise_conversion_error(
                    "INVALID_STEP0006_DISCARD_SUMMARY_ROW_COUNT",
                    "step0006 discard summary must contain one header and one data row. "
                    + "Path = " + pszInputPath,
                )
            arrHeader = arrLines[0].split("\t")
            arrFields = arrLines[1].split("\t")
            if arrHeader != arrExpectedHeader or len(arrFields) != len(arrExpectedHeader):
                raise_conversion_error(
                    "INVALID_STEP0006_DISCARD_SUMMARY_COLUMNS",
                    "unexpected step0006 discard summary columns. Path = " + pszInputPath,
                )
            if parse_year_month(arrFields[0], arrFields[1], 2) != (
                objKey[2], objKey[3]
            ):
                raise_conversion_error(
                    "STEP0006_DISCARD_SUMMARY_YEAR_MONTH_MISMATCH",
                    "file name and row year/month do not match. Path = " + pszInputPath,
                )
            if arrFields[2:4] != ["", "総計"]:
                raise_conversion_error(
                    "INVALID_STEP0006_DISCARD_SUMMARY_GROUP",
                    "expected an empty group number and total group name. Path = "
                    + pszInputPath,
                )
            iFirstSubtotal = arrHeader.index("小計")
            iLastSubtotal = len(arrHeader) - 1 - arrHeader[::-1].index("小計")
            objSubtotalValues: List[Decimal] = []
            for iColumnIndex in (iFirstSubtotal, iLastSubtotal):
                objValue = parse_decimal_value(
                    arrFields[iColumnIndex].replace(",", ""),
                    2,
                    iColumnIndex,
                    "小計",
                )
                if objValue is None:
                    raise_conversion_error(
                        "INVALID_STEP0006_DISCARD_SUMMARY_SUBTOTAL",
                        "subtotal value was blank. Path = " + pszInputPath,
                    )
                objSubtotalValues.append(objValue)
            objTotal = sum(objSubtotalValues, Decimal("0"))
            pszTotal = (
                format(objTotal.quantize(Decimal("0.1"), rounding=ROUND_HALF_UP), ",.1f")
                if bWeight
                else format(Decimal(format_sum_value(objTotal)), ",f")
            )
            pszOutputPath = create_step0007_discard_summary_output_path(
                pszInputPath, pszSection
            )
            write_step0007_discard_summary_tsv(
                pszOutputPath, [arrHeader + ["合計"], arrFields + [pszTotal]]
            )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0007_DISCARD_SUMMARY_CLEANUP_FAILED",
                "step0007 discard summary creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0007_discard_summary_upper_files(
    arrStep0006UpperPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    arrExpectedHeader = (
        ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数", "平均体重"]
        + STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS + ["小計"]
        + STEP0004_TOTAL_COUNT_ITEMS + ["小計"]
    )
    return create_step0007_discard_summary_files(
        arrStep0006UpperPaths, "上段", arrExpectedHeader, False
    )


def create_step0007_discard_summary_lower_files(
    arrStep0006LowerPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    arrExpectedHeader = (
        ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "出荷羽数", "出荷重量"]
        + STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS + ["小計"]
        + STEP0004_TOTAL_WEIGHT_ITEMS + ["小計"]
    )
    return create_step0007_discard_summary_files(
        arrStep0006LowerPaths, "下段", arrExpectedHeader, True
    )


def create_step0008_discard_summary_output_path(pszUpperInputPath: str) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0007_総計_廃棄データ一覧表_上段_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv",
        os.path.basename(pszUpperInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0007_DISCARD_SUMMARY_UPPER_FILE_NAME",
            "invalid step0007 upper-section file name. Path = " + pszUpperInputPath,
        )
    pszStem = (
        f"{objMatch.group(1)}_step0008_総計_廃棄データ一覧表_"
        f"{objMatch.group(2)}_{objMatch.group(3)}{objMatch.group(4)}"
    )
    pszDirectory = os.path.dirname(pszUpperInputPath)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def normalize_step0008_discard_summary_headers(
    arrUpperHeader: List[str], arrLowerHeader: List[str]
) -> Tuple[List[str], List[str]]:
    arrUpperExpected = (
        ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数", "平均体重"]
        + STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS + ["小計"]
        + STEP0004_TOTAL_COUNT_ITEMS + ["小計", "合計"]
    )
    arrLowerExpected = (
        ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "出荷羽数", "出荷重量"]
        + STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS + ["小計"]
        + STEP0004_TOTAL_WEIGHT_ITEMS + ["小計", "合計"]
    )
    if arrUpperHeader != arrUpperExpected or arrLowerHeader != arrLowerExpected:
        raise_conversion_error(
            "INVALID_STEP0007_DISCARD_SUMMARY_COLUMNS",
            "step0007 upper or lower summary has unexpected columns.",
        )

    arrUpperOutput = list(arrUpperHeader)
    arrUpperOutput[:4] = ["", "", "", ""]
    arrUpperOutput[6] = "懸鳥前"
    arrUpperOutput[7] = "懸鳥前"
    iUpperDiscardStart = (
        4 + 2 + len(STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS) + 1
    )
    for iColumn in range(
        iUpperDiscardStart,
        iUpperDiscardStart + len(STEP0004_TOTAL_COUNT_ITEMS),
    ):
        arrUpperOutput[iColumn] = ""
    arrUpperOutput[iUpperDiscardStart] = "自主廃棄＋検査廃棄"

    arrLowerOutput = list(arrLowerHeader)
    arrLowerOutput[6] = "死鶏"
    arrLowerOutput[7] = "発育不良"
    iLowerDiscardStart = (
        4 + 2 + len(STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS) + 1
    )
    for iColumn in range(
        iLowerDiscardStart,
        iLowerDiscardStart + len(STEP0004_TOTAL_WEIGHT_ITEMS),
    ):
        pszHeader = arrLowerOutput[iColumn]
        if pszHeader.endswith("羽数"):
            arrLowerOutput[iColumn] = pszHeader[:-2]
        elif pszHeader.endswith("重量"):
            arrLowerOutput[iColumn] = pszHeader[:-2]
    return arrUpperOutput, arrLowerOutput


def create_step0008_discard_summary_files(
    arrStep0007UpperPaths: List[str], arrStep0007LowerPaths: List[str]
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """step0007の上段・下段を結合し、step0008廃棄一覧表を作成します。"""
    objPattern = re.compile(
        r"(.+)_step0007_総計_廃棄データ一覧表_(上段|下段)_"
        r"(\d{4})_(\d{2})((?:_\d+)*)\.tsv"
    )
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszSection, arrPaths in (
        ("上段", arrStep0007UpperPaths), ("下段", arrStep0007LowerPaths)
    ):
        for pszInputPath in arrPaths:
            objMatch = objPattern.fullmatch(os.path.basename(pszInputPath))
            if objMatch is None or objMatch.group(2) != pszSection:
                raise_conversion_error(
                    "INVALID_STEP0007_DISCARD_SUMMARY_FILE_NAME",
                    "invalid step0007 summary file name. Path = " + pszInputPath,
                )
            iYear, iMonth = parse_year_month(objMatch.group(3), objMatch.group(4), 1)
            objKey = (
                os.path.dirname(os.path.abspath(pszInputPath)), objMatch.group(1),
                iYear, iMonth, objMatch.group(5),
            )
            dicPair = dicInputs.setdefault(objKey, {})
            if pszSection in dicPair:
                raise_conversion_error(
                    "DUPLICATE_STEP0007_DISCARD_SUMMARY_FILE",
                    "duplicate step0007 summary file. Path = " + pszInputPath,
                )
            dicPair[pszSection] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda key: (key[2], key[3], key[0], key[1], key[4])):
            dicPair = dicInputs[objKey]
            if set(dicPair) != {"上段", "下段"}:
                raise_conversion_error(
                    "MISSING_STEP0007_DISCARD_SUMMARY_PAIR",
                    "matching upper and lower step0007 files are required. Key = "
                    + str(objKey),
                )

            arrLoaded: Dict[str, List[List[str]]] = {}
            for pszSection in ("上段", "下段"):
                pszInputPath = dicPair[pszSection]
                try:
                    with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
                        arrLines = [strip_line_ending(line) for line in objFile]
                except Exception as objException:
                    raise_conversion_error(
                        "STEP0007_DISCARD_SUMMARY_READ_FAILED",
                        "failed to read step0007 summary. Path = "
                        + pszInputPath + "; Detail = " + str(objException),
                    )
                if len(arrLines) != 2:
                    raise_conversion_error(
                        "INVALID_STEP0007_DISCARD_SUMMARY_ROW_COUNT",
                        "step0007 summary must contain a header and one data row. Path = "
                        + pszInputPath,
                    )
                arrLoaded[pszSection] = [line.split("\t") for line in arrLines]

            arrUpperHeader, arrUpperValues = arrLoaded["上段"]
            arrLowerHeader, arrLowerValues = arrLoaded["下段"]
            arrUpperHeader, arrLowerHeader = normalize_step0008_discard_summary_headers(
                arrUpperHeader, arrLowerHeader
            )
            if (
                len(arrUpperValues) != len(arrUpperHeader)
                or len(arrLowerValues) != len(arrLowerHeader)
            ):
                raise_conversion_error(
                    "INVALID_STEP0007_DISCARD_SUMMARY_COLUMNS",
                    "step0007 data row column count does not match its header. Key = "
                    + str(objKey),
                )
            for arrValues in (arrUpperValues, arrLowerValues):
                if parse_year_month(
                    arrValues[0], arrValues[1], 2
                ) != (objKey[2], objKey[3]):
                    raise_conversion_error(
                        "STEP0007_DISCARD_SUMMARY_YEAR_MONTH_MISMATCH",
                        "file name and row year/month do not match. Key = " + str(objKey),
                    )
                if arrValues[2:4] != ["", "総計"]:
                    raise_conversion_error(
                        "INVALID_STEP0007_DISCARD_SUMMARY_GROUP",
                        "expected an empty group number and total group name. Key = "
                        + str(objKey),
                    )
            if arrUpperValues[:4] != arrLowerValues[:4]:
                raise_conversion_error(
                    "STEP0007_DISCARD_SUMMARY_COMMON_COLUMNS_MISMATCH",
                    "upper and lower year/month/group values do not match. Key = "
                    + str(objKey),
                )

            arrOutputRows = [
                arrUpperHeader, arrLowerHeader, arrUpperValues, arrLowerValues
            ]
            pszOutputPath = create_step0008_discard_summary_output_path(dicPair["上段"])
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrOutputRows[0]),
                    ["\t".join(row) for row in arrOutputRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0008_DISCARD_SUMMARY_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0008_DISCARD_SUMMARY_CLEANUP_FAILED",
                "step0008 discard summary creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0009_discard_summary_output_path(pszInputPath: str) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0008_総計_廃棄データ一覧表_(\d{4})_(\d{2})"
        r"((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0008_DISCARD_SUMMARY_FILE_NAME",
            "invalid step0008 discard summary file name. Path = " + pszInputPath,
        )
    pszStem = (
        f"{objMatch.group(1)}_step0009_総計_廃棄データ一覧表_"
        f"{objMatch.group(2)}_{objMatch.group(3)}{objMatch.group(4)}"
    )
    pszDirectory = os.path.dirname(pszInputPath)
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def format_step0009_discard_weight_subtotal(pszValue: str, iColumn: int) -> str:
    objValue = parse_decimal_value(
        pszValue.replace(",", ""), 4, iColumn, "小計・合計"
    )
    if objValue is None:
        raise_conversion_error(
            "INVALID_STEP0008_DISCARD_SUMMARY_TOTAL",
            "weight subtotal or total was blank. Column = " + str(iColumn + 1),
        )
    return format(
        objValue.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP), ",.2f"
    )


def create_step0009_discard_summary_files(
    arrStep0008DiscardSummaryPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """step0008の総計一覧表からグループ列を除き、羽数・重量区分列を加えます。"""
    objPattern = re.compile(
        r"(.+)_step0008_総計_廃棄データ一覧表_(\d{4})_(\d{2})"
        r"((?:_\d+)*)\.tsv"
    )
    arrUpperSourceHeader = (
        ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数", "平均体重"]
        + STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS + ["小計"]
        + STEP0004_TOTAL_COUNT_ITEMS + ["小計", "合計"]
    )
    arrLowerSourceHeader = (
        ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "出荷羽数", "出荷重量"]
        + STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS + ["小計"]
        + STEP0004_TOTAL_WEIGHT_ITEMS + ["小計", "合計"]
    )
    arrUpperHeader, arrLowerHeader = normalize_step0008_discard_summary_headers(
        arrUpperSourceHeader, arrLowerSourceHeader
    )
    arrExpectedRows = [arrUpperHeader, arrLowerHeader]
    dicInputs: Dict[Tuple[str, str, int, int, str], str] = {}
    for pszInputPath in arrStep0008DiscardSummaryPaths:
        objMatch = objPattern.fullmatch(os.path.basename(pszInputPath))
        if objMatch is None:
            raise_conversion_error(
                "INVALID_STEP0008_DISCARD_SUMMARY_FILE_NAME",
                "invalid step0008 discard summary file name. Path = " + pszInputPath,
            )
        iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
        objKey = (
            os.path.dirname(os.path.abspath(pszInputPath)), objMatch.group(1),
            iYear, iMonth, objMatch.group(4),
        )
        if objKey in dicInputs:
            raise_conversion_error(
                "DUPLICATE_STEP0008_DISCARD_SUMMARY_FILE",
                "duplicate step0008 discard summary file. Path = " + pszInputPath,
            )
        dicInputs[objKey] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(
            dicInputs, key=lambda key: (key[2], key[3], key[0], key[1], key[4])
        ):
            pszInputPath = dicInputs[objKey]
            try:
                with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
                    arrLines = [strip_line_ending(line) for line in objFile]
            except Exception as objException:
                raise_conversion_error(
                    "STEP0008_DISCARD_SUMMARY_READ_FAILED",
                    "failed to read step0008 discard summary. Path = "
                    + pszInputPath + "; Detail = " + str(objException),
                )
            if len(arrLines) != 4:
                raise_conversion_error(
                    "INVALID_STEP0008_DISCARD_SUMMARY_ROW_COUNT",
                    "step0008 discard summary must contain four rows. Path = "
                    + pszInputPath,
                )
            arrRows = [line.split("\t") for line in arrLines]
            if arrRows[:2] != arrExpectedRows:
                raise_conversion_error(
                    "INVALID_STEP0008_DISCARD_SUMMARY_COLUMNS",
                    "unexpected step0008 discard summary title rows. Path = "
                    + pszInputPath,
                )
            if (
                len(arrRows[0]) != len(arrUpperHeader)
                or len(arrRows[1]) != len(arrLowerHeader)
                or len(arrRows[2]) != len(arrUpperHeader)
                or len(arrRows[3]) != len(arrLowerHeader)
            ):
                raise_conversion_error(
                    "INVALID_STEP0008_DISCARD_SUMMARY_COLUMNS",
                    "step0008 rows must match upper/lower header widths. Path = "
                    + pszInputPath,
                )

            for iRow in (2, 3):
                arrValues = arrRows[iRow]
                if parse_year_month(arrValues[0], arrValues[1], iRow + 1) != (
                    objKey[2], objKey[3]
                ):
                    raise_conversion_error(
                        "STEP0008_DISCARD_SUMMARY_YEAR_MONTH_MISMATCH",
                        "file name and row year/month do not match. Path = "
                        + pszInputPath,
                    )
                if arrValues[2:4] != ["", "総計"]:
                    raise_conversion_error(
                        "INVALID_STEP0008_DISCARD_SUMMARY_GROUP",
                        "expected an empty group number and total group name. Path = "
                        + pszInputPath,
                    )
            if arrRows[2][:4] != arrRows[3][:4]:
                raise_conversion_error(
                    "STEP0008_DISCARD_SUMMARY_COMMON_COLUMNS_MISMATCH",
                    "upper and lower year/month/group values do not match. Path = "
                    + pszInputPath,
                )

            arrOutputRows: List[List[str]] = []
            for iRow, arrSourceRow in enumerate(arrRows):
                arrOutputRow = arrSourceRow[:2] + arrSourceRow[4:]
                arrOutputRow.insert(4, "")
                if iRow == 2:
                    arrOutputRow[4] = "羽数"
                elif iRow == 3:
                    arrOutputRow[4] = "重量"
                    for iColumn, pszHeader in enumerate(arrOutputRows[1]):
                        if pszHeader in {"小計", "合計"}:
                            arrOutputRow[iColumn] = format_step0009_discard_weight_subtotal(
                                arrOutputRow[iColumn], iColumn
                            )
                arrOutputRows.append(arrOutputRow)

            for iRow in (2, 3):
                objYearMonth = parse_year_month(
                    arrOutputRows[iRow][0], arrOutputRows[iRow][1], iRow + 1
                )
                arrOutputRows[iRow][0] = str(objYearMonth[0])
                arrOutputRows[iRow][1] = str(objYearMonth[1])

            pszOutputPath = create_step0009_discard_summary_output_path(pszInputPath)
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrOutputRows[0]),
                    ["\t".join(row) for row in arrOutputRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0009_DISCARD_SUMMARY_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0009_DISCARD_SUMMARY_CLEANUP_FAILED",
                "step0009 discard summary creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0010_discard_summary_output_path(
    pszDirectory: str, pszBaseName: str
) -> str:
    pszStem = f"{pszBaseName}_step0010_総計_廃棄データ一覧表"
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def create_step0010_discard_summary_files(
    arrStep0009DiscardSummaryPaths: List[str],
) -> List[str]:
    """step0009月別総計一覧表を新しい順に1つのstep0010 TSVへ連結します。"""
    objPattern = re.compile(
        r"(.+)_step0009_総計_廃棄データ一覧表_(\d{4})_(\d{2})"
        r"((?:_\d+)*)\.tsv"
    )
    arrUpperSourceHeader = (
        ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "入雛羽数", "平均体重"]
        + STEP0004_TOTAL_PRE_SLAUGHTER_COUNT_ITEMS + ["小計"]
        + STEP0004_TOTAL_COUNT_ITEMS + ["小計", "合計"]
    )
    arrLowerSourceHeader = (
        ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", "出荷羽数", "出荷重量"]
        + STEP0004_TOTAL_PRE_SLAUGHTER_WEIGHT_ITEMS + ["小計"]
        + STEP0004_TOTAL_WEIGHT_ITEMS + ["小計", "合計"]
    )
    arrUpperStep8Header, arrLowerStep8Header = (
        normalize_step0008_discard_summary_headers(
            arrUpperSourceHeader, arrLowerSourceHeader
        )
    )
    arrExpectedHeaders: List[List[str]] = []
    for arrHeader in (arrUpperStep8Header, arrLowerStep8Header):
        arrExpectedHeader = arrHeader[:2] + arrHeader[4:]
        arrExpectedHeader.insert(4, "")
        arrExpectedHeaders.append(arrExpectedHeader)

    dicInputs: Dict[Tuple[str, str], Dict[Tuple[int, int], str]] = {}
    for pszInputPath in arrStep0009DiscardSummaryPaths:
        objMatch = objPattern.fullmatch(os.path.basename(pszInputPath))
        if objMatch is None:
            raise_conversion_error(
                "INVALID_STEP0009_DISCARD_SUMMARY_FILE_NAME",
                "invalid step0009 discard summary file name. Path = " + pszInputPath,
            )
        iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
        objKey = (os.path.dirname(os.path.abspath(pszInputPath)), objMatch.group(1))
        objYearMonth = (iYear, iMonth)
        dicByMonth = dicInputs.setdefault(objKey, {})
        if objYearMonth in dicByMonth:
            raise_conversion_error(
                "DUPLICATE_STEP0009_DISCARD_SUMMARY_MONTH",
                "multiple step0009 discard summaries have the same year/month. Key = "
                + str(objKey) + ", YearMonth = " + str(objYearMonth),
            )
        dicByMonth[objYearMonth] = pszInputPath

    arrCreatedPaths: List[str] = []
    try:
        for objKey in sorted(dicInputs):
            dicByMonth = dicInputs[objKey]
            arrRowsByMonth: Dict[Tuple[int, int], List[List[str]]] = {}
            arrTitleRows: Optional[List[List[str]]] = None
            for objYearMonth in sorted(dicByMonth, reverse=True):
                pszInputPath = dicByMonth[objYearMonth]
                try:
                    with open(
                        pszInputPath, "r", encoding="utf-8", newline=""
                    ) as objFile:
                        arrLines = [strip_line_ending(line) for line in objFile]
                except Exception as objException:
                    raise_conversion_error(
                        "STEP0009_DISCARD_SUMMARY_READ_FAILED",
                        "failed to read step0009 discard summary. Path = "
                        + pszInputPath + "; Detail = " + str(objException),
                    )
                if len(arrLines) != 4:
                    raise_conversion_error(
                        "INVALID_STEP0009_DISCARD_SUMMARY_ROW_COUNT",
                        "step0009 discard summary must contain four rows. Path = "
                        + pszInputPath,
                    )
                arrRows = [line.split("\t") for line in arrLines]
                if arrRows[:2] != arrExpectedHeaders:
                    raise_conversion_error(
                        "INVALID_STEP0009_DISCARD_SUMMARY_COLUMNS",
                        "unexpected step0009 discard summary title rows. Path = "
                        + pszInputPath,
                    )
                if (
                    len(arrRows[0]) != len(arrExpectedHeaders[0])
                    or len(arrRows[1]) != len(arrExpectedHeaders[1])
                    or len(arrRows[2]) != len(arrExpectedHeaders[0])
                    or len(arrRows[3]) != len(arrExpectedHeaders[1])
                ):
                    raise_conversion_error(
                        "INVALID_STEP0009_DISCARD_SUMMARY_COLUMNS",
                        "step0009 rows must match their upper/lower header widths. Path = "
                        + pszInputPath,
                    )
                for iRow in (2, 3):
                    if parse_year_month(
                        arrRows[iRow][0], arrRows[iRow][1], iRow + 1
                    ) != objYearMonth:
                        raise_conversion_error(
                            "STEP0009_DISCARD_SUMMARY_YEAR_MONTH_MISMATCH",
                            "file name and row year/month do not match. Path = "
                            + pszInputPath,
                        )
                if arrRows[2][4] != "羽数" or arrRows[3][4] != "重量":
                    raise_conversion_error(
                        "INVALID_STEP0009_DISCARD_SUMMARY_KIND",
                        "expected 羽数 and 重量 data rows. Path = " + pszInputPath,
                    )
                if arrRows[2][:2] != arrRows[3][:2]:
                    raise_conversion_error(
                        "STEP0009_DISCARD_SUMMARY_COMMON_COLUMNS_MISMATCH",
                        "the two data rows must have matching year/month. Path = "
                        + pszInputPath,
                    )
                if arrTitleRows is None:
                    arrTitleRows = arrRows[:2]
                elif arrTitleRows != arrRows[:2]:
                    raise_conversion_error(
                        "STEP0009_DISCARD_SUMMARY_TITLE_MISMATCH",
                        "title rows differ between monthly files. Path = "
                        + pszInputPath,
                    )

                arrLowerValues = list(arrRows[3])
                arrLowerValues[0] = ""
                arrLowerValues[1] = ""
                arrRowsByMonth[objYearMonth] = [arrRows[2], arrLowerValues]

            if arrTitleRows is None:
                continue
            arrOutputRows = [list(arrRow) for arrRow in arrTitleRows]
            for iColumn, pszExpectedTitle in (
                (7, "小計"), (23, "小計"), (24, "合計")
            ):
                if arrOutputRows[0][iColumn] != pszExpectedTitle:
                    raise_conversion_error(
                        "INVALID_STEP0009_DISCARD_SUMMARY_COLUMNS",
                        "unexpected step0010 upper title at column "
                        + str(iColumn + 1) + ".",
                    )
                arrOutputRows[0][iColumn] = ""
            for objYearMonth in sorted(arrRowsByMonth, reverse=True):
                arrOutputRows.extend(arrRowsByMonth[objYearMonth])

            pszOutputPath = create_step0010_discard_summary_output_path(
                objKey[0], objKey[1]
            )
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrOutputRows[0]),
                    ["\t".join(row) for row in arrOutputRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0010_DISCARD_SUMMARY_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0010_DISCARD_SUMMARY_CLEANUP_FAILED",
                "step0010 discard summary creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0011_discard_summary_output_path(pszInputPath: str) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0010_総計_廃棄データ一覧表((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0010_DISCARD_SUMMARY_FILE_NAME",
            "invalid step0010 discard summary file name. Path = " + pszInputPath,
        )
    pszDirectory = os.path.dirname(pszInputPath)
    pszBaseName, pszSuffix = objMatch.groups()
    pszOutputStem = os.path.join(
        pszDirectory,
        f"{pszBaseName}_step0011_総計_廃棄データ一覧表",
    )
    pszOutputPath = pszOutputStem + pszSuffix + ".tsv"
    if os.path.exists(pszOutputPath):
        arrSuffixParts = [part for part in pszSuffix.split("_") if part]
        iSuffix = int(arrSuffixParts[-1]) + 1 if arrSuffixParts else 2
        while True:
            pszOutputPath = f"{pszOutputStem}_{iSuffix:02d}.tsv"
            if not os.path.exists(pszOutputPath):
                break
            iSuffix += 1
    return pszOutputPath


def create_step0011_discard_summary_files(
    arrStep0010DiscardSummaryPaths: List[str],
) -> List[str]:
    """step0010総計一覧表の右端に前年同月比(羽数)見出し列を追加します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszInputPath in arrStep0010DiscardSummaryPaths:
            try:
                with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
                    arrRows = [
                        strip_line_ending(pszLine).split("\t")
                        for pszLine in objFile
                    ]
            except Exception as objException:
                raise_conversion_error(
                    "STEP0010_DISCARD_SUMMARY_READ_FAILED",
                    "failed to read step0010 discard summary. Path = "
                    + pszInputPath + "; Detail = " + str(objException),
                )
            if len(arrRows) < 2 or any(
                len(arrRow) < 25 or len(arrRow) > 26
                or (len(arrRow) == 26 and arrRow[25] != "")
                for arrRow in arrRows
            ):
                raise_conversion_error(
                    "INVALID_STEP0010_DISCARD_SUMMARY_COLUMNS",
                    "step0010 discard summary rows must have 25 columns and may "
                    "have only a blank trailing Z cell. Path = " + pszInputPath,
                )

            arrOutputRows = [
                list(arrRow) + [""] * (26 - len(arrRow)) for arrRow in arrRows
            ]
            arrOutputRows[0][25] = ""
            arrOutputRows[1][25] = "前年同月比(羽数)"
            for arrRow in arrOutputRows[2:]:
                arrRow[25] = ""

            pszOutputPath = create_step0011_discard_summary_output_path(pszInputPath)
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrOutputRows[0]),
                    ["\t".join(arrRow) for arrRow in arrOutputRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0011_DISCARD_SUMMARY_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0011_DISCARD_SUMMARY_CLEANUP_FAILED",
                "step0011 discard summary creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0012_discard_summary_output_path(pszInputPath: str) -> str:
    objMatch = re.fullmatch(
        r"(.+)_step0011_総計_廃棄データ一覧表((?:_\d+)*)\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0011_DISCARD_SUMMARY_FILE_NAME",
            "invalid step0011 discard summary file name. Path = " + pszInputPath,
        )
    pszDirectory = os.path.dirname(pszInputPath)
    pszBaseName, pszSuffix = objMatch.groups()
    pszOutputStem = os.path.join(
        pszDirectory,
        f"{pszBaseName}_step0012_総計_廃棄データ一覧表",
    )
    pszOutputPath = pszOutputStem + pszSuffix + ".tsv"
    if os.path.exists(pszOutputPath):
        arrSuffixParts = [part for part in pszSuffix.split("_") if part]
        iSuffix = int(arrSuffixParts[-1]) + 1 if arrSuffixParts else 2
        while True:
            pszOutputPath = f"{pszOutputStem}_{iSuffix:02d}.tsv"
            if not os.path.exists(pszOutputPath):
                break
            iSuffix += 1
    return pszOutputPath


def create_step0012_discard_summary_files(
    arrStep0011DiscardSummaryPaths: List[str],
) -> List[str]:
    """前年同月比(羽数)をZ列に追加し、step0012総計一覧表を出力します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszInputPath in arrStep0011DiscardSummaryPaths:
            try:
                with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
                    arrRows = [
                        strip_line_ending(pszLine).split("\t")
                        for pszLine in objFile
                    ]
            except Exception as objException:
                raise_conversion_error(
                    "STEP0011_DISCARD_SUMMARY_READ_FAILED",
                    "failed to read step0011 discard summary. Path = "
                    + pszInputPath + "; Detail = " + str(objException),
                )
            if (
                len(arrRows) < 2
                or any(len(arrRow) != 26 for arrRow in arrRows)
                or arrRows[1][25] != "前年同月比(羽数)"
            ):
                raise_conversion_error(
                    "INVALID_STEP0011_DISCARD_SUMMARY_COLUMNS",
                    "step0011 discard summary must have 26 columns and the "
                    "前年同月比(羽数) header in Z2. Path = " + pszInputPath,
                )

            dicCountRows: Dict[Tuple[int, int], int] = {}
            arrYearMonthRows: List[Tuple[int, int, int]] = []
            for iRowIndex, arrRow in enumerate(arrRows[2:], start=3):
                pszYear = arrRow[0].strip(" \u3000\t")
                pszMonth = arrRow[1].strip(" \u3000\t")
                if pszYear == "" and pszMonth == "":
                    continue
                if pszYear == "" or pszMonth == "":
                    raise_conversion_error(
                        "INVALID_STEP0011_DISCARD_SUMMARY_YEAR_MONTH",
                        "step0011 count row must have both year and month. Line = "
                        + str(iRowIndex),
                    )
                objYearMonth = parse_year_month(pszYear, pszMonth, iRowIndex)
                if objYearMonth in dicCountRows:
                    raise_conversion_error(
                        "DUPLICATE_STEP0011_DISCARD_SUMMARY_YEAR_MONTH",
                        "duplicate year/month in step0011 discard summary. Line = "
                        + str(iRowIndex),
                    )
                dicCountRows[objYearMonth] = iRowIndex - 1
                arrYearMonthRows.append((iRowIndex - 1, *objYearMonth))

            for iRowIndex, iYear, iMonth in arrYearMonthRows:
                iPreviousRowIndex = dicCountRows.get((iYear - 1, iMonth))
                if iPreviousRowIndex is None:
                    arrRows[iRowIndex][25] = "－"
                    continue
                objCurrentTotal = parse_decimal_value(
                    arrRows[iRowIndex][24].replace(",", ""),
                    iRowIndex + 1,
                    24,
                    "合計",
                )
                objPreviousTotal = parse_decimal_value(
                    arrRows[iPreviousRowIndex][24].replace(",", ""),
                    iPreviousRowIndex + 1,
                    24,
                    "合計",
                )
                if (
                    objCurrentTotal is None
                    or objPreviousTotal is None
                    or objPreviousTotal == 0
                ):
                    arrRows[iRowIndex][25] = "－"
                    continue
                objRate = (objCurrentTotal / objPreviousTotal * Decimal("100")).quantize(
                    Decimal("0.01"), rounding=ROUND_HALF_UP
                )
                arrRows[iRowIndex][25] = format(objRate, ".2f") + "%"

            pszOutputPath = create_step0012_discard_summary_output_path(pszInputPath)
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrRows[0]),
                    ["\t".join(arrRow) for arrRow in arrRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0012_DISCARD_SUMMARY_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0012_DISCARD_SUMMARY_CLEANUP_FAILED",
                "step0012 discard summary creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0012_discard_summary_excel_files(
    arrStep0012DiscardSummaryPaths: List[str],
) -> List[str]:
    """step0012総計廃棄データ一覧表TSVから同名のExcelファイルを作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszTsvPath in arrStep0012DiscardSummaryPaths:
            try:
                with open(pszTsvPath, "r", encoding="utf-8", newline="") as objFile:
                    arrRows = [
                        strip_line_ending(pszLine).split("\t")
                        for pszLine in objFile
                    ]
            except Exception as objException:
                raise_conversion_error(
                    "STEP0012_DISCARD_SUMMARY_EXCEL_READ_FAILED",
                    "failed to read step0012 discard summary TSV. Path = "
                    + pszTsvPath + "; Detail = " + str(objException),
                )
            if (
                len(arrRows) < 3
                or any(len(arrRow) != 26 for arrRow in arrRows)
                or arrRows[1][25].strip(" \u3000\t") != "前年同月比(羽数)"
            ):
                raise_conversion_error(
                    "INVALID_STEP0012_DISCARD_SUMMARY_EXCEL_INPUT",
                    "step0012 discard summary TSV must contain 26 columns "
                    "and the expected second header row. Path = " + pszTsvPath,
                )

            pszOutputPath = os.path.splitext(pszTsvPath)[0] + ".xlsx"
            pszTemporaryPath = ""
            bOutputReserved = False
            objWorkbook: Optional[Workbook] = None
            try:
                try:
                    iFileDescriptor = os.open(
                        pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
                    )
                    os.close(iFileDescriptor)
                    bOutputReserved = True
                except FileExistsError:
                    raise_conversion_error(
                        "STEP0012_DISCARD_SUMMARY_EXCEL_CREATE_FAILED",
                        "step0012 discard summary Excel already exists. Path = "
                        + pszOutputPath,
                    )
                iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                    prefix="group_shipment_summary_step0012_discard_",
                    suffix=".xlsx",
                    dir=os.path.dirname(pszOutputPath),
                )
                os.close(iTemporaryFileDescriptor)
                objWorkbook = Workbook()
                objWorksheet = objWorkbook.active
                objWorksheet.title = "総計"
                for iRowIndex, arrFields in enumerate(arrRows, start=1):
                    pszUnit = arrFields[4].strip(" \u3000\t") if iRowIndex >= 3 else ""
                    for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                        objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                        pszNormalizedValue = pszValue.strip(" \u3000\t")
                        if pszNormalizedValue == "":
                            objCell.value = None
                            continue
                        if iRowIndex <= 2 or pszNormalizedValue == "－" or "%" in pszNormalizedValue:
                            objCell.value = pszNormalizedValue
                            continue
                        try:
                            objValue = Decimal(pszNormalizedValue.replace(",", ""))
                        except InvalidOperation:
                            objCell.value = pszNormalizedValue
                            continue
                        objCell.value = (
                            int(objValue)
                            if objValue == objValue.to_integral_value()
                            else float(objValue)
                        )
                        if iColumnIndex == 1:
                            objCell.number_format = "0"
                        elif iColumnIndex == 2:
                            objCell.number_format = "00"
                        elif iColumnIndex == 4 and pszUnit == "羽数":
                            objCell.number_format = "0.00"
                        elif iColumnIndex == 3 or pszUnit == "羽数":
                            objCell.number_format = "#,##0"
                        elif pszUnit == "重量" or iColumnIndex in {4, 25}:
                            iDecimalPlaces = max(
                                0, -objValue.as_tuple().exponent
                            )
                            objCell.number_format = (
                                "#,##0." + "0" * iDecimalPlaces
                                if iDecimalPlaces else "#,##0"
                            )
                        else:
                            objCell.number_format = "0.#####"
                for iColumnIndex in range(1, objWorksheet.max_column + 1):
                    iMaximumDisplayWidth = 0
                    for objCell in objWorksheet.iter_cols(
                        min_col=iColumnIndex,
                        max_col=iColumnIndex,
                        min_row=1,
                        max_row=objWorksheet.max_row,
                    ):
                        for objWidthCell in objCell:
                            objValue = objWidthCell.value
                            if objValue is None:
                                continue
                            if isinstance(objValue, (int, float, Decimal)):
                                objDecimalValue = Decimal(str(objValue))
                                pszFormat = objWidthCell.number_format.split(";")[0]
                                pszIntegerFormat, _, pszFractionFormat = pszFormat.partition(".")
                                iFractionPlaces = sum(
                                    1 for pszCharacter in pszFractionFormat
                                    if pszCharacter in "0#"
                                )
                                pszDisplayValue = format(
                                    objDecimalValue,
                                    ",." + str(iFractionPlaces) + "f"
                                    if "," in pszIntegerFormat
                                    else "." + str(iFractionPlaces) + "f",
                                )
                                if iFractionPlaces:
                                    iMinimumFractionPlaces = pszFractionFormat.count("0")
                                    while (
                                        len(pszDisplayValue.rsplit(".", 1)[-1])
                                        > iMinimumFractionPlaces
                                        and pszDisplayValue.endswith("0")
                                    ):
                                        pszDisplayValue = pszDisplayValue[:-1]
                                    if pszDisplayValue.endswith("."):
                                        pszDisplayValue = pszDisplayValue[:-1]
                            else:
                                pszDisplayValue = str(objValue)
                            iDisplayWidth = sum(
                                2 if unicodedata.east_asian_width(pszCharacter)
                                in {"F", "W"}
                                else 1
                                for pszCharacter in pszDisplayValue
                            )
                            iMaximumDisplayWidth = max(
                                iMaximumDisplayWidth, iDisplayWidth
                            )
                    objWorksheet.column_dimensions[
                        get_column_letter(iColumnIndex)
                    ].width = min(iMaximumDisplayWidth + 2, 255)
                objWorkbook.save(pszTemporaryPath)
                objWorkbook.close()
                objWorkbook = None
                os.replace(pszTemporaryPath, pszOutputPath)
                pszTemporaryPath = ""
                bOutputReserved = False
            except ConversionError:
                raise
            except Exception as objException:
                raise_conversion_error(
                    "STEP0012_DISCARD_SUMMARY_EXCEL_WRITE_FAILED",
                    "failed to create step0012 discard summary Excel. Path = "
                    + pszOutputPath + "; Detail = " + str(objException),
                )
            finally:
                if objWorkbook is not None:
                    objWorkbook.close()
                if pszTemporaryPath and os.path.exists(pszTemporaryPath):
                    os.remove(pszTemporaryPath)
                if bOutputReserved and os.path.exists(pszOutputPath):
                    os.remove(pszOutputPath)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0012_DISCARD_SUMMARY_EXCEL_CLEANUP_FAILED",
                "step0012 discard summary Excel creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def apply_group_summary_yearly_borders(
    pszExcelPath: str, bDeduction: bool = False
) -> None:
    """見本と同じ31行のグループ別表に罫線だけを設定し、上書き保存します。"""
    objWorkbook = load_workbook(pszExcelPath)
    pszTemporaryPath = ""
    try:
        objWorksheet = objWorkbook["グループ別"]
        if objWorksheet.max_row != 31:
            return
        objThin = Side(style="thin", color=Color(indexed=64))
        objMedium = Side(style="medium", color=Color(indexed=64))
        objEmpty = Side()
        for iRow in range(1, 32):
            for iColumn in range(1, 16):
                objWorksheet.cell(iRow, iColumn).border = Border(
                    left=(
                        objMedium if iColumn == 1 else
                        objThin if iColumn >= 3 else objEmpty
                    ),
                    right=objMedium if iColumn == 15 else objEmpty,
                    top=(
                        objMedium if iRow == 1 else
                        objThin if iRow in {3, 8, 13, 18, 24, 26} else objEmpty
                    ),
                    bottom=(
                        objMedium if iRow == 31 else
                        objThin if iRow in {2, 7, 12, 23, 25}
                        or (bDeduction and iRow == 17) else objEmpty
                    ),
                    diagonal=Side(),
                )
        iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
            prefix="group_shipment_summary_group_borders_",
            suffix=".xlsx",
            dir=os.path.dirname(os.path.abspath(pszExcelPath)),
        )
        os.close(iFileDescriptor)
        objWorkbook.save(pszTemporaryPath)
        objWorkbook.close()
        os.replace(pszTemporaryPath, pszExcelPath)
        pszTemporaryPath = ""
    finally:
        objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)


def create_group_summary_yearly_named_copies(arrExcelPaths: List[str]) -> List[str]:
    """作成済みグループ別年間Excelを対象年付きの別名でコピーします。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszExcelPath in arrExcelPaths:
            objMatch = re.fullmatch(
                r"(.+)_step0013_グループ別_([0-9]{4})_年間(?:_[0-9]+)*\.xlsx",
                os.path.basename(pszExcelPath),
            )
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_GROUP_SUMMARY_NAMED_COPY_FILENAME", pszExcelPath
                )
            iYear = int(objMatch.group(2))
            pszDirectory = os.path.dirname(pszExcelPath)
            pszStem = f"〇グループ別集計_{iYear:04d}年"
            pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
            iSuffix = 2
            while os.path.exists(pszOutputPath):
                pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx")
                iSuffix += 1
            shutil.copy2(pszExcelPath, pszOutputPath)
            arrCreatedPaths.append(pszOutputPath)
            apply_group_summary_yearly_borders(pszOutputPath)
        return arrCreatedPaths
    except Exception:
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except OSError:
                pass
        raise


def create_variety_summary_yearly_named_copies(arrExcelPaths: List[str]) -> List[str]:
    """作成済み品種別年間Excelを対象年付きの別名でコピーします。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszExcelPath in arrExcelPaths:
            objMatch = re.fullmatch(
                r"(.+)_step0013_品種別_([0-9]{4})_年間(?:_[0-9]+)*\.xlsx",
                os.path.basename(pszExcelPath),
            )
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_VARIETY_SUMMARY_NAMED_COPY_FILENAME", pszExcelPath
                )
            iYear = int(objMatch.group(2))
            pszDirectory = os.path.dirname(pszExcelPath)
            pszStem = f"〇品種別実績集計_{iYear:04d}年"
            pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
            iSuffix = 2
            while os.path.exists(pszOutputPath):
                pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx")
                iSuffix += 1
            shutil.copy2(pszExcelPath, pszOutputPath)
            arrCreatedPaths.append(pszOutputPath)
            apply_variety_deduction_summary_yearly_borders(pszOutputPath)
        return arrCreatedPaths
    except Exception:
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except OSError:
                pass
        raise


def apply_total_count_history_borders(pszExcelPath: str) -> None:
    """過去5年の羽数表15ブロックに罫線だけを設定し、上書き保存します。"""
    objWorkbook = load_workbook(pszExcelPath)
    pszTemporaryPath = ""
    try:
        objWorksheet = objWorkbook["総計羽数過去5年"]
        objThin = Side(style="thin", color=Color(indexed=64))
        objBorder = Border(
            left=objThin, right=objThin, top=objThin, bottom=objThin,
            diagonal=Side(),
        )
        for iStartRow in range(2, 115, 8):
            for arrCells in objWorksheet.iter_rows(
                min_row=iStartRow, max_row=iStartRow + 5,
                min_col=1, max_col=13,
            ):
                for objCell in arrCells:
                    objCell.border = objBorder
        iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
            prefix="group_shipment_summary_history_borders_",
            suffix=".xlsx",
            dir=os.path.dirname(os.path.abspath(pszExcelPath)),
        )
        os.close(iFileDescriptor)
        objWorkbook.save(pszTemporaryPath)
        objWorkbook.close()
        os.replace(pszTemporaryPath, pszExcelPath)
        pszTemporaryPath = ""
    finally:
        objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)


def create_total_count_history_named_copies(arrExcelPaths: List[str]) -> List[str]:
    """作成済み出荷廃棄羽数過去05年Excelを別名でコピーします。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszExcelPath in arrExcelPaths:
            objMatch = re.fullmatch(
                r"(.+)_step0005_総計_出荷廃棄_羽数_過去05年(?:_[0-9]+)*\.xlsx",
                os.path.basename(pszExcelPath),
            )
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_TOTAL_COUNT_HISTORY_NAMED_COPY_FILENAME", pszExcelPath
                )
            pszDirectory = os.path.dirname(pszExcelPath)
            pszStem = "〇出荷廃棄_羽数_過去05年"
            pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
            iSuffix = 2
            while os.path.exists(pszOutputPath):
                pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx")
                iSuffix += 1
            shutil.copy2(pszExcelPath, pszOutputPath)
            arrCreatedPaths.append(pszOutputPath)
            apply_total_count_history_borders(pszOutputPath)
        return arrCreatedPaths
    except Exception:
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except OSError:
                pass
        raise


def apply_variety_deduction_summary_yearly_borders(pszExcelPath: str) -> None:
    """見本と同じ48行の品種別控除廃棄表に罫線だけを設定します。"""
    objWorkbook = load_workbook(pszExcelPath)
    pszTemporaryPath = ""
    try:
        objWorksheet = objWorkbook["品種別"]
        # 39行構成など、今回の見本と異なる表は変更しません。
        if objWorksheet.max_row != 48:
            return
        objThin = Side(style="thin", color=Color(indexed=64))
        objMedium = Side(style="medium", color=Color(indexed=64))
        objDouble = Side(style="double", color=Color(indexed=64))
        objEmpty = Side()
        arrSeparatorRows = {2, 8, 14, 20, 26, 32, 39}
        for iRow in range(1, 49):
            for iColumn in range(1, 16):
                objWorksheet.cell(iRow, iColumn).border = Border(
                    left=objMedium if iColumn == 1 else objEmpty,
                    right=(
                        objMedium if iColumn == 15 else
                        objDouble if iColumn == 2 else
                        objThin if iColumn >= 3 else objEmpty
                    ),
                    top=(
                        objMedium if iRow in {1, 42} else
                        objThin if iRow - 1 in arrSeparatorRows else objEmpty
                    ),
                    bottom=(
                        objMedium if iRow in {41, 48} else
                        objThin if iRow in arrSeparatorRows else objEmpty
                    ),
                    diagonal=Side(),
                )
        iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
            prefix="group_shipment_summary_variety_deduction_borders_",
            suffix=".xlsx",
            dir=os.path.dirname(os.path.abspath(pszExcelPath)),
        )
        os.close(iFileDescriptor)
        objWorkbook.save(pszTemporaryPath)
        objWorkbook.close()
        os.replace(pszTemporaryPath, pszExcelPath)
        pszTemporaryPath = ""
    finally:
        objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)


def create_variety_deduction_summary_yearly_named_copies(arrExcelPaths: List[str]) -> List[str]:
    """作成済み控除廃棄版品種別年間Excelを対象年付きの別名でコピーします。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszExcelPath in arrExcelPaths:
            objMatch = re.fullmatch(
                r"(.+)_step0013_品種別_控除廃棄_([0-9]{4})_年間(?:_[0-9]+)*\.xlsx",
                os.path.basename(pszExcelPath),
            )
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_VARIETY_DEDUCTION_SUMMARY_NAMED_COPY_FILENAME", pszExcelPath
                )
            iYear = int(objMatch.group(2))
            pszDirectory = os.path.dirname(pszExcelPath)
            pszStem = f"〇品種別実績集計_控除廃棄_{iYear:04d}年"
            pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
            iSuffix = 2
            while os.path.exists(pszOutputPath):
                pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx")
                iSuffix += 1
            shutil.copy2(pszExcelPath, pszOutputPath)
            arrCreatedPaths.append(pszOutputPath)
            apply_variety_deduction_summary_yearly_borders(pszOutputPath)
        return arrCreatedPaths
    except Exception:
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except OSError:
                pass
        raise


def create_group_deduction_summary_yearly_named_copies(arrExcelPaths: List[str]) -> List[str]:
    """作成済み控除廃棄版グループ別年間Excelを対象年付きの別名でコピーします。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszExcelPath in arrExcelPaths:
            objMatch = re.fullmatch(
                r"(.+)_step0017_グループ別_控除廃棄_([0-9]{4})_年間(?:_[0-9]+)*\.xlsx",
                os.path.basename(pszExcelPath),
            )
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_GROUP_DEDUCTION_SUMMARY_NAMED_COPY_FILENAME", pszExcelPath
                )
            iYear = int(objMatch.group(2))
            pszDirectory = os.path.dirname(pszExcelPath)
            pszStem = f"〇グループ別集計_控除廃棄_{iYear:04d}年"
            pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
            iSuffix = 2
            while os.path.exists(pszOutputPath):
                pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx")
                iSuffix += 1
            shutil.copy2(pszExcelPath, pszOutputPath)
            arrCreatedPaths.append(pszOutputPath)
            apply_group_summary_yearly_borders(pszOutputPath, bDeduction=True)
        return arrCreatedPaths
    except Exception:
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except OSError:
                pass
        raise


def apply_discard_summary_monthly_borders(pszExcelPath: str) -> None:
    """年月付き一覧表の罫線だけを設定し、同じExcelへ上書き保存します。"""
    objWorkbook = load_workbook(pszExcelPath)
    pszTemporaryPath = ""
    try:
        objWorksheet = objWorkbook["総計"]
        iLastRow = objWorksheet.max_row
        objThin = Side(style="thin", color=Color(indexed=64))
        objMedium = Side(style="medium", color=Color(indexed=64))
        objEmpty = Side()
        for iRow in range(1, iLastRow + 1):
            for iColumn in range(1, 27):
                objWorksheet.cell(iRow, iColumn).border = Border(
                    left=(
                        objMedium if iColumn == 1 else
                        objThin if iColumn in {3, 5, 9, 24, 25} else objEmpty
                    ),
                    right=(
                        objMedium if iColumn == 26 else
                        objThin if iColumn in {4, 8, 23, 24, 25} else objEmpty
                    ),
                    top=(
                        objMedium if iRow == 1 else
                        objThin if iRow >= 3 and iRow % 2 == 1 else objEmpty
                    ),
                    bottom=(
                        objMedium if iRow == iLastRow else
                        objThin if iRow % 2 == 0 else objEmpty
                    ),
                )
        iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
            prefix="group_shipment_summary_discard_borders_",
            suffix=".xlsx",
            dir=os.path.dirname(os.path.abspath(pszExcelPath)),
        )
        os.close(iFileDescriptor)
        objWorkbook.save(pszTemporaryPath)
        objWorkbook.close()
        os.replace(pszTemporaryPath, pszExcelPath)
        pszTemporaryPath = ""
    finally:
        objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)


def create_discard_summary_monthly_named_copies(
    arrExcelPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszExcelPath in arrExcelPaths:
            objWorkbook = load_workbook(pszExcelPath, read_only=True, data_only=True)
            try:
                objWorksheet = objWorkbook.active
                arrYearMonths: List[Tuple[int, int]] = []
                for arrValues in objWorksheet.iter_rows(
                    min_row=1, min_col=1, max_col=2, values_only=True
                ):
                    objYear, objMonth = arrValues
                    try:
                        iYear = int(objYear)
                        iMonth = int(objMonth)
                    except (TypeError, ValueError, OverflowError):
                        continue
                    if 1900 <= iYear <= 9999 and 1 <= iMonth <= 12:
                        arrYearMonths.append((iYear, iMonth))
            finally:
                objWorkbook.close()
            if not arrYearMonths:
                raise_conversion_error(
                    "STEP0012_DISCARD_NAMED_COPY_DATE_MISSING",
                    "no valid year/month was found in columns A and B. Path = "
                    + pszExcelPath,
                )
            iYear, iMonth = max(arrYearMonths)
            pszDirectory = os.path.dirname(pszExcelPath)
            pszStem = f"〇廃棄データ月別一覧表～{iYear:04d}年{iMonth:02d}月"
            pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
            iSuffix = 2
            while os.path.exists(pszOutputPath):
                pszOutputPath = os.path.join(
                    pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
                )
                iSuffix += 1
            shutil.copy2(pszExcelPath, pszOutputPath)
            arrCreatedPaths.append(pszOutputPath)
            apply_discard_summary_monthly_borders(pszOutputPath)
        return arrCreatedPaths
    except Exception:
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except OSError:
                pass
        raise


STEP0004_TOTAL_PLACEMENT_AVERAGE_ITEMS: List[str] = [
    "入雛羽数", "平均体重",
]


def create_step0004_total_placement_average_output_path(
    pszFirstInputPath: str,
) -> str:
    pszBaseName, _, iYear, iMonth, pszSuffix = parse_step0003_total_file_name(
        pszFirstInputPath
    )
    pszDirectory: str = os.path.dirname(pszFirstInputPath)
    pszStem: str = (
        f"{pszBaseName}_step0004_総計_羽数データ別_入雛羽数・平均体重_{iYear:04d}_{iMonth:02d}"
        + pszSuffix
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def read_step0003_total_placement_average_item(
    pszInputPath: str, pszExpectedItem: str,
) -> Tuple[List[str], List[str]]:
    _, pszItemName, iFileYear, iFileMonth, _ = parse_step0003_total_file_name(
        pszInputPath
    )
    if pszItemName != pszExpectedItem:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ITEM", "unexpected item. Path = " + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0003_TOTAL_READ_FAILED",
            "failed to read step0003 total TSV. Path = " + pszInputPath
            + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ROW_COUNT",
            "step0003 total TSV must contain one header and one data row. Path = "
            + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    if len(arrHeader) != 5 or len(arrFields) != 5:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_COLUMN_COUNT",
            "step0003 total TSV must contain five columns. Path = " + pszInputPath,
        )
    if arrHeader != ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", pszExpectedItem]:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_HEADER",
            "unexpected step0003 total header. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0003_TOTAL_YEAR_MONTH_MISMATCH",
            "file name and data year/month do not match. Path = " + pszInputPath,
        )
    if arrFields[2:4] != ["", "総計"]:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_GROUP",
            "expected an empty group number and total group name. Path = " + pszInputPath,
        )
    if (parse_decimal_value(arrFields[4], 2, 4, pszExpectedItem) is None
            and pszExpectedItem != "平均体重"):
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ITEM_VALUE",
            "placement count was blank. Path = " + pszInputPath,
        )
    return arrHeader, arrFields


def create_step0004_total_placement_average_files(
    arrStep0003TotalPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """同じ年月の入雛羽数と平均体重を横結合し、値と平均体重の空欄を保持します。"""
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszInputPath in arrStep0003TotalPaths:
        pszBaseName, pszItemName, iYear, iMonth, pszSuffix = (
            parse_step0003_total_file_name(pszInputPath)
        )
        objKey = (os.path.dirname(os.path.abspath(pszInputPath)),
                  pszBaseName, iYear, iMonth, pszSuffix)
        dicItems = dicInputs.setdefault(objKey, {})
        if pszItemName not in STEP0004_TOTAL_PLACEMENT_AVERAGE_ITEMS:
            continue
        if pszItemName in dicItems:
            raise_conversion_error(
                "DUPLICATE_STEP0003_TOTAL_PLACEMENT_AVERAGE_ITEM",
                "duplicate placement/average weight item. Path = " + pszInputPath,
            )
        dicItems[pszItemName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicItems = dicInputs[objKey]
            arrMissing = [pszItem for pszItem in STEP0004_TOTAL_PLACEMENT_AVERAGE_ITEMS
                          if pszItem not in dicItems]
            if arrMissing:
                raise_conversion_error(
                    "MISSING_STEP0003_TOTAL_PLACEMENT_AVERAGE_ITEMS",
                    "missing placement/average weight items for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissing),
                )
            arrOutputHeader: List[str] = []
            arrOutputFields: List[str] = []
            for pszItem in STEP0004_TOTAL_PLACEMENT_AVERAGE_ITEMS:
                arrHeader, arrFields = read_step0003_total_placement_average_item(
                    dicItems[pszItem], pszItem
                )
                if not arrOutputFields:
                    arrOutputHeader = arrHeader[:4]
                    arrOutputFields = arrFields[:4]
                elif arrFields[:4] != arrOutputFields[:4]:
                    raise_conversion_error(
                        "STEP0003_TOTAL_COMMON_COLUMNS_MISMATCH",
                        "common columns do not match. Path = " + dicItems[pszItem],
                    )
                arrOutputHeader.append(arrHeader[4])
                arrOutputFields.append(arrFields[4])
            pszOutputPath = create_step0004_total_placement_average_output_path(
                dicItems[STEP0004_TOTAL_PLACEMENT_AVERAGE_ITEMS[0]]
            )
            write_step0004_tsv(
                pszOutputPath, "\t".join(arrOutputHeader), "\t".join(arrOutputFields)
            )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_TOTAL_PLACEMENT_AVERAGE_CLEANUP_FAILED",
                "placement/average weight combination failed. Original error = " + str(objException)
                + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


STEP0004_TOTAL_SHIPMENT_COUNT_WEIGHT_ITEMS: List[str] = [
    "出荷羽数", "出荷重量",
]


def create_step0004_total_shipment_count_weight_output_path(
    pszFirstInputPath: str,
) -> str:
    pszBaseName, _, iYear, iMonth, pszSuffix = parse_step0003_total_file_name(
        pszFirstInputPath
    )
    pszDirectory: str = os.path.dirname(pszFirstInputPath)
    pszStem: str = (
        f"{pszBaseName}_step0004_総計_羽数データ別_出荷羽数・出荷重量_{iYear:04d}_{iMonth:02d}"
        + pszSuffix
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def read_step0003_total_shipment_count_weight_item(
    pszInputPath: str, pszExpectedItem: str,
) -> Tuple[List[str], List[str]]:
    _, pszItemName, iFileYear, iFileMonth, _ = parse_step0003_total_file_name(
        pszInputPath
    )
    if pszItemName != pszExpectedItem:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ITEM", "unexpected item. Path = " + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0003_TOTAL_READ_FAILED",
            "failed to read step0003 total TSV. Path = " + pszInputPath
            + "; Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ROW_COUNT",
            "step0003 total TSV must contain one header and one data row. Path = "
            + pszInputPath,
        )
    arrHeader = arrLines[0].split("\t")
    arrFields = arrLines[1].split("\t")
    if len(arrHeader) != 5 or len(arrFields) != 5:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_COLUMN_COUNT",
            "step0003 total TSV must contain five columns. Path = " + pszInputPath,
        )
    if arrHeader != ["年", "月", "ｸﾞﾙｰﾌﾟ№", "ｸﾞﾙｰﾌﾟ名", pszExpectedItem]:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_HEADER",
            "unexpected step0003 total header. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (iFileYear, iFileMonth):
        raise_conversion_error(
            "STEP0003_TOTAL_YEAR_MONTH_MISMATCH",
            "file name and data year/month do not match. Path = " + pszInputPath,
        )
    if arrFields[2:4] != ["", "総計"]:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_GROUP",
            "expected an empty group number and total group name. Path = " + pszInputPath,
        )
    if parse_decimal_value(arrFields[4], 2, 4, pszExpectedItem) is None:
        raise_conversion_error(
            "INVALID_STEP0003_TOTAL_ITEM_VALUE",
            "shipment count/weight was blank. Path = " + pszInputPath,
        )
    return arrHeader, arrFields


def create_step0004_total_shipment_count_weight_files(
    arrStep0003TotalPaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """同じ年月の出荷羽数と出荷重量を横結合し、入力値の表記を保持します。"""
    dicInputs: Dict[Tuple[str, str, int, int, str], Dict[str, str]] = {}
    for pszInputPath in arrStep0003TotalPaths:
        pszBaseName, pszItemName, iYear, iMonth, pszSuffix = (
            parse_step0003_total_file_name(pszInputPath)
        )
        objKey = (os.path.dirname(os.path.abspath(pszInputPath)),
                  pszBaseName, iYear, iMonth, pszSuffix)
        dicItems = dicInputs.setdefault(objKey, {})
        if pszItemName not in STEP0004_TOTAL_SHIPMENT_COUNT_WEIGHT_ITEMS:
            continue
        if pszItemName in dicItems:
            raise_conversion_error(
                "DUPLICATE_STEP0003_TOTAL_SHIPMENT_COUNT_WEIGHT_ITEM",
                "duplicate shipment count/weight item. Path = " + pszInputPath,
            )
        dicItems[pszItemName] = pszInputPath

    arrCreatedPaths: List[str] = []
    arrYearMonths: List[Tuple[int, int]] = []
    try:
        for objKey in sorted(dicInputs, key=lambda objKey: (
            objKey[2], objKey[3], objKey[0], objKey[1], objKey[4]
        )):
            dicItems = dicInputs[objKey]
            arrMissing = [pszItem for pszItem in STEP0004_TOTAL_SHIPMENT_COUNT_WEIGHT_ITEMS
                          if pszItem not in dicItems]
            if arrMissing:
                raise_conversion_error(
                    "MISSING_STEP0003_TOTAL_SHIPMENT_COUNT_WEIGHT_ITEMS",
                    "missing shipment count/weight items for "
                    + f"{objKey[1]} {objKey[2]:04d}_{objKey[3]:02d}{objKey[4]}: "
                    + ", ".join(arrMissing),
                )
            arrOutputHeader: List[str] = []
            arrOutputFields: List[str] = []
            for pszItem in STEP0004_TOTAL_SHIPMENT_COUNT_WEIGHT_ITEMS:
                arrHeader, arrFields = read_step0003_total_shipment_count_weight_item(
                    dicItems[pszItem], pszItem
                )
                if not arrOutputFields:
                    arrOutputHeader = arrHeader[:4]
                    arrOutputFields = arrFields[:4]
                elif arrFields[:4] != arrOutputFields[:4]:
                    raise_conversion_error(
                        "STEP0003_TOTAL_COMMON_COLUMNS_MISMATCH",
                        "common columns do not match. Path = " + dicItems[pszItem],
                    )
                arrOutputHeader.append(arrHeader[4])
                arrOutputFields.append(arrFields[4])
            pszOutputPath = create_step0004_total_shipment_count_weight_output_path(
                dicItems[STEP0004_TOTAL_SHIPMENT_COUNT_WEIGHT_ITEMS[0]]
            )
            write_step0004_tsv(
                pszOutputPath, "\t".join(arrOutputHeader), "\t".join(arrOutputFields)
            )
            arrCreatedPaths.append(pszOutputPath)
            arrYearMonths.append((objKey[2], objKey[3]))
        return arrCreatedPaths, arrYearMonths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_TOTAL_SHIPMENT_COUNT_WEIGHT_CLEANUP_FAILED",
                "shipment count/weight combination failed. Original error = " + str(objException)
                + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0004_output_path(
    pszSourceTsvPath: str,
    pszGroupName: str,
    iYear: int,
    iMonth: int,
) -> str:
    pszDirectory: str = os.path.dirname(pszSourceTsvPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem: str = (
        f"{pszBaseName}_step0004_グループ別_{pszGroupName}_{iYear:04d}_{iMonth:02d}"
    )
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def write_step0004_tsv(
    pszOutputPath: str, pszHeader: str, pszRow: str
) -> None:
    try:
        write_monthly_tsv(pszOutputPath, pszHeader, [pszRow])
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0004_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0004_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0004_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0004_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0004_WRITE_FAILED"),
            str(objException),
        )


def build_step0003_path_map(
    arrStep0003Paths: List[str], arrStep0003YearMonths: List[Tuple[int, int]]
) -> Dict[Tuple[int, int, str], str]:
    iExpectedCount: int = len(arrStep0003YearMonths) * len(STEP0003_ITEMS)
    if len(arrStep0003Paths) != iExpectedCount:
        raise_conversion_error(
            "INVALID_STEP0003_PATH_COUNT",
            "step0003 path count did not match year/month and item counts. Expected = "
            + str(iExpectedCount)
            + ", Actual = "
            + str(len(arrStep0003Paths)),
        )
    dicPaths: Dict[Tuple[int, int, str], str] = {}
    iPathIndex: int = 0
    for iYear, iMonth in arrStep0003YearMonths:
        for pszItemName, _, _, _ in STEP0003_ITEMS:
            objKey: Tuple[int, int, str] = (iYear, iMonth, pszItemName)
            if objKey in dicPaths:
                raise_conversion_error(
                    "DUPLICATE_STEP0003_PATH_KEY",
                    "duplicate step0003 path key. Key = " + str(objKey),
                )
            dicPaths[objKey] = arrStep0003Paths[iPathIndex]
            iPathIndex += 1
    return dicPaths


def read_step0003_for_step0004(
    pszInputPath: str,
    pszExpectedItemName: str,
    bAllowBlank: bool,
    iExpectedYear: int,
    iExpectedMonth: int,
    bVariety: bool = False,
) -> Dict[Tuple[str, str], List[str]]:
    arrCategories = SPLIT_VARIETIES if bVariety else GROUPS
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0003_READ_FAILED",
            "failed to read step0003 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    if len(arrLines) != len(arrCategories) + 1:
        raise_conversion_error(
            "INVALID_STEP0003_ROW_COUNT",
            "step0003 must contain one header and "
            + ("five" if bVariety else "three") + " data rows. Path = "
            + pszInputPath
            + ", Actual rows = "
            + str(len(arrLines)),
        )
    arrHeader: List[str] = arrLines[0].split("\t")
    arrExpectedHeader: List[str] = [
        "年",
        "月",
        "品種No." if bVariety else "ｸﾞﾙｰﾌﾟ№",
        "品種名" if bVariety else "ｸﾞﾙｰﾌﾟ名",
        pszExpectedItemName,
    ]
    if len(arrHeader) != 5 or [
        pszValue.strip(" \u3000\t") for pszValue in arrHeader
    ] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0003_HEADER",
            "unexpected step0003 header. Path = "
            + pszInputPath
            + ", Expected = "
            + " | ".join(arrExpectedHeader)
            + ", Actual = "
            + " | ".join(arrHeader),
        )

    setExpectedGroups = set(arrCategories)
    dicRows: Dict[Tuple[str, str], List[str]] = {}
    for iLineNumber, pszLine in enumerate(arrLines[1:], start=2):
        arrFields: List[str] = pszLine.split("\t")
        if len(arrFields) != 5:
            raise_conversion_error(
                "INVALID_STEP0003_COLUMN_COUNT",
                "step0003 row must contain 5 columns. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber)
                + ", Actual = "
                + str(len(arrFields)),
            )
        iYear, iMonth = parse_year_month(arrFields[0], arrFields[1], iLineNumber)
        if (iYear, iMonth) != (iExpectedYear, iExpectedMonth):
            raise_conversion_error(
                "INVALID_STEP0003_YEAR_MONTH",
                "unexpected step0003 year/month. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber),
            )
        pszGroupNumber: str = arrFields[2].strip(" \u3000\t")
        pszGroupName: str = arrFields[3].strip(" \u3000\t")
        objExpectedKey = (
            arrCategories[iLineNumber - 2] if bVariety
            else (pszGroupNumber, pszGroupName)
        )
        if (
            (pszGroupNumber, pszGroupName) not in setExpectedGroups
            or (pszGroupNumber, pszGroupName) != objExpectedKey
        ):
            raise_conversion_error(
                "INVALID_STEP0003_VARIETY" if bVariety else "INVALID_STEP0003_GROUP",
                "unexpected step0003 " + ("variety" if bVariety else "group")
                + " mapping. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber),
            )
        objGroupKey: Tuple[str, str] = (pszGroupNumber, pszGroupName)
        if objGroupKey in dicRows:
            raise_conversion_error(
                "INVALID_STEP0003_VARIETY" if bVariety else "INVALID_STEP0003_GROUP",
                "duplicate step0003 " + ("variety" if bVariety else "group") + ". Path = "
                + pszInputPath
                + ", Group = "
                + str(objGroupKey),
            )
        try:
            objValue = parse_decimal_value(
                arrFields[4], iLineNumber, 4, pszExpectedItemName
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0003_ITEM_VALUE",
                "invalid step0003 item value. Path = "
                + pszInputPath
                + ", Item = "
                + pszExpectedItemName
                + ". Detail = "
                + str(objException),
            )
        if objValue is None and not bAllowBlank:
            raise_conversion_error(
                "INVALID_STEP0003_ITEM_VALUE",
                "step0003 item value was blank. Path = "
                + pszInputPath
                + ", Item = "
                + pszExpectedItemName
                + ", Line = "
                + str(iLineNumber),
            )
        dicRows[objGroupKey] = arrFields

    if set(dicRows) != setExpectedGroups:
        raise_conversion_error(
            "INVALID_STEP0003_VARIETY" if bVariety else "INVALID_STEP0003_GROUP",
            "step0003 " + ("varieties" if bVariety else "groups")
            + " were missing or unexpected. Path = " + pszInputPath,
        )
    return dicRows


def create_step0004_files(
    pszSourceTsvPath: str,
    arrStep0003Paths: List[str],
    arrStep0003YearMonths: List[Tuple[int, int]],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    dicStep0003Paths = build_step0003_path_map(
        arrStep0003Paths, arrStep0003YearMonths
    )
    arrCreatedPaths: List[str] = []
    try:
        for iYear, iMonth in arrStep0003YearMonths:
            dicItemRows: Dict[str, Dict[Tuple[str, str], List[str]]] = {}
            for pszItemName, bAllowBlank in STEP0004_ITEMS:
                objPathKey: Tuple[int, int, str] = (iYear, iMonth, pszItemName)
                if objPathKey not in dicStep0003Paths:
                    raise_conversion_error(
                        "STEP0003_INPUT_MISSING",
                        "required step0003 input was missing. Key = "
                        + str(objPathKey),
                    )
                pszInputPath: str = dicStep0003Paths[objPathKey]
                dicItemRows[pszItemName] = read_step0003_for_step0004(
                    pszInputPath,
                    pszItemName,
                    bAllowBlank,
                    iYear,
                    iMonth,
                )

            for pszGroupNumber, pszGroupName in GROUPS:
                objGroupKey: Tuple[str, str] = (pszGroupNumber, pszGroupName)
                arrCommonFields: List[str] = dicItemRows[STEP0004_ITEMS[0][0]][
                    objGroupKey
                ][:4]
                for pszItemName, _ in STEP0004_ITEMS[1:]:
                    if dicItemRows[pszItemName][objGroupKey][:4] != arrCommonFields:
                        raise_conversion_error(
                            "STEP0003_GROUP_KEY_MISMATCH",
                            "step0003 group keys did not match. Year/month = "
                            + f"{iYear:04d}_{iMonth:02d}"
                            + ", Group = "
                            + str(objGroupKey)
                            + ", Item = "
                            + pszItemName,
                        )
                arrOutputFields: List[str] = arrCommonFields + [
                    dicItemRows[pszItemName][objGroupKey][4]
                    for pszItemName, _ in STEP0004_ITEMS
                ]
                pszOutputPath: str = create_step0004_output_path(
                    pszSourceTsvPath,
                    pszGroupName,
                    iYear,
                    iMonth,
                )
                write_step0004_tsv(
                    pszOutputPath,
                    STEP0004_HEADER,
                    "\t".join(arrOutputFields),
                )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, list(arrStep0003YearMonths)
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_CLEANUP_FAILED",
                "step0004 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0004_variety_output_path(
    pszSourceTsvPath: str,
    pszVarietyName: str,
    iYear: int,
    iMonth: int,
) -> str:
    pszDirectory = os.path.dirname(pszSourceTsvPath)
    pszBaseName = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem = (
        f"{pszBaseName}_step0004_品種別_{pszVarietyName}_"
        f"{iYear:04d}_{iMonth:02d}"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0004_variety_files(
    pszSourceTsvPath: str,
    arrStep0003VarietyPaths: List[str],
    arrStep0003YearMonths: List[Tuple[int, int]],
    arrStep0003VarietyHusbandryAveragePaths: List[str],
    arrStep0003VarietyPreDiscardGrowthRatePaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """品種別step0003の6項目を品種ごとに結合したstep0004を作成します。"""
    dicStep0003Paths = build_step0003_path_map(
        arrStep0003VarietyPaths, arrStep0003YearMonths
    )

    def build_additional_path_map(
        arrPaths: List[str], pszFileItemName: str, pszErrorCode: str
    ) -> Dict[Tuple[int, int], str]:
        dicPaths: Dict[Tuple[int, int], str] = {}
        pszPattern = (
            r".+_step0003_品種別_" + re.escape(pszFileItemName)
            + r"_(\d{4})_(\d{2})(?:_\d+)*\.tsv"
        )
        for pszPath in arrPaths:
            objMatch = re.fullmatch(pszPattern, os.path.basename(pszPath))
            if objMatch is None:
                raise_conversion_error(
                    pszErrorCode,
                    "invalid variety step0003 file name. Path = " + pszPath,
                )
            iYear, iMonth = parse_year_month(
                objMatch.group(1), objMatch.group(2), 1
            )
            objKey = (iYear, iMonth)
            if objKey in dicPaths:
                raise_conversion_error(
                    pszErrorCode,
                    "duplicate variety step0003 item. Year/month = "
                    + f"{iYear:04d}_{iMonth:02d}"
                    + ", Item = " + pszFileItemName,
                )
            dicPaths[objKey] = pszPath
        if set(dicPaths) != set(arrStep0003YearMonths):
            raise_conversion_error(
                pszErrorCode,
                "variety step0003 inputs did not match expected year/months. "
                "Item = " + pszFileItemName,
            )
        return dicPaths

    dicHusbandryAveragePaths = build_additional_path_map(
        arrStep0003VarietyHusbandryAveragePaths,
        "飼育日数平均",
        "STEP0004_VARIETY_HUSBANDRY_AVERAGE_INPUT_MISMATCH",
    )
    dicPreDiscardGrowthRatePaths = build_additional_path_map(
        arrStep0003VarietyPreDiscardGrowthRatePaths,
        "廃棄控除前育成率",
        "STEP0004_VARIETY_PRE_DISCARD_GROWTH_RATE_INPUT_MISMATCH",
    )
    arrCreatedPaths: List[str] = []
    try:
        for iYear, iMonth in arrStep0003YearMonths:
            dicItemRows: Dict[str, Dict[Tuple[str, str], List[str]]] = {}
            for pszOutputItemName, pszInputItemName, bAllowBlank in STEP0004_VARIETY_ITEMS:
                if pszInputItemName == "日令":
                    pszInputPath = dicHusbandryAveragePaths[(iYear, iMonth)]
                elif pszInputItemName == "育成率":
                    pszInputPath = dicPreDiscardGrowthRatePaths[(iYear, iMonth)]
                else:
                    pszInputPath = dicStep0003Paths[
                        (iYear, iMonth, pszInputItemName)
                    ]
                dicItemRows[pszOutputItemName] = read_step0003_for_step0004(
                    pszInputPath,
                    pszInputItemName,
                    bAllowBlank,
                    iYear,
                    iMonth,
                    True,
                )

            for pszVarietyNumber, pszVarietyName in SPLIT_VARIETIES:
                objVarietyKey = (pszVarietyNumber, pszVarietyName)
                arrCommonFields = dicItemRows[STEP0004_VARIETY_ITEMS[0][0]][
                    objVarietyKey
                ][:4]
                for pszOutputItemName, _, _ in STEP0004_VARIETY_ITEMS[1:]:
                    if (
                        dicItemRows[pszOutputItemName][objVarietyKey][:4]
                        != arrCommonFields
                    ):
                        raise_conversion_error(
                            "STEP0003_VARIETY_KEY_MISMATCH",
                            "step0003 variety keys did not match. Year/month = "
                            + f"{iYear:04d}_{iMonth:02d}"
                            + ", Variety = " + str(objVarietyKey)
                            + ", Item = " + pszOutputItemName,
                        )
                arrOutputFields = arrCommonFields + [
                    dicItemRows[pszOutputItemName][objVarietyKey][4]
                    for pszOutputItemName, _, _ in STEP0004_VARIETY_ITEMS
                ]
                pszOutputPath = create_step0004_variety_output_path(
                    pszSourceTsvPath,
                    pszVarietyName,
                    iYear,
                    iMonth,
                )
                write_step0004_tsv(
                    pszOutputPath,
                    STEP0004_VARIETY_HEADER,
                    "\t".join(arrOutputFields),
                )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, list(arrStep0003YearMonths)
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_VARIETY_CLEANUP_FAILED",
                "variety step0004 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0004_variety_deduction_output_path(
    pszSourceTsvPath: str,
    pszVarietyName: str,
    iYear: int,
    iMonth: int,
) -> str:
    pszDirectory = os.path.dirname(pszSourceTsvPath)
    pszBaseName = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem = (
        f"{pszBaseName}_step0004_品種別_控除廃棄_{pszVarietyName}_"
        f"{iYear:04d}_{iMonth:02d}"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0004_variety_deduction_files(
    pszSourceTsvPath: str,
    arrStep0003VarietyPaths: List[str],
    arrStep0003YearMonths: List[Tuple[int, int]],
    arrStep0007VarietyPaths: List[str],
    arrStep0003VarietyHusbandryAveragePaths: List[str],
    arrStep0003VarietyDeductionGrowthRatePaths: List[str],
) -> Tuple[List[str], List[Tuple[int, int]]]:
    """廃棄控除後の品種別step0003項目をstep0004にまとめます。"""
    dicStep0003Paths = build_step0003_path_map(
        arrStep0003VarietyPaths, arrStep0003YearMonths
    )

    def build_item_path_map(
        arrPaths: List[str],
        pszFileItemName: str,
        pszErrorCode: str,
        bIgnoreOtherItems: bool = False,
    ) -> Dict[Tuple[int, int], str]:
        dicPaths: Dict[Tuple[int, int], str] = {}
        pszPattern = (
            r".+_step0003_品種別_" + re.escape(pszFileItemName)
            + r"_(\d{4})_(\d{2})(?:_\d+)*\.tsv"
        )
        for pszPath in arrPaths:
            objMatch = re.fullmatch(pszPattern, os.path.basename(pszPath))
            if objMatch is None:
                if bIgnoreOtherItems:
                    continue
                raise_conversion_error(
                    pszErrorCode,
                    "invalid variety step0003 file name. Path = " + pszPath,
                )
            iYear, iMonth = parse_year_month(
                objMatch.group(1), objMatch.group(2), 1
            )
            objKey = (iYear, iMonth)
            if objKey in dicPaths:
                raise_conversion_error(
                    pszErrorCode,
                    "duplicate variety step0003 item. Year/month = "
                    + f"{iYear:04d}_{iMonth:02d}"
                    + ", Item = " + pszFileItemName,
                )
            dicPaths[objKey] = pszPath
        if set(dicPaths) != set(arrStep0003YearMonths):
            raise_conversion_error(
                pszErrorCode,
                "variety step0003 inputs did not match expected year/months. "
                "Item = " + pszFileItemName,
            )
        return dicPaths

    dicProcessedCountPaths = build_item_path_map(
        arrStep0007VarietyPaths,
        "処理羽数",
        "STEP0004_VARIETY_PROCESSED_COUNT_INPUT_MISMATCH",
        True,
    )
    dicProcessedWeightPaths = build_item_path_map(
        arrStep0007VarietyPaths,
        "処理重量",
        "STEP0004_VARIETY_PROCESSED_WEIGHT_INPUT_MISMATCH",
        True,
    )
    dicPostDiscardAverageWeightPaths = build_item_path_map(
        arrStep0007VarietyPaths,
        "控除後平均体重",
        "STEP0004_VARIETY_POST_DISCARD_AVERAGE_WEIGHT_INPUT_MISMATCH",
        True,
    )
    dicHusbandryAveragePaths = build_item_path_map(
        arrStep0003VarietyHusbandryAveragePaths,
        "飼育日数平均",
        "STEP0004_VARIETY_HUSBANDRY_AVERAGE_INPUT_MISMATCH",
    )
    dicPostDiscardGrowthRatePaths = build_item_path_map(
        arrStep0003VarietyDeductionGrowthRatePaths,
        "廃棄控除後育成率",
        "STEP0004_VARIETY_POST_DISCARD_GROWTH_RATE_INPUT_MISMATCH",
    )
    dicFileItemPaths = {
        "入雛羽数": {
            (iYear, iMonth): dicStep0003Paths[(iYear, iMonth, "入雛羽数")]
            for iYear, iMonth in arrStep0003YearMonths
        },
        "処理羽数": dicProcessedCountPaths,
        "処理重量": dicProcessedWeightPaths,
        "控除後平均体重": dicPostDiscardAverageWeightPaths,
        "飼育日数平均": dicHusbandryAveragePaths,
        "廃棄控除後育成率": dicPostDiscardGrowthRatePaths,
    }
    dicInputHeaders = {
        "入雛羽数": "入雛羽数",
        "処理羽数": "処理羽数",
        "処理重量": "処理重量",
        "控除後平均体重": "平均体重",
        "飼育日数平均": "日令",
        "廃棄控除後育成率": "育成率",
    }
    arrCreatedPaths: List[str] = []
    try:
        for iYear, iMonth in arrStep0003YearMonths:
            dicItemRows: Dict[str, Dict[Tuple[str, str], List[str]]] = {}
            for pszOutputItemName, _, bAllowBlank in STEP0004_VARIETY_DEDUCTION_ITEMS:
                pszInputPath = dicFileItemPaths[pszOutputItemName][(iYear, iMonth)]
                dicItemRows[pszOutputItemName] = read_step0003_for_step0004(
                    pszInputPath,
                    dicInputHeaders[pszOutputItemName],
                    bAllowBlank,
                    iYear,
                    iMonth,
                    True,
                )

            for pszVarietyNumber, pszVarietyName in SPLIT_VARIETIES:
                objVarietyKey = (pszVarietyNumber, pszVarietyName)
                arrCommonFields = dicItemRows[
                    STEP0004_VARIETY_DEDUCTION_ITEMS[0][0]
                ][objVarietyKey][:4]
                for pszOutputItemName, _, _ in STEP0004_VARIETY_DEDUCTION_ITEMS[1:]:
                    if (
                        dicItemRows[pszOutputItemName][objVarietyKey][:4]
                        != arrCommonFields
                    ):
                        raise_conversion_error(
                            "STEP0003_VARIETY_DEDUCTION_KEY_MISMATCH",
                            "step0003 deduction variety keys did not match. "
                            "Year/month = " + f"{iYear:04d}_{iMonth:02d}"
                            + ", Variety = " + str(objVarietyKey)
                            + ", Item = " + pszOutputItemName,
                        )
                arrOutputFields = arrCommonFields + [
                    dicItemRows[pszOutputItemName][objVarietyKey][4]
                    for pszOutputItemName, _, _ in STEP0004_VARIETY_DEDUCTION_ITEMS
                ]
                pszOutputPath = create_step0004_variety_deduction_output_path(
                    pszSourceTsvPath,
                    pszVarietyName,
                    iYear,
                    iMonth,
                )
                write_step0004_tsv(
                    pszOutputPath,
                    STEP0004_VARIETY_DEDUCTION_HEADER,
                    "\t".join(arrOutputFields),
                )
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, list(arrStep0003YearMonths)
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0004_VARIETY_DEDUCTION_CLEANUP_FAILED",
                "variety deduction step0004 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0005_output_path(
    pszSourceTsvPath: str, pszGroupName: str, iYear: int
) -> str:
    pszDirectory: str = os.path.dirname(pszSourceTsvPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem: str = f"{pszBaseName}_step0005_グループ別_{pszGroupName}_{iYear:04d}_年間"
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv")
        iSuffix += 1
    return pszOutputPath


def write_step0005_tsv(
    pszOutputPath: str, pszHeader: str, arrRows: List[str]
) -> None:
    try:
        write_monthly_tsv(pszOutputPath, pszHeader, arrRows)
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0005_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0005_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0005_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0005_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0005_WRITE_FAILED"),
            str(objException),
        )


def build_step0004_path_map(
    arrStep0004Paths: List[str], arrStep0004YearMonths: List[Tuple[int, int]]
) -> Dict[Tuple[int, int, str, str], str]:
    iExpectedCount: int = len(arrStep0004YearMonths) * len(GROUPS)
    if len(arrStep0004Paths) != iExpectedCount:
        raise_conversion_error(
            "INVALID_STEP0004_PATH_COUNT",
            "step0004 path count did not match year/month and group counts. Expected = "
            + str(iExpectedCount)
            + ", Actual = "
            + str(len(arrStep0004Paths)),
        )
    dicPaths: Dict[Tuple[int, int, str, str], str] = {}
    iPathIndex: int = 0
    for iYear, iMonth in arrStep0004YearMonths:
        for pszGroupNumber, pszGroupName in GROUPS:
            objKey: Tuple[int, int, str, str] = (
                iYear,
                iMonth,
                pszGroupNumber,
                pszGroupName,
            )
            if objKey in dicPaths:
                raise_conversion_error(
                    "DUPLICATE_STEP0004_YEAR_MONTH_GROUP",
                    "duplicate step0004 year/month/group key. Key = " + str(objKey),
                )
            dicPaths[objKey] = arrStep0004Paths[iPathIndex]
            iPathIndex += 1
    return dicPaths


def read_step0004_for_step0005(
    pszInputPath: str,
    iExpectedYear: int,
    iExpectedMonth: int,
    pszExpectedGroupNumber: str,
    pszExpectedGroupName: str,
) -> Tuple[str, str]:
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines: List[str] = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0004_READ_FAILED",
            "failed to read step0004 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0004_ROW_COUNT",
            "step0004 must contain one header and one data row. Path = "
            + pszInputPath
            + ", Actual rows = "
            + str(len(arrLines)),
        )
    arrHeader: List[str] = arrLines[0].split("\t")
    arrExpectedHeader: List[str] = STEP0004_HEADER.split("\t")
    if len(arrHeader) != 8 or [
        pszValue.strip(" \u3000\t") for pszValue in arrHeader
    ] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0004_HEADER",
            "unexpected step0004 header. Path = "
            + pszInputPath
            + ", Expected = "
            + " | ".join(arrExpectedHeader)
            + ", Actual = "
            + " | ".join(arrHeader),
        )
    arrFields: List[str] = arrLines[1].split("\t")
    if len(arrFields) != 8:
        raise_conversion_error(
            "INVALID_STEP0004_COLUMN_COUNT",
            "step0004 data row must contain 8 columns. Path = "
            + pszInputPath
            + ", Actual = "
            + str(len(arrFields)),
        )
    iYear, iMonth = parse_year_month(arrFields[0], arrFields[1], 2)
    if (iYear, iMonth) != (iExpectedYear, iExpectedMonth):
        raise_conversion_error(
            "INVALID_STEP0004_YEAR_MONTH",
            "unexpected step0004 year/month. Path = " + pszInputPath,
        )
    pszGroupNumber: str = arrFields[2].strip(" \u3000\t")
    pszGroupName: str = arrFields[3].strip(" \u3000\t")
    if (pszGroupNumber, pszGroupName) != (
        pszExpectedGroupNumber,
        pszExpectedGroupName,
    ):
        raise_conversion_error(
            "INVALID_STEP0004_GROUP",
            "unexpected step0004 group. Path = " + pszInputPath,
        )
    for iColumnIndex, (pszItemName, bAllowBlank) in enumerate(
        STEP0004_ITEMS, start=4
    ):
        try:
            objValue = parse_decimal_value(
                arrFields[iColumnIndex], 2, iColumnIndex, pszItemName
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0004_ITEM_VALUE",
                "invalid step0004 item value. Path = "
                + pszInputPath
                + ", Item = "
                + pszItemName
                + ". Detail = "
                + str(objException),
            )
        if objValue is None and not bAllowBlank:
            raise_conversion_error(
                "INVALID_STEP0004_ITEM_VALUE",
                "step0004 item value was blank. Path = "
                + pszInputPath
                + ", Item = "
                + pszItemName,
            )
    return arrLines[0], arrLines[1]


def create_step0005_files(
    pszSourceTsvPath: str,
    arrStep0004Paths: List[str],
    arrStep0004YearMonths: List[Tuple[int, int]],
) -> Tuple[List[str], List[int]]:
    dicStep0004Paths = build_step0004_path_map(
        arrStep0004Paths, arrStep0004YearMonths
    )
    arrYears: List[int] = sorted({iYear for iYear, _ in arrStep0004YearMonths})
    arrCreatedPaths: List[str] = []
    try:
        for iYear in arrYears:
            arrMonths: List[int] = sorted(
                {
                    iMonth
                    for iTargetYear, iMonth in arrStep0004YearMonths
                    if iTargetYear == iYear
                }
            )
            for pszGroupNumber, pszGroupName in GROUPS:
                pszCommonHeader: str = ""
                arrRows: List[str] = []
                for iMonth in arrMonths:
                    objPathKey: Tuple[int, int, str, str] = (
                        iYear,
                        iMonth,
                        pszGroupNumber,
                        pszGroupName,
                    )
                    if objPathKey not in dicStep0004Paths:
                        raise_conversion_error(
                            "STEP0004_INPUT_MISSING",
                            "required step0004 input was missing. Key = "
                            + str(objPathKey),
                        )
                    pszHeader, pszRow = read_step0004_for_step0005(
                        dicStep0004Paths[objPathKey],
                        iYear,
                        iMonth,
                        pszGroupNumber,
                        pszGroupName,
                    )
                    if not pszCommonHeader:
                        pszCommonHeader = pszHeader
                    elif pszHeader != pszCommonHeader:
                        raise_conversion_error(
                            "INVALID_STEP0004_HEADER",
                            "step0004 headers did not match within annual group. Year = "
                            + str(iYear)
                            + ", Group = "
                            + pszGroupName,
                        )
                    arrRows.append(pszRow)
                if not arrRows:
                    continue
                pszOutputPath: str = create_step0005_output_path(
                    pszSourceTsvPath, pszGroupName, iYear
                )
                write_step0005_tsv(pszOutputPath, pszCommonHeader, arrRows)
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, arrYears
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0005_CLEANUP_FAILED",
                "step0005 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0005_variety_output_path(
    pszSourceTsvPath: str,
    pszVarietyName: str,
    iYear: int,
) -> str:
    pszDirectory = os.path.dirname(pszSourceTsvPath)
    pszBaseName = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem = (
        f"{pszBaseName}_step0005_品種別_{pszVarietyName}_"
        f"{iYear:04d}_年間"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def build_step0004_variety_path_map(
    arrStep0004Paths: List[str], arrStep0004YearMonths: List[Tuple[int, int]]
) -> Dict[Tuple[int, int, str, str], str]:
    setExpectedKeys = {
        (iYear, iMonth, pszVarietyNumber, pszVarietyName)
        for iYear, iMonth in arrStep0004YearMonths
        for pszVarietyNumber, pszVarietyName in SPLIT_VARIETIES
    }
    dicPaths: Dict[Tuple[int, int, str, str], str] = {}
    for pszPath in arrStep0004Paths:
        objMatch = re.fullmatch(
            r".+_step0004_品種別_" + VARIETY_NAME_PATTERN + r"_"
            r"(\d{4})_(\d{2})(?:_\d+)?\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error(
                "INVALID_STEP0004_VARIETY_FILE_NAME",
                "invalid variety step0004 file name. Path = " + pszPath,
            )
        pszVarietyName = objMatch.group(1)
        iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
        dicVarietyNumbers = {
            pszName: pszNumber for pszNumber, pszName in SPLIT_VARIETIES
        }
        objKey = (
            iYear,
            iMonth,
            dicVarietyNumbers[pszVarietyName],
            pszVarietyName,
        )
        if objKey in dicPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0004_VARIETY_YEAR_MONTH",
                "duplicate variety step0004 year/month input. Key = "
                + str(objKey),
            )
        dicPaths[objKey] = pszPath
    if set(dicPaths) != setExpectedKeys:
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_PATH_COUNT",
            "step0004 variety paths did not match expected year/months and "
            "varieties. Expected = " + str(len(setExpectedKeys))
            + ", Actual = " + str(len(dicPaths)),
        )
    return dicPaths


def read_step0004_variety_for_step0005(
    pszInputPath: str,
    iExpectedYear: int,
    iExpectedMonth: int,
    pszExpectedVarietyNumber: str,
    pszExpectedVarietyName: str,
) -> Tuple[str, str]:
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0004_VARIETY_READ_FAILED",
            "failed to read variety step0004 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_ROW_COUNT",
            "variety step0004 must contain one header and one data row. Path = "
            + pszInputPath + ", Actual rows = " + str(len(arrLines)),
        )
    arrExpectedHeader = STEP0004_VARIETY_HEADER.split("\t")
    arrHeader = arrLines[0].split("\t")
    if len(arrHeader) != len(arrExpectedHeader) or [
        pszValue.strip(" \u3000\t") for pszValue in arrHeader
    ] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_HEADER",
            "unexpected variety step0004 header. Path = " + pszInputPath,
        )
    arrFields = arrLines[1].split("\t")
    if len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_COLUMN_COUNT",
            "variety step0004 data row must contain "
            + str(len(arrExpectedHeader))
            + " columns. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (
        iExpectedYear,
        iExpectedMonth,
    ):
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_YEAR_MONTH",
            "unexpected variety step0004 year/month. Path = " + pszInputPath,
        )
    if arrFields[2:4] != [pszExpectedVarietyNumber, pszExpectedVarietyName]:
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY",
            "unexpected step0004 variety. Path = " + pszInputPath,
        )
    for iColumnIndex, (pszItemName, _, bAllowBlank) in enumerate(
        STEP0004_VARIETY_ITEMS, start=4
    ):
        try:
            objValue = parse_decimal_value(
                arrFields[iColumnIndex], 2, iColumnIndex, pszItemName
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0004_VARIETY_ITEM_VALUE",
                "invalid variety step0004 item value. Path = "
                + pszInputPath + ", Item = " + pszItemName
                + ". Detail = " + str(objException),
            )
        if objValue is None and not bAllowBlank:
            raise_conversion_error(
                "INVALID_STEP0004_VARIETY_ITEM_VALUE",
                "variety step0004 item value was blank. Path = "
                + pszInputPath + ", Item = " + pszItemName,
            )
    return arrLines[0], arrLines[1]


def create_step0005_variety_files(
    pszSourceTsvPath: str,
    arrStep0004VarietyPaths: List[str],
    arrStep0004VarietyYearMonths: List[Tuple[int, int]],
) -> Tuple[List[str], List[int]]:
    """品種別step0004の各月行を、品種ごとに年単位で連結します。"""
    dicStep0004Paths = build_step0004_variety_path_map(
        arrStep0004VarietyPaths, arrStep0004VarietyYearMonths
    )
    arrYears = sorted({iYear for iYear, _ in arrStep0004VarietyYearMonths})
    arrCreatedPaths: List[str] = []
    try:
        for iYear in arrYears:
            arrMonths = sorted(
                {
                    iMonth
                    for iTargetYear, iMonth in arrStep0004VarietyYearMonths
                    if iTargetYear == iYear
                }
            )
            for pszVarietyNumber, pszVarietyName in SPLIT_VARIETIES:
                pszCommonHeader = ""
                arrRows: List[str] = []
                for iMonth in arrMonths:
                    objPathKey = (
                        iYear,
                        iMonth,
                        pszVarietyNumber,
                        pszVarietyName,
                    )
                    if objPathKey not in dicStep0004Paths:
                        raise_conversion_error(
                            "STEP0004_VARIETY_INPUT_MISSING",
                            "required variety step0004 input was missing. Key = "
                            + str(objPathKey),
                        )
                    pszHeader, pszRow = read_step0004_variety_for_step0005(
                        dicStep0004Paths[objPathKey],
                        iYear,
                        iMonth,
                        pszVarietyNumber,
                        pszVarietyName,
                    )
                    if not pszCommonHeader:
                        pszCommonHeader = pszHeader
                    elif pszHeader != pszCommonHeader:
                        raise_conversion_error(
                            "INVALID_STEP0004_VARIETY_HEADER",
                            "variety step0004 headers did not match within "
                            "annual variety. Year = " + str(iYear)
                            + ", Variety = " + pszVarietyName,
                        )
                    arrRows.append(pszRow)
                if not arrRows:
                    continue
                pszOutputPath = create_step0005_variety_output_path(
                    pszSourceTsvPath,
                    pszVarietyName,
                    iYear,
                )
                write_step0005_tsv(pszOutputPath, pszCommonHeader, arrRows)
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, arrYears
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0005_VARIETY_CLEANUP_FAILED",
                "variety step0005 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0005_variety_deduction_output_path(
    pszSourceTsvPath: str,
    pszVarietyName: str,
    iYear: int,
) -> str:
    pszDirectory = os.path.dirname(pszSourceTsvPath)
    pszBaseName = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem = (
        f"{pszBaseName}_step0005_品種別_控除廃棄_{pszVarietyName}_"
        f"{iYear:04d}_年間"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def build_step0004_variety_deduction_path_map(
    arrStep0004Paths: List[str], arrStep0004YearMonths: List[Tuple[int, int]]
) -> Dict[Tuple[int, int, str, str], str]:
    setExpectedKeys = {
        (iYear, iMonth, pszVarietyNumber, pszVarietyName)
        for iYear, iMonth in arrStep0004YearMonths
        for pszVarietyNumber, pszVarietyName in SPLIT_VARIETIES
    }
    dicVarietyNumbers = {
        pszName: pszNumber for pszNumber, pszName in SPLIT_VARIETIES
    }
    dicPaths: Dict[Tuple[int, int, str, str], str] = {}
    for pszPath in arrStep0004Paths:
        objMatch = re.fullmatch(
            r".+_step0004_品種別_控除廃棄_" + VARIETY_NAME_PATTERN + r"_(\d{4})_(\d{2})"
            r"(?:_\d+)?\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error(
                "INVALID_STEP0004_VARIETY_DEDUCTION_FILE_NAME",
                "invalid variety deduction step0004 file name. Path = "
                + pszPath,
            )
        pszVarietyName = objMatch.group(1)
        iYear, iMonth = parse_year_month(objMatch.group(2), objMatch.group(3), 1)
        objKey = (
            iYear,
            iMonth,
            dicVarietyNumbers[pszVarietyName],
            pszVarietyName,
        )
        if objKey in dicPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0004_VARIETY_DEDUCTION_YEAR_MONTH",
                "duplicate variety deduction step0004 input. Key = "
                + str(objKey),
            )
        dicPaths[objKey] = pszPath
    if set(dicPaths) != setExpectedKeys:
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_DEDUCTION_PATH_COUNT",
            "step0004 variety deduction paths did not match expected "
            "year/months and varieties. Expected = " + str(len(setExpectedKeys))
            + ", Actual = " + str(len(dicPaths)),
        )
    return dicPaths


def read_step0004_variety_deduction_for_step0005(
    pszInputPath: str,
    iExpectedYear: int,
    iExpectedMonth: int,
    pszExpectedVarietyNumber: str,
    pszExpectedVarietyName: str,
) -> Tuple[str, str]:
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrLines = [strip_line_ending(pszLine) for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0004_VARIETY_DEDUCTION_READ_FAILED",
            "failed to read variety deduction step0004 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    if len(arrLines) != 2:
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_DEDUCTION_ROW_COUNT",
            "variety deduction step0004 must contain one header and one data row. "
            "Path = " + pszInputPath,
        )
    arrExpectedHeader = STEP0004_VARIETY_DEDUCTION_HEADER.split("\t")
    arrHeader = arrLines[0].split("\t")
    if len(arrHeader) != len(arrExpectedHeader) or [
        pszValue.strip(" \u3000\t") for pszValue in arrHeader
    ] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_DEDUCTION_HEADER",
            "unexpected variety deduction step0004 header. Path = "
            + pszInputPath,
        )
    arrFields = arrLines[1].split("\t")
    if len(arrFields) != len(arrExpectedHeader):
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_DEDUCTION_COLUMN_COUNT",
            "variety deduction step0004 data row must contain "
            + str(len(arrExpectedHeader)) + " columns. Path = " + pszInputPath,
        )
    if parse_year_month(arrFields[0], arrFields[1], 2) != (
        iExpectedYear,
        iExpectedMonth,
    ):
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_DEDUCTION_YEAR_MONTH",
            "unexpected variety deduction step0004 year/month. Path = "
            + pszInputPath,
        )
    if arrFields[2:4] != [pszExpectedVarietyNumber, pszExpectedVarietyName]:
        raise_conversion_error(
            "INVALID_STEP0004_VARIETY_DEDUCTION_VARIETY",
            "unexpected variety in deduction step0004. Path = " + pszInputPath,
        )
    for iColumnIndex, (pszItemName, _, bAllowBlank) in enumerate(
        STEP0004_VARIETY_DEDUCTION_ITEMS, start=4
    ):
        try:
            objValue = parse_decimal_value(
                arrFields[iColumnIndex], 2, iColumnIndex, pszItemName
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0004_VARIETY_DEDUCTION_ITEM_VALUE",
                "invalid variety deduction step0004 item value. Path = "
                + pszInputPath + ", Item = " + pszItemName
                + ". Detail = " + str(objException),
            )
        if objValue is None and not bAllowBlank:
            raise_conversion_error(
                "INVALID_STEP0004_VARIETY_DEDUCTION_ITEM_VALUE",
                "variety deduction step0004 item value was blank. Path = "
                + pszInputPath + ", Item = " + pszItemName,
            )
    return arrLines[0], arrLines[1]


def create_step0005_variety_deduction_files(
    pszSourceTsvPath: str,
    arrStep0004VarietyDeductionPaths: List[str],
    arrStep0004VarietyDeductionYearMonths: List[Tuple[int, int]],
) -> Tuple[List[str], List[int]]:
    """控除廃棄版の品種別step0004を、品種ごとに年単位で連結します。"""
    dicStep0004Paths = build_step0004_variety_deduction_path_map(
        arrStep0004VarietyDeductionPaths,
        arrStep0004VarietyDeductionYearMonths,
    )
    arrYears = sorted(
        {iYear for iYear, _ in arrStep0004VarietyDeductionYearMonths}
    )
    arrCreatedPaths: List[str] = []
    try:
        for iYear in arrYears:
            arrMonths = sorted(
                {
                    iMonth
                    for iTargetYear, iMonth in arrStep0004VarietyDeductionYearMonths
                    if iTargetYear == iYear
                }
            )
            for pszVarietyNumber, pszVarietyName in SPLIT_VARIETIES:
                pszCommonHeader = ""
                arrRows: List[str] = []
                for iMonth in arrMonths:
                    objPathKey = (
                        iYear,
                        iMonth,
                        pszVarietyNumber,
                        pszVarietyName,
                    )
                    if objPathKey not in dicStep0004Paths:
                        raise_conversion_error(
                            "STEP0004_VARIETY_DEDUCTION_INPUT_MISSING",
                            "required variety deduction step0004 input was "
                            "missing. Key = " + str(objPathKey),
                        )
                    pszHeader, pszRow = read_step0004_variety_deduction_for_step0005(
                        dicStep0004Paths[objPathKey],
                        iYear,
                        iMonth,
                        pszVarietyNumber,
                        pszVarietyName,
                    )
                    if not pszCommonHeader:
                        pszCommonHeader = pszHeader
                    elif pszHeader != pszCommonHeader:
                        raise_conversion_error(
                            "INVALID_STEP0004_VARIETY_DEDUCTION_HEADER",
                            "variety deduction step0004 headers did not match "
                            "within annual variety. Year = " + str(iYear)
                            + ", Variety = " + pszVarietyName,
                        )
                    arrRows.append(pszRow)
                if not arrRows:
                    continue
                pszOutputPath = create_step0005_variety_deduction_output_path(
                    pszSourceTsvPath,
                    pszVarietyName,
                    iYear,
                )
                write_step0005_tsv(pszOutputPath, pszCommonHeader, arrRows)
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, arrYears
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0005_VARIETY_DEDUCTION_CLEANUP_FAILED",
                "variety deduction step0005 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0006_output_path(pszStep0005Path: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0005Path)
    pszFileName: str = os.path.basename(pszStep0005Path)
    if "_step0005_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0005_FILENAME",
            "step0005 file name did not contain the expected marker. Path = "
            + pszStep0005Path,
        )
    pszStep0006FileName: str = pszFileName.replace(
        "_step0005_", "_step0006_", 1
    )
    pszStem, pszExtension = os.path.splitext(pszStep0006FileName)
    pszOutputPath: str = os.path.join(pszDirectory, pszStep0006FileName)
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def read_step0005_for_step0006(pszInputPath: str) -> List[List[str]]:
    objFileNameMatch = re.search(
        r"_step0005_グループ別_(契約|直営|フリー)_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0005_FILENAME",
            "unexpected step0005 file name. Path = " + pszInputPath,
        )
    pszExpectedGroupName: str = objFileNameMatch.group(1)
    iExpectedYear: int = int(objFileNameMatch.group(2))
    dicExpectedGroups: Dict[str, str] = {
        pszGroupName: pszGroupNumber for pszGroupNumber, pszGroupName in GROUPS
    }
    pszExpectedGroupNumber: str = dicExpectedGroups[pszExpectedGroupName]

    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0005_READ_FAILED",
            "failed to read step0005 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    if len(arrRows) < 2:
        raise_conversion_error(
            "INVALID_STEP0005_ROW_COUNT",
            "step0005 must contain one header and at least one data row. Path = "
            + pszInputPath,
        )
    arrExpectedHeader: List[str] = STEP0004_HEADER.split("\t")
    if [pszValue.strip(" \u3000\t") for pszValue in arrRows[0]] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0005_HEADER",
            "unexpected step0005 header. Path = " + pszInputPath,
        )

    iPreviousMonth: int = 0
    setMonths = set()
    for iLineNumber, arrFields in enumerate(arrRows[1:], start=2):
        if len(arrFields) != len(arrExpectedHeader):
            raise_conversion_error(
                "INVALID_STEP0005_COLUMN_COUNT",
                "step0005 row must contain 8 columns. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber)
                + ", Actual = "
                + str(len(arrFields)),
            )
        iYear, iMonth = parse_year_month(
            arrFields[0], arrFields[1], iLineNumber
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0005_YEAR",
                "step0005 year did not match its file name. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber),
            )
        if iMonth in setMonths:
            raise_conversion_error(
                "DUPLICATE_STEP0005_MONTH",
                "duplicate month in step0005. Path = "
                + pszInputPath
                + ", Month = "
                + str(iMonth),
            )
        if iMonth <= iPreviousMonth:
            raise_conversion_error(
                "INVALID_STEP0005_MONTH_ORDER",
                "step0005 months were not in ascending order. Path = "
                + pszInputPath,
            )
        setMonths.add(iMonth)
        iPreviousMonth = iMonth
        if (
            arrFields[2].strip(" \u3000\t") != pszExpectedGroupNumber
            or arrFields[3].strip(" \u3000\t") != pszExpectedGroupName
        ):
            raise_conversion_error(
                "INVALID_STEP0005_GROUP",
                "step0005 group did not match its file name. Path = "
                + pszInputPath
                + ", Line = "
                + str(iLineNumber),
            )
        for iColumnIndex, (pszItemName, bAllowBlank) in enumerate(
            STEP0004_ITEMS, start=4
        ):
            try:
                objValue = parse_decimal_value(
                    arrFields[iColumnIndex],
                    iLineNumber,
                    iColumnIndex,
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0005_ITEM_VALUE",
                    "invalid step0005 item value. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iLineNumber)
                    + ", Item = "
                    + pszItemName
                    + ". Detail = "
                    + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0005_ITEM_VALUE",
                    "step0005 item value was blank. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iLineNumber)
                    + ", Item = "
                    + pszItemName,
                )
    return arrRows


def transpose_tsv_rows(arrRows: List[List[str]]) -> List[List[str]]:
    if not arrRows:
        raise_conversion_error(
            "INVALID_STEP0005_ROW_COUNT", "step0005 did not contain any rows."
        )
    iColumnCount: int = len(arrRows[0])
    if any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0005_COLUMN_COUNT",
            "step0005 rows did not have a consistent column count.",
        )
    return [
        [arrRows[iRowIndex][iColumnIndex] for iRowIndex in range(len(arrRows))]
        for iColumnIndex in range(iColumnCount)
    ]


def write_step0006_tsv(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    if not arrRows:
        raise_conversion_error(
            "INVALID_STEP0006_ROW_COUNT", "step0006 did not contain any rows."
        )
    try:
        write_monthly_tsv(
            pszOutputPath,
            "\t".join(arrRows[0]),
            ["\t".join(arrRow) for arrRow in arrRows[1:]],
        )
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0006_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0006_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0006_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0006_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0006_WRITE_FAILED"),
            str(objException),
        )


def create_step0006_files(arrStep0005Paths: List[str]) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0005Path in arrStep0005Paths:
            arrInputRows: List[List[str]] = read_step0005_for_step0006(
                pszStep0005Path
            )
            arrTransposedRows: List[List[str]] = transpose_tsv_rows(arrInputRows)
            pszOutputPath: str = create_step0006_output_path(pszStep0005Path)
            write_step0006_tsv(pszOutputPath, arrTransposedRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0006_CLEANUP_FAILED",
                "step0006 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0005_variety_for_step0006(
    pszInputPath: str,
) -> List[List[str]]:
    """品種別step0005年間TSVを検証して読み込みます。"""
    objFileNameMatch = re.search(
        r"_step0005_品種別_" + VARIETY_NAME_PATTERN + r"_"
        r"(\d{4})_年間(?:_\d+)?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0005_VARIETY_FILENAME",
            "unexpected variety step0005 file name. Path = " + pszInputPath,
        )
    pszExpectedVarietyName = objFileNameMatch.group(1)
    iExpectedYear = int(objFileNameMatch.group(2))
    dicVarietyNumbers = {
        pszName: pszNumber for pszNumber, pszName in SPLIT_VARIETIES
    }
    pszExpectedVarietyNumber = dicVarietyNumbers[pszExpectedVarietyName]

    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0005_VARIETY_READ_FAILED",
            "failed to read variety step0005 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    arrExpectedHeader = STEP0004_VARIETY_HEADER.split("\t")
    if len(arrRows) < 2:
        raise_conversion_error(
            "INVALID_STEP0005_VARIETY_ROW_COUNT",
            "variety step0005 must contain one header and at least one data row. Path = "
            + pszInputPath,
        )
    if [pszValue.strip(" \u3000\t") for pszValue in arrRows[0]] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0005_VARIETY_HEADER",
            "unexpected variety step0005 header. Path = " + pszInputPath,
        )

    iPreviousMonth = 0
    setMonths = set()
    for iLineNumber, arrFields in enumerate(arrRows[1:], start=2):
        if len(arrFields) != len(arrExpectedHeader):
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_COLUMN_COUNT",
                "variety step0005 row must contain 10 columns. Path = "
                + pszInputPath + ", Line = " + str(iLineNumber),
            )
        iYear, iMonth = parse_year_month(
            arrFields[0], arrFields[1], iLineNumber
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_YEAR",
                "variety step0005 year did not match its file name. Path = "
                + pszInputPath + ", Line = " + str(iLineNumber),
            )
        if iMonth in setMonths or iMonth <= iPreviousMonth:
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_MONTH_ORDER",
                "variety step0005 months must be unique and ascending. Path = "
                + pszInputPath,
            )
        setMonths.add(iMonth)
        iPreviousMonth = iMonth
        if arrFields[2:4] != [
            pszExpectedVarietyNumber,
            pszExpectedVarietyName,
        ]:
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY",
                "variety step0005 variety did not match its file name. Path = "
                + pszInputPath + ", Line = " + str(iLineNumber),
            )
        for iColumnIndex, (pszItemName, _, bAllowBlank) in enumerate(
            STEP0004_VARIETY_ITEMS, start=4
        ):
            try:
                objValue = parse_decimal_value(
                    arrFields[iColumnIndex],
                    iLineNumber,
                    iColumnIndex,
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0005_VARIETY_ITEM_VALUE",
                    "invalid variety step0005 item value. Path = "
                    + pszInputPath + ", Line = " + str(iLineNumber)
                    + ", Item = " + pszItemName
                    + ". Detail = " + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0005_VARIETY_ITEM_VALUE",
                    "variety step0005 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iLineNumber)
                    + ", Item = " + pszItemName,
                )
    return arrRows


def create_step0006_variety_files(
    arrStep0005VarietyPaths: List[str],
) -> List[str]:
    """通常の品種別step0005年間TSVを行列入れ替えしてstep0006を作成します。"""
    dicInputs: Dict[Tuple[int, str], str] = {}
    for pszPath in arrStep0005VarietyPaths:
        objMatch = re.search(
            r"_step0005_品種別_" + VARIETY_NAME_PATTERN + r"_"
            r"(\d{4})_年間(?:_\d+)?\.tsv$",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_FILENAME",
                "unexpected variety step0005 file name. Path = " + pszPath,
            )
        objKey = (int(objMatch.group(2)), objMatch.group(1))
        if objKey in dicInputs:
            raise_conversion_error(
                "DUPLICATE_STEP0005_VARIETY_YEAR",
                "duplicate variety step0005 year input. Key = " + str(objKey),
            )
        dicInputs[objKey] = pszPath

    setYears = {iYear for iYear, _ in dicInputs}
    setExpectedKeys = {
        (iYear, pszVarietyName)
        for iYear in setYears
        for _, pszVarietyName in SPLIT_VARIETIES
    }
    if set(dicInputs) != setExpectedKeys:
        raise_conversion_error(
            "INVALID_STEP0005_VARIETY_PATH_COUNT",
            "variety step0005 paths must contain all five varieties for each year. "
            "Expected = " + str(len(setExpectedKeys))
            + ", Actual = " + str(len(dicInputs)),
        )

    arrCreatedPaths: List[str] = []
    try:
        for iYear in sorted(setYears):
            for _, pszVarietyName in SPLIT_VARIETIES:
                pszInputPath = dicInputs[(iYear, pszVarietyName)]
                arrInputRows = read_step0005_variety_for_step0006(pszInputPath)
                arrTransposedRows = transpose_tsv_rows(arrInputRows)
                pszOutputPath = create_step0006_output_path(pszInputPath)
                write_step0006_tsv(pszOutputPath, arrTransposedRows)
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0006_VARIETY_CLEANUP_FAILED",
                "variety step0006 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0005_variety_deduction_for_step0006(
    pszInputPath: str,
) -> List[List[str]]:
    """控除廃棄版の品種別step0005年間TSVを検証して読み込みます。"""
    objFileNameMatch = re.search(
        r"_step0005_品種別_控除廃棄_" + VARIETY_NAME_PATTERN + r"_"
        r"(\d{4})_年間(?:_\d+)?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0005_VARIETY_DEDUCTION_FILENAME",
            "unexpected variety deduction step0005 file name. Path = "
            + pszInputPath,
        )
    pszExpectedVarietyName = objFileNameMatch.group(1)
    iExpectedYear = int(objFileNameMatch.group(2))
    dicVarietyNumbers = {
        pszName: pszNumber for pszNumber, pszName in SPLIT_VARIETIES
    }
    pszExpectedVarietyNumber = dicVarietyNumbers[pszExpectedVarietyName]

    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0005_VARIETY_DEDUCTION_READ_FAILED",
            "failed to read variety deduction step0005 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    arrExpectedHeader = STEP0004_VARIETY_DEDUCTION_HEADER.split("\t")
    if len(arrRows) < 2:
        raise_conversion_error(
            "INVALID_STEP0005_VARIETY_DEDUCTION_ROW_COUNT",
            "variety deduction step0005 must contain one header and at least "
            "one data row. Path = " + pszInputPath,
        )
    if [pszValue.strip(" \u3000\t") for pszValue in arrRows[0]] != arrExpectedHeader:
        raise_conversion_error(
            "INVALID_STEP0005_VARIETY_DEDUCTION_HEADER",
            "unexpected variety deduction step0005 header. Path = "
            + pszInputPath,
        )

    iPreviousMonth = 0
    setMonths = set()
    for iLineNumber, arrFields in enumerate(arrRows[1:], start=2):
        if len(arrFields) != len(arrExpectedHeader):
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_DEDUCTION_COLUMN_COUNT",
                "variety deduction step0005 row must contain 10 columns. "
                "Path = " + pszInputPath + ", Line = " + str(iLineNumber),
            )
        iYear, iMonth = parse_year_month(
            arrFields[0], arrFields[1], iLineNumber
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_DEDUCTION_YEAR",
                "variety deduction step0005 year did not match its file name. "
                "Path = " + pszInputPath + ", Line = " + str(iLineNumber),
            )
        if iMonth in setMonths or iMonth <= iPreviousMonth:
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_DEDUCTION_MONTH_ORDER",
                "variety deduction step0005 months must be unique and "
                "ascending. Path = " + pszInputPath,
            )
        setMonths.add(iMonth)
        iPreviousMonth = iMonth
        if arrFields[2:4] != [
            pszExpectedVarietyNumber,
            pszExpectedVarietyName,
        ]:
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_DEDUCTION_VARIETY",
                "variety deduction step0005 variety did not match its file "
                "name. Path = " + pszInputPath + ", Line = " + str(iLineNumber),
            )
        for iColumnIndex, (pszItemName, _, bAllowBlank) in enumerate(
            STEP0004_VARIETY_DEDUCTION_ITEMS, start=4
        ):
            try:
                objValue = parse_decimal_value(
                    arrFields[iColumnIndex],
                    iLineNumber,
                    iColumnIndex,
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0005_VARIETY_DEDUCTION_ITEM_VALUE",
                    "invalid variety deduction step0005 item value. Path = "
                    + pszInputPath + ", Line = " + str(iLineNumber)
                    + ", Item = " + pszItemName
                    + ". Detail = " + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0005_VARIETY_DEDUCTION_ITEM_VALUE",
                    "variety deduction step0005 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iLineNumber)
                    + ", Item = " + pszItemName,
                )
    return arrRows


def create_step0006_variety_deduction_files(
    arrStep0005VarietyDeductionPaths: List[str],
) -> List[str]:
    """控除廃棄版の品種別step0005年間TSVを行列入れ替えしてstep0006を作成します。"""
    dicInputs: Dict[Tuple[int, str], str] = {}
    for pszPath in arrStep0005VarietyDeductionPaths:
        objMatch = re.search(
            r"_step0005_品種別_控除廃棄_" + VARIETY_NAME_PATTERN + r"_"
            r"(\d{4})_年間(?:_\d+)?\.tsv$",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error(
                "INVALID_STEP0005_VARIETY_DEDUCTION_FILENAME",
                "unexpected variety deduction step0005 file name. Path = "
                + pszPath,
            )
        objKey = (int(objMatch.group(2)), objMatch.group(1))
        if objKey in dicInputs:
            raise_conversion_error(
                "DUPLICATE_STEP0005_VARIETY_DEDUCTION_YEAR",
                "duplicate variety deduction step0005 year input. Key = "
                + str(objKey),
            )
        dicInputs[objKey] = pszPath

    setYears = {iYear for iYear, _ in dicInputs}
    setExpectedKeys = {
        (iYear, pszVarietyName)
        for iYear in setYears
        for _, pszVarietyName in SPLIT_VARIETIES
    }
    if set(dicInputs) != setExpectedKeys:
        raise_conversion_error(
            "INVALID_STEP0005_VARIETY_DEDUCTION_PATH_COUNT",
            "variety deduction step0005 paths must contain all five "
            "varieties for each year. Expected = " + str(len(setExpectedKeys))
            + ", Actual = " + str(len(dicInputs)),
        )

    arrCreatedPaths: List[str] = []
    try:
        for iYear in sorted(setYears):
            for _, pszVarietyName in SPLIT_VARIETIES:
                pszInputPath = dicInputs[(iYear, pszVarietyName)]
                arrInputRows = read_step0005_variety_deduction_for_step0006(
                    pszInputPath
                )
                arrTransposedRows = transpose_tsv_rows(arrInputRows)
                pszOutputPath = create_step0006_output_path(pszInputPath)
                write_step0006_tsv(pszOutputPath, arrTransposedRows)
                arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0006_VARIETY_DEDUCTION_CLEANUP_FAILED",
                "variety deduction step0006 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0007_output_path(pszSourceTsvPath: str, iYear: int) -> str:
    pszDirectory: str = os.path.dirname(pszSourceTsvPath)
    pszBaseName: str = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem: str = f"{pszBaseName}_step0007_グループ別_{iYear:04d}_年間"
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def parse_step0006_file_name(pszStep0006Path: str) -> Tuple[int, str]:
    objMatch = re.search(
        r"_step0006_グループ別_(契約|直営|フリー)_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszStep0006Path),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0006_FILENAME",
            "unexpected step0006 file name. Path = " + pszStep0006Path,
        )
    return int(objMatch.group(2)), objMatch.group(1)


def read_step0006_for_step0007(
    pszInputPath: str, iExpectedYear: int, pszExpectedGroupName: str
) -> Dict[str, List[str]]:
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0006_READ_FAILED",
            "failed to read step0006 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    arrExpectedRowHeaders: List[str] = STEP0004_HEADER.split("\t")
    if len(arrRows) != len(arrExpectedRowHeaders):
        raise_conversion_error(
            "INVALID_STEP0006_ROW_COUNT",
            "step0006 must contain 8 rows. Path = "
            + pszInputPath
            + ", Actual = "
            + str(len(arrRows)),
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount < 2 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0006_COLUMN_COUNT",
            "step0006 rows must have an equal column count and at least one data column. Path = "
            + pszInputPath,
        )
    arrActualRowHeaders: List[str] = [
        arrRow[0].strip(" \u3000\t") for arrRow in arrRows
    ]
    if arrActualRowHeaders != arrExpectedRowHeaders:
        raise_conversion_error(
            "INVALID_STEP0006_ROW_HEADER",
            "unexpected step0006 row headers. Path = " + pszInputPath,
        )
    dicRows: Dict[str, List[str]] = {
        arrRow[0].strip(" \u3000\t"): arrRow[1:] for arrRow in arrRows
    }
    dicExpectedGroups: Dict[str, str] = {
        pszGroupName: pszGroupNumber for pszGroupNumber, pszGroupName in GROUPS
    }
    pszExpectedGroupNumber: str = dicExpectedGroups[pszExpectedGroupName]
    arrMonths: List[int] = []
    for iValueIndex in range(iColumnCount - 1):
        iLineNumber: int = iValueIndex + 2
        iYear, iMonth = parse_year_month(
            dicRows["年"][iValueIndex], dicRows["月"][iValueIndex], iLineNumber
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0006_YEAR",
                "step0006 year did not match its file name. Path = "
                + pszInputPath
                + ", Column = "
                + str(iValueIndex + 2),
            )
        if iMonth in arrMonths:
            raise_conversion_error(
                "DUPLICATE_STEP0006_MONTH",
                "duplicate month in step0006. Path = "
                + pszInputPath
                + ", Month = "
                + str(iMonth),
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0006_MONTH",
                "step0006 months were not in ascending order. Path = "
                + pszInputPath,
            )
        arrMonths.append(iMonth)
        if (
            dicRows["ｸﾞﾙｰﾌﾟ№"][iValueIndex].strip(" \u3000\t")
            != pszExpectedGroupNumber
            or dicRows["ｸﾞﾙｰﾌﾟ名"][iValueIndex].strip(" \u3000\t")
            != pszExpectedGroupName
        ):
            raise_conversion_error(
                "INVALID_STEP0006_GROUP",
                "step0006 group did not match its file name. Path = "
                + pszInputPath
                + ", Column = "
                + str(iValueIndex + 2),
            )
        for pszItemName, bAllowBlank in STEP0004_ITEMS:
            try:
                objValue = parse_decimal_value(
                    dicRows[pszItemName][iValueIndex],
                    iLineNumber,
                    arrExpectedRowHeaders.index(pszItemName),
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0006_ITEM_VALUE",
                    "invalid step0006 item value. Path = "
                    + pszInputPath
                    + ", Item = "
                    + pszItemName
                    + ". Detail = "
                    + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0006_ITEM_VALUE",
                    "step0006 item value was blank. Path = "
                    + pszInputPath
                    + ", Item = "
                    + pszItemName
                    + ", Column = "
                    + str(iValueIndex + 2),
                )
    dicRows["月"] = [f"{iMonth:02d}" for iMonth in arrMonths]
    return dicRows


def format_step0007_average_weight(pszValue: str) -> str:
    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
    if pszNormalizedValue == "":
        return ""
    try:
        objValue = Decimal(pszNormalizedValue)
    except InvalidOperation:
        objValue = Decimal("NaN")
    if not objValue.is_finite():
        raise_conversion_error(
            "INVALID_STEP0006_ITEM_VALUE",
            "average weight was not a finite number. Value = " + pszValue,
        )
    return format(
        objValue.quantize(Decimal("0.001"), rounding=ROUND_HALF_UP), ".3f"
    )


def create_step0007_group_rows(
    pszGroupName: str, dicStep0006Rows: Dict[str, List[str]]
) -> List[str]:
    arrOutputRows: List[str] = []
    for iItemIndex, (pszItemName, _) in enumerate(STEP0004_ITEMS):
        arrValues: List[str] = list(dicStep0006Rows[pszItemName])
        if pszItemName == "平均体重":
            arrValues = [
                format_step0007_average_weight(pszValue) for pszValue in arrValues
            ]
        arrOutputRows.append(
            "\t".join(
                ([pszGroupName if iItemIndex == 0 else "", pszItemName] + arrValues)
            )
        )
    return arrOutputRows


def write_step0007_tsv(pszOutputPath: str, arrRows: List[str]) -> None:
    if not arrRows:
        raise_conversion_error(
            "INVALID_STEP0007_ROW_COUNT", "step0007 did not contain any rows."
        )
    try:
        write_monthly_tsv(pszOutputPath, arrRows[0], arrRows[1:])
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0007_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0007_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0007_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0007_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0007_WRITE_FAILED"),
            str(objException),
        )


def create_step0007_files(
    pszSourceTsvPath: str, arrStep0006Paths: List[str]
) -> Tuple[List[str], List[int]]:
    dicPaths: Dict[Tuple[int, str], str] = {}
    for pszPath in arrStep0006Paths:
        iYear, pszGroupName = parse_step0006_file_name(pszPath)
        objKey: Tuple[int, str] = (iYear, pszGroupName)
        if objKey in dicPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0006_YEAR_GROUP",
                "duplicate step0006 year/group. Key = " + str(objKey),
            )
        dicPaths[objKey] = pszPath
    arrYears: List[int] = sorted({iYear for iYear, _ in dicPaths})
    arrCreatedPaths: List[str] = []
    try:
        for iYear in arrYears:
            dicGroupRows: Dict[str, Dict[str, List[str]]] = {}
            for _, pszGroupName in GROUPS:
                objKey = (iYear, pszGroupName)
                if objKey not in dicPaths:
                    raise_conversion_error(
                        "MISSING_STEP0006_GROUP",
                        "required step0006 group was missing. Key = " + str(objKey),
                    )
                dicGroupRows[pszGroupName] = read_step0006_for_step0007(
                    dicPaths[objKey], iYear, pszGroupName
                )
            arrReferenceMonths: List[str] = dicGroupRows[GROUPS[0][1]]["月"]
            for _, pszGroupName in GROUPS[1:]:
                if dicGroupRows[pszGroupName]["月"] != arrReferenceMonths:
                    raise_conversion_error(
                        "STEP0006_MONTH_MISMATCH",
                        "step0006 month columns did not match between groups. Year = "
                        + str(iYear)
                        + ", Group = "
                        + pszGroupName,
                    )
            iMonthCount: int = len(arrReferenceMonths)
            arrRows: List[str] = [
                "\t".join(["", "年"] + [f"{iYear:04d}"] * iMonthCount),
                "\t".join(["", "月"] + arrReferenceMonths),
            ]
            for _, pszGroupName in GROUPS:
                arrRows.extend(
                    create_step0007_group_rows(
                        pszGroupName, dicGroupRows[pszGroupName]
                    )
                )
            arrAdditionalMonths: List[str] = [
                f"{iMonth:02d}" for iMonth in range(int(arrReferenceMonths[-1]) + 1, 13)
            ]
            if arrAdditionalMonths:
                dicFillValues: Dict[str, str] = {
                    "入雛羽数": "0",
                    "出荷羽数": "0",
                    "正味重量": "0.0",
                    "平均体重": "0.000",
                }
                for iRowIndex, pszRow in enumerate(arrRows):
                    pszItemName: str = pszRow.split("\t")[1]
                    if pszItemName == "年":
                        arrAdditionalValues = [f"{iYear:04d}"] * len(arrAdditionalMonths)
                    elif pszItemName == "月":
                        arrAdditionalValues = arrAdditionalMonths
                    else:
                        arrAdditionalValues = [dicFillValues[pszItemName]] * len(arrAdditionalMonths)
                    arrRows[iRowIndex] = pszRow + "\t" + "\t".join(arrAdditionalValues)
            pszOutputPath: str = create_step0007_output_path(
                pszSourceTsvPath, iYear
            )
            write_step0007_tsv(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, arrYears
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0007_CLEANUP_FAILED",
                "step0007 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0007_variety_annual_output_path(
    pszSourceTsvPath: str, iYear: int
) -> str:
    pszDirectory = os.path.dirname(pszSourceTsvPath)
    pszBaseName = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem = f"{pszBaseName}_step0007_品種別_{iYear:04d}_年間"
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def read_step0006_variety_annual_for_step0007(
    pszInputPath: str, iExpectedYear: int, pszExpectedVarietyName: str
) -> Dict[str, List[str]]:
    objMatch = re.search(
        r"_step0006_品種別_" + VARIETY_NAME_PATTERN + r"_"
        r"(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objMatch is None or objMatch.group(1) != pszExpectedVarietyName:
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_FILENAME",
            "unexpected variety step0006 file name. Path = " + pszInputPath,
        )
    if int(objMatch.group(2)) != iExpectedYear:
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_YEAR",
            "variety step0006 year did not match the expected year. Path = "
            + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0006_VARIETY_READ_FAILED",
            "failed to read variety step0006 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    arrExpectedHeaders = STEP0004_VARIETY_HEADER.split("\t")
    if len(arrRows) != len(arrExpectedHeaders):
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_ROW_COUNT",
            "variety step0006 must contain 10 rows. Path = " + pszInputPath,
        )
    iColumnCount = len(arrRows[0])
    if iColumnCount < 2 or any(
        len(arrRow) != iColumnCount for arrRow in arrRows
    ):
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_COLUMN_COUNT",
            "variety step0006 rows must have equal columns and at least one "
            "month. Path = " + pszInputPath,
        )
    arrActualHeaders = [
        arrRow[0].strip(" \u3000\t") for arrRow in arrRows
    ]
    if arrActualHeaders != arrExpectedHeaders:
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_ROW_HEADER",
            "unexpected variety step0006 row headers. Path = " + pszInputPath,
        )
    dicRows = {
        arrRow[0].strip(" \u3000\t"): arrRow[1:] for arrRow in arrRows
    }
    dicVarietyNumbers = {
        pszName: pszNumber for pszNumber, pszName in SPLIT_VARIETIES
    }
    pszExpectedVarietyNumber = dicVarietyNumbers[pszExpectedVarietyName]
    arrMonths: List[int] = []
    for iValueIndex in range(iColumnCount - 1):
        iLineNumber = iValueIndex + 2
        iYear, iMonth = parse_year_month(
            dicRows["年"][iValueIndex], dicRows["月"][iValueIndex], iLineNumber
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0006_VARIETY_YEAR",
                "variety step0006 data year did not match the file name. "
                "Path = " + pszInputPath + ", Column = "
                + str(iValueIndex + 2),
            )
        if iMonth in arrMonths or (arrMonths and iMonth <= arrMonths[-1]):
            raise_conversion_error(
                "INVALID_STEP0006_VARIETY_MONTH_ORDER",
                "variety step0006 months must be unique and ascending. Path = "
                + pszInputPath,
            )
        arrMonths.append(iMonth)
        if (
            dicRows["品種No."][iValueIndex].strip(" \u3000\t")
            != pszExpectedVarietyNumber
            or dicRows["品種名"][iValueIndex].strip(" \u3000\t")
            != pszExpectedVarietyName
        ):
            raise_conversion_error(
                "INVALID_STEP0006_VARIETY_IDENTITY",
                "variety step0006 identity did not match its file name. Path = "
                + pszInputPath + ", Column = " + str(iValueIndex + 2),
            )
        for iItemIndex, (pszItemName, _, bAllowBlank) in enumerate(
            STEP0004_VARIETY_ITEMS, start=4
        ):
            try:
                objValue = parse_decimal_value(
                    dicRows[pszItemName][iValueIndex],
                    iLineNumber,
                    iItemIndex,
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0006_VARIETY_ITEM_VALUE",
                    "invalid variety step0006 item value. Path = "
                    + pszInputPath + ", Item = " + pszItemName
                    + ". Detail = " + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0006_VARIETY_ITEM_VALUE",
                    "variety step0006 item value was blank. Path = "
                    + pszInputPath + ", Item = " + pszItemName
                    + ", Column = " + str(iValueIndex + 2),
                )
    dicRows["月"] = [f"{iMonth:02d}" for iMonth in arrMonths]
    return dicRows


def create_step0007_variety_annual_rows(
    pszVarietyName: str, dicStep0006Rows: Dict[str, List[str]]
) -> List[str]:
    arrOutputRows: List[str] = []
    for iItemIndex, (pszItemName, _, _) in enumerate(STEP0004_VARIETY_ITEMS):
        arrValues = list(dicStep0006Rows[pszItemName])
        if pszItemName == "平均体重":
            arrValues = [
                format_step0007_average_weight(pszValue)
                for pszValue in arrValues
            ]
        arrOutputRows.append(
            "\t".join(
                [pszVarietyName if iItemIndex == 0 else "", pszItemName]
                + arrValues
            )
        )
    return arrOutputRows


def create_step0007_variety_annual_files(
    pszSourceTsvPath: str, arrStep0006VarietyPaths: List[str]
) -> Tuple[List[str], List[int]]:
    """品種別step0006を照合して、年ごとの品種別step0007表を作成します。"""
    dicPaths: Dict[Tuple[int, str], str] = {}
    for pszPath in arrStep0006VarietyPaths:
        objMatch = re.search(
            r"_step0006_品種別_" + VARIETY_NAME_PATTERN + r"_"
            r"(\d{4})_年間(?:_\d{2})?\.tsv$",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error(
                "INVALID_STEP0006_VARIETY_FILENAME",
                "unexpected variety step0006 file name. Path = " + pszPath,
            )
        objKey = (int(objMatch.group(2)), objMatch.group(1))
        if objKey in dicPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0006_VARIETY_YEAR",
                "duplicate step0006 year/variety. Key = " + str(objKey),
            )
        dicPaths[objKey] = pszPath

    arrYears = sorted({iYear for iYear, _ in dicPaths})
    setExpectedKeys = {
        (iYear, pszVarietyName)
        for iYear in arrYears
        for _, pszVarietyName in SPLIT_VARIETIES
    }
    if set(dicPaths) != setExpectedKeys:
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_PATH_COUNT",
            "step0006 variety paths must contain all five varieties for each "
            "year. Expected = " + str(len(setExpectedKeys))
            + ", Actual = " + str(len(dicPaths)),
        )

    arrCreatedPaths: List[str] = []
    try:
        for iYear in arrYears:
            dicVarietyRows: Dict[str, Dict[str, List[str]]] = {}
            for _, pszVarietyName in SPLIT_VARIETIES:
                pszInputPath = dicPaths[(iYear, pszVarietyName)]
                dicVarietyRows[pszVarietyName] = (
                    read_step0006_variety_annual_for_step0007(
                        pszInputPath, iYear, pszVarietyName
                    )
                )
            arrReferenceMonths = dicVarietyRows[SPLIT_VARIETIES[0][1]]["月"]
            for _, pszVarietyName in SPLIT_VARIETIES[1:]:
                if dicVarietyRows[pszVarietyName]["月"] != arrReferenceMonths:
                    raise_conversion_error(
                        "STEP0006_VARIETY_MONTH_MISMATCH",
                        "step0006 month columns did not match between "
                        "varieties. Year = " + str(iYear)
                        + ", Variety = " + pszVarietyName,
                    )

            iMonthCount = len(arrReferenceMonths)
            arrRows = [
                "\t".join(["", "年"] + [f"{iYear:04d}"] * iMonthCount),
                "\t".join(["", "月"] + arrReferenceMonths),
            ]
            for _, pszVarietyName in SPLIT_VARIETIES:
                arrRows.extend(
                    create_step0007_variety_annual_rows(
                        pszVarietyName, dicVarietyRows[pszVarietyName]
                    )
                )

            arrAdditionalMonths = [
                f"{iMonth:02d}"
                for iMonth in range(int(arrReferenceMonths[-1]) + 1, 13)
            ]
            if arrAdditionalMonths:
                dicFillValues = {
                    "入雛羽数": "0",
                    "出荷羽数": "0",
                    "正味重量": "0.0",
                    "平均体重": "0.000",
                    "飼育日数平均": "0",
                    "廃棄控除前育成率": "0",
                }
                for iRowIndex, pszRow in enumerate(arrRows):
                    pszItemName = pszRow.split("\t")[1]
                    if pszItemName == "年":
                        arrAdditionalValues = [
                            f"{iYear:04d}"
                        ] * len(arrAdditionalMonths)
                    elif pszItemName == "月":
                        arrAdditionalValues = arrAdditionalMonths
                    else:
                        arrAdditionalValues = [dicFillValues[pszItemName]] * len(
                            arrAdditionalMonths
                        )
                    arrRows[iRowIndex] += "\t" + "\t".join(arrAdditionalValues)

            pszOutputPath = create_step0007_variety_annual_output_path(
                pszSourceTsvPath, iYear
            )
            write_step0007_tsv(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, arrYears
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0007_VARIETY_CLEANUP_FAILED",
                "variety step0007 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0007_variety_deduction_annual_output_path(
    pszSourceTsvPath: str, iYear: int
) -> str:
    pszDirectory = os.path.dirname(pszSourceTsvPath)
    pszBaseName = os.path.splitext(os.path.basename(pszSourceTsvPath))[0]
    pszStem = (
        f"{pszBaseName}_step0007_品種別_控除廃棄_{iYear:04d}_年間"
    )
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".tsv")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.tsv"
        )
        iSuffix += 1
    return pszOutputPath


def read_step0006_variety_deduction_annual_for_step0007(
    pszInputPath: str, iExpectedYear: int, pszExpectedVarietyName: str
) -> Dict[str, List[str]]:
    objMatch = re.search(
        r"_step0006_品種別_控除廃棄_" + VARIETY_NAME_PATTERN + r"_"
        r"(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objMatch is None or objMatch.group(1) != pszExpectedVarietyName:
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_DEDUCTION_FILENAME",
            "unexpected variety deduction step0006 file name. Path = "
            + pszInputPath,
        )
    if int(objMatch.group(2)) != iExpectedYear:
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_DEDUCTION_YEAR",
            "variety deduction step0006 year did not match the expected year. "
            "Path = " + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0006_VARIETY_DEDUCTION_READ_FAILED",
            "failed to read variety deduction step0006 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    arrExpectedHeaders = STEP0004_VARIETY_DEDUCTION_HEADER.split("\t")
    if len(arrRows) != len(arrExpectedHeaders):
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_DEDUCTION_ROW_COUNT",
            "variety deduction step0006 must contain 10 rows. Path = "
            + pszInputPath,
        )
    iColumnCount = len(arrRows[0])
    if iColumnCount < 2 or any(
        len(arrRow) != iColumnCount for arrRow in arrRows
    ):
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_DEDUCTION_COLUMN_COUNT",
            "variety deduction step0006 rows must have equal columns and at "
            "least one month. Path = " + pszInputPath,
        )
    arrActualHeaders = [
        arrRow[0].strip(" \u3000\t") for arrRow in arrRows
    ]
    if arrActualHeaders != arrExpectedHeaders:
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_DEDUCTION_ROW_HEADER",
            "unexpected variety deduction step0006 row headers. Path = "
            + pszInputPath,
        )
    dicRows = {
        arrRow[0].strip(" \u3000\t"): arrRow[1:] for arrRow in arrRows
    }
    dicVarietyNumbers = {
        pszName: pszNumber for pszNumber, pszName in SPLIT_VARIETIES
    }
    pszExpectedVarietyNumber = dicVarietyNumbers[pszExpectedVarietyName]
    arrMonths: List[int] = []
    for iValueIndex in range(iColumnCount - 1):
        iLineNumber = iValueIndex + 2
        iYear, iMonth = parse_year_month(
            dicRows["年"][iValueIndex], dicRows["月"][iValueIndex], iLineNumber
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0006_VARIETY_DEDUCTION_YEAR",
                "variety deduction step0006 data year did not match its file "
                "name. Path = " + pszInputPath + ", Column = "
                + str(iValueIndex + 2),
            )
        if iMonth in arrMonths or (arrMonths and iMonth <= arrMonths[-1]):
            raise_conversion_error(
                "INVALID_STEP0006_VARIETY_DEDUCTION_MONTH_ORDER",
                "variety deduction step0006 months must be unique and "
                "ascending. Path = " + pszInputPath,
            )
        arrMonths.append(iMonth)
        if (
            dicRows["品種No."][iValueIndex].strip(" \u3000\t")
            != pszExpectedVarietyNumber
            or dicRows["品種名"][iValueIndex].strip(" \u3000\t")
            != pszExpectedVarietyName
        ):
            raise_conversion_error(
                "INVALID_STEP0006_VARIETY_DEDUCTION_IDENTITY",
                "variety deduction step0006 identity did not match its file "
                "name. Path = " + pszInputPath + ", Column = "
                + str(iValueIndex + 2),
            )
        for iItemIndex, (pszItemName, _, bAllowBlank) in enumerate(
            STEP0004_VARIETY_DEDUCTION_ITEMS, start=4
        ):
            try:
                objValue = parse_decimal_value(
                    dicRows[pszItemName][iValueIndex],
                    iLineNumber,
                    iItemIndex,
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0006_VARIETY_DEDUCTION_ITEM_VALUE",
                    "invalid variety deduction step0006 item value. Path = "
                    + pszInputPath + ", Item = " + pszItemName
                    + ". Detail = " + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0006_VARIETY_DEDUCTION_ITEM_VALUE",
                    "variety deduction step0006 item value was blank. Path = "
                    + pszInputPath + ", Item = " + pszItemName
                    + ", Column = " + str(iValueIndex + 2),
                )
    dicRows["月"] = [f"{iMonth:02d}" for iMonth in arrMonths]
    return dicRows


def create_step0007_variety_deduction_annual_rows(
    pszVarietyName: str, dicStep0006Rows: Dict[str, List[str]]
) -> List[str]:
    arrOutputRows: List[str] = []
    for iItemIndex, (pszItemName, _, _) in enumerate(
        STEP0004_VARIETY_DEDUCTION_ITEMS
    ):
        arrValues = list(dicStep0006Rows[pszItemName])
        if pszItemName == "控除後平均体重":
            arrValues = [
                format_step0007_average_weight(pszValue)
                for pszValue in arrValues
            ]
        arrOutputRows.append(
            "\t".join(
                [pszVarietyName if iItemIndex == 0 else "", pszItemName]
                + arrValues
            )
        )
    return arrOutputRows


def create_step0007_variety_deduction_annual_files(
    pszSourceTsvPath: str, arrStep0006VarietyDeductionPaths: List[str]
) -> Tuple[List[str], List[int]]:
    """控除廃棄版の品種別step0006を照合し、年ごとのstep0007を作成します。"""
    dicPaths: Dict[Tuple[int, str], str] = {}
    for pszPath in arrStep0006VarietyDeductionPaths:
        objMatch = re.search(
            r"_step0006_品種別_控除廃棄_" + VARIETY_NAME_PATTERN + r"_"
            r"(\d{4})_年間(?:_\d{2})?\.tsv$",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error(
                "INVALID_STEP0006_VARIETY_DEDUCTION_FILENAME",
                "unexpected variety deduction step0006 file name. Path = "
                + pszPath,
            )
        objKey = (int(objMatch.group(2)), objMatch.group(1))
        if objKey in dicPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0006_VARIETY_DEDUCTION_YEAR",
                "duplicate step0006 year/variety. Key = " + str(objKey),
            )
        dicPaths[objKey] = pszPath

    arrYears = sorted({iYear for iYear, _ in dicPaths})
    setExpectedKeys = {
        (iYear, pszVarietyName)
        for iYear in arrYears
        for _, pszVarietyName in SPLIT_VARIETIES
    }
    if set(dicPaths) != setExpectedKeys:
        raise_conversion_error(
            "INVALID_STEP0006_VARIETY_DEDUCTION_PATH_COUNT",
            "variety deduction step0006 paths must contain all five varieties "
            "for each year. Expected = " + str(len(setExpectedKeys))
            + ", Actual = " + str(len(dicPaths)),
        )

    arrCreatedPaths: List[str] = []
    try:
        for iYear in arrYears:
            dicVarietyRows: Dict[str, Dict[str, List[str]]] = {}
            for _, pszVarietyName in SPLIT_VARIETIES:
                pszInputPath = dicPaths[(iYear, pszVarietyName)]
                dicVarietyRows[pszVarietyName] = (
                    read_step0006_variety_deduction_annual_for_step0007(
                        pszInputPath, iYear, pszVarietyName
                    )
                )
            arrReferenceMonths = dicVarietyRows[SPLIT_VARIETIES[0][1]]["月"]
            for _, pszVarietyName in SPLIT_VARIETIES[1:]:
                if dicVarietyRows[pszVarietyName]["月"] != arrReferenceMonths:
                    raise_conversion_error(
                        "STEP0006_VARIETY_DEDUCTION_MONTH_MISMATCH",
                        "step0006 month columns did not match between "
                        "deduction varieties. Year = " + str(iYear)
                        + ", Variety = " + pszVarietyName,
                    )

            iMonthCount = len(arrReferenceMonths)
            arrRows = [
                "\t".join(["", "年"] + [f"{iYear:04d}"] * iMonthCount),
                "\t".join(["", "月"] + arrReferenceMonths),
            ]
            for _, pszVarietyName in SPLIT_VARIETIES:
                arrRows.extend(
                    create_step0007_variety_deduction_annual_rows(
                        pszVarietyName, dicVarietyRows[pszVarietyName]
                    )
                )

            arrAdditionalMonths = [
                f"{iMonth:02d}"
                for iMonth in range(int(arrReferenceMonths[-1]) + 1, 13)
            ]
            if arrAdditionalMonths:
                dicFillValues = {
                    "入雛羽数": "0",
                    "処理羽数": "0",
                    "処理重量": "0.0",
                    "控除後平均体重": "0.000",
                    "飼育日数平均": "0",
                    "廃棄控除後育成率": "0",
                }
                for iRowIndex, pszRow in enumerate(arrRows):
                    pszItemName = pszRow.split("\t")[1]
                    if pszItemName == "年":
                        arrAdditionalValues = [
                            f"{iYear:04d}"
                        ] * len(arrAdditionalMonths)
                    elif pszItemName == "月":
                        arrAdditionalValues = arrAdditionalMonths
                    else:
                        arrAdditionalValues = [dicFillValues[pszItemName]] * len(
                            arrAdditionalMonths
                        )
                    arrRows[iRowIndex] += "\t" + "\t".join(arrAdditionalValues)

            pszOutputPath = create_step0007_variety_deduction_annual_output_path(
                pszSourceTsvPath, iYear
            )
            write_step0007_tsv(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths, arrYears
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0007_VARIETY_DEDUCTION_CLEANUP_FAILED",
                "variety deduction step0007 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0007_excel_output_path(pszStep0007TsvPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0007TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0007TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def read_step0007_for_excel(
    pszInputPath: str,
) -> Tuple[List[List[str]], int]:
    objFileNameMatch = re.search(
        r"_step0007_グループ別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0007_FILENAME",
            "unexpected step0007 file name. Path = " + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0007_TSV_READ_FAILED",
            "failed to read step0007 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    arrExpectedStructure: List[Tuple[str, str]] = [
        ("", "年"),
        ("", "月"),
        ("契約", "入雛羽数"),
        ("", "出荷羽数"),
        ("", "正味重量"),
        ("", "平均体重"),
        ("直営", "入雛羽数"),
        ("", "出荷羽数"),
        ("", "正味重量"),
        ("", "平均体重"),
        ("フリー", "入雛羽数"),
        ("", "出荷羽数"),
        ("", "正味重量"),
        ("", "平均体重"),
    ]
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0007_ROW_COUNT",
            "step0007 must contain 14 rows. Path = "
            + pszInputPath
            + ", Actual = "
            + str(len(arrRows)),
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount < 3 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0007_COLUMN_COUNT",
            "step0007 rows must have an equal column count and at least one data column. Path = "
            + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        objActual: Tuple[str, str] = (
            arrRows[iRowIndex][0].strip(" \u3000\t"),
            arrRows[iRowIndex][1].strip(" \u3000\t"),
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0007_ROW_STRUCTURE",
                "unexpected step0007 row structure. Path = "
                + pszInputPath
                + ", Line = "
                + str(iRowIndex + 1),
            )

    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex],
            arrRows[1][iColumnIndex],
            iColumnIndex + 1,
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0007_YEAR",
                "step0007 year did not match its file name. Path = "
                + pszInputPath
                + ", Column = "
                + str(iColumnIndex + 1),
            )
        if iMonth in arrMonths or (arrMonths and iMonth <= arrMonths[-1]):
            raise_conversion_error(
                "INVALID_STEP0007_MONTH",
                "step0007 months must be unique and ascending. Path = "
                + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            pszItemName: str = arrRows[iRowIndex][1].strip(" \u3000\t")
            bAllowBlank: bool = pszItemName == "平均体重"
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1,
                    iColumnIndex,
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0007_ITEM_VALUE",
                    "invalid step0007 item value. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iRowIndex + 1)
                    + ", Column = "
                    + str(iColumnIndex + 1)
                    + ". Detail = "
                    + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0007_ITEM_VALUE",
                    "step0007 item value was blank. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iRowIndex + 1)
                    + ", Column = "
                    + str(iColumnIndex + 1),
                )
    return arrRows, iExpectedYear


def convert_step0007_excel_numeric_value(pszValue: str) -> Optional[Any]:
    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
    if pszNormalizedValue == "":
        return None
    try:
        objValue = Decimal(pszNormalizedValue)
    except InvalidOperation:
        objValue = Decimal("NaN")
    if not objValue.is_finite():
        raise_conversion_error(
            "INVALID_STEP0007_ITEM_VALUE",
            "step0007 value was not a finite number. Value = " + pszValue,
        )
    if objValue == objValue.to_integral_value():
        return int(objValue)
    return float(objValue)


def write_step0007_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0007_EXCEL_CREATE_FAILED",
                "step0007 Excel already exists. Path = " + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0007_EXCEL_CREATE_FAILED",
                "failed to reserve step0007 Excel. Detail = " + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0007_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "グループ別"
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2:
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    objCell.value = convert_step0007_excel_numeric_value(pszValue)
                    if objCell.value is None:
                        continue
                    if pszRowType == "年":
                        objCell.number_format = "0"
                    elif pszRowType == "月":
                        objCell.number_format = "00"
                    elif pszRowType in {"入雛羽数", "出荷羽数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "正味重量":
                        objCell.number_format = "0.##########"
                    elif pszRowType == "平均体重":
                        objCell.number_format = "0.000"
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except ConversionError:
            raise
        except Exception as objException:
            raise_conversion_error(
                "STEP0007_EXCEL_WRITE_FAILED",
                "failed to write step0007 Excel. Detail = " + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0007_EXCEL_FINALIZE_FAILED",
                "failed to finalize step0007 Excel. Detail = " + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0007_excel_files(
    arrStep0007TsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0007TsvPath in arrStep0007TsvPaths:
            arrRows, _ = read_step0007_for_excel(pszStep0007TsvPath)
            pszOutputPath: str = create_step0007_excel_output_path(
                pszStep0007TsvPath
            )
            write_step0007_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0007_EXCEL_CLEANUP_FAILED",
                "step0007 Excel creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0008_output_path(pszStep0007Path: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0007Path)
    pszFileName: str = os.path.basename(pszStep0007Path)
    if "_step0007_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0007_FOR_STEP0008",
            "step0007 marker was not found. Path = " + pszStep0007Path,
        )
    pszStep0008FileName: str = pszFileName.replace(
        "_step0007_", "_step0008_", 1
    )
    pszStem, pszExtension = os.path.splitext(pszStep0008FileName)
    pszOutputPath: str = os.path.join(pszDirectory, pszStep0008FileName)
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def calculate_growth_rate(
    pszPlacementCount: str, pszShippingCount: str
) -> Tuple[str, bool]:
    try:
        objPlacementCount = Decimal(pszPlacementCount.strip(" \u3000\t"))
        objShippingCount = Decimal(pszShippingCount.strip(" \u3000\t"))
    except InvalidOperation:
        raise_conversion_error(
            "INVALID_GROWTH_RATE_SOURCE_VALUE",
            "placement or shipping count was not numeric.",
        )
    if not objPlacementCount.is_finite() or not objShippingCount.is_finite():
        raise_conversion_error(
            "INVALID_GROWTH_RATE_SOURCE_VALUE",
            "placement or shipping count was not finite.",
        )
    if objPlacementCount < 0 or objShippingCount < 0:
        raise_conversion_error(
            "INVALID_GROWTH_RATE_SOURCE_VALUE",
            "placement and shipping counts must not be negative.",
        )
    if objPlacementCount == 0:
        return "", True
    objRate: Decimal = objShippingCount / objPlacementCount
    return (
        format(
            objRate.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP),
            "7.4f",
        ),
        False,
    )


def write_step0008_tsv(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    if not arrRows:
        raise_conversion_error(
            "INVALID_STEP0007_FOR_STEP0008",
            "step0008 did not contain any rows.",
        )
    try:
        write_monthly_tsv(
            pszOutputPath,
            "\t".join(arrRows[0]),
            ["\t".join(arrRow) for arrRow in arrRows[1:]],
        )
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0008_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0008_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0008_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0008_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0008_WRITE_FAILED"),
            str(objException),
        )


def format_growth_rate_warning(
    iWarningNumber: int, dicWarning: Dict[str, str]
) -> str:
    if "variety_name" in dicWarning:
        pszTargetInfo = (
            "品種No.: " + dicWarning["variety_number"] + "\r\n"
            + "品種名: " + dicWarning["variety_name"] + "\r\n"
        )
    else:
        pszTargetInfo = (
            "グループ№: " + dicWarning["group_number"] + "\r\n"
            + "グループ名: " + dicWarning["group_name"] + "\r\n"
        )
    return (
        "\r\n[警告 "
        + str(iWarningNumber)
        + "]\r\n"
        + "警告種別: "
        + dicWarning["type"]
        + "\r\n入力step0007: "
        + dicWarning["input"]
        + "\r\n出力step0008: "
        + dicWarning["output"]
        + "\r\n処理年月: "
        + dicWarning["month"]
        + "\r\n"
        + pszTargetInfo
        + "\r\n対象列: 入雛羽数, 出荷羽数\r\n"
        + "警告理由: 入雛羽数が0のため育成率を計算できません。\r\n"
        + "出力値: empty\r\n"
    )


def write_growth_rate_warnings(
    pszInputExcelPath: str,
    pszExistingWarningPath: str,
    iExistingWarningCount: int,
    arrWarnings: List[Dict[str, str]],
) -> str:
    if not arrWarnings:
        return pszExistingWarningPath
    iTotalWarningCount: int = iExistingWarningCount + len(arrWarnings)
    if pszExistingWarningPath:
        pszTemporaryPath: str = ""
        try:
            with open(
                pszExistingWarningPath, "r", encoding="utf-8", newline=""
            ) as objFile:
                pszExistingText: str = objFile.read()
            objCountMatch = re.search(r"警告件数: (\d+)", pszExistingText)
            if objCountMatch is None:
                raise_conversion_error(
                    "WARNING_FILE_UPDATE_FAILED",
                    "warning count was not found in existing warning file. Path = "
                    + pszExistingWarningPath,
                )
            pszExistingText = (
                pszExistingText[: objCountMatch.start()]
                + "警告件数: "
                + str(iTotalWarningCount)
                + pszExistingText[objCountMatch.end() :]
            )
            for iIndex, dicWarning in enumerate(
                arrWarnings, start=iExistingWarningCount + 1
            ):
                pszExistingText += format_growth_rate_warning(iIndex, dicWarning)
            iFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_warning_",
                suffix=".txt",
                dir=os.path.dirname(pszExistingWarningPath),
            )
            os.close(iFileDescriptor)
            with open(
                pszTemporaryPath, "w", encoding="utf-8", newline=""
            ) as objFile:
                objFile.write(pszExistingText)
            os.replace(pszTemporaryPath, pszExistingWarningPath)
            pszTemporaryPath = ""
            return pszExistingWarningPath
        except ConversionError:
            raise
        except Exception as objException:
            raise_conversion_error(
                "WARNING_FILE_UPDATE_FAILED",
                "failed to update warning file. Detail = " + str(objException),
            )
        finally:
            if pszTemporaryPath and os.path.exists(pszTemporaryPath):
                os.remove(pszTemporaryPath)

    pszWarningPath: str = create_warning_path(pszInputExcelPath)
    try:
        with open(pszWarningPath, "x", encoding="utf-8", newline="") as objFile:
            objFile.write("育成率追加が警告付きで完了しました。\r\n")
            objFile.write("入力Excel: " + pszInputExcelPath + "\r\n")
            objFile.write("警告件数: " + str(len(arrWarnings)) + "\r\n")
            for iIndex, dicWarning in enumerate(arrWarnings, start=1):
                objFile.write(format_growth_rate_warning(iIndex, dicWarning))
        return pszWarningPath
    except Exception as objException:
        if os.path.exists(pszWarningPath):
            os.remove(pszWarningPath)
        raise_conversion_error(
            "WARNING_FILE_CREATE_FAILED",
            "failed to create growth-rate warning file. Detail = "
            + str(objException),
        )


def create_step0008_files(
    arrStep0007Paths: List[str],
    pszInputExcelPath: str,
    pszExistingWarningPath: str,
    iExistingWarningCount: int,
) -> Tuple[List[str], str, int]:
    arrCreatedPaths: List[str] = []
    arrWarnings: List[Dict[str, str]] = []
    try:
        for pszStep0007Path in arrStep0007Paths:
            arrInputRows, iYear = read_step0007_for_excel(pszStep0007Path)
            pszOutputPath: str = create_step0008_output_path(pszStep0007Path)
            arrOutputRows: List[List[str]] = [
                list(arrInputRows[0]), list(arrInputRows[1])
            ]
            arrMonths: List[str] = arrInputRows[1][2:]
            for iGroupIndex, (pszGroupNumber, pszGroupName) in enumerate(GROUPS):
                iGroupStart: int = 2 + iGroupIndex * 4
                arrGroupRows: List[List[str]] = [
                    list(arrRow)
                    for arrRow in arrInputRows[iGroupStart : iGroupStart + 4]
                ]
                arrGrowthRateRow: List[str] = ["", "育成率"]
                for iValueIndex, pszMonth in enumerate(arrMonths, start=2):
                    pszGrowthRate, bDivisionByZero = calculate_growth_rate(
                        arrGroupRows[0][iValueIndex],
                        arrGroupRows[1][iValueIndex],
                    )
                    arrGrowthRateRow.append(pszGrowthRate)
                    if bDivisionByZero:
                        arrWarnings.append(
                            {
                                "type": "GROWTH_RATE_DIVISION_BY_ZERO",
                                "input": pszStep0007Path,
                                "output": pszOutputPath,
                                "month": f"{iYear:04d}{int(pszMonth):02d}",
                                "group_number": pszGroupNumber,
                                "group_name": pszGroupName,
                            }
                        )
                arrOutputRows.extend(arrGroupRows)
                arrOutputRows.append(arrGrowthRateRow)
            write_step0008_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)
        pszWarningPath: str = write_growth_rate_warnings(
            pszInputExcelPath,
            pszExistingWarningPath,
            iExistingWarningCount,
            arrWarnings,
        )
        return (
            arrCreatedPaths,
            pszWarningPath,
            iExistingWarningCount + len(arrWarnings),
        )
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0008_CLEANUP_FAILED",
                "step0008 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0007_variety_annual_for_step0008(
    pszInputPath: str,
) -> Tuple[List[List[str]], int]:
    objFileNameMatch = re.search(
        r"_step0007_品種別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0007_VARIETY_FILENAME",
            "unexpected variety step0007 file name. Path = " + pszInputPath,
        )
    iExpectedYear = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0007_VARIETY_READ_FAILED",
            "failed to read variety step0007 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for _, pszVarietyName in SPLIT_VARIETIES:
        arrExpectedStructure.extend(
            (pszVarietyName if iItemIndex == 0 else "", pszItemName)
            for iItemIndex, (pszItemName, _, _) in enumerate(
                STEP0004_VARIETY_ITEMS
            )
        )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0007_VARIETY_ROW_COUNT",
            "variety step0007 must contain 32 rows. Path = " + pszInputPath,
        )
    iColumnCount = len(arrRows[0])
    if iColumnCount < 3 or any(
        len(arrRow) != iColumnCount for arrRow in arrRows
    ):
        raise_conversion_error(
            "INVALID_STEP0007_VARIETY_COLUMN_COUNT",
            "variety step0007 rows must have equal columns and at least one "
            "data month. Path = " + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0007_VARIETY_ROW_STRUCTURE",
                "unexpected variety step0007 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )

    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex],
            arrRows[1][iColumnIndex],
            iColumnIndex + 1,
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0007_VARIETY_YEAR",
                "variety step0007 year did not match its file name. Path = "
                + pszInputPath + ", Column = " + str(iColumnIndex + 1),
            )
        if iMonth in arrMonths or (arrMonths and iMonth <= arrMonths[-1]):
            raise_conversion_error(
                "INVALID_STEP0007_VARIETY_MONTH",
                "variety step0007 months must be unique and ascending. Path = "
                + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iVarietyIndex, (_, pszVarietyName) in enumerate(SPLIT_VARIETIES):
            iVarietyStart = 2 + iVarietyIndex * len(STEP0004_VARIETY_ITEMS)
            for iItemOffset, (pszItemName, _, bAllowBlank) in enumerate(
                STEP0004_VARIETY_ITEMS
            ):
                iRowIndex = iVarietyStart + iItemOffset
                try:
                    objValue = parse_decimal_value(
                        arrRows[iRowIndex][iColumnIndex],
                        iRowIndex + 1,
                        iColumnIndex,
                        pszItemName,
                    )
                except ConversionError as objException:
                    raise_conversion_error(
                        "INVALID_STEP0007_VARIETY_ITEM_VALUE",
                        "invalid variety step0007 item value. Path = "
                        + pszInputPath + ", Variety = " + pszVarietyName
                        + ", Item = " + pszItemName
                        + ". Detail = " + str(objException),
                    )
                if objValue is None and not bAllowBlank:
                    raise_conversion_error(
                        "INVALID_STEP0007_VARIETY_ITEM_VALUE",
                        "variety step0007 item value was blank. Path = "
                        + pszInputPath + ", Variety = " + pszVarietyName
                        + ", Item = " + pszItemName
                        + ", Column = " + str(iColumnIndex + 1),
                    )
    return arrRows, iExpectedYear


def create_step0008_variety_annual_files(
    arrStep0007VarietyPaths: List[str],
    pszInputExcelPath: str,
    pszExistingWarningPath: str,
    iExistingWarningCount: int,
) -> Tuple[List[str], str, int]:
    """品種別step0007年間TSVの育成率項目名を変更してstep0008を作成します。"""
    arrCreatedPaths: List[str] = []
    arrWarnings: List[Dict[str, str]] = []
    try:
        for pszStep0007Path in arrStep0007VarietyPaths:
            arrInputRows, iYear = read_step0007_variety_annual_for_step0008(
                pszStep0007Path
            )
            pszOutputPath = create_step0008_output_path(pszStep0007Path)
            arrOutputRows = [list(arrInputRows[0]), list(arrInputRows[1])]
            for iVarietyIndex, (pszVarietyNumber, pszVarietyName) in enumerate(
                SPLIT_VARIETIES
            ):
                iVarietyStart = 2 + iVarietyIndex * len(STEP0004_VARIETY_ITEMS)
                arrVarietyRows = [
                    list(arrRow)
                    for arrRow in arrInputRows[
                        iVarietyStart : iVarietyStart
                        + len(STEP0004_VARIETY_ITEMS)
                    ]
                ]
                arrVarietyRows[-1][1] = "育成率"
                arrOutputRows.extend(arrVarietyRows)
            write_step0008_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)

        pszWarningPath = write_growth_rate_warnings(
            pszInputExcelPath,
            pszExistingWarningPath,
            iExistingWarningCount,
            arrWarnings,
        )
        return (
            arrCreatedPaths,
            pszWarningPath,
            iExistingWarningCount + len(arrWarnings),
        )
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0008_VARIETY_CLEANUP_FAILED",
                "variety step0008 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0008_variety_annual_for_step0009(
    pszInputPath: str,
) -> Tuple[List[List[str]], int]:
    """品種別step0008年間TSVを検証して読み込みます。"""
    objFileNameMatch = re.search(
        r"_step0008_品種別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0008_VARIETY_FILENAME",
            "unexpected variety step0008 file name. Path = " + pszInputPath,
        )
    iExpectedYear = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0008_VARIETY_READ_FAILED",
            "failed to read variety step0008 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    arrExpectedItems = STEP0008_VARIETY_ITEM_NAMES
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for _, pszVarietyName in SPLIT_VARIETIES:
        arrExpectedStructure.extend(
            (pszVarietyName if iItemIndex == 0 else "", pszItemName)
            for iItemIndex, pszItemName in enumerate(arrExpectedItems)
        )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0008_VARIETY_ROW_COUNT",
            "variety step0008 must contain 32 rows. Path = " + pszInputPath,
        )
    iColumnCount = len(arrRows[0])
    if iColumnCount < 3 or any(
        len(arrRow) != iColumnCount for arrRow in arrRows
    ):
        raise_conversion_error(
            "INVALID_STEP0008_VARIETY_COLUMN_COUNT",
            "variety step0008 rows must have equal columns and at least one "
            "data month. Path = " + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0008_VARIETY_ROW_STRUCTURE",
                "unexpected variety step0008 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )

    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex],
            arrRows[1][iColumnIndex],
            iColumnIndex + 1,
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0008_VARIETY_YEAR",
                "variety step0008 year did not match its file name. Path = "
                + pszInputPath + ", Column = " + str(iColumnIndex + 1),
            )
        if iMonth in arrMonths or (arrMonths and iMonth <= arrMonths[-1]):
            raise_conversion_error(
                "INVALID_STEP0008_VARIETY_MONTH",
                "variety step0008 months must be unique and ascending. Path = "
                + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iVarietyIndex, (_, pszVarietyName) in enumerate(SPLIT_VARIETIES):
            iVarietyStart = 2 + iVarietyIndex * len(arrExpectedItems)
            for iItemOffset, pszItemName in enumerate(arrExpectedItems):
                iRowIndex = iVarietyStart + iItemOffset
                bAllowBlank = STEP0004_VARIETY_ITEMS[iItemOffset][2]
                try:
                    objValue = parse_decimal_value(
                        arrRows[iRowIndex][iColumnIndex],
                        iRowIndex + 1,
                        iColumnIndex,
                        pszItemName,
                    )
                except ConversionError as objException:
                    raise_conversion_error(
                        "INVALID_STEP0008_VARIETY_ITEM_VALUE",
                        "invalid variety step0008 item value. Path = "
                        + pszInputPath + ", Variety = " + pszVarietyName
                        + ", Item = " + pszItemName
                        + ". Detail = " + str(objException),
                    )
                if objValue is None and not bAllowBlank:
                    raise_conversion_error(
                        "INVALID_STEP0008_VARIETY_ITEM_VALUE",
                        "variety step0008 item value was blank. Path = "
                        + pszInputPath + ", Variety = " + pszVarietyName
                        + ", Item = " + pszItemName
                        + ", Column = " + str(iColumnIndex + 1),
                    )
    return arrRows, iExpectedYear


def read_step0007_variety_deduction_annual_for_step0008(
    pszInputPath: str,
) -> Tuple[List[List[str]], int]:
    """控除廃棄版の品種別step0007年間TSVを検証して読み込みます。"""
    objFileNameMatch = re.search(
        r"_step0007_品種別_控除廃棄_(\d{4})_年間(?:_\d+)?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0007_VARIETY_DEDUCTION_FILENAME",
            "unexpected variety deduction step0007 file name. Path = "
            + pszInputPath,
        )
    iExpectedYear = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0007_VARIETY_DEDUCTION_READ_FAILED",
            "failed to read variety deduction step0007 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for _, pszVarietyName in SPLIT_VARIETIES:
        arrExpectedStructure.extend(
            (pszVarietyName if iItemIndex == 0 else "", pszItemName)
            for iItemIndex, (pszItemName, _, _) in enumerate(
                STEP0004_VARIETY_DEDUCTION_ITEMS
            )
        )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0007_VARIETY_DEDUCTION_ROW_COUNT",
            "variety deduction step0007 must contain 32 rows. Path = "
            + pszInputPath,
        )
    iColumnCount = len(arrRows[0])
    if iColumnCount < 3 or any(
        len(arrRow) != iColumnCount for arrRow in arrRows
    ):
        raise_conversion_error(
            "INVALID_STEP0007_VARIETY_DEDUCTION_COLUMN_COUNT",
            "variety deduction step0007 rows must have equal columns and at "
            "least one data month. Path = " + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0007_VARIETY_DEDUCTION_ROW_STRUCTURE",
                "unexpected variety deduction step0007 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )

    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex],
            arrRows[1][iColumnIndex],
            iColumnIndex + 1,
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0007_VARIETY_DEDUCTION_YEAR",
                "variety deduction step0007 year did not match its file name. "
                "Path = " + pszInputPath + ", Column = "
                + str(iColumnIndex + 1),
            )
        if iMonth in arrMonths or (arrMonths and iMonth <= arrMonths[-1]):
            raise_conversion_error(
                "INVALID_STEP0007_VARIETY_DEDUCTION_MONTH",
                "variety deduction step0007 months must be unique and "
                "ascending. Path = " + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iVarietyIndex, (_, pszVarietyName) in enumerate(SPLIT_VARIETIES):
            iVarietyStart = (
                2 + iVarietyIndex * len(STEP0004_VARIETY_DEDUCTION_ITEMS)
            )
            for iItemOffset, (pszItemName, _, bAllowBlank) in enumerate(
                STEP0004_VARIETY_DEDUCTION_ITEMS
            ):
                iRowIndex = iVarietyStart + iItemOffset
                try:
                    objValue = parse_decimal_value(
                        arrRows[iRowIndex][iColumnIndex],
                        iRowIndex + 1,
                        iColumnIndex,
                        pszItemName,
                    )
                except ConversionError as objException:
                    raise_conversion_error(
                        "INVALID_STEP0007_VARIETY_DEDUCTION_ITEM_VALUE",
                        "invalid variety deduction step0007 item value. Path = "
                        + pszInputPath + ", Variety = " + pszVarietyName
                        + ", Item = " + pszItemName
                        + ". Detail = " + str(objException),
                    )
                if objValue is None and not bAllowBlank:
                    raise_conversion_error(
                        "INVALID_STEP0007_VARIETY_DEDUCTION_ITEM_VALUE",
                        "variety deduction step0007 item value was blank. "
                        "Path = " + pszInputPath + ", Variety = "
                        + pszVarietyName + ", Item = " + pszItemName
                        + ", Column = " + str(iColumnIndex + 1),
                    )
    return arrRows, iExpectedYear


def create_step0008_variety_deduction_annual_files(
    arrStep0007VarietyDeductionPaths: List[str],
    pszInputExcelPath: str,
    pszExistingWarningPath: str,
    iExistingWarningCount: int,
) -> Tuple[List[str], str, int]:
    """控除廃棄版step0007の育成率項目名を変更してstep0008を作成します。"""
    arrCreatedPaths: List[str] = []
    arrWarnings: List[Dict[str, str]] = []
    try:
        for pszStep0007Path in arrStep0007VarietyDeductionPaths:
            arrInputRows, iYear = (
                read_step0007_variety_deduction_annual_for_step0008(
                    pszStep0007Path
                )
            )
            pszOutputPath = create_step0008_output_path(pszStep0007Path)
            arrOutputRows = [list(arrInputRows[0]), list(arrInputRows[1])]
            for iVarietyIndex, (pszVarietyNumber, pszVarietyName) in enumerate(
                SPLIT_VARIETIES
            ):
                iVarietyStart = (
                    2 + iVarietyIndex * len(STEP0004_VARIETY_DEDUCTION_ITEMS)
                )
                arrVarietyRows = [
                    list(arrRow)
                    for arrRow in arrInputRows[
                        iVarietyStart : iVarietyStart
                        + len(STEP0004_VARIETY_DEDUCTION_ITEMS)
                    ]
                ]
                arrVarietyRows[-1][1] = "育成率"
                arrOutputRows.extend(arrVarietyRows)
            write_step0008_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)

        pszWarningPath = write_growth_rate_warnings(
            pszInputExcelPath,
            pszExistingWarningPath,
            iExistingWarningCount,
            arrWarnings,
        )
        return (
            arrCreatedPaths,
            pszWarningPath,
            iExistingWarningCount + len(arrWarnings),
        )
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0008_VARIETY_DEDUCTION_CLEANUP_FAILED",
                "variety deduction step0008 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0008_for_step0009(
    pszInputPath: str,
) -> List[List[str]]:
    objFileNameMatch = re.search(
        r"_step0008_グループ別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0008_FILENAME",
            "unexpected step0008 file name. Path = " + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0008_READ_FAILED",
            "failed to read step0008 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for _, pszGroupName in GROUPS:
        arrExpectedStructure.extend(
            [(pszGroupName, "入雛羽数"), ("", "出荷羽数"),
             ("", "正味重量"), ("", "平均体重"), ("", "育成率")]
        )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0008_ROW_COUNT",
            "step0008 must contain 17 rows. Path = " + pszInputPath,
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount < 3 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0008_COLUMN_COUNT",
            "step0008 rows must have equal columns and at least one data column. Path = "
            + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        if tuple(pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]) != objExpected:
            raise_conversion_error(
                "INVALID_STEP0008_ROW_STRUCTURE",
                "unexpected step0008 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0008_YEAR",
                "step0008 year did not match its file name. Path = " + pszInputPath,
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0008_MONTH",
                "step0008 months must be unique and ascending. Path = " + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            pszItemName: str = arrExpectedStructure[iRowIndex][1]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1, iColumnIndex, pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error("INVALID_STEP0008_ITEM_VALUE", str(objException))
            if objValue is None and pszItemName not in {"平均体重", "育成率"}:
                raise_conversion_error(
                    "INVALID_STEP0008_ITEM_VALUE",
                    "step0008 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1),
                )
    return arrRows


def create_step0008_excel_output_path(pszStep0008TsvPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0008TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0008TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_step0008_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0008_EXCEL_CREATE_FAILED",
                "step0008 Excel already exists. Path = " + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0008_EXCEL_CREATE_FAILED",
                "failed to reserve step0008 Excel. Detail = " + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0008_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "グループ別"
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2:
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue: Decimal = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue) if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszRowType in {"年", "入雛羽数", "出荷羽数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "月":
                        objCell.number_format = "00"
                    elif pszRowType == "正味重量":
                        iDecimalPlaces: int = max(0, -objValue.as_tuple().exponent)
                        objCell.number_format = (
                            "0." + "0" * iDecimalPlaces if iDecimalPlaces else "0"
                        )
                    elif pszRowType == "平均体重":
                        objCell.number_format = "0.000"
                    elif pszRowType == "育成率":
                        objCell.number_format = "0.0000"
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                "STEP0008_EXCEL_WRITE_FAILED",
                "failed to write step0008 Excel. Detail = " + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0008_EXCEL_FINALIZE_FAILED",
                "failed to finalize step0008 Excel. Detail = " + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0008_excel_files(
    arrStep0008TsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0008TsvPath in arrStep0008TsvPaths:
            arrRows = read_step0008_for_step0009(pszStep0008TsvPath)
            pszOutputPath: str = create_step0008_excel_output_path(
                pszStep0008TsvPath
            )
            write_step0008_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0008_EXCEL_CLEANUP_FAILED",
                "step0008 Excel creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0009_output_path(pszStep0008Path: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0008Path)
    pszFileName: str = os.path.basename(pszStep0008Path)
    if "_step0008_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0008_FOR_STEP0009",
            "step0008 marker was not found. Path = " + pszStep0008Path,
        )
    pszFileName = pszFileName.replace("_step0008_", "_step0009_", 1)
    pszStem, pszExtension = os.path.splitext(pszFileName)
    pszOutputPath: str = os.path.join(pszDirectory, pszFileName)
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0009_rows(arrInputRows: List[List[str]]) -> List[List[str]]:
    arrOutputRows: List[List[str]] = [
        arrRow[:2] + [
            format_sum_value(Decimal(pszValue.strip(" \u3000\t")))
            if pszValue.strip(" \u3000\t") else ""
            for pszValue in arrRow[2:]
        ]
        for arrRow in arrInputRows
    ]
    arrTotalRows: List[List[str]] = [
        ["合計", "入雛羽数"], ["", "出荷羽数"], ["", "正味重量"],
        ["", "平均体重"], ["", "育成率"],
    ]
    for iColumnIndex in range(2, len(arrInputRows[0])):
        arrTotals: List[Decimal] = [
            sum(
                (Decimal(arrInputRows[2 + iGroupIndex * 5 + iItemIndex][iColumnIndex].strip(" \u3000\t"))
                 for iGroupIndex in range(len(GROUPS))),
                Decimal("0"),
            )
            for iItemIndex in range(3)
        ]
        objAverageWeight: Decimal = (
            (arrTotals[2] / arrTotals[1]).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
            if arrTotals[1] != 0 else Decimal("0")
        )
        objGrowthRate: Decimal = (
            (arrTotals[1] / arrTotals[0]).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
            if arrTotals[0] != 0 else Decimal("0")
        )
        for arrRow, objValue in zip(arrTotalRows, arrTotals + [objAverageWeight, objGrowthRate]):
            arrRow.append(format_sum_value(objValue))
    return arrOutputRows + arrTotalRows


def create_step0009_variety_annual_rows(
    arrInputRows: List[List[str]],
) -> List[List[str]]:
    """品種別step0008に5区分合計行を加えてstep0009行を作成します。"""
    iItemsPerVariety = len(STEP0008_VARIETY_ITEM_NAMES)
    arrItemNames = STEP0008_VARIETY_ITEM_NAMES
    arrOutputRows: List[List[str]] = [
        arrRow[:2] + [
            format_sum_value(Decimal(pszValue.strip(" \u3000\t")))
            if pszValue.strip(" \u3000\t") else ""
            for pszValue in arrRow[2:]
        ]
        for arrRow in arrInputRows
    ]
    arrTotalRows: List[List[str]] = [
        ["合計" if iItemIndex == 0 else "", pszItemName]
        for iItemIndex, pszItemName in enumerate(arrItemNames)
    ]
    for iColumnIndex in range(2, len(arrInputRows[0])):
        arrTotals: List[Decimal] = [
            sum(
                (
                    Decimal(
                        arrInputRows[
                            2 + iVarietyIndex * iItemsPerVariety + iItemIndex
                        ][iColumnIndex].strip(" \u3000\t")
                    )
                    for iVarietyIndex in range(len(SPLIT_VARIETIES))
                ),
                Decimal("0"),
            )
            for iItemIndex in range(3)
        ]
        objAverageWeight: Decimal = (
            (arrTotals[2] / arrTotals[1]).quantize(
                Decimal("0.001"), rounding=ROUND_HALF_UP
            )
            if arrTotals[1] != 0 else Decimal("0")
        )
        objAgeWeightedTotal = Decimal("0")
        objAgeWeightTotal = Decimal("0")
        for iVarietyIndex in range(len(SPLIT_VARIETIES)):
            iVarietyStart = 2 + iVarietyIndex * iItemsPerVariety
            pszAge = arrInputRows[iVarietyStart + 4][iColumnIndex].strip(
                " \u3000\t"
            )
            if pszAge:
                objShippingCount = Decimal(
                    arrInputRows[iVarietyStart + 1][iColumnIndex].strip(
                        " \u3000\t"
                    )
                )
                objAgeWeightedTotal += Decimal(pszAge) * objShippingCount
                objAgeWeightTotal += objShippingCount
        objAverageAge: Decimal = (
            (objAgeWeightedTotal / objAgeWeightTotal).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            if objAgeWeightTotal != 0 else Decimal("0")
        )
        objGrowthRate: Decimal = (
            (arrTotals[1] / arrTotals[0]).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            )
            if arrTotals[0] != 0 else Decimal("0")
        )
        arrTotalValues = (
            arrTotals[:3]
            + [objAverageWeight, objAverageAge, objGrowthRate]
        )
        for arrRow, objValue in zip(arrTotalRows, arrTotalValues):
            arrRow.append(format_sum_value(objValue))
    return arrOutputRows + arrTotalRows


def write_step0009_tsv(pszOutputPath: str, arrRows: List[List[str]]) -> None:
    try:
        write_monthly_tsv(
            pszOutputPath, "\t".join(arrRows[0]),
            ["\t".join(arrRow) for arrRow in arrRows[1:]],
        )
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0009_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0009_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0009_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0009_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0009_WRITE_FAILED"),
            str(objException),
        )


def create_step0009_files(arrStep0008Paths: List[str]) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0008Path in arrStep0008Paths:
            arrInputRows = read_step0008_for_step0009(pszStep0008Path)
            pszOutputPath: str = create_step0009_output_path(pszStep0008Path)
            write_step0009_tsv(pszOutputPath, create_step0009_rows(arrInputRows))
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0009_CLEANUP_FAILED",
                "step0009 creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0009_variety_annual_files(
    arrStep0008VarietyAnnualPaths: List[str],
) -> List[str]:
    """品種別step0008年間TSVを読み込み、合計行を加えたstep0009を作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0008Path in arrStep0008VarietyAnnualPaths:
            arrInputRows, _ = read_step0008_variety_annual_for_step0009(
                pszStep0008Path
            )
            pszOutputPath = create_step0009_output_path(pszStep0008Path)
            arrOutputRows = create_step0009_variety_annual_rows(arrInputRows)
            write_step0009_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0009_VARIETY_CLEANUP_FAILED",
                "variety step0009 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0009_for_excel(
    pszInputPath: str,
) -> List[List[str]]:
    objFileNameMatch = re.search(
        r"_step0009_グループ別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0009_FILENAME",
            "unexpected step0009 file name. Path = " + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0009_READ_FAILED",
            "failed to read step0009 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for pszGroupName in [pszName for _, pszName in GROUPS] + ["合計"]:
        arrExpectedStructure.extend(
            [(pszGroupName, "入雛羽数"), ("", "出荷羽数"),
             ("", "正味重量"), ("", "平均体重"), ("", "育成率")]
        )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0009_ROW_COUNT",
            "step0009 must contain 22 rows. Path = " + pszInputPath,
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount < 3 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0009_COLUMN_COUNT",
            "step0009 rows must have equal columns and at least one data column. Path = "
            + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        if tuple(pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]) != objExpected:
            raise_conversion_error(
                "INVALID_STEP0009_ROW_STRUCTURE",
                "unexpected step0009 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0009_YEAR",
                "step0009 year did not match its file name. Path = " + pszInputPath,
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0009_MONTH",
                "step0009 months must be unique and ascending. Path = " + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            pszItemName: str = arrExpectedStructure[iRowIndex][1]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1, iColumnIndex, pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error("INVALID_STEP0009_ITEM_VALUE", str(objException))
            if objValue is None and not (
                iRowIndex < 17 and pszItemName in {"平均体重", "育成率"}
            ):
                raise_conversion_error(
                    "INVALID_STEP0009_ITEM_VALUE",
                    "step0009 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1),
                )
    return arrRows


def create_step0009_excel_output_path(pszStep0009TsvPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0009TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0009TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_step0009_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0009_EXCEL_CREATE_FAILED",
                "step0009 Excel already exists. Path = " + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0009_EXCEL_CREATE_FAILED",
                "failed to reserve step0009 Excel. Detail = " + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0009_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "グループ別"
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2:
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue: Decimal = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue) if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszRowType in {"年", "入雛羽数", "出荷羽数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "月":
                        objCell.number_format = "0"
                    elif pszRowType == "正味重量":
                        iDecimalPlaces: int = max(0, -objValue.as_tuple().exponent)
                        objCell.number_format = (
                            "0." + "#" * iDecimalPlaces if iDecimalPlaces else "0"
                        )
                    elif pszRowType == "平均体重":
                        objCell.number_format = "0" if objValue == objValue.to_integral_value() else "0.###"
                    elif pszRowType == "育成率":
                        objCell.number_format = "0" if objValue == objValue.to_integral_value() else "0.####"
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                "STEP0009_EXCEL_WRITE_FAILED",
                "failed to write step0009 Excel. Detail = " + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0009_EXCEL_FINALIZE_FAILED",
                "failed to finalize step0009 Excel. Detail = " + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0009_excel_files(
    arrStep0009TsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0009TsvPath in arrStep0009TsvPaths:
            arrRows = read_step0009_for_excel(pszStep0009TsvPath)
            pszOutputPath: str = create_step0009_excel_output_path(
                pszStep0009TsvPath
            )
            write_step0009_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0009_EXCEL_CLEANUP_FAILED",
                "step0009 Excel creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0008_variety_deduction_for_step0009(
    pszInputPath: str,
) -> List[List[str]]:
    """控除廃棄版品種別step0008年間TSVを検証して読み込みます。"""
    objFileNameMatch = re.search(
        r"_step0008_品種別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0008_VARIETY_DEDUCTION_FILENAME",
            "unexpected variety deduction step0008 file name. Path = "
            + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0008_VARIETY_DEDUCTION_READ_FAILED",
            "failed to read variety deduction step0008 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    arrItemNames: List[str] = STEP0008_VARIETY_DEDUCTION_ITEM_NAMES
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for _, pszVarietyName in SPLIT_VARIETIES:
        for iItemIndex, pszItemName in enumerate(arrItemNames):
            arrExpectedStructure.append(
                (pszVarietyName if iItemIndex == 0 else "", pszItemName)
            )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0008_VARIETY_DEDUCTION_ROW_COUNT",
            "variety deduction step0008 must contain 32 rows. Path = "
            + pszInputPath,
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount < 3 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0008_VARIETY_DEDUCTION_COLUMN_COUNT",
            "variety deduction step0008 rows must have equal columns and at "
            "least one data month. Path = " + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0008_VARIETY_DEDUCTION_ROW_STRUCTURE",
                "unexpected variety deduction step0008 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )

    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0008_VARIETY_DEDUCTION_YEAR",
                "variety deduction step0008 year did not match its file name. "
                "Path = " + pszInputPath,
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0008_VARIETY_DEDUCTION_MONTH",
                "variety deduction step0008 months must be unique and "
                "ascending. Path = " + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            iItemOffset: int = (iRowIndex - 2) % len(arrItemNames)
            pszItemName: str = arrItemNames[iItemOffset]
            bAllowBlank: bool = STEP0004_VARIETY_DEDUCTION_ITEMS[iItemOffset][2]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1, iColumnIndex, pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0008_VARIETY_DEDUCTION_ITEM_VALUE",
                    "invalid variety deduction step0008 item value. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1)
                    + ". Detail = " + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0008_VARIETY_DEDUCTION_ITEM_VALUE",
                    "variety deduction step0008 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1)
                    + ", Column = " + str(iColumnIndex + 1),
                )
    return arrRows


def create_step0009_variety_deduction_annual_rows(
    arrInputRows: List[List[str]],
) -> List[List[str]]:
    """控除廃棄版step0008に5区分合計行を加えてstep0009を作成します。"""
    iItemsPerVariety: int = len(STEP0008_VARIETY_DEDUCTION_ITEM_NAMES)
    arrItemNames: List[str] = STEP0008_VARIETY_DEDUCTION_ITEM_NAMES
    arrOutputRows: List[List[str]] = [
        arrRow[:2] + [
            format_sum_value(Decimal(pszValue.strip(" \u3000\t")))
            if pszValue.strip(" \u3000\t") else ""
            for pszValue in arrRow[2:]
        ]
        for arrRow in arrInputRows
    ]
    arrTotalRows: List[List[str]] = [
        ["合計" if iItemIndex == 0 else "", pszItemName]
        for iItemIndex, pszItemName in enumerate(arrItemNames)
    ]
    for iColumnIndex in range(2, len(arrInputRows[0])):
        arrTotals: List[Decimal] = [
            sum(
                (
                    Decimal(
                        arrInputRows[
                            2 + iVarietyIndex * iItemsPerVariety + iItemIndex
                        ][iColumnIndex].strip(" \u3000\t")
                    )
                    for iVarietyIndex in range(len(SPLIT_VARIETIES))
                ),
                Decimal("0"),
            )
            for iItemIndex in range(3)
        ]
        objAverageWeight: Decimal = (
            (arrTotals[2] / arrTotals[1]).quantize(
                Decimal("0.001"), rounding=ROUND_HALF_UP
            )
            if arrTotals[1] != 0 else Decimal("0")
        )
        objAgeWeightedTotal: Decimal = Decimal("0")
        objAgeWeightTotal: Decimal = Decimal("0")
        for iVarietyIndex in range(len(SPLIT_VARIETIES)):
            iVarietyStart: int = 2 + iVarietyIndex * iItemsPerVariety
            pszAge: str = arrInputRows[iVarietyStart + 4][iColumnIndex].strip(
                " \u3000\t"
            )
            if pszAge:
                objProcessedCount: Decimal = Decimal(
                    arrInputRows[iVarietyStart + 1][iColumnIndex].strip(
                        " \u3000\t"
                    )
                )
                objAgeWeightedTotal += Decimal(pszAge) * objProcessedCount
                objAgeWeightTotal += objProcessedCount
        objAverageAge: Decimal = (
            (objAgeWeightedTotal / objAgeWeightTotal).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            if objAgeWeightTotal != 0 else Decimal("0")
        )
        objGrowthRate: Decimal = (
            (arrTotals[1] / arrTotals[0]).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            )
            if arrTotals[0] != 0 else Decimal("0")
        )
        arrTotalValues: List[Decimal] = (
            arrTotals[:3]
            + [objAverageWeight, objAverageAge, objGrowthRate]
        )
        for arrRow, objValue in zip(arrTotalRows, arrTotalValues):
            arrRow.append(format_sum_value(objValue))
    return arrOutputRows + arrTotalRows


def create_step0009_variety_deduction_annual_files(
    arrStep0008VarietyDeductionPaths: List[str],
) -> List[str]:
    """控除廃棄版品種別step0008年間TSVから合計付きstep0009を作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0008Path in arrStep0008VarietyDeductionPaths:
            arrInputRows = read_step0008_variety_deduction_for_step0009(
                pszStep0008Path
            )
            pszOutputPath = create_step0009_output_path(pszStep0008Path)
            arrOutputRows = create_step0009_variety_deduction_annual_rows(
                arrInputRows
            )
            write_step0009_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0009_VARIETY_DEDUCTION_CLEANUP_FAILED",
                "variety deduction step0009 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0009_variety_for_excel(
    pszInputPath: str,
) -> List[List[str]]:
    objFileNameMatch = re.search(
        r"_step0009_品種別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_FILENAME",
            "unexpected variety step0009 file name. Path = " + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0009_VARIETY_READ_FAILED",
            "failed to read variety step0009 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    arrItemNames: List[str] = STEP0008_VARIETY_ITEM_NAMES
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for pszVarietyName in [pszName for _, pszName in SPLIT_VARIETIES] + ["合計"]:
        for iItemIndex, pszItemName in enumerate(arrItemNames):
            arrExpectedStructure.append(
                (pszVarietyName if iItemIndex == 0 else "", pszItemName)
            )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_ROW_COUNT",
            "variety step0009 must contain 38 rows. Path = " + pszInputPath,
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount < 3 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_COLUMN_COUNT",
            "variety step0009 rows must have equal columns and at least one data column. Path = "
            + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        if tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        ) != objExpected:
            raise_conversion_error(
                "INVALID_STEP0009_VARIETY_ROW_STRUCTURE",
                "unexpected variety step0009 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0009_VARIETY_YEAR",
                "variety step0009 year did not match its file name. Path = "
                + pszInputPath,
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0009_VARIETY_MONTH",
                "variety step0009 months must be unique and ascending. Path = "
                + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            iItemOffset: int = (iRowIndex - 2) % len(arrItemNames)
            pszItemName: str = arrItemNames[iItemOffset]
            bAllowBlank: bool = STEP0004_VARIETY_ITEMS[iItemOffset][2]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1, iColumnIndex, pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0009_VARIETY_ITEM_VALUE", str(objException)
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0009_VARIETY_ITEM_VALUE",
                    "variety step0009 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1),
                )
    return arrRows


def create_step0009_variety_excel_output_path(
    pszStep0009TsvPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszStep0009TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0009TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_step0009_variety_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0009_VARIETY_EXCEL_CREATE_FAILED",
                "variety step0009 Excel already exists. Path = " + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0009_VARIETY_EXCEL_CREATE_FAILED",
                "failed to reserve variety step0009 Excel. Detail = "
                + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0009_variety_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "品種別"
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2:
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue: Decimal = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue) if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszRowType in {"年", "入雛羽数", "出荷羽数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "月":
                        objCell.number_format = "0"
                    elif pszRowType == "正味重量":
                        iDecimalPlaces: int = max(0, -objValue.as_tuple().exponent)
                        objCell.number_format = (
                            "0." + "#" * iDecimalPlaces if iDecimalPlaces else "0"
                        )
                    elif pszRowType in {"平均体重", "飼育日数平均"}:
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0.###"
                        )
                    elif pszRowType in {"廃棄控除前育成率", "育成率"}:
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0.####"
                        )
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                "STEP0009_VARIETY_EXCEL_WRITE_FAILED",
                "failed to write variety step0009 Excel. Detail = "
                + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0009_VARIETY_EXCEL_FINALIZE_FAILED",
                "failed to finalize variety step0009 Excel. Detail = "
                + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0009_variety_excel_files(
    arrStep0009VarietyTsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0009TsvPath in arrStep0009VarietyTsvPaths:
            arrRows = read_step0009_variety_for_excel(pszStep0009TsvPath)
            pszOutputPath = create_step0009_variety_excel_output_path(
                pszStep0009TsvPath
            )
            write_step0009_variety_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0009_VARIETY_EXCEL_CLEANUP_FAILED",
                "variety step0009 Excel creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0009_variety_deduction_for_excel(
    pszInputPath: str,
) -> List[List[str]]:
    """控除廃棄版品種別step0009年間TSVを検証して読み込みます。"""
    objFileNameMatch = re.search(
        r"_step0009_品種別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_DEDUCTION_FILENAME",
            "unexpected variety deduction step0009 file name. Path = "
            + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0009_VARIETY_DEDUCTION_READ_FAILED",
            "failed to read variety deduction step0009 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    arrItemNames: List[str] = STEP0008_VARIETY_DEDUCTION_ITEM_NAMES
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for pszVarietyName in [pszName for _, pszName in SPLIT_VARIETIES] + ["合計"]:
        for iItemIndex, pszItemName in enumerate(arrItemNames):
            arrExpectedStructure.append(
                (pszVarietyName if iItemIndex == 0 else "", pszItemName)
            )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_DEDUCTION_ROW_COUNT",
            "variety deduction step0009 must contain 38 rows. Path = "
            + pszInputPath,
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount < 3 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_DEDUCTION_COLUMN_COUNT",
            "variety deduction step0009 rows must have equal columns and at "
            "least one data month. Path = " + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0009_VARIETY_DEDUCTION_ROW_STRUCTURE",
                "unexpected variety deduction step0009 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
    arrMonths: List[int] = []
    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0009_VARIETY_DEDUCTION_YEAR",
                "variety deduction step0009 year did not match its file name. "
                "Path = " + pszInputPath,
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0009_VARIETY_DEDUCTION_MONTH",
                "variety deduction step0009 months must be unique and "
                "ascending. Path = " + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            iItemOffset: int = (iRowIndex - 2) % len(arrItemNames)
            pszItemName: str = arrItemNames[iItemOffset]
            bAllowBlank: bool = STEP0004_VARIETY_DEDUCTION_ITEMS[iItemOffset][2]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1, iColumnIndex, pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0009_VARIETY_DEDUCTION_ITEM_VALUE",
                    "invalid variety deduction step0009 item value. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1)
                    + ". Detail = " + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0009_VARIETY_DEDUCTION_ITEM_VALUE",
                    "variety deduction step0009 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1),
                )
    return arrRows


def create_step0009_variety_deduction_excel_output_path(
    pszStep0009TsvPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszStep0009TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0009TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_step0009_variety_deduction_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0009_VARIETY_DEDUCTION_EXCEL_CREATE_FAILED",
                "variety deduction step0009 Excel already exists. Path = "
                + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0009_VARIETY_DEDUCTION_EXCEL_CREATE_FAILED",
                "failed to reserve variety deduction step0009 Excel. Detail = "
                + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0009_variety_deduction_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "品種別"
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2:
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue: Decimal = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue) if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszRowType in {"年", "月", "入雛羽数", "処理羽数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "処理重量":
                        iDecimalPlaces: int = max(0, -objValue.as_tuple().exponent)
                        objCell.number_format = (
                            "0." + "#" * iDecimalPlaces if iDecimalPlaces else "0"
                        )
                    elif pszRowType in {"控除後平均体重", "飼育日数平均"}:
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0.###"
                        )
                    elif pszRowType in {"廃棄控除後育成率", "育成率"}:
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0.####"
                        )
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                "STEP0009_VARIETY_DEDUCTION_EXCEL_WRITE_FAILED",
                "failed to write variety deduction step0009 Excel. Detail = "
                + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0009_VARIETY_DEDUCTION_EXCEL_FINALIZE_FAILED",
                "failed to finalize variety deduction step0009 Excel. Detail = "
                + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0009_variety_deduction_excel_files(
    arrStep0009VarietyDeductionTsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0009TsvPath in arrStep0009VarietyDeductionTsvPaths:
            arrRows = read_step0009_variety_deduction_for_excel(
                pszStep0009TsvPath
            )
            pszOutputPath = create_step0009_variety_deduction_excel_output_path(
                pszStep0009TsvPath
            )
            write_step0009_variety_deduction_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0009_VARIETY_DEDUCTION_EXCEL_CLEANUP_FAILED",
                "variety deduction step0009 Excel creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0009_for_step0010(pszInputPath: str) -> List[List[str]]:
    arrRows = read_step0009_for_excel(pszInputPath)
    if len(arrRows[0]) != 14 or [
        int(Decimal(pszMonth.strip(" \u3000\t"))) for pszMonth in arrRows[1][2:]
    ] != list(range(1, 13)):
        raise_conversion_error(
            "INVALID_STEP0009_FOR_STEP0010",
            "step0010 requires all 12 months in ascending order. Path = " + pszInputPath,
        )
    for iRowIndex, arrRow in enumerate(arrRows[2:], start=3):
        if arrRow[1].strip(" \u3000\t") in {"入雛羽数", "出荷羽数"}:
            for pszValue in arrRow[2:]:
                objValue: Decimal = Decimal(pszValue.strip(" \u3000\t"))
                if objValue != objValue.to_integral_value():
                    raise_conversion_error(
                        "INVALID_STEP0010_COUNT_VALUE",
                        "placement and shipping counts must be integers. Path = "
                        + pszInputPath + ", Line = " + str(iRowIndex),
                    )
    return arrRows


def create_step0010_output_path(pszStep0009Path: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0009Path)
    pszFileName: str = os.path.basename(pszStep0009Path)
    if "_step0009_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0009_FOR_STEP0010",
            "step0009 marker was not found. Path = " + pszStep0009Path,
        )
    pszFileName = pszFileName.replace("_step0009_", "_step0010_", 1)
    pszStem, pszExtension = os.path.splitext(pszFileName)
    pszOutputPath: str = os.path.join(pszDirectory, pszFileName)
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0010_rows(arrInputRows: List[List[str]]) -> List[List[str]]:
    arrOutputRows: List[List[str]] = [list(arrRow) for arrRow in arrInputRows]
    arrOutputRows[0].append("")
    arrOutputRows[1].append("合計")
    for iGroupIndex in range(len(GROUPS) + 1):
        iGroupStart: int = 2 + iGroupIndex * 5
        arrTotals: List[Decimal] = [
            sum(
                (Decimal(pszValue.strip(" \u3000\t"))
                 for pszValue in arrInputRows[iGroupStart + iItemIndex][2:]),
                Decimal("0"),
            )
            for iItemIndex in range(3)
        ]
        objAverageWeight: Decimal = (
            (arrTotals[2] / arrTotals[1]).quantize(Decimal("0.001"), rounding=ROUND_HALF_UP)
            if arrTotals[1] != 0 else Decimal("0")
        )
        objGrowthRate: Decimal = (
            (arrTotals[1] / arrTotals[0]).quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)
            if arrTotals[0] != 0 else Decimal("0")
        )
        for iItemIndex, objValue in enumerate(arrTotals + [objAverageWeight, objGrowthRate]):
            arrOutputRows[iGroupStart + iItemIndex].append(format_sum_value(objValue))
    return arrOutputRows


def write_step0010_tsv(pszOutputPath: str, arrRows: List[List[str]]) -> None:
    try:
        write_monthly_tsv(
            pszOutputPath, "\t".join(arrRows[0]),
            ["\t".join(arrRow) for arrRow in arrRows[1:]],
        )
    except ConversionError as objException:
        dicErrorCodes: Dict[str, str] = {
            "MONTHLY_OUTPUT_ALREADY_EXISTS": "STEP0010_CREATE_FAILED",
            "MONTHLY_TSV_CREATE_FAILED": "STEP0010_CREATE_FAILED",
            "MONTHLY_TSV_WRITE_FAILED": "STEP0010_WRITE_FAILED",
            "MONTHLY_TSV_FINALIZE_FAILED": "STEP0010_FINALIZE_FAILED",
        }
        raise_conversion_error(
            dicErrorCodes.get(objException.pszErrorCode, "STEP0010_WRITE_FAILED"),
            str(objException),
        )


def create_step0010_files(arrStep0009Paths: List[str]) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0009Path in arrStep0009Paths:
            arrInputRows = read_step0009_for_step0010(pszStep0009Path)
            pszOutputPath: str = create_step0010_output_path(pszStep0009Path)
            write_step0010_tsv(pszOutputPath, create_step0010_rows(arrInputRows))
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0010_CLEANUP_FAILED",
                "step0010 creation failed. Original error = "
                + str(objException) + "; cleanup errors = " + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0009_variety_for_step0010(
    pszInputPath: str,
) -> List[List[str]]:
    """品種別step0009年間TSVを検証し、step0010用に読み込みます。"""
    arrRows = read_step0009_variety_for_excel(pszInputPath)
    if len(arrRows[0]) != 14 or [
        int(Decimal(pszMonth.strip(" \u3000\t"))) for pszMonth in arrRows[1][2:]
    ] != list(range(1, 13)):
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_FOR_STEP0010",
            "variety step0010 requires all 12 months in ascending order. Path = "
            + pszInputPath,
        )
    for iRowIndex, arrRow in enumerate(arrRows[2:], start=3):
        if arrRow[1].strip(" \u3000\t") in {"入雛羽数", "出荷羽数"}:
            for pszValue in arrRow[2:]:
                objValue: Decimal = Decimal(pszValue.strip(" \u3000\t"))
                if objValue != objValue.to_integral_value():
                    raise_conversion_error(
                        "INVALID_STEP0010_VARIETY_COUNT_VALUE",
                        "placement and shipping counts must be integers. Path = "
                        + pszInputPath + ", Line = " + str(iRowIndex),
                    )
    return arrRows


def create_step0010_variety_rows(
    arrInputRows: List[List[str]],
) -> List[List[str]]:
    """品種別step0009年間TSVの右端に品種別年間合計列を追加します。"""
    iItemsPerVariety: int = len(STEP0008_VARIETY_ITEM_NAMES)
    arrOutputRows: List[List[str]] = [list(arrRow) for arrRow in arrInputRows]
    arrOutputRows[0].append("")
    arrOutputRows[1].append("合計")
    for iVarietyIndex in range(len(SPLIT_VARIETIES) + 1):
        iVarietyStart: int = 2 + iVarietyIndex * iItemsPerVariety
        arrTotals: List[Decimal] = [
            sum(
                (
                    Decimal(pszValue.strip(" \u3000\t"))
                    for pszValue in arrInputRows[
                        iVarietyStart + iItemIndex
                    ][2:]
                ),
                Decimal("0"),
            )
            for iItemIndex in range(3)
        ]
        objAverageWeight: Decimal = (
            (arrTotals[2] / arrTotals[1]).quantize(
                Decimal("0.001"), rounding=ROUND_HALF_UP
            )
            if arrTotals[1] != 0 else Decimal("0")
        )
        objAgeWeightedTotal: Decimal = Decimal("0")
        objAgeWeightTotal: Decimal = Decimal("0")
        for iMonthColumnIndex in range(2, len(arrInputRows[0])):
            pszAge: str = arrInputRows[iVarietyStart + 4][
                iMonthColumnIndex
            ].strip(" \u3000\t")
            if pszAge:
                objShippingCount: Decimal = Decimal(
                    arrInputRows[iVarietyStart + 1][
                        iMonthColumnIndex
                    ].strip(" \u3000\t")
                )
                objAgeWeightedTotal += Decimal(pszAge) * objShippingCount
                objAgeWeightTotal += objShippingCount
        objAverageAge: Decimal = (
            (objAgeWeightedTotal / objAgeWeightTotal).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            if objAgeWeightTotal != 0 else Decimal("0")
        )
        objGrowthRate: Decimal = (
            (arrTotals[1] / arrTotals[0]).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            )
            if arrTotals[0] != 0 else Decimal("0")
        )
        arrAnnualValues: List[Decimal] = (
            arrTotals[:3]
            + [objAverageWeight, objAverageAge, objGrowthRate]
        )
        for iItemIndex, objValue in enumerate(arrAnnualValues):
            arrOutputRows[iVarietyStart + iItemIndex].append(
                format_sum_value(objValue)
            )
    return arrOutputRows


def create_step0010_variety_output_path(pszStep0009Path: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0009Path)
    pszFileName: str = os.path.basename(pszStep0009Path)
    if "_step0009_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_FOR_STEP0010",
            "variety step0009 marker was not found. Path = " + pszStep0009Path,
        )
    pszFileName = pszFileName.replace("_step0009_", "_step0010_", 1)
    pszStem, pszExtension = os.path.splitext(pszFileName)
    pszOutputPath: str = os.path.join(pszDirectory, pszFileName)
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0010_variety_annual_files(
    arrStep0009VarietyAnnualPaths: List[str],
) -> List[str]:
    """品種別step0009年間TSVから年間合計列付きstep0010を作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0009Path in arrStep0009VarietyAnnualPaths:
            arrInputRows = read_step0009_variety_for_step0010(pszStep0009Path)
            pszOutputPath = create_step0010_variety_output_path(
                pszStep0009Path
            )
            arrOutputRows = create_step0010_variety_rows(arrInputRows)
            write_step0010_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0010_VARIETY_CLEANUP_FAILED",
                "variety step0010 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0010_variety_for_excel(
    pszInputPath: str,
) -> List[List[str]]:
    """品種別step0010年間TSVを検証してExcel用に読み込みます。"""
    objFileNameMatch = re.search(
        r"_step0010_品種別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_FILENAME",
            "unexpected variety step0010 file name. Path = " + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0010_VARIETY_READ_FAILED",
            "failed to read variety step0010 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    arrItemNames: List[str] = STEP0008_VARIETY_ITEM_NAMES
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for pszVarietyName in [pszName for _, pszName in SPLIT_VARIETIES] + ["合計"]:
        for iItemIndex, pszItemName in enumerate(arrItemNames):
            arrExpectedStructure.append(
                (pszVarietyName if iItemIndex == 0 else "", pszItemName)
            )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_ROW_COUNT",
            "variety step0010 must contain 38 rows. Path = " + pszInputPath,
        )
    if any(len(arrRow) != 15 for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_COLUMN_COUNT",
            "variety step0010 rows must contain exactly 15 columns. Path = "
            + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        if tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        ) != objExpected:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_ROW_STRUCTURE",
                "unexpected variety step0010 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
    if arrRows[0][14].strip(" \u3000\t") != "" or arrRows[1][14].strip(
        " \u3000\t"
    ) != "合計":
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_TOTAL_HEADER",
            "variety step0010 annual column header is invalid. Path = "
            + pszInputPath,
        )
    arrMonths: List[int] = []
    for iColumnIndex in range(2, 14):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_YEAR",
                "variety step0010 year did not match its file name. Path = "
                + pszInputPath,
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_MONTH",
                "variety step0010 months must be unique and ascending. Path = "
                + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            iItemOffset: int = (iRowIndex - 2) % len(arrItemNames)
            pszItemName: str = arrItemNames[iItemOffset]
            bAllowBlank: bool = STEP0004_VARIETY_ITEMS[iItemOffset][2]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1, iColumnIndex + 1, pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0010_VARIETY_ITEM_VALUE", str(objException)
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0010_VARIETY_ITEM_VALUE",
                    "variety step0010 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1),
                )
    if arrMonths != list(range(1, 13)):
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_MONTH",
            "variety step0010 must contain months 1 through 12. Path = "
            + pszInputPath,
        )
    for iRowIndex in range(2, len(arrRows)):
        iItemOffset = (iRowIndex - 2) % len(arrItemNames)
        pszItemName = arrItemNames[iItemOffset]
        try:
            objAnnualValue = parse_decimal_value(
                arrRows[iRowIndex][14], iRowIndex + 1, 15, pszItemName
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_ITEM_VALUE", str(objException)
            )
        if objAnnualValue is None:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_ITEM_VALUE",
                "variety step0010 annual value was blank. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
        if pszItemName in {"入雛羽数", "出荷羽数"}:
            for pszValue in arrRows[iRowIndex][2:]:
                objCount: Decimal = Decimal(pszValue.strip(" \u3000\t"))
                if objCount != objCount.to_integral_value():
                    raise_conversion_error(
                        "INVALID_STEP0010_VARIETY_COUNT_VALUE",
                        "variety step0010 counts must be integers. Path = "
                        + pszInputPath + ", Line = " + str(iRowIndex + 1),
                    )
    return arrRows


def create_step0010_variety_excel_output_path(
    pszStep0010TsvPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszStep0010TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0010TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_step0010_variety_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0010_VARIETY_EXCEL_CREATE_FAILED",
                "variety step0010 Excel already exists. Path = " + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0010_VARIETY_EXCEL_CREATE_FAILED",
                "failed to reserve variety step0010 Excel. Detail = "
                + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0010_variety_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "品種別"
            pszCurrentVariety: str = ""
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                if arrFields[0].strip(" \u3000\t"):
                    pszCurrentVariety = arrFields[0].strip(" \u3000\t")
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2 or (
                        iRowIndex == 2 and iColumnIndex == 15
                    ):
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    if (
                        pszCurrentVariety == BROILER_BREAKDOWNS[1][1]
                        and pszRowType == "飼育日数平均"
                    ):
                        continue
                    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue: Decimal = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue) if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszRowType in {"年", "入雛羽数", "出荷羽数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "月":
                        objCell.number_format = "0"
                    elif pszRowType == "正味重量":
                        iDecimalPlaces: int = max(
                            0, -objValue.as_tuple().exponent
                        )
                        objCell.number_format = (
                            "0." + "#" * iDecimalPlaces
                            if iDecimalPlaces else "0"
                        )
                    elif pszRowType in {"平均体重", "飼育日数平均"}:
                        iDecimalPlaces = (
                            max(3, -objValue.as_tuple().exponent)
                            if iColumnIndex < 15 else 3
                        )
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0." + "#" * iDecimalPlaces
                        )
                    elif pszRowType in {"廃棄控除前育成率", "育成率"}:
                        iDecimalPlaces = (
                            max(4, -objValue.as_tuple().exponent)
                            if iColumnIndex < 15 else 4
                        )
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0." + "#" * iDecimalPlaces
                        )
            for iColumnIndex in range(1, objWorksheet.max_column + 1):
                iMaximumDisplayWidth = 0
                for objWidthColumn in objWorksheet.iter_cols(
                    min_col=iColumnIndex,
                    max_col=iColumnIndex,
                    min_row=1,
                    max_row=objWorksheet.max_row,
                ):
                    for objWidthCell in objWidthColumn:
                        objWidthValue = objWidthCell.value
                        if objWidthValue is None:
                            continue
                        if isinstance(objWidthValue, (int, float, Decimal)):
                            objWidthDecimal = Decimal(str(objWidthValue))
                            pszFractionFormat = (
                                objWidthCell.number_format.split(";")[0]
                                .partition(".")[2]
                            )
                            iFractionPlaces = sum(
                                1 for pszCharacter in pszFractionFormat
                                if pszCharacter in "0#"
                            )
                            pszDisplayValue = format(
                                objWidthDecimal,
                                "." + str(iFractionPlaces) + "f",
                            )
                            if iFractionPlaces:
                                iMinimumFractionPlaces = pszFractionFormat.count("0")
                                while (
                                    len(pszDisplayValue.rsplit(".", 1)[-1])
                                    > iMinimumFractionPlaces
                                    and pszDisplayValue.endswith("0")
                                ):
                                    pszDisplayValue = pszDisplayValue[:-1]
                                if pszDisplayValue.endswith("."):
                                    pszDisplayValue = pszDisplayValue[:-1]
                        else:
                            pszDisplayValue = str(objWidthValue)
                        iDisplayWidth = sum(
                            2 if unicodedata.east_asian_width(pszCharacter)
                            in {"F", "W"}
                            else 1
                            for pszCharacter in pszDisplayValue
                        )
                        iMaximumDisplayWidth = max(
                            iMaximumDisplayWidth, iDisplayWidth
                        )
                objWorksheet.column_dimensions[
                    get_column_letter(iColumnIndex)
                ].width = min(max(iMaximumDisplayWidth + 2, 8), 255)
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                "STEP0010_VARIETY_EXCEL_WRITE_FAILED",
                "failed to write variety step0010 Excel. Detail = "
                + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0010_VARIETY_EXCEL_FINALIZE_FAILED",
                "failed to finalize variety step0010 Excel. Detail = "
                + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0010_variety_excel_files(
    arrStep0010VarietyTsvPaths: List[str],
) -> List[str]:
    """品種別step0010年間TSVを読み込み年間Excelを作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0010TsvPath in arrStep0010VarietyTsvPaths:
            arrRows = read_step0010_variety_for_excel(pszStep0010TsvPath)
            pszOutputPath = create_step0010_variety_excel_output_path(
                pszStep0010TsvPath
            )
            write_step0010_variety_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0010_VARIETY_EXCEL_CLEANUP_FAILED",
                "variety step0010 Excel creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0009_variety_deduction_for_step0010(
    pszInputPath: str,
) -> List[List[str]]:
    """控除廃棄版品種別step0009を検証し、step0010用に読み込みます。"""
    arrRows = read_step0009_variety_deduction_for_excel(pszInputPath)
    if len(arrRows[0]) != 14 or [
        int(Decimal(pszMonth.strip(" \u3000\t"))) for pszMonth in arrRows[1][2:]
    ] != list(range(1, 13)):
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_DEDUCTION_FOR_STEP0010",
            "variety deduction step0010 requires all 12 months in ascending "
            "order. Path = " + pszInputPath,
        )
    for iRowIndex, arrRow in enumerate(arrRows[2:], start=3):
        if arrRow[1].strip(" \u3000\t") in {"入雛羽数", "処理羽数"}:
            for pszValue in arrRow[2:]:
                objValue: Decimal = Decimal(pszValue.strip(" \u3000\t"))
                if objValue != objValue.to_integral_value():
                    raise_conversion_error(
                        "INVALID_STEP0010_VARIETY_DEDUCTION_COUNT_VALUE",
                        "placement and processed counts must be integers. "
                        "Path = " + pszInputPath + ", Line = "
                        + str(iRowIndex),
                    )
    return arrRows


def create_step0010_variety_deduction_rows(
    arrInputRows: List[List[str]],
) -> List[List[str]]:
    """控除廃棄版品種別step0009に年間合計列を追加します。"""
    iItemsPerVariety: int = len(STEP0008_VARIETY_DEDUCTION_ITEM_NAMES)
    arrOutputRows: List[List[str]] = [list(arrRow) for arrRow in arrInputRows]
    arrOutputRows[0].append("")
    arrOutputRows[1].append("合計")
    for iVarietyIndex in range(len(SPLIT_VARIETIES) + 1):
        iVarietyStart: int = 2 + iVarietyIndex * iItemsPerVariety
        arrTotals: List[Decimal] = [
            sum(
                (
                    Decimal(pszValue.strip(" \u3000\t"))
                    for pszValue in arrInputRows[
                        iVarietyStart + iItemIndex
                    ][2:]
                ),
                Decimal("0"),
            )
            for iItemIndex in range(3)
        ]
        objAverageWeight: Decimal = (
            (arrTotals[2] / arrTotals[1]).quantize(
                Decimal("0.001"), rounding=ROUND_HALF_UP
            )
            if arrTotals[1] != 0 else Decimal("0")
        )
        objAgeWeightedTotal: Decimal = Decimal("0")
        objAgeWeightTotal: Decimal = Decimal("0")
        for iMonthColumnIndex in range(2, len(arrInputRows[0])):
            pszAge: str = arrInputRows[iVarietyStart + 4][
                iMonthColumnIndex
            ].strip(" \u3000\t")
            if pszAge:
                objProcessedCount: Decimal = Decimal(
                    arrInputRows[iVarietyStart + 1][
                        iMonthColumnIndex
                    ].strip(" \u3000\t")
                )
                objAgeWeightedTotal += Decimal(pszAge) * objProcessedCount
                objAgeWeightTotal += objProcessedCount
        objAverageAge: Decimal = (
            (objAgeWeightedTotal / objAgeWeightTotal).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            if objAgeWeightTotal != 0 else Decimal("0")
        )
        objGrowthRate: Decimal = (
            (arrTotals[1] / arrTotals[0]).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            )
            if arrTotals[0] != 0 else Decimal("0")
        )
        arrAnnualValues: List[Decimal] = (
            arrTotals[:3]
            + [objAverageWeight, objAverageAge, objGrowthRate]
        )
        for iItemIndex, objValue in enumerate(arrAnnualValues):
            arrOutputRows[iVarietyStart + iItemIndex].append(
                format_sum_value(objValue)
            )
    return arrOutputRows


def create_step0010_variety_deduction_output_path(
    pszStep0009Path: str,
) -> str:
    """控除廃棄版品種別step0009のstep0010出力先を作成します。"""
    pszDirectory: str = os.path.dirname(pszStep0009Path)
    pszFileName: str = os.path.basename(pszStep0009Path)
    if "_step0009_品種別_控除廃棄_" not in pszFileName:
        raise_conversion_error(
            "INVALID_STEP0009_VARIETY_DEDUCTION_FOR_STEP0010",
            "variety deduction step0009 marker was not found. Path = "
            + pszStep0009Path,
        )
    pszFileName = pszFileName.replace("_step0009_", "_step0010_", 1)
    pszStem, pszExtension = os.path.splitext(pszFileName)
    pszOutputPath: str = os.path.join(pszDirectory, pszFileName)
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0010_variety_deduction_annual_files(
    arrStep0009VarietyDeductionAnnualPaths: List[str],
) -> List[str]:
    """控除廃棄版品種別step0009から年間合計列付きstep0010を作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0009Path in arrStep0009VarietyDeductionAnnualPaths:
            arrInputRows = read_step0009_variety_deduction_for_step0010(
                pszStep0009Path
            )
            pszOutputPath = create_step0010_variety_deduction_output_path(
                pszStep0009Path
            )
            arrOutputRows = create_step0010_variety_deduction_rows(
                arrInputRows
            )
            write_step0010_tsv(pszOutputPath, arrOutputRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0010_VARIETY_DEDUCTION_CLEANUP_FAILED",
                "variety deduction step0010 creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0010_variety_deduction_for_excel(
    pszInputPath: str,
) -> List[List[str]]:
    """控除廃棄版品種別step0010年間TSVを検証して読み込みます。"""
    objFileNameMatch = re.search(
        r"_step0010_品種別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_DEDUCTION_FILENAME",
            "unexpected variety deduction step0010 file name. Path = "
            + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0010_VARIETY_DEDUCTION_READ_FAILED",
            "failed to read variety deduction step0010 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    arrItemNames: List[str] = STEP0008_VARIETY_DEDUCTION_ITEM_NAMES
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for pszVarietyName in [pszName for _, pszName in SPLIT_VARIETIES] + ["合計"]:
        for iItemIndex, pszItemName in enumerate(arrItemNames):
            arrExpectedStructure.append(
                (pszVarietyName if iItemIndex == 0 else "", pszItemName)
            )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_DEDUCTION_ROW_COUNT",
            "variety deduction step0010 must contain 38 rows. Path = "
            + pszInputPath,
        )
    if any(len(arrRow) != 15 for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_DEDUCTION_COLUMN_COUNT",
            "variety deduction step0010 rows must contain exactly 15 columns. "
            "Path = " + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_DEDUCTION_ROW_STRUCTURE",
                "unexpected variety deduction step0010 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
    if arrRows[0][14].strip(" \u3000\t") != "" or arrRows[1][14].strip(
        " \u3000\t"
    ) != "合計":
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_DEDUCTION_TOTAL_HEADER",
            "variety deduction step0010 annual column header is invalid. "
            "Path = " + pszInputPath,
        )
    arrMonths: List[int] = []
    for iColumnIndex in range(2, 14):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_DEDUCTION_YEAR",
                "variety deduction step0010 year did not match its file name. "
                "Path = " + pszInputPath,
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_DEDUCTION_MONTH",
                "variety deduction step0010 months must be unique and "
                "ascending. Path = " + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            iItemOffset: int = (iRowIndex - 2) % len(arrItemNames)
            pszItemName: str = arrItemNames[iItemOffset]
            bAllowBlank: bool = STEP0004_VARIETY_DEDUCTION_ITEMS[iItemOffset][2]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1, iColumnIndex + 1, pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0010_VARIETY_DEDUCTION_ITEM_VALUE",
                    "invalid variety deduction step0010 item value. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1)
                    + ". Detail = " + str(objException),
                )
            if objValue is None and not bAllowBlank:
                raise_conversion_error(
                    "INVALID_STEP0010_VARIETY_DEDUCTION_ITEM_VALUE",
                    "variety deduction step0010 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1),
                )
    if arrMonths != list(range(1, 13)):
        raise_conversion_error(
            "INVALID_STEP0010_VARIETY_DEDUCTION_MONTH",
            "variety deduction step0010 must contain months 1 through 12. "
            "Path = " + pszInputPath,
        )
    for iRowIndex in range(2, len(arrRows)):
        iItemOffset = (iRowIndex - 2) % len(arrItemNames)
        pszItemName = arrItemNames[iItemOffset]
        try:
            objAnnualValue = parse_decimal_value(
                arrRows[iRowIndex][14], iRowIndex + 1, 15, pszItemName
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_DEDUCTION_ITEM_VALUE",
                "invalid variety deduction step0010 annual value. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1)
                + ". Detail = " + str(objException),
            )
        if objAnnualValue is None:
            raise_conversion_error(
                "INVALID_STEP0010_VARIETY_DEDUCTION_ITEM_VALUE",
                "variety deduction step0010 annual value was blank. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
        if pszItemName in {"入雛羽数", "処理羽数"}:
            for pszValue in arrRows[iRowIndex][2:]:
                objCount: Decimal = Decimal(pszValue.strip(" \u3000\t"))
                if objCount != objCount.to_integral_value():
                    raise_conversion_error(
                        "INVALID_STEP0010_VARIETY_DEDUCTION_COUNT_VALUE",
                        "variety deduction step0010 counts must be integers. "
                        "Path = " + pszInputPath + ", Line = "
                        + str(iRowIndex + 1),
                    )
    return arrRows


def read_step0013_variety_for_excel(pszInputPath: str) -> List[List[str]]:
    """前年実績・前年比付き通常版年間表を検証します。"""
    objMatch = re.fullmatch(
        r"(.+)_step0013_品種別_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error("INVALID_STEP0013_VARIETY_EXCEL_FILENAME", pszInputPath)
    iYear = int(objMatch.group(2))
    try:
        with open(pszInputPath, "r", encoding="utf-8-sig", newline="") as objFile:
            arrRows = [strip_line_ending(line).split("\t") for line in objFile]
    except (OSError, UnicodeError) as objException:
        raise_conversion_error("STEP0013_VARIETY_EXCEL_READ_FAILED", str(objException))
    arrStructure = [("", "年"), ("", "月")]
    for pszName in [name for _, name in SPLIT_VARIETIES] + ["合計"]:
        arrStructure.extend((pszName if i == 0 else "", item)
                            for i, item in enumerate(STEP0008_VARIETY_ITEM_NAMES))
    arrStructure.append(("", "稼動日数"))
    if iYear != 2022:
        arrStructure += [("出荷羽数", "前年比"), ("", "(一日平均)"),
                         ("前年実績", "入雛羽数"), (f"{iYear - 1}年", "出荷羽数"),
                         ("", "正味重量"), ("", "平均体重"),
                         ("", "飼育日数平均"), ("", "育成率"), ("", "稼動日数")]
    if (len(arrRows) != len(arrStructure) or any(len(row) != 15 for row in arrRows)
            or [tuple(row[:2]) for row in arrRows] != arrStructure):
        raise_conversion_error("INVALID_STEP0013_VARIETY_EXCEL_STRUCTURE", pszInputPath)
    if (arrRows[0] != ["", "年"] + [str(iYear)] * 12 + [""]
            or arrRows[1] != ["", "月"] + [str(m) for m in range(1, 13)] + ["合計"]):
        raise_conversion_error("INVALID_STEP0013_VARIETY_EXCEL_HEADER", pszInputPath)
    for iRow, row in enumerate(arrRows[2:], start=2):
        if row[1] == "(一日平均)":
            if any(row[2:]):
                raise_conversion_error("INVALID_STEP0013_VARIETY_EXCEL_NOTE", pszInputPath)
            continue
        for value in row[2:]:
            if value == "":
                continue
            try:
                objValue = Decimal(value)
                if not objValue.is_finite():
                    raise ValueError("non-finite value")
                if row[1] in {"入雛羽数", "出荷羽数", "稼動日数"} and objValue != objValue.to_integral_value():
                    raise ValueError("count must be integer")
                if row[1] == "稼動日数" and objValue < 0:
                    raise ValueError("negative operating days")
            except (ArithmeticError, ValueError):
                raise_conversion_error("INVALID_STEP0013_VARIETY_EXCEL_VALUE",
                                       pszInputPath + ": " + str(iRow + 1))
    return arrRows


def create_step0013_variety_excel_output_path(
    pszStep0013TsvPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszStep0013TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0013TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_step0013_variety_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0013_VARIETY_EXCEL_CREATE_FAILED",
                "variety step0013 Excel already exists. Path = "
                + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0013_VARIETY_EXCEL_CREATE_FAILED",
                "failed to reserve variety step0013 Excel. Detail = "
                + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0013_variety_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "品種別"
            pszCurrentVariety: str = ""
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                if arrFields[0].strip(" \u3000\t"):
                    pszCurrentVariety = arrFields[0].strip(" \u3000\t")
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2 or (
                        iRowIndex == 2 and iColumnIndex == 15
                    ):
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    if (
                        pszCurrentVariety == BROILER_BREAKDOWNS[1][1]
                        and pszRowType == "飼育日数平均"
                    ):
                        continue
                    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue: Decimal = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue) if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszRowType in {"年", "月", "入雛羽数", "出荷羽数", "稼動日数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "前年比":
                        objCell.number_format = "0.0000"
                    elif pszRowType == "正味重量":
                        iDecimalPlaces: int = max(
                            0, -objValue.as_tuple().exponent
                        )
                        objCell.number_format = (
                            "0." + "#" * iDecimalPlaces
                            if iDecimalPlaces else "0"
                        )
                    elif pszRowType in {"平均体重", "飼育日数平均"}:
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0.###"
                        )
                    elif pszRowType in {"育成率", "育成率"}:
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0.####"
                        )
            for iColumnIndex in range(1, objWorksheet.max_column + 1):
                iMaximumDisplayWidth = 0
                for objWidthColumn in objWorksheet.iter_cols(
                    min_col=iColumnIndex,
                    max_col=iColumnIndex,
                    min_row=1,
                    max_row=objWorksheet.max_row,
                ):
                    for objWidthCell in objWidthColumn:
                        objWidthValue = objWidthCell.value
                        if objWidthValue is None:
                            continue
                        if isinstance(objWidthValue, (int, float, Decimal)):
                            objWidthDecimal = Decimal(str(objWidthValue))
                            pszFractionFormat = (
                                objWidthCell.number_format.split(";")[0]
                                .partition(".")[2]
                            )
                            iFractionPlaces = sum(
                                1 for pszCharacter in pszFractionFormat
                                if pszCharacter in "0#"
                            )
                            pszDisplayValue = format(
                                objWidthDecimal,
                                "." + str(iFractionPlaces) + "f",
                            )
                            if iFractionPlaces:
                                iMinimumFractionPlaces = pszFractionFormat.count("0")
                                while (
                                    len(pszDisplayValue.rsplit(".", 1)[-1])
                                    > iMinimumFractionPlaces
                                    and pszDisplayValue.endswith("0")
                                ):
                                    pszDisplayValue = pszDisplayValue[:-1]
                                if pszDisplayValue.endswith("."):
                                    pszDisplayValue = pszDisplayValue[:-1]
                        else:
                            pszDisplayValue = str(objWidthValue)
                        iDisplayWidth = sum(
                            2 if unicodedata.east_asian_width(pszCharacter)
                            in {"F", "W"}
                            else 1
                            for pszCharacter in pszDisplayValue
                        )
                        iMaximumDisplayWidth = max(
                            iMaximumDisplayWidth, iDisplayWidth
                        )
                objWorksheet.column_dimensions[
                    get_column_letter(iColumnIndex)
                ].width = min(max(iMaximumDisplayWidth + 2, 8), 255)
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                "STEP0013_VARIETY_EXCEL_WRITE_FAILED",
                "failed to write variety step0013 Excel. Detail = "
                + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0013_VARIETY_EXCEL_FINALIZE_FAILED",
                "failed to finalize variety step0013 Excel. Detail = "
                + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0013_variety_excel_files(
    arrStep0013VarietyTsvPaths: List[str],
) -> List[str]:
    """通常版品種別step0013年間TSVからExcelを作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0013TsvPath in arrStep0013VarietyTsvPaths:
            arrRows = read_step0013_variety_for_excel(
                pszStep0013TsvPath
            )
            pszOutputPath = create_step0013_variety_excel_output_path(
                pszStep0013TsvPath
            )
            write_step0013_variety_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0013_VARIETY_EXCEL_CLEANUP_FAILED",
                "variety step0013 Excel creation failed. "
                "Original error = " + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0013_variety_deduction_for_excel(pszInputPath: str) -> List[List[str]]:
    """前年実績・前年比付き控除廃棄版年間表を検証します。"""
    objMatch = re.fullmatch(
        r"(.+)_step0013_品種別_控除廃棄_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error("INVALID_STEP0013_VARIETY_DEDUCTION_EXCEL_FILENAME", pszInputPath)
    iYear = int(objMatch.group(2))
    try:
        with open(pszInputPath, "r", encoding="utf-8-sig", newline="") as objFile:
            arrRows = [strip_line_ending(line).split("\t") for line in objFile]
    except (OSError, UnicodeError) as objException:
        raise_conversion_error("STEP0013_VARIETY_DEDUCTION_EXCEL_READ_FAILED", str(objException))
    arrStructure = [("", "年"), ("", "月")]
    for pszName in [name for _, name in SPLIT_VARIETIES] + ["合計"]:
        arrStructure.extend((pszName if i == 0 else "", item)
                            for i, item in enumerate(STEP0008_VARIETY_DEDUCTION_ITEM_NAMES))
    arrStructure.append(("", "稼動日数"))
    if iYear != 2022:
        arrStructure += [("処理羽数", "前年比"), ("", "(一日平均)"),
                         ("前年実績", "入雛羽数"), (f"{iYear - 1}年", "処理羽数"),
                         ("", "処理重量"), ("", "控除後平均体重"),
                         ("", "飼育日数平均"), ("", "育成率"), ("", "稼動日数")]
    if (len(arrRows) != len(arrStructure) or any(len(row) != 15 for row in arrRows)
            or [tuple(row[:2]) for row in arrRows] != arrStructure):
        raise_conversion_error("INVALID_STEP0013_VARIETY_DEDUCTION_EXCEL_STRUCTURE", pszInputPath)
    if (arrRows[0] != ["", "年"] + [str(iYear)] * 12 + [""]
            or arrRows[1] != ["", "月"] + [str(m) for m in range(1, 13)] + ["合計"]):
        raise_conversion_error("INVALID_STEP0013_VARIETY_DEDUCTION_EXCEL_HEADER", pszInputPath)
    for iRow, row in enumerate(arrRows[2:], start=2):
        if row[1] == "(一日平均)":
            if any(row[2:]):
                raise_conversion_error("INVALID_STEP0013_VARIETY_DEDUCTION_EXCEL_NOTE", pszInputPath)
            continue
        for value in row[2:]:
            if value == "":
                continue
            try:
                objValue = Decimal(value)
                if not objValue.is_finite():
                    raise ValueError("non-finite value")
                if row[1] in {"入雛羽数", "処理羽数", "稼動日数"} and objValue != objValue.to_integral_value():
                    raise ValueError("count must be integer")
                if row[1] == "稼動日数" and objValue < 0:
                    raise ValueError("negative operating days")
            except (ArithmeticError, ValueError):
                raise_conversion_error("INVALID_STEP0013_VARIETY_DEDUCTION_EXCEL_VALUE",
                                       pszInputPath + ": " + str(iRow + 1))
    return arrRows


def create_step0013_variety_deduction_excel_output_path(
    pszStep0013TsvPath: str,
) -> str:
    pszDirectory: str = os.path.dirname(pszStep0013TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0013TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_step0013_variety_deduction_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0010_VARIETY_DEDUCTION_EXCEL_CREATE_FAILED",
                "variety deduction step0013 Excel already exists. Path = "
                + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0010_VARIETY_DEDUCTION_EXCEL_CREATE_FAILED",
                "failed to reserve variety deduction step0013 Excel. Detail = "
                + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0013_variety_deduction_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "品種別"
            pszCurrentVariety: str = ""
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                if arrFields[0].strip(" \u3000\t"):
                    pszCurrentVariety = arrFields[0].strip(" \u3000\t")
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2 or (
                        iRowIndex == 2 and iColumnIndex == 15
                    ):
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    if (
                        pszCurrentVariety == BROILER_BREAKDOWNS[1][1]
                        and pszRowType == "飼育日数平均"
                    ):
                        continue
                    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue: Decimal = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue) if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszRowType in {"年", "月", "入雛羽数", "処理羽数", "稼動日数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "前年比":
                        objCell.number_format = "0.0000"
                    elif pszRowType == "処理重量":
                        iDecimalPlaces: int = max(
                            0, -objValue.as_tuple().exponent
                        )
                        objCell.number_format = (
                            "0." + "#" * iDecimalPlaces
                            if iDecimalPlaces else "0"
                        )
                    elif pszRowType in {"控除後平均体重", "飼育日数平均"}:
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0.###"
                        )
                    elif pszRowType in {"廃棄控除後育成率", "育成率"}:
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0.####"
                        )
            for iColumnIndex in range(1, objWorksheet.max_column + 1):
                iMaximumDisplayWidth = 0
                for objWidthColumn in objWorksheet.iter_cols(
                    min_col=iColumnIndex,
                    max_col=iColumnIndex,
                    min_row=1,
                    max_row=objWorksheet.max_row,
                ):
                    for objWidthCell in objWidthColumn:
                        objWidthValue = objWidthCell.value
                        if objWidthValue is None:
                            continue
                        if isinstance(objWidthValue, (int, float, Decimal)):
                            objWidthDecimal = Decimal(str(objWidthValue))
                            pszFractionFormat = (
                                objWidthCell.number_format.split(";")[0]
                                .partition(".")[2]
                            )
                            iFractionPlaces = sum(
                                1 for pszCharacter in pszFractionFormat
                                if pszCharacter in "0#"
                            )
                            pszDisplayValue = format(
                                objWidthDecimal,
                                "." + str(iFractionPlaces) + "f",
                            )
                            if iFractionPlaces:
                                iMinimumFractionPlaces = pszFractionFormat.count("0")
                                while (
                                    len(pszDisplayValue.rsplit(".", 1)[-1])
                                    > iMinimumFractionPlaces
                                    and pszDisplayValue.endswith("0")
                                ):
                                    pszDisplayValue = pszDisplayValue[:-1]
                                if pszDisplayValue.endswith("."):
                                    pszDisplayValue = pszDisplayValue[:-1]
                        else:
                            pszDisplayValue = str(objWidthValue)
                        iDisplayWidth = sum(
                            2 if unicodedata.east_asian_width(pszCharacter)
                            in {"F", "W"}
                            else 1
                            for pszCharacter in pszDisplayValue
                        )
                        iMaximumDisplayWidth = max(
                            iMaximumDisplayWidth, iDisplayWidth
                        )
                objWorksheet.column_dimensions[
                    get_column_letter(iColumnIndex)
                ].width = min(max(iMaximumDisplayWidth + 2, 8), 255)
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                "STEP0010_VARIETY_DEDUCTION_EXCEL_WRITE_FAILED",
                "failed to write variety deduction step0013 Excel. Detail = "
                + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0010_VARIETY_DEDUCTION_EXCEL_FINALIZE_FAILED",
                "failed to finalize variety deduction step0013 Excel. Detail = "
                + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0013_variety_deduction_excel_files(
    arrStep0013VarietyDeductionTsvPaths: List[str],
) -> List[str]:
    """控除廃棄版品種別step0013年間TSVからExcelを作成します。"""
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0013TsvPath in arrStep0013VarietyDeductionTsvPaths:
            arrRows = read_step0013_variety_deduction_for_excel(
                pszStep0013TsvPath
            )
            pszOutputPath = create_step0013_variety_deduction_excel_output_path(
                pszStep0013TsvPath
            )
            write_step0013_variety_deduction_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0010_VARIETY_DEDUCTION_EXCEL_CLEANUP_FAILED",
                "variety deduction step0013 Excel creation failed. "
                "Original error = " + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0010_for_excel(
    pszInputPath: str,
) -> List[List[str]]:
    objFileNameMatch = re.search(
        r"_step0010_グループ別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0010_FILENAME",
            "unexpected step0010 file name. Path = " + pszInputPath,
        )
    iExpectedYear: int = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows: List[List[str]] = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0010_READ_FAILED",
            "failed to read step0010 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    arrExpectedStructure: List[Tuple[str, str]] = [("", "年"), ("", "月")]
    for pszGroupName in [pszName for _, pszName in GROUPS] + ["合計"]:
        arrExpectedStructure.extend(
            [(pszGroupName, "入雛羽数"), ("", "出荷羽数"),
             ("", "正味重量"), ("", "平均体重"), ("", "育成率")]
        )
    if len(arrRows) != len(arrExpectedStructure):
        raise_conversion_error(
            "INVALID_STEP0010_ROW_COUNT",
            "step0010 must contain 22 rows. Path = " + pszInputPath,
        )
    iColumnCount: int = len(arrRows[0])
    if iColumnCount != 15 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0010_COLUMN_COUNT",
            "step0010 rows must contain exactly 15 columns. Path = "
            + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(arrExpectedStructure):
        if tuple(pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]) != objExpected:
            raise_conversion_error(
                "INVALID_STEP0010_ROW_STRUCTURE",
                "unexpected step0010 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
    if arrRows[0][14] != "" or arrRows[1][14] != "合計":
        raise_conversion_error(
            "INVALID_STEP0010_TOTAL_HEADER",
            "step0010 annual column must have a blank year and a total label. Path = " + pszInputPath,
        )
    arrMonths: List[int] = []
    for iColumnIndex in range(2, 14):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear:
            raise_conversion_error(
                "INVALID_STEP0010_YEAR",
                "step0010 year did not match its file name. Path = " + pszInputPath,
            )
        if arrMonths and iMonth <= arrMonths[-1]:
            raise_conversion_error(
                "INVALID_STEP0010_MONTH",
                "step0010 months must be unique and ascending. Path = " + pszInputPath,
            )
        arrMonths.append(iMonth)
        for iRowIndex in range(2, len(arrRows)):
            pszItemName: str = arrExpectedStructure[iRowIndex][1]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1, iColumnIndex, pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error("INVALID_STEP0010_ITEM_VALUE", str(objException))
            if objValue is None and not (
                iRowIndex < 17 and pszItemName in {"平均体重", "育成率"}
            ):
                raise_conversion_error(
                    "INVALID_STEP0010_ITEM_VALUE",
                    "step0010 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1),
                )
    if arrMonths != list(range(1, 13)):
        raise_conversion_error(
            "INVALID_STEP0010_MONTH",
            "step0010 must contain months 1 through 12. Path = " + pszInputPath,
        )
    for iRowIndex in range(2, len(arrRows)):
        pszItemName = arrExpectedStructure[iRowIndex][1]
        try:
            objAnnualValue = parse_decimal_value(
                arrRows[iRowIndex][14], iRowIndex + 1, 14, pszItemName
            )
        except ConversionError as objException:
            raise_conversion_error("INVALID_STEP0010_ITEM_VALUE", str(objException))
        if objAnnualValue is None:
            raise_conversion_error(
                "INVALID_STEP0010_ITEM_VALUE",
                "step0010 annual value was blank. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )
        if pszItemName in {"入雛羽数", "出荷羽数"}:
            for iColumnIndex in range(2, 15):
                objCount: Decimal = Decimal(arrRows[iRowIndex][iColumnIndex].strip(" \u3000\t"))
                if objCount != objCount.to_integral_value():
                    raise_conversion_error(
                        "INVALID_STEP0010_COUNT_VALUE",
                        "step0010 counts must be integers. Path = "
                        + pszInputPath + ", Line = " + str(iRowIndex + 1),
                    )
    return arrRows


def create_step0010_excel_output_path(pszStep0010TsvPath: str) -> str:
    pszDirectory: str = os.path.dirname(pszStep0010TsvPath)
    pszStem: str = os.path.splitext(os.path.basename(pszStep0010TsvPath))[0]
    pszOutputPath: str = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix: int = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_step0010_excel(
    pszOutputPath: str, arrRows: List[List[str]]
) -> None:
    pszTemporaryPath: str = ""
    bOutputReserved: bool = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor: int = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                "STEP0010_EXCEL_CREATE_FAILED",
                "step0010 Excel already exists. Path = " + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                "STEP0010_EXCEL_CREATE_FAILED",
                "failed to reserve step0010 Excel. Detail = " + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix="group_shipment_summary_step0010_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "グループ別"
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                pszRowType: str = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2:
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    if iRowIndex == 2 and iColumnIndex == 15:
                        objCell.value = pszValue
                        continue
                    pszNormalizedValue: str = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue: Decimal = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue) if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszRowType in {"年", "入雛羽数", "出荷羽数"}:
                        objCell.number_format = "0"
                    elif pszRowType == "月":
                        objCell.number_format = "0"
                    elif pszRowType == "正味重量":
                        iDecimalPlaces: int = max(0, -objValue.as_tuple().exponent)
                        objCell.number_format = (
                            "0." + "#" * iDecimalPlaces if iDecimalPlaces else "0"
                        )
                    elif pszRowType == "平均体重":
                        iDecimalPlaces = max(3, -objValue.as_tuple().exponent) if iColumnIndex < 15 else 3
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0." + "#" * iDecimalPlaces
                        )
                    elif pszRowType == "育成率":
                        iDecimalPlaces = max(4, -objValue.as_tuple().exponent) if iColumnIndex < 15 else 4
                        objCell.number_format = (
                            "0" if objValue == objValue.to_integral_value()
                            else "0." + "#" * iDecimalPlaces
                        )
            for iColumnIndex in range(1, objWorksheet.max_column + 1):
                iMaximumDisplayWidth = 0
                for objWidthColumn in objWorksheet.iter_cols(
                    min_col=iColumnIndex,
                    max_col=iColumnIndex,
                    min_row=1,
                    max_row=objWorksheet.max_row,
                ):
                    for objWidthCell in objWidthColumn:
                        objWidthValue = objWidthCell.value
                        if objWidthValue is None:
                            continue
                        if isinstance(objWidthValue, (int, float, Decimal)):
                            objWidthDecimal = Decimal(str(objWidthValue))
                            pszFractionFormat = (
                                objWidthCell.number_format.split(";")[0]
                                .partition(".")[2]
                            )
                            iFractionPlaces = sum(
                                1 for pszCharacter in pszFractionFormat
                                if pszCharacter in "0#"
                            )
                            pszDisplayValue = format(
                                objWidthDecimal,
                                "." + str(iFractionPlaces) + "f",
                            )
                            if iFractionPlaces:
                                iMinimumFractionPlaces = pszFractionFormat.count("0")
                                while (
                                    len(pszDisplayValue.rsplit(".", 1)[-1])
                                    > iMinimumFractionPlaces
                                    and pszDisplayValue.endswith("0")
                                ):
                                    pszDisplayValue = pszDisplayValue[:-1]
                                if pszDisplayValue.endswith("."):
                                    pszDisplayValue = pszDisplayValue[:-1]
                        else:
                            pszDisplayValue = str(objWidthValue)
                        iDisplayWidth = sum(
                            2 if unicodedata.east_asian_width(pszCharacter)
                            in {"F", "W"}
                            else 1
                            for pszCharacter in pszDisplayValue
                        )
                        iMaximumDisplayWidth = max(
                            iMaximumDisplayWidth, iDisplayWidth
                        )
                objWorksheet.column_dimensions[
                    get_column_letter(iColumnIndex)
                ].width = min(max(iMaximumDisplayWidth + 2, 8), 255)
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                "STEP0010_EXCEL_WRITE_FAILED",
                "failed to write step0010 Excel. Detail = " + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                "STEP0010_EXCEL_FINALIZE_FAILED",
                "failed to finalize step0010 Excel. Detail = " + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0010_excel_files(
    arrStep0010TsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0010TsvPath in arrStep0010TsvPaths:
            arrRows = read_step0010_for_excel(pszStep0010TsvPath)
            pszOutputPath: str = create_step0010_excel_output_path(
                pszStep0010TsvPath
            )
            write_step0010_excel(pszOutputPath, arrRows)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0010_EXCEL_CLEANUP_FAILED",
                "step0010 Excel creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


STEP0013_GROUP_EXCEL_CURRENT_STRUCTURE: List[Tuple[str, str]] = [
    ("", "年"), ("", "月")
]
for _, pszGroupName in GROUPS:
    STEP0013_GROUP_EXCEL_CURRENT_STRUCTURE.extend(
        [
            (pszGroupName, "入雛羽数"),
            ("", "出荷羽数"),
            ("", "正味重量"),
            ("", "平均体重"),
            ("", "育成率"),
        ]
    )
STEP0013_GROUP_EXCEL_CURRENT_STRUCTURE.extend(
    [
        ("合計", "入雛羽数"),
        ("", "出荷羽数"),
        ("", "正味重量"),
        ("", "平均体重"),
        ("", "育成率"),
        ("", "稼動日数"),
    ]
)


def read_step0013_group_for_excel(pszInputPath: str) -> List[List[str]]:
    objFileNameMatch = re.search(
        r"_step0013_グループ別_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_EXCEL_FILENAME",
            "unexpected step0013 group file name. Path = " + pszInputPath,
        )
    iExpectedYear = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8-sig", newline="") as objFile:
            arrRows = [strip_line_ending(pszLine).split("\t") for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0013_GROUP_EXCEL_READ_FAILED",
            "failed to read step0013 group TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )

    bHasPreviousYear = iExpectedYear != 2022
    iExpectedRowCount = 31 if bHasPreviousYear else 23
    if len(arrRows) != iExpectedRowCount:
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_EXCEL_ROW_COUNT",
            "step0013 group TSV must contain " + str(iExpectedRowCount)
            + " rows. Path = " + pszInputPath,
        )
    if any(len(arrRow) != 15 for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_EXCEL_COLUMN_COUNT",
            "step0013 group TSV rows must contain exactly 15 columns. Path = "
            + pszInputPath,
        )

    for iRowIndex, objExpected in enumerate(STEP0013_GROUP_EXCEL_CURRENT_STRUCTURE):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0013_GROUP_EXCEL_ROW_STRUCTURE",
                "unexpected step0013 group row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )

    if arrRows[0][14] != "" or arrRows[1][14] != "合計":
        raise_conversion_error(
            "INVALID_STEP0013_GROUP_EXCEL_TOTAL_HEADER",
            "step0013 group annual header is invalid. Path = " + pszInputPath,
        )
    for iColumnIndex in range(2, 14):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear or iMonth != iColumnIndex - 1:
            raise_conversion_error(
                "INVALID_STEP0013_GROUP_EXCEL_YEAR_MONTH",
                "step0013 group year or month did not match its file name or column. Path = "
                + pszInputPath + ", Column = " + str(iColumnIndex + 1),
            )

    def validate_numeric_cell(iRowIndex: int, iColumnIndex: int, pszItemName: str,
                              bAllowBlank: bool) -> None:
        try:
            objValue = parse_decimal_value(
                arrRows[iRowIndex][iColumnIndex],
                iRowIndex + 1,
                iColumnIndex,
                pszItemName,
            )
        except ConversionError as objException:
            raise_conversion_error(
                "INVALID_STEP0013_GROUP_EXCEL_ITEM_VALUE",
                "invalid step0013 group value. Path = " + pszInputPath
                + ", Line = " + str(iRowIndex + 1)
                + ", Column = " + str(iColumnIndex + 1)
                + ", Detail = " + str(objException),
            )
        if objValue is None and not bAllowBlank:
            raise_conversion_error(
                "INVALID_STEP0013_GROUP_EXCEL_ITEM_VALUE",
                "step0013 group value was blank. Path = " + pszInputPath
                + ", Line = " + str(iRowIndex + 1)
                + ", Column = " + str(iColumnIndex + 1),
            )
        if (objValue is not None
                and pszItemName in {"入雛羽数", "出荷羽数", "稼動日数"}
                and objValue != objValue.to_integral_value()):
            raise_conversion_error(
                "INVALID_STEP0013_GROUP_EXCEL_COUNT_VALUE",
                "step0013 group counts must be integers. Path = " + pszInputPath
                + ", Line = " + str(iRowIndex + 1)
                + ", Column = " + str(iColumnIndex + 1),
            )

    for iRowIndex in range(2, 23):
        pszItemName = STEP0013_GROUP_EXCEL_CURRENT_STRUCTURE[iRowIndex][1]
        for iColumnIndex in range(2, 15):
            validate_numeric_cell(
                iRowIndex,
                iColumnIndex,
                pszItemName,
                pszItemName in {"平均体重", "育成率", "稼動日数"}
                and iColumnIndex < 14,
            )

    if bHasPreviousYear:
        if (arrRows[23][:2] != ["出荷羽数", "前年比"]
                or arrRows[24][:2] != ["", "(一日平均)"]
                or any(arrRows[24][2:])):
            raise_conversion_error(
                "INVALID_STEP0013_GROUP_EXCEL_RATIO_ROW",
                "step0013 group ratio rows are invalid. Path = " + pszInputPath,
            )
        for iColumnIndex in range(2, 15):
            validate_numeric_cell(23, iColumnIndex, "前年比", True)
        arrPreviousStructure = [
            ("前年実績", "入雛羽数"),
            (f"{iExpectedYear - 1}年", "出荷羽数"),
            ("", "正味重量"),
            ("", "平均体重"),
            ("", "育成率"),
            ("", "稼動日数"),
        ]
        for iOffset, objExpected in enumerate(arrPreviousStructure):
            iRowIndex = 25 + iOffset
            objActual = tuple(
                pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
            )
            if objActual != objExpected:
                raise_conversion_error(
                    "INVALID_STEP0013_GROUP_EXCEL_PREVIOUS_STRUCTURE",
                    "step0013 group previous-year row structure is invalid. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1),
                )
            pszItemName = objExpected[1]
            for iColumnIndex in range(2, 15):
                validate_numeric_cell(
                    iRowIndex,
                    iColumnIndex,
                    pszItemName,
                    pszItemName in {"平均体重", "育成率", "稼動日数"}
                    and iColumnIndex < 14,
                )
    return arrRows


def create_step0013_group_excel_output_path(pszStep0013TsvPath: str) -> str:
    pszDirectory = os.path.dirname(pszStep0013TsvPath)
    pszStem = os.path.splitext(os.path.basename(pszStep0013TsvPath))[0]
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx")
        iSuffix += 1
    return pszOutputPath


def create_step0013_group_excel_files(
    arrStep0013TsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0013TsvPath in arrStep0013TsvPaths:
            arrRows = read_step0013_group_for_excel(pszStep0013TsvPath)
            pszOutputPath = create_step0013_group_excel_output_path(
                pszStep0013TsvPath
            )
            write_group_deduction_excel(pszOutputPath, arrRows, 13)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0013_GROUP_EXCEL_CLEANUP_FAILED",
                "step0013 group Excel creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


STEP0013_EXCEL_ROW_STRUCTURE: List[Tuple[str, str]] = [("", "年"), ("", "月")]
for _, pszGroupName in GROUPS:
    STEP0013_EXCEL_ROW_STRUCTURE.extend(
        [
            (pszGroupName, "入雛羽数"),
            ("", "処理羽数"),
            ("", "処理重量"),
            ("", "平均体重"),
            ("", "育成率"),
        ]
    )
STEP0013_EXCEL_ROW_STRUCTURE.extend(
    [
        ("合計", "入雛羽数"),
        ("", "処理羽数"),
        ("", "処理重量"),
        ("", "平均体重"),
        ("", "育成率"),
    ]
)


def read_step0013_for_excel(pszInputPath: str) -> List[List[str]]:
    objFileNameMatch = re.search(
        r"_step0013_グループ別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objFileNameMatch is None:
        raise_conversion_error(
            "INVALID_STEP0013_EXCEL_FILENAME",
            "unexpected step0013 group deduction file name. Path = " + pszInputPath,
        )
    iExpectedYear = int(objFileNameMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [strip_line_ending(pszLine).split("\t") for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0013_EXCEL_READ_FAILED",
            "failed to read step0013 TSV. Path = "
            + pszInputPath + ", Detail = " + str(objException),
        )
    if len(arrRows) != len(STEP0013_EXCEL_ROW_STRUCTURE):
        raise_conversion_error(
            "INVALID_STEP0013_EXCEL_ROW_COUNT",
            "step0013 TSV must contain 22 rows. Path = " + pszInputPath,
        )
    iColumnCount = len(arrRows[0])
    if iColumnCount != 14 or any(len(arrRow) != iColumnCount for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0013_EXCEL_COLUMN_COUNT",
            "step0013 TSV must contain 14 columns with equal row lengths. Path = "
            + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(STEP0013_EXCEL_ROW_STRUCTURE):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0013_EXCEL_ROW_STRUCTURE",
                "unexpected step0013 row structure. Path = "
                + pszInputPath + ", Line = " + str(iRowIndex + 1),
            )

    for iColumnIndex in range(2, iColumnCount):
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear or iMonth != iColumnIndex - 1:
            raise_conversion_error(
                "INVALID_STEP0013_EXCEL_YEAR_MONTH",
                "step0013 year or month did not match its file name or column. Path = "
                + pszInputPath + ", Column = " + str(iColumnIndex + 1),
            )
        for iRowIndex in range(2, len(arrRows)):
            pszItemName = STEP0013_EXCEL_ROW_STRUCTURE[iRowIndex][1]
            try:
                objValue = parse_decimal_value(
                    arrRows[iRowIndex][iColumnIndex],
                    iRowIndex + 1,
                    iColumnIndex,
                    pszItemName,
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "INVALID_STEP0013_EXCEL_ITEM_VALUE",
                    "invalid step0013 item value. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1)
                    + ", Month = " + str(iMonth) + ", Detail = " + str(objException),
                )
            if objValue is None and pszItemName not in {"平均体重", "育成率"}:
                raise_conversion_error(
                    "INVALID_STEP0013_EXCEL_ITEM_VALUE",
                    "step0013 item value was blank. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1)
                    + ", Month = " + str(iMonth),
                )
            if (
                objValue is not None
                and pszItemName in {"入雛羽数", "処理羽数"}
                and objValue != objValue.to_integral_value()
            ):
                raise_conversion_error(
                    "INVALID_STEP0013_EXCEL_COUNT_VALUE",
                    "step0013 counts must be integers. Path = "
                    + pszInputPath + ", Line = " + str(iRowIndex + 1)
                    + ", Month = " + str(iMonth),
                )
    return arrRows


def create_step0013_group_deduction_excel_output_path(
    pszStep0013TsvPath: str,
) -> str:
    pszDirectory = os.path.dirname(pszStep0013TsvPath)
    pszStem = os.path.splitext(os.path.basename(pszStep0013TsvPath))[0]
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def write_group_deduction_excel(
    pszOutputPath: str, arrRows: List[List[str]], iStepNumber: int
) -> None:
    pszTemporaryPath = ""
    bOutputReserved = False
    objWorkbook: Optional[Workbook] = None
    try:
        try:
            iFileDescriptor = os.open(
                pszOutputPath, os.O_CREAT | os.O_EXCL | os.O_WRONLY
            )
            os.close(iFileDescriptor)
            bOutputReserved = True
        except FileExistsError:
            raise_conversion_error(
                f"STEP{iStepNumber:04d}_EXCEL_CREATE_FAILED",
                f"step{iStepNumber:04d} Excel already exists. Path = " + pszOutputPath,
            )
        except OSError as objException:
            raise_conversion_error(
                f"STEP{iStepNumber:04d}_EXCEL_CREATE_FAILED",
                f"failed to reserve step{iStepNumber:04d} Excel. Detail = " + str(objException),
            )
        try:
            iTemporaryFileDescriptor, pszTemporaryPath = tempfile.mkstemp(
                prefix=f"group_shipment_summary_step{iStepNumber:04d}_",
                suffix=".xlsx",
                dir=os.path.dirname(pszOutputPath),
            )
            os.close(iTemporaryFileDescriptor)
            objWorkbook = Workbook()
            objWorksheet = objWorkbook.active
            objWorksheet.title = "グループ別"
            for iRowIndex, arrFields in enumerate(arrRows, start=1):
                pszItemName = arrFields[1].strip(" \u3000\t")
                for iColumnIndex, pszValue in enumerate(arrFields, start=1):
                    objCell = objWorksheet.cell(iRowIndex, iColumnIndex)
                    if iColumnIndex <= 2:
                        objCell.value = pszValue if pszValue != "" else None
                        continue
                    if iRowIndex == 2 and iColumnIndex == 15 and pszValue == "合計":
                        objCell.value = pszValue
                        continue
                    pszNormalizedValue = pszValue.strip(" \u3000\t")
                    if pszNormalizedValue == "":
                        continue
                    objValue = Decimal(pszNormalizedValue)
                    objCell.value = (
                        int(objValue)
                        if objValue == objValue.to_integral_value()
                        else float(objValue)
                    )
                    if pszItemName in {"年", "月", "入雛羽数", "出荷羽数", "処理羽数", "稼動日数"}:
                        objCell.number_format = "0"
                    elif pszItemName in {"正味重量", "処理重量"}:
                        iDecimalPlaces = max(0, -objValue.as_tuple().exponent)
                        objCell.number_format = (
                            "0." + "#" * iDecimalPlaces if iDecimalPlaces else "0"
                        )
                    elif pszItemName == "平均体重":
                        iDecimalPlaces = max(3, -objValue.as_tuple().exponent)
                        objCell.number_format = "0." + "#" * iDecimalPlaces
                    elif pszItemName == "育成率":
                        iDecimalPlaces = max(4, -objValue.as_tuple().exponent)
                        objCell.number_format = "0." + "#" * iDecimalPlaces
                    elif pszItemName == "前年比":
                        objCell.number_format = "0.0000"
            for iColumnIndex in range(1, objWorksheet.max_column + 1):
                iMaximumDisplayWidth = 0
                for objWidthCell in objWorksheet.iter_cols(
                    min_col=iColumnIndex,
                    max_col=iColumnIndex,
                    min_row=1,
                    max_row=objWorksheet.max_row,
                ):
                    for objCell in objWidthCell:
                        objValue = objCell.value
                        if objValue is None:
                            continue
                        if isinstance(objValue, (int, float, Decimal)):
                            objDecimalValue = Decimal(str(objValue))
                            pszFractionFormat = objCell.number_format.split(";")[0].partition(".")[2]
                            iFractionPlaces = sum(
                                1 for pszCharacter in pszFractionFormat
                                if pszCharacter in "0#"
                            )
                            pszDisplayValue = format(
                                objDecimalValue, "." + str(iFractionPlaces) + "f"
                            )
                            if iFractionPlaces:
                                iMinimumFractionPlaces = pszFractionFormat.count("0")
                                while (
                                    len(pszDisplayValue.rsplit(".", 1)[-1])
                                    > iMinimumFractionPlaces
                                    and pszDisplayValue.endswith("0")
                                ):
                                    pszDisplayValue = pszDisplayValue[:-1]
                                if pszDisplayValue.endswith("."):
                                    pszDisplayValue = pszDisplayValue[:-1]
                        else:
                            pszDisplayValue = str(objValue)
                        iDisplayWidth = sum(
                            2 if unicodedata.east_asian_width(pszCharacter)
                            in {"F", "W"}
                            else 1
                            for pszCharacter in pszDisplayValue
                        )
                        iMaximumDisplayWidth = max(
                            iMaximumDisplayWidth, iDisplayWidth
                        )
                objWorksheet.column_dimensions[
                    get_column_letter(iColumnIndex)
                ].width = min(max(iMaximumDisplayWidth + 2, 8), 255)
            objWorkbook.save(pszTemporaryPath)
            objWorkbook.close()
            objWorkbook = None
        except Exception as objException:
            raise_conversion_error(
                f"STEP{iStepNumber:04d}_EXCEL_WRITE_FAILED",
                f"failed to write step{iStepNumber:04d} Excel. Detail = " + str(objException),
            )
        try:
            os.replace(pszTemporaryPath, pszOutputPath)
            pszTemporaryPath = ""
            bOutputReserved = False
        except Exception as objException:
            raise_conversion_error(
                f"STEP{iStepNumber:04d}_EXCEL_FINALIZE_FAILED",
                f"failed to finalize step{iStepNumber:04d} Excel. Detail = " + str(objException),
            )
    finally:
        if objWorkbook is not None:
            objWorkbook.close()
        if pszTemporaryPath and os.path.exists(pszTemporaryPath):
            os.remove(pszTemporaryPath)
        if bOutputReserved and os.path.exists(pszOutputPath):
            os.remove(pszOutputPath)


def create_step0013_group_deduction_excel_files(
    arrStep0013TsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0013TsvPath in arrStep0013TsvPaths:
            arrRows = read_step0013_for_excel(pszStep0013TsvPath)
            pszOutputPath = create_step0013_group_deduction_excel_output_path(
                pszStep0013TsvPath
            )
            write_group_deduction_excel(pszOutputPath, arrRows, 13)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0013_EXCEL_CLEANUP_FAILED",
                "step0013 Excel creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def read_step0014_for_excel(pszInputPath: str) -> List[List[str]]:
    objMatch = re.search(
        r"_step0014_グループ別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0014_EXCEL_FILENAME",
            "unexpected step0014 group deduction file name. Path = " + pszInputPath,
        )
    iExpectedYear = int(objMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8", newline="") as objFile:
            arrRows = [strip_line_ending(pszLine).split("\t") for pszLine in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0014_EXCEL_READ_FAILED",
            "failed to read step0014 TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    if len(arrRows) != len(STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE):
        raise_conversion_error(
            "INVALID_STEP0014_EXCEL_ROW_COUNT",
            "step0014 TSV must contain 22 rows. Path = " + pszInputPath,
        )
    if any(len(arrRow) != 15 for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0014_EXCEL_COLUMN_COUNT",
            "step0014 TSV must contain 15 columns. Path = " + pszInputPath,
        )
    for iRowIndex, objExpected in enumerate(STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE):
        objActual = tuple(
            pszValue.strip(" \u3000\t") for pszValue in arrRows[iRowIndex][:2]
        )
        if objActual != objExpected:
            raise_conversion_error(
                "INVALID_STEP0014_EXCEL_ROW_STRUCTURE",
                "unexpected step0014 row structure. Path = "
                + pszInputPath
                + ", Line = "
                + str(iRowIndex + 1),
            )
    if arrRows[0][14] != "" or arrRows[1][14] != "合計":
        raise_conversion_error(
            "INVALID_STEP0014_EXCEL_ANNUAL_HEADER",
            "O1は空欄、O2は「合計」である必要があります。"
            "旧形式のTSVはstep0014から再作成してください。Path = "
            + pszInputPath,
        )
    for iMonthIndex in range(12):
        iColumnIndex = iMonthIndex + 2
        iYear, iMonth = parse_year_month(
            arrRows[0][iColumnIndex], arrRows[1][iColumnIndex], iColumnIndex + 1
        )
        if iYear != iExpectedYear or iMonth != iMonthIndex + 1:
            raise_conversion_error(
                "INVALID_STEP0014_EXCEL_YEAR_MONTH",
                "step0014 year or month did not match its file name or column. Path = "
                + pszInputPath
                + ", Column = "
                + str(iColumnIndex + 1),
            )
    for iRowIndex in range(2, len(arrRows)):
        pszItemName = STEP0014_GROUP_DEDUCTION_ROW_STRUCTURE[iRowIndex][1]
        for iColumnIndex in list(range(2, 14)) + [14]:
            objValue = parse_decimal_value(
                arrRows[iRowIndex][iColumnIndex],
                iRowIndex + 1,
                iColumnIndex,
                pszItemName,
            )
            bAnnualValue = iColumnIndex == 14
            if objValue is None and (
                bAnnualValue or pszItemName not in {"平均体重", "育成率"}
            ):
                raise_conversion_error(
                    "INVALID_STEP0014_EXCEL_ITEM_VALUE",
                    "step0014 required value was blank. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iRowIndex + 1),
                )
            if (
                objValue is not None
                and pszItemName in {"入雛羽数", "処理羽数"}
                and objValue != objValue.to_integral_value()
            ):
                raise_conversion_error(
                    "INVALID_STEP0014_EXCEL_COUNT_VALUE",
                    "step0014 counts must be integers. Path = "
                    + pszInputPath
                    + ", Line = "
                    + str(iRowIndex + 1),
                )
    return arrRows


def create_step0014_group_deduction_excel_output_path(
    pszStep0014TsvPath: str,
) -> str:
    pszDirectory = os.path.dirname(pszStep0014TsvPath)
    pszStem = os.path.splitext(os.path.basename(pszStep0014TsvPath))[0]
    pszOutputPath = os.path.join(pszDirectory, pszStem + ".xlsx")
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}.xlsx"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0014_group_deduction_excel_files(
    arrStep0014TsvPaths: List[str],
) -> List[str]:
    arrCreatedPaths: List[str] = []
    try:
        for pszStep0014TsvPath in arrStep0014TsvPaths:
            arrRows = read_step0014_for_excel(pszStep0014TsvPath)
            pszOutputPath = create_step0014_group_deduction_excel_output_path(
                pszStep0014TsvPath
            )
            write_group_deduction_excel(pszOutputPath, arrRows, 14)
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0014_EXCEL_CLEANUP_FAILED",
                "step0014 Excel creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


STEP0017_GROUP_DEDUCTION_EXCEL_STRUCTURE: List[Tuple[str, str]] = [
    ("", "年"), ("", "月")
]
for _, pszGroupName in GROUPS:
    STEP0017_GROUP_DEDUCTION_EXCEL_STRUCTURE.extend(
        [(pszGroupName, "入雛羽数"), ("", "処理羽数"), ("", "処理重量"),
         ("", "平均体重"), ("", "育成率")]
    )
STEP0017_GROUP_DEDUCTION_EXCEL_STRUCTURE.extend(
    [("合計", "入雛羽数"), ("", "処理羽数"), ("", "処理重量"),
     ("", "平均体重"), ("", "育成率"), ("", "稼動日数")]
)


def read_step0017_group_deduction_for_excel(
    pszInputPath: str,
) -> List[List[str]]:
    objMatch = re.search(
        r"_step0017_グループ別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv$",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0017_EXCEL_FILENAME",
            "unexpected step0017 group deduction file name. Path = " + pszInputPath,
        )
    iYear = int(objMatch.group(1))
    try:
        with open(pszInputPath, "r", encoding="utf-8-sig", newline="") as objFile:
            arrRows = [strip_line_ending(line).split("\t") for line in objFile]
    except Exception as objException:
        raise_conversion_error(
            "STEP0017_EXCEL_READ_FAILED",
            "failed to read step0017 TSV. Path = " + pszInputPath
            + ", Detail = " + str(objException),
        )
    bHasPrevious = iYear != 2022
    iExpectedRows = 31 if bHasPrevious else 23
    if len(arrRows) != iExpectedRows or any(len(row) != 15 for row in arrRows):
        raise_conversion_error(
            "INVALID_STEP0017_EXCEL_DIMENSIONS",
            "step0017 TSV must contain " + str(iExpectedRows)
            + " rows and 15 columns. Path = " + pszInputPath,
        )
    for i, objExpected in enumerate(STEP0017_GROUP_DEDUCTION_EXCEL_STRUCTURE):
        if tuple(value.strip(" \u3000\t") for value in arrRows[i][:2]) != objExpected:
            raise_conversion_error(
                "INVALID_STEP0017_EXCEL_ROW_STRUCTURE",
                "unexpected step0017 row structure. Path = " + pszInputPath
                + ", Line = " + str(i + 1),
            )
    if arrRows[0][14] != "" or arrRows[1][14] != "合計":
        raise_conversion_error(
            "INVALID_STEP0017_EXCEL_ANNUAL_HEADER",
            "O1は空欄、O2は「合計」である必要があります。"
            "旧形式のTSVはstep0014から再作成してください。Path = " + pszInputPath,
        )
    for month_index in range(12):
        col = month_index + 2
        year, month = parse_year_month(arrRows[0][col], arrRows[1][col], col + 1)
        if year != iYear or month != month_index + 1:
            raise_conversion_error(
                "INVALID_STEP0017_EXCEL_YEAR_MONTH",
                "step0017 year/month header is invalid. Path = " + pszInputPath,
            )

    def validate_value(row_index: int, col: int, item: str, allow_blank: bool) -> None:
        try:
            value = parse_decimal_value(arrRows[row_index][col], row_index + 1, col, item)
        except ConversionError as exc:
            raise_conversion_error("INVALID_STEP0017_EXCEL_VALUE", str(exc))
        if value is None and not allow_blank:
            raise_conversion_error(
                "INVALID_STEP0017_EXCEL_VALUE",
                "step0017 required value was blank. Path = " + pszInputPath
                + ", Line = " + str(row_index + 1) + ", Column = " + str(col + 1),
            )
        if value is not None and item in {"入雛羽数", "処理羽数", "稼動日数"}:
            if value != value.to_integral_value():
                raise_conversion_error(
                    "INVALID_STEP0017_EXCEL_COUNT",
                    "step0017 count must be an integer. Path = " + pszInputPath,
                )

    for row_index in range(2, 23):
        item = STEP0017_GROUP_DEDUCTION_EXCEL_STRUCTURE[row_index][1]
        for col in range(2, 15):
            validate_value(
                row_index, col, item,
                item in {"平均体重", "育成率", "稼動日数"} and col < 14,
            )
    if bHasPrevious:
        if (arrRows[23][:2] != ["処理羽数", "前年比"]
                or arrRows[24][:2] != ["", "(一日平均)"]
                or any(arrRows[24][2:])):
            raise_conversion_error(
                "INVALID_STEP0017_EXCEL_RATIO_ROW",
                "step0017 ratio rows are invalid. Path = " + pszInputPath,
            )
        for col in range(2, 15):
            validate_value(23, col, "前年比", True)
        previous = [("前年実績", "入雛羽数"), (f"{iYear - 1}年", "処理羽数"),
                    ("", "処理重量"), ("", "平均体重"), ("", "育成率"),
                    ("", "稼動日数")]
        for offset, expected in enumerate(previous):
            row_index = 25 + offset
            if tuple(v.strip(" \u3000\t") for v in arrRows[row_index][:2]) != expected:
                raise_conversion_error(
                    "INVALID_STEP0017_EXCEL_PREVIOUS_STRUCTURE",
                    "step0017 previous-year rows are invalid. Path = " + pszInputPath,
                )
            for col in range(2, 15):
                validate_value(
                    row_index, col, expected[1],
                    expected[1] in {"平均体重", "育成率", "稼動日数"} and col < 14,
                )
    return arrRows


def create_step0017_group_deduction_excel_files(
    arrTsvPaths: List[str],
) -> List[str]:
    arrCreated: List[str] = []
    try:
        for tsv_path in arrTsvPaths:
            rows = read_step0017_group_deduction_for_excel(tsv_path)
            directory = os.path.dirname(tsv_path)
            stem = os.path.splitext(os.path.basename(tsv_path))[0]
            output = os.path.join(directory, stem + ".xlsx")
            suffix = 2
            while os.path.exists(output):
                output = os.path.join(directory, f"{stem}_{suffix:02d}.xlsx")
                suffix += 1
            write_group_deduction_excel(output, rows, 17)
            arrCreated.append(output)
        return arrCreated
    except Exception:
        for path in arrCreated:
            try:
                os.remove(path)
            except OSError:
                pass
        raise


def parse_step0014_group_deduction_file_name(
    pszInputPath: str,
) -> Tuple[str, int]:
    objMatch = re.fullmatch(
        r"(.+)_step0014_グループ別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0014_GROUP_DEDUCTION_FILENAME",
            "unexpected step0014 group deduction file name. Path = " + pszInputPath,
        )
    return objMatch.group(1), int(objMatch.group(2))


def parse_daily_operation_file_name(pszInputPath: str) -> Tuple[str, int]:
    objMatch = re.fullmatch(
        r"(.+)_step0001_日別稼動判定_(\d{4})(?:_\d{2})?\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_DAILY_OPERATION_FILENAME",
            "unexpected daily operation file name. Path = " + pszInputPath,
        )
    return objMatch.group(1), int(objMatch.group(2))


def read_daily_operation_row_for_step0015(
    pszInputPath: str, pszExpectedPrefix: str, iExpectedYear: int
) -> List[str]:
    pszActualPrefix, iActualYear = parse_daily_operation_file_name(pszInputPath)
    if (pszActualPrefix, iActualYear) != (pszExpectedPrefix, iExpectedYear):
        raise_conversion_error(
            "STEP0015_DAILY_OPERATION_KEY_MISMATCH",
            "daily operation file did not match step0014 base name and year. Path = "
            + pszInputPath,
        )
    try:
        with open(pszInputPath, "r", encoding="utf-8-sig", newline="") as objFile:
            arrRows = [
                strip_line_ending(pszLine).split("\t") for pszLine in objFile
            ]
    except Exception as objException:
        raise_conversion_error(
            "STEP0015_DAILY_OPERATION_READ_FAILED",
            "failed to read daily operation TSV. Path = "
            + pszInputPath
            + ", Detail = "
            + str(objException),
        )
    if len(arrRows) != 33 or any(len(arrRow) != 13 for arrRow in arrRows):
        raise_conversion_error(
            "INVALID_STEP0015_DAILY_OPERATION_STRUCTURE",
            "daily operation TSV must contain 33 rows and 13 columns. Path = "
            + pszInputPath,
        )
    arrExpectedMonths = [f"{iMonth}月" for iMonth in range(1, 13)]
    if (
        arrRows[0][0].strip(" \u3000\t") != f"{iExpectedYear}年"
        or [pszValue.strip(" \u3000\t") for pszValue in arrRows[0][1:]]
        != arrExpectedMonths
    ):
        raise_conversion_error(
            "INVALID_STEP0015_DAILY_OPERATION_HEADER",
            "daily operation year or month header did not match the expected year. Path = "
            + pszInputPath,
        )
    if arrRows[1][0].strip(" \u3000\t") != "稼動日数":
        raise_conversion_error(
            "INVALID_STEP0015_DAILY_OPERATION_ROW",
            "daily operation TSV did not contain the 稼動日数 row in row 2. Path = "
            + pszInputPath,
        )
    arrDayCounts: List[str] = []
    for iMonth, pszValue in enumerate(arrRows[1][1:], start=1):
        pszNormalizedValue = pszValue.strip(" \u3000\t")
        if pszNormalizedValue:
            if not pszNormalizedValue.isdigit() or int(pszNormalizedValue) > 31:
                raise_conversion_error(
                    "INVALID_STEP0015_DAILY_OPERATION_VALUE",
                    "invalid operating day count. Path = "
                    + pszInputPath
                    + ", Month = "
                    + str(iMonth),
                )
        arrDayCounts.append(pszNormalizedValue)
    return arrDayCounts


def create_step0015_group_deduction_output_path(
    pszStep0014Path: str,
) -> str:
    pszDirectory = os.path.dirname(pszStep0014Path)
    pszFileName = os.path.basename(pszStep0014Path)
    if re.search(r"_step0014_", pszFileName) is None:
        raise_conversion_error(
            "INVALID_STEP0014_GROUP_DEDUCTION_FILENAME",
            "step0014 marker was not found. Path = " + pszStep0014Path,
        )
    pszOutputName = re.sub(
        r"_step0014_", "_step0015_", pszFileName, count=1
    )
    pszStem, pszExtension = os.path.splitext(pszOutputName)
    pszOutputPath = os.path.join(pszDirectory, pszOutputName)
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0015_group_deduction_files(
    arrStep0014Paths: List[str], arrDailyOperationPaths: List[str]
) -> List[str]:
    dicDailyPaths: Dict[Tuple[str, str, int], str] = {}
    for pszDailyPath in arrDailyOperationPaths:
        pszPrefix, iYear = parse_daily_operation_file_name(pszDailyPath)
        objKey = (
            os.path.dirname(os.path.abspath(pszDailyPath)),
            pszPrefix,
            iYear,
        )
        if objKey in dicDailyPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0015_DAILY_OPERATION_INPUT",
                "duplicate daily operation input. Path = " + pszDailyPath,
            )
        dicDailyPaths[objKey] = pszDailyPath

    arrCreatedPaths: List[str] = []
    try:
        for pszStep0014Path in arrStep0014Paths:
            pszPrefix, iYear = parse_step0014_group_deduction_file_name(
                pszStep0014Path
            )
            objKey = (
                os.path.dirname(os.path.abspath(pszStep0014Path)),
                pszPrefix,
                iYear,
            )
            pszDailyPath = dicDailyPaths.get(objKey)
            if pszDailyPath is None:
                raise_conversion_error(
                    "MISSING_STEP0015_DAILY_OPERATION_INPUT",
                    "matching daily operation TSV was not found for step0014. Path = "
                    + pszStep0014Path,
                )
            arrRows = read_step0014_for_excel(pszStep0014Path)
            arrDayCounts = read_daily_operation_row_for_step0015(
                pszDailyPath, pszPrefix, iYear
            )
            arrRows.append(["", "稼動日数"] + arrDayCounts + [""])
            pszOutputPath = create_step0015_group_deduction_output_path(
                pszStep0014Path
            )
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrRows[0]),
                    ["\t".join(arrRow) for arrRow in arrRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0015_GROUP_DEDUCTION_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0015_GROUP_DEDUCTION_CLEANUP_FAILED",
                "step0015 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def parse_step0010_group_file_name(pszInputPath: str) -> Tuple[str, int]:
    objMatch = re.fullmatch(
        r"(.+)_step0010_グループ別_(\d{4})_年間(?:_\d{2})?\.tsv",
        os.path.basename(pszInputPath),
    )
    if objMatch is None:
        raise_conversion_error(
            "INVALID_STEP0010_GROUP_FILENAME",
            "unexpected step0010 group file name. Path = " + pszInputPath,
        )
    return objMatch.group(1), int(objMatch.group(2))


def create_step0011_group_output_path(pszStep0010Path: str) -> str:
    pszDirectory = os.path.dirname(pszStep0010Path)
    pszOutputName = re.sub(
        r"_step0010_", "_step0011_", os.path.basename(pszStep0010Path), count=1
    )
    pszStem, pszExtension = os.path.splitext(pszOutputName)
    pszOutputPath = os.path.join(pszDirectory, pszOutputName)
    iSuffix = 2
    while os.path.exists(pszOutputPath):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszStem}_{iSuffix:02d}{pszExtension}"
        )
        iSuffix += 1
    return pszOutputPath


def create_step0011_group_files(
    arrStep0010Paths: List[str], arrDailyOperationPaths: List[str]
) -> List[str]:
    """step0010グループ別年間TSVに稼動日数行を加え、step0011を作成します。"""
    dicDailyPaths: Dict[Tuple[str, str, int], str] = {}
    for pszDailyPath in arrDailyOperationPaths:
        pszPrefix, iYear = parse_daily_operation_file_name(pszDailyPath)
        objKey = (os.path.dirname(os.path.abspath(pszDailyPath)), pszPrefix, iYear)
        if objKey in dicDailyPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0011_DAILY_OPERATION_INPUT",
                "duplicate daily operation input. Path = " + pszDailyPath,
            )
        dicDailyPaths[objKey] = pszDailyPath

    arrCreatedPaths: List[str] = []
    try:
        for pszStep0010Path in arrStep0010Paths:
            pszPrefix, iYear = parse_step0010_group_file_name(pszStep0010Path)
            objKey = (
                os.path.dirname(os.path.abspath(pszStep0010Path)), pszPrefix, iYear
            )
            pszDailyPath = dicDailyPaths.get(objKey)
            if pszDailyPath is None:
                raise_conversion_error(
                    "MISSING_STEP0011_DAILY_OPERATION_INPUT",
                    "matching daily operation TSV was not found for step0010. Path = "
                    + pszStep0010Path,
                )
            arrRows = read_step0010_for_excel(pszStep0010Path)
            arrDayCounts = read_daily_operation_row_for_step0015(
                pszDailyPath, pszPrefix, iYear
            )
            arrRows.append(["", "稼動日数"] + arrDayCounts + [""])
            pszOutputPath = create_step0011_group_output_path(pszStep0010Path)
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrRows[0]),
                    ["\t".join(arrRow) for arrRow in arrRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error(
                    "STEP0011_GROUP_WRITE_FAILED", str(objException)
                )
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0011_GROUP_CLEANUP_FAILED",
                "step0011 creation failed. Original error = "
                + str(objException)
                + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0011_variety_files(
    arrStep0010Paths: List[str], arrDailyOperationPaths: List[str]
) -> List[str]:
    """品種別step0010年間TSVに全品種共通の稼動日数行を追加します。"""
    dicDailyPaths: Dict[Tuple[str, str, int], str] = {}
    for pszDailyPath in arrDailyOperationPaths:
        pszPrefix, iYear = parse_daily_operation_file_name(pszDailyPath)
        objKey = (os.path.dirname(os.path.abspath(pszDailyPath)), pszPrefix, iYear)
        if objKey in dicDailyPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0011_VARIETY_DAILY_INPUT",
                "duplicate daily operation input. Path = " + pszDailyPath,
            )
        dicDailyPaths[objKey] = pszDailyPath

    arrCreatedPaths: List[str] = []
    try:
        for pszStep0010Path in arrStep0010Paths:
            objMatch = re.fullmatch(
                r"(.+)_step0010_品種別_(\d{4})_年間(?:_\d{2})?\.tsv",
                os.path.basename(pszStep0010Path),
            )
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_STEP0010_VARIETY_FILENAME",
                    "unexpected variety step0010 file name. Path = "
                    + pszStep0010Path,
                )
            pszPrefix, iYearText = objMatch.groups()
            iYear = int(iYearText)
            objKey = (
                os.path.dirname(os.path.abspath(pszStep0010Path)), pszPrefix, iYear
            )
            pszDailyPath = dicDailyPaths.get(objKey)
            if pszDailyPath is None:
                raise_conversion_error(
                    "MISSING_STEP0011_VARIETY_DAILY_INPUT",
                    "matching daily operation TSV was not found. Path = "
                    + pszStep0010Path,
                )
            arrRows = read_step0010_variety_for_excel(pszStep0010Path)
            arrDayCounts = read_daily_operation_row_for_step0015(
                pszDailyPath, pszPrefix, iYear
            )
            arrRows.append(["", "稼動日数"] + arrDayCounts + [""])
            pszOutputPath = os.path.join(
                os.path.dirname(pszStep0010Path),
                os.path.basename(pszStep0010Path).replace(
                    "_step0010_", "_step0011_", 1
                ),
            )
            pszStem, pszExtension = os.path.splitext(pszOutputPath)
            iSuffix = 2
            while os.path.exists(pszOutputPath):
                pszOutputPath = f"{pszStem}_{iSuffix:02d}{pszExtension}"
                iSuffix += 1
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrRows[0]),
                    ["\t".join(row) for row in arrRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error("STEP0011_VARIETY_WRITE_FAILED", str(objException))
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0011_VARIETY_CLEANUP_FAILED",
                "step0011 variety creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0011_variety_deduction_files(
    arrStep0010Paths: List[str], arrDailyOperationPaths: List[str]
) -> List[str]:
    """控除廃棄版品種別step0010年間TSVに全品種共通の稼動日数行を追加します。"""
    dicDailyPaths: Dict[Tuple[str, str, int], str] = {}
    for pszDailyPath in arrDailyOperationPaths:
        pszPrefix, iYear = parse_daily_operation_file_name(pszDailyPath)
        objKey = (os.path.dirname(os.path.abspath(pszDailyPath)), pszPrefix, iYear)
        if objKey in dicDailyPaths:
            raise_conversion_error(
                "DUPLICATE_STEP0011_VARIETY_DEDUCTION_DAILY_INPUT",
                "duplicate daily operation input. Path = " + pszDailyPath,
            )
        dicDailyPaths[objKey] = pszDailyPath

    arrCreatedPaths: List[str] = []
    try:
        for pszStep0010Path in arrStep0010Paths:
            objMatch = re.fullmatch(
                r"(.+)_step0010_品種別_控除廃棄_(\d{4})_年間(?:_\d{2})?\.tsv",
                os.path.basename(pszStep0010Path),
            )
            if objMatch is None:
                raise_conversion_error(
                    "INVALID_STEP0010_VARIETY_DEDUCTION_FILENAME",
                    "unexpected variety deduction step0010 file name. Path = "
                    + pszStep0010Path,
                )
            pszPrefix, iYearText = objMatch.groups()
            iYear = int(iYearText)
            objKey = (
                os.path.dirname(os.path.abspath(pszStep0010Path)), pszPrefix, iYear
            )
            pszDailyPath = dicDailyPaths.get(objKey)
            if pszDailyPath is None:
                raise_conversion_error(
                    "MISSING_STEP0011_VARIETY_DEDUCTION_DAILY_INPUT",
                    "matching daily operation TSV was not found. Path = "
                    + pszStep0010Path,
                )
            arrRows = read_step0010_variety_deduction_for_excel(pszStep0010Path)
            arrDayCounts = read_daily_operation_row_for_step0015(
                pszDailyPath, pszPrefix, iYear
            )
            arrRows.append(["", "稼動日数"] + arrDayCounts + [""])
            pszOutputPath = os.path.join(
                os.path.dirname(pszStep0010Path),
                os.path.basename(pszStep0010Path).replace(
                    "_step0010_", "_step0011_", 1
                ),
            )
            pszStem, pszExtension = os.path.splitext(pszOutputPath)
            iSuffix = 2
            while os.path.exists(pszOutputPath):
                pszOutputPath = f"{pszStem}_{iSuffix:02d}{pszExtension}"
                iSuffix += 1
            try:
                write_monthly_tsv(
                    pszOutputPath,
                    "\t".join(arrRows[0]),
                    ["\t".join(row) for row in arrRows[1:]],
                )
            except ConversionError as objException:
                raise_conversion_error("STEP0011_VARIETY_DEDUCTION_WRITE_FAILED", str(objException))
            arrCreatedPaths.append(pszOutputPath)
        return arrCreatedPaths
    except Exception as objException:
        arrCleanupErrors: List[str] = []
        for pszPath in arrCreatedPaths:
            try:
                os.remove(pszPath)
            except Exception as objCleanupException:
                arrCleanupErrors.append(pszPath + ": " + str(objCleanupException))
        if arrCleanupErrors:
            raise_conversion_error(
                "STEP0011_VARIETY_DEDUCTION_CLEANUP_FAILED",
                "step0011 variety deduction creation failed. Original error = "
                + str(objException) + "; cleanup errors = "
                + " | ".join(arrCleanupErrors),
            )
        raise


def create_step0012_variety_files(arrStep0011Paths: List[str]) -> List[str]:
    """品種別当年表に空行2行と前年合計・稼動日数7行を追加します。"""
    dicInputs: Dict[Tuple[str, str, int], str] = {}
    for pszPath in arrStep0011Paths:
        objMatch = re.fullmatch(
            r"(.+)_step0011_品種別_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error("STEP0012_VARIETY_INVALID_NAME", "Path = " + pszPath)
        pszPrefix, pszYear = objMatch.groups()
        objKey = (os.path.dirname(os.path.abspath(pszPath)), pszPrefix, int(pszYear))
        if objKey in dicInputs:
            raise_conversion_error("STEP0012_VARIETY_DUPLICATE_YEAR", "Path = " + pszPath)
        dicInputs[objKey] = pszPath

    def read_annual(pszPath: str, iYear: int) -> Tuple[bytes, List[List[str]]]:
        with open(pszPath, "rb") as objFile:
            objData = objFile.read()
        arrRows = [line.split("\t") for line in objData.decode("utf-8-sig").splitlines()]
        arrStructure = [("", "年"), ("", "月")]
        for pszName in [name for _, name in SPLIT_VARIETIES] + ["合計"]:
            arrStructure.extend(
                (pszName if i == 0 else "", item)
                for i, item in enumerate(STEP0008_VARIETY_ITEM_NAMES)
            )
        arrStructure.append(("", "稼動日数"))
        if (len(arrRows) != 39 or any(len(row) != 15 for row in arrRows)
                or [tuple(row[:2]) for row in arrRows] != arrStructure
                or arrRows[0] != ["", "年"] + [str(iYear)] * 12 + [""]
                or arrRows[1] != ["", "月"] + [str(m) for m in range(1, 13)] + ["合計"]):
            raise ValueError("品種別年間表の行数・列数・年月見出し・項目構成が不正です。Path = " + pszPath)
        for value in arrRows[38][2:14]:
            if value != "" and re.fullmatch(r"[0-9]+", value) is None:
                raise ValueError("稼動日数は空欄または非負の整数である必要があります。Path = " + pszPath)
        return objData, arrRows

    arrCreatedPaths: List[str] = []
    arrErrors: List[str] = []
    for (pszDirectory, pszPrefix, iYear), pszInputPath in sorted(dicInputs.items()):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszPrefix}_step0012_品種別_{iYear:04d}_年間.tsv"
        )
        bReserved = False
        try:
            objCurrent, arrCurrentRows = read_annual(pszInputPath, iYear)
            pszNewline = "\r\n" if b"\r\n" in objCurrent else ("\r" if b"\r" in objCurrent else "\n")
            objNewline = pszNewline.encode("ascii")
            arrLines = objCurrent.splitlines(keepends=True)
            objBody = arrLines[38].rstrip(b"\r\n")
            arrFields = objBody.split(b"\t")
            arrFields[14] = str(sum(int(v) for v in arrCurrentRows[38][2:14] if v)).encode("ascii")
            arrLines[38] = b"\t".join(arrFields) + arrLines[38][len(objBody):]
            objOutput = b"".join(arrLines)
            if iYear != 2022:
                pszPreviousPath = dicInputs.get(
                    (pszDirectory, pszPrefix, iYear - 1),
                    os.path.join(pszDirectory, f"{pszPrefix}_step0011_品種別_{iYear - 1:04d}_年間.tsv"),
                )
                _, arrPreviousRows = read_annual(pszPreviousPath, iYear - 1)
                arrTotals = [list(row) for row in arrPreviousRows[32:39]]
                arrTotals[0][0] = "前年実績"
                arrTotals[1][0] = f"{iYear - 1}年"
                arrTotals[6][14] = str(sum(int(v) for v in arrTotals[6][2:14] if v))
                objOutput = objOutput.rstrip(b"\r\n") + objNewline * 3
                objOutput += (pszNewline.join("\t".join(row) for row in arrTotals) + pszNewline).encode("utf-8")
            with open(pszOutputPath, "xb") as objFile:
                bReserved = True
                objFile.write(objOutput)
            bReserved = False
            arrCreatedPaths.append(pszOutputPath)
        except (OSError, UnicodeError, ValueError) as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszInputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("STEP0012_VARIETY_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def find_step0011_variety_deduction_files(
    arrStep0010Paths: List[str],
) -> List[str]:
    """今回の年間出力に対応する、作成済みの控除廃棄版step0011を探します。"""
    arrPaths: List[str] = []
    objSeen: set = set()
    for pszPath in arrStep0010Paths:
        objMatch = re.fullmatch(
            r"(.+)_step0010_品種別_控除廃棄_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error("STEP0012_VARIETY_DEDUCTION_INVALID_SOURCE_NAME", "Path = " + pszPath)
        pszPrefix, pszYear = objMatch.groups()
        pszDirectory = os.path.dirname(os.path.abspath(pszPath))
        objKey = (pszDirectory, pszPrefix, pszYear)
        if objKey in objSeen:
            continue
        objSeen.add(objKey)
        pszPattern = (
            re.escape(pszPrefix) + r"_step0011_品種別_控除廃棄_"
            + pszYear + r"_年間(?:_[0-9]+)?\.tsv"
        )
        for pszName in sorted(os.listdir(pszDirectory)):
            if re.fullmatch(pszPattern, pszName):
                arrPaths.append(os.path.join(pszDirectory, pszName))
    return arrPaths


def create_step0012_variety_deduction_files(arrStep0011Paths: List[str]) -> List[str]:
    """控除廃棄版品種別当年表に空行2行と前年合計・稼動日数7行を追加します。"""
    dicInputs: Dict[Tuple[str, str, int], str] = {}
    for pszPath in arrStep0011Paths:
        objMatch = re.fullmatch(
            r"(.+)_step0011_品種別_控除廃棄_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error("STEP0012_VARIETY_DEDUCTION_INVALID_NAME", "Path = " + pszPath)
        pszPrefix, pszYear = objMatch.groups()
        objKey = (os.path.dirname(os.path.abspath(pszPath)), pszPrefix, int(pszYear))
        if objKey in dicInputs:
            raise_conversion_error("STEP0012_VARIETY_DEDUCTION_DUPLICATE_YEAR", "Path = " + pszPath)
        dicInputs[objKey] = pszPath

    def read_annual(pszPath: str, iYear: int) -> Tuple[bytes, List[List[str]]]:
        with open(pszPath, "rb") as objFile:
            objData = objFile.read()
        arrRows = [line.split("\t") for line in objData.decode("utf-8-sig").splitlines()]
        arrStructure = [("", "年"), ("", "月")]
        for pszName in [name for _, name in SPLIT_VARIETIES] + ["合計"]:
            arrStructure.extend(
                (pszName if i == 0 else "", item)
                for i, item in enumerate(STEP0008_VARIETY_DEDUCTION_ITEM_NAMES)
            )
        arrStructure.append(("", "稼動日数"))
        if (len(arrRows) != 39 or any(len(row) != 15 for row in arrRows)
                or [tuple(row[:2]) for row in arrRows] != arrStructure
                or arrRows[0] != ["", "年"] + [str(iYear)] * 12 + [""]
                or arrRows[1] != ["", "月"] + [str(m) for m in range(1, 13)] + ["合計"]):
            raise ValueError("控除廃棄版品種別年間表の行数・列数・年月見出し・項目構成が不正です。Path = " + pszPath)
        for value in arrRows[38][2:14]:
            if value != "" and re.fullmatch(r"[0-9]+", value) is None:
                raise ValueError("稼動日数は空欄または非負の整数である必要があります。Path = " + pszPath)
        return objData, arrRows

    arrCreatedPaths: List[str] = []
    arrErrors: List[str] = []
    for (pszDirectory, pszPrefix, iYear), pszInputPath in sorted(dicInputs.items()):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszPrefix}_step0012_品種別_控除廃棄_{iYear:04d}_年間.tsv"
        )
        bReserved = False
        try:
            objCurrent, arrCurrentRows = read_annual(pszInputPath, iYear)
            pszNewline = "\r\n" if b"\r\n" in objCurrent else ("\r" if b"\r" in objCurrent else "\n")
            objNewline = pszNewline.encode("ascii")
            arrLines = objCurrent.splitlines(keepends=True)
            objBody = arrLines[38].rstrip(b"\r\n")
            arrFields = objBody.split(b"\t")
            arrFields[14] = str(sum(int(v) for v in arrCurrentRows[38][2:14] if v)).encode("ascii")
            arrLines[38] = b"\t".join(arrFields) + arrLines[38][len(objBody):]
            objOutput = b"".join(arrLines)
            if iYear != 2022:
                pszPreviousPath = dicInputs.get(
                    (pszDirectory, pszPrefix, iYear - 1),
                    os.path.join(pszDirectory, f"{pszPrefix}_step0011_品種別_控除廃棄_{iYear - 1:04d}_年間.tsv"),
                )
                _, arrPreviousRows = read_annual(pszPreviousPath, iYear - 1)
                arrTotals = [list(row) for row in arrPreviousRows[32:39]]
                arrTotals[0][0] = "前年実績"
                arrTotals[1][0] = f"{iYear - 1}年"
                arrTotals[6][14] = str(sum(int(v) for v in arrTotals[6][2:14] if v))
                objOutput = objOutput.rstrip(b"\r\n") + objNewline * 3
                objOutput += (pszNewline.join("\t".join(row) for row in arrTotals) + pszNewline).encode("utf-8")
            with open(pszOutputPath, "xb") as objFile:
                bReserved = True
                objFile.write(objOutput)
            bReserved = False
            arrCreatedPaths.append(pszOutputPath)
        except (OSError, UnicodeError, ValueError) as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszInputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("STEP0012_VARIETY_DEDUCTION_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def create_step0012_group_files(arrStep0011Paths: List[str]) -> List[str]:
    """当年表に前年合計6行を追加し、稼動日数の年間合計を出力します。"""
    dicInputs: Dict[Tuple[str, str, int], str] = {}
    for pszPath in arrStep0011Paths:
        objMatch = re.fullmatch(
            r"(.+)_step0011_グループ別_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error("STEP0012_GROUP_INVALID_NAME", "Path = " + pszPath)
        pszPrefix, pszYear = objMatch.groups()
        objKey = (os.path.dirname(os.path.abspath(pszPath)), pszPrefix, int(pszYear))
        if objKey in dicInputs:
            raise_conversion_error("STEP0012_GROUP_DUPLICATE_YEAR", "Path = " + pszPath)
        dicInputs[objKey] = pszPath

    arrCreatedPaths: List[str] = []
    arrErrors: List[str] = []
    for (pszDirectory, pszPrefix, iYear), pszInputPath in sorted(dicInputs.items()):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszPrefix}_step0012_グループ別_{iYear:04d}_年間.tsv"
        )
        bReserved = False
        try:
            with open(pszInputPath, "rb") as objFile:
                objCurrent = objFile.read()
            objOutput = objCurrent
            if iYear != 2022:
                pszPreviousPath = dicInputs.get(
                    (pszDirectory, pszPrefix, iYear - 1),
                    os.path.join(pszDirectory, f"{pszPrefix}_step0011_グループ別_{iYear - 1:04d}_年間.tsv"),
                )
                with open(pszPreviousPath, "rb") as objFile:
                    objPrevious = objFile.read()
                arrCurrentLines = objCurrent.decode("utf-8-sig").splitlines()
                arrPreviousLines = objPrevious.decode("utf-8-sig").splitlines()
                arrCurrentRows = [line.split("\t") for line in arrCurrentLines]
                arrPreviousRows = [line.split("\t") for line in arrPreviousLines]
                for pszPath, arrRows in (
                    (pszInputPath, arrCurrentRows), (pszPreviousPath, arrPreviousRows)
                ):
                    if (len(arrRows) < 2 or len(arrRows[0]) != 15
                            or arrRows[0][:2] != ["", "年"]
                            or arrRows[1] != ["", "月"] + [str(m) for m in range(1, 13)] + ["合計"]
                            or any(len(row) != 15 for row in arrRows)):
                        raise ValueError("年間表の列構成が不正です。Path = " + pszPath)
                arrTotalIndexes = [
                    i for i, row in enumerate(arrPreviousRows) if row[0] == "合計"
                ]
                if len(arrTotalIndexes) != 1:
                    raise ValueError("前年の合計ブロックが不足または重複しています。Path = " + pszPreviousPath)
                iStart = arrTotalIndexes[0]
                arrTotals = [list(row) for row in arrPreviousRows[iStart:iStart + 6]]
                arrItems = ["入雛羽数", "出荷羽数", "正味重量", "平均体重", "育成率", "稼動日数"]
                if (len(arrTotals) != 6 or [row[1] for row in arrTotals] != arrItems
                        or any(row[0] for row in arrTotals[1:])
                        or iStart + 6 != len(arrPreviousRows)):
                    raise ValueError("前年の合計6行が不正です。Path = " + pszPreviousPath)
                arrTotals[0][0] = "前年実績"
                arrTotals[1][0] = f"{iYear - 1}年"
                pszNewline = "\r\n" if b"\r\n" in objCurrent else ("\r" if b"\r" in objCurrent else "\n")
                objNewline = pszNewline.encode("ascii")
                # 元の内容を保持し、最終行との間に空行をちょうど2行設けます。
                objOutput = objCurrent.rstrip(b"\r\n") + objNewline * 3
                objOutput += (pszNewline.join("\t".join(row) for row in arrTotals) + pszNewline).encode("utf-8")
            # 稼動日数行のO列だけを書き換え、他のバイト列・改行は保持します。
            arrOutputLines = objOutput.splitlines(keepends=True)
            for iLineIndex, objLine in enumerate(arrOutputLines):
                objBody = objLine.rstrip(b"\r\n")
                arrFields = objBody.split(b"\t")
                if len(arrFields) >= 2 and arrFields[1] == "稼動日数".encode("utf-8"):
                    if len(arrFields) != 15:
                        raise ValueError("稼動日数行は15列である必要があります。")
                    iTotalDays = 0
                    for objValue in arrFields[2:14]:
                        if objValue == b"":
                            continue
                        if re.fullmatch(rb"[0-9]+", objValue) is None:
                            raise ValueError("稼動日数は空欄または非負の整数である必要があります。")
                        iTotalDays += int(objValue)
                    arrFields[14] = str(iTotalDays).encode("ascii")
                    arrOutputLines[iLineIndex] = b"\t".join(arrFields) + objLine[len(objBody):]
            objOutput = b"".join(arrOutputLines)
            with open(pszOutputPath, "xb") as objFile:
                bReserved = True
                objFile.write(objOutput)
            bReserved = False
            arrCreatedPaths.append(pszOutputPath)
        except (OSError, UnicodeError, ValueError) as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszInputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("STEP0012_GROUP_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def create_step0016_group_deduction_files(arrStep0015Paths: List[str]) -> List[str]:
    """当年表に前年合計6行を追加し、稼動日数の年間合計を出力します。"""
    dicInputs: Dict[Tuple[str, str, int], str] = {}
    for pszPath in arrStep0015Paths:
        objMatch = re.fullmatch(
            r"(.+)_step0015_グループ別_控除廃棄_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
            os.path.basename(pszPath),
        )
        if objMatch is None:
            raise_conversion_error("STEP0016_GROUP_DEDUCTION_INVALID_NAME", "Path = " + pszPath)
        pszPrefix, pszYear = objMatch.groups()
        objKey = (os.path.dirname(os.path.abspath(pszPath)), pszPrefix, int(pszYear))
        if objKey in dicInputs:
            raise_conversion_error("STEP0016_GROUP_DEDUCTION_DUPLICATE_YEAR", "Path = " + pszPath)
        dicInputs[objKey] = pszPath

    arrCreatedPaths: List[str] = []
    arrErrors: List[str] = []
    for (pszDirectory, pszPrefix, iYear), pszInputPath in sorted(dicInputs.items()):
        pszOutputPath = os.path.join(
            pszDirectory, f"{pszPrefix}_step0016_グループ別_控除廃棄_{iYear:04d}_年間.tsv"
        )
        bReserved = False
        try:
            with open(pszInputPath, "rb") as objFile:
                objCurrent = objFile.read()
            objOutput = objCurrent
            arrCurrentRows = [
                line.split("\t") for line in objCurrent.decode("utf-8-sig").splitlines()
            ]
            if (len(arrCurrentRows) < 2
                    or any(len(row) != 15 for row in arrCurrentRows[:2])
                    or arrCurrentRows[0][14] != "" or arrCurrentRows[1][14] != "合計"):
                raise ValueError(
                    "O1は空欄、O2は「合計」である必要があります。"
                    "旧形式のTSVはstep0014から再作成してください。Path = " + pszInputPath
                )
            if iYear != 2022:
                pszPreviousPath = dicInputs.get(
                    (pszDirectory, pszPrefix, iYear - 1),
                    os.path.join(pszDirectory, f"{pszPrefix}_step0015_グループ別_控除廃棄_{iYear - 1:04d}_年間.tsv"),
                )
                with open(pszPreviousPath, "rb") as objFile:
                    objPrevious = objFile.read()
                arrPreviousLines = objPrevious.decode("utf-8-sig").splitlines()
                arrPreviousRows = [line.split("\t") for line in arrPreviousLines]
                for pszPath, arrRows, iExpectedYear in (
                    (pszInputPath, arrCurrentRows, iYear),
                    (pszPreviousPath, arrPreviousRows, iYear - 1),
                ):
                    if (len(arrRows) < 2
                            or any(len(row) != 15 for row in arrRows)
                            or arrRows[0] != ["", "年"] + [str(iExpectedYear)] * 12 + [""]
                            or arrRows[1][:2] != ["", "月"]
                            or arrRows[1][14] != "合計"):
                        raise ValueError(
                            "年間表の列数・年月見出しが不正です（15列、O1は空欄、O2は「合計」）。"
                            "旧形式のTSVはstep0014から再作成してください。Path = "
                            + pszPath
                        )
                    arrMonths = [
                        int(pszMonth) if re.fullmatch(r"[0-9]+", pszMonth) else 0
                        for pszMonth in arrRows[1][2:14]
                    ]
                    if arrMonths != list(range(1, 13)):
                        raise ValueError(
                            "年間表の月見出しが不正です（01～12または1～12の月順）。Path = "
                            + pszPath
                        )
                arrTotalIndexes = [
                    i for i, row in enumerate(arrPreviousRows) if row[0] == "合計"
                ]
                if len(arrTotalIndexes) != 1:
                    raise ValueError("前年の合計ブロックが不足または重複しています。Path = " + pszPreviousPath)
                iStart = arrTotalIndexes[0]
                arrTotals = [list(row) for row in arrPreviousRows[iStart:iStart + 6]]
                arrItems = ["入雛羽数", "処理羽数", "処理重量", "平均体重", "育成率", "稼動日数"]
                if (len(arrTotals) != 6 or [row[1] for row in arrTotals] != arrItems
                        or any(row[0] for row in arrTotals[1:])
                        or iStart + 6 != len(arrPreviousRows)):
                    raise ValueError("前年の合計6行が不正です。Path = " + pszPreviousPath)
                arrTotals[0][0] = "前年実績"
                arrTotals[1][0] = f"{iYear - 1}年"
                pszNewline = "\r\n" if b"\r\n" in objCurrent else ("\r" if b"\r" in objCurrent else "\n")
                objNewline = pszNewline.encode("ascii")
                # 元の内容を保持し、最終行との間に空行をちょうど2行設けます。
                objOutput = objCurrent.rstrip(b"\r\n") + objNewline * 3
                objOutput += (pszNewline.join("\t".join(row) for row in arrTotals) + pszNewline).encode("utf-8")
            # 稼動日数行のO列だけを書き換え、他のバイト列・改行は保持します。
            arrOutputLines = objOutput.splitlines(keepends=True)
            for iLineIndex, objLine in enumerate(arrOutputLines):
                objBody = objLine.rstrip(b"\r\n")
                arrFields = objBody.split(b"\t")
                if len(arrFields) >= 2 and arrFields[1] == "稼動日数".encode("utf-8"):
                    if len(arrFields) != 15:
                        raise ValueError("稼動日数行は15列である必要があります。")
                    iTotalDays = 0
                    for objValue in arrFields[2:14]:
                        if objValue == b"":
                            continue
                        if re.fullmatch(rb"[0-9]+", objValue) is None:
                            raise ValueError("稼動日数は空欄または非負の整数である必要があります。")
                        iTotalDays += int(objValue)
                    arrFields[14] = str(iTotalDays).encode("ascii")
                    arrOutputLines[iLineIndex] = b"\t".join(arrFields) + objLine[len(objBody):]
            objOutput = b"".join(arrOutputLines)
            with open(pszOutputPath, "xb") as objFile:
                bReserved = True
                objFile.write(objOutput)
            bReserved = False
            arrCreatedPaths.append(pszOutputPath)
        except (OSError, UnicodeError, ValueError) as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszInputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("STEP0016_GROUP_DEDUCTION_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def create_step0017_group_deduction_files(arrStep0016Paths: List[str]) -> List[str]:
    """step0016の空行2行を一日平均処理羽数の前年比に置き換えます。"""
    arrCreatedPaths: List[str] = []
    arrErrors: List[str] = []
    for pszInputPath in arrStep0016Paths:
        bReserved = False
        pszOutputPath = ""
        try:
            objMatch = re.fullmatch(
                r"(.+)_step0016_グループ別_控除廃棄_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
                os.path.basename(pszInputPath),
            )
            if objMatch is None:
                raise ValueError("step0016のファイル名が不正です。")
            pszPrefix, pszYear = objMatch.groups()
            iYear = int(pszYear)
            pszOutputPath = os.path.join(
                os.path.dirname(pszInputPath),
                f"{pszPrefix}_step0017_グループ別_控除廃棄_{iYear:04d}_年間.tsv",
            )
            with open(pszInputPath, "rb") as objFile:
                objInput = objFile.read()
            objOutput = objInput
            arrRows = [line.split("\t") for line in objInput.decode("utf-8-sig").splitlines()]
            if (len(arrRows) < 2 or any(len(row) != 15 for row in arrRows[:2])
                    or arrRows[0][14] != "" or arrRows[1][14] != "合計"):
                raise ValueError(
                    "O1は空欄、O2は「合計」である必要があります。"
                    "旧形式のTSVはstep0014から再作成してください。Path = " + pszInputPath
                )
            if iYear != 2022:
                arrLines = objInput.splitlines(keepends=True)
                if (len(arrRows) != 31
                        or any(len(row) != 15 for i, row in enumerate(arrRows) if i not in (23, 24))):
                    raise ValueError("step0016は31行で、空行以外は15列である必要があります。")
                if any(value for i in (23, 24) for value in arrRows[i]):
                    raise ValueError("24・25行目は空行である必要があります。")
                if (arrRows[0] != ["", "年"] + [str(iYear)] * 12 + [""]
                        or arrRows[1][:2] != ["", "月"] or arrRows[1][14] != "合計"
                        or any(re.fullmatch(r"[0-9]+", v) is None for v in arrRows[1][2:14])
                        or [int(v) for v in arrRows[1][2:14]] != list(range(1, 13))):
                    raise ValueError("step0016の年月見出しが不正です。")
                arrItems = ["入雛羽数", "処理羽数", "処理重量", "平均体重", "育成率", "稼動日数"]
                if (arrRows[17][:2] != ["合計", "入雛羽数"]
                        or [row[1] for row in arrRows[17:23]] != arrItems
                        or [row[1] for row in arrRows[25:31]] != arrItems
                        or arrRows[25][0] != "前年実績"
                        or arrRows[26][0] != f"{iYear - 1}年"):
                    raise ValueError("当年・前年の合計項目または前年表示が不正です。")
                arrRatios: List[str] = []
                for iColumn in range(2, 15):
                    try:
                        arrValues = [Decimal(arrRows[iRow][iColumn] or "0") for iRow in (18, 22, 26, 30)]
                        if any(not value.is_finite() for value in arrValues):
                            raise ValueError("non-finite value")
                        objCurrent, objCurrentDays, objPrevious, objPreviousDays = arrValues
                        if objCurrentDays == 0 or objPreviousDays == 0 or objPrevious == 0:
                            raise ValueError("division by zero")
                        objRatio = ((objCurrent / objCurrentDays) / (objPrevious / objPreviousDays)).quantize(
                            Decimal("0.0001"), rounding=ROUND_HALF_UP
                        )
                        arrRatios.append(format(objRatio, ".4f"))
                    except (ArithmeticError, ValueError):
                        arrRatios.append("")
                objNewline = b"\r\n" if b"\r\n" in objInput else (b"\r" if b"\r" in objInput else b"\n")
                arrLines[23] = "\t".join(["処理羽数", "前年比"] + arrRatios).encode("utf-8") + objNewline
                arrLines[24] = "\t".join(["", "(一日平均)"] + [""] * 13).encode("utf-8") + objNewline
                objOutput = b"".join(arrLines)
            with open(pszOutputPath, "xb") as objFile:
                bReserved = True
                objFile.write(objOutput)
            bReserved = False
            arrCreatedPaths.append(pszOutputPath)
        except (OSError, UnicodeError, ValueError) as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszInputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("STEP0017_GROUP_DEDUCTION_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def create_step0013_variety_files(arrStep0012Paths: List[str]) -> List[str]:
    """step0012の空行2行を一日平均出荷羽数の前年比に置き換えます。"""
    arrCreatedPaths: List[str] = []
    arrErrors: List[str] = []
    for pszInputPath in arrStep0012Paths:
        bReserved = False
        pszOutputPath = ""
        try:
            objMatch = re.fullmatch(
                r"(.+)_step0012_品種別_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
                os.path.basename(pszInputPath),
            )
            if objMatch is None:
                raise ValueError("step0012のファイル名が不正です。")
            pszPrefix, pszYear = objMatch.groups()
            iYear = int(pszYear)
            pszOutputPath = os.path.join(
                os.path.dirname(pszInputPath),
                f"{pszPrefix}_step0013_品種別_{iYear:04d}_年間.tsv",
            )
            with open(pszInputPath, "rb") as objFile:
                objInput = objFile.read()
            objOutput = objInput
            if iYear != 2022:
                arrLines = objInput.splitlines(keepends=True)
                arrRows = [line.split("\t") for line in objInput.decode("utf-8-sig").splitlines()]
                if (len(arrRows) != 48
                        or any(len(row) != 15 for i, row in enumerate(arrRows) if i not in (39, 40))):
                    raise ValueError("step0012は48行で、空行以外は15列である必要があります。")
                if any(value for i in (39, 40) for value in arrRows[i]):
                    raise ValueError("40・41行目は空行である必要があります。")
                if (arrRows[0] != ["", "年"] + [str(iYear)] * 12 + [""]
                        or arrRows[1] != ["", "月"] + [str(m) for m in range(1, 13)] + ["合計"]):
                    raise ValueError("step0012の年月見出しが不正です。")
                arrStructure = [("", "年"), ("", "月")]
                for pszName in [name for _, name in SPLIT_VARIETIES] + ["合計"]:
                    arrStructure.extend(
                        (pszName if i == 0 else "", item)
                        for i, item in enumerate(STEP0008_VARIETY_ITEM_NAMES)
                    )
                arrStructure.append(("", "稼動日数"))
                arrPreviousStructure = [
                    ("前年実績", "入雛羽数"), (f"{iYear - 1}年", "出荷羽数"),
                    ("", "正味重量"), ("", "平均体重"), ("", "飼育日数平均"),
                    ("", "育成率"), ("", "稼動日数"),
                ]
                if ([tuple(row[:2]) for row in arrRows[:39]] != arrStructure
                        or [tuple(row[:2]) for row in arrRows[41:48]] != arrPreviousStructure):
                    raise ValueError("品種・当年合計・前年実績の項目構成が不正です。")
                for iRow in (38, 47):
                    if any(v != "" and re.fullmatch(r"[0-9]+", v) is None
                           for v in arrRows[iRow][2:14]):
                        raise ValueError("稼動日数は空欄または非負の整数である必要があります。")
                arrRatios: List[str] = []
                for iColumn in range(2, 15):
                    try:
                        arrValues = [Decimal(arrRows[iRow][iColumn] or "0") for iRow in (33, 38, 42, 47)]
                        if any(not value.is_finite() for value in arrValues):
                            raise ValueError("non-finite value")
                        objCurrent, objCurrentDays, objPrevious, objPreviousDays = arrValues
                        if objCurrentDays == 0 or objPreviousDays == 0 or objPrevious == 0:
                            raise ValueError("division by zero")
                        objRatio = ((objCurrent / objCurrentDays) / (objPrevious / objPreviousDays)).quantize(
                            Decimal("0.0001"), rounding=ROUND_HALF_UP
                        )
                        arrRatios.append(format(objRatio, ".4f"))
                    except (ArithmeticError, ValueError):
                        arrRatios.append("")
                objNewline = b"\r\n" if b"\r\n" in objInput else (b"\r" if b"\r" in objInput else b"\n")
                arrLines[39] = "\t".join(["出荷羽数", "前年比"] + arrRatios).encode("utf-8") + objNewline
                arrLines[40] = "\t".join(["", "(一日平均)"] + [""] * 13).encode("utf-8") + objNewline
                objOutput = b"".join(arrLines)
            with open(pszOutputPath, "xb") as objFile:
                bReserved = True
                objFile.write(objOutput)
            bReserved = False
            arrCreatedPaths.append(pszOutputPath)
        except (OSError, UnicodeError, ValueError) as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszInputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("STEP0013_VARIETY_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def create_step0013_variety_deduction_files(arrStep0012Paths: List[str]) -> List[str]:
    """step0012の空行2行を一日平均処理羽数の前年比に置き換えます。"""
    arrCreatedPaths: List[str] = []
    arrErrors: List[str] = []
    for pszInputPath in arrStep0012Paths:
        bReserved = False
        pszOutputPath = ""
        try:
            objMatch = re.fullmatch(
                r"(.+)_step0012_品種別_控除廃棄_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
                os.path.basename(pszInputPath),
            )
            if objMatch is None:
                raise ValueError("step0012のファイル名が不正です。")
            pszPrefix, pszYear = objMatch.groups()
            iYear = int(pszYear)
            pszOutputPath = os.path.join(
                os.path.dirname(pszInputPath),
                f"{pszPrefix}_step0013_品種別_控除廃棄_{iYear:04d}_年間.tsv",
            )
            with open(pszInputPath, "rb") as objFile:
                objInput = objFile.read()
            objOutput = objInput
            if iYear != 2022:
                arrLines = objInput.splitlines(keepends=True)
                arrRows = [line.split("\t") for line in objInput.decode("utf-8-sig").splitlines()]
                if (len(arrRows) != 48
                        or any(len(row) != 15 for i, row in enumerate(arrRows) if i not in (39, 40))):
                    raise ValueError("step0012は48行で、空行以外は15列である必要があります。")
                if any(value for i in (39, 40) for value in arrRows[i]):
                    raise ValueError("40・41行目は空行である必要があります。")
                if (arrRows[0] != ["", "年"] + [str(iYear)] * 12 + [""]
                        or arrRows[1] != ["", "月"] + [str(m) for m in range(1, 13)] + ["合計"]):
                    raise ValueError("step0012の年月見出しが不正です。")
                arrStructure = [("", "年"), ("", "月")]
                for pszName in [name for _, name in SPLIT_VARIETIES] + ["合計"]:
                    arrStructure.extend(
                        (pszName if i == 0 else "", item)
                        for i, item in enumerate(STEP0008_VARIETY_DEDUCTION_ITEM_NAMES)
                    )
                arrStructure.append(("", "稼動日数"))
                arrPreviousStructure = [
                    ("前年実績", "入雛羽数"), (f"{iYear - 1}年", "処理羽数"),
                    ("", "処理重量"), ("", "控除後平均体重"), ("", "飼育日数平均"),
                    ("", "育成率"), ("", "稼動日数"),
                ]
                if ([tuple(row[:2]) for row in arrRows[:39]] != arrStructure
                        or [tuple(row[:2]) for row in arrRows[41:48]] != arrPreviousStructure):
                    raise ValueError("品種・当年合計・前年実績の項目構成が不正です。")
                for iRow in (38, 47):
                    if any(v != "" and re.fullmatch(r"[0-9]+", v) is None
                           for v in arrRows[iRow][2:14]):
                        raise ValueError("稼動日数は空欄または非負の整数である必要があります。")
                arrRatios: List[str] = []
                for iColumn in range(2, 15):
                    try:
                        arrValues = [Decimal(arrRows[iRow][iColumn] or "0") for iRow in (33, 38, 42, 47)]
                        if any(not value.is_finite() for value in arrValues):
                            raise ValueError("non-finite value")
                        objCurrent, objCurrentDays, objPrevious, objPreviousDays = arrValues
                        if objCurrentDays == 0 or objPreviousDays == 0 or objPrevious == 0:
                            raise ValueError("division by zero")
                        objRatio = ((objCurrent / objCurrentDays) / (objPrevious / objPreviousDays)).quantize(
                            Decimal("0.0001"), rounding=ROUND_HALF_UP
                        )
                        arrRatios.append(format(objRatio, ".4f"))
                    except (ArithmeticError, ValueError):
                        arrRatios.append("")
                objNewline = b"\r\n" if b"\r\n" in objInput else (b"\r" if b"\r" in objInput else b"\n")
                arrLines[39] = "\t".join(["処理羽数", "前年比"] + arrRatios).encode("utf-8") + objNewline
                arrLines[40] = "\t".join(["", "(一日平均)"] + [""] * 13).encode("utf-8") + objNewline
                objOutput = b"".join(arrLines)
            with open(pszOutputPath, "xb") as objFile:
                bReserved = True
                objFile.write(objOutput)
            bReserved = False
            arrCreatedPaths.append(pszOutputPath)
        except (OSError, UnicodeError, ValueError) as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszInputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("STEP0013_VARIETY_DEDUCTION_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def create_step0013_group_files(arrStep0012Paths: List[str]) -> List[str]:
    """step0012の空行2行を一日平均出荷羽数の前年比に置き換えます。"""
    arrCreatedPaths: List[str] = []
    arrErrors: List[str] = []
    for pszInputPath in arrStep0012Paths:
        bReserved = False
        pszOutputPath = ""
        try:
            objMatch = re.fullmatch(
                r"(.+)_step0012_グループ別_([0-9]{4})_年間(?:_[0-9]+)?\.tsv",
                os.path.basename(pszInputPath),
            )
            if objMatch is None:
                raise ValueError("step0012のファイル名が不正です。")
            pszPrefix, pszYear = objMatch.groups()
            iYear = int(pszYear)
            pszOutputPath = os.path.join(
                os.path.dirname(pszInputPath),
                f"{pszPrefix}_step0013_グループ別_{iYear:04d}_年間.tsv",
            )
            with open(pszInputPath, "rb") as objFile:
                objInput = objFile.read()
            objOutput = objInput
            if iYear != 2022:
                arrLines = objInput.splitlines(keepends=True)
                arrRows = [line.split("\t") for line in objInput.decode("utf-8-sig").splitlines()]
                if (len(arrRows) != 31
                        or any(len(row) != 15 for i, row in enumerate(arrRows) if i not in (23, 24))):
                    raise ValueError("step0012は31行で、空行以外は15列である必要があります。")
                if any(value for i in (23, 24) for value in arrRows[i]):
                    raise ValueError("24・25行目は空行である必要があります。")
                if (arrRows[0] != ["", "年"] + [str(iYear)] * 12 + [""]
                        or arrRows[1][:2] != ["", "月"] or arrRows[1][14] != "合計"
                        or any(re.fullmatch(r"[0-9]+", v) is None for v in arrRows[1][2:14])
                        or [int(v) for v in arrRows[1][2:14]] != list(range(1, 13))):
                    raise ValueError("step0012の年月見出しが不正です。")
                arrItems = ["入雛羽数", "出荷羽数", "正味重量", "平均体重", "育成率", "稼動日数"]
                if (arrRows[17][:2] != ["合計", "入雛羽数"]
                        or [row[1] for row in arrRows[17:23]] != arrItems
                        or [row[1] for row in arrRows[25:31]] != arrItems
                        or arrRows[25][0] != "前年実績"
                        or arrRows[26][0] != f"{iYear - 1}年"):
                    raise ValueError("当年・前年の合計項目または前年表示が不正です。")
                arrRatios: List[str] = []
                for iColumn in range(2, 15):
                    try:
                        arrValues = [Decimal(arrRows[iRow][iColumn] or "0") for iRow in (18, 22, 26, 30)]
                        if any(not value.is_finite() for value in arrValues):
                            raise ValueError("non-finite value")
                        objCurrent, objCurrentDays, objPrevious, objPreviousDays = arrValues
                        if objCurrentDays == 0 or objPreviousDays == 0 or objPrevious == 0:
                            raise ValueError("division by zero")
                        objRatio = ((objCurrent / objCurrentDays) / (objPrevious / objPreviousDays)).quantize(
                            Decimal("0.0001"), rounding=ROUND_HALF_UP
                        )
                        arrRatios.append(format(objRatio, ".4f"))
                    except (ArithmeticError, ValueError):
                        arrRatios.append("")
                objNewline = b"\r\n" if b"\r\n" in objInput else (b"\r" if b"\r" in objInput else b"\n")
                arrLines[23] = "\t".join(["出荷羽数", "前年比"] + arrRatios).encode("utf-8") + objNewline
                arrLines[24] = "\t".join(["", "(一日平均)"] + [""] * 13).encode("utf-8") + objNewline
                objOutput = b"".join(arrLines)
            with open(pszOutputPath, "xb") as objFile:
                bReserved = True
                objFile.write(objOutput)
            bReserved = False
            arrCreatedPaths.append(pszOutputPath)
        except (OSError, UnicodeError, ValueError) as objException:
            if bReserved:
                os.remove(pszOutputPath)
            arrErrors.append(pszInputPath + ": " + str(objException))
    if arrErrors:
        raise_conversion_error("STEP0013_GROUP_FAILED", "\n".join(arrErrors))
    return arrCreatedPaths


def convert_excel_to_tsv(pszOutputTsvPath: str, pszInputExcelPath: str) -> Tuple[int, int]:
    try:
        objFormulaWorkbook = load_workbook(
            pszInputExcelPath, data_only=False, read_only=False
        )
        objValueWorkbook = load_workbook(
            pszInputExcelPath, data_only=True, read_only=False
        )
    except Exception as objException:
        pszDetail: str = str(objException)
        pszCode: str = (
            "EXCEL_PASSWORD_PROTECTED"
            if "password" in pszDetail.lower() or "encrypted" in pszDetail.lower()
            else "EXCEL_OPEN_FAILED"
        )
        raise_conversion_error(pszCode, "failed to open Excel workbook. Detail = " + pszDetail)

    pszTemporaryTsvPath: str = ""
    bOutputReserved: bool = False
    try:
        if TARGET_SHEET_NAME not in objFormulaWorkbook.sheetnames:
            raise_conversion_error(
                "DETAIL_SHEET_NOT_FOUND",
                "worksheet was not found. Sheet = " + TARGET_SHEET_NAME,
            )

        objFormulaWorksheet = objFormulaWorkbook[TARGET_SHEET_NAME]
        objValueWorksheet = objValueWorkbook[TARGET_SHEET_NAME]
        objBounds = find_data_bounds(objFormulaWorksheet, objValueWorksheet)
        if objBounds is None:
            raise_conversion_error(
                "EMPTY_DETAIL_SHEET",
                "no data was found in worksheet. Sheet = " + TARGET_SHEET_NAME,
            )
        iLastRow, iLastColumn = objBounds

        reserve_output_path(pszOutputTsvPath)
        bOutputReserved = True
        try:
            iTemporaryFileDescriptor, pszTemporaryTsvPath = tempfile.mkstemp(
                prefix="group_shipment_summary_",
                suffix=".tsv",
                dir=os.path.dirname(pszOutputTsvPath),
            )
            os.close(iTemporaryFileDescriptor)
        except OSError as objException:
            raise_conversion_error(
                "TSV_TEMP_CREATE_FAILED",
                "failed to create temporary TSV. Detail = " + str(objException),
            )

        try:
            with open(pszTemporaryTsvPath, "w", encoding="utf-8", newline="") as objTsvFile:
                for iRow in range(1, iLastRow + 1):
                    arrValues = []
                    for iColumn in range(1, iLastColumn + 1):
                        objFormulaCell = objFormulaWorksheet.cell(iRow, iColumn)
                        objValueCell = objValueWorksheet.cell(iRow, iColumn)
                        if objFormulaCell.data_type == "f" and objValueCell.value is None:
                            raise_conversion_error(
                                "FORMULA_RESULT_UNAVAILABLE",
                                "saved formula result was not available. Cell = "
                                + objFormulaCell.coordinate,
                            )
                        arrValues.append(format_cell_value(objFormulaCell, objValueCell.value))
                    objTsvFile.write("\t".join(arrValues) + "\r\n")
            os.replace(pszTemporaryTsvPath, pszOutputTsvPath)
            pszTemporaryTsvPath = ""
            bOutputReserved = False
        except ConversionError:
            raise
        except Exception as objException:
            raise_conversion_error(
                "TSV_WRITE_FAILED", "failed to write TSV. Detail = " + str(objException)
            )

        print("対象シート: " + TARGET_SHEET_NAME)
        print(
            "出力範囲: A1:"
            + get_column_letter(iLastColumn)
            + str(iLastRow)
        )
        print("出力行数: " + str(iLastRow))
        print("出力列数: " + str(iLastColumn))
        return iLastRow, iLastColumn
    finally:
        objFormulaWorkbook.close()
        objValueWorkbook.close()
        if pszTemporaryTsvPath and os.path.exists(pszTemporaryTsvPath):
            os.remove(pszTemporaryTsvPath)
        if bOutputReserved and os.path.exists(pszOutputTsvPath):
            os.remove(pszOutputTsvPath)


def main() -> int:
    if len(sys.argv) != 3:
        print(
            "Usage: python group_shipment_summary_Cmd.py "
            "<output_tsv_path> <input_excel_path>",
            file=sys.stderr,
        )
        return 1

    pszOutputTsvPath: str = os.path.abspath(sys.argv[1])
    pszInputExcelPath: str = os.path.abspath(sys.argv[2])

    try:
        validate_paths(pszOutputTsvPath, pszInputExcelPath)
        convert_excel_to_tsv(pszOutputTsvPath, pszInputExcelPath)
        (
            arrMonthlyTsvPaths,
            objFirstYearMonth,
            objLastYearMonth,
        ) = split_tsv_by_month(pszOutputTsvPath)
        arrDailyOperationPaths = create_daily_operation_files(arrMonthlyTsvPaths)
        (
            arrVarietyTsvPaths,
            pszWarningPath,
            iWarningCount,
        ) = aggregate_step0001_files(arrMonthlyTsvPaths, pszInputExcelPath)
        (
            arrAggregatedTsvPaths,
            pszWarningPath,
            iWarningCount,
        ) = aggregate_step0001_files(
            arrMonthlyTsvPaths, pszInputExcelPath, False, pszWarningPath, iWarningCount
        )
        (
            arrTotalTsvPaths,
            pszWarningPath,
            iWarningCount,
        ) = create_step0002_total_files(
            arrMonthlyTsvPaths, pszInputExcelPath, pszWarningPath, iWarningCount
        )
        arrStep0003VarietyPaths, arrStep0003VarietyYearMonths = create_step0003_files(
            arrVarietyTsvPaths, objLastYearMonth[0], True
        )
        arrStep0003VarietyPreDiscardGrowthRatePaths = (
            create_step0003_variety_pre_discard_growth_rate_files(
                arrStep0003VarietyPaths,
                arrStep0003VarietyYearMonths,
            )
        )
        arrStep0003VarietyAdditionalPaths = (
            create_step0003_variety_additional_files(
                arrVarietyTsvPaths, objLastYearMonth[0]
            )
        )
        arrStep0003VarietyDayAgePaths = [
            pszPath
            for pszPath in arrStep0003VarietyAdditionalPaths
            if re.fullmatch(
                r".+_step0003_品種別_日令_\d{4}_\d{2}(?:_\d+)*\.tsv",
                os.path.basename(pszPath),
            )
        ]
        arrStep0003VarietyHusbandryAveragePaths = (
            create_step0003_variety_husbandry_average_files(
                arrStep0003VarietyDayAgePaths
            )
        )
        arrStep0007VarietyPaths = create_step0007_variety_files(
            arrStep0003VarietyPaths, arrStep0003VarietyYearMonths, arrMonthlyTsvPaths
        )
        arrStep0003VarietyDeductionGrowthRatePaths = (
            create_step0003_variety_deduction_growth_rate_files(
                arrStep0003VarietyPaths,
                arrStep0007VarietyPaths,
                arrStep0003VarietyYearMonths,
            )
        )
        (
            arrStep0004VarietyPaths,
            arrStep0004VarietyYearMonths,
        ) = create_step0004_variety_files(
            pszOutputTsvPath,
            arrStep0003VarietyPaths,
            arrStep0003VarietyYearMonths,
            arrStep0003VarietyHusbandryAveragePaths,
            arrStep0003VarietyPreDiscardGrowthRatePaths,
        )
        (
            arrStep0005VarietyPaths,
            arrStep0005VarietyYears,
        ) = create_step0005_variety_files(
            pszOutputTsvPath,
            arrStep0004VarietyPaths,
            arrStep0004VarietyYearMonths,
        )
        arrStep0006VarietyPaths = create_step0006_variety_files(
            arrStep0005VarietyPaths
        )
        (
            arrStep0007VarietyAnnualPaths,
            arrStep0007VarietyAnnualYears,
        ) = create_step0007_variety_annual_files(
            pszOutputTsvPath, arrStep0006VarietyPaths
        )
        (
            arrStep0008VarietyAnnualPaths,
            pszWarningPath,
            iWarningCount,
        ) = create_step0008_variety_annual_files(
            arrStep0007VarietyAnnualPaths,
            pszInputExcelPath,
            pszWarningPath,
            iWarningCount,
        )
        arrStep0009VarietyAnnualPaths = create_step0009_variety_annual_files(
            arrStep0008VarietyAnnualPaths
        )
        arrStep0010VarietyAnnualPaths = create_step0010_variety_annual_files(
            arrStep0009VarietyAnnualPaths
        )
        arrStep0011VarietyPaths = create_step0011_variety_files(
            arrStep0010VarietyAnnualPaths, arrDailyOperationPaths
        )
        arrStep0012VarietyPaths = create_step0012_variety_files(arrStep0011VarietyPaths)
        arrStep0013VarietyPaths = create_step0013_variety_files(arrStep0012VarietyPaths)
        arrStep0013VarietyExcelPaths = create_step0013_variety_excel_files(
            arrStep0013VarietyPaths
        )
        arrNamedVarietySummaryExcelPaths = create_variety_summary_yearly_named_copies(
            arrStep0013VarietyExcelPaths
        )
        (
            arrStep0004VarietyDeductionPaths,
            arrStep0004VarietyDeductionYearMonths,
        ) = create_step0004_variety_deduction_files(
            pszOutputTsvPath,
            arrStep0003VarietyPaths,
            arrStep0003VarietyYearMonths,
            arrStep0007VarietyPaths,
            arrStep0003VarietyHusbandryAveragePaths,
            arrStep0003VarietyDeductionGrowthRatePaths,
        )
        (
            arrStep0005VarietyDeductionPaths,
            arrStep0005VarietyDeductionYears,
        ) = create_step0005_variety_deduction_files(
            pszOutputTsvPath,
            arrStep0004VarietyDeductionPaths,
            arrStep0004VarietyDeductionYearMonths,
        )
        arrStep0006VarietyDeductionPaths = (
            create_step0006_variety_deduction_files(
                arrStep0005VarietyDeductionPaths
            )
        )
        (
            arrStep0007VarietyDeductionPaths,
            arrStep0007VarietyDeductionYears,
        ) = create_step0007_variety_deduction_annual_files(
            pszOutputTsvPath, arrStep0006VarietyDeductionPaths
        )
        (
            arrStep0008VarietyDeductionPaths,
            pszWarningPath,
            iWarningCount,
        ) = create_step0008_variety_deduction_annual_files(
            arrStep0007VarietyDeductionPaths,
            pszInputExcelPath,
            pszWarningPath,
            iWarningCount,
        )
        arrStep0009VarietyDeductionPaths = (
            create_step0009_variety_deduction_annual_files(
                arrStep0008VarietyDeductionPaths
            )
        )
        arrStep0010VarietyDeductionAnnualPaths = (
            create_step0010_variety_deduction_annual_files(
                arrStep0009VarietyDeductionPaths
            )
        )
        arrStep0011VarietyDeductionPaths = create_step0011_variety_deduction_files(
            arrStep0010VarietyDeductionAnnualPaths, arrDailyOperationPaths
        )
        arrStep0012VarietyDeductionPaths = create_step0012_variety_deduction_files(
            arrStep0011VarietyDeductionPaths
        )
        arrStep0013VarietyDeductionPaths = create_step0013_variety_deduction_files(
            arrStep0012VarietyDeductionPaths
        )
        arrStep0013VarietyDeductionExcelPaths = create_step0013_variety_deduction_excel_files(
            arrStep0013VarietyDeductionPaths
        )
        arrNamedVarietyDeductionSummaryExcelPaths = create_variety_deduction_summary_yearly_named_copies(
            arrStep0013VarietyDeductionExcelPaths
        )
        arrStep0003Paths, arrStep0003YearMonths = create_step0003_files(
            arrAggregatedTsvPaths, objLastYearMonth[0]
        )
        arrStep0003TotalPaths, arrStep0003TotalYearMonths = create_step0003_total_files(
            arrTotalTsvPaths, objLastYearMonth[0]
        )
        (
            arrStep0004PlacementAveragePaths,
            arrStep0004PlacementAverageYearMonths,
        ) = create_step0004_total_placement_average_files(arrStep0003TotalPaths)
        (
            arrStep0004ShipmentCountWeightPaths,
            arrStep0004ShipmentCountWeightYearMonths,
        ) = create_step0004_total_shipment_count_weight_files(arrStep0003TotalPaths)
        (
            arrStep0004TotalWeightPaths,
            arrStep0004TotalWeightYearMonths,
        ) = create_step0004_total_weight_files(arrStep0003TotalPaths)
        (
            arrStep0004TotalCountPaths,
            arrStep0004TotalCountYearMonths,
        ) = create_step0004_total_count_files(arrStep0003TotalPaths)
        arrStep0004TotalCountHistoryPaths = create_step0004_total_count_history_files(
            arrStep0003TotalPaths, objLastYearMonth[0]
        )
        arrStep0005TotalCountHistoryPaths = (
            create_step0005_total_count_history_files(
                arrStep0004TotalCountHistoryPaths
            )
        )
        arrStep0005TotalCountHistoryExcelPaths = (
            create_step0005_total_count_history_excel_files(
                arrStep0005TotalCountHistoryPaths
            )
        )
        arrNamedTotalCountHistoryExcelPaths = create_total_count_history_named_copies(
            arrStep0005TotalCountHistoryExcelPaths
        )
        arrStep0004GroupDiscardCountPaths = create_step0004_group_discard_count_files(
            arrStep0003Paths, arrStep0003YearMonths
        )
        arrStep0005GroupDiscardCountPaths = create_step0005_group_discard_count_files(
            arrStep0004GroupDiscardCountPaths
        )
        arrStep0006GroupDiscardCountPaths = create_step0006_group_discard_count_files(
            arrStep0005GroupDiscardCountPaths
        )
        arrStep0007GroupTotalDiscardCountPaths = (
            create_step0007_group_total_discard_count_files(
                arrStep0006GroupDiscardCountPaths
            )
        )
        arrStep0004GroupDiscardWeightPaths = create_step0004_group_discard_weight_files(
            arrStep0003Paths, arrStep0003YearMonths
        )
        arrStep0005GroupDiscardWeightPaths = create_step0005_group_discard_weight_files(
            arrStep0004GroupDiscardWeightPaths
        )
        arrStep0006GroupDiscardWeightPaths = create_step0006_group_discard_weight_files(
            arrStep0005GroupDiscardWeightPaths
        )
        arrStep0007GroupTotalDiscardWeightPaths = (
            create_step0007_group_total_discard_weight_files(
                arrStep0006GroupDiscardWeightPaths
            )
        )
        (
            arrStep0004TotalPreSlaughterCountPaths,
            arrStep0004TotalPreSlaughterCountYearMonths,
        ) = create_step0004_total_pre_slaughter_count_files(arrStep0003TotalPaths)
        (
            arrStep0004TotalPreSlaughterWeightPaths,
            arrStep0004TotalPreSlaughterWeightYearMonths,
        ) = create_step0004_total_pre_slaughter_weight_files(arrStep0003TotalPaths)
        (
            arrStep0005PreSlaughterCountPaths,
            arrStep0005PreSlaughterWeightPaths,
            arrStep0005PreSlaughterYearMonths,
        ) = create_step0005_pre_slaughter_subtotal_files(
            arrStep0004TotalPreSlaughterCountPaths, arrStep0004TotalPreSlaughterWeightPaths
        )
        (
            arrStep0005DiscardCountPaths,
            arrStep0005DiscardWeightPaths,
            arrStep0005DiscardYearMonths,
        ) = create_step0005_discard_subtotal_files(
            arrStep0004TotalCountPaths, arrStep0004TotalWeightPaths
        )
        (
            arrStep0006DiscardSummaryLowerPaths,
            arrStep0006DiscardSummaryLowerYearMonths,
        ) = create_step0006_discard_summary_lower_files(
            arrStep0004ShipmentCountWeightPaths,
            arrStep0005PreSlaughterWeightPaths,
            arrStep0005DiscardWeightPaths,
        )
        (
            arrStep0006DiscardSummaryUpperPaths,
            arrStep0006DiscardSummaryUpperYearMonths,
        ) = create_step0006_discard_summary_upper_files(
            arrStep0004PlacementAveragePaths,
            arrStep0005PreSlaughterCountPaths,
            arrStep0005DiscardCountPaths,
        )
        (
            arrStep0007DiscardSummaryUpperPaths,
            arrStep0007DiscardSummaryUpperYearMonths,
        ) = create_step0007_discard_summary_upper_files(
            arrStep0006DiscardSummaryUpperPaths
        )
        (
            arrStep0007DiscardSummaryLowerPaths,
            arrStep0007DiscardSummaryLowerYearMonths,
        ) = create_step0007_discard_summary_lower_files(
            arrStep0006DiscardSummaryLowerPaths
        )
        (
            arrStep0008DiscardSummaryPaths,
            arrStep0008DiscardSummaryYearMonths,
        ) = create_step0008_discard_summary_files(
            arrStep0007DiscardSummaryUpperPaths,
            arrStep0007DiscardSummaryLowerPaths,
        )
        (
            arrStep0009DiscardSummaryPaths,
            arrStep0009DiscardSummaryYearMonths,
        ) = create_step0009_discard_summary_files(arrStep0008DiscardSummaryPaths)
        arrStep0010DiscardSummaryPaths = create_step0010_discard_summary_files(
            arrStep0009DiscardSummaryPaths
        )
        arrStep0011DiscardSummaryPaths = create_step0011_discard_summary_files(
            arrStep0010DiscardSummaryPaths
        )
        arrStep0012DiscardSummaryPaths = create_step0012_discard_summary_files(
            arrStep0011DiscardSummaryPaths
        )
        arrStep0012DiscardSummaryExcelPaths = (
            create_step0012_discard_summary_excel_files(
                arrStep0012DiscardSummaryPaths
            )
        )
        arrNamedDiscardSummaryExcelPaths = (
            create_discard_summary_monthly_named_copies(
                arrStep0012DiscardSummaryExcelPaths
            )
        )
        arrStep0004Paths, arrStep0004YearMonths = create_step0004_files(
            pszOutputTsvPath, arrStep0003Paths, arrStep0003YearMonths
        )
        arrStep0008GroupDiscardDeductionPaths = (
            create_step0008_group_discard_deduction_files(
                arrStep0004Paths,
                arrStep0007GroupTotalDiscardCountPaths,
                arrStep0007GroupTotalDiscardWeightPaths,
                arrStep0003Paths,
                arrMonthlyTsvPaths,
            )
        )
        arrStep0009GroupDiscardPaths = create_step0009_group_discard_files(
            arrStep0008GroupDiscardDeductionPaths
        )
        arrStep0010GroupDiscardAnnualPaths = (
            create_step0010_group_discard_annual_files(
                arrStep0009GroupDiscardPaths
            )
        )
        arrStep0011GroupDeductionPaths = (
            create_step0011_group_deduction_files(
                arrStep0010GroupDiscardAnnualPaths
            )
        )
        arrStep0012GroupDeductionPaths, _ = (
            create_step0012_group_deduction_files(
                pszOutputTsvPath, arrStep0011GroupDeductionPaths
            )
        )
        arrStep0013GroupDeductionPaths, _ = (
            create_step0013_group_deduction_files(
                arrStep0012GroupDeductionPaths
            )
        )
        arrStep0013GroupDeductionExcelPaths: List[str] = []
        arrStep0014GroupDeductionPaths = (
            create_step0014_group_deduction_files(
                arrStep0013GroupDeductionPaths
            )
        )
        arrStep0014GroupDeductionExcelPaths: List[str] = []
        arrStep0015GroupDeductionPaths = (
            create_step0015_group_deduction_files(
                arrStep0014GroupDeductionPaths, arrDailyOperationPaths
            )
        )
        arrStep0016GroupDeductionPaths = create_step0016_group_deduction_files(
            arrStep0015GroupDeductionPaths
        )
        arrStep0017GroupDeductionPaths = create_step0017_group_deduction_files(
            arrStep0016GroupDeductionPaths
        )
        arrStep0017GroupDeductionExcelPaths = (
            create_step0017_group_deduction_excel_files(
                arrStep0017GroupDeductionPaths
            )
        )
        arrNamedGroupDeductionSummaryExcelPaths = create_group_deduction_summary_yearly_named_copies(
            arrStep0017GroupDeductionExcelPaths
        )
        arrStep0005Paths, arrStep0005Years = create_step0005_files(
            pszOutputTsvPath, arrStep0004Paths, arrStep0004YearMonths
        )
        arrStep0006Paths = create_step0006_files(arrStep0005Paths)
        arrStep0007Paths, arrStep0007Years = create_step0007_files(
            pszOutputTsvPath, arrStep0006Paths
        )
        arrStep0007ExcelPaths = create_step0007_excel_files(arrStep0007Paths)
        (
            arrStep0008Paths,
            pszWarningPath,
            iWarningCount,
        ) = create_step0008_files(
            arrStep0007Paths,
            pszInputExcelPath,
            pszWarningPath,
            iWarningCount,
        )
        arrStep0008ExcelPaths: List[str] = []
        arrStep0009Paths = create_step0009_files(arrStep0008Paths)
        arrStep0009ExcelPaths: List[str] = []
        arrStep0010Paths = create_step0010_files(arrStep0009Paths)
        arrStep0011GroupPaths = create_step0011_group_files(
            arrStep0010Paths, arrDailyOperationPaths
        )
        arrStep0012GroupPaths = create_step0012_group_files(arrStep0011GroupPaths)
        arrStep0013GroupPaths = create_step0013_group_files(arrStep0012GroupPaths)
        arrStep0013GroupExcelPaths = create_step0013_group_excel_files(
            arrStep0013GroupPaths
        )
        arrNamedGroupSummaryExcelPaths = create_group_summary_yearly_named_copies(
            arrStep0013GroupExcelPaths
        )
    except Exception as objException:
        pszErrorCode: str = getattr(
            objException, "pszErrorCode", "UNEXPECTED_CONVERSION_ERROR"
        )
        pszErrorDetail: str = str(objException)
        print(
            "Error: failed to create, split, aggregate, extract, combine, concatenate, transpose, join annual groups, create annual Excel, add growth rates, create deduction variety step0008 and step0009 annual TSVs, create variety step0009 and step0010 annual TSVs or step0010 Excel files, create variety deduction step0009 and step0010 annual TSVs or step0010 Excel files, add annual totals, or create step0011 and step0013 group annual TSV or Excel files, step0014 group deduction annual TSV, step0015 group deduction annual TSV, or step0017 group deduction annual TSV or Excel files. Detail = "
            + pszErrorDetail,
            file=sys.stderr,
        )
        try:
            pszErrorPath: str = write_error_file(
                pszInputExcelPath, pszErrorCode, pszErrorDetail
            )
            print("Error file: " + pszErrorPath, file=sys.stderr)
        except Exception as objErrorFileException:
            print(
                "Error: failed to create error file. Detail = "
                + str(objErrorFileException),
                file=sys.stderr,
            )
        return 1

    print("元TSV: " + pszOutputTsvPath)
    print("step0001ファイル数: " + str(len(arrMonthlyTsvPaths)))
    print("日別稼動判定ファイル数: " + str(len(arrDailyOperationPaths)))
    for pszPath in arrDailyOperationPaths:
        print("日別稼動判定ファイル: " + pszPath)
    print("step0002品種別ファイル数: " + str(len(arrVarietyTsvPaths)))
    for pszPath in arrVarietyTsvPaths:
        print("step0002品種別ファイル: " + pszPath)
    print("品種集計順: 阿波尾鶏、赤どり、地養鳥、ブロイラー、ブロイラー（貞光分）、ブロイラー(外部購入)")
    print("step0002グループ別ファイル数: " + str(len(arrAggregatedTsvPaths)))
    print("step0002総計ファイル数: " + str(len(arrTotalTsvPaths)))
    for pszPath in arrTotalTsvPaths:
        print("step0002総計ファイル: " + pszPath)
    print("step0003品種別ファイル数: " + str(len(arrStep0003VarietyPaths)))
    print("step0004品種別ファイル数: " + str(len(arrStep0004VarietyPaths)))
    print(
        "step0004品種別控除廃棄ファイル数: "
        + str(len(arrStep0004VarietyDeductionPaths))
    )
    print(
        "step0005品種別控除廃棄年間ファイル数: "
        + str(len(arrStep0005VarietyDeductionPaths))
    )
    print(
        "step0006品種別控除廃棄年間ファイル数: "
        + str(len(arrStep0006VarietyDeductionPaths))
    )
    print(
        "step0007品種別控除廃棄対象年数: "
        + str(len(arrStep0007VarietyDeductionYears))
    )
    print(
        "step0007品種別控除廃棄年間ファイル数: "
        + str(len(arrStep0007VarietyDeductionPaths))
    )
    print(
        "step0008品種別控除廃棄年間ファイル数: "
        + str(len(arrStep0008VarietyDeductionPaths))
    )
    print(
        "step0009品種別控除廃棄年間ファイル数: "
        + str(len(arrStep0009VarietyDeductionPaths))
    )
    print(
        "step0010品種別控除廃棄年間ファイル数: "
        + str(len(arrStep0010VarietyDeductionAnnualPaths))
    )
    print(
        "step0013品種別控除廃棄年間Excelファイル数: "
        + str(len(arrStep0013VarietyDeductionExcelPaths))
    )
    print("step0005品種別年間ファイル数: " + str(len(arrStep0005VarietyPaths)))
    print("step0006品種別年間ファイル数: " + str(len(arrStep0006VarietyPaths)))
    print("step0007品種別対象年数: " + str(len(arrStep0007VarietyAnnualYears)))
    print(
        "step0007品種別年間ファイル数: "
        + str(len(arrStep0007VarietyAnnualPaths))
    )
    print(
        "step0008品種別年間ファイル数: "
        + str(len(arrStep0008VarietyAnnualPaths))
    )
    print(
        "step0009品種別年間ファイル数: "
        + str(len(arrStep0009VarietyAnnualPaths))
    )
    print(
        "step0010品種別年間ファイル数: "
        + str(len(arrStep0010VarietyAnnualPaths))
    )
    print(
        "step0013品種別年間Excelファイル数: "
        + str(len(arrStep0013VarietyExcelPaths))
    )
    for pszPath in arrStep0013VarietyDeductionExcelPaths:
        print("step0013品種別控除廃棄年間Excelファイル: " + pszPath)
    print("年付き品種別控除廃棄実績集計Excelファイル数: " + str(len(arrNamedVarietyDeductionSummaryExcelPaths)))
    for pszPath in arrNamedVarietyDeductionSummaryExcelPaths:
        print("年付き品種別控除廃棄実績集計Excelファイル: " + pszPath)
    print("step0011品種別控除廃棄年間ファイル数: " + str(len(arrStep0011VarietyDeductionPaths)))
    for pszPath in arrStep0011VarietyDeductionPaths:
        print("step0011品種別控除廃棄年間ファイル: " + pszPath)
    print("step0012品種別控除廃棄年間ファイル数: " + str(len(arrStep0012VarietyDeductionPaths)))
    for pszPath in arrStep0012VarietyDeductionPaths:
        print("step0012品種別控除廃棄年間ファイル: " + pszPath)
    print("step0013品種別控除廃棄年間ファイル数: " + str(len(arrStep0013VarietyDeductionPaths)))
    for pszPath in arrStep0013VarietyDeductionPaths:
        print("step0013品種別控除廃棄年間ファイル: " + pszPath)
    for pszPath in arrStep0013VarietyExcelPaths:
        print("step0013品種別年間Excelファイル: " + pszPath)
    print("年付き品種別実績集計Excelファイル数: " + str(len(arrNamedVarietySummaryExcelPaths)))
    for pszPath in arrNamedVarietySummaryExcelPaths:
        print("年付き品種別実績集計Excelファイル: " + pszPath)
    print("step0011品種別年間ファイル数: " + str(len(arrStep0011VarietyPaths)))
    for pszPath in arrStep0011VarietyPaths:
        print("step0011品種別年間ファイル: " + pszPath)
    print("step0012品種別年間ファイル数: " + str(len(arrStep0012VarietyPaths)))
    for pszPath in arrStep0012VarietyPaths:
        print("step0012品種別年間ファイル: " + pszPath)
    print("step0013品種別年間ファイル数: " + str(len(arrStep0013VarietyPaths)))
    for pszPath in arrStep0013VarietyPaths:
        print("step0013品種別年間ファイル: " + pszPath)
    print(
        "step0003品種別廃棄控除前育成率ファイル数: "
        + str(len(arrStep0003VarietyPreDiscardGrowthRatePaths))
    )
    print(
        "step0003品種別追加抽出ファイル数: "
        + str(len(arrStep0003VarietyAdditionalPaths))
    )
    print(
        "step0003品種別飼育日数平均ファイル数: "
        + str(len(arrStep0003VarietyHusbandryAveragePaths))
    )
    print("step0003品種別控除後計算ファイル数: " + str(len(arrStep0007VarietyPaths)))
    print(
        "step0003品種別控除廃棄育成率ファイル数: "
        + str(len(arrStep0003VarietyDeductionGrowthRatePaths))
    )
    print("集計グループ数: 3")
    print("最初の年月: " + f"{objFirstYearMonth[0]:04d}_{objFirstYearMonth[1]:02d}")
    print("最後の年月: " + f"{objLastYearMonth[0]:04d}_{objLastYearMonth[1]:02d}")
    print("step0003最新年: " + str(objLastYearMonth[0]))
    print("step0003前年: " + str(objLastYearMonth[0] - 1))
    print(
        "step0003対象年: "
        + ", ".join(
            str(iYear) for iYear in range(objLastYearMonth[0] - 4, objLastYearMonth[0] + 1)
        )
    )
    print("step0003項目数: " + str(len(STEP0003_ITEMS)))
    print("step0003ファイル数: " + str(len(arrStep0003Paths)))
    print("step0003総計ファイル数: " + str(len(arrStep0003TotalPaths)))
    print("step0004総計入雛羽数・平均体重ファイル数: " + str(len(arrStep0004PlacementAveragePaths)))
    print("step0004総計出荷羽数・出荷重量ファイル数: " + str(len(arrStep0004ShipmentCountWeightPaths)))
    print("step0004総計重量ファイル数: " + str(len(arrStep0004TotalWeightPaths)))
    print("step0004総計羽数ファイル数: " + str(len(arrStep0004TotalCountPaths)))
    print(
        "step0004総計羽数過去5年ファイル数: "
        + str(len(arrStep0004TotalCountHistoryPaths))
    )
    for pszPath in arrStep0004TotalCountHistoryPaths:
        print("step0004総計羽数過去5年ファイル: " + pszPath)
    print(
        "step0005総計出荷廃棄羽数過去5年ファイル数: "
        + str(len(arrStep0005TotalCountHistoryPaths))
    )
    for pszPath in arrStep0005TotalCountHistoryPaths:
        print("step0005総計出荷廃棄羽数過去5年ファイル: " + pszPath)
    print(
        "step0005総計出荷廃棄羽数過去5年Excelファイル数: "
        + str(len(arrStep0005TotalCountHistoryExcelPaths))
    )
    for pszPath in arrStep0005TotalCountHistoryExcelPaths:
        print("step0005総計出荷廃棄羽数過去5年Excelファイル: " + pszPath)
    print("出荷廃棄羽数過去5年集計Excelファイル数: " + str(len(arrNamedTotalCountHistoryExcelPaths)))
    for pszPath in arrNamedTotalCountHistoryExcelPaths:
        print("出荷廃棄羽数過去5年集計Excelファイル: " + pszPath)
    print(
        "step0004グループ別廃棄羽数ファイル数: "
        + str(len(arrStep0004GroupDiscardCountPaths))
    )
    print(
        "step0004グループ別廃棄重量ファイル数: "
        + str(len(arrStep0004GroupDiscardWeightPaths))
    )
    print(
        "step0008グループ別控除廃棄ファイル数: "
        + str(len(arrStep0008GroupDiscardDeductionPaths))
    )
    print(
        "step0009グループ別控除廃棄ファイル数: "
        + str(len(arrStep0009GroupDiscardPaths))
    )
    print(
        "step0010グループ別控除廃棄年間ファイル数: "
        + str(len(arrStep0010GroupDiscardAnnualPaths))
    )
    print(
        "step0011グループ別控除廃棄年間行列入れ替えファイル数: "
        + str(len(arrStep0011GroupDeductionPaths))
    )
    print(
        "step0012グループ別控除廃棄年間結合ファイル数: "
        + str(len(arrStep0012GroupDeductionPaths))
    )
    print(
        "step0013グループ別控除廃棄年間合計ファイル数: "
        + str(len(arrStep0013GroupDeductionPaths))
    )
    print("step0013グループ別控除廃棄年間Excelは作成しません。")
    print(
        "step0014グループ別控除廃棄年間ファイル数: "
        + str(len(arrStep0014GroupDeductionPaths))
    )
    print("step0014グループ別控除廃棄年間Excelは作成しません。")
    print(
        "step0015グループ別控除廃棄年間ファイル数: "
        + str(len(arrStep0015GroupDeductionPaths))
    )
    print("step0017グループ別控除廃棄年間ファイル数: " + str(len(arrStep0017GroupDeductionPaths)))
    for pszPath in arrStep0017GroupDeductionPaths:
        print("step0017グループ別控除廃棄年間ファイル: " + pszPath)
    print("step0017グループ別控除廃棄年間Excelファイル数: " + str(len(arrStep0017GroupDeductionExcelPaths)))
    for pszPath in arrStep0017GroupDeductionExcelPaths:
        print("step0017グループ別控除廃棄年間Excelファイル: " + pszPath)
    print("年付きグループ別控除廃棄集計Excelファイル数: " + str(len(arrNamedGroupDeductionSummaryExcelPaths)))
    for pszPath in arrNamedGroupDeductionSummaryExcelPaths:
        print("年付きグループ別控除廃棄集計Excelファイル: " + pszPath)
    print("step0016グループ別控除廃棄年間ファイル数: " + str(len(arrStep0016GroupDeductionPaths)))
    for pszPath in arrStep0016GroupDeductionPaths:
        print("step0016グループ別控除廃棄年間ファイル: " + pszPath)
    print(
        "step0005グループ別廃棄羽数ファイル数: "
        + str(len(arrStep0005GroupDiscardCountPaths))
    )
    print(
        "step0006グループ別廃棄羽数ファイル数: "
        + str(len(arrStep0006GroupDiscardCountPaths))
    )
    print(
        "step0007グループ別全廃棄羽数ファイル数: "
        + str(len(arrStep0007GroupTotalDiscardCountPaths))
    )
    print(
        "step0005グループ別廃棄重量ファイル数: "
        + str(len(arrStep0005GroupDiscardWeightPaths))
    )
    print(
        "step0006グループ別廃棄重量ファイル数: "
        + str(len(arrStep0006GroupDiscardWeightPaths))
    )
    print(
        "step0007グループ別全廃棄重量ファイル数: "
        + str(len(arrStep0007GroupTotalDiscardWeightPaths))
    )
    print(
        "step0004総計懸鳥前羽数ファイル数: "
        + str(len(arrStep0004TotalPreSlaughterCountPaths))
    )
    print(
        "step0004総計懸鳥前重量ファイル数: "
        + str(len(arrStep0004TotalPreSlaughterWeightPaths))
    )
    print("step0005総計懸鳥前羽数ファイル数: " + str(len(arrStep0005PreSlaughterCountPaths)))
    print("step0005総計懸鳥前重量ファイル数: " + str(len(arrStep0005PreSlaughterWeightPaths)))
    print("step0005総計廃棄羽数ファイル数: " + str(len(arrStep0005DiscardCountPaths)))
    print("step0005総計廃棄重量ファイル数: " + str(len(arrStep0005DiscardWeightPaths)))
    print("step0006総計廃棄データ一覧表上段ファイル数: " + str(len(arrStep0006DiscardSummaryUpperPaths)))
    print("step0006総計廃棄データ一覧表下段ファイル数: " + str(len(arrStep0006DiscardSummaryLowerPaths)))
    print("step0007総計廃棄データ一覧表上段ファイル数: " + str(len(arrStep0007DiscardSummaryUpperPaths)))
    print("step0007総計廃棄データ一覧表下段ファイル数: " + str(len(arrStep0007DiscardSummaryLowerPaths)))
    print("step0008総計廃棄データ一覧表ファイル数: " + str(len(arrStep0008DiscardSummaryPaths)))
    print("step0009総計廃棄データ一覧表ファイル数: " + str(len(arrStep0009DiscardSummaryPaths)))
    print("step0010総計廃棄データ一覧表ファイル数: " + str(len(arrStep0010DiscardSummaryPaths)))
    print("step0011総計廃棄データ一覧表ファイル数: " + str(len(arrStep0011DiscardSummaryPaths)))
    print("step0012総計廃棄データ一覧表ファイル数: " + str(len(arrStep0012DiscardSummaryPaths)))
    print(
        "step0012総計廃棄データ一覧表Excelファイル数: "
        + str(len(arrStep0012DiscardSummaryExcelPaths))
    )
    print("年月付き廃棄データ一覧表Excelファイル数: " + str(len(arrNamedDiscardSummaryExcelPaths)))
    for pszPath in arrNamedDiscardSummaryExcelPaths:
        print("年月付き廃棄データ一覧表Excelファイル: " + pszPath)
    if arrStep0003YearMonths:
        print(
            "step0003最初の年月: "
            + f"{arrStep0003YearMonths[0][0]:04d}_{arrStep0003YearMonths[0][1]:02d}"
        )
        print(
            "step0003最後の年月: "
            + f"{arrStep0003YearMonths[-1][0]:04d}_{arrStep0003YearMonths[-1][1]:02d}"
        )
    print("step0004対象項目: 入雛羽数, 出荷羽数, 正味重量, 平均体重")
    print("step0004グループ数: " + str(len(GROUPS)))
    print("step0004ファイル数: " + str(len(arrStep0004Paths)))
    if arrStep0004YearMonths:
        print(
            "step0004最初の年月: "
            + f"{arrStep0004YearMonths[0][0]:04d}_{arrStep0004YearMonths[0][1]:02d}"
        )
        print(
            "step0004最後の年月: "
            + f"{arrStep0004YearMonths[-1][0]:04d}_{arrStep0004YearMonths[-1][1]:02d}"
        )
    print("step0005対象年数: " + str(len(arrStep0005Years)))
    print("step0005グループ数: " + str(len(GROUPS)))
    print("step0005ファイル数: " + str(len(arrStep0005Paths)))
    if arrStep0005Years:
        print("step0005最初の年: " + str(arrStep0005Years[0]))
        print("step0005最後の年: " + str(arrStep0005Years[-1]))
    print("step0006入力ファイル数: " + str(len(arrStep0005Paths)))
    print("step0006ファイル数: " + str(len(arrStep0006Paths)))
    print("step0007対象年数: " + str(len(arrStep0007Years)))
    print("step0007ファイル数: " + str(len(arrStep0007Paths)))
    if arrStep0007Years:
        print("step0007最初の年: " + str(arrStep0007Years[0]))
        print("step0007最後の年: " + str(arrStep0007Years[-1]))
    print("step0007 Excel入力ファイル数: " + str(len(arrStep0007Paths)))
    print("step0007 Excelファイル数: " + str(len(arrStep0007ExcelPaths)))
    print("step0008入力ファイル数: " + str(len(arrStep0007Paths)))
    print("step0008ファイル数: " + str(len(arrStep0008Paths)))
    print("step0008グループ別年間Excelは作成しません。")
    print("step0009入力ファイル数: " + str(len(arrStep0008Paths)))
    print("step0009ファイル数: " + str(len(arrStep0009Paths)))
    print("step0009グループ別Excelは作成しません。")
    print("step0010入力ファイル数: " + str(len(arrStep0009Paths)))
    print("step0010ファイル数: " + str(len(arrStep0010Paths)))
    print("step0013グループ別Excel入力ファイル数: " + str(len(arrStep0013GroupPaths)))
    print("step0013グループ別Excelファイル数: " + str(len(arrStep0013GroupExcelPaths)))
    for pszPath in arrStep0013GroupExcelPaths:
        print("step0013グループ別Excelファイル: " + pszPath)
    print("年付きグループ別集計Excelファイル数: " + str(len(arrNamedGroupSummaryExcelPaths)))
    for pszPath in arrNamedGroupSummaryExcelPaths:
        print("年付きグループ別集計Excelファイル: " + pszPath)
    print("step0011グループ別年間ファイル数: " + str(len(arrStep0011GroupPaths)))
    print("step0013グループ別年間ファイル数: " + str(len(arrStep0013GroupPaths)))
    for pszPath in arrStep0013GroupPaths:
        print("step0013グループ別年間ファイル: " + pszPath)
    print("step0012グループ別年間ファイル数: " + str(len(arrStep0012GroupPaths)))
    for pszPath in arrStep0012GroupPaths:
        print("step0012グループ別年間ファイル: " + pszPath)
    print("警告件数: " + str(iWarningCount))
    if pszWarningPath:
        print("警告ファイル: " + pszWarningPath)
    print("◇ 変換・年月別分割・グループ別集計・43項目抽出・グループ別4項目結合・グループ別年間連結・年間行列入れ替え・年間グループ連結・育成率追加・控除廃棄品種別step0008・step0009・step0010年間TSV作成・step0010年間Excel作成・品種別step0009・step0010年間TSV・step0010年間Excel作成・控除廃棄品種別step0010年間TSV・Excel作成・年間合計追加・年間合計列追加・step0013グループ別年間Excel作成・step0011・step0013グループ別年間TSV・step0013グループ別控除廃棄年間TSV・step0014グループ別控除廃棄年間TSV・step0017グループ別控除廃棄年間Excel・step0015グループ別控除廃棄年間TSV作成完了 ◇")
    print("入力Excel: " + pszInputExcelPath)
    print("出力フォルダ: " + os.path.dirname(pszOutputTsvPath))
    return 0


if __name__ == "__main__":
    sys.exit(main())
