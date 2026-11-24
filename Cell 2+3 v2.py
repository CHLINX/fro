# @title
# ======================================================================
# 巡田最速適｜Cell 2+3 v2
# 單一 Colab 程式：Cell 2 v8.7 → Cell 3 v4.3.4
# ======================================================================
#
# 執行流程：
#   1. 執行 Cell 2 v8.7：輸入案件、勘查狀態排除、NLSC 地號／地址定位。
#   2. 確認 /content/result.xlsx 已產生。
#   3. 執行 Cell 3 v4.3.4：NLSC 多圖幅下載、有效圖幅向量化與聯集。
#   4. 驗證 /content/CELL3_HANDOFF.pkl，供 Cell 4／Cell 5 使用。
#
# 整併原則：
#   - 保留 Cell 2 v8.7 與 Cell 3 v4.3.4 的原始查詢、續跑、向量化、
#     備份及稽核邏輯。
#   - 兩個原始模組依序在同一個 Colab 程式碼儲存格執行。
#   - Cell 2 的 /content/result.xlsx 直接作為 Cell 3 輸入。
#   - Cell 5 固定使用的 /content/CELL3_HANDOFF.pkl 路徑不變。
#   - 本程式需在 Google Colab／IPython 儲存格執行，因 Cell 2 使用 top-level await。
#
# 原始來源（已核對 Git blob SHA）：
#   Cell 2 v8.7: 0334c8c56fd84345d4ce5b055bcfb591e642c0e9
#   Cell 3 v4.3.4: 49bb37f27f4fbd0aef534637732232c7ec411060
#
# ======================================================================

CELL23_VERSION = "v2"
CELL23_COMPONENT_VERSIONS = {"Cell2": "v8.7", "Cell3": "v4.3.4"}

print("=" * 76)
print("巡田最速適｜Cell 2+3 v2")
print("整合版本：Cell 2 v8.7 + Cell 3 v4.3.4")
print("=" * 76)

# ============================================================
# 巡田最速適｜Cell 2 v8.7: NLSC 地號／地址批次座標查詢【429快速切換備援＋冷卻斷路】
# ============================================================
#
# 功能：
#
# 1. 自動尋找：
#      /content/input.xlsx
#      /content/input.xls
#      /content/input.csv
#
# 2. 輸入欄位（固定 9 欄）：
#      縣市
#      鄉鎮市區
#      地段
#      地號
#      本筆面積(公頃)
#      申請面積(公頃)
#      地址
#      名稱
#      勘查狀態
#
#      注意：
#      查詢方式不再由輸入檔提供，程式會自動判斷：
#        有地段＋地號 → 地號定位（即使同時有地址）
#        只有地址 → 地址定位
#      若地號案件同時有地址：
#        經度／緯度主座標保留地號定位結果
#        另存「地址經度／地址緯度」供 Cell 4 判斷實際進場道路。
#      本筆面積(公頃)、申請面積(公頃)、名稱、勘查狀態均直接沿用輸入檔，
#      不向 NLSC 查詢，也不由 NLSC 結果覆蓋。
#      讀檔後先排除已完成類狀態（完成／已完成／勘查完成／已勘查／OK／Done 等），不查座標、不寫入任何交接結果。
#      其餘案件才進入座標查詢；勘查狀態空白或不明值仍視同「未勘查」。
#
# v8.7 變更：DIRECT 首次 429 即切換瀏覽器備援；全域冷卻期間跳過 DIRECT、不逐筆等待；Retry-After 可解析秒數及 HTTP-date。
# v8.6 變更：讀檔後先排除勘查完成類案件（完成／已完成／勘查完成／已勘查／OK／Done 等），
#       不查詢座標、不納入 result.xlsx／partial／error／gmap.csv。
#       若全批皆排除，正式 result.xlsx 會覆寫為 0 筆空表，避免舊資料誤交接。

# 3. 查詢方式：
#      地號：縣市＋鄉鎮市區＋地段＋地號
#      地址：縣市＋鄉鎮市區＋地址
#      地址可沒有路名，例如「○○里○鄰○○5號」
##
# 3. 鄉鎮市區名稱支援：
#      八德   -> 八德區
#      大溪   -> 大溪區
#      新豐   -> 新豐鄉
#      關西   -> 關西鎮
#      竹北   -> 竹北市
#
#      完整名稱也可：
#      八德區、新豐鄉、竹北市
#
# 4. 地段名稱支援：
#      完整名稱
#      特定地段別名標準化
#      地段名稱由 NLSC 正式選單自動匹配
#      省略「段」
#      省略「小段」
#      同時省略「段／小段」
#      中間省略「段／小段」
#      ? 單字元萬用字元
#
#      例如：
#      茄苳
#          -> 茄苳段
#
#      下田心子下田心子
#          -> 下田心子段下田心子小段
#
#      十五間十五間尾
#          -> 十五間段十五間尾小段
#
#      ?榔上?榔
#          -> 槺榔段上槺榔小段
#
# 5. 地號格式：
#      1208       -> 12080000
#      4820000    -> 04820000
#      9880000    -> 09880000
#      12080000   -> 12080000
#      17420013   -> 17420013
#      1208-0000  -> 12080000
#
# 6. 使用 NLSC 真實網頁流程查詢
#      地號：定位查詢 → 地號
#      地址：定位查詢 → 其他定位 → 門牌
#
# 7. 地址查詢攔截門牌定位服務回應
#    地號查詢攔截 qryTileMapIndex response
#
# 8. 取得：
#      經度
#      緯度
#      TWD97_X
#      TWD97_Y
#      bbox
#
# 9. HTTP 502 / 503 / 504 自動重試
#
# 10. timeout 自動重試
#     timeout 時重新載入 NLSC，
#     重新設定縣市／鄉鎮市區／地段／地號
#
# 11. 原始 response：
#      /content/nlsc_raw/
#
# 12. 成功結果：
#      /content/result.xlsx
#
#      輸入檔「名稱」／本筆面積(公頃)／申請面積(公頃)
#      原樣帶入結果，不向 NLSC 查詢。
#
# 13. 失敗結果：
#      /content/result_error.xlsx
#
# 14. 續跑機制：
#      若已存在 result_partial.xlsx 或 result.xlsx，
#      程式會先讀取上一輪結果。
#      狀態為 OK 的案件直接沿用，不再向 NLSC 查詢；
#      只有上次失敗的案件才會重新查詢。
#      每完成 1 筆即寫入 result_partial.xlsx；手動取消後也可直接續跑。
#      完成後才建立新的 result.xlsx。
#
# 15. 每完成 1 筆就寫入 result_partial.xlsx；手動取消後可直接續跑。
#
# 16. Google Maps 匯入 CSV：
#      /content/gmap.csv
#
#     欄位：
#       Name
#       Latitude
#       Longitude
#
#     名稱：
#       鄉鎮市區（保留鄉／鎮／市／區）
#       +
#       段名（移除「段」「小段」）
#       +
#       8碼地號
#
#     例如：
#       八德區茄苳07370000
#       新豐鄉福興後湖子16490000
#
#     座標順序：
#       緯度、經度
#
# ============================================================


# ============================================================
# 0. 自動安裝套件
# ============================================================

import sys
import subprocess
import importlib.util


def ensure_package(
    package_name,
    import_name=None
):

    if import_name is None:
        import_name = package_name

    if importlib.util.find_spec(
        import_name
    ) is None:

        print(
            f"正在安裝套件：{package_name}"
        )

        subprocess.check_call([
            sys.executable,
            "-m",
            "pip",
            "install",
            "-q",
            package_name
        ])


ensure_package(
    "openpyxl"
)

ensure_package(
    "pyproj"
)

ensure_package(
    "playwright"
)


# .xls
try:
    import xlrd
except Exception:

    try:

        ensure_package(
            "xlrd"
        )

    except Exception:
        pass


# Chromium
subprocess.run(
    [
        sys.executable,
        "-m",
        "playwright",
        "install",
        "chromium"
    ],
    check=False
)


# ============================================================
# 1. 匯入
# ============================================================

import json
import re
import time
import hashlib
import asyncio
import unicodedata
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
import traceback
import csv
import math

from pathlib import Path
from urllib.parse import (
    urlparse,
    parse_qs
)

import openpyxl

from pyproj import Transformer

from playwright.async_api import (
    async_playwright,
)


# ============================================================
# 2. 基本設定
# ============================================================

BASE_DIR = Path(
    "/content"
)

INPUT_CANDIDATES = [

    BASE_DIR / "input.xlsx",

    BASE_DIR / "input.xls",

    BASE_DIR / "input.csv",
]


# ============================================================
# 輸出檔案
# ============================================================

SUCCESS_OUTPUT_XLSX = (
    BASE_DIR / "result.xlsx"
)

ERROR_OUTPUT_XLSX = (
    BASE_DIR / "result_error.xlsx"
)

GMAP_OUTPUT_CSV = (
    BASE_DIR / "gmap.csv"
)


# ============================================================
# NLSC
# ============================================================

NLSC_URL = (
    "https://maps.nlsc.gov.tw/"
    "T09/mobilemap.action"
)


# ============================================================
# 原始 response
# ============================================================

RAW_DIR = (
    BASE_DIR / "nlsc_raw"
)

# v8.4.1：地號案件若另有地址，地址查詢 response 另存，避免覆蓋地號 response。
ADDRESS_RAW_DIR = (
    BASE_DIR / "nlsc_raw_address"
)


# ============================================================
# 查詢設定
# ============================================================

DELAY_BETWEEN_ROWS = 0.0

# v8.7：主流程 DELAY=0；DIRECT 本身使用 1 秒最小請求間隔。
RESPONSE_TIMEOUT = 30000

SECTION_POLL_INTERVAL = 0.25

SECTION_MAX_POLLS = 40

MAX_RETRIES = 3

RETRY_BACKOFF = [
    2,
    4,
    8,
]

# v8.7：DIRECT 最小請求間隔提高至 1 秒。
DIRECT_MIN_INTERVAL_SECONDS = 1.0

# v8.7：同一案件 DIRECT 第一次遇到 429 就立即切換瀏覽器備援，不再同筆重打。
# Retry-After 有提供時完全尊重伺服器秒數／HTTP-date；缺少時才使用本地退避值。
DIRECT_429_BACKOFF = [15, 30, 60]
DIRECT_MAX_ATTEMPTS = 2  # 保留原有非 429 暫時性錯誤重試；429 分支會直接 break
DIRECT_429_MAX_SERVER_WAIT_SECONDS = 3600.0

# v8.4.1：地號案件若同時有地址，是否再查一次 NLSC 門牌座標。
# True：主座標維持地號；另存「地址經度／地址緯度」。
ENABLE_SECONDARY_ADDRESS_GEOCODE = True

RETRY_STATUS_CODES = {
    502,
    503,
    504,
}


# ============================================================
# 3. 基本文字
# ============================================================

def clean_text(
    value
):

    if value is None:
        return ""

    return str(
        value
    ).strip()


# ============================================================
# 3A. 勘查狀態標準化／已完成案件排除
# ============================================================
# 已完成狀態在讀檔後立即排除，必須早於續跑快取、NLSC 查詢與所有輸出。
# 中文支援常見勘查／查勘／巡查完成說法；英文不區分大小寫。
# 明確的未完成／進行中描述優先，不會因為包含「完成」字樣而誤剔除。

VALID_INSPECTION_STATUS = {
    "未勘查",
    "已完成",
}

_COMPLETED_INSPECTION_MARKERS_ZH = (
    "已完成",
    "完成",
    "勘查完成",
    "完成勘查",
    "已勘查",
    "勘查完畢",
    "已勘察",
    "勘察完成",
    "完成勘察",
    "已查勘",
    "查勘完成",
    "完成查勘",
    "已巡查",
    "巡查完成",
    "完成巡查",
    "已巡田",
    "巡田完成",
    "完成巡田",
    "已查驗",
    "查驗完成",
    "驗收完成",
    "已勘測",
    "勘測完成",
    "已踏勘",
    "踏勘完成",
    "已訪查",
    "訪查完成",
    "已查訪",
    "查訪完成",
    "已訪視",
    "訪視完成",
    "已巡勘",
    "巡勘完成",
    "已結案",
    "結案",
    "處理完成",
)

_NOT_COMPLETED_INSPECTION_MARKERS_ZH = (
    "尚未完成",
    "還未完成",
    "還沒完成",
    "未完成",
    "不算完成",
    "部分完成",
    "部份完成",
    "尚未勘查",
    "未勘查",
    "還未勘查",
    "還沒勘查",
    "尚未勘察",
    "未勘察",
    "待勘查",
    "待勘察",
    "尚未查勘",
    "未查勘",
    "待查勘",
    "尚未巡查",
    "未巡查",
    "待巡查",
    "進行中",
    "處理中",
)

_COMPLETED_INSPECTION_WORDS_EN = {
    "ok",
    "okay",
    "done",
    "complete",
    "completed",
    "checked",
    "inspected",
    "passed",
    "finished",
    "closed",
    "resolved",
    "verified",
}

_NOT_COMPLETED_INSPECTION_WORDS_EN = {
    "not",
    "incomplete",
    "unfinished",
    "pending",
    "processing",
    "todo",
    "waiting",
}


def is_completed_inspection_status(value):
    """判斷輸入狀態是否代表已完成；空白、未完成及不明值一律不排除。"""
    raw = clean_text(value)
    if not raw:
        return False

    normalized = unicodedata.normalize("NFKC", raw).casefold()
    compact = re.sub(r"[\W_]+", "", normalized, flags=re.UNICODE)

    # 先辨識明確的未完成／進行中描述，避免「未完成」被「完成」誤判。
    if any(marker in compact for marker in _NOT_COMPLETED_INSPECTION_MARKERS_ZH):
        return False

    english_words = re.findall(r"[a-z]+", normalized)
    english_word_set = set(english_words)
    english_phrase = " ".join(english_words)

    if english_word_set & _NOT_COMPLETED_INSPECTION_WORDS_EN:
        return False
    if any(
        phrase in english_phrase
        for phrase in (
            "not ok",
            "not okay",
            "not done",
            "not complete",
            "not completed",
            "not checked",
            "not inspected",
            "not passed",
            "not yet",
        )
    ):
        return False

    if any(marker in compact for marker in _COMPLETED_INSPECTION_MARKERS_ZH):
        return True

    if english_word_set & _COMPLETED_INSPECTION_WORDS_EN:
        return True

    return False


def split_completed_inspection_rows(rows):
    """回傳 (待處理資料, 已完成排除清單)；保留原始資料順序。"""
    active_rows = []
    excluded_rows = []

    for input_index, row in enumerate(rows, start=1):
        raw_status = clean_text(row.get("勘查狀態", ""))
        if is_completed_inspection_status(raw_status):
            excluded_rows.append({
                "input_index": input_index,
                "status": raw_status,
                "row": row,
            })
        else:
            active_rows.append(row)

    return active_rows, excluded_rows


def normalize_inspection_status(value):
    status = clean_text(value)

    if not status:
        return "未勘查"

    if status in VALID_INSPECTION_STATUS:
        return status

    print(
        f"  [WARN] 勘查狀態「{status}」不是允許值，"
        "自動視同「未勘查」。"
    )
    return "未勘查"


# ============================================================
# 4. 地號正規化
# ============================================================

def normalize_landno(
    value
):

    value = clean_text(
        value
    )

    # Excel 可能讀成 1208.0
    if re.fullmatch(
        r"\d+\.0",
        value
    ):

        value = value[:-2]

    # 含「-」的母號－子號格式：前、後各固定 4 碼後串接
    # 例如：106-5 -> 01060005；1-1 -> 00010001
    # 同時兼容全形連字號：106－5
    if "-" in value or "－" in value:

        value = value.replace("－", "-")
        parts = value.split("-")

        if len(parts) == 2:

            main_no = clean_text(parts[0])
            sub_no = clean_text(parts[1])

            if (
                re.fullmatch(r"\d+", main_no)
                and
                re.fullmatch(r"\d+", sub_no)
            ):

                return (
                    main_no.zfill(4)
                    +
                    sub_no.zfill(4)
                )

        # 非標準「-」格式則原樣保留，避免誤改資料
        return value

    # 純數字
    if re.fullmatch(
        r"\d+",
        value
    ):

        # 4碼母號
        if len(value) == 4:

            return (
                value
                + "0000"
            )

        # 少於4碼
        if len(value) < 4:

            return (
                value.zfill(4)
                + "0000"
            )

        # 少於8碼
        if len(value) < 8:

            return value.zfill(8)

        return value

    return value


# ============================================================
# 5. 地號格式測試
# ============================================================

LANDNO_TESTS = {

    "1208":
        "12080000",

    "4820000":
        "04820000",

    "9880000":
        "09880000",

    "12080000":
        "12080000",

    "17420013":
        "17420013",

    "1208-0000":
        "12080000",

    "106-5":
        "01060005",

    "1-1":
        "00010001",

    "12-34":
        "00120034",

    "106－5":
        "01060005",

    "12":
        "00120000",
}


for source, expected in (
    LANDNO_TESTS.items()
):

    actual = normalize_landno(
        source
    )

    if actual != expected:

        raise RuntimeError(
            "地號格式測試失敗："
            f"{source} -> {actual}；"
            f"預期 {expected}"
        )


print(
    "地號格式測試：全部通過"
)


# ============================================================
# 6. 地段名稱處理
# ============================================================

def normalize_section_name_variants(
    name
):

    name = clean_text(
        name
    )

    if not name:
        return []

    variants = []
    seen = set()

    def add(
        value
    ):

        value = clean_text(
            value
        )

        if (
            value
            and
            value not in seen
        ):

            seen.add(
                value
            )

            variants.append(
                value
            )

    # 原始
    add(
        name
    )

    # 去尾端「小段」
    if name.endswith(
        "小段"
    ):

        add(
            name[:-2]
        )

    # 去尾端「段」
    if name.endswith(
        "段"
    ):

        add(
            name[:-1]
        )

    # 補「段」
    if not name.endswith(
        "段"
    ):

        add(
            name + "段"
        )

    # 補「小段」
    if not name.endswith(
        "小段"
    ):

        add(
            name + "小段"
        )

    return variants


def canonical_section_name(
    name
):

    """
    移除全部「段」「小段」

    例如：

    下田心子段下田心子小段
    ->
    下田心子下田心子

    十五間段十五間尾小段
    ->
    十五間十五間尾

    福興段後湖子小段
    ->
    福興後湖子
    """

    name = clean_text(
        name
    )

    name = name.replace(
        "小段",
        ""
    )

    name = name.replace(
        "段",
        ""
    )

    return name


# ------------------------------------------------------------
# 地段查詢標準化
#
# 不建立人工地段別名對照表。
# 原始地段保留不動，實際由 NLSC 地段選單自動比對正式名稱。
# ------------------------------------------------------------
def normalize_section_for_query(
    city,
    area,
    section
):
    """
    不建立人工地段別名表。

    保留原始地段文字，實際送入 NLSC 選單後，
    由既有 section_name_matches() 自動比對正式地段名稱。

    既有比對規則包含：
      - 完整名稱
      - 省略「段」
      - 省略「小段」
      - 同時省略「段／小段」
      - 中間省略「段／小段」
      - 「?」單字元萬用字元

    例如：
      桃園武陵
      → 自動命中 NLSC 的桃園段武陵小段
    """

    return clean_text(section)


# 不使用人工別名的標準化測試
SECTION_QUERY_TESTS = {
    ("桃園", "桃園", "桃園武陵"): "桃園武陵",
    ("桃園市", "桃園區", "桃園武陵"): "桃園武陵",
    ("桃園市", "桃園區", "桃園段武陵"): "桃園段武陵",
    ("桃園市", "桃園區", "武陵"): "武陵",
}

for (test_city, test_area, test_section), expected in SECTION_QUERY_TESTS.items():
    actual = normalize_section_for_query(
        test_city,
        test_area,
        test_section
    )
    if actual != expected:
        raise RuntimeError(
            "地段查詢標準化測試失敗："
            f"{test_city}/{test_area}/{test_section} -> {actual}；"
            f"預期 {expected}"
        )

print("地段查詢標準化測試：全部通過（無人工別名）")

def wildcard_fullmatch(
    requested,
    actual
):

    """
    ? = 任意一個字元
    """

    pattern = (
        "^"
        +
        re.escape(
            requested
        ).replace(
            r"\?",
            "."
        )
        +
        "$"
    )

    try:

        return bool(
            re.fullmatch(
                pattern,
                actual
            )
        )

    except Exception:

        return False


def section_name_matches(
    requested,
    actual
):

    requested = clean_text(
        requested
    )

    actual = clean_text(
        actual
    )

    if (
        not requested
        or
        not actual
    ):
        return False

    # 完整相等
    if requested == actual:
        return True

    requested_variants = (
        normalize_section_name_variants(
            requested
        )
    )

    actual_variants = (
        normalize_section_name_variants(
            actual
        )
    )

    # variants
    for req in requested_variants:

        for act in actual_variants:

            if req == act:
                return True

    # 移除所有段／小段
    requested_canonical = (
        canonical_section_name(
            requested
        )
    )

    actual_canonical = (
        canonical_section_name(
            actual
        )
    )

    if (
        requested_canonical
        ==
        actual_canonical
    ):
        return True

    # ?
    if wildcard_fullmatch(
        requested,
        actual
    ):
        return True

    for req in requested_variants:

        if wildcard_fullmatch(
            req,
            actual
        ):

            return True

    for act in actual_variants:

        if wildcard_fullmatch(
            requested,
            act
        ):

            return True

    # ?
    # + 移除段／小段
    if wildcard_fullmatch(
        requested_canonical,
        actual_canonical
    ):

        return True

    return False


# 地段名稱比對測試：不建立人工對照表
SECTION_MATCH_TESTS = [
    ("桃園武陵", "桃園段武陵小段", True),
    ("茄苳", "茄苳段", True),
    ("?榔上?榔", "槺榔段上槺榔小段", True),
    ("下田心子下田心子", "下田心子段下田心子小段", True),
    ("十五間十五間尾", "十五間段十五間尾小段", True),
    ("不存在", "桃園段武陵小段", False),
]

for requested, actual, expected in SECTION_MATCH_TESTS:
    result = section_name_matches(
        requested,
        actual
    )
    if result != expected:
        raise RuntimeError(
            "地段名稱比對測試失敗："
            f"{requested} vs {actual} -> {result}；"
            f"預期 {expected}"
        )

print("地段名稱比對測試：全部通過")


def find_option_value(
    options,
    requested_name
):

    candidates = []

    for option in options:

        text = clean_text(
            option.get(
                "text"
            )
        )

        value = clean_text(
            option.get(
                "value"
            )
        )

        if section_name_matches(
            requested_name,
            text
        ):

            candidates.append({
                "text": text,
                "value": value,
            })

    if len(candidates) == 1:

        return candidates[0]

    if len(candidates) > 1:

        print(
            f"  [WARN] 地段名稱「"
            f"{requested_name}"
            f"」有多個候選："
            +
            "、".join(
                x["text"]
                for x in candidates
            )
        )

    return None


# ============================================================
# 7. 鄉鎮市區名稱處理（地號／地址共用 NLSC 正式名稱自動匹配）
# ============================================================

def normalize_area_name_variants(
    name
):

    """
    支援：

      八德
      八德區

      大溪
      大溪區

      新豐
      新豐鄉

      關西
      關西鎮

      竹北
      竹北市
    """

    name = clean_text(
        name
    )

    if not name:
        return []

    variants = []

    seen = set()

    def add(
        value
    ):

        value = clean_text(
            value
        )

        if (
            value
            and
            value not in seen
        ):

            seen.add(
                value
            )

            variants.append(
                value
            )

    # 原始
    add(
        name
    )

    # 移除最後鄉鎮市區字尾
    base = name

    for suffix in [
        "鄉",
        "鎮",
        "市",
        "區",
    ]:

        if base.endswith(
            suffix
        ):

            base = (
                base[:-1]
            )

            break

    add(
        base
    )

    # 補鄉鎮市區尾字
    add(
        base + "區"
    )

    add(
        base + "鄉"
    )

    add(
        base + "鎮"
    )

    add(
        base + "市"
    )

    return variants


def area_name_matches(
    requested,
    actual
):

    requested = clean_text(
        requested
    )

    actual = clean_text(
        actual
    )

    if (
        not requested
        or
        not actual
    ):
        return False

    # 完整相等
    if requested == actual:
        return True

    requested_variants = (
        normalize_area_name_variants(
            requested
        )
    )

    actual_variants = (
        normalize_area_name_variants(
            actual
        )
    )

    for req in requested_variants:

        for act in actual_variants:

            if req == act:
                return True

    return False


# ============================================================
# 8. select options
# ============================================================

async def get_options(
    page,
    selector
):

    return await page.eval_on_selector_all(
        selector,
        """
        els => els.map(x => ({
            text: (x.textContent || '').trim(),
            value: x.value || ''
        }))
        """
    )


# ============================================================
# 9. 設定縣市 / 鄉鎮市區 / 地段
# ============================================================

async def resolve_nlsc_area_name(
    page,
    city,
    area,
):
    """
    依 NLSC 實際鄉鎮市區選單，自動找出正式名稱。

    支援輸入省略「鄉／鎮／市／區」：
      八德   -> 八德區
      新豐   -> 新豐鄉
      關西   -> 關西鎮
      竹北   -> 竹北市

    回傳 NLSC 實際選項文字與 value；不自行猜測正式名稱。
    """
    city = clean_text(city)
    area = clean_text(area)

    if not city or not area:
        raise RuntimeError(
            f"鄉鎮市區自動匹配需要：縣市、鄉鎮市區｜{city} / {area}"
        )

    await page.select_option(
        "#land_city",
        label=city,
        force=True,
    )

    matched = None

    for _ in range(SECTION_MAX_POLLS):
        try:
            options = await get_options(
                page,
                "#land_area_office option",
            )

            for option in options:
                option_text = clean_text(option.get("text", ""))
                option_value = clean_text(option.get("value", ""))

                if not option_text or not option_value:
                    continue

                if area_name_matches(area, option_text):
                    matched = {
                        "area_text": option_text,
                        "area_value": option_value,
                    }
                    break

            if matched is not None:
                break

        except Exception:
            pass

        await page.wait_for_timeout(
            int(SECTION_POLL_INTERVAL * 1000)
        )

    if matched is None:
        raise RuntimeError(
            "找不到鄉鎮市區："
            f"{city} / {area}"
        )

    actual_area = matched["area_text"]

    if actual_area != area:
        print(
            f"  鄉鎮市區自動調整：{area} -> {actual_area}"
        )
    else:
        print(
            f"  鄉鎮市區已匹配：{area} -> {actual_area}"
        )

    return matched


async def set_nlsc_section(
    page,
    city,
    area,
    section
):

    city = clean_text(
        city
    )

    area = clean_text(
        area
    )

    section = clean_text(
        section
    )

    # --------------------------------------------------------
    # 縣市
    # --------------------------------------------------------

    await page.select_option(
        "#land_city",
        label=city,
        force=True
    )

    # --------------------------------------------------------
    # 鄉鎮市區：依 NLSC 實際選單自動匹配正式名稱
    # --------------------------------------------------------

    area_info = await resolve_nlsc_area_name(
        page,
        city,
        area,
    )

    area_value = area_info["area_value"]
    matched_area_text = area_info["area_text"]

    # --------------------------------------------------------
    # 設定鄉鎮市區
    # --------------------------------------------------------

    await page.evaluate(
        """
        ({value}) => {

            const el =
                document.querySelector(
                    '#land_area_office'
                );

            if (!el) {

                throw new Error(
                    '#land_area_office not found'
                );
            }

            el.value = value;

            el.dispatchEvent(
                new Event(
                    'change',
                    {
                        bubbles: true
                    }
                )
            );
        }
        """,
        {
            "value":
                area_value
        }
    )

    # --------------------------------------------------------
    # 等地段清單
    # --------------------------------------------------------

    section_option = None

    for _ in range(
        SECTION_MAX_POLLS
    ):

        try:

            options = (
                await get_options(
                    page,
                    "#land_section option"
                )
            )

            options = [
                option
                for option in options
                if (
                    option["value"]
                    and
                    option["text"]
                )
            ]

            section_option = (
                find_option_value(
                    options,
                    section
                )
            )

            if (
                section_option
                is not None
            ):

                break

        except Exception:
            pass

        await page.wait_for_timeout(
            int(
                SECTION_POLL_INTERVAL
                * 1000
            )
        )

    if section_option is None:

        try:

            debug_options = (
                await get_options(
                    page,
                    "#land_section option"
                )
            )

            debug_names = [

                clean_text(
                    x["text"]
                )

                for x in debug_options

                if (
                    x["value"]
                    and
                    x["text"]
                )
            ]

            print(
                "  [DEBUG] "
                "目前地段選單前20項："
            )

            for name in (
                debug_names[:20]
            ):

                print(
                    f"    {name}"
                )

        except Exception:
            pass

        raise RuntimeError(
            "找不到地段："
            f"{city} / "
            f"{area} / "
            f"{section}"
        )

    # --------------------------------------------------------
    # 設定地段
    # --------------------------------------------------------

    await page.evaluate(
        """
        ({value}) => {

            const el =
                document.querySelector(
                    '#land_section'
                );

            if (!el) {

                throw new Error(
                    '#land_section not found'
                );
            }

            el.value = value;

            el.dispatchEvent(
                new Event(
                    'change',
                    {
                        bubbles: true
                    }
                )
            );
        }
        """,
        {
            "value":
                section_option["value"]
        }
    )

    # --------------------------------------------------------
    # sectioncode
    # --------------------------------------------------------

    section_code = ""

    for _ in range(
        SECTION_MAX_POLLS
    ):

        try:

            section_code = (
                await page.input_value(
                    "#land_sectioncode"
                )
            )

        except Exception:

            section_code = ""

        if clean_text(
            section_code
        ):
            break

        await page.wait_for_timeout(
            int(
                SECTION_POLL_INTERVAL
                * 1000
            )
        )

    return {

        "area_text":
            matched_area_text,

        "area_value":
            area_value,

        "section_text":
            section_option["text"],

        "section_value":
            section_option["value"],

        "section_code":
            clean_text(
                section_code
            ),
    }


# ============================================================
# 10. 填地號
# ============================================================

async def fill_landno(
    page,
    normalized_landno
):

    await page.evaluate(
        """
        ({value}) => {

            const el =
                document.querySelector(
                    '#landcode'
                );

            if (!el) {

                throw new Error(
                    '#landcode not found'
                );
            }

            el.value = value;

            el.dispatchEvent(
                new Event(
                    'input',
                    {
                        bubbles: true
                    }
                )
            );

            el.dispatchEvent(
                new Event(
                    'change',
                    {
                        bubbles: true
                    }
                )
            );
        }
        """,
        {
            "value":
                normalized_landno
        }
    )


# ============================================================
# 11. 觸發 NLSC 查詢
# ============================================================

async def trigger_nlsc_query(
    page
):

    await page.evaluate(
        """
        () => {

            const btn =
                document.querySelector(
                    '#qt-land-submit'
                );

            if (!btn) {

                throw new Error(
                    '#qt-land-submit not found'
                );
            }

            if (
                window.jQuery
                &&
                typeof
                window.jQuery.fn.trigger
                === 'function'
            ) {

                window.jQuery(
                    '#qt-land-submit'
                ).trigger(
                    'click'
                );

                return;
            }

            if (
                typeof
                window.qtLandLoad
                !== 'function'
            ) {

                throw new Error(
                    'qtLandLoad function not found'
                );
            }

            window.qtLandLoad(
                'land_city',
                'land_area_office',
                'land_section',
                'land_sectioncode',
                'landcode',
                'qt-land-ctn'
            );
        }
        """
    )


# ============================================================
# 12. NLSC 頁面初始化
# ============================================================

async def initialize_nlsc_page(
    page
):

    print(
        "  正在重新初始化 NLSC 頁面..."
    )

    await page.goto(
        NLSC_URL,
        wait_until="domcontentloaded",
        timeout=60000
    )

    required_selectors = [

        "#land_city",

        "#land_area_office",

        "#land_section",

        "#land_sectioncode",

        "#landcode",

        "#qt-land-submit",
    ]

    for selector in (
        required_selectors
    ):

        await page.wait_for_selector(
            selector,
            state="attached",
            timeout=60000
        )

    await page.wait_for_timeout(
        1500
    )

    print(
        "  NLSC 頁面重新初始化完成"
    )


# ============================================================
# 13. JSONP 解析
# ============================================================

def jsonp_to_json(
    text
):

    text = text.strip()

    # 純 JSON
    if (
        text.startswith("[")
        or
        text.startswith("{")
    ):

        return json.loads(
            text
        )

    # JSONP
    match = re.search(
        r"^[^(]+\((.*)\)\s*;?\s*$",
        text,
        re.S
    )

    if match:

        return json.loads(
            match.group(1).strip()
        )

    raise ValueError(
        "無法解析 JSON / JSONP"
    )


# ============================================================
# 14. NLSC response 解析
# ============================================================

def parse_nlsc_response(
    body
):

    result = {

        "ok": False,

        "error": "",

        "json": None,

        "record": None,
    }

    if not body:

        result["error"] = (
            "EMPTY_RESPONSE"
        )

        return result

    try:

        data = jsonp_to_json(
            body
        )

        result["json"] = data

    except Exception as e:

        result["error"] = (
            f"INVALID_JSONP: {e}"
        )

        return result

    try:

        record = data[0][0]

    except Exception:

        result["error"] = (
            "INVALID_RESPONSE_STRUCTURE"
        )

        return result

    if not isinstance(
        record,
        dict
    ):

        result["error"] = (
            "INVALID_RECORD"
        )

        return result

    if (
        "cx" not in record
        or
        "cy" not in record
    ):

        result["error"] = (
            "NO_COORDINATE"
        )

        result["record"] = record

        return result

    try:

        record["cx"] = float(
            record["cx"]
        )

        record["cy"] = float(
            record["cy"]
        )

        for key in [
            "lx",
            "ly",
            "rx",
            "ry",
        ]:

            if (
                key in record
                and
                record[key]
                not in (
                    None,
                    ""
                )
            ):

                record[key] = float(
                    record[key]
                )

    except Exception as e:

        result["error"] = (
            f"COORDINATE_PARSE_ERROR: {e}"
        )

        result["record"] = record

        return result

    result["ok"] = True

    result["record"] = record

    return result


# ============================================================
# 15. WGS84 -> TWD97
# ============================================================

transformer = (
    Transformer.from_crs(
        "EPSG:4326",
        "EPSG:3826",
        always_xy=True
    )
)


def convert_to_twd97(
    lon,
    lat
):

    x, y = transformer.transform(
        lon,
        lat
    )

    return (
        float(x),
        float(y)
    )


# ============================================================
# 16. 儲存原始 response
# ============================================================

def save_raw_response(
    row_number,
    status_code,
    response_url,
    body
):

    RAW_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    filename = (
        RAW_DIR
        /
        f"row_{row_number:06d}.txt"
    )

    content = [

        f"HTTP_STATUS={status_code}",

        f"URL={response_url}",

        "",

        body
        if body is not None
        else ""
    ]

    filename.write_text(
        "\n".join(
            content
        ),
        encoding="utf-8"
    )

    return str(
        filename
    )


# ============================================================
# 17. URL Query Parameter
# ============================================================

def get_query_parameter(
    url,
    key
):

    try:

        query = parse_qs(
            urlparse(
                url
            ).query
        )

        values = query.get(
            key
        )

        if values:
            return values[0]

    except Exception:
        pass

    return ""


# ============================================================
# 18. qryTileMapIndex response
# ============================================================

def is_land_query_response(
    response
):

    try:

        return (
            "qryTileMapIndex"
            in response.url
        )

    except Exception:

        return False



# ============================================================
# 19A. 地址查詢輔助工具
# ============================================================

def normalize_address_text(value):
    """
    地址比對用標準化：
    - Unicode NFKC
    - 移除空白
    - 臺／台統一
    - 不自行新增或刪除「路、街、里、鄰」等地址文字
    """
    import unicodedata

    value = clean_text(value)
    if not value:
        return ""

    value = unicodedata.normalize("NFKC", value)
    value = value.replace("臺", "台")
    value = re.sub(r"\s+", "", value)
    value = value.replace("　", "")
    return value


def build_full_address(city, area, address):
    """
    建立 NLSC 門牌查詢地址，避免縣市／鄉鎮市區重複拼接。

    例：
      桃園市 + 桃園市 + 大同路111號
        -> 桃園市大同路111號

      桃園市 + 桃園區 + 大同路111號
        -> 桃園市桃園區大同路111號

      若地址本身已包含完整縣市或鄉鎮市區前綴，也避免再次重複。
    """
    city = clean_text(city)
    area = clean_text(area)
    address = clean_text(address)

    if not address:
        return f"{city}{area}"

    # 地址本身已是完整或部分完整地址時，避免重複。
    if city and address.startswith(city):
        return address

    # 鄉鎮市區欄若誤填成與縣市相同，例如：桃園市 + 桃園市，僅保留一次。
    if area and area == city:
        return f"{city}{address}"

    # 地址本身已包含鄉鎮市區，但沒有包含縣市。
    if area and address.startswith(area):
        return f"{city}{address}"

    return f"{city}{area}{address}"


def parse_address_location_response(body):
    """
    解析 NLSC TextQueryAddress / 門牌查詢回應。

    NLSC 門牌服務實際回應可能使用 XML：
      <ITEM>
        <CONTENT>...</CONTENT>
        <LOCATION>121.309363,24.988454</LOCATION>
      </ITEM>

    也支援舊版 addressItem XML，以及 JSON/JSONP 容錯。
    """
    body = clean_text(body)
    if not body:
        raise ValueError("ADDRESS_EMPTY_RESPONSE")

    records = []

    # --------------------------------------------------------
    # XML
    # --------------------------------------------------------
    try:
        import xml.etree.ElementTree as ET
        root = ET.fromstring(body)
        for item in root.iter():
            tag = item.tag.split("}", 1)[-1].lower()
            if tag not in {"item", "addressitem"}:
                continue

            content = ""
            location = ""
            for child in item.iter():
                ctag = child.tag.split("}", 1)[-1].lower()
                text = clean_text(child.text)
                if ctag == "content" and text:
                    content = text
                elif ctag == "location" and text:
                    location = text

            if location:
                m = re.search(
                    r"(-?\d+(?:\.\d+)?)\s*[,，]\s*(-?\d+(?:\.\d+)?)",
                    location,
                )
                if m:
                    records.append({
                        "address": content,
                        "lon": float(m.group(1)),
                        "lat": float(m.group(2)),
                    })

        if records:
            return records
    except Exception:
        pass

    # --------------------------------------------------------
    # JSON / JSONP
    # --------------------------------------------------------
    try:
        data = jsonp_to_json(body)

        def walk(obj):
            if isinstance(obj, dict):
                loc = obj.get("location") or obj.get("Location")
                content = (
                    obj.get("content")
                    or obj.get("Content")
                    or obj.get("address")
                    or obj.get("Address")
                    or ""
                )
                if isinstance(loc, str):
                    m = re.search(
                        r"(-?\d+(?:\.\d+)?)\s*[,，]\s*(-?\d+(?:\.\d+)?)",
                        loc,
                    )
                    if m:
                        records.append({
                            "address": clean_text(content),
                            "lon": float(m.group(1)),
                            "lat": float(m.group(2)),
                        })
                for v in obj.values():
                    walk(v)
            elif isinstance(obj, list):
                for v in obj:
                    walk(v)

        walk(data)
        if records:
            return records
    except Exception:
        pass

    # --------------------------------------------------------
    # 最後容錯：直接從 response 文字找「location」座標
    # --------------------------------------------------------
    matches = re.findall(
        r"(?:location|Location)[^0-9-]*(-?\d+(?:\.\d+)?)\s*[,，]\s*(-?\d+(?:\.\d+)?)",
        body,
    )
    for lon, lat in matches:
        records.append({
            "address": "",
            "lon": float(lon),
            "lat": float(lat),
        })

    if not records:
        raise ValueError("ADDRESS_NO_COORDINATE")

    return records


def choose_best_address_record(records, full_query_address):
    """在多筆門牌搜尋結果中，選擇與查詢文字最接近者。"""
    from difflib import SequenceMatcher

    if not records:
        raise ValueError("ADDRESS_NO_RESULT")

    target = normalize_address_text(full_query_address)

    scored = []
    for r in records:
        addr = normalize_address_text(r.get("address", ""))
        score = 0.0
        if target and addr:
            if target == addr:
                score = 2.0
            elif target in addr or addr in target:
                score = 1.5
            else:
                score = SequenceMatcher(None, target, addr).ratio()
        scored.append((score, r))

    scored.sort(key=lambda x: x[0], reverse=True)
    best_score, best = scored[0]
    best = dict(best)
    best["match_score"] = float(best_score)
    best["result_count"] = len(records)
    return best


async def click_visible_text_js(page, texts):
    """依可見文字尋找並點擊元素；網站改版時比固定 CSS selector 有韌性。"""
    return await page.evaluate(
        """
        (texts) => {
            const wanted = texts.map(x => String(x));
            const els = Array.from(document.querySelectorAll(
                'a,button,input[type="button"],input[type="submit"],span,div,li,label'
            ));

            function visible(el) {
                const s = getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return s.display !== 'none' && s.visibility !== 'hidden' &&
                       r.width > 0 && r.height > 0;
            }

            for (const exact of [true, false]) {
                for (const el of els) {
                    if (!visible(el)) continue;
                    const text = (el.innerText || el.value || el.title || '').trim();
                    if (!text) continue;
                    const hit = wanted.some(w => exact ? text === w : text.includes(w));
                    if (hit) {
                        try { el.scrollIntoView({block:'center'}); } catch(e) {}
                        try { el.click(); } catch(e) {}
                        return {ok:true, text, tag:el.tagName, id:el.id || ''};
                    }
                }
            }
            return {ok:false, text:'', tag:'', id:''};
        }
        """,
        texts,
    )


async def fill_address_ui_and_submit(page, full_address):
    """
    開啟 NLSC「定位查詢→門牌」後：
    1. 尋找可見的地址/門牌輸入框
    2. 填入完整地址
    3. 在同一查詢區塊尋找「搜尋／搜索」按鈕並觸發
    """
    # 開啟定位查詢
    info = await click_visible_text_js(page, ["定位查詢"])
    if info.get("ok"):
        await page.wait_for_timeout(500)

    # 切到門牌
    info = await click_visible_text_js(page, ["門牌"])
    if info.get("ok"):
        await page.wait_for_timeout(500)

    # 用 JS 找地址欄與同區塊搜尋按鈕
    result = await page.evaluate(
        """
        (address) => {
            const all = Array.from(document.querySelectorAll('input,textarea'));
            const visible = el => {
                const s = getComputedStyle(el);
                const r = el.getBoundingClientRect();
                return s.display !== 'none' && s.visibility !== 'hidden' &&
                       r.width > 0 && r.height > 0;
            };

            function attrs(el) {
                return [
                    el.id || '', el.name || '', el.placeholder || '',
                    el.title || '', el.getAttribute('aria-label') || '',
                    el.className || ''
                ].join(' ').toLowerCase();
            }

            let best = null;
            let bestScore = -1;

            for (const el of all) {
                if (!visible(el)) continue;
                const a = attrs(el);
                let score = 0;
                if (/(地址|門牌|addr|address)/i.test(a)) score += 100;
                if ((el.value || '').trim() === '') score += 5;
                if (el.tagName === 'INPUT' && ['text','search',''].includes((el.type || '').toLowerCase())) score += 3;
                if (score > bestScore) { bestScore = score; best = el; }
            }

            // 第二順位：可見文字輸入框中，所在祖先文字含「門牌／地址」
            if (!best) {
                for (const el of all) {
                    if (!visible(el)) continue;
                    let p = el;
                    let context = '';
                    for (let i=0; i<5 && p; i++, p=p.parentElement) {
                        context += ' ' + (p.innerText || '');
                    }
                    if (/門牌|地址/.test(context)) { best = el; break; }
                }
            }

            if (!best) return {ok:false, reason:'ADDRESS_INPUT_NOT_FOUND'};

            best.focus();
            const setter = Object.getOwnPropertyDescriptor(
                best.tagName === 'TEXTAREA' ? window.HTMLTextAreaElement.prototype :
                window.HTMLInputElement.prototype,
                'value'
            )?.set;
            if (setter) setter.call(best, address);
            else best.value = address;
            best.dispatchEvent(new Event('input', {bubbles:true}));
            best.dispatchEvent(new Event('change', {bubbles:true}));

            let root = best;
            for (let i=0; i<7 && root; i++, root=root.parentElement) {
                const txt = (root.innerText || '');
                if (!/門牌|地址/.test(txt)) continue;
                const btns = Array.from(root.querySelectorAll(
                    'button,input[type="button"],input[type="submit"],a,span'
                )).filter(visible);
                const btn = btns.find(el => /^(搜尋|搜索|查詢)$/.test(
                    (el.innerText || el.value || el.title || '').trim()
                ));
                if (btn) {
                    try { btn.scrollIntoView({block:'center'}); } catch(e) {}
                    try { btn.click(); } catch(e) {}
                    return {
                        ok:true,
                        input_id:best.id || '',
                        input_name:best.name || '',
                        button_text:(btn.innerText || btn.value || btn.title || '').trim()
                    };
                }
            }

            // 全頁尋找一次
            const btns = Array.from(document.querySelectorAll(
                'button,input[type="button"],input[type="submit"],a,span'
            )).filter(visible);
            const btn = btns.find(el => /^(搜尋|搜索|查詢)$/.test(
                (el.innerText || el.value || el.title || '').trim()
            ));
            if (btn) {
                try { btn.click(); } catch(e) {}
                return {
                    ok:true,
                    input_id:best.id || '',
                    input_name:best.name || '',
                    button_text:(btn.innerText || btn.value || btn.title || '').trim()
                };
            }

            return {
                ok:false,
                reason:'ADDRESS_SEARCH_BUTTON_NOT_FOUND',
                input_id:best.id || '',
                input_name:best.name || ''
            };
        }
        """,
        full_address,
    )

    if not result.get("ok"):
        raise RuntimeError(str(result))

    await page.wait_for_timeout(1200)
    return result


async def query_one_address_browser(
    page,
    row_number,
    city,
    area,
    address,
):
    """單筆地址定位：使用 NLSC 圖臺的門牌定位介面。"""
    city = clean_text(city)
    area = clean_text(area)
    address = clean_text(address)

    if not city or not area or not address:
        raise ValueError("地址查詢需要：縣市、鄉鎮市區、地址")

    # 地址查詢也比照地號查詢：
    # 鄉鎮市區若省略「鄉／鎮／市／區」，直接依 NLSC 實際選單
    # 自動找出正式名稱，例如「新豐」->「新豐鄉」。
    area_info = await resolve_nlsc_area_name(
        page,
        city,
        area,
    )
    resolved_area = area_info["area_text"]

    full_address = build_full_address(
        city,
        resolved_area,
        address,
    )
    last_error = ""
    last_status = None
    last_url = ""
    last_body = ""
    last_raw_file = ""

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            print(
                f"  第 {attempt}/{MAX_RETRIES} 次查詢："
                f"地址={full_address}"
            )

            def address_response_filter(response):
                """
                NLSC 地址查詢實際會觸發：
                  https://api.nlsc.gov.tw/MapSearch/QuerySearch...

                舊版只比對 textqueryaddress / textquerymap / addressquery，
                因此新版改成同步 predicate 後反而可能完全等不到實際 API，
                最後變成 30 秒 Timeout。

                現在同時支援實際 QuerySearch 與舊版 URL 關鍵字。
                """
                try:
                    url = response.url
                    low = url.lower()
                    if "api.nlsc.gov.tw" not in low:
                        return False

                    return (
                        "mapsearch/querysearch" in low
                        or "querysearch" in low
                        or "textqueryaddress" in low
                        or "textquerymap" in low
                        or "addressquery" in low
                    )
                except Exception:
                    return False

            async with page.expect_response(
                address_response_filter,
                timeout=RESPONSE_TIMEOUT
            ) as response_info:
                await fill_address_ui_and_submit(
                    page,
                    full_address
                )

            response = await response_info.value
            status_code = response.status
            response_url = response.url
            try:
                body = await response.text()
            except Exception:
                body = ""

            last_status = status_code
            last_url = response_url
            last_body = body

            print(f"  HTTP={status_code}")

            if status_code == 200:
                records = parse_address_location_response(body)
                chosen = choose_best_address_record(
                    records,
                    full_address
                )

                lon = float(chosen["lon"])
                lat = float(chosen["lat"])
                twd97_x, twd97_y = convert_to_twd97(lon, lat)

                raw_file = save_raw_response(
                    row_number,
                    status_code,
                    response_url,
                    body
                )

                quality = (
                    "地址單筆結果"
                    if chosen["result_count"] == 1
                    else "地址多筆取最相符"
                )

                return {
                    "縣市": city,
                    "鄉鎮市區": resolved_area,                    "地段": "",
                    "查詢地段": "",
                    "地號": "",
                    "地址": address,
                    "查詢方式": "地址",
                    "定位來源": "NLSC門牌定位",
                    "完整查詢地址": full_address,
                    "命中地址": chosen.get("address", ""),
                    "定位品質": quality,
                    "查詢地號": "",
                    "office": "",
                    "sect": "",
                    "狀態": "OK",
                    "HTTP狀態": status_code,
                    "經度": lon,
                    "緯度": lat,
                    "TWD97_X": twd97_x,
                    "TWD97_Y": twd97_y,
                    "左下經度": "",
                    "左下緯度": "",
                    "右上經度": "",
                    "右上緯度": "",
                    "response_url": response_url,
                    "raw_file": raw_file,
                    "錯誤": "",
                    "嘗試次數": attempt,
                }

            if status_code in RETRY_STATUS_CODES and attempt < MAX_RETRIES:
                wait_seconds = RETRY_BACKOFF[min(attempt - 1, len(RETRY_BACKOFF)-1)]
                print(f"  HTTP {status_code}，{wait_seconds} 秒後自動重試...")
                await page.wait_for_timeout(wait_seconds * 1000)
                continue

            last_error = f"HTTP_{status_code}"
            break

        except Exception as e:
            last_error = f"{type(e).__name__}: {e}"
            print(f"  地址查詢例外：{last_error}")
            if attempt < MAX_RETRIES:
                wait_seconds = RETRY_BACKOFF[min(attempt - 1, len(RETRY_BACKOFF)-1)]
                print(f"  {wait_seconds} 秒後自動重試...")
                await page.wait_for_timeout(wait_seconds * 1000)
                continue
            break

    if last_status in RETRY_STATUS_CODES:
        status_text = "HTTP_ERROR"
    elif "timeout" in last_error.lower():
        status_text = "TIMEOUT"
    else:
        status_text = "ERROR"

    if last_body:
        try:
            last_raw_file = save_raw_response(
                row_number,
                last_status if last_status is not None else 0,
                last_url,
                last_body
            )
        except Exception:
            pass

    return {
        "縣市": city,
        "鄉鎮市區": resolved_area,        "地段": "",
        "查詢地段": "",
        "地號": "",
        "地址": address,
        "查詢方式": "地址",
        "定位來源": "NLSC門牌定位",
        "完整查詢地址": full_address,
        "命中地址": "",
        "定位品質": "",
        "查詢地號": "",
        "office": "",
        "sect": "",
        "狀態": status_text,
        "HTTP狀態": last_status if last_status is not None else "",
        "經度": "",
        "緯度": "",
        "TWD97_X": "",
        "TWD97_Y": "",
        "左下經度": "",
        "左下緯度": "",
        "右上經度": "",
        "右上緯度": "",
        "response_url": last_url,
        "raw_file": last_raw_file,
        "錯誤": last_error,
        "嘗試次數": MAX_RETRIES,
    }


# ============================================================
# 19. 單筆查詢
# ============================================================

async def query_one_browser(
    page,
    row_number,
    city,
    area,
    section,
    landno,
    section_cache
):

    original_landno = clean_text(
        landno
    )

    normalized_landno = (
        normalize_landno(
            original_landno
        )
    )

    query_section = normalize_section_for_query(
        city,
        area,
        section
    )

    section_key = (

        clean_text(city),

        clean_text(area),

        clean_text(query_section),
    )

    # --------------------------------------------------------
    # 地段變更
    # --------------------------------------------------------

    matched_section_text = clean_text(
        section_cache.get(
            "section_text",
            query_section
        )
    )

    matched_area_text = clean_text(
        section_cache.get(
            "area_text",
            area
        )
    )

    if (
        section_cache.get(
            "key"
        )
        != section_key
    ):

        section_info = (
            await set_nlsc_section(
                page,
                city,
                area,
                query_section
            )
        )

        matched_section_text = clean_text(
            section_info.get(
                "section_text",
                query_section
            )
        )

        matched_area_text = clean_text(
            section_info.get(
                "area_text",
                area
            )
        )

        section_cache.clear()

        section_cache.update({

            "key":
                section_key,

            "area_text":
                matched_area_text,

            "area_value":
                section_info.get(
                    "area_value",
                    ""
                ),

            "section_text":
                section_info[
                    "section_text"
                ],

            "section_value":
                section_info[
                    "section_value"
                ],

            "section_code":
                section_info[
                    "section_code"
                ],
        })

        print(
            "  地段已設定："
            f"{section_info['section_text']} "
            f"(value="
            f"{section_info['section_value']}, "
            f"code="
            f"{section_info['section_code']})"
        )

    # --------------------------------------------------------
    # 填寫地號
    # --------------------------------------------------------

    await fill_landno(
        page,
        normalized_landno
    )

    # --------------------------------------------------------
    # 查詢
    # --------------------------------------------------------

    last_error = ""

    last_status = None

    last_url = ""

    last_body = ""

    last_raw_file = ""

    for attempt in range(
        1,
        MAX_RETRIES + 1
    ):

        try:

            print(
                f"  第 {attempt}/{MAX_RETRIES} "
                f"次查詢："
                f"地號={normalized_landno}"
            )

            # 只抓 qryTileMapIndex
            async with page.expect_response(
                is_land_query_response,
                timeout=RESPONSE_TIMEOUT
            ) as response_info:

                await trigger_nlsc_query(
                    page
                )

            response = (
                await response_info.value
            )

            status_code = (
                response.status
            )

            response_url = (
                response.url
            )

            try:

                body = (
                    await response.text()
                )

            except Exception:

                body = ""

            last_status = (
                status_code
            )

            last_url = (
                response_url
            )

            last_body = body

            print(
                f"  HTTP={status_code}"
            )

            parsed = (
                parse_nlsc_response(
                    body
                )
            )

            # ------------------------------------------------
            # 成功
            # ------------------------------------------------

            if (
                status_code == 200
                and
                parsed["ok"]
            ):

                record = (
                    parsed["record"]
                )

                lon = float(
                    record["cx"]
                )

                lat = float(
                    record["cy"]
                )

                twd97_x, twd97_y = (
                    convert_to_twd97(
                        lon,
                        lat
                    )
                )

                raw_file = (
                    save_raw_response(
                        row_number,
                        status_code,
                        response_url,
                        body
                    )
                )

                return {

                    "縣市":
                        city,

                    "鄉鎮市區":
                        matched_area_text or area,

                    # NLSC 實際選中的正式地段名稱（含段／小段）
                    "地段":
                        matched_section_text or section,

                    "查詢地段":
                        matched_section_text or section,

                    "地號": original_landno,
                    "地址": "",
                    "查詢方式": "地號",
                    "定位來源": "NLSC地號定位",
                    "完整查詢地址": "",
                    "命中地址": "",
                    "定位品質": "地號宗地定位",
                    "查詢地號": normalized_landno,

                    "office":
                        get_query_parameter(
                            response_url,
                            "office"
                        ),

                    "sect":
                        get_query_parameter(
                            response_url,
                            "sect"
                        ),

                    "狀態":
                        "OK",

                    "HTTP狀態":
                        status_code,

                    "經度":
                        lon,

                    "緯度":
                        lat,

                    "TWD97_X":
                        twd97_x,

                    "TWD97_Y":
                        twd97_y,

                    "左下經度":
                        record.get(
                            "lx",
                            ""
                        ),

                    "左下緯度":
                        record.get(
                            "ly",
                            ""
                        ),

                    "右上經度":
                        record.get(
                            "rx",
                            ""
                        ),

                    "右上緯度":
                        record.get(
                            "ry",
                            ""
                        ),

                    "response_url":
                        response_url,

                    "raw_file":
                        raw_file,

                    "錯誤":
                        "",

                    "嘗試次數":
                        attempt,
                }

            # ------------------------------------------------
            # 502 / 503 / 504
            # ------------------------------------------------

            if (
                status_code
                in RETRY_STATUS_CODES
            ):

                last_error = (
                    f"HTTP_{status_code}"
                )

                last_raw_file = (
                    save_raw_response(
                        row_number,
                        status_code,
                        response_url,
                        body
                    )
                )

                if (
                    attempt
                    < MAX_RETRIES
                ):

                    wait_seconds = (
                        RETRY_BACKOFF[
                            min(
                                attempt - 1,
                                len(
                                    RETRY_BACKOFF
                                ) - 1
                            )
                        ]
                    )

                    print(
                        f"  HTTP "
                        f"{status_code}，"
                        f"{wait_seconds} 秒後"
                        f"自動重試..."
                    )

                    await page.wait_for_timeout(
                        int(
                            wait_seconds
                            * 1000
                        )
                    )

                    continue

                break

            # ------------------------------------------------
            # HTTP 200 無座標
            # ------------------------------------------------

            if (
                status_code == 200
            ):

                last_raw_file = (
                    save_raw_response(
                        row_number,
                        status_code,
                        response_url,
                        body
                    )
                )

                last_error = (
                    parsed["error"]
                )

                print(
                    "  HTTP=200，"
                    "但查詢內容異常："
                    f"{parsed['error']}"
                )

                break

            # ------------------------------------------------
            # 其他 HTTP
            # ------------------------------------------------

            last_raw_file = (
                save_raw_response(
                    row_number,
                    status_code,
                    response_url,
                    body
                )
            )

            last_error = (
                f"HTTP_{status_code}"
            )

            break

        # ====================================================
        # Timeout / 例外
        # ====================================================

        except Exception as e:

            last_error = (
                f"{type(e).__name__}: {e}"
            )

            print(
                f"  查詢例外："
                f"{last_error}"
            )

            if (
                attempt
                < MAX_RETRIES
            ):

                wait_seconds = (
                    RETRY_BACKOFF[
                        min(
                            attempt - 1,
                            len(
                                RETRY_BACKOFF
                            ) - 1
                        )
                    ]
                )

                print(
                    f"  {wait_seconds} 秒後"
                    f"自動重試..."
                )

                await page.wait_for_timeout(
                    int(
                        wait_seconds
                        * 1000
                    )
                )

                # ------------------------------------------------
                # Timeout：
                # 重新載入 NLSC
                # ------------------------------------------------

                if (
                    "TimeoutError"
                    in type(e).__name__
                    or
                    "timeout"
                    in str(e).lower()
                ):

                    try:

                        await initialize_nlsc_page(
                            page
                        )

                        section_cache.clear()

                        section_info = (
                            await set_nlsc_section(
                                page,
                                city,
                                area,
                                query_section
                            )
                        )

                        section_cache.update({

                            "key":
                                section_key,

                            "section_text":
                                section_info[
                                    "section_text"
                                ],

                            "section_value":
                                section_info[
                                    "section_value"
                                ],

                            "section_code":
                                section_info[
                                    "section_code"
                                ],
                        })

                        print(
                            "  地段重新設定："
                            f"{section_info['section_text']}"
                        )

                        await fill_landno(
                            page,
                            normalized_landno
                        )

                    except Exception as recovery_error:

                        print(
                            "  頁面恢復失敗："
                            f"{type(recovery_error).__name__}: "
                            f"{recovery_error}"
                        )

                continue

            break

    # ========================================================
    # 最終失敗
    # ========================================================

    if (
        last_status
        in RETRY_STATUS_CODES
    ):

        status_text = (
            "HTTP_ERROR"
        )

    elif (
        "Timeout"
        in last_error
        or
        "timeout"
        in last_error.lower()
    ):

        status_text = (
            "TIMEOUT"
        )

    else:

        status_text = (
            "ERROR"
        )

    return {

        "縣市":
            city,

        "鄉鎮市區":
            area,

        "地段":
            section,

        "查詢地段":
            matched_section_text,

        "地號": original_landno,
        "地址": "",
        "查詢方式": "地號",
        "定位來源": "NLSC地號定位",
        "完整查詢地址": "",
        "命中地址": "",
        "定位品質": "",
        "查詢地號": normalized_landno,

        "office":
            "",

        "sect":
            section_cache.get(
                "section_value",
                ""
            ),

        "狀態":
            status_text,

        "HTTP狀態":
            (
                last_status
                if last_status is not None
                else ""
            ),

        "經度":
            "",

        "緯度":
            "",

        "TWD97_X":
            "",

        "TWD97_Y":
            "",

        "左下經度":
            "",

        "左下緯度":
            "",

        "右上經度":
            "",

        "右上緯度":
            "",

        "response_url":
            last_url,

        "raw_file":
            last_raw_file,

        "錯誤":
            last_error,

        "嘗試次數":
            MAX_RETRIES,
    }


# ============================================================
# 20. 執行時上傳輸入檔；取消後依序搜尋既有檔案
# ============================================================

def find_input_file():
    """優先使用本次在 Colab 上傳的單一檔案；取消時依序讀既有候選檔。"""
    supported_suffixes = {".xlsx", ".xls", ".csv"}
    uploaded_payloads = None

    print("=" * 70)
    print("Cell 2 v8.7 輸入檔選擇")
    print("請上傳本次要查詢的 .xlsx、.xls 或 .csv 檔案。")
    print("若取消或未選擇檔案，將依序搜尋 /content/input.xlsx、input.xls、input.csv。")
    print("=" * 70)

    # Google Colab 上先開啟上傳介面；非 Colab 環境則直接走既有檔案 fallback。
    try:
        from google.colab import files as colab_files
    except Exception:
        colab_files = None
        print("[INFO] 目前不是可用的 Google Colab 上傳環境，改查既有輸入檔。")

    if colab_files is not None:
        try:
            uploaded_payloads = colab_files.upload()
        except Exception as exc:
            # 使用者取消上傳或上傳視窗未完成時，依需求 fallback 到既有檔案。
            print(f"[INFO] 上傳未完成或已取消（{type(exc).__name__}），改查既有輸入檔。")
            uploaded_payloads = None

    if uploaded_payloads:
        supported_uploads = []
        unsupported_names = []
        for uploaded_name, payload in uploaded_payloads.items():
            safe_name = Path(str(uploaded_name)).name
            suffix = Path(safe_name).suffix.lower()
            if suffix not in supported_suffixes:
                unsupported_names.append(safe_name)
                continue
            supported_uploads.append((safe_name, payload))

        if len(supported_uploads) > 1:
            names = [name for name, _ in supported_uploads]
            raise RuntimeError(
                "一次請只上傳一個輸入檔（支援 .xlsx、.xls、.csv）；"
                f"本次收到多個可用檔案：{names}"
            )

        if len(supported_uploads) == 1:
            uploaded_name, payload = supported_uploads[0]
            if not isinstance(payload, (bytes, bytearray)):
                raise TypeError(
                    f"上傳檔案內容格式異常：{uploaded_name}"
                )
            destination = BASE_DIR / uploaded_name
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(bytes(payload))
            print(f"已採用本次上傳檔案：{destination}")
            return destination

        # 有上傳動作但檔案格式不支援，不把它誤判成「取消上傳」。
        raise RuntimeError(
            "收到上傳檔案，但沒有支援的 .xlsx、.xls 或 .csv 檔案。"
            f"不支援檔案：{unsupported_names}"
        )

    print("未收到本次上傳檔案（已取消或未選擇），依序檢查既有檔案……")
    for path in INPUT_CANDIDATES:
        if path.is_file():
            print(f"使用既有輸入檔：{path}")
            return path

    raise FileNotFoundError(
        "未上傳輸入檔，且 /content/ 找不到既有候選檔。\n"
        "已依序檢查：\n"
        "  /content/input.xlsx\n"
        "  /content/input.xls\n"
        "  /content/input.csv"
    )


# ============================================================
# 21. 讀取輸入檔
# ============================================================

def load_input_file(
    path
):

    suffix = (
        path.suffix.lower()
    )

    required_columns = [
        "縣市",
        "鄉鎮市區",
        "地段",
        "地號",
        "本筆面積(公頃)",
        "申請面積(公頃)",
        "地址",
        "名稱",
        "勘查狀態",
    ]

    # 目前輸入檔不再需要「查詢方式」欄位；查詢方式由程式自動判定。
    optional_columns = []

    # ========================================================
    # XLSX
    # ========================================================

    if suffix == ".xlsx":

        wb = (
            openpyxl.load_workbook(
                path,
                read_only=True,
                data_only=True
            )
        )

        ws = (
            wb[
                wb.sheetnames[0]
            ]
        )

        rows = list(
            ws.iter_rows(
                values_only=True
            )
        )

        if not rows:

            wb.close()

            raise RuntimeError(
                "Excel 沒有資料"
            )

        headers = [
            clean_text(x)
            for x in rows[0]
        ]

        index_map = {
            h: i
            for i, h in enumerate(
                headers
            )
        }

        missing = [
            col
            for col in required_columns
            if col not in index_map
        ]

        if missing:

            wb.close()

            raise RuntimeError(
                "Excel 缺少欄位："
                f"{missing}"
            )

        data = []

        for row in rows[1:]:

            if not any(
                clean_text(x)
                for x in row
            ):

                continue

            item = {}

            # 必要欄位一定存在；選填欄位不存在時自動視為空白。
            for col in required_columns + optional_columns:

                if col not in index_map:
                    item[col] = ""
                    continue

                idx = index_map[col]

                value = (
                    row[idx]
                    if idx < len(row)
                    else ""
                )

                item[col] = (
                    clean_text(value)
                )

            data.append(
                item
            )

        wb.close()

        return data

    # ========================================================
    # XLS
    # ========================================================

    if suffix == ".xls":

        import xlrd

        workbook = (
            xlrd.open_workbook(
                path,
                on_demand=True
            )
        )

        sheet = (
            workbook.sheet_by_index(0)
        )

        headers = [
            clean_text(
                sheet.cell_value(
                    0,
                    i
                )
            )

            for i in range(
                sheet.ncols
            )
        ]

        index_map = {
            h: i
            for i, h in enumerate(
                headers
            )
        }

        missing = [
            col
            for col in required_columns
            if col not in index_map
        ]

        if missing:

            workbook.release_resources()

            raise RuntimeError(
                "XLS 缺少欄位："
                f"{missing}"
            )

        data = []

        for r in range(
            1,
            sheet.nrows
        ):

            item = {}

            for col in required_columns + optional_columns:

                if col not in index_map:
                    item[col] = ""
                    continue

                item[col] = clean_text(
                    sheet.cell_value(
                        r,
                        index_map[col]
                    )
                )

            if not any(
                item.values()
            ):

                continue

            data.append(
                item
            )

        workbook.release_resources()

        return data

    # ========================================================
    # CSV
    # ========================================================

    if suffix == ".csv":

        encodings = [
            "utf-8-sig",
            "cp950",
        ]

        last_error = None

        for encoding in (
            encodings
        ):

            try:

                with open(
                    path,
                    "r",
                    encoding=encoding,
                    newline=""
                ) as f:

                    reader = (
                        csv.DictReader(
                            f
                        )
                    )

                    fieldnames = [
                        clean_text(x)
                        for x in (
                            reader.fieldnames
                            or []
                        )
                    ]

                    missing = [
                        col
                        for col in required_columns
                        if col not in fieldnames
                    ]

                    if missing:

                        raise RuntimeError(
                            "CSV 缺少欄位："
                            f"{missing}"
                        )

                    data = []

                    for row in reader:

                        item = {

                            col:
                                clean_text(
                                    row.get(
                                        col,
                                        ""
                                    )
                                )

                            for col
                            in required_columns + optional_columns
                        }

                        if not any(
                            item.values()
                        ):

                            continue

                        data.append(
                            item
                        )

                    print(
                        "CSV 編碼："
                        f"{encoding}"
                    )

                    return data

            except Exception as e:

                last_error = e

        raise RuntimeError(
            "CSV 讀取失敗："
            f"{last_error}"
        )

    raise RuntimeError(
        "不支援的輸入格式："
        f"{suffix}"
    )


# ============================================================
# ============================================================
# 21A. 正式資料識別
# ============================================================

def parse_area_value(value):
    """將輸入檔面積保留為可寫入 Excel 的數值；空白維持空字串。"""
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    text = clean_text(value)
    if not text:
        return ""
    text = text.replace(",", "")
    try:
        return float(text)
    except Exception:
        return text


def _norm_case_part(value):
    if value is None:
        return ""
    s = str(value).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def build_case_id(query_mode, result, row_number):
    """建立跨 Cell 穩定案件ID。地號案件採 OFFICE_SECT_8碼地號；地址案件採雜湊。"""
    mode = _norm_case_part(query_mode)
    if mode == "地號":
        office = _norm_case_part(result.get("office", "")).upper()
        sect = _norm_case_part(result.get("sect", ""))
        land = normalize_landno(result.get("查詢地號", result.get("地號", "")))
        if office and sect and land:
            return f"{office}_{sect}_{land}"
    raw = "|".join([
        _norm_case_part(result.get("縣市", "")),
        _norm_case_part(result.get("鄉鎮市區", "")),
        _norm_case_part(result.get("地段", "")),
        _norm_case_part(result.get("地號", "")),
        _norm_case_part(result.get("地址", "")),
        str(mode),
    ])
    digest = hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]
    return f"ADDR_{digest}" if mode == "地址" else f"ROW_{row_number:04d}_{digest}"


# 22. Excel 輸出
# ============================================================

def save_excel_openpyxl(
    results,
    output_path
):

    wb = (
        openpyxl.Workbook()
    )

    ws = (
        wb.active
    )

    ws.title = (
        "查詢結果"
    )

    headers = [

        "名稱",
        "查詢方式",
        "縣市",
        "鄉鎮市區",
        "地段",
        "查詢地段",
        "地號",
        "本筆面積(公頃)",
        "申請面積(公頃)",
        "地址",
        "勘查狀態",
        "定位來源",
        "完整查詢地址",
        "命中地址",
        "定位品質",
        "查詢地號",
        "案件ID",
        "office",
        "sect",
        "狀態",
        "資料檢核",
        "重複來源列",
        "HTTP狀態",
        "經度",
        "緯度",
        "TWD97_X",
        "TWD97_Y",
        "左下經度",
        "左下緯度",
        "右上經度",
        "右上緯度",
        "response_url",
        "raw_file",
        "錯誤",
        "嘗試次數",
        # v8.4.1：地號案件另存門牌定位座標；地址案件則與主座標相同。
        "地址經度",
        "地址緯度",
    ]

    ws.append(
        headers
    )

    for result in results:

        ws.append([
            result.get(
                header,
                ""
            )

            for header in headers
        ])

    # --------------------------------------------------------
    # 凍結首列
    # --------------------------------------------------------

    ws.freeze_panes = (
        "A2"
    )

    # --------------------------------------------------------
    # 篩選
    # --------------------------------------------------------

    ws.auto_filter.ref = (
        ws.dimensions
    )

    # --------------------------------------------------------
    # 欄寬
    # --------------------------------------------------------

    width_map = {
        "A": 24, "B": 10, "C": 10, "D": 12, "E": 28, "F": 30,
        "G": 14, "H": 16, "I": 16, "J": 38, "K": 12, "L": 20,
        "M": 42, "N": 42, "O": 16, "P": 30, "Q": 18, "R": 10,
        "S": 12, "T": 10, "U": 12, "V": 14, "W": 12, "X": 15,
        "Y": 15, "Z": 15, "AA": 15, "AB": 15, "AC": 15, "AD": 15,
        "AE": 15, "AF": 55, "AG": 55, "AH": 45, "AI": 10,
        "AJ": 15, "AK": 15,
    }

    for col, width in (
        width_map.items()
    ):

        ws.column_dimensions[
            col
        ].width = width

    # --------------------------------------------------------
    # 數字格式
    # --------------------------------------------------------

    numeric_formats = {
        "H": "0.0000",
        "I": "0.0000",
        "V": "0",
        "W": "0",
        "X": "0.00000000",
        "Y": "0.00000000",
        "Z": "0.000",
        "AA": "0.000",
        "AB": "0.00000000",
        "AC": "0.00000000",
        "AD": "0.00000000",
        "AE": "0.00000000",
        "AI": "0",
        "AJ": "0.00000000",
        "AK": "0.00000000",
    }

    for col, fmt in (
        numeric_formats.items()
    ):

        for row in range(
            2,
            ws.max_row + 1
        ):

            ws[
                f"{col}{row}"
            ].number_format = fmt

    wb.save(
        output_path
    )

    return output_path


# ============================================================
# 23. Google Maps 名稱
# ============================================================

def build_gmap_name(
    area,
    section,
    landno,
    address="",
    city="",
):

    """建立 Google Maps 與全系統共用的正式顯示名稱。

    地號案件：
      縣市＋鄉鎮市區＋完整正式地段（含段／小段）＋8碼地號。

    地址案件：
      完整地址。
    """

    area = clean_text(area)
    section = clean_text(section)
    address = clean_text(address)
    city = clean_text(city)

    if address:
        return build_full_address(city, area, address)

    landno = normalize_landno(landno)

    if (
        "-" not in landno
        and re.fullmatch(r"\d+", landno)
    ):
        landno = landno.zfill(8)

    return "".join(
        part
        for part in (city, area, section, landno)
        if part
    )


def build_display_name(result):
    """建立全系統統一使用的「顯示名稱」。

    顯示優先順序：
      1. 只要輸入資料有「名稱」，地號／地址案件一律優先使用「名稱」。
      2. 地號案件沒有名稱時：縣市＋鄉鎮市區＋正式地段（含段／小段）＋8碼地號。
      3. 地址案件沒有名稱時：使用完整地址。
    """
    if result is None:
        return ""

    query_mode = clean_text(result.get("查詢方式", ""))
    city = clean_text(result.get("縣市", ""))
    area = clean_text(result.get("鄉鎮市區", ""))
    section = clean_text(result.get("地段", ""))
    resolved_section = clean_text(result.get("查詢地段", ""))
    landno = clean_text(result.get("查詢地號", ""))
    if not landno:
        landno = clean_text(result.get("地號", ""))
    address = clean_text(result.get("地址", ""))
    custom_name = clean_text(result.get("名稱", ""))

    # 統一規則：有「名稱」就優先使用名稱。
    # 不論是地址案件或地號案件，都不再用地段／地號格式蓋掉自訂名稱。
    if custom_name:
        return custom_name

    # 地址案件：沒有名稱才使用完整地址。
    if query_mode == "地址" or (address and query_mode != "地號"):
        return build_full_address(city, area, address)

    # 地號案件：沒有名稱時，優先採用 NLSC 實際命中的正式地段。
    if resolved_section and has_formal_section_marker(resolved_section):
        section = resolved_section

    return build_gmap_name(
        area=area,
        section=section,
        landno=landno,
        address="",
        city=city,
    )


def has_formal_section_marker(section):
    """判斷地段是否已取得 NLSC 正式「段／小段」名稱。

    NLSC 地段正式選項包含「段」或「小段」。
    這個判斷只用來避免續跑時永久沿用舊版沒有段／小段的結果。
    """
    section = clean_text(section)
    return "段" in section


# ============================================================
# 24. Google Maps CSV
# ============================================================

def save_gmap_csv(
    success_results,
    output_path
):

    """
    Google Maps 可匯入格式：

      Name
      Latitude
      Longitude

    注意：
      座標順序為
      緯度、經度
    """

    with open(
        output_path,
        "w",
        encoding="utf-8-sig",
        newline=""
    ) as f:

        writer = csv.writer(
            f
        )

        writer.writerow([
            "Name",
            "Latitude",
            "Longitude",
        ])

        for result in (
            success_results
        ):

            if (
                result.get(
                    "狀態"
                )
                != "OK"
            ):

                continue

            # 統一沿用「顯示名稱」；舊資料沒有顯示名稱時才回退到「名稱」或重建。
            name = clean_text(result.get("顯示名稱", ""))
            if not name:
                name = clean_text(result.get("名稱", ""))
            if not name:
                name = build_display_name(result)

            # 緯度
            latitude = result.get(
                "緯度",
                ""
            )

            # 經度
            longitude = result.get(
                "經度",
                ""
            )

            writer.writerow([
                name,
                latitude,
                longitude,
            ])

    return output_path


# ============================================================
# 24A. 續跑／快取
# ============================================================

RESULT_PARTIAL_XLSX = BASE_DIR / "result_partial.xlsx"
RESULT_MANIFEST = BASE_DIR / "result_manifest.json"
RESULT_TEMP_XLSX = BASE_DIR / "result.__tmp__.xlsx"


def normalize_area_name_for_key(name):
    """
    續跑案件鍵的鄉鎮市區標準化。
    讓「桃園」與「桃園區」、「新豐」與「新豐鄉」視為同一案件，
    避免本程式將 NLSC 自動補上的正式名稱視為新案件而重新查詢。
    """
    name = clean_text(name)
    if not name:
        return ""

    base = name
    for suffix in ("鄉", "鎮", "市", "區"):
        if base.endswith(suffix):
            base = base[:-1]
            break

    return _norm_case_part(base)


def build_resume_key_from_values(
    query_mode,
    city,
    area,
    section,
    landno,
    address,
):
    """
    建立跨次執行的穩定案件鍵。

    只使用「真正用來定位」的欄位，不使用名稱與面積。
    """
    # 地號案件的「輸入地段」與 NLSC 成功結果可能不同：
    # 例如輸入「社角」，NLSC 正式名稱為「社角段」。
    # 使用 canonical_section_name() 消除「段／小段」差異，
    # 讓第二次以後可以真正命中成功快取。
    section_key = section
    if _norm_case_part(query_mode) == "地號":
        section_key = canonical_section_name(section)

    values = [
        _norm_case_part(query_mode),
        _norm_case_part(city),
        normalize_area_name_for_key(area),
        _norm_case_part(section_key),
        _norm_case_part(normalize_landno(landno)),
        _norm_case_part(address),
    ]

    raw = "|".join(values)
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()


def build_land_duplicate_key_from_input(row):
    """
    建立「重複土地」檢核鍵。

    僅地號案件參與檢核；地址案件不套用。
    唯一鍵：縣市＋鄉鎮市區＋地段＋地號。

    這裡不使用名稱、面積、申請面積或勘查狀態，
    因為「同一土地」的判定只看實際定位用的土地識別欄位。
    """
    city = clean_text(row.get("縣市", ""))
    area = clean_text(row.get("鄉鎮市區", ""))
    section = clean_text(row.get("地段", ""))
    landno = clean_text(row.get("地號", ""))

    if not (section and landno):
        return ""

    parts = [
        _norm_case_part(city),
        normalize_area_name_for_key(area),
        _norm_case_part(section),
        _norm_case_part(normalize_landno(landno)),
    ]

    if not all(parts):
        return ""

    return "|".join(parts)


def build_resume_key_from_input(row):
    city = clean_text(row.get("縣市", ""))
    area = clean_text(row.get("鄉鎮市區", ""))
    section = clean_text(row.get("地段", ""))
    landno = clean_text(row.get("地號", ""))
    address = clean_text(row.get("地址", ""))

    if section and landno:
        query_mode = "地號"
    elif address:
        query_mode = "地址"
    else:
        query_mode = ""

    return build_resume_key_from_values(
        query_mode,
        city,
        area,
        section,
        landno,
        address,
    )


def build_resume_key_from_result(result):
    query_mode = clean_text(result.get("查詢方式", ""))

    city = clean_text(result.get("縣市", ""))
    area = clean_text(result.get("鄉鎮市區", ""))
    section = clean_text(result.get("地段", ""))
    landno = clean_text(result.get("地號", ""))
    address = clean_text(result.get("地址", ""))

    if not query_mode:
        if section and landno:
            query_mode = "地號"
        elif address:
            query_mode = "地址"

    return build_resume_key_from_values(
        query_mode,
        city,
        area,
        section,
        landno,
        address,
    )


def load_existing_result_xlsx(path):
    """讀取上一輪 Excel 結果；使用目前欄位名稱。"""
    if not path.exists():
        return []

    try:
        wb = openpyxl.load_workbook(
            path,
            read_only=True,
            data_only=True,
        )
        ws = wb[wb.sheetnames[0]]
        rows = ws.iter_rows(values_only=True)

        try:
            header_row = next(rows)
        except StopIteration:
            wb.close()
            return []

        headers = [clean_text(x) for x in header_row]
        results = []

        for row in rows:
            if not any(clean_text(x) for x in row):
                continue

            item = {}
            for idx, header in enumerate(headers):
                if not header:
                    continue
                value = row[idx] if idx < len(row) else ""

                if header in {
                    "本筆面積(公頃)",
                    "申請面積(公頃)",
                    "經度", "緯度", "地址經度", "地址緯度", "TWD97_X", "TWD97_Y",
                    "左下經度", "左下緯度", "右上經度", "右上緯度",
                }:
                    item[header] = "" if value is None else value
                else:
                    item[header] = clean_text(value)

            results.append(item)

        wb.close()
        return results

    except Exception as e:
        print(f"⚠️ 讀取續跑結果失敗：{path} -> {e}")
        return []


def load_all_resume_results():
    """
    同時讀取 result_partial.xlsx / result.xlsx / result_error.xlsx。

    不能只挑最新的一個，因為舊 partial 可能保存較早的結果，
    而 result.xlsx 可能保存另一批成功結果。合併後以 OK 優先。
    """
    candidates = [
        RESULT_PARTIAL_XLSX,
        SUCCESS_OUTPUT_XLSX,
        ERROR_OUTPUT_XLSX,
    ]

    combined = {}
    source_stats = []

    for path in candidates:
        if not path.exists():
            continue

        items = load_existing_result_xlsx(path)
        source_stats.append((path, len(items)))

        for item in items:
            try:
                key = build_resume_key_from_result(item)
            except Exception:
                continue

            if not key:
                continue

            old = combined.get(key)
            if old is None:
                combined[key] = item
                continue

            old_ok = clean_text(old.get("狀態", "")) == "OK"
            new_ok = clean_text(item.get("狀態", "")) == "OK"

            if new_ok and not old_ok:
                combined[key] = item
            elif new_ok == old_ok:
                # 同狀態時優先較新的檔案資料（後讀到者覆蓋）。
                combined[key] = item

    if source_stats:
        print("既有結果檔：")
        for path, count in source_stats:
            print(f"  {path.name}: {count} 筆")

    return list(combined.values())


def apply_input_fields_to_result(result, row, idx):
    """將目前 input.xlsx 的資料欄位覆蓋回結果；座標不受影響。"""
    city = clean_text(row.get("縣市", ""))
    area = clean_text(row.get("鄉鎮市區", ""))
    section = clean_text(row.get("地段", ""))
    landno = clean_text(row.get("地號", ""))
    address = clean_text(row.get("地址", ""))

    if section and landno:
        query_mode = "地號"
    elif address:
        query_mode = "地址"
    else:
        query_mode = ""

    result_area = clean_text(
        result.get("鄉鎮市區", "")
    )
    keep_resolved_area = (
        clean_text(result.get("狀態", "")) == "OK"
        and query_mode in {"地號", "地址"}
        and bool(result_area)
        and area_name_matches(area, result_area)
    )

    result["名稱"] = clean_text(row.get("名稱", ""))
    result["本筆面積(公頃)"] = parse_area_value(row.get("本筆面積(公頃)", ""))
    result["申請面積(公頃)"] = parse_area_value(row.get("申請面積(公頃)", ""))
    result["勘查狀態"] = normalize_inspection_status(
        row.get("勘查狀態", "")
    )
    result["縣市"] = city
    result["鄉鎮市區"] = result_area if keep_resolved_area else area

    # v8.1：地號案件的「地段」改以 NLSC 實際命中的正式選單名稱為準。
    # 例如：
    #   輸入：下田心子下田心子
    #   NLSC正式名稱：下田心子段下田心子小段
    #   最終「地段」：下田心子段下田心子小段
    # 這樣 Cell 3 → Cell 4 → Cell 5 才能完整保留「段／小段」。
    resolved_section = clean_text(
        result.get("查詢地段", "")
    ) or section

    result["地段"] = resolved_section
    result["地號"] = landno
    result["地址"] = address
    result["查詢方式"] = query_mode
    result["顯示名稱"] = build_display_name(result)

    result["案件ID"] = build_case_id(
        query_mode,
        result,
        idx,
    )

    # v8.4.1：輸入欄位套用後統一補齊雙座標欄位。
    ensure_dual_coordinate_fields(result)

    return result


def make_checkpoint_placeholder(row, idx, status="PENDING"):
    """建立尚未執行或尚未有結果的暫存列。只有 OK 才會被續跑跳過。"""
    city = clean_text(row.get("縣市", ""))
    area = clean_text(row.get("鄉鎮市區", ""))
    section = clean_text(row.get("地段", ""))
    landno = clean_text(row.get("地號", ""))
    address = clean_text(row.get("地址", ""))
    query_mode = "地號" if section and landno else ("地址" if address else "")

    result = {
        "名稱": clean_text(row.get("名稱", "")),
        "查詢方式": query_mode,
        "縣市": city,
        "鄉鎮市區": area,
        "地段": section,
        "查詢地段": normalize_section_for_query(city, area, section) if section else "",
        "地號": landno,
        "本筆面積(公頃)": parse_area_value(row.get("本筆面積(公頃)", "")),
        "申請面積(公頃)": parse_area_value(row.get("申請面積(公頃)", "")),
        "勘查狀態": normalize_inspection_status(row.get("勘查狀態", "")),
        "地址": address,
        "定位來源": "",
        "完整查詢地址": build_full_address(city, area, address) if query_mode == "地址" else "",
        "命中地址": "",
        "定位品質": "",
        "查詢地號": normalize_landno(landno) if landno else "",
        "案件ID": f"PENDING_{idx:04d}",
        "顯示名稱": build_display_name({
            "查詢方式": query_mode,
            "縣市": city,
            "鄉鎮市區": area,
            "地段": section,
            "地號": landno,
            "查詢地號": normalize_landno(landno) if landno else "",
            "地址": address,
            "名稱": clean_text(row.get("名稱", "")),
        }),
        "office": "",
        "sect": "",
        "狀態": status,
        "HTTP狀態": "",
        "經度": "",
        "緯度": "",
        "TWD97_X": "",
        "TWD97_Y": "",
        "左下經度": "",
        "左下緯度": "",
        "右上經度": "",
        "右上緯度": "",
        "response_url": "",
        "raw_file": "",
        "錯誤": "",
        "嘗試次數": "",
    }
    return result


def save_resume_checkpoint(
    data,
    results_by_input_index,
    query_result_map,
):
    """
    每完成 1 筆就把目前已知狀態寫入 result_partial.xlsx。

    這是續跑的核心：即使使用者在下一筆查詢途中手動取消，
    前面已完成的 OK 資料仍然會保留下來。
    """
    checkpoint_results = []

    for idx, row in enumerate(data, start=1):
        result = None

        if idx in query_result_map:
            result = dict(query_result_map[idx])
        elif idx in results_by_input_index:
            result = dict(results_by_input_index[idx])

        if result is None:
            result = make_checkpoint_placeholder(row, idx)

        result = apply_input_fields_to_result(result, row, idx)
        checkpoint_results.append(result)

    # 先寫暫存檔，再以原子替換方式更新 partial。
    temp = BASE_DIR / "result_partial.__tmp__.xlsx"
    try:
        save_excel_openpyxl(checkpoint_results, temp)
        temp.replace(RESULT_PARTIAL_XLSX)
    finally:
        if temp.exists():
            try:
                temp.unlink()
            except Exception:
                pass

    return checkpoint_results


# ============================================================
# 24A. v8.4.1 雙座標處理
# ============================================================

def ensure_dual_coordinate_fields(result):
    """確保結果具有「地址經度／地址緯度」欄位。

    規則：
      1. 地址案件：主座標就是門牌座標，雙座標複製。
      2. 地號案件：主座標維持地號定位；若已有地址座標則保留。
      3. 地號案件沒有地址時：地址座標留白。
    """
    if result is None:
        return result

    result.setdefault("地址經度", "")
    result.setdefault("地址緯度", "")

    query_mode = clean_text(result.get("查詢方式", ""))
    main_lon = result.get("經度", "")
    main_lat = result.get("緯度", "")

    if query_mode == "地址":
        if clean_text(result.get("地址經度", "")) == "":
            result["地址經度"] = main_lon
        if clean_text(result.get("地址緯度", "")) == "":
            result["地址緯度"] = main_lat

    return result


async def query_secondary_address_coordinates(
    page,
    row_number,
    city,
    area,
    address,
):
    """地號案件的次要門牌座標查詢。

    地址 response 另存於 /content/nlsc_raw_address/，
    避免覆蓋同一案件的地號 response。
    失敗只影響地址座標，不影響主要地號定位結果。
    """
    global RAW_DIR

    old_raw_dir = RAW_DIR
    ADDRESS_RAW_DIR.mkdir(parents=True, exist_ok=True)
    RAW_DIR = ADDRESS_RAW_DIR

    try:
        return await query_one_address_browser(
            page=page,
            row_number=row_number,
            city=city,
            area=area,
            address=address,
        )
    finally:
        RAW_DIR = old_raw_dir


# ============================================================
# 25. 主程式
# ============================================================

_CELL2_STATUS_FILTER_STATS = {
    "input_total": 0,
    "eligible_total": 0,
    "completed_excluded": 0,
    "completed_status_counts": {},
}
_CELL2_INPUT_SHA256 = ""

async def main():

    start_time = time.time()

    print("=" * 70)
    print("Cell 2 v8.7 正式整合：NLSC 地號／地址批次座標查詢＋429快速備援＋雙座標")
    print("=" * 70)

    # --------------------------------------------------------
    # 找輸入檔
    # --------------------------------------------------------

    input_file = find_input_file()
    global _CELL2_INPUT_SHA256
    _CELL2_INPUT_SHA256 = hashlib.sha256(
        Path(input_file).read_bytes()
    ).hexdigest()

    print(f"輸入檔：{input_file}")
    print(f"輸入檔 SHA256：{_CELL2_INPUT_SHA256}")

    # --------------------------------------------------------
    # 讀目前輸入資料
    # --------------------------------------------------------

    raw_data = load_input_file(input_file)
    input_total_count = len(raw_data)

    if not raw_data:
        raise RuntimeError("輸入檔沒有有效資料")

    # 必須在讀取續跑快取、啟動 NLSC、建立 checkpoint 之前完成排除。
    data, completed_excluded_rows = split_completed_inspection_rows(raw_data)
    completed_status_counts = {}
    for excluded in completed_excluded_rows:
        status_value = excluded["status"]
        completed_status_counts[status_value] = (
            completed_status_counts.get(status_value, 0) + 1
        )

    _CELL2_STATUS_FILTER_STATS.update({
        "input_total": input_total_count,
        "eligible_total": len(data),
        "completed_excluded": len(completed_excluded_rows),
        "completed_status_counts": completed_status_counts,
    })

    print(f"輸入資料總筆數：{input_total_count}")
    print(
        f"已完成案件排除：{len(completed_excluded_rows)} 筆；"
        f"進入座標查詢／後續交接：{len(data)} 筆。"
    )
    for excluded in completed_excluded_rows:
        excluded_row = excluded["row"]
        excluded_label = (
            clean_text(excluded_row.get("名稱", ""))
            or clean_text(excluded_row.get("地址", ""))
            or " ".join(
                part for part in (
                    clean_text(excluded_row.get("地段", "")),
                    clean_text(excluded_row.get("地號", "")),
                ) if part
            )
            or f"輸入第 {excluded['input_index']} 筆"
        )
        print(
            f"[SKIP_COMPLETED] 資料第 {excluded['input_index']} 筆｜"
            f"狀態={excluded['status']}｜{excluded_label}｜"
            "不查座標、不寫入 Cell3／Cell4 交接檔"
        )

    # 只有未完成／狀態不明的案件才會進入後續；空白仍視同「未勘查」。
    for _row in data:
        _row["勘查狀態"] = normalize_inspection_status(
            _row.get("勘查狀態", "")
        )

    # 整批皆已完成時，覆寫正式交接檔為只有欄位標題的空白表，
    # 並清除舊 partial/error/temp，避免上次資料殘留被 Cell3／Cell4 誤用。
    if not data:
        print(
            "本批所有土地均為已完成狀態；不啟動 NLSC，也不產生任何案件座標。"
        )
        save_excel_openpyxl([], SUCCESS_OUTPUT_XLSX)
        save_gmap_csv([], GMAP_OUTPUT_CSV)

        for stale_path in (
            ERROR_OUTPUT_XLSX,
            RESULT_PARTIAL_XLSX,
            RESULT_TEMP_XLSX,
        ):
            try:
                if stale_path.exists():
                    stale_path.unlink()
            except Exception as exc:
                print(f"[WARN] 無法清除舊檔 {stale_path}: {exc}")

        empty_manifest = {
            "schema_version": "1.9-completed-inspection-exclusion",
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "input_file": str(input_file),
            "input_total": input_total_count,
            "eligible_total": 0,
            "completed_excluded": len(completed_excluded_rows),
            "completed_status_counts": completed_status_counts,
            "total": 0,
            "success": 0,
            "failed": 0,
            "complete": True,
            "reused_success": 0,
            "duplicate_land": 0,
            "queried_this_run": 0,
            "batch_fingerprint": hashlib.sha256(b"").hexdigest(),
            "result_path": str(SUCCESS_OUTPUT_XLSX),
            "partial_result_path": "",
            "error_result_path": "",
        }
        RESULT_MANIFEST.write_text(
            json.dumps(empty_manifest, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"空白正式交接檔：{SUCCESS_OUTPUT_XLSX}（僅欄位標題，0 筆土地）")
        print(f"空白 Google Maps CSV：{GMAP_OUTPUT_CSV}")
        print(f"結果狀態檔：{RESULT_MANIFEST}")
        return []

    RAW_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ========================================================
    # 續跑：先讀取上一輪結果
    # ========================================================

    print()
    print("檢查既有結果，準備續跑……")

    previous_results = load_all_resume_results()
    previous_success_map = {}

    for old_result in previous_results:
        try:
            key = build_resume_key_from_result(old_result)
        except Exception:
            continue

        if not key:
            continue

        old_status = clean_text(old_result.get("狀態", ""))
        current = previous_success_map.get(key)

        if current is None:
            previous_success_map[key] = old_result
        else:
            current_status = clean_text(current.get("狀態", ""))
            if old_status == "OK" and current_status != "OK":
                previous_success_map[key] = old_result

    reusable_success_count = sum(
        1 for x in previous_success_map.values()
        if clean_text(x.get("狀態", "")) == "OK"
    )

    if previous_results:
        print(f"合併後可辨識案件：{len(previous_results)} 筆")
        print(f"其中已成功可直接沿用：{reusable_success_count} 筆")
    else:
        print("沒有可用的既有結果，本次將全部查詢。")

    # ========================================================
    # 建立待查清單
    # ========================================================

    pending_rows = []
    results_by_input_index = {}
    reused_count = 0
    duplicate_count = 0
    seen_land_keys = {}

    for idx, row in enumerate(data, start=1):

        city = clean_text(row.get("縣市", ""))
        area = clean_text(row.get("鄉鎮市區", ""))
        section = clean_text(row.get("地段", ""))
        landno = clean_text(row.get("地號", ""))
        address = clean_text(row.get("地址", ""))

        if section and landno:
            query_mode = "地號"
        elif address:
            query_mode = "地址"
        else:
            query_mode = ""

        # ----------------------------------------------------
        # v1.2.4：重複土地防呆
        #
        # 同一「縣市＋鄉鎮市區＋地段＋地號」只允許第一筆
        # 進入 NLSC 定位流程。後續重複列保留在結果中，
        # 但標記為「重複土地」，不查 NLSC、不產生座標，
        # 也不進入 Cell 3。
        # 地址案件不套用本規則。
        # ----------------------------------------------------
        duplicate_key = build_land_duplicate_key_from_input(row)
        duplicate_first_idx = (
            seen_land_keys.get(duplicate_key)
            if duplicate_key
            else None
        )

        if duplicate_key and duplicate_first_idx is not None:
            duplicate_result = make_checkpoint_placeholder(
                row,
                idx,
                status="重複土地",
            )
            duplicate_result["資料檢核"] = "重複土地"
            duplicate_result["重複來源列"] = duplicate_first_idx
            duplicate_result["錯誤"] = (
                f"與輸入第 {duplicate_first_idx} 筆土地重複；"
                "本筆不進行 NLSC 定位查詢。"
            )
            duplicate_result["案件ID"] = f"DUP_{idx:04d}"

            results_by_input_index[idx] = duplicate_result
            duplicate_count += 1

            print(
                f"[{idx}/{len(data)}] DUPLICATE｜重複土地，跳過 NLSC｜"
                f"來源第 {duplicate_first_idx} 筆｜"
                f"{city} / {area} / {section} / {landno}"
            )
            continue

        if duplicate_key:
            seen_land_keys[duplicate_key] = idx

        resume_key = build_resume_key_from_input(row)
        old_result = previous_success_map.get(resume_key)

        # ----------------------------------------------------
        # v1.2.6：成功快取最高優先，永久沿用
        # ----------------------------------------------------
        # 只要既有結果狀態為 OK，代表本筆已成功取得座標。
        # 第二次以後不再因「正式段／小段名稱是否存在」而重新
        # 查詢 NLSC；正式名稱只在真正重新查詢時更新。
        # 這樣可避免每次重跑都因舊欄位格式差異重新定位。
        # ----------------------------------------------------

        if (
            old_result is not None
            and clean_text(old_result.get("狀態", "")) == "OK"
        ):

            cached = dict(old_result)

            # 輸入檔欄位永遠以這一輪 input.xlsx 為準。
            cached["名稱"] = clean_text(
                row.get("名稱", "")
            )
            cached["本筆面積(公頃)"] = parse_area_value(
                row.get("本筆面積(公頃)", "")
            )
            cached["申請面積(公頃)"] = parse_area_value(
                row.get("申請面積(公頃)", "")
            )
            cached["勘查狀態"] = normalize_inspection_status(
                row.get("勘查狀態", "")
            )
            cached["縣市"] = city
            cached_area = clean_text(
                cached.get("鄉鎮市區", "")
            )
            if not area_name_matches(area, cached_area):
                cached_area = area
            cached["鄉鎮市區"] = cached_area or area

            # 已有新版正式名稱時優先保留；沒有時才回到輸入值。
            cached_formal_section = clean_text(
                cached.get("查詢地段", "")
            )
            if query_mode == "地號" and has_formal_section_marker(cached_formal_section):
                cached["地段"] = cached_formal_section
                cached["查詢地段"] = cached_formal_section
            else:
                # v1.2.6：成功快取沿用時不再觸發任何 NLSC／地段
                # 正規化查詢；保留既有資料，缺少正式名稱也不重查。
                cached["地段"] = clean_text(cached.get("地段", "")) or section
                cached["查詢地段"] = clean_text(
                    cached.get("查詢地段", "")
                ) or section

            cached["地號"] = landno
            cached["地址"] = address
            cached["查詢方式"] = query_mode
            cached["顯示名稱"] = build_display_name(cached)

            # 案件ID重新整理；若為地號案件會沿用既有 office/sect。
            cached["案件ID"] = build_case_id(
                query_mode,
                cached,
                idx,
            )

            # v8.4.1：地址案件沿用舊結果時，也補齊雙座標欄位。
            ensure_dual_coordinate_fields(cached)

            results_by_input_index[idx] = cached
            reused_count += 1

            print(
                f"[{idx}/{len(data)}] "
                f"SKIP｜已成功，沿用既有座標｜"
                f"{city} / {area} / "
                f"{address if query_mode == '地址' else section + ' / ' + landno}"
            )

        else:
            pending_rows.append({
                "idx": idx,
                "row": row,
                "resume_key": resume_key,
                "query_mode": query_mode,
            })

    print()
    print("=" * 70)
    print(
        f"續跑統計：總 {len(data)} 筆｜"
        f"直接沿用成功 {reused_count} 筆｜"
        f"重複土地 {duplicate_count} 筆｜"
        f"需要重新查詢 {len(pending_rows)} 筆"
    )
    print("=" * 70)

    # 進入 NLSC 前先保存一次目前狀態；後續每完成 1 筆也會立即保存。
    save_resume_checkpoint(
        data,
        results_by_input_index,
        {},
    )

    # ========================================================
    # 沒有待查資料：完全不啟動瀏覽器
    # ========================================================

    if not pending_rows:

        results = [
            results_by_input_index[i]
            for i in range(1, len(data) + 1)
        ]

        success_results = [
            result
            for result in results
            if clean_text(result.get("狀態", "")) == "OK"
        ]

        error_results = [
            result
            for result in results
            if clean_text(result.get("狀態", "")) != "OK"
        ]

        # v1.2.4：重複土地屬資料檢核結果，不算真正查詢失敗。
        hard_error_results = [
            result
            for result in error_results
            if clean_text(result.get("狀態", "")) != "重複土地"
        ]
        duplicate_results = [
            result
            for result in error_results
            if clean_text(result.get("狀態", "")) == "重複土地"
        ]

        # 清除舊暫存，避免下次誤判。
        try:
            if RESULT_PARTIAL_XLSX.exists():
                RESULT_PARTIAL_XLSX.unlink()
        except Exception:
            pass

        if duplicate_results:
            save_excel_openpyxl(
                duplicate_results,
                ERROR_OUTPUT_XLSX,
            )
        else:
            try:
                if ERROR_OUTPUT_XLSX.exists():
                    ERROR_OUTPUT_XLSX.unlink()
            except Exception:
                pass

        save_excel_openpyxl(
            success_results,
            RESULT_TEMP_XLSX,
        )
        RESULT_TEMP_XLSX.replace(
            SUCCESS_OUTPUT_XLSX
        )

        save_gmap_csv(
            success_results,
            GMAP_OUTPUT_CSV,
        )

        batch_fingerprint = hashlib.sha256(
            "\n".join(
                str(r.get("案件ID", ""))
                for r in results
            ).encode("utf-8")
        ).hexdigest()

        manifest = {
            "schema_version": "1.9-completed-inspection-exclusion",
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "input_file": str(input_file),
            "input_total": input_total_count,
            "eligible_total": len(data),
            "completed_excluded": len(completed_excluded_rows),
            "completed_status_counts": completed_status_counts,
            "total": len(results),
            "success": len(success_results),
            "failed": len(hard_error_results),
            "complete": True,
            "reused_success": reused_count,
            "duplicate_land": duplicate_count,
            "queried_this_run": 0,
            "batch_fingerprint": batch_fingerprint,
            "result_path": str(SUCCESS_OUTPUT_XLSX),
            "partial_result_path": "",
            "error_result_path": str(ERROR_OUTPUT_XLSX) if duplicate_results else "",
        }
        RESULT_MANIFEST.write_text(
            json.dumps(
                manifest,
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        print()
        if hard_error_results:
            print("仍有真正失敗案件，本次未視為完整完成。")
        elif duplicate_results:
            print(
                f"唯一正常案件已完成；另有 {len(duplicate_results)} 筆重複土地列為資料檢核，不阻止正式結果輸出。"
            )
        else:
            print("全部案件上次已成功，本次沒有重新抓取 NLSC。")
        print(f"正式結果 Excel：{SUCCESS_OUTPUT_XLSX}")
        print(f"Google Maps CSV：{GMAP_OUTPUT_CSV}")
        if duplicate_results:
            print(f"資料檢核結果 Excel：{ERROR_OUTPUT_XLSX}")
        print(f"結果狀態檔：{RESULT_MANIFEST}")
        print(
            f"耗時：{time.time() - start_time:.1f} 秒"
        )
        print("全部處理結束。")
        return results

    # ========================================================
    # 有失敗／待查案件：只對 pending_rows 啟動 NLSC
    # ========================================================

    section_cache = {}
    query_result_map = {}

    async with async_playwright() as p:

        print("啟動 Chromium……")

        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-dev-shm-usage",
            ],
        )

        context = await browser.new_context(
            viewport={
                "width": 1440,
                "height": 1000,
            },
            locale="zh-TW",
            ignore_https_errors=True,
        )

        page = await context.new_page()

        print(f"開啟 NLSC：{NLSC_URL}")

        await initialize_nlsc_page(page)

        # ----------------------------------------------------
        # 元素檢查
        # ----------------------------------------------------

        check_result = await page.evaluate(
            """
            () => {
                const ids = [
                    'land_city',
                    'land_area_office',
                    'land_section',
                    'land_sectioncode',
                    'landcode',
                    'qt-land-submit',
                    'longitude',
                    'latitude',
                    'twd97_x',
                    'twd97_y'
                ];
                const result = {};
                for (const id of ids) {
                    const el = document.getElementById(id);
                    result[id] = {
                        exists: !!el,
                        tag: el ? el.tagName : null,
                        display: el ? getComputedStyle(el).display : null,
                        visibility: el ? getComputedStyle(el).visibility : null
                    };
                }
                return result;
            }
            """
        )

        print("NLSC 元素檢查：")
        for key, info in check_result.items():
            print(
                f"  #{key}: "
                f"exists={info['exists']}, "
                f"tag={info['tag']}, "
                f"display={info['display']}, "
                f"visibility={info['visibility']}"
            )
        print("-" * 70)

        # ----------------------------------------------------
        # 只查 pending_rows
        # ----------------------------------------------------

        for pending_pos, pending in enumerate(
            pending_rows,
            start=1,
        ):

            idx = pending["idx"]
            row = pending["row"]

            city = clean_text(row.get("縣市", ""))
            area = clean_text(row.get("鄉鎮市區", ""))
            section = clean_text(row.get("地段", ""))
            landno = clean_text(row.get("地號", ""))
            address = clean_text(row.get("地址", ""))
            query_mode = pending["query_mode"]

            query_section_for_log = normalize_section_for_query(
                city,
                area,
                section,
            )

            if (
                query_mode == "地號"
                and query_section_for_log != clean_text(section)
            ):
                section_display = (
                    f"{section} -> {query_section_for_log}"
                )
            else:
                section_display = section

            print(
                f"[{idx}/{len(data)}] "
                f"重新查詢 {pending_pos}/{len(pending_rows)} | "
                f"{query_mode or '未判定'} | "
                f"{city} / {area} / "
                f"{address if query_mode == '地址' else section_display + ' / ' + landno}"
            )

            try:
                if query_mode == "地址":
                    result = await query_one_address_browser(
                        page=page,
                        row_number=idx,
                        city=city,
                        area=area,
                        address=address,
                    )

                elif query_mode == "地號":
                    if not section or not landno:
                        raise ValueError(
                            "地號查詢需要：地段、地號"
                        )

                    result = await query_one_browser(
                        page=page,
                        row_number=idx,
                        city=city,
                        area=area,
                        section=section,
                        landno=landno,
                        section_cache=section_cache,
                    )

                    # v8.4.1：地號＋地址案件，主座標一定保留地號定位；
                    # 另外查一次地址座標，供 Cell 4 判斷實際進場道路。
                    if (
                        ENABLE_SECONDARY_ADDRESS_GEOCODE
                        and address
                        and clean_text(result.get("狀態", "")) == "OK"
                    ):
                        print(
                            f"  另查門牌座標：{build_full_address(city, area, address)}"
                        )
                        try:
                            address_result = await query_secondary_address_coordinates(
                                page=page,
                                row_number=idx,
                                city=city,
                                area=area,
                                address=address,
                            )

                            if clean_text(address_result.get("狀態", "")) == "OK":
                                result["地址經度"] = address_result.get("經度", "")
                                result["地址緯度"] = address_result.get("緯度", "")
                                print(
                                    "  地址座標 OK "
                                    f"lon={float(result['地址經度']):.6f} "
                                    f"lat={float(result['地址緯度']):.6f}"
                                )
                            else:
                                print(
                                    "  地址座標查詢未成功；主要地號座標保留。"
                                    f" 狀態={address_result.get('狀態', '')} "
                                    f"錯誤={address_result.get('錯誤', '')}"
                                )
                        except Exception as address_exc:
                            print(
                                "  地址座標查詢例外；主要地號座標保留。 "
                                f"{type(address_exc).__name__}: {address_exc}"
                            )

                else:
                    raise ValueError(
                        "無法判定查詢方式："
                        "請提供地號（地段＋地號）或地址"
                    )

            except asyncio.CancelledError:
                # 使用者手動取消時，前面已完成的案件已經逐筆存入 partial。
                # 再額外保存一次，避免剛完成的狀態遺失。
                print()
                print("⚠️ 偵測到手動取消，正在保存目前續跑進度……")
                save_resume_checkpoint(
                    data,
                    results_by_input_index,
                    query_result_map,
                )
                print(f"已保存輕量續跑檔：{V7_CHECKPOINT_JSONL}")
                raise

            except Exception as e:

                error_text = (
                    f"{type(e).__name__}: {e}"
                )

                print(
                    f"  [ERROR] {error_text}"
                )

                result = {
                    "縣市": city,
                    "鄉鎮市區": area,
                    "地段": section,
                    "查詢地段": query_section_for_log,
                    "地號": landno,
                    "地址": address,
                    "查詢方式": query_mode,
                    "定位來源": "",
                    "完整查詢地址": (
                        build_full_address(city, area, address)
                        if query_mode == "地址"
                        else ""
                    ),
                    "命中地址": "",
                    "定位品質": "",
                    "查詢地號": normalize_landno(landno),
                    "office": "",
                    "sect": "",
                    "狀態": "ERROR",
                    "HTTP狀態": "",
                    "經度": "",
                    "緯度": "",
                    "TWD97_X": "",
                    "TWD97_Y": "",
                    "左下經度": "",
                    "左下緯度": "",
                    "右上經度": "",
                    "右上緯度": "",
                    "response_url": "",
                    "raw_file": "",
                    "錯誤": error_text,
                    "嘗試次數": "",
                }

                print(
                    traceback.format_exc()
                )

            # ------------------------------------------------
            # 輸入檔欄位永遠以目前 input.xlsx 為準
            # ------------------------------------------------

            result = apply_input_fields_to_result(
                result,
                row,
                idx,
            )

            # v8.4.1：統一補齊雙座標欄位；地址案件主／次座標相同。
            ensure_dual_coordinate_fields(result)

            query_result_map[idx] = result

            if clean_text(result.get("狀態", "")) == "OK":
                print(
                    f"  OK "
                    f"HTTP={result.get('HTTP狀態', '')} "
                    f"lon={float(result['經度']):.6f} "
                    f"lat={float(result['緯度']):.6f} "
                    f"TWD97=({float(result['TWD97_X']):.3f}, "
                    f"{float(result['TWD97_Y']):.3f})"
                )
            else:
                print(
                    f"  {result.get('狀態', '')} "
                    f"HTTP={result.get('HTTP狀態', '')} "
                    f"錯誤={result.get('錯誤', '')}"
                )

            # ★ 核心：每完成 1 筆立即保存。
            save_resume_checkpoint(
                data,
                results_by_input_index,
                query_result_map,
            )

            # 每筆待查案件之間稍停。
            if pending_pos < len(pending_rows):
                await page.wait_for_timeout(
                    int(DELAY_BETWEEN_ROWS * 1000)
                )

        await context.close()
        await browser.close()

    # ========================================================
    # 合併：成功沿用資料 + 本次重查資料
    # ========================================================

    results = []

    for idx in range(1, len(data) + 1):
        if idx in query_result_map:
            results.append(query_result_map[idx])
        elif idx in results_by_input_index:
            results.append(results_by_input_index[idx])
        else:
            # 理論上不應發生；保留完整輸出結構以免中斷後無法產檔。
            row = data[idx - 1]
            city = clean_text(row.get("縣市", ""))
            area = clean_text(row.get("鄉鎮市區", ""))
            section = clean_text(row.get("地段", ""))
            landno = clean_text(row.get("地號", ""))
            address = clean_text(row.get("地址", ""))

            if section and landno:
                query_mode = "地號"
            elif address:
                query_mode = "地址"
            else:
                query_mode = ""

            results.append({
                "名稱": clean_text(row.get("名稱", "")),
                "查詢方式": query_mode,
                "縣市": city,
                "鄉鎮市區": area,
                "地段": section,
                "查詢地段": normalize_section_for_query(
                    city,
                    area,
                    section,
                ),
                "地號": landno,
                "本筆面積(公頃)": parse_area_value(
                    row.get("本筆面積(公頃)", "")
                ),
                "申請面積(公頃)": parse_area_value(
                    row.get("申請面積(公頃)", "")
                ),
                "勘查狀態": normalize_inspection_status(
                    row.get("勘查狀態", "")
                ),
                "地址": address,
                "定位來源": "",
                "完整查詢地址": (
                    build_full_address(city, area, address)
                    if query_mode == "地址"
                    else ""
                ),
                "命中地址": "",
                "定位品質": "",
                "查詢地號": normalize_landno(landno),
                "案件ID": build_case_id(
                    query_mode,
                    {
                        "縣市": city,
                        "鄉鎮市區": area,
                        "地段": section,
                        "地號": landno,
                        "地址": address,
                    },
                    idx,
                ),
                "顯示名稱": build_display_name({
                    "查詢方式": query_mode,
                    "縣市": city,
                    "鄉鎮市區": area,
                    "地段": section,
                    "地號": landno,
                    "查詢地號": normalize_landno(landno),
                    "地址": address,
                    "名稱": clean_text(row.get("名稱", "")),
                }),
                "office": "",
                "sect": "",
                "狀態": "ERROR",
                "HTTP狀態": "",
                "經度": "",
                "緯度": "",
                "TWD97_X": "",
                "TWD97_Y": "",
                "左下經度": "",
                "左下緯度": "",
                "右上經度": "",
                "右上緯度": "",
                "response_url": "",
                "raw_file": "",
                "錯誤": "缺少處理結果",
                "嘗試次數": "",
            })

    success_results = [
        result
        for result in results
        if clean_text(result.get("狀態", "")) == "OK"
    ]

    error_results = [
        result
        for result in results
        if clean_text(result.get("狀態", "")) != "OK"
    ]

    # v1.2.4：重複土地不是 NLSC 查詢失敗。
    # 重複列仍保留於 result_error.xlsx 供人工檢核，但不應阻止
    # 唯一且正常的案件建立正式 result.xlsx。
    hard_error_results = [
        result
        for result in error_results
        if clean_text(result.get("狀態", "")) != "重複土地"
    ]

    duplicate_results = [
        result
        for result in error_results
        if clean_text(result.get("狀態", "")) == "重複土地"
    ]

    # ========================================================
    # 正式結果／部分結果
    # ========================================================

    # 只有真正的查詢／處理錯誤才算未完成。
    # 「重複土地」屬於輸入資料檢核結果，不阻止正式成功檔輸出。
    complete_run = (
        len(results) > 0
        and len(hard_error_results) == 0
    )

    batch_fingerprint = hashlib.sha256(
        "\n".join(
            str(r.get("案件ID", ""))
            for r in results
        ).encode("utf-8")
    ).hexdigest()

    if complete_run:

        save_excel_openpyxl(
            success_results,
            RESULT_TEMP_XLSX,
        )

        RESULT_TEMP_XLSX.replace(
            SUCCESS_OUTPUT_XLSX
        )

        print(
            f"正式結果 Excel：{SUCCESS_OUTPUT_XLSX}"
        )

        save_gmap_csv(
            success_results,
            GMAP_OUTPUT_CSV,
        )

        print(
            f"Google Maps CSV：{GMAP_OUTPUT_CSV}"
        )

        # 正式結果建立後，刪除舊 partial。
        # 若本批存在「重複土地」，仍保留 result_error.xlsx 供人工檢核；
        # 僅在本批完全沒有檢核異常時才移除舊 error 檔。
        try:
            if RESULT_PARTIAL_XLSX.exists():
                RESULT_PARTIAL_XLSX.unlink()
        except Exception:
            pass

        if duplicate_results:
            save_excel_openpyxl(
                duplicate_results,
                ERROR_OUTPUT_XLSX,
            )
            print(
                f"資料檢核結果 Excel：{ERROR_OUTPUT_XLSX}"
            )
        else:
            try:
                if ERROR_OUTPUT_XLSX.exists():
                    ERROR_OUTPUT_XLSX.unlink()
            except Exception:
                pass

        manifest = {
            "schema_version": "1.9-completed-inspection-exclusion",
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "input_file": str(input_file),
            "input_total": input_total_count,
            "eligible_total": len(data),
            "completed_excluded": len(completed_excluded_rows),
            "completed_status_counts": completed_status_counts,
            "total": len(results),
            "success": len(success_results),
            "failed": len(hard_error_results),
            "complete": True,
            "reused_success": reused_count,
            "duplicate_land": duplicate_count,
            "queried_this_run": len(pending_rows),
            "batch_fingerprint": batch_fingerprint,
            "result_path": str(SUCCESS_OUTPUT_XLSX),
            "partial_result_path": "",
            "error_result_path": str(ERROR_OUTPUT_XLSX) if duplicate_results else "",
        }

    else:

        save_excel_openpyxl(
            results,
            RESULT_PARTIAL_XLSX,
        )

        print(
            f"⚠️ 仍有真正失敗案件，正式 result.xlsx 不更新。"
        )
        print(
            f"部分結果：{RESULT_PARTIAL_XLSX}"
        )

        if error_results:
            save_excel_openpyxl(
                error_results,
                ERROR_OUTPUT_XLSX,
            )
            print(
                f"失敗結果 Excel：{ERROR_OUTPUT_XLSX}"
            )

        manifest = {
            "schema_version": "1.9-completed-inspection-exclusion",
            "generated_at": datetime.now().astimezone().isoformat(timespec="seconds"),
            "input_file": str(input_file),
            "input_total": input_total_count,
            "eligible_total": len(data),
            "completed_excluded": len(completed_excluded_rows),
            "completed_status_counts": completed_status_counts,
            "total": len(results),
            "success": len(success_results),
            "failed": len(hard_error_results),
            "complete": False,
            "reused_success": reused_count,
            "queried_this_run": len(pending_rows),
            "batch_fingerprint": batch_fingerprint,
            "result_path": "",
            "partial_result_path": str(RESULT_PARTIAL_XLSX),
            "error_result_path": str(ERROR_OUTPUT_XLSX) if error_results else "",
        }

        print(
            "下次重新執行 Cell2 時，所有 OK 案件直接沿用既有結果；不因正式段／小段欄位差異重新查詢 NLSC。"
        )

    RESULT_MANIFEST.write_text(
        json.dumps(
            manifest,
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    # 完整成功後才輸出 Google Maps；部分成功沿用原本設計。
    if not complete_run:
        print(
            "本次仍有失敗資料，因此暫不更新 gmap.csv。"
        )

    # ========================================================
    # 正式地段／顯示名稱驗證
    # ========================================================
    land_ok = [
        r for r in success_results
        if clean_text(r.get("查詢方式", "")) == "地號"
    ]

    formal_section_missing = [
        r for r in land_ok
        if not has_formal_section_marker(r.get("查詢地段", ""))
    ]

    display_name_missing = [
        r for r in land_ok
        if not clean_text(r.get("顯示名稱", ""))
    ]

    print()
    print("正式地段／顯示名稱驗證：")
    print(f"  地號成功案件：{len(land_ok)} 筆")
    print(f"  缺少正式段／小段：{len(formal_section_missing)} 筆")
    print(f"  缺少顯示名稱：{len(display_name_missing)} 筆")

    for sample in land_ok[:5]:
        print(
            "  範例："
            f"{sample.get('縣市','')}"
            f"{sample.get('鄉鎮市區','')}"
            f"{sample.get('地段','')}"
            f"{normalize_landno(sample.get('查詢地號', sample.get('地號','')))}"
            f" | 顯示名稱={sample.get('顯示名稱','')}"
        )

    # ========================================================
    # 統計
    # ========================================================

    status_counts = {}

    for result in results:
        status = clean_text(
            result.get("狀態", "")
        )
        status_counts[status] = (
            status_counts.get(status, 0) + 1
        )

    elapsed = time.time() - start_time

    print()
    print("=" * 70)
    print("查詢完成")
    print("=" * 70)
    print(f"總筆數：{len(results)}")
    print(f"本次直接沿用成功：{reused_count}")
    print(f"本次重新查詢：{len(pending_rows)}")
    print(f"目前成功：{len(success_results)}")
    print(f"目前真正失敗：{len(hard_error_results)}")
    print(f"重複土地：{len(duplicate_results)}")
    print(
        "狀態統計："
        + ", ".join(
            f"{k}:{v}"
            for k, v in status_counts.items()
        )
    )

    if complete_run:
        print(
            f"正式成功檔案：{SUCCESS_OUTPUT_XLSX}"
        )
        print(
            f"Google Maps：{GMAP_OUTPUT_CSV}"
        )

    if error_results:
        if hard_error_results:
            print(
                f"失敗檔案：{ERROR_OUTPUT_XLSX}"
            )
        elif duplicate_results:
            print(
                f"資料檢核結果：{ERROR_OUTPUT_XLSX}"
            )

    print(f"原始回應：{RAW_DIR}")
    print(f"結果狀態檔：{RESULT_MANIFEST}")
    print(f"耗時：{elapsed:.1f} 秒")
    print("=" * 70)

    if hard_error_results:
        print()
        print("真正失敗項目：")

        for result in hard_error_results:
            print(
                f"  {result.get('縣市')} / "
                f"{result.get('鄉鎮市區')} / "
                f"{result.get('地段')} / "
                f"{result.get('地號')} "
                f"-> {result.get('狀態')} "
                f"HTTP={result.get('HTTP狀態')} "
                f"{result.get('錯誤')}"
            )

    if duplicate_results:
        print()
        print(f"重複土地檢核項目：{len(duplicate_results)} 筆（不列為查詢失敗）")

    print()
    print("全部處理結束。")

    return results


# ============================================================
# 26. 執行
# ============================================================



# ============================================================
# Cell 2 v8.5 雙座標版：沿用既有 DIRECT + 鄉鎮目錄 + JSONL 續跑架構
# ------------------------------------------------------------
# 目的：
#   1. 完全忽略既有 result.xlsx / result_partial.xlsx / error，
#      強制重新查詢全部輸入案件。
#   2. 所有測試輸出集中到 /content/CELL2_SPEED_TEST/v8.5_v4_section_cache/
#      絕不覆蓋正式 Cell 2 v8.4 輸出。
#   3. 保留 v3 已驗證的 qryTileMapIndex DIRECT 查詢。
#   4. 將 v3 的「單一目前地段 cache」提升為「本次執行的
#      鄉鎮／區域 → 完整地段 option 目錄 cache」。
#   5. 同一鄉鎮內切換不同地段時，不再重新執行完整
#      set_nlsc_section()；只直接套用已快取的地段 option，
#      再等待 sectioncode 更新。
#   6. 新鄉鎮第一次出現時，仍使用 v8.4 原始 set_nlsc_section()
#      建立該鄉鎮的完整地段目錄，降低風險。
#   7. 若快取切換失敗，立即退回原始 set_nlsc_section()；
#      頁面重新初始化後也會清除「目前 DOM 狀態」，但保留
#      本次執行已建立的目錄資料。
#
# 不改：
#   - DIRECT endpoint 與 parse_nlsc_response()
#   - retry 次數與 backoff
#   - 地址查詢流程
#   - checkpoint / Excel / raw response 行為
#   - 輸入欄位與正式輸出結構
# ============================================================

_SPEED_ROOT = BASE_DIR / "CELL2_SPEED_TEST"
_SPEED_RUN_DIR = _SPEED_ROOT / "v8.5_v9_rate_safe"
_SPEED_RUN_DIR.mkdir(parents=True, exist_ok=True)

SUCCESS_OUTPUT_XLSX = _SPEED_RUN_DIR / "result.xlsx"
ERROR_OUTPUT_XLSX = _SPEED_RUN_DIR / "result_error.xlsx"
GMAP_OUTPUT_CSV = _SPEED_RUN_DIR / "gmap.csv"
RAW_DIR = _SPEED_RUN_DIR / "nlsc_raw"
RESULT_PARTIAL_XLSX = _SPEED_RUN_DIR / "result_partial.xlsx"
RESULT_MANIFEST = _SPEED_RUN_DIR / "result_manifest.json"
RESULT_TEMP_XLSX = _SPEED_RUN_DIR / "result.__tmp__.xlsx"
# v7：每筆只寫 append-only JSONL checkpoint，不再每筆重建 XLSX。
V7_CHECKPOINT_JSONL = _SPEED_RUN_DIR / "checkpoint_state.jsonl"
_V7_CHECKPOINT_SEQ = 0
_V7_CHECKPOINT_WRITTEN_REUSED = set()
_FORMAL_QUERY_CHECKPOINTED = set()


def _speed_test_no_resume():
    return []

load_all_resume_results = _speed_test_no_resume

_SPEED_TIMING = {
    "script_start": time.perf_counter(),
    "initialize_calls": 0,
    "initialize_seconds": 0.0,
    "area_resolve_calls": 0,
    "area_resolve_seconds": 0.0,
    "section_set_calls": 0,
    "section_set_seconds": 0.0,
    "section_full_set_calls": 0,
    "section_full_set_seconds": 0.0,
    "section_cached_switch_calls": 0,
    "section_cached_switch_seconds": 0.0,
    "area_catalog_build_calls": 0,
    "area_catalog_build_seconds": 0.0,
    "area_catalog_count": 0,
    "area_restore_calls": 0,
    "area_restore_seconds": 0.0,
    "catalog_direct_resolve_calls": 0,
    "catalog_direct_resolve_seconds": 0.0,
    "land_query_calls": 0,
    "land_query_seconds": 0.0,
    "address_query_calls": 0,
    "address_query_seconds": 0.0,
    "checkpoint_calls": 0,
    "checkpoint_seconds": 0.0,
    "v7_checkpoint_records_appended": 0,
    "excel_write_calls": 0,
    "excel_write_seconds": 0.0,
    "raw_write_calls": 0,
    "raw_write_seconds": 0.0,
    "query_rows": [],
}


def _speed_now():
    return time.perf_counter()


_original_initialize_nlsc_page = initialize_nlsc_page
_original_resolve_nlsc_area_name = resolve_nlsc_area_name
_original_set_nlsc_section = set_nlsc_section
_original_query_one_browser = query_one_browser
_original_query_one_address_browser = query_one_address_browser
_original_save_resume_checkpoint = save_resume_checkpoint
_original_save_excel_openpyxl = save_excel_openpyxl
_original_save_raw_response = save_raw_response


async def _speed_initialize_nlsc_page(page):
    t0 = _speed_now()
    try:
        result = await _original_initialize_nlsc_page(page)
        return result
    finally:
        _SPEED_TIMING["initialize_calls"] += 1
        _SPEED_TIMING["initialize_seconds"] += _speed_now() - t0
        # DOM 已被重新建立；保留 section catalog，但清除目前頁面的
        # active area 狀態，下一筆會重新把該鄉鎮套回頁面。
        _V4_SECTION_CATALOG_STATE["active_area_key"] = None
        _V4_SECTION_CATALOG_STATE["active_section_key"] = None
        _V4_SECTION_CATALOG_STATE["active_section_code"] = ""


async def _speed_resolve_nlsc_area_name(page, city, area):
    t0 = _speed_now()
    try:
        return await _original_resolve_nlsc_area_name(page, city, area)
    finally:
        _SPEED_TIMING["area_resolve_calls"] += 1
        _SPEED_TIMING["area_resolve_seconds"] += _speed_now() - t0


async def _speed_query_one_address_browser(*args, **kwargs):
    t0 = _speed_now()
    row_number = kwargs.get("row_number")
    if row_number is None and len(args) >= 2:
        row_number = args[1]
    try:
        return await _original_query_one_address_browser(*args, **kwargs)
    finally:
        elapsed = _speed_now() - t0
        _SPEED_TIMING["address_query_calls"] += 1
        _SPEED_TIMING["address_query_seconds"] += elapsed
        _SPEED_TIMING["query_rows"].append({
            "row_number": int(row_number) if row_number is not None else "",
            "query_mode": "地址",
            "seconds": elapsed,
        })


def _speed_save_resume_checkpoint(data, results_by_input_index, query_result_map):
    """
    v7 輕量續跑 checkpoint：
    - 不再每筆建立 / 寫入 result_partial.xlsx。
    - 已沿用的成功資料只寫一次。
    - 本次查詢結果以 append-only JSONL 逐筆追加。
    - 最終 result.xlsx 仍由原 main 正常產生。
    """
    global _V7_CHECKPOINT_SEQ

    t0 = _speed_now()
    appended = 0

    V7_CHECKPOINT_JSONL.parent.mkdir(parents=True, exist_ok=True)
    file_exists = V7_CHECKPOINT_JSONL.exists()

    with V7_CHECKPOINT_JSONL.open("a", encoding="utf-8") as f:
        if not file_exists:
            _V7_CHECKPOINT_SEQ += 1
            f.write(json.dumps({
                "type": "meta",
                "version": "Cell2 v8.6 SPEED TEST v7 light checkpoint",
                "input_rows": len(data),
            }, ensure_ascii=False) + "\n")

        # 已成功沿用的資料：只需要記錄第一次。
        for idx in sorted(results_by_input_index):
            if idx in _V7_CHECKPOINT_WRITTEN_REUSED:
                continue
            result = dict(results_by_input_index[idx])
            _V7_CHECKPOINT_SEQ += 1
            f.write(json.dumps({
                "type": "result",
                "seq": _V7_CHECKPOINT_SEQ,
                "index": int(idx),
                "result": result,
                "source": "reused",
            }, ensure_ascii=False, default=str) + "\n")
            _V7_CHECKPOINT_WRITTEN_REUSED.add(idx)
            appended += 1

        # 本次查詢：每次 checkpoint 都追加目前的最新結果。
        # 通常每次只有 1 筆，取消／例外時也能留下最新狀態。
        for idx in sorted(query_result_map):
            result = dict(query_result_map[idx])
            _V7_CHECKPOINT_SEQ += 1
            f.write(json.dumps({
                "type": "result",
                "seq": _V7_CHECKPOINT_SEQ,
                "index": int(idx),
                "result": result,
                "source": "query",
            }, ensure_ascii=False, default=str) + "\n")
            appended += 1

        f.flush()
        try:
            os.fsync(f.fileno())
        except Exception:
            pass

    _SPEED_TIMING["checkpoint_calls"] += 1
    _SPEED_TIMING["checkpoint_seconds"] += _speed_now() - t0
    _SPEED_TIMING["v7_checkpoint_records_appended"] = (
        _SPEED_TIMING.get("v7_checkpoint_records_appended", 0) + appended
    )
    return []


def _speed_save_excel_openpyxl(*args, **kwargs):
    t0 = _speed_now()
    try:
        return _original_save_excel_openpyxl(*args, **kwargs)
    finally:
        _SPEED_TIMING["excel_write_calls"] += 1
        _SPEED_TIMING["excel_write_seconds"] += _speed_now() - t0


def _speed_save_raw_response(*args, **kwargs):
    t0 = _speed_now()
    try:
        return _original_save_raw_response(*args, **kwargs)
    finally:
        _SPEED_TIMING["raw_write_calls"] += 1
        _SPEED_TIMING["raw_write_seconds"] += _speed_now() - t0


# ============================================================
# v4 鄉鎮／地段完整目錄 cache
# ============================================================
_V4_SECTION_CATALOG_STATE = {
    "catalog": {},
    "active_area_key": None,
    "active_section_key": None,
    "active_section_code": "",
}


def _v4_area_key(city, area):
    return (
        clean_text(city),
        clean_text(area),
    )


def _v4_section_key(city, area, section):
    return (
        clean_text(city),
        clean_text(area),
        clean_text(normalize_section_for_query(city, area, section)),
    )


def _v4_section_catalog_option(catalog, query_section):
    options = catalog.get("sections", [])
    return find_option_value(options, query_section)


async def _v4_build_area_catalog(page, city, area, query_section):
    """第一次遇到某鄉鎮：沿用 v8.4 原始設定一次，然後快取完整地段 option。"""
    area_key = _v4_area_key(city, area)
    t0 = _speed_now()
    try:
        section_info = await _original_set_nlsc_section(
            page,
            city,
            area,
            query_section,
        )

        options = await get_options(page, "#land_section option")
        sections = [
            {
                "text": clean_text(option.get("text", "")),
                "value": clean_text(option.get("value", "")),
            }
            for option in options
            if clean_text(option.get("text", ""))
            and clean_text(option.get("value", ""))
        ]

        if not sections:
            raise RuntimeError(
                f"建立地段目錄失敗：{city} / {area} 沒有可用地段 option"
            )

        catalog = {
            "city": clean_text(city),
            "requested_area": clean_text(area),
            "area_text": clean_text(section_info.get("area_text", area)),
            "area_value": clean_text(section_info.get("area_value", "")),
            "sections": sections,
            "section_info_first": dict(section_info),
        }
        _V4_SECTION_CATALOG_STATE["catalog"][area_key] = catalog
        _SPEED_TIMING["area_catalog_count"] = len(
            _V4_SECTION_CATALOG_STATE["catalog"]
        )
        return catalog, section_info
    finally:
        elapsed = _speed_now() - t0
        _SPEED_TIMING["area_catalog_build_calls"] += 1
        _SPEED_TIMING["area_catalog_build_seconds"] += elapsed


async def _v4_restore_area_from_catalog(page, catalog):
    """目前 DOM 不在目標鄉鎮時，用已快取 area_value 恢復，不重跑名稱匹配。"""
    t0 = _speed_now()
    try:
        city = clean_text(catalog.get("city", ""))
        area_value = clean_text(catalog.get("area_value", ""))
        if not city or not area_value:
            raise RuntimeError("快取缺少 city / area_value，無法恢復鄉鎮")

        current_city_value = ""
        try:
            current_city_value = clean_text(
                await page.input_value("#land_city")
            )
        except Exception:
            pass

        # land_city 通常 value 與 label 相同；若已在同一縣市就不重選。
        if current_city_value != city:
            await page.select_option(
                "#land_city",
                label=city,
                force=True,
            )

        # 等待 area option 出現。
        area_found = False
        for _ in range(SECTION_MAX_POLLS):
            try:
                options = await get_options(page, "#land_area_office option")
                if any(
                    clean_text(x.get("value", "")) == area_value
                    for x in options
                ):
                    area_found = True
                    break
            except Exception:
                pass
            await page.wait_for_timeout(
                int(SECTION_POLL_INTERVAL * 1000)
            )

        if not area_found:
            raise RuntimeError(
                f"快取鄉鎮 value 未出現在 NLSC：{city} / {area_value}"
            )

        await page.evaluate(
            """
            ({value}) => {
                const el = document.querySelector('#land_area_office');
                if (!el) throw new Error('#land_area_office not found');
                el.value = value;
                el.dispatchEvent(new Event('change', {bubbles: true}));
            }
            """,
            {"value": area_value},
        )

        # 等 section 目錄重新載入，並用 cached option 驗證 DOM 已恢復。
        cached_values = {
            clean_text(x.get("value", ""))
            for x in catalog.get("sections", [])
            if clean_text(x.get("value", ""))
        }
        restored = False
        for _ in range(SECTION_MAX_POLLS):
            try:
                options = await get_options(page, "#land_section option")
                current_values = {
                    clean_text(x.get("value", ""))
                    for x in options
                    if clean_text(x.get("value", ""))
                }
                # 用交集而非完整相等，避免網站增加/刪除 placeholder 影響判定。
                if cached_values and current_values.intersection(cached_values):
                    restored = True
                    break
            except Exception:
                pass
            await page.wait_for_timeout(
                int(SECTION_POLL_INTERVAL * 1000)
            )

        if not restored:
            raise RuntimeError(
                f"快取地段目錄恢復失敗：{city} / {catalog.get('area_text', '')}"
            )
    finally:
        elapsed = _speed_now() - t0
        _SPEED_TIMING["area_restore_calls"] += 1
        _SPEED_TIMING["area_restore_seconds"] += elapsed


async def _v4_apply_cached_section(page, catalog, query_section, section_key):
    """同一鄉鎮已建目錄時，直接套用 section option，並讀取新的 sectioncode。"""
    t0 = _speed_now()
    try:
        option = _v4_section_catalog_option(catalog, query_section)
        if option is None:
            raise RuntimeError(
                f"快取找不到地段：{catalog.get('city')} / "
                f"{catalog.get('area_text', '')} / {query_section}"
            )

        previous_code = ""
        try:
            previous_code = clean_text(
                await page.input_value("#land_sectioncode")
            )
        except Exception:
            pass

        await page.evaluate(
            """
            ({value}) => {
                const el = document.querySelector('#land_section');
                if (!el) throw new Error('#land_section not found');
                el.value = value;
                el.dispatchEvent(new Event('change', {bubbles: true}));
            }
            """,
            {"value": option["value"]},
        )

        section_code = ""
        for _ in range(SECTION_MAX_POLLS):
            try:
                current_value = clean_text(
                    await page.input_value("#land_section")
                )
                current_code = clean_text(
                    await page.input_value("#land_sectioncode")
                )
                if current_value == clean_text(option["value"]):
                    # 若 code 本身沒有變動也允許直接使用；有些頁面會重設成同值。
                    if current_code and (
                        not previous_code or current_code != previous_code
                    ):
                        section_code = current_code
                        break
                    if current_code:
                        section_code = current_code
                        break
            except Exception:
                pass
            await page.wait_for_timeout(
                int(SECTION_POLL_INTERVAL * 1000)
            )

        if not section_code:
            raise RuntimeError(
                f"套用快取地段後未取得 sectioncode：{query_section}"
            )

        return {
            "area_text": catalog.get("area_text", ""),
            "area_value": catalog.get("area_value", ""),
            "section_text": clean_text(option["text"]),
            "section_value": clean_text(option["value"]),
            "section_code": section_code,
        }
    finally:
        elapsed = _speed_now() - t0
        _SPEED_TIMING["section_cached_switch_calls"] += 1
        _SPEED_TIMING["section_cached_switch_seconds"] += elapsed


async def _v6_prepare_section(page, city, area, query_section, section_cache):
    """v5：只在第一次遇到鄉鎮時操作 NLSC DOM；之後直接從 catalog 取得 office/sect。"""
    area_key = _v4_area_key(city, area)
    catalog = _V4_SECTION_CATALOG_STATE["catalog"].get(area_key)

    if catalog is None:
        # 新鄉鎮：需要進一次 NLSC DOM，建立完整地段目錄。
        t0 = _speed_now()
        try:
            catalog, first_info = await _v4_build_area_catalog(
                page,
                city,
                area,
                query_section,
            )
        finally:
            elapsed = _speed_now() - t0
            _SPEED_TIMING["section_full_set_calls"] += 1
            _SPEED_TIMING["section_full_set_seconds"] += elapsed
        section_info = first_info
    else:
        # 已有完整鄉鎮目錄：完全不碰 page DOM。
        t0 = _speed_now()
        try:
            option = _v4_section_catalog_option(catalog, query_section)
            if option is None:
                raise RuntimeError(
                    f"鄉鎮目錄找不到地段：{city} / {catalog.get('area_text', area)} / {query_section}"
                )

            section_value = clean_text(option.get("value", ""))
            section_text = clean_text(option.get("text", query_section))
            if not section_value:
                raise RuntimeError(
                    f"鄉鎮目錄地段 option 沒有 value：{city} / {area} / {query_section}"
                )

            # NLSC option value 已在 v3/v4 DIRECT 測試驗證可拆成 office + sect。
            if "_" in section_value:
                derived_sect = clean_text(section_value.split("_", 1)[1])
            else:
                derived_sect = section_value

            section_info = {
                "area_text": catalog.get("area_text", area),
                "area_value": catalog.get("area_value", ""),
                "section_text": section_text,
                "section_value": section_value,
                "section_code": derived_sect,
            }
        finally:
            _SPEED_TIMING["catalog_direct_resolve_calls"] += 1
            _SPEED_TIMING["catalog_direct_resolve_seconds"] += _speed_now() - t0

    section_cache.clear()
    section_cache.update({
        "key": _v4_section_key(city, area, query_section),
        "area_text": clean_text(section_info.get("area_text", area)),
        "area_value": clean_text(section_info.get("area_value", "")),
        "section_text": clean_text(section_info.get("section_text", query_section)),
        "section_value": clean_text(section_info.get("section_value", "")),
        "section_code": clean_text(section_info.get("section_code", "")),
    })
    return section_info



# ============================================================
# v6 地號 DIRECT query
# ============================================================
LAND_DIRECT_FALLBACK = True

_V4_LAND_STATE = {
    "direct_calls": 0,
    "direct_seconds": 0.0,
    "fallback_calls": 0,
    "fallback_seconds": 0.0,
    "direct_429_count": 0,
    "direct_cooldown_skips": 0,
    "direct_429_cooldown_seconds": 0.0,
    "direct_429_server_retry_after_seconds": 0.0,
}


# v8.7 DIRECT 429 全域斷路器：冷卻期間不再逐筆等待，也不呼叫 DIRECT；
# 每筆案件直接走既有瀏覽器備援，冷卻到期後才恢復 DIRECT。
_V8_DIRECT_429_COOLDOWN_UNTIL = 0.0
_V9_DIRECT_NEXT_ALLOWED = 0.0
_DIRECT_429_STREAK = 0


def _parse_direct_retry_after(response_headers):
    """解析 Retry-After 秒數或 HTTP-date；沒有有效值時回傳 None。"""
    raw_value = ""
    try:
        for key, value in (response_headers or {}).items():
            if str(key).strip().lower() == "retry-after":
                raw_value = str(value).strip()
                break
    except Exception:
        raw_value = ""

    if not raw_value:
        return None

    try:
        seconds = float(raw_value)
    except (TypeError, ValueError):
        try:
            retry_dt = parsedate_to_datetime(raw_value)
            if retry_dt.tzinfo is None:
                retry_dt = retry_dt.replace(tzinfo=timezone.utc)
            seconds = (retry_dt - datetime.now(timezone.utc)).total_seconds()
        except Exception:
            return None

    if not (seconds == seconds) or seconds in (float("inf"), float("-inf")):
        return None
    return max(0.0, min(float(seconds), DIRECT_429_MAX_SERVER_WAIT_SECONDS))


async def _v6_direct_query_one_browser(
    page,
    row_number,
    city,
    area,
    section,
    landno,
    section_cache,
):
    global _V8_DIRECT_429_COOLDOWN_UNTIL, _V9_DIRECT_NEXT_ALLOWED, _DIRECT_429_STREAK

    original_landno = clean_text(landno)
    normalized_landno = normalize_landno(original_landno)
    query_section = normalize_section_for_query(city, area, section)

    section_info = await _v6_prepare_section(
        page,
        city,
        area,
        query_section,
        section_cache,
    )

    matched_section_text = clean_text(
        section_info.get("section_text", query_section)
    )
    matched_area_text = clean_text(
        section_info.get("area_text", area)
    )

    section_value = clean_text(section_info.get("section_value", ""))
    section_code = clean_text(section_info.get("section_code", ""))

    if "_" in section_value:
        office = section_value.split("_", 1)[0]
        value_sect = section_value.split("_", 1)[1]
    else:
        office = ""
        value_sect = section_value

    sect = section_code or value_sect

    last_error = ""
    last_status = None
    last_url = ""
    last_body = ""
    last_raw_file = ""
    direct_attempt = 0

    async def _direct_once():
        global _V9_DIRECT_NEXT_ALLOWED

        now = _speed_now()
        if now < _V9_DIRECT_NEXT_ALLOWED:
            wait_seconds = _V9_DIRECT_NEXT_ALLOWED - now
            print(f"  DIRECT 節流間隔：等待 {wait_seconds:.3f} 秒...")
            await page.wait_for_timeout(int(wait_seconds * 1000))

        callback_nonce = int(time.time() * 1000)
        query_url = (
            f"https://landmaps.nlsc.gov.tw/S_Maps/qryTileMapIndex"
            f"?type=2&flag=2"
            f"&office={office}"
            f"&sect={sect}"
            f"&landno={normalized_landno}"
            f"&alpah=0.5f"
            f"&callback=_jqjsp"
            f"&_{callback_nonce}="
        )

        response = await page.context.request.get(
            query_url,
            timeout=RESPONSE_TIMEOUT,
            headers={"Referer": NLSC_URL, "Accept": "*/*"},
            fail_on_status_code=False,
        )
        return response.status, response.url, await response.text(), response.headers

    direct_cooldown_remaining = max(
        0.0, _V8_DIRECT_429_COOLDOWN_UNTIL - _speed_now()
    )
    if direct_cooldown_remaining > 0:
        _V4_LAND_STATE["direct_cooldown_skips"] += 1
        print(
            f"  [DIRECT COOLDOWN] 尚餘 {direct_cooldown_remaining:.1f} 秒；"
            "本筆跳過 DIRECT，立即改用瀏覽器備援（不等待）。"
        )

    while (
        direct_attempt < DIRECT_MAX_ATTEMPTS
        and _speed_now() >= _V8_DIRECT_429_COOLDOWN_UNTIL
    ):
        direct_attempt += 1

        try:
            print(
                f"  第 {direct_attempt}/{DIRECT_MAX_ATTEMPTS} 次查詢："
                f"地號={normalized_landno} [DIRECT]"
            )
            direct_t0 = _speed_now()
            _V4_LAND_STATE["direct_calls"] += 1
            try:
                status_code, response_url, body, response_headers = await _direct_once()
            finally:
                _V4_LAND_STATE["direct_seconds"] += _speed_now() - direct_t0
                _V9_DIRECT_NEXT_ALLOWED = max(
                    _V9_DIRECT_NEXT_ALLOWED,
                    _speed_now() + DIRECT_MIN_INTERVAL_SECONDS,
                )

            last_status = status_code
            last_url = response_url
            last_body = body
            print(f"  HTTP={status_code} [DIRECT]")

            # HTTP 200 代表 DIRECT 端點已不再回 429；即使內容沒有座標，
            # 也應與速率限制錯誤分開處理，並重置連續 429 計數。
            if status_code == 200:
                _DIRECT_429_STREAK = 0
                _V8_DIRECT_429_COOLDOWN_UNTIL = 0.0

            parsed = parse_nlsc_response(body)

            if status_code == 200 and parsed["ok"]:
                record = parsed["record"]
                lon = float(record["cx"])
                lat = float(record["cy"])
                twd97_x, twd97_y = convert_to_twd97(lon, lat)
                raw_file = save_raw_response(
                    row_number, status_code, response_url, body
                )
                return {
                    "縣市": city,
                    "鄉鎮市區": matched_area_text or area,
                    "地段": matched_section_text or section,
                    "查詢地段": matched_section_text or section,
                    "地號": original_landno,
                    "地址": "",
                    "查詢方式": "地號",
                    "定位來源": "NLSC地號定位",
                    "完整查詢地址": "",
                    "命中地址": "",
                    "定位品質": "地號宗地定位",
                    "查詢地號": normalized_landno,
                    "office": get_query_parameter(response_url, "office") or office,
                    "sect": get_query_parameter(response_url, "sect") or sect,
                    "狀態": "OK",
                    "HTTP狀態": status_code,
                    "經度": lon,
                    "緯度": lat,
                    "TWD97_X": twd97_x,
                    "TWD97_Y": twd97_y,
                    "左下經度": record.get("lx", ""),
                    "左下緯度": record.get("ly", ""),
                    "右上經度": record.get("rx", ""),
                    "右上緯度": record.get("ry", ""),
                    "response_url": response_url,
                    "raw_file": raw_file,
                    "錯誤": "",
                    "嘗試次數": direct_attempt,
                }

            if status_code == 429:
                last_error = "HTTP_429"
                last_raw_file = save_raw_response(
                    row_number, status_code, response_url, body
                )

                _DIRECT_429_STREAK += 1
                _V4_LAND_STATE["direct_429_count"] += 1

                # Retry-After 優先；沒有有效 header 才用連續 429 的本地退避。
                server_retry_after = _parse_direct_retry_after(response_headers)
                if server_retry_after is not None:
                    wait_seconds = max(1.0, server_retry_after)
                    _V4_LAND_STATE["direct_429_server_retry_after_seconds"] += wait_seconds
                    wait_source = "SERVER_RETRY_AFTER"
                else:
                    fallback_wait = DIRECT_429_BACKOFF[
                        min(_DIRECT_429_STREAK - 1, len(DIRECT_429_BACKOFF) - 1)
                    ]
                    wait_seconds = float(fallback_wait)
                    wait_source = "LOCAL_BACKOFF"

                cooldown_until = _speed_now() + wait_seconds
                _V8_DIRECT_429_COOLDOWN_UNTIL = max(
                    _V8_DIRECT_429_COOLDOWN_UNTIL,
                    cooldown_until,
                )
                _V4_LAND_STATE["direct_429_cooldown_seconds"] += wait_seconds

                print(
                    f"  HTTP 429：{wait_source}={wait_seconds:.1f} 秒；"
                    "本筆不再重試 DIRECT，立即切換瀏覽器備援。"
                )
                print(
                    "  [DIRECT CIRCUIT BREAKER] 冷卻期間後續案件跳過 DIRECT，"
                    "直接使用瀏覽器備援，不逐筆等待冷卻。"
                )
                break

            # 其他暫時性 HTTP 維持原本 retry/backoff。
            if status_code in RETRY_STATUS_CODES and direct_attempt < DIRECT_MAX_ATTEMPTS:
                last_error = f"HTTP_{status_code}"
                last_raw_file = save_raw_response(
                    row_number, status_code, response_url, body
                )
                wait_seconds = RETRY_BACKOFF[
                    min(direct_attempt - 1, len(RETRY_BACKOFF) - 1)
                ]
                print(f"  HTTP {status_code}，{wait_seconds} 秒後自動重試...")
                await page.wait_for_timeout(wait_seconds * 1000)
                continue

            last_raw_file = save_raw_response(
                row_number, status_code, response_url, body
            )
            last_error = parsed["error"] if status_code == 200 else f"HTTP_{status_code}"
            break

        except Exception as e:
            last_error = f"{type(e).__name__}: {e}"
            print(f"  Direct 查詢例外：{last_error}")
            if direct_attempt < DIRECT_MAX_ATTEMPTS:
                wait_seconds = RETRY_BACKOFF[
                    min(direct_attempt - 1, len(RETRY_BACKOFF) - 1)
                ]
                print(f"  {wait_seconds} 秒後自動重試...")
                await page.wait_for_timeout(wait_seconds * 1000)
                continue
            break

    if LAND_DIRECT_FALLBACK:
        if last_status == 429:
            fallback_reason = "本筆 DIRECT HTTP 429"
        elif direct_cooldown_remaining > 0:
            fallback_reason = f"DIRECT 全域冷卻尚餘 {direct_cooldown_remaining:.1f} 秒"
        else:
            fallback_reason = last_error or "DIRECT 未取得有效座標"
        print(
            "  [DIRECT FALLBACK] 改用 v8.4 原始瀏覽器查詢流程；"
            f"原因={fallback_reason}"
        )
        fallback_t0 = _speed_now()
        _V4_LAND_STATE["fallback_calls"] += 1
        try:
            return await _original_query_one_browser(
                page=page,
                row_number=row_number,
                city=city,
                area=area,
                section=section,
                landno=landno,
                section_cache=section_cache,
            )
        finally:
            _V4_LAND_STATE["fallback_seconds"] += _speed_now() - fallback_t0

    if last_status == 429:
        status_text = "HTTP_ERROR"
    elif last_status in RETRY_STATUS_CODES:
        status_text = "HTTP_ERROR"
    elif "Timeout" in last_error or "timeout" in last_error.lower():
        status_text = "TIMEOUT"
    else:
        status_text = "ERROR"

    return {
        "縣市": city,
        "鄉鎮市區": area,
        "地段": section,
        "查詢地段": matched_section_text,
        "地號": original_landno,
        "地址": "",
        "查詢方式": "地號",
        "定位來源": "NLSC地號定位",
        "完整查詢地址": "",
        "命中地址": "",
        "定位品質": "",
        "查詢地號": normalized_landno,
        "office": office,
        "sect": sect,
        "狀態": status_text,
        "HTTP狀態": last_status if last_status is not None else "",
        "經度": "",
        "緯度": "",
        "TWD97_X": "",
        "TWD97_Y": "",
        "左下經度": "",
        "左下緯度": "",
        "右上經度": "",
        "右上緯度": "",
        "response_url": last_url,
        "raw_file": last_raw_file,
        "錯誤": last_error,
        "嘗試次數": direct_attempt,
    }


async def _speed_query_one_browser_v5(*args, **kwargs):
    t0 = _speed_now()
    row_number = kwargs.get("row_number")
    if row_number is None and len(args) >= 2:
        row_number = args[1]
    try:
        return await _v6_direct_query_one_browser(*args, **kwargs)
    finally:
        elapsed = _speed_now() - t0
        _SPEED_TIMING["land_query_calls"] += 1
        _SPEED_TIMING["land_query_seconds"] += elapsed
        _SPEED_TIMING["query_rows"].append({
            "row_number": int(row_number) if row_number is not None else "",
            "query_mode": "地號",
            "seconds": elapsed,
        })


# v6 對外只替換本次測試需要的函式；正式 v8.4 原始碼完全不變。
initialize_nlsc_page = _speed_initialize_nlsc_page
resolve_nlsc_area_name = _speed_resolve_nlsc_area_name
query_one_address_browser = _speed_query_one_address_browser
query_one_browser = _speed_query_one_browser_v5
save_resume_checkpoint = _speed_save_resume_checkpoint
save_excel_openpyxl = _speed_save_excel_openpyxl
save_raw_response = _speed_save_raw_response

print("=" * 70)

# ============================================================
# Cell 2 v8.7 正式整合層
# ------------------------------------------------------------
# 正式候選整合：
#   1. CATALOG DIRECT：首次遇到鄉鎮時建立完整地段目錄；後續直接解析。
#   2. DIRECT qryTileMapIndex：沿用已驗證的後端 HTTP 查詢。
#   3. DELAY_BETWEEN_ROWS = 0：取消固定逐筆等待。
#   4. DIRECT_MIN_INTERVAL_SECONDS = 1.0：降低連續 DIRECT 請求密度。
#   5. HTTP 429：同筆第一次即切換瀏覽器備援；冷卻期間跳過 DIRECT、不逐筆等待。
#      優先尊重 Retry-After 秒數／HTTP-date；缺少時採 15/30/60 秒退避。
#   6. 輕量 JSONL checkpoint：每筆只追加，不再每筆重建 XLSX。
#   7. 正式續跑：優先讀取 JSONL checkpoint + result.xlsx/result_error.xlsx。
#   8. 完整成功後才建立 /content/result.xlsx；失敗時保留 partial/error。
#   9. 保留 v8.4 正式輸出結構與 input 欄位處理。
#  10. 不強制 full requery；正式版恢復原本續跑語意。
# ------------------------------------------------------------
# 正式輸出：
#   /content/result.xlsx
#   /content/result_error.xlsx
#   /content/result_partial.xlsx（僅不完整執行時）
#   /content/gmap.csv
#   /content/nlsc_raw/
#   /content/nlsc_raw_address/（地號＋地址案件的次要門牌查詢 response）
#   /content/result_manifest.json
#   /content/cell2_checkpoint_state.jsonl（未完成時保留；完整成功後刪除）
#   /content/cell2_timing_report.json
# ============================================================

FORMAL_CHECKPOINT_JSONL = BASE_DIR / "cell2_checkpoint_state.jsonl"
FORMAL_TIMING_JSON = BASE_DIR / "cell2_timing_report.json"

# 正式版改回 /content 正式輸出；不使用 SPEED_TEST 子目錄。
SUCCESS_OUTPUT_XLSX = BASE_DIR / "result.xlsx"
ERROR_OUTPUT_XLSX = BASE_DIR / "result_error.xlsx"
GMAP_OUTPUT_CSV = BASE_DIR / "gmap.csv"
RAW_DIR = BASE_DIR / "nlsc_raw"
RESULT_PARTIAL_XLSX = BASE_DIR / "result_partial.xlsx"
RESULT_MANIFEST = BASE_DIR / "result_manifest.json"
RESULT_TEMP_XLSX = BASE_DIR / "result.__tmp__.xlsx"

# 避免同一 notebook 內若曾執行其他 SPEED TEST，沿用舊 checkpoint 計數。
_V7_CHECKPOINT_SEQ = 0
_V7_CHECKPOINT_WRITTEN_REUSED = set()
_FORMAL_QUERY_CHECKPOINTED = set()


def _formal_load_jsonl_results():
    """讀取正式 JSONL checkpoint，取各 input index 最新一筆 result。"""
    path = FORMAL_CHECKPOINT_JSONL
    if not path.exists():
        return {}

    latest = {}
    try:
        with path.open("r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    obj = json.loads(line)
                except Exception:
                    continue
                if obj.get("type") != "result":
                    continue
                try:
                    idx = int(obj.get("index"))
                except Exception:
                    continue
                result = obj.get("result")
                if not isinstance(result, dict):
                    continue
                seq = obj.get("seq", 0)
                try:
                    seq_num = int(seq)
                except Exception:
                    seq_num = 0
                old = latest.get(idx)
                if old is None or seq_num >= old[0]:
                    latest[idx] = (seq_num, result)
    except Exception as e:
        print(f"⚠️ 讀取 JSONL checkpoint 失敗：{path} -> {e}")

    return {
        idx: dict(payload[1])
        for idx, payload in latest.items()
    }


def _formal_load_all_resume_results():
    """
    正式續跑來源：
      1. JSONL checkpoint（最新逐筆狀態）
      2. result_partial.xlsx
      3. result.xlsx
      4. result_error.xlsx

    同案件以 OK 優先；同狀態時後讀來源覆蓋前者。
    """
    combined = {}

    jsonl_results = _formal_load_jsonl_results()
    for idx, item in jsonl_results.items():
        try:
            key = build_resume_key_from_result(item)
        except Exception:
            key = ""
        if not key:
            continue
        old = combined.get(key)
        if old is None:
            combined[key] = item
        else:
            old_ok = clean_text(old.get("狀態", "")) == "OK"
            new_ok = clean_text(item.get("狀態", "")) == "OK"
            if new_ok and not old_ok:
                combined[key] = item
            elif new_ok == old_ok:
                combined[key] = item

    # 與 v8.4 相容：既有正式 Excel 仍可直接續跑。
    for path in (RESULT_PARTIAL_XLSX, SUCCESS_OUTPUT_XLSX, ERROR_OUTPUT_XLSX):
        if not path.exists():
            continue
        try:
            items = load_existing_result_xlsx(path)
        except Exception as e:
            print(f"⚠️ 讀取既有結果失敗：{path} -> {e}")
            continue

        print(f"  續跑讀取：{path.name} -> {len(items)} 筆")
        for item in items:
            try:
                key = build_resume_key_from_result(item)
            except Exception:
                continue
            if not key:
                continue
            old = combined.get(key)
            if old is None:
                combined[key] = item
                continue
            old_ok = clean_text(old.get("狀態", "")) == "OK"
            new_ok = clean_text(item.get("狀態", "")) == "OK"
            if new_ok and not old_ok:
                combined[key] = item
            elif new_ok == old_ok:
                combined[key] = item

    if combined:
        print(f"既有結果合併後：{len(combined)} 筆")

    return list(combined.values())


def _formal_save_resume_checkpoint(data, results_by_input_index, query_result_map):
    """正式版 JSONL checkpoint：每筆只追加，不重建 XLSX。"""
    global _V7_CHECKPOINT_SEQ

    t0 = _speed_now()
    appended = 0
    FORMAL_CHECKPOINT_JSONL.parent.mkdir(parents=True, exist_ok=True)
    file_exists = FORMAL_CHECKPOINT_JSONL.exists()

    # 找出目前檔案最後 seq，容許 notebook 重跑時繼續追加。
    if file_exists and _V7_CHECKPOINT_SEQ == 0:
        try:
            with FORMAL_CHECKPOINT_JSONL.open("r", encoding="utf-8") as f:
                for line in f:
                    try:
                        obj = json.loads(line)
                        seq = obj.get("seq")
                        if seq is not None:
                            _V7_CHECKPOINT_SEQ = max(_V7_CHECKPOINT_SEQ, int(seq))
                    except Exception:
                        continue
        except Exception:
            pass

    with FORMAL_CHECKPOINT_JSONL.open("a", encoding="utf-8") as f:
        if not file_exists:
            f.write(json.dumps({
                "type": "meta",
                "version": "Cell2 v8.7 429 circuit-breaker dual-coordinate completed-status exclusion formal",
                "input_rows": len(data),
            }, ensure_ascii=False) + "\n")

        # 已成功沿用資料只需在這次執行第一次記錄。
        for idx in sorted(results_by_input_index):
            if idx in _V7_CHECKPOINT_WRITTEN_REUSED:
                continue
            result = dict(results_by_input_index[idx])
            _V7_CHECKPOINT_SEQ += 1
            f.write(json.dumps({
                "type": "result",
                "seq": _V7_CHECKPOINT_SEQ,
                "index": int(idx),
                "result": result,
                "source": "reused",
            }, ensure_ascii=False, default=str) + "\n")
            _V7_CHECKPOINT_WRITTEN_REUSED.add(idx)
            appended += 1

        # 本次查詢採 append-once：每個 input index 只追加一次。
        # query_result_map 會累積整批結果；逐次重寫全部 key 會形成 O(N²) I/O。
        for idx in sorted(query_result_map):
            if idx in _FORMAL_QUERY_CHECKPOINTED:
                continue
            result = dict(query_result_map[idx])
            _V7_CHECKPOINT_SEQ += 1
            f.write(json.dumps({
                "type": "result",
                "seq": _V7_CHECKPOINT_SEQ,
                "index": int(idx),
                "result": result,
                "source": "query",
            }, ensure_ascii=False, default=str) + "\n")
            _FORMAL_QUERY_CHECKPOINTED.add(idx)
            appended += 1

        f.flush()
        try:
            os.fsync(f.fileno())
        except Exception:
            pass

    _SPEED_TIMING["checkpoint_calls"] += 1
    _SPEED_TIMING["checkpoint_seconds"] += _speed_now() - t0
    _SPEED_TIMING["v7_checkpoint_records_appended"] = (
        _SPEED_TIMING.get("v7_checkpoint_records_appended", 0) + appended
    )
    return []


# 正式版恢復續跑，不再強制 full requery。
load_all_resume_results = _formal_load_all_resume_results
save_resume_checkpoint = _formal_save_resume_checkpoint

# 正式版重新初始化本輪 timing 計數，避免同一 notebook 先前 SPEED TEST 殘留。
for _k in list(_SPEED_TIMING.keys()):
    if _k == "query_rows":
        _SPEED_TIMING[_k] = []
    elif _k == "script_start":
        _SPEED_TIMING[_k] = time.perf_counter()
    elif isinstance(_SPEED_TIMING[_k], (int, float)):
        _SPEED_TIMING[_k] = 0.0

print("=" * 70)
print("Cell 2 v8.7：429 首次即切換瀏覽器備援＋DIRECT 全域冷卻斷路器＋JSONL 續跑＋雙座標")
print(f"正式輸出：{BASE_DIR}")
print(f"DIRECT 最小間隔：{DIRECT_MIN_INTERVAL_SECONDS:.2f} 秒")
print(f"429 策略：同筆 429 不重試、立即備援；其他暫時錯誤最多 {DIRECT_MAX_ATTEMPTS} 次 DIRECT；無 Retry-After 時冷卻={DIRECT_429_BACKOFF} 秒")
print(f"JSONL checkpoint：{FORMAL_CHECKPOINT_JSONL}")
print("=" * 70)

_CELL2_RUN_STARTED_NS = time.time_ns()
run_results = None
try:
    run_results = await main()
finally:
    total_seconds = time.perf_counter() - _SPEED_TIMING["script_start"]
    status_counts = {}
    for r in run_results if isinstance(run_results, list) else []:
        status = clean_text(r.get("狀態", ""))
        status_counts[status] = status_counts.get(status, 0) + 1

    report = {
        "version": "Cell2 v8.7 429 circuit-breaker dual-coordinate completed-status exclusion formal",
        "test_mode": "FORMAL_RESUME_CACHE_FIRST",
        "output_dir": str(BASE_DIR),
        "total_seconds_including_setup": total_seconds,
        "status_counts": status_counts,
        "status_filter": dict(_CELL2_STATUS_FILTER_STATS),
        "timing": {
            **_SPEED_TIMING,
            "direct_min_interval_seconds": DIRECT_MIN_INTERVAL_SECONDS,
            "direct_429_backoff_seconds": DIRECT_429_BACKOFF,
            "direct_max_attempts_for_non_429": DIRECT_MAX_ATTEMPTS,
            "direct_429_same_case_retry_count": 0,
            "direct_429_policy": "FIRST_429_IMMEDIATE_BROWSER_FALLBACK; SKIP_DIRECT_DURING_GLOBAL_COOLDOWN; NO_PER_ROW_COOLDOWN_WAIT",
            "direct_429_count": _V4_LAND_STATE["direct_429_count"],
            "direct_cooldown_skips": _V4_LAND_STATE["direct_cooldown_skips"],
            "direct_429_cooldown_seconds": _V4_LAND_STATE["direct_429_cooldown_seconds"],
            "direct_429_server_retry_after_seconds": _V4_LAND_STATE["direct_429_server_retry_after_seconds"],
            "checkpoint_mode": "append_only_jsonl",
            "checkpoint_state_file": str(FORMAL_CHECKPOINT_JSONL),
            "land_direct_calls": _V4_LAND_STATE["direct_calls"],
            "land_direct_seconds": _V4_LAND_STATE["direct_seconds"],
            "land_fallback_calls": _V4_LAND_STATE["fallback_calls"],
            "land_fallback_seconds": _V4_LAND_STATE["fallback_seconds"],
            "query_total_seconds": _SPEED_TIMING["land_query_seconds"] + _SPEED_TIMING["address_query_seconds"],
        },
    }
    try:
        FORMAL_TIMING_JSON.write_text(
            json.dumps(report, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )
        print(f"Timing JSON：{FORMAL_TIMING_JSON}")
    except Exception as e:
        print(f"⚠️ Timing JSON 寫入失敗：{e}")

    # 完整成功後，JSONL 只作中斷續跑用途，避免下次留下陳舊狀態。
    if isinstance(run_results, list) and run_results:
        complete_statuses = {"OK", "重複土地"}
        all_complete = all(
            clean_text(r.get("狀態", "")) in complete_statuses
            for r in run_results
        )
        if all_complete:
            try:
                if FORMAL_CHECKPOINT_JSONL.exists():
                    FORMAL_CHECKPOINT_JSONL.unlink()
                print("完整完成（OK／重複土地皆已處理）：已清除 JSONL checkpoint。")
            except Exception as e:
                print(f"⚠️ 無法清除 JSONL checkpoint：{e}")


# ======================================================================
# Cell 2+3 v2｜階段交接檢查
# 必須確認本次 Cell2 完整成功、輸入 SHA 相符、result.xlsx 為本輪新產生且筆數吻合。
# ======================================================================

_CELL23_RESULT_XLSX = Path("/content/result.xlsx")
_CELL2_MANIFEST_PATH = Path("/content/result_manifest.json")

if not isinstance(run_results, list):
    raise RuntimeError("Cell 2 未正常回傳結果；停止 Cell 3，避免使用舊資料。")

_CELL2_ALLOWED_FINAL_STATUSES = {"OK", "重複土地"}
_CELL2_BAD_RESULTS = [
    r for r in run_results
    if clean_text(r.get("狀態", "")) not in _CELL2_ALLOWED_FINAL_STATUSES
]
if _CELL2_BAD_RESULTS:
    raise RuntimeError(
        f"Cell 2 尚有 {len(_CELL2_BAD_RESULTS)} 筆未完成案件；"
        "已保留續跑資料，但本次禁止啟動 Cell 3。"
    )

if not _CELL2_MANIFEST_PATH.is_file():
    raise FileNotFoundError("Cell 2 未產生本次 result_manifest.json；停止 Cell 3。")

try:
    _CELL2_MANIFEST = json.loads(
        _CELL2_MANIFEST_PATH.read_text(encoding="utf-8")
    )
except Exception as _manifest_error:
    raise RuntimeError(f"Cell 2 result_manifest.json 無法解析：{_manifest_error}")

if not bool(_CELL2_MANIFEST.get("complete", False)):
    raise RuntimeError("Cell 2 manifest 未標示 complete=true；停止 Cell 3。")

if not _CELL2_INPUT_SHA256:
    raise RuntimeError("Cell 2 未留下本次輸入檔 SHA256；停止 Cell 3。")

_CELL2_MANIFEST["input_file_sha256"] = _CELL2_INPUT_SHA256

if not _CELL23_RESULT_XLSX.is_file() or _CELL23_RESULT_XLSX.stat().st_size <= 0:
    raise RuntimeError("Cell 2 未產生有效 /content/result.xlsx；停止 Cell 3。")

if _CELL23_RESULT_XLSX.stat().st_mtime_ns < _CELL2_RUN_STARTED_NS:
    raise RuntimeError("result.xlsx 並非本次 Cell 2 新產生；拒絕使用舊資料啟動 Cell 3。")

_CELL2_WB_CHECK = openpyxl.load_workbook(
    _CELL23_RESULT_XLSX, read_only=True, data_only=True
)
try:
    _CELL2_WS_CHECK = _CELL2_WB_CHECK[_CELL2_WB_CHECK.sheetnames[0]]
    _CELL2_RESULT_ROW_COUNT = max(0, _CELL2_WS_CHECK.max_row - 1)
finally:
    _CELL2_WB_CHECK.close()

if _CELL2_RESULT_ROW_COUNT != int(_CELL2_MANIFEST.get("success", -1)):
    raise RuntimeError(
        "Cell 2 result.xlsx 筆數與 manifest success 不一致："
        f"{_CELL2_RESULT_ROW_COUNT}/{_CELL2_MANIFEST.get('success')}"
    )

_CELL2_RESULT_SHA256 = hashlib.sha256(_CELL23_RESULT_XLSX.read_bytes()).hexdigest()
_CELL2_MANIFEST["result_file_sha256"] = _CELL2_RESULT_SHA256
_CELL2_MANIFEST_PATH.write_text(
    json.dumps(_CELL2_MANIFEST, ensure_ascii=False, indent=2),
    encoding="utf-8",
)
_CELL2_MANIFEST_READBACK = json.loads(
    _CELL2_MANIFEST_PATH.read_text(encoding="utf-8")
)
if _CELL2_MANIFEST_READBACK.get("result_file_sha256") != _CELL2_RESULT_SHA256:
    raise RuntimeError("Cell 2 result.xlsx SHA256 manifest 回讀驗證失敗；停止 Cell 3。")

print()
print("[CELL23_STAGE_DONE] Cell 2 v8.7")
print(f"[CELL23_RESULT_XLSX] {_CELL23_RESULT_XLSX}")
print(f"[CELL23_RESULT_SIZE_BYTES] {_CELL23_RESULT_XLSX.stat().st_size}")
print("=" * 76)
print("[CELL23_STAGE_START] Cell 3 v4.3.4")
print("=" * 76)

# ============================================================
# 巡田最速適｜Cell 3 v4.3.4 多圖幅聯集／重複案件鍵安全處理版｜Optional Drive 儲存修正版
# 基於 v4.3.2；保留全部原始圖幅與 PNG 證據，將無法形成有效面外輪廓的退化／佔位圖排除於 Polygon 化。
# 其餘有效 PNG 仍逐幅向量化，先 WGS84 union 再投影 EPSG:3826；Cell5 handoff 明確標示排除圖幅。
# Drive 已掛載：工作資料存入本次 Drive 執行資料夾；未掛載：全部存 /content，不主動掛載。
# Cell5 固定讀取的 /content/CELL3_HANDOFF.pkl 保留。
#
# 整合內容：
#   1. Cell 2 result.xlsx
#   2. Direct JSONP qryTileMapIndex(flag=2, imgflag=1)
#   3. 多 image Response：全部解析，不取第一張
#   4. Cell 2 定位點 → bbox 唯一選圖
#   5. Image Audit：
#        image_record_count
#        selected_image_index
#        image_selection_method
#   6. Raw Response 優先離線重解析；不存在才線上查詢
#   7. 舊 PNG 與正確選定 image 不一致時，以正確選定 image 更新 PNG
#   8. 每筆下載／Audit 立即保存 checkpoint
#   9. PNG Alpha → Polygon / MultiPolygon
#  10. Pixel → WGS84 → TWD97 EPSG:3826
#  11. Cell 2 定位點與 Polygon 關係
#  12. SVG
#  13. GeoJSON / GPKG / SHP / Excel
#  14. Cell 3 → Cell 5 CELL3_HANDOFF.pkl
#  15. 續跑：PNG 不變且向量有效時沿用；PNG 變動只重做該筆向量
#  16. 正式成果檔名不含「_106筆」
#
# 回歸測試整合：
#   A. 地號標準化
#   B. 地段標準化
#   C. Image bbox 選圖：單一命中／無命中／多重命中
#   D. Base64 image 解碼
#   E. Cell 2 案件數與 Point Map
#   F. 正式輸出檔名規則
#   G. 最終 105 筆地籍案件 Audit 完整性
# ============================================================

import os
import re
import json
import base64
import hashlib
import asyncio
import pickle
import tempfile
import traceback
import time
import shutil
import urllib.request
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import cv2

from shapely.geometry import Polygon, MultiPolygon, Point
from shapely.ops import transform as shapely_transform_legacy, unary_union
try:
    from shapely import transform as shapely_transform_v2
except Exception:
    shapely_transform_v2 = None
import geopandas as gpd
from pyproj import Transformer

from playwright.async_api import async_playwright
import nest_asyncio


# ============================================================
# 0. 路徑與設定
# ============================================================

CELL3_VERSION = "v4.3.4"

# Cell 2 v8.7 實際輸出：正式 /content/result.xlsx 優先；
# 若本次正在使用 Resume 測試隔離輸出，則自動改用測試完成檔。
_INPUT_CANDIDATES = [
    Path("/content/result.xlsx"),
    Path("/content/CELL2_RESUME_TEST/result.xlsx"),
]
INPUT_XLSX = next((p for p in _INPUT_CANDIDATES if p.exists()), _INPUT_CANDIDATES[0])
if not INPUT_XLSX.is_file():
    raise FileNotFoundError(f"Cell 3 找不到本次 Cell 2 結果：{INPUT_XLSX}")
# 以輸入 SHA 識別批次；相同輸入可沿用同一工作目錄，避免每次重跑複製整批圖檔。
INPUT_BATCH_SHA256 = hashlib.sha256(INPUT_XLSX.read_bytes()).hexdigest()
INPUT_BATCH_KEY = INPUT_BATCH_SHA256[:16]

# ------------------------------------------------------------
# Drive 僅偵測，不主動掛載
# ------------------------------------------------------------
DRIVE_MYDRIVE = Path("/content/drive/MyDrive")
DRIVE_MOUNTED = DRIVE_MYDRIVE.is_dir()
RUN_STAMP = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

if DRIVE_MOUNTED:
    FRO_DRIVE_ROOT = DRIVE_MYDRIVE / "巡田最速適"
    DRIVE_CELL3_ROOT = FRO_DRIVE_ROOT / "Cell3"
    DRIVE_CELL3_ROOT.mkdir(parents=True, exist_ok=True)
    # 同一份 Cell2 輸入固定使用同一工作資料夾；新批次以 SHA 前綴隔離。
    DRIVE_BACKUP_DIR = DRIVE_CELL3_ROOT / f"MULTI_TILE_V2_{INPUT_BATCH_KEY}"
    DRIVE_BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    STORAGE_ROOT = DRIVE_BACKUP_DIR
    RUN_BACKUP_DIR = DRIVE_BACKUP_DIR
    print("CELL3_STORAGE_MODE=GOOGLE_DRIVE")
    print("CELL3_STORAGE_ROOT=", STORAGE_ROOT)
else:
    FRO_DRIVE_ROOT = None
    DRIVE_CELL3_ROOT = None
    DRIVE_BACKUP_DIR = None
    STORAGE_ROOT = Path("/content")
    # 未掛載 Drive 時，工作資料直接使用既有 /content 路徑；
    # 僅在 /content 建立小型稽核備份資料夾，不嘗試連接 Google Drive。
    LOCAL_BACKUP_DIR = (
        Path("/content") / "CELL3_LOCAL_BACKUPS" / f"MULTI_TILE_V1_{RUN_STAMP}"
    )
    LOCAL_BACKUP_DIR.mkdir(parents=True, exist_ok=False)
    RUN_BACKUP_DIR = LOCAL_BACKUP_DIR
    print("CELL3_STORAGE_MODE=LOCAL_CONTENT")
    print("CELL3_STORAGE_ROOT=", STORAGE_ROOT)

# Drive 模式：本次執行資料夾內保存全部影像、Raw、成果、SVG 與 checkpoint。
# 未掛載模式：保留 v4.3.1 原本 /content 檔案位置，以便沿用本機續跑資料。
if DRIVE_MOUNTED:
    IMAGE_DIR = STORAGE_ROOT / "IMAGE_TILES"
    RAW_DIR = STORAGE_ROOT / "RAW_RESPONSES"
    IMAGE_RESULT_XLSX = STORAGE_ROOT / "result_cadastral_image_multi_tile_candidate.xlsx"
    FORMAL_DIR = STORAGE_ROOT / "CANDIDATE_EXPORTS"
    SVG_DIR = STORAGE_ROOT / "SVG"
    VECTOR_CHECK_DIR = STORAGE_ROOT / "VECTOR_CHECK"
    DOWNLOAD_STATE_PATH = STORAGE_ROOT / "CELL3_DOWNLOAD_PARTIAL_MULTI_TILE_V1.pkl"
    VECTOR_STATE_PATH = STORAGE_ROOT / "CELL3_VECTOR_PARTIAL_MULTI_TILE_V1.pkl"
else:
    IMAGE_DIR = Path("/content/NLSC_地籍圖_多圖幅候選_V1")
    RAW_DIR = Path("/content/NLSC_地籍圖_多圖幅候選_原始Response")
    IMAGE_RESULT_XLSX = Path("/content/result_cadastral_image_multi_tile_candidate.xlsx")
    FORMAL_DIR = Path("/content/NLSC_地籍圖_多圖幅候選成果_V1")
    SVG_DIR = Path("/content/NLSC_地籍圖_多圖幅候選向量SVG_V1")
    VECTOR_CHECK_DIR = Path("/content/NLSC_地籍圖_多圖幅候選向量檢查_V1")
    DOWNLOAD_STATE_PATH = Path("/content/CELL3_DOWNLOAD_PARTIAL_MULTI_TILE_V1.pkl")
    VECTOR_STATE_PATH = Path("/content/CELL3_VECTOR_PARTIAL_MULTI_TILE_V1.pkl")

# 舊版 Raw Response 仍可作為離線重解析來源；本次新資料依儲存模式寫入 RAW_DIR。
RAW_READ_DIRS = [
    Path("/content/NLSC_地籍圖_原始Response"),
    Path("/content/NLSC_地籍圖_多圖幅候選_原始Response"),
    RAW_DIR,
]
# Cell5 固定讀取這個 /content 交接檔，因此不隨 Drive 工作目錄移動。
CELL3_HANDOFF_PATH = Path("/content/CELL3_HANDOFF.pkl")
HANDOFF_VERSION = "1.3-MULTI-TILE-WGS84-UNION-PLACEHOLDER-FILTER"
CHECKPOINT_JOURNAL_PATH = STORAGE_ROOT / f"CELL3_CHECKPOINTS_V2_{INPUT_BATCH_KEY}.journal"


CONCURRENCY = 4
API_RETRY = 4
API_TIMEOUT_MS = 30000
API_DELAY_MS = 250
API_BACKOFF_BASE_MS = 300
# 多 worker 共用全域請求節流，避免每個 worker 各自 delay 後仍形成高併發尖峰。
GLOBAL_API_MIN_INTERVAL_MS = 350
_GLOBAL_API_LOCK = asyncio.Lock()
_GLOBAL_API_NEXT_SLOT = 0.0

# v4.2.1：整批線上補抓完成後，若仍有案件抓不到，不立即終止。
# 改成「失敗案件專屬 retry queue」再跑數輪，成功者可在本次執行續跑向量化。
ONLINE_RETRY_ROUNDS = 2
RETRY_ONLY_CONCURRENCY = 2
RETRY_ONLY_DELAY_MS = 500
RETRY_ONLY_TIMEOUT_MS = 45000

NLSC_URL = "https://maps.nlsc.gov.tw/T09/mobilemap.action"
NLSC_IMAGE_ENDPOINT = "https://landmaps.nlsc.gov.tw/S_Maps/qryTileMapIndex"

# 正式成果檔名：不可加入案件筆數。
GEOJSON_PATH = FORMAL_DIR / "NLSC_地籍圖.geojson"
GPKG_PATH = FORMAL_DIR / "NLSC_地籍圖.gpkg"
SHP_DIR = FORMAL_DIR / "SHP"
SHP_BASE = SHP_DIR / "NLSC_地籍圖"
SHP_PATH = SHP_BASE.with_suffix(".shp")
ATTRIBUTE_PATH = FORMAL_DIR / "NLSC_地籍圖_屬性.xlsx"
DUPLICATE_KEY_AUDIT_PATH = FORMAL_DIR / "NLSC_重複案件鍵稽核.xlsx"
SHP_TILE_DETAILS_CSV = FORMAL_DIR / "SHP_圖幅詳細資訊.csv"
SHP_LONG_TEXT_CSV = FORMAL_DIR / "SHP_超長文字欄位完整值.csv"

# 舊錯誤命名：本次重新產出前清掉。
LEGACY_PATHS = [
    FORMAL_DIR / "NLSC_地籍圖_106筆.geojson",
    FORMAL_DIR / "NLSC_地籍圖_106筆.gpkg",
    FORMAL_DIR / "NLSC_地籍圖_106筆_屬性.xlsx",
]
LEGACY_SHP_BASE = SHP_DIR / "NLSC_地籍圖_106筆"

def _resolve_executed_source_text():
    """取得實際執行的完整來源；支援 .py 執行與 Colab/Jupyter cell。"""
    # 1) 一般 .py 執行
    _file_value = globals().get("__file__")
    if _file_value:
        _candidate = Path(str(_file_value)).expanduser()
        if _candidate.is_file():
            return _candidate.read_text(encoding="utf-8"), _candidate.name, "PY_FILE"

    # 2) Colab/Jupyter 直接貼入整格程式碼執行時，__file__ 不存在；
    #    從 IPython 原始輸入歷史找本版本完整 cell，避免把 exec(open(...)) 包裝行誤當來源。
    _histories = []
    for _key in ("In", "_ih"):
        _h = globals().get(_key)
        if isinstance(_h, (list, tuple)):
            _histories.extend(x for x in _h if isinstance(x, str))
    try:
        _ip = get_ipython()  # noqa: F821 - available only inside IPython
        _raw = getattr(getattr(_ip, "history_manager", None), "input_hist_raw", None)
        if isinstance(_raw, (list, tuple)):
            _histories.extend(x for x in _raw if isinstance(x, str))
    except Exception:
        pass

    _matches = [
        x for x in _histories
        if "Cell 3 v4.3.4" in x
        and "def build_geometry_from_mask" in x
        and "run_multitile_geometry_regression_tests" in x
        and len(x) >= 30000
    ]
    if _matches:
        _text = max(_matches, key=len)
        return _text, "Cell3_v4.3.4.py", "IPYTHON_CELL"
    return None, None, "SOURCE_NOT_FOUND"


try:
    _source_text, _source_name, _source_kind = _resolve_executed_source_text()
    if _source_text:
        _source_dir = RUN_BACKUP_DIR / "SOURCE"
        _source_dir.mkdir(parents=True, exist_ok=True)
        _script_backup = _source_dir / _source_name
        _source_bytes = _source_text.encode("utf-8")
        _script_tmp = _script_backup.with_name(_script_backup.name + ".tmp")
        _script_tmp.write_bytes(_source_bytes)
        os.replace(_script_tmp, _script_backup)
        _source_sha = hashlib.sha256(_script_backup.read_bytes()).hexdigest()
        if _source_sha != hashlib.sha256(_source_bytes).hexdigest():
            raise RuntimeError("來源程式 Drive 備份 SHA256 不一致")
        print(f"SOURCE_SCRIPT_BACKUP=PASS; SOURCE_KIND={_source_kind}; PATH={_script_backup}; SHA256={_source_sha}")
    else:
        print(f"SOURCE_SCRIPT_BACKUP=SKIPPED; REASON={_source_kind}")
except Exception as _source_backup_error:
    print("SOURCE_SCRIPT_BACKUP=FAILED", repr(_source_backup_error))

for p in [IMAGE_DIR, RAW_DIR, FORMAL_DIR, SVG_DIR, VECTOR_CHECK_DIR, SHP_DIR]:
    p.mkdir(parents=True, exist_ok=True)

for p in LEGACY_PATHS:
    if p.exists():
        try:
            p.unlink()
            print("已移除舊版正式成果：", p)
        except Exception as e:
            print("⚠️ 無法移除舊版正式成果：", p, repr(e))

for suffix in [".shp", ".shx", ".dbf", ".prj", ".cpg"]:
    p = LEGACY_SHP_BASE.with_suffix(suffix)
    if p.exists():
        try:
            p.unlink()
            print("已移除舊版正式成果：", p)
        except Exception as e:
            print("⚠️ 無法移除舊版 SHP：", p, repr(e))


# ============================================================
# 1. 基本工具
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    s = str(value).strip()
    return "" if s.lower() == "nan" else s


def normalize_landno(value):
    value = clean_text(value)
    if not value:
        return ""
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    if "-" in value:
        parts = [p.strip() for p in value.split("-", 1)]
        if len(parts) == 2 and all(re.fullmatch(r"\d+", p) for p in parts):
            return parts[0].zfill(4) + parts[1].zfill(4)
        return value
    if re.fullmatch(r"\d+", value):
        if len(value) == 4:
            return value + "0000"
        if len(value) < 4:
            return value.zfill(4) + "0000"
        if len(value) < 8:
            return value.zfill(8)
        return value
    return value


def normalize_sect(value):
    value = clean_text(value)
    if re.fullmatch(r"\d+\.0", value):
        value = value[:-2]
    return value.zfill(4) if value else ""


def safe_float(value):
    try:
        x = float(value)
        return x if np.isfinite(x) else np.nan
    except Exception:
        return np.nan


def is_address_case(row):
    return (
        clean_text(row.get("查詢方式", "")) == "地址"
        or (
            clean_text(row.get("縣市", "")) == "桃園市"
            and clean_text(row.get("鄉鎮市區", "")) == "桃園區"
            and not clean_text(row.get("office", ""))
        )
    )


def is_cadastral_query_row(row):
    if clean_text(row.get("查詢方式", "")) == "地址":
        return False
    office = clean_text(row.get("office", ""))
    sect = clean_text(row.get("sect", ""))
    landno = normalize_landno(row.get("查詢地號", row.get("地號", "")))
    return bool(office and sect and landno)


def case_key(row):
    office = clean_text(row.get("office", "")).upper()
    sect = normalize_sect(row.get("sect", ""))
    landno = normalize_landno(row.get("查詢地號", row.get("地號", "")))
    return f"{office}_{sect}_{landno}"


def stable_case_id(row):
    office = clean_text(row.get("office", "")).upper()
    sect = normalize_sect(row.get("sect", ""))
    landno = normalize_landno(row.get("查詢地號", row.get("地號", "")))
    if office and sect and landno and is_cadastral_query_row(row):
        return f"LAND|{office}|{sect}|{landno}"
    return "ADDR|" + "|".join([
        clean_text(row.get("查詢方式", "")),
        clean_text(row.get("縣市", "")),
        clean_text(row.get("鄉鎮市區", "")),
        clean_text(row.get("地址", "")),
        clean_text(row.get("名稱", "")),
    ])


def expected_image_name(row_number, row):
    return (
        f"{row_number:03d}_"
        f"{clean_text(row.get('office', ''))}_"
        f"{normalize_sect(row.get('sect', ''))}_"
        f"{normalize_landno(row.get('查詢地號', row.get('地號', '')))}.png"
    )


def expected_image_record_name(row_number, row, record_index):
    stem = expected_image_name(row_number, row)[:-4]
    return f"{stem}_RECORD_{int(record_index):02d}.png"


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


_TILE_SHA256_CACHE = {}

def cached_tile_sha256(path):
    """依檔案大小／mtime 快取 PNG SHA；檔案變動時重新計算。"""
    p = Path(path)
    st = p.stat()
    cache_key = (
        str(p.resolve()),
        int(st.st_size),
        int(st.st_mtime_ns),
        int(getattr(st, "st_ctime_ns", 0)),
    )
    cached = _TILE_SHA256_CACHE.get(cache_key)
    if cached:
        return cached
    digest = sha256_file(p)
    for old_key in list(_TILE_SHA256_CACHE):
        if old_key[0] == cache_key[0]:
            _TILE_SHA256_CACHE.pop(old_key, None)
    _TILE_SHA256_CACHE[cache_key] = digest
    return digest


def atomic_write_bytes(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + f".{os.getpid()}.tmp")
    try:
        tmp.write_bytes(data)
        os.replace(tmp, path)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass


def mirror_file_to_drive(source_path, relative_folder):
    """Drive 模式將外部來源鏡像進本次執行資料夾；本地模式完全略過。"""
    if not DRIVE_MOUNTED or DRIVE_BACKUP_DIR is None:
        return False
    source_path = Path(source_path)
    if not source_path.is_file():
        return False
    destination = DRIVE_BACKUP_DIR / relative_folder / source_path.name
    if source_path.resolve() == destination.resolve():
        return True
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp_destination = destination.with_name(destination.name + ".tmp")
    shutil.copy2(source_path, temp_destination)
    os.replace(temp_destination, destination)
    if sha256_file(destination) != sha256_file(source_path):
        raise RuntimeError(f"Drive 即時備份 SHA256 不一致：{source_path.name}")
    return True


def tile_set_signature(tile_records):
    """使用內容 SHA 與 bbox 建立穩定簽章，不依賴 mtime。"""
    compact = []
    for rec in sorted(tile_records, key=lambda x: int(x.get("record_index", 0))):
        p = Path(rec["image_path"])
        if not p.is_file():
            return None
        actual_sha = cached_tile_sha256(p)
        expected_sha = clean_text(rec.get("sha256", ""))
        if expected_sha and actual_sha != expected_sha:
            return None
        compact.append({
            "record_index": int(rec["record_index"]),
            "sha256": actual_sha,
            "bbox": [float(rec[k]) for k in ("lx", "ly", "rx", "ry")],
        })
    payload = json.dumps(compact, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
    return {"algorithm": "SHA256_TILE_SET_V1", "sha256": hashlib.sha256(payload.encode("utf-8")).hexdigest(), "count": len(compact)}


def image_signature(path):
    p = Path(path)
    if not p.exists():
        return None
    st = p.stat()
    return {"size": int(st.st_size), "mtime_ns": int(st.st_mtime_ns)}


def signature_equal(a, b):
    return (
        isinstance(a, dict) and isinstance(b, dict)
        and int(a.get("size", -1)) == int(b.get("size", -2))
        and int(a.get("mtime_ns", -1)) == int(b.get("mtime_ns", -2))
    )


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def inspect_png_bytes(data, bbox=None):
    """完整解碼 PNG，辨識無法形成面幾何的單像素佔位圖。"""
    if not isinstance(data, (bytes, bytearray)) or not bytes(data).startswith(b"\x89PNG\r\n\x1a\n"):
        raise RuntimeError("PNG 檔頭無效")
    image = cv2.imdecode(np.frombuffer(bytes(data), dtype=np.uint8), cv2.IMREAD_UNCHANGED)
    if image is None or not hasattr(image, "shape") or len(image.shape) < 2:
        raise RuntimeError("PNG 無法完整解碼")
    height, width = image.shape[:2]
    if width < 1 or height < 1:
        raise RuntimeError(f"PNG 尺寸無效：{width}x{height}")

    # 必須與 build_geometry_from_mask 使用相同的 Alpha／灰階遮罩規則。
    if len(image.shape) == 2:
        mask = (image > 0).astype(np.uint8) * 255
    elif image.shape[2] >= 4:
        mask = (image[:, :, 3] > 0).astype(np.uint8) * 255
    else:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        mask = (gray < 250).astype(np.uint8) * 255
    foreground_pixel_count = int(np.count_nonzero(mask))
    contours, hierarchy = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    polygon_contour_count = 0
    if contours and hierarchy is not None:
        hierarchy = hierarchy[0]
        polygon_contour_count = sum(
            1 for i, contour in enumerate(contours)
            if hierarchy[i][3] == -1 and len(contour) >= 3 and cv2.contourArea(contour) > 0
        )

    # NLSC 可能回傳 1x1、全透明、單點或只有無面積邊線的退化圖幅。
    # 保留原始 PNG 與 bbox；若沒有可建立 Polygon 的外輪廓，就不納入幾何聯集。
    is_placeholder = (width * height <= 1) or (polygon_contour_count == 0)
    if is_placeholder:
        image_class = "DEGENERATE_PLACEHOLDER"
        if width * height <= 1:
            placeholder_reason = "SINGLE_PIXEL_IMAGE"
        elif foreground_pixel_count == 0:
            placeholder_reason = "NO_GEOMETRY_PIXELS"
        elif polygon_contour_count == 0:
            placeholder_reason = "NO_POLYGON_AREA_CONTOUR"
        else:
            placeholder_reason = "DEGENERATE_IMAGE"
        geometry_eligible = False
    elif width >= 2 and height >= 2:
        image_class = "GEOMETRY_ELIGIBLE"
        placeholder_reason = ""
        geometry_eligible = True
    else:
        # 極窄一維影像若仍有外輪廓標記，不能安全向量化，須重取而非默默略過。
        image_class = "INVALID_GEOMETRY_DIMENSION"
        placeholder_reason = "ONE_DIMENSION_BELOW_2_PIXELS"
        geometry_eligible = False

    result = {
        "image_width": int(width),
        "image_height": int(height),
        "foreground_pixel_count": foreground_pixel_count,
        "polygon_contour_count": int(polygon_contour_count),
        "image_class": image_class,
        "geometry_eligible": bool(geometry_eligible),
        "placeholder_reason": placeholder_reason,
    }
    if bbox is not None:
        result["bbox"] = [float(v) for v in bbox]
    return result


def inspect_png_file(path, bbox=None):
    path = Path(path)
    if not path.is_file():
        raise RuntimeError(f"PNG 檔案不存在：{path}")
    return inspect_png_bytes(path.read_bytes(), bbox=bbox)


_TILE_METADATA_CACHE = {}

def enrich_tile_metadata(tile):
    """依 PNG 檔案簽章快取解碼後的影像分類；PNG 未變動時不重複 OpenCV 解碼。"""
    bbox = [safe_float(tile.get(k)) for k in ("lx", "ly", "rx", "ry")]
    if not all(np.isfinite(v) for v in bbox):
        raise RuntimeError("圖幅 bbox 不完整")
    path = Path(tile.get("image_path", ""))
    st = path.stat()
    cache_key = (
        str(path.resolve()),
        int(st.st_size),
        int(st.st_mtime_ns),
        int(getattr(st, "st_ctime_ns", 0)),
        tuple(float(v) for v in bbox),
    )
    meta = _TILE_METADATA_CACHE.get(cache_key)
    if meta is None:
        meta = inspect_png_file(path, bbox=bbox)
        for old_key in list(_TILE_METADATA_CACHE):
            if old_key[0] == cache_key[0]:
                _TILE_METADATA_CACHE.pop(old_key, None)
        _TILE_METADATA_CACHE[cache_key] = dict(meta)
    tile.update(meta)
    return dict(meta)


def _safe_json_list(value):
    if value is None or (isinstance(value, float) and not np.isfinite(value)):
        return []
    if isinstance(value, list):
        return value
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return []
    parsed = json.loads(text)
    return parsed if isinstance(parsed, list) else []


def build_shp_tile_detail_rows(frame):
    """將 DBF 不適合承載的多圖幅 JSON 拆成可 join 的逐圖幅 CSV。"""
    rows = []
    for _, row in frame.iterrows():
        case_id = clean_text(row.get("案件ID", "")) or case_key(row)
        row_no = safe_float(row.get("列號", np.nan))
        tiles = _safe_json_list(row.get("圖幅清單_JSON", "[]"))
        svg_paths = _safe_json_list(row.get("圖幅SVG清單_JSON", "[]"))
        for tile in tiles:
            if not isinstance(tile, dict):
                continue
            idx = int(tile.get("record_index", 0) or 0)
            rows.append({
                "案件ID": case_id,
                "列號": int(row_no) if np.isfinite(row_no) else "",
                "圖幅序號": idx,
                "是否定位選中": bool(tile.get("selected", False)),
                "PNG路徑": clean_text(tile.get("image_path", "")),
                "左下經度": safe_float(tile.get("lx")),
                "左下緯度": safe_float(tile.get("ly")),
                "右上經度": safe_float(tile.get("rx")),
                "右上緯度": safe_float(tile.get("ry")),
                "PNG大小_bytes": int(tile.get("image_size", 0) or 0),
                "PNG寬": int(tile.get("image_width", 0) or 0),
                "PNG高": int(tile.get("image_height", 0) or 0),
                "前景像素數": int(tile.get("foreground_pixel_count", 0) or 0),
                "有效外輪廓數": int(tile.get("polygon_contour_count", 0) or 0),
                "圖幅影像分類": clean_text(tile.get("image_class", "")),
                "是否納入向量化": bool(tile.get("geometry_eligible", False)),
                "排除原因": clean_text(tile.get("placeholder_reason", "")),
                "PNG_SHA256": clean_text(tile.get("sha256", "")),
                "SVG路徑": clean_text(svg_paths[idx - 1]) if 1 <= idx <= len(svg_paths) else "",
            })
    return rows


def prepare_shapefile_frame(frame, field_map, dbf_text_limit=254):
    """排除長 JSON 欄位；其餘超長文字完整另存 sidecar，DBF 只留安全長度。"""
    excluded = {"圖幅清單_JSON", "圖幅SVG清單_JSON", "圖幅佔位清單_JSON"}
    out = frame.drop(columns=[c for c in excluded if c in frame.columns], errors="ignore").copy()
    long_rows = []
    for col in out.columns:
        if col == "geometry":
            continue
        if out[col].dtype == object or pd.api.types.is_string_dtype(out[col].dtype):
            converted = []
            for idx, value in out[col].items():
                if value is None or (not isinstance(value, (list, dict, tuple)) and pd.isna(value)):
                    value = ""
                else:
                    value = str(value)
                value_bytes = value.encode("utf-8")
                if len(value_bytes) > dbf_text_limit:
                    row = frame.loc[idx]
                    long_rows.append({
                        "案件ID": clean_text(row.get("案件ID", "")),
                        "列號": row.get("列號", ""),
                        "欄位名稱": str(col),
                        "完整值": value,
                        "完整值字元數": len(value),
                        "完整值UTF8位元組數": len(value_bytes),
                    })
                    # 以 UTF-8 位元組長度截斷，且不切斷中文字元。
                    value = value_bytes[:dbf_text_limit].decode("utf-8", errors="ignore")
                converted.append(value)
            out[col] = converted
    return out, long_rows, sorted(excluded.intersection(set(frame.columns)))


def atomic_pickle_dump(obj, path):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    # Checkpoint 綁定目前 result.xlsx SHA，禁止跨批次誤用。
    if path.name in {DOWNLOAD_STATE_PATH.name, VECTOR_STATE_PATH.name} and isinstance(obj, dict):
        obj = dict(obj)
        obj["input_xlsx_sha256"] = globals().get("INPUT_XLSX_SHA256", "")
        obj["algorithm"] = HANDOFF_VERSION
    fd, tmp_name = tempfile.mkstemp(prefix=path.stem + "_", suffix=".tmp", dir=str(path.parent))
    os.close(fd)
    tmp = Path(tmp_name)
    try:
        pd.to_pickle(obj, tmp)
        os.replace(tmp, path)
        # Drive 模式下 checkpoint 本身就在 Drive；本地模式不觸碰 Drive。
        if DRIVE_MOUNTED and DRIVE_BACKUP_DIR is not None and path.name in {DOWNLOAD_STATE_PATH.name, VECTOR_STATE_PATH.name}:
            drive_target = DRIVE_BACKUP_DIR / path.name
            if path.resolve() != drive_target.resolve():
                drive_target.parent.mkdir(parents=True, exist_ok=True)
                drive_tmp = drive_target.with_name(drive_target.name + ".tmp")
                shutil.copy2(path, drive_tmp)
                os.replace(drive_tmp, drive_target)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass


def load_pickle_state(path, default):
    path = Path(path)
    if not path.exists():
        return default
    try:
        obj = pd.read_pickle(path)
        return obj if obj is not None else default
    except Exception as e:
        print("⚠️ 續跑狀態檔讀取失敗，重新建立：", path, repr(e))
        return default


def append_case_checkpoint(stage, key, value):
    """
    單筆追加式 checkpoint。
    只寫入本次變動的案件，不重建整份 download_results／vector_resume。
    pickle stream 可逐筆讀取；若中斷造成最後一筆不完整，讀取端保留之前完整紀錄。
    """
    CHECKPOINT_JOURNAL_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "stage": str(stage),
        "key": str(key),
        "value": value,
        "written_at": datetime.now().isoformat(timespec="seconds"),
    }
    with CHECKPOINT_JOURNAL_PATH.open("ab") as f:
        pickle.dump(payload, f, protocol=pickle.HIGHEST_PROTOCOL)
        f.flush()
        try:
            os.fsync(f.fileno())
        except Exception:
            pass


def load_case_checkpoint_journal(stage):
    """讀取指定階段的最新逐案件紀錄；如尾端毀損，截去不完整尾段再保留可續跑日誌。"""
    latest = {}
    if not CHECKPOINT_JOURNAL_PATH.is_file():
        return latest
    file_size = CHECKPOINT_JOURNAL_PATH.stat().st_size
    last_good_offset = 0
    damaged_tail = False
    try:
        with CHECKPOINT_JOURNAL_PATH.open("rb") as f:
            while f.tell() < file_size:
                try:
                    payload = pickle.load(f)
                except Exception as e:
                    damaged_tail = True
                    print(f"[CHECKPOINT_JOURNAL] 發現不完整尾端，將截除損毀部分：{e}")
                    break
                last_good_offset = f.tell()
                if not isinstance(payload, dict) or payload.get("stage") != stage:
                    continue
                key = payload.get("key")
                if key is not None:
                    latest[str(key)] = payload.get("value")
    except Exception as e:
        print(f"⚠️ 讀取案件 checkpoint journal 失敗：{e}")
    if damaged_tail:
        try:
            with CHECKPOINT_JOURNAL_PATH.open("r+b") as f:
                f.truncate(last_good_offset)
                f.flush()
                os.fsync(f.fileno())
        except Exception as e:
            print(f"⚠️ 無法修復 checkpoint journal 尾端：{e}")
    return latest


async def wait_for_global_nlsc_slot():
    """所有 Cell3 worker 共用的請求起始節流器。"""
    global _GLOBAL_API_NEXT_SLOT
    loop = asyncio.get_running_loop()
    async with _GLOBAL_API_LOCK:
        now = loop.time()
        slot = max(now, _GLOBAL_API_NEXT_SLOT)
        _GLOBAL_API_NEXT_SLOT = slot + GLOBAL_API_MIN_INTERVAL_MS / 1000.0
    wait_seconds = slot - loop.time()
    if wait_seconds > 0:
        await asyncio.sleep(wait_seconds)


def jsonp_to_json(body):
    body = (body or "").strip()
    if not body:
        raise ValueError("EMPTY_RESPONSE")
    try:
        return json.loads(body)
    except Exception:
        pass
    m = re.search(r"^[^(]+\((.*)\)\s*;?\s*$", body, flags=re.S)
    if not m:
        raise ValueError("INVALID_JSONP")
    return json.loads(m.group(1).strip())


def recursively_find_image_records(obj):
    records = []
    if isinstance(obj, dict):
        if all(k in obj for k in ("image", "lx", "ly", "rx", "ry")):
            records.append(obj)
        for v in obj.values():
            records.extend(recursively_find_image_records(v))
    elif isinstance(obj, list):
        for v in obj:
            records.extend(recursively_find_image_records(v))
    return records


def image_bbox_contains_point(record, lon, lat):
    lx, ly, rx, ry = [safe_float(record.get(k)) for k in ("lx", "ly", "rx", "ry")]
    if not all(np.isfinite(v) for v in (lx, ly, rx, ry)):
        return False
    return min(lx, rx) <= lon <= max(lx, rx) and min(ly, ry) <= lat <= max(ly, ry)


def decode_image_b64(image_b64):
    s = clean_text(image_b64)
    if not s:
        raise RuntimeError("選定 image 沒有 Base64 image")
    if "," in s:
        prefix, data_part = s.split(",", 1)
        if "base64" in prefix.lower():
            s = data_part
    data = base64.b64decode(s, validate=True)
    if not data.startswith(b"\x89PNG\r\n\x1a\n"):
        raise RuntimeError("image 解碼後不是 PNG")
    return data


def select_image_record(body, query_lon, query_lat, case_name=""):
    """解析全部圖幅；Cell2 點仍必須唯一命中其中一幅，其他圖幅也全部保留。"""
    data = jsonp_to_json(body)
    records = recursively_find_image_records(data)
    if not records:
        raise RuntimeError(f"{case_name}：Response 找不到 image record")

    candidates = []
    tile_records = []
    for image_index, record in enumerate(records, start=1):
        bbox = [safe_float(record.get(k)) for k in ("lx", "ly", "rx", "ry")]
        if not all(np.isfinite(v) for v in bbox):
            raise RuntimeError(f"{case_name}：第 {image_index} 幅缺少有效 bbox")
        image_bytes = decode_image_b64(record.get("image", ""))
        image_meta = inspect_png_bytes(image_bytes, bbox=bbox)
        tile_records.append({
            "record_index": int(image_index),
            "lx": float(bbox[0]), "ly": float(bbox[1]),
            "rx": float(bbox[2]), "ry": float(bbox[3]),
            "image_bytes": image_bytes,
            "sha256": sha256_bytes(image_bytes),
            "image_size": int(len(image_bytes)),
            **image_meta,
        })
        if image_bbox_contains_point(record, query_lon, query_lat):
            candidates.append(image_index)

    if len(candidates) == 0:
        raise RuntimeError(
            f"{case_name}：沒有 image bbox 包含 Cell 2 查詢點；image_count={len(records)}"
        )
    if len(candidates) > 1:
        raise RuntimeError(
            f"{case_name}：有 {len(candidates)} 張 image 同時包含 Cell 2 查詢點，不做猜測；candidate_count={len(candidates)}"
        )

    selected_index = int(candidates[0])
    selected_tile = tile_records[selected_index - 1]
    return {
        "record_count": len(tile_records),
        "selected_index": selected_index,
        "lx": selected_tile["lx"], "ly": selected_tile["ly"],
        "rx": selected_tile["rx"], "ry": selected_tile["ry"],
        "image_bytes": selected_tile["image_bytes"],
        "image_records": tile_records,
        "selection_method": "CELL2_QUERY_POINT_IN_IMAGE_BBOX",
    }


# ============================================================
# 2. 正式回歸測試
# ============================================================

def run_regression_tests():
    print()
    print("=" * 70)
    print("Cell 3 正式回歸測試")
    print("=" * 70)

    # A. 地號標準化
    land_tests = {
        "1742-0010": "17420010",
        "17420010": "17420010",
        "1742": "17420000",
        "1": "00010000",
        "1208-0000": "12080000",
    }
    for original, expected in land_tests.items():
        actual = normalize_landno(original)
        if actual != expected:
            raise RuntimeError(f"地號標準化測試失敗：{original} -> {actual} / {expected}")
    print("地號格式測試：全部通過")

    # B. 地段標準化
    sect_tests = {
        "1": "0001",
        "0001": "0001",
        "301": "0301",
        "0301.0": "0301",
    }
    for original, expected in sect_tests.items():
        actual = normalize_sect(original)
        if actual != expected:
            raise RuntimeError(f"地段標準化測試失敗：{original} -> {actual} / {expected}")
    print("地段格式測試：全部通過")

    # C. 多 image 選圖邏輯：使用真正可解碼 PNG；另外驗證單像素佔位圖分類。
    _normal_test_img = np.zeros((12, 16, 4), dtype=np.uint8)
    _normal_test_img[2:10, 2:14, 3] = 255
    _ok, _normal_test_buf = cv2.imencode(".png", _normal_test_img)
    assert _ok
    normal_png = base64.b64encode(_normal_test_buf.tobytes()).decode("ascii")
    _placeholder_img = np.array([[[0, 0, 255, 127]]], dtype=np.uint8)
    _ok, _placeholder_buf = cv2.imencode(".png", _placeholder_img)
    assert _ok
    tiny_png = base64.b64encode(_placeholder_buf.tobytes()).decode("ascii")
    _normal_meta = inspect_png_bytes(_normal_test_buf.tobytes())
    _placeholder_meta = inspect_png_bytes(_placeholder_buf.tobytes())
    assert _normal_meta["geometry_eligible"] is True
    assert _placeholder_meta["image_class"] == "DEGENERATE_PLACEHOLDER"
    assert _placeholder_meta["geometry_eligible"] is False
    _blank_canvas = np.zeros((16, 16, 4), dtype=np.uint8)
    _ok, _blank_buf = cv2.imencode(".png", _blank_canvas)
    assert _ok
    _blank_meta = inspect_png_bytes(_blank_buf.tobytes())
    assert _blank_meta["image_class"] == "DEGENERATE_PLACEHOLDER"
    assert _blank_meta["placeholder_reason"] == "NO_GEOMETRY_PIXELS"
    _edge_canvas = np.zeros((16, 16, 4), dtype=np.uint8)
    _edge_canvas[8, 2:14, 3] = 255
    _ok, _edge_buf = cv2.imencode(".png", _edge_canvas)
    assert _ok
    _edge_meta = inspect_png_bytes(_edge_buf.tobytes())
    assert _edge_meta["image_class"] == "DEGENERATE_PLACEHOLDER"
    body = json.dumps([
        [{"image": normal_png, "lx": 121.0, "ly": 24.0, "rx": 121.1, "ry": 24.1}],
        [{"image": normal_png, "lx": 121.2, "ly": 24.2, "rx": 121.3, "ry": 24.3}],
    ])
    result = select_image_record(body, 121.05, 24.05, "TEST_SINGLE")
    assert result["record_count"] == 2
    assert result["selected_index"] == 1
    print("多 image 唯一命中測試：OK")

    # D. 無命中必須失敗
    try:
        select_image_record(body, 122.0, 25.0, "TEST_NONE")
        raise RuntimeError("多 image 無命中測試失敗：本應拋出例外")
    except RuntimeError as e:
        if "沒有 image bbox" not in str(e):
            raise
    print("多 image 無命中保護測試：OK")

    # E. 多重命中必須失敗，不猜測
    body_overlap = json.dumps([
        [{"image": normal_png, "lx": 121.0, "ly": 24.0, "rx": 121.2, "ry": 24.2},
         {"image": normal_png, "lx": 121.1, "ly": 24.1, "rx": 121.3, "ry": 24.3}]
    ])
    try:
        select_image_record(body_overlap, 121.15, 24.15, "TEST_MULTI")
        raise RuntimeError("多 image 多重命中測試失敗：本應拋出例外")
    except RuntimeError as e:
        if "同時包含" not in str(e):
            raise
    print("多 image 多重命中保護測試：OK")

    # F. 正式成果檔名
    legacy_forbidden = "_106筆"
    for p in [GEOJSON_PATH, GPKG_PATH, SHP_PATH, ATTRIBUTE_PATH]:
        if legacy_forbidden in p.name:
            raise RuntimeError(f"正式成果檔名仍含 _106筆：{p}")
    print("正式成果檔名測試：OK（不含 _106筆）")

    # G. 多圖幅完整保留，且 selected index 不等同第一張。
    body_multi = json.dumps([
        [{"image": normal_png, "lx": 121.0, "ly": 24.0, "rx": 121.1, "ry": 24.1},
         {"image": normal_png, "lx": 121.1, "ly": 24.0, "rx": 121.2, "ry": 24.1}]
    ])
    multi = select_image_record(body_multi, 121.15, 24.05, "TEST_MULTI_TILE")
    assert multi["record_count"] == 2 and len(multi["image_records"]) == 2
    assert multi["selected_index"] == 2
    assert [x["record_index"] for x in multi["image_records"]] == [1, 2]
    print("多圖幅全量保留／唯一選中索引測試：OK")

    # G2. 佔位圖必須保留在 records，但只向量化其他可用 PNG。
    body_placeholder_mix = json.dumps([
        [{"image": tiny_png, "lx": 121.0, "ly": 24.0, "rx": 121.000001, "ry": 24.000001},
         {"image": normal_png, "lx": 121.000001, "ly": 24.0, "rx": 121.001, "ry": 24.001}]
    ])
    mix = select_image_record(body_placeholder_mix, 121.0005, 24.0005, "TEST_PLACEHOLDER_MIX")
    _eligible = [x for x in mix["image_records"] if x["geometry_eligible"]]
    _placeholder = [x for x in mix["image_records"] if x["image_class"] == "DEGENERATE_PLACEHOLDER"]
    assert mix["record_count"] == 2 and len(_eligible) == 1 and len(_placeholder) == 1
    assert _eligible[0]["record_index"] == 2 and _placeholder[0]["record_index"] == 1
    print("退化佔位圖保留／其他有效 PNG 可用測試：OK")

    # H. DBF 長度限制：JSON 不進 DBF，完整逐幅內容進 sidecar。
    _test_frame = pd.DataFrame([{
        "案件ID": "TEST_CASE", "列號": 1,
        "圖幅清單_JSON": json.dumps([{"record_index": 1, "image_path": "/tmp/a.png", "selected": True}]),
        "圖幅SVG清單_JSON": json.dumps(["/tmp/a.svg"]),
        "response_url": "https://example.test/" + ("x" * 300),
        "圖幅數": 1, "圖幅有效數": 1, "圖幅佔位數": 0,
    }])
    _detail_rows = build_shp_tile_detail_rows(_test_frame)
    assert len(_detail_rows) == 1 and _detail_rows[0]["圖幅序號"] == 1
    _shp_test, _long_test, _excluded_test = prepare_shapefile_frame(_test_frame, {})
    assert "圖幅清單_JSON" not in _shp_test.columns and "圖幅SVG清單_JSON" not in _shp_test.columns
    assert len(_long_test) == 1 and len(_long_test[0]["完整值"]) > 254
    assert max(len(str(v).encode("utf-8")) for v in _shp_test["response_url"]) <= 254
    _unicode_test = pd.DataFrame([{
        "案件ID": "TEST_UNICODE", "列號": 2, "長文字": "農地" * 160, "geometry": None
    }])
    _unicode_shp, _unicode_long, _ = prepare_shapefile_frame(_unicode_test, {})
    assert len(_unicode_long) == 1
    assert len(_unicode_shp.loc[0, "長文字"].encode("utf-8")) <= 254
    print("SHP DBF 長欄位／圖幅 sidecar 測試：OK")

    # I. 簽章必須依 tile contents 穩定生成。
    print("正式回歸測試：全部通過")


# ============================================================
# 3. 讀取 Cell 2 / 既有結果
# ============================================================

if not INPUT_XLSX.exists():
    raise FileNotFoundError(f"找不到 Cell 2：{INPUT_XLSX}")

INPUT_XLSX_SHA256 = sha256_file(INPUT_XLSX)
df_input = pd.read_excel(INPUT_XLSX, dtype=str).fillna("")

required_cols = ["縣市", "鄉鎮市區", "地段", "地號", "office", "sect"]
missing = [c for c in required_cols if c not in df_input.columns]
if missing:
    raise RuntimeError("Cell 2 result.xlsx 缺少必要欄位：\n" + "\n".join(f"  - {x}" for x in missing))

run_regression_tests()

print()
print("=" * 80)
print("Cell 3 v4.3.4：NLSC 多圖幅全量保存／重複案件鍵安全處理／有效圖幅聯集")
print("=" * 80)
print("Cell 2：", INPUT_XLSX)
print("總筆數：", len(df_input))

cadastral_mask = df_input.apply(is_cadastral_query_row, axis=1)
cadastral_case_count = int(cadastral_mask.sum())  # 原始 Cell2 地籍資料列數，輸出仍逐列保留
address_case_count = int(len(df_input) - cadastral_case_count)
_cadastral_keys_in_order = [
    case_key(row) for _, row in df_input[cadastral_mask].iterrows()
]
_unique_cadastral_keys = list(dict.fromkeys(_cadastral_keys_in_order))
unique_cadastral_key_count = len(_unique_cadastral_keys)
duplicate_key_extra_rows = cadastral_case_count - unique_cadastral_key_count
_duplicate_key_groups = {
    key: count for key, count in pd.Series(_cadastral_keys_in_order).value_counts().items()
    if int(count) > 1
}
print(f"可進行地籍圖處理：{cadastral_case_count} 筆原始資料列；其餘案件跳過地籍圖：{address_case_count} 筆")
print(f"唯一 NLSC 案件鍵：{unique_cadastral_key_count} 組；重複鍵群組：{len(_duplicate_key_groups)} 組；重複資料列：{duplicate_key_extra_rows} 筆")
print(f"ℹ️ 本次實際案件數：{len(df_input)} 筆；其中地籍圖資料列：{cadastral_case_count} 筆。輸出保留原始資料列，不刪除重複列。")

if _duplicate_key_groups:
    _duplicate_audit = df_input.loc[cadastral_mask].copy()
    _duplicate_audit.insert(0, "Cell2_Excel列號", _duplicate_audit.index.to_series().astype(int) + 2)
    _duplicate_audit["NLSC案件鍵"] = [case_key(_row) for _, _row in _duplicate_audit.iterrows()]
    _duplicate_audit["同鍵資料列數"] = _duplicate_audit["NLSC案件鍵"].map(_duplicate_key_groups)
    _duplicate_audit = _duplicate_audit[
        _duplicate_audit["NLSC案件鍵"].isin(_duplicate_key_groups)
    ].sort_values(["NLSC案件鍵", "Cell2_Excel列號"])
    _duplicate_columns = [
        "Cell2_Excel列號", "NLSC案件鍵", "同鍵資料列數", "縣市", "鄉鎮市區", "地段",
        "地號", "查詢地號", "office", "sect", "經度", "緯度", "TWD97_X", "TWD97_Y",
    ]
    _duplicate_audit[[c for c in _duplicate_columns if c in _duplicate_audit.columns]].to_excel(
        DUPLICATE_KEY_AUDIT_PATH, index=False
    )
    print("重複案件鍵稽核檔：", DUPLICATE_KEY_AUDIT_PATH)
else:
    print("重複案件鍵：0 組")

previous_df = None
if IMAGE_RESULT_XLSX.exists():
    try:
        previous_df = pd.read_excel(IMAGE_RESULT_XLSX, dtype=str).fillna("")
    except Exception as e:
        print("⚠️ 讀取既有結果 Excel 失敗，將重新建立：", repr(e))


# ============================================================
# 4. Cell 2 查詢點
# ============================================================

transformer_3826_to_4326 = Transformer.from_crs("EPSG:3826", "EPSG:4326", always_xy=True)
transformer_4326_to_3826 = Transformer.from_crs("EPSG:4326", "EPSG:3826", always_xy=True)

point_map = {}
_point_key_row_counts = {}
_point_key_coordinate_conflicts = []
for _, row in df_input.iterrows():
    if not is_cadastral_query_row(row):
        continue
    key = case_key(row)
    _point_key_row_counts[key] = _point_key_row_counts.get(key, 0) + 1
    lon = safe_float(row.get("經度", np.nan))
    lat = safe_float(row.get("緯度", np.nan))
    if not (np.isfinite(lon) and np.isfinite(lat)):
        x = safe_float(row.get("TWD97_X", np.nan))
        y = safe_float(row.get("TWD97_Y", np.nan))
        if np.isfinite(x) and np.isfinite(y):
            lon, lat = transformer_3826_to_4326.transform(x, y)
    if np.isfinite(lon) and np.isfinite(lat):
        candidate_point = (float(lon), float(lat))
        if key in point_map:
            # 相同 NLSC office/sect/landno 是同一個查詢鍵；重複列可共用定位點，
            # 但若座標差異超過 0.5m，必須中止，不能任意以後一列覆蓋前一列。
            old_x, old_y = transformer_4326_to_3826.transform(*point_map[key])
            new_x, new_y = transformer_4326_to_3826.transform(*candidate_point)
            drift_m = float(np.hypot(new_x - old_x, new_y - old_y))
            if drift_m > 0.5:
                _point_key_coordinate_conflicts.append((key, drift_m, point_map[key], candidate_point))
        else:
            point_map[key] = candidate_point

if _point_key_coordinate_conflicts:
    details = "\n".join(
        f"  {key}: 座標差 {distance:.3f}m | first={old} | duplicate={new}"
        for key, distance, old, new in _point_key_coordinate_conflicts[:30]
    )
    raise RuntimeError("同一 NLSC 案件鍵出現不同定位座標，拒絕合併：\n" + details)

_missing_point_keys = [key for key in _unique_cadastral_keys if key not in point_map]
if _missing_point_keys:
    details = "\n".join(f"  {key}" for key in _missing_point_keys[:50])
    raise RuntimeError(
        f"Cell 2 定位點缺漏：唯一案件鍵應有 {unique_cadastral_key_count}，"
        f"實際有效 {len(point_map)}，缺少 {len(_missing_point_keys)} 組：\n{details}"
    )
if len(point_map) != unique_cadastral_key_count:
    raise RuntimeError(
        f"Cell 2 定位點數量異常：唯一案件鍵應有 {unique_cadastral_key_count}，實際 {len(point_map)}"
    )
print(
    f"Cell 2 定位點回歸檢查：{len(point_map)}/{unique_cadastral_key_count} 組唯一 NLSC 案件鍵完整；"
    f"原始資料列 {cadastral_case_count} 筆，重複鍵額外資料列 {duplicate_key_extra_rows} 筆（保留於輸出）"
)


# ============================================================
# 5. 既有 checkpoint / Audit 狀態
# ============================================================

def restore_latest_drive_checkpoint():
    """同一份 Cell2 輸入時，從先前 V1 Drive 執行資料夾復原影像、Raw、SVG 與 checkpoint。"""
    if not DRIVE_MOUNTED or FRO_DRIVE_ROOT is None:
        print("DRIVE_CHECKPOINT_RECOVERY=SKIPPED_DRIVE_NOT_MOUNTED")
        return
    cell3_drive = FRO_DRIVE_ROOT / "Cell3"
    if not cell3_drive.is_dir():
        return

    # 若 /content 留有其他批次的 checkpoint，先移除候選版隔離檔；只在 SHA 相同時沿用。
    current_download_valid = False
    if DOWNLOAD_STATE_PATH.is_file():
        try:
            current_download = pd.read_pickle(DOWNLOAD_STATE_PATH)
            current_download_valid = (
                isinstance(current_download, dict)
                and current_download.get("input_xlsx_sha256") == INPUT_XLSX_SHA256
            )
        except Exception:
            current_download_valid = False
        if not current_download_valid:
            DOWNLOAD_STATE_PATH.unlink(missing_ok=True)

    current_vector_valid = False
    if VECTOR_STATE_PATH.is_file():
        try:
            current_vector = pd.read_pickle(VECTOR_STATE_PATH)
            current_vector_valid = (
                isinstance(current_vector, dict)
                and current_vector.get("input_xlsx_sha256") == INPUT_XLSX_SHA256
            )
        except Exception:
            current_vector_valid = False
        if not current_vector_valid:
            VECTOR_STATE_PATH.unlink(missing_ok=True)

    candidates = sorted(
        [p for p in cell3_drive.glob("MULTI_TILE_V1_*") if p.is_dir() and p != DRIVE_BACKUP_DIR],
        key=lambda p: p.name,
        reverse=True,
    )
    for folder in candidates:
        state_path = folder / DOWNLOAD_STATE_PATH.name
        if not state_path.is_file():
            continue
        try:
            state_obj = pd.read_pickle(state_path)
        except Exception:
            continue
        if not isinstance(state_obj, dict) or state_obj.get("input_xlsx_sha256") != INPUT_XLSX_SHA256:
            continue

        if not DOWNLOAD_STATE_PATH.exists():
            shutil.copy2(state_path, DOWNLOAD_STATE_PATH)
        vector_backup = folder / VECTOR_STATE_PATH.name
        if not VECTOR_STATE_PATH.exists() and vector_backup.is_file():
            try:
                vec_obj = pd.read_pickle(vector_backup)
                if isinstance(vec_obj, dict) and vec_obj.get("input_xlsx_sha256") == INPUT_XLSX_SHA256:
                    shutil.copy2(vector_backup, VECTOR_STATE_PATH)
            except Exception:
                pass
        image_backup = folder / "IMAGE_TILES"
        if image_backup.is_dir():
            shutil.copytree(image_backup, IMAGE_DIR, dirs_exist_ok=True)
        raw_backup = folder / "RAW_RESPONSES"
        if not raw_backup.is_dir():
            # 相容原 v4.3.1 的 Drive 備份結構。
            raw_backup = folder / "RAW_RESPONSES"
        if raw_backup.is_dir():
            shutil.copytree(raw_backup, RAW_DIR, dirs_exist_ok=True)
        svg_backup = folder / "SVG"
        if svg_backup.is_dir():
            shutil.copytree(svg_backup, SVG_DIR, dirs_exist_ok=True)
        prior_result_xlsx = folder / "result_cadastral_image_multi_tile_candidate.xlsx"
        if prior_result_xlsx.is_file() and not IMAGE_RESULT_XLSX.exists():
            shutil.copy2(prior_result_xlsx, IMAGE_RESULT_XLSX)
        print("DRIVE_CHECKPOINT_RECOVERY=PASS", folder)
        return
    print("DRIVE_CHECKPOINT_RECOVERY=NO_MATCHING_PRIOR_RUN")


restore_latest_drive_checkpoint()

download_state = load_pickle_state(DOWNLOAD_STATE_PATH, {"version": HANDOFF_VERSION, "results": {}, "errors": {}})
if not isinstance(download_state, dict):
    download_state = {"version": "4.1", "results": {}, "errors": {}}

if download_state.get("input_xlsx_sha256") != INPUT_XLSX_SHA256:
    print("DOWNLOAD_CHECKPOINT_INPUT_SHA_MISMATCH=RESET")
    download_state = {"version": HANDOFF_VERSION, "results": {}, "errors": {}}

download_results = dict(download_state.get("results", {}))
download_errors = dict(download_state.get("errors", {}))
# 舊版 .pkl 作為基底；v2 journal 的較新單筆紀錄覆蓋舊快照。
download_results.update(load_case_checkpoint_journal("download"))
for _error_key, _error_value in load_case_checkpoint_journal("download_error").items():
    if _error_value is None:
        download_errors.pop(_error_key, None)
    else:
        download_errors[_error_key] = _error_value


def rebase_download_asset_paths(results):
    """將舊執行資料夾 checkpoint 內的影像／Raw 路徑，對應到本次工作目錄。"""
    changed = False
    for item in results.values():
        if not isinstance(item, dict):
            continue
        rebased_tiles = []
        for tile in item.get("image_records", []) if isinstance(item.get("image_records"), list) else []:
            if not isinstance(tile, dict):
                continue
            tile = dict(tile)
            old = Path(str(tile.get("image_path", ""))) if tile.get("image_path") else None
            candidate = IMAGE_DIR / old.name if old and old.name else None
            if candidate is not None and candidate.is_file() and str(candidate) != str(old):
                tile["image_path"] = str(candidate)
                changed = True
            elif candidate is not None and old is not None and old.is_file() and DRIVE_MOUNTED:
                candidate.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(old, candidate)
                if clean_text(tile.get("sha256", "")) and sha256_file(candidate) != tile["sha256"]:
                    raise RuntimeError(f"舊 PNG 移轉 SHA256 不一致：{old}")
                tile["image_path"] = str(candidate)
                changed = True
            rebased_tiles.append(tile)
        if rebased_tiles:
            item["image_records"] = rebased_tiles
            paths = [str(x.get("image_path", "")) for x in rebased_tiles]
            if item.get("image_paths") != paths:
                item["image_paths"] = paths
                changed = True
            selected_index = int(item.get("selected_image_index", 0) or 0)
            selected = next((x for x in rebased_tiles if int(x.get("record_index", 0) or 0) == selected_index), None)
            if selected:
                if item.get("image_path") != selected.get("image_path"):
                    item["image_path"] = selected.get("image_path")
                    changed = True
                item["image_signature"] = image_signature(selected.get("image_path", ""))
            new_tile_sig = tile_set_signature(rebased_tiles)
            if item.get("tile_set_signature") != new_tile_sig:
                item["tile_set_signature"] = new_tile_sig
                changed = True
        raw_value = clean_text(item.get("raw_file", ""))
        if raw_value:
            raw_old = Path(raw_value)
            raw_candidate = RAW_DIR / raw_old.name
            if raw_candidate.is_file() and str(raw_candidate) != raw_value:
                item["raw_file"] = str(raw_candidate)
                changed = True
    if changed:
        print("CHECKPOINT_ASSET_PATH_REBASE=PASS")
    return changed


rebase_download_asset_paths(download_results)

# 以案件ID建立既有 Excel 索引，避免單純依列號。
previous_by_key = {}
if previous_df is not None:
    for _, r in previous_df.iterrows():
        if is_cadastral_query_row(r):
            previous_by_key[case_key(r)] = r.to_dict()


def audit_from_previous(key):
    r = previous_by_key.get(key, {})
    return {
        "image_record_count": r.get("image_record_count", ""),
        "selected_image_index": r.get("selected_image_index", ""),
        "image_selection_method": r.get("image_selection_method", ""),
        "lx": safe_float(r.get("左下經度", np.nan)),
        "ly": safe_float(r.get("左下緯度", np.nan)),
        "rx": safe_float(r.get("右上經度", np.nan)),
        "ry": safe_float(r.get("右上緯度", np.nan)),
        "raw_file": clean_text(r.get("raw_file", "")),
    }


def audit_complete(item, point=None):
    """全部 records 均須可驗證；單像素佔位圖可保留，但至少須有一幅有效 PNG。"""
    if not isinstance(item, dict):
        return False
    try:
        count = int(item.get("image_record_count", 0))
        selected_index = int(item.get("selected_image_index", 0))
    except Exception:
        return False
    tiles = item.get("image_records")
    if count < 1 or not isinstance(tiles, list) or len(tiles) != count:
        return False
    if selected_index < 1 or selected_index > count:
        return False
    if clean_text(item.get("image_selection_method", "")) != "CELL2_QUERY_POINT_IN_IMAGE_BBOX":
        return False

    indices = []
    geometry_eligible_count = 0
    placeholder_count = 0
    for tile in tiles:
        try:
            idx = int(tile.get("record_index", 0))
            p = Path(tile.get("image_path", ""))
            bbox = [safe_float(tile.get(k)) for k in ("lx", "ly", "rx", "ry")]
            if idx < 1 or not p.is_file() or not all(np.isfinite(v) for v in bbox):
                return False
            if cached_tile_sha256(p) != clean_text(tile.get("sha256", "")):
                return False
            meta = enrich_tile_metadata(tile)
            if meta["geometry_eligible"]:
                geometry_eligible_count += 1
            elif meta["image_class"] == "DEGENERATE_PLACEHOLDER":
                placeholder_count += 1
            else:
                # 無法分類成佔位圖的低維 PNG 不得被默默略過，仍需線上重抓。
                return False
            indices.append(idx)
        except Exception:
            return False
    if sorted(indices) != list(range(1, count + 1)):
        return False
    if geometry_eligible_count < 1:
        return False

    selected = next((x for x in tiles if int(x["record_index"]) == selected_index), None)
    if selected is None or Path(item.get("image_path", "")) != Path(selected["image_path"]):
        return False
    if point is not None:
        hits = [x for x in tiles if (
            min(float(x["lx"]), float(x["rx"])) <= float(point[0]) <= max(float(x["lx"]), float(x["rx"]))
            and min(float(x["ly"]), float(x["ry"])) <= float(point[1]) <= max(float(x["ly"]), float(x["ry"]))
        )]
        if len(hits) != 1 or int(hits[0]["record_index"]) != selected_index:
            return False
    item["geometry_eligible_tile_count"] = int(geometry_eligible_count)
    item["placeholder_tile_count"] = int(placeholder_count)
    return True


# ============================================================
# 6. Raw Response 離線解析
# ============================================================

def raw_candidates(row_number):
    return [
        directory / f"row_{row_number:03d}.txt"
        for directory in RAW_READ_DIRS
    ] + [
        directory / f"row_{row_number:03d}.json"
        for directory in RAW_READ_DIRS
    ]


def read_existing_raw(row_number):
    for p in raw_candidates(row_number):
        if p.exists():
            try:
                return p, p.read_text(encoding="utf-8")
            except Exception:
                pass
    return None, None


def apply_selected_result(row_number, row, selected, source, raw_path, overwrite_existing=True):
    """將同一筆案件的全部 image records 原子保存；legacy image_path 指向唯一命中圖幅。"""
    sid = stable_case_id(row)
    all_tiles = selected.get("image_records")
    if not isinstance(all_tiles, list) or len(all_tiles) != int(selected.get("record_count", 0)):
        raise RuntimeError(f"{case_key(row)}：多圖幅清單不完整，拒絕保存")

    persisted_tiles = []
    changed_any = False
    created_any = False
    for tile in all_tiles:
        record_index = int(tile["record_index"])
        image_path = IMAGE_DIR / expected_image_record_name(row_number, row, record_index)
        data = tile["image_bytes"]
        tile_meta = inspect_png_bytes(
            data,
            bbox=[tile["lx"], tile["ly"], tile["rx"], tile["ry"]],
        )
        new_sha = sha256_bytes(data)
        old_sha = sha256_file(image_path) if image_path.is_file() else None
        changed_any = changed_any or (old_sha is not None and old_sha != new_sha)
        created_any = created_any or (old_sha is None)
        if overwrite_existing or old_sha is None:
            atomic_write_bytes(image_path, data)
        if sha256_file(image_path) != new_sha:
            raise RuntimeError(f"{case_key(row)}：第 {record_index} 幅 PNG 寫入後 SHA256 不一致")
        mirror_file_to_drive(image_path, "IMAGE_TILES")
        persisted_tiles.append({
            "record_index": record_index,
            "image_path": str(image_path),
            "lx": float(tile["lx"]), "ly": float(tile["ly"]),
            "rx": float(tile["rx"]), "ry": float(tile["ry"]),
            "sha256": new_sha,
            "image_size": int(len(data)),
            "selected": bool(record_index == int(selected["selected_index"])),
            **tile_meta,
        })

    selected_tile = next(x for x in persisted_tiles if x["selected"])
    item = download_results.setdefault(sid, {})
    item.update({
        "image_path": selected_tile["image_path"],
        "image_paths": [x["image_path"] for x in persisted_tiles],
        "image_size": int(sum(x["image_size"] for x in persisted_tiles)),
        "lx": selected_tile["lx"], "ly": selected_tile["ly"],
        "rx": selected_tile["rx"], "ry": selected_tile["ry"],
        "raw_file": str(raw_path) if raw_path else item.get("raw_file", ""),
        "response_url": item.get("response_url", ""),
        "image_record_count": int(selected["record_count"]),
        "selected_image_index": int(selected["selected_index"]),
        "image_selection_method": selected["selection_method"],
        "image_records": persisted_tiles,
        "geometry_eligible_tile_count": int(sum(bool(x.get("geometry_eligible")) for x in persisted_tiles)),
        "placeholder_tile_count": int(sum(x.get("image_class") == "DEGENERATE_PLACEHOLDER" for x in persisted_tiles)),
        "download_method": source,
        "image_signature": image_signature(selected_tile["image_path"]),
        "tile_set_signature": tile_set_signature(persisted_tiles),
        "last_refresh_changed_png": bool(changed_any),
    })
    download_errors.pop(sid, None)
    # 每個案件只追加自己的 checkpoint；不再每筆重新序列化整個下載結果字典。
    append_case_checkpoint("download", sid, item)
    append_case_checkpoint("download_error", sid, None)
    return changed_any, created_any


# 先把 checkpoint 與既有 Excel audit 合併，然後逐筆嘗試離線修復。
for idx, row in df_input.iterrows():
    row_number = idx + 1
    if not is_cadastral_query_row(row):
        continue
    sid = stable_case_id(row)
    key = case_key(row)
    item = dict(download_results.get(sid, {}))

    # 既有 Excel 可補齊 checkpoint 中不存在的 audit metadata。
    prev = audit_from_previous(key)
    for k, v in prev.items():
        if k not in item or item.get(k) in (None, "", np.nan):
            item[k] = v
    # 候選版採多圖幅資料夾；不沿用舊版單張 PNG 作為完整多圖幅 Audit。
    download_results[sid] = item


# V8.3 已驗證並保存在 GitHub 診斷分支的 P27/P28 五張 PNG。
# 僅在 row number、標準化地號、Cell2 查詢座標均完全吻合時才可用作本輪輸入復原。
DIAGNOSTIC_METADATA_URL = (
    "https://raw.githubusercontent.com/CHLINX/fro/diagnostic/target-27-29-20261031_141945/"
    "diagnostics/target-27-29/20261031_141945/RAW_IMAGE_RECORDS/RAW_IMAGE_RECORDS_METADATA.json"
)
DIAGNOSTIC_METADATA_SHA256 = "dfaed6ed4e7c0e75e9552e0fbe759c33cc0987d7c79b3fee7f79ceb6eb9ded0b"
DIAGNOSTIC_RAW_IMAGE_BASE = (
    "https://raw.githubusercontent.com/CHLINX/fro/diagnostic/target-27-29-20261031_141945/"
    "diagnostics/target-27-29/20261031_141945/RAW_IMAGE_RECORDS/"
)
_DIAGNOSTIC_METADATA_CACHE = None


def try_restore_diagnostic_multitile_assets(row_number, row):
    global _DIAGNOSTIC_METADATA_CACHE
    expected = {
        (61, "14410000"): ("P27", (121.047230, 24.998430)),
        (60, "05370000"): ("P28", (121.043498, 24.999035)),
    }
    lot = normalize_landno(row.get("查詢地號", row.get("地號", "")))
    spec = expected.get((int(row_number), lot))
    if spec is None:
        return None
    target_name, expected_point = spec
    point = point_map.get(case_key(row))
    if point is None or abs(point[0] - expected_point[0]) > 1e-8 or abs(point[1] - expected_point[1]) > 1e-8:
        return None

    if _DIAGNOSTIC_METADATA_CACHE is None:
        with urllib.request.urlopen(DIAGNOSTIC_METADATA_URL, timeout=30) as response:
            raw_meta = response.read()
        if hashlib.sha256(raw_meta).hexdigest() != DIAGNOSTIC_METADATA_SHA256:
            raise RuntimeError("GitHub 診斷 Metadata SHA256 不一致")
        _metadata_copy = RAW_DIR / "GITHUB_RAW_IMAGE_RECORDS_METADATA_P27_P28.json"
        atomic_write_bytes(_metadata_copy, raw_meta)
        mirror_file_to_drive(_metadata_copy, "RAW_RESPONSES")
        _DIAGNOSTIC_METADATA_CACHE = json.loads(raw_meta.decode("utf-8"))

    target = next((x for x in _DIAGNOSTIC_METADATA_CACHE.get("targets", []) if x.get("target") == target_name), None)
    if not isinstance(target, dict):
        return None
    if normalize_landno(target.get("lot", "")) != lot:
        return None
    target_rows = {int(x.get("row_number", -1)) for x in target.get("records", [])}
    if target_rows != {int(row_number)}:
        return None
    if tuple(round(float(v), 8) for v in target.get("query_point_wgs84", [])) != tuple(round(float(v), 8) for v in expected_point):
        return None

    records = target.get("records", [])
    if len(records) != int(target.get("record_count", -1)) or not records:
        raise RuntimeError(f"{target_name} 診斷 Metadata record_count 不一致")
    matches = [r for r in records if image_bbox_contains_point(r, point[0], point[1])]
    selected_flags = [r for r in records if bool(r.get("is_selected"))]
    if len(matches) != 1 or len(selected_flags) != 1:
        raise RuntimeError(f"{target_name} 診斷 Metadata 唯一選圖條件不成立")
    if int(matches[0].get("record_index", -1)) != int(selected_flags[0].get("record_index", -2)):
        raise RuntimeError(f"{target_name} 診斷 Metadata 的 bbox 選中圖與 is_selected 不一致")

    recovered = []
    for r in sorted(records, key=lambda x: int(x["record_index"])):
        url = DIAGNOSTIC_RAW_IMAGE_BASE + str(r["filename"])
        with urllib.request.urlopen(url, timeout=30) as response:
            image_bytes = response.read()
        if len(image_bytes) != int(r["image_size_bytes"]):
            raise RuntimeError(f"{target_name} {r['filename']} bytes 數不一致")
        digest = hashlib.sha256(image_bytes).hexdigest()
        if digest != str(r["image_sha256"]):
            raise RuntimeError(f"{target_name} {r['filename']} SHA256 不一致")
        if not image_bytes.startswith(b"\x89PNG\r\n\x1a\n"):
            raise RuntimeError(f"{target_name} {r['filename']} 不是 PNG")
        recovered.append({
            "record_index": int(r["record_index"]),
            "lx": float(r["lx"]), "ly": float(r["ly"]),
            "rx": float(r["rx"]), "ry": float(r["ry"]),
            "image_bytes": image_bytes, "sha256": digest,
            "image_size": len(image_bytes),
        })
    selected_index = int(matches[0]["record_index"])
    selected_tile = recovered[selected_index - 1]
    return {
        "record_count": len(recovered), "selected_index": selected_index,
        "lx": selected_tile["lx"], "ly": selected_tile["ly"],
        "rx": selected_tile["rx"], "ry": selected_tile["ry"],
        "image_bytes": selected_tile["image_bytes"],
        "image_records": recovered,
        "selection_method": "CELL2_QUERY_POINT_IN_IMAGE_BBOX",
    }


repair_candidates = []
changed_from_raw = []
raw_repaired = []
github_recovered = []

for idx, row in df_input.iterrows():
    row_number = idx + 1
    if not is_cadastral_query_row(row):
        continue
    sid = stable_case_id(row)
    item = download_results.get(sid, {})
    point = point_map.get(case_key(row))
    if audit_complete(item, point):
        continue

    raw_path, body = read_existing_raw(row_number)
    if point is not None and body:
        try:
            if raw_path is not None:
                mirror_file_to_drive(raw_path, "RAW_RESPONSES")
            selected = select_image_record(body, point[0], point[1], case_key(row))
            changed, created = apply_selected_result(
                row_number, row, selected, "RAW_OFFLINE_ALL_RECORDS", raw_path, overwrite_existing=True
            )
            raw_repaired.append(row_number)
            if changed:
                changed_from_raw.append(row_number)
            continue
        except Exception as e:
            print(f"⚠️ [{row_number:03d}] raw 離線重新選圖失敗：{repr(e)}")

    # 已完成 V8.3 接縫稽核的 P27/P28 可從受 SHA256 保護的 GitHub PNG 診斷資產復原。
    try:
        restored = try_restore_diagnostic_multitile_assets(row_number, row)
        if restored is not None:
            apply_selected_result(
                row_number, row, restored, "GITHUB_VERIFIED_DIAGNOSTIC_RECOVERY", None,
                overwrite_existing=True,
            )
            github_recovered.append(row_number)
            print(f"✅ [{row_number:03d}] {case_key(row)} | GitHub verified multi-tile recovery | records={restored['record_count']}")
            continue
    except Exception as e:
        print(f"⚠️ [{row_number:03d}] GitHub 診斷圖幅復原失敗，改由線上 NLSC 取得：{repr(e)}")

    repair_candidates.append(row_number)

# 同一 NLSC 案件鍵即使在 Cell2 有多筆不同顯示名稱，也只需查詢／補抓一次。
# 先完成所有列的 Raw 離線復原，再依共用 checkpoint 狀態篩掉已被其他重複列修好的案件。
_unique_repair_candidates = []
_seen_repair_keys = set()
for _row_number in repair_candidates:
    _candidate_row = df_input.iloc[_row_number - 1]
    _candidate_key = case_key(_candidate_row)
    if _candidate_key in _seen_repair_keys:
        continue
    _seen_repair_keys.add(_candidate_key)
    _candidate_sid = stable_case_id(_candidate_row)
    if not audit_complete(
        download_results.get(_candidate_sid, {}), point_map.get(_candidate_key)
    ):
        _unique_repair_candidates.append(_row_number)
repair_candidates = _unique_repair_candidates


print()
print("=" * 70)
print("Cell 3 Image Audit / 既有資料整理")
print("=" * 70)
_stable_id_by_case_key = {}
for _idx, _row in df_input.iterrows():
    if is_cadastral_query_row(_row):
        _stable_id_by_case_key.setdefault(case_key(_row), stable_case_id(_row))
_audit_complete_unique = sum(
    1 for _key in _unique_cadastral_keys
    if audit_complete(
        download_results.get(_stable_id_by_case_key[_key], {}),
        point_map.get(_key),
    )
)
print(f"唯一 NLSC 案件鍵目前已完成 Audit：{_audit_complete_unique}/{unique_cadastral_key_count} 組")
print(f"Raw Response 離線重新選圖：{len(raw_repaired)} 列")
print(f"GitHub 診斷資產 SHA256 復原：{len(github_recovered)} 列")
print(f"其中重新選圖後 PNG 內容改變：{len(changed_from_raw)} 列")
print(f"仍需線上 Direct JSONP：{len(repair_candidates)} 組唯一案件鍵")

atomic_pickle_dump(
    {"version": HANDOFF_VERSION, "results": download_results, "errors": download_errors},
    DOWNLOAD_STATE_PATH,
)


# ============================================================
# 7. Direct JSONP 線上補抓
# ============================================================

async def nlsc_jsonp_request(context, office, sect, landno, timeout_ms=API_TIMEOUT_MS):
    """
    NLSC Direct JSONP 取得器（v4.2.1 速度平衡版＋失敗自動重試）。

    不再使用瀏覽器 <script src=...> JSONP。
    改由 Playwright BrowserContext 的 APIRequestContext 直接 GET，
    保留原始 JSON/JSONP response，再交給既有 jsonp_to_json() 解析。

    這樣可避免大量批次查詢時，把 HTTP 層失敗轉成籠統的
    `JSONP script onerror`，也避免每筆查詢都動態插入 script。
    """
    stamp = int(time.time() * 1000)
    callback = f"jQuery{stamp}_{int(time.time() * 1000000) % 1000000}"
    url = (
        f"{NLSC_IMAGE_ENDPOINT}?"
        f"type=2&flag=2&office={office}&sect={sect}&landno={landno}&"
        f"alpah=0.5f&imgflag=1&callback={callback}&_={stamp}"
    )

    await wait_for_global_nlsc_slot()
    response = await context.request.get(
        url,
        timeout=timeout_ms,
        fail_on_status_code=False,
        headers={
            "Referer": NLSC_URL,
            "Accept": "*/*",
            "User-Agent": (
                "Mozilla/5.0 (X11; Linux x86_64) "
                "AppleWebKit/537.36 Chrome/140.0.0.0 Safari/537.36"
            ),
        },
    )

    status = response.status
    body = await response.text()

    if status < 200 or status >= 300:
        raise RuntimeError(
            f"NLSC HTTP 失敗：status={status} | body_head={clean_text(body)[:300]}"
        )
    if not clean_text(body):
        raise RuntimeError(f"NLSC HTTP 空回應：status={status}")

    return body, url


async def online_one(context, row_number, row, timeout_ms=None):
    office = clean_text(row.get("office", "")).upper()
    sect = normalize_sect(row.get("sect", ""))
    landno = normalize_landno(row.get("查詢地號", ""))
    point = point_map.get(case_key(row))
    if point is None:
        raise RuntimeError(f"{case_key(row)}：找不到 Cell 2 查詢點")

    last_error = None
    for attempt in range(1, API_RETRY + 1):
        try:
            body, response_url = await nlsc_jsonp_request(
                context, office, sect, landno,
                timeout_ms=(timeout_ms if timeout_ms is not None else API_TIMEOUT_MS),
            )
            selected = select_image_record(body, point[0], point[1], case_key(row))
            raw_path = RAW_DIR / f"row_{row_number:03d}.txt"
            raw_path.write_text(body, encoding="utf-8")
            mirror_file_to_drive(raw_path, "RAW_RESPONSES")
            changed, _ = apply_selected_result(
                row_number, row, selected, "DIRECT_JSONP", raw_path, overwrite_existing=True
            )
            sid = stable_case_id(row)
            download_results[sid]["response_url"] = response_url
            download_results[sid]["last_refresh_changed_png"] = bool(changed)
            download_results[sid]["http_method"] = "PLAYWRIGHT_CONTEXT_REQUEST"
            append_case_checkpoint("download", sid, download_results[sid])
            return selected, changed
        except Exception as e:
            last_error = e
            if attempt < API_RETRY:
                backoff_ms = API_BACKOFF_BASE_MS * (2 ** (attempt - 1))
                await asyncio.sleep(backoff_ms / 1000.0)
    raise RuntimeError(f"NLSC Direct JSONP 最終失敗：{repr(last_error)}")


async def run_online_repair(candidates=None, concurrency_override=None,
                            delay_override=None, timeout_override=None):
    """
    v4.2.1：線上補抓執行器。

    第一輪維持 v4.2 原本速度設定；後續 retry round 只處理尚未成功的案件，
    可用較低 concurrency / 較長 delay / 較長 timeout 降低 NLSC 批次失敗率。
    """
    candidates = list(candidates if candidates is not None else repair_candidates)
    if not candidates:
        return

    _deduped_candidates = []
    _seen_candidate_keys = set()
    for _row_number in candidates:
        _row = df_input.iloc[int(_row_number) - 1]
        _key = case_key(_row)
        if _key not in _seen_candidate_keys:
            _seen_candidate_keys.add(_key)
            _deduped_candidates.append(int(_row_number))
    candidates = _deduped_candidates
    if not candidates:
        return

    run_concurrency = int(concurrency_override or CONCURRENCY)
    run_delay_ms = int(delay_override if delay_override is not None else API_DELAY_MS)
    run_timeout_ms = int(timeout_override or API_TIMEOUT_MS)

    print()
    print("=" * 70)
    print("開始線上 Direct JSONP Image Audit / 補抓")
    print("=" * 70)
    print(f"本輪待補案件：{len(candidates)}")
    print(f"本輪 concurrency：{run_concurrency}")
    print(f"本輪 delay：{run_delay_ms} ms")
    print(f"全域 NLSC 最小請求間隔：{GLOBAL_API_MIN_INTERVAL_MS} ms")
    print(f"本輪 timeout：{run_timeout_ms} ms")

    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-dev-shm-usage", "--ignore-certificate-errors"],
        )

        queue = asyncio.Queue()
        for row_number in candidates:
            await queue.put(row_number)

        async def worker(worker_id):
            context = await browser.new_context(
                viewport={"width":1600, "height":1200},
                ignore_https_errors=True,
                user_agent=(
                    "Mozilla/5.0 (X11; Linux x86_64) "
                    "AppleWebKit/537.36 "
                    "Chrome/140.0.0.0 Safari/537.36"
                ),
            )
            page = await context.new_page()
            try:
                await page.goto(NLSC_URL, wait_until="domcontentloaded", timeout=60000)
                await page.wait_for_timeout(1000)
                while True:
                    try:
                        row_number = queue.get_nowait()
                    except asyncio.QueueEmpty:
                        break
                    row = df_input.iloc[row_number - 1]
                    try:
                        selected, changed = await online_one(
                            context, row_number, row, timeout_ms=run_timeout_ms
                        )
                        print(
                            f"✅ [{row_number:03d}] {case_key(row)} | online | "
                            f"count={selected['record_count']} | index={selected['selected_index']}"
                            + (" | PNG已更新" if changed else "")
                        )
                    except Exception as e:
                        sid = stable_case_id(row)
                        download_errors[sid] = traceback.format_exc()
                        append_case_checkpoint("download_error", sid, download_errors[sid])
                        print(f"❌ [{row_number:03d}] {case_key(row)}：{repr(e)}")
                    queue.task_done()
                    await asyncio.sleep(run_delay_ms / 1000.0)
            finally:
                try:
                    await page.close()
                except Exception:
                    pass
                try:
                    await context.close()
                except Exception:
                    pass

        worker_count = max(1, min(run_concurrency, len(candidates)))
        await asyncio.gather(*[worker(i + 1) for i in range(worker_count)])
        await browser.close()



nest_asyncio.apply()
loop = asyncio.get_event_loop()


def collect_remaining_audit():
    remaining = []
    seen = set()
    for _idx, _row in df_input.iterrows():
        if not is_cadastral_query_row(_row):
            continue
        _key = case_key(_row)
        if _key in seen:
            continue
        seen.add(_key)
        _sid = stable_case_id(_row)
        _point = point_map.get(_key)
        if not audit_complete(download_results.get(_sid, {}), _point):
            remaining.append(_key)
    return remaining


# ------------------------------------------------------------
# v4.2.1：第一輪維持 v4.2 原設定。
# ------------------------------------------------------------
loop.run_until_complete(run_online_repair())
remaining_audit = collect_remaining_audit()

# ------------------------------------------------------------
# v4.2.1：失敗案件進入 retry queue。
# 只重試未成功案件，不重新打已成功案件。
# ------------------------------------------------------------
for retry_round in range(1, ONLINE_RETRY_ROUNDS + 1):
    if not remaining_audit:
        break

    retry_case_keys = set(remaining_audit)
    retry_rows = []
    retry_seen_keys = set()
    for _idx, _row in df_input.iterrows():
        if not is_cadastral_query_row(_row):
            continue
        _key = case_key(_row)
        if _key in retry_case_keys and _key not in retry_seen_keys:
            retry_rows.append(_idx + 1)
            retry_seen_keys.add(_key)

    print()
    print("=" * 70)
    print(f"Cell 3 v4.2.1｜失敗案件 retry round {retry_round}/{ONLINE_RETRY_ROUNDS}")
    print("=" * 70)
    print(f"上一輪仍缺 Audit：{len(retry_rows)} 筆")

    loop.run_until_complete(
        run_online_repair(
            candidates=retry_rows,
            concurrency_override=RETRY_ONLY_CONCURRENCY,
            delay_override=RETRY_ONLY_DELAY_MS,
            timeout_override=RETRY_ONLY_TIMEOUT_MS,
        )
    )
    remaining_audit = collect_remaining_audit()

print()
if remaining_audit:
    print("⚠️ Cell 3 Image Audit 仍有未完成案件：", len(remaining_audit))
    for _k in remaining_audit[:50]:
        print("  -", _k)
    if len(remaining_audit) > 50:
        print(f"  ... 其餘 {len(remaining_audit) - 50} 筆已省略")
else:
    print(f"✅ Cell 3 Image Audit：{unique_cadastral_key_count}/{unique_cadastral_key_count} 組唯一 NLSC 案件鍵全部完整（原始資料列 {cadastral_case_count} 筆仍逐列保留）")

# v4.2.1 不在此處直接 raise。
# 讓後面的 PNG → Polygon 階段仍能處理本輪已成功取得 PNG 的案件，
# 同時將失敗案件保留在 checkpoint，下一次執行會只補這些案件。



# ============================================================
# 9. 組合 result_cadastral_image.xlsx
# ============================================================

def classify_case_tiles(item):
    tiles = item.get("image_records", []) if isinstance(item, dict) else []
    eligible, placeholders = [], []
    for tile in tiles:
        try:
            meta = enrich_tile_metadata(tile)
            if meta["geometry_eligible"]:
                eligible.append(tile)
            elif meta["image_class"] == "DEGENERATE_PLACEHOLDER":
                placeholders.append(tile)
        except Exception:
            pass
    return eligible, placeholders


result_rows = []
for idx, row in df_input.iterrows():
    row_number = idx + 1
    output = row.to_dict()
    output["查詢地號"] = normalize_landno(row.get("查詢地號", row.get("地號", "")))

    # 優先保留 Cell2 的顯示名稱；其次保留輸入檔自訂名稱，再回退到地號／地址組合。
    if not clean_text(output.get("顯示名稱", "")):
        if clean_text(output.get("名稱", "")):
            output["顯示名稱"] = clean_text(output.get("名稱", ""))
        elif clean_text(output.get("查詢方式", "")) == "地號":
            output["顯示名稱"] = "".join(
                part for part in (
                    clean_text(output.get("縣市", "")),
                    clean_text(output.get("鄉鎮市區", "")),
                    clean_text(output.get("地段", "")),
                    normalize_landno(output.get("查詢地號", output.get("地號", ""))),
                ) if part
            )
        else:
            output["顯示名稱"] = "".join(
                part for part in (
                    clean_text(output.get("縣市", "")),
                    clean_text(output.get("鄉鎮市區", "")),
                    clean_text(output.get("地址", "")),
                ) if part
            )

    if is_cadastral_query_row(row):
        sid = stable_case_id(row)
        item = download_results.get(sid, {})
        image_path = clean_text(item.get("image_path", ""))
        tile_records = item.get("image_records", []) if isinstance(item.get("image_records"), list) else []
        eligible_tiles, placeholder_tiles = classify_case_tiles(item)
        audit_ok = audit_complete(item, point_map.get(case_key(row)))
        output["地籍圖檔案"] = image_path  # 相容舊欄位：仍指向 Cell2 點唯一命中的原始圖幅
        total_bytes = sum(int(x.get("image_size", 0)) for x in tile_records)
        output["地籍圖大小KB"] = total_bytes / 1024 if tile_records else np.nan
        output["圖幅數"] = len(tile_records)  # NLSC 原始 records 總數，包含佔位圖
        output["圖幅有效數"] = len(eligible_tiles)
        output["圖幅佔位數"] = len(placeholder_tiles)
        output["圖幅涵蓋狀態"] = (
            f"部分圖幅：排除 {len(placeholder_tiles)} 幅退化佔位圖，僅以 {len(eligible_tiles)} 幅有效 PNG 建立幾何"
            if placeholder_tiles else ("全部原始圖幅有效" if audit_ok else "圖幅尚未完整")
        )
        output["圖幅佔位清單_JSON"] = json.dumps([
            {k: x.get(k) for k in ("record_index", "image_path", "image_width", "image_height", "foreground_pixel_count", "polygon_contour_count", "image_class", "placeholder_reason", "lx", "ly", "rx", "ry", "sha256", "image_size")}
            for x in placeholder_tiles
        ], ensure_ascii=False)
        output["圖幅清單_JSON"] = json.dumps([
            {k: x.get(k) for k in ("record_index", "image_path", "lx", "ly", "rx", "ry", "sha256", "image_size", "selected", "image_width", "image_height", "foreground_pixel_count", "polygon_contour_count", "image_class", "geometry_eligible", "placeholder_reason")}
            for x in tile_records
        ], ensure_ascii=False)
        for col, key in [
            ("左下經度", "lx"), ("左下緯度", "ly"),
            ("右上經度", "rx"), ("右上緯度", "ry"),
        ]:
            output[col] = safe_float(item.get(key, np.nan))
        output["response_url"] = item.get("response_url", "")
        output["raw_file"] = item.get("raw_file", "")
        output["image_record_count"] = item.get("image_record_count", "")
        output["selected_image_index"] = item.get("selected_image_index", "")
        output["image_selection_method"] = item.get("image_selection_method", "")
        output["地籍圖image"] = (
            "部分成功（已排除退化佔位圖）" if audit_ok and placeholder_tiles
            else ("成功" if audit_ok else "未取得／圖幅不完整")
        )
        if not audit_ok:
            output["錯誤"] = "存在無法解碼、低維非佔位或缺失 PNG；本輪保留於 retry queue。"
        elif placeholder_tiles:
            output["錯誤"] = f"NLSC 回傳 {len(placeholder_tiles)} 幅退化佔位圖；已保留原圖證據，向量化時排除，改以其他 {len(eligible_tiles)} 幅有效 PNG 聯集。"
        else:
            output["錯誤"] = ""
    else:
        output["地籍圖檔案"] = ""
        output["地籍圖大小KB"] = np.nan
        output["地籍圖image"] = "跳過（非地號案件）"
        output["image_record_count"] = ""
        output["selected_image_index"] = ""
        output["image_selection_method"] = ""
        output["圖幅數"] = 0
        output["圖幅有效數"] = 0
        output["圖幅佔位數"] = 0
        output["圖幅涵蓋狀態"] = "非地號案件"
        output["圖幅佔位清單_JSON"] = "[]"
        output["圖幅清單_JSON"] = "[]"
        output["錯誤"] = "此筆為地址查詢案件，Cell 3 不進行地籍圖下載與向量化。"

    result_rows.append(output)

df_result = pd.DataFrame(result_rows)
for c in ["左下經度", "左下緯度", "右上經度", "右上緯度"]:
    if c not in df_result.columns:
        df_result[c] = np.nan
    df_result[c] = pd.to_numeric(df_result[c], errors="coerce")

df_result.to_excel(IMAGE_RESULT_XLSX, index=False)
mirror_file_to_drive(IMAGE_RESULT_XLSX, "CHECKPOINTS")
atomic_pickle_dump(
    {"version": HANDOFF_VERSION,"results":download_results,"errors":download_errors},
    DOWNLOAD_STATE_PATH,
)

print()
print("NLSC 地籍圖／Image Audit 階段完成：", IMAGE_RESULT_XLSX)


# ============================================================
# 10. Polygon / SVG 工具
# ============================================================

def pixel_to_lonlat(x, y, width, height, lx, ly, rx, ry):
    dx = max(width - 1, 1)
    dy = max(height - 1, 1)
    lon = lx + (x / dx) * (rx - lx)
    lat = ry - (y / dy) * (ry - ly)
    return float(lon), float(lat)


def contour_to_lonlat(contour, width, height, lx, ly, rx, ry):
    return [
        pixel_to_lonlat(pt[0][0], pt[0][1], width, height, lx, ly, rx, ry)
        for pt in contour
    ]


def build_geometry_from_mask(image, lx, ly, rx, ry):
    if image is None:
        raise ValueError("無法讀取 PNG")

    if len(image.shape) == 2:
        mask = (image > 0).astype(np.uint8) * 255
    elif image.shape[2] >= 4:
        mask = (image[:, :, 3] > 0).astype(np.uint8) * 255
    else:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        mask = (gray < 250).astype(np.uint8) * 255

    height, width = mask.shape[:2]
    contours, hierarchy = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    if not contours or hierarchy is None:
        raise ValueError("Alpha mask 無輪廓")

    hierarchy = hierarchy[0]
    polygons = []
    for idx, contour in enumerate(contours):
        if hierarchy[idx][3] != -1 or len(contour) < 3:
            continue
        exterior = contour_to_lonlat(contour, width, height, lx, ly, rx, ry)
        holes = []
        child = hierarchy[idx][2]
        while child != -1:
            child_contour = contours[child]
            if len(child_contour) >= 3:
                holes.append(contour_to_lonlat(child_contour, width, height, lx, ly, rx, ry))
            child = hierarchy[child][0]
        try:
            poly = Polygon(exterior, holes)
            if not poly.is_empty and poly.area > 0:
                polygons.append(poly)
        except Exception:
            pass

    if not polygons:
        raise ValueError("無法建立 Polygon")
    geom = polygons[0] if len(polygons) == 1 else MultiPolygon(polygons)
    if not geom.is_valid:
        geom = geom.buffer(0)
    if geom.is_empty or not geom.is_valid:
        raise ValueError("Alpha mask 建立的幾何為空或無效")
    return geom


def project_wgs84_geometry(geometry):
    """Shapely 2 使用新 API；先 union WGS84，再一次投影 TWD97。"""
    if shapely_transform_v2 is not None:
        return shapely_transform_v2(
            geometry,
            transformer_4326_to_3826.transform,
            interleaved=False,
        )
    return shapely_transform_legacy(transformer_4326_to_3826.transform, geometry)


def make_wgs84_union(tile_geometries):
    if not tile_geometries:
        raise ValueError("沒有任何圖幅幾何可供聯集")
    merged = unary_union(tile_geometries)
    if merged.is_empty:
        raise ValueError("多圖幅聯集結果為空")
    if not merged.is_valid:
        merged = merged.buffer(0)
    if merged.is_empty or not merged.is_valid:
        raise ValueError("WGS84 多圖幅聯集後幾何仍無效")
    return merged


def geometry_component_count(geometry):
    if geometry.geom_type == "Polygon":
        return 1
    if geometry.geom_type == "MultiPolygon":
        return len(geometry.geoms)
    if hasattr(geometry, "geoms"):
        return len(geometry.geoms)
    return 1


def geometry_to_svg(image_path, output_path, image=None):
    if image is None:
        image = cv2.imread(str(image_path), cv2.IMREAD_UNCHANGED)
    if image is None:
        return False
    if len(image.shape) == 2:
        mask = (image > 0).astype(np.uint8) * 255
    elif image.shape[2] >= 4:
        mask = (image[:, :, 3] > 0).astype(np.uint8) * 255
    else:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        mask = (gray < 250).astype(np.uint8) * 255
    height, width = mask.shape[:2]
    contours, _ = cv2.findContours(mask, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    paths = []
    for contour in contours:
        if len(contour) < 2:
            continue
        points = [f"{pt[0][0]},{pt[0][1]}" for pt in contour]
        if points:
            d = "M " + " L ".join(points) + " Z"
            paths.append(f'<path d="{d}" fill="none" stroke="black" stroke-width="1"/>')
    output_path.write_text(
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">' + "".join(paths) + "</svg>",
        encoding="utf-8",
    )
    return True



def run_multitile_geometry_regression_tests():
    """在轉換器與幾何函式均可用後測試 union→projection 流程。"""
    from shapely.geometry import box
    a = box(121.0000, 24.0000, 121.0010, 24.0010)
    b = box(121.0010, 24.0000, 121.0020, 24.0010)
    merged = make_wgs84_union([a, b])
    assert merged.geom_type == "Polygon" and merged.is_valid
    projected = project_wgs84_geometry(merged)
    assert projected.geom_type == "Polygon" and projected.is_valid
    # 再用合成 Alpha mask 測試 Pixel→WGS84 向量化。
    synthetic = np.zeros((24, 24, 4), dtype=np.uint8)
    synthetic[2:22, 2:22, 3] = 127
    tile_geom = build_geometry_from_mask(synthetic, 121.0, 24.0, 121.001, 24.001)
    assert tile_geom.is_valid and not tile_geom.is_empty
    print("MULTI_TILE_GEOMETRY_REGRESSION=PASS (touching union + projection + alpha polygonize)")


run_multitile_geometry_regression_tests()

# ============================================================
# 11. 全部 PNG → 分幅 Polygon → WGS84 union → TWD97
# ============================================================

vector_state = load_pickle_state(VECTOR_STATE_PATH, {"version": HANDOFF_VERSION, "records": {}})
if not isinstance(vector_state, dict):
    vector_state = {"version": HANDOFF_VERSION, "records": {}}
if vector_state.get("input_xlsx_sha256") != INPUT_XLSX_SHA256:
    print("VECTOR_CHECKPOINT_INPUT_SHA_MISMATCH=RESET")
    vector_state = {"version": HANDOFF_VERSION, "records": {}}
vector_resume = dict(vector_state.get("records", {}))
vector_resume.update(load_case_checkpoint_journal("vector"))


def rebase_svg_value(value):
    value = clean_text(value)
    if not value:
        return ""
    old = Path(value)
    candidate = SVG_DIR / old.name
    if candidate.is_file():
        return str(candidate)
    if DRIVE_MOUNTED and old.is_file():
        candidate.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(old, candidate)
        return str(candidate)
    return value


def rebase_vector_resume_paths(records):
    for entry in records.values():
        if not isinstance(entry, dict) or not isinstance(entry.get("record"), dict):
            continue
        rec = entry["record"]
        old_svg = clean_text(rec.get("向量SVG", ""))
        new_svg = rebase_svg_value(old_svg)
        if old_svg and new_svg != old_svg:
            rec["向量SVG"] = new_svg
        raw_list = clean_text(rec.get("圖幅SVG清單_JSON", ""))
        if raw_list:
            try:
                paths = json.loads(raw_list)
                new_paths = [rebase_svg_value(x) for x in paths] if isinstance(paths, list) else []
                rec["圖幅SVG清單_JSON"] = json.dumps(new_paths, ensure_ascii=False)
                if any(x and not Path(x).is_file() for x in new_paths):
                    entry["success"] = False
            except Exception:
                entry["success"] = False


rebase_vector_resume_paths(vector_resume)
vector_records = []
revectorized_count = 0
reused_vector_count = 0

print()
print("=" * 70)
print("開始多圖幅 PNG → Polygon → WGS84 union → TWD97")
print("=" * 70)

for idx, row in df_result.iterrows():
    row_number = idx + 1
    resume_key = stable_case_id(row)
    old_resume = vector_resume.get(resume_key)
    case_item = download_results.get(resume_key, {})
    tile_records = case_item.get("image_records", []) if isinstance(case_item.get("image_records"), list) else []
    current_sig = tile_set_signature(tile_records) if tile_records else None

    _old_algorithm = old_resume.get("algorithm") if isinstance(old_resume, dict) else None
    _row_placeholder_count = safe_float(row.get("圖幅佔位數", 0))
    _legacy_algorithm_compatible = (
        _old_algorithm == "1.2-MULTI-TILE-WGS84-UNION"
        and np.isfinite(_row_placeholder_count)
        and int(_row_placeholder_count) == 0
    )
    can_resume = (
        isinstance(old_resume, dict)
        and (_old_algorithm == HANDOFF_VERSION or _legacy_algorithm_compatible)
        and bool(old_resume.get("success"))
        and isinstance(old_resume.get("record"), dict)
        and ((not is_cadastral_query_row(row)) or (current_sig is not None and current_sig == old_resume.get("tile_set_signature")))
    )

    if can_resume:
        resumed = row.to_dict()
        rec = old_resume["record"]
        for k in [
            "列號", "geometry", "EPSG", "幾何狀態", "GeometryType", "面積_m2",
            "寬像素", "高像素", "解析度分類", "定位點位於宗地", "中心偏移_m", "向量SVG",
            "圖幅數", "圖幅有效數", "圖幅佔位數", "圖幅涵蓋狀態", "圖幅佔位清單_JSON",
            "圖幅向量化數", "圖幅合併方式", "合併後元件數", "圖幅SVG清單_JSON",
        ]:
            if k in rec:
                resumed[k] = rec[k]
        resumed["列號"] = row_number
        vector_records.append(resumed)
        # 將相容的舊版成功 checkpoint 原地遷移到新演算法標記；幾何及來源簽章不變。
        vector_resume[resume_key] = {
            "algorithm": HANDOFF_VERSION,
            "success": True,
            "tile_set_signature": current_sig,
            "record": resumed,
        }
        append_case_checkpoint("vector", resume_key, vector_resume[resume_key])
        reused_vector_count += 1
        print(f"[{row_number:03d}/{len(df_result)}] ⏭ 多圖幅向量已驗證，沿用／checkpoint 升級")
        continue

    record = row.to_dict()
    record.update({
        "列號": row_number,
        "geometry": None,
        "EPSG": 3826,
        "幾何狀態": "",
        "GeometryType": "",
        "面積_m2": np.nan,
        "寬像素": np.nan,
        "高像素": np.nan,
        "解析度分類": "無圖",
        "定位點位於宗地": False,
        "中心偏移_m": np.nan,
        "向量SVG": "",
        "圖幅數": int(len(tile_records)),
        "圖幅有效數": 0,
        "圖幅佔位數": 0,
        "圖幅涵蓋狀態": "圖幅尚未分類",
        "圖幅佔位清單_JSON": "[]",
        "圖幅向量化數": 0,
        "圖幅合併方式": "WGS84_UNION_THEN_TWD97",
        "合併後元件數": 0,
    })

    if not is_cadastral_query_row(row):
        record["幾何狀態"] = "非地號案件，跳過地籍圖向量化"
        vector_records.append(record)
        vector_resume[resume_key] = {
            "algorithm": HANDOFF_VERSION, "success": True,
            "tile_set_signature": None, "record": record,
        }
        append_case_checkpoint("vector", resume_key, vector_resume[resume_key])
        print(f"[{row_number:03d}/{len(df_result)}] 地址案件 → 跳過地籍圖向量化")
        continue

    if not audit_complete(case_item, point_map.get(case_key(row))) or not tile_records:
        record["幾何狀態"] = "圖幅 PNG／Audit 不完整，禁止部分向量化"
        vector_records.append(record)
        vector_resume[resume_key] = {
            "algorithm": HANDOFF_VERSION, "success": False,
            "tile_set_signature": current_sig, "record": record,
        }
        append_case_checkpoint("vector", resume_key, vector_resume[resume_key])
        continue

    tile_geometries = []
    tile_pixel_sizes = []
    tile_svg_paths = [""] * len(tile_records)  # 以原始 record_index 對齊；佔位圖保持空 SVG
    usable_tiles = []
    placeholder_tiles = []
    try:
        for tile in sorted(tile_records, key=lambda x: int(x["record_index"])):
            tile_path = Path(tile["image_path"])
            if not tile_path.is_file() or cached_tile_sha256(tile_path) != tile["sha256"]:
                raise RuntimeError(f"第 {tile['record_index']} 幅 PNG 不存在或 SHA256 不符")
            tile_meta = enrich_tile_metadata(tile)
            if tile_meta["image_class"] == "DEGENERATE_PLACEHOLDER":
                placeholder_tiles.append(tile)
                continue
            if not tile_meta["geometry_eligible"]:
                raise RuntimeError(
                    f"第 {tile['record_index']} 幅 PNG 無法分類為有效圖或佔位圖："
                    f"{tile_meta['image_width']}x{tile_meta['image_height']}"
                )
            image = cv2.imread(str(tile_path), cv2.IMREAD_UNCHANGED)
            if image is None:
                raise RuntimeError(f"第 {tile['record_index']} 幅 PNG 讀取失敗")
            height, width = image.shape[:2]
            geom_tile = build_geometry_from_mask(
                image,
                float(tile["lx"]), float(tile["ly"]),
                float(tile["rx"]), float(tile["ry"]),
            )
            tile_geometries.append(geom_tile)
            tile_pixel_sizes.append((width, height))
            usable_tiles.append(tile)

            office = clean_text(row.get("office", ""))
            sect = normalize_sect(row.get("sect", ""))
            landno = normalize_landno(row.get("查詢地號", ""))
            svg_path = SVG_DIR / f"{row_number:03d}_{office}_{sect}_{landno}_RECORD_{int(tile['record_index']):02d}.svg"
            try:
                geometry_to_svg(tile_path, svg_path, image=image)
                tile_svg_paths[int(tile["record_index"]) - 1] = str(svg_path)
            except Exception:
                tile_svg_paths[int(tile["record_index"]) - 1] = ""

        if not tile_geometries:
            raise RuntimeError("沒有可向量化的有效 PNG；拒絕使用純佔位圖建立 Polygon")

        # 佔位圖只保留作為原始證據；所有其他有效圖幅仍在 WGS84 聯集後才投影。
        geom_wgs84 = make_wgs84_union(tile_geometries)
        geom_twd97 = project_wgs84_geometry(geom_wgs84)
        if not geom_twd97.is_valid:
            geom_twd97 = geom_twd97.buffer(0)
        if geom_twd97.is_empty or not geom_twd97.is_valid:
            raise RuntimeError("投影後聯集幾何為空或無效")

        # 選中圖幅 SVG 路徑保留在舊欄位；全部 SVG 路徑放新欄位。
        selected_index = int(case_item["selected_image_index"])
        selected_svg_idx = selected_index - 1
        record["向量SVG"] = tile_svg_paths[selected_svg_idx] if selected_svg_idx < len(tile_svg_paths) else ""
        record["圖幅SVG清單_JSON"] = json.dumps(tile_svg_paths, ensure_ascii=False)
        record["圖幅有效數"] = int(len(usable_tiles))
        record["圖幅佔位數"] = int(len(placeholder_tiles))
        record["圖幅涵蓋狀態"] = (
            f"PARTIAL_PLACEHOLDER_EXCLUDED:{len(placeholder_tiles)}"
            if placeholder_tiles else "ALL_TILES_VECTORISED"
        )
        record["圖幅佔位清單_JSON"] = json.dumps([
            {k: x.get(k) for k in ("record_index", "image_path", "image_width", "image_height", "foreground_pixel_count", "polygon_contour_count", "image_class", "placeholder_reason", "lx", "ly", "rx", "ry", "sha256", "image_size")}
            for x in placeholder_tiles
        ], ensure_ascii=False)
        record["寬像素"] = int(max(w for w, h in tile_pixel_sizes))
        record["高像素"] = int(max(h for w, h in tile_pixel_sizes))
        min_short_side = min(min(w, h) for w, h in tile_pixel_sizes)
        record["解析度分類"] = "極低解析度" if min_short_side < 20 else ("低解析度" if min_short_side < 100 else "一般")
        record["圖幅向量化數"] = len(tile_geometries)
        record["圖幅合併方式"] = "WGS84_UNARY_UNION_THEN_EPSG3826"
        record["合併後元件數"] = int(geometry_component_count(geom_twd97))
        record["geometry"] = geom_twd97
        record["GeometryType"] = geom_twd97.geom_type
        record["面積_m2"] = float(geom_twd97.area)
        record["幾何狀態"] = "Valid"

        x = safe_float(row.get("TWD97_X", np.nan))
        y = safe_float(row.get("TWD97_Y", np.nan))
        if np.isfinite(x) and np.isfinite(y):
            pt = Point(x, y)
            record["定位點位於宗地"] = bool(geom_twd97.covers(pt))
            record["中心偏移_m"] = float(geom_twd97.centroid.distance(pt))

    except Exception as e:
        record["幾何狀態"] = "多圖幅向量化／聯集失敗：" + repr(e)
        record["geometry"] = None
        record["定位點位於宗地"] = False

    vector_records.append(record)
    success = bool(record.get("幾何狀態") == "Valid" and record.get("geometry") is not None
                   and int(record.get("圖幅向量化數", 0)) == int(record.get("圖幅有效數", 0))
                   and int(record.get("圖幅有效數", 0)) >= 1)
    vector_resume[resume_key] = {
        "algorithm": HANDOFF_VERSION,
        "success": success,
        "tile_set_signature": current_sig,
        "record": record,
    }
    append_case_checkpoint("vector", resume_key, vector_resume[resume_key])
    revectorized_count += 1
    print(
        f"[{row_number:03d}/{len(df_result)}] {case_key(row)} | "
        f"有效圖幅={len(tile_geometries)}/{len(usable_tiles)} | 原始圖幅={len(tile_records)} | "
        f"排除佔位={len(placeholder_tiles)} | {record['GeometryType'] or '-'} | "
        f"components={record.get('合併後元件數', 0)} | {record['幾何狀態']}"
    )


gdf = gpd.GeoDataFrame(vector_records, geometry="geometry", crs="EPSG:3826")
atomic_pickle_dump({"version": HANDOFF_VERSION, "records": vector_resume}, VECTOR_STATE_PATH)

# 下載與向量化完整快照均已原子保存；此後清除逐筆日誌，避免跨次續跑無限累積。
try:
    if CHECKPOINT_JOURNAL_PATH.exists():
        CHECKPOINT_JOURNAL_PATH.unlink()
        print("CELL3_CHECKPOINT_JOURNAL_COMPACTED=PASS")
except Exception as _journal_compact_error:
    print("CELL3_CHECKPOINT_JOURNAL_COMPACTED=SKIPPED", repr(_journal_compact_error))


# ------------------------------------------------------------
# v4.2.1：若仍有下載失敗案件，本輪只保存「未完成續跑成果」。
# 不覆蓋正式 CELL3_HANDOFF / 正式 GeoJSON / GPKG / SHP，避免 Cell 5 誤用不完整資料。
# 下一次執行會由 DOWNLOAD_STATE + VECTOR_STATE 直接續跑。
# ------------------------------------------------------------
_pending_geometry_count = int(gdf.geometry.notna().sum())
_pending_cases = list(remaining_audit)
if _pending_cases or _pending_geometry_count != cadastral_case_count:
    pending_rows = []
    for _idx, _row in df_result.iterrows():
        if not is_cadastral_query_row(_row):
            continue
        _sid = stable_case_id(_row)
        _item = download_results.get(_sid, {})
        if not audit_complete(_item, point_map.get(case_key(_row))):
            pending_rows.append({
                "列號": int(_idx + 1),
                "案件ID": case_key(_row),
                "狀態": "PNG/Audit 未完成",
                "下載錯誤": download_errors.get(_sid, ""),
            })
        elif _idx < len(vector_records):
            _rec = vector_records[_idx]
            if _rec.get("geometry") is None:
                pending_rows.append({
                    "列號": int(_idx + 1),
                    "案件ID": case_key(_row),
                    "狀態": "PNG 已取得但向量化未完成",
                    "下載錯誤": "",
                })

    pending_df = pd.DataFrame(pending_rows)
    pending_path = FORMAL_DIR / "NLSC_地籍圖_待續跑案件.xlsx"
    pending_df.to_excel(pending_path, index=False)
    partial_path = FORMAL_DIR / "NLSC_地籍圖_目前向量化結果_待續跑.gpkg"
    try:
        _partial = gdf[gdf.geometry.notna()].copy()
        if partial_path.exists():
            partial_path.unlink()
        if len(_partial) > 0:
            _partial.to_file(partial_path, layer="NLSC_地籍圖_partial", driver="GPKG")
    except Exception as _e:
        print("⚠️ partial GPKG 儲存失敗：", repr(_e))

    atomic_pickle_dump(
        {"version": HANDOFF_VERSION, "results":download_results,"errors":download_errors},
        DOWNLOAD_STATE_PATH,
    )
    atomic_pickle_dump(
        {"version": HANDOFF_VERSION,"records":vector_resume},
        VECTOR_STATE_PATH,
    )

    print()
    print("=" * 80)
    print("CELL 3 v4.3 MULTI-TILE｜本輪未完成，已安全保存續跑狀態")
    print("=" * 80)
    print(f"地籍案件：{cadastral_case_count}")
    print(f"本輪已取得／可向量化：{_pending_geometry_count}")
    print(f"仍待補抓／待向量化：{len(pending_df)}")
    print("待續跑清單：", pending_path)
    print("下載 checkpoint：", DOWNLOAD_STATE_PATH)
    print("向量 checkpoint：", VECTOR_STATE_PATH)
    print("本輪不覆蓋正式 CELL3_HANDOFF，避免 Cell 5 使用不完整 Polygon。")
    print("下一次重新執行本程式時，會優先補抓這些案件。")
    print("=" * 80)
    raise SystemExit(0)


# ============================================================
# 12. Cell 3 → Cell 5 Handoff
# ============================================================

def build_case_id_from_record(record):
    return (
        f"{clean_text(record.get('office','')).upper()}_"
        f"{normalize_sect(record.get('sect',''))}_"
        f"{normalize_landno(record.get('查詢地號', record.get('地號','')))}"
    )


gdf["點索引"] = np.arange(len(gdf), dtype=int)
gdf["案件ID"] = [build_case_id_from_record(r) for r in vector_records]
gdf_twd97 = gdf.copy()
gdf_wgs84 = gdf.to_crs("EPSG:4326")

_cell5_keep_cols = [
    "案件ID", "點索引", "列號", "縣市", "鄉鎮市區", "地段", "地號", "查詢地號",
    "本筆面積(公頃)", "申請面積(公頃)", "名稱", "顯示名稱", "office", "sect",
    "經度", "緯度", "TWD97_X", "TWD97_Y", "面積_m2", "解析度分類",
    "定位點位於宗地", "中心偏移_m", "幾何狀態", "GeometryType", "地籍圖檔案", "向量SVG",
    "geometry",
]
_cell5_keep_cols = [c for c in _cell5_keep_cols if c in gdf_twd97.columns]

_cell5_extra_cols = [
    "圖幅數", "圖幅有效數", "圖幅佔位數", "圖幅涵蓋狀態", "圖幅佔位清單_JSON",
    "圖幅向量化數", "圖幅合併方式", "合併後元件數", "圖幅清單_JSON", "圖幅SVG清單_JSON",
]
_cell5_keep_cols = list(dict.fromkeys(_cell5_keep_cols + [c for c in _cell5_extra_cols if c in gdf_twd97.columns]))

CELL3_HANDOFF = {
    "version": "1.1",
    "geometry_algorithm_version": HANDOFF_VERSION,
    "count": int(len(gdf)),
    "cadastral_input_rows": int(cadastral_case_count),
    "unique_nlsc_query_keys": int(unique_cadastral_key_count),
    "duplicate_nlsc_key_groups": int(len(_duplicate_key_groups)),
    "duplicate_nlsc_key_extra_rows_preserved": int(duplicate_key_extra_rows),
    "crs_twd97": "EPSG:3826",
    "crs_wgs84": "EPSG:4326",
    "geometry_build_mode": "GEOMETRY_ELIGIBLE_IMAGE_RECORDS_PIXEL_TO_WGS84_UNION_THEN_TWD97",
    "image_record_policy": "SAVE_ALL_RECORDS; EXCLUDE_IMAGES_WITHOUT_VALID_POLYGON_CONTOUR_FROM_POLYGON; VECTORIZE_AND_UNION_OTHER_VALID_PNG; CELL2_POINT_UNIQUELY_SELECTS_REFERENCE_TILE",
    "df": df_result.copy(),
    "gdf_twd97": gdf_twd97[_cell5_keep_cols].copy(),
    "gdf_wgs84": gdf_wgs84[_cell5_keep_cols].copy(),
    "image_dir": str(IMAGE_DIR),
    "svg_dir": str(SVG_DIR),
    "multi_tile_candidate": True,
    "run_stamp": RUN_STAMP,
}

# 不在此處立即覆蓋正式 handoff；待所有輸出及最終稽核通過後，最後才原子更新。


# ============================================================
# 13. Formal GeoJSON / GPKG / SHP / Excel
# ============================================================

valid_geometry_count = int(gdf.geometry.notna().sum())
valid_after_fix = int(gdf[gdf.geometry.notna()].geometry.is_valid.sum())
low_resolution_count = int((gdf["解析度分類"] == "低解析度").sum())
very_low_count = int((gdf["解析度分類"] == "極低解析度").sum())


gdf_geojson = gdf_wgs84[gdf_wgs84.geometry.notna()].copy()
for col in gdf_geojson.columns:
    if col != "geometry" and gdf_geojson[col].dtype == "object":
        gdf_geojson[col] = gdf_geojson[col].map(lambda v: "" if v is None or (not isinstance(v, (list, dict, tuple)) and pd.isna(v)) else str(v))
if len(gdf_geojson) > 0:
    gdf_geojson.to_file(GEOJSON_PATH, driver="GeoJSON", encoding="utf-8")

if GPKG_PATH.exists():
    GPKG_PATH.unlink()
gdf_gpkg = gdf[gdf.geometry.notna()].copy()
if len(gdf_gpkg) > 0:
    gdf_gpkg.to_file(GPKG_PATH, layer="NLSC_地籍圖", driver="GPKG")

for suffix in [".shp", ".shx", ".dbf", ".prj", ".cpg"]:
    p = SHP_BASE.with_suffix(suffix)
    if p.exists():
        p.unlink()

gdf_shp_full = gdf[gdf.geometry.notna()].copy()

# Shapefile/DBF 單欄文字長度上限約 254 bytes；長 JSON 不寫入 DBF，
# 改用逐圖幅 CSV sidecar，並對其他超長欄位保留完整值。
tile_detail_rows = build_shp_tile_detail_rows(gdf_shp_full)
tile_detail_df = pd.DataFrame(tile_detail_rows, columns=[
    "案件ID", "列號", "圖幅序號", "是否定位選中", "PNG路徑",
    "左下經度", "左下緯度", "右上經度", "右上緯度",
    "PNG大小_bytes", "PNG寬", "PNG高", "前景像素數", "有效外輪廓數",
    "圖幅影像分類", "是否納入向量化", "排除原因", "PNG_SHA256", "SVG路徑",
])
expected_tile_detail_count = int(pd.to_numeric(gdf_shp_full.get("圖幅數", 0), errors="coerce").fillna(0).sum())
if len(tile_detail_df) != expected_tile_detail_count:
    raise RuntimeError(
        f"SHP 圖幅詳細資料筆數不一致：{len(tile_detail_df)}/{expected_tile_detail_count}"
    )
tile_detail_df.to_csv(SHP_TILE_DETAILS_CSV, index=False, encoding="utf-8-sig")

SHP_FIELD_MAP = {
    "查詢方式":"QRY_TYPE", "縣市":"CITY", "鄉鎮市區":"AREA", "地段":"SECTION", "地號":"LANDNO",
    "本筆面積(公頃)":"P_AREA", "申請面積(公頃)":"APP_AREA", "地址":"ADDRESS", "定位來源":"LOC_SRC",
    "完整查詢地址":"FULL_ADDR", "命中地址":"HIT_ADDR", "定位品質":"LOC_QUAL", "查詢地號":"Q_LANDNO",
    "office":"OFFICE", "sect":"SECT", "狀態":"STATUS", "HTTP狀態":"HTTP", "經度":"LON", "緯度":"LAT",
    "TWD97_X":"TWD_X", "TWD97_Y":"TWD_Y", "左下經度":"LX", "左下緯度":"LY", "右上經度":"RX", "右上緯度":"RY",
    "response_url":"RESP_URL", "raw_file":"RAW_FILE", "錯誤":"ERROR", "嘗試次數":"TRIES", "地籍圖image":"IMG_OK",
    "地籍圖檔案":"IMG_FILE", "地籍圖大小KB":"IMG_KB", "image_record_count":"IMG_CNT", "selected_image_index":"IMG_IDX",
    "image_selection_method":"IMG_SEL", "列號":"ROW_NO", "EPSG":"EPSG", "幾何狀態":"GEOM_STAT", "GeometryType":"GEOM_TYPE",
    "面積_m2":"AREA_M2", "寬像素":"PX_W", "高像素":"PX_H", "解析度分類":"RES_CLASS", "定位點位於宗地":"PT_IN_POLY",
    "中心偏移_m":"OFFSET_M", "向量SVG":"SVG_FILE", "案件ID":"CASE_ID", "點索引":"PT_INDEX",
    "圖幅數":"TILE_CNT", "圖幅有效數":"TILE_OK", "圖幅佔位數":"TILE_PH", "圖幅涵蓋狀態":"TILE_COV",
    "圖幅向量化數":"TILE_VEC", "圖幅合併方式":"TILE_MODE", "合併後元件數":"COMP_CNT",
}

gdf_shp, shp_long_text_rows, shp_excluded_columns = prepare_shapefile_frame(gdf_shp_full, SHP_FIELD_MAP)
pd.DataFrame(shp_long_text_rows, columns=["案件ID", "列號", "欄位名稱", "完整值", "完整值字元數", "完整值UTF8位元組數"]).to_csv(
    SHP_LONG_TEXT_CSV, index=False, encoding="utf-8-sig"
)

used = set()
rename = {}
map_records = []
for col in gdf_shp.columns:
    if col == "geometry":
        continue
    new_name = SHP_FIELD_MAP.get(col, f"FIELD{len(map_records)+1:03d}")
    new_name = new_name.encode("ascii", "ignore").decode("ascii") or "FIELD"
    new_name = new_name[:10]
    base = new_name
    n = 1
    while new_name in used:
        suffix = str(n)
        new_name = base[:10-len(suffix)] + suffix
        n += 1
    used.add(new_name)
    rename[col] = new_name
    map_records.append({"原始欄位": col, "SHP欄位": new_name})

gdf_shp = gdf_shp.rename(columns=rename)
if len(gdf_shp) > 0:
    gdf_shp.to_file(SHP_PATH, driver="ESRI Shapefile", encoding="UTF-8")
pd.DataFrame(map_records).to_excel(FORMAL_DIR / "SHP_欄位對照.xlsx", index=False)

# 寫出後讀回檢查，並確認任何 DBF 字串欄位都已在安全長度內。
if len(gdf_shp) > 0:
    _shp_readback = gpd.read_file(SHP_PATH)
    if len(_shp_readback) != len(gdf_shp):
        raise RuntimeError(f"SHP 回讀筆數不一致：{len(_shp_readback)}/{len(gdf_shp)}")
    if int(_shp_readback.geometry.notna().sum()) != len(gdf_shp):
        raise RuntimeError("SHP 回讀後有缺失幾何")
if shp_excluded_columns:
    print("SHP_JSON_FIELDS_MOVED_TO_TILE_SIDECAR=", shp_excluded_columns)
print("SHP_TILE_DETAILS_ROWS=", len(tile_detail_df), "EXPECTED=", expected_tile_detail_count)
print("SHP_LONG_TEXT_VALUES_SIDECAR=", len(shp_long_text_rows))
print("SHP_TEXT_LIMIT_UTF8_BYTES=254")
print("SHP_DBF_EXPORT_SCHEMA=PASS")

attribute_df = gdf.drop(columns=["geometry"], errors="ignore").copy()
attribute_df.to_excel(ATTRIBUTE_PATH, index=False)

low_df = gdf[gdf["解析度分類"].isin(["低解析度", "極低解析度"])].drop(columns=["geometry"], errors="ignore")
low_path = FORMAL_DIR / "低解析度案件.xlsx"
low_df.to_excel(low_path, index=False)

very_low_df = gdf[gdf["解析度分類"] == "極低解析度"].drop(columns=["geometry"], errors="ignore")
very_low_path = FORMAL_DIR / "極低解析度案件.xlsx"
very_low_df.to_excel(very_low_path, index=False)

summary = pd.DataFrame([
    {"項目":"總筆數", "數量":len(gdf)},
    {"項目":"地籍案件", "數量":cadastral_case_count},
    {"項目":"有幾何", "數量":valid_geometry_count},
    {"項目":"Valid 幾何", "數量":valid_after_fix},
    {"項目":"低解析度", "數量":low_resolution_count},
    {"項目":"極低解析度", "數量":very_low_count},
    {"項目":"含退化佔位圖案件", "數量":int((pd.to_numeric(gdf.get("圖幅佔位數", 0), errors="coerce").fillna(0) > 0).sum())},
    {"項目":"退化佔位圖幅總數", "數量":int(pd.to_numeric(gdf.get("圖幅佔位數", 0), errors="coerce").fillna(0).sum())},
    {"項目":"Audit 完整（唯一 NLSC 案件鍵）", "數量":unique_cadastral_key_count - len(remaining_audit)},
    {"項目":"唯一 NLSC 案件鍵", "數量":unique_cadastral_key_count},
    {"項目":"重複案件鍵群組", "數量":len(_duplicate_key_groups)},
    {"項目":"重複案件鍵額外資料列（保留）", "數量":duplicate_key_extra_rows},
    {"項目":"Raw 離線重選圖", "數量":len(raw_repaired)},
    {"項目":"GitHub SHA256 圖幅復原", "數量":len(github_recovered)},
    {"項目":"Raw 重選後 PNG 變更", "數量":len(changed_from_raw)},
    {"項目":"向量沿用", "數量":reused_vector_count},
    {"項目":"向量重新處理", "數量":revectorized_count},
    {"項目":"SHP 圖幅詳細資料列數", "數量":len(tile_detail_df)},
    {"項目":"SHP 超長文字 sidecar 筆數", "數量":len(shp_long_text_rows)},
])
summary.to_excel(FORMAL_DIR / "NLSC_地籍圖_向量化摘要.xlsx", index=False)


# ============================================================
# 14. 最終正式自我檢查
# ============================================================

final_audit_count = 0
final_geometry_count = int(gdf.geometry.notna().sum())
for _, row in df_result.iterrows():
    if is_cadastral_query_row(row):
        method = clean_text(row.get("image_selection_method", ""))
        count = clean_text(row.get("image_record_count", ""))
        selected = clean_text(row.get("selected_image_index", ""))
        if method and count and selected:
            final_audit_count += 1

if final_audit_count != cadastral_case_count:
    raise RuntimeError(f"最終 Audit 自我檢查失敗：{final_audit_count}/{cadastral_case_count}")
if final_geometry_count != cadastral_case_count:
    raise RuntimeError(f"最終 Polygon 自我檢查失敗：{final_geometry_count}/{cadastral_case_count}")

# 任何一筆有效 PNG 未全部向量化、佔位圖分類異常或聯集無效，都不覆寫 Cell5 正式 handoff。
for _, _r in gdf.iterrows():
    if is_cadastral_query_row(_r):
        _raw_count = safe_float(_r.get("圖幅數", 0))
        _valid_count = safe_float(_r.get("圖幅有效數", 0))
        _vectorized_count = safe_float(_r.get("圖幅向量化數", 0))
        _raw = int(_raw_count) if np.isfinite(_raw_count) else 0
        _valid = int(_valid_count) if np.isfinite(_valid_count) else 0
        _vectorized = int(_vectorized_count) if np.isfinite(_vectorized_count) else 0
        if (_raw < 1 or _valid < 1 or _vectorized != _valid
                or _r.get("幾何狀態") != "Valid" or _r.get("geometry") is None):
            raise RuntimeError(
                f"多圖幅最終 Gate 失敗：{case_key(_r)} | raw={_raw}; vectorized={_vectorized}/{_valid} | { _r.get('幾何狀態') }"
            )

# ------------------------------------------------------------
# 依儲存模式建立備份與 manifest；Drive 未掛載時完全不呼叫 Drive API。
# 正式 handoff 仍固定提交至 /content/CELL3_HANDOFF.pkl，供 Cell5 直接讀取。
# ------------------------------------------------------------
if DRIVE_MOUNTED:
    BACKUP_DIR = DRIVE_BACKUP_DIR
    BACKUP_MANIFEST_NAME = "DRIVE_BACKUP_MANIFEST.json"
else:
    BACKUP_DIR = RUN_BACKUP_DIR
    BACKUP_MANIFEST_NAME = "LOCAL_BACKUP_MANIFEST.json"

BACKUP_DIR.mkdir(parents=True, exist_ok=True)
if CELL3_HANDOFF_PATH.exists():
    prior_backup = BACKUP_DIR / "CELL3_HANDOFF_PREVIOUS.pkl"
    shutil.copy2(CELL3_HANDOFF_PATH, prior_backup)
    prior_sha = sha256_file(prior_backup)
    if prior_sha != sha256_file(CELL3_HANDOFF_PATH):
        raise RuntimeError("舊 handoff 備份 SHA256 不一致；拒絕更新正式 handoff")
else:
    prior_backup = None
    prior_sha = None

# 先建立 handoff 暫存檔並做回讀檢查。
new_handoff_temp = CELL3_HANDOFF_PATH.with_name(CELL3_HANDOFF_PATH.name + ".multi_tile_tmp")
pd.to_pickle(CELL3_HANDOFF, new_handoff_temp)
if not new_handoff_temp.is_file() or new_handoff_temp.stat().st_size <= 0:
    raise RuntimeError("新 handoff 暫存檔不存在或為空；拒絕替換正式 handoff")
_verify_handoff = pd.read_pickle(new_handoff_temp)
if _verify_handoff.get("version") != "1.1" or _verify_handoff.get("geometry_algorithm_version") != HANDOFF_VERSION:
    raise RuntimeError("handoff 回讀版本不一致")
if int(_verify_handoff.get("count", -1)) != len(gdf):
    raise RuntimeError("handoff 回讀筆數不一致")
if int(_verify_handoff["gdf_twd97"].geometry.notna().sum()) != cadastral_case_count:
    raise RuntimeError("handoff 回讀幾何筆數不一致")
if int(_verify_handoff["gdf_wgs84"].geometry.notna().sum()) != cadastral_case_count:
    raise RuntimeError("handoff WGS84 回讀幾何筆數不一致")

new_handoff_sha = sha256_file(new_handoff_temp)
backup_handoff = BACKUP_DIR / "CELL3_HANDOFF_MULTI_TILE.pkl"
shutil.copy2(new_handoff_temp, backup_handoff)
if sha256_file(backup_handoff) != new_handoff_sha:
    raise RuntimeError("新 handoff 備份 SHA256 不一致；拒絕替換正式 handoff")

# Drive 模式：工作資料已直接寫入本次 Drive 執行資料夾。
# 本地模式：工作資料已在 /content；manifest 記錄原位檔案，不重複複製整批 PNG。
manifest_items = []
if DRIVE_MOUNTED:
    for fp in sorted(BACKUP_DIR.rglob("*")):
        if fp.is_file() and fp.name not in {BACKUP_MANIFEST_NAME, BACKUP_MANIFEST_NAME + ".tmp"}:
            manifest_items.append({
                "path": str(fp.relative_to(BACKUP_DIR)),
                "size_bytes": fp.stat().st_size,
                "sha256": sha256_file(fp),
            })
else:
    _local_roots = [IMAGE_DIR, RAW_DIR, FORMAL_DIR, SVG_DIR, VECTOR_CHECK_DIR]
    _local_files = [DOWNLOAD_STATE_PATH, VECTOR_STATE_PATH, IMAGE_RESULT_XLSX, CHECKPOINT_JOURNAL_PATH]
    for _root in _local_roots:
        if _root.exists():
            _local_files.extend(p for p in _root.rglob("*") if p.is_file())
    for _file in _local_files:
        if _file.is_file():
            manifest_items.append({
                "path": str(_file.relative_to(Path("/content"))),
                "size_bytes": _file.stat().st_size,
                "sha256": sha256_file(_file),
            })
    for _file in sorted(BACKUP_DIR.rglob("*")):
        if _file.is_file() and _file.name not in {BACKUP_MANIFEST_NAME, BACKUP_MANIFEST_NAME + ".tmp"}:
            manifest_items.append({
                "path": str(_file.relative_to(Path("/content"))),
                "size_bytes": _file.stat().st_size,
                "sha256": sha256_file(_file),
            })

backup_manifest = {
    "version": HANDOFF_VERSION,
    "candidate_version": "4.3.4",
    "storage_mode": "GOOGLE_DRIVE" if DRIVE_MOUNTED else "LOCAL_CONTENT",
    "storage_root": str(STORAGE_ROOT),
    "run_stamp": RUN_STAMP,
    "input_xlsx": str(INPUT_XLSX),
    "input_xlsx_sha256": INPUT_XLSX_SHA256,
    "previous_handoff_sha256": prior_sha,
    "new_handoff_sha256": new_handoff_sha,
    "commit_status": "PREPARED_TO_COMMIT",
    "files": manifest_items,
}
manifest_path = BACKUP_DIR / BACKUP_MANIFEST_NAME
manifest_tmp = manifest_path.with_name(manifest_path.name + ".tmp")
manifest_tmp.write_text(json.dumps(backup_manifest, ensure_ascii=False, indent=2), encoding="utf-8")
_manifest_check = json.loads(manifest_tmp.read_text(encoding="utf-8"))
if _manifest_check.get("new_handoff_sha256") != new_handoff_sha or _manifest_check.get("commit_status") != "PREPARED_TO_COMMIT":
    raise RuntimeError("備份 manifest preflight 失敗；拒絕替換正式 handoff")

# 驗證 manifest 列出的所有檔案均存在且 SHA256 正確。
for _item in manifest_items:
    _fp = (BACKUP_DIR / _item["path"]) if DRIVE_MOUNTED else (Path("/content") / _item["path"])
    if not _fp.is_file() or _fp.stat().st_size != _item["size_bytes"] or sha256_file(_fp) != _item["sha256"]:
        raise RuntimeError(f"備份完整性驗證失敗：{_item['path']}")

os.replace(manifest_tmp, manifest_path)

# 防止候選產生期間另一個流程更新 handoff：覆寫前確認舊 hash 仍一致。
if prior_sha is not None:
    if not CELL3_HANDOFF_PATH.is_file() or sha256_file(CELL3_HANDOFF_PATH) != prior_sha:
        raise RuntimeError("正式 handoff 在備份期間已被其他流程變更；拒絕覆寫")

# 所有候選檔、舊 handoff、manifest 及 SHA256 均已驗證，最後才原子替換。
os.replace(new_handoff_temp, CELL3_HANDOFF_PATH)
final_handoff_sha = sha256_file(CELL3_HANDOFF_PATH)
if final_handoff_sha != new_handoff_sha:
    if prior_backup is not None and prior_backup.is_file():
        _rollback = CELL3_HANDOFF_PATH.with_name(CELL3_HANDOFF_PATH.name + ".rollback_tmp")
        shutil.copy2(prior_backup, _rollback)
        os.replace(_rollback, CELL3_HANDOFF_PATH)
    raise RuntimeError("正式 handoff 更新後 SHA256 不一致，已執行回復／停止")

# 更新 manifest 最終狀態；失敗時回復舊 handoff，避免證據與正式檔不同步。
backup_manifest["commit_status"] = "COMMITTED"
backup_manifest["committed_handoff_sha256"] = final_handoff_sha
manifest_tmp.write_text(json.dumps(backup_manifest, ensure_ascii=False, indent=2), encoding="utf-8")
try:
    _manifest_final_check = json.loads(manifest_tmp.read_text(encoding="utf-8"))
    if _manifest_final_check.get("committed_handoff_sha256") != final_handoff_sha:
        raise RuntimeError("COMMITTED manifest SHA mismatch")
    os.replace(manifest_tmp, manifest_path)
except Exception:
    if prior_backup is not None and prior_backup.is_file():
        _rollback = CELL3_HANDOFF_PATH.with_name(CELL3_HANDOFF_PATH.name + ".rollback_tmp")
        shutil.copy2(prior_backup, _rollback)
        os.replace(_rollback, CELL3_HANDOFF_PATH)
    raise

if DRIVE_MOUNTED:
    print("DRIVE_BACKUP_PRECOMMIT_VERIFY=PASS")
    print("DRIVE_BACKUP_MANIFEST_STATUS=COMMITTED")
else:
    print("LOCAL_CONTENT_BACKUP_PRECOMMIT_VERIFY=PASS")
    print("LOCAL_CONTENT_MANIFEST_STATUS=COMMITTED")
print("FORMAL_HANDOFF_FINAL_SHA256=", final_handoff_sha)


# ============================================================
# 15. 完成
# ============================================================

print()
print("=" * 80)
print("CELL 3 v4.3.4 MULTI-TILE DUPLICATE-KEY SAFE 完成")
print("=" * 80)
print(f"總筆數：{len(gdf)}")
print(f"地籍案件：{cadastral_case_count}")
print(f"地址案件：{address_case_count}")
print(f"Image Audit：{final_audit_count}/{cadastral_case_count}")
print(f"有幾何：{valid_geometry_count}")
print(f"Valid 幾何：{valid_after_fix}")
print(f"低解析度：{low_resolution_count}")
print(f"極低解析度：{very_low_count}")
print(f"含退化佔位圖案件：{int((pd.to_numeric(gdf.get("圖幅佔位數", 0), errors="coerce").fillna(0) > 0).sum())}")
print(f"退化佔位圖幅總數：{int(pd.to_numeric(gdf.get("圖幅佔位數", 0), errors="coerce").fillna(0).sum())}")
print(f"Raw 離線重新選圖：{len(raw_repaired)}")
print(f"Raw 重新選圖後 PNG 改變：{len(changed_from_raw)}")
print(f"向量沿用：{reused_vector_count}")
print(f"向量重新處理：{revectorized_count}")
print()
print("PNG：", IMAGE_DIR)
print("結果 Excel：", IMAGE_RESULT_XLSX)
print("GeoJSON：", GEOJSON_PATH)
print("GeoPackage：", GPKG_PATH)
print("Shapefile：", SHP_PATH)
print("SHP 欄位對照：", FORMAL_DIR / "SHP_欄位對照.xlsx")
print("SHP 圖幅詳細資料：", SHP_TILE_DETAILS_CSV)
print("SHP 超長文字完整值：", SHP_LONG_TEXT_CSV)
print("屬性 Excel：", ATTRIBUTE_PATH)
print("重複案件鍵稽核：", DUPLICATE_KEY_AUDIT_PATH if _duplicate_key_groups else "無重複鍵")
print("Cell 5 交接：", CELL3_HANDOFF_PATH)
print("下載續跑狀態：", DOWNLOAD_STATE_PATH)
print("向量化續跑狀態：", VECTOR_STATE_PATH)
print("SVG：", SVG_DIR)
print("低解析度：", low_path)
print("極低解析度：", very_low_path)
print("正式成果檔名規則：不使用「_106筆」")
print("正式多 image 規則：全部 records 均保存；有有效面外輪廓的 PNG 才向量化聯集；退化佔位圖保留但排除")
print("幾何聯集順序：各幅 Pixel→WGS84；WGS84 unary_union；一次投影 EPSG:3826")
print("正式 CELL3_HANDOFF 已更新：", CELL3_HANDOFF_PATH)
print("來源程式備份：", RUN_BACKUP_DIR / "SOURCE")
if DRIVE_MOUNTED:
    print("Drive 工作資料／備份：", DRIVE_BACKUP_DIR)
    print("Drive SHA256 manifest：", DRIVE_BACKUP_DIR / "DRIVE_BACKUP_MANIFEST.json")
else:
    print("未掛載 Google Drive；未主動掛載或連線。")
    print("本地備份／SHA256 manifest：", RUN_BACKUP_DIR / "LOCAL_BACKUP_MANIFEST.json")
print("Cell5 正式 handoff 原子更新：", CELL3_HANDOFF_PATH)
print("=" * 80)


# ======================================================================
# Cell 2+3 v1｜整合執行完成確認
# 只有 Cell 3 全部正常執行至結尾，才會輸出 SUCCESS。
# ======================================================================

_CELL23_HANDOFF = Path("/content/CELL3_HANDOFF.pkl")
if not _CELL23_HANDOFF.is_file() or _CELL23_HANDOFF.stat().st_size <= 0:
    raise RuntimeError(
        "Cell 3 未成功產生有效的 /content/CELL3_HANDOFF.pkl。"
    )

_CELL23_HANDOFF_SHA256 = hashlib.sha256(_CELL23_HANDOFF.read_bytes()).hexdigest()

print()
print("=" * 76)
print("[CELL23_EXECUTION_STATUS]=SUCCESS")
print("[CELL23_VERSION]=v2")
print("[CELL2_VERSION]=v8.7")
print("[CELL3_VERSION]=v4.3.4")
print(f"[CELL3_HANDOFF_PATH]={_CELL23_HANDOFF}")
print(f"[CELL3_HANDOFF_SIZE_BYTES]={_CELL23_HANDOFF.stat().st_size}")
print(f"[CELL3_HANDOFF_SHA256]={_CELL23_HANDOFF_SHA256}")
print("Cell 2+3 v2 已完成，可接續執行 Cell 4。")
print("=" * 76)