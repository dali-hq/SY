import json
from io import StringIO
from pathlib import Path

import pandas as pd

# ==================================================
# 基本設定：每次處理不同月份，只需改這裡
# ==================================================

folder = Path(__file__).parent

input_filename = "出入歷史紀錄報表10月.xls"
output_filename = "員工出勤資料_10月.json"

REPORT_YEAR = 2026
REPORT_MONTH = 10

input_file = folder / input_filename
output_file = folder / output_filename

# 老闆帳號：保留打卡資料，但不判定工時、遲到、加班、狀態
EXEMPT_EMPLOYEES = {"10001", "10002"}

# ==================================================
# 門別名稱
# ==================================================

location_names = {
    "10001大門": "實毅總公司",
    "11203大里農會": "台中大里農會",
    "11211九樓": "實毅總公司",
    "11301南七": "南科第七座配水池",
    "11305北港臉型機": "北港水資源中心",
    "11402嘉義配水池": "嘉義園區配水池"
}

location_short_names = {
    "11305北港臉型機": "北港",
    "11203大里農會": "大里",
    "11301南七": "南七",
    "11402嘉義配水池": "嘉配",
    "10001大門": "總公司",
    "11211九樓": "總公司"
}

name_corrections = {
    "33007": "陳廷儹",
    "33036": "黃明山"
}

# ==================================================
# 文字與日期時間工具
# ==================================================

def clean_text(value):
    """清理儲存格文字。"""
    if pd.isna(value):
        return ""

    return (
        str(value)
        .replace("\u3000", "")
        .replace("\xa0", "")
        .replace("\\", "")
        .strip()
    )

def clean_column_name(column_name):
    """清理欄位名稱，移除表頭中的空白與換行。"""
    return (
        str(column_name)
        .replace("\u3000", "")
        .replace("\xa0", "")
        .replace(" ", "")
        .replace("\n", "")
        .replace("\r", "")
        .strip()
    )

def normalize_location_name(location):
    """統一門別名稱格式。"""
    return clean_text(location).replace(" ", "")

def parse_excel_date(value):
    """轉換 Excel 日期為 Timestamp。"""
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value.normalize()

    if isinstance(value, str):
        text = clean_text(value)

        for fmt in [
            "%Y/%m/%d",
            "%Y-%m-%d",
            "%Y/%m/%d %H:%M:%S",
            "%Y-%m-%d %H:%M:%S"
        ]:
            try:
                return pd.to_datetime(text, format=fmt).normalize()
            except Exception:
                continue

        try:
            return pd.to_datetime(text).normalize()
        except Exception:
            return pd.NaT

    try:
        return pd.to_datetime(value).normalize()
    except Exception:
        return pd.NaT

def parse_excel_time(value):
    """轉換 Excel 時間為固定基準日的 Timestamp。"""
    if pd.isna(value):
        return pd.NaT

    base = pd.Timestamp("1900-01-01")

    if isinstance(value, pd.Timestamp):
        return base + pd.Timedelta(
            hours=value.hour,
            minutes=value.minute,
            seconds=value.second
        )

    if hasattr(value, "hour") and hasattr(value, "minute"):
        return base + pd.Timedelta(
            hours=value.hour,
            minutes=value.minute,
            seconds=getattr(value, "second", 0)
        )

    if isinstance(value, str):
        text = clean_text(value)

        for fmt in ["%H:%M:%S", "%H:%M"]:
            try:
                parsed = pd.to_datetime(text, format=fmt)
                return base + pd.Timedelta(
                    hours=parsed.hour,
                    minutes=parsed.minute,
                    seconds=parsed.second
                )
            except Exception:
                continue

        return pd.NaT

    # Excel 有時會把時間讀成 0 到 1 的浮點數
    try:
        number = float(value)

        if 0 <= number < 1:
            return base + pd.to_timedelta(number, unit="D")

        parsed = pd.to_datetime(value)
        return base + pd.Timedelta(
            hours=parsed.hour,
            minutes=parsed.minute,
            seconds=parsed.second
        )
    except Exception:
        return pd.NaT

def combine_locations(values):
    """同一人同一天多門別時，合併為不重複的文字。"""
    unique_values = []

    for value in values:
        text = clean_text(value)

        if text and text not in unique_values:
            unique_values.append(text)

    return "、".join(unique_values)

def get_location_short_name(location_full):
    """取得門別簡稱。"""
    for original_name, short_name in location_short_names.items():
        if original_name in location_full:
            return short_name

    return ""

# ==================================================
# 讀取 Excel / HTML 偽裝 XLS
# ==================================================

def has_required_headers(values):
    """判斷某一列是否是報表欄位列。"""
    clean_values = [clean_column_name(value) for value in values]

    return (
        "員編" in clean_values
        and "日期" in clean_values
        and "時間" in clean_values
    )

def make_dataframe_from_header_row(raw_df, header_index):
    """將指定列設為欄名，後面的列設為資料。"""
    headers = [
        clean_column_name(value)
        for value in raw_df.iloc[header_index].tolist()
    ]

    data = raw_df.iloc[header_index + 1:].copy()
    data.columns = headers

    # 避免完全空白列影響結果
    data = data.dropna(how="all").reset_index(drop=True)

    return data

def read_html_attendance_file(file_path):
    """讀取 HTML 格式但副檔名為 XLS 的報表。"""
    raw_bytes = file_path.read_bytes()

    for encoding in ["cp950", "big5", "utf-8", "latin1"]:
        try:
            content = raw_bytes.decode(encoding, errors="ignore")
            tables = pd.read_html(StringIO(content), header=None)

            for table in tables:
                for index in range(min(30, len(table))):
                    if has_required_headers(table.iloc[index].tolist()):
                        return make_dataframe_from_header_row(table, index)
        except Exception:
            continue

    raise ValueError(
        "檔案可能是 HTML 格式，但找不到包含「員編、日期、時間」的欄位列。"
    )

def read_excel_attendance_file(file_path):
    """讀取原生 XLS / XLSX，並自動找出欄位名稱列。"""
    last_error = None

    for engine in ["xlrd", None]:
        try:
            if engine is None:
                raw_df = pd.read_excel(file_path, header=None)
            else:
                raw_df = pd.read_excel(
                    file_path,
                    header=None,
                    engine=engine
                )

            for index in range(min(30, len(raw_df))):
                if has_required_headers(raw_df.iloc[index].tolist()):
                    return make_dataframe_from_header_row(raw_df, index)

        except Exception as error:
            last_error = error

    raise ValueError(
        f"無法以 Excel 格式讀取或找不到欄位列。最後錯誤：{last_error}"
    )

def read_attendance_file(file_path):
    """自動判斷 HTML 偽裝 XLS 或原生 Excel。"""
    try:
        header = file_path.read_bytes()[:2000].lower()
    except Exception:
        header = b""

    is_html = any(
        marker in header
        for marker in [b"<html", b"<table", b"<!doctype"]
    )

    if is_html:
        return read_html_attendance_file(file_path)

    return read_excel_attendance_file(file_path)

# ==================================================
# 出勤工時計算
# ==================================================

def overlap_seconds(left_start, left_end, right_start, right_end):
    """計算兩段時間重疊秒數。"""
    start = max(left_start, right_start)
    end = min(left_end, right_end)

    if start >= end:
        return 0

    return (end - start).total_seconds()

def adjust_clock_in(clock_in):
    """
    新的上班與遲到規則：

    hh:00       → 從 hh:00 起算，遲到 0 分
    hh:01-hh:20 → 該小時仍算完整 1 小時，另記實際遲到分鐘
    hh:21-hh:30 → 從 hh:30 起算，該小時算 0.5 小時
    hh:31-hh:59 → 從下一個整點起算，該小時不算工時
    """
    hour_start = clock_in.replace(minute=0, second=0, microsecond=0)
    minute = clock_in.minute

    if minute == 0:
        return hour_start, 0

    if 1 <= minute <= 20:
        return hour_start, minute

    if 21 <= minute <= 30:
        return hour_start + pd.Timedelta(minutes=30), 0

    return hour_start + pd.Timedelta(hours=1), 0

def compute_attendance_fields(earliest_ts, latest_ts):
    """
    計算工時、系統加班、狀態、遲到分鐘。

    正常工時區間：
    08:00-09:00
    09:00-10:00
    10:00-11:00
    11:00-12:00
    13:00-14:00
    14:00-15:00
    15:00-16:00
    16:00-17:00
    """
    if pd.isna(earliest_ts) or pd.isna(latest_ts):
        return {
            "工時": None,
            "系統加班": None,
            "狀態": "休",
            "遲到分鐘": 0
        }

    base = pd.Timestamp("1900-01-01")
    clock_in = earliest_ts
    clock_out = latest_ts

    # 單邊卡、相同卡或下班早於上班：先列為異常
    if clock_out <= clock_in:
        return {
            "工時": None,
            "系統加班": None,
            "狀態": "異常",
            "遲到分鐘": 0
        }

    effective_start, late_minutes = adjust_clock_in(clock_in)

    normal_intervals = []

    for start_hour in [8, 9, 10, 11, 13, 14, 15, 16]:
        normal_intervals.append((
            base + pd.Timedelta(hours=start_hour),
            base + pd.Timedelta(hours=start_hour + 1)
        ))

    total_seconds = 0

    for interval_start, interval_end in normal_intervals:
        total_seconds += overlap_seconds(
            effective_start,
            clock_out,
            interval_start,
            interval_end
        )

    raw_work_hours = total_seconds / 3600

    # 只以 0.5 小時為單位顯示，最大為正常 8 小時
    work_hours = min(8, int(raw_work_hours * 2) / 2)

    # 系統加班：
    # 早上 07:00 前上班 → 1 小時
    # 晚上 18:00 後下班 → 1 小時
    # 沒有可判定工時時不計
    morning_ot = clock_in < base + pd.Timedelta(hours=7)
    evening_ot = clock_out >= base + pd.Timedelta(hours=18)

    if work_hours <= 0:
        system_ot = None
    elif morning_ot and evening_ot:
        system_ot = 2
    elif morning_ot or evening_ot:
        system_ot = 1
    else:
        system_ot = None

    if work_hours <= 0:
        status = "異常"
    elif work_hours == 4:
        status = "上半天"
    elif work_hours == 5.5:
        status = "待確認"
    else:
        status = ""

    return {
        "工時": work_hours if work_hours > 0 else None,
        "系統加班": system_ot,
        "狀態": status,
        "遲到分鐘": late_minutes
    }

# ==================================================
# 讀取並清理原始資料
# ==================================================

if not input_file.exists():
    raise FileNotFoundError(f"找不到輸入檔案：{input_file}")

if input_file.stat().st_size == 0:
    raise ValueError(f"輸入檔案為空：{input_file}")

print("讀取檔案：", input_file.name)
print("檔案大小：", input_file.stat().st_size, "bytes")

df = read_attendance_file(input_file)

print("原始欄位：", list(df.columns))

df.columns = [
    clean_column_name(column)
    for column in df.columns
]

print("清理後欄位：", list(df.columns))

if "受訪者" in df.columns:
    print("發現「受訪者」欄位，將不納入出勤計算。")

# 「受訪者」存在於原始檔，但出勤程式不需要它。
expected_columns = [
    "序號",
    "門別",
    "狀態",
    "員編",
    "姓名",
    "部門編號",
    "部門",
    "卡號",
    "日期",
    "時間"
]

missing_columns = [
    column
    for column in expected_columns
    if column not in df.columns
]

if missing_columns:
    raise ValueError(
        f"缺少必要欄位：{missing_columns}\n"
        f"目前欄位：{list(df.columns)}"
    )

df = df[expected_columns].copy()

for column in ["員編", "姓名", "門別", "卡號"]:
    df[column] = df[column].apply(clean_text)

df["門別"] = df["門別"].apply(normalize_location_name)

df["員編"] = (
    df["員編"]
    .str.replace(".0", "", regex=False)
    .str.strip()
)

df["姓名"] = (
    df["員編"]
    .map(name_corrections)
    .fillna(df["姓名"])
)

df["_日期"] = df["日期"].apply(parse_excel_date)
df["_時間"] = df["時間"].apply(parse_excel_time)

df = df[
    df["員編"].notna()
    & (df["員編"] != "")
    & (df["員編"].str.lower() != "nan")
    & df["_日期"].notna()
    & df["_時間"].notna()
].copy()

df = df[
    (df["_日期"].dt.year == REPORT_YEAR)
    & (df["_日期"].dt.month == REPORT_MONTH)
].copy()

if df.empty:
    raise ValueError(
        f"找不到 {REPORT_YEAR}/{REPORT_MONTH:02d} 的有效出勤資料。"
    )

# 自動取當月最後一個「有打卡紀錄」的日期。
last_record_date = df["_日期"].max().normalize()

print("原始有效打卡筆數：", len(df))
print("員工人數：", df["員編"].nunique())
print("最後有紀錄日期：", last_record_date.strftime("%Y/%m/%d"))

df["_原始順序"] = range(len(df))

# ==================================================
# 員工主資料
# ==================================================

last_info = (
    df.sort_values("_原始順序")
    .groupby("員編", sort=False)
    .tail(1)
    [["員編", "姓名", "門別", "卡號"]]
    .copy()
)

last_info["顯示門別"] = (
    last_info["門別"]
    .map(location_names)
    .fillna(last_info["門別"])
)

# ==================================================
# 每日彙整
# ==================================================

daily_attendance = (
    df.groupby(["員編", "_日期"], as_index=False)
    .agg(
        門別=("門別", combine_locations),
        最早時間=("_時間", "min"),
        最晚時間=("_時間", "max")
    )
)

daily_lookup = {
    (row["員編"], row["_日期"].date()): row
    for _, row in daily_attendance.iterrows()
}

# 從該月 1 日，列到實際最後有打卡紀錄的日期。
all_dates = pd.date_range(
    start=pd.Timestamp(
        year=REPORT_YEAR,
        month=REPORT_MONTH,
        day=1
    ),
    end=last_record_date,
    freq="D"
)

# ==================================================
# 建立網頁 JSON
# ==================================================

web_data = []

for _, employee in last_info.iterrows():
    employee_id = employee["員編"]
    attendance = []
    total_late_minutes = 0

    for date in all_dates:
        daily_row = daily_lookup.get(
            (employee_id, date.date())
        )

        if daily_row is None:
            location_full = ""
            location_short = ""
            earliest_time = ""
            latest_time = ""
            work_hours = None
            system_ot = None
            status = "休"
            late_minutes = 0

        else:
            location_full = daily_row["門別"]
            location_short = get_location_short_name(location_full)

            earliest_ts = daily_row["最早時間"]
            latest_ts = daily_row["最晚時間"]

            earliest_time = earliest_ts.strftime("%H:%M:%S")
            latest_time = latest_ts.strftime("%H:%M:%S")

            if employee_id in EXEMPT_EMPLOYEES:
                work_hours = None
                system_ot = None
                status = ""
                late_minutes = 0

            else:
                result = compute_attendance_fields(
                    earliest_ts,
                    latest_ts
                )

                work_hours = result["工時"]
                system_ot = result["系統加班"]
                status = result["狀態"]
                late_minutes = result["遲到分鐘"]

        total_late_minutes += late_minutes

        attendance.append({
            "日期": date.strftime("%Y/%m/%d"),
            "門別": location_full,
            "門別簡稱": location_short,
            "最早時間": earliest_time,
            "最晚時間": latest_time,
            "工時": work_hours,
            "系統加班": system_ot,
            "狀態": status,
            "遲到分鐘": late_minutes,
            "備註": ""
        })

    web_data.append({
        "門別": employee["門別"],
        "顯示門別": employee["顯示門別"],
        "員編": employee_id,
        "姓名": employee["姓名"],
        "卡號": employee["卡號"],
        "遲到累計": total_late_minutes,
        "出勤": attendance
    })

# ==================================================
# 員工排序
# ==================================================

def employee_sort_key(employee):
    employee_id = str(employee["員編"])

    try:
        employee_number = int(employee_id)
    except ValueError:
        employee_number = 999999999

    return (
        employee["顯示門別"],
        employee_number
    )

web_data.sort(key=employee_sort_key)

# ==================================================
# 輸出：indent=2 讓 JSON 方便檢查
# ==================================================

with open(output_file, "w", encoding="utf-8") as file:
    json.dump(
        web_data,
        file,
        ensure_ascii=False,
        indent=2
    )

print("轉換完成：", output_file.name)
print("輸出員工人數：", len(web_data))
print("JSON 檔案大小：", output_file.stat().st_size, "bytes")