import json
from datetime import datetime, time
from pathlib import Path

import pandas as pd

# =========================
# 基本設定
# =========================

folder = Path(__file__).parent

# 每次處理不同月份，只要修改這四行
input_filename = "出入歷史紀錄報表10月.xls"
input_file = folder / input_filename
output_file = folder / "員工出勤資料_10月.json"

REPORT_YEAR = 2026
REPORT_MONTH = 10

# True：由 Python 計算工時、系統加班、狀態、遲到
# False：僅輸出最早/最晚時間，交由 HTML 計算
USE_SCHEME_A = True

# =========================
# 門別顯示名稱與簡稱
# =========================

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

# =========================
# 工具函式
# =========================

def clean_text(value):
    """清理文字中的空白、全形空白、反斜線。"""
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
    """欄名清理，用於辨識受訪者等欄位。"""
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
    """門別移除空白，確保對照表可以匹配。"""
    return clean_text(location).replace(" ", "")


def read_attendance_file(file_path):
    """自動辨識原生 XLS 或 HTML 格式報表。"""

    try:
        file_header = file_path.read_bytes()[:500].lower()
    except Exception:
        file_header = b""

    is_html = any(
        marker in file_header
        for marker in [
            b"<html",
            b"<meta",
            b"<table",
            b"<!doctype"
        ]
    )

    if is_html:
        print("偵測到 HTML 格式報表，使用 read_html 讀取")

        # HTML 報表的標題列可能被讀為第一筆資料。
        tables = pd.read_html(file_path, header=None)

        if not tables:
            raise ValueError(
                f"找不到 HTML 表格：{file_path.name}"
            )

        df = tables[0].copy()

        print("HTML 原始前 3 列：")
        print(df.head(3))

        # 第一列作為欄位標題。
        header_row = df.iloc[0].tolist()

        df = df.iloc[1:].copy()
        df.columns = header_row
        df = df.reset_index(drop=True)

        print("HTML 調整後欄位：", list(df.columns))

        return df

    print("偵測到 XLS 格式報表，使用 xlrd 讀取")

    return pd.read_excel(
        file_path,
        engine="xlrd",
        header=0
    )


def parse_excel_date(value):
    """將 Excel 日期欄位轉為當日 00:00:00 的 Timestamp。"""
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value.normalize()

    if isinstance(value, datetime):
        return pd.Timestamp(value).normalize()

    if isinstance(value, (int, float)):
        try:
            return (
                pd.Timestamp("1899-12-30")
                + pd.to_timedelta(float(value), unit="D")
            ).normalize()
        except Exception:
            return pd.NaT

    text = clean_text(value)

    if not text or text.lower() == "nan":
        return pd.NaT

    parsed = pd.to_datetime(text, errors="coerce")

    if pd.isna(parsed):
        return pd.NaT

    return pd.Timestamp(parsed).normalize()


def parse_excel_time(value):
    """將 Excel 時間欄位轉為可取用時分秒的 Timestamp。"""
    if pd.isna(value):
        return pd.NaT

    if isinstance(value, pd.Timestamp):
        return value

    if isinstance(value, datetime):
        return pd.Timestamp(value)

    if isinstance(value, time):
        return pd.Timestamp(
            f"1900-01-01 {value.strftime('%H:%M:%S')}"
        )

    if isinstance(value, (int, float)):
        try:
            seconds = round(float(value) * 86400)

            return (
                pd.Timestamp("1900-01-01")
                + pd.to_timedelta(seconds, unit="s")
            )
        except Exception:
            return pd.NaT

    text = clean_text(value)

    if not text or text.lower() == "nan":
        return pd.NaT

    parsed = pd.to_datetime(
        text,
        format="%H:%M:%S",
        errors="coerce"
    )

    if pd.isna(parsed):
        parsed = pd.to_datetime(
            text,
            errors="coerce"
        )

    return parsed


def combine_locations(values):
    """同日多門別時，保留不重複的原始門別並以頓號連結。"""
    locations = []

    for value in values:
        location = normalize_location_name(value)

        if (
            location
            and location != "nan"
            and location not in locations
        ):
            locations.append(location)

    return "、".join(locations)


def get_location_short_name(location_full):
    """取得每日出勤欄位要使用的門別簡稱。"""
    location_full = normalize_location_name(location_full)

    if not location_full or location_full == "nan":
        return ""

    # 同日跨門別時，依第一個門別顯示簡稱
    first_location = location_full.split("、")[0]

    if first_location in location_short_names:
        return location_short_names[first_location]

    if "北港" in first_location:
        return "北港"

    if "大里" in first_location:
        return "大里"

    if "南七" in first_location:
        return "南七"

    if "嘉義" in first_location:
        return "嘉配"

    if "大門" in first_location or "九樓" in first_location:
        return "總公司"

    return first_location


def to_base_time(timestamp_value):
    """把任意日期的時間移到 1900-01-01 作區間比較。"""
    base_date = pd.Timestamp("1900-01-01")

    return base_date + pd.to_timedelta(
        timestamp_value.hour * 3600
        + timestamp_value.minute * 60
        + timestamp_value.second,
        unit="s"
    )


def overlap_seconds(start, end, segment_start, segment_end):
    """計算兩個時間區間交集秒數。"""
    left = max(start, segment_start)
    right = min(end, segment_end)

    if left >= right:
        return 0

    return (right - left).total_seconds()

def compute_attendance_fields(earliest_ts, latest_ts):
    """
    使用每日最早與最晚打卡計算：
    - 工時：8 / 4 / 3.5 / 實際時數
    - 系統加班：早上 7-8、下午 17-18
    - 狀態：休 / 異常 / 上半天 / 待確認 / 空白
    - 遲到：只計算 08:00-09:00、13:00-14:00
    """

    if pd.isna(earliest_ts) or pd.isna(latest_ts):
        return {
            "工時": None,
            "系統加班": None,
            "狀態": "休",
            "遲到分鐘": 0
        }

    start_time = to_base_time(earliest_ts)
    end_time = to_base_time(latest_ts)

    base_date = pd.Timestamp("1900-01-01")

    morning_start = base_date + pd.Timedelta(hours=8)
    morning_end = base_date + pd.Timedelta(hours=12)

    afternoon_start = base_date + pd.Timedelta(hours=13)
    afternoon_end = base_date + pd.Timedelta(hours=17)

    # 計算上午 08:00-12:00 工時
    morning_seconds = overlap_seconds(
        start_time,
        end_time,
        morning_start,
        morning_end
    )

    morning_hours = morning_seconds / 3600

    # 計算下午 13:00-17:00 工時
    afternoon_seconds = overlap_seconds(
        start_time,
        end_time,
        afternoon_start,
        afternoon_end
    )

    afternoon_hours = afternoon_seconds / 3600

    # 總工時，不計算 12:00-13:00
    total_hours = morning_hours + afternoon_hours

    # 判斷工時與狀態
    if total_hours >= 7.5:
        work_hours = 8
        status = ""

    elif total_hours >= 4:
        work_hours = round(total_hours, 1)
        status = ""

    elif total_hours >= 3.75:
        work_hours = 4
        status = "上半天"

    elif total_hours >= 3.25:
        work_hours = 3.5
        status = "上半天"

    elif total_hours >= 1:
        work_hours = round(total_hours, 1)
        status = "待確認"

    else:
        work_hours = None
        status = "異常"

    # 系統加班：
    # 早上 07:00 前上班 算 1 小時
    # 下午 18:00 後下班 算 1 小時
    morning_ot_start = base_date + pd.Timedelta(hours=7)
    evening_ot_end = base_date + pd.Timedelta(hours=18)

    morning_ot = (
        work_hours is not None
        and start_time < morning_ot_start  # 07:00 前
    )

    evening_ot = (
        work_hours is not None
        and end_time >= evening_ot_end  # 18:00 後
    )

    if morning_ot and evening_ot:
        system_ot = 2
    elif morning_ot or evening_ot:
        system_ot = 1
    else:
        system_ot = None

    # 遲到只計算兩個區間：
    # 上午 08:00-09:00
    # 下午 13:00-14:00
    late_minutes = 0

    morning_late_end = base_date + pd.Timedelta(hours=9)

    afternoon_late_start = base_date + pd.Timedelta(hours=13)
    afternoon_late_end = base_date + pd.Timedelta(hours=14)

    if (
        start_time > morning_start
        and start_time <= morning_late_end
    ):
        late_minutes = int(
            (start_time - morning_start).total_seconds() // 60
        )

    elif (
        start_time > afternoon_late_start
        and start_time <= afternoon_late_end
    ):
        late_minutes = int(
            (start_time - afternoon_late_start).total_seconds() // 60
        )

    return {
        "工時": work_hours,
        "系統加班": system_ot,
        "狀態": status,
        "遲到分鐘": late_minutes
    }

# =========================
# 讀取原始檔案
# =========================

if not input_file.exists():
    raise FileNotFoundError(
        f"找不到輸入檔案：{input_file}"
    )

if input_file.stat().st_size == 0:
    raise ValueError(
        f"輸入檔案為空：{input_file}"
    )

print("讀取檔案：", input_file.name)
print("檔案大小：", input_file.stat().st_size, "bytes")

df = read_attendance_file(input_file)

print("原始欄位：", list(df.columns))
print("原始欄位數量：", len(df.columns))

# =========================
# 自動清理欄位與忽略受訪者
# =========================

original_columns = list(df.columns)

df.columns = [
    clean_column_name(column)
    for column in df.columns
]

print("清理後欄位：", list(df.columns))

# 僅在實際有「受訪者」欄位時刪除。
# 9 月整理過的檔案沒有此欄，不會錯刪 F 欄部門編號。
if "受訪者" in df.columns:
    df = df.drop(columns=["受訪者"])
    print("已忽略「受訪者」欄位")
else:
    print("未發現「受訪者」欄位，不刪除任何必要欄位")

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
        "缺少必要欄位："
        f"{missing_columns}\n"
        f"目前欄位：{list(df.columns)}\n"
        f"原始欄位：{original_columns}"
    )

# 忽略所有非必要欄位，並固定資料欄位順序。
df = df[expected_columns].copy()

print("最後使用欄位：", list(df.columns))
print("最後欄位數量：", len(df.columns))

# =========================
# 清理資料內容
# =========================

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

df = df[
    df["員編"].notna()
    & (df["員編"] != "")
    & (df["員編"] != "nan")
].copy()

# =========================
# 日期與時間處理
# =========================

df["_日期"] = df["日期"].apply(parse_excel_date)
df["_時間"] = df["時間"].apply(parse_excel_time)

df = df[
    df["_日期"].notna()
    & df["_時間"].notna()
].copy()

df = df[
    (df["_日期"].dt.year == REPORT_YEAR)
    & (df["_日期"].dt.month == REPORT_MONTH)
].copy()

if df.empty:
    raise ValueError(
        f"找不到 {REPORT_YEAR}/{REPORT_MONTH:02d} 的有效出勤資料"
    )

print("原始有效打卡筆數：", len(df))
print("員工人數：", df["員編"].nunique())

df["_原始順序"] = range(len(df))

# =========================
# 員工主資料
# =========================

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

# =========================
# 每日彙整：只保留最早、最晚
# =========================

daily_attendance = (
    df.groupby(
        ["員編", "_日期"],
        as_index=False
    )
    .agg(
        門別=("門別", combine_locations),
        最早時間=("_時間", "min"),
        最晚時間=("_時間", "max")
    )
)

print("每日彙整資料筆數：", len(daily_attendance))

daily_lookup = {
    (row["員編"], row["_日期"].date()): row
    for _, row in daily_attendance.iterrows()
}

# =========================
# 建立整月日期
# =========================

month_start = pd.Timestamp(
    year=REPORT_YEAR,
    month=REPORT_MONTH,
    day=1
)

month_end = month_start + pd.offsets.MonthEnd(1)

all_dates = pd.date_range(
    start=month_start,
    end=month_end,
    freq="D"
)

# =========================
# 建立 JSON 資料
# =========================

web_data = []
# =========================
# 統計每個工地每天的人數
# =========================

site_day_counts = {}

for _, row in daily_attendance.iterrows():
    location = row["門別"]
    date = row["_日期"].date()
    key = (location, date)
    site_day_counts[key] = site_day_counts.get(key, 0) + 1

print("工地人數統計完成，總筆數：", len(site_day_counts))


for _, employee in last_info.iterrows():
    employee_id = employee["員編"]
    attendance = []
    total_late_minutes = 0

    # 找出該員工最後有打卡的日期
    employee_records = daily_attendance[daily_attendance["員編"] == employee_id]
    if not employee_records.empty:
        last_record_date = employee_records["_日期"].max()
    else:
        last_record_date = None

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

            if USE_SCHEME_A:
                result = compute_attendance_fields(
                    earliest_ts,
                    latest_ts
                )

                work_hours = result["工時"]
                system_ot = result["系統加班"]
                status = result["狀態"]
                late_minutes = result["遲到分鐘"]
            else:
                work_hours = None
                system_ot = None
                status = None
                late_minutes = 0

            # 只累計到最後有打卡的日期
            if last_record_date is not None and date <= last_record_date:
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

# =========================
# 員工排序
# =========================

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

# =========================
# 輸出壓縮 JSON
# =========================

with open(output_file, "w", encoding="utf-8") as file:
    json.dump(
        web_data,
        file,
        ensure_ascii=False,
        separators=(",", ":")
    )

print("轉換完成：", output_file.name)
print("輸出員工人數：", len(web_data))
print("JSON 檔案大小：", output_file.stat().st_size, "bytes")