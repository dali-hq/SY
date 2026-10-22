import json
from io import StringIO
from pathlib import Path

import pandas as pd

# ==================================================
# 基本設定：每次處理不同月份，只需修改此區
# ==================================================

folder = Path(__file__).parent

REPORT_YEAR = 2026
REPORT_MONTH = 10

input_filename = "出入歷史紀錄報表10月.xls"
output_filename = "員工出勤資料_10月.json"

input_file = folder / input_filename
output_file = folder / output_filename

# 老闆帳號：保留打卡資料，不計算工時、系統加班及遲到
EXEMPT_EMPLOYEES = {"10001", "10002"}

# 各年月應休天數
# 國定假日日期日後可再補；目前先使用已確認的月應休總數。
MONTHLY_REQUIRED_REST_DAYS = {
    (2026, 9): { "工地": 10, "總公司": 10 },
    (2026, 10): { "工地": 10,  "總公司": 11 },
    (2026, 11): { "工地": 8,  "總公司": 9 },
    (2026, 12): { "工地": 9,  "總公司": 9 }
}

# 總公司門別代碼
HEAD_OFFICE_GATES = { "10001大門", "11211九樓" }

# 林千惠的個別班表
# 不加「完整工時」文字標示；
# 有兩筆以上打卡且涵蓋 09:00–15:30 時，正常工時記為 5.5。
EMPLOYEE_SCHEDULES = {
    "10008": {
        "姓名": "林千惠",
        "上班": "09:00",
        "下班": "15:30",
        "工時": 5.5
    }
}

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
    "33036": "黃明山",
    "33026": "阮曰山"
}

# ==================================================
# 文字、日期、時間工具
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
    """清理欄位名稱中的空白與換行。"""
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
    """轉換 Excel 日期為 pandas Timestamp。"""
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
                return pd.to_datetime(
                    text,
                    format=fmt
                ).normalize()
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
    """將 Excel 時間轉成固定基準日的 pandas Timestamp。"""
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

    try:
        number = float(value)

        # Excel 時間常以 0 至 1 的小數表示
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

def get_location_full_name(location):
    """取得門別完整顯示名稱。"""
    return location_names.get(location, location or "")

def get_location_short_name(location):
    """取得門別簡稱。"""
    if not location:
        return ""

    if location in location_short_names:
        return location_short_names[location]

    for original_name, short_name in location_short_names.items():
        if original_name in location:
            return short_name

    return location

def is_head_office(employee):
    """依員工最後門別判斷是否屬總公司。"""
    return employee["門別"] in HEAD_OFFICE_GATES

# ==================================================
# 讀取 Excel 或 HTML 偽裝 XLS
# ==================================================

def has_required_headers(values):
    """判斷一列是否包含必要欄位名稱。"""
    headers = [clean_column_name(value) for value in values]

    return (
        "員編" in headers
        and "日期" in headers
        and "時間" in headers
    )

def make_dataframe_from_header_row(raw_df, header_index):
    """將找到的欄位列設為表頭。"""
    headers = [
        clean_column_name(value)
        for value in raw_df.iloc[header_index].tolist()
    ]

    data = raw_df.iloc[header_index + 1:].copy()
    data.columns = headers

    return data.dropna(how="all").reset_index(drop=True)

def read_html_attendance_file(file_path):
    """讀取副檔名為 XLS 的 HTML 報表。"""
    raw_bytes = file_path.read_bytes()

    for encoding in ["cp950", "big5", "utf-8", "latin1"]:
        try:
            content = raw_bytes.decode(encoding, errors="ignore")
            tables = pd.read_html(StringIO(content), header=None)

            for table in tables:
                for index in range(min(30, len(table))):
                    if has_required_headers(table.iloc[index].tolist()):
                        return make_dataframe_from_header_row(
                            table,
                            index
                        )
        except Exception:
            continue

    raise ValueError(
        "HTML 報表中找不到「員編、日期、時間」欄位列。"
    )

def read_excel_attendance_file(file_path):
    """讀取原生 XLS/XLSX，並自動尋找欄位列。"""
    last_error = None

    for engine in ["xlrd", None]:
        try:
            if engine is None:
                raw_df = pd.read_excel(
                    file_path,
                    header=None
                )
            else:
                raw_df = pd.read_excel(
                    file_path,
                    header=None,
                    engine=engine
                )

            for index in range(min(30, len(raw_df))):
                if has_required_headers(raw_df.iloc[index].tolist()):
                    return make_dataframe_from_header_row(
                        raw_df,
                        index
                    )

        except Exception as error:
            last_error = error

    raise ValueError(
        "無法讀取 Excel 或找不到必要欄位列。"
        f"最後錯誤：{last_error}"
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
# 工時、加班及狀態計算
# ==================================================

def overlap_seconds(left_start, left_end, right_start, right_end):
    """計算兩段時間重疊秒數。"""
    start = max(left_start, right_start)
    end = min(left_end, right_end)

    if start >= end:
        return 0

    return (end - start).total_seconds()

def round_half_hour_down(value):
    """將時間向下對齊至半小時。"""
    minute = value.minute
    return value.replace(
        minute=0 if minute < 30 else 30,
        second=0,
        microsecond=0
    )

def round_half_hour_up(value):
    """將時間向上對齊至半小時。"""
    rounded = round_half_hour_down(value)
    if rounded < value:
        rounded += pd.Timedelta(minutes=30)
    return rounded

def overlap_hours(start, end, interval_start, interval_end):
    """計算兩個時間區間的重疊小時數。"""
    overlap_start = max(start, interval_start)
    overlap_end = min(end, interval_end)
    if overlap_start >= overlap_end:
        return 0.0
    return (overlap_end - overlap_start).total_seconds() / 3600

def compute_attendance_fields(
    employee_id,
    earliest_ts,
    latest_ts,
    punch_count
):
    """
    計算工時、系統加班、狀態、遲到分鐘與計休時數。

    - 未打卡由外層判定為「休」。
    - 單邊打卡或上下班間隔小於 30 分鐘，由外層判定為「異常/判休」。
    - 一般班別：08:00–17:00，午休 12:00–13:00 不計。
    - 接小孩班別：09:00–15:30，完整上班為 5.5 小時。
    - 遲到、早退與加班均適用 ±2 分鐘容許值。
    - 上班與加班以半小時為最小單位。
    """
    clock_in = earliest_ts
    clock_out = latest_ts
    tolerance = pd.Timedelta(minutes=2)
    base = pd.Timestamp("1900-01-01")

    if pd.isna(clock_in) or pd.isna(clock_out):
        return {
            "工時": None,
            "系統加班": None,
            "狀態": "異常/判休",
            "遲到分鐘": 0,
            "計休時數": 8.0
        }

    if employee_id in EMPLOYEE_SCHEDULES:
        schedule = EMPLOYEE_SCHEDULES[employee_id]
        schedule_start = base + pd.to_timedelta(
            int(schedule["上班"][:2]) * 3600
            + int(schedule["上班"][3:]) * 60,
            unit="s"
        )
        schedule_end = base + pd.to_timedelta(
            int(schedule["下班"][:2]) * 3600
            + int(schedule["下班"][3:]) * 60,
            unit="s"
        )
        required_hours = float(schedule["工時"])

        effective_in = (
            schedule_start
            if clock_in <= schedule_start + tolerance
            else round_half_hour_up(clock_in)
        )
        effective_out = (
            schedule_end
            if clock_out >= schedule_end - tolerance
            else round_half_hour_down(clock_out)
        )

        work_intervals = [
            (schedule_start, base + pd.Timedelta(hours=12)),
            (base + pd.Timedelta(hours=13), schedule_end)
        ]
        work_hours = sum(
            overlap_hours(effective_in, effective_out, start, end)
            for start, end in work_intervals
        )
        work_hours = max(0.0, min(required_hours, work_hours))
        work_hours = int(work_hours * 2) / 2

        late_minutes = 0
       
        rest_hours = max(0.0, required_hours - work_hours)

        return {
            "工時": work_hours if work_hours > 0 else None,
            "系統加班": None,
            "狀態": "",
            "遲到分鐘": late_minutes,
            "計休時數": rest_hours
        }

    schedule_start = base + pd.Timedelta(hours=8)
    schedule_end = base + pd.Timedelta(hours=17)
    required_hours = 8.0

    effective_in = (
        schedule_start
        if clock_in <= schedule_start + tolerance
        else round_half_hour_up(clock_in)
    )
    effective_out = (
        schedule_end
        if clock_out >= schedule_end - tolerance
        else round_half_hour_down(clock_out)
    )

    work_intervals = [
        (base + pd.Timedelta(hours=8), base + pd.Timedelta(hours=12)),
        (base + pd.Timedelta(hours=13), base + pd.Timedelta(hours=17))
    ]
    work_hours = sum(
        overlap_hours(effective_in, effective_out, start, end)
        for start, end in work_intervals
    )
    work_hours = max(0.0, min(required_hours, work_hours))
    work_hours = int(work_hours * 2) / 2

    late_minutes = 0
    
    # 加班：07:00–08:00 與 17:00–18:00，以半小時為單位。
    morning_ot = overlap_hours(
        clock_in,
        clock_out,
        base + pd.Timedelta(hours=7),
        base + pd.Timedelta(hours=8)
    )
    evening_ot = overlap_hours(
        clock_in,
        clock_out,
        base + pd.Timedelta(hours=17),
        base + pd.Timedelta(hours=18)
    )
    morning_ot = int(morning_ot * 2) / 2
    evening_ot = int(evening_ot * 2) / 2
    system_ot = morning_ot + evening_ot
    system_ot = system_ot if system_ot > 0 else None

    rest_hours = max(0.0, required_hours - work_hours)

    return {
        "工時": work_hours if work_hours > 0 else None,
        "系統加班": system_ot,
        "狀態": "",
        "遲到分鐘": late_minutes,
        "計休時數": rest_hours
    }

# ==================================================
# 檢查並讀取檔案
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

# 原始表的「受訪者」欄位不參與計算
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

df["_原始順序"] = range(len(df))
last_record_date = df["_日期"].max().normalize()

print("有效打卡筆數：", len(df))
print("員工人數：", df["員編"].nunique())
print("最後有紀錄日期：", last_record_date.strftime("%Y/%m/%d"))

# ==================================================
# 員工主資料：取報表中最後一筆門別
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
# 每日彙整：最早、最晚刷卡各自帶出門別
# ==================================================

daily_rows = []

# 上下班間隔小於 30 分鐘，或只有單筆打卡，判定為異常/判休。
# 單筆打卡：12:00 含以前顯示為上班；12:01 以後顯示為下班。
SHORT_PUNCH_LIMIT_MINUTES = 30
MIDDAY_SPLIT = pd.Timestamp("1900-01-01 12:00:00")

for (employee_id, work_date), group in df.groupby(
    ["員編", "_日期"],
    sort=False
):
    group = group.sort_values(
        ["_時間", "_原始順序"],
        ascending=[True, True]
    )

    first_punch = group.iloc[0]
    last_punch = group.iloc[-1]
    first_time = first_punch["_時間"]
    last_time = last_punch["_時間"]
    punch_count = len(group)
    interval_minutes = (last_time - first_time).total_seconds() / 60

    in_time = first_time
    out_time = last_time
    in_gate = first_punch["門別"]
    out_gate = last_punch["門別"]
    short_interval_abnormal = (
        punch_count == 1 or interval_minutes < SHORT_PUNCH_LIMIT_MINUTES
    )

    if short_interval_abnormal:
        if first_time <= MIDDAY_SPLIT:
            # 12:00 含以前：只填上班，下班留白。
            in_time = first_time
            in_gate = first_punch["門別"]
            out_time = pd.NaT
            out_gate = ""
        else:
            # 12:01 以後：只填下班，上班留白。
            in_time = pd.NaT
            in_gate = ""
            out_time = last_time
            out_gate = last_punch["門別"]

    daily_rows.append({
        "員編": employee_id,
        "_日期": work_date,
        "打卡筆數": punch_count,
        "上班門別": in_gate,
        "下班門別": out_gate,
        "最早時間": in_time,
        "最晚時間": out_time,
        "短間距異常": short_interval_abnormal
    })

daily_attendance = pd.DataFrame(daily_rows)

daily_lookup = {
    (row["員編"], row["_日期"].date()): row
    for _, row in daily_attendance.iterrows()
}

# 只列至當月最後一筆有效打卡日期，不列未來日期
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

month_rest_days = MONTHLY_REQUIRED_REST_DAYS.get(
    (REPORT_YEAR, REPORT_MONTH),
    {"工地": 0, "總公司": 0}
)

def format_rest_display(hours):
    """將休假時數顯示為「X天YH」；整數小時不顯示 H 以外的小數。"""
    days = int(hours // 8)
    remaining_hours = hours - days * 8

    if remaining_hours == 0:
        return f"{days}天"

    remaining_text = f"{remaining_hours:g}"
    return f"{days}天{remaining_text}H"

web_data = []

for _, employee in last_info.iterrows():
    employee_id = employee["員編"]
    attendance = []

    total_late_minutes = 0
    total_system_overtime = 0
    rest_hours_total = 0.0

    employee_is_exempt = employee_id in EXEMPT_EMPLOYEES
    employee_is_head_office = is_head_office(employee)
    rest_type = "總公司" if employee_is_head_office else "工地"
    required_rest_days = month_rest_days[rest_type]

    for date in all_dates:
        rest_hours = 0.0
        daily_row = daily_lookup.get(
            (employee_id, date.date())
        )

        if daily_row is None:
            # 完全沒有打卡時標示休
            in_gate = ""
            out_gate = ""
            in_gate_short = ""
            out_gate_short = ""
            earliest_time = ""
            latest_time = ""
            punch_count = 0
            work_hours = None
            system_ot = None
            status = "休"
            late_minutes = 0
            rest_hours = 8.0

        else:
            in_gate = daily_row["上班門別"]
            out_gate = daily_row["下班門別"]

            in_gate_short = get_location_short_name(in_gate)
            out_gate_short = get_location_short_name(out_gate)

            earliest_ts = daily_row["最早時間"]
            latest_ts = daily_row["最晚時間"]
            punch_count = int(daily_row["打卡筆數"])

            earliest_time = (
                "" if pd.isna(earliest_ts)
                    else earliest_ts.strftime("%H:%M:%S")
                    )
            latest_time = (
                    "" if pd.isna(latest_ts)
                        else latest_ts.strftime("%H:%M:%S")
                        )
            if employee_is_exempt:
                # 老闆保留原始打卡資料，不進行出勤判定
                work_hours = None
                system_ot = None
                status = ""
                late_minutes = 0

            else:
                # 短間距或單筆資料已於每日摘要階段標示異常；
                # 傳入原始MIN/MAX以保留既有工時計算行為，再強制狀態為異常。
                result = compute_attendance_fields(
                    employee_id=employee_id,
                    earliest_ts=earliest_ts,
                    latest_ts=latest_ts,
                    punch_count=punch_count
                )

                work_hours = result["工時"]
                system_ot = result["系統加班"]
                status = (
                    "異常/判休"
                    if bool(daily_row.get("短間距異常", False))
                    else result["狀態"]
                )
                late_minutes = result["遲到分鐘"]
                rest_hours = (
                    8.0
                    if status == "異常/判休"
                    else result["計休時數"]
                )

        total_late_minutes += late_minutes

        if system_ot is not None:
            total_system_overtime += system_ot

        rest_hours_total += rest_hours

        attendance.append({
            "日期": date.strftime("%Y/%m/%d"),
            "上班門別": in_gate,
            "上班門別簡稱": in_gate_short,
            "最早時間": earliest_time,
            "下班門別": out_gate,
            "下班門別簡稱": out_gate_short,
            "最晚時間": latest_time,
            "打卡筆數": punch_count,
            "工時": work_hours,
            "系統加班": system_ot,
            "狀態": status,
            "遲到分鐘": late_minutes,
            "計休時數": rest_hours,
            "備註": ""
        })

    web_data.append({
        "門別": employee["門別"],
        "顯示門別": employee["顯示門別"],
        "員編": employee_id,
        "姓名": employee["姓名"],
        "卡號": employee["卡號"],
        "休假類別": rest_type,
        "應休天數": required_rest_days,
        "已休時數": rest_hours_total,
        "已休假": format_rest_display(rest_hours_total),
        "實際休假天數": rest_hours_total / 8,
        "未休時數": max(
            0.0,
            required_rest_days * 8 - rest_hours_total
        ),
        "未休假": format_rest_display(
            max(
                0.0,
                required_rest_days * 8 - rest_hours_total
            )
        ),
        "遲到累計": total_late_minutes,
        "系統加班累計": total_system_overtime,
        "出勤": attendance
    })

# ==================================================
# 排序：顯示門別升冪，再依員編升冪
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
# 輸出縮排 JSON
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
print("報表最後有紀錄日期：", last_record_date.strftime("%Y/%m/%d"))
print("JSON 檔案大小：", output_file.stat().st_size, "bytes")