from __future__ import annotations

import io
import json
from dataclasses import dataclass
from datetime import datetime, timedelta, date, time
from calendar import monthrange
from typing import Dict, List, Tuple, Optional, Set
import sqlite3
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="製造排程系統",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.actual-full-card{
    border: 1.8px solid #22c55e;
    border-radius: 22px;
    background: #f8fafc;
    padding: 14px 14px 16px 14px;
    min-height: 520px;
    margin-bottom: 14px;
}
.actual-card-station{
    font-size: 22px;
    font-weight: 800;
    color: #0f172a;
    margin-bottom: 10px;
}
.actual-card-machine{
    font-size: 18px;
    font-weight: 800;
    color: #0f172a;
    text-align: center;
    margin-bottom: 8px;
}
.actual-card-meta{
    font-size: 14px;
    color: #475569;
    text-align: center;
    margin-bottom: 2px;
}
.actual-card-label{
    font-size: 14px;
    font-weight: 800;
    color: #0f172a;
    margin-top: 12px;
    margin-bottom: 4px;
    text-align: left;
}
.actual-card-submeta{
    font-size: 12px;
    color: #64748b;
    margin-top: 6px;
    margin-bottom: 2px;
    text-align: left;
}
.actual-progress-pct{
    font-size: 12px;
    color: #475569;
    margin-bottom: 3px;
    text-align: left;
}
.actual-track{
    width: 100%;
    height: 10px;
    background: #e5e7eb;
    border-radius: 999px;
    overflow: hidden;
}
.actual-fill{
    height: 100%;
    border-radius: 999px;
}
.status-badge{
    display:inline-block;
    padding:4px 10px;
    border-radius:999px;
    font-size:12px;
    font-weight:800;
    color:#ffffff;
    margin-top:8px;
}
.layout-grid-cell{
    border:2px solid #111827;
    border-radius:10px;
    background:#ffffff;
    padding:10px 8px;
    min-height:88px;
    margin-bottom:10px;
}
.layout-grid-title{
    font-size:14px;
    font-weight:800;
    color:#0f172a;
    margin-bottom:6px;
}
.layout-grid-product{
    font-size:13px;
    color:#334155;
    min-height:18px;
    margin-bottom:6px;
    word-break:break-all;
}
.layout-grid-status{
    font-size:12px;
    font-weight:800;
}
</style>
""", unsafe_allow_html=True)


st.markdown("""
<style>
input[type="checkbox"] { accent-color: #22c55e !important; }
.process-header-cell {
    color: #0f172a;
    font-weight: 800;
    border-bottom: 2px solid #cbd5e1;
    padding-bottom: 6px;
    margin-bottom: 2px;
}
.process-body-cell {
    color: #0f172a;
    font-weight: 700;
    padding-top: 6px;
}
.process-row-divider {
    border-bottom: 1px solid #e5e7eb;
    margin: 6px 0 4px 0;
}
</style>
""", unsafe_allow_html=True)


DB_PATH = Path("scheduler.db")

st.markdown("""
<style>
.block-container {padding-top: 1rem; padding-bottom: 1.4rem; max-width: 1650px;}
[data-testid="stSidebar"] {background: linear-gradient(180deg, #0b1220 0%, #111827 55%, #172033 100%); border-right: 1px solid rgba(255,255,255,0.08);}
[data-testid="stSidebar"] * {color: #e5edf8 !important;}
[data-testid="stSidebar"] .stCaption {color: #9fb0cb !important;}
.sidebar-brand {padding: 0.2rem 0 0.6rem 0;}
.sidebar-brand-title {font-size: 1.25rem; font-weight: 800; color: #ffffff; line-height: 1.2;}
.sidebar-brand-sub {color: #94a3b8; font-size: 0.85rem; margin-top: 0.2rem;}
[data-testid="stSidebar"] details {background: linear-gradient(180deg, rgba(255,255,255,0.045) 0%, rgba(255,255,255,0.02) 100%); border: 1px solid rgba(148,163,184,0.16); border-radius: 16px; padding: 6px 8px 8px 8px; margin-bottom: 12px;}
[data-testid="stSidebar"] summary {color: #000000 !important; font-weight: 800; font-size: 0.96rem;}
[data-testid="stSidebar"] .stButton > button {width: 100%; border-radius: 12px; min-height: 42px; text-align: left; justify-content: flex-start; font-weight: 700; background: rgba(255,255,255,0.02) !important; color: #dbe7f7 !important; border: 1px solid rgba(148,163,184,0.16) !important; box-shadow: none !important;}
[data-testid="stSidebar"] .stButton > button:hover {border-color: rgba(96,165,250,0.55) !important; background: rgba(59,130,246,0.14) !important; color: #ffffff !important;}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important; color: #ffffff !important; border: 1px solid rgba(147,197,253,0.45) !important;}
.app-title {font-size: 1.7rem; font-weight: 800; color: #0f172a; margin-bottom: 0.15rem;}
.app-subtitle {color: #475569; margin-bottom: 1rem;}
.kpi-card {background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%); border: 1px solid #e2e8f0; border-radius: 18px; padding: 16px 18px; box-shadow: 0 6px 18px rgba(15, 23, 42, 0.04);}
.kpi-label {font-size: 0.82rem; color: #64748b; margin-bottom: 0.35rem;}
.kpi-value {font-size: 1.65rem; font-weight: 800; color: #0f172a;}
.panel {background: white; border: 1px solid #e5e7eb; border-radius: 18px; padding: 18px 18px 8px 18px; box-shadow: 0 6px 18px rgba(15, 23, 42, 0.04); margin-bottom: 14px;}
.panel-title {font-size: 1.02rem; font-weight: 800; color: #0f172a; margin-bottom: 0.35rem;}
.panel-subtitle {color: #64748b; font-size: 0.92rem; margin-bottom: 0.8rem;}
.tag {display: inline-block; font-size: 0.76rem; font-weight: 700; color: #1d4ed8; background: #dbeafe; padding: 4px 10px; border-radius: 999px; margin-right: 6px;}
.section-divider {margin-top: 0.5rem; margin-bottom: 0.75rem; border-top: 1px solid rgba(148,163,184,0.18);}
.small-note {color: #64748b; font-size: 0.86rem;}
.station-card {background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%); border: 2px solid #22c55e; border-radius: 20px; padding: 14px 14px 12px 14px; min-height: 310px; box-shadow: 0 8px 18px rgba(15, 23, 42, 0.06);}
.station-top {display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px;}
.station-top-left {color: #64748b; font-size: 0.82rem; font-weight: 700;}
.station-slot {font-size: 1.45rem; font-weight: 900; color: #0f172a; margin-top: 6px;}
.station-chip {padding: 5px 10px; border-radius: 999px; font-size: 0.8rem; font-weight: 800; color: #166534; background: #dcfce7; border: 1px solid #4ade80;}
.station-order {font-size: 0.98rem; font-weight: 900; color: #0f172a; text-align: center; margin-top: 4px;}
.station-meta {text-align: center; color: #64748b; font-size: 0.84rem; line-height: 1.45; margin-top: 6px;}
.station-row {display: flex; justify-content: space-between; align-items: center; color: #0f172a; font-weight: 800; font-size: 0.9rem; margin-top: 12px;}
.progress-wrap {width: 100%; height: 8px; background: #d1d5db; border-radius: 999px; overflow: hidden; margin-top: 6px;}
.progress-fill-step {height: 100%; background: linear-gradient(90deg, #7c3aed 0%, #4f46e5 100%); border-radius: 999px;}
.progress-fill-total {height: 100%; background: linear-gradient(90deg, #6366f1 0%, #8b5cf6 100%); border-radius: 999px;}
.station-foot {text-align: center; color: #6b7280; font-size: 0.8rem; line-height: 1.35; margin-top: 12px;}
.empty-card {border: 1px dashed #cbd5e1; border-radius: 20px; padding: 30px 18px; min-height: 240px; background: #f8fafc; text-align: center; color: #64748b;}
</style>
""", unsafe_allow_html=True)

DEFAULT_START_DATETIME = datetime(2026, 1, 1, 8, 0)

DEFAULT_AREA_CONFIG = {
    "A": {"name": "A區 / FIMS-1", "positions": 6, "group_a_count": 3, "group_a_skill": "FIMS-1 組裝", "group_b_count": 3, "group_b_skill": "FIMS-1 支援"},
    "B": {"name": "B區 / FIMS-2", "positions": 6, "group_a_count": 3, "group_a_skill": "FIMS-2 組裝", "group_b_count": 3, "group_b_skill": "FIMS-2 支援"},
    "C": {"name": "C區 / Robot", "positions": 5, "group_a_count": 3, "group_a_skill": "Robot 組裝", "group_b_count": 2, "group_b_skill": "Robot 配線"},
    "D": {"name": "D區 / STK4", "positions": 10, "group_a_count": 4, "group_a_skill": "STK4 組裝", "group_b_count": 3, "group_b_skill": "STK4 檢證"},
    "E": {"name": "E區 / STK5", "positions": 6, "group_a_count": 3, "group_a_skill": "STK5 組裝", "group_b_count": 2, "group_b_skill": "STK5 檢證"},
}

DEFAULT_HOURS_CONFIG = pd.DataFrame([
    {"模組": "STK4", "標準工時(hrs)": 144.0},
    {"模組": "FR301", "標準工時(hrs)": 60.0},
    {"模組": "FIMS-1-4", "標準工時(hrs)": 34.5},
    {"模組": "FIMS-2-4", "標準工時(hrs)": 37.0},
    {"模組": "STK5", "標準工時(hrs)": 218.0},
    {"模組": "FV501", "標準工時(hrs)": 71.5},
    {"模組": "FIMS-1-5", "標準工時(hrs)": 37.0},
    {"模組": "FIMS-2-5", "標準工時(hrs)": 34.5},
])

DEFAULT_CALENDAR_SETTINGS = {
    "default_daily_hours": 8.0,
    "weekday_ot_hours": 0.0,
    "holiday_work_hours": 0.0,
    "special_workdays_text": "",
    "holidays_text": "",
    "stk_precheck_ratio": 0.80,
}


DEFAULT_SETUP_TIME_CONFIG = pd.DataFrame([
    {"區域代碼": "A", "換線時間(hrs)": 0.5},
    {"區域代碼": "B", "換線時間(hrs)": 0.5},
    {"區域代碼": "C", "換線時間(hrs)": 0.5},
    {"區域代碼": "D", "換線時間(hrs)": 1.0},
    {"區域代碼": "E", "換線時間(hrs)": 1.0},
])

DEFAULT_QUALITY_TIME_SETTINGS = {
    "ipqc_hours": 0.5,
    "fqc_hours": 0.5,
    "oqc_hours": 1.0,
}


def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def ensure_table_columns(conn, table_name: str, required_columns: dict):
    info = pd.read_sql_query(f"PRAGMA table_info({table_name})", conn)
    existing_cols = set(info["name"].tolist()) if not info.empty else set()
    for col, col_type in required_columns.items():
        if col not in existing_cols:
            conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {col} {col_type}")
    conn.commit()

def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
    CREATE TABLE IF NOT EXISTS official_schedule (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        version_no INTEGER,
        project_no TEXT,
        product_type TEXT,
        unit_no INTEGER,
        priority INTEGER,
        step TEXT,
        step_label TEXT,
        area TEXT,
        area_name TEXT,
        slot_no INTEGER,
        hours REAL,
        planned_start TEXT,
        planned_finish TEXT,
        due_date TEXT,
        machine_no TEXT,
        estimated_ship_datetime TEXT,
        estimated_ship_date TEXT,
        can_meet_due INTEGER,
        days_late INTEGER
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS machine_master (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        machine_no TEXT UNIQUE,
        product_type TEXT,
        station_code TEXT,
        start_date TEXT,
        robot_due_date TEXT,
        fims_due_date TEXT,
        crate_due_date TEXT,
        module_due_date TEXT,
        created_at TEXT,
        updated_at TEXT
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS production_report (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        machine_no TEXT,
        report_date TEXT,
        product_type TEXT,
        current_process_code TEXT,
        current_process_label TEXT,
        current_process_pct REAL,
        abnormal_type TEXT,
        current_status TEXT,
        abnormal_reason TEXT,
        abnormals_json TEXT,
        daily_note TEXT,
        checklist_json TEXT,
        progress_pct REAL,
        completed_std_hours REAL,
        total_std_hours REAL,
        created_at TEXT
    )
    """)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS machine_realtime_status (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        machine_no TEXT UNIQUE,
        status_code TEXT,
        updated_at TEXT
    )
    """)
    conn.commit()

    ensure_table_columns(conn, "machine_master", {
        "station_code": "TEXT",
        "start_date": "TEXT",
        "robot_due_date": "TEXT",
        "fims_due_date": "TEXT",
        "crate_due_date": "TEXT",
        "module_due_date": "TEXT",
        "created_at": "TEXT",
        "updated_at": "TEXT",
    })
    ensure_table_columns(conn, "production_report", {
        "current_process_code": "TEXT",
        "current_process_label": "TEXT",
        "current_process_pct": "REAL",
        "abnormal_type": "TEXT",
        "current_status": "TEXT",
        "abnormal_reason": "TEXT",
        "abnormals_json": "TEXT",
        "daily_note": "TEXT",
        "checklist_json": "TEXT",
        "progress_pct": "REAL",
        "completed_std_hours": "REAL",
        "total_std_hours": "REAL",
        "created_at": "TEXT",
    })
    conn.close()

def load_official_schedule_db(version_no: int | None = None) -> pd.DataFrame:
    conn = get_conn()
    if version_no is None:
        q = "SELECT * FROM official_schedule ORDER BY version_no, planned_start"
        df = pd.read_sql_query(q, conn)
    else:
        q = "SELECT * FROM official_schedule WHERE version_no = ? ORDER BY planned_start"
        df = pd.read_sql_query(q, conn, params=(version_no,))
    conn.close()
    if df.empty:
        return df
    for col in ["planned_start","planned_finish","due_date","estimated_ship_datetime","estimated_ship_date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    if "can_meet_due" in df.columns:
        df["can_meet_due"] = df["can_meet_due"].astype(bool)
    return df

def next_official_version_no() -> int:
    conn = get_conn()
    row = conn.execute("SELECT COALESCE(MAX(version_no),0) FROM official_schedule").fetchone()
    conn.close()
    return int(row[0]) + 1

def save_official_schedule_db(df: pd.DataFrame):
    if df.empty:
        return
    conn = get_conn()
    save = df.copy()
    save["version_no"] = next_official_version_no()
    for col in ["planned_start","planned_finish","due_date","estimated_ship_datetime","estimated_ship_date"]:
        if col in save.columns:
            save[col] = pd.to_datetime(save[col], errors="coerce").astype(str)
    if "can_meet_due" in save.columns:
        save["can_meet_due"] = save["can_meet_due"].astype(int)
    save.to_sql("official_schedule", conn, if_exists="append", index=False)
    conn.close()

def list_official_versions() -> list[int]:
    conn = get_conn()
    rows = conn.execute("SELECT DISTINCT version_no FROM official_schedule ORDER BY version_no DESC").fetchall()
    conn.close()
    return [int(r[0]) for r in rows]

PRODUCT_OPTIONS = ["STK4","STK5","單獨FIMS-1-4","單獨FIMS-2-4","單獨FR301","單獨FIMS-1-5","單獨FIMS-2-5","單獨FV501"]

STK4_REPORT_PROCESS_MASTER = pd.DataFrame([
    {"分類": "FRAME", "工序編號": "100", "作業人數": 1, "標準工時(時)": 13.0},
    {"分類": "螢幕", "工序編號": "200", "作業人數": 1, "標準工時(時)": 5.0},
    {"分類": "LoadPort", "工序編號": "300", "作業人數": 1, "標準工時(時)": 23.0},
    {"分類": "FIMS", "工序編號": "500", "作業人數": 1, "標準工時(時)": 1.0},
    {"分類": "層板", "工序編號": "600", "作業人數": 1, "標準工時(時)": 12.0},
    {"分類": "SHUTTER", "工序編號": "700", "作業人數": 1, "標準工時(時)": 3.0},
    {"分類": "電盤", "工序編號": "800", "作業人數": 1, "標準工時(時)": 25.0},
    {"分類": "robot", "工序編號": "1000", "作業人數": 1, "標準工時(時)": 0.0},
    {"分類": "檢證作業", "工序編號": "檢證作業", "作業人數": 1, "標準工時(時)": 26.0},
    {"分類": "自檢", "工序編號": "自檢", "作業人數": 1, "標準工時(時)": 8.0},
    {"分類": "IPQC", "工序編號": "IPQC", "作業人數": 1, "標準工時(時)": 10.0},
    {"分類": "FQC", "工序編號": "FQC", "作業人數": 1, "標準工時(時)": 6.0},
    {"分類": "捆包作業", "工序編號": "捆包+OQC", "作業人數": 1, "標準工時(時)": 12.0},
])

REPORT_ABNORMAL_OPTIONS = ["正常", "料件異常", "人員異常", "其他"]
REPORT_STATUS_OPTIONS = ["作業", "待料", "檢驗", "異常處理", "暫停", "完成"]

COLUMN_MAP = {
    "project_no": "專案號",
    "product_type": "產品",
    "qty": "數量",
    "priority": "優先級",
    "unit_no": "序號",
    "machine_no": "機台編號",
    "area": "區域代碼",
    "area_name": "區域名稱",
    "slot_no": "工位",
    "step": "工序代碼",
    "step_label": "工序",
    "hours": "工時",
    "planned_start": "開始時間",
    "planned_finish": "結束時間",
    "estimated_ship_date": "預計出貨日",
    "due_date": "需求交期",
    "can_meet_due": "可否準時出貨",
    "days_late": "延遲天數",
    "positions": "工位數",
    "manpower": "總人數",
    "group_a_count": "A組人數",
    "group_a_skill": "A組職能",
    "group_b_count": "B組人數",
    "group_b_skill": "B組職能",
    "effective_parallel_capacity": "有效同時產能",
    "busy_slots": "忙碌工位數",
    "utilization_pct": "負荷率",
    "date": "日期",
    "total_hours": "總工時",
    "display": "內容",
}

@dataclass
class Task:
    project_no: str
    product_type: str
    unit_no: int
    priority: int
    step: str
    step_label: str
    area: str
    hours: float
    due_date: date
    depends_on: List[str]
    sequence: int

def zh(df: pd.DataFrame) -> pd.DataFrame:
    return df.rename(columns={c: COLUMN_MAP.get(c, c) for c in df.columns})

def parse_date_text(text: str) -> Set[date]:
    out = set()
    if not text:
        return out
    raw = text.replace("，", ",").replace("\n", ",").split(",")
    for item in raw:
        x = item.strip()
        if not x:
            continue
        try:
            out.add(pd.to_datetime(x).date())
        except Exception:
            pass
    return out

def get_area_config() -> Dict[str, Dict]:
    cfg = {}
    for area, item in st.session_state["area_config"].items():
        group_a = int(item.get("group_a_count", 0))
        group_b = int(item.get("group_b_count", 0))
        cfg[area] = {
            "name": item["name"],
            "positions": int(item["positions"]),
            "group_a_count": group_a,
            "group_a_skill": item.get("group_a_skill", ""),
            "group_b_count": group_b,
            "group_b_skill": item.get("group_b_skill", ""),
            "manpower": group_a + group_b,
        }
    return cfg

def get_hours_map() -> Dict[str, float]:
    df = st.session_state["hours_config_df"].copy()
    df["標準工時(hrs)"] = pd.to_numeric(df["標準工時(hrs)"], errors="coerce").fillna(0.0)
    return {str(r["模組"]): float(r["標準工時(hrs)"]) for _, r in df.iterrows()}

def get_calendar_settings() -> Dict:
    return st.session_state["calendar_settings"]

def effective_manpower(cfg: Dict) -> int:
    manpower = cfg.get("manpower", None)
    if manpower is not None:
        try:
            if pd.notna(manpower):
                return int(manpower)
        except Exception:
            pass
    return int(cfg.get("group_a_count", 0)) + int(cfg.get("group_b_count", 0))

def is_regular_workday(d: date, holidays: Set[date]) -> bool:
    return d.weekday() < 5 and d not in holidays

def is_workday(d: date, settings: Dict) -> bool:
    holidays = parse_date_text(settings["holidays_text"])
    special_workdays = parse_date_text(settings["special_workdays_text"])
    if d in special_workdays:
        return True
    return is_regular_workday(d, holidays)

def daily_work_hours(d: date, settings: Dict) -> float:
    holidays = parse_date_text(settings["holidays_text"])
    special_workdays = parse_date_text(settings["special_workdays_text"])
    if d in special_workdays and not is_regular_workday(d, holidays):
        return float(settings["holiday_work_hours"])
    if is_regular_workday(d, holidays):
        return float(settings["default_daily_hours"]) + float(settings["weekday_ot_hours"])
    return 0.0

def next_work_start(dt: datetime, settings: Dict) -> datetime:
    current = dt.replace(second=0, microsecond=0)
    while True:
        d = current.date()
        hours = daily_work_hours(d, settings)
        if hours <= 0:
            current = datetime.combine(d + timedelta(days=1), time(8, 0))
            continue
        day_start = datetime.combine(d, time(8, 0))
        day_end = day_start + timedelta(hours=hours)
        if current < day_start:
            return day_start
        if current >= day_end:
            current = datetime.combine(d + timedelta(days=1), time(8, 0))
            continue
        return current

def add_work_hours_forward(start_dt: datetime, hours: float, settings: Dict) -> datetime:
    current = next_work_start(start_dt, settings)
    remaining = float(hours)
    while remaining > 1e-9:
        d = current.date()
        day_start = datetime.combine(d, time(8, 0))
        day_end = day_start + timedelta(hours=daily_work_hours(d, settings))
        available = max((day_end - current).total_seconds() / 3600.0, 0.0)
        if available <= 1e-9:
            current = next_work_start(datetime.combine(d + timedelta(days=1), time(8, 0)), settings)
            continue
        consume = min(available, remaining)
        current += timedelta(hours=consume)
        remaining -= consume
        if remaining > 1e-9:
            current = next_work_start(datetime.combine(d + timedelta(days=1), time(8, 0)), settings)
    return current

def infer_product_type(project_no: str) -> Optional[str]:
    s = str(project_no).upper().replace(" ", "").replace("-", "")
    if "STK4" in s or s.endswith("4"):
        return "STK4"
    if "STK5" in s or s.endswith("5"):
        return "STK5"
    return None

def build_slot_code(area: str, slot_no: int) -> str:
    return f"{area}{int(slot_no):02d}"


def normalize_forecast(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "專案號": "project_no", "產品": "product_type", "產品型號": "product_type",
        "數量": "qty", "台數": "qty", "預計出貨日": "due_date", "需求交期": "due_date",
        "出貨日": "due_date", "優先級": "priority", "是否含FIMS": "include_fims", "是否含Robot": "include_robot",
        "project_no": "project_no", "product_type": "product_type", "qty": "qty", "due_date": "due_date", "priority": "priority",
        "include_fims": "include_fims", "include_robot": "include_robot",
    }
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})
    required = {"project_no", "qty", "due_date", "priority", "product_type"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Forecast 缺少必要欄位：{', '.join(sorted(missing))}")
    df = df.copy()
    df["project_no"] = df["project_no"].astype(str)
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce")
    df["priority"] = pd.to_numeric(df["priority"], errors="coerce")
    df["due_date"] = pd.to_datetime(df["due_date"], errors="coerce")
    df["product_type"] = df["product_type"].astype(str)
    if "include_fims" not in df.columns:
        df["include_fims"] = df["product_type"].isin(["STK4","STK5"])
    if "include_robot" not in df.columns:
        df["include_robot"] = df["product_type"].isin(["STK4","STK5"])
    df["include_fims"] = df["include_fims"].fillna(False).astype(bool)
    df["include_robot"] = df["include_robot"].fillna(False).astype(bool)
    df = df[df["project_no"].notna() & df["qty"].notna() & df["priority"].notna() & df["due_date"].notna() & df["product_type"].notna()]
    df["qty"] = df["qty"].astype(int)
    df["priority"] = df["priority"].astype(int)
    df["due_date"] = pd.to_datetime(df["due_date"]).dt.date
    return df[["project_no", "product_type", "qty", "due_date", "priority", "include_fims", "include_robot"]].sort_values(["priority", "due_date", "project_no"]).reset_index(drop=True)



def make_routes(hours_map: Dict[str, float], stk_precheck_ratio: float) -> Dict[str, List[Dict]]:
    r = max(0.0, min(1.0, float(stk_precheck_ratio)))
    return {
        "STK4": [
            {"step": "FIMS-1-4", "label": "FIMS-1-4", "area": "A", "hours": hours_map["FIMS-1-4"], "depends_on": []},
            {"step": "FIMS-2-4", "label": "FIMS-2-4", "area": "B", "hours": hours_map["FIMS-2-4"], "depends_on": ["FIMS-1-4"]},
            {"step": "FR301", "label": "FR301", "area": "C", "hours": hours_map["FR301"], "depends_on": []},
            {"step": "STK4_PRE", "label": "STK4前段", "area": "D", "hours": hours_map["STK4"] * r, "depends_on": []},
            {"step": "STK4_QC", "label": "STK4檢證", "area": "D", "hours": hours_map["STK4"] * (1 - r), "depends_on": ["STK4_PRE", "FIMS-2-4", "FR301"]},
        ],
        "STK5": [
            {"step": "FIMS-1-5", "label": "FIMS-1-5", "area": "A", "hours": hours_map["FIMS-1-5"], "depends_on": []},
            {"step": "FIMS-2-5", "label": "FIMS-2-5", "area": "B", "hours": hours_map["FIMS-2-5"], "depends_on": ["FIMS-1-5"]},
            {"step": "FV501", "label": "FV501", "area": "C", "hours": hours_map["FV501"], "depends_on": []},
            {"step": "STK5_PRE", "label": "STK5前段", "area": "E", "hours": hours_map["STK5"] * r, "depends_on": []},
            {"step": "STK5_QC", "label": "STK5檢證", "area": "E", "hours": hours_map["STK5"] * (1 - r), "depends_on": ["STK5_PRE", "FIMS-2-5", "FV501"]},
        ],
        "單獨FIMS-1-4": [
            {"step": "FIMS-1-4", "label": "FIMS-1-4", "area": "A", "hours": hours_map["FIMS-1-4"], "depends_on": []},
        ],
        "單獨FIMS-2-4": [
            {"step": "FIMS-1-4", "label": "FIMS-1-4(前置)", "area": "A", "hours": hours_map["FIMS-1-4"], "depends_on": []},
            {"step": "FIMS-2-4", "label": "FIMS-2-4", "area": "B", "hours": hours_map["FIMS-2-4"], "depends_on": ["FIMS-1-4"]},
        ],
        "單獨FR301": [
            {"step": "FR301", "label": "FR301", "area": "C", "hours": hours_map["FR301"], "depends_on": []},
        ],
        "單獨FIMS-1-5": [
            {"step": "FIMS-1-5", "label": "FIMS-1-5", "area": "A", "hours": hours_map["FIMS-1-5"], "depends_on": []},
        ],
        "單獨FIMS-2-5": [
            {"step": "FIMS-1-5", "label": "FIMS-1-5(前置)", "area": "A", "hours": hours_map["FIMS-1-5"], "depends_on": []},
            {"step": "FIMS-2-5", "label": "FIMS-2-5", "area": "B", "hours": hours_map["FIMS-2-5"], "depends_on": ["FIMS-1-5"]},
        ],
        "單獨FV501": [
            {"step": "FV501", "label": "FV501", "area": "C", "hours": hours_map["FV501"], "depends_on": []},
        ],
    }



def explode_forecast_to_tasks(forecast_df: pd.DataFrame, routes: Dict[str, List[Dict]]) -> List[Task]:
    tasks: List[Task] = []
    for _, row in forecast_df.iterrows():
        base_route = routes[row["product_type"]]
        route = []
        if row["product_type"] == "STK4":
            for step in base_route:
                if step["step"] in ["FIMS-1-4", "FIMS-2-4"] and not bool(row.get("include_fims", True)):
                    continue
                if step["step"] == "FR301" and not bool(row.get("include_robot", True)):
                    continue
                if step["step"] == "STK4_QC":
                    deps = ["STK4_PRE"]
                    if bool(row.get("include_fims", True)):
                        deps.append("FIMS-2-4")
                    if bool(row.get("include_robot", True)):
                        deps.append("FR301")
                    new_step = dict(step)
                    new_step["depends_on"] = deps
                    route.append(new_step)
                else:
                    route.append(step)
        elif row["product_type"] == "STK5":
            for step in base_route:
                if step["step"] in ["FIMS-1-5", "FIMS-2-5"] and not bool(row.get("include_fims", True)):
                    continue
                if step["step"] == "FV501" and not bool(row.get("include_robot", True)):
                    continue
                if step["step"] == "STK5_QC":
                    deps = ["STK5_PRE"]
                    if bool(row.get("include_fims", True)):
                        deps.append("FIMS-2-5")
                    if bool(row.get("include_robot", True)):
                        deps.append("FV501")
                    new_step = dict(step)
                    new_step["depends_on"] = deps
                    route.append(new_step)
                else:
                    route.append(step)
        else:
            route = base_route
        for unit_no in range(1, int(row["qty"]) + 1):
            for sequence, step in enumerate(route, start=1):
                tasks.append(Task(
                    project_no=str(row["project_no"]),
                    product_type=str(row["product_type"]),
                    unit_no=int(unit_no),
                    priority=int(row["priority"]),
                    step=str(step["step"]),
                    step_label=str(step["label"]),
                    area=str(step["area"]),
                    hours=float(step["hours"]),
                    due_date=row["due_date"],
                    depends_on=list(step["depends_on"]),
                    sequence=sequence,
                ))
    return tasks


class CapacityScheduler:
    def __init__(self, start_datetime: datetime, area_config: Dict[str, Dict], calendar_settings: Dict):
        self.start_datetime = next_work_start(max(start_datetime, datetime.now()), calendar_settings)
        self.area_config = area_config
        self.calendar_settings = calendar_settings
        self.area_slots = {
            area: [dict(last_finish=self.start_datetime) for _ in range(min(cfg["positions"], effective_manpower(cfg)))]
            for area, cfg in area_config.items()
        }
        official_df = st.session_state.get("official_schedule_df", pd.DataFrame())
        if isinstance(official_df, pd.DataFrame) and not official_df.empty:
            tmp = official_df.copy()
            tmp["planned_finish"] = pd.to_datetime(tmp["planned_finish"], errors="coerce")
            tmp = tmp[tmp["planned_finish"].notna()]
            for area, cfg in area_config.items():
                eff = min(cfg["positions"], effective_manpower(cfg))
                for slot_no in range(1, eff + 1):
                    area_rows = tmp[(tmp["area"] == area) & (tmp["slot_no"] == slot_no)]
                    if not area_rows.empty:
                        last_finish = max(area_rows["planned_finish"].max().to_pydatetime(), self.start_datetime)
                        self.area_slots[area][slot_no - 1]["last_finish"] = last_finish

    def schedule_forward(self, tasks: List[Task]) -> pd.DataFrame:
        rows = []
        finish_map: Dict[Tuple[str, int, str], datetime] = {}

        def sort_key(t: Task):
            return (t.priority, t.due_date, t.project_no, t.unit_no, t.sequence)

        for t in sorted(tasks, key=sort_key):
            predecessors = [finish_map[(t.project_no, t.unit_no, dep)] for dep in t.depends_on] if t.depends_on else []
            earliest = max(predecessors) if predecessors else self.start_datetime

            best_slot, best_start, best_finish = None, None, None
            for idx, slot in enumerate(self.area_slots[t.area]):
                candidate_start = max(slot["last_finish"], earliest)
                candidate_start = next_work_start(candidate_start, self.calendar_settings)
                candidate_finish = add_work_hours_forward(candidate_start, t.hours, self.calendar_settings)
                if best_finish is None or candidate_finish < best_finish:
                    best_slot, best_start, best_finish = idx, candidate_start, candidate_finish

            self.area_slots[t.area][best_slot]["last_finish"] = best_finish
            finish_map[(t.project_no, t.unit_no, t.step)] = best_finish

            rows.append({
                "project_no": t.project_no,
                "product_type": t.product_type,
                "unit_no": t.unit_no,
                "priority": t.priority,
                "step": t.step,
                "step_label": t.step_label,
                "area": t.area,
                "area_name": self.area_config[t.area]["name"],
                "slot_no": best_slot + 1,
                "hours": round(t.hours, 2),
                "planned_start": best_start,
                "planned_finish": best_finish,
                "due_date": t.due_date,
            })

        df = pd.DataFrame(rows)
        if df.empty:
            return df

        ship_df = (
            df[df["step"].isin(["STK4_QC", "STK5_QC"])]
            [["project_no", "unit_no", "planned_finish"]]
            .rename(columns={"planned_finish": "estimated_ship_datetime"})
        )
        df = df.merge(ship_df, on=["project_no", "unit_no"], how="left")
        df["estimated_ship_date"] = pd.to_datetime(df["estimated_ship_datetime"]).dt.date
        df["machine_no"] = df["project_no"] + "-" + df["unit_no"].astype(str).str.zfill(3)
        df["can_meet_due"] = df["estimated_ship_date"] <= df["due_date"]
        df["days_late"] = (pd.to_datetime(df["estimated_ship_date"]) - pd.to_datetime(df["due_date"])).dt.days.clip(lower=0)
        return df.sort_values(["priority", "planned_start", "area", "slot_no", "project_no", "unit_no"]).reset_index(drop=True)

    def build_daily_station_view(self, schedule_df: pd.DataFrame) -> pd.DataFrame:
        records = []
        if schedule_df.empty:
            return pd.DataFrame()
        for _, row in schedule_df.iterrows():
            d = row["planned_start"].date()
            end_d = row["planned_finish"].date()
            while d <= end_d:
                if is_workday(d, self.calendar_settings):
                    records.append({
                        "date": d,
                        "area": row["area"],
                        "area_name": row["area_name"],
                        "slot_no": row["slot_no"],
                        "project_no": row["project_no"],
                        "product_type": row["product_type"],
                        "unit_no": row["unit_no"],
                        "priority": row["priority"],
                        "machine_no": row["machine_no"],
                        "step": row["step"],
                        "step_label": row["step_label"],
                        "estimated_ship_date": row["estimated_ship_date"],
                        "display": f'{row["machine_no"]} / {row["step_label"]}',
                    })
                d += timedelta(days=1)
        return pd.DataFrame(records).sort_values(["date", "priority", "area", "slot_no"]).reset_index(drop=True)

    def build_project_summary(self, schedule_df: pd.DataFrame) -> pd.DataFrame:
        if schedule_df.empty:
            return pd.DataFrame()
        return (
            schedule_df.groupby(["project_no", "product_type", "unit_no", "machine_no", "priority"], as_index=False)
            .agg(
                planned_start=("planned_start", "min"),
                planned_finish=("planned_finish", "max"),
                due_date=("due_date", "max"),
                estimated_ship_date=("estimated_ship_date", "max"),
                can_meet_due=("can_meet_due", "max"),
                days_late=("days_late", "max"),
                total_hours=("hours", "sum"),
            )
            .sort_values(["priority", "due_date", "project_no", "unit_no"])
            .reset_index(drop=True)
        )

    def build_area_load(self, schedule_df: pd.DataFrame) -> pd.DataFrame:
        if schedule_df.empty:
            return pd.DataFrame()
        rows = []
        for area, cfg in self.area_config.items():
            area_df = schedule_df[schedule_df["area"] == area].copy()
            if area_df.empty:
                continue
            start_date = area_df["planned_start"].min().date()
            end_date = area_df["planned_finish"].max().date()
            eff_capacity = min(cfg["positions"], effective_manpower(cfg))
            d = start_date
            while d <= end_date:
                if is_workday(d, self.calendar_settings):
                    active = area_df[(area_df["planned_start"].dt.date <= d) & (area_df["planned_finish"].dt.date >= d)]
                    busy = active["slot_no"].nunique()
                    rows.append({
                        "date": d,
                        "area": area,
                        "area_name": cfg["name"],
                        "positions": cfg["positions"],
                        "manpower": effective_manpower(cfg),
                        "effective_parallel_capacity": eff_capacity,
                        "busy_slots": busy,
                        "utilization_pct": busy / eff_capacity if eff_capacity else 0,
                    })
                d += timedelta(days=1)
        return pd.DataFrame(rows).sort_values(["date", "area"]).reset_index(drop=True)

def to_excel_bytes(forecast_df, summary_df, schedule_df, daily_df, load_df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        zh(forecast_df).to_excel(writer, sheet_name="Forecast", index=False)
        zh(summary_df).to_excel(writer, sheet_name="專案彙總", index=False)
        zh(schedule_df).to_excel(writer, sheet_name="排程清單", index=False)
        zh(daily_df).to_excel(writer, sheet_name="工位日報", index=False)
        zh(load_df).to_excel(writer, sheet_name="產能負荷", index=False)
        zh(build_machine_metrics(schedule_df)).to_excel(writer, sheet_name="製程指標", index=False)
        routing = build_routing_flow_table().rename(columns={"product_type":"產品","routing_flow":"製程順序"})
        routing.to_excel(writer, sheet_name="RoutingFlow", index=False)
        setup = build_setup_events(schedule_df)
        if not setup.empty:
            setup.to_excel(writer, sheet_name="換線時間", index=False)
        rework = get_rework_log_df()
        if isinstance(rework, pd.DataFrame) and not rework.empty:
            rework.to_excel(writer, sheet_name="重工返修", index=False)
        erp = st.session_state.get("erp_receipt_df", pd.DataFrame())
        if isinstance(erp, pd.DataFrame) and not erp.empty:
            erp.to_excel(writer, sheet_name="ERP入庫", index=False)
        reports = load_latest_production_reports()
        if isinstance(reports, pd.DataFrame) and not reports.empty:
            reports.to_excel(writer, sheet_name="產線回報", index=False)
    return buffer.getvalue()


def get_setup_time_map() -> Dict[str, float]:
    df = st.session_state.get("setup_time_config_df", pd.DataFrame()).copy()
    if df.empty:
        return {}
    df["換線時間(hrs)"] = pd.to_numeric(df["換線時間(hrs)"], errors="coerce").fillna(0.0)
    return {str(r["區域代碼"]): float(r["換線時間(hrs)"]) for _, r in df.iterrows()}

def get_quality_time_settings() -> Dict[str, float]:
    return st.session_state.get("quality_time_settings", DEFAULT_QUALITY_TIME_SETTINGS.copy())

def build_machine_metrics(schedule_df: pd.DataFrame) -> pd.DataFrame:
    if schedule_df.empty:
        return pd.DataFrame()
    quality = get_quality_time_settings()
    out = []
    work = schedule_df.sort_values(["project_no", "unit_no", "planned_start"]).copy()
    for (project_no, unit_no), grp in work.groupby(["project_no", "unit_no"], as_index=False):
        grp = grp.sort_values("planned_start")
        queue_hours = 0.0
        prev_finish = None
        for _, row in grp.iterrows():
            if prev_finish is not None:
                gap = (row["planned_start"] - prev_finish).total_seconds() / 3600.0
                if gap > 0:
                    queue_hours += gap
            prev_finish = row["planned_finish"]
        cycle_hours = (grp["planned_finish"].max() - grp["planned_start"].min()).total_seconds() / 3600.0
        quality_hours = float(quality.get("ipqc_hours", 0)) + float(quality.get("fqc_hours", 0)) + float(quality.get("oqc_hours", 0))
        out.append({
            "project_no": project_no,
            "unit_no": unit_no,
            "machine_no": grp["machine_no"].iloc[0],
            "product_type": grp["product_type"].iloc[0],
            "priority": grp["priority"].iloc[0],
            "planned_start": grp["planned_start"].min(),
            "planned_finish": grp["planned_finish"].max(),
            "cycle_time_hours": round(cycle_hours, 2),
            "queue_time_hours": round(queue_hours, 2),
            "process_time_hours": round(float(grp["hours"].sum()), 2),
            "quality_time_hours": round(quality_hours, 2),
            "due_date": grp["due_date"].max(),
            "estimated_ship_date": grp["estimated_ship_date"].max(),
            "can_meet_due": bool(pd.to_numeric(grp["can_meet_due"], errors="coerce").fillna(0).max()),
            "days_late": int(pd.to_numeric(grp["days_late"], errors="coerce").fillna(0).max()),
        })
    return pd.DataFrame(out)

def build_setup_events(schedule_df: pd.DataFrame) -> pd.DataFrame:
    if schedule_df.empty:
        return pd.DataFrame()
    setup_map = get_setup_time_map()
    events = []
    work = schedule_df.sort_values(["area", "slot_no", "planned_start"]).copy()
    for (area, slot_no), grp in work.groupby(["area", "slot_no"], as_index=False):
        grp = grp.sort_values("planned_start")
        prev_product = None
        prev_machine = None
        prev_finish = None
        for _, row in grp.iterrows():
            if prev_product is not None and row["product_type"] != prev_product:
                setup_hours = float(setup_map.get(area, 0.0))
                events.append({
                    "area": area,
                    "slot_no": slot_no,
                    "from_machine": prev_machine,
                    "to_machine": row["machine_no"],
                    "from_product": prev_product,
                    "to_product": row["product_type"],
                    "setup_time_hours": round(setup_hours, 2),
                    "setup_start": prev_finish,
                    "setup_end": prev_finish + pd.Timedelta(hours=setup_hours) if prev_finish is not None else pd.NaT,
                })
            prev_product = row["product_type"]
            prev_machine = row["machine_no"]
            prev_finish = row["planned_finish"]
    return pd.DataFrame(events)

def build_routing_flow_table() -> pd.DataFrame:
    routes = make_routes(get_hours_map(), get_calendar_settings()["stk_precheck_ratio"])
    rows = []
    for product, steps in routes.items():
        rows.append({
            "product_type": product,
            "routing_flow": " → ".join([s["label"] for s in steps]),
        })
    return pd.DataFrame(rows)

def get_rework_log_df() -> pd.DataFrame:
    return st.session_state.get("rework_log_df", pd.DataFrame())

def throughput_by_date(summary_df: pd.DataFrame) -> pd.DataFrame:
    if summary_df.empty:
        return pd.DataFrame(columns=["estimated_ship_date", "throughput"])
    out = summary_df.groupby("estimated_ship_date", as_index=False).size().rename(columns={"size": "throughput"})
    return out.sort_values("estimated_ship_date")

def wip_on_date(schedule_df: pd.DataFrame, selected_date: date) -> int:
    if schedule_df.empty:
        return 0
    active = schedule_df[(schedule_df["planned_start"].dt.date <= selected_date) & (schedule_df["planned_finish"].dt.date >= selected_date)]
    return int(active[["project_no", "unit_no"]].drop_duplicates().shape[0])

def bottleneck_summary(load_df: pd.DataFrame) -> tuple[str, float]:
    if load_df.empty:
        return "-", 0.0
    grp = load_df.groupby("area_name", as_index=False)["utilization_pct"].mean().sort_values("utilization_pct", ascending=False)
    return grp.iloc[0]["area_name"], float(grp.iloc[0]["utilization_pct"])


def week_bounds(any_date: date) -> tuple[date, date]:
    start = any_date - timedelta(days=any_date.weekday())
    end = start + timedelta(days=6)
    return start, end

def build_week_receipt_report(summary_df: pd.DataFrame, erp_df: pd.DataFrame, target_date: date) -> tuple[pd.DataFrame, dict]:
    # 全部統一轉成 pandas Timestamp(normalized) 再比較，避免 datetime64 與 date 型別衝突
    week_start, week_end = week_bounds(target_date)
    week_start = pd.Timestamp(week_start).normalize()
    week_end = pd.Timestamp(week_end).normalize()

    if summary_df.empty:
        planned_df = pd.DataFrame(columns=["機台編號", "產品", "優先級", "預計出貨日"])
    else:
        planned_df = summary_df.copy()
        planned_df["estimated_ship_date"] = pd.to_datetime(planned_df["estimated_ship_date"], errors="coerce").dt.normalize()
        planned_df = planned_df[planned_df["estimated_ship_date"].notna()].copy()
        planned_df = planned_df[
            (planned_df["estimated_ship_date"] >= week_start) &
            (planned_df["estimated_ship_date"] <= week_end)
        ].copy()
        planned_df = planned_df.rename(columns={
            "machine_no": "機台編號",
            "product_type": "產品",
            "priority": "優先級",
            "estimated_ship_date": "預計出貨日",
        })
        planned_df = planned_df[["機台編號", "產品", "優先級", "預計出貨日"]]
        if not planned_df.empty:
            planned_df["預計出貨日"] = pd.to_datetime(planned_df["預計出貨日"]).dt.date

    actual_df = erp_df.copy() if isinstance(erp_df, pd.DataFrame) else pd.DataFrame()
    if not actual_df.empty:
        actual_df["入庫日期"] = pd.to_datetime(actual_df["入庫日期"], errors="coerce").dt.normalize()
        actual_df["數量"] = pd.to_numeric(actual_df["數量"], errors="coerce").fillna(0)
        actual_df = actual_df[actual_df["入庫日期"].notna()].copy()
        actual_df = actual_df[
            (actual_df["入庫日期"] >= week_start) &
            (actual_df["入庫日期"] <= week_end)
        ].copy()
        if not actual_df.empty:
            actual_df["入庫日期"] = pd.to_datetime(actual_df["入庫日期"]).dt.date
    else:
        actual_df = pd.DataFrame(columns=["入庫日期", "機台編號", "數量", "來源"])

    planned_units = int(planned_df["機台編號"].nunique()) if not planned_df.empty else 0
    actual_units = int(actual_df["數量"].sum()) if not actual_df.empty else 0
    gap = planned_units - actual_units
    achieve = round((actual_units / planned_units) * 100, 1) if planned_units > 0 else 0.0

    plan_set = set(planned_df["機台編號"].astype(str)) if not planned_df.empty else set()
    actual_set = set(actual_df["機台編號"].astype(str)) if not actual_df.empty else set()

    compare_rows = []
    for m in sorted(plan_set | actual_set):
        compare_rows.append({
            "機台編號": m,
            "本周計畫": "是" if m in plan_set else "否",
            "實際入庫": "是" if m in actual_set else "否",
            "狀態": "已達成" if (m in plan_set and m in actual_set) else ("未入庫" if m in plan_set else "非計畫入庫"),
        })
    compare_df = pd.DataFrame(compare_rows)

    summary = {
        "week_start": week_start.date(),
        "week_end": week_end.date(),
        "planned_units": planned_units,
        "actual_units": actual_units,
        "gap_units": gap,
        "achieve_rate": achieve,
    }
    return planned_df.copy(), {"actual_df": actual_df, "compare_df": compare_df, "summary": summary}



def load_machine_master() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query("SELECT * FROM machine_master ORDER BY machine_no", conn)
    except Exception:
        df = pd.DataFrame(columns=["machine_no", "product_type", "start_date", "robot_due_date", "fims_due_date", "crate_due_date", "module_due_date", "created_at", "updated_at"])
    conn.close()
    for col in ["machine_no", "product_type", "station_code", "start_date", "robot_due_date", "fims_due_date", "crate_due_date", "module_due_date", "created_at", "updated_at"]:
        if col not in df.columns:
            df[col] = None
    return df


def get_station_options() -> list[str]:
    area_cfg = st.session_state.get("area_config", DEFAULT_AREA_CONFIG)
    options = []
    for area, cfg in area_cfg.items():
        try:
            positions = int(cfg.get("positions", 0))
        except Exception:
            positions = 0
        for i in range(1, positions + 1):
            options.append(f"{area}{i:02d}")
    return options

def progress_bar_html(value, color="#22c55e", track="#e5e7eb"):
    try:
        pct = max(0, min(100, int(round(float(value)))))
    except Exception:
        pct = 0
    return f"""
    <div style='margin-top:4px;margin-bottom:8px'>
      <div style='display:flex;justify-content:space-between;font-size:12px;color:#475569;margin-bottom:3px'>
        <span>{pct}%</span>
      </div>
      <div style='width:100%;height:10px;background:{track};border-radius:999px;overflow:hidden'>
        <div style='width:{pct}%;height:100%;background:{color};border-radius:999px'></div>
      </div>
    </div>
    """

def build_actual_station_cards_df() -> pd.DataFrame:
    mm = load_machine_master().copy()
    if mm.empty:
        return pd.DataFrame()

    mm["product_type"] = mm["product_type"].astype(str)
    mm["station_code"] = mm["station_code"].fillna("").astype(str).str.strip().str.upper()
    mm["crate_due_date"] = pd.to_datetime(mm.get("crate_due_date"), errors="coerce")
    mm["robot_due_date"] = pd.to_datetime(mm.get("robot_due_date"), errors="coerce")
    mm["fims_due_date"] = pd.to_datetime(mm.get("fims_due_date"), errors="coerce")
    mm["module_due_date"] = pd.to_datetime(mm.get("module_due_date"), errors="coerce")
    mm = mm[mm["station_code"] != ""].copy()
    if mm.empty:
        return pd.DataFrame()

    reports = load_latest_production_reports()
    if reports.empty:
        reports = pd.DataFrame(columns=["machine_no", "report_date", "created_at", "progress_pct", "checklist_json"])
    else:
        reports["report_date"] = pd.to_datetime(reports["report_date"], errors="coerce")
        reports["created_at"] = pd.to_datetime(reports["created_at"], errors="coerce")
        reports = reports.sort_values(["machine_no", "report_date", "created_at"]).groupby("machine_no", as_index=False).tail(1)

    latest_report_map = {}
    for _, r in reports.iterrows():
        latest_report_map[str(r.get("machine_no", ""))] = r.to_dict()

    def extract_step_progress(machine_no: str, step_code: str) -> float:
        row = latest_report_map.get(str(machine_no))
        if not row:
            return 0.0
        checklist_raw = row.get("checklist_json", "[]")
        try:
            checklist = json.loads(checklist_raw or "[]")
        except Exception:
            checklist = []
        for item in checklist:
            if str(item.get("工序編號", "")).strip() == str(step_code):
                try:
                    return float(item.get("完成%", 0) or 0)
                except Exception:
                    return 0.0
        return 0.0

    rt = load_realtime_status_df()
    rt_map = {}
    if not rt.empty:
        rt_map = {str(r.get("machine_no", "")): str(r.get("status_code", "") or "") for _, r in rt.iterrows()}

    stk = mm[mm["product_type"].isin(["STK4", "STK5"])].copy()
    if stk.empty:
        return pd.DataFrame()

    cards = []
    for _, row in stk.iterrows():
        machine_no = str(row["machine_no"])
        whole_row = latest_report_map.get(machine_no, {})
        whole_progress = float(pd.to_numeric(whole_row.get("progress_pct", 0), errors="coerce") or 0)
        fims_progress = extract_step_progress(machine_no, "500")
        robot_progress = extract_step_progress(machine_no, "1000")

        cards.append({
            "station_code": str(row["station_code"]),
            "machine_no": machine_no,
            "product_type": str(row["product_type"]),
            "crate_due_date": row["crate_due_date"],
            "robot_due_date": row["robot_due_date"],
            "fims_due_date": row["fims_due_date"],
            "whole_progress": whole_progress,
            "fims_progress": fims_progress,
            "robot_progress": robot_progress,
        })
    return pd.DataFrame(cards)

def save_machine_master(df: pd.DataFrame):
    conn = get_conn()
    conn.execute("DELETE FROM machine_master")
    conn.commit()
    if not df.empty:
        x = df.copy()
        for col in ["machine_no", "product_type", "station_code", "start_date", "robot_due_date", "fims_due_date", "crate_due_date", "module_due_date", "created_at", "updated_at"]:
            if col not in x.columns:
                x[col] = None
        now_ts = datetime.now().isoformat(timespec="seconds")
        x["updated_at"] = now_ts
        x["created_at"] = x["created_at"].fillna(now_ts)
        for col in ["start_date", "robot_due_date", "fims_due_date", "crate_due_date", "module_due_date"]:
            x[col] = pd.to_datetime(x[col], errors="coerce").astype(str)
        x = x[["machine_no", "product_type", "station_code", "start_date", "robot_due_date", "fims_due_date", "crate_due_date", "module_due_date", "created_at", "updated_at"]]
        x.to_sql("machine_master", conn, if_exists="append", index=False)
    conn.close()

def get_stk4_process_master() -> pd.DataFrame:
    return STK4_REPORT_PROCESS_MASTER.copy()

def list_stk4_machine_options() -> list[str]:
    frames = []
    mm = load_machine_master()
    if not mm.empty:
        tmp = mm[mm["product_type"].astype(str) == "STK4"][["machine_no"]].drop_duplicates()
        if not tmp.empty:
            frames.append(tmp)
    official = st.session_state.get("official_schedule_df", pd.DataFrame())
    draft = st.session_state.get("draft_schedule_df", pd.DataFrame())
    for df in [official, draft]:
        if isinstance(df, pd.DataFrame) and not df.empty and "machine_no" in df.columns and "product_type" in df.columns:
            tmp = df[df["product_type"] == "STK4"][["machine_no"]].drop_duplicates()
            if not tmp.empty:
                frames.append(tmp)
    if not frames:
        return []
    return sorted(pd.concat(frames, ignore_index=True)["machine_no"].dropna().astype(str).unique().tolist())

def load_latest_production_reports() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query("SELECT * FROM production_report ORDER BY report_date DESC, created_at DESC", conn)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    if df.empty:
        return df
    if "report_date" in df.columns:
        df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    if "created_at" in df.columns:
        df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce")
    return df

def save_production_report(payload: dict):
    conn = get_conn()
    conn.execute("DELETE FROM production_report WHERE machine_no = ? AND report_date = ?", (payload["machine_no"], payload["report_date"]))
    conn.execute(
        """
        INSERT INTO production_report (
            machine_no, report_date, product_type, current_process_code, current_process_label, current_process_pct,
            abnormal_type, current_status, abnormal_reason, abnormals_json, daily_note, checklist_json,
            progress_pct, completed_std_hours, total_std_hours, created_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            payload["machine_no"], payload["report_date"], payload["product_type"], payload["current_process_code"],
            payload["current_process_label"], payload["current_process_pct"], payload["abnormal_type"],
            payload["current_status"], payload["abnormal_reason"], payload["abnormals_json"], payload["daily_note"], payload["checklist_json"],
            payload["progress_pct"], payload["completed_std_hours"], payload["total_std_hours"], payload["created_at"]
        )
    )
    conn.commit()
    conn.close()

def latest_report_up_to(selected_date: date) -> pd.DataFrame:
    df = load_latest_production_reports()
    if df.empty:
        return df
    df = df[pd.to_datetime(df["report_date"], errors="coerce").dt.date <= selected_date].copy()
    if df.empty:
        return df
    df = df.sort_values(["machine_no", "report_date", "created_at"]).groupby("machine_no", as_index=False).tail(1)
    return df


def calc_stk4_progress(checklist_df: pd.DataFrame) -> tuple[float, float, float]:
    work = checklist_df.copy()
    work["標準工時(時)"] = pd.to_numeric(work["標準工時(時)"], errors="coerce").fillna(0.0)
    work["完成"] = work["完成"].fillna(False).astype(bool)
    work["完成%"] = pd.to_numeric(work["完成%"], errors="coerce").fillna(0).clip(lower=0, upper=100).round(0).astype(int)
    work.loc[work["完成"] == True, "完成%"] = 100
    total_std = float(work["標準工時(時)"].sum())
    completed_hours = float((work["標準工時(時)"] * work["完成%"] / 100.0).sum())
    progress_pct = round((completed_hours / total_std) * 100, 0) if total_std > 0 else 0.0
    return progress_pct, round(completed_hours, 2), round(total_std, 2)

def get_default_report_form(machine_no: str, report_date: date) -> dict:
    reports = load_latest_production_reports()
    if reports.empty:
        return {}
    same = reports[(reports["machine_no"].astype(str) == str(machine_no)) & (pd.to_datetime(reports["report_date"]).dt.date == report_date)]
    if same.empty:
        return {}
    row = same.sort_values("created_at").iloc[-1]
    try:
        checklist = json.loads(row.get("checklist_json", "[]") or "[]")
    except Exception:
        checklist = []
    try:
        abnormals = json.loads(row.get("abnormals_json", "[]") or "[]")
    except Exception:
        abnormals = []
    return {
        "current_process_code": row.get("current_process_code", ""),
        "current_process_pct": float(row.get("current_process_pct", 0) or 0),
        "abnormal_type": row.get("abnormal_type", "正常"),
        "current_status": row.get("current_status", "作業"),
        "abnormal_reason": row.get("abnormal_reason", ""),
        "daily_note": row.get("daily_note", ""),
        "checklist": checklist,
        "abnormals": abnormals,
    }



REALTIME_STATUS_OPTIONS = ["正常生產", "檢驗中", "有風險", "異常 / 延遲"]

def load_realtime_status_df() -> pd.DataFrame:
    conn = get_conn()
    try:
        df = pd.read_sql_query("SELECT * FROM machine_realtime_status ORDER BY machine_no", conn)
    except Exception:
        df = pd.DataFrame(columns=["machine_no", "status_code", "updated_at"])
    conn.close()
    for col in ["machine_no", "status_code", "updated_at"]:
        if col not in df.columns:
            df[col] = None
    return df

def save_machine_realtime_status(machine_no: str, status_code: str):
    conn = get_conn()
    now_ts = datetime.now().isoformat(timespec="seconds")
    conn.execute(
        """
        INSERT INTO machine_realtime_status (machine_no, status_code, updated_at)
        VALUES (?, ?, ?)
        ON CONFLICT(machine_no) DO UPDATE SET
            status_code=excluded.status_code,
            updated_at=excluded.updated_at
        """,
        (str(machine_no), str(status_code), now_ts),
    )
    conn.commit()
    conn.close()

def status_color_hex(status_code: str) -> str:
    mapping = {
        "正常生產": "#22c55e",
        "檢驗中": "#16a34a",
        "有風險": "#f59e0b",
        "異常 / 延遲": "#ef4444",
    }
    return mapping.get(str(status_code), "#94a3b8")

def status_emoji(status_code: str) -> str:
    mapping = {
        "正常生產": "🟢",
        "檢驗中": "🟢",
        "有風險": "🟡",
        "異常 / 延遲": "🔴",
    }
    return mapping.get(str(status_code), "⚪")


def build_production_report_list_df() -> pd.DataFrame:
    reports = load_latest_production_reports()
    if reports.empty:
        return reports
    mm = load_machine_master()
    if not mm.empty:
        mm2 = mm.copy()
        for c in ["crate_due_date", "module_due_date", "start_date", "robot_due_date", "fims_due_date"]:
            if c in mm2.columns:
                mm2[c] = pd.to_datetime(mm2[c], errors="coerce")
        reports = reports.merge(mm2[["machine_no", "product_type", "crate_due_date"]].rename(columns={"product_type":"machine_product_type"}), on="machine_no", how="left")
    else:
        reports["crate_due_date"] = pd.NaT
        reports["machine_product_type"] = None

    def extract_abnormal_items(row):
        try:
            arr = json.loads(row.get("abnormals_json", "[]") or "[]")
        except Exception:
            arr = []
        items = [str(x.get("異常類別","")).strip() for x in arr if str(x.get("異常類別","")).strip() and str(x.get("異常類別","")).strip() != "正常"]
        reasons = [str(x.get("異常原因","")).strip() for x in arr if str(x.get("異常原因","")).strip()]
        return pd.Series({
            "abnormal_items": "、".join(items) if items else "正常",
            "abnormal_reasons_joined": "；".join(reasons) if reasons else ""
        })

    extra = reports.apply(extract_abnormal_items, axis=1)
    reports = pd.concat([reports, extra], axis=1)
    reports["report_date"] = pd.to_datetime(reports["report_date"], errors="coerce")
    reports["crate_due_date"] = pd.to_datetime(reports["crate_due_date"], errors="coerce")
    return reports
def build_actual_board_df(selected_date: date) -> pd.DataFrame:
    base = st.session_state.get("official_daily_df", pd.DataFrame()).copy()
    if base.empty:
        base = st.session_state.get("draft_daily_df", pd.DataFrame()).copy()
    if base.empty:
        return pd.DataFrame()
    base = base[base["date"] == selected_date].copy()
    if base.empty:
        return base
    reports = latest_report_up_to(selected_date)
    if reports.empty:
        base["actual_process_label"] = base["step_label"]
        base["progress_pct"] = None
        base["abnormal_type"] = ""
        base["current_status"] = ""
        base["daily_note"] = ""
        return base
    merge_cols = ["machine_no", "current_process_label", "progress_pct", "abnormal_type", "current_status", "daily_note", "abnormal_reason", "report_date"]
    merged = base.merge(reports[merge_cols], on="machine_no", how="left")
    merged["actual_process_label"] = merged["current_process_label"].fillna(merged["step_label"])
    return merged
def seed_forecast() -> pd.DataFrame:
    return pd.DataFrame([
        {"project_no": "S026-0021", "product_type": "STK4", "qty": 2, "due_date": date(2026, 5, 13), "priority": 2, "include_fims": True, "include_robot": True},
        {"project_no": "PJ-STK5-002", "product_type": "STK5", "qty": 3, "due_date": date(2026, 5, 20), "priority": 3, "include_fims": True, "include_robot": True},
        {"project_no": "FR301-ONLY-01", "product_type": "單獨FR301", "qty": 1, "due_date": date(2026, 5, 22), "priority": 1, "include_fims": False, "include_robot": True},
    ])

def kpi(label, value, caption=""):
    st.markdown(f'<div class="kpi-card"><div class="kpi-label">{label}</div><div class="kpi-value">{value}</div><div class="small-note">{caption}</div></div>', unsafe_allow_html=True)

def page_header(title: str, subtitle: str, tags: list[str] | None = None):
    st.markdown(f'<div class="app-title">{title}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="app-subtitle">{subtitle}</div>', unsafe_allow_html=True)
    if tags:
        st.markdown("".join([f'<span class="tag">{x}</span>' for x in tags]), unsafe_allow_html=True)
        st.write("")

def panel_open(title: str, subtitle: str = ""):
    st.markdown('<div class="panel">', unsafe_allow_html=True)
    st.markdown(f'<div class="panel-title">{title}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="panel-subtitle">{subtitle}</div>', unsafe_allow_html=True)

def panel_close():
    st.markdown('</div>', unsafe_allow_html=True)

def current_step_progress_percent(row: pd.Series, selected_date: date) -> int:
    start = row["planned_start"]; finish = row["planned_finish"]
    if selected_date <= start.date(): return 0
    if selected_date >= finish.date(): return 100
    total = (finish - start).total_seconds()
    elapsed = (datetime.combine(selected_date, time(16, 0)) - start).total_seconds()
    if total <= 0: return 100
    return int(max(0, min(100, round(elapsed / total * 100))))

def total_progress_percent(schedule_df: pd.DataFrame, project_no: str, unit_no: int, selected_date: date) -> int:
    unit_df = schedule_df[(schedule_df["project_no"] == project_no) & (schedule_df["unit_no"] == unit_no)].copy()
    if unit_df.empty: return 0
    total_hours = unit_df["hours"].sum()
    complete_hours = 0.0
    for _, r in unit_df.iterrows():
        if selected_date > r["planned_finish"].date():
            complete_hours += float(r["hours"])
        elif r["planned_start"].date() <= selected_date <= r["planned_finish"].date():
            complete_hours += float(r["hours"]) * current_step_progress_percent(r, selected_date) / 100.0
    if total_hours <= 0: return 0
    return int(round(complete_hours / total_hours * 100))

def render_station_card(row: pd.Series, selected_date: date, schedule_df: pd.DataFrame):
    step_pct = current_step_progress_percent(row, selected_date)
    total_pct = total_progress_percent(schedule_df, row["project_no"], row["unit_no"], selected_date)
    slot_code = build_slot_code(row["area"], row["slot_no"])
    chip = "生產中" if step_pct < 100 else "已完工"
    st.markdown(f"""
    <div class="station-card">
        <div class="station-top">
            <div><div class="station-top-left">工位</div><div class="station-slot">{slot_code}</div></div>
            <div class="station-chip">{chip}</div>
        </div>
        <div class="station-order">{row["machine_no"]}</div>
        <div class="station-meta">工序：{row["step_label"]}<br>預計出貨：{row["estimated_ship_date"]}</div>
        <div class="station-row"><span>目前工序進度</span><span>{step_pct}%</span></div>
        <div class="progress-wrap"><div class="progress-fill-step" style="width:{step_pct}%;"></div></div>
        <div class="station-row"><span>整機完成度</span><span>{total_pct}%</span></div>
        <div class="progress-wrap"><div class="progress-fill-total" style="width:{total_pct}%;"></div></div>
        <div class="station-foot">區域：{row["area_name"]}<br>日期：{selected_date}</div>
    </div>
    """, unsafe_allow_html=True)

def month_key(mode: str) -> str:
    return f"{mode}_calendar_month"

def render_calendar_selector(mode: str, label: str, selected_dates: Set[date]):
    st.markdown(f"**{label}**")
    if month_key(mode) not in st.session_state:
        today = date.today().replace(day=1)
        st.session_state[month_key(mode)] = today
    current_month = st.session_state[month_key(mode)]

    c1, c2, c3 = st.columns([1,2,1])
    with c1:
        if st.button("◀ 上月", key=f"{mode}_prev_month", use_container_width=True):
            y = current_month.year
            m = current_month.month - 1
            if m == 0:
                y -= 1
                m = 12
            st.session_state[month_key(mode)] = date(y, m, 1)
            current_month = st.session_state[month_key(mode)]
    with c2:
        st.markdown(f"<div style='text-align:center;font-weight:800;padding-top:0.4rem'>{current_month.year} / {current_month.month:02d}</div>", unsafe_allow_html=True)
    with c3:
        if st.button("下月 ▶", key=f"{mode}_next_month", use_container_width=True):
            y = current_month.year
            m = current_month.month + 1
            if m == 13:
                y += 1
                m = 1
            st.session_state[month_key(mode)] = date(y, m, 1)
            current_month = st.session_state[month_key(mode)]

    week_names = ["一","二","三","四","五","六","日"]
    header_cols = st.columns(7)
    for i, w in enumerate(week_names):
        header_cols[i].markdown(f"<div style='text-align:center;font-weight:700;color:#475569'>{w}</div>", unsafe_allow_html=True)

    first_weekday, days_in_month = monthrange(current_month.year, current_month.month)
    # python monthrange Monday=0
    day = 1
    started = False
    for week in range(6):
        cols = st.columns(7)
        for wd in range(7):
            if not started and wd == first_weekday:
                started = True
            if started and day <= days_in_month:
                d = date(current_month.year, current_month.month, day)
                active = d in selected_dates
                label = f"{day}"
                if active:
                    label = f"✅ {day}"
                if cols[wd].button(label, key=f"{mode}_{d.isoformat()}", use_container_width=True):
                    if d in selected_dates:
                        selected_dates.remove(d)
                    else:
                        selected_dates.add(d)
                day += 1
            else:
                cols[wd].markdown(" ")
        if day > days_in_month:
            break

    if selected_dates:
        show = sorted(selected_dates)
        st.caption("已選日期：" + ", ".join([x.isoformat() for x in show[:12]]) + (" ..." if len(show) > 12 else ""))
    else:
        st.caption("尚未選擇日期")
    return selected_dates


def run_planning():
    forecast_df = normalize_forecast(st.session_state["forecast_df"])
    routes = make_routes(get_hours_map(), get_calendar_settings()["stk_precheck_ratio"])
    scheduler = CapacityScheduler(DEFAULT_START_DATETIME, get_area_config(), get_calendar_settings())
    tasks = explode_forecast_to_tasks(forecast_df, routes)
    schedule_df = scheduler.schedule_forward(tasks)
    daily_df = scheduler.build_daily_station_view(schedule_df)
    summary_df = scheduler.build_project_summary(schedule_df)
    load_df = scheduler.build_area_load(schedule_df)

    st.session_state["forecast_df"] = forecast_df
    st.session_state["draft_schedule_df"] = schedule_df
    st.session_state["draft_daily_df"] = daily_df
    st.session_state["draft_summary_df"] = summary_df
    st.session_state["draft_load_df"] = load_df
    if st.session_state.get("official_schedule_df", pd.DataFrame()).empty:
        st.session_state["use_official_schedule"] = False


def ensure_state():
    init_db()
    if "forecast_df" not in st.session_state:
        st.session_state["forecast_df"] = seed_forecast()
    if "active_page" not in st.session_state:
        st.session_state["active_page"] = "dashboard"
    if "area_config" not in st.session_state:
        st.session_state["area_config"] = {k: v.copy() for k, v in DEFAULT_AREA_CONFIG.items()}
    if "area_config_editor_df" not in st.session_state:
        st.session_state["area_config_editor_df"] = pd.DataFrame([
            {"區域代碼": area, "區域名稱": cfg["name"], "工位數": cfg["positions"], "A組人數": cfg.get("group_a_count", 0), "A組職能": cfg.get("group_a_skill", "組裝"),
             "B組人數": cfg.get("group_b_count", 0), "B組職能": cfg.get("group_b_skill", "支援"), "人力數": effective_manpower(cfg)}
            for area, cfg in st.session_state["area_config"].items()
        ])
    if "hours_config_df" not in st.session_state:
        st.session_state["hours_config_df"] = DEFAULT_HOURS_CONFIG.copy()
    if "calendar_settings" not in st.session_state:
        st.session_state["calendar_settings"] = DEFAULT_CALENDAR_SETTINGS.copy()
    if "setup_time_config_df" not in st.session_state:
        st.session_state["setup_time_config_df"] = DEFAULT_SETUP_TIME_CONFIG.copy()
    if "quality_time_settings" not in st.session_state:
        st.session_state["quality_time_settings"] = DEFAULT_QUALITY_TIME_SETTINGS.copy()
    if "rework_log_df" not in st.session_state:
        st.session_state["rework_log_df"] = pd.DataFrame([{"日期": pd.NaT, "專案號": "", "機台編號": "", "原因": "", "返修工時": 0.0, "狀態": "待處理"}])
    if "erp_receipt_df" not in st.session_state:
        st.session_state["erp_receipt_df"] = pd.DataFrame([{"入庫日期": pd.NaT, "機台編號": "", "數量": 1, "來源": "ERP"}])
    if "holiday_dates_temp" not in st.session_state:
        st.session_state["holiday_dates_temp"] = parse_date_text(st.session_state["calendar_settings"]["holidays_text"])
    if "special_workday_dates_temp" not in st.session_state:
        st.session_state["special_workday_dates_temp"] = parse_date_text(st.session_state["calendar_settings"]["special_workdays_text"])
    if "draft_schedule_df" not in st.session_state:
        st.session_state["draft_schedule_df"] = pd.DataFrame()
    if "draft_summary_df" not in st.session_state:
        st.session_state["draft_summary_df"] = pd.DataFrame()
    if "draft_daily_df" not in st.session_state:
        st.session_state["draft_daily_df"] = pd.DataFrame()
    if "draft_load_df" not in st.session_state:
        st.session_state["draft_load_df"] = pd.DataFrame()
    official = load_official_schedule_db()
    st.session_state["official_schedule_df"] = official.copy()
    if not official.empty:
        sched = CapacityScheduler(DEFAULT_START_DATETIME, get_area_config(), get_calendar_settings())
        st.session_state["official_daily_df"] = sched.build_daily_station_view(official)
        st.session_state["official_summary_df"] = sched.build_project_summary(official)
        st.session_state["official_load_df"] = sched.build_area_load(official)
    else:
        st.session_state["official_daily_df"] = pd.DataFrame()
        st.session_state["official_summary_df"] = pd.DataFrame()
        st.session_state["official_load_df"] = pd.DataFrame()
    if "use_official_schedule" not in st.session_state:
        st.session_state["use_official_schedule"] = False

    if "draft_schedule_df" not in st.session_state:
        st.session_state["draft_schedule_df"] = pd.DataFrame()
    if "draft_summary_df" not in st.session_state:
        st.session_state["draft_summary_df"] = pd.DataFrame()
    if "draft_daily_df" not in st.session_state:
        st.session_state["draft_daily_df"] = pd.DataFrame()
    if "draft_load_df" not in st.session_state:
        st.session_state["draft_load_df"] = pd.DataFrame()
    if "official_schedule_df" not in st.session_state:
        st.session_state["official_schedule_df"] = pd.DataFrame()
    if "official_summary_df" not in st.session_state:
        st.session_state["official_summary_df"] = pd.DataFrame()
    if "official_daily_df" not in st.session_state:
        st.session_state["official_daily_df"] = pd.DataFrame()
    if "official_load_df" not in st.session_state:
        st.session_state["official_load_df"] = pd.DataFrame()
    if "use_official_schedule" not in st.session_state:
        st.session_state["use_official_schedule"] = False



def get_current_schedule_df() -> pd.DataFrame:
    official = st.session_state.get("official_schedule_df", pd.DataFrame())
    if st.session_state.get("use_official_schedule", False) and not official.empty:
        return official
    return st.session_state.get("draft_schedule_df", pd.DataFrame())

def get_current_summary_df() -> pd.DataFrame:
    official = st.session_state.get("official_summary_df", pd.DataFrame())
    if st.session_state.get("use_official_schedule", False) and not official.empty:
        return official
    return st.session_state.get("draft_summary_df", pd.DataFrame())

def get_current_daily_df() -> pd.DataFrame:
    official = st.session_state.get("official_daily_df", pd.DataFrame())
    if st.session_state.get("use_official_schedule", False) and not official.empty:
        return official
    return st.session_state.get("draft_daily_df", pd.DataFrame())

def get_current_load_df() -> pd.DataFrame:
    official = st.session_state.get("official_load_df", pd.DataFrame())
    if st.session_state.get("use_official_schedule", False) and not official.empty:
        return official
    return st.session_state.get("draft_load_df", pd.DataFrame())

def has_result() -> bool:
    current = get_current_schedule_df()
    return isinstance(current, pd.DataFrame) and not current.empty


def publish_current_schedule():
    draft = st.session_state.get("draft_schedule_df", pd.DataFrame())
    if not isinstance(draft, pd.DataFrame) or draft.empty:
        raise ValueError("目前沒有可發布的排程版本")
    save_official_schedule_db(draft)
    official = load_official_schedule_db()
    st.session_state["official_schedule_df"] = official.copy()
    sched = CapacityScheduler(DEFAULT_START_DATETIME, get_area_config(), get_calendar_settings())
    st.session_state["official_daily_df"] = sched.build_daily_station_view(official)
    st.session_state["official_summary_df"] = sched.build_project_summary(official)
    st.session_state["official_load_df"] = sched.build_area_load(official)
    st.session_state["use_official_schedule"] = True



def nav_button(label: str, key: str, icon: str):
    active = st.session_state.get("active_page") == key
    if st.button(f"{icon}  {label}", key=f"nav_{key}", use_container_width=True, type="primary" if active else "secondary"):
        st.session_state["active_page"] = key

def sidebar():
    with st.sidebar:
        st.markdown('<div class="sidebar-brand"><div class="sidebar-brand-title">🏭 製造排程系統</div><div class="sidebar-brand-sub">Production Planning Suite</div></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
        with st.expander("📊 總覽中心", expanded=True):
            nav_button("營運儀表板", "dashboard", "📊")
        with st.expander("🧭 排程中心", expanded=False):
            nav_button("Forecast輸入", "forecast", "📝")
            nav_button("排程清單", "schedule_list", "📋")
            nav_button("工位看板", "station_board", "🏗️")
            nav_button("工位圖", "station_cards", "🪪")
        with st.expander("🛠️ 裝機進度", expanded=False):
            nav_button("機台主檔", "machine_master", "🧾")
            nav_button("產線日報", "production_report", "📝")
            nav_button("日報清單", "production_report_list", "📚")
            nav_button("實際工位圖", "actual_station_board", "📍")
        with st.expander("📈 報表中心", expanded=False):
            nav_button("產能負荷報表", "capacity_report", "📈")
            nav_button("出貨達交報表", "delivery_report", "🚚")
            nav_button("經營指標報表", "management_report", "📌")
            nav_button("週入庫追蹤報表", "weekly_receipt_report", "📦")
            nav_button("工序甘特總覽", "gantt", "🗂️")
            nav_button("排程審核發布", "review_publish", "✅")
        with st.expander("⚙️ 系統設定", expanded=False):
            nav_button("主檔設定", "settings_master", "⚙️")
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
        st.caption("版本：V9.1 依賴修正版")

def render_dashboard():
    page_header("營運儀表板", "依優先級、交期、人力與工時排程後的整體概況。", ["總覽", "KPI"])
    if not has_result():
        st.info("請先到「Forecast輸入」完成排程。")
        return
    summary_df = get_current_summary_df(); load_df = get_current_load_df()
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: kpi("總機台數", int(summary_df.shape[0]), "本次排程總台數")
    with c2: kpi("準時出貨台數", int(summary_df["can_meet_due"].sum()), "符合交期")
    with c3: kpi("延遲台數", int((~summary_df["can_meet_due"]).sum()), "需關注案件")
    with c4: kpi("最晚出貨日", str(summary_df["estimated_ship_date"].max()), "最後完工日")
    with c5: kpi("平均負荷率", f'{round(load_df["utilization_pct"].mean() * 100, 1)}%', "各區平均")
    panel_open("專案彙總")
    st.dataframe(zh(summary_df), use_container_width=True, height=420, hide_index=True)
    panel_close()


def render_forecast():
    page_header("Forecast輸入", "保留原本所有功能，並增加是否含 FIMS / Robot 選項；新 forecast 會依已發布正式排程自動續排。", ["排程", "輸入", "續排"])
    left, right = st.columns([1.45, 1])

    with left:
        panel_open("Forecast資料維護", "可建立 STK4 / STK5 專案，也可選擇單獨出貨 FIMS 或 Robot。正式排程發布後，下一批 Forecast 會接著排。")
        display_df = st.session_state["forecast_df"].rename(columns={
            "project_no": "專案號", "product_type": "產品", "qty": "數量", "due_date": "出貨日", "priority": "優先級", "include_fims": "是否含FIMS", "include_robot": "是否含Robot",
        })
        if "是否含FIMS" not in display_df.columns:
            display_df["是否含FIMS"] = display_df["產品"].isin(["STK4", "STK5"])
        if "是否含Robot" not in display_df.columns:
            display_df["是否含Robot"] = display_df["產品"].isin(["STK4", "STK5"])
        edited = st.data_editor(
            display_df, num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={
                "專案號": st.column_config.TextColumn("專案號"),
                "產品": st.column_config.SelectboxColumn("產品", options=PRODUCT_OPTIONS),
                "數量": st.column_config.NumberColumn("數量", min_value=1, step=1),
                "出貨日": st.column_config.DateColumn("出貨日"),
                "優先級": st.column_config.NumberColumn("優先級", min_value=1, step=1),
                "是否含FIMS": st.column_config.CheckboxColumn("是否含FIMS"),
                "是否含Robot": st.column_config.CheckboxColumn("是否含Robot"),
            },
            key="forecast_editor_v14",
        )
        st.session_state["forecast_df"] = edited.rename(columns={
            "專案號": "project_no", "產品": "product_type", "數量": "qty", "出貨日": "due_date", "優先級": "priority",
            "是否含FIMS": "include_fims", "是否含Robot": "include_robot",
        })

        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("開始排程", type="primary", use_container_width=True):
                try:
                    run_planning()
                    st.session_state["use_official_schedule"] = False
                    st.success("草稿排程已完成，並已依正式排程占用狀況續排。")
                except Exception as e:
                    st.error(f"排程失敗：{e}")
        with c2:
            if st.button("載入示範資料", use_container_width=True):
                st.session_state["forecast_df"] = seed_forecast()
                st.success("已載入示範 Forecast。")
        with c3:
            sample = st.session_state["forecast_df"].rename(columns={
                "project_no": "專案號", "product_type": "產品", "qty": "數量", "due_date": "出貨日", "priority": "優先級",
                "include_fims": "是否含FIMS", "include_robot": "是否含Robot",
            }).to_csv(index=False).encode("utf-8-sig")
            st.download_button("下載範例CSV", data=sample, file_name="forecast_sample.csv", mime="text/csv", use_container_width=True)
        panel_close()

    with right:
        panel_open("排程前檢查")
        st.dataframe(display_df, use_container_width=True, height=340, hide_index=True)
        st.write(f"已發布正式版本數：{len(list_official_versions())}")
        official_rows = st.session_state.get("official_schedule_df", pd.DataFrame())
        st.write(f"正式排程累計工序筆數：{len(official_rows)}")
        panel_close()

def render_schedule_list():
    page_header("排程清單", "依優先級、交期、工時與產線人力排出的詳細工序清單。", ["排程", "清單"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    df = get_current_schedule_df().copy()
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1:
        areas = st.multiselect("區域", options=sorted(df["area"].unique()), default=sorted(df["area"].unique()))
    with c2:
        products = st.multiselect("產品", options=sorted(df["product_type"].unique()), default=sorted(df["product_type"].unique()))
    with c3:
        selected_date = st.date_input("日期", value=df["planned_start"].min().date(), key="schedule_list_date")
    with c4:
        keyword = st.text_input("搜尋機台/專案", placeholder="例如 S026-0021 或 PJ-STK5-002")
    filtered = df[
        df["area"].isin(areas) &
        df["product_type"].isin(products) &
        (df["planned_start"].dt.date <= selected_date) &
        (df["planned_finish"].dt.date >= selected_date)
    ]
    if keyword:
        filtered = filtered[
            filtered["project_no"].str.contains(keyword, case=False, na=False) |
            filtered["machine_no"].str.contains(keyword, case=False, na=False)
        ]
    filtered = filtered.copy()
    filtered["slot_no"] = filtered.apply(lambda r: build_slot_code(r["area"], r["slot_no"]), axis=1)
    panel_open("詳細排程資訊")
    st.dataframe(zh(filtered[["priority", "area_name", "slot_no", "machine_no", "product_type", "step_label", "planned_start", "planned_finish", "estimated_ship_date"]]), use_container_width=True, height=560, hide_index=True)
    panel_close()

def render_station_board():
    page_header("工位看板", "查看指定日期各工位正在裝配哪一台機台與哪一道工序。", ["工位", "矩陣"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    daily_df = get_current_daily_df().copy()
    c1, c2 = st.columns([1, 2])
    with c1:
        selected_date = st.date_input("查看日期", value=daily_df["date"].min(), key="station_board_date")
    with c2:
        area_filter = st.multiselect("區域篩選", options=sorted(daily_df["area"].unique()), default=sorted(daily_df["area"].unique()), key="station_board_area")
    filtered = daily_df[(daily_df["date"] == selected_date) & (daily_df["area"].isin(area_filter))].copy()
    if filtered.empty:
        st.warning("該日沒有生產安排。")
        return
    filtered["工位代號"] = filtered.apply(lambda r: build_slot_code(r["area"], r["slot_no"]), axis=1)
    panel_open("工位矩陣")
    pivot = filtered.pivot_table(index=["area_name"], columns="工位代號", values="display", aggfunc="first").reset_index().rename(columns={"area_name": "區域名稱"})
    st.dataframe(pivot, use_container_width=True, height=360, hide_index=True)
    panel_close()

def render_station_cards():
    page_header("工位圖", "以卡片方式呈現指定日期各工位正在裝配的機台，資料與排程清單連動。", ["工位", "卡片"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    schedule_df = get_current_schedule_df().copy()
    c1, c2, c3 = st.columns([1, 1.3, 1.2])
    with c1:
        selected_date = st.date_input("查看日期", value=schedule_df["planned_start"].min().date(), key="station_cards_date")
    with c2:
        area_filter = st.multiselect("區域", options=sorted(schedule_df["area"].unique()), default=sorted(schedule_df["area"].unique()), key="station_cards_area")
    with c3:
        max_cards = st.selectbox("每列卡片數", [3, 4, 5], index=1)
    active_rows = schedule_df[
        (schedule_df["area"].isin(area_filter)) &
        (schedule_df["planned_start"].dt.date <= selected_date) &
        (schedule_df["planned_finish"].dt.date >= selected_date)
    ].copy()
    panel_open("工位卡片")
    if active_rows.empty:
        st.markdown('<div class="empty-card">該日期目前沒有工位在製資料。</div>', unsafe_allow_html=True)
        panel_close()
        return
    active_rows = active_rows.sort_values(["priority", "area", "slot_no", "planned_start"])
    groups = [active_rows.iloc[i:i + max_cards] for i in range(0, len(active_rows), max_cards)]
    for batch in groups:
        cols = st.columns(max_cards)
        for idx in range(max_cards):
            with cols[idx]:
                if idx < len(batch):
                    render_station_card(batch.iloc[idx], selected_date, schedule_df)
    panel_close()





def render_machine_master():
    page_header("機台主檔", "依機種顯示不同日期欄位。STK 與模組機種使用不同輸入表單。", ["主檔", "機台"])
    mm = load_machine_master()
    if mm.empty:
        mm = pd.DataFrame(columns=["machine_no", "product_type", "station_code", "start_date", "robot_due_date", "fims_due_date", "crate_due_date", "module_due_date", "created_at", "updated_at"])

    panel_open("機台主檔維護", "新增時輸入產品編號；修改既有資料時直接顯示產品編號避免混淆。")
    machine_list = ["新增機台"] + sorted(mm["machine_no"].dropna().astype(str).unique().tolist())
    selected_machine = st.selectbox("選擇要新增或修改的產品編號", options=machine_list, key="machine_master_select_v26")

    current_row = {}
    if selected_machine != "新增機台":
        row_df = mm[mm["machine_no"].astype(str) == selected_machine]
        if not row_df.empty:
            current_row = row_df.iloc[-1].to_dict()

    default_machine_no = "" if selected_machine == "新增機台" else str(current_row.get("machine_no", ""))
    default_product_type = str(current_row.get("product_type", "STK4") or "STK4")
    default_station_code = str(current_row.get("station_code", "") or "")

    product_type = st.selectbox(
        "機種",
        options=["STK4", "STK5", "FR301", "FV501", "FIMS-4前", "FIMS-4後", "FIMS-5前", "FIMS-5後"],
        index=["STK4", "STK5", "FR301", "FV501", "FIMS-4前", "FIMS-4後", "FIMS-5前", "FIMS-5後"].index(default_product_type) if default_product_type in ["STK4", "STK5", "FR301", "FV501", "FIMS-4前", "FIMS-4後", "FIMS-5前", "FIMS-5後"] else 0,
        key="machine_master_product_v26"
    )
    if selected_machine == "新增機台":
        machine_no = st.text_input("產品編號", value=default_machine_no, key="machine_master_no_v26")
    else:
        machine_no = default_machine_no
        st.markdown(f"**產品編號：** {machine_no}")
    station_options = get_station_options()
    station_choice_options = [""] + station_options
    station_code = st.selectbox(
        "工位資訊",
        options=station_choice_options,
        index=station_choice_options.index(default_station_code) if default_station_code in station_choice_options else 0,
        key="machine_master_station_v26"
    )

    def dt_or_none(v):
        x = pd.to_datetime(v, errors="coerce")
        return None if pd.isna(x) else x.date()

    start_date_default = dt_or_none(current_row.get("start_date"))
    robot_due_default = dt_or_none(current_row.get("robot_due_date"))
    fims_due_default = dt_or_none(current_row.get("fims_due_date"))
    crate_due_default = dt_or_none(current_row.get("crate_due_date"))
    module_due_default = dt_or_none(current_row.get("module_due_date"))

    start_date = st.date_input("開工日", value=start_date_default or None, key="machine_master_start_v26")

    robot_due_date = None
    fims_due_date = None
    crate_due_date = None
    module_due_date = None

    if product_type in ["STK4", "STK5"]:
        c1, c2, c3 = st.columns(3)
        with c1:
            robot_due_date = st.date_input("robot預計交付日", value=robot_due_default or None, key="machine_master_robot_due_v26")
        with c2:
            fims_due_date = st.date_input("FIMS預計交付日", value=fims_due_default or None, key="machine_master_fims_due_v26")
        with c3:
            crate_due_date = st.date_input("預計釘箱日", value=crate_due_default or None, key="machine_master_crate_due_v26")
        st.caption("目前機種為 STK4 / STK5，因此顯示：開工日 / robot預計交付日 / FIMS預計交付日 / 預計釘箱日")
    else:
        module_due_date = st.date_input("預計交付日", value=module_due_default or None, key="machine_master_module_due_v26")
        st.caption("目前機種為模組機種，因此顯示：開工日 / 預計交付日")

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("新增 / 更新存檔", type="primary", use_container_width=True):
            machine_no_clean = str(machine_no).strip()
            if not machine_no_clean:
                st.error("請輸入產品編號。")
            else:
                save_df = mm.copy()
                save_df = save_df[save_df["machine_no"].astype(str) != machine_no_clean].copy()
                new_row = pd.DataFrame([{
                    "machine_no": machine_no_clean,
                    "product_type": product_type,
                    "station_code": str(station_code).strip(),
                    "start_date": start_date,
                    "robot_due_date": robot_due_date if product_type in ["STK4", "STK5"] else None,
                    "fims_due_date": fims_due_date if product_type in ["STK4", "STK5"] else None,
                    "crate_due_date": crate_due_date if product_type in ["STK4", "STK5"] else None,
                    "module_due_date": module_due_date if product_type not in ["STK4", "STK5"] else None,
                    "created_at": current_row.get("created_at", None),
                    "updated_at": None,
                }])
                save_df = pd.concat([save_df, new_row], ignore_index=True)
                save_machine_master(save_df)
                st.success("機台主檔已儲存。")
                st.rerun()
    with c2:
        if st.button("清空輸入", use_container_width=True):
            st.session_state["machine_master_select_v26"] = "新增機台"
            st.rerun()
    with c3:
        if selected_machine != "新增機台" and st.button("刪除此機台", use_container_width=True):
            save_df = mm[mm["machine_no"].astype(str) != selected_machine].copy()
            save_machine_master(save_df)
            st.success("機台主檔已刪除。")
            st.rerun()
    panel_close()

    panel_open("機台主檔清單")
    display = mm.copy()
    if not display.empty:
        stk_mask = display["product_type"].isin(["STK4", "STK5"])
        module_mask = ~stk_mask
        display.loc[module_mask, ["robot_due_date", "fims_due_date", "crate_due_date"]] = None
        display.loc[stk_mask, ["module_due_date"]] = None
        display = display.rename(columns={
            "machine_no": "產品編號",
            "product_type": "機種",
            "station_code": "工位資訊",
            "start_date": "開工日",
            "robot_due_date": "robot預計交付日",
            "fims_due_date": "FIMS預計交付日",
            "crate_due_date": "預計釘箱日",
            "module_due_date": "預計交付日",
        })
        st.dataframe(
            display[["產品編號", "機種", "工位資訊", "開工日", "robot預計交付日", "FIMS預計交付日", "預計釘箱日", "預計交付日"]],
            use_container_width=True,
            hide_index=True,
            height=320
        )
    else:
        st.info("目前尚無機台主檔資料。")
    panel_close()







def render_production_report():
    page_header("產線日報", "提供產線每日回報實際進度。這版先支援 STK4 機種，工序完成度採完成勾選與數字輸入同步控制。", ["日報", "STK4"])
    machine_options = list_stk4_machine_options()
    if not machine_options:
        st.info("目前尚無 STK4 機台可回報。請先到「機台主檔」建立機台編號。")
        return

    panel_open("每日進度回報", "欄位包含機台編號、日期、異常類別、異常原因、當天進度說明；工序完成度以完成勾選與數字輸入同步。")
    top1, top2 = st.columns(2)
    with top1:
        machine_no = st.selectbox("機台編號", options=machine_options, key="report_machine_no")
    with top2:
        report_date = st.date_input("日期", value=date.today(), key="report_date")

    default_form = get_default_report_form(machine_no, report_date)
    process_master = get_stk4_process_master().copy()
    process_master["分類"] = process_master["分類"].replace({"警幕": "螢幕"})

    default_abnormals = default_form.get("abnormals", [])
    if not isinstance(default_abnormals, list) or len(default_abnormals) == 0:
        default_abnormals = [{"異常類別": "正常", "異常原因": ""}]
    abnormal_df = pd.DataFrame(default_abnormals)

    st.markdown("#### 異常回報")
    abnormal_edit = st.data_editor(
        abnormal_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "異常類別": st.column_config.SelectboxColumn("異常類別", options=REPORT_ABNORMAL_OPTIONS),
            "異常原因": st.column_config.TextColumn("異常原因"),
        },
        key=f"abnormal_edit_{machine_no}_{report_date.isoformat()}"
    )

    daily_note = st.text_area("當天進度說明", value=str(default_form.get("daily_note", "")), height=120, key="report_daily_note")

    st.markdown("#### STK4 工序進度")
    sync_key = f"stk4_sync_rows_{machine_no}_{report_date.isoformat()}"
    if sync_key not in st.session_state:
        base = process_master.copy()
        base["完成"] = False
        base["完成%"] = 0
        if default_form.get("checklist"):
            loaded = pd.DataFrame(default_form["checklist"])
            if not loaded.empty:
                base = loaded.copy()
                if "完成" not in base.columns:
                    base["完成"] = False
                if "完成%" not in base.columns:
                    base["完成%"] = base["完成"].apply(lambda x: 100 if bool(x) else 0)
        base["分類"] = base["分類"].replace({"警幕": "螢幕"})
        base["完成"] = base["完成"].fillna(False).astype(bool)
        base["完成%"] = pd.to_numeric(base["完成%"], errors="coerce").fillna(0).clip(lower=0, upper=100).round(0).astype(int)
        st.session_state[sync_key] = base[["分類", "工序編號", "完成", "完成%"]].to_dict(orient="records")

    rows = st.session_state[sync_key]

    def on_done_change(done_key: str, pct_key: str):
        if st.session_state.get(done_key, False):
            st.session_state[pct_key] = 100
        else:
            if int(st.session_state.get(pct_key, 0)) == 100:
                st.session_state[pct_key] = 0

    def on_pct_change(done_key: str, pct_key: str):
        pct_val = int(st.session_state.get(pct_key, 0))
        st.session_state[done_key] = pct_val >= 100

    hdr = st.columns([1.0, 2.0, 2.0, 1.2])
    hdr[0].markdown('<div class="process-header-cell">完成</div>', unsafe_allow_html=True)
    hdr[1].markdown('<div class="process-header-cell">分類</div>', unsafe_allow_html=True)
    hdr[2].markdown('<div class="process-header-cell">工序編號</div>', unsafe_allow_html=True)
    hdr[3].markdown('<div class="process-header-cell">完成%</div>', unsafe_allow_html=True)

    updated_rows = []
    for i, row in enumerate(rows):
        done_key = f"{sync_key}_done_{i}"
        pct_key = f"{sync_key}_pct_{i}"

        if done_key not in st.session_state:
            st.session_state[done_key] = bool(row.get("完成", False))
        if pct_key not in st.session_state:
            st.session_state[pct_key] = int(row.get("完成%", 0) or 0)

        c1, c2, c3, c4 = st.columns([1.0, 2.0, 2.0, 1.2])
        with c1:
            st.checkbox("完成", key=done_key, on_change=on_done_change, args=(done_key, pct_key), label_visibility="collapsed")
        with c2:
            st.markdown(f'<div class="process-body-cell">{row["分類"]}</div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="process-body-cell">{row["工序編號"]}</div>', unsafe_allow_html=True)
        with c4:
            st.number_input("完成%", min_value=0, max_value=100, step=1, key=pct_key, label_visibility="collapsed", on_change=on_pct_change, args=(done_key, pct_key))

        updated_rows.append({
            "分類": row["分類"],
            "工序編號": row["工序編號"],
            "完成": bool(st.session_state[done_key]),
            "完成%": int(st.session_state[pct_key]),
        })
        st.markdown('<div class="process-row-divider"></div>', unsafe_allow_html=True)

    checklist_edit = pd.DataFrame(updated_rows)
    st.session_state[sync_key] = checklist_edit.to_dict(orient="records")

    checklist_for_calc = checklist_edit.merge(
        process_master[["分類", "工序編號", "標準工時(時)"]],
        on=["分類", "工序編號"],
        how="left"
    )
    progress_pct, completed_std_hours, total_std_hours = calc_stk4_progress(checklist_for_calc)

    current_rows = checklist_edit[checklist_edit["完成%"] < 100]
    if current_rows.empty:
        current_process_code = str(checklist_edit.iloc[-1]["工序編號"])
        current_process_label = str(checklist_edit.iloc[-1]["分類"])
        current_process_pct = 100
    else:
        current_process_code = str(current_rows.iloc[0]["工序編號"])
        current_process_label = str(current_rows.iloc[0]["分類"])
        current_process_pct = int(current_rows.iloc[0]["完成%"])

    abnormal_types = [str(x) for x in abnormal_edit["異常類別"].dropna().tolist()] if not abnormal_edit.empty else ["正常"]
    overall_abnormal_type = "正常" if all(x == "正常" for x in abnormal_types) else "異常"

    k1, k2, k3 = st.columns(3)
    with k1:
        kpi("整機完成度", f"{int(progress_pct)}%", "以 STK4 標準工時回推")
    with k2:
        kpi("已完成工時", completed_std_hours, "標準工時")
    with k3:
        kpi("總標準工時", total_std_hours, "STK4 全工序")

    if st.button("送出日報並同步實際工位圖", type="primary", use_container_width=True):
        payload = {
            "machine_no": machine_no,
            "report_date": report_date.isoformat(),
            "product_type": "STK4",
            "current_process_code": current_process_code,
            "current_process_label": current_process_label,
            "current_process_pct": int(current_process_pct),
            "abnormal_type": overall_abnormal_type,
            "current_status": "作業" if int(progress_pct) < 100 else "完成",
            "abnormal_reason": "；".join([str(x) for x in abnormal_edit["異常原因"].fillna("").tolist() if str(x).strip() != ""]),
            "abnormals_json": json.dumps(abnormal_edit.to_dict(orient="records"), ensure_ascii=False),
            "daily_note": daily_note,
            "checklist_json": json.dumps(checklist_edit.to_dict(orient="records"), ensure_ascii=False),
            "progress_pct": float(progress_pct),
            "completed_std_hours": float(completed_std_hours),
            "total_std_hours": float(total_std_hours),
            "created_at": datetime.now().isoformat(timespec="seconds"),
        }
        save_production_report(payload)
        st.success("已儲存產線日報，實際工位圖與日報清單會同步更新。")
        st.rerun()
    panel_close()

    panel_open("近期日報紀錄")
    reports = load_latest_production_reports()
    if reports.empty:
        st.info("目前尚無回報資料。")
    else:
        show_cols = [c for c in ["machine_no", "report_date", "current_process_code", "current_process_label", "current_process_pct", "progress_pct", "abnormal_type", "abnormal_reason", "daily_note"] if c in reports.columns]
        st.dataframe(reports[show_cols], use_container_width=True, height=260, hide_index=True)
    panel_close()




def render_production_report_list():
    page_header("日報清單", "可依日期區間、產品編號、機種篩選日報內容。", ["日報", "查詢"])
    df = build_production_report_list_df()
    if df.empty:
        st.info("目前尚無日報資料。")
        return

    df["report_date"] = pd.to_datetime(df["report_date"], errors="coerce")
    valid_dates = df["report_date"].dropna()
    min_date = valid_dates.dt.date.min() if not valid_dates.empty else date.today()
    max_date = valid_dates.dt.date.max() if not valid_dates.empty else date.today()

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        start_date = st.date_input("起始日期", value=min_date, key="report_list_start_date")
    with c2:
        end_date = st.date_input("結束日期", value=max_date, key="report_list_end_date")
    with c3:
        machine_options = sorted(df["machine_no"].dropna().astype(str).unique().tolist())
        selected_machine = st.selectbox("產品編號", options=["全部"] + machine_options, key="report_list_machine")
    with c4:
        prod_series = df["machine_product_type"] if "machine_product_type" in df.columns else df["product_type"]
        prod_options = sorted(pd.Series(prod_series).dropna().astype(str).unique().tolist())
        selected_product = st.selectbox("機種", options=["全部"] + prod_options, key="report_list_product")

    filtered = df.copy()
    filtered = filtered[
        (pd.to_datetime(filtered["report_date"], errors="coerce").dt.date >= start_date) &
        (pd.to_datetime(filtered["report_date"], errors="coerce").dt.date <= end_date)
    ]
    if selected_machine != "全部":
        filtered = filtered[filtered["machine_no"].astype(str) == str(selected_machine)]
    if selected_product != "全部":
        prod_series = filtered["machine_product_type"] if "machine_product_type" in filtered.columns else filtered["product_type"]
        filtered = filtered[pd.Series(prod_series).astype(str) == str(selected_product)]

    show = filtered.copy()
    show["report_date"] = pd.to_datetime(show["report_date"], errors="coerce").dt.date
    if "crate_due_date" in show.columns:
        show["crate_due_date"] = pd.to_datetime(show["crate_due_date"], errors="coerce").dt.date
    show = show.rename(columns={
        "machine_no": "產品編號",
        "report_date": "日期",
        "abnormal_items": "異常項目",
        "abnormal_reasons_joined": "異常原因",
        "daily_note": "當天進度說明",
        "progress_pct": "整機完成度",
        "crate_due_date": "預計釘箱日",
        "machine_product_type": "機種",
    })
    cols = [c for c in ["產品編號", "日期", "機種", "異常項目", "異常原因", "當天進度說明", "整機完成度", "預計釘箱日"] if c in show.columns]
    st.dataframe(show[cols], use_container_width=True, height=420, hide_index=True)





def build_line_layout_df_static() -> pd.DataFrame:
    df = build_actual_station_cards_df().copy()
    if df.empty:
        return pd.DataFrame(columns=["layout_group", "layout_label", "station_code", "machine_no", "status_code", "status_color", "sort_order"])

    for col in ["station_code", "machine_no", "status_code"]:
        if col not in df.columns:
            df[col] = ""

    rt = load_realtime_status_df()
    rt_map = {}
    if not rt.empty:
        rt_map = {str(r.get("machine_no", "")): str(r.get("status_code", "") or "") for _, r in rt.iterrows()}

    mapping_rows = []
    for i in range(1, 11):
        mapping_rows.append({"layout_group": "STK4", "layout_label": f"STK4-{i}", "station_code": f"D{i:02d}", "sort_order": i})
    for i in range(1, 6):
        mapping_rows.append({"layout_group": "STK5", "layout_label": f"STK5-{i}", "station_code": f"E{i:02d}", "sort_order": i})

    map_df = pd.DataFrame(mapping_rows)
    base = map_df.merge(df[["station_code", "machine_no", "status_code"]], on="station_code", how="left")
    base["machine_no"] = base["machine_no"].fillna("")

    # 以即時狀態資料表為主，沒有時才退回工位卡資料
    base["status_code"] = base.apply(
        lambda r: rt_map.get(str(r.get("machine_no", "")), str(r.get("status_code", "") or "")),
        axis=1
    )
    base["status_code"] = base["status_code"].replace("", "未設定").fillna("未設定")
    base["status_color"] = base["status_code"].apply(status_color_hex)
    return base



def render_static_line_layout():
    layout_df = build_line_layout_df_static()
    if layout_df.empty:
        st.info("目前沒有可顯示的產線工位資料。")
        return

    st.markdown("#### STK4 實際產線工位")
    stk4 = layout_df[layout_df["layout_group"] == "STK4"].sort_values("sort_order").reset_index(drop=True)
    cols = st.columns(5)
    for idx, (_, row) in enumerate(stk4.iterrows()):
        with cols[idx % 5]:
            machine_display = row["machine_no"] if str(row["machine_no"]).strip() else "-"
            status_display = row["status_code"] if str(row["status_code"]).strip() else "未設定"
            st.markdown(f"""
            <div class="layout-grid-cell">
                <div class="layout-grid-title">{row["layout_label"]}</div>
                <div class="layout-grid-product">產品：{machine_display}</div>
                <div class="layout-grid-status" style="color:{row["status_color"]}">● {status_display}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    st.markdown("#### STK5 實際產線工位")
    stk5 = layout_df[layout_df["layout_group"] == "STK5"].sort_values("sort_order").reset_index(drop=True)
    cols = st.columns(5)
    for idx, (_, row) in enumerate(stk5.iterrows()):
        with cols[idx % 5]:
            machine_display = row["machine_no"] if str(row["machine_no"]).strip() else "-"
            status_display = row["status_code"] if str(row["status_code"]).strip() else "未設定"
            st.markdown(f"""
            <div class="layout-grid-cell">
                <div class="layout-grid-title">{row["layout_label"]}</div>
                <div class="layout-grid-product">產品：{machine_display}</div>
                <div class="layout-grid-status" style="color:{row["status_color"]}">● {status_display}</div>
            </div>
            """, unsafe_allow_html=True)


def render_actual_station_board():
    page_header("實際工位圖", "可查看工位卡，也可查看對應實際產線工位畫面。", ["實績", "工位"])
    actual_df = build_actual_station_cards_df()
    if actual_df.empty:
        st.info("目前沒有可顯示的工位卡資料。請先在機台主檔設定工位資訊，並建立產線日報。")
        return

    c1, c2, c3 = st.columns(3)
    with c1:
        station_query = st.text_input("搜尋工位代碼", value="", key="actual_station_query")
    with c2:
        machine_query = st.text_input("搜尋產品編號", value="", key="actual_machine_query")
    with c3:
        product_options = ["全部"] + sorted(actual_df["product_type"].dropna().astype(str).unique().tolist())
        product_filter = st.selectbox("機種", options=product_options, key="actual_product_filter")

    view = actual_df.copy()
    if station_query.strip():
        view = view[view["station_code"].astype(str).str.contains(station_query.strip(), case=False, na=False)]
    if machine_query.strip():
        view = view[view["machine_no"].astype(str).str.contains(machine_query.strip(), case=False, na=False)]
    if product_filter != "全部":
        view = view[view["product_type"].astype(str) == product_filter]

    tab1, tab2 = st.tabs(["工位卡畫面", "實際產線工位畫面"])

    with tab1:
        if view.empty:
            st.warning("查無符合條件的工位卡。")
        else:
            view = view.sort_values(["station_code", "machine_no"]).reset_index(drop=True)

            def fmt_date(v):
                x = pd.to_datetime(v, errors="coerce")
                return "-" if pd.isna(x) else str(x.date())

            def pct_num(v):
                try:
                    return max(0, min(100, int(round(float(v)))))
                except Exception:
                    return 0

            def bar_html(label, pct, color, due_text=None):
                due_html = f"<div class='actual-card-submeta'>{due_text}</div>" if due_text else ""
                return f"""
                {due_html}
                <div class='actual-card-label'>{label}</div>
                <div class='actual-progress-pct'>{pct}%</div>
                <div class='actual-track'>
                    <div class='actual-fill' style='width:{pct}%;background:{color}'></div>
                </div>
                """

            panel_open("實際工位卡片", "一張卡代表一個工位，完整顯示該工位機台與進度資訊。")
            max_cards = 2
            groups = [view.iloc[i:i + max_cards] for i in range(0, len(view), max_cards)]
            for batch in groups:
                cols = st.columns(max_cards)
                for idx, (_, row) in enumerate(batch.iterrows()):
                    with cols[idx]:
                        crate_date = fmt_date(row.get("crate_due_date"))
                        fims_due = fmt_date(row.get("fims_due_date"))
                        robot_due = fmt_date(row.get("robot_due_date"))
                        whole_pct = pct_num(row.get("whole_progress", 0))
                        fims_pct = pct_num(row.get("fims_progress", 0))
                        robot_pct = pct_num(row.get("robot_progress", 0))
                        db_status = str(row.get("status_code", "正常生產") or "正常生產")
                        widget_key = f'main_realtime_status_{row["machine_no"]}'
                        if widget_key not in st.session_state:
                            st.session_state[widget_key] = db_status
                        selected = st.selectbox(
                            "當下狀態",
                            options=REALTIME_STATUS_OPTIONS,
                            index=REALTIME_STATUS_OPTIONS.index(st.session_state.get(widget_key, db_status)) if st.session_state.get(widget_key, db_status) in REALTIME_STATUS_OPTIONS else 0,
                            key=widget_key
                        )
                        current_status = selected if selected else db_status

                        if st.button("更新狀態", key=f'main_update_status_{row["machine_no"]}', use_container_width=True):
                            save_machine_realtime_status(str(row["machine_no"]), current_status)
                            row["status_code"] = current_status
                            st.success(f'已更新 {row["machine_no"]} 狀態為：{current_status}')
                            st.rerun()

                        card_html = f"""
                        <div class='actual-full-card' style='border-color:{status_color_hex(current_status)}'>
                            <div class='actual-card-station'>{row["station_code"]}</div>
                            <div class='actual-card-machine'>{row["machine_no"]}</div>
                            <div class='actual-card-meta'>機種：{row["product_type"]}</div>
                            <div class='actual-card-meta'>預計釘箱日：{crate_date}</div>
                            <div style='text-align:center'>
                                <span class='status-badge' style='background:{status_color_hex(current_status)}'>{status_emoji(current_status)} {current_status}</span>
                            </div>
                            {bar_html("整機裝機進度", whole_pct, "#2563eb")}
                            {bar_html("FIMS進度", fims_pct, "#22c55e", f"FIMS預計交付日：{fims_due}")}
                            {bar_html("Robot進度", robot_pct, "#f59e0b", f"Robot預計交付日：{robot_due}")}
                        </div>
                        """
                        st.markdown(card_html, unsafe_allow_html=True)
            panel_close()

    with tab2:
        panel_open("實際產線工位畫面", "先顯示對應產線工位、產品編號與燈號，不影響原本工位卡功能。")
        render_static_line_layout()
        panel_close()

    panel_open("實際工位明細")
    show = build_actual_station_cards_df().copy()
    if station_query.strip():
        show = show[show["station_code"].astype(str).str.contains(station_query.strip(), case=False, na=False)]
    if machine_query.strip():
        show = show[show["machine_no"].astype(str).str.contains(machine_query.strip(), case=False, na=False)]
    if product_filter != "全部":
        show = show[show["product_type"].astype(str) == product_filter]
    show["crate_due_date"] = pd.to_datetime(show["crate_due_date"], errors="coerce").dt.date
    show["fims_due_date"] = pd.to_datetime(show["fims_due_date"], errors="coerce").dt.date
    show["robot_due_date"] = pd.to_datetime(show["robot_due_date"], errors="coerce").dt.date
    show = show.rename(columns={
        "station_code": "工位代碼",
        "machine_no": "產品編號",
        "product_type": "機種",
        "status_code": "當下狀態",
        "crate_due_date": "預計釘箱日",
        "fims_due_date": "FIMS預計交付日",
        "robot_due_date": "Robot預計交付日",
        "whole_progress": "整機完成度",
        "fims_progress": "FIMS進度",
        "robot_progress": "Robot進度",
    })
    cols = [c for c in ["工位代碼", "產品編號", "機種", "當下狀態", "預計釘箱日", "FIMS預計交付日", "Robot預計交付日", "整機完成度", "FIMS進度", "Robot進度"] if c in show.columns]
    st.dataframe(show[cols], use_container_width=True, height=320, hide_index=True)
    panel_close()

def render_capacity_report():
    page_header("產能負荷報表", "依工作日定義、平日加班、假日工時計算各區負荷。", ["報表", "產能"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    load_df = get_current_load_df().copy()
    show = zh(load_df)
    if "負荷率" in show.columns:
        show["負荷率"] = (load_df["utilization_pct"] * 100).round(1).astype(str) + "%"
    panel_open("產能負荷明細")
    st.dataframe(show, use_container_width=True, height=460, hide_index=True)
    panel_close()

def render_delivery_report():
    page_header("出貨達交報表", "顯示依優先級與交期排出的最終出貨結果。", ["報表", "出貨"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    summary_df = get_current_summary_df().copy()
    c1, c2, c3 = st.columns(3)
    with c1: kpi("總出貨台數", int(summary_df.shape[0]), "全部案件")
    with c2: kpi("準交率 OTD", f"{round(float(summary_df["can_meet_due"].mean())*100,1) if not summary_df.empty else 0}%", "準時出貨比例")
    with c3: kpi("延遲台數", int((~summary_df["can_meet_due"]).sum()) if not summary_df.empty else 0, "未準時案件")
    panel_open("出貨總表")
    st.dataframe(zh(summary_df), use_container_width=True, height=460, hide_index=True)
    panel_close()
    excel_bytes = to_excel_bytes(st.session_state["forecast_df"], get_current_summary_df(), get_current_schedule_df(), get_current_daily_df(), get_current_load_df())
    st.download_button("下載完整排程 Excel", data=excel_bytes, file_name="schedule_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")



def render_management_report():
    page_header("經營指標報表", "整合製程時間、等待時間、交期、產出、準交率、在製品、瓶頸與路由流程。", ["KPI", "管理"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return

    schedule_df = get_current_schedule_df().copy()
    summary_df = get_current_summary_df().copy()
    load_df = get_current_load_df().copy()
    machine_metrics = build_machine_metrics(schedule_df)
    setup_events = build_setup_events(schedule_df)
    routing_df = build_routing_flow_table()
    rework_df = get_rework_log_df().copy()

    today_default = summary_df["planned_start"].min().date() if not summary_df.empty else date.today()
    selected_date = st.date_input("WIP查看日期", value=today_default, key="management_wip_date")

    throughput_df = throughput_by_date(summary_df)
    wip_value = wip_on_date(schedule_df, selected_date)
    otd = round(float(summary_df["can_meet_due"].mean()) * 100, 1) if not summary_df.empty else 0.0
    bottleneck_area, bottleneck_util = bottleneck_summary(load_df)
    total_setup = round(float(setup_events["setup_time_hours"].sum()), 2) if not setup_events.empty else 0.0
    total_rework = round(float(pd.to_numeric(rework_df.get("返修工時", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()), 2) if not rework_df.empty else 0.0

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: kpi("準交率 OTD", f"{otd}%", "準時出貨比例")
    with c2: kpi("在製品 WIP", wip_value, f"{selected_date} 當日")
    with c3: kpi("瓶頸負荷率", f"{round(bottleneck_util*100,1)}%", bottleneck_area)
    with c4: kpi("總換線時間", f"{total_setup} hrs", "依產品切換估算")
    with c5: kpi("總返修工時", f"{total_rework} hrs", "重工 / 返修紀錄")

    left, right = st.columns([1.3, 1])
    with left:
        panel_open("製程與等待時間", "Cycle Time / Queue Time / 品檢時間 / Due Date")
        if not machine_metrics.empty:
            show = machine_metrics.rename(columns={
                "cycle_time_hours": "製程時間(hrs)",
                "queue_time_hours": "等待時間(hrs)",
                "process_time_hours": "加工工時(hrs)",
                "quality_time_hours": "品檢時間(hrs)",
            })
            st.dataframe(zh(show[["project_no", "machine_no", "product_type", "priority", "planned_start", "planned_finish", "製程時間(hrs)", "等待時間(hrs)", "加工工時(hrs)", "品檢時間(hrs)", "due_date", "estimated_ship_date", "can_meet_due", "days_late"]]), use_container_width=True, height=420, hide_index=True)
        panel_close()

    with right:
        panel_open("產出 Throughput", "依預計出貨日統計每日完成台數")
        if not throughput_df.empty:
            st.bar_chart(throughput_df.set_index("estimated_ship_date"))
        else:
            st.info("目前沒有產出資料。")
        panel_close()

    panel_open("瓶頸負荷與換線時間")
    c1, c2 = st.columns(2)
    with c1:
        if not load_df.empty:
            plot_load = load_df.copy()
            plot_load["utilization_pct"] = plot_load["utilization_pct"] * 100
            st.line_chart(plot_load.pivot(index="date", columns="area_name", values="utilization_pct"))
    with c2:
        if not setup_events.empty:
            setup_show = setup_events.copy()
            setup_show["slot_no"] = setup_show.apply(lambda r: build_slot_code(r["area"], r["slot_no"]), axis=1)
            st.dataframe(setup_show.rename(columns={
                "area": "區域代碼",
                "slot_no": "工位",
                "from_machine": "前機台",
                "to_machine": "後機台",
                "from_product": "前產品",
                "to_product": "後產品",
                "setup_time_hours": "換線時間(hrs)",
                "setup_start": "換線開始",
                "setup_end": "換線結束",
            }), use_container_width=True, height=300, hide_index=True)
        else:
            st.info("目前沒有換線事件。")
    panel_close()

    panel_open("Routing Flow / 重工返修")
    c1, c2 = st.columns(2)
    with c1:
        st.dataframe(routing_df.rename(columns={"product_type": "產品", "routing_flow": "製程順序"}), use_container_width=True, hide_index=True)
    with c2:
        if rework_df.empty:
            st.info("目前沒有重工 / 返修紀錄。")
        else:
            st.dataframe(rework_df, use_container_width=True, height=240, hide_index=True)
    panel_close()



def render_weekly_receipt_report():
    page_header("週入庫追蹤報表", "監控本周原定安排台數與 ERP 實際入庫台數的差異。", ["報表", "入庫", "ERP"])
    summary_df = get_current_summary_df().copy()
    if summary_df.empty:
        st.info("請先完成排程，系統才知道本周原定安排幾台。")
        return

    c1, c2 = st.columns([1, 1.4])
    with c1:
        ship_dates = pd.to_datetime(summary_df["estimated_ship_date"], errors="coerce").dropna()
        default_target_date = ship_dates.min().date() if not ship_dates.empty else date.today()
        target_date = st.date_input("查詢週別", value=default_target_date, key="weekly_receipt_date")
    with c2:
        st.caption("本報表用排程的預計出貨日當作『本周原定安排』，實際入庫則由 ERP 入庫資料匯入或手動維護。")

    week_start, week_end = week_bounds(target_date)
    st.write(f"統計區間：**{week_start} ~ {week_end}**")

    tab1, tab2 = st.tabs(["ERP入庫資料", "週追蹤結果"])

    with tab1:
        panel_open("ERP 入庫資料維護", "可先手動輸入或上傳 ERP 入庫清單，後續再串 ERP。")
        mode = st.segmented_control("資料來源", ["手動維護", "上傳檔案"], default="手動維護", key="erp_mode")
        if mode == "手動維護":
            erp_edit = st.data_editor(
                st.session_state["erp_receipt_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "入庫日期": st.column_config.DateColumn("入庫日期"),
                    "機台編號": st.column_config.TextColumn("機台編號"),
                    "數量": st.column_config.NumberColumn("數量", min_value=0, step=1),
                    "來源": st.column_config.TextColumn("來源"),
                },
                key="erp_receipt_editor_v13",
            )
            st.session_state["erp_receipt_df"] = erp_edit
        else:
            uploaded = st.file_uploader("上傳 ERP 入庫 Excel / CSV", type=["xlsx", "xls", "csv"], key="erp_receipt_upload")
            if uploaded is not None:
                df = pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv") else pd.read_excel(uploaded)
                rename_map = {"日期": "入庫日期", "入庫日": "入庫日期", "machine_no": "機台編號", "數量": "數量", "qty": "數量"}
                df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})
                keep = [c for c in ["入庫日期", "機台編號", "數量", "來源"] if c in df.columns]
                if "來源" not in keep:
                    df["來源"] = "ERP"
                    keep = [c for c in ["入庫日期", "機台編號", "數量", "來源"] if c in df.columns]
                st.session_state["erp_receipt_df"] = df[keep].copy()
                st.success("ERP 入庫資料已載入。")
        panel_close()

    with tab2:
        planned_df, payload = build_week_receipt_report(summary_df, st.session_state["erp_receipt_df"], target_date)
        actual_df = payload["actual_df"]
        compare_df = payload["compare_df"]
        s = payload["summary"]

        k1, k2, k3, k4 = st.columns(4)
        with k1: kpi("本周原定安排", s["planned_units"], "台")
        with k2: kpi("實際入庫", s["actual_units"], "台")
        with k3: kpi("差異台數", s["gap_units"], "計畫 - 實績")
        with k4: kpi("達成率", f'{s["achieve_rate"]}%', "實際 / 計畫")

        panel_open("本周計畫 vs 實際入庫")
        x1, x2 = st.columns(2)
        with x1:
            st.markdown("**本周原定安排機台**")
            st.dataframe(planned_df, use_container_width=True, height=280, hide_index=True)
        with x2:
            st.markdown("**本周實際入庫機台**")
            st.dataframe(actual_df, use_container_width=True, height=280, hide_index=True)
        panel_close()

        panel_open("差異追蹤")
        st.dataframe(compare_df, use_container_width=True, height=320, hide_index=True)
        panel_close()


def render_gantt():
    page_header("工序甘特總覽", "以甘特圖顯示各工序開始、完成、工位與交期。", ["報表", "甘特圖"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return

    schedule_df = get_current_schedule_df().copy()
    if schedule_df.empty:
        st.info("目前沒有可顯示的排程。")
        return

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        products = st.multiselect("產品", options=sorted(schedule_df["product_type"].unique()), default=sorted(schedule_df["product_type"].unique()), key="gantt_products")
    with c2:
        areas = st.multiselect("區域", options=sorted(schedule_df["area"].unique()), default=sorted(schedule_df["area"].unique()), key="gantt_areas")
    with c3:
        keyword = st.text_input("搜尋機台/專案", placeholder="例如 S026-0021 或 PJ-STK5-002", key="gantt_keyword")

    filtered = schedule_df[schedule_df["product_type"].isin(products) & schedule_df["area"].isin(areas)].copy()
    if keyword:
        filtered = filtered[
            filtered["project_no"].str.contains(keyword, case=False, na=False) |
            filtered["machine_no"].str.contains(keyword, case=False, na=False)
        ]

    filtered["任務名稱"] = filtered["machine_no"] + " / " + filtered["step_label"] + " / " + filtered.apply(lambda r: build_slot_code(r["area"], r["slot_no"]), axis=1)
    filtered["區域工位"] = filtered["area_name"] + " / " + filtered.apply(lambda r: build_slot_code(r["area"], r["slot_no"]), axis=1)

    panel_open("工序甘特圖", "這裡是圖形化甘特圖，可搭配下方表格人工審核。")
    if filtered.empty:
        st.warning("目前篩選條件下沒有資料。")
    else:
        fig = px.timeline(
            filtered,
            x_start="planned_start",
            x_end="planned_finish",
            y="任務名稱",
            color="area_name",
            hover_data=["project_no", "product_type", "priority", "step_label", "due_date", "estimated_ship_date", "slot_no"],
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(height=max(500, 45 * len(filtered)), margin=dict(l=20, r=20, t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

    show_df = filtered[["priority", "machine_no", "product_type", "step_label", "area_name", "slot_no", "planned_start", "planned_finish", "estimated_ship_date"]].copy()
    show_df["slot_no"] = filtered.apply(lambda r: build_slot_code(r["area"], r["slot_no"]), axis=1) if not filtered.empty else show_df.get("slot_no", [])
    st.dataframe(zh(show_df), use_container_width=True, height=420, hide_index=True)
    panel_close()


def render_review_publish():
    page_header("排程審核發布", "系統先自動排一版，人工確認與調整後，再發布成正式執行版本。", ["審核", "發布"])
    draft_df = st.session_state.get("draft_schedule_df", pd.DataFrame()).copy()
    official_df = st.session_state.get("official_schedule_df", pd.DataFrame()).copy()

    top1, top2, top3 = st.columns(3)
    with top1:
        st.metric("草稿排程筆數", len(draft_df))
    with top2:
        st.metric("正式排程筆數", len(official_df))
    published_versions = list_official_versions()
    with top3:
        st.metric("目前顯示版本", "正式版" if st.session_state.get("use_official_schedule") and not official_df.empty else "草稿版")
        st.caption(f"已發布版本數：{len(published_versions)}")

    if draft_df.empty:
        st.info("請先到 Forecast輸入 產生草稿排程。")
        return

    panel_open("人工調整草稿排程", "可直接修改工位與開始時間；系統會自動重算結束時間。發布後才是正式版本。")

    editable = draft_df[["priority", "project_no", "machine_no", "product_type", "step", "step_label", "area", "slot_no", "planned_start", "hours", "due_date"]].copy()
    editable["slot_display"] = editable.apply(lambda r: build_slot_code(r["area"], r["slot_no"]), axis=1)

    edited = st.data_editor(
        editable.rename(columns={
            "priority": "優先級",
            "project_no": "專案號",
            "machine_no": "機台編號",
            "product_type": "產品",
            "step": "工序代碼",
            "step_label": "工序",
            "area": "區域代碼",
            "slot_display": "工位",
            "planned_start": "開始時間",
            "hours": "工時",
            "due_date": "需求交期",
        }),
        use_container_width=True,
        hide_index=True,
        disabled=["優先級", "專案號", "機台編號", "產品", "工序代碼", "工序", "工時", "需求交期"],
        column_config={
            "區域代碼": st.column_config.SelectboxColumn("區域代碼", options=sorted(get_area_config().keys())),
            "工位": st.column_config.TextColumn("工位"),
            "開始時間": st.column_config.DatetimeColumn("開始時間"),
        },
        key="draft_review_editor",
    )

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("套用人工調整", type="primary", use_container_width=True):
            edited_df = edited.rename(columns={
                "優先級": "priority",
                "專案號": "project_no",
                "機台編號": "machine_no",
                "產品": "product_type",
                "工序代碼": "step",
                "工序": "step_label",
                "區域代碼": "area",
                "工位": "slot_display",
                "開始時間": "planned_start",
                "工時": "hours",
                "需求交期": "due_date",
            }).copy()

            area_cfg = get_area_config()
            for idx, row in edited_df.iterrows():
                slot_display = str(row["slot_display"]).strip().upper()
                slot_no = int(slot_display[1:]) if len(slot_display) >= 2 and slot_display[1:].isdigit() else draft_df.loc[idx, "slot_no"]
                area = str(row["area"]).strip().upper()
                start_dt = pd.to_datetime(row["planned_start"])
                finish_dt = add_work_hours_forward(start_dt, float(row["hours"]), get_calendar_settings())
                draft_df.loc[idx, "area"] = area
                draft_df.loc[idx, "area_name"] = area_cfg.get(area, {}).get("name", draft_df.loc[idx, "area_name"])
                draft_df.loc[idx, "slot_no"] = slot_no
                draft_df.loc[idx, "planned_start"] = start_dt
                draft_df.loc[idx, "planned_finish"] = finish_dt

            ship_df = (
                draft_df[draft_df["step"].isin(["STK4_QC", "STK5_QC"])]
                [["project_no", "unit_no", "planned_finish"]]
                .rename(columns={"planned_finish": "estimated_ship_datetime"})
            )
            draft_df = draft_df.drop(columns=[c for c in ["estimated_ship_datetime", "estimated_ship_date", "can_meet_due", "days_late"] if c in draft_df.columns], errors="ignore")
            draft_df = draft_df.merge(ship_df, on=["project_no", "unit_no"], how="left")
            draft_df["estimated_ship_date"] = pd.to_datetime(draft_df["estimated_ship_datetime"]).dt.date
            draft_df["can_meet_due"] = draft_df["estimated_ship_date"] <= draft_df["due_date"]
            draft_df["days_late"] = (pd.to_datetime(draft_df["estimated_ship_date"]) - pd.to_datetime(draft_df["due_date"])).dt.days.clip(lower=0)
            st.session_state["draft_schedule_df"] = draft_df.sort_values(["priority", "planned_start", "area", "slot_no", "project_no", "unit_no"]).reset_index(drop=True)

            # rebuild draft daily/summary/load from adjusted draft
            sched = CapacityScheduler(DEFAULT_START_DATETIME, get_area_config(), get_calendar_settings())
            st.session_state["draft_daily_df"] = sched.build_daily_station_view(st.session_state["draft_schedule_df"])
            st.session_state["draft_summary_df"] = sched.build_project_summary(st.session_state["draft_schedule_df"])
            st.session_state["draft_load_df"] = sched.build_area_load(st.session_state["draft_schedule_df"])
            st.success("人工調整已套用到草稿排程。")

    with c2:
        if st.button("發布正式版本", use_container_width=True):
            try:
                publish_current_schedule()
                st.session_state["official_schedule_df"] = load_official_schedule_db()
                st.success("草稿排程已發布為正式版本，並存入系統資料庫。")
            except Exception as e:
                st.error(str(e))

    with c3:
        if st.button("切換顯示正式版/草稿版", use_container_width=True):
            if not official_df.empty:
                st.session_state["use_official_schedule"] = not st.session_state.get("use_official_schedule", False)
                st.success("已切換顯示版本。")
            else:
                st.warning("目前還沒有正式版可切換。")

    st.dataframe(zh(st.session_state["draft_schedule_df"][["priority", "machine_no", "product_type", "step_label", "area_name", "slot_no", "planned_start", "planned_finish", "estimated_ship_date"]]), use_container_width=True, height=380, hide_index=True)
    panel_close()


def render_settings_master():
    page_header("主檔設定", "可調整區域產能、標準工時、平日加班、假日工時與工作日設定。", ["設定"])
    tab1, tab2, tab3, tab4 = st.tabs(["區域產能", "標準工時", "行事曆與工時", "補充參數"])

    with tab1:
        panel_open("區域產能與職能主檔", "每區至少區分 A 組 / B 組，系統以兩組總人數計算產能，職能欄位供管理與後續擴充。")
        editor_df = st.data_editor(
            st.session_state["area_config_editor_df"],
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            column_config={
                "區域代碼": st.column_config.TextColumn("區域代碼", disabled=True),
                "區域名稱": st.column_config.TextColumn("區域名稱"),
                "工位數": st.column_config.NumberColumn("工位數", min_value=1, step=1),
                "A組人數": st.column_config.NumberColumn("A組人數", min_value=0, step=1),
                "A組職能": st.column_config.TextColumn("A組職能"),
                "B組人數": st.column_config.NumberColumn("B組人數", min_value=0, step=1),
                "B組職能": st.column_config.TextColumn("B組職能"),
            },
            key="area_master_editor_v10"
        )
        st.session_state["area_config_editor_df"] = pd.DataFrame(editor_df)
        if st.button("儲存區域產能與職能", type="primary", use_container_width=True):
            new_cfg = {}
            for _, row in st.session_state["area_config_editor_df"].iterrows():
                new_cfg[str(row["區域代碼"])] = {
                    "name": str(row["區域名稱"]),
                    "positions": int(row["工位數"]),
                    "group_a_count": int(row["A組人數"]),
                    "group_a_skill": str(row["A組職能"]),
                    "group_b_count": int(row["B組人數"]),
                    "group_b_skill": str(row["B組職能"]),
                }
            st.session_state["area_config"] = new_cfg
            st.success("區域產能與職能設定已儲存。重新排程後會套用。")
        panel_close()

    with tab2:
        panel_open("標準工時主檔")
        hours_df = st.data_editor(
            st.session_state["hours_config_df"],
            num_rows="fixed",
            use_container_width=True,
            hide_index=True,
            column_config={"模組": st.column_config.TextColumn("模組"), "標準工時(hrs)": st.column_config.NumberColumn("標準工時(hrs)", min_value=0.0, step=0.5)},
            key="hours_editor_v10"
        )
        st.session_state["hours_config_df"] = pd.DataFrame(hours_df)
        if st.button("儲存標準工時", type="primary", use_container_width=True):
            st.success("標準工時已儲存。重新排程後會套用。")
        panel_close()

    with tab3:
        panel_open("工作日與工時設定", "排程只會往現在之後排，不會排到過去日期。可直接點選日曆把日期設成假日或加班工作日。")
        settings = st.session_state["calendar_settings"]
        c1, c2 = st.columns(2)
        with c1:
            default_daily_hours = st.number_input("每日基本工時", min_value=1.0, max_value=24.0, step=0.5, value=float(settings["default_daily_hours"]))
            weekday_ot_hours = st.number_input("平日加班工時", min_value=0.0, max_value=12.0, step=0.5, value=float(settings["weekday_ot_hours"]))
            holiday_work_hours = st.number_input("假日工時（若假日設為工作日）", min_value=0.0, max_value=24.0, step=0.5, value=float(settings["holiday_work_hours"]))
            stk_precheck_ratio = st.slider("STK前段可先行比例", min_value=0.5, max_value=0.95, step=0.05, value=float(settings["stk_precheck_ratio"]))
        with c2:
            st.markdown("**說明**")
            st.caption("1. 假日：系統不排程。")
            st.caption("2. 加班工作日：原本非工作日，但你希望可排程。")
            st.caption("3. 平日工時 = 每日基本工時 + 平日加班工時。")

        if "holiday_dates_temp" not in st.session_state:
            st.session_state["holiday_dates_temp"] = parse_date_text(settings["holidays_text"])
        if "special_workday_dates_temp" not in st.session_state:
            st.session_state["special_workday_dates_temp"] = parse_date_text(settings["special_workdays_text"])

        holiday_dates = set(st.session_state["holiday_dates_temp"])
        special_workday_dates = set(st.session_state["special_workday_dates_temp"])

        sub1, sub2 = st.tabs(["假日點選", "加班工作日點選"])
        with sub1:
            holiday_dates = render_calendar_selector("holiday", "點選要設為假日的日期", holiday_dates)
            st.session_state["holiday_dates_temp"] = set(holiday_dates)
            if st.button("清空假日", key="clear_holiday_dates", use_container_width=True):
                st.session_state["holiday_dates_temp"] = set()
                st.rerun()
        with sub2:
            special_workday_dates = render_calendar_selector("special_workday", "點選要設為加班工作日的日期", special_workday_dates)
            st.session_state["special_workday_dates_temp"] = set(special_workday_dates)
            if st.button("清空加班工作日", key="clear_special_workdays", use_container_width=True):
                st.session_state["special_workday_dates_temp"] = set()
                st.rerun()

        if st.button("儲存工作日與工時設定", type="primary", use_container_width=True):
            holiday_dates = set(st.session_state["holiday_dates_temp"])
            special_workday_dates = set(st.session_state["special_workday_dates_temp"])
            st.session_state["calendar_settings"] = {
                "default_daily_hours": float(default_daily_hours),
                "weekday_ot_hours": float(weekday_ot_hours),
                "holiday_work_hours": float(holiday_work_hours),
                "special_workdays_text": ",".join(sorted([d.isoformat() for d in special_workday_dates])),
                "holidays_text": ",".join(sorted([d.isoformat() for d in holiday_dates])),
                "stk_precheck_ratio": float(stk_precheck_ratio),
            }
            st.success("工作日與工時設定已儲存。重新排程後會套用。")
        panel_close()



    with tab4:
        panel_open("補充參數設定", "補充管理報表用的換線時間、品檢時間與重工返修紀錄。")
        left, right = st.columns([1, 1])
        with left:
            st.markdown("**換線時間設定**")
            setup_df = st.data_editor(
                st.session_state["setup_time_config_df"],
                num_rows="fixed",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "區域代碼": st.column_config.TextColumn("區域代碼"),
                    "換線時間(hrs)": st.column_config.NumberColumn("換線時間(hrs)", min_value=0.0, step=0.5),
                },
                key="setup_time_editor_v12",
            )
            st.session_state["setup_time_config_df"] = pd.DataFrame(setup_df)

            q = st.session_state["quality_time_settings"]
            ipqc_hours = st.number_input("IPQC 時間(hrs)", min_value=0.0, max_value=24.0, step=0.5, value=float(q["ipqc_hours"]))
            fqc_hours = st.number_input("FOC / FQC 時間(hrs)", min_value=0.0, max_value=24.0, step=0.5, value=float(q["fqc_hours"]))
            oqc_hours = st.number_input("OQC 時間(hrs)", min_value=0.0, max_value=24.0, step=0.5, value=float(q["oqc_hours"]))

            if st.button("儲存補充參數", type="primary", use_container_width=True):
                st.session_state["quality_time_settings"] = {
                    "ipqc_hours": float(ipqc_hours),
                    "fqc_hours": float(fqc_hours),
                    "oqc_hours": float(oqc_hours),
                }
                st.success("補充參數已儲存。")

        with right:
            st.markdown("**重工 / 返修紀錄**")
            rework_df = st.data_editor(
                st.session_state["rework_log_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "日期": st.column_config.DateColumn("日期"),
                    "專案號": st.column_config.TextColumn("專案號"),
                    "機台編號": st.column_config.TextColumn("機台編號"),
                    "原因": st.column_config.TextColumn("原因"),
                    "返修工時": st.column_config.NumberColumn("返修工時", min_value=0.0, step=0.5),
                    "狀態": st.column_config.SelectboxColumn("狀態", options=["待處理", "處理中", "已完成"]),
                },
                key="rework_log_editor_v12",
            )
            st.session_state["rework_log_df"] = pd.DataFrame(rework_df)
        panel_close()

PAGE_RENDERERS = {
    "dashboard": render_dashboard, "forecast": render_forecast, "schedule_list": render_schedule_list,
    "station_board": render_station_board, "station_cards": render_station_cards,
    "machine_master": render_machine_master,
    "production_report": render_production_report,
    "production_report_list": render_production_report_list,
    "actual_station_board": render_actual_station_board, "capacity_report": render_capacity_report,
    "delivery_report": render_delivery_report, "management_report": render_management_report, "weekly_receipt_report": render_weekly_receipt_report, "gantt": render_gantt, "review_publish": render_review_publish, "settings_master": render_settings_master,
}

ensure_state()
sidebar()
if st.session_state.get("active_page") not in PAGE_RENDERERS:
    st.session_state["active_page"] = "dashboard"
PAGE_RENDERERS[st.session_state["active_page"]]()
