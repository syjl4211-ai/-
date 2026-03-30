from __future__ import annotations

import io
from dataclasses import dataclass
from datetime import datetime, timedelta, date, time
from calendar import monthrange
from typing import Dict, List, Tuple, Optional, Set

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
        "出貨日": "due_date", "優先級": "priority", "project_no": "project_no",
        "product_type": "product_type", "qty": "qty", "due_date": "due_date", "priority": "priority",
    }
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})
    required = {"project_no", "qty", "due_date", "priority"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Forecast 缺少必要欄位：{', '.join(sorted(missing))}")
    df = df.copy()
    df["project_no"] = df["project_no"].astype(str)
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce")
    df["priority"] = pd.to_numeric(df["priority"], errors="coerce")
    df["due_date"] = pd.to_datetime(df["due_date"], errors="coerce")
    if "product_type" not in df.columns:
        df["product_type"] = df["project_no"].apply(infer_product_type)
    else:
        df["product_type"] = df["product_type"].fillna(df["project_no"].apply(infer_product_type))
    df = df[df["project_no"].notna() & df["qty"].notna() & df["priority"].notna() & df["due_date"].notna()]
    bad = df[df["product_type"].isna()]
    if not bad.empty:
        raise ValueError("以下專案無法判斷產品型號，請補產品欄位（STK4 或 STK5）：\n" + "\n".join(bad["project_no"].astype(str).tolist()))
    df["qty"] = df["qty"].astype(int)
    df["priority"] = df["priority"].astype(int)
    df["product_type"] = df["product_type"].astype(str).str.upper()
    df["due_date"] = pd.to_datetime(df["due_date"]).dt.date
    return df[["project_no", "product_type", "qty", "due_date", "priority"]].sort_values(["priority", "due_date", "project_no"]).reset_index(drop=True)

def make_routes(hours_map: Dict[str, float], stk_precheck_ratio: float) -> Dict[str, List[Dict]]:
    r = max(0.0, min(1.0, float(stk_precheck_ratio)))
    return {
        "STK4": [
            {"step": "FIMS-1-4", "label": "FIMS-1-4", "area": "A", "hours": hours_map["FIMS-1-4"], "depends_on": []},
            {"step": "FIMS-2-4", "label": "FIMS-2-4", "area": "B", "hours": hours_map["FIMS-2-4"], "depends_on": ["FIMS-1-4"]},
            {"step": "FR301", "label": "FR301", "area": "C", "hours": hours_map["FR301"], "depends_on": []},
            {"step": "STK4_PRE", "label": "STK4前段", "area": "D", "hours": hours_map["STK4"] * r, "depends_on": []},
            {"step": "STK4_QC", "label": "STK4檢證", "area": "D", "hours": hours_map["STK4"] * (1 - r), "depends_on": ["FIMS-2-4", "FR301", "STK4_PRE"]},
        ],
        "STK5": [
            {"step": "FIMS-1-5", "label": "FIMS-1-5", "area": "A", "hours": hours_map["FIMS-1-5"], "depends_on": []},
            {"step": "FIMS-2-5", "label": "FIMS-2-5", "area": "B", "hours": hours_map["FIMS-2-5"], "depends_on": ["FIMS-1-5"]},
            {"step": "FV501", "label": "FV501", "area": "C", "hours": hours_map["FV501"], "depends_on": []},
            {"step": "STK5_PRE", "label": "STK5前段", "area": "E", "hours": hours_map["STK5"] * r, "depends_on": []},
            {"step": "STK5_QC", "label": "STK5檢證", "area": "E", "hours": hours_map["STK5"] * (1 - r), "depends_on": ["FIMS-2-5", "FV501", "STK5_PRE"]},
        ],
    }

def explode_forecast_to_tasks(forecast_df: pd.DataFrame, routes: Dict[str, List[Dict]]) -> List[Task]:
    tasks: List[Task] = []
    for _, row in forecast_df.iterrows():
        route = routes[row["product_type"]]
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
            "can_meet_due": bool(grp["can_meet_due"].max()),
            "days_late": int(grp["days_late"].max()),
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
        planned_df = planned_df[
            planned_df["estimated_ship_date"].notna() &
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
        actual_df = actual_df[
            actual_df["入庫日期"].notna() &
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

def seed_forecast() -> pd.DataFrame:
    return pd.DataFrame([
        {"project_no": "S026-0021", "product_type": "STK4", "qty": 2, "due_date": date(2026, 5, 13), "priority": 2},
        {"project_no": "PJ-STK5-002", "product_type": "STK5", "qty": 3, "due_date": date(2026, 5, 20), "priority": 3},
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
        st.session_state["rework_log_df"] = pd.DataFrame([
            {"日期": pd.NaT, "專案號": "", "機台編號": "", "原因": "", "返修工時": 0.0, "狀態": "待處理"}
        ])
    if "erp_receipt_df" not in st.session_state:
        st.session_state["erp_receipt_df"] = pd.DataFrame([
            {"入庫日期": pd.NaT, "機台編號": "", "數量": 1, "來源": "ERP"}
        ])

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
    st.session_state["official_schedule_df"] = st.session_state["draft_schedule_df"].copy()
    st.session_state["official_summary_df"] = st.session_state["draft_summary_df"].copy()
    st.session_state["official_daily_df"] = st.session_state["draft_daily_df"].copy()
    st.session_state["official_load_df"] = st.session_state["draft_load_df"].copy()
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
    page_header("Forecast輸入", "輸入產品、數量、出貨日、優先級，系統依產能與工時自動排程。", ["排程", "輸入"])
    left, right = st.columns([1.5, 1])
    with left:
        panel_open("Forecast資料維護", "欄位：產品、數量、出貨日、優先級。")
        display_df = st.session_state["forecast_df"].rename(columns={
            "project_no": "專案號", "product_type": "產品", "qty": "數量", "due_date": "出貨日", "priority": "優先級",
        })
        edited = st.data_editor(
            display_df, num_rows="dynamic", use_container_width=True, hide_index=True,
            column_config={
                "專案號": st.column_config.TextColumn("專案號"),
                "產品": st.column_config.SelectboxColumn("產品", options=["STK4", "STK5"]),
                "數量": st.column_config.NumberColumn("數量", min_value=1, step=1),
                "出貨日": st.column_config.DateColumn("出貨日"),
                "優先級": st.column_config.NumberColumn("優先級", min_value=1, step=1),
            },
            key="forecast_editor_v91",
        )
        st.session_state["forecast_df"] = edited.rename(columns={
            "專案號": "project_no", "產品": "product_type", "數量": "qty", "出貨日": "due_date", "優先級": "priority",
        })
        c1, c2 = st.columns(2)
        with c1:
            if st.button("開始排程", type="primary", use_container_width=True):
                try:
                    run_planning()
                    st.success("排程完成，所有頁面資料已同步更新。")
                except Exception as e:
                    st.error(f"排程失敗：{e}")
        with c2:
            if st.button("載入示範資料", use_container_width=True):
                st.session_state["forecast_df"] = seed_forecast()
                st.success("已載入示範 Forecast。")
        panel_close()
    with right:
        panel_open("排程前檢查")
        st.dataframe(display_df, use_container_width=True, height=340, hide_index=True)
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
        target_date = st.date_input("查詢週別", value=pd.to_datetime(summary_df["estimated_ship_date"]).dt.date.min(), key="weekly_receipt_date")
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
    with top3:
        st.metric("目前顯示版本", "正式版" if st.session_state.get("use_official_schedule") and not official_df.empty else "草稿版")

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
                st.success("草稿排程已發布為正式版本。")
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
    "station_board": render_station_board, "station_cards": render_station_cards, "capacity_report": render_capacity_report,
    "delivery_report": render_delivery_report, "management_report": render_management_report, "weekly_receipt_report": render_weekly_receipt_report, "gantt": render_gantt, "review_publish": render_review_publish, "settings_master": render_settings_master,
}

ensure_state()
sidebar()
PAGE_RENDERERS[st.session_state["active_page"]]()
