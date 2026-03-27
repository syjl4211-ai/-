from __future__ import annotations

import io
from dataclasses import dataclass
from datetime import datetime, timedelta, date
from typing import Dict, List, Tuple, Optional

import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="製造排程管理系統",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.block-container {
    padding-top: 1.0rem;
    padding-bottom: 1.4rem;
    max-width: 1650px;
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0b1220 0%, #111827 55%, #172033 100%);
    border-right: 1px solid rgba(255,255,255,0.08);
}
[data-testid="stSidebar"] * {
    color: #e5edf8 !important;
}
[data-testid="stSidebar"] .stCaption {
    color: #9fb0cb !important;
}
.sidebar-brand {
    padding: 0.2rem 0 0.6rem 0;
}
.sidebar-brand-title {
    font-size: 1.25rem;
    font-weight: 800;
    color: #ffffff;
    line-height: 1.2;
}
.sidebar-brand-sub {
    color: #94a3b8;
    font-size: 0.85rem;
    margin-top: 0.2rem;
}
[data-testid="stSidebar"] details {
    background: linear-gradient(180deg, rgba(255,255,255,0.045) 0%, rgba(255,255,255,0.02) 100%);
    border: 1px solid rgba(148, 163, 184, 0.16);
    border-radius: 16px;
    padding: 6px 8px 8px 8px;
    margin-bottom: 12px;
}
[data-testid="stSidebar"] summary {
    color: #f8fafc !important;
    font-weight: 800;
    font-size: 0.96rem;
}
[data-testid="stSidebar"] .stButton > button {
    width: 100%;
    border-radius: 12px;
    min-height: 42px;
    text-align: left;
    justify-content: flex-start;
    font-weight: 700;
    background: rgba(255,255,255,0.02) !important;
    color: #dbe7f7 !important;
    border: 1px solid rgba(148, 163, 184, 0.16) !important;
    box-shadow: none !important;
}
[data-testid="stSidebar"] .stButton > button:hover {
    border-color: rgba(96,165,250,0.55) !important;
    background: rgba(59,130,246,0.14) !important;
    color: #ffffff !important;
}
[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
    color: #ffffff !important;
    border: 1px solid rgba(147,197,253,0.45) !important;
}
.app-title {
    font-size: 1.7rem;
    font-weight: 800;
    color: #0f172a;
    margin-bottom: 0.15rem;
}
.app-subtitle {
    color: #475569;
    margin-bottom: 1rem;
}
.kpi-card {
    background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%);
    border: 1px solid #e2e8f0;
    border-radius: 18px;
    padding: 16px 18px;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.04);
}
.kpi-label {
    font-size: 0.82rem;
    color: #64748b;
    margin-bottom: 0.35rem;
}
.kpi-value {
    font-size: 1.65rem;
    font-weight: 800;
    color: #0f172a;
}
.panel {
    background: white;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 18px 18px 8px 18px;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.04);
    margin-bottom: 14px;
}
.panel-title {
    font-size: 1.02rem;
    font-weight: 800;
    color: #0f172a;
    margin-bottom: 0.35rem;
}
.panel-subtitle {
    color: #64748b;
    font-size: 0.92rem;
    margin-bottom: 0.8rem;
}
.tag {
    display: inline-block;
    font-size: 0.76rem;
    font-weight: 700;
    color: #1d4ed8;
    background: #dbeafe;
    padding: 4px 10px;
    border-radius: 999px;
    margin-right: 6px;
}
.section-divider {
    margin-top: 0.5rem;
    margin-bottom: 0.75rem;
    border-top: 1px solid rgba(148,163,184,0.18);
}
.small-note {
    color: #64748b;
    font-size: 0.86rem;
}
.station-card {
    background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
    border: 2px solid #22c55e;
    border-radius: 20px;
    padding: 14px 14px 12px 14px;
    min-height: 310px;
    box-shadow: 0 8px 18px rgba(15, 23, 42, 0.06);
}
.station-top {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 8px;
}
.station-top-left {
    color: #64748b;
    font-size: 0.82rem;
    font-weight: 700;
}
.station-slot {
    font-size: 1.45rem;
    font-weight: 900;
    color: #0f172a;
    margin-top: 6px;
}
.station-chip {
    padding: 5px 10px;
    border-radius: 999px;
    font-size: 0.8rem;
    font-weight: 800;
    color: #166534;
    background: #dcfce7;
    border: 1px solid #4ade80;
}
.station-order {
    font-size: 0.98rem;
    font-weight: 900;
    color: #0f172a;
    text-align: center;
    margin-top: 4px;
}
.station-meta {
    text-align: center;
    color: #64748b;
    font-size: 0.84rem;
    line-height: 1.45;
    margin-top: 6px;
}
.station-row {
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: #0f172a;
    font-weight: 800;
    font-size: 0.9rem;
    margin-top: 12px;
}
.progress-wrap {
    width: 100%;
    height: 8px;
    background: #d1d5db;
    border-radius: 999px;
    overflow: hidden;
    margin-top: 6px;
}
.progress-fill-step {
    height: 100%;
    background: linear-gradient(90deg, #7c3aed 0%, #4f46e5 100%);
    border-radius: 999px;
}
.progress-fill-total {
    height: 100%;
    background: linear-gradient(90deg, #6366f1 0%, #8b5cf6 100%);
    border-radius: 999px;
}
.station-foot {
    text-align: center;
    color: #6b7280;
    font-size: 0.8rem;
    line-height: 1.35;
    margin-top: 12px;
}
.empty-card {
    border: 1px dashed #cbd5e1;
    border-radius: 20px;
    padding: 30px 18px;
    min-height: 240px;
    background: #f8fafc;
    text-align: center;
    color: #64748b;
}
</style>
""", unsafe_allow_html=True)

DEFAULT_START_DATETIME = datetime(2026, 1, 1, 8, 0)
WORK_START_HOUR = 8
WORK_END_HOUR = 16
WORK_DAYS = {0, 1, 2, 3, 4}

DEFAULT_AREA_CONFIG = {
    "A": {"name": "A區 / FIMS-1", "positions": 6, "manpower": 6},
    "B": {"name": "B區 / FIMS-2", "positions": 6, "manpower": 6},
    "C": {"name": "C區 / Robot", "positions": 5, "manpower": 5},
    "D": {"name": "D區 / STK4", "positions": 10, "manpower": 7},
    "E": {"name": "E區 / STK5", "positions": 6, "manpower": 5},
}

PROCESS_ROUTES = {
    "STK4": [
        {"step": "FIMS-1-4", "area": "A", "hours": 37.0, "depends_on": []},
        {"step": "FIMS-2-4", "area": "B", "hours": 34.5, "depends_on": ["FIMS-1-4"]},
        {"step": "robot4", "area": "C", "hours": 71.5, "depends_on": []},
        {"step": "STK4", "area": "D", "hours": 218.0, "depends_on": ["FIMS-2-4", "robot4"]},
    ],
    "STK5": [
        {"step": "FIMS-1-5", "area": "A", "hours": 37.0, "depends_on": []},
        {"step": "FIMS-2-5", "area": "B", "hours": 34.5, "depends_on": ["FIMS-1-5"]},
        {"step": "robot5", "area": "C", "hours": 60.0, "depends_on": []},
        {"step": "STK5", "area": "E", "hours": 144.0, "depends_on": ["FIMS-2-5", "robot5"]},
    ],
}

STEP_LABELS = {
    "FIMS-1-4": "FIMS-1-4",
    "FIMS-2-4": "FIMS-2-4",
    "robot4": "Robot4",
    "STK4": "STK4 組裝",
    "FIMS-1-5": "FIMS-1-5",
    "FIMS-2-5": "FIMS-2-5",
    "robot5": "Robot5",
    "STK5": "STK5 組裝",
}

@dataclass
class Task:
    project_no: str
    product_type: str
    unit_no: int
    step: str
    area: str
    hours: float
    due_date: date
    depends_on: List[str]

def get_area_config() -> Dict[str, Dict]:
    return st.session_state["area_config"]

def next_workday(d: date) -> date:
    x = d
    while x.weekday() not in WORK_DAYS:
        x += timedelta(days=1)
    return x

def align_to_work_time(dt: datetime) -> datetime:
    if dt.weekday() not in WORK_DAYS:
        return datetime.combine(next_workday(dt.date()), datetime.min.time()).replace(hour=WORK_START_HOUR)
    if dt.hour < WORK_START_HOUR:
        return dt.replace(hour=WORK_START_HOUR, minute=0, second=0, microsecond=0)
    if dt.hour >= WORK_END_HOUR:
        return datetime.combine(next_workday(dt.date() + timedelta(days=1)), datetime.min.time()).replace(hour=WORK_START_HOUR)
    return dt.replace(second=0, microsecond=0)

def add_work_hours_forward(start_dt: datetime, hours: float) -> datetime:
    current = align_to_work_time(start_dt)
    remaining = float(hours)
    while remaining > 1e-9:
        current = align_to_work_time(current)
        day_end = current.replace(hour=WORK_END_HOUR, minute=0, second=0, microsecond=0)
        available = (day_end - current).total_seconds() / 3600.0
        if available <= 1e-9:
            current = datetime.combine(next_workday(current.date() + timedelta(days=1)), datetime.min.time()).replace(hour=WORK_START_HOUR)
            continue
        consume = min(available, remaining)
        current += timedelta(hours=consume)
        remaining -= consume
        if remaining > 1e-9 and current >= day_end:
            current = datetime.combine(next_workday(current.date() + timedelta(days=1)), datetime.min.time()).replace(hour=WORK_START_HOUR)
    return current

def infer_product_type(project_no: str) -> Optional[str]:
    s = str(project_no).upper().replace(" ", "").replace("-", "")
    if "STK4" in s or s.endswith("4"):
        return "STK4"
    if "STK5" in s or s.endswith("5"):
        return "STK5"
    return None

def normalize_forecast(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "專案號": "project_no", "project_no": "project_no",
        "產品型號": "product_type", "product_type": "product_type",
        "台數": "qty", "qty": "qty", "數量": "qty",
        "預計出貨日": "due_date", "due_date": "due_date", "交期": "due_date",
    }
    df = df.rename(columns={c: rename_map.get(c, c) for c in df.columns})
    required = {"project_no", "qty", "due_date"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Forecast 缺少必要欄位：{', '.join(sorted(missing))}")
    df = df.copy()
    df["project_no"] = df["project_no"].astype(str)
    df["qty"] = pd.to_numeric(df["qty"], errors="coerce")
    df["due_date"] = pd.to_datetime(df["due_date"], errors="coerce")
    if "product_type" not in df.columns:
        df["product_type"] = df["project_no"].apply(infer_product_type)
    else:
        df["product_type"] = df["product_type"].fillna(df["project_no"].apply(infer_product_type))
    df = df[df["project_no"].notna() & df["qty"].notna() & df["due_date"].notna()]
    bad = df[df["product_type"].isna()]
    if not bad.empty:
        raise ValueError("以下專案無法判斷產品型號，請補 product_type（STK4 或 STK5）：\n" + "\n".join(bad["project_no"].astype(str).tolist()))
    df["qty"] = df["qty"].astype(int)
    df["product_type"] = df["product_type"].astype(str).str.upper()
    df["due_date"] = pd.to_datetime(df["due_date"]).dt.date
    return df[["project_no", "product_type", "qty", "due_date"]].sort_values(["due_date", "project_no"]).reset_index(drop=True)

def explode_forecast_to_tasks(forecast_df: pd.DataFrame) -> List[Task]:
    tasks: List[Task] = []
    for _, row in forecast_df.iterrows():
        route = PROCESS_ROUTES[row["product_type"]]
        for unit_no in range(1, int(row["qty"]) + 1):
            for step in route:
                tasks.append(Task(
                    project_no=str(row["project_no"]),
                    product_type=str(row["product_type"]),
                    unit_no=int(unit_no),
                    step=str(step["step"]),
                    area=str(step["area"]),
                    hours=float(step["hours"]),
                    due_date=row["due_date"],
                    depends_on=list(step["depends_on"]),
                ))
    return tasks

class CapacityScheduler:
    def __init__(self, start_datetime: datetime, area_config: Dict[str, Dict]):
        self.start_datetime = start_datetime
        self.area_config = area_config
        self.area_slots = {
            area: [dict(last_finish=start_datetime) for _ in range(min(cfg["positions"], cfg["manpower"]))]
            for area, cfg in area_config.items()
        }

    def schedule_forward(self, tasks: List[Task]) -> pd.DataFrame:
        rows = []
        finish_map: Dict[Tuple[str, int, str], datetime] = {}

        def sort_key(t: Task):
            sequence = {route["step"]: idx for idx, route in enumerate(PROCESS_ROUTES[t.product_type], start=1)}
            return (t.due_date, t.project_no, t.unit_no, sequence[t.step], t.step)

        for t in sorted(tasks, key=sort_key):
            predecessors = [finish_map[(t.project_no, t.unit_no, dep)] for dep in t.depends_on] if t.depends_on else []
            earliest = max(predecessors) if predecessors else self.start_datetime
            best_slot, best_start, best_finish = None, None, None
            for idx, slot in enumerate(self.area_slots[t.area]):
                candidate_start = max(slot["last_finish"], earliest)
                candidate_start = align_to_work_time(candidate_start)
                candidate_finish = add_work_hours_forward(candidate_start, t.hours)
                if best_finish is None or candidate_finish < best_finish:
                    best_slot, best_start, best_finish = idx, candidate_start, candidate_finish
            self.area_slots[t.area][best_slot]["last_finish"] = best_finish
            finish_map[(t.project_no, t.unit_no, t.step)] = best_finish
            rows.append({
                "project_no": t.project_no,
                "product_type": t.product_type,
                "unit_no": t.unit_no,
                "step": t.step,
                "step_label": STEP_LABELS.get(t.step, t.step),
                "area": t.area,
                "area_name": self.area_config[t.area]["name"],
                "slot_no": best_slot + 1,
                "hours": t.hours,
                "planned_start": best_start,
                "planned_finish": best_finish,
                "customer_due_date": datetime.combine(t.due_date, datetime.min.time()),
            })

        df = pd.DataFrame(rows)
        if df.empty:
            return df
        final_steps = df[df["step"].isin(["STK4", "STK5"])][["project_no", "unit_no", "planned_finish"]].rename(columns={"planned_finish": "estimated_ship_datetime"})
        df = df.merge(final_steps, on=["project_no", "unit_no"], how="left")
        df["estimated_ship_date"] = pd.to_datetime(df["estimated_ship_datetime"]).dt.date
        df["due_date"] = pd.to_datetime(df["customer_due_date"]).dt.date
        df["machine_no"] = df["project_no"] + "-" + df["unit_no"].astype(str).str.zfill(3)
        df["can_meet_due"] = df["estimated_ship_date"] <= df["due_date"]
        df["days_late"] = (pd.to_datetime(df["estimated_ship_date"]) - pd.to_datetime(df["due_date"])).dt.days.clip(lower=0)
        return df.sort_values(["planned_start", "area", "slot_no", "project_no", "unit_no"]).reset_index(drop=True)

    def build_daily_station_view(self, schedule_df: pd.DataFrame) -> pd.DataFrame:
        records = []
        if schedule_df.empty:
            return pd.DataFrame()
        for _, row in schedule_df.iterrows():
            d = row["planned_start"].date()
            end_d = row["planned_finish"].date()
            while d <= end_d:
                if d.weekday() in WORK_DAYS:
                    records.append({
                        "date": d,
                        "area": row["area"],
                        "area_name": row["area_name"],
                        "slot_no": row["slot_no"],
                        "project_no": row["project_no"],
                        "product_type": row["product_type"],
                        "unit_no": row["unit_no"],
                        "machine_no": row["machine_no"],
                        "step": row["step"],
                        "step_label": row["step_label"],
                        "estimated_ship_date": row["estimated_ship_date"],
                        "display": f'{row["machine_no"]} / {row["step_label"]}',
                    })
                d += timedelta(days=1)
        return pd.DataFrame(records).sort_values(["date", "area", "slot_no"]).reset_index(drop=True)

    def build_project_summary(self, schedule_df: pd.DataFrame) -> pd.DataFrame:
        if schedule_df.empty:
            return pd.DataFrame()
        return (
            schedule_df.groupby(["project_no", "product_type", "unit_no", "machine_no"], as_index=False)
            .agg(
                planned_start=("planned_start", "min"),
                planned_finish=("planned_finish", "max"),
                due_date=("due_date", "max"),
                estimated_ship_date=("estimated_ship_date", "max"),
                can_meet_due=("can_meet_due", "max"),
                days_late=("days_late", "max"),
                total_hours=("hours", "sum"),
            )
            .sort_values(["due_date", "project_no", "unit_no"])
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
            eff_capacity = min(cfg["positions"], cfg["manpower"])
            d = start_date
            while d <= end_date:
                if d.weekday() in WORK_DAYS:
                    active = area_df[(area_df["planned_start"].dt.date <= d) & (area_df["planned_finish"].dt.date >= d)]
                    busy = active["slot_no"].nunique()
                    rows.append({
                        "date": d, "area": area, "area_name": cfg["name"],
                        "positions": cfg["positions"], "manpower": cfg["manpower"],
                        "effective_parallel_capacity": eff_capacity,
                        "busy_slots": busy,
                        "utilization_pct": busy / eff_capacity if eff_capacity else 0,
                    })
                d += timedelta(days=1)
        return pd.DataFrame(rows).sort_values(["date", "area"]).reset_index(drop=True)

def to_excel_bytes(forecast_df, summary_df, schedule_df, daily_df, load_df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        forecast_df.to_excel(writer, sheet_name="Forecast", index=False)
        summary_df.to_excel(writer, sheet_name="ProjectSummary", index=False)
        schedule_df.to_excel(writer, sheet_name="TaskSchedule", index=False)
        daily_df.to_excel(writer, sheet_name="StationDailyView", index=False)
        load_df.to_excel(writer, sheet_name="AreaLoad", index=False)
    return buffer.getvalue()

def seed_forecast() -> pd.DataFrame:
    return pd.DataFrame([
        {"project_no": "S026-0003", "product_type": "STK5", "qty": 1, "due_date": date(2026, 4, 15)},
        {"project_no": "PJ-STK4-001", "product_type": "STK4", "qty": 2, "due_date": date(2026, 4, 30)},
        {"project_no": "PJ-STK5-002", "product_type": "STK5", "qty": 3, "due_date": date(2026, 5, 20)},
    ])

def kpi(label, value, caption=""):
    st.markdown(f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="small-note">{caption}</div>
        </div>
    """, unsafe_allow_html=True)

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

def build_slot_code(area: str, slot_no: int) -> str:
    return f"{area}{int(slot_no):02d}"

def current_step_progress_percent(row: pd.Series, selected_date: date) -> int:
    start = row["planned_start"]
    finish = row["planned_finish"]
    if selected_date <= start.date():
        return 0
    if selected_date >= finish.date():
        return 100
    total = (finish - start).total_seconds()
    elapsed = (datetime.combine(selected_date, datetime.min.time()).replace(hour=16) - start).total_seconds()
    if total <= 0:
        return 100
    return int(max(0, min(100, round(elapsed / total * 100))))

def total_progress_percent(schedule_df: pd.DataFrame, project_no: str, unit_no: int, selected_date: date) -> int:
    unit_df = schedule_df[(schedule_df["project_no"] == project_no) & (schedule_df["unit_no"] == unit_no)].copy()
    if unit_df.empty:
        return 0
    total_hours = unit_df["hours"].sum()
    complete_hours = 0.0
    for _, r in unit_df.iterrows():
        if selected_date > r["planned_finish"].date():
            complete_hours += float(r["hours"])
        elif r["planned_start"].dt.date if False else False:
            pass
        elif r["planned_start"].date() <= selected_date <= r["planned_finish"].date():
            complete_hours += float(r["hours"]) * current_step_progress_percent(r, selected_date) / 100.0
    if total_hours <= 0:
        return 0
    return int(round(complete_hours / total_hours * 100))

def render_station_card(row: pd.Series, selected_date: date, schedule_df: pd.DataFrame):
    step_pct = current_step_progress_percent(row, selected_date)
    total_pct = total_progress_percent(schedule_df, row["project_no"], row["unit_no"], selected_date)
    slot_code = build_slot_code(row["area"], row["slot_no"])
    chip = "生產中" if step_pct < 100 else "已完工"
    st.markdown(f"""
    <div class="station-card">
        <div class="station-top">
            <div>
                <div class="station-top-left">工位</div>
                <div class="station-slot">{slot_code}</div>
            </div>
            <div class="station-chip">{chip}</div>
        </div>
        <div class="station-order">{row["machine_no"]}</div>
        <div class="station-meta">
            工序：{row["step_label"]}<br>
            預計出貨：{row["estimated_ship_date"]}
        </div>
        <div class="station-row"><span>目前工序進度</span><span>{step_pct}%</span></div>
        <div class="progress-wrap"><div class="progress-fill-step" style="width:{step_pct}%;"></div></div>
        <div class="station-row"><span>整機完成度</span><span>{total_pct}%</span></div>
        <div class="progress-wrap"><div class="progress-fill-total" style="width:{total_pct}%;"></div></div>
        <div class="station-foot">
            區域：{row["area_name"]}<br>
            日期：{selected_date}
        </div>
    </div>
    """, unsafe_allow_html=True)

def run_planning():
    forecast_df = normalize_forecast(st.session_state["forecast_df"])
    scheduler = CapacityScheduler(DEFAULT_START_DATETIME, get_area_config())
    tasks = explode_forecast_to_tasks(forecast_df)
    schedule_df = scheduler.schedule_forward(tasks)
    daily_df = scheduler.build_daily_station_view(schedule_df)
    summary_df = scheduler.build_project_summary(schedule_df)
    load_df = scheduler.build_area_load(schedule_df)
    st.session_state["forecast_df"] = forecast_df
    st.session_state["schedule_df"] = schedule_df
    st.session_state["daily_df"] = daily_df
    st.session_state["summary_df"] = summary_df
    st.session_state["load_df"] = load_df

def ensure_state():
    if "forecast_df" not in st.session_state:
        st.session_state["forecast_df"] = seed_forecast()
    if "active_page" not in st.session_state:
        st.session_state["active_page"] = "dashboard"
    if "area_config" not in st.session_state:
        st.session_state["area_config"] = {k: v.copy() for k, v in DEFAULT_AREA_CONFIG.items()}
    if "area_config_editor_df" not in st.session_state:
        st.session_state["area_config_editor_df"] = pd.DataFrame([
            {"area": area, "area_name": cfg["name"], "positions": cfg["positions"], "manpower": cfg["manpower"]}
            for area, cfg in st.session_state["area_config"].items()
        ])

def has_result() -> bool:
    return all(k in st.session_state for k in ["schedule_df", "daily_df", "summary_df", "load_df"])

def nav_button(label: str, key: str, icon: str):
    active = st.session_state.get("active_page") == key
    if st.button(f"{icon}  {label}", key=f"nav_{key}", use_container_width=True, type="primary" if active else "secondary"):
        st.session_state["active_page"] = key

def sidebar():
    with st.sidebar:
        st.markdown("""
        <div class="sidebar-brand">
            <div class="sidebar-brand-title">🏭 製造排程系統</div>
            <div class="sidebar-brand-sub">Production Planning Suite</div>
        </div>
        """, unsafe_allow_html=True)
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
            nav_button("工序甘特總覽", "gantt", "🗂️")
        with st.expander("⚙️ 系統設定", expanded=False):
            nav_button("主檔設定", "settings_master", "⚙️")
            nav_button("規則說明", "settings_rules", "📘")
        st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)
        st.caption("版本：V5 工位圖連動版")

def render_dashboard():
    page_header("營運儀表板", "提供目前產線、交期、產能與排程健康度的管理視圖。", ["Dashboard", "KPI", "Management"])
    if not has_result():
        st.info("請先到「Forecast輸入」完成排程。")
        return
    summary_df = st.session_state["summary_df"]
    load_df = st.session_state["load_df"]
    daily_df = st.session_state["daily_df"]
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: kpi("總機台數", int(summary_df.shape[0]), "本次排程總台數")
    with c2: kpi("準時出貨", int(summary_df["can_meet_due"].sum()), "符合交期")
    with c3: kpi("延遲台數", int((~summary_df["can_meet_due"]).sum()), "需關注案件")
    with c4: kpi("最晚出貨日", str(summary_df["estimated_ship_date"].max()), "最後完工日")
    with c5: kpi("平均負荷率", f'{round(load_df["utilization_pct"].mean() * 100, 1)}%', "各區平均")

def render_forecast():
    page_header("Forecast輸入", "生管可手動輸入或匯入業務 forecast，系統依工時與產能自動排程。", ["Planning", "Input", "Forecast"])
    left, right = st.columns([1.45, 1])
    with left:
        panel_open("Forecast資料維護")
        mode = st.segmented_control("輸入方式", ["手動輸入", "上傳檔案"], default="手動輸入")
        if mode == "手動輸入":
            edited = st.data_editor(
                st.session_state["forecast_df"],
                num_rows="dynamic",
                use_container_width=True,
                hide_index=True,
                column_config={
                    "project_no": st.column_config.TextColumn("專案號"),
                    "product_type": st.column_config.SelectboxColumn("產品型號", options=["STK4", "STK5"]),
                    "qty": st.column_config.NumberColumn("台數", min_value=1, step=1),
                    "due_date": st.column_config.DateColumn("預計出貨日"),
                },
                key="forecast_editor_v5",
            )
            st.session_state["forecast_df"] = pd.DataFrame(edited)
        else:
            uploaded = st.file_uploader("上傳 Excel / CSV", type=["xlsx", "xls", "csv"])
            if uploaded is not None:
                df = pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv") else pd.read_excel(uploaded)
                st.session_state["forecast_df"] = normalize_forecast(df)
                st.success("Forecast 已成功載入。")
        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("開始排程", type="primary", use_container_width=True):
                try:
                    run_planning()
                    st.success("排程完成，左側各頁面已同步更新。")
                except Exception as e:
                    st.error(f"排程失敗：{e}")
        with c2:
            if st.button("載入示範資料", use_container_width=True):
                st.session_state["forecast_df"] = seed_forecast()
                st.success("已載入示範 Forecast。")
        with c3:
            sample = seed_forecast().to_csv(index=False).encode("utf-8-sig")
            st.download_button("下載範例CSV", data=sample, file_name="forecast_sample.csv", mime="text/csv", use_container_width=True)
        panel_close()
    with right:
        panel_open("排程前檢查")
        df = st.session_state["forecast_df"].copy()
        st.dataframe(df, use_container_width=True, height=340, hide_index=True)
        panel_close()

def render_schedule_list():
    page_header("排程清單", "查看指定日期各工位實際正在裝配的機台與工序，並與工位圖連動。", ["Schedule", "Execution", "Linked"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    df = st.session_state["schedule_df"].copy()
    c1, c2, c3, c4 = st.columns([1, 1, 1, 2])
    with c1:
        areas = st.multiselect("區域", options=sorted(df["area"].unique()), default=sorted(df["area"].unique()))
    with c2:
        products = st.multiselect("型號", options=sorted(df["product_type"].unique()), default=sorted(df["product_type"].unique()))
    with c3:
        selected_date = st.date_input("日期", value=df["planned_start"].min().date(), key="schedule_list_date")
    with c4:
        keyword = st.text_input("搜尋機台/專案", placeholder="例如 S026-0003 或 PJ-STK4-001")
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
    panel_open("工序排程明細", "這裡與工位圖使用相同的日期過濾邏輯。")
    st.dataframe(
        filtered[["area_name", "slot_no", "machine_no", "product_type", "step_label", "planned_start", "planned_finish", "estimated_ship_date"]],
        use_container_width=True, height=560
    )
    panel_close()

def render_station_board():
    page_header("工位看板", "以工位矩陣方式查看指定日期各站別正在生產哪一台機台。", ["Shopfloor", "Board", "Matrix"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    daily_df = st.session_state["daily_df"].copy()
    c1, c2 = st.columns([1, 2])
    with c1:
        selected_date = st.date_input("查看日期", value=daily_df["date"].min(), key="station_board_date")
    with c2:
        area_filter = st.multiselect("區域篩選", options=sorted(daily_df["area"].unique()), default=sorted(daily_df["area"].unique()), key="station_board_area")
    filtered = daily_df[(daily_df["date"] == selected_date) & (daily_df["area"].isin(area_filter))].copy()
    if filtered.empty:
        st.warning("該日沒有生產安排。")
        return
    panel_open("工位矩陣")
    pivot = filtered.pivot_table(index=["area", "area_name"], columns="slot_no", values="display", aggfunc="first").reset_index()
    st.dataframe(pivot, use_container_width=True, height=320)
    panel_close()

def render_station_cards():
    page_header("工位圖", "以卡片方式呈現指定日期各工位正在裝配的機台，並與排程清單同日同資料連動。", ["Cards", "Linked"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    schedule_df = st.session_state["schedule_df"].copy()
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
    panel_open("工位卡片", "這裡顯示的工位、機台、工序，會與同一天的排程清單一致。")
    if active_rows.empty:
        st.markdown('<div class="empty-card">該日期目前沒有工位在製資料。</div>', unsafe_allow_html=True)
        panel_close()
        return
    active_rows = active_rows.sort_values(["area", "slot_no", "planned_start"])
    groups = [active_rows.iloc[i:i + max_cards] for i in range(0, len(active_rows), max_cards)]
    for batch in groups:
        cols = st.columns(max_cards)
        for idx in range(max_cards):
            with cols[idx]:
                if idx < len(batch):
                    render_station_card(batch.iloc[idx], selected_date, schedule_df)
    panel_close()

def render_capacity_report():
    page_header("產能負荷報表", "依日期與區域查看有效產能、忙碌工位數與利用率。", ["Report", "Capacity"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    load_df = st.session_state["load_df"].copy()
    st.dataframe(load_df, use_container_width=True, height=420)

def render_delivery_report():
    page_header("出貨達交報表", "給主管看的交期達成率、風險案件與最終完工日期。", ["Report", "Delivery"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    summary_df = st.session_state["summary_df"].copy()
    st.dataframe(summary_df, use_container_width=True, height=420)

def render_gantt():
    page_header("工序甘特總覽", "用管理視角檢視各工序的開始、完成、所屬工位與交期。", ["Visual", "Timeline"])
    if not has_result():
        st.info("請先到「Forecast輸入」執行排程。")
        return
    schedule_df = st.session_state["schedule_df"].copy()
    schedule_df["task_name"] = schedule_df["machine_no"] + " / " + schedule_df["step_label"]
    st.dataframe(schedule_df[["task_name", "area_name", "slot_no", "planned_start", "planned_finish", "hours", "due_date", "can_meet_due"]], use_container_width=True, height=560)

def render_settings_master():
    page_header("主檔設定", "管理區域工位、人力與產品製程路由。", ["Settings", "Master Data"])
    editor_df = st.data_editor(
        st.session_state["area_config_editor_df"],
        num_rows="fixed",
        use_container_width=True,
        hide_index=True,
        column_config={
            "area": st.column_config.TextColumn("區域代碼", disabled=True),
            "area_name": st.column_config.TextColumn("區域名稱"),
            "positions": st.column_config.NumberColumn("工位數", min_value=1, step=1),
            "manpower": st.column_config.NumberColumn("人力數", min_value=1, step=1),
        },
        key="area_master_editor_v5"
    )
    st.session_state["area_config_editor_df"] = pd.DataFrame(editor_df)
    if st.button("儲存區域主檔", type="primary"):
        new_cfg = {}
        for _, row in st.session_state["area_config_editor_df"].iterrows():
            new_cfg[str(row["area"])] = {"name": str(row["area_name"]), "positions": int(row["positions"]), "manpower": int(row["manpower"])}
        st.session_state["area_config"] = new_cfg
        st.success("區域產能主檔已儲存。重新執行排程後會套用。")

def render_settings_rules():
    page_header("規則說明", "目前內建排程規則。", ["Rules"])
    st.markdown("FIMS 與 Robot 可平行開工；FIMS-1 完成後 FIMS-2 才能開始；STK 最後組裝。")

PAGE_RENDERERS = {
    "dashboard": render_dashboard,
    "forecast": render_forecast,
    "schedule_list": render_schedule_list,
    "station_board": render_station_board,
    "station_cards": render_station_cards,
    "capacity_report": render_capacity_report,
    "delivery_report": render_delivery_report,
    "gantt": render_gantt,
    "settings_master": render_settings_master,
    "settings_rules": render_settings_rules,
}

ensure_state()
sidebar()
PAGE_RENDERERS[st.session_state["active_page"]]()
