
# ko_tao_dashboard_streamlit_2569_v10_plus_final.py
# Ko Tao EMS Executive + Dispatch + Fuel Dashboard (2569) – v10+
# Fixes vs v9:
# - EMS DG energy is always computed from Dispatch Profiles_15min (DG_Total_MW, DG_NamSaeng_MW, DG_PEA_MW, DG_Mobile_MW)
# - Base(2568) DG energy + fuel cost are taken from KoTao_StandardValue_Charts.xlsx (sheet: Inputs)
# - Monthly merge uses Month number only => no duplicate month rows / no sawtooth charts
# - Dispatch tab: optional day filter + Group A/B background + includes Headroom and DG split lines
# - Cost tab: clean Base vs EMS monthly table + monthly charts + Group A/B hours
#
# Run:
#   cd "E:\#NR_TRAINING BY CHINA\Python EMS\Import_Grid_P_2024\kohtao"
#   streamlit run .\ko_tao_dashboard_streamlit_2569_v10_plus_final.py

from __future__ import annotations

from pathlib import Path
import os
import re
import datetime as _dt
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt


APP_TITLE = "Ko Tao EMS 2569 – Executive + Dispatch + Fuel Dashboard (v10+)"

# -----------------------------
# Defaults (edit to your local paths)
# -----------------------------
DEFAULT_OUT_DIR = str((Path(__file__).resolve().parent / "data").resolve())

DEFAULT_EXEC_XLSX_NAME = "KoTao_Executive_Charts_2567_2568_2569_EMS_v4.xlsx"
DEFAULT_DISPATCH_XLSX_NAME = "KoTao_Dispatch_GroupAB_2569_v3.xlsx"
DEFAULT_STANDARD_XLSX_NAME = "KoTao_StandardValue_Charts.xlsx"

MONTH_LABELS_TH = {
    1: "ม.ค.", 2: "ก.พ.", 3: "มี.ค.", 4: "เม.ย.", 5: "พ.ค.", 6: "มิ.ย.",
    7: "ก.ค.", 8: "ส.ค.", 9: "ก.ย.", 10: "ต.ค.", 11: "พ.ย.", 12: "ธ.ค."
}
MONTH_LABELS_EN = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}
MONTH_ORDER = list(range(1, 13))

TH_FULL_TO_MONTH = {
    "มกราคม": 1,
    "กุมภาพันธ์": 2,
    "มีนาคม": 3,
    "เมษายน": 4,
    "พฤษภาคม": 5,
    "มิถุนายน": 6,
    "กรกฎาคม": 7,
    "สิงหาคม": 8,
    "กันยายน": 9,
    "ตุลาคม": 10,
    "พฤศจิกายน": 11,
    "ธันวาคม": 12,
    "รวม": None,
}

# -----------------------------
# Utilities
# -----------------------------
def _strip_quotes(s: str) -> str:
    return str(s).strip().strip('"').strip("'")


def _pick_default_excel(out_dir: str, candidates: list[str]) -> str:
    out = Path(out_dir)
    for name in candidates:
        cand = out / name
        if cand.exists():
            return str(cand)
    return str(out / candidates[0])  # fallback (even if missing)


def _default_exec_path() -> str:
    env = os.environ.get("KOTAO_EXEC_XLSX")
    if env:
        return _strip_quotes(env)
    return _pick_default_excel(DEFAULT_OUT_DIR, [
        DEFAULT_EXEC_XLSX_NAME,
        DEFAULT_EXEC_XLSX_NAME.replace(".xlsx", ".xlsm"),
        "KoTao_Executive_Charts_2567_2568_2569_EMS_v3.xlsx",
    ])


def _default_dispatch_path() -> str:
    env = os.environ.get("KOTAO_DISPATCH_XLSX")
    if env:
        return _strip_quotes(env)
    out = Path(DEFAULT_OUT_DIR)
    candidates = [
        DEFAULT_DISPATCH_XLSX_NAME,
        DEFAULT_DISPATCH_XLSX_NAME.replace(".xlsx", ".xlsm"),
        "KoTao_Dispatch_Test3_withbaseline.xlsx",
        "KoTao_Dispatch_Test2.xlsx",
        "KoTao_Dispatch_GroupAB_REAL.xlsx",
    ]
    p = _pick_default_excel(DEFAULT_OUT_DIR, candidates)
    if Path(p).exists():
        return p
    if out.exists():
        matches = sorted(out.glob("KoTao_Dispatch*.xls*"))
        if matches:
            return str(matches[0])
    return p


def _default_standard_path() -> str:
    env = os.environ.get("KOTAO_STD_XLSX")
    if env:
        return _strip_quotes(env)
    return _pick_default_excel(DEFAULT_OUT_DIR, [
        DEFAULT_STANDARD_XLSX_NAME,
        "KoTao_StandardValue_Charts_rev1.xlsx",
        "KoTao_StandardValue.xlsx",
    ])


def _list_xlsx_in_outdir(out_dir: str) -> list[str]:
    p = Path(out_dir)
    if not p.exists():
        return []
    files = sorted([x.name for x in p.glob("*.xls*")])
    return files[:250]


def _ensure_dt(df: pd.DataFrame, col: str = "Datetime") -> pd.DataFrame:
    out = df.copy()
    if col in out.columns:
        out[col] = pd.to_datetime(out[col], errors="coerce")
    return out


def _infer_dt_h(dt_series: pd.Series, default: str = "15min") -> float:
    try:
        freq = pd.infer_freq(pd.to_datetime(dt_series))
        if not freq:
            freq = default
        return pd.Timedelta(str(freq)).total_seconds() / 3600.0
    except Exception:
        return pd.Timedelta(default).total_seconds() / 3600.0


# -----------------------------
# Loaders
# -----------------------------
@st.cache_data(show_spinner=False)
def load_exec_pack(path: str) -> dict[str, pd.DataFrame]:
    xlsx = pd.ExcelFile(path)
    out: dict[str, pd.DataFrame] = {}
    for sh in ["Exec_Summary", "Monthly", "TypicalDay_PeakMonth", "Inputs"]:
        if sh in xlsx.sheet_names:
            out[sh] = pd.read_excel(xlsx, sheet_name=sh)
    return out


@st.cache_data(show_spinner=False)
def load_dispatch_pack(path: str) -> dict[str, pd.DataFrame]:
    xlsx = pd.ExcelFile(path)
    out: dict[str, pd.DataFrame] = {}
    for sh in [
        "Profiles_15min",
        "Baseline_15min",
        "Monthly_Summary_EMS",
        "Monthly_Summary_Base",
        "TypicalDay_PeakMonth",
    ]:
        if sh in xlsx.sheet_names:
            out[sh] = pd.read_excel(xlsx, sheet_name=sh)
    return out


@st.cache_data(show_spinner=False)
def load_standard_value_2568(path: str) -> pd.DataFrame:
    """
    Parse KoTao_StandardValue_Charts.xlsx sheet 'Inputs' which has a header row inside the sheet.
    Output columns:
      Month (1..12), Month_TH, Month_EN,
      Fuel_L_2568, FuelCost_THB_2568, Energy_kWh_2568, AllIn_THB_per_kWh_2568
    """
    xlsx = pd.ExcelFile(path)
    if "Inputs" not in xlsx.sheet_names:
        raise ValueError("StandardValue file must contain sheet 'Inputs'")

    raw = pd.read_excel(path, sheet_name="Inputs", header=None)

    header_row = None
    for i in range(raw.shape[0]):
        if raw.iloc[i].astype(str).str.contains(r"\bMonth\b", regex=True).any():
            header_row = i
            break
    if header_row is None:
        raise ValueError("Cannot find header row in StandardValue Inputs sheet")

    header = raw.iloc[header_row].tolist()
    data = raw.iloc[header_row + 1:].copy()
    data.columns = [str(x).strip() for x in header]

    # keep only month rows
    data = data.rename(columns={"Month (TH)": "Month_TH_full", "Month": "Month_EN"})
    data["Month_TH_full"] = data["Month_TH_full"].astype(str).str.strip()
    data = data[~data["Month_TH_full"].isin(["nan", "NaN", "None"])].copy()

    data["Month"] = data["Month_TH_full"].map(TH_FULL_TO_MONTH)
    data = data[data["Month"].notna()].copy()
    data["Month"] = data["Month"].astype(int)

    # select columns (tolerant)
    def _col(name: str) -> str:
        if name in data.columns:
            return name
        raise KeyError(f"Missing column '{name}' in Inputs")

    cols = {
        "Fuel_L_2568": _col("Fuel (L) 2568"),
        "FuelCost_THB_2568": _col("Fuel Cost (THB) 2568"),
        "Energy_kWh_2568": _col("Energy (kWh) 2568"),
        "AllIn_THB_per_kWh_2568": _col("All-in (THB/kWh) 2568"),
    }

    out = pd.DataFrame({
        "Month": data["Month"].to_numpy(int),
        "Month_TH": [MONTH_LABELS_TH[int(m)] for m in data["Month"].to_numpy(int)],
        "Month_EN": [MONTH_LABELS_EN[int(m)] for m in data["Month"].to_numpy(int)],
        "Fuel_L_2568": pd.to_numeric(data[cols["Fuel_L_2568"]], errors="coerce").fillna(0.0),
        "FuelCost_THB_2568": pd.to_numeric(data[cols["FuelCost_THB_2568"]], errors="coerce").fillna(0.0),
        "Energy_kWh_2568": pd.to_numeric(data[cols["Energy_kWh_2568"]], errors="coerce").fillna(0.0),
        "AllIn_THB_per_kWh_2568": pd.to_numeric(data[cols["AllIn_THB_per_kWh_2568"]], errors="coerce").fillna(0.0),
    }).sort_values("Month")

    # guard: ensure 12 rows
    if len(out) != 12:
        # still allow, but warn in UI
        pass

    return out


# -----------------------------
# Fuel-price helpers (All-in proxy)
# -----------------------------
def _parse_month_prices(text: str) -> dict[int, float]:
    """Parse:
      1=10.8
      2:10.7
      or "1=10.8,2=10.7"
    """
    if not text or not str(text).strip():
        return {}
    t = str(text).strip()
    t = t.replace(",", "\n").replace(";", "\n")
    out: dict[int, float] = {}
    for raw in t.splitlines():
        s = raw.strip()
        if not s:
            continue
        m = re.match(r"^\s*(\d{1,2})\s*[:=]\s*([0-9.]+)\s*$", s)
        if not m:
            continue
        mm = int(m.group(1))
        if 1 <= mm <= 12:
            out[mm] = float(m.group(2))
    return out


def _month_series_from_dt(dt: pd.Series, month_values: dict[int, float], fallback: float) -> np.ndarray:
    months = pd.to_datetime(dt).dt.month.to_numpy()
    out = np.full(len(months), float(fallback), dtype=float)
    if month_values:
        for m, v in month_values.items():
            out[months == int(m)] = float(v)
    return out


# -----------------------------
# Dispatch normalizer (DG split + Mobile + caps)
# -----------------------------
def _ensure_dg_split(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    for c in ["DG_Total_MW", "DG_NamSaeng_MW", "DG_PEA_MW"]:
        if c in d.columns:
            d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
        else:
            d[c] = 0.0

    # caps (per your rules)
    d["DG_NamSaeng_MW"] = d["DG_NamSaeng_MW"].clip(lower=0.0, upper=5.0)
    d["DG_PEA_MW"] = d["DG_PEA_MW"].clip(lower=0.0, upper=8.0)

    if "DG_Mobile_MW" in d.columns:
        d["DG_Mobile_MW"] = pd.to_numeric(d["DG_Mobile_MW"], errors="coerce").fillna(0.0)
    else:
        d["DG_Mobile_MW"] = (d["DG_Total_MW"] - d["DG_NamSaeng_MW"] - d["DG_PEA_MW"]).fillna(0.0)

    d["DG_Mobile_MW"] = d["DG_Mobile_MW"].clip(lower=0.0, upper=2.0)

    # DG_Total fallback if missing (should not happen)
    if "DG_Total_MW" not in df.columns or df["DG_Total_MW"].isna().all():
        d["DG_Total_MW"] = d["DG_NamSaeng_MW"] + d["DG_PEA_MW"] + d["DG_Mobile_MW"]

    return d


# -----------------------------
# Monthly aggregations
# -----------------------------
def _monthly_from_dispatch(df: pd.DataFrame, dt_h: float, allin_fallback: float, allin_month_override: dict[int, float], cost_mode: str) -> pd.DataFrame:
    """
    df: Profiles_15min (EMS)
    cost_mode:
      - "proxy_allin": DG_kWh * All-in
      - "from_dispatch": use FuelCost_THB column if present, else fallback to proxy
    """
    d = df.copy()
    d["Datetime"] = pd.to_datetime(d["Datetime"])
    d["Month"] = d["Datetime"].dt.month.astype(int)

    d = _ensure_dg_split(d)

    # energy per step
    d["DG_kWh_step"] = d["DG_Total_MW"].to_numpy(float) * 1000.0 * float(dt_h)
    d["NS_kWh_step"] = d["DG_NamSaeng_MW"].to_numpy(float) * 1000.0 * float(dt_h)
    d["PEA_kWh_step"] = d["DG_PEA_MW"].to_numpy(float) * 1000.0 * float(dt_h)
    d["Mobile_kWh_step"] = d["DG_Mobile_MW"].to_numpy(float) * 1000.0 * float(dt_h)

    # all-in
    allin = _month_series_from_dt(d["Datetime"], allin_month_override, float(allin_fallback))
    d["AllIn_THB_per_kWh"] = allin

    d["FuelCost_proxy_THB"] = d["DG_kWh_step"] * allin

    if cost_mode == "from_dispatch" and "FuelCost_THB" in d.columns:
        d["FuelCost_THB"] = pd.to_numeric(d["FuelCost_THB"], errors="coerce").fillna(0.0)
    else:
        d["FuelCost_THB"] = d["FuelCost_proxy_THB"]

    if "Fuel_L" in d.columns:
        d["Fuel_L"] = pd.to_numeric(d["Fuel_L"], errors="coerce").fillna(0.0)
    else:
        d["Fuel_L"] = 0.0

    g = d.groupby("Month", as_index=False).agg(
        DG_kWh=("DG_kWh_step", "sum"),
        DG_MWh=("DG_kWh_step", lambda x: float(np.nansum(x)) / 1000.0),
        NS_MWh=("NS_kWh_step", lambda x: float(np.nansum(x)) / 1000.0),
        PEA_MWh=("PEA_kWh_step", lambda x: float(np.nansum(x)) / 1000.0),
        Mobile_MWh=("Mobile_kWh_step", lambda x: float(np.nansum(x)) / 1000.0),
        Fuel_L=("Fuel_L", "sum"),
        FuelCost_THB=("FuelCost_THB", "sum"),
    )

    g["AllIn_THB_per_kWh"] = [float(allin_month_override.get(int(m), float(allin_fallback))) for m in g["Month"].to_numpy(int)]
    g["FuelIntensity_L_per_kWh"] = np.where(g["DG_kWh"] > 0, g["Fuel_L"] / g["DG_kWh"], 0.0)
    g["TotalCost_MTHB"] = g["FuelCost_THB"] / 1_000_000.0

    # add labels
    g["Month_TH"] = g["Month"].map(lambda m: MONTH_LABELS_TH.get(int(m), str(m)))
    g["Month_EN"] = g["Month"].map(lambda m: MONTH_LABELS_EN.get(int(m), str(m)))

    # ensure full 12 months (fill 0)
    full = pd.DataFrame({"Month": MONTH_ORDER})
    out = full.merge(g, on="Month", how="left").fillna(0.0)
    out["Month_TH"] = out["Month"].map(lambda m: MONTH_LABELS_TH[int(m)])
    out["Month_EN"] = out["Month"].map(lambda m: MONTH_LABELS_EN[int(m)])
    return out


def _monthly_base_from_standard(std2568: pd.DataFrame) -> pd.DataFrame:
    d = std2568.copy()
    out = pd.DataFrame({
        "Month": d["Month"].astype(int),
        "Month_TH": d["Month"].map(lambda m: MONTH_LABELS_TH[int(m)]),
        "Month_EN": d["Month"].map(lambda m: MONTH_LABELS_EN[int(m)]),
        "DG_kWh_base": pd.to_numeric(d["Energy_kWh_2568"], errors="coerce").fillna(0.0),
        "DG_MWh_base": pd.to_numeric(d["Energy_kWh_2568"], errors="coerce").fillna(0.0) / 1000.0,
        "Fuel_L_base": pd.to_numeric(d["Fuel_L_2568"], errors="coerce").fillna(0.0),
        "FuelCost_THB_base": pd.to_numeric(d["FuelCost_THB_2568"], errors="coerce").fillna(0.0),
        "AllIn_base_2568": pd.to_numeric(d["AllIn_THB_per_kWh_2568"], errors="coerce").fillna(0.0),
        "FuelIntensity_base_2568": np.where(
            pd.to_numeric(d["Energy_kWh_2568"], errors="coerce").fillna(0.0) > 0,
            pd.to_numeric(d["Fuel_L_2568"], errors="coerce").fillna(0.0) / pd.to_numeric(d["Energy_kWh_2568"], errors="coerce").fillna(0.0),
            0.0
        ),
    }).sort_values("Month")

    # ensure 12 months (fill missing)
    full = pd.DataFrame({"Month": MONTH_ORDER})
    out = full.merge(out, on="Month", how="left").fillna(0.0)
    out["Month_TH"] = out["Month"].map(lambda m: MONTH_LABELS_TH[int(m)])
    out["Month_EN"] = out["Month"].map(lambda m: MONTH_LABELS_EN[int(m)])
    return out


# -----------------------------
# Charts (Altair with legend-toggle)
# -----------------------------
def _line_chart_legend_toggle(df_long: pd.DataFrame, title: str, y_title: str = "") -> alt.Chart:
    """
    df_long: columns: x (Month or Datetime), metric, value
    Uses legend-click selection to toggle series.
    """
    sel = alt.selection_point(fields=["metric"], bind="legend")
    base = alt.Chart(df_long).mark_line().encode(
        x=alt.X("x:O", sort=MONTH_ORDER, title="Month"),
        y=alt.Y("value:Q", title=y_title),
        color=alt.Color("metric:N", legend=alt.Legend(orient="bottom")),
        opacity=alt.condition(sel, alt.value(1.0), alt.value(0.12)),
        tooltip=["x:O", "metric:N", alt.Tooltip("value:Q", format=",.2f")],
    ).add_params(sel).properties(height=320, title=title)
    return base


def _stacked_bar(df_wide: pd.DataFrame, title: str, cols: list[str]) -> alt.Chart:
    d = df_wide[["Month"] + cols].copy()
    long = d.melt("Month", var_name="metric", value_name="value")
    return (
        alt.Chart(long)
        .mark_bar()
        .encode(
            x=alt.X("Month:O", sort=MONTH_ORDER, title="Month"),
            y=alt.Y("value:Q", title="Hours", stack=True),
            color=alt.Color("metric:N", legend=alt.Legend(orient="bottom")),
            tooltip=["Month:O", "metric:N", alt.Tooltip("value:Q", format=",.1f")],
        )
        .properties(height=320, title=title)
    )


def _build_mode_segments(df: pd.DataFrame, dt_h: float) -> pd.DataFrame:
    d = df[["Datetime", "Mode_final"]].copy()
    d["Datetime"] = pd.to_datetime(d["Datetime"])
    d["Mode_final"] = d["Mode_final"].astype(str).str.upper().fillna("A")
    d["chg"] = (d["Mode_final"] != d["Mode_final"].shift(1)).cumsum()
    seg = d.groupby("chg").agg(
        start=("Datetime", "min"),
        end=("Datetime", "max"),
        Mode=("Mode_final", "first"),
    ).reset_index(drop=True)
    seg["end"] = seg["end"] + pd.to_timedelta(dt_h, unit="h")
    return seg


def _ts_chart_with_mode_bg(df: pd.DataFrame, metrics: list[str], title: str, mode_colors: dict[str, str], mode_opacity: float) -> alt.Chart:
    d = df.copy()
    d["Datetime"] = pd.to_datetime(d["Datetime"])
    dt_h = _infer_dt_h(d["Datetime"], default="15min")

    layers = []
    if "Mode_final" in d.columns:
        seg = _build_mode_segments(d, dt_h)
        mode_scale = alt.Scale(domain=list(mode_colors.keys()), range=list(mode_colors.values()))
        rect = alt.Chart(seg).mark_rect(opacity=float(mode_opacity)).encode(
            x="start:T",
            x2="end:T",
            color=alt.Color("Mode:N", scale=mode_scale, legend=alt.Legend(title="Group", orient="top")),
            tooltip=["Mode:N", "start:T", "end:T"],
        )
        layers.append(rect)

    cols = [c for c in metrics if c in d.columns]
    if not cols:
        return alt.Chart(pd.DataFrame({"Datetime": [], "metric": [], "value": []})).mark_line()

    long = d[["Datetime"] + cols].melt("Datetime", var_name="metric", value_name="value")

    sel = alt.selection_point(fields=["metric"], bind="legend")
    line = alt.Chart(long).mark_line().encode(
        x=alt.X("Datetime:T", title=None),
        y=alt.Y("value:Q", title=None),
        color=alt.Color("metric:N", legend=alt.Legend(orient="bottom")),
        opacity=alt.condition(sel, alt.value(1.0), alt.value(0.12)),
        tooltip=["Datetime:T", "metric:N", alt.Tooltip("value:Q", format=",.3f")],
    ).add_params(sel)

    layers.append(line)
    chart = alt.layer(*layers).properties(height=360, title=title).resolve_scale(color="independent")
    return chart.interactive()


# -----------------------------
# Gen schedule (hours)
# -----------------------------
def _monthly_group_hours(df: pd.DataFrame, dt_h: float) -> pd.DataFrame:
    d = df.copy()
    d["Datetime"] = pd.to_datetime(d["Datetime"])
    d["Month"] = d["Datetime"].dt.month.astype(int)
    if "Mode_final" not in d.columns:
        return pd.DataFrame({"Month": MONTH_ORDER, "A_hours": 0.0, "B_hours": 0.0})
    mf = d["Mode_final"].astype(str).str.upper().fillna("A")
    out = d.assign(mf=mf).groupby("Month", as_index=False).agg(
        A_hours=("mf", lambda x: float((x == "A").sum()) * dt_h),
        B_hours=("mf", lambda x: float((x == "B").sum()) * dt_h),
    )
    full = pd.DataFrame({"Month": MONTH_ORDER})
    out = full.merge(out, on="Month", how="left").fillna(0.0)
    return out


# -----------------------------
# App
# -----------------------------
def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    st.title(APP_TITLE)

    with st.sidebar:
        st.header("Data sources")
        st.caption("สำหรับ Streamlit Cloud: ให้วางไฟล์ Excel ไว้ในโฟลเดอร์ **data/** ใน GitHub repo (อยู่โฟลเดอร์เดียวกับ app.py)")

        data_dir = Path(DEFAULT_OUT_DIR)
        files = _list_xlsx_in_outdir(str(data_dir))

        if not files:
            st.warning("ยังไม่พบไฟล์ .xlsx/.xlsm ในโฟลเดอร์ data/ — กรุณาอัปโหลดไฟล์ Excel ทั้ง 3 ไฟล์เข้า repo ก่อน แล้ว Deploy ใหม่")
            exec_path = ""
            dispatch_path = ""
            std_path = ""
        else:
            def _pick(label: str, default_name: str) -> str:
                idx = 0
                if default_name in files:
                    idx = files.index(default_name)
                choice = st.selectbox(label, options=files, index=idx)
                return str(data_dir / choice)

            exec_path = _pick("Executive pack (xlsx)", DEFAULT_EXEC_XLSX_NAME)
            dispatch_path = _pick("Dispatch output (xlsx)", DEFAULT_DISPATCH_XLSX_NAME)
            std_path = _pick("StandardValue (Base 2568) (xlsx)", DEFAULT_STANDARD_XLSX_NAME)

        with st.expander("Files found in data/ (ช่วยตรวจชื่อไฟล์)"):
            if files:
                st.code("\n".join(files))
            else:
                st.write("ไม่พบไฟล์ .xlsx ในโฟลเดอร์:", str(data_dir))

        st.divider()
        st.header("Fuel cost settings (EMS)")

        st.caption("EMS Fuel cost มี 2 โหมด:\n- **proxy_allin**: DG_kWh × All-in(THB/kWh)\n- **from_dispatch**: ใช้คอลัมน์ FuelCost_THB จากไฟล์ dispatch (ถ้ามี) ไม่งั้น fallback ไป proxy")
        cost_mode = st.selectbox("EMS cost mode", ["proxy_allin", "from_dispatch"], index=0)

        allin_ems = st.number_input("All-in EMS (THB/kWh) – default", value=10.12, step=0.01, format="%.2f")
        st.caption("Optional: ใส่ All-in รายเดือน (override) เช่น 1=10.10,2=10.08,...")
        month_text_ems = st.text_area("All-in monthly EMS (THB/kWh)", value="", height=110)
        month_values_ems = _parse_month_prices(month_text_ems)

        st.divider()
        st.header("Chart styling (Group background)")
        col_a = st.color_picker("Group A background", value="#FFE066")  # yellow
        col_b = st.color_picker("Group B background", value="#B197FC")  # purple
        bg_opacity = st.slider("Background opacity", min_value=0.0, max_value=0.35, value=0.12, step=0.01)

        st.divider()
        st.header("Dispatch view controls")
        month_opts = ["All"] + [f"{m:02d} ({MONTH_LABELS_TH[m]})" for m in range(1, 13)]
        sel_month = st.selectbox("เลือกเดือน (สำหรับกราฟ 15-min)", month_opts, index=0)

    # -----------------------------
    # Load data
    # -----------------------------
    exec_data: dict[str, pd.DataFrame] = {}
    dispatch_data: dict[str, pd.DataFrame] = {}
    std2568: Optional[pd.DataFrame] = None

    if exec_path and Path(exec_path).exists():
        exec_data = load_exec_pack(exec_path)

    if dispatch_path and Path(dispatch_path).exists():
        dispatch_data = load_dispatch_pack(dispatch_path)

    std_error = None
    if std_path and Path(std_path).exists():
        try:
            std2568 = load_standard_value_2568(std_path)
        except Exception as e:
            std_error = str(e)
            std2568 = None

    # -----------------------------
    # Tabs
    # -----------------------------
    tab_exec, tab_dispatch, tab_cost = st.tabs(["📊 Executive pack", "🧭 Dispatch (15-min)", "💸 DG energy & Fuel cost"])

    # =============================
    # Executive pack
    # =============================
    with tab_exec:
        if not exec_data:
            st.info("Executive pack ไม่จำเป็นสำหรับการเทียบ DG Base vs EMS (คุณสามารถเว้นได้) แต่ถ้ามีจะโชว์ KPI เพิ่มเติม")
        else:
            summary = exec_data.get("Exec_Summary", pd.DataFrame())
            monthly = exec_data.get("Monthly", pd.DataFrame())
            typical = exec_data.get("TypicalDay_PeakMonth", pd.DataFrame())

            if not summary.empty:
                st.subheader("Executive KPIs")
                st.dataframe(summary, use_container_width=True)

            if not monthly.empty:
                st.subheader("Monthly (Executive pack)")
                st.dataframe(monthly, use_container_width=True)

            if not typical.empty:
                st.subheader("Typical day (peak month)")
                st.dataframe(typical, use_container_width=True)

            with st.expander("Inputs"):
                if "Inputs" in exec_data:
                    st.dataframe(exec_data["Inputs"], use_container_width=True)

    # =============================
    # Dispatch viewer
    # =============================
    with tab_dispatch:
        if not dispatch_data or "Profiles_15min" not in dispatch_data or dispatch_data["Profiles_15min"].empty:
            st.error("ต้องระบุ Dispatch output ที่มี sheet 'Profiles_15min' (EMS dispatch) ก่อน")
            st.stop()

        prof = _ensure_dt(dispatch_data["Profiles_15min"], "Datetime")
        prof = _ensure_dg_split(prof)

        # dt
        dt_h = _infer_dt_h(prof["Datetime"], default="15min")

        # month filter
        prof["Month"] = prof["Datetime"].dt.month.astype(int)
        if sel_month == "All":
            prof_m = prof.copy()
        else:
            mm = int(sel_month.split()[0])
            prof_m = prof[prof["Month"] == mm].copy()

        # day selector
        prof_m["Date"] = prof_m["Datetime"].dt.date
        dates = sorted(prof_m["Date"].dropna().unique().tolist())
        if not dates:
            st.warning("ไม่มีข้อมูลในช่วงเดือนที่เลือก")
            st.stop()

        left, right = st.columns([1, 3])
        with left:
            day_mode = st.radio("มุมมองเวลา", ["ทั้งช่วงที่เลือก", "เลือกวัน"], index=0, horizontal=True)
            if day_mode == "เลือกวัน":
                picked = st.date_input("เลือกวันที่", value=dates[0], min_value=min(dates), max_value=max(dates))
                if picked not in dates:
                    ords = np.array([d.toordinal() for d in dates], dtype=int)
                    picked_ord = int(picked.toordinal())
                    nearest = dates[int(np.argmin(np.abs(ords - picked_ord)))]
                    st.info(f"วันที่ {picked} ไม่มีข้อมูล → เลือกวันใกล้สุด: {nearest}")
                    picked = nearest
                prof_d = prof_m[prof_m["Date"] == picked].copy()
            else:
                prof_d = prof_m.copy()

            # metric selector
            default_metrics = [
                "DG_Total_MW", "DG_NamSaeng_MW", "DG_PEA_MW", "DG_Mobile_MW",
                "Sub_MW", "Load_KoTao_MW", "Headroom_115kV_MW", "LoadShed_MW",
            ]
            available = [m for m in default_metrics if m in prof_d.columns]
            sel_metrics = st.multiselect("เลือกเส้นที่จะโชว์ (Legend คลิกซ่อน/แสดงได้)", available, default=available)

        # KPI for the selected window
        dg_kwh = float(np.nansum(prof_d["DG_Total_MW"].to_numpy(float) * 1000.0 * dt_h))
        allin_row = _month_series_from_dt(prof_d["Datetime"], month_values_ems, float(allin_ems))
        fuelcost_proxy = float(np.nansum((prof_d["DG_Total_MW"].to_numpy(float) * 1000.0 * dt_h) * allin_row))

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("DG Energy (MWh)", f"{dg_kwh/1000.0:,.2f}")
        k2.metric("Fuel (L) – from sim", f"{float(np.nansum(prof_d.get('Fuel_L', 0.0))):,.0f}")
        k3.metric("Fuel cost (THB) – proxy", f"{fuelcost_proxy:,.0f}")
        k4.metric("All-in effective (THB/kWh)", f"{(fuelcost_proxy/dg_kwh if dg_kwh>0 else 0.0):,.2f}")

        st.subheader("Dispatch time series (with 115kV headroom + Group A/B background)")
        chart = _ts_chart_with_mode_bg(
            prof_d,
            metrics=sel_metrics,
            title="(Legend คลิกซ่อน/แสดง series ได้)",
            mode_colors={"A": col_a, "B": col_b},
            mode_opacity=float(bg_opacity),
        )
        st.altair_chart(chart, use_container_width=True)

        with st.expander("ดูข้อมูล 15-min (raw)"):
            st.dataframe(prof_d.drop(columns=["Month", "Date"], errors="ignore"), use_container_width=True)

    # =============================
    # Fuel cost tab
    # =============================
    with tab_cost:
        if std_error:
            st.error(f"อ่าน StandardValue ไม่ได้: {std_error}")
        if std2568 is None:
            st.warning("เพื่อให้ Base(2568) ถูกต้อง (Fuel cost 399,200,313) กรุณาระบุไฟล์ KoTao_StandardValue_Charts.xlsx ให้ถูกต้อง")
        if not dispatch_data or "Profiles_15min" not in dispatch_data or dispatch_data["Profiles_15min"].empty:
            st.error("ต้องระบุ Dispatch output (Profiles_15min) ก่อน")
            st.stop()

        prof = _ensure_dt(dispatch_data["Profiles_15min"], "Datetime")
        dt_h = _infer_dt_h(prof["Datetime"], default="15min")

        ems_month = _monthly_from_dispatch(
            prof,
            dt_h=dt_h,
            allin_fallback=float(allin_ems),
            allin_month_override=month_values_ems,
            cost_mode=str(cost_mode),
        )

        # base monthly from standard
        if std2568 is not None:
            base_month = _monthly_base_from_standard(std2568)
        else:
            base_month = pd.DataFrame({"Month": MONTH_ORDER})
            base_month["DG_MWh_base"] = 0.0
            base_month["FuelCost_THB_base"] = 0.0
            base_month["Fuel_L_base"] = 0.0
            base_month["Month_TH"] = base_month["Month"].map(lambda m: MONTH_LABELS_TH[int(m)])
            base_month["Month_EN"] = base_month["Month"].map(lambda m: MONTH_LABELS_EN[int(m)])

        # merge on Month only (critical)
        wide = base_month.merge(
            ems_month.rename(columns={
                "DG_MWh": "DG_MWh_ems",
                "DG_kWh": "DG_kWh_ems",
                "Fuel_L": "Fuel_L_ems",
                "FuelCost_THB": "FuelCost_THB_ems",
                "AllIn_THB_per_kWh": "AllIn_ems",
                "NS_MWh": "NS_MWh_ems",
                "PEA_MWh": "PEA_MWh_ems",
                "Mobile_MWh": "Mobile_MWh_ems",
            }),
            on="Month",
            how="left",
            suffixes=("", "_dup"),
        ).fillna(0.0)

        wide["Month_TH"] = wide["Month"].map(lambda m: MONTH_LABELS_TH[int(m)])
        wide["Month_EN"] = wide["Month"].map(lambda m: MONTH_LABELS_EN[int(m)])

        # savings
        wide["Saving_THB"] = wide["FuelCost_THB_base"] - wide["FuelCost_THB_ems"]
        wide["Saving_pct"] = np.where(wide["FuelCost_THB_base"] > 0, wide["Saving_THB"] / wide["FuelCost_THB_base"] * 100.0, 0.0)

        # KPIs year
        tot_base = float(np.nansum(wide["FuelCost_THB_base"]))
        tot_ems = float(np.nansum(wide["FuelCost_THB_ems"]))
        saving = tot_base - tot_ems
        saving_pct = (saving / tot_base * 100.0) if tot_base > 0 else 0.0

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Fuel cost Base 2568 (THB)", f"{tot_base:,.0f}")
        k2.metric("Fuel cost EMS 2569 (THB)", f"{tot_ems:,.0f}")
        k3.metric("Saving (THB)", f"{saving:,.0f}")
        k4.metric("Saving (%)", f"{saving_pct:,.1f}%")

        with st.expander("ที่มาของตัวเลข (Base vs EMS)"):
            st.write(
                "**Base 2568**: อ่านค่ารายเดือนจาก KoTao_StandardValue_Charts.xlsx (sheet: Inputs) โดยใช้คอลัมน์\n"
                "- Energy (kWh) 2568 → DG_MWh_base\n"
                "- Fuel Cost (THB) 2568 → FuelCost_THB_base\n"
                "- Fuel (L) 2568 → Fuel_L_base\n\n"
                "**EMS 2569**: อ่านพลังงาน DG จาก Dispatch Profiles_15min แล้วรวมรายเดือน\n"
                "- DG_Total_MW, DG_NamSaeng_MW, DG_PEA_MW, DG_Mobile_MW\n\n"
                "ค่า Fuel cost EMS เลือกได้ 2 โหมด:\n"
                "- proxy_allin: DG_kWh × All-in\n"
                "- from_dispatch: ใช้ FuelCost_THB ในไฟล์ dispatch (ถ้ามี)\n"
            )

        # monthly charts
        c1, c2 = st.columns(2)
        with c1:
            long = pd.DataFrame({
                "x": wide["Month"].astype(int),
                "DG_MWh_base": wide["DG_MWh_base"].astype(float),
                "DG_MWh_ems": wide["DG_MWh_ems"].astype(float),
            }).melt("x", var_name="metric", value_name="value").rename(columns={"x": "x"})
            st.altair_chart(_line_chart_legend_toggle(long, "DG energy (MWh): Base vs EMS", "MWh"), use_container_width=True)

        with c2:
            long = pd.DataFrame({
                "x": wide["Month"].astype(int),
                "FuelCost_THB_base": wide["FuelCost_THB_base"].astype(float),
                "FuelCost_THB_ems": wide["FuelCost_THB_ems"].astype(float),
            }).melt("x", var_name="metric", value_name="value")
            st.altair_chart(_line_chart_legend_toggle(long, "Fuel cost (THB): Base vs EMS", "THB"), use_container_width=True)

        # group hours chart from EMS dispatch
        gh = _monthly_group_hours(_ensure_dt(dispatch_data["Profiles_15min"], "Datetime"), dt_h)
        st.altair_chart(_stacked_bar(gh, "Group A vs Group B (hours / month) – EMS", ["A_hours", "B_hours"]), use_container_width=True)

        # monthly table
        st.subheader("Monthly table (Base 2568 vs EMS 2569)")
        cols = [
            "Month_TH", "Month_EN", "Month",
            "DG_MWh_base", "DG_MWh_ems",
            "NS_MWh_ems", "PEA_MWh_ems", "Mobile_MWh_ems",
            "Fuel_L_base", "Fuel_L_ems",
            "FuelCost_THB_base", "FuelCost_THB_ems",
            "Saving_THB", "Saving_pct",
            "AllIn_ems",
        ]
        cols = [c for c in cols if c in wide.columns]
        st.dataframe(wide[cols].sort_values("Month"), use_container_width=True, height=480)

        # split share chart (MWh) – EMS
        st.subheader("DG split energy (MWh) – EMS")
        split = wide[["Month", "NS_MWh_ems", "PEA_MWh_ems", "Mobile_MWh_ems"]].copy()
        split_long = split.melt("Month", var_name="metric", value_name="value").rename(columns={"Month": "x"})
        st.altair_chart(_line_chart_legend_toggle(split_long, "DG split (MWh) – EMS", "MWh"), use_container_width=True)

    st.caption("หมายเหตุ: Legend คลิกเพื่อซ่อน/แสดงเส้น (Altair selection)")

if __name__ == "__main__":
    main()
