# -*- coding: utf-8 -*-
"""
=============================================================================
CONTRACT LABOUR COMPLIANCE CONTROL TOWER
Version : 2.0  |  Deployment Ready
Developed for : HR Contractor Cell
Purpose : Upload CLM monthly Excel exports -> monitor vendor compliance,
          track issues per month, run analytics, generate Excel reports.
=============================================================================
"""

import io
import re
import sqlite3
import datetime

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =============================================================================
# CONSTANTS & CONFIGURATION
# =============================================================================

APP_TITLE   = "Contract Labour Compliance Control Tower"
APP_VERSION = "2.0"
DB_PATH     = "clm_control_tower.db"

MONTH_ORDER = {
    "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
    "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
    "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12,
}

REQUIRED_COLUMNS = [
    "Region", "Location", "Vendor Name", "V Code",
    "No of Employees", "Active gatepass", "Difference",
    "Wage Compliance", "PF Compliance", "ESIC Compliance",
    "Reason for Non-Compliance",
]

STATUS_OPTIONS = [
    "New",
    "Waiting for Vendor Documents",
    "Vendor Responded",
    "Under HR Review",
    "Under Contractor Cell Review",
    "PO Pending",
    "Medical Pending",
    "Safety Training Pending",
    "Gate Pass Under Process",
    "Escalated",
    "Resolved",
    "Closed",
]

OWNER_OPTIONS = [
    "HR", "Contractor Cell", "Vendor",
    "Procurement", "Security", "Medical", "Safety", "Finance",
]

HEALTH_BANDS = [
    (100, 100, "Healthy",         "#27ae60"),
    (75,   99, "Minor Issues",    "#f39c12"),
    (50,   74, "Needs Attention", "#e67e22"),
    (0,    49, "Critical",        "#e74c3c"),
]

# Reason keyword -> category mapping for root cause grouping
REASON_KEYWORDS = {
    "vendor not responding": "Vendor Not Responding",
    "vendor delay":          "Vendor Not Responding",
    "po pending":            "PO Pending",
    "po deleted":            "PO Pending",
    "po not released":       "PO Pending",
    "contract expired":      "Contract Expired",
    "clm":                   "CLM Implementation Pending",
    "medical":               "Medical Pending",
    "safety":                "Safety Training Pending",
    "gatepass":              "Gatepass Under Process",
    "wages not fed":         "Wage Submission Delay",
    "wage":                  "Wage Submission Delay",
    "vendor not finalised":  "Vendor Not Finalised",
    "punch not happening":   "Attendance/Punch Issue",
}

# =============================================================================
# PAGE CONFIG  (must be first Streamlit call)
# =============================================================================

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="CT",
    layout="wide",
    initial_sidebar_state="expanded",
)

# =============================================================================
# GLOBAL CSS
# =============================================================================

st.markdown("""
<style>
/* ── global ── */
html, body, [class*="css"] { font-family: 'Segoe UI', Arial, sans-serif; }
.stApp { background-color: #f0f2f6; color: #1a1a1a; }
p, span, li, td, th, div { color: #1a1a1a; }
h1, h2, h3, h4, h5 { color: #1a1a1a !important; font-weight: 700; }
[data-testid="stMetricValue"],
[data-testid="stMetricLabel"],
[data-testid="stCaptionContainer"] { color: #1a1a1a !important; }
[data-testid="stAlert"] p { color: #1a1a1a !important; }
[data-testid="stDataFrame"] td,
[data-testid="stDataFrame"] th { color: #1a1a1a !important; }
[data-testid="stSelectbox"] label,
[data-testid="stTextInput"] label,
[data-testid="stTextArea"] label { color: #1a1a1a !important; font-weight: 600; }
[data-testid="stExpander"] summary p { color: #1a1a1a !important; font-weight: 600; }

/* ── sidebar ── */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1a3c5e 0%, #2c6fad 100%);
}
[data-testid="stSidebar"] *,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] span,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] div,
[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] caption,
[data-testid="stSidebar"] [data-testid="stCaptionContainer"] {
    color: #ffffff !important;
}
[data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div {
    background: rgba(255,255,255,0.15);
    color: #ffffff;
    border: 1px solid rgba(255,255,255,0.3);
}

/* ── KPI cards ── */
.kpi-card {
    background: #ffffff;
    border-radius: 10px;
    padding: 18px 14px;
    text-align: center;
    border-left: 5px solid #2c6fad;
    box-shadow: 0 2px 8px rgba(0,0,0,.10);
    min-height: 90px;
}
.kpi-card.red    { border-left-color: #c0392b; }
.kpi-card.green  { border-left-color: #1e8449; }
.kpi-card.amber  { border-left-color: #d68910; }
.kpi-card.purple { border-left-color: #6c3483; }
.kpi-card.teal   { border-left-color: #117a65; }
.kpi-value { font-size: 1.85rem; font-weight: 800; color: #1a1a1a !important; line-height: 1.1; }
.kpi-label { font-size: .78rem; color: #333333 !important; margin-top: 6px; font-weight: 600; text-transform: uppercase; letter-spacing: .04em; }

/* ── section header ── */
.section-header {
    background: linear-gradient(90deg, #1a3c5e, #2c6fad);
    color: #ffffff !important;
    padding: 10px 18px;
    border-radius: 6px;
    font-weight: 700;
    font-size: 1rem;
    margin-bottom: 16px;
}

/* ── month banner ── */
.month-banner {
    background: #1a3c5e;
    color: #ffffff !important;
    padding: 8px 16px;
    border-radius: 6px;
    font-weight: 600;
    margin-bottom: 12px;
    font-size: 0.92rem;
}

/* ── tabs ── */
[data-baseweb="tab-list"] { background: #ffffff; border-radius: 8px; padding: 2px; }
[data-baseweb="tab"] { font-weight: 700; color: #1a1a1a !important; }
[data-baseweb="tab"][aria-selected="true"] { color: #1a3c5e !important; }

/* ── landing page ── */
.feature-box {
    background: #ffffff;
    border-radius: 8px;
    padding: 16px 20px;
    border-left: 4px solid #2c6fad;
    margin-bottom: 12px;
    box-shadow: 0 1px 4px rgba(0,0,0,.06);
}
.feature-box h4 { margin: 0 0 4px 0; color: #1a3c5e !important; }
.feature-box p  { margin: 0; font-size: .88rem; color: #444 !important; }
</style>
""", unsafe_allow_html=True)


# =============================================================================
# CHART THEME  -- applied to every figure
# =============================================================================

_CHART_FONT = dict(family="Segoe UI, Arial, sans-serif", color="#1a1a1a", size=12)
_AXIS_STYLE = dict(
    tickfont=dict(color="#1a1a1a", size=11),
    title_font=dict(color="#1a1a1a", size=12),
    linecolor="#cccccc",
    gridcolor="#e8e8e8",
)


def _apply_theme(fig, height=320, margin=None, show_legend=True):
    """Apply consistent readable theme to any Plotly figure."""
    if margin is None:
        margin = dict(t=48, b=16, l=16, r=16)
    fig.update_layout(
        height=height,
        margin=margin,
        paper_bgcolor="#ffffff",
        plot_bgcolor="#ffffff",
        font=_CHART_FONT,
        title_font=dict(color="#1a1a1a", size=14, family="Segoe UI, Arial, sans-serif"),
        legend=dict(font=dict(color="#1a1a1a", size=11), bgcolor="rgba(255,255,255,0.8)"),
        showlegend=show_legend,
        xaxis=_AXIS_STYLE,
        yaxis=_AXIS_STYLE,
    )
    return fig


# =============================================================================
# DATABASE
# =============================================================================

@st.cache_resource
def get_db():
    """Return singleton SQLite connection; create schema on first run."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    cur  = conn.cursor()
    cur.executescript("""
        CREATE TABLE IF NOT EXISTS TrackingData (
            id             INTEGER PRIMARY KEY AUTOINCREMENT,
            report_month   TEXT,
            vendor_name    TEXT,
            current_status TEXT DEFAULT 'New',
            owner          TEXT DEFAULT 'HR',
            remarks        TEXT DEFAULT '',
            last_updated   TEXT,
            UNIQUE(report_month, vendor_name)
        );
        CREATE TABLE IF NOT EXISTS AuditLog (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            ts           TEXT,
            report_month TEXT,
            vendor_name  TEXT,
            field        TEXT,
            old_value    TEXT,
            new_value    TEXT
        );
        CREATE TABLE IF NOT EXISTS UploadHistory (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            report_month TEXT UNIQUE,
            uploaded_at  TEXT,
            row_count    INTEGER,
            vendor_count INTEGER
        );
        CREATE TABLE IF NOT EXISTS MonthlyVendorData (
            id           INTEGER PRIMARY KEY AUTOINCREMENT,
            report_month TEXT UNIQUE,
            vendor_json  TEXT
        );
    """)
    conn.commit()
    return conn


def upsert_tracking(conn, report_month, vendor, status, owner, remarks):
    today = datetime.date.today().isoformat()
    cur   = conn.cursor()
    cur.execute(
        "SELECT current_status, owner, remarks FROM TrackingData "
        "WHERE report_month=? AND vendor_name=?",
        (report_month, vendor)
    )
    existing = cur.fetchone()
    if existing:
        for field, old, new in [
            ("current_status", existing[0], status),
            ("owner",          existing[1], owner),
            ("remarks",        existing[2], remarks),
        ]:
            if old != new:
                cur.execute(
                    "INSERT INTO AuditLog(ts,report_month,vendor_name,field,old_value,new_value)"
                    " VALUES(?,?,?,?,?,?)",
                    (today, report_month, vendor, field, old, new)
                )
        cur.execute(
            "UPDATE TrackingData SET current_status=?,owner=?,remarks=?,last_updated=?"
            " WHERE report_month=? AND vendor_name=?",
            (status, owner, remarks, today, report_month, vendor)
        )
    else:
        cur.execute(
            "INSERT INTO TrackingData(report_month,vendor_name,current_status,owner,remarks,last_updated)"
            " VALUES(?,?,?,?,?,?)",
            (report_month, vendor, status, owner, remarks, today)
        )
    conn.commit()


def get_tracking(conn, report_month):
    return pd.read_sql_query(
        "SELECT vendor_name, current_status, owner, remarks, last_updated "
        "FROM TrackingData WHERE report_month=?",
        conn, params=(report_month,)
    )


def save_vendor_snapshot(conn, report_month, vendor_df):
    conn.execute(
        "INSERT OR REPLACE INTO MonthlyVendorData(report_month, vendor_json) VALUES(?,?)",
        (report_month, vendor_df.to_json(orient="records"))
    )
    conn.commit()


def load_vendor_snapshot(conn, report_month):
    cur = conn.cursor()
    cur.execute("SELECT vendor_json FROM MonthlyVendorData WHERE report_month=?", (report_month,))
    row = cur.fetchone()
    if row:
        df = pd.read_json(io.StringIO(row[0]), orient="records")
        for col in ["No of Employees", "Active gatepass", "Difference", "Health Score"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        return df
    return None


def get_all_months(conn):
    cur = conn.cursor()
    cur.execute("SELECT report_month FROM MonthlyVendorData")
    months = [r[0] for r in cur.fetchall()]

    def _sort_key(m):
        parts = m.strip().split()
        if len(parts) == 2:
            return (int(parts[1]), MONTH_ORDER.get(parts[0][:3], 0))
        return (9999, 0)

    return sorted(months, key=_sort_key)


def log_upload(conn, report_month, row_count, vendor_count):
    conn.execute(
        "INSERT OR REPLACE INTO UploadHistory(report_month,uploaded_at,row_count,vendor_count)"
        " VALUES(?,?,?,?)",
        (report_month, datetime.datetime.now().isoformat(), row_count, vendor_count)
    )
    conn.commit()


def get_upload_history(conn):
    return pd.read_sql_query(
        "SELECT report_month, uploaded_at, row_count, vendor_count "
        "FROM UploadHistory ORDER BY uploaded_at DESC",
        conn
    )


# =============================================================================
# COLUMN FUZZY MATCHING
# =============================================================================

def _canonical(col):
    """Strip spaces/hyphens/underscores and lowercase for fuzzy match."""
    return re.sub(r"[\s\-_/]+", "", str(col).lower())


_REQ_CANON = {_canonical(c): c for c in REQUIRED_COLUMNS}


def normalise_columns(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
    rename = {}
    for col in df.columns:
        target = _REQ_CANON.get(_canonical(col))
        if target and col != target:
            rename[col] = target
    if rename:
        df.rename(columns=rename, inplace=True)
    return df


def validate_columns(df):
    return [c for c in REQUIRED_COLUMNS if c not in df.columns]


# =============================================================================
# FILE PARSING
# =============================================================================

def parse_upload(file):
    """
    Parse CLM Excel/CSV export.
    Returns dict {month_label: raw_dataframe} and sorted month list.
    Reads ALL monthly sheets; ignores Dashboard/Trend/Guideline sheets.
    Uses openpyxl engine so Excel formulas resolve to calculated values.
    """
    name = file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file)
        df = normalise_columns(df)
        return {"CSV Upload": df}, ["CSV Upload"]

    xls   = pd.ExcelFile(file, engine="openpyxl")
    months = [
        s for s in xls.sheet_names
        if re.match(r"^[A-Za-z]{3}\s+\d{2}$", s.strip())
    ]
    if not months:
        st.error("No monthly sheets detected. Expected names like Jan 26, Feb 26, etc.")
        st.stop()

    def _mk(m):
        parts = m.strip().split()
        return (int(parts[1]), MONTH_ORDER.get(parts[0][:3], 0)) if len(parts) == 2 else (9999, 0)

    sorted_months = sorted(months, key=_mk)
    month_data    = {}
    for m in sorted_months:
        df = pd.read_excel(file, sheet_name=m, engine="openpyxl")
        df = normalise_columns(df)
        df = df.dropna(subset=["Vendor Name"])
        month_data[m] = df

    return month_data, sorted_months


# =============================================================================
# BUSINESS LOGIC
# =============================================================================

def _complied(val):
    return str(val).strip().lower() == "complied"


def compute_health_score(row):
    s = 0
    if _complied(row.get("Wage Compliance", "")): s += 25
    if _complied(row.get("PF Compliance",   "")): s += 25
    if _complied(row.get("ESIC Compliance", "")): s += 25
    if row.get("Difference", 1) == 0:             s += 25
    return s


def health_label(score):
    for lo, hi, label, _ in HEALTH_BANDS:
        if lo <= score <= hi:
            return label
    return "Unknown"


def health_color(score):
    for lo, hi, _, color in HEALTH_BANDS:
        if lo <= score <= hi:
            return color
    return "#999"


def get_issues(row):
    issues = []
    if not _complied(row.get("Wage Compliance", "")): issues.append("Wage Non-Compliant")
    if not _complied(row.get("PF Compliance",   "")): issues.append("PF Non-Compliant")
    if not _complied(row.get("ESIC Compliance", "")): issues.append("ESIC Non-Compliant")
    d = row.get("Difference", 0)
    if d != 0: issues.append(f"Gatepass Gap ({int(d)})")
    return issues


def get_priority(issues):
    if not issues:                                                    return "-"
    if any(k in " ".join(issues) for k in ["Wage", "PF", "ESIC"]): return "Critical"
    if "Gatepass" in " ".join(issues):                              return "High"
    return "Medium"


def get_recommendation(row):
    recs = []
    if not _complied(row.get("PF Compliance",   "")): recs.append("Verify PF challan immediately")
    if not _complied(row.get("ESIC Compliance", "")): recs.append("Verify ESIC contributions")
    if not _complied(row.get("Wage Compliance", "")): recs.append("Obtain wage payment evidence")
    if row.get("Difference", 0) != 0:                 recs.append("Reconcile gatepass vs headcount")
    return " | ".join(recs) if recs else "No Action Required"


def categorise_reason(reason_text):
    """Map free-text reason to a standard category."""
    t = str(reason_text).lower()
    for kw, cat in REASON_KEYWORDS.items():
        if kw in t:
            return cat
    return "Other"


def build_vendor_df(raw):
    """Aggregate raw row-level data to one row per vendor."""
    agg = {
        "Region"                   : "first",
        "Location"                 : "first",
        "No of Employees"          : "sum",
        "Active gatepass"          : "sum",
        "Difference"               : "sum",
        "Wage Compliance"          : "first",
        "PF Compliance"            : "first",
        "ESIC Compliance"          : "first",
        "Reason for Non-Compliance": lambda x: " | ".join(
            [str(v) for v in x.dropna().unique() if str(v).strip() and str(v) != "nan"]
        ),
    }
    agg = {k: v for k, v in agg.items() if k in raw.columns}

    vendor = raw.groupby("Vendor Name", as_index=False).agg(agg)

    if "V Code" in raw.columns:
        vc     = raw.groupby("Vendor Name")["V Code"].first().reset_index()
        vendor = vendor.merge(vc, on="Vendor Name", how="left")

    vendor["Health Score"]       = vendor.apply(compute_health_score, axis=1)
    vendor["Status"]             = vendor["Health Score"].apply(health_label)
    vendor["Issues"]             = vendor.apply(lambda r: " | ".join(get_issues(r)) or "None", axis=1)
    vendor["Priority"]           = vendor.apply(lambda r: get_priority(get_issues(r)), axis=1)
    vendor["Recommended Action"] = vendor.apply(get_recommendation, axis=1)

    return vendor


def compute_kpis(vendor):
    total    = len(vendor)
    comp     = int((vendor["Health Score"] == 100).sum())
    emp      = int(vendor["No of Employees"].sum())
    risk     = int(vendor.loc[vendor["Health Score"] < 100, "No of Employees"].sum())
    gap      = int(vendor["Difference"].sum())
    critical = int((vendor["Health Score"] < 50).sum())
    pct      = round(comp / total * 100, 1) if total else 0.0
    return dict(total=total, comp=comp, non=total-comp, emp=emp,
                risk=risk, gap=gap, pct=pct, critical=critical)


def compute_aging(val):
    s = str(val).strip()
    if s in ("", "nan", "None"): return "Not updated"
    try:
        d = (datetime.date.today() - datetime.date.fromisoformat(s)).days
        if d == 0:  return "Today"
        if d == 1:  return "1 day"
        if d < 7:   return f"{d} days"
        if d < 30:  return f"{d // 7} week{'s' if d // 7 > 1 else ''}"
        return f"{d // 30} month{'s' if d // 30 > 1 else ''}"
    except Exception:
        return "-"


# =============================================================================
# MULTI-MONTH TREND BUILDER
# =============================================================================

def build_trend_df(conn, months):
    """Build month-over-month KPI dataframe from stored snapshots."""
    rows = []
    for m in months:
        v = load_vendor_snapshot(conn, m)
        if v is None:
            continue
        k = compute_kpis(v)
        wage_nc  = int((v["Wage Compliance"].str.lower().str.strip() != "complied").sum())
        pf_nc    = int((v["PF Compliance"].str.lower().str.strip()   != "complied").sum())
        esic_nc  = int((v["ESIC Compliance"].str.lower().str.strip() != "complied").sum())
        gate_nc  = int((v["Difference"] != 0).sum())
        rows.append({
            "Month": m, "Total": k["total"], "Compliant": k["comp"],
            "Non_Compliant": k["non"], "Compliance_Pct": k["pct"],
            "Employees": k["emp"], "At_Risk": k["risk"],
            "Gatepass_Gap": k["gap"],
            "Wage_NC": wage_nc, "PF_NC": pf_nc,
            "ESIC_NC": esic_nc, "Gatepass_NC": gate_nc,
        })
    return pd.DataFrame(rows)


# =============================================================================
# CHARTS  -- data-driven, based on actual field analysis
# =============================================================================

# -- DASHBOARD CHARTS --

def chart_compliance_donut(comp, non):
    """Overall compliant vs non-compliant split."""
    fig = go.Figure(go.Pie(
        labels=["Compliant", "Non-Compliant"],
        values=[comp, non],
        hole=0.62,
        marker_colors=["#27ae60", "#e74c3c"],
        textinfo="percent+label",
        textfont=dict(color="#1a1a1a", size=12),
        hovertemplate="%{label}: %{value} vendors<extra></extra>",
    ))
    fig.update_layout(
        title="Overall Compliance Split",
        annotations=[dict(text=f"<b>{comp+non}</b><br>Vendors",
                          x=0.5, y=0.5, font_size=13, showarrow=False,
                          font_color="#1a1a1a")],
    )
    return _apply_theme(fig)


def chart_health_distribution(vendor):
    """Vendor health band distribution -- 4 bands."""
    counts     = vendor["Status"].value_counts().reset_index()
    counts.columns = ["Status", "Count"]
    order      = ["Healthy", "Minor Issues", "Needs Attention", "Critical"]
    colour_map = {"Healthy": "#27ae60", "Minor Issues": "#f39c12",
                  "Needs Attention": "#e67e22", "Critical": "#e74c3c"}
    counts["Status"] = pd.Categorical(counts["Status"], categories=order, ordered=True)
    counts = counts.sort_values("Status")

    fig = px.bar(counts, x="Status", y="Count", color="Status",
                 color_discrete_map=colour_map,
                 title="Vendor Health Band Distribution",
                 text="Count")
    fig.update_traces(textposition="outside", textfont_color="#1a1a1a")
    fig.update_layout(showlegend=False)
    return _apply_theme(fig)


def chart_region_compliance(vendor):
    """
    Region-wise compliance rate (%) -- horizontal bar.
    Derived from: Region + all 3 compliance cols + Difference.
    Shows compliance % per region, making West's 0% immediately visible.
    """
    grp = vendor.groupby("Region").apply(
        lambda df: pd.Series({
            "Total":     len(df),
            "Compliant": (df["Health Score"] == 100).sum(),
            "Employees": int(df["No of Employees"].sum()),
        })
    ).reset_index()
    grp["Compliance_Pct"] = (grp["Compliant"] / grp["Total"] * 100).round(1)
    grp = grp.sort_values("Compliance_Pct")

    fig = px.bar(grp, y="Region", x="Compliance_Pct", orientation="h",
                 color="Compliance_Pct",
                 color_continuous_scale=["#e74c3c", "#f39c12", "#27ae60"],
                 range_color=[0, 100],
                 title="Region-wise Compliance Rate (%)",
                 text=grp["Compliance_Pct"].astype(str) + "%")
    fig.update_traces(textposition="outside", textfont_color="#1a1a1a")
    fig.update_layout(coloraxis_showscale=False, xaxis=dict(range=[0, 115]))
    return _apply_theme(fig, margin=dict(t=48, b=16, l=80, r=40))


def chart_component_compliance(vendor):
    """
    Wage / PF / ESIC / Gatepass -- grouped bar showing complied vs non-complied.
    Derived directly from 4 compliance fields. Reveals which component
    is the biggest driver of non-compliance.
    """
    total  = len(vendor)
    wage_c = (_complied_series(vendor["Wage Compliance"])).sum()
    pf_c   = (_complied_series(vendor["PF Compliance"])).sum()
    esic_c = (_complied_series(vendor["ESIC Compliance"])).sum()
    gate_c = (vendor["Difference"] == 0).sum()

    df = pd.DataFrame({
        "Component":     ["Wage", "PF", "ESIC", "Gatepass"],
        "Compliant":     [wage_c,       pf_c,   esic_c,   gate_c],
        "Non-Compliant": [total-wage_c, total-pf_c, total-esic_c, total-gate_c],
    })
    fig = go.Figure(data=[
        go.Bar(name="Compliant",     x=df["Component"], y=df["Compliant"],
               marker_color="#27ae60", text=df["Compliant"],
               textposition="outside", textfont=dict(color="#1a1a1a")),
        go.Bar(name="Non-Compliant", x=df["Component"], y=df["Non-Compliant"],
               marker_color="#e74c3c", text=df["Non-Compliant"],
               textposition="outside", textfont=dict(color="#1a1a1a")),
    ])
    fig.update_layout(barmode="group", title="Compliance by Component")
    return _apply_theme(fig)


def _complied_series(series):
    return series.astype(str).str.strip().str.lower() == "complied"


def chart_employees_at_risk(vendor):
    """
    Employees at risk by Region.
    Derived from No of Employees + Health Score < 100.
    Critical for understanding workforce exposure, not just vendor count.
    """
    emp = (vendor[vendor["Health Score"] < 100]
           .groupby("Region")["No of Employees"]
           .sum().reset_index(name="Employees at Risk"))
    emp = emp.sort_values("Employees at Risk", ascending=False)

    fig = px.bar(emp, x="Region", y="Employees at Risk",
                 color="Employees at Risk",
                 color_continuous_scale=["#f39c12", "#e74c3c"],
                 title="Employees at Risk by Region",
                 text="Employees at Risk")
    fig.update_traces(textposition="outside", textfont_color="#1a1a1a")
    fig.update_layout(coloraxis_showscale=False)
    return _apply_theme(fig)


def chart_gatepass_gap_location(vendor):
    """
    Gatepass gap (Difference) by Location -- top 10.
    Derived from Difference column and Location.
    Highlights where deployed headcount vs gatepass mismatch is worst.
    """
    loc = (vendor[vendor["Difference"] > 0]
           .groupby("Location")["Difference"]
           .sum().reset_index(name="Gatepass Gap")
           .sort_values("Gatepass Gap", ascending=False)
           .head(10))
    if loc.empty:
        return None

    fig = px.bar(loc, y="Location", x="Gatepass Gap", orientation="h",
                 color="Gatepass Gap",
                 color_continuous_scale=["#f39c12", "#e74c3c"],
                 title="Gatepass Gap by Location (Top 10)",
                 text="Gatepass Gap")
    fig.update_traces(textposition="outside", textfont_color="#1a1a1a")
    fig.update_layout(coloraxis_showscale=False,
                      yaxis=dict(autorange="reversed"))
    return _apply_theme(fig, height=340, margin=dict(t=48, b=16, l=140, r=40))


def chart_root_cause(vendor):
    """
    Categorised root cause -- from Reason for Non-Compliance field.
    Applies REASON_KEYWORDS map to normalise free text into categories.
    Shows actual operational reason distribution.
    """
    reasons = vendor["Reason for Non-Compliance"].dropna().astype(str)
    cats    = reasons.apply(categorise_reason).value_counts().reset_index()
    cats.columns = ["Category", "Count"]
    cats = cats[cats["Category"] != "nan"]
    if cats.empty:
        return None

    fig = px.bar(cats, y="Category", x="Count", orientation="h",
                 color="Count",
                 color_continuous_scale=["#f39c12", "#e74c3c"],
                 title="Root Cause Category Distribution",
                 text="Count")
    fig.update_traces(textposition="outside", textfont_color="#1a1a1a")
    fig.update_layout(coloraxis_showscale=False,
                      yaxis=dict(autorange="reversed"))
    return _apply_theme(fig, height=340, margin=dict(t=48, b=16, l=180, r=40))


def chart_priority_breakdown(vendor):
    """Critical / High / Medium non-compliant vendor count."""
    nc = vendor[vendor["Health Score"] < 100]
    if nc.empty:
        return None
    pc = nc["Priority"].value_counts().reset_index()
    pc.columns = ["Priority", "Count"]
    cmap = {"Critical": "#e74c3c", "High": "#e67e22", "Medium": "#f39c12"}
    fig = px.bar(pc, x="Priority", y="Count", color="Priority",
                 color_discrete_map=cmap, title="Non-Compliant Vendors by Priority",
                 text="Count")
    fig.update_traces(textposition="outside", textfont_color="#1a1a1a")
    fig.update_layout(showlegend=False)
    return _apply_theme(fig, height=300)


# -- ANALYTICS CHARTS --

def chart_trend_compliance(trend_df):
    """
    Month-over-month compliance % line chart.
    Derived from stored monthly snapshots.
    Shows whether compliance is improving or deteriorating.
    """
    if trend_df.empty or len(trend_df) < 2:
        return None
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=trend_df["Month"], y=trend_df["Compliance_Pct"],
        mode="lines+markers+text",
        text=trend_df["Compliance_Pct"].astype(str) + "%",
        textposition="top center",
        textfont=dict(color="#1a1a1a"),
        line=dict(color="#2c6fad", width=3),
        marker=dict(size=10, color="#1a3c5e"),
        name="Compliance %",
    ))
    fig.add_hline(y=80, line_dash="dash", line_color="#27ae60",
                  annotation_text="Target 80%",
                  annotation_font_color="#27ae60")
    fig.update_layout(title="Monthly Compliance % Trend",
                      yaxis=dict(range=[0, 110], ticksuffix="%"))
    return _apply_theme(fig, height=320)


def chart_trend_components(trend_df):
    """
    Month-over-month non-compliance per component (Wage/PF/ESIC/Gatepass).
    Shows which components are improving or worsening over time.
    """
    if trend_df.empty or len(trend_df) < 2:
        return None
    fig = go.Figure()
    components = [
        ("Wage_NC",    "Wage",     "#e74c3c"),
        ("PF_NC",      "PF",       "#e67e22"),
        ("ESIC_NC",    "ESIC",     "#f39c12"),
        ("Gatepass_NC","Gatepass", "#8e44ad"),
    ]
    for col, name, colour in components:
        if col in trend_df.columns:
            fig.add_trace(go.Scatter(
                x=trend_df["Month"], y=trend_df[col],
                mode="lines+markers",
                name=name,
                line=dict(color=colour, width=2),
                marker=dict(size=8),
            ))
    fig.update_layout(title="Non-Compliance Trend by Component")
    return _apply_theme(fig, height=320)


def chart_trend_employees(trend_df):
    """Total employees vs employees at risk per month."""
    if trend_df.empty or len(trend_df) < 2:
        return None
    fig = go.Figure(data=[
        go.Bar(name="Total Employees", x=trend_df["Month"],
               y=trend_df["Employees"], marker_color="#2c6fad",
               text=trend_df["Employees"], textposition="outside",
               textfont=dict(color="#1a1a1a")),
        go.Bar(name="At Risk",         x=trend_df["Month"],
               y=trend_df["At_Risk"],  marker_color="#e74c3c",
               text=trend_df["At_Risk"], textposition="outside",
               textfont=dict(color="#1a1a1a")),
    ])
    fig.update_layout(barmode="group", title="Employees: Total vs At Risk by Month")
    return _apply_theme(fig, height=320)


def chart_location_heatmap(vendor):
    """
    Location vs Compliance component heatmap.
    Derived from Location, Wage/PF/ESIC compliance + Difference.
    Shows exactly which location fails which component.
    """
    cols = {
        "Wage":     "Wage Compliance",
        "PF":       "PF Compliance",
        "ESIC":     "ESIC Compliance",
        "Gatepass": "Difference",
    }
    locs = vendor["Location"].unique()
    matrix = []
    for loc in locs:
        sub = vendor[vendor["Location"] == loc]
        row = {"Location": loc}
        row["Wage"]     = int((_complied_series(sub["Wage Compliance"])).sum())
        row["PF"]       = int((_complied_series(sub["PF Compliance"])).sum())
        row["ESIC"]     = int((_complied_series(sub["ESIC Compliance"])).sum())
        row["Gatepass"] = int((sub["Difference"] == 0).sum())
        row["Total"]    = len(sub)
        matrix.append(row)

    mat_df = pd.DataFrame(matrix)
    for col in ["Wage","PF","ESIC","Gatepass"]:
        mat_df[col] = (mat_df[col] / mat_df["Total"] * 100).round(0).astype(int)
    mat_df = mat_df.sort_values("Wage").head(15)

    z    = mat_df[["Wage","PF","ESIC","Gatepass"]].values.tolist()
    fig  = go.Figure(go.Heatmap(
        z=z,
        x=["Wage","PF","ESIC","Gatepass"],
        y=mat_df["Location"].tolist(),
        colorscale=[[0,"#e74c3c"],[0.5,"#f39c12"],[1,"#27ae60"]],
        zmin=0, zmax=100,
        text=[[f"{v}%" for v in row] for row in z],
        texttemplate="%{text}",
        textfont=dict(color="#1a1a1a", size=11),
        hovertemplate="Location: %{y}<br>Component: %{x}<br>Compliant: %{z}%<extra></extra>",
        colorbar=dict(tickfont=dict(color="#1a1a1a"),
                      title=dict(text="Compliance%", font=dict(color="#1a1a1a"))),
    ))
    fig.update_layout(title="Location x Component Compliance Heatmap (% Compliant)",
                      yaxis=dict(autorange="reversed"))
    return _apply_theme(fig, height=420, margin=dict(t=48, b=16, l=140, r=16))


def chart_scatter_employees_vs_score(vendor):
    """
    Employee count vs Health Score scatter -- bubble size = employees.
    Identifies high-employee vendors with low compliance (highest risk).
    """
    fig = px.scatter(
        vendor, x="No of Employees", y="Health Score",
        color="Status", size="No of Employees",
        hover_name="Vendor Name",
        hover_data={"Region": True, "Location": True,
                    "No of Employees": True, "Health Score": True},
        color_discrete_map={
            "Healthy": "#27ae60", "Minor Issues": "#f39c12",
            "Needs Attention": "#e67e22", "Critical": "#e74c3c",
        },
        title="Risk Matrix: Employee Exposure vs Health Score",
        labels={"No of Employees": "No of Employees",
                "Health Score": "Health Score (0-100)"},
    )
    fig.add_hline(y=75, line_dash="dot", line_color="#e67e22",
                  annotation_text="Score 75", annotation_font_color="#e67e22")
    fig.update_layout(yaxis=dict(range=[-5, 110]))
    return _apply_theme(fig, height=380)


def chart_vendor_ranking(vendor, top_n=15):
    """
    Bottom N vendors by health score -- identifies most critical vendors.
    Derived from Health Score field.
    """
    bottom = vendor.nsmallest(top_n, "Health Score")[
        ["Vendor Name", "Health Score", "Region", "Priority"]
    ].copy()
    bottom["Label"] = (bottom["Vendor Name"].str[:22] + " (" + bottom["Region"] + ")")
    colour_map = {"Critical": "#e74c3c", "High": "#e67e22", "Medium": "#f39c12", "-": "#95a5a6"}

    fig = px.bar(bottom, y="Label", x="Health Score",
                 color="Priority", color_discrete_map=colour_map,
                 orientation="h",
                 title=f"Bottom {top_n} Vendors by Health Score",
                 text="Health Score")
    fig.update_traces(textposition="outside", textfont_color="#1a1a1a")
    fig.update_layout(yaxis=dict(autorange="reversed"), xaxis=dict(range=[0, 120]))
    return _apply_theme(fig, height=420, margin=dict(t=48, b=16, l=220, r=40))


# =============================================================================
# EXCEL REPORT BUILDER
# =============================================================================

def _hdr(ws, row, cols, fill_hex="1a3c5e", font_hex="FFFFFF"):
    fill  = PatternFill("solid", fgColor=fill_hex)
    font  = Font(bold=True, color=font_hex, size=10)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin  = Border(left=Side(style="thin"), right=Side(style="thin"),
                   top=Side(style="thin"), bottom=Side(style="thin"))
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.fill = fill; cell.font = font
        cell.alignment = align; cell.border = thin


def _zebra(ws, start, end, cols, even_hex="EAF0F6"):
    fill = PatternFill("solid", fgColor=even_hex)
    for r in range(start, end+1):
        for c in range(1, cols+1):
            cell = ws.cell(row=r, column=c)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if r % 2 == 0:
                cell.fill = fill


def _auto_w(ws, mn=10, mx=42):
    for col in ws.columns:
        w = max((len(str(c.value or "")) for c in col), default=mn)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(w+2, mn), mx)


def _excel_bytes(wb):
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def sheet_executive_summary(wb, kpis, report_month):
    ws = wb.create_sheet("Executive Summary")
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:G1")
    ws["A1"] = f"Contract Labour Compliance - Executive Summary  |  {report_month}"
    ws["A1"].font      = Font(bold=True, size=13, color="FFFFFF")
    ws["A1"].fill      = PatternFill("solid", fgColor="1a3c5e")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    ws["A2"] = f"Generated: {datetime.date.today().strftime('%d %b %Y')}"
    ws["A2"].font = Font(italic=True, size=9, color="000000")

    headers = ["Total Vendors","Total Employees","Compliant","Non-Compliant",
               "Employees at Risk","Gatepass Gap","Compliance %"]
    values  = [kpis["total"], kpis["emp"], kpis["comp"], kpis["non"],
               kpis["risk"], kpis["gap"], f"{kpis['pct']}%"]
    colours = ["2c6fad","16a085","27ae60","e74c3c","e67e22","8e44ad","1e8449"]
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 34
    for i, (h, v, c) in enumerate(zip(headers, values, colours), start=1):
        ch = ws.cell(row=4, column=i, value=h)
        cv = ws.cell(row=5, column=i, value=v)
        ch.fill = PatternFill("solid", fgColor=c)
        ch.font = Font(bold=True, color="FFFFFF", size=9)
        ch.alignment = Alignment(horizontal="center", vertical="center")
        cv.font = Font(bold=True, size=15, color=c)
        cv.alignment = Alignment(horizontal="center", vertical="center")
    _auto_w(ws)
    return ws


def sheet_vendor_status(wb, vendor):
    ws = wb.create_sheet("Vendor Status")
    ws.sheet_view.showGridLines = False
    wanted = ["Vendor Name","V Code","Region","Location",
              "No of Employees","Active gatepass","Difference",
              "Health Score","Status","Issues","Priority","Recommended Action"]
    cols = [c for c in wanted if c in vendor.columns]
    for ci, h in enumerate(cols, 1):
        ws.cell(row=1, column=ci, value=h)
    _hdr(ws, 1, len(cols))
    s_col = {"Healthy":"27ae60","Minor Issues":"f39c12","Needs Attention":"e67e22","Critical":"e74c3c"}
    for ri, (_, row) in enumerate(vendor[cols].iterrows(), start=2):
        for ci, col in enumerate(cols, 1):
            cell = ws.cell(row=ri, column=ci, value=row[col])
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if col == "Status":
                cell.fill = PatternFill("solid", fgColor=s_col.get(str(row[col]),"999999"))
                cell.font = Font(bold=True, color="FFFFFF")
        if ri % 2 == 0:
            for ci in range(1, len(cols)+1):
                c2 = ws.cell(row=ri, column=ci)
                if not c2.fill.fgColor or c2.fill.fgColor.rgb in ("00000000","FFFFFFFF",""):
                    c2.fill = PatternFill("solid", fgColor="EAF0F6")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"
    _auto_w(ws)
    return ws


def sheet_action_center(wb, ac_df):
    ws = wb.create_sheet("Action Center")
    ws.sheet_view.showGridLines = False
    wanted = ["Vendor Name","Region","Issues","Priority",
              "Current Status","Owner","Remarks","Last Updated","Status Aging","Recommended Action"]
    cols = [c for c in wanted if c in ac_df.columns]
    for ci, h in enumerate(cols, 1):
        ws.cell(row=1, column=ci, value=h)
    _hdr(ws, 1, len(cols), fill_hex="c0392b")
    pf = {"Critical": Font(bold=True, color="c0392b"),
          "High":     Font(bold=True, color="e67e22"),
          "Medium":   Font(bold=True, color="d68910")}
    for ri, (_, row) in enumerate(ac_df[cols].iterrows(), start=2):
        for ci, col in enumerate(cols, 1):
            cell = ws.cell(row=ri, column=ci, value=row.get(col,""))
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            if col == "Priority":
                cell.font = pf.get(str(row.get(col,"")), Font(size=10))
        if ri % 2 == 0:
            for ci in range(1, len(cols)+1):
                ws.cell(row=ri, column=ci).fill = PatternFill("solid", fgColor="FDEDEC")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"
    _auto_w(ws)
    return ws


def build_full_report(vendor, kpis, ac_df, report_month):
    wb = Workbook(); wb.remove(wb.active)
    sheet_executive_summary(wb, kpis, report_month)
    sheet_vendor_status(wb, vendor)
    sheet_action_center(wb, ac_df)
    ws_raw = wb.create_sheet("Raw Data")
    ws_raw.sheet_view.showGridLines = False
    for ci, h in enumerate(vendor.columns, 1):
        ws_raw.cell(row=1, column=ci, value=h)
    _hdr(ws_raw, 1, len(vendor.columns), fill_hex="2c3e50")
    for ri, (_, row) in enumerate(vendor.iterrows(), start=2):
        for ci, val in enumerate(row, 1):
            ws_raw.cell(row=ri, column=ci, value=val)
    _zebra(ws_raw, 2, len(vendor)+1, len(vendor.columns))
    _auto_w(ws_raw)
    return _excel_bytes(wb)


# =============================================================================
# SIDEBAR
# =============================================================================

def render_sidebar(conn):
    with st.sidebar:
        st.markdown(f"### {APP_TITLE}")
        st.markdown(f"*Version {APP_VERSION}*")
        st.markdown("---")
        st.markdown("**Upload CLM Export**")
        uploaded = st.file_uploader(
            "Excel (.xlsx) or CSV (.csv)",
            type=["xlsx", "csv"],
            label_visibility="collapsed",
        )
        st.markdown("---")
        stored = get_all_months(conn)
        selected = None
        if stored:
            st.markdown("**Select Report Month**")
            selected = st.selectbox(
                "month", options=stored[::-1],
                label_visibility="collapsed", key="sel_month"
            )
            st.markdown("---")
            st.markdown("**Available Months**")
            for m in stored[::-1]:
                st.caption(f"- {m}")
        st.markdown("---")
        st.caption("Data stored in local SQLite\nOriginal files never modified\nEach month tracked independently")
    return uploaded, selected


# =============================================================================
# PAGE: LANDING
# =============================================================================

def page_landing(conn):
    st.markdown(f"## {APP_TITLE}")
    col_l, col_r = st.columns([3, 1])
    with col_l:
        st.markdown("### Getting Started")
        features = [
            ("Multi-Month Tracking",
             "Upload your CLM Excel workbook once. All monthly sheets (Jan 26, Feb 26 …) "
             "are read and stored separately. Switch between months using the sidebar."),
            ("Automated Compliance Scoring",
             "Each vendor receives a Health Score (0-100) based on Wage, PF, ESIC, "
             "and Gatepass Difference. Rule engine classifies: Healthy / Minor Issues / "
             "Needs Attention / Critical."),
            ("Action Center",
             "Per-vendor status tracking with dropdown for Current Status, Owner, "
             "and Remarks. Every change is audit-logged with timestamp."),
            ("Data-Driven Analytics",
             "9 charts derived from your actual CLM fields: compliance trend, "
             "component breakdown, gatepass gap, root cause, location heatmap, "
             "and risk scatter."),
            ("One-Click Reports",
             "Download Executive Summary, Vendor Status, Action Center, or Complete "
             "Report as formatted Excel files."),
        ]
        for title, desc in features:
            st.markdown(
                f'<div class="feature-box"><h4>{title}</h4><p>{desc}</p></div>',
                unsafe_allow_html=True
            )
    with col_r:
        st.markdown("#### Quick Stats")
        hist = get_upload_history(conn)
        if not hist.empty:
            st.metric("Months Stored", len(hist))
            st.metric("Latest Month",  hist.iloc[0]["report_month"])
            st.metric("Last Uploaded", hist.iloc[0]["uploaded_at"][:10])
        else:
            st.info("No data uploaded yet.\n\nUpload your CLM Excel file using the sidebar to begin.")
        st.markdown("---")
        st.markdown("**Required Columns**")
        for c in REQUIRED_COLUMNS:
            st.caption(f"- {c}")


# =============================================================================
# TAB: DASHBOARD
# =============================================================================

def tab_dashboard(vendor, kpis, month):
    sh = lambda t: st.markdown(f'<div class="section-header">{t}</div>', unsafe_allow_html=True)
    sh(f"Executive Dashboard - {month}")

    # KPI row
    cols = st.columns(7)
    kpi_data = [
        ("Total Vendors",     kpis["total"],    "blue"),
        ("Total Employees",   kpis["emp"],      "blue"),
        ("Compliant",         kpis["comp"],     "green"),
        ("Non-Compliant",     kpis["non"],      "red"),
        ("Employees at Risk", kpis["risk"],     "amber"),
        ("Gatepass Gap",      kpis["gap"],      "purple"),
        ("Compliance %",      f"{kpis['pct']}%",
         "green" if kpis["pct"] >= 80 else "amber"),
    ]
    for col, (label, val, colour) in zip(cols, kpi_data):
        col.markdown(
            f'<div class="kpi-card {colour}">'
            f'<div class="kpi-value">{val}</div>'
            f'<div class="kpi-label">{label}</div></div>',
            unsafe_allow_html=True,
        )
    st.markdown("<br>", unsafe_allow_html=True)

    # Row 1: donut + health bands
    c1, c2 = st.columns(2)
    with c1:
        st.plotly_chart(chart_compliance_donut(kpis["comp"], kpis["non"]),
                        use_container_width=True, key="d_donut")
    with c2:
        st.plotly_chart(chart_health_distribution(vendor),
                        use_container_width=True, key="d_health")

    # Row 2: region compliance % + component grouped bar
    c3, c4 = st.columns(2)
    with c3:
        st.plotly_chart(chart_region_compliance(vendor),
                        use_container_width=True, key="d_region")
    with c4:
        st.plotly_chart(chart_component_compliance(vendor),
                        use_container_width=True, key="d_component")

    # Row 3: employees at risk + gatepass gap by location
    c5, c6 = st.columns(2)
    with c5:
        if not vendor[vendor["Health Score"] < 100].empty:
            st.plotly_chart(chart_employees_at_risk(vendor),
                            use_container_width=True, key="d_emp_risk")
    with c6:
        fig_gp = chart_gatepass_gap_location(vendor)
        if fig_gp:
            st.plotly_chart(fig_gp, use_container_width=True, key="d_gatepass")

    # Row 4: root cause + priority
    c7, c8 = st.columns(2)
    with c7:
        fig_rc = chart_root_cause(vendor)
        if fig_rc:
            st.plotly_chart(fig_rc, use_container_width=True, key="d_root_cause")
    with c8:
        fig_pb = chart_priority_breakdown(vendor)
        if fig_pb:
            st.plotly_chart(fig_pb, use_container_width=True, key="d_priority")

    # Region summary table
    with st.expander("Region-wise Summary Table"):
        grp = vendor.groupby("Region").agg(
            Vendors=("Vendor Name", "count"),
            Compliant=("Health Score", lambda x: (x==100).sum()),
            Non_Compliant=("Health Score", lambda x: (x<100).sum()),
            Employees=("No of Employees", "sum"),
            At_Risk=("No of Employees", lambda x: x[vendor.loc[x.index,"Health Score"]<100].sum()),
            Gatepass_Gap=("Difference","sum"),
        ).reset_index()
        grp["Compliance %"] = (grp["Compliant"]/grp["Vendors"]*100).round(1).astype(str)+"%"
        st.dataframe(grp, use_container_width=True, hide_index=True)


# =============================================================================
# TAB: VENDOR STATUS
# =============================================================================

def tab_vendor_status(vendor):
    st.markdown('<div class="section-header">Vendor Status</div>', unsafe_allow_html=True)
    fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 3])
    regions    = ["All"] + sorted(vendor["Region"].dropna().unique().tolist())
    statuses   = ["All","Healthy","Minor Issues","Needs Attention","Critical"]
    priorities = ["All","Critical","High","Medium","-"]

    rf = fc1.selectbox("Region",   regions,    key="vs_reg")
    sf = fc2.selectbox("Status",   statuses,   key="vs_sta")
    pf = fc3.selectbox("Priority", priorities, key="vs_pri")
    q  = fc4.text_input("Search Vendor Name", key="vs_q")

    filt = vendor.copy()
    if rf != "All": filt = filt[filt["Region"]   == rf]
    if sf != "All": filt = filt[filt["Status"]   == sf]
    if pf != "All": filt = filt[filt["Priority"] == pf]
    if q:           filt = filt[filt["Vendor Name"].str.contains(q, case=False, na=False)]

    st.caption(f"Showing **{len(filt)}** of **{len(vendor)}** vendors")
    show = [c for c in [
        "Vendor Name","V Code","Region","Location","No of Employees",
        "Active gatepass","Difference","Health Score","Status","Priority",
        "Issues","Wage Compliance","PF Compliance","ESIC Compliance",
        "Reason for Non-Compliance",
    ] if c in filt.columns]
    st.dataframe(filt[show].reset_index(drop=True), use_container_width=True, height=480)

    wb = Workbook(); wb.remove(wb.active)
    sheet_vendor_status(wb, filt)
    st.download_button(
        "Download Filtered Vendor Status (Excel)",
        data=_excel_bytes(wb),
        file_name=f"vendor_status_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =============================================================================
# TAB: ACTION CENTER
# =============================================================================

def tab_action_center(vendor, month, conn):
    st.markdown('<div class="section-header">Action Center - Non-Compliant Vendors</div>',
                unsafe_allow_html=True)

    nc = vendor[vendor["Health Score"] < 100].copy()
    if nc.empty:
        st.success("All vendors are fully compliant. No actions required.")
        return

    tracking = get_tracking(conn, month)
    if not tracking.empty:
        nc = nc.merge(tracking.rename(columns={"vendor_name":"Vendor Name"}),
                      on="Vendor Name", how="left")

    for col in ["current_status","owner","remarks","last_updated"]:
        if col not in nc.columns:
            nc[col] = None
    nc["current_status"] = nc["current_status"].fillna("New")
    nc["owner"]          = nc["owner"].fillna("HR")
    nc["remarks"]        = nc["remarks"].fillna("")
    nc["last_updated"]   = nc["last_updated"].fillna("")
    nc["Status Aging"]   = nc["last_updated"].apply(compute_aging)

    f1, f2 = st.columns([2, 3])
    pf = f1.selectbox("Filter by Priority", ["All","Critical","High","Medium"], key="ac_pri")
    sf = f2.selectbox("Filter by Status",   ["All"]+STATUS_OPTIONS, key="ac_sta")
    view = nc.copy()
    if pf != "All": view = view[view["Priority"]       == pf]
    if sf != "All": view = view[view["current_status"] == sf]

    st.caption(f"**{len(view)}** vendor(s) requiring attention")
    st.markdown("---")

    saved = False
    for _, row in view.sort_values(["Priority","Vendor Name"]).iterrows():
        vn  = row["Vendor Name"]
        tag = {"Critical":"[CRITICAL]","High":"[HIGH]","Medium":"[MEDIUM]"}.get(row["Priority"],"[--]")
        with st.expander(f"{tag}  {vn}  |  {row['Region']}  |  Score: {row['Health Score']}"):
            a, b = st.columns([3, 2])
            with a:
                st.write(f"**Issues:** {row['Issues']}")
                st.write(f"**Recommendation:** {row['Recommended Action']}")
                st.write(f"**Root Cause:** {row.get('Reason for Non-Compliance', '-')}")
                st.write(f"**Last Updated:** {row['last_updated'] or 'Never'}  |  **Aging:** {row['Status Aging']}")
            with b:
                kp = f"{month}_{vn}".replace(" ", "_")
                cs = st.selectbox("Current Status", STATUS_OPTIONS,
                    index=STATUS_OPTIONS.index(row["current_status"])
                          if row["current_status"] in STATUS_OPTIONS else 0,
                    key=f"s_{kp}")
                co = st.selectbox("Owner", OWNER_OPTIONS,
                    index=OWNER_OPTIONS.index(row["owner"])
                          if row["owner"] in OWNER_OPTIONS else 0,
                    key=f"o_{kp}")
                cr = st.text_area("Remarks", value=row["remarks"],
                                   key=f"r_{kp}", height=80)
                if st.button("Save", key=f"sv_{kp}"):
                    upsert_tracking(conn, month, vn, cs, co, cr)
                    st.success(f"Saved - {vn}")
                    saved = True
    if saved:
        st.rerun()

    st.markdown("---")
    ac_dl = view[[c for c in [
        "Vendor Name","Region","Issues","Priority",
        "current_status","owner","remarks","last_updated",
        "Status Aging","Recommended Action"
    ] if c in view.columns]].rename(columns={
        "current_status":"Current Status","owner":"Owner",
        "remarks":"Remarks","last_updated":"Last Updated",
    })
    wb = Workbook(); wb.remove(wb.active)
    sheet_action_center(wb, ac_dl)
    st.download_button(
        "Download Action Center (Excel)",
        data=_excel_bytes(wb),
        file_name=f"action_center_{month.replace(' ','_')}_{datetime.date.today()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# =============================================================================
# TAB: ANALYTICS
# =============================================================================

def tab_analytics(vendor, month, conn):
    st.markdown('<div class="section-header">Analytics and Deep Dive</div>',
                unsafe_allow_html=True)
    all_months = get_all_months(conn)
    trend_df   = build_trend_df(conn, all_months)

    # ── Trend section (only when 2+ months available) ──
    if len(all_months) >= 2:
        st.markdown("**Month-over-Month Trends**")
        t1, t2 = st.columns(2)
        with t1:
            fig = chart_trend_compliance(trend_df)
            if fig: st.plotly_chart(fig, use_container_width=True, key="an_trend_comp")
        with t2:
            fig = chart_trend_components(trend_df)
            if fig: st.plotly_chart(fig, use_container_width=True, key="an_trend_comp2")
        fig = chart_trend_employees(trend_df)
        if fig: st.plotly_chart(fig, use_container_width=True, key="an_trend_emp")
        st.markdown("---")

    # ── Current month deep dive ──
    st.markdown(f"**Deep Dive - {month}**")
    r1, r2 = st.columns(2)
    with r1:
        st.plotly_chart(chart_scatter_employees_vs_score(vendor),
                        use_container_width=True, key="an_scatter")
    with r2:
        st.plotly_chart(chart_vendor_ranking(vendor),
                        use_container_width=True, key="an_ranking")

    # Location heatmap
    st.plotly_chart(chart_location_heatmap(vendor),
                    use_container_width=True, key="an_heatmap")

    # Root cause + component
    rc1, rc2 = st.columns(2)
    with rc1:
        fig = chart_root_cause(vendor)
        if fig: st.plotly_chart(fig, use_container_width=True, key="an_root")
    with rc2:
        st.plotly_chart(chart_component_compliance(vendor),
                        use_container_width=True, key="an_component")

    # Trend table
    if not trend_df.empty:
        with st.expander("Month-over-Month Data Table"):
            st.dataframe(trend_df, use_container_width=True, hide_index=True)


# =============================================================================
# TAB: REPORTS
# =============================================================================

def tab_reports(vendor, kpis, month, conn):
    st.markdown('<div class="section-header">Reports and Downloads</div>',
                unsafe_allow_html=True)
    st.info("Reports are generated from current data. Original uploaded files are never modified.")

    # Prepare action center data
    nc = vendor[vendor["Health Score"] < 100].copy()
    tr = get_tracking(conn, month)
    if not tr.empty and not nc.empty:
        nc = nc.merge(tr.rename(columns={"vendor_name":"Vendor Name"}),
                      on="Vendor Name", how="left")
    for col in ["current_status","owner","remarks","last_updated"]:
        if col not in nc.columns: nc[col] = "-"
    nc["current_status"] = nc["current_status"].fillna("New")
    nc["Status Aging"]   = nc.get("last_updated","").apply(compute_aging)
    ac_dl = nc.rename(columns={"current_status":"Current Status","owner":"Owner",
                                "remarks":"Remarks","last_updated":"Last Updated"})

    st.markdown("---")
    r1, r2 = st.columns(2)
    with r1:
        st.markdown("#### Executive Summary")
        st.caption("KPI cards snapshot for management.")
        wb = Workbook(); wb.remove(wb.active)
        sheet_executive_summary(wb, kpis, month)
        st.download_button("Download Executive Summary", data=_excel_bytes(wb),
            file_name=f"exec_summary_{month.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with r2:
        st.markdown("#### Vendor Status Report")
        st.caption("Full vendor list with health scores and compliance breakdown.")
        wb = Workbook(); wb.remove(wb.active)
        sheet_vendor_status(wb, vendor)
        st.download_button("Download Vendor Status", data=_excel_bytes(wb),
            file_name=f"vendor_status_{month.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    st.markdown("---")
    r3, r4 = st.columns(2)
    with r3:
        st.markdown("#### Action Center Report")
        st.caption("Non-compliant vendors with status, owner, and remarks.")
        wb = Workbook(); wb.remove(wb.active)
        sheet_action_center(wb, ac_dl)
        st.download_button("Download Action Center", data=_excel_bytes(wb),
            file_name=f"action_center_{month.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)
    with r4:
        st.markdown("#### Complete Report (All Sheets)")
        st.caption("Executive Summary + Vendor Status + Action Center + Raw Data.")
        st.download_button("Download Complete Report",
            data=build_full_report(vendor, kpis, ac_dl, month),
            file_name=f"complete_report_{month.replace(' ','_')}_{datetime.date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True)

    # Audit log
    st.markdown("---")
    st.markdown("#### Audit Log")
    audit = pd.read_sql_query(
        "SELECT ts,report_month,vendor_name,field,old_value,new_value "
        "FROM AuditLog ORDER BY ts DESC LIMIT 300", conn)
    if audit.empty:
        st.caption("No changes recorded yet.")
    else:
        st.dataframe(audit, use_container_width=True, height=260)
        st.download_button("Download Audit Log (CSV)",
            data=audit.to_csv(index=False).encode(), file_name="audit_log.csv", mime="text/csv")


# =============================================================================
# TAB: SETTINGS
# =============================================================================

def tab_settings(conn):
    st.markdown('<div class="section-header">Settings and Administration</div>',
                unsafe_allow_html=True)

    st.markdown("#### Upload History")
    hist = get_upload_history(conn)
    if hist.empty:
        st.caption("No uploads yet.")
    else:
        st.dataframe(hist, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("#### Business Rules Reference")
    st.markdown("""
| Component | Score Contribution | Non-Complied Classification |
|---|---|---|
| Wage Compliance = Complied | +25 | Critical |
| PF Compliance = Complied | +25 | Critical |
| ESIC Compliance = Complied | +25 | Critical |
| Difference = 0 | +25 | High |

| Health Score | Band |
|---|---|
| 100 | Healthy |
| 75 to 99 | Minor Issues |
| 50 to 74 | Needs Attention |
| 0 to 49 | Critical |
    """)

    st.markdown("---")
    st.markdown("#### Required Columns")
    st.code("\n".join(REQUIRED_COLUMNS))

    st.markdown("---")
    st.markdown("#### Clear Tracking Data")
    st.warning("Deletes all status updates and audit log. Cannot be undone.")
    if st.button("Clear Tracking Data and Audit Log", type="secondary"):
        conn.execute("DELETE FROM TrackingData")
        conn.execute("DELETE FROM AuditLog")
        conn.commit()
        st.success("Tracking data cleared.")
        st.rerun()

    st.markdown("---")
    st.markdown("#### Clear All Stored Data")
    st.error("This removes ALL monthly snapshots. You will need to re-upload all files.")
    if st.button("Clear All Data (Full Reset)", type="secondary"):
        conn.execute("DELETE FROM TrackingData")
        conn.execute("DELETE FROM AuditLog")
        conn.execute("DELETE FROM MonthlyVendorData")
        conn.execute("DELETE FROM UploadHistory")
        conn.commit()
        st.success("All data cleared. Please re-upload your CLM files.")
        st.rerun()


# =============================================================================
# MAIN
# =============================================================================

def main():
    conn               = get_db()
    uploaded, selected = render_sidebar(conn)

    # ── Process new upload ──
    if uploaded is not None:
        with st.spinner("Reading file and processing all monthly sheets..."):
            month_data, sorted_months = parse_upload(uploaded)

        errors  = []
        success = []
        for month, raw_df in month_data.items():
            missing = validate_columns(raw_df)
            if missing:
                errors.append(f"Sheet '{month}': missing columns - {', '.join(missing)}")
                continue
            vendor_df = build_vendor_df(raw_df)
            save_vendor_snapshot(conn, month, vendor_df)
            log_upload(conn, month, len(raw_df), len(vendor_df))
            success.append(month)

        if success:
            st.success(
                f"Upload complete. {len(success)} sheet(s) stored: {', '.join(success)}. "
                "Use the month selector in the sidebar to switch months."
            )
        for e in errors:
            st.error(e)

    # ── Check if any data exists ──
    stored = get_all_months(conn)
    if not stored:
        page_landing(conn)
        st.stop()

    # ── Resolve selected month ──
    if not selected or selected not in stored:
        selected = stored[-1]

    vendor = load_vendor_snapshot(conn, selected)
    if vendor is None:
        st.error(f"Could not load data for {selected}. Please re-upload.")
        st.stop()

    kpis = compute_kpis(vendor)

    # Month banner
    st.markdown(
        f'<div class="month-banner">'
        f'Viewing: <b>{selected}</b> &nbsp;|&nbsp; '
        f'Vendors: {kpis["total"]} &nbsp;|&nbsp; '
        f'Compliance: {kpis["pct"]}% &nbsp;|&nbsp; '
        f'Non-Compliant: {kpis["non"]} &nbsp;|&nbsp; '
        f'Critical: {kpis["critical"]} &nbsp;|&nbsp; '
        f'Gatepass Gap: {kpis["gap"]}'
        f'</div>',
        unsafe_allow_html=True,
    )

    tabs = st.tabs(["Dashboard","Vendor Status","Action Center","Analytics","Reports","Settings"])
    with tabs[0]: tab_dashboard(vendor, kpis, selected)
    with tabs[1]: tab_vendor_status(vendor)
    with tabs[2]: tab_action_center(vendor, selected, conn)
    with tabs[3]: tab_analytics(vendor, selected, conn)
    with tabs[4]: tab_reports(vendor, kpis, selected, conn)
    with tabs[5]: tab_settings(conn)


if __name__ == "__main__":
    main()
