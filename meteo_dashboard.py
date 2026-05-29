import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import numpy as np
import re
import csv
import hashlib
from io import BytesIO
from pathlib import Path
from datetime import datetime, time
from fpdf import FPDF
import smtplib
from email.message import EmailMessage

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Meteorological Operations Suite",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# =========================================================
# STYLING
# =========================================================
st.markdown("""
<style>
.block-container {
    padding-top: 2.4rem;
    padding-bottom: 1rem;
    max-width: 97%;
}
.hero-panel {
    margin-top: 0.8rem;
}
:root {
    --bg: #EFF4F9;
    --surface: #FFFFFF;
    --surface-soft: #F8FBFE;
    --line: #D9E2EC;
    --line-soft: #E9EEF5;
    --text: #172B4D;
    --muted: #6B778C;
    --title: #0F1C2E;
    --primary: #3FA9F5;
    --primary-dark: #1F5E8C;
    --shadow-sm: 0 1px 2px rgba(15, 23, 42, 0.05);
    --shadow-md: 0 10px 24px rgba(15, 23, 42, 0.07);
    --radius-lg: 22px;
}
html, body, [class*="css"] {
    font-family: "Segoe UI", Arial, sans-serif;
    color: var(--text);
}
.stApp {
    background:
        linear-gradient(180deg, rgba(63,169,245,0.12) 0%, rgba(239,244,249,1) 240px),
        var(--bg);
}

#MainMenu, footer {
    visibility: hidden !important;
}
.main-shell {
    margin-bottom: 0.8rem;
}
.hero-panel {
    background: linear-gradient(135deg, rgba(255,255,255,0.98) 0%, rgba(249,252,255,0.99) 100%);
    border: 1px solid rgba(217,226,236,0.95);
    border-radius: 28px;
    padding: 20px 22px 16px 22px;
    box-shadow: var(--shadow-md);
    margin-bottom: 0.75rem;
}
.hero-title {
    font-size: 1.5rem;
    font-weight: 760;
    color: var(--title);
    margin-bottom: 0.15rem;
    letter-spacing: -0.02em;
}
.hero-subtitle {
    font-size: 0.93rem;
    color: var(--muted);
    line-height: 1.5;
    max-width: 980px;
    margin-bottom: 0.75rem;
}
.pill {
    display: inline-block;
    padding: 6px 12px;
    border-radius: 999px;
    font-size: 0.75rem;
    font-weight: 650;
    margin-right: 8px;
    margin-top: 4px;
    border: 1px solid var(--line-soft);
    background: #FFFFFF;
    color: var(--text);
}
.section-label {
    font-size: 1rem;
    font-weight: 720;
    color: var(--title);
    margin: 0.2rem 0 0.45rem 0;
}
.subtle-note {
    background: rgba(255,255,255,0.78);
    border: 1px solid var(--line);
    border-left: 4px solid var(--primary);
    color: var(--text);
    padding: 10px 12px;
    border-radius: 12px;
    font-size: 0.88rem;
    margin-bottom: 0.8rem;
    box-shadow: var(--shadow-sm);
}
.site-shell {
    background: linear-gradient(180deg, rgba(255,255,255,0.94) 0%, rgba(252,253,254,0.98) 100%);
    border: 1px solid var(--line);
    border-radius: 22px;
    padding: 12px 12px 10px 12px;
    box-shadow: var(--shadow-md);
    margin-bottom: 0.9rem;
}
.site-title {
    font-size: 1.04rem;
    font-weight: 740;
    color: var(--title);
    margin-bottom: 0.08rem;
}
.site-subtitle {
    color: var(--muted);
    font-size: 0.82rem;
}
.metric-chip {
    display: inline-block;
    padding: 4px 10px;
    border-radius: 999px;
    font-size: 0.73rem;
    font-weight: 650;
    margin-right: 6px;
    margin-top: 6px;
    border: 1px solid var(--line-soft);
    background: #FFFFFF;
    color: #334155;
}
.range-badge {
    display: inline-block;
    padding: 4px 10px;
    border-radius: 999px;
    background: rgba(63,169,245,0.10);
    color: var(--primary-dark);
    font-size: 0.74rem;
    font-weight: 700;
    border: 1px solid rgba(63,169,245,0.18);
}
.helper-note {
    background: #F8FBFE;
    border: 1px dashed #C8D9EA;
    color: #46627F;
    padding: 9px 11px;
    border-radius: 12px;
    font-size: 0.82rem;
    margin: 0.25rem 0 0.7rem 0;
}
.small-panel {
    background: linear-gradient(180deg, #FFFFFF 0%, #FCFDFE 100%);
    border: 1px solid var(--line);
    border-radius: 18px;
    padding: 10px 12px;
    box-shadow: var(--shadow-sm);
}
.small-title {
    font-size: 0.84rem;
    font-weight: 700;
    color: var(--title);
    margin-bottom: 0.35rem;
}
.small-caption {
    color: var(--muted);
    font-size: 0.76rem;
    margin-bottom: 0.45rem;
}
div[data-testid="stMetric"] {
    background: linear-gradient(180deg, #FFFFFF 0%, #FBFCFE 100%);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 8px 10px;
    box-shadow: var(--shadow-sm);
}
div[data-testid="stMetricLabel"] {
    color: var(--muted) !important;
    font-weight: 650 !important;
}
div[data-testid="stMetricValue"] {
    color: var(--title) !important;
    font-weight: 760 !important;
}
div[data-testid="stDataFrame"],
div[data-testid="stExpander"] {
    border: 1px solid var(--line);
    border-radius: 14px;
    overflow: hidden;
    background: #FFFFFF;
}
.stButton > button,
.stDownloadButton > button,
.stFormSubmitButton > button {
    border-radius: 12px !important;
    border: 1px solid var(--line) !important;
    background: #FFFFFF !important;
    color: var(--text) !important;
    font-weight: 650 !important;
    box-shadow: none !important;
}
.stFormSubmitButton > button {
    background: linear-gradient(135deg, #52B6F6 0%, #2F95DA 100%) !important;
    color: white !important;
    border-color: #52B6F6 !important;
}
div[data-testid="stRadio"] > div {
    flex-direction: row;
    gap: 8px;
}
div[data-testid="stRadio"] label {
    background: #FFFFFF;
    border: 1px solid var(--line);
    padding: 6px 10px;
    border-radius: 10px;
}
hr {
    border: none;
    border-top: 1px solid var(--line-soft);
    margin: 0.55rem 0;
}
</style>
""", unsafe_allow_html=True)

# =========================================================
# PATHS
# =========================================================
BASE_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
UPLOAD_DIR = BASE_DIR / "data_uploads"
PROCESSED_DIR = BASE_DIR / "processed"
REPORT_DIR = BASE_DIR / "Meteorological Reports" / "Daily"
LOG_FILE = BASE_DIR / "upload_log.csv"

for folder in [UPLOAD_DIR, PROCESSED_DIR, REPORT_DIR]:
    folder.mkdir(parents=True, exist_ok=True)

# =========================================================
# CONSTANTS
# =========================================================
METRIC_PATTERN = re.compile(
    r"^(.*?)\s(Wind Speed|Wind Direction|Irradiance|Precipitation)\s\[(.*?)\]$",
    flags=re.IGNORECASE
)

COLOR_MAP = {
    "Wind Speed": "#58AEE8",
    "Wind Direction": "#F1A55B",
    "Irradiance": "#4CC9B0",
    "Precipitation": "#9B7BF2"
}

# =========================================================
# HELPERS
# =========================================================
def send_test_email_outlook(to_email):
    sender = st.secrets["OUTLOOK_SENDER"]
    password = st.secrets["OUTLOOK_PASSWORD"]

    msg = EmailMessage()
    msg["Subject"] = "Test email from Meteorological Dashboard"
    msg["From"] = sender
    msg["To"] = to_email
    msg.set_content("This is a test email sent from your hosted Streamlit app.")

    with smtplib.SMTP("smtp-mail.outlook.com", 587) as server:
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(sender, password)
        server.send_message(msg)

def safe_filename(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", str(name))

def get_file_hash(file_bytes: bytes) -> str:
    return hashlib.md5(file_bytes).hexdigest()

def append_upload_log(record: dict):
    df_new = pd.DataFrame([record])
    if LOG_FILE.exists():
        try:
            df_old = pd.read_csv(LOG_FILE)
            df_all = pd.concat([df_old, df_new], ignore_index=True)
        except Exception:
            df_all = df_new
    else:
        df_all = df_new
    df_all.to_csv(LOG_FILE, index=False)

def load_upload_log():
    if LOG_FILE.exists():
        return pd.read_csv(LOG_FILE)
    return pd.DataFrame()

def split_variable(col_name):
    text = str(col_name).strip()
    match = METRIC_PATTERN.match(text)
    if match:
        site = match.group(1).strip()
        metric = match.group(2).strip().title()
        unit = match.group(3).strip()
        return site, metric, unit
    return text, "Value", ""

def parse_semicolon_table(lines):
    clean_lines = [ln for ln in lines if str(ln).strip()]
    if not clean_lines:
        raise ValueError("No readable lines found in the uploaded file.")

    reader = csv.reader(clean_lines, delimiter=";")
    rows = [list(r) for r in reader if r]

    if len(rows) < 2:
        raise ValueError("The file does not contain enough rows to parse.")

    header = [str(h).strip() for h in rows[0]]
    parsed_rows = []

    for row in rows[1:]:
        row = [str(x).strip() for x in row]
        if len(row) < len(header):
            row = row + [None] * (len(header) - len(row))
        elif len(row) > len(header):
            row = row[:len(header)]
        parsed_rows.append(row)

    df = pd.DataFrame(parsed_rows, columns=header)

    begin_col = next((c for c in df.columns if str(c).lower().startswith("begin")), None)
    end_col = next((c for c in df.columns if str(c).lower().startswith("end")), None)

    if begin_col is None:
        raise ValueError("Could not find the 'begin' timestamp column.")

    df[begin_col] = pd.to_datetime(df[begin_col], dayfirst=True, errors="coerce")
    if end_col:
        df[end_col] = pd.to_datetime(df[end_col], dayfirst=True, errors="coerce")

    for col in df.columns:
        if col not in [begin_col, end_col]:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.dropna(subset=[begin_col]).sort_values(begin_col).reset_index(drop=True)
    return df, begin_col, end_col

def wide_to_long(df, begin_col, end_col):
    id_vars = [begin_col] + ([end_col] if end_col else [])
    value_vars = [c for c in df.columns if c not in id_vars]
    long_df = df.melt(id_vars=id_vars, value_vars=value_vars, var_name="parameter", value_name="value")
    meta = long_df["parameter"].apply(lambda x: pd.Series(split_variable(x), index=["site", "metric", "unit"]))
    long_df = pd.concat([long_df, meta], axis=1)

    rename_map = {begin_col: "begin"}
    if end_col:
        rename_map[end_col] = "end"

    long_df = long_df.rename(columns=rename_map)
    long_df = long_df.dropna(subset=["begin", "value"]).sort_values(["site", "metric", "begin"]).reset_index(drop=True)
    return long_df

@st.cache_data(show_spinner=False)
def parse_uploaded_content(file_bytes: bytes, file_name: str):
    ext = Path(file_name).suffix.lower()

    if ext in [".xlsx", ".xls"]:
        xdf = pd.read_excel(BytesIO(file_bytes), header=None, dtype=str)
        lines = []
        for _, row in xdf.iterrows():
            vals = [str(v).strip() for v in row.tolist() if pd.notna(v) and str(v).strip() != ""]
            if not vals:
                continue
            if len(vals) == 1:
                cell_text = vals[0].replace("\r\n", "\n").replace("\r", "\n")
                lines.extend([ln.strip() for ln in cell_text.split("\n") if ln.strip()])
            else:
                lines.append(";".join(vals))
    else:
        text = file_bytes.decode("utf-8-sig", errors="ignore")
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    wide_df, begin_col, end_col = parse_semicolon_table(lines)
    long_df = wide_to_long(wide_df, begin_col, end_col)
    return wide_df, long_df, begin_col, end_col

@st.cache_data(show_spinner=False)
def prepare_long_df(long_df):
    df = long_df.copy()
    df["site"] = df["site"].astype(str).str.strip()
    df["metric"] = df["metric"].astype(str).str.strip()
    df["unit"] = df["unit"].astype(str).replace("nan", "")
    return df.sort_values(["site", "metric", "begin"]).reset_index(drop=True)

def save_upload_bundle(uploaded_file, file_bytes, wide_df, long_df):
    ts = datetime.now()
    stamp = ts.strftime("%Y%m%d_%H%M%S")
    day_folder = UPLOAD_DIR / ts.strftime("%Y-%m-%d")
    day_folder.mkdir(parents=True, exist_ok=True)

    cleaned_name = safe_filename(uploaded_file.name)
    base_name = Path(cleaned_name).stem

    raw_path = day_folder / f"{stamp}_{cleaned_name}"
    wide_path = PROCESSED_DIR / f"{stamp}_{base_name}_wide.csv"
    long_path = PROCESSED_DIR / f"{stamp}_{base_name}_long.csv"

    with open(raw_path, "wb") as f:
        f.write(file_bytes)

    wide_df.to_csv(wide_path, index=False)
    long_df.to_csv(long_path, index=False)

    append_upload_log({
        "upload_timestamp": ts.strftime("%Y-%m-%d %H:%M:%S"),
        "original_file_name": uploaded_file.name,
        "saved_raw_file": str(raw_path),
        "saved_wide_csv": str(wide_path),
        "saved_long_csv": str(long_path),
        "rows_wide": len(wide_df),
        "rows_long": len(long_df),
        "sites": long_df["site"].nunique(),
        "metrics": long_df["metric"].nunique(),
        "start_time": str(long_df["begin"].min()),
        "end_time": str(long_df["begin"].max()),
        "file_hash": get_file_hash(file_bytes)
    })

def normalize_series(s):
    s = pd.to_numeric(s, errors="coerce")
    s_min = s.min()
    s_max = s.max()
    if pd.isna(s_min) or pd.isna(s_max):
        return pd.Series(np.nan, index=s.index)
    if s_max == s_min:
        return pd.Series(50.0, index=s.index)
    return ((s - s_min) / (s_max - s_min)) * 100.0

def get_filtered_site_metric_df(site_df, selected_metrics):
    return site_df[site_df["metric"].isin(selected_metrics)].copy().dropna(subset=["begin"]).sort_values("begin")

def get_last_n_days_df(site_df, selected_metrics, n_days):
    df = get_filtered_site_metric_df(site_df, selected_metrics)
    if df.empty:
        return df
    max_ts = df["begin"].max()
    start_ts = max_ts - pd.Timedelta(days=n_days)
    return df[df["begin"] >= start_ts].copy()

def get_custom_range_df(site_df, selected_metrics, start_dt, end_dt):
    df = get_filtered_site_metric_df(site_df, selected_metrics)
    if df.empty or start_dt is None or end_dt is None:
        return df.copy()
    return df[(df["begin"] >= start_dt) & (df["begin"] <= end_dt)].copy()

def build_latest_table(df):
    if df.empty:
        return pd.DataFrame(columns=["metric", "unit", "latest_time", "latest_value"])

    latest_df = (
        df.sort_values("begin")
        .groupby(["metric", "unit"], as_index=False)
        .tail(1)[["metric", "unit", "begin", "value"]]
        .rename(columns={"begin": "latest_time", "value": "latest_value"})
        .sort_values("metric")
        .reset_index(drop=True)
    )
    latest_df["latest_value"] = latest_df["latest_value"].round(2)
    return latest_df

def build_window_summary(df):
    if df.empty:
        return pd.DataFrame(columns=["metric", "unit", "avg", "min", "max", "latest_time"])

    summary_df = (
        df.groupby(["metric", "unit"], as_index=False)
        .agg(avg=("value", "mean"), min=("value", "min"), max=("value", "max"), latest_time=("begin", "max"))
        .sort_values("metric")
        .reset_index(drop=True)
    )

    for col in ["avg", "min", "max"]:
        summary_df[col] = summary_df[col].round(2)
    return summary_df

def build_combined_chart(df, site_name, scale_mode, chart_id, title_suffix):
    if df.empty:
        return None

    plot_df = df.copy().sort_values("begin")

    if scale_mode == "Normalized (0-100)":
        plot_df["plot_value"] = plot_df.groupby("metric")["value"].transform(normalize_series)
        y_title = "Normalized scale (0-100)"
        y_range = [0, 100]
    else:
        plot_df["plot_value"] = plot_df["value"]
        y_title = "Actual values"
        y_range = None

    fig = go.Figure()

    for metric in sorted(plot_df["metric"].dropna().unique()):
        mdf = plot_df[plot_df["metric"] == metric].copy()
        if mdf.empty:
            continue

        unit_mode = mdf["unit"].dropna().astype(str).mode()
        unit = unit_mode.iloc[0] if len(unit_mode) > 0 else ""
        label = f"{metric} [{unit}]" if str(unit).strip() else metric
        color = COLOR_MAP.get(metric, "#64748B")

        fig.add_trace(go.Scatter(
            x=mdf["begin"],
            y=mdf["plot_value"],
            mode="lines",
            name=label,
            line=dict(width=2.4, color=color, shape="spline", smoothing=0.35),
            connectgaps=True,
            hovertemplate=f"<b>{label}</b><br>Time: %{{x}}<br>Value: %{{customdata:.2f}}<extra></extra>",
            customdata=mdf["value"]
        ))

    fig.update_layout(
        title=f"{site_name} | {title_suffix}",
        template="plotly_white",
        height=455,
        margin=dict(l=10, r=10, t=35, b=60),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=-0.18,
            xanchor="center",
            x=0.5,
            bgcolor="rgba(255,255,255,0.0)",
            font=dict(size=11)
        ),
        legend_title="",
        hovermode="x unified",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#FFFFFF",
        uirevision=chart_id,
        font=dict(color="#172B4D")
    )

    fig.update_xaxes(title="Timestamp", showgrid=True, gridcolor="#E7EEF5", zeroline=False, showline=True, linecolor="#D9E2EC")
    fig.update_yaxes(title=y_title, showgrid=True, gridcolor="#E7EEF5", zeroline=False, showline=True, linecolor="#D9E2EC")

    if y_range is not None:
        fig.update_yaxes(range=y_range)

    return fig

def make_metric_summary(df_metric):
    if df_metric.empty:
        return pd.DataFrame()

    summary = (
        df_metric.groupby("site", as_index=False)
        .agg(
            count=("value", "count"),
            avg=("value", "mean"),
            min=("value", "min"),
            max=("value", "max"),
            std=("value", "std"),
            total=("value", "sum")
        )
    )

    latest = (
        df_metric.sort_values("begin")
        .groupby("site", as_index=False)
        .tail(1)[["site", "begin", "value"]]
        .rename(columns={"begin": "latest_time", "value": "latest_value"})
    )

    summary = summary.merge(latest, on="site", how="left")
    for col in ["avg", "min", "max", "std", "total", "latest_value"]:
        summary[col] = summary[col].round(2)

    return summary.sort_values("avg", ascending=False).reset_index(drop=True)

def build_insights(df_metric, metric_name, unit):
    insights = []
    if df_metric.empty:
        return insights

    avg_by_site = df_metric.groupby("site")["value"].mean().sort_values(ascending=False)
    max_by_site = df_metric.groupby("site")["value"].max().sort_values(ascending=False)
    std_by_site = df_metric.groupby("site")["value"].std().sort_values(ascending=False)

    if not avg_by_site.empty:
        insights.append(f"Highest average {metric_name.lower()}: {avg_by_site.index[0]} ({avg_by_site.iloc[0]:.2f} {unit})")
    if not max_by_site.empty:
        insights.append(f"Highest observed {metric_name.lower()}: {max_by_site.index[0]} ({max_by_site.iloc[0]:.2f} {unit})")

    std_non_na = std_by_site.dropna()
    if not std_non_na.empty:
        insights.append(f"Most variable {metric_name.lower()}: {std_non_na.index[0]} (std dev {std_non_na.iloc[0]:.2f} {unit})")

    if metric_name.lower() == "precipitation":
        total_by_site = df_metric.groupby("site")["value"].sum().sort_values(ascending=False)
        if not total_by_site.empty:
            insights.append(f"Highest cumulative precipitation: {total_by_site.index[0]} ({total_by_site.iloc[0]:.2f} {unit})")

    return insights

def add_chart_page(pdf, page_title, image_path, source_file_name, selected_metrics, scale_mode):
    pdf.add_page()
    pdf.set_font("Helvetica", style="B", size=13)
    pdf.cell(0, 8, page_title, ln=True)
    pdf.set_font("Helvetica", size=9)
    pdf.cell(0, 5, f"Source file: {source_file_name}", ln=True)
    pdf.cell(0, 5, f"Metrics: {', '.join(selected_metrics)}", ln=True)
    pdf.cell(0, 5, f"Scale mode: {scale_mode}", ln=True)
    pdf.cell(0, 5, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True)
    pdf.ln(3)
    pdf.image(str(image_path), x=10, y=pdf.get_y(), w=190)

def add_no_data_page(pdf, page_title, source_file_name, selected_metrics, scale_mode):
    pdf.add_page()
    pdf.set_font("Helvetica", style="B", size=13)
    pdf.cell(0, 8, page_title, ln=True)
    pdf.set_font("Helvetica", size=9)
    pdf.cell(0, 5, f"Source file: {source_file_name}", ln=True)
    pdf.cell(0, 5, f"Metrics: {', '.join(selected_metrics)}", ln=True)
    pdf.cell(0, 5, f"Scale mode: {scale_mode}", ln=True)
    pdf.ln(8)
    pdf.multi_cell(0, 6, "No data available for this chart window.")

def build_site_pdf_report_bytes(site_name, site_df, selected_metrics, scale_mode, source_file_name, report_timestamp):
    safe_site = safe_filename(site_name)
    temp_dir = REPORT_DIR / "_temp_images"
    temp_dir.mkdir(parents=True, exist_ok=True)

    chart_15_df = get_last_n_days_df(site_df, selected_metrics, 15)
    chart_2_df = get_last_n_days_df(site_df, selected_metrics, 2)

    fig_15 = build_combined_chart(
        df=chart_15_df,
        site_name=site_name,
        scale_mode=scale_mode,
        chart_id=f"{safe_site}_pdf_15_{report_timestamp}",
        title_suffix="Last 15 Days"
    )
    fig_2 = build_combined_chart(
        df=chart_2_df,
        site_name=site_name,
        scale_mode=scale_mode,
        chart_id=f"{safe_site}_pdf_2_{report_timestamp}",
        title_suffix="Last 2 Days"
    )

    img_15_path = temp_dir / f"{safe_site}_{report_timestamp}_15day.png"
    img_2_path = temp_dir / f"{safe_site}_{report_timestamp}_2day.png"
    created_images = []

    try:
        if fig_15 is not None:
            fig_15.write_image(str(img_15_path), format="png", width=1600, height=900)
            created_images.append(img_15_path)

        if fig_2 is not None:
            fig_2.write_image(str(img_2_path), format="png", width=1600, height=900)
            created_images.append(img_2_path)

        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=10)

        if img_15_path.exists():
            add_chart_page(
                pdf=pdf,
                page_title=f"{site_name} | Last 15 Days",
                image_path=img_15_path,
                source_file_name=source_file_name,
                selected_metrics=selected_metrics,
                scale_mode=scale_mode
            )
        else:
            add_no_data_page(
                pdf=pdf,
                page_title=f"{site_name} | Last 15 Days",
                source_file_name=source_file_name,
                selected_metrics=selected_metrics,
                scale_mode=scale_mode
            )

        if img_2_path.exists():
            add_chart_page(
                pdf=pdf,
                page_title=f"{site_name} | Last 2 Days",
                image_path=img_2_path,
                source_file_name=source_file_name,
                selected_metrics=selected_metrics,
                scale_mode=scale_mode
            )
        else:
            add_no_data_page(
                pdf=pdf,
                page_title=f"{site_name} | Last 2 Days",
                source_file_name=source_file_name,
                selected_metrics=selected_metrics,
                scale_mode=scale_mode
            )

        result = pdf.output(dest="S")
        return result if isinstance(result, (bytes, bytearray)) else result.encode("latin1")

    finally:
        for image_path in created_images:
            if image_path.exists():
                try:
                    image_path.unlink()
                except Exception:
                    pass

# =========================================================
# STATE
# =========================================================
def init_dashboard_state(current_hash, sites, metrics):
    default_metrics = [m for m in ["Wind Speed", "Wind Direction"] if m in metrics]
    if not default_metrics and metrics:
        default_metrics = [metrics[0]]

    if st.session_state.get("loaded_file_hash") != current_hash:
        st.session_state["loaded_file_hash"] = current_hash
        st.session_state["applied_filters"] = {
            "scale_mode": "Normalized (0-100)",
            "selected_metrics": default_metrics,
            "displayed_sites": sites,
            "report_sites": sites
        }
        st.session_state["site_view_state"] = {}
        st.session_state["generated_site_pdfs"] = []

    st.session_state.setdefault("applied_filters", {
        "scale_mode": "Normalized (0-100)",
        "selected_metrics": default_metrics,
        "displayed_sites": sites,
        "report_sites": sites
    })
    st.session_state.setdefault("site_view_state", {})
    st.session_state.setdefault("generated_site_pdfs", [])

def ensure_site_state(site, site_min, site_max):
    if site not in st.session_state["site_view_state"]:
        st.session_state["site_view_state"][site] = {
            "mode": "15 Day",
            "custom_start_date": site_min.date() if pd.notna(site_min) else datetime.now().date(),
            "custom_end_date": site_max.date() if pd.notna(site_max) else datetime.now().date(),
            "custom_start_time": time(0, 0),
            "custom_end_time": time(23, 59)
        }

# =========================================================
# HEADER
# =========================================================
st.markdown("""
<div class="main-shell">
    <div class="hero-panel">
        <div class="hero-title">Meteorological Operations Suite</div>
        <div class="hero-subtitle">
            Compact operational monitoring dashboard for multi-site meteorological data.
            Focused on clean charting, fast filtering, and minimal on-screen clutter.
        </div>
        <span class="pill">Compact UI</span>
        <span class="pill">Instant 15/2 Day</span>
        <span class="pill">Custom Range View</span>
        <span class="pill">Per-Site PDF Reports</span>
    </div>
</div>
""", unsafe_allow_html=True)

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.markdown("### Data Input")
uploaded_file = st.sidebar.file_uploader("Upload CSV / Excel", type=["csv", "xlsx", "xls"])
show_summary_tables = st.sidebar.toggle("Show metric summaries", value=False)
show_raw = st.sidebar.toggle("Show processed raw data", value=False)
show_upload_history = st.sidebar.toggle("Show upload history", value=False)

if uploaded_file is None:
    st.info("Upload your latest meteorological file to begin.")
    if show_upload_history:
        log_df = load_upload_log()
        if not log_df.empty:
            st.dataframe(
                log_df.sort_values("upload_timestamp", ascending=False).head(10),
                use_container_width=True,
                hide_index=True
            )
    st.stop()

# =========================================================
# PROCESS FILE
# =========================================================
try:
    file_bytes = uploaded_file.getvalue()
    with st.spinner("Processing uploaded file..."):
        wide_df, long_df, begin_col, end_col = parse_uploaded_content(file_bytes, uploaded_file.name)
        long_df = prepare_long_df(long_df)
except Exception as e:
    st.error(f"File could not be parsed. Error: {e}")
    st.stop()

current_hash = get_file_hash(file_bytes)
if st.session_state.get("last_saved_hash") != current_hash:
    save_upload_bundle(uploaded_file, file_bytes, wide_df, long_df)
    st.session_state["last_saved_hash"] = current_hash
    st.success(f"Uploaded and saved: {uploaded_file.name}")

sites = sorted(long_df["site"].dropna().unique().tolist())
metrics = sorted(long_df["metric"].dropna().unique().tolist())

init_dashboard_state(current_hash, sites, metrics)
applied = st.session_state["applied_filters"]

# =========================================================
# FILTERS
# =========================================================
st.sidebar.markdown("### Analysis Controls")
with st.sidebar.form("dashboard_filter_form"):
    form_scale_mode = st.selectbox(
        "Combined graph scale",
        ["Normalized (0-100)", "Actual Values"],
        index=["Normalized (0-100)", "Actual Values"].index(applied["scale_mode"]) if applied["scale_mode"] in ["Normalized (0-100)", "Actual Values"] else 0
    )

    form_selected_metrics = st.multiselect(
        "Parameters",
        options=metrics,
        default=[m for m in applied["selected_metrics"] if m in metrics]
    )

    form_displayed_sites = st.multiselect(
        "Sites",
        options=sites,
        default=[s for s in applied["displayed_sites"] if s in sites]
    )

    form_report_sites = st.multiselect(
        "Report sites",
        options=sites,
        default=[s for s in applied["report_sites"] if s in sites]
    )

    view_dashboard = st.form_submit_button("View Dashboard", use_container_width=True)

if view_dashboard:
    st.session_state["applied_filters"] = {
        "scale_mode": form_scale_mode,
        "selected_metrics": form_selected_metrics,
        "displayed_sites": form_displayed_sites,
        "report_sites": form_report_sites
    }
    applied = st.session_state["applied_filters"]

scale_mode = applied["scale_mode"]
selected_metrics = applied["selected_metrics"]
displayed_sites = applied["displayed_sites"]
selected_sites_for_reports = applied["report_sites"]

filtered_long = long_df[
    (long_df["site"].isin(selected_sites_for_reports)) &
    (long_df["metric"].isin(selected_metrics))
].copy()

summary_tables = {}

if show_summary_tables and selected_metrics:
    st.markdown('<div class="section-label">Metric Summaries</div>', unsafe_allow_html=True)
    for metric in selected_metrics:
        metric_df = filtered_long[filtered_long["metric"] == metric].copy()
        if metric_df.empty:
            continue

        unit_mode = metric_df["unit"].dropna().astype(str).mode()
        unit = unit_mode.iloc[0] if len(unit_mode) > 0 else ""
        summary_df = make_metric_summary(metric_df)
        summary_tables[metric] = summary_df

        insights = build_insights(metric_df, metric, unit)
        with st.expander(f"{metric} summary", expanded=False):
            for item in insights:
                st.write(f"- {item}")
            st.dataframe(summary_df, use_container_width=True, hide_index=True)

# =========================================================
# TOP NOTE
# =========================================================
st.markdown("""
<div class="subtle-note">
Apply filters from the sidebar, then use <b>View Dashboard</b>. Each site card supports 15 Day, 2 Day, and Custom views.
The PDF generator creates one PDF per selected report site and includes exactly two charts per PDF: 15 Day and 2 Day.
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-label">Site Analysis</div>', unsafe_allow_html=True)

# =========================================================
# SITE CARDS
# =========================================================
if not selected_metrics:
    st.warning("Please select at least one parameter and click View Dashboard.")
else:
    for site in displayed_sites:
        site_df = long_df[long_df["site"] == site].copy()
        if site_df.empty:
            continue

        site_min = site_df["begin"].min()
        site_max = site_df["begin"].max()
        ensure_site_state(site, site_min, site_max)
        site_state = st.session_state["site_view_state"][site]
        safe_site = safe_filename(site)

        metrics_present = sorted(site_df[site_df["metric"].isin(selected_metrics)]["metric"].dropna().unique().tolist())
        chips_html = "".join([f'<span class="metric-chip">{m}</span>' for m in metrics_present])

        st.markdown('<div class="site-shell">', unsafe_allow_html=True)
        st.markdown(f"""
            <div class="site-title">{site}</div>
            <div class="site-subtitle">Available range: {str(site_min)[:16] if pd.notna(site_min) else "-"} to {str(site_max)[:16] if pd.notna(site_max) else "-"}</div>
            <div style="margin-top:4px;">{chips_html}</div>
        """, unsafe_allow_html=True)

        current_mode = st.radio(
            "View range",
            options=["15 Day", "2 Day", "Custom"],
            index=["15 Day", "2 Day", "Custom"].index(site_state["mode"]) if site_state["mode"] in ["15 Day", "2 Day", "Custom"] else 0,
            horizontal=True,
            key=f"{safe_site}_mode_radio",
            label_visibility="collapsed"
        )
        site_state["mode"] = current_mode

        st.markdown(
            f'<div style="margin:6px 0 10px 0;"><span class="range-badge">Current View: {site_state["mode"]}</span></div>',
            unsafe_allow_html=True
        )

        if current_mode == "Custom":
            st.markdown('<div class="helper-note">Choose dates and time, then click <b>View Custom Range</b>.</div>', unsafe_allow_html=True)
            with st.form(f"custom_form_{safe_site}"):
                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    custom_start_date = st.date_input(
                        "Start date",
                        value=site_state["custom_start_date"],
                        min_value=site_min.date() if pd.notna(site_min) else None,
                        max_value=site_max.date() if pd.notna(site_max) else None,
                        key=f"{safe_site}_custom_start_date_widget"
                    )
                with c2:
                    custom_end_date = st.date_input(
                        "End date",
                        value=site_state["custom_end_date"],
                        min_value=site_min.date() if pd.notna(site_min) else None,
                        max_value=site_max.date() if pd.notna(site_max) else None,
                        key=f"{safe_site}_custom_end_date_widget"
                    )
                with c3:
                    custom_start_time = st.time_input(
                        "Start time",
                        value=site_state["custom_start_time"],
                        key=f"{safe_site}_custom_start_time_widget"
                    )
                with c4:
                    custom_end_time = st.time_input(
                        "End time",
                        value=site_state["custom_end_time"],
                        key=f"{safe_site}_custom_end_time_widget"
                    )
                view_custom = st.form_submit_button("View Custom Range", use_container_width=True)

            if view_custom:
                site_state["custom_start_date"] = custom_start_date
                site_state["custom_end_date"] = custom_end_date
                site_state["custom_start_time"] = custom_start_time
                site_state["custom_end_time"] = custom_end_time

            custom_start_dt = pd.Timestamp(datetime.combine(site_state["custom_start_date"], site_state["custom_start_time"]))
            custom_end_dt = pd.Timestamp(datetime.combine(site_state["custom_end_date"], site_state["custom_end_time"]))

            if custom_end_dt < custom_start_dt:
                st.warning("Custom range end must be after start.")
                range_df = pd.DataFrame()
                title_suffix = "Custom Range"
            else:
                range_df = get_custom_range_df(site_df, selected_metrics, custom_start_dt, custom_end_dt)
                title_suffix = f"Custom Range | {custom_start_dt} to {custom_end_dt}"
        elif current_mode == "2 Day":
            range_df = get_last_n_days_df(site_df, selected_metrics, 2)
            title_suffix = "Last 2 Days"
        else:
            range_df = get_last_n_days_df(site_df, selected_metrics, 15)
            title_suffix = "Last 15 Days"

        fig = build_combined_chart(
            df=range_df,
            site_name=site,
            scale_mode=scale_mode,
            chart_id=f"{safe_site}_{current_mode}_{scale_mode}_{len(range_df)}",
            title_suffix=title_suffix
        )

        if fig is not None:
            st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False, "responsive": True})
        else:
            st.info("No data available for the current site view.")

        latest_df = build_latest_table(range_df)
        summary_df = build_window_summary(range_df)

        tab1, tab2 = st.tabs(["Latest values", "Window summary"])
        with tab1:
            if latest_df.empty:
                st.caption("No latest values available for this view.")
            else:
                st.dataframe(latest_df, use_container_width=True, hide_index=True)

        with tab2:
            if summary_df.empty:
                st.caption("No summary available for this view.")
            else:
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

        st.markdown("</div>", unsafe_allow_html=True)

# =========================================================
# GENERATE SITE PDFS
# =========================================================
st.markdown('<div class="section-label">Generate Site PDFs</div>', unsafe_allow_html=True)

if st.button("Generate Site PDF Reports"):
    if not selected_metrics:
        st.warning("Please select at least one parameter before generating PDFs.")
    elif not selected_sites_for_reports:
        st.warning("Please select at least one report site before generating PDFs.")
    else:
        generated_reports = []
        report_timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

        with st.spinner("Generating site PDF reports..."):
            for site in selected_sites_for_reports:
                site_df = long_df[
                    (long_df["site"] == site) &
                    (long_df["metric"].isin(selected_metrics))
                ].copy()

                if site_df.empty:
                    continue

                try:
                    pdf_bytes = build_site_pdf_report_bytes(
                        site_name=site,
                        site_df=site_df,
                        selected_metrics=selected_metrics,
                        scale_mode=scale_mode,
                        source_file_name=uploaded_file.name,
                        report_timestamp=report_timestamp
                    )

                    pdf_name = f"{safe_filename(site)}_{report_timestamp}.pdf"
                    pdf_path = REPORT_DIR / pdf_name

                    with open(pdf_path, "wb") as f:
                        f.write(pdf_bytes)

                    generated_reports.append({
                        "site": site,
                        "file_name": pdf_name,
                        "path": str(pdf_path),
                        "bytes": pdf_bytes
                    })

                except Exception as e:
                    st.error(f"Could not generate PDF for {site}: {e}")

        st.session_state["generated_site_pdfs"] = generated_reports

        if generated_reports:
            st.success(f"Generated {len(generated_reports)} site PDF report(s) in: {REPORT_DIR}")
        else:
            st.warning("No site PDFs were generated.")

# =========================================================
# PDF DOWNLOADS
# =========================================================
generated_site_pdfs = st.session_state.get("generated_site_pdfs", [])

if generated_site_pdfs:
    st.markdown('<div class="section-label">PDF Downloads</div>', unsafe_allow_html=True)

    for report in generated_site_pdfs:
        c1, c2 = st.columns([2.4, 1])
        with c1:
            st.write(f"{report['site']}")
            st.caption(report["path"])
        with c2:
            pdf_path = report["path"]

        with open(pdf_path, "rb") as pdf_file:
            pdf_data = pdf_file.read()
        
        st.download_button(
            label=f"Download {report['site']} PDF",
            data=pdf_data,
            file_name=report["file_name"],
            mime="application/pdf",
            key=f"download_{safe_filename(report['site'])}_{report['file_name']}",
            use_container_width=True
        )

st.markdown("### Email Test")

test_email_to = st.text_input("Send test email to", value="")

if st.button("Send Test Email"):
    if not test_email_to.strip():
        st.warning("Please enter a recipient email address.")
    else:
        try:
            send_test_email_outlook(test_email_to.strip())
            st.success(f"Test email sent to {test_email_to.strip()}")
        except Exception as e:
            st.error(f"Email failed: {e}")


# =========================================================
# OPTIONAL DATA VIEWS
# =========================================================
if show_raw:
    st.markdown('<div class="section-label">Processed Raw Data</div>', unsafe_allow_html=True)
    st.dataframe(filtered_long, use_container_width=True, hide_index=True)

if show_upload_history:
    log_df = load_upload_log()
    if not log_df.empty:
        st.markdown('<div class="section-label">Upload History</div>', unsafe_allow_html=True)
        st.dataframe(
            log_df.sort_values("upload_timestamp", ascending=False).head(20),
            use_container_width=True,
            hide_index=True
        )
