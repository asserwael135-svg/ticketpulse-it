"""
TicketPulse IT - Ticket Telemetry & Integrity Audit
Requires: streamlit>=1.40, pandas, numpy, openpyxl
Run:      streamlit run app.py
"""
import hashlib
import html as _html
import inspect
import time
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st

# =========================================================
# 0. Settings you may want to edit
# =========================================================
DEFAULT_FILE = "Closed_Tickets_20261017_160931.xlsx"   # used when nothing is uploaded
TICKET_PREFIX = "#IT-"
SLA_TARGET_MIN = 25.0        # SLA target shown on the speed card (minutes)
FAST_CLOSE_SEC = 30          # tickets closed faster than this are "Fast-Closed"
ANOMALY_THRESHOLD = 10.0     # % suspicious rate that triggers a review
FCR_MAX_MIN = 30             # proxy for first-contact resolution (genuine tickets closed within X min)
PAGE_SIZE = 6
# 'use_container_width' was replaced by width='stretch' in newer Streamlit; support both
STRETCH = {"width": "stretch"} if "width" in inspect.signature(st.button).parameters else {"use_container_width": True}

NAV = ["Overview & Fraud Matrix", "Live Ticket Stream", "Agent Audit", "Rules Engine", "Reports"]
TAGS = ["Fast-Closed", "Self-Assigned", "Generic Solution", "Missing Solution"]
TAB_LABELS = {"all": "All Flagged", "Fast-Closed": "Closed < 30s",
              "Self-Assigned": "Self-Assigned", "Generic Solution": "Generic Solution", "Missing Solution": "Missing Solution"}

st.set_page_config(
    page_title="TicketPulse IT v2.4 - Telemetry & Integrity Audit",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="collapsed",
)

for _k, _v in {"theme": "dark", "page": 1, "dismissed": set()}.items():
    st.session_state.setdefault(_k, _v)

# =========================================================
# 1. Theme + CSS
# =========================================================
THEMES = {
    "dark": dict(bg="#0A0D14", glow="rgba(99,102,241,.10)", panel="#0E121A", card="#11161F", head="#0D1118",
                 chip="rgba(148,163,184,.12)", border="#212836", border2="#2C3445", text="#E6E9F0",
                 muted="#8A93A6", accent="#6366F1", accent_text="#A5B4FC", soft="rgba(99,102,241,.16)",
                 green="#34D399", green_bg="rgba(52,211,153,.13)", red="#F87171", red_bg="rgba(248,113,113,.14)",
                 track="#1B2230"),
    "light": dict(bg="#F4F5F9", glow="rgba(99,102,241,.05)", panel="#FFFFFF", card="#FFFFFF", head="#F6F7FB",
                  chip="#F1F3F7", border="#E4E7EE", border2="#D5D9E3", text="#111827", muted="#6B7280",
                  accent="#5B5BF0", accent_text="#4F46E5", soft="rgba(91,91,240,.10)",
                  green="#059669", green_bg="rgba(16,185,129,.12)", red="#DC2626", red_bg="rgba(239,68,68,.10)",
                  track="#E8EBF2"),
}

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=JetBrains+Mono:wght@400;500;600&display=swap');
html, body, .stApp {background: radial-gradient(900px 320px at 55% 0%, var(--glow), transparent 70%), var(--bg) !important;
  color: var(--text); font-family: 'Inter', sans-serif;}
[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stSidebar"], [data-testid="collapsedControl"],
[data-testid="stSidebarCollapsedControl"], footer, #MainMenu {display: none !important;}
.block-container, [data-testid="stMainBlockContainer"] {padding: 0 !important; max-width: 100% !important;}
[data-testid="stColumn"] {min-width: 0;}
[data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p {color: var(--text); font-family: 'Inter', sans-serif;}
[data-testid="stMarkdownContainer"] p {margin: 0;}
.mono {font-family: 'JetBrains Mono', monospace;}
svg.ic {vertical-align: -3px; flex: none;}

/* ---------- top bar / layout ---------- */
.st-key-topbar {background: var(--panel); border-bottom: 1px solid var(--border); padding: 10px 24px;}
.st-key-body [data-testid="stHorizontalBlock"] {gap: 0;}
.st-key-sidenav {background: var(--panel); border-right: 1px solid var(--border); padding: 24px 16px;
  min-height: calc(100vh - 66px); display: flex; flex-direction: column; gap: 6px;}
.st-key-sidenav > :last-child {margin-top: auto;}
.st-key-main {padding: 24px; gap: 20px;}
.st-key-main [data-testid="stHorizontalBlock"] {gap: 12px;}
.st-key-hero {background: var(--card); border: 1px solid var(--border); border-radius: 16px; padding: 22px 24px;}
.brand {display: flex; align-items: center; gap: 10px; white-space: nowrap;}
.logo {width: 32px; height: 32px; border-radius: 9px; background: var(--accent); color: #fff; display: inline-flex;
  align-items: center; justify-content: center;}
.brand b {font-size: 17px; font-weight: 600;}
.ver {font: 500 11px 'JetBrains Mono', monospace; color: var(--muted); background: var(--chip); border: 1px solid var(--border);
  padding: 2px 8px; border-radius: 6px;}
.pill-g {display: inline-flex; align-items: center; gap: 8px; font: 500 11px 'JetBrains Mono', monospace; color: var(--green);
  background: var(--green-bg); padding: 6px 12px; border-radius: 20px; white-space: nowrap;}
.pill-g i {width: 7px; height: 7px; border-radius: 50%; background: var(--green);}
.user {display: flex; align-items: center; gap: 10px; justify-content: flex-end;}
.user .av {width: 34px; height: 34px; background: linear-gradient(135deg, #6366F1, #A855F7); color: #fff; font-size: 12px;}
.nav-l {font: 600 11px 'Inter', sans-serif; letter-spacing: .06em; color: var(--muted); padding: 0 12px 10px;}
.ingest {background: var(--card); border: 1px solid var(--border); border-radius: 10px; padding: 12px 14px;}
.ingest .row {display: flex; justify-content: space-between; font-size: 12px; margin-bottom: 8px;}
.ingest .row.s {font: 400 10.5px 'JetBrains Mono', monospace; color: var(--muted); margin: 8px 0 0;}

/* ---------- generic pieces ---------- */
.card {background: var(--card); border: 1px solid var(--border); border-radius: 14px; padding: 20px 22px;}
.h1 {font-size: 25px; font-weight: 500; letter-spacing: -.01em; margin: 0 0 10px;}
.row-w {display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin-top: 8px;}
.tag-r {font: 500 11px 'JetBrains Mono', monospace; color: var(--red); background: var(--red-bg); padding: 3px 10px; border-radius: 6px;}
.chipx {display: inline-flex; align-items: center; gap: 7px; font: 500 12px 'Inter', sans-serif; color: var(--text);
  background: var(--chip); border: 1px solid var(--border); padding: 5px 10px; border-radius: 8px; max-width: 100%;}
.chipx.m {font: 400 11.5px 'JetBrains Mono', monospace;}
.chipx .e {overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 170px;}
.hash {font: 400 11.5px 'JetBrains Mono', monospace; color: var(--green); margin-top: 10px; display: flex; gap: 6px; align-items: center;}
.chip {display: inline-flex; align-items: center; gap: 5px; padding: 3px 9px; border-radius: 8px; white-space: nowrap;
  font: 500 11px 'JetBrains Mono', monospace; background: var(--chip); color: var(--muted);}
.chip.g {color: var(--green); background: var(--green-bg);}
.chip.r {color: var(--red); background: var(--red-bg);}
.chip.b {color: var(--accent-text); background: var(--soft);}
.g {color: var(--green);} .r {color: var(--red);}
.bar {display: flex; height: 6px; border-radius: 6px; overflow: hidden; background: var(--track);}
.bar i {display: block; height: 100%;}

/* ---------- KPI cards ---------- */
.kpi-grid {display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px;}
@media (max-width: 1100px) {.kpi-grid {grid-template-columns: repeat(2, 1fr);}}
.kpi {background: var(--card); border: 1px solid var(--border); border-radius: 14px; padding: 18px 20px; display: flex; flex-direction: column;}
.kpi-h {display: flex; justify-content: space-between; align-items: flex-start; gap: 8px; color: var(--muted);
  font: 600 11px 'Inter', sans-serif; letter-spacing: .06em; text-transform: uppercase; line-height: 1.4; min-height: 30px;}
.kpi-v {display: flex; align-items: center; gap: 10px; margin-top: 10px; flex-wrap: wrap; justify-content: space-between;}
.num {font-size: 30px; font-weight: 600; letter-spacing: -.02em; font-variant-numeric: tabular-nums; line-height: 1.1;}
.kpi-f {display: flex; justify-content: space-between; gap: 8px; font: 400 11px 'JetBrains Mono', monospace; color: var(--muted); margin-top: 8px;}
.tri {display: grid; grid-template-columns: repeat(4, auto); justify-content: space-between; margin: 12px 0 6px;
  font: 400 10.5px 'JetBrains Mono', monospace; color: var(--muted);}
.tri b {display: block; font-weight: 600; color: var(--text); margin-top: 2px;}
.spark {width: 100%; height: 44px; display: block; margin: 10px 0 6px;}

/* ---------- leaderboard + categories ---------- */
.mid {display: grid; grid-template-columns: 1.45fr 1fr; gap: 16px; align-items: stretch;}
@media (max-width: 1100px) {.mid {grid-template-columns: 1fr;}}
.mid .card {display: flex; flex-direction: column;}
.sec-h {display: flex; justify-content: space-between; align-items: flex-start; gap: 12px;}
.sec-t {font-size: 18px; font-weight: 500; display: flex; align-items: center; gap: 9px;}
.sec-s {color: var(--muted); font-size: 13px; margin-top: 6px; max-width: 460px; line-height: 1.45;}
.lb {margin-top: 18px; border: 1px solid var(--border); border-radius: 10px; overflow: hidden;}
.lb-r {display: grid; grid-template-columns: 2.6fr .8fr 1.3fr 1fr 1.2fr; align-items: center; gap: 6px;
  padding: 12px 14px; border-top: 1px solid var(--border);}
.lb-r.hd {background: var(--head); color: var(--muted); font: 600 10.5px 'Inter', sans-serif; letter-spacing: .06em; border-top: 0;}
.lb-r .c {text-align: center;}
.who {display: flex; align-items: center; gap: 10px; min-width: 0;}
.av {width: 28px; height: 28px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center;
  font: 600 10px 'JetBrains Mono', monospace; background: var(--chip); color: var(--muted); flex: none;}
.av.r {background: var(--red-bg); color: var(--red);} .av.g {background: var(--green-bg); color: var(--green);}
.av.b {background: var(--soft); color: var(--accent-text);}
.nm {font-size: 13px; font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;}
.sub {font: 400 10.5px 'JetBrains Mono', monospace; color: var(--muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap;}
.foot {display: flex; justify-content: space-between; gap: 12px; margin-top: 14px; font: 400 11px 'JetBrains Mono', monospace; color: var(--muted);}
.cat {margin-top: 16px;}
.cat-t {display: flex; justify-content: space-between; font-size: 12.5px; margin-bottom: 7px; gap: 8px;}
.cat-t span:last-child {font: 400 11.5px 'JetBrains Mono', monospace; color: var(--muted);}
.cat-t span:last-child.hot {color: var(--red); font-weight: 600;}
.donut {display: flex; align-items: center; gap: 22px; margin-top: auto; padding-top: 28px;}
.legend {display: grid; grid-template-columns: 1fr 1fr; gap: 8px 18px; font: 400 11.5px 'JetBrains Mono', monospace;}
.legend i {display: inline-block; width: 8px; height: 8px; border-radius: 50%; margin-right: 7px;}

/* ---------- inspection table ---------- */
.st-key-inspect {background: var(--card); border: 1px solid var(--border); border-radius: 14px; padding: 20px 22px; gap: 14px;}
.st-key-tablewrap {border: 1px solid var(--border); border-radius: 10px; overflow: hidden; gap: 0 !important;}
.st-key-thead, .st-key-tbl {gap: 0 !important;}
.st-key-thead [data-testid="stHorizontalBlock"] {background: var(--head); padding: 10px 14px; gap: 6px !important;}
.st-key-tbl [data-testid="stHorizontalBlock"] {padding: 8px 14px; border-top: 1px solid var(--border); gap: 6px !important;}
.th {color: var(--muted); font: 600 10.5px 'Inter', sans-serif; letter-spacing: .06em;}
.tid {font: 500 12px 'JetBrains Mono', monospace; color: var(--accent-text);}
.cell {font-size: 12.5px; display: flex; align-items: center; gap: 8px; min-width: 0;}
.note {font: 400 11.5px 'JetBrains Mono', monospace; color: var(--text); overflow: hidden; text-overflow: ellipsis; white-space: nowrap;}
.empty {padding: 28px; text-align: center; color: var(--muted); font-size: 13px;}

/* ---------- Streamlit widgets ---------- */
.stApp .stButton > button, .stApp .stDownloadButton > button, .stApp [data-testid="stPopover"] button {
  background: transparent; color: var(--text); border: 1px solid transparent; border-radius: 8px; box-shadow: none;
  font: 500 12.5px 'Inter', sans-serif; padding: 6px 12px; min-height: 0; white-space: nowrap;}
.stApp .stButton > button:hover, .stApp .stDownloadButton > button:hover, .stApp [data-testid="stPopover"] button:hover {
  background: var(--chip); border-color: var(--border); color: var(--text);}
.stApp .stButton > button[kind="primary"], .stApp .stDownloadButton > button[kind="primary"],
.stApp button[data-testid="stBaseButton-primary"] {background: var(--accent); color: #fff !important; border: 0;}
.stApp button[data-testid="stBaseButton-primary"]:hover {background: var(--accent); filter: brightness(1.1);}
.stApp button p {color: inherit !important; font: inherit !important;}
.stApp button:disabled {opacity: .4;}
.st-key-pager button {min-width: 32px; padding: 4px 8px !important; border-color: var(--border) !important;}
[data-testid="stFileUploader"] section, [data-testid="stFileUploaderDropzone"] {background: var(--card); border: 1px dashed var(--border2);
  border-radius: 10px; padding: 6px 10px; min-height: 0; align-items: center;}
[data-testid="stFileUploaderDropzoneInstructions"] {display: none;}
[data-testid="stFileUploaderDropzone"]::before {content: "Drop IT_Dump.xlsx or browse"; font: 400 12px 'JetBrains Mono', monospace;
  color: var(--muted); margin-right: auto;}
[data-testid="stFileUploaderDropzone"] button {font-size: 11px; padding: 3px 8px; border: 1px solid var(--border); color: var(--text);}
[data-testid="stFileUploaderFile"] * {color: var(--text) !important;}
[data-baseweb="input"], [data-baseweb="input"] > div {background: var(--card) !important; border: 1px solid var(--border) !important; border-radius: 8px !important;}
[data-baseweb="input"] input {color: var(--text) !important; font-size: 13px; background: transparent !important;}
[data-testid="InputInstructions"] {display: none;}
.stApp [data-testid="stSelectbox"] > div > div {background: var(--card); color: var(--text); border-color: var(--border);}
[data-testid="stRadio"] [role="radiogroup"] {gap: 4px;}
[data-testid="stRadio"] label {padding: 8px 12px; border-radius: 8px; cursor: pointer; margin: 0; width: 100%;}
[data-testid="stRadio"] label > div:first-child {display: none;}
[data-testid="stRadio"] label p {font-size: 13.5px; color: var(--muted) !important;}
[data-testid="stRadio"] label:hover {background: var(--chip);}
[data-testid="stRadio"] label:has(input:checked) {background: var(--soft);}
[data-testid="stRadio"] label:has(input:checked) p {color: var(--accent-text) !important; font-weight: 600;}
.st-key-tabs [role="radiogroup"] {flex-direction: row; flex-wrap: wrap;}
.st-key-tabs label {width: auto; padding: 6px 12px;}
.st-key-tabs label p {font-size: 12.5px;}
.st-key-tabs label:has(input:checked) {background: var(--accent);}
.st-key-tabs label:has(input:checked) p {color: #fff !important;}
.st-key-tabs code {font: 500 10.5px 'JetBrains Mono', monospace; background: var(--chip); color: inherit; padding: 1px 7px; border-radius: 10px;}
.st-key-tabs label:has(input:checked) code {background: rgba(255,255,255,.22); color: #fff;}
[data-testid="stCheckbox"] {margin: 0;}
[data-testid="stMetric"] {background: var(--card); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px;}
[data-testid="stMetricLabel"] *, [data-testid="stMetricValue"] {color: var(--text) !important;}

/* ---------- popovers / dropdowns render in a portal, so theme them explicitly ---------- */
[data-testid="stPopoverBody"], [data-baseweb="popover"] > div {background: var(--card) !important; border: 1px solid var(--border) !important;
  border-radius: 12px !important; color: var(--text) !important;}
[data-testid="stPopoverBody"] p, [data-testid="stPopoverBody"] strong, [data-testid="stPopoverBody"] span {color: var(--text) !important;}
[data-testid="stPopoverBody"] [data-testid="stAlert"] {background: var(--soft) !important; border: 1px solid var(--border);}
[data-testid="stPopoverBody"] [data-testid="stAlert"] * {color: var(--accent-text) !important;}
[data-baseweb="menu"], [data-baseweb="menu"] li, [data-baseweb="select"] > div {background: var(--card) !important; color: var(--text) !important;}
[data-baseweb="menu"] li:hover {background: var(--chip) !important;}
"""


def inject_css(theme: str):
    tokens = "".join(f"--{k.replace('_', '-')}:{v};" for k, v in THEMES[theme].items())
    st.markdown(f"<style>:root{{{tokens}}}{CSS}</style>", unsafe_allow_html=True)


# =========================================================
# 2. Small helpers
# =========================================================
ICONS = {
    "pulse": '<path d="M22 12h-4l-3 9L9 3l-3 9H2"/>',
    "calendar": '<rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/>',
    "file": '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><path d="M14 2v6h6M16 13H8M16 17H8"/>',
    "check": '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><path d="M22 4L12 14.01l-3-3"/>',
    "timer": '<circle cx="12" cy="13" r="8"/><path d="M12 9v4l2 2M9 2h6"/>',
    "inbox": '<path d="M22 12h-6l-2 3h-4l-2-3H2"/><path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>',
    "users": '<path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75"/>',
    "pie": '<path d="M21.21 15.89A10 10 0 1 1 8 2.83"/><path d="M22 12A10 10 0 0 0 12 2v10z"/>',
    "flag": '<path d="M4 15s1-1 4-1 5 2 8 2 4-1 4-1V3s-1 1-4 1-5-2-8-2-4 1-4 1z"/><path d="M4 22v-7"/>',
    "bolt": '<path d="M13 2L3 14h9l-1 8 10-12h-9z"/>',
    "up": '<path d="M23 6l-9.5 9.5-5-5L1 18"/><path d="M17 6h6v6"/>',
    "down": '<path d="M12 5v14M19 12l-7 7-7-7"/>',
}


def icon(name, size=16):
    return (f'<svg class="ic" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            f'stroke-width="2" stroke-linecap="round" stroke-linejoin="round">{ICONS[name]}</svg>')


def esc(v):
    """Escape data for safe use inside HTML rendered through st.markdown."""
    return " ".join(_html.escape(str(v), quote=True).split()).replace("$", "&#36;")


def md(s: str):
    """Render an HTML snippet (must start with a <div>) through st.markdown."""
    st.markdown("\n".join(l.strip() for l in s.strip().splitlines() if l.strip()), unsafe_allow_html=True)


def fmt_dur(minutes, force_min=False):
    if minutes is None or pd.isna(minutes):
        return "N/A"
    s = max(int(round(float(minutes) * 60)), 0)
    if s >= 86400:
        return f"{s // 86400}d {s % 86400 // 3600}h"
    if s >= 3600:
        return f"{s // 3600}h {s % 3600 // 60:02d}m"
    if s >= 60 or force_min:
        return f"{s // 60}m {s % 60:02d}s"
    return f"{s}s"


def fmt_day(d):
    return f"{d:%b} {d.day}"


def fmt_dt(x):
    return "N/A" if pd.isna(x) else pd.Timestamp(x).strftime("%Y-%m-%d %H:%M")


def initials(name):
    return "".join(w[0] for w in str(name).split()[:2]).upper() or "?"


def tone(rate):
    """(status label, colour class) for an agent's suspicious rate."""
    if rate > 20:
        return "Critical Flag", "r"
    if rate > ANOMALY_THRESHOLD:
        return "Under Review", "b"
    if rate < 1:
        return "Exemplary", "g"
    if rate <= 3:
        return "Compliant", "g"
    return "Normal", "n"


def sparkline(vals):
    w, h = 240, 44
    if len(vals) < 2:
        return ""
    v = np.asarray(vals, float)
    rng = (v.max() - v.min()) or 1.0
    xs = np.linspace(0, w, len(v))
    ys = h - 4 - (v - v.min()) / rng * (h - 10)
    pts = " ".join(f"{x:.1f},{y:.1f}" for x, y in zip(xs, ys))
    return (f'<svg class="spark" viewBox="0 0 {w} {h}" preserveAspectRatio="none"><defs>'
            '<linearGradient id="sg" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#6366F1" stop-opacity=".35"/>'
            '<stop offset="1" stop-color="#6366F1" stop-opacity="0"/></linearGradient></defs>'
            f'<polygon points="0,{h} {pts} {w},{h}" fill="url(#sg)"/>'
            f'<polyline points="{pts}" fill="none" stroke="#818CF8" stroke-width="2" vector-effect="non-scaling-stroke"/></svg>')


def donut(parts, big, small):
    c, cum, circles = 2 * np.pi * 42, 0.0, ""
    total = sum(p[1] for p in parts) or 1
    for _, val, color in parts:
        f = val / total
        if f > 0:
            circles += (f'<circle cx="52" cy="52" r="42" fill="none" stroke-width="9" style="stroke:{color}" '
                        f'stroke-dasharray="{f * c:.2f} {c:.2f}" stroke-dashoffset="{-cum * c:.2f}"/>')
        cum += f
    return (f'<svg width="104" height="104" viewBox="0 0 104 104"><g transform="rotate(-90 52 52)">'
            f'<circle cx="52" cy="52" r="42" fill="none" stroke-width="9" style="stroke:var(--track)"/>{circles}</g>'
            f'<text x="52" y="52" text-anchor="middle" style="fill:var(--text);font:600 15px Inter,sans-serif">{big}</text>'
            f'<text x="52" y="66" text-anchor="middle" style="fill:var(--muted);font:500 7px JetBrains Mono,monospace">{small}</text></svg>')


@st.cache_data(show_spinner=False)
def build_xlsx(sheets: dict) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, index=False, sheet_name=name[:31])
    return buf.getvalue()


# =========================================================
# 3. Data: demo fallback + forensic engine
# =========================================================
@st.cache_data(show_spinner=False)
def demo_bytes() -> bytes:
    """Small synthetic sheet so the dashboard renders when no file is available."""
    rng = np.random.default_rng(7)
    agents = [("Marcus Vance", .22), ("Chloe Zhao", .14), ("Elena Rostova", .01),
              ("Devon Larsson", .02), ("Tariq Kanso", .05), ("Aisha Al-Mansoor", .005)]
    cats = ["Network & Wi-Fi", "Account & Access", "Hardware & Peripherals", "Software & Licenses",
            "Email & M365 Collaboration", "ERP & Internal Systems"]
    good = ["Replaced faulty patch cable and verified connectivity.", "Reset MFA token and confirmed sign-in with the user.",
            "Reinstalled the client and applied the latest update.", "Cleared mailbox cache and re-synced Outlook profile."]
    bad = ["Resolved per standard policy.", "Fixed.", "Done."]
    rows = []
    start = pd.Timestamp("2026-10-01")
    for i in range(480):
        name, risk = agents[rng.integers(len(agents))]
        opened = start + pd.Timedelta(minutes=int(rng.integers(0, 30 * 24 * 60)))
        assigned = opened + pd.Timedelta(minutes=int(rng.integers(1, 30)))
        fast = rng.random() < risk
        mins = float(rng.uniform(0.1, 0.5)) if fast else float(rng.gamma(2.5, 9)) + 1
        sol = bad[rng.integers(len(bad))] if (fast or rng.random() < .02) else good[rng.integers(len(good))]
        rows.append({"Ticket ID": 8000 + i, "Subject": f"Support request {8000 + i}", "Category": cats[rng.integers(len(cats))],
                     "Desk Agent": name, "Created By": name if rng.random() < .03 else f"user{rng.integers(1, 200)}", "Open Date": opened,
                     "Assign Date": assigned, "Close Date": assigned + pd.Timedelta(minutes=mins), "Solution": sol,
                     "Resolved": "Solved", "SLA Result": "Met" if mins < SLA_TARGET_MIN else "Breached"})
    buf = BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def process_excel_sheet(raw: bytes) -> pd.DataFrame:
    df = pd.read_excel(BytesIO(raw))

    for col, default in {"Ticket ID": None, "Desk Agent": "Unassigned", "Category": "Uncategorized",
                         "Subject": "", "Created By": "Unknown", "Solution": np.nan}.items():
        if col not in df.columns:
            df[col] = df.index + 1 if col == "Ticket ID" else default
    for col in ["Open Date", "Close Date", "Assign Date"]:
        df[col] = pd.to_datetime(df[col], errors="coerce") if col in df.columns else pd.NaT

    df["Assign_To_Close_Sec"] = (df["Close Date"] - df["Assign Date"]).dt.total_seconds()
    if "Elapsed Minutes" in df.columns:
        df["Resolution_Min"] = pd.to_numeric(df["Elapsed Minutes"], errors="coerce")
    else:
        df["Resolution_Min"] = df["Assign_To_Close_Sec"] / 60.0

    # Rule 1 - closed almost instantly
    rule_fast = (df["Assign_To_Close_Sec"] < FAST_CLOSE_SEC) | (df["Resolution_Min"] < FAST_CLOSE_SEC / 60)

    # Rule 2 - empty / generic solution text
    generic_words = {"done", "ok", "fixed", "تم", "حل", "tam", "dne", "done.", "fixed.", "no", "na", "n/a", ".", ".."}

    def is_generic(sol):
        if pd.isna(sol) or not str(sol).strip():
            return False  # handled by the "missing" rule
        s = str(sol).strip().lower()
        return len(s) <= 4 or s in generic_words

    rule_generic = df["Solution"].apply(is_generic)
    # Rule 3 - no solution text at all
    rule_missing = df["Solution"].isna() | df["Solution"].astype(str).str.strip().eq("")

    # Rule 4 - self-assigned: the ticket creator and the desk agent are the same person
    norm = lambda col: df[col].astype(str).str.strip().str.lower()
    rule_self = (norm("Created By") == norm("Desk Agent")) & df["Created By"].notna() & df["Desk Agent"].notna()

    reasons, tags = [], []
    for f, m, g, sa in zip(rule_fast, rule_missing, rule_generic, rule_self):
        r, t = [], []
        if f:
            r.append("Fast-Closed (<30s)"); t.append("Fast-Closed")
        if sa:
            r.append("Self-Assigned"); t.append("Self-Assigned")
        if m:
            r.append("Missing Solution Text"); t.append("Missing Solution")
        if g:
            r.append("Generic/Spam Solution"); t.append("Generic Solution")
        reasons.append(" | ".join(r) if r else "Verified Genuine")
        tags.append(t[0] if t else "Genuine")

    df["Detection_Reason"] = reasons
    df["Primary_Flag"] = tags
    df["Is_Fake"] = df["Primary_Flag"] != "Genuine"
    return df


# =========================================================
# 4. Callbacks
# =========================================================
def toggle_theme():
    st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"


def rerun_detection():
    process_excel_sheet.clear()
    st.session_state.dismissed = set()
    st.toast("Anomaly detection re-run complete", icon=":material/check_circle:")


def bulk_dismiss():
    keys = [k for k in list(st.session_state) if str(k).startswith("sel_")]
    st.session_state.dismissed |= {int(k[4:]) for k in keys if st.session_state[k]}
    for k in keys:
        del st.session_state[k]


def go_page(p):
    st.session_state.page = p


# =========================================================
# 5. Top bar
# =========================================================
inject_css(st.session_state.theme)
dark = st.session_state.theme == "dark"

with st.container(key="topbar"):
    c = st.columns([1.9, 2.3, 1.6, 1.4, 1.2, 0.4], vertical_alignment="center", gap="small")
    with c[0]:
        md(f'<div class="brand"><span class="logo">{icon("pulse", 18)}</span><b>TicketPulse IT</b>'
           f'<span class="ver">v2.4 Live Telemetry</span></div>')
    with c[1]:
        upload = st.file_uploader("Upload sheet", type=["xlsx", "xls"], label_visibility="collapsed", key="upload")
    with c[2]:
        st.button("Process & Analyze Sheet", type="primary", icon=":material/analytics:", key="process",
                  on_click=rerun_detection, **STRETCH)
    with c[3]:
        md('<div style="text-align:center"><span class="pill-g"><i></i>Anomaly Engine Active</span></div>')
    with c[5]:
        st.button("", icon=":material/light_mode:" if dark else ":material/dark_mode:", key="theme_icon", on_click=toggle_theme)

# =========================================================
# 6. Load + analyse data
# =========================================================
if upload is not None:
    raw, fname = upload.getvalue(), upload.name
elif Path(DEFAULT_FILE).exists():
    raw, fname = Path(DEFAULT_FILE).read_bytes(), DEFAULT_FILE
else:
    raw, fname = demo_bytes(), "Demo_Helpdesk_Log.xlsx"

t0 = time.perf_counter()
try:
    df = process_excel_sheet(raw)
except Exception as exc:  # noqa: BLE001
    st.error(f"Could not read the spreadsheet: {exc}")
    st.stop()
sync_ms = int((time.perf_counter() - t0) * 1000)

if df.empty:
    st.warning("The spreadsheet has no rows to analyse.")
    st.stop()

total = len(df)
flagged_df = df[df["Is_Fake"]].copy()
fake_n = len(flagged_df)
genuine_n = total - fake_n
fake_pct, genuine_pct = fake_n / total * 100, (total - fake_n) / total * 100
tag_n = {t: int((df["Primary_Flag"] == t).sum()) for t in TAGS}
avg_min, med_min = df["Resolution_Min"].mean(), df["Resolution_Min"].median()
fcr_n = int(((df["Resolution_Min"] <= FCR_MAX_MIN) & ~df["Is_Fake"]).sum())
sha = hashlib.sha256(raw).hexdigest()[:7]

date_col = next((c for c in ["Close Date", "Open Date", "Assign Date"] if df[c].notna().any()), None)
daily = pd.Series(dtype=float)
if date_col:
    daily = df[date_col].dropna().dt.normalize().value_counts().sort_index()
    daily = daily.reindex(pd.date_range(daily.index.min(), daily.index.max()), fill_value=0)
ingest_pct = df[date_col].notna().mean() * 100 if date_col else 100.0

# =========================================================
# 7. Layout: left navigation + main area
# =========================================================
with st.container(key="body"):
    nav_col, main_col = st.columns([1, 4.2], gap="small")

with nav_col:
    with st.container(key="sidenav"):
        md('<div class="nav-l">OPERATIONS CONSOLE</div>')
        page_name = st.radio("Navigation", NAV, label_visibility="collapsed", key="nav")
        md(f'<div class="ingest"><div class="row"><span>Ingestion Stream</span><b class="g mono">{ingest_pct:.2f}%</b></div>'
           f'<div class="bar"><i style="width:{ingest_pct:.2f}%;background:var(--green)"></i></div>'
           f'<div class="row s"><span>Sync: {sync_ms}ms</span><span>Batch: {total:,}</span></div></div>')


def page_header(title, sub):
    md(f'<div class="card"><div class="h1" style="margin:0">{esc(title)}</div><div class="sec-s">{esc(sub)}</div></div>')


# ---------------------------------------------------------
# Overview
# ---------------------------------------------------------
def page_overview():
    # ----- hero -----
    if date_col:
        range_txt = f"Last {len(daily)} Days ({fmt_day(daily.index.min())} – {fmt_day(daily.index.max())})"
    else:
        range_txt = "All time"
    size_mb = len(raw) / 1e6
    with st.container(key="hero"):
        left, right = st.columns([1.45, 1.55], vertical_alignment="center")
        with left:
            md(f'<div><div class="h1">Ticket Telemetry &amp; Integrity Audit</div>'
               f'<span class="tag-r">● Live Forensic Matrix</span>'
               f'<div class="row-w"><span class="chipx">{icon("calendar", 14)}{esc(range_txt)}</span></div>'
               f'<div class="row-w"><span class="chipx m">{icon("file", 14)}Dataset: <span class="e" title="{esc(fname)}">{esc(fname)}</span>'
               f' <span style="color:var(--muted)">{size_mb:.1f} MB · {total:,} rows</span></span></div>'
               f'<div class="hash">{icon("check", 14)}Validated Hash: sha256:{sha}</div></div>')
        with right:
            b1, b2, b3 = st.columns([1, 1.35, 1.7])
            b1.button("Dark Mode" if dark else "Light Mode", icon=":material/dark_mode:" if dark else ":material/light_mode:",
                      key="theme_btn", on_click=toggle_theme, **STRETCH)
            b2.button("Re-Run Anomaly Detection", icon=":material/sync:", key="rerun", on_click=rerun_detection,
                      **STRETCH)
            report = build_xlsx({"All Processed Tickets": df.assign(Dismissed=df.index.isin(st.session_state.dismissed)),
                                 "Flagged Tickets Only": flagged_df})
            b3.download_button("Export Cleaned Report (.XLSX)", report, "TicketPulse_IT_Full_Audit_Report.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               icon=":material/download:", type="primary", **STRETCH, key="dl_report")

    # ----- KPI cards -----
    if len(daily) >= 4:
        h = len(daily) // 2
        a, b = daily.iloc[:h].mean(), daily.iloc[h:].mean()
        delta = (b - a) / a * 100 if a else 0.0
        delta_chip = f'<span class="chip {"g" if delta >= 0 else "r"}">{icon("up", 12)}{delta:+.1f}% vs prior half</span>'
    else:
        delta_chip = ""
    avg_daily = f"{daily.mean():.1f}" if len(daily) else "N/A"
    peak = f"{int(daily.max())} ({fmt_day(daily.idxmax())})" if len(daily) else "N/A"
    risk = "High Risk" if fake_pct >= 5 else ("Medium Risk" if fake_pct >= 2 else "Low Risk")
    diff = avg_min - SLA_TARGET_MIN if pd.notna(avg_min) else 0
    spd_chip = (f'<span class="chip {"g" if diff <= 0 else "r"}">{icon("down" if diff <= 0 else "up", 12)}'
                f'{"-" if diff <= 0 else "+"}{fmt_dur(abs(diff)).split()[0]} vs SLA target</span>')
    mix = [("var(--red)", tag_n["Fast-Closed"]), ("#8B8CF6", tag_n["Generic Solution"]), ("#F59E0B", tag_n["Self-Assigned"]), ("#94A3B8", tag_n["Missing Solution"])]
    mix_bar = "".join(f'<i style="width:{v / max(fake_n, 1) * 100:.1f}%;background:{c}"></i>' for c, v in mix)
    md(f'''<div class="kpi-grid">
<div class="kpi"><div class="kpi-h"><span>Total Tickets Ingested</span>{icon("inbox", 18)}</div>
<div class="kpi-v"><span class="num">{total:,}</span>{delta_chip}</div>
{sparkline(daily.values)}
<div class="kpi-f"><span>Avg Daily: {avg_daily}</span><span>Peak: {peak}</span></div></div>
<div class="kpi"><div class="kpi-h"><span>Genuine &amp; Validated</span><span class="chip g">{genuine_pct:.1f}% Verified</span></div>
<div class="kpi-v"><span class="num">{genuine_n:,}</span><span class="g" style="font-size:12px">{icon("check", 14)} High Integrity</span></div>
<div class="bar" style="margin-top:22px"><i style="width:{genuine_pct:.1f}%;background:var(--green)"></i></div>
<div class="kpi-f"><span>{fcr_n:,} first-contact resolution</span><b class="g">{fcr_n / max(genuine_n, 1) * 100:.1f}% FCR</b></div></div>
<div class="kpi"><div class="kpi-h"><span>Suspicious / Fake Flagged</span><span class="chip r">{fake_pct:.1f}% Anomaly</span></div>
<div class="kpi-v"><span class="num">{fake_n:,}</span><span class="r mono" style="font-size:11px">{risk}</span></div>
<div class="tri"><span>Fast-Closed (&lt;{FAST_CLOSE_SEC}s):<b class="r">{tag_n["Fast-Closed"]}</b></span>
<span>Spam:<b>{tag_n["Generic Solution"]}</b></span><span>Self:<b>{tag_n["Self-Assigned"]}</b></span><span>Ghost:<b>{tag_n["Missing Solution"]}</b></span></div>
<div class="bar">{mix_bar}</div></div>
<div class="kpi"><div class="kpi-h"><span>Avg. Resolution Speed</span><span style="color:var(--muted)">{icon("timer", 18)}</span></div>
<div class="kpi-v"><span class="num">{fmt_dur(avg_min, True)}</span>{spd_chip}</div>
<div class="kpi-f" style="margin-top:auto;padding-top:14px"><span>Median Speed:<br><b style="color:var(--text)">{fmt_dur(med_min, True)}</b></span>
<span>SLA Target:<br><b style="color:var(--text)">{fmt_dur(SLA_TARGET_MIN, True)}</b></span></div></div>
</div>''')

    # ----- leaderboard + categories -----
    ag = df.groupby("Desk Agent").agg(closed=("Is_Fake", "size"), flagged=("Is_Fake", "sum"),
                                      avg=("Resolution_Min", "mean")).reset_index()
    ag["rate"] = ag["flagged"] / ag["closed"] * 100
    top_cat = df.groupby("Desk Agent")["Category"].agg(lambda s: s.mode().iat[0] if not s.mode().empty else "")
    ag = ag.sort_values("closed", ascending=False)
    rows = ""
    for _, r in ag.head(8).iterrows():
        label, cls = tone(r["rate"])
        hot = pd.notna(r["avg"]) and r["avg"] < 1
        dur = (f'<span class="r">{icon("bolt", 13)} {fmt_dur(r["avg"], True)}</span>' if hot else fmt_dur(r["avg"], True))
        rows += (f'<div class="lb-r"><div class="who"><span class="av {cls}">{esc(initials(r["Desk Agent"]))}</span>'
                 f'<div style="min-width:0"><div class="nm">{esc(r["Desk Agent"])}</div>'
                 f'<div class="sub">{esc(top_cat[r["Desk Agent"]])}</div></div></div>'
                 f'<div class="c mono" style="font-size:12.5px">{int(r["closed"])}</div>'
                 f'<div class="c"><span class="chip {cls}">{r["rate"]:.1f}% ({int(r["flagged"])})</span></div>'
                 f'<div class="c mono" style="font-size:12px">{dur}</div>'
                 f'<div class="c"><span class="chip {cls}">{label}</span></div></div>')
    n_crit = int((ag["rate"] > 20).sum())
    esc_txt = f'{n_crit} agent{"s" if n_crit != 1 else ""} {"require" if n_crit != 1 else "requires"} executive escalation' if n_crit else ""

    cats = df.groupby("Category").agg(n=("Is_Fake", "size"), f=("Is_Fake", "sum")).reset_index().sort_values("n", ascending=False)
    cat_rows = ""
    for _, r in cats.head(6).iterrows():
        rate = r["f"] / r["n"] * 100
        hot = rate >= ANOMALY_THRESHOLD
        note = f'{int(r["n"])} · {rate:.1f}% {"HIGH SPIKE" if hot else "flagged"}'
        cat_rows += (f'<div class="cat"><div class="cat-t"><span>{esc(r["Category"])}</span><span class="{"hot" if hot else ""}">{note}</span></div>'
                     f'<div class="bar"><i style="width:{100 - rate:.1f}%;background:var(--green)"></i>'
                     f'<i style="width:{rate:.1f}%;background:var(--red)"></i></div></div>')
    parts = [("Genuine", genuine_n, "var(--green)"), ("Generic", tag_n["Generic Solution"], "#8B8CF6"),
             ("Fast-Closed", tag_n["Fast-Closed"], "#FB7185"), ("Self-Assigned", tag_n["Self-Assigned"], "#F59E0B"), ("Missing", tag_n["Missing Solution"], "#94A3B8")]
    legend = "".join(f'<span><i style="background:{c}"></i>{n} ({v / total * 100:.1f}%)</span>' for n, v, c in parts)
    md(f'''<div class="mid">
<div class="card"><div class="sec-h"><div><div class="sec-t">{icon("users", 20)}Employee Leaderboard &amp; Risk Audit</div>
<div class="sec-s">Monitoring agent velocity, closure anomalies, and repetitive templated resolutions.</div></div>
<span class="chip">{len(ag)} Active Analysts</span></div>
<div class="lb"><div class="lb-r hd"><div>AGENT PROFILE</div><div class="c">CLOSED</div><div class="c">SUSPICIOUS RATE</div>
<div class="c">AVG DURATION</div><div class="c">AUDIT STATUS</div></div>{rows}</div>
<div class="foot"><span>Anomaly Threshold: &gt; {ANOMALY_THRESHOLD:.0f}% Suspicious Rate</span><span class="r">{esc_txt}</span></div></div>
<div class="card"><div class="sec-h"><div><div class="sec-t">{icon("pie", 20)}Category Breakdown &amp; Integrity</div>
<div class="sec-s">Correlating ticket volume with anomaly spike velocity across departmental queues.</div></div>
<span class="chip">{len(cats)} Incident Types</span></div>{cat_rows}
<div class="donut">{donut(parts, f"{genuine_pct:.1f}%", "CLEAN")}<div class="legend">{legend}</div></div></div>
</div>''')

    # ----- deep inspection -----
    dismissed = st.session_state.dismissed
    remaining = flagged_df[~flagged_df.index.isin(dismissed)]
    counts = {"all": len(remaining), **{t: int((remaining["Primary_Flag"] == t).sum()) for t in TAGS}}

    with st.container(key="inspect"):
        h1, h2, h3, h4 = st.columns([1.5, 1.5, 0.8, 0.9], vertical_alignment="center")
        with h1:
            md(f'<div class="sec-t">{icon("flag", 20)}<span>Suspicious Tickets Deep Inspection</span>'
               f'<span class="chip r">{counts["all"]} Anomalies Detected</span></div>')
        with h2:
            q = st.text_input("Search", placeholder="Search ID, agent, solution...", label_visibility="collapsed", key="q")
        view = remaining
        if q.strip():
            hay = (remaining["Ticket ID"].astype(str) + " " + remaining["Desk Agent"].astype(str) + " "
                   + remaining["Solution"].fillna("").astype(str) + " " + remaining["Subject"].astype(str)).str.lower()
            view = remaining[hay.str.contains(q.strip().lower(), regex=False)]
        with h3:
            st.download_button("Export XLSX", build_xlsx({"Flagged Tickets": view.drop(columns=["Assign_To_Close_Sec"])}),
                               "TicketPulse_Flagged_Tickets.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               icon=":material/download:", key="dl_flagged", **STRETCH)
        with h4:
            st.button("Bulk Dismiss", icon=":material/done_all:", key="bulk", on_click=bulk_dismiss, **STRETCH)

        with st.container(key="tabs"):
            tab = st.radio("Filter", ["all"] + TAGS, horizontal=True, label_visibility="collapsed", key="tab",
                           format_func=lambda k: f"{TAB_LABELS[k]} `{counts[k]}`")
        if tab != "all":
            view = view[view["Primary_Flag"] == tab]
        view = view.sort_values(date_col, ascending=False, na_position="last") if date_col else view.sort_index(ascending=False)

        sig = (tab, q, len(dismissed))
        if st.session_state.get("sig") != sig:
            st.session_state.page, st.session_state.sig = 1, sig
        pages = max(1, -(-len(view) // PAGE_SIZE))
        page = min(max(st.session_state.page, 1), pages)
        chunk = view.iloc[(page - 1) * PAGE_SIZE: page * PAGE_SIZE]

        W = [0.35, 1.0, 1.7, 1.6, 0.9, 3.1, 1.7, 0.45]
        with st.container(key="tablewrap"):
            with st.container(key="thead"):
                for col, title in zip(st.columns(W, vertical_alignment="center"),
                                      ["", "TICKET ID", "CATEGORY", "ASSIGNED EMPLOYEE", "DURATION",
                                       "DOCUMENTED SOLUTION NOTE", "DETECTION FLAG", ""]):
                    col.markdown(f'<div class="th">{title}</div>', unsafe_allow_html=True)
            with st.container(key="tbl"):
                if chunk.empty:
                    md('<div class="empty">No flagged tickets match this view.</div>')
                for i, r in chunk.iterrows():
                    cols = st.columns(W, vertical_alignment="center")
                    cols[0].checkbox("Select ticket", key=f"sel_{i}", label_visibility="collapsed")
                    cols[1].markdown(f'<div class="tid">{esc(TICKET_PREFIX)}{esc(r["Ticket ID"])}</div>', unsafe_allow_html=True)
                    cols[2].markdown(f'<div class="cell">{esc(r["Category"])}</div>', unsafe_allow_html=True)
                    cols[3].markdown(f'<div class="cell"><span class="av {"r" if r["Is_Fake"] else ""}">{esc(initials(r["Desk Agent"]))}</span>'
                                     f'<span class="nm">{esc(r["Desk Agent"])}</span></div>', unsafe_allow_html=True)
                    fast = r["Primary_Flag"] == "Fast-Closed" or "Fast-Closed" in r["Detection_Reason"]
                    d = fmt_dur(r["Resolution_Min"])
                    cols[4].markdown(f'<div class="cell mono">{f"""<span class="chip r">{icon("timer", 12)}{d}</span>""" if fast else d}</div>',
                                     unsafe_allow_html=True)
                    sol = r["Solution"]
                    sol_txt = "(no solution recorded)" if pd.isna(sol) or not str(sol).strip() else f'"{sol}"'
                    cols[5].markdown(f'<div class="note" title="{esc(sol_txt)}">{esc(sol_txt)}</div>', unsafe_allow_html=True)
                    reasons = r["Detection_Reason"].split(" | ")
                    chip_cls = {"Fast-Closed": "r", "Self-Assigned": "r", "Generic Solution": "b"}.get(r["Primary_Flag"], "")
                    more = f" +{len(reasons) - 1}" if len(reasons) > 1 else ""
                    cols[6].markdown(f'<div class="cell"><span class="chip {chip_cls}">{esc(reasons[0])}{more}</span></div>',
                                     unsafe_allow_html=True)
                    with cols[7].popover("", icon=":material/visibility:"):
                        st.markdown(f"**Ticket {TICKET_PREFIX}{r['Ticket ID']}**")
                        st.markdown(f"**Subject:** {r['Subject']}\n\n**Desk Agent:** {r['Desk Agent']}\n\n"
                                    f"**Category:** {r['Category']}\n\n**Created By:** {r.get('Created By', 'Unknown')}\n\n"
                                    f"**Logged Duration:** {fmt_dur(r['Resolution_Min'], True)}\n\n"
                                    f"**Opened:** {fmt_dt(r['Open Date'])}\n\n**Closed:** {fmt_dt(r['Close Date'])}\n\n"
                                    f"**SLA Result:** {r.get('SLA Result', 'N/A')}\n\n**Detection:** {r['Detection_Reason']}")
                        st.info(sol_txt)

        f1, f2 = st.columns([2.2, 2.8], vertical_alignment="center")
        with f1:
            md(f'<div class="foot" style="margin:0">Showing {len(chunk)} of {len(view)} flagged tickets · Page {page} of {pages}</div>')
        with f2:
            with st.container(key="pager"):
                start = min(max(1, page - 1), max(1, pages - 2))
                nums = list(range(start, min(start + 2, pages) + 1))
                pc = st.columns([3, 1.4] + [0.7] * len(nums) + [1], vertical_alignment="center", gap="small")
                pc[1].button("Previous", key="prev", disabled=page <= 1, on_click=go_page, args=(page - 1,), **STRETCH)
                for col, n in zip(pc[2:2 + len(nums)], nums):
                    col.button(str(n), key=f"pg_{n}", type="primary" if n == page else "secondary", on_click=go_page, args=(n,),
                               **STRETCH)
                pc[-1].button("Next", key="next", disabled=page >= pages, on_click=go_page, args=(page + 1,), **STRETCH)


# ---------------------------------------------------------
# Other pages
# ---------------------------------------------------------
def page_stream():
    page_header("Live Ticket Stream", "Every ticket in the loaded sheet with its detection result.")
    c1, c2 = st.columns([1, 2])
    status = c1.selectbox("Status", ["All tickets", "Flagged only", "Genuine only"])
    term = c2.text_input("Search tickets", placeholder="Search ID, agent, subject or solution...")
    view = df if status == "All tickets" else (flagged_df if status == "Flagged only" else df[~df["Is_Fake"]])
    if term.strip():
        hay = (view["Ticket ID"].astype(str) + " " + view["Desk Agent"].astype(str) + " " + view["Subject"].astype(str)
               + " " + view["Solution"].fillna("").astype(str)).str.lower()
        view = view[hay.str.contains(term.strip().lower(), regex=False)]
    cols = [c for c in ["Ticket ID", "Category", "Desk Agent", "Subject", "Resolution_Min", "Solution", "Detection_Reason"] if c in view]
    st.dataframe(view[cols], hide_index=True, **STRETCH,
                 column_config={"Resolution_Min": st.column_config.NumberColumn("Duration (min)", format="%.1f"),
                                "Detection_Reason": "Detection Flag"})


def page_agents():
    page_header("Agent Audit", "Detailed risk analytics for a single analyst.")
    agent = st.selectbox("Select an agent", sorted(df["Desk Agent"].dropna().unique()))
    adf = df[df["Desk Agent"] == agent]
    a_fake = int(adf["Is_Fake"].sum())
    m1, m2, m3 = st.columns(3)
    m1.metric("Total tickets", len(adf))
    m2.metric("Suspicious tickets", a_fake)
    m3.metric("Suspicious rate", f"{a_fake / len(adf) * 100:.1f}%")
    cols = [c for c in ["Ticket ID", "Category", "Subject", "Resolution_Min", "Solution", "Is_Fake", "Detection_Reason"] if c in adf]
    st.dataframe(adf[cols], hide_index=True, **STRETCH,
                 column_config={"Resolution_Min": st.column_config.NumberColumn("Duration (min)", format="%.1f"),
                                "Is_Fake": "Flagged", "Detection_Reason": "Detection Flag"})


def page_rules():
    page_header("Rules Engine", "The checks that decide whether a closed ticket is flagged.")
    md(f'''<div class="card"><div class="nm" style="font-size:14px">Fast-Closed</div>
<div class="sec-s">Assign-to-close time under {FAST_CLOSE_SEC} seconds, or a logged duration under {FAST_CLOSE_SEC / 60:.1f} minutes.</div>
<div class="nm" style="font-size:14px;margin-top:16px">Self-Assigned</div>
<div class="sec-s">The ticket creator (Created By) and the desk agent are the same person.</div>
<div class="nm" style="font-size:14px;margin-top:16px">Missing Solution</div>
<div class="sec-s">The solution field is empty.</div>
<div class="nm" style="font-size:14px;margin-top:16px">Generic Solution</div>
<div class="sec-s">The solution is 4 characters or fewer, or a known filler such as "done", "ok" or "fixed".</div>
<div class="nm" style="font-size:14px;margin-top:16px">Agent audit status</div>
<div class="sec-s">Critical Flag above 20% suspicious, Under Review above {ANOMALY_THRESHOLD:.0f}%, Exemplary below 1%,
Compliant up to 3%, otherwise Normal.</div></div>''')


def page_reports():
    page_header("Reports", "Download the audited data as Excel workbooks.")
    st.download_button("Full audit report (.XLSX)", build_xlsx({"All Processed Tickets": df, "Flagged Tickets Only": flagged_df}),
                       "TicketPulse_IT_Full_Audit_Report.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       icon=":material/download:", type="primary", key="dl_full")
    st.download_button("Flagged tickets only (.XLSX)", build_xlsx({"Flagged Tickets Only": flagged_df}),
                       "TicketPulse_Flagged_Tickets.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       icon=":material/download:", key="dl_flag_only")


with main_col:
    with st.container(key="main"):
        {"Overview & Fraud Matrix": page_overview, "Live Ticket Stream": page_stream, "Agent Audit": page_agents,
         "Rules Engine": page_rules, "Reports": page_reports}[page_name]()