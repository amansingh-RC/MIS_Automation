"""
Shared UI / branding for the RCL MIS Automation app.

Keeps all styling in one place so app.py stays focused on logic and adding a
new tab is a few lines. Royal Chain palette: deep navy + gold.
"""

from __future__ import annotations

import base64
from functools import lru_cache
from pathlib import Path

import streamlit as st

_LOGO_PATH = Path(__file__).parent / "assets" / "royal_chain_logo.png"

_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cinzel:wght@600;700;900&family=Rajdhani:wght@400;500;600;700&display=swap');

:root {
  --gold: #C9A84C;
  --gold-light: #F0D080;
  --dark-navy: #090E1A;
  --navy: #0D1B3E;
  --navy-mid: #1A2B5E;
  --navy-card: #111827;
  --border: rgba(201,168,76,0.22);
  --text-white: #F5F0E8;
  --text-muted: #8899BB;
}

/* App background with subtle radial glows */
.stApp {
  background:
    radial-gradient(ellipse 60% 40% at 8% 12%, rgba(201,168,76,0.06) 0%, transparent 60%),
    radial-gradient(ellipse 50% 55% at 92% 88%, rgba(26,43,94,0.45) 0%, transparent 65%),
    var(--dark-navy);
}
[data-testid="stHeader"] { background: transparent; }

/* Center + constrain content */
.block-container { max-width: 1080px; padding-top: 1.2rem; }

* { font-family: 'Rajdhani', sans-serif; }

/* ---- Branded header ---- */
.rcl-header {
  text-align: center;
  padding: 26px 0 18px;
  margin-bottom: 8px;
}
.rcl-logo-card {
  display: inline-block;
  background: #FFFFFF;
  border: 1px solid var(--border);
  border-radius: 14px;
  padding: 16px 26px;
  box-shadow: 0 8px 28px rgba(0,0,0,0.35), 0 0 22px rgba(201,168,76,0.18);
}
.rcl-logo-card img { height: 84px; width: auto; display: block; }
.rcl-sub {
  margin-top: 14px;
  font-size: 12px; letter-spacing: 0.35em; color: var(--text-muted);
  text-transform: uppercase; font-weight: 600;
}
.rcl-rule {
  width: 130px; height: 2px; margin: 14px auto 0;
  background: linear-gradient(90deg, transparent, var(--gold), transparent);
}

/* ---- Section heading ---- */
.rcl-section-title {
  font-family: 'Cinzel', serif; font-size: 18px; color: var(--gold-light);
  letter-spacing: 0.06em; margin: 6px 0 2px;
}
.rcl-section-sub { font-size: 14px; color: var(--text-muted); margin-bottom: 14px; }

/* ---- Tabs ---- */
[data-testid="stTabs"] [role="tablist"] {
  gap: 6px; border-bottom: 1px solid var(--border); flex-wrap: wrap;
}
[data-testid="stTabs"] [role="tab"] {
  background: rgba(13,27,62,0.45);
  border: 1px solid var(--border); border-bottom: none;
  border-radius: 10px 10px 0 0;
  padding: 9px 20px; color: var(--text-muted);
  font-family: 'Cinzel', serif; font-size: 13px; letter-spacing: 0.06em;
}
[data-testid="stTabs"] [role="tab"]:hover { color: var(--gold-light); }
[data-testid="stTabs"] [aria-selected="true"] {
  background: linear-gradient(135deg, var(--navy-mid), var(--navy));
  color: var(--gold) !important;
  box-shadow: 0 -2px 14px rgba(201,168,76,0.12);
}
[data-testid="stTabs"] [role="tab"] p { font-size: 13px !important; font-weight: 700; }

/* ---- File uploader ---- */
[data-testid="stFileUploaderDropzone"] {
  background: rgba(13,27,62,0.4);
  border: 2px dashed var(--border); border-radius: 12px;
  transition: all 0.25s ease;
}
[data-testid="stFileUploaderDropzone"]:hover {
  border-color: var(--gold); background: rgba(201,168,76,0.06);
}

/* ---- Buttons (process / generic) ---- */
.stButton > button, .stFormSubmitButton > button {
  background: linear-gradient(135deg, var(--navy-mid), var(--navy));
  border: 1px solid var(--gold); color: var(--gold);
  font-family: 'Cinzel', serif; letter-spacing: 0.1em; font-weight: 700;
  border-radius: 9px; transition: all 0.25s ease;
}
.stButton > button:hover {
  background: rgba(201,168,76,0.12); color: var(--gold-light);
  box-shadow: 0 0 20px rgba(201,168,76,0.2);
}

/* ---- Download button (solid gold) ---- */
[data-testid="stDownloadButton"] > button {
  background: linear-gradient(135deg, var(--gold), #A8862E);
  border: none; color: var(--dark-navy);
  font-family: 'Cinzel', serif; font-weight: 700; letter-spacing: 0.08em;
  border-radius: 8px;
}
[data-testid="stDownloadButton"] > button:hover {
  transform: translateY(-1px); box-shadow: 0 6px 18px rgba(201,168,76,0.4);
  color: var(--dark-navy);
}

/* ---- Metric cards ---- */
[data-testid="stMetric"] {
  background: var(--navy-card); border: 1px solid var(--border);
  border-radius: 10px; padding: 14px 16px; position: relative; overflow: hidden;
}
[data-testid="stMetric"]::before {
  content: ''; position: absolute; top: 0; left: 0; right: 0; height: 2px;
  background: linear-gradient(90deg, transparent, var(--gold), transparent);
}
[data-testid="stMetricValue"] { color: var(--gold-light); font-weight: 700; }
[data-testid="stMetricLabel"] { color: var(--text-muted); }

/* ---- Alerts, expanders, dataframe ---- */
[data-testid="stExpander"] {
  border: 1px solid var(--border); border-radius: 10px;
  background: rgba(13,27,62,0.3);
}
[data-testid="stExpander"] summary { color: var(--gold-light); }
[data-testid="stDataFrame"] {
  border: 1px solid var(--border); border-radius: 10px;
}

/* ---- Footer ---- */
.rcl-footer {
  text-align: center; margin-top: 40px; padding-top: 18px;
  border-top: 1px solid var(--border);
  font-size: 12px; color: var(--text-muted); letter-spacing: 0.06em;
}
</style>
"""


def inject_css() -> None:
    st.markdown(_CSS, unsafe_allow_html=True)


@lru_cache(maxsize=1)
def _logo_data_uri() -> str:
    """Base64 data URI of the local logo (embedded so it works offline)."""
    try:
        data = _LOGO_PATH.read_bytes()
    except OSError:
        return ""
    return "data:image/png;base64," + base64.b64encode(data).decode("ascii")


def render_header() -> None:
    uri = _logo_data_uri()
    logo = (f'<div class="rcl-logo-card"><img src="{uri}" alt="Royal Chain '
            f'Limited"></div>' if uri else
            '<div class="rcl-name">Royal Chain Limited</div>')
    st.markdown(
        f"""
        <div class="rcl-header">
          {logo}
          <div class="rcl-sub">MIS Report Automation</div>
          <div class="rcl-rule"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_title(title: str, subtitle: str = "") -> None:
    html = f'<div class="rcl-section-title">{title}</div>'
    if subtitle:
        html += f'<div class="rcl-section-sub">{subtitle}</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_footer() -> None:
    st.markdown(
        '<div class="rcl-footer">© Royal Chain Limited &nbsp;·&nbsp; '
        'MIS Report Automation &nbsp;·&nbsp; Internal Tool</div>',
        unsafe_allow_html=True,
    )
