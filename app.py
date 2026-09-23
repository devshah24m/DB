"""
Booked Profit Dashboard — Streamlit version (premium dark theme).

Runs entirely in the cloud (Streamlit Community Cloud, free): logs into
Angel One, holds a live WebSocket price feed in a background thread, and
renders the same open/closed positions view as the original local HTML
dashboard — but reachable from any device, with your PC turned off.

Local files reused unchanged from the original project:
    positions_builder.py   — FIFO matching (open + closed positions)
    token_resolver.py      — Angel One instrument-master token lookup

Secrets required (set in Streamlit Cloud's "Secrets" panel, never in code):
    APP_PASSWORD        — shared password gate for viewers
    ANGEL_API_KEY
    ANGEL_CLIENT_CODE
    ANGEL_PASSWORD
    ANGEL_TOTP_SECRET

The trade ledger is fetched live as JSON from a Google Apps Script Web App
bound to the Sheet (see Code.gs — deploy it, paste the /exec URL into the
sidebar, or default it via st.secrets["APPS_SCRIPT_URL"]) rather than an
uploaded/local Excel file or a CSV export link. This means: no "Anyone with
the link" sharing requirement (the script runs under your own account), no
CSV parsing/format guessing, and today's ledger updates without a redeploy
— just edit the sheet and hit "Restart feed / refetch sheet".

LIVE UPDATES
    Instead of refreshing the whole page every few seconds (old
    st_autorefresh approach — flickers, resets scroll position, reruns
    everything including the sidebar), the KPI cards + tables now live
    inside an @st.fragment(run_every=...) block. Only that fragment reruns
    on its own clock, reading whatever ticks have landed in the background
    WebSocket thread since the last run — so the screen updates in near
    real time, tick by tick, without touching the rest of the app.
    Requires streamlit >= 1.33.
"""
import os
import re
import threading
import time
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from html import escape as _html_escape
from urllib.parse import quote_plus
from zoneinfo import ZoneInfo

import altair as alt
import pandas as pd
import pyotp
import requests
import streamlit as st
from SmartApi import SmartConnect
from SmartApi.smartWebSocketV2 import SmartWebSocketV2

from positions_builder import load_trade_ledger, load_trade_ledger_from_records, build_positions
from token_resolver import resolve_all

IST = ZoneInfo("Asia/Kolkata")


def _fetch_ledger_records(apps_script_url: str, sheet_name: str | None = None, raw: bool = False):
    """Hits the Apps Script Web App's /exec URL and returns its JSON body
    directly — no CSV parsing involved. See Code.gs for the doGet handler
    this talks to. Normally returns a list of row-dicts (header-keyed);
    with raw=True (used for the block-structured per-client Rollover tab,
    which has no single header row) it instead returns the raw 2-D grid
    of cell values, unmodified.
    """
    params = {"sheet": sheet_name} if sheet_name else {}
    if raw:
        params["raw"] = "1"
    params = params or None

    last_err = None
    for attempt in range(3):
        try:
            resp = requests.get(apps_script_url, params=params, timeout=45)
            resp.raise_for_status()
            data = resp.json()
            if isinstance(data, dict) and "error" in data:
                raise ValueError(f"Apps Script error: {data['error']}")
            return data
        except requests.exceptions.ReadTimeout as e:
            last_err = e
            if attempt < 2:
                time.sleep(2 * (attempt + 1))  # 2s, then 4s backoff
                continue
            raise TimeoutError(
                "The ledger service (Google Apps Script) didn't respond in time "
                "after 3 attempts. It may be cold-starting or the sheet is large — "
                "please try refreshing in a moment."
            ) from e
    raise last_err

st.set_page_config(
    page_title="Booked Profit Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

MASTER_CACHE_PATH = "instrument_master_cache.json"
MASTER_CACHE_MAX_AGE_HOURS = 20
SUBSCRIBE_MODE = 2
MAX_TOKENS_PER_SUBSCRIBE = 1000
TICK_REFRESH_SECONDS = 1  # how often the live fragment re-renders


# ── Premium theme ──────────────────────────────────────────────────────────
def inject_theme():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Carlito:ital,wght@0,400;0,500;0,600;0,700;1,400&display=swap');

        :root {
            --bg: #070b14;
            --panel: #10161f;
            --panel-2: #131b27;
            --border: #1e2733;
            --accent: #34d5c8;
            --accent-2: #7c5cff;
            --pos: #2ee6a6;
            --neg: #ff5c7a;
            --muted: #7f8ba3;
            --text: #eaf0f7;
        }

        html, body, [class*="css"]  { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; }

        .stApp {
            background:
                radial-gradient(circle at 15% 0%, rgba(124,92,255,0.10), transparent 40%),
                radial-gradient(circle at 85% 10%, rgba(52,213,200,0.08), transparent 35%),
                var(--bg);
            color: var(--text);
            overflow-x: hidden; /* belt-and-braces: nothing on this page should ever force a horizontal swipe */
        }
        html, body { overflow-x: hidden; }

        section[data-testid="stSidebar"] {
            background: var(--panel);
            border-right: 1px solid var(--border);
        }

        div.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1300px; }
        @media (max-width: 640px) {
            div.block-container { padding-left: 0.9rem; padding-right: 0.9rem; padding-top: 1rem; }
        }

        /* Header */
        .db-header {
            display: flex; align-items: center; justify-content: space-between;
            padding-bottom: 6px; margin-bottom: 22px;
            border-bottom: 1px solid var(--border);
        }
        .db-title { display: flex; align-items: center; gap: 12px; }
        .db-title h1 {
            font-size: 1.65rem; font-weight: 800; letter-spacing: -0.02em; margin: 0;
            background: linear-gradient(90deg, #ffffff, #b9c4d6);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }
        .db-icon {
            width: 40px; height: 40px; border-radius: 12px;
            background: linear-gradient(135deg, var(--accent), var(--accent-2));
            display: flex; align-items: center; justify-content: center;
            font-size: 1.15rem; box-shadow: 0 0 24px rgba(52,213,200,0.35);
        }

        .status-pill {
            display: inline-flex; align-items: center; gap: 8px;
            padding: 6px 14px; border-radius: 999px; font-size: 0.78rem; font-weight: 600;
            border: 1px solid var(--border); background: var(--panel-2); color: var(--muted);
        }
        .status-dot { width: 8px; height: 8px; border-radius: 50%; background: var(--muted); }
        .status-live .status-dot { background: var(--pos); box-shadow: 0 0 8px var(--pos); animation: pulse 1.4s ease-in-out infinite; }
        .status-live { color: var(--pos); border-color: rgba(46,230,166,0.25); }
        .status-error .status-dot { background: var(--neg); }
        .status-error { color: var(--neg); border-color: rgba(255,92,122,0.25); }
        .status-warn .status-dot { background: #f5b942; }
        .status-warn { color: #f5b942; border-color: rgba(245,185,66,0.25); }
        @keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.35; } }

        /* KPI cards */
        .kpi-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin-bottom: 26px; }
        .kpi-card {
            background: linear-gradient(180deg, var(--panel-2), var(--panel));
            border: 1px solid var(--border); border-radius: 16px; padding: 18px 20px;
            position: relative;
            container-type: inline-size; /* lets kpi-value size off THIS card's actual width */
        }
        .kpi-card::before {
            content: ""; position: absolute; top: 0; left: 0; right: 0; height: 3px;
            background: linear-gradient(90deg, var(--accent), var(--accent-2)); opacity: 0.85;
            border-radius: 16px 16px 0 0; /* rounds the bar itself now that the card no longer clips overflow */
        }
        .kpi-label { font-size: 0.74rem; font-weight: 600; text-transform: uppercase; letter-spacing: 0.06em; color: var(--muted); margin-bottom: 8px; white-space: nowrap; }
        /* cqw = % of this card's own width, so the value shrinks exactly as much as its
           card needs regardless of how many KPI cards share the row. No overflow/ellipsis
           here on purpose — a clipped digit on a money figure is worse than a smaller font. */
        .kpi-value { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: clamp(0.72rem, 8.5cqw, 1.55rem); font-weight: 700; letter-spacing: -0.01em; white-space: nowrap; display: block; }
        .kpi-pos { color: var(--pos); }
        .kpi-neg { color: var(--neg); }
        .kpi-sub { font-size: 0.72rem; color: var(--muted); margin-top: 6px; font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; }

        /* Section headers */
        .section-label {
            display: flex; align-items: center; gap: 8px;
            font-size: 0.95rem; font-weight: 700; margin: 6px 0 10px 0; color: var(--text);
        }
        .section-label .badge {
            font-size: 0.68rem; font-weight: 600; padding: 2px 9px; border-radius: 999px;
            background: var(--panel-2); border: 1px solid var(--border); color: var(--muted);
        }

        /* Tabs */
        .stTabs [data-baseweb="tab-list"] { gap: 4px; border-bottom: 1px solid var(--border); }
        .stTabs [data-baseweb="tab"] {
            background: transparent; color: var(--muted); font-weight: 600; padding: 10px 18px;
        }
        .stTabs [aria-selected="true"] {
            color: var(--accent) !important; border-bottom: 2px solid var(--accent) !important;
        }

        /* Top status + live clock bar */
        .top-bar {
            display: flex; align-items: center; justify-content: space-between;
            flex-wrap: wrap; background: var(--panel-2); border: 1px solid var(--border);
            border-radius: 12px; padding: 10px 18px; margin-bottom: 20px; gap: 8px 14px;
        }
        .top-clock {
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.82rem; color: var(--text);
            display: flex; align-items: center; gap: 10px; flex-wrap: nowrap; white-space: nowrap;
        }
        .top-clock .date-part { color: var(--muted); }
        .top-clock .tz-badge {
            font-size: 0.65rem; font-weight: 700; color: var(--accent);
            background: rgba(52,213,200,0.1); border: 1px solid rgba(52,213,200,0.25);
            padding: 1px 7px; border-radius: 999px;
        }
        /* On phones, drop the weekday ("Thursday, ") so the whole clock
           strip ("17 Sep 2026  11:37:33  IST") reliably fits one line
           instead of wrapping across three. */
        @media (max-width: 640px) {
            .top-bar { padding: 10px 14px; }
            .top-clock .weekday-part { display: none; }
        }

        /* Segment-scoped one-line summary (shown inside each open/closed
           segment tab — reflects ONLY that segment, not the whole book).
           A responsive grid rather than a free-flowing flex row, so on a
           narrow phone the stats stack into a clean 1-column list instead
           of wrapping unevenly mid-row. */
        .seg-summary {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr));
            gap: 10px 22px; align-items: center;
            background: var(--panel-2); border: 1px solid var(--border); border-radius: 12px;
            padding: 12px 18px; margin-bottom: 16px;
        }
        .seg-stat { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
        @media (max-width: 480px) {
            .seg-summary { grid-template-columns: 1fr; }
        }
        .seg-stat-label {
            font-size: 0.68rem; font-weight: 600; text-transform: uppercase;
            letter-spacing: 0.05em; color: var(--muted);
        }
        .seg-stat-value {
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif;
            font-size: 0.88rem; font-weight: 700; white-space: nowrap;
        }

        /* Position cards */
        .pos-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 8px; }
        @media (max-width: 640px) { .pos-grid { grid-template-columns: 1fr; } }
        .pos-card {
            background: linear-gradient(180deg, var(--panel-2), var(--panel));
            border: 1px solid var(--border); border-radius: 14px; padding: 14px 16px;
            transition: border-color 0.2s ease;
        }
        .pos-card:hover { border-color: rgba(52,213,200,0.35); }

        /* Highest daily gain card — animated RGB glow so it pops out of the grid.
           Recomputed every refresh, so this always follows whichever position
           currently has the top daily gain rather than sitting on one symbol. */
        .pos-card.top-gain {
            position: relative;
            border-color: transparent;
            background:
                linear-gradient(180deg, var(--panel-2), var(--panel)) padding-box,
                conic-gradient(from var(--rgb-angle, 0deg), #ff3cac, #784ba0, #2b86c5, #2ee6a6, #f5b942, #ff3cac) border-box;
            border: 2px solid transparent;
            animation: rgb-spin 4s linear infinite;
            box-shadow: 0 0 22px rgba(124,92,255,0.35), 0 0 42px rgba(52,213,200,0.18);
        }
        @property --rgb-angle {
            syntax: '<angle>'; inherits: false; initial-value: 0deg;
        }
        @keyframes rgb-spin {
            to { --rgb-angle: 360deg; }
        }
        .top-gain-badge {
            position: absolute; top: -10px; right: 14px;
            font-size: 0.6rem; font-weight: 800; letter-spacing: 0.04em;
            padding: 2px 9px; border-radius: 999px; text-transform: uppercase;
            background: linear-gradient(90deg, #ff3cac, #784ba0, #2b86c5, #2ee6a6);
            background-size: 300% 100%; animation: rgb-shift 3s linear infinite;
            color: #05070d; box-shadow: 0 0 10px rgba(124,92,255,0.5);
        }
        @keyframes rgb-shift {
            0% { background-position: 0% 50%; }
            100% { background-position: 300% 50%; }
        }
        .pos-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 10px; }
        .pos-symbol { font-weight: 700; font-size: 0.98rem; letter-spacing: -0.01em; }
        .pos-tags { display: flex; gap: 5px; margin-top: 4px; }
        .pos-chip {
            font-size: 0.62rem; font-weight: 700; padding: 1px 7px; border-radius: 999px;
            background: var(--panel); border: 1px solid var(--border); color: var(--muted);
            text-transform: uppercase; letter-spacing: 0.03em;
        }
        .pos-chip.long { color: var(--pos); border-color: rgba(46,230,166,0.3); }
        .pos-chip.short { color: var(--neg); border-color: rgba(255,92,122,0.3); }
        .pos-daychg {
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.72rem; font-weight: 700;
            padding: 3px 9px; border-radius: 999px; white-space: nowrap;
        }
        .day-pos { background: rgba(46,230,166,0.12); color: var(--pos); }
        .day-neg { background: rgba(255,92,122,0.12); color: var(--neg); }
        .pos-rows { display: grid; grid-template-columns: 1fr 1fr; gap: 6px 14px; margin-bottom: 10px; }
        .pos-row-label { font-size: 0.68rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
        .pos-row-value { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.82rem; font-weight: 600; }
        .pos-mtm-block {
            display: flex; justify-content: space-between; align-items: center;
            border-top: 1px solid var(--border); padding-top: 10px;
        }
        .pos-mtm-label { font-size: 0.68rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.03em; }
        .pos-mtm-value { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 1.02rem; font-weight: 700; }
        .pos-tick { font-size: 0.66rem; color: var(--muted); font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; }

        /* Closed position cards */
        .closed-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-bottom: 20px; }
        @media (max-width: 640px) { .closed-grid { grid-template-columns: 1fr; } }
        .closed-card {
            background: linear-gradient(180deg, var(--panel-2), var(--panel));
            border: 1px solid var(--border); border-radius: 14px; padding: 13px 15px;
        }
        .closed-top { display: flex; justify-content: space-between; align-items: flex-start; margin-bottom: 8px; }
        .closed-symbol { font-weight: 700; font-size: 0.92rem; }
        .closed-badge {
            font-size: 0.62rem; font-weight: 700; padding: 1px 7px; border-radius: 999px;
            background: var(--panel); border: 1px solid var(--border); color: var(--muted);
        }
        .closed-rows { display: flex; justify-content: space-between; font-size: 0.72rem; color: var(--muted); margin-bottom: 4px; font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; }
        .closed-pnl { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-weight: 700; font-size: 0.95rem; margin-top: 8px; }

        /* Position table — borderless rows, just a thin separator line
           between each position, replacing the boxed-card layout. */
        .pos-table-wrap { overflow-x: auto; margin-bottom: 20px; }
        .pos-table { width: 100%; min-width: 760px; border-collapse: collapse; }
        .pos-table-head, .pos-table-row {
            display: grid;
            align-items: center;
            column-gap: 14px;
        }
        .pos-table-head.cols-open, .pos-table-row.cols-open {
            grid-template-columns: 2fr 0.8fr 0.8fr 1fr 1fr 1fr 1.2fr 1fr;
        }
        .pos-table-head.cols-open-opt, .pos-table-row.cols-open-opt {
            grid-template-columns: 1.7fr 0.65fr 0.5fr 0.8fr 0.7fr 0.9fr 0.9fr 0.8fr 1.1fr 0.9fr;
        }
        .pos-table-head.cols-closed, .pos-table-row.cols-closed {
            grid-template-columns: 1.5fr 0.8fr 0.8fr 1fr 1fr 1fr 1.2fr;
        }
        .pos-table-head {
            padding: 6px 6px 10px 6px;
            border-bottom: 1px solid var(--border);
        }
        .pos-table-head > div {
            font-size: 0.66rem; font-weight: 700; text-transform: uppercase;
            letter-spacing: 0.05em; color: var(--muted);
        }
        .pos-table-row {
            padding: 13px 6px;
            border-bottom: 1px solid var(--border);
            transition: background 0.15s ease;
        }
        .pos-table-row:hover { background: rgba(255,255,255,0.025); }
        .pos-table-row:last-child { border-bottom: none; }
        .pt-symbol { display: flex; align-items: baseline; gap: 8px; flex-wrap: wrap; overflow: visible; min-width: 0; }
        .pt-symbol-name { font-weight: 700; font-size: 0.9rem; letter-spacing: -0.01em; white-space: nowrap; }
        .pt-tag {
            font-size: 0.58rem; font-weight: 700; padding: 1px 6px; border-radius: 999px;
            border: 1px solid var(--border); color: var(--muted); text-transform: uppercase;
            white-space: nowrap; flex-shrink: 0;
        }
        .pt-tag.long { color: var(--pos); border-color: rgba(46,230,166,0.3); }
        .pt-tag.short { color: var(--neg); border-color: rgba(255,92,122,0.3); }
        .pt-tag.rolled {
            color: var(--accent-2); border-color: rgba(124,92,255,0.35);
            background: rgba(124,92,255,0.08); cursor: default;
            padding: 1px 5px; font-size: 0.68rem; line-height: 1;
        }
        .pt-cell { font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.85rem; font-weight: 600; }
        .pt-cell.muted { color: var(--muted); font-weight: 500; font-size: 0.78rem; }
        .pt-cell.pos { color: var(--pos); }
        .pt-cell.neg { color: var(--neg); }
        .pt-arrow { font-size: 0.72rem; margin-right: 2px; }

        /* Leader row (top gain / least loss) — no box, just a soft glowing
           left accent bar + tinted background so it stands out among plain
           rows without reintroducing card borders. */
        .pos-table-row.top-gain-row {
            position: relative;
            background: linear-gradient(90deg, rgba(124,92,255,0.10), transparent 60%);
            border-left: 3px solid;
            border-image: linear-gradient(180deg, #ff3cac, #784ba0, #2b86c5, #2ee6a6) 1;
        }
        .pt-leader-badge {
            display: inline-flex; align-items: center; gap: 4px; margin-left: 8px;
            font-size: 0.58rem; font-weight: 800; letter-spacing: 0.04em;
            padding: 1px 8px; border-radius: 999px; text-transform: uppercase;
            background: linear-gradient(90deg, #ff3cac, #784ba0, #2b86c5, #2ee6a6);
            background-size: 300% 100%; animation: rgb-shift 3s linear infinite;
            color: #05070d; white-space: nowrap; flex-shrink: 0;
        }

        /* Collapsible F&O summary row — a stock/index with multiple lots
           (e.g. rolled from one expiry to the next) collapses into ONE
           qty-weighted-average line by default; expanding it reveals the
           original per-lot rows with their real sell dates. Built on the
           native <details>/<summary> pair, so no JS is needed. */
        details.pt-details { border-bottom: 1px solid var(--border); }
        details.pt-details:last-child { border-bottom: none; }
        summary.pos-table-row { list-style: none; cursor: pointer; border-bottom: none; }
        summary.pos-table-row::-webkit-details-marker { display: none; }
        summary.pos-table-row:hover { background: rgba(255,255,255,0.025); }
        .pt-expand-chevron {
            display: inline-block; font-size: 0.65rem; color: var(--muted);
            transition: transform 0.18s ease; flex-shrink: 0;
        }
        details.pt-details[open] summary .pt-expand-chevron { transform: rotate(90deg); color: var(--accent); }
        .pt-lot-count {
            font-size: 0.6rem; font-weight: 700; color: var(--accent);
            background: rgba(52,213,200,0.1); border: 1px solid rgba(52,213,200,0.25);
            padding: 1px 8px; border-radius: 999px; white-space: nowrap; flex-shrink: 0;
        }
        .pt-details-body { background: rgba(255,255,255,0.015); padding: 2px 0 6px 0; }
        .pt-details-body .pos-table-row:last-child { border-bottom: none; }
        .pt-details-label {
            font-size: 0.64rem; font-weight: 600; color: var(--muted);
            text-transform: uppercase; letter-spacing: 0.05em; padding: 8px 6px 2px 6px;
        }

        /* Mobile reflow — desktop grid (above) is untouched. Below 640px the
           8-col (open) / 7-col (closed) grid no longer fits, so instead of
           forcing a horizontal scroll we restack each row into a compact
           2-column card: symbol on top, then labelled qty/avg/cmp pairs,
           then the P&L full-width at the bottom since that's the number
           that actually matters at a glance on a phone. */
        @media (max-width: 640px) {
            .pos-table-wrap { overflow-x: visible; }
            .pos-table { min-width: 0; }
            .pos-table-head.cols-open, .pos-table-head.cols-closed, .pos-table-head.cols-open-opt { display: none; }

            .pos-table-row.cols-open {
                grid-template-columns: 1fr 1fr;
                grid-template-areas:
                    "sym   sym"
                    "exch  tick"
                    "qty   avg"
                    "cmp   day"
                    "mtm   mtm";
                row-gap: 6px;
                column-gap: 10px;
                padding: 14px 10px;
            }
            .pos-table-row.cols-open > div:nth-child(1) { grid-area: sym; }
            .pos-table-row.cols-open > div:nth-child(2) { grid-area: exch; }
            .pos-table-row.cols-open > div:nth-child(3) { grid-area: qty; }
            .pos-table-row.cols-open > div:nth-child(4) { grid-area: avg; }
            .pos-table-row.cols-open > div:nth-child(5) { grid-area: cmp; }
            .pos-table-row.cols-open > div:nth-child(6) { grid-area: day; text-align: right; }
            .pos-table-row.cols-open > div:nth-child(7) {
                grid-area: mtm; text-align: right; font-size: 1rem; margin-top: 2px;
            }
            .pos-table-row.cols-open > div:nth-child(8) { grid-area: tick; text-align: right; }
            .pos-table-row.cols-open > div:nth-child(3)::before { content: "Qty "; color: var(--muted); font-weight: 500; }
            .pos-table-row.cols-open > div:nth-child(4)::before { content: "Avg "; color: var(--muted); font-weight: 500; }
            .pos-table-row.cols-open > div:nth-child(5)::before { content: "CMP "; color: var(--muted); font-weight: 500; }

            .pos-table-row.cols-open-opt {
                grid-template-columns: 1fr 1fr;
                grid-template-areas:
                    "sym    sym"
                    "strike type"
                    "exp    tick"
                    "qty    avg"
                    "cmp    day"
                    "mtm    mtm";
                row-gap: 6px;
                column-gap: 10px;
                padding: 14px 10px;
            }
            .pos-table-row.cols-open-opt > div:nth-child(1) { grid-area: sym; }
            .pos-table-row.cols-open-opt > div:nth-child(2) { grid-area: strike; }
            .pos-table-row.cols-open-opt > div:nth-child(3) { grid-area: type; }
            .pos-table-row.cols-open-opt > div:nth-child(4) { grid-area: exp; }
            .pos-table-row.cols-open-opt > div:nth-child(5) { grid-area: qty; }
            .pos-table-row.cols-open-opt > div:nth-child(6) { grid-area: avg; }
            .pos-table-row.cols-open-opt > div:nth-child(7) { grid-area: cmp; }
            .pos-table-row.cols-open-opt > div:nth-child(8) { grid-area: day; text-align: right; }
            .pos-table-row.cols-open-opt > div:nth-child(9) {
                grid-area: mtm; text-align: right; font-size: 1rem; margin-top: 2px;
            }
            .pos-table-row.cols-open-opt > div:nth-child(10) { grid-area: tick; text-align: right; }
            .pos-table-row.cols-open-opt > div:nth-child(2)::before { content: "Strike "; color: var(--muted); font-weight: 500; }
            .pos-table-row.cols-open-opt > div:nth-child(5)::before { content: "Qty "; color: var(--muted); font-weight: 500; }
            .pos-table-row.cols-open-opt > div:nth-child(6)::before { content: "Avg "; color: var(--muted); font-weight: 500; }
            .pos-table-row.cols-open-opt > div:nth-child(7)::before { content: "CMP "; color: var(--muted); font-weight: 500; }

            .pos-table-row.cols-closed {
                grid-template-columns: 1fr 1fr;
                grid-template-areas:
                    "sym   sym"
                    "exch  date"
                    "qty   buy"
                    "sell  sell"
                    "pnl   pnl";
                row-gap: 6px;
                column-gap: 10px;
                padding: 14px 10px;
            }
            .pos-table-row.cols-closed > div:nth-child(1) { grid-area: sym; }
            .pos-table-row.cols-closed > div:nth-child(2) { grid-area: exch; }
            .pos-table-row.cols-closed > div:nth-child(3) { grid-area: qty; }
            .pos-table-row.cols-closed > div:nth-child(4) { grid-area: buy; }
            .pos-table-row.cols-closed > div:nth-child(5) { grid-area: sell; }
            .pos-table-row.cols-closed > div:nth-child(6) { grid-area: date; text-align: right; }
            .pos-table-row.cols-closed > div:nth-child(7) {
                grid-area: pnl; text-align: right; font-size: 1rem; margin-top: 2px;
            }
            .pos-table-row.cols-closed > div:nth-child(3)::before { content: "Qty "; color: var(--muted); font-weight: 500; }
            .pos-table-row.cols-closed > div:nth-child(4)::before { content: "Buy "; color: var(--muted); font-weight: 500; }
            .pos-table-row.cols-closed > div:nth-child(5)::before { content: "Sell "; color: var(--muted); font-weight: 500; }
        }

        /* Chart cards: st.markdown('<div class="chart-card">') + st.altair_chart(...) +
           st.markdown('</div>') used to be 3 SEPARATE calls — Streamlit renders each call
           as its own DOM node, so that div opened and closed empty and the real chart sat
           outside it, unstyled. Charts now render inside `with st.container(border=True):`,
           which is a genuine parent element, and we restyle Streamlit's own wrapper for it
           below so it matches the rest of the theme instead of Streamlit's default grey box. */
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background: var(--panel-2) !important; border: 1px solid var(--border) !important;
            border-radius: 14px !important; margin-bottom: 18px !important;
        }
        div[data-testid="stVerticalBlockBorderWrapper"] > div { border-radius: 14px !important; }

        /* Dataframe polish (still used where a raw table makes sense) */
        div[data-testid="stDataFrame"] {
            border: 1px solid var(--border); border-radius: 12px; overflow: hidden;
        }

        div[data-testid="stMetric"] { display: none; }  /* using custom KPI cards instead */

        /* Hero MTM block — big headline number + status pill, used for the
           top-of-dashboard "Live position MTM" and the closed-tab "Booked
           ledger" summary. */
        .hero-block {
            display: flex; justify-content: space-between; align-items: flex-start;
            flex-wrap: wrap; gap: 14px; margin-bottom: 18px;
        }
        .hero-label {
            display: flex; align-items: center; gap: 8px;
            font-size: 0.72rem; font-weight: 700; text-transform: uppercase;
            letter-spacing: 0.08em; color: var(--muted); margin-bottom: 10px;
        }
        .hero-label .dot { width: 6px; height: 6px; border-radius: 50%; }
        .hero-label .dot.positive { background: var(--pos); box-shadow: 0 0 8px var(--pos); }
        .hero-label .dot.negative { background: var(--neg); box-shadow: 0 0 8px var(--neg); }
        .hero-value {
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif;
            font-size: clamp(1.9rem, 4.2vw, 3rem); font-weight: 800; letter-spacing: -0.02em;
            line-height: 1.1;
        }
        .hero-value.positive { color: var(--pos); }
        .hero-value.negative { color: var(--neg); }
        .hero-sub { font-size: 0.82rem; color: var(--muted); margin-top: 6px; }
        .hero-pill {
            display: inline-flex; align-items: center; gap: 6px;
            padding: 7px 16px; border-radius: 999px; font-size: 0.72rem; font-weight: 700;
            letter-spacing: 0.04em; text-transform: uppercase; border: 1px solid; white-space: nowrap;
        }
        .hero-pill.positive { color: var(--pos); border-color: rgba(46,230,166,0.4); background: rgba(46,230,166,0.08); }
        .hero-pill.negative { color: var(--neg); border-color: rgba(255,92,122,0.4); background: rgba(255,92,122,0.08); }

        /* Two-value hero variant — used for "Live Position MTM" so the
           headline shows current MTM *with* booked profit folded in right
           next to the running MTM of open positions only, instead of
           forcing the person to scan down to the small tiles to tell them
           apart. */
        .hero-dual { display: flex; flex-wrap: wrap; gap: 26px; }
        .hero-dual-item { min-width: 170px; }
        .hero-dual-item .hero-value { font-size: clamp(1.5rem, 3.2vw, 2.35rem); }
        .hero-dual-divider { width: 1px; align-self: stretch; background: var(--border); min-height: 60px; }
        @media (max-width: 640px) {
            .hero-dual { gap: 16px; }
            .hero-dual-divider { display: none; }
        }

        /* Plain bordered stat tiles under a hero block. Wraps into as many
           rows as needed (auto-fit) so every tile is always visible without
           ever needing a horizontal swipe — on mobile this collapses to a
           tidy 2-column grid instead of a cut-off horizontally-scrolling
           strip. */
        .hero-tiles {
            display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 10px; margin-bottom: 26px;
        }
        @media (max-width: 640px) {
            .hero-tiles { grid-template-columns: repeat(2, 1fr); gap: 8px; }
        }
        .hero-tile {
            background: var(--panel-2); border: 1px solid var(--border); border-radius: 12px;
            padding: 11px 14px; min-width: 0;
        }
        .hero-tile-label {
            font-size: 0.58rem; font-weight: 700; text-transform: uppercase;
            letter-spacing: 0.04em; color: var(--muted); margin-bottom: 6px;
            white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
        }
        .hero-tile-value {
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif;
            font-size: 0.88rem; font-weight: 700; white-space: nowrap;
            overflow: hidden; text-overflow: ellipsis;
        }
        .hero-tile-value.positive { color: var(--pos); }
        .hero-tile-value.negative { color: var(--neg); }

        /* Rollover-chain leg — a compact stat card replacing a wide
           st.dataframe (Date/From/To/Qty/prices...) that forced horizontal
           scrolling on mobile. Same visual language as .hero-tile. */
        .roll-chain-card {
            background: var(--panel-2); border: 1px solid var(--border); border-radius: 12px;
            padding: 12px 14px; margin-bottom: 10px;
        }
        .roll-chain-head {
            display: flex; align-items: center; justify-content: space-between;
            margin-bottom: 10px; padding-bottom: 8px; border-bottom: 1px solid var(--border);
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif; font-size: 0.8rem; font-weight: 700;
        }
        .roll-chain-arrow { color: var(--accent); font-weight: 700; }
        .roll-chain-grid {
            display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px 8px;
        }
        @media (max-width: 640px) {
            .roll-chain-grid { grid-template-columns: repeat(2, 1fr); }
        }
        .roll-chain-stat-label {
            font-size: 0.6rem; font-weight: 600; text-transform: uppercase;
            letter-spacing: 0.04em; color: var(--muted); margin-bottom: 3px;
        }
        .roll-chain-stat-value {
            font-family: 'Calibri', 'Carlito', 'Segoe UI', sans-serif;
            font-size: 0.8rem; font-weight: 700;
        }

        .news-item {
            padding: 12px 10px 12px 12px; border-bottom: 1px solid var(--border);
            border-left: 3px solid transparent; transition: background 0.15s;
        }
        .news-item:last-child { border-bottom: none; }
        /* Most recent (<24h) gets the strongest highlight, "recent" (1-3
           days) a subtler one, and 3-7 days old is left plain — reusing
           the app's existing accent/accent-2 colors rather than a new hue. */
        .news-item.news-new { border-left-color: var(--accent); background: rgba(52,213,200,0.07); }
        .news-item.news-recent { border-left-color: var(--accent-2); background: rgba(124,92,255,0.05); }
        .news-title {
            color: var(--text); font-weight: 600; font-size: 0.92rem;
            text-decoration: none; line-height: 1.4;
        }
        .news-title:hover { color: var(--accent, #4da3ff); text-decoration: underline; }
        .news-badge {
            display: inline-block; font-size: 0.58rem; font-weight: 800;
            letter-spacing: 0.04em; text-transform: uppercase; margin-left: 8px;
            padding: 1px 7px; border-radius: 999px; vertical-align: middle; white-space: nowrap;
        }
        .news-badge-new {
            background: rgba(52,213,200,0.15); color: var(--accent);
            border: 1px solid rgba(52,213,200,0.4);
        }
        .news-badge-recent {
            background: rgba(124,92,255,0.15); color: var(--accent-2);
            border: 1px solid rgba(124,92,255,0.4);
        }
        .news-badge-stock {
            background: var(--panel-2); color: var(--muted);
            border: 1px solid var(--border); font-weight: 700;
        }
        .news-meta {
            color: var(--muted); font-size: 0.74rem; margin-top: 4px;
        }
        </style>
        <style>
        """,
        unsafe_allow_html=True,
    )


def status_pill(state: str, label: str) -> str:
    cls = {"live": "status-live", "error": "status-error", "warn": "status-warn"}.get(state, "")
    return f'<span class="status-pill {cls}"><span class="status-dot"></span>{label}</span>'


def flat(html: str) -> str:
    """Collapse a multi-line, indented HTML f-string to a single line.

    Streamlit's markdown renderer runs HTML through a CommonMark parser
    before allowing it through. An HTML block continues being treated as
    raw HTML only until a blank line; after that, any line indented 4+
    spaces (which our nested Python f-strings produce naturally) is read
    as an *indented code block* and shown as literal text instead of being
    rendered — this is what caused the closed-position cards to print
    "<div class=..." as visible text after the first card. Stripping all
    newlines/indentation removes any chance of a stray blank line or deep
    indentation confusing the parser, regardless of how deeply the
    generating Python code is nested.
    """
    return re.sub(r"\s*\n\s*", "", html.strip())


def kpi_card(label, value, positive=None, sub=None):
    cls = "kpi-pos" if positive is True else ("kpi-neg" if positive is False else "")
    sub_html = f'<div class="kpi-sub">{sub}</div>' if sub else ""
    return flat(f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value {cls}">{value}</div>
            {sub_html}
        </div>
    """)


def seg_summary_line(items):
    """items: list of (label, value_html, extra_css_class) tuples — a single
    horizontal strip of stats scoped to whichever segment tab it's rendered
    inside, so it never mixes numbers across segments."""
    stats = "".join(
        f'<div class="seg-stat"><span class="seg-stat-label">{lbl}</span>'
        f'<span class="seg-stat-value {cls}">{val}</span></div>'
        for lbl, val, cls in items
    )
    return flat(f'<div class="seg-summary">{stats}</div>')


def hero_block(label, value, sub, is_positive, pill_text):
    sign_cls = "positive" if is_positive else "negative"
    return flat(f"""
        <div class="hero-block">
            <div>
                <div class="hero-label"><span class="dot {sign_cls}"></span>{label}</div>
                <div class="hero-value {sign_cls}">{value}</div>
                <div class="hero-sub">{sub}</div>
            </div>
            <div class="hero-pill {sign_cls}">{'▲' if is_positive else '▼'} {pill_text}</div>
        </div>
    """)


def hero_block_dual(
    primary_label, primary_value, primary_sub, primary_is_positive,
    secondary_label, secondary_value, secondary_sub, secondary_is_positive,
    pill_text,
):
    """Two headline values side by side instead of one — e.g. current MTM
    WITH booked profit folded in, next to the running MTM of open positions
    ONLY (no booked profit). The status pill always follows the primary
    (booked + open) figure, since that's the number the pill's Profit/Loss
    language describes."""
    p_cls = "positive" if primary_is_positive else "negative"
    s_cls = "positive" if secondary_is_positive else "negative"
    return flat(f"""
        <div class="hero-block">
            <div class="hero-dual">
                <div class="hero-dual-item">
                    <div class="hero-label"><span class="dot {p_cls}"></span>{primary_label}</div>
                    <div class="hero-value {p_cls}">{primary_value}</div>
                    <div class="hero-sub">{primary_sub}</div>
                </div>
                <div class="hero-dual-divider"></div>
                <div class="hero-dual-item">
                    <div class="hero-label"><span class="dot {s_cls}"></span>{secondary_label}</div>
                    <div class="hero-value {s_cls}">{secondary_value}</div>
                    <div class="hero-sub">{secondary_sub}</div>
                </div>
            </div>
            <div class="hero-pill {p_cls}">{'▲' if primary_is_positive else '▼'} {pill_text}</div>
        </div>
    """)


def hero_tiles(items):
    """items: list of (label, value_html, extra_css_class) tuples rendered
    as plain bordered stat boxes under a hero_block."""
    tiles = "".join(
        f'<div class="hero-tile"><div class="hero-tile-label">{lbl}</div>'
        f'<div class="hero-tile-value {cls}">{val}</div></div>'
        for lbl, val, cls in items
    )
    return flat(f'<div class="hero-tiles">{tiles}</div>')


# ── Access log ─────────────────────────────────────────────────────────────
# Shared across every visitor's session via st.cache_resource (a
# session_state list would only be visible to that one visitor's own
# browser tab). In-memory only: resets whenever the app restarts/redeploys,
# same lifetime as get_engine's cache. Capped so it can't grow unbounded
# over a long-running deployment.
_ACCESS_LOG_MAX = 500


@st.cache_resource
def _access_log_store():
    return {"lock": threading.Lock(), "entries": []}


def _record_access(name):
    store = _access_log_store()
    with store["lock"]:
        store["entries"].append({"name": name, "ts": datetime.now(IST)})
        if len(store["entries"]) > _ACCESS_LOG_MAX:
            del store["entries"][: len(store["entries"]) - _ACCESS_LOG_MAX]


def _get_access_log():
    store = _access_log_store()
    with store["lock"]:
        return list(store["entries"])


# ── Client-wise login ─────────────────────────────────────────────────────
#
# Secrets format (set in Streamlit Cloud → Secrets):
#
# [clients.ROHITH]
# password       = "secret123"
# display_name   = "Rohith Sir"
# apps_script_url = "https://script.google.com/macros/s/ABC.../exec"
# sheet_name     = "Rohith"          # optional — sheet tab name
#
# [clients.PRIYA]
# password       = "priya456"
# display_name   = "Priya"
# apps_script_url = "https://script.google.com/macros/s/XYZ.../exec"
# sheet_name     = "Priya"
#
# Optionally, one admin client sees ALL positions concatenated:
# [clients.ADMIN]
# password        = "adminpass"
# display_name    = "Admin"
# is_admin        = true              # sees all clients' data merged
# apps_script_url = "..."
#
# Falls back to the old flat APP_PASSWORD / APPS_SCRIPT_URL secrets if no
# [clients] table is defined (backward-compatible).

def _get_clients():
    """Return dict of client_id -> config dict from st.secrets."""
    try:
        raw = st.secrets.get("clients", {})
        if not raw:
            # Backward-compat: single shared password
            return {
                "DEFAULT": {
                    "password":        st.secrets.get("APP_PASSWORD", ""),
                    "display_name":    "User",
                    "apps_script_url": st.secrets.get("APPS_SCRIPT_URL", ""),
                    "sheet_name":      st.secrets.get("APPS_SCRIPT_SHEET_NAME", ""),
                    "is_admin":        False,
                }
            }
        default_url = st.secrets.get("APPS_SCRIPT_URL", "")
        clients = {}
        for k, v in raw.items():
            cfg = dict(v)
            cfg.setdefault("apps_script_url", default_url)
            if not cfg.get("apps_script_url"):
                cfg["apps_script_url"] = default_url
            clients[k.upper()] = cfg
        return clients
    except Exception:
        return {}


def check_password():
    """Show login form. On success stores client config in session_state.
    Returns True only when a valid client is logged in."""

    if st.session_state.get("_client_ok"):
        return True

    inject_theme()
    st.markdown(
        flat("""
        <div class="db-title" style="justify-content:center; margin: 60px 0 24px 0;">
            <div class="db-icon">📊</div>
            <h1>Booked Profit Dashboard</h1>
        </div>
        """),
        unsafe_allow_html=True,
    )

    clients = _get_clients()

    def _submit():
        cid   = st.session_state.get("_login_id", "").strip().upper()
        pw    = st.session_state.get("_login_pw", "")
        cfg   = clients.get(cid)
        if cfg and pw == cfg.get("password", ""):
            st.session_state["_client_ok"]     = True
            st.session_state["_client_id"]     = cid
            st.session_state["_client_cfg"]    = cfg
            st.session_state["_client_name"]   = cfg.get("display_name", cid)
            _record_access(cfg.get("display_name", cid))
        else:
            st.session_state["_client_ok"]    = False
            st.session_state["_login_failed"] = True

    c1, c2, c3 = st.columns([1, 1.1, 1])
    with c2:
        with st.container(border=True):
            st.markdown(
                "<p style='text-align:center;color:var(--muted);font-size:.85rem;"
                "margin-bottom:18px;'>Sign in to view your portfolio</p>",
                unsafe_allow_html=True,
            )
            st.text_input("Client ID", key="_login_id",
                          placeholder="e.g. ROHITH",
                          help="Your unique client code — ask your broker for this.")
            st.text_input("Password", type="password", key="_login_pw",
                          on_change=_submit,
                          placeholder="Enter your password")
            st.button("Sign in →", on_click=_submit, use_container_width=True, type="primary")

            if st.session_state.get("_login_failed") and not st.session_state.get("_client_ok"):
                st.error("Invalid Client ID or password.")

    return False


def current_client_cfg() -> dict:
    """Return the logged-in client's config dict."""
    return st.session_state.get("_client_cfg", {})


def current_client_name() -> str:
    return st.session_state.get("_client_name", "")


def is_admin() -> bool:
    return current_client_cfg().get("is_admin", False)


# ── Background engine: login + FIFO build + token resolve + live WS feed ──
_ROLLOVER_FIELDS = [
    "Date", "From Series", "To Series", "Qty Out",
    "Last Month Price", "Roll Sell Price", "Roll Buy Price",
    "Roll Diff", "Carry-Forward Price",
]


def _parse_rollover_grid(values):
    """Parse the raw grid of a per-client Rollover tab into
    {STOCK: [entry, ...]}. The tab is laid out as a human-readable report
    (matching the client-facing Excel export): a row with just the
    stock/index name, then a blank row, then a header row starting with
    'Date', then that stock's data rows, then a blank row, repeated per
    stock. The header LABELS after 'Date' vary block to block (e.g. 'Qty
    Out (sold)' vs 'Qty Rolled', 'Carry-Forward Price' vs 'Carry-Fwd
    Price'), but their COLUMN POSITION doesn't, so this reads positionally
    off the header row's first non-blank column rather than matching label
    text: Date, From Series, To Series, Qty Out, Last Month Price, Roll
    Sell Price, Roll Buy Price, Roll Diff, Carry-Forward Price, in that
    fixed order.
    """
    by_stock = {}
    current_stock = None
    data_col = None

    for row in values:
        non_blank = [(i, c) for i, c in enumerate(row) if str(c).strip() != ""]
        if not non_blank:
            continue  # blank row — separates blocks
        if len(non_blank) == 1:
            _, val = non_blank[0]
            sval = str(val).strip()
            if sval.lower() == "date":
                continue  # a lone 'Date' cell isn't a real header row
            current_stock = sval.upper()
            data_col = None
            continue
        first_idx, first_val = non_blank[0]
        if str(first_val).strip().lower() == "date":
            data_col = first_idx  # this block's header row — remember where data starts
            continue
        if current_stock is None or data_col is None:
            continue  # a data-looking row before we know which block it belongs to
        entry = {
            field: (row[data_col + offset] if data_col + offset < len(row) else None)
            for offset, field in enumerate(_ROLLOVER_FIELDS)
        }
        by_stock.setdefault(current_stock, []).append(entry)

    return by_stock


# ── Shared live feed: ONE Angel One login + ONE websocket for the whole
# app, shared by every client ───────────────────────────────────────────
# Angel One's SmartAPI only allows a single live session per client code:
# logging in a second time silently invalidates the previous JWT/feed
# token, which drops that earlier websocket ("WebSocket connection
# error"). Every client dashboard authenticates with the SAME broker
# account (ANGEL_API_KEY / ANGEL_CLIENT_CODE / ANGEL_PASSWORD /
# ANGEL_TOTP_SECRET are global secrets, not per-client), so each visitor
# used to trigger their own login+socket and kick everyone else off.
# SharedFeed logs in exactly once — cached process-wide via
# st.cache_resource, so it's the same instance no matter how many
# visitors/sessions hit the app — and multiplexes every client's tokens
# over that one connection. Raw ticks (ltp/open/close/volume) are the
# same for every client and are cached here; qty/avgPrice/mtm are
# per-client and stay in each client's own LiveEngine (see snapshot()).
# ETF / bond classification. The ledger's Type column (BOND / ETF / EQ) is
# carried as InstrumentType; positions_builder tags every NSE cash row as
# "Equity", which is why ETFs and SGBs never reached the Bonds/ETF tab.
# Symbols that must always show under Bonds/ETF (add more here any time).
BOND_ETF_SYMBOLS = {
    "SGBJAN29IX", "SGBMAY29I", "SILVERBEES", "SILVERIETF", "JUNIORBEES",
    "NIFTYBEES", "BANKBEES", "BANKNIFTY1", "CASHIETF", "GOLDBEES", "LIQUIDBEES",
}
_BOND_ETF_TYPES = {"BOND", "BONDS", "ETF", "SGB", "GSEC", "G-SEC", "GOLDBOND", "MF"}


def effective_segment(segment, instrument_type="", symbol=""):
    """Return 'Bonds/ETF' for ETF/bond rows; otherwise keep the segment."""
    if segment == "F&O":
        return segment
    it = str(instrument_type or "").strip().upper()
    # Angel One trading symbols carry a series suffix ("SILVERBEES-EQ");
    # strip it so name-based matching sees the bare ETF name.
    sym = re.sub(r"-(EQ|BE|BL|BZ|SM|N\d)$", "", str(symbol or "").strip().upper())
    if sym in BOND_ETF_SYMBOLS or it in _BOND_ETF_TYPES or sym.startswith("SGB") or sym.endswith("BEES") or sym.endswith("ETF"):
        return "Bonds/ETF"
    return segment


class SharedFeed:
    def __init__(self):
        self.lock = threading.Lock()
        self.latest_raw = {}       # token -> raw tick dict (ltp/open/close/volume/ts)
        self.status = "starting"   # starting | live | disconnected | error
        self.error = None
        self.last_tick_ts = None
        self._sws = None
        self._connected = False
        self._exch_by_token = {}   # token -> exchangeType, everything subscribed so far
        self._pending = {}         # token -> exchangeType, registered but not yet subscribed
        threading.Thread(target=self._login_and_connect, daemon=True).start()

    def _login_and_connect(self):
        try:
            totp = pyotp.TOTP(st.secrets["ANGEL_TOTP_SECRET"]).now()
            sc = SmartConnect(api_key=st.secrets["ANGEL_API_KEY"])
            data = sc.generateSession(st.secrets["ANGEL_CLIENT_CODE"], st.secrets["ANGEL_PASSWORD"], totp)
            if not data.get("status"):
                with self.lock:
                    self.status = "error"
                    self.error = f"Angel One login failed: {data}"
                return
            jwt_token = data["data"]["jwtToken"]
            feed_token = data["data"]["feedToken"]
            self._start_ws(jwt_token, feed_token)
        except Exception as e:
            with self.lock:
                self.status = "error"
                self.error = str(e)

    def _start_ws(self, jwt_token, feed_token):
        sws = SmartWebSocketV2(
            auth_token=jwt_token, api_key=st.secrets["ANGEL_API_KEY"],
            client_code=st.secrets["ANGEL_CLIENT_CODE"], feed_token=feed_token,
        )

        def on_open(wsapp):
            with self.lock:
                self._connected = True
                self.status = "live"
                to_subscribe = dict(self._pending)
                self._pending = {}
            self._subscribe(to_subscribe)

        def on_data(wsapp, message):
            token = str(message.get("token"))
            with self.lock:
                prev = self.latest_raw.get(token, {})
                now = datetime.now(IST)
                self.latest_raw[token] = {
                    "ltp": message.get("last_traded_price", 0) / 100.0,
                    "prev_ltp": prev.get("ltp"),
                    "close": (message.get("closed_price", 0) / 100.0 if message.get("closed_price") else prev.get("close")),
                    "open": (message.get("open_price_of_the_day", 0) / 100.0 if message.get("open_price_of_the_day") else prev.get("open")),
                    "volume": message.get("volume_trade_for_the_day"),
                    "ts": now.isoformat(timespec="seconds"),
                }
                self.last_tick_ts = now

        def on_error(wsapp, error):
            with self.lock:
                self.status = "error"
                self.error = f"WebSocket error: {error}"

        def on_close(wsapp, *a):
            with self.lock:
                self._connected = False
                self.status = "disconnected"

        sws.on_open = on_open
        sws.on_data = on_data
        sws.on_error = on_error
        sws.on_close = on_close
        with self.lock:
            self._sws = sws
        threading.Thread(target=sws.connect, daemon=True).start()

    def _chunk(self, exch_by_token: dict):
        by_exch = {}
        for token, exch_type in exch_by_token.items():
            by_exch.setdefault(exch_type, []).append(token)
        batches, current, current_count = [], [], 0
        for exch_type, tokens in by_exch.items():
            for i in range(0, len(tokens), 200):
                chunk = tokens[i:i + 200]
                if current_count + len(chunk) > MAX_TOKENS_PER_SUBSCRIBE:
                    batches.append(current)
                    current, current_count = [], 0
                current.append({"exchangeType": exch_type, "tokens": chunk})
                current_count += len(chunk)
        if current:
            batches.append(current)
        return batches

    def _subscribe(self, exch_by_token: dict):
        if not exch_by_token or self._sws is None:
            return
        for i, batch in enumerate(self._chunk(exch_by_token)):
            self._sws.subscribe(f"feed_{int(time.time() * 1000)}_{i}", SUBSCRIBE_MODE, batch)
            time.sleep(0.3)

    def register(self, exch_by_token: dict):
        """Add a client's tokens to the shared subscription. Safe to call
        repeatedly (e.g. every time a client's engine is (re)built) —
        tokens already subscribed are skipped, so this only ever grows
        the subscription and never triggers a fresh login/reconnect."""
        with self.lock:
            new_tokens = {t: e for t, e in exch_by_token.items() if t not in self._exch_by_token}
            self._exch_by_token.update(exch_by_token)
            if not new_tokens:
                return
            if not self._connected:
                self._pending.update(new_tokens)
                return
        self._subscribe(new_tokens)

    def snapshot_for(self, tokens):
        """Raw ticks for just `tokens` (one client's token set), plus the
        last-tick timestamp across the whole shared feed (a fine liveness
        indicator even though it may belong to another client's symbol)."""
        with self.lock:
            raw = {t: self.latest_raw[t] for t in tokens if t in self.latest_raw}
            return raw, self.last_tick_ts


@st.cache_resource(show_spinner=False)
def get_shared_feed() -> "SharedFeed":
    return SharedFeed()


class LiveEngine:
    def __init__(self, ledger_rows, apps_script_url=None, ledger_sheet_name=None):
        self.lock = threading.Lock()
        self.token_to_symbol = {}     # token -> meta (qty/avgPrice/segment/... for THIS client)
        self.closed_positions = []
        self.booked_mtm_total = 0.0
        self._own_status = "starting"  # this client's own ledger/token-resolve phase
        self._own_error = None
        self.apps_script_url = apps_script_url
        # Your Code.gs reads one tab per client (e.g. the "Rohan" tab for
        # that client's ledger) via ?sheet=<tab name>. Rollover data is
        # per-client too ("I will add rollover sheet as per client name"),
        # so the natural tab name is "<ledger tab> Rollover" — e.g. a
        # "Rohan Rollover" tab alongside the existing "Rohan" tab. Adjust
        # this one line if you end up naming it differently.
        self.rollover_sheet_name = f"{ledger_sheet_name} Rollover" if ledger_sheet_name else "Rollover"
        self.rollovers = {}           # underlying symbol -> list of rollover-chain dicts
        # Positions/token-resolution/login are all network calls and can take
        # several seconds. Do them in a background thread so get_engine()
        # (and therefore the page) returns to the browser immediately instead
        # of blocking behind Streamlit's "connecting" spinner. The UI polls
        # self.status via the live fragment until this flips to "live".
        threading.Thread(target=self._start, args=(ledger_rows,), daemon=True).start()

    @property
    def status(self):
        # While THIS client's own ledger/positions/token-resolve work is
        # still running (or failed), that's authoritative. Once it's
        # done, defer to the shared feed's connection state — that's the
        # actual websocket, shared by every client.
        if self._own_status in ("starting", "error"):
            return self._own_status
        return get_shared_feed().status

    @property
    def error(self):
        if self._own_status == "error":
            return self._own_error
        return get_shared_feed().error

    def _load_rollovers(self):
        """Best-effort fetch of this client's Rollover tab, if one exists.
        Unlike the ledger tab, this one is a human-readable report — a
        row with just the stock/index name, then a header row ('Date',
        'From...', ...), then that stock's data rows, then a blank row,
        repeated per stock (no single header row for the whole sheet, so
        the generic doGet's header-keyed JSON doesn't apply here). This
        fetches the tab in raw-grid mode (Code.gs's ?raw=1) and parses the
        blocks in _parse_rollover_grid.
        Most clients won't have this tab yet, so a 404 ("Sheet not
        found") or any other failure here is swallowed — this must never
        break the main ledger/positions load, it only enables the
        optional rollover-history expander once a client's tab has real
        data."""
        if not self.apps_script_url:
            return
        try:
            grid = _fetch_ledger_records(self.apps_script_url, self.rollover_sheet_name, raw=True)
        except Exception:
            return
        if not isinstance(grid, list):
            return
        try:
            by_stock = _parse_rollover_grid(grid)
        except Exception:
            return

        with self.lock:
            self.rollovers = by_stock

    def _start(self, ledger_rows):
        try:
            open_positions, closed_positions = build_positions(ledger_rows, verbose=False)
            self.closed_positions = closed_positions
            self.booked_mtm_total = round(sum(c["BookedPnL"] for c in closed_positions), 2)
            self._load_rollovers()

            resolved, unresolved = resolve_all(open_positions, MASTER_CACHE_PATH, MASTER_CACHE_MAX_AGE_HOURS)
            if not resolved:
                self._own_status = "error"
                self._own_error = "Nothing resolved from the ledger — check symbol/exchange/expiry spelling."
                return

            for r in resolved:
                # Match back to the original open_positions entry using
                # Exchange + Qty + AvgPrice rather than Symbol. Angel One's
                # instrument-master resolution returns the FULL contract
                # tradingsymbol for F&O (e.g. "NIFTY28APR2622500CE"), which
                # never equals the bare underlying symbol ("NIFTY") that
                # positions_builder.py stores — so a Symbol-based match
                # always missed for F&O and silently blanked out Expiry,
                # Segment, and PositionType for every F&O row. Qty/AvgPrice
                # are copied through resolve_all unchanged from the input
                # position, so they're a reliable join key; Symbol is kept
                # as a secondary check only for Equity, where it's still
                # accurate and helps disambiguate ties.
                def _matches(p, r=r):
                    same_exch = p["Exchange"] == r["exchange"]
                    same_qty = abs(float(p.get("Qty") or 0)) == abs(float(r.get("qty") or 0))
                    avg_p, avg_r = p.get("AvgPrice"), r.get("avgPrice")
                    same_avg = avg_p is not None and avg_r is not None and round(float(avg_p), 2) == round(float(avg_r), 2)
                    return same_exch and same_qty and same_avg

                src = next((p for p in open_positions if _matches(p)), {})
                if not src:
                    # Fall back to the old Symbol+Exchange match (covers
                    # Equity, where symbol formats do line up).
                    src = next((p for p in open_positions
                                if p["Symbol"] == r["symbol"] and p["Exchange"] == r["exchange"]), {})
                # If the Exchange+Qty+AvgPrice join missed (or missed the
                # type), look the ledger row up by bare symbol so its
                # BOND/ETF type still gets through.
                if not src or not src.get("InstrumentType"):
                    _base = re.sub(r"-(EQ|BE|BL|BZ|SM|N\d)$", "", str(r["symbol"]).upper())
                    _alt = next((p for p in open_positions
                                 if str(p.get("Symbol", "")).upper() == _base
                                 and p.get("InstrumentType")), None)
                    if _alt:
                        src = {**src, "InstrumentType": _alt["InstrumentType"],
                               "Segment": src.get("Segment") or _alt.get("Segment", "Other")}
                self.token_to_symbol[r["token"]] = {
                    "symbol": r["symbol"], "exchange": r["exchange"],
                    "qty": r.get("qty"), "avgPrice": r.get("avgPrice"),
                    "segment": effective_segment(src.get("Segment", "Other"),
                                                 src.get("InstrumentType", ""), r.get("symbol", "")),
                    "positionType": src.get("PositionType", "LONG"),
                    # Expiry may come back from the instrument-master lookup
                    # (F&O tokens resolve with an expiry) or from the ledger
                    # row itself, depending on which one has it.
                    "expiry": r.get("expiry") or src.get("Expiry") or src.get("ExpiryDate") or "",
                    # positions_builder.py keeps Strike/OptionType per contract
                    # (it's part of the FIFO contract key) — carry them through
                    # so the UI can show "BANKNIFTY 56000 CE" instead of just
                    # the bare underlying name. These were previously dropped
                    # here, which is why strike/CE-PE never reached the display.
                    "optionType": src.get("OptionType", "") or "",
                    "strike": src.get("Strike", "") or "",
                    "instrumentType": src.get("InstrumentType", "") or "",
                }


            # Register our tokens on the ONE shared Angel One
            # login/websocket instead of logging in ourselves — see
            # SharedFeed above for why per-client logins break things.
            # Safe/cheap even if the shared feed is already live: it only
            # subscribes whatever tokens aren't already covered.
            get_shared_feed().register({r["token"]: r["exchangeType"] for r in resolved})
            self._own_status = "live"
        except Exception as e:
            self._own_status = "error"
            self._own_error = str(e)

    def snapshot(self):
        """Merge this client's own qty/avgPrice/segment metadata with the
        raw ltp/open/close/volume ticks coming off the shared feed —
        keeps mtm calculations correct per-client even when two clients
        hold the same symbol at different qty/avg price."""
        token_meta = self.token_to_symbol
        raw, last_tick_ts = get_shared_feed().snapshot_for(token_meta.keys())
        ticks = []
        for token, meta in token_meta.items():
            r = raw.get(token)
            if not r:
                continue
            qty, avg_price, ltp = meta.get("qty"), meta.get("avgPrice"), r["ltp"]
            ticks.append({
                "token": token, "symbol": meta.get("symbol", token),
                "exchange": meta.get("exchange", ""), "segment": meta.get("segment", "Other"),
                "positionType": meta.get("positionType", "LONG"), "expiry": meta.get("expiry", ""), "ltp": ltp,
                "optionType": meta.get("optionType", ""), "strike": meta.get("strike", ""),
                "instrumentType": meta.get("instrumentType", ""),
                "prev_ltp": r.get("prev_ltp"),
                "close": r.get("close"), "open": r.get("open"), "volume": r.get("volume"),
                "qty": qty, "avgPrice": avg_price,
                "mtm": round((ltp - avg_price) * qty, 2) if (qty is not None and avg_price is not None) else None,
                "ts": r.get("ts"),
            })
        return ticks, last_tick_ts


@st.cache_resource(show_spinner="Fetching ledger from Google Sheet (Apps Script)...")
def get_engine(apps_script_url: str, sheet_name: str | None):
    records = _fetch_ledger_records(apps_script_url, sheet_name)
    from corporate_actions import apply_confirmed_actions
    records = apply_confirmed_actions(records)  # bake in confirmed Split/Bonus before FIFO
    rows = load_trade_ledger_from_records(records)
    return LiveEngine(rows, apps_script_url=apps_script_url, ledger_sheet_name=sheet_name)


# ── News ─────────────────────────────────────────────────────────────────
# No single free API covers NSE + BSE + Mint + Business Standard + Times of
# India + more with one key, so instead this queries Google News' public RSS
# search (no key/auth needed), restricted to those publishers. Google News
# indexes each outlet's full archive, not just today, so this also surfaces
# older ("past") coverage for a symbol, not only breaking news.
#
# IMPORTANT: this deliberately does ONE separate query PER source rather
# than one query with all sources OR'd together. A single query like
# "SYMBOL stock (site:a.com OR site:b.com OR ... 9 sites)" makes Google
# News quietly loosen/drop terms it can't satisfy well across that many
# OR'd site: filters — which is how unrelated pages (a generic BSE listing
# page, an F1 racing article, an "Option Chain" static page) were slipping
# into results. Splitting into 9 narrow per-source queries — each requiring
# the ticker literally in the headline via intitle: — avoids that entirely.
NEWS_SOURCES = {
    "NSE": "nseindia.com",
    "BSE": "bseindia.com",
    "Mint": "livemint.com",
    "Business Standard": "business-standard.com",
    "Times of India": "timesofindia.indiatimes.com",
    "Economic Times": "economictimes.indiatimes.com",
    "Moneycontrol": "moneycontrol.com",
    "CNBC-TV18": "cnbctv18.com",
    "Reuters": "reuters.com",
}

# Several outlets (Economic Times especially) auto-publish a templated
# "<Stock> Share Price Today, <Stock> Stock Price Live NSE/BSE Updates"
# page per stock, regenerated daily — not an actual news story. Filtered
# both in the query itself (-intitle: exclusions) and again client-side as
# a safety net, since Google doesn't reliably honor every negative filter.
NEWS_TITLE_BLOCKLIST = (
    "share price today", "stock price live", "share price live",
    "nse/bse updates", "share price nse", "stock price nse",
)

# Broad market-recap headlines. These only ever leak in for index symbols
# (NIFTY, BANKNIFTY, SENSEX, FINNIFTY, ...) since intitle:"<symbol>" happily
# matches "Nifty ends lower", "Sensex settles ... points" daily-wrap stories
# that aren't actually news about a specific holding — they're the general
# market-close roundup that runs every trading day regardless. Applied only
# when the symbol itself is an index, so it never touches real stock tickers.
INDEX_SYMBOLS = {"NIFTY", "BANKNIFTY", "SENSEX", "FINNIFTY", "NIFTY50", "NIFTY 50"}
MARKET_WRAP_BLOCKLIST = (
    "closing bell", "sensex settles", "sensex falls", "sensex ends",
    "sensex closes", "stock market today", "market today", "nifty ends",
    "nifty trades", "nifty settles", "nifty closes", "taking stock",
    "opening bell", "market wrap", "sensex, nifty", "nifty, sensex",
)
NEWS_MAX_AGE_DAYS = 7


def _parse_news_rss(xml_bytes):
    root = ET.fromstring(xml_bytes)
    items = []
    for item in root.findall(".//item"):
        title = (item.findtext("title") or "").strip()
        link = (item.findtext("link") or "").strip()
        pub_date = (item.findtext("pubDate") or "").strip()
        source_el = item.find("source")
        source = source_el.text.strip() if source_el is not None and source_el.text else ""
        # Google News titles are formatted "Headline - Source". Strip that
        # trailing " - Source" for a clean headline regardless of whether
        # <source> was present (normal case) or we have to fall back to
        # splitting the title itself (source tag missing).
        if source and title.endswith(f" - {source}"):
            title = title[: -(len(source) + 3)].strip()
        elif not source and " - " in title:
            title, source = title.rsplit(" - ", 1)
        try:
            # Google always sends this in GMT; strptime's %Z doesn't attach
            # real tzinfo, so dt comes back naive but is UTC-equivalent —
            # fine as long as we compare it against another naive UTC value
            # (see the utcnow() cutoff below), never against IST/aware time.
            dt = datetime.strptime(pub_date, "%a, %d %b %Y %H:%M:%S %Z")
        except ValueError:
            dt = None
        items.append({"title": title, "source": source, "link": link, "dt": dt, "pub_date_raw": pub_date})
    return items


def _fetch_one_source(symbol, domain, max_items):
    # intitle: forces the ticker to literally appear in the headline —
    # this is what actually keeps results on-topic. The -intitle: terms
    # knock out the auto-generated "Share Price Today" template pages at
    # the source, before they even count against max_items.
    exclusions = " ".join(f'-intitle:"{p}"' for p in ("Share Price Today", "Stock Price Live"))
    query = f'intitle:"{symbol}" {exclusions} site:{domain}'
    url = f"https://news.google.com/rss/search?q={quote_plus(query)}&hl=en-IN&gl=IN&ceid=IN:en"
    resp = requests.get(url, timeout=10, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return _parse_news_rss(resp.content)[:max_items]


@st.cache_data(ttl=1800, show_spinner=False)
def fetch_stock_news(symbol: str, per_source_max: int = 8, total_max: int = 25):
    """Latest coverage for one symbol from the past NEWS_MAX_AGE_DAYS days
    only, queried separately per source in NEWS_SOURCES (see module note
    above for why), merged and sorted newest-first. Cached 30 min per
    symbol so switching tabs or the live-tick refresh loop doesn't re-hit
    Google News on every rerun. Returns a list of
    {title, source, link, published, recency} dicts — recency is one of
    "new" (<24h old), "recent" (1-3 days), or "week" (3-7 days), computed
    at fetch time; since the cache TTL is 30 min, a bucket can drift stale
    by at most that much, which isn't visible at day-scale bucket sizes.
    """
    merged = []
    with ThreadPoolExecutor(max_workers=len(NEWS_SOURCES)) as pool:
        futures = {
            pool.submit(_fetch_one_source, symbol, domain, per_source_max): name
            for name, domain in NEWS_SOURCES.items()
        }
        for fut in as_completed(futures):
            try:
                merged.extend(fut.result())
            except requests.RequestException:
                continue  # one source failing shouldn't sink the whole fetch

    now = datetime.utcnow()
    cutoff = now - timedelta(days=NEWS_MAX_AGE_DAYS)
    is_index = symbol.strip().upper() in INDEX_SYMBOLS
    filtered = []
    seen_keys = set()  # (normalized title, link) — Google News RSS can hand
                        # back the same story twice for one source query
                        # (syndication/pagination overlap); drop the repeat.
    for it in merged:
        # Strict 7-day window: an item we can't date can't be verified as
        # within it, so — unlike a "best effort" feed — it gets dropped
        # rather than kept.
        if it["dt"] is None or it["dt"] < cutoff:
            continue
        low = it["title"].lower()
        if any(p in low for p in NEWS_TITLE_BLOCKLIST):
            continue
        if is_index and any(p in low for p in MARKET_WRAP_BLOCKLIST):
            continue
        dedup_key = (re.sub(r"\s+", " ", low).strip(), it["link"])
        if dedup_key in seen_keys:
            continue
        seen_keys.add(dedup_key)
        filtered.append(it)

    filtered.sort(key=lambda it: it["dt"], reverse=True)

    out = []
    for it in filtered[:total_max]:
        age = now - it["dt"]
        if age <= timedelta(hours=24):
            recency = "new"
        elif age <= timedelta(days=3):
            recency = "recent"
        else:
            recency = "week"
        out.append({
            "title": it["title"], "source": it["source"], "link": it["link"],
            "published": it["dt"].strftime("%d %b %Y, %I:%M %p"), "recency": recency,
            "dt": it["dt"],  # kept (not just the formatted string) so the "All
                             # stocks" view can merge+re-sort across symbols
        })
    return out


ALL_STOCKS_OPTION = "🌐 All my stocks"


def fetch_all_stock_news(symbols: list[str], per_symbol_max: int = 6, total_max: int = 60):
    """Merges fetch_stock_news across every symbol into one recency-sorted
    feed, each item tagged with which stock it's about. Bounded to
    per_symbol_max items per stock before merging, so one heavily-covered
    stock can't crowd out everything else. Runs the per-symbol fetches
    concurrently (each of which is itself already cached 30 min and
    internally parallel across sources) capped to a modest worker count —
    fetch_stock_news's own internal pool already opens up to 9 connections
    per symbol, so fanning out too many symbols at once would multiply
    that unnecessarily.
    """
    merged = []
    with ThreadPoolExecutor(max_workers=min(6, max(1, len(symbols)))) as pool:
        futures = {pool.submit(fetch_stock_news, sym): sym for sym in symbols}
        for fut in as_completed(futures):
            sym = futures[fut]
            try:
                items = fut.result()
            except Exception:
                continue  # one stock failing shouldn't sink the whole feed
            for it in items[:per_symbol_max]:
                merged.append({**it, "symbol": sym})

    merged.sort(key=lambda it: it["dt"], reverse=True)
    return merged[:total_max]


def render_news_section(symbols: list[str], *, scope_key: str, scope_noun: str):
    """Renders one news feed (selectbox + headline list) scoped to `symbols`.

    scope_key   — unique suffix for widget keys, so the Open-positions and
                  Closed-positions news pickers (rendered in separate tabs)
                  don't collide in Streamlit's session state.
    scope_noun  — used in empty-state / caption copy, e.g. "open positions".
    """
    if not symbols:
        st.caption(f"No {scope_noun} to show news for yet.")
        return
    options = [ALL_STOCKS_OPTION] + symbols
    picked = st.selectbox("Stock", options=options, key=f"_news_symbol_{scope_key}")
    if not picked:
        return

    show_all = picked == ALL_STOCKS_OPTION
    with st.spinner(f"Fetching news for {'all your ' + scope_noun if show_all else picked}..."):
        try:
            items = fetch_all_stock_news(symbols) if show_all else fetch_stock_news(picked)
        except Exception as exc:
            st.error(f"Couldn't fetch news: {exc}")
            return

    if not items:
        scope = f"your {scope_noun}" if show_all else f"\"{picked}\""
        st.caption(f"No headlines with {scope} in the title from the last "
                   f"{NEWS_MAX_AGE_DAYS} days across the tracked sources.")
        return

    RECENCY_BADGE = {
        "new": '<span class="news-badge news-badge-new">🟢 New</span>',
        "recent": '<span class="news-badge news-badge-recent">🕒 Recent</span>',
        "week": "",
    }
    for it in items:
        meta = it["source"] + (" · " + it["published"] if it["published"] else "")
        badge = RECENCY_BADGE.get(it["recency"], "")
        stock_badge = f'<span class="news-badge news-badge-stock">{it["symbol"]}</span>' if show_all else ""
        st.markdown(
            flat(f"""
            <div class="news-item news-{it['recency']}">
                <a href="{it['link']}" target="_blank" rel="noopener noreferrer" class="news-title">{it['title']}</a>{stock_badge}{badge}
                <div class="news-meta">{meta}</div>
            </div>
            """),
            unsafe_allow_html=True,
        )

    scope_text = f"across all {len(symbols)} of your {scope_noun}" if show_all else f"with \"{picked}\" in the title"
    st.caption(f"Headlines {scope_text}, past {NEWS_MAX_AGE_DAYS} days, across NSE, BSE, "
               "Mint, Business Standard, Times of India, Economic Times, Moneycontrol, CNBC-TV18 and "
               "Reuters — excluding auto-generated price-update pages. Cached 30 min per stock.")



def fmt_money(x):
    if x is None:
        return "-"
    return f"(₹{abs(x):,.2f})" if x < 0 else f"₹{x:,.2f}"


def fmt_qty(qty):
    if qty is None:
        return "-"
    return f"{abs(round(qty)):,}"


def fmt_time(ts):
    """ts is an ISO datetime string (possibly tz-aware); show HH:MM:SS only."""
    if not ts:
        return "-"
    try:
        dt = datetime.fromisoformat(ts)
        return dt.strftime("%H:%M:%S")
    except ValueError:
        return ts


def fmt_datetime(ts):
    """ts is an ISO datetime string (possibly tz-aware); show date + time."""
    if not ts:
        return "-"
    try:
        dt = datetime.fromisoformat(ts)
        return dt.strftime("%d %b, %H:%M:%S")
    except ValueError:
        return ts


def fmt_sell_date(d):
    """d may be a date/datetime object or an ISO-ish string; show a short date."""
    if not d:
        return "-"
    if isinstance(d, str):
        try:
            d = datetime.fromisoformat(d)
        except ValueError:
            return d
    try:
        return d.strftime("%d %b %Y")
    except AttributeError:
        return str(d)


def fmt_expiry(e):
    """e may be a date/datetime object or a string like '25AUG2026'/'2026-08-25'/
    '2026-08-25 00:00:00' (the last of these is what positions_builder's
    _norm() produces for CSV-sourced — i.e. Google Sheets — ledger rows,
    since it str()s a full Timestamp rather than a bare date)."""
    if not e:
        return "-"
    if isinstance(e, str):
        for pattern in ("%d%b%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return datetime.strptime(e, pattern).strftime("%d %b %Y")
            except ValueError:
                continue
        # Last resort: let pandas' looser parser have a go before giving up.
        parsed = pd.to_datetime(e, errors="coerce")
        if pd.notna(parsed):
            return parsed.strftime("%d %b %Y")
        return e  # unrecognized format — show as-is rather than hide it
    try:
        return e.strftime("%d %b %Y")
    except AttributeError:
        return str(e)


_FNO_SUFFIX_RE = re.compile(r"^([A-Z&\-]+?)\d{1,2}[A-Z]{3,9}\d{2,4}(?:\d+(?:CE|PE)|FUT)$")


def underlying_symbol(symbol: str) -> str:
    """Strip the expiry/strike/option-type suffix off an F&O trading symbol
    to get the underlying stock/index name — e.g. 'HDFCBANK30JUN26FUT' ->
    'HDFCBANK', 'NIFTY28APR2622500CE' -> 'NIFTY'. Used to group a stock's
    F&O contracts (including ones rolled from one expiry series to the
    next) together the way the client-facing Excel report does. Equity
    symbols have no such suffix and are returned unchanged."""
    if not symbol:
        return symbol
    m = _FNO_SUFFIX_RE.match(symbol.strip().upper())
    return m.group(1) if m else symbol


def _parse_num(val):
    """Best-effort float parse for values coming out of a Google Sheet cell
    (may already be a number, or a comma-formatted string like '2,040')."""
    if val is None or val == "":
        return None
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def contract_display(r: dict) -> str:
    """Human label for an F&O contract: underlying + strike + CE/PE (e.g.
    'BANKNIFTY 56000 CE'), or underlying + FUT — instead of just the bare
    underlying name. positions_builder.py keeps OptionType/Strike/
    InstrumentType as part of the FIFO contract key, and LiveEngine now
    carries them through (optionType/strike/instrumentType) onto every
    resolved position, so those are the reliable source; a handful of
    alternate field-name fallbacks are kept in case a row lands here
    without going through that same path.
    """
    symbol = str(r.get("symbol") or "-").strip()
    # Already a full Angel One-style trading symbol (e.g. "BANKNIFTY25SEP2666000CE")?
    if _FNO_SUFFIX_RE.match(symbol.upper()):
        return symbol

    stock = underlying_symbol(symbol) or symbol
    strike = (
        r.get("strike") or r.get("Strike") or r.get("strikePrice")
        or r.get("StrikePrice") or r.get("strike_price")
    )
    opt_raw = (
        r.get("optionType") or r.get("optiontype") or r.get("OptionType")
        or r.get("Option Type") or r.get("right") or r.get("Right")
        or r.get("CE/PE") or r.get("ce_pe")
    )
    opt = str(opt_raw).strip().upper() if opt_raw else ""
    instrument_type = str(r.get("instrumentType") or r.get("InstrumentType") or "").strip().upper()

    strike_num = _parse_num(strike)
    if opt in ("CE", "PE") and strike_num is not None and strike_num > 0:
        return f"{stock} {fmt_qty(strike_num)} {opt}"
    if instrument_type.startswith("FUT") or opt in ("FUT", "FUTURE", "FUTURES") or "FUT" in symbol.upper():
        return f"{stock} FUT"
    return stock


def roll_chain_card_html(row: dict) -> str:
    """One rollover-chain leg (Date, From Series → To Series, Qty rolled,
    last-month price, roll sell/buy price, roll diff, carry-forward price)
    rendered as a compact stat card instead of a row in a wide st.dataframe
    — the dataframe's ~7 columns forced horizontal scrolling on mobile;
    this reflows into 2-3 columns per screen width with nothing to swipe."""
    date_txt = fmt_expiry(row.get("Date")) if row.get("Date") else "-"
    from_s = row.get("From Series") or "-"
    to_s = row.get("To Series") or "-"

    def _money(key):
        v = _parse_num(row.get(key))
        return fmt_money(v) if v is not None else "-"

    qty_out = _parse_num(row.get("Qty Out"))
    stats = [
        ("Qty Rolled", fmt_qty(qty_out) if qty_out is not None else "-"),
        ("Last Month Price", _money("Last Month Price")),
        ("Roll Sell Price", _money("Roll Sell Price")),
        ("Roll Buy Price", _money("Roll Buy Price")),
        ("Roll Diff", _money("Roll Diff")),
        ("Carry-Fwd Price", _money("Carry-Forward Price")),
    ]
    stats_html = "".join(
        f'<div><div class="roll-chain-stat-label">{lbl}</div>'
        f'<div class="roll-chain-stat-value">{val}</div></div>'
        for lbl, val in stats
    )
    return flat(f"""
        <div class="roll-chain-card">
            <div class="roll-chain-head">
                <span>{date_txt}</span>
                <span class="roll-chain-arrow">{from_s} → {to_s}</span>
            </div>
            <div class="roll-chain-grid">{stats_html}</div>
        </div>
    """)


def rollover_badge_html(symbol: str, rollovers: dict) -> str:
    """A small '🔄 Rolled' pill for a symbol's row (open positions, live
    table) when its underlying stock has rollover history — so anyone
    scanning the table can see at a glance that a position was carried
    forward rather than freshly opened, without having to go check the
    Closed Positions tab. Hovering shows the most recent roll (series +
    roll diff) as a tooltip. Empty string when there's no rollover data
    for this symbol's underlying (the normal case until a client's
    Rollover sheet has entries)."""
    chain = rollovers.get(underlying_symbol(symbol or "-"))
    if not chain:
        return ""
    last = chain[-1]
    detail = ""
    if last.get("From Series") and last.get("To Series"):
        detail = f"{last['From Series']} → {last['To Series']}"
    elif last.get("Date"):
        detail = str(last["Date"])
    if last.get("Roll Diff") not in (None, ""):
        detail = f"{detail} · diff {last['Roll Diff']}" if detail else f"diff {last['Roll Diff']}"
    tooltip = f"Rolled {len(chain)}x — last: {detail}" if detail else f"Rolled {len(chain)}x"
    return f'<span class="pt-tag rolled" title="{_html_escape(tooltip, quote=True)}">🔄 Rolled</span>'


def alt_dark(chart):
    """Apply a shared dark, transparent-background theme to an Altair chart."""
    return (
        chart.configure_view(strokeWidth=0)
        .configure_axis(
            gridColor="#1e2733", domainColor="#1e2733", tickColor="#1e2733",
            labelColor="#7f8ba3", titleColor="#7f8ba3", labelFontSize=10.5, titleFontSize=11,
        )
        .configure_legend(labelColor="#7f8ba3", titleColor="#7f8ba3")
        .properties(background="transparent")
    )


def style_pnl_table(df, cols):
    """Return a pandas Styler that colors P&L-type columns green/red."""
    def _color(v):
        if pd.isna(v):
            return ""
        return "color: #2ee6a6; font-weight:600;" if v >= 0 else "color: #ff5c7a; font-weight:600;"

    fmt = {c: "₹{:,.2f}".format for c in cols if c in df.columns}
    styler = df.style.format(fmt)
    # pandas >=2.1 renamed Styler.applymap -> Styler.map (and removed
    # applymap entirely in pandas 3.x), so pick whichever exists at runtime.
    color_fn = styler.map if hasattr(styler, "map") else styler.applymap
    for c in cols:
        if c in df.columns:
            styler = color_fn(_color, subset=[c])
            color_fn = styler.map if hasattr(styler, "map") else styler.applymap
    return styler


# ── Live fragment: KPI cards + open/closed tables, refreshes on its own ───
@st.fragment(run_every=TICK_REFRESH_SECONDS)
def render_live(engine: "LiveEngine"):
    if engine.status == "error":
        st.markdown(status_pill("error", "Feed error"), unsafe_allow_html=True)
        st.error(f"Feed failed to start: {engine.error}")
        return
    elif engine.status == "starting":
        st.markdown(status_pill("warn", "Starting up..."), unsafe_allow_html=True)
        st.caption("Resolving instrument tokens and logging into Angel One — this can take a few seconds on a cold start.")
        return

    ticks, last_tick_ts = engine.snapshot()

    now_ist = datetime.now(IST)
    if engine.status == "disconnected":
        status_html = status_pill("error", "Disconnected — restart feed in sidebar")
    else:
        age = f"{(now_ist - last_tick_ts).seconds}s ago" if last_tick_ts else "waiting for first tick..."
        status_html = status_pill("live", f"Live • last tick {age}")

    st.markdown(
        flat(f"""
        <div class="top-bar">
            {status_html}
            <div class="top-clock">
                <span class="date-part"><span class="weekday-part">{now_ist.strftime("%A")}, </span>{now_ist.strftime("%d %b %Y")}</span>
                <span>{now_ist.strftime("%H:%M:%S")}</span>
                <span class="tz-badge">IST</span>
            </div>
        </div>
        """),
        unsafe_allow_html=True,
    )

    for t in ticks:
        t["segment"] = effective_segment(t.get("segment"), t.get("instrumentType", ""), t.get("symbol", ""))
    equity = [t for t in ticks if t.get("segment") == "Equity"]
    bonds_etf = [t for t in ticks if t.get("segment") not in ("Equity", "F&O")]
    fo_all = [t for t in ticks if t.get("segment") == "F&O"]
    # Options carry a CE/PE optionType (now passed through from
    # positions_builder.py via LiveEngine); anything F&O without CE/PE
    # (plain FUT, or blank) is a future.
    futures = [t for t in fo_all if str(t.get("optionType") or "").strip().upper() not in ("CE", "PE")]
    options = [t for t in fo_all if str(t.get("optionType") or "").strip().upper() in ("CE", "PE")]

    def seg_totals(rows):
        # Capital deployed must use abs(qty): a short position has a
        # negative qty, and summing signed qty*avgPrice was silently
        # *subtracting* short positions' cost basis from the total instead
        # of adding it — that was the "wrong Investment Value" bug.
        buy_value = sum(abs(r["qty"] or 0) * (r["avgPrice"] or 0) for r in rows)
        mtm = sum(r["mtm"] for r in rows if r["mtm"] is not None)
        # Today's move only (ltp vs previous close), signed qty on purpose:
        # for a short, a falling price is a gain, and (ltp-close) is
        # negative while qty is negative too, so the product comes out
        # positive automatically.
        day_pnl = sum(
            (r["ltp"] - r["close"]) * r["qty"]
            for r in rows if r.get("close") not in (None, 0) and r.get("qty") is not None
        )
        return buy_value, mtm, day_pnl

    eq_buy, eq_mtm, eq_day = seg_totals(equity)
    bonds_buy, bonds_mtm, bonds_day = seg_totals(bonds_etf)
    fut_buy, fut_mtm, fut_day = seg_totals(futures)
    opt_buy, opt_mtm, opt_day = seg_totals(options)
    seg_open_totals = {
        "Equity": (eq_buy, eq_mtm, eq_day),
        "Bonds/ETF": (bonds_buy, bonds_mtm, bonds_day),
        "Futures": (fut_buy, fut_mtm, fut_day),
        "Options": (opt_buy, opt_mtm, opt_day),
    }
    # Headline totals must cover EVERY open position — no segment silently
    # excluded. "Open MTM" is labeled "current MTM — booked plus open", so
    # it has to include Bonds/ETF too, not just Equity/Futures/Options.
    investment_value = eq_buy + bonds_buy + fut_buy + opt_buy
    current_mtm = eq_mtm + bonds_mtm + fut_mtm + opt_mtm
    total_mtm = engine.booked_mtm_total + current_mtm
    day_pnl_total = eq_day + bonds_day + fut_day + opt_day
    day_pnl_pct = (day_pnl_total / investment_value * 100) if investment_value else 0.0

    st.markdown(
        hero_block_dual(
            "Current MTM (Booked + Open)",
            fmt_money(total_mtm),
            "includes booked profit from closed positions",
            total_mtm >= 0,
            "Running MTM (Open Only)",
            fmt_money(current_mtm),
            "live open positions, excludes booked profit",
            current_mtm >= 0,
            pill_text="Profit" if total_mtm >= 0 else "Loss",
        ),
        unsafe_allow_html=True,
    )
    st.markdown(
        hero_tiles([
            ("Investment Value", fmt_money(investment_value), ""),
            (
                "Today's P&L",
                f"{fmt_money(day_pnl_total)} ({'+' if day_pnl_pct >= 0 else ''}{day_pnl_pct:.2f}%)",
                "positive" if day_pnl_total >= 0 else "negative",
            ),
            ("Booked MTM", fmt_money(engine.booked_mtm_total), "positive" if engine.booked_mtm_total >= 0 else "negative"),
            ("Open MTM", fmt_money(current_mtm), "positive" if current_mtm >= 0 else "negative"),
            ("Current MTM", fmt_money(total_mtm), "positive" if total_mtm >= 0 else "negative"),
        ]),
        unsafe_allow_html=True,
    )

    tab_open, tab_closed, tab_news, tab_corp = st.tabs(
        ["📈 Open positions", "✅ Closed positions", "📰 News", "🏢 Corporate Actions"]
    )

    with tab_open:
        open_segments = [
            ("Equity", equity),
            ("Bonds/ETF", bonds_etf),
            ("Futures", futures),
            ("Options", options),
        ]
        # One sub-tab per segment so picking "F&O" shows only F&O, not every
        # segment stacked one after another.
        open_seg_tabs = st.tabs([f"{label} ({len(rows)})" for label, rows in open_segments])
        for (label, rows), seg_tab in zip(open_segments, open_seg_tabs):
          with seg_tab:
            count_badge = f'<span class="badge">{len(rows)}</span>'
            st.markdown(f'<div class="section-label">{label} {count_badge}</div>', unsafe_allow_html=True)

            # One-line summary scoped to THIS segment only — switching to the
            # F&O tab shows F&O's own investment/running-P&L/today's-P&L,
            # not the combined book total.
            seg_buy, seg_mtm, seg_day = seg_open_totals.get(label, (0.0, 0.0, 0.0))
            seg_mtm_pct = (seg_mtm / seg_buy * 100) if seg_buy else 0.0
            seg_day_pct = (seg_day / seg_buy * 100) if seg_buy else 0.0
            st.markdown(
                seg_summary_line([
                    (f"{label} Investment", fmt_money(seg_buy), ""),
                    (
                        "Running P&L",
                        f"{fmt_money(seg_mtm)} ({'+' if seg_mtm >= 0 else ''}{seg_mtm_pct:.2f}%)",
                        "kpi-pos" if seg_mtm >= 0 else "kpi-neg",
                    ),
                    (
                        "Today's P&L",
                        f"{fmt_money(seg_day)} ({'+' if seg_day >= 0 else ''}{seg_day_pct:.2f}%)",
                        "kpi-pos" if seg_day >= 0 else "kpi-neg",
                    ),
                ]),
                unsafe_allow_html=True,
            )

            if not rows:
                st.caption("No open positions in this segment.")
                continue

            def _daily_gain_pct(x):
                # Price-move % (ltp vs prev close), then flipped for shorts:
                # a short position LOSES when the price rises, so its
                # "daily gain" is the mirror image of the raw price move.
                close = x.get("close")
                if not close:
                    return None
                raw_pct = (x["ltp"] - close) / close * 100
                pos_type = (x.get("positionType") or "LONG").upper()
                return raw_pct if pos_type == "LONG" else -raw_pct

            # Futures and Options both care about expiry more than exchange
            # (always NFO/BFO); Equity/Bonds-ETF show exchange as before.
            # Options gets its own wider row template with explicit
            # Strike/Type/Expiry columns instead of folding them into one
            # "second column".
            is_fo = label in ("Futures", "Options")
            is_options = label == "Options"
            # Options don't roll in this workflow — only Futures do — so the
            # "Rolled" badge and the rollover expander are scoped to Futures
            # only, even though Options still uses the F&O-style symbol/
            # expiry labeling above.
            show_rollover = label == "Futures"
            second_col_label = "Expiry" if label == "Futures" else "Exchange"

            # Highest daily gain first, always — never a fixed/pinned order.
            # Positions with no price yet (None) sort to the bottom.
            rows_html = []
            sorted_rows = sorted(
                rows,
                key=lambda x: (_daily_gain_pct(x) is None, -(_daily_gain_pct(x) or 0)),
            )
            for idx, r in enumerate(sorted_rows):
                pos_type = (r.get("positionType") or "LONG").upper()
                type_cls = "long" if pos_type == "LONG" else "short"
                close = r.get("close")
                day_pct = _daily_gain_pct(r)  # this position's actual gain/loss %, short-adjusted
                day_cls = "pt-cell pos" if (day_pct or 0) >= 0 else "pt-cell neg"
                day_txt = f"{'+' if (day_pct or 0) >= 0 else ''}{day_pct:.2f}%" if day_pct is not None else "–"
                # The single best-performing row in this segment always gets
                # the glow — if it's a genuine gain it reads "Top Gain"; if
                # every position in the segment is red today, the least-bad
                # one still gets marked so there's always a clear leader.
                is_top = idx == 0 and day_pct is not None
                row_variant = "cols-open-opt" if is_options else "cols-open"
                row_cls = f"pos-table-row {row_variant} top-gain-row" if is_top else f"pos-table-row {row_variant}"
                if is_top and day_pct > 0:
                    leader_badge = f'<span class="pt-leader-badge">🔥 Top Gain</span>'
                elif is_top:
                    leader_badge = f'<span class="pt-leader-badge">🛡️ Least Loss</span>'
                else:
                    leader_badge = ""
                mtm = r.get("mtm")
                mtm_cls = "pt-cell pos" if (mtm or 0) >= 0 else "pt-cell neg"
                mtm_arrow = "▲" if (mtm or 0) >= 0 else "▼"
                roll_badge = rollover_badge_html(r.get("symbol", ""), engine.rollovers) if show_rollover else ""
                symbol_label = _html_escape(contract_display(r)) if is_fo else _html_escape(str(r.get("symbol", "-")))

                if is_options:
                    strike_num = _parse_num(r.get("strike"))
                    strike_txt = fmt_qty(strike_num) if strike_num else "-"
                    opt_txt = str(r.get("optionType") or "-").strip().upper()
                    expiry_txt = fmt_expiry(r.get("expiry"))
                    rows_html.append(flat(f"""
                        <div class="{row_cls}">
                            <div class="pt-symbol">
                                <span class="pt-symbol-name">{symbol_label}</span>
                                <span class="pt-tag {type_cls}">{pos_type}</span>
                                {roll_badge}
                                {leader_badge}
                            </div>
                            <div class="pt-cell muted">{strike_txt}</div>
                            <div class="pt-cell muted">{opt_txt}</div>
                            <div class="pt-cell muted">{expiry_txt}</div>
                            <div class="pt-cell">{fmt_qty(r.get('qty'))}</div>
                            <div class="pt-cell">{fmt_money(r.get('avgPrice'))}</div>
                            <div class="pt-cell">{fmt_money(r.get('ltp'))}</div>
                            <div class="{day_cls}">{day_txt}</div>
                            <div class="{mtm_cls}">{mtm_arrow} {fmt_money(mtm)}</div>
                            <div class="pt-cell muted">{fmt_datetime(r.get('ts'))}</div>
                        </div>
                    """))
                else:
                    second_col_value = fmt_expiry(r.get("expiry")) if label == "Futures" else r.get("exchange", "-")
                    rows_html.append(flat(f"""
                        <div class="{row_cls}">
                            <div class="pt-symbol">
                                <span class="pt-symbol-name">{symbol_label}</span>
                                <span class="pt-tag {type_cls}">{pos_type}</span>
                                {roll_badge}
                                {leader_badge}
                            </div>
                            <div class="pt-cell muted">{second_col_value}</div>
                            <div class="pt-cell">{fmt_qty(r.get('qty'))}</div>
                            <div class="pt-cell">{fmt_money(r.get('avgPrice'))}</div>
                            <div class="pt-cell">{fmt_money(r.get('ltp'))}</div>
                            <div class="{day_cls}">{day_txt}</div>
                            <div class="{mtm_cls}">{mtm_arrow} {fmt_money(mtm)}</div>
                            <div class="pt-cell muted">{fmt_datetime(r.get('ts'))}</div>
                        </div>
                    """))
            if is_options:
                table_html = flat(f"""
                    <div class="pos-table-wrap">
                        <div class="pos-table">
                            <div class="pos-table-head cols-open-opt">
                                <div>Symbol</div><div>Strike</div><div>CE/PE</div><div>Expiry</div>
                                <div>Qty</div><div>Avg Price</div><div>CMP</div><div>Day Chg %</div>
                                <div>MTM P&amp;L</div><div>Last Tick</div>
                            </div>
                            {"".join(rows_html)}
                        </div>
                    </div>
                """)
            else:
                table_html = flat(f"""
                    <div class="pos-table-wrap">
                        <div class="pos-table">
                            <div class="pos-table-head cols-open">
                                <div>Symbol</div><div>{second_col_label}</div><div>Qty</div>
                                <div>Avg Price</div><div>CMP</div><div>Day Chg %</div>
                                <div>MTM P&amp;L</div><div>Last Tick</div>
                            </div>
                            {"".join(rows_html)}
                        </div>
                    </div>
                """)
            st.markdown(table_html, unsafe_allow_html=True)

            # Rollover positions — a real st.expander (not raw HTML) so it
            # stays open across this fragment's 1s reruns, unlike the inline
            # per-row badge which is just a quick visual flag. Shows every
            # currently-open F&O position that's been rolled, each with its
            # live MTM plus the roll chain (from/to series, roll diff, ...)
            # that got it there.
            if show_rollover:
                rolled_rows = [r for r in rows if engine.rollovers.get(underlying_symbol(r.get("symbol", "")))]
                if rolled_rows:
                    # Group by underlying stock first — engine.rollovers holds
                    # ALL roll-chain rows for an underlying (e.g. "BANKNIFTY")
                    # merged into one bucket, because the sheet's block header
                    # is just the bare stock name for every contract (no
                    # strike/CE-PE/qty in the header itself to tell separate
                    # contracts apart). So within each stock, we further match
                    # each position to its OWN chain rows by "Qty Out" — the
                    # one field that reliably lines up with a specific
                    # contract's current open quantity (confirmed against the
                    # client's ledger: qty 660/1200/1560/1560/2040 each map to
                    # a distinct strike+side). If nothing matches by qty, we
                    # fall back to showing the full merged chain rather than
                    # hiding data.
                    by_stock = {}
                    for r in rolled_rows:
                        stock = underlying_symbol(r.get("symbol", ""))
                        by_stock.setdefault(stock, []).append(r)

                    with st.expander(f"🔄 Rollover positions ({len(rolled_rows)})"):
                        for stock in sorted(by_stock.keys()):
                            stock_rows = sorted(by_stock[stock], key=lambda x: (_parse_num(x.get("qty")) or 0))
                            full_chain = engine.rollovers.get(stock, [])
                            st.markdown(f"**{stock}**")
                            for r in stock_rows:
                                r_qty = _parse_num(r.get("qty"))
                                matched_chain = [
                                    row for row in full_chain
                                    if r_qty is not None and _parse_num(row.get("Qty Out")) == r_qty
                                ] or full_chain
                                contract_label = contract_display(r)
                                mtm_txt = fmt_money(r.get("mtm"))
                                st.markdown(
                                    f"**{contract_label}** — Qty {fmt_qty(r.get('qty'))} · "
                                    f"Avg {fmt_money(r.get('avgPrice'))} · CMP {fmt_money(r.get('ltp'))} · "
                                    f"MTM {mtm_txt}"
                                )
                                st.markdown(
                                    "".join(roll_chain_card_html(row) for row in matched_chain),
                                    unsafe_allow_html=True,
                                )
                            st.divider()

    with tab_closed:
        if not engine.closed_positions:
            st.caption("No closed positions in this ledger.")
        else:
            closed = engine.closed_positions
            wins = sum(1 for c in closed if c["BookedPnL"] >= 0)
            losses = len(closed) - wins
            win_rate = wins / len(closed) * 100 if closed else 0

            # All realized FIFO trade legs across every closed position —
            # used for the leg count, best/worst single trade, and the
            # average P&L per realized trade below.
            all_legs = [
                (t.get("Pnl", 0.0), p.get("Symbol", "-"))
                for p in closed for t in p.get("Trades", [])
            ]
            realized_trades = len(all_legs)
            if all_legs:
                best_pnl, best_symbol = max(all_legs, key=lambda x: x[0])
                worst_pnl, worst_symbol = min(all_legs, key=lambda x: x[0])
            else:
                best_pnl = worst_pnl = 0.0
                best_symbol = worst_symbol = "-"
            avg_pnl_trade = (engine.booked_mtm_total / realized_trades) if realized_trades else 0.0

            st.markdown(
                hero_block(
                    "Booked Ledger — from trade ledger FIFO",
                    fmt_money(engine.booked_mtm_total),
                    "total booked profit",
                    is_positive=engine.booked_mtm_total >= 0,
                    pill_text="Booked",
                ),
                unsafe_allow_html=True,
            )
            st.markdown(
                hero_tiles([
                    ("Closed Positions", str(len(closed)), ""),
                    ("Realized Trades (FIFO Legs)", str(realized_trades), ""),
                    ("Win Rate", f"{win_rate:.0f}%", ""),
                    ("Best Trade", f"{best_symbol} · {fmt_money(best_pnl)}", "positive" if best_pnl >= 0 else "negative"),
                    ("Worst Trade", f"{worst_symbol} · {fmt_money(worst_pnl)}", "positive" if worst_pnl >= 0 else "negative"),
                    ("Avg P&L / Trade", fmt_money(avg_pnl_trade), "positive" if avg_pnl_trade >= 0 else "negative"),
                    ("Current Open MTM", fmt_money(current_mtm), "positive" if current_mtm >= 0 else "negative"),
                ]),
                unsafe_allow_html=True,
            )

            def _closed_sell_date(c):
                # Trades[].SellDate comes out of positions_builder.py's
                # _fmt_date() as an ISO string ("2026-09-29"), or is absent
                # entirely for a leg with no parseable sell date. Returning
                # that string directly here meant some rows in a stock group
                # keyed on a str (max() of the string dates) while OTHER rows
                # (no sell date at all) fell back to `datetime.min` at the
                # call site — mixing str and datetime keys in the same
                # sorted() call, which raises TypeError. Parsing to real
                # datetime objects here (or None) keeps every key the same
                # comparable type.
                dates = []
                for t in c.get("Trades", []):
                    raw = t.get("SellDate")
                    if not raw:
                        continue
                    parsed = pd.to_datetime(raw, errors="coerce")
                    if pd.notna(parsed):
                        dates.append(parsed.to_pydatetime())
                return max(dates) if dates else None

            def closed_row_html(c, is_top=False):
                pnl = c["BookedPnL"]
                pnl_cls = "pt-cell pos" if pnl >= 0 else "pt-cell neg"
                pnl_arrow = "▲" if pnl >= 0 else "▼"
                row_cls = "pos-table-row cols-closed top-gain-row" if is_top else "pos-table-row cols-closed"
                if is_top and pnl > 0:
                    leader_badge = '<span class="pt-leader-badge">🔥 Top Gain</span>'
                elif is_top:
                    leader_badge = '<span class="pt-leader-badge">🛡️ Least Loss</span>'
                else:
                    leader_badge = ""
                sell_date = fmt_sell_date(_closed_sell_date(c))
                return flat(f"""
                    <div class="{row_cls}">
                        <div class="pt-symbol">
                            <span class="pt-symbol-name">{c.get('Symbol', '-')}</span>
                            {leader_badge}
                        </div>
                        <div class="pt-cell muted">{c.get('Exchange', '-')}</div>
                        <div class="pt-cell">{fmt_qty(c.get('Qty'))}</div>
                        <div class="pt-cell">{fmt_money(c.get('AvgBuyPrice'))}</div>
                        <div class="pt-cell">{fmt_money(c.get('AvgSellPrice'))}</div>
                        <div class="pt-cell muted">{sell_date}</div>
                        <div class="{pnl_cls}">{pnl_arrow} {fmt_money(pnl)}</div>
                    </div>
                """)

            def fno_group_summary(stock, stock_rows):
                """Collapse a stock's/index's closed F&O lots (e.g. a position
                rolled from one expiry series to the next) into ONE row: qty
                as the sum across lots, buy/sell price as the qty-weighted
                average across lots, and the sell date shown as a range. Pure
                display aggregation — the underlying FIFO-matched lots
                (stock_rows, straight from build_positions/closed_positions)
                are left completely unchanged and remain visible when this
                summary row is expanded."""
                total_qty = sum(abs(c.get("Qty") or 0) for c in stock_rows)
                if total_qty:
                    avg_buy = sum(abs(c.get("Qty") or 0) * (c.get("AvgBuyPrice") or 0) for c in stock_rows) / total_qty
                    avg_sell = sum(abs(c.get("Qty") or 0) * (c.get("AvgSellPrice") or 0) for c in stock_rows) / total_qty
                else:
                    avg_buy = avg_sell = 0.0
                dates = [_closed_sell_date(c) for c in stock_rows if _closed_sell_date(c)]
                if dates:
                    d_min, d_max = min(dates), max(dates)
                    date_disp = fmt_sell_date(d_min) if d_min == d_max else f"{fmt_sell_date(d_min)} – {fmt_sell_date(d_max)}"
                else:
                    date_disp = "-"
                return {
                    "Symbol": stock,
                    "Exchange": stock_rows[0].get("Exchange", "-") if stock_rows else "-",
                    "Qty": total_qty,
                    "AvgBuyPrice": avg_buy,
                    "AvgSellPrice": avg_sell,
                    "SellDateDisplay": date_disp,
                    "BookedPnL": sum(x["BookedPnL"] for x in stock_rows),
                }

            def fno_summary_row_html(summary, lot_count, is_top=False):
                pnl = summary["BookedPnL"]
                pnl_cls = "pt-cell pos" if pnl >= 0 else "pt-cell neg"
                pnl_arrow = "▲" if pnl >= 0 else "▼"
                row_cls = "pos-table-row cols-closed" + (" top-gain-row" if is_top else "")
                if is_top and pnl > 0:
                    leader_badge = '<span class="pt-leader-badge">🔥 Top Gain</span>'
                elif is_top:
                    leader_badge = '<span class="pt-leader-badge">🛡️ Least Loss</span>'
                else:
                    leader_badge = ""
                return flat(f"""
                    <summary class="{row_cls}">
                        <div class="pt-symbol">
                            <span class="pt-expand-chevron">▶</span>
                            <span class="pt-symbol-name">{summary['Symbol']}</span>
                            <span class="pt-lot-count">{lot_count} lots · wtd avg</span>
                            {leader_badge}
                        </div>
                        <div class="pt-cell muted">{summary['Exchange']}</div>
                        <div class="pt-cell">{fmt_qty(summary['Qty'])}</div>
                        <div class="pt-cell">{fmt_money(summary['AvgBuyPrice'])}</div>
                        <div class="pt-cell">{fmt_money(summary['AvgSellPrice'])}</div>
                        <div class="pt-cell muted">{summary['SellDateDisplay']}</div>
                        <div class="{pnl_cls}">{pnl_arrow} {fmt_money(pnl)}</div>
                    </summary>
                """)

            for c in closed:
                c["Segment"] = effective_segment(c.get("Segment"), c.get("InstrumentType", ""), c.get("Symbol", ""))
            closed_equity = [c for c in closed if c.get("Segment") == "Equity"]
            closed_fo = [c for c in closed if c.get("Segment") == "F&O"]
            closed_segments = [("Equity", closed_equity), ("F&O", closed_fo)]
            closed_other = [c for c in closed if c.get("Segment") not in ("Equity", "F&O")]
            if closed_other:
                closed_segments.append(("Other", closed_other))

            # One sub-tab per segment so picking "F&O" shows only F&O, not every
            # segment stacked one after another.
            closed_seg_tabs = st.tabs([f"{label} ({len(rows)})" for label, rows in closed_segments])
            for (label, rows), seg_tab in zip(closed_segments, closed_seg_tabs):
              with seg_tab:
                if not rows:
                    st.markdown(f'<div class="section-label">{label} <span class="badge">0</span></div>', unsafe_allow_html=True)
                    st.caption(f"No closed {label.lower()} positions in this ledger.")
                    continue
                seg_win_rate = sum(1 for c in rows if c["BookedPnL"] >= 0) / len(rows) * 100
                st.markdown(
                    f'<div class="section-label">{label} <span class="badge">{len(rows)}</span></div>',
                    unsafe_allow_html=True,
                )
                # One-line summary scoped to THIS segment's closed positions
                # only — cost basis deployed here vs. what was booked here.
                seg_invested = sum(abs(c.get("Qty") or 0) * (c.get("AvgBuyPrice") or 0) for c in rows)
                seg_booked = sum(c["BookedPnL"] for c in rows)
                seg_booked_pct = (seg_booked / seg_invested * 100) if seg_invested else 0.0
                st.markdown(
                    seg_summary_line([
                        (f"{label} Investment", fmt_money(seg_invested), ""),
                        (
                            "Booked P&L",
                            f"{fmt_money(seg_booked)} ({'+' if seg_booked >= 0 else ''}{seg_booked_pct:.2f}%)",
                            "kpi-pos" if seg_booked >= 0 else "kpi-neg",
                        ),
                        ("Win rate", f"{seg_win_rate:.0f}%", ""),
                    ]),
                    unsafe_allow_html=True,
                )
                # Whichever row has the single highest booked P&L in this
                # segment always gets the glow — a genuine gain reads "Top
                # Gain", otherwise the least-bad loss reads "Least Loss".
                best_pnl = max((c["BookedPnL"] for c in rows), default=None)

                closed_table_head = flat("""
                    <div class="pos-table-head cols-closed">
                        <div>Symbol</div><div>Exchange</div><div>Qty</div>
                        <div>Avg Buy</div><div>Avg Sell</div><div>Sell Date</div>
                        <div>Booked P&amp;L</div>
                    </div>
                """)

                def closed_total_row_html(total_pnl):
                    total_cls = "pt-cell pos" if total_pnl >= 0 else "pt-cell neg"
                    return flat(f"""
                        <div class="pos-table-row cols-closed" style="font-weight:700;border-top:1px solid var(--border);">
                            <div class="pt-symbol"><span class="pt-symbol-name">Total</span></div>
                            <div class="pt-cell muted">-</div>
                            <div class="pt-cell">-</div>
                            <div class="pt-cell">-</div>
                            <div class="pt-cell">-</div>
                            <div class="pt-cell muted">-</div>
                            <div class="{total_cls}">{fmt_money(total_pnl)}</div>
                        </div>
                    """)

                if label == "F&O":
                    # Group each stock's/index's contracts together — e.g. a
                    # position rolled from one expiry series to the next
                    # shows as one block with a Total row, matching the
                    # client-facing Excel report layout — instead of one
                    # flat list of unrelated-looking rows.
                    groups = {}
                    for c in rows:
                        groups.setdefault(underlying_symbol(c.get("Symbol", "-")), []).append(c)
                    stock_order = sorted(
                        groups.keys(),
                        key=lambda s: sum(x["BookedPnL"] for x in groups[s]),
                        reverse=True,
                    )
                    for stock in stock_order:
                        stock_rows = sorted(
                            groups[stock],
                            key=lambda x: _closed_sell_date(x) or datetime.min,
                        )
                        stock_total = sum(x["BookedPnL"] for x in stock_rows)
                        # Top gain is decided by the STOCK's aggregate
                        # (weighted-average) booked P&L — stock_order is
                        # already sorted by that same total, so the winner
                        # is always its first entry — never by whichever
                        # single leg inside a multi-lot group happens to
                        # have the best individual number (a big winning
                        # leg can still net out below another stock's
                        # smaller-but-consistent total once its losing
                        # legs are included).
                        stock_is_top = stock == stock_order[0]
                        st.markdown(
                            f'<div class="section-label" style="margin-top:16px;">{stock} '
                            f'<span class="badge">{len(stock_rows)}</span></div>',
                            unsafe_allow_html=True,
                        )
                        if len(stock_rows) > 1:
                            # Multiple lots (typically a rolled position) —
                            # one weighted-average line by default; expand to
                            # see each original FIFO-matched lot with its
                            # actual sell date, exactly as build_positions
                            # produced it (rows_html below is untouched).
                            summary = fno_group_summary(stock, stock_rows)
                            summary_html = fno_summary_row_html(summary, len(stock_rows), is_top=stock_is_top)
                            rows_html = [
                                # Individual legs never get their own
                                # top-gain badge now — only the group's
                                # aggregate summary row above does (see
                                # stock_is_top). A single big leg inside a
                                # losing/mediocre group used to get flagged
                                # here on its own, which was misleading.
                                closed_row_html(c)
                                for c in stock_rows
                            ]
                            detail_html = flat(f"""
                                <div class="pos-table-wrap">
                                    <div class="pos-table">
                                        <details class="pt-details">
                                            {summary_html}
                                            {closed_table_head}
                                            <div class="pt-details-body">
                                                {"".join(rows_html)}
                                                {closed_total_row_html(stock_total)}
                                            </div>
                                        </details>
                                    </div>
                                </div>
                            """)
                            st.markdown(detail_html, unsafe_allow_html=True)
                        else:
                            table_html = flat(f"""
                                <div class="pos-table-wrap">
                                    <div class="pos-table">
                                        {closed_table_head}
                                        {closed_row_html(stock_rows[0], is_top=stock_is_top)}
                                    </div>
                                </div>
                            """)
                            st.markdown(table_html, unsafe_allow_html=True)
                else:
                    sorted_rows = sorted(rows, key=lambda x: x["BookedPnL"], reverse=True)
                    rows_html = [
                        closed_row_html(
                            c,
                            is_top=(best_pnl is not None and c["BookedPnL"] == best_pnl),
                        )
                        for c in sorted_rows
                    ]
                    table_html = flat(f"""
                        <div class="pos-table-wrap">
                            <div class="pos-table">
                                {closed_table_head}
                                {"".join(rows_html)}
                            </div>
                        </div>
                    """)
                    st.markdown(table_html, unsafe_allow_html=True)

            # ── Charts — rendered below every position table so tables stay
            # the primary focus. Chart 1 stays full-width since a date axis
            # needs the room; charts 2 and 3 are compact enough to share a row.
            # Chart 1: cumulative booked profit over time, EQUITY ONLY — F&O
            # legs are intentionally excluded so the trend line reflects pure
            # equity performance and isn't skewed by F&O's larger swings.
            legs = []
            for p in closed:
                if p.get("Segment") != "Equity":
                    continue
                for t in p.get("Trades", []):
                    if t.get("SellDate"):
                        legs.append({"SellDate": t["SellDate"], "Pnl": t["Pnl"], "Segment": p.get("Segment", "Other")})

            if legs:
                # Multiple FIFO legs can share the exact same SellDate (e.g.
                # several lots/symbols closed on one day). If we cumsum one
                # row per leg, ties on SellDate get an arbitrary order, which
                # produces a zig-zagging line and — since the tooltip snaps to
                # the nearest x — a hover value that doesn't match what's
                # visually plotted at that point. Collapse to one point per
                # calendar day first so the x-axis is strictly increasing.
                chart_df = pd.DataFrame(legs)
                chart_df["SellDate"] = pd.to_datetime(chart_df["SellDate"]).dt.normalize()
                chart_df = chart_df.groupby("SellDate", as_index=False)["Pnl"].sum().sort_values("SellDate")
                chart_df["Cumulative"] = chart_df["Pnl"].cumsum()
                with st.container(border=True):
                    st.markdown('<div class="section-label">Cumulative booked profit over time (Equity)</div>', unsafe_allow_html=True)
                    area = alt.Chart(chart_df).mark_area(
                        line={"color": "#34d5c8", "strokeWidth": 2},
                        interpolate="monotone",
                        fillOpacity=0.15,
                        color=alt.Gradient(
                            gradient="linear",
                            stops=[alt.GradientStop(color="#34d5c8", offset=0), alt.GradientStop(color="transparent", offset=1)],
                            x1=1, x2=1, y1=1, y2=0,
                        ),
                    ).encode(
                        x=alt.X("SellDate:T", title=None),
                        y=alt.Y("Cumulative:Q", title="Cumulative ₹"),
                        tooltip=[alt.Tooltip("SellDate:T", title="Date"), alt.Tooltip("Cumulative:Q", title="Cumulative", format=",.0f")],
                    ).properties(height=240)
                    st.altair_chart(alt_dark(area), use_container_width=True)
            else:
                st.caption("No sell-dated trade legs available yet to plot a cumulative profit trend.")

            # Charts 2 + 3 share one row.
            top_df = pd.DataFrame(closed)[["Symbol", "BookedPnL"]].copy()
            top_df["AbsPnl"] = top_df["BookedPnL"].abs()
            top_df = top_df.sort_values("AbsPnl", ascending=False).head(12).drop(columns="AbsPnl")
            top_df["Direction"] = top_df["BookedPnL"].apply(lambda v: "Profit" if v >= 0 else "Loss")
            win_df = pd.DataFrame({"Outcome": ["Profitable", "Loss-making"], "Count": [wins, losses]})

            col_movers, col_donut = st.columns(2)
            with col_movers:
                with st.container(border=True):
                    st.markdown('<div class="section-label">Biggest movers — booked P&amp;L</div>', unsafe_allow_html=True)
                    bar = alt.Chart(top_df).mark_bar(cornerRadiusEnd=4).encode(
                        x=alt.X("BookedPnL:Q", title="Booked P&L (₹)"),
                        y=alt.Y("Symbol:N", sort="-x", title=None),
                        color=alt.Color(
                            "Direction:N",
                            scale=alt.Scale(domain=["Profit", "Loss"], range=["#2ee6a6", "#ff5c7a"]),
                            legend=None,
                        ),
                        tooltip=[alt.Tooltip("Symbol:N"), alt.Tooltip("BookedPnL:Q", title="Booked P&L", format=",.0f")],
                    ).properties(height=max(220, 24 * len(top_df)))
                    st.altair_chart(alt_dark(bar), use_container_width=True)

            with col_donut:
                with st.container(border=True):
                    st.markdown('<div class="section-label">Win / loss split</div>', unsafe_allow_html=True)
                    donut = alt.Chart(win_df).mark_arc(innerRadius=65, cornerRadius=3).encode(
                        theta=alt.Theta("Count:Q"),
                        color=alt.Color(
                            "Outcome:N",
                            scale=alt.Scale(domain=["Profitable", "Loss-making"], range=["#2ee6a6", "#ff5c7a"]),
                            legend=alt.Legend(orient="right", title=None),
                        ),
                        tooltip=[alt.Tooltip("Outcome:N"), alt.Tooltip("Count:Q")],
                    ).properties(height=220)
                    st.altair_chart(alt_dark(donut), use_container_width=True)

    with tab_news:
        st.markdown('<div class="section-label">News by stock</div>', unsafe_allow_html=True)
        open_symbols = sorted({t.get("symbol") for t in ticks if t.get("symbol")})
        closed_symbols = sorted({c.get("Symbol") for c in engine.closed_positions if c.get("Symbol")})

        news_tab_open, news_tab_closed = st.tabs([
            f"📈 Open positions ({len(open_symbols)})",
            f"✅ Closed positions ({len(closed_symbols)})",
        ])
        with news_tab_open:
            render_news_section(open_symbols, scope_key="open", scope_noun="open positions")
        with news_tab_closed:
            render_news_section(closed_symbols, scope_key="closed", scope_noun="closed positions")

    with tab_corp:
        from corporate_actions import (
            get_pending_actions, get_applied_actions, get_dividend_log,
            sync_pending_from_nse, confirm_action, discard_pending,
        )

        st.markdown('<div class="section-label">Corporate actions</div>', unsafe_allow_html=True)
        corp_symbols = sorted({t.get("symbol") for t in ticks if t.get("symbol")})

        if st.button("🔄 Check NSE for new actions", key="corp_sync_btn"):
            try:
                added = sync_pending_from_nse(corp_symbols)
                st.success(f"Found {added} new action(s).") if added else st.info("Nothing new from NSE.")
            except Exception as e:
                st.error(f"NSE fetch failed: {e}")

        open_qty_by_symbol = {}
        for t in ticks:
            if t.get("symbol"):
                open_qty_by_symbol[t["symbol"]] = open_qty_by_symbol.get(t["symbol"], 0) + (t.get("qty") or 0)

        st.markdown("**Pending — awaiting your confirmation**")
        pending = get_pending_actions()
        if not pending:
            st.caption("No pending corporate actions.")
        for a in pending:
            with st.container(border=True):
                st.write(f"**{a['symbol']}** — {a['type']} · Ex-date {a['ex_date']}")
                st.caption(a.get("raw_purpose", ""))
                if a.get("confidence") == "low":
                    st.warning("NSE text didn't parse cleanly — verify ratio/amount before applying.")
                held_qty = open_qty_by_symbol.get(a["symbol"], 0.0)
                st.caption(f"Your current open qty (used for bonus/dividend calc): {held_qty}")
                c1, c2 = st.columns(2)
                if c1.button("✅ Apply", key=f"corp_apply_{a['id']}"):
                    confirm_action(a["id"], held_qty_hint=held_qty)
                    get_engine.clear()
                    st.rerun()
                if c2.button("✖ Discard", key=f"corp_discard_{a['id']}"):
                    discard_pending(a["id"])
                    st.rerun()

        st.markdown("**Applied history**")
        applied = get_applied_actions()
        if applied:
            st.dataframe(pd.DataFrame(applied), hide_index=True, use_container_width=True)
        else:
            st.caption("No corporate actions applied yet.")

        st.markdown("**Dividends received** &nbsp;<span style='color:var(--muted);font-size:.78rem;'>(income only — kept separate from booked P&amp;L)</span>", unsafe_allow_html=True)
        divs = get_dividend_log()
        if divs:
            st.dataframe(pd.DataFrame(divs), hide_index=True, use_container_width=True)
            st.metric("Total dividends received", f"₹{sum(d['amount'] for d in divs):,.0f}")
        else:
            st.caption("No dividends logged yet.")


# ── UI ──────────────────────────────────────────────────────────────────
def main():
    if not check_password():
        return

    inject_theme()

    client_name_disp = current_client_name()
    st.markdown(
        flat(f"""
        <div class="db-header">
            <div class="db-title">
                <div class="db-icon">📊</div>
                <h1>Booked Profit Dashboard</h1>
            </div>
            <div style="font-size:.82rem;color:var(--muted);">
                Portfolio of &nbsp;<strong style="color:var(--text);">{client_name_disp}</strong>
            </div>
        </div>
        """),
        unsafe_allow_html=True,
    )

    cfg        = current_client_cfg()
    client_name = current_client_name()
    script_url  = cfg.get("apps_script_url", "")
    sheet_tab   = cfg.get("sheet_name", "") or ""

    with st.sidebar:
        # ── Client badge ────────────────────────────────────────
        st.markdown(
            flat(f"""
            <div style="background:var(--panel-2);border:1px solid var(--border);
                        border-radius:12px;padding:12px 14px;margin-bottom:14px;">
              <div style="font-size:.68rem;font-weight:700;text-transform:uppercase;
                          letter-spacing:.06em;color:var(--muted);margin-bottom:4px;">
                Logged in as
              </div>
              <div style="font-size:1rem;font-weight:700;color:var(--text);">
                {client_name}
              </div>
              <div style="font-size:.72rem;color:var(--muted);margin-top:2px;">
                ID: {st.session_state.get("_client_id","?")}
                {"&nbsp;&nbsp;🔑 Admin" if is_admin() else ""}
              </div>
            </div>
            """),
            unsafe_allow_html=True,
        )

        if st.button("🚪 Sign out", use_container_width=True):
            for k in ["_client_ok","_client_id","_client_cfg","_client_name","_login_failed"]:
                st.session_state.pop(k, None)
            get_engine.clear()
            st.rerun()

        st.divider()

        # Admin: let them pick a different client to view
        if is_admin():
            clients = _get_clients()
            non_admin = {k: v for k, v in clients.items() if not v.get("is_admin")}
            if non_admin:
                chosen = st.selectbox(
                    "View client",
                    options=["(All merged)"] + list(non_admin.keys()),
                    format_func=lambda k: k if k == "(All merged)"
                                          else non_admin[k].get("display_name", k),
                    key="_admin_client_view",
                )
                if chosen != "(All merged)":
                    script_url = non_admin[chosen].get("apps_script_url", script_url)
                    sheet_tab  = non_admin[chosen].get("sheet_name", "") or ""

        if st.button("🔄 Restart feed / refetch sheet", use_container_width=True):
            get_engine.clear()
            # Only reset the shared Angel One connection if it's actually
            # down — everyone shares one feed now, so we don't want a
            # routine per-client refresh to force a fresh broker login
            # (and disrupt other viewers) when the feed is perfectly fine.
            if get_shared_feed().status in ("error", "disconnected"):
                get_shared_feed.clear()
            st.rerun()
        st.caption(f"Live tables refresh every {TICK_REFRESH_SECONDS}s, tick by tick.")

        st.divider()
        access_log = _get_access_log()
        with st.expander(f"🔐 Access log ({len(access_log)})"):
            if not access_log:
                st.caption("No entries yet.")
            else:
                log_df = pd.DataFrame(
                    [{"Name": e["name"], "Time": e["ts"].strftime("%d %b %Y, %I:%M:%S %p")}
                     for e in reversed(access_log[-100:])]
                )
                st.dataframe(log_df, hide_index=True, use_container_width=True)
                st.caption(
                    f"Showing latest {min(100, len(access_log))} of {len(access_log)} (IST). "
                    "In-memory only, resets on app restart/redeploy."
                )

    if not script_url:
        st.warning(
            f"No Apps Script URL configured for client **{client_name}**. "
            "Ask your administrator to add it to the app secrets."
        )
        return

    engine = get_engine(script_url, sheet_tab or None)
    render_live(engine)


if __name__ == "__main__":
    main()
