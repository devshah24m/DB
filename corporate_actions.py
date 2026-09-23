"""
corporate_actions.py
=====================
Fetches corporate actions (Split / Bonus / Dividend / Rights / other) for
your held symbols directly from NSE India, parses them into a structured
form, stores them locally (Pending -> Applied), and — once you confirm —
mutates the trade ledger rows in the exact shape `positions_builder.py`
expects, so `build_positions()` downstream needs no changes at all.

INTEGRATION POINTS (in app__17_.py):

  1. After you fetch `records` from the Apps Script (the raw ledger rows,
     before `load_trade_ledger_from_records`), call:

         records = apply_confirmed_actions(records, get_applied_actions())

     This rewrites/injects rows for every action you've already clicked
     "Apply" on, BEFORE FIFO matching runs — so open/closed positions,
     avg price, everything downstream is automatically correct.

  2. Once a day (or on "Restart feed"), refresh what NSE has for your
     symbols:

         symbols = sorted({r.get("symbol") or r.get("Symbol") for r in records if ...})
         sync_pending_from_nse(symbols)

  3. Render a "Corporate Actions" section (new tab) using:

         get_pending_actions()   -> list to show with an "Apply" button each
         get_applied_actions()   -> history table
         get_dividend_log()      -> dividend income, kept OUT of booked P&L

     When the user clicks Apply on a pending action, call:

         confirm_action(action_id, held_qty_hint=<open qty for that symbol>)

WHY THIS IS SEPARATE FROM YOUR REAL TRADES:
Every row this module injects/rewrites carries `"_source": "Corporate
Action - <Type>"` so you can filter/tag it distinctly anywhere in the UI
(e.g. `[r for r in closed_or_open if r.get("_source","").startswith("Corporate Action")]`).
"""
from __future__ import annotations

import json
import os
import re
import time
from datetime import datetime, date

import pandas as pd
import requests

# ── Storage ──────────────────────────────────────────────────────────────
STORE_PATH = os.environ.get("CORP_ACTIONS_STORE", "corporate_actions_store.json")

_EMPTY_STORE = {"pending": [], "applied": [], "dividends": [], "seen_keys": []}


def _load_store() -> dict:
    if not os.path.exists(STORE_PATH):
        return dict(_EMPTY_STORE)
    try:
        with open(STORE_PATH, "r") as f:
            data = json.load(f)
        for k, v in _EMPTY_STORE.items():
            data.setdefault(k, v if not isinstance(v, list) else [])
        return data
    except Exception:
        return dict(_EMPTY_STORE)


def _save_store(store: dict):
    tmp = STORE_PATH + ".tmp"
    with open(tmp, "w") as f:
        json.dump(store, f, indent=2, default=str)
    os.replace(tmp, STORE_PATH)


def get_pending_actions() -> list:
    return _load_store()["pending"]


def get_applied_actions() -> list:
    return _load_store()["applied"]


def get_dividend_log() -> list:
    return _load_store()["dividends"]


# ── NSE fetch ────────────────────────────────────────────────────────────
NSE_BASE = "https://www.nseindia.com"
NSE_ACTIONS_URL = f"{NSE_BASE}/api/corporates-corporateActions"

_HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
    "Referer": f"{NSE_BASE}/companies-listing/corporate-filings-actions",
}


def _nse_session() -> requests.Session:
    """NSE blocks bare API calls with a 403 unless you first hold cookies
    from a normal page load — this mimics that handshake."""
    s = requests.Session()
    s.headers.update(_HEADERS)
    s.get(NSE_BASE, timeout=10)
    s.get(f"{NSE_BASE}/companies-listing/corporate-filings-actions", timeout=10)
    return s


def fetch_actions_for_symbol(symbol: str, session: requests.Session | None = None,
                              retries: int = 2) -> list[dict]:
    """Raw NSE JSON rows for one symbol. Each row typically has keys like
    'symbol', 'subject' or 'purpose', 'exDate', 'recDate', 'faceVal',
    'series' — NSE's exact key names have shifted before, so callers should
    use .get(...) defensively (see _extract_purpose/_extract_exdate below).
    """
    last_err = None
    for attempt in range(retries + 1):
        s = session or _nse_session()
        try:
            r = s.get(NSE_ACTIONS_URL, params={"index": "equities", "symbol": symbol}, timeout=15)
            r.raise_for_status()
            data = r.json()
            return data if isinstance(data, list) else data.get("data", [])
        except Exception as e:
            last_err = e
            session = None  # force a fresh cookie handshake on retry
            time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f"NSE corporate-actions fetch failed for {symbol}: {last_err}")


def _extract_purpose(raw: dict) -> str:
    for k in ("subject", "purpose", "comp", "desc"):
        if raw.get(k):
            return str(raw[k])
    return ""


def _extract_exdate(raw: dict):
    for k in ("exDate", "exdt", "exDt", "recDate"):
        if raw.get(k):
            return pd.to_datetime(raw[k], errors="coerce", dayfirst=True)
    return pd.NaT


# ── Parsing "Purpose" free text into a structured action ──────────────────
_SPLIT_RE = re.compile(
    r"FACE\s*VALUE\s*SPLIT.*?FROM\s*RS\.?\s*([\d.]+).*?TO\s*RS\.?\s*([\d.]+)", re.I)
_BONUS_RE = re.compile(r"BONUS\s*(?:ISSUE\s*)?(\d+)\s*[:/]\s*(\d+)", re.I)
_DIV_RE = re.compile(r"DIVIDEND[^0-9]{0,20}RS\.?\s*([\d.]+)\s*/?-?\s*PER\s*SHARE", re.I)


def parse_purpose(purpose: str) -> dict | None:
    """Returns {"type": "Split"|"Bonus"|"Dividend", ..., "confidence": "high"|"low"}
    or None if it's not one of these three action types (rights/merger/AGM etc.
    are left for you to add manually)."""
    if not purpose:
        return None
    p = purpose.upper()

    m = _SPLIT_RE.search(p)
    if m:
        old_fv, new_fv = float(m.group(1)), float(m.group(2))
        if old_fv > 0 and new_fv > 0 and old_fv != new_fv:
            return {"type": "Split", "ratio": round(old_fv / new_fv, 6),
                    "raw": purpose, "confidence": "high"}

    m = _BONUS_RE.search(p)
    if m:
        n, d = float(m.group(1)), float(m.group(2))
        if d > 0:
            return {"type": "Bonus", "ratio": round(n / d, 6),
                    "raw": purpose, "confidence": "high"}

    m = _DIV_RE.search(p)
    if m:
        return {"type": "Dividend", "per_share": float(m.group(1)),
                "raw": purpose, "confidence": "high"}

    # Couldn't parse a clean number — still surface it, flagged for manual entry.
    if "DIVIDEND" in p:
        return {"type": "Dividend", "per_share": None, "raw": purpose, "confidence": "low"}
    if "BONUS" in p:
        return {"type": "Bonus", "ratio": None, "raw": purpose, "confidence": "low"}
    if "SPLIT" in p or "SUB-DIVISION" in p or "SUB DIVISION" in p:
        return {"type": "Split", "ratio": None, "raw": purpose, "confidence": "low"}
    return None


# ── Sync: pull NSE actions for held symbols into the Pending list ─────────
def sync_pending_from_nse(symbols: list[str]) -> int:
    """Fetches NSE actions for each symbol, parses them, and adds any new
    ones to Pending (deduped by symbol+exDate+raw purpose). Returns count
    of newly added pending actions. Safe to call repeatedly."""
    store = _load_store()
    seen = set(store["seen_keys"])
    added = 0
    session = _nse_session()

    for sym in symbols:
        try:
            raw_rows = fetch_actions_for_symbol(sym, session=session)
        except Exception as e:
            store.setdefault("fetch_errors", {})[sym] = str(e)
            continue

        for raw in raw_rows:
            purpose = _extract_purpose(raw)
            parsed = parse_purpose(purpose)
            if not parsed:
                continue
            ex_date = _extract_exdate(raw)
            if pd.isna(ex_date):
                continue
            ex_date_str = ex_date.strftime("%Y-%m-%d")
            key = f"{sym}|{ex_date_str}|{purpose}"
            if key in seen:
                continue
            seen.add(key)
            store["pending"].append({
                "id": key,
                "symbol": sym,
                "type": parsed["type"],
                "ratio": parsed.get("ratio"),
                "per_share": parsed.get("per_share"),
                "confidence": parsed["confidence"],
                "raw_purpose": purpose,
                "ex_date": ex_date_str,
                "fetched_at": datetime.now().isoformat(timespec="seconds"),
                "status": "pending",
            })
            added += 1

    store["seen_keys"] = list(seen)
    _save_store(store)
    return added


# ── Applying a confirmed action to the ledger ──────────────────────────────
def _norm(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip().upper()


def _apply_split(rows: list[dict], symbol: str, ratio: float, ex_date: str) -> list[dict]:
    """Rewrites qty/price of every existing buy/sell leg for `symbol` dated
    strictly before ex_date: Quantity *= ratio, Price /= ratio. Invested
    value is unchanged; this keeps FIFO matching correct going forward."""
    ex_ts = pd.to_datetime(ex_date)
    out = []
    for r in rows:
        r = dict(r)
        if _norm(r.get("Symbol") or r.get("symbol")) == _norm(symbol):
            for price_col, date_col in (("BuyPrice", "BuyDate"), ("Buy Price", "Buy Date"),
                                         ("SellPrice", "SellDate"), ("Sell Price", "Sell Date")):
                if price_col in r and date_col in r:
                    d = pd.to_datetime(r.get(date_col), errors="coerce")
                    px = r.get(price_col)
                    if pd.notna(d) and d < ex_ts and px not in ("", None) and not pd.isna(px):
                        qty_col = "Quantity" if "Quantity" in r else "Qty"
                        r[qty_col] = float(r.get(qty_col) or 0) * ratio
                        r[price_col] = float(px) / ratio
                        r["_source"] = r.get("_source", "") or "Corporate Action - Split"
        out.append(r)
    return out


def _apply_bonus(rows: list[dict], symbol: str, exchange: str, ratio: float,
                  ex_date: str, held_qty_on_ex_date: float) -> list[dict]:
    """Injects one synthetic BUY row: qty = held_qty * ratio, price = 0,
    dated on ex_date, tagged as a Corporate Action so it's distinguishable
    from real trades everywhere downstream."""
    bonus_qty = round(held_qty_on_ex_date * ratio, 4)
    if bonus_qty <= 0:
        return rows
    synthetic = {
        "Symbol": symbol, "Exchange": exchange or "NSE", "InstrumentType": "EQ",
        "Expiry": "", "OptionType": "", "Strike": "",
        "Quantity": bonus_qty, "BuyPrice": 0.0, "BuyDate": ex_date,
        "SellPrice": "", "SellDate": "",
        "_source": "Corporate Action - Bonus",
    }
    return rows + [synthetic]


def confirm_action(action_id: str, held_qty_hint: float = 0.0, exchange_hint: str = "NSE",
                    manual_ratio: float | None = None, manual_per_share: float | None = None) -> dict:
    """Marks a pending action as applied and returns the action dict — the
    caller (app__17_.py) is responsible for calling apply_confirmed_actions()
    on the NEXT ledger load so the split/bonus rewrite actually takes effect
    (this function only updates the store, it doesn't touch live `records`
    in memory since the caller owns that).

    held_qty_hint: your current open quantity for this symbol — required for
    Bonus (to compute how many free shares you get) and Dividend (to compute
    the payout). Pass the qty as displayed in Open Positions for that symbol
    at the time you click Apply.
    manual_ratio / manual_per_share: fill these in if confidence was "low"
    and NSE's text didn't parse cleanly — lets you correct it before applying.
    """
    store = _load_store()
    action = next((a for a in store["pending"] if a["id"] == action_id), None)
    if not action:
        raise ValueError(f"No pending action with id {action_id}")

    if manual_ratio is not None:
        action["ratio"] = manual_ratio
    if manual_per_share is not None:
        action["per_share"] = manual_per_share

    action["status"] = "applied"
    action["applied_at"] = datetime.now().isoformat(timespec="seconds")
    action["held_qty_at_apply"] = held_qty_hint
    action["exchange"] = exchange_hint

    if action["type"] == "Dividend":
        amount = round((held_qty_hint or 0) * (action.get("per_share") or 0), 2)
        store["dividends"].append({
            "symbol": action["symbol"], "ex_date": action["ex_date"],
            "per_share": action.get("per_share"), "qty_held": held_qty_hint,
            "amount": amount, "applied_at": action["applied_at"],
        })

    store["pending"] = [a for a in store["pending"] if a["id"] != action_id]
    store["applied"].append(action)
    _save_store(store)
    return action


def apply_confirmed_actions(rows: list[dict], applied_actions: list[dict] | None = None) -> list[dict]:
    """Call this on your raw ledger `records` BEFORE
    load_trade_ledger_from_records()/build_positions(). Replays every
    applied Split/Bonus (Dividend needs no ledger change) in ex_date order
    so multiple splits/bonuses on the same symbol stack correctly."""
    applied_actions = applied_actions if applied_actions is not None else get_applied_actions()
    actions = sorted(
        [a for a in applied_actions if a["type"] in ("Split", "Bonus")],
        key=lambda a: a["ex_date"],
    )
    out = rows
    for a in actions:
        if a["type"] == "Split" and a.get("ratio"):
            out = _apply_split(out, a["symbol"], a["ratio"], a["ex_date"])
        elif a["type"] == "Bonus" and a.get("ratio"):
            out = _apply_bonus(out, a["symbol"], a.get("exchange", "NSE"),
                                a["ratio"], a["ex_date"], a.get("held_qty_at_apply", 0.0))
    return out


def discard_pending(action_id: str):
    """Dismiss a pending action without applying it (e.g. it's irrelevant,
    or you already exited the position before ex-date)."""
    store = _load_store()
    store["pending"] = [a for a in store["pending"] if a["id"] != action_id]
    _save_store(store)
