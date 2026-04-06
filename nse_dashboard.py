"""
NSE Options Dashboard Generator
================================
Fetches real End-of-Day data from NSE for BHARTIARTL & TATAMOTORS
and generates a beautiful HTML dashboard.

Run after 3:45 PM IST on any trading day.
"""

import requests
import json
import os
import math
import time
import webbrowser
from datetime import datetime, date
from pathlib import Path

# ─────────────────────────────────────────────
#  CONFIG — edit stocks here if needed
# ─────────────────────────────────────────────
STOCKS = {
    "BHARTIARTL": {"name": "BHARTIARTL", "sector": "Telecom", "lot": 1},
    "TATAMOTORS":  {"name": "TMPV",   "sector": "Auto",    "lot": 1},
}

OUTPUT_FILE = "nse_options_dashboard.html"

# ─────────────────────────────────────────────
#  NSE SESSION SETUP  (mimics browser headers)
# ─────────────────────────────────────────────
def get_nse_session():
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Referer": "https://www.nseindia.com/",
        "X-Requested-With": "XMLHttpRequest",
    })
    # Prime the session with a cookie
    print("  Connecting to NSE...")
    session.get("https://www.nseindia.com", timeout=10)
    time.sleep(1)
    session.get("https://www.nseindia.com/option-chain", timeout=10)
    time.sleep(1)
    return session


# ─────────────────────────────────────────────
#  FETCH OPTION CHAIN
# ─────────────────────────────────────────────
def fetch_option_chain(session, symbol):
    url = f"https://www.nseindia.com/api/option-chain-equities?symbol={symbol}"
    try:
        resp = session.get(url, timeout=15)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print(f"  ⚠  Could not fetch option chain for {symbol}: {e}")
        return None


# ─────────────────────────────────────────────
#  FETCH EQUITY QUOTE (spot price, delivery%)
# ─────────────────────────────────────────────
def fetch_equity_quote(session, symbol):
    url = f"https://www.nseindia.com/api/quote-equity?symbol={symbol}&section=trade_info"
    try:
        resp = session.get(url, timeout=10)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        print(f"  ⚠  Could not fetch quote for {symbol}: {e}")
        return None


# ─────────────────────────────────────────────
#  PARSE & CALCULATE
# ─────────────────────────────────────────────
def parse_data(oc_data, quote_data, symbol):
    if not oc_data:
        return None

    records  = oc_data.get("records", {})
    filtered = oc_data.get("filtered", {})
    spot     = records.get("underlyingValue", 0)

    # Delivery % from quote
    delivery_pct = 0
    delivery_trend = "NEUTRAL"
    if quote_data:
        try:
            ti = quote_data.get("marketDeptOrderBook", {})
            td = quote_data.get("securityWiseDP", {})
            delivery_pct   = float(td.get("deliveryToTradedQuantity", 0))
            delivery_trend = (
                "ACCUMULATION" if delivery_pct > 55
                else "DISTRIBUTION" if delivery_pct < 35
                else "NEUTRAL"
            )
        except Exception:
            pass

    # Nearest expiry
    expiry_dates = records.get("expiryDates", [])
    expiry       = expiry_dates[0] if expiry_dates else "N/A"
    try:
        exp_dt       = datetime.strptime(expiry, "%d-%b-%Y")
        days_to_exp  = max(1, (exp_dt.date() - date.today()).days)
        expiry_type  = "weekly" if days_to_exp <= 7 else "monthly"
    except Exception:
        days_to_exp = 0
        expiry_type = "monthly"

    # Build strike map for nearest expiry
    strike_map = {}
    total_call_oi = 0
    total_put_oi  = 0
    total_call_oi_chg = 0
    total_put_oi_chg  = 0

    for item in records.get("data", []):
        if item.get("expiryDate") != expiry:
            continue
        strike = item["strikePrice"]
        if strike not in strike_map:
            strike_map[strike] = {}
        if "CE" in item:
            strike_map[strike]["CE"] = item["CE"]
            total_call_oi     += item["CE"].get("openInterest", 0)
            total_call_oi_chg += item["CE"].get("changeinOpenInterest", 0)
        if "PE" in item:
            strike_map[strike]["PE"] = item["PE"]
            total_put_oi     += item["PE"].get("openInterest", 0)
            total_put_oi_chg += item["PE"].get("changeinOpenInterest", 0)

    # PCR
    pcr = round(total_put_oi / total_call_oi, 2) if total_call_oi else 1.0

    # ATM strike
    all_strikes = sorted(strike_map.keys())
    atm_strike  = min(all_strikes, key=lambda x: abs(x - spot)) if all_strikes else spot

    # Select 7 strikes: 3 ITM, 1 ATM, 3 OTM
    atm_idx = all_strikes.index(atm_strike)
    selected_idxs = range(
        max(0, atm_idx - 3),
        min(len(all_strikes), atm_idx + 4)
    )
    selected_strikes = [all_strikes[i] for i in selected_idxs]

    strikes_out = []
    atm_iv      = 0
    atm_ce_ltp  = 0
    atm_pe_ltp  = 0

    for s in selected_strikes:
        ce = strike_map.get(s, {}).get("CE", {})
        pe = strike_map.get(s, {}).get("PE", {})
        is_atm = (s == atm_strike)
        if is_atm:
            atm_iv     = ce.get("impliedVolatility", 0)
            atm_ce_ltp = ce.get("lastPrice", 0)
            atm_pe_ltp = pe.get("lastPrice", 0)
        strikes_out.append({
            "strike":      s,
            "isATM":       is_atm,
            "callOI":      ce.get("openInterest", 0),
            "callOIChg":   ce.get("changeinOpenInterest", 0),
            "callIV":      ce.get("impliedVolatility", 0),
            "callLTP":     ce.get("lastPrice", 0),
            "callVolume":  ce.get("totalTradedVolume", 0),
            "putOI":       pe.get("openInterest", 0),
            "putOIChg":    pe.get("changeinOpenInterest", 0),
            "putIV":       pe.get("impliedVolatility", 0),
            "putLTP":      pe.get("lastPrice", 0),
            "putVolume":   pe.get("totalTradedVolume", 0),
        })

    # Max pain
    max_pain_strike = calc_max_pain(strike_map, all_strikes)

    # IV Percentile (approximation using current ATM IV)
    iv_pct = min(99, max(1, round((atm_iv / 50) * 100))) if atm_iv else 50

    # Greeks (Black-Scholes ATM approximation)
    greeks = calc_greeks(spot, atm_strike, atm_iv / 100, days_to_exp / 365)

    # Support / Resistance (using high-OI strikes)
    sorted_by_put_oi  = sorted(strike_map.items(),
                                key=lambda x: x[1].get("PE", {}).get("openInterest", 0),
                                reverse=True)
    sorted_by_call_oi = sorted(strike_map.items(),
                                key=lambda x: x[1].get("CE", {}).get("openInterest", 0),
                                reverse=True)
    support    = sorted([s for s, _ in sorted_by_put_oi[:2]  if s < spot])
    resistance = sorted([s for s, _ in sorted_by_call_oi[:2] if s > spot])

    # Trend & Signal
    trend, signal, strength = derive_signal(pcr, iv_pct, delivery_pct,
                                             total_call_oi_chg, total_put_oi_chg,
                                             spot, max_pain_strike)

    # Spreads
    spread = calc_spreads(spot, atm_strike, all_strikes, strike_map, atm_ce_ltp, atm_pe_ltp)

    # Spot change
    prev = records.get("underlyingValue", spot)  # fallback
    change     = round(filtered.get("CE", {}).get("change", 0), 2)
    change_pct = round((change / spot * 100), 2) if spot else 0

    return {
        "symbol":        symbol,
        "name":          STOCKS[symbol]["name"],
        "sector":        STOCKS[symbol]["sector"],
        "lot":           STOCKS[symbol]["lot"],
        "spotPrice":     spot,
        "change":        change,
        "changePct":     change_pct,
        "expiry":        expiry,
        "expiry_type":   expiry_type,
        "daysToExpiry":  days_to_exp,
        "deliveryPct":   round(delivery_pct, 1),
        "deliveryTrend": delivery_trend,
        "pcr":           pcr,
        "ivPercentile":  iv_pct,
        "maxPain":       max_pain_strike,
        "trend":         trend,
        "signal":        signal,
        "signalStrength":strength,
        "totalCallOI":   total_call_oi,
        "totalPutOI":    total_put_oi,
        "oiChange":      total_call_oi_chg - total_put_oi_chg,
        "strikes":       strikes_out,
        "greeks":        greeks,
        "support":       support[:2],
        "resistance":    resistance[:2],
        "spread":        spread,
        "atm_iv":        atm_iv,
    }


def calc_max_pain(strike_map, all_strikes):
    min_pain  = float("inf")
    max_pain_strike = all_strikes[len(all_strikes)//2] if all_strikes else 0
    for candidate in all_strikes:
        total_pain = 0
        for s, data in strike_map.items():
            ce_oi = data.get("CE", {}).get("openInterest", 0)
            pe_oi = data.get("PE", {}).get("openInterest", 0)
            if candidate > s:
                total_pain += ce_oi * (candidate - s)
            if candidate < s:
                total_pain += pe_oi * (s - candidate)
        if total_pain < min_pain:
            min_pain = total_pain
            max_pain_strike = candidate
    return max_pain_strike


def calc_greeks(spot, strike, iv, T):
    if T <= 0 or iv <= 0:
        return {"delta": 0.5, "gamma": 0, "theta": 0, "vega": 0, "iv": round(iv*100, 1)}
    r  = 0.065  # India risk-free rate
    d1 = (math.log(spot / strike) + (r + 0.5 * iv**2) * T) / (iv * math.sqrt(T))
    d2 = d1 - iv * math.sqrt(T)
    nd1     = norm_cdf(d1)
    nd1_pdf = norm_pdf(d1)
    delta   = round(nd1, 4)
    gamma   = round(nd1_pdf / (spot * iv * math.sqrt(T)), 6)
    theta   = round((-spot * nd1_pdf * iv / (2 * math.sqrt(T))
                     - r * strike * math.exp(-r * T) * norm_cdf(d2)) / 365, 4)
    vega    = round(spot * nd1_pdf * math.sqrt(T) / 100, 4)
    return {"delta": delta, "gamma": gamma, "theta": theta, "vega": vega, "iv": round(iv*100, 1)}


def norm_cdf(x):
    return 0.5 * (1 + math.erf(x / math.sqrt(2)))

def norm_pdf(x):
    return math.exp(-0.5 * x**2) / math.sqrt(2 * math.pi)


def calc_spreads(spot, atm, all_strikes, strike_map, ce_ltp, pe_ltp):
    strikes = sorted(all_strikes)
    atm_idx = strikes.index(atm) if atm in strikes else len(strikes)//2
    spread  = {}
    # Bull Call: buy ATM CE, sell next OTM CE
    if atm_idx + 1 < len(strikes):
        sell_strike = strikes[atm_idx + 1]
        sell_ce_ltp = strike_map.get(sell_strike, {}).get("CE", {}).get("lastPrice", 0)
        net = round(ce_ltp - sell_ce_ltp, 2)
        spread["bullCall"] = {
            "buyStrike":  atm,
            "sellStrike": sell_strike,
            "netPremium": net,
            "maxProfit":  round((sell_strike - atm - net) * STOCKS.get("BHARTIARTL", {}).get("lot", 1), 2),
            "maxLoss":    round(net * STOCKS.get("BHARTIARTL", {}).get("lot", 1), 2),
            "breakeven":  round(atm + net, 2),
        }
    # Bear Put: buy ATM PE, sell next ITM PE
    if atm_idx - 1 >= 0:
        sell_strike = strikes[atm_idx - 1]
        sell_pe_ltp = strike_map.get(sell_strike, {}).get("PE", {}).get("lastPrice", 0)
        net = round(pe_ltp - sell_pe_ltp, 2)
        spread["bearPut"] = {
            "buyStrike":  atm,
            "sellStrike": sell_strike,
            "netPremium": net,
            "maxProfit":  round((atm - sell_strike - net) * STOCKS.get("TATAMOTORS", {}).get("lot", 1), 2),
            "maxLoss":    round(net * STOCKS.get("TATAMOTORS", {}).get("lot", 1), 2),
            "breakeven":  round(atm - net, 2),
        }
    return spread


def derive_signal(pcr, iv_pct, delivery_pct, call_oi_chg, put_oi_chg, spot, max_pain):
    score = 0
    if pcr > 1.3:   score += 2
    elif pcr < 0.7: score -= 2
    elif pcr > 1.0: score += 1
    else:           score -= 1

    if put_oi_chg > call_oi_chg: score += 1
    else:                         score -= 1

    if delivery_pct > 55: score += 1
    elif delivery_pct < 35: score -= 1

    if spot > max_pain: score += 1
    elif spot < max_pain: score -= 1

    if score >= 3:
        return "BULLISH", "BUY CE", min(10, 5 + score)
    elif score <= -3:
        return "BEARISH", "BUY PE", min(10, 5 + abs(score))
    elif score > 0:
        return "BULLISH", "BUY CE", 4 + score
    elif score < 0:
        return "BEARISH", "BUY PE", 4 + abs(score)
    else:
        return "NEUTRAL", "WAIT", 3


# ─────────────────────────────────────────────
#  HTML GENERATOR
# ─────────────────────────────────────────────
def generate_html(all_data):
    now   = datetime.now().strftime("%d %b %Y, %I:%M %p")
    cards = "".join(stock_card_html(d) for d in all_data if d)

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>NSE Options Dashboard — {now}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=JetBrains+Mono:wght@400;600;700&display=swap" rel="stylesheet">
<style>
  :root {{
    --bg:       #010409;
    --surface:  #0d1117;
    --card:     #161b22;
    --border:   #21262d;
    --border2:  #30363d;
    --text:     #e6edf3;
    --muted:    #8b949e;
    --green:    #00e676;
    --red:      #ff5252;
    --blue:     #58a6ff;
    --yellow:   #ffd740;
    --purple:   #bc8cff;
    --orange:   #ff9100;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ background: var(--bg); color: var(--text); font-family: 'Syne', sans-serif; min-height: 100vh; }}
  ::-webkit-scrollbar {{ width: 6px; height: 6px; }}
  ::-webkit-scrollbar-track {{ background: var(--bg); }}
  ::-webkit-scrollbar-thumb {{ background: var(--border2); border-radius: 3px; }}

  /* ── TOP BAR ── */
  .topbar {{
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    padding: 14px 28px;
    display: flex; justify-content: space-between; align-items: center;
    position: sticky; top: 0; z-index: 100;
  }}
  .logo {{ font-weight: 800; font-size: 20px; letter-spacing: -0.5px; }}
  .logo span.accent {{ color: var(--blue); }}
  .logo span.dim {{ color: var(--muted); font-weight: 400; font-size: 14px; margin-left: 8px; }}
  .timestamp {{ color: var(--muted); font-size: 12px; font-family: 'JetBrains Mono', monospace; }}
  .badge-live {{
    background: #00e67622; color: var(--green);
    border: 1px solid #00e67644; border-radius: 4px;
    padding: 3px 10px; font-size: 11px; font-weight: 700; letter-spacing: 1px;
    margin-left: 10px;
  }}

  /* ── DISCLAIMER ── */
  .disclaimer {{
    background: #ffd74011; border: 1px solid #ffd74033;
    margin: 16px 28px 0; border-radius: 8px;
    padding: 8px 16px; font-size: 11px; color: var(--yellow);
  }}

  /* ── LAYOUT ── */
  .container {{ padding: 20px 28px 60px; display: flex; flex-direction: column; gap: 24px; }}

  /* ── STOCK CARD ── */
  .stock-card {{
    background: var(--card);
    border: 1px solid var(--border2);
    border-radius: 14px; overflow: hidden;
  }}
  .card-header {{
    background: var(--surface); border-bottom: 1px solid var(--border);
    padding: 16px 22px; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 12px;
  }}
  .symbol {{ font-size: 22px; font-weight: 800; font-family: 'JetBrains Mono', monospace; }}
  .sector {{ color: var(--muted); font-size: 12px; margin-top: 2px; }}
  .spot-price {{ font-size: 28px; font-weight: 800; font-family: 'JetBrains Mono', monospace; text-align: right; }}
  .spot-change {{ font-size: 13px; text-align: right; margin-top: 2px; }}
  .card-body {{ padding: 18px 22px; display: flex; flex-direction: column; gap: 18px; }}

  /* ── PILL ── */
  .pill {{
    display: inline-block; border-radius: 4px;
    padding: 2px 10px; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase;
  }}

  /* ── STAT GRID ── */
  .stat-grid {{ display: flex; flex-wrap: wrap; gap: 10px; }}
  .stat-box {{
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 8px; padding: 12px 16px; min-width: 110px; flex: 1;
  }}
  .stat-label {{ color: var(--muted); font-size: 10px; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 4px; }}
  .stat-value {{ font-size: 18px; font-weight: 700; font-family: 'JetBrains Mono', monospace; }}
  .stat-sub {{ color: var(--muted); font-size: 11px; margin-top: 2px; }}

  /* ── OI BAR ── */
  .oi-section {{ background: var(--surface); border: 1px solid var(--border); border-radius: 8px; padding: 14px 16px; }}
  .oi-labels {{ display: flex; justify-content: space-between; font-size: 11px; margin-bottom: 6px; font-family: 'JetBrains Mono', monospace; }}
  .oi-bar-track {{ height: 7px; background: #ff525244; border-radius: 4px; overflow: hidden; }}
  .oi-bar-fill {{ height: 100%; background: var(--green); border-radius: 4px; transition: width 0.5s; }}
  .section-title {{ color: var(--muted); font-size: 11px; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 10px; }}

  /* ── GREEKS ── */
  .greeks-row {{ display: flex; flex-wrap: wrap; gap: 10px; }}
  .greek-box {{
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 6px; padding: 8px 14px; text-align: center; min-width: 80px;
  }}
  .greek-label {{ color: var(--muted); font-size: 9px; letter-spacing: 0.5px; margin-bottom: 3px; }}
  .greek-value {{ font-size: 15px; font-weight: 700; font-family: 'JetBrains Mono', monospace; }}

  /* ── S/R ── */
  .sr-row {{ display: flex; flex-wrap: wrap; gap: 8px; }}
  .sr-box {{
    border-radius: 6px; padding: 6px 14px; font-size: 12px; font-family: 'JetBrains Mono', monospace;
  }}

  /* ── OPTION CHAIN TABLE ── */
  .chain-wrap {{ background: var(--surface); border: 1px solid var(--border); border-radius: 8px; overflow-x: auto; padding: 12px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 11px; font-family: 'JetBrains Mono', monospace; }}
  thead tr {{ border-bottom: 1px solid var(--border); }}
  th {{ padding: 6px 8px; color: var(--muted); font-weight: 600; }}
  td {{ padding: 5px 8px; }}
  tr.atm-row {{ background: #ffd74011; }}
  .strike-cell {{
    background: #ffd74011; border-left: 1px solid var(--border); border-right: 1px solid var(--border);
    font-weight: 700; text-align: center;
  }}
  .atm-tag {{ font-size: 9px; color: var(--yellow); margin-left: 4px; }}

  /* ── SPREADS ── */
  .spread-grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
  .spread-card {{ border-radius: 8px; padding: 14px; }}
  .spread-title {{ font-weight: 700; font-size: 12px; letter-spacing: 1px; margin-bottom: 10px; }}
  .spread-fields {{ display: grid; grid-template-columns: 1fr 1fr; gap: 8px; }}
  .spread-field-label {{ color: var(--muted); font-size: 9px; letter-spacing: 0.5px; }}
  .spread-field-value {{ font-family: 'JetBrains Mono', monospace; font-weight: 600; font-size: 13px; }}

  /* ── FOOTER ── */
  footer {{ text-align: center; color: #30363d; font-size: 11px; padding-top: 8px; }}

  @media (max-width: 600px) {{
    .spread-grid {{ grid-template-columns: 1fr; }}
    .topbar {{ flex-direction: column; align-items: flex-start; gap: 8px; }}
  }}
</style>
</head>
<body>

<div class="topbar">
  <div>
    <div class="logo">
      <span class="accent">NSE</span> Options Dashboard
      <span class="badge-live">EOD</span>
    </div>
    <div class="timestamp">Generated: {now} IST</div>
  </div>
  <div class="timestamp">BHARTIARTL · TATAMOTORS · F&amp;O · OI · Greeks · Delivery</div>
</div>

<div class="disclaimer">
  ⚠ End-of-day data from NSE. Not SEBI-registered investment advice. Verify before trading.
</div>

<div class="container">
  {cards}
</div>

<footer>NSE Options Dashboard — Python + NSE public data · {now}</footer>
</body>
</html>"""


def pill_html(label, color):
    return f'<span class="pill" style="background:{color}22;color:{color};border:1px solid {color}44">{label}</span>'


def color_for(val, *, high_good=True, thresholds=(0.8, 1.2)):
    lo, hi = thresholds
    if high_good:
        if val >= hi:   return "#00e676"
        if val <= lo:   return "#ff5252"
        return "#ffd740"
    else:
        if val <= lo:   return "#00e676"
        if val >= hi:   return "#ff5252"
        return "#ffd740"


TREND_COLOR  = {"BULLISH": "#00e676", "BEARISH": "#ff5252", "NEUTRAL": "#ffd740"}
SIGNAL_COLOR = {
    "BUY CE": "#00e676", "BUY PE": "#ff5252",
    "SELL CE": "#ff9100", "SELL PE": "#ea80fc", "WAIT": "#ffd740"
}

def stock_card_html(d):
    tc   = TREND_COLOR.get(d["trend"], "#8b949e")
    sc   = SIGNAL_COLOR.get(d["signal"], "#8b949e")
    chg_c = "#00e676" if d["change"] >= 0 else "#ff5252"
    arrow = "▲" if d["change"] >= 0 else "▼"

    # OI bar
    total_oi = d["totalCallOI"] + d["totalPutOI"]
    call_pct  = round(d["totalCallOI"] / total_oi * 100, 1) if total_oi else 50
    oi_chg_c  = "#00e676" if d["oiChange"] >= 0 else "#ff5252"

    # Greeks
    g = d["greeks"]
    greeks_html = "".join(
        f'<div class="greek-box"><div class="greek-label">{k}</div>'
        f'<div class="greek-value" style="color:{c}">{v}</div></div>'
        for k, v, c in [
            ("Δ Delta", g["delta"],  "#58a6ff"),
            ("Γ Gamma", g["gamma"],  "#bc8cff"),
            ("Θ Theta", g["theta"],  "#ff5252"),
            ("ν Vega",  g["vega"],   "#ffd740"),
            ("IV %",    f'{g["iv"]}%', "#00e676"),
        ]
    )

    # S/R
    sr_html = ""
    for i, s in enumerate(d["support"]):
        sr_html += f'<div class="sr-box" style="background:#00e67611;border:1px solid #00e67633"><span style="color:#8b949e;font-size:10px">S{i+1} </span><span style="color:#00e676;font-weight:700">₹{s:,.1f}</span></div>'
    for i, r in enumerate(d["resistance"]):
        sr_html += f'<div class="sr-box" style="background:#ff525211;border:1px solid #ff525233"><span style="color:#8b949e;font-size:10px">R{i+1} </span><span style="color:#ff5252;font-weight:700">₹{r:,.1f}</span></div>'

    # Strike table
    rows = ""
    for s in d["strikes"]:
        atm_cls  = ' class="atm-row"' if s["isATM"] else ""
        atm_tag  = '<span class="atm-tag">ATM</span>' if s["isATM"] else ""
        c_oi_c   = "#58a6ff" if s["strike"] < d["spotPrice"] else "#e6edf3"
        p_oi_c   = "#58a6ff" if s["strike"] > d["spotPrice"] else "#e6edf3"
        c_chg_c  = "#00e676" if s["callOIChg"] >= 0 else "#ff5252"
        p_chg_c  = "#00e676" if s["putOIChg"]  >= 0 else "#ff5252"
        c_chg_s  = "+" if s["callOIChg"] >= 0 else ""
        p_chg_s  = "+" if s["putOIChg"]  >= 0 else ""
        rows += f"""<tr{atm_cls}>
          <td style="text-align:right;color:{c_oi_c}">{s['callOI']:,}</td>
          <td style="text-align:right;color:{c_chg_c}">{c_chg_s}{s['callOIChg']:,}</td>
          <td style="text-align:right;color:#8b949e">{s['callIV']:.1f}</td>
          <td style="text-align:right;color:#e6edf3;{'font-weight:700' if s['isATM'] else ''}">{s['callLTP']:.2f}</td>
          <td class="strike-cell" style="color:{'#ffd740' if s['isATM'] else '#e6edf3'}">{s['strike']:,}{atm_tag}</td>
          <td style="color:#e6edf3;{'font-weight:700' if s['isATM'] else ''}">{s['putLTP']:.2f}</td>
          <td style="color:#8b949e">{s['putIV']:.1f}</td>
          <td style="color:{p_chg_c}">{p_chg_s}{s['putOIChg']:,}</td>
          <td style="color:{p_oi_c}">{s['putOI']:,}</td>
        </tr>"""

    # Spreads
    spread       = d.get("spread", {})
    bull         = spread.get("bullCall", {})
    bear         = spread.get("bearPut",  {})
    spread_html  = '<div class="spread-grid">'
    if bull:
        spread_html += f"""
        <div class="spread-card" style="background:#0d1117;border:1px solid #00e67633">
          <div class="spread-title" style="color:#00e676">Bull Call Spread</div>
          <div class="spread-fields">
            {"".join(f'<div><div class="spread-field-label">{k}</div><div class="spread-field-value" style="color:#e6edf3">{v}</div></div>'
              for k, v in [("Buy Strike", f"₹{bull.get('buyStrike',0):,}"),
                            ("Sell Strike", f"₹{bull.get('sellStrike',0):,}"),
                            ("Net Premium", f"₹{bull.get('netPremium',0):.2f}"),
                            ("Breakeven", f"₹{bull.get('breakeven',0):.2f}"),
                            ("Max Profit", f"₹{bull.get('maxProfit',0):,.0f}"),
                            ("Max Loss", f"₹{bull.get('maxLoss',0):,.0f}")])}
          </div>
        </div>"""
    if bear:
        spread_html += f"""
        <div class="spread-card" style="background:#0d1117;border:1px solid #ff525233">
          <div class="spread-title" style="color:#ff5252">Bear Put Spread</div>
          <div class="spread-fields">
            {"".join(f'<div><div class="spread-field-label">{k}</div><div class="spread-field-value" style="color:#e6edf3">{v}</div></div>'
              for k, v in [("Buy Strike", f"₹{bear.get('buyStrike',0):,}"),
                            ("Sell Strike", f"₹{bear.get('sellStrike',0):,}"),
                            ("Net Premium", f"₹{bear.get('netPremium',0):.2f}"),
                            ("Breakeven", f"₹{bear.get('breakeven',0):.2f}"),
                            ("Max Profit", f"₹{bear.get('maxProfit',0):,.0f}"),
                            ("Max Loss", f"₹{bear.get('maxLoss',0):,.0f}")])}
          </div>
        </div>"""
    spread_html += "</div>"

    pcr_c   = color_for(d["pcr"], thresholds=(0.8, 1.2))
    iv_c    = "#ff5252" if d["ivPercentile"] > 70 else "#00e676" if d["ivPercentile"] < 30 else "#ffd740"
    del_c   = "#00e676" if d["deliveryPct"] > 50 else "#8b949e"

    return f"""
<div class="stock-card">
  <div class="card-header">
    <div>
      <div style="display:flex;align-items:center;gap:10px">
        <span class="symbol">{d['symbol']}</span>
        <span style="color:#8b949e;font-size:12px">{d['sector']}</span>
        {pill_html(d['trend'], tc)}
      </div>
      <div class="sector">{d['name']} · Lot: {d['lot']} · {d['expiry']} ({d['daysToExpiry']}d to expiry)</div>
    </div>
    <div>
      <div class="spot-price" style="color:{chg_c}">₹{d['spotPrice']:,.2f}</div>
      <div class="spot-change" style="color:{chg_c}">{arrow} ₹{abs(d['change']):.2f} ({abs(d['changePct']):.2f}%)</div>
    </div>
  </div>
  <div class="card-body">

    <!-- STAT GRID -->
    <div class="stat-grid">
      <div class="stat-box">
        <div class="stat-label">Signal</div>
        <div class="stat-value" style="color:{sc}">{d['signal']}</div>
        <div class="stat-sub">Strength: {d['signalStrength']}/10</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">PCR</div>
        <div class="stat-value" style="color:{pcr_c}">{d['pcr']:.2f}</div>
        <div class="stat-sub">Put/Call Ratio</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">IV Percentile</div>
        <div class="stat-value" style="color:{iv_c}">{d['ivPercentile']:.1f}%</div>
        <div class="stat-sub">{'High Vol' if d['ivPercentile']>50 else 'Low Vol'}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Max Pain</div>
        <div class="stat-value" style="color:#bc8cff">₹{d['maxPain']:,}</div>
        <div class="stat-sub">Strike</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">Delivery%</div>
        <div class="stat-value" style="color:{del_c}">{d['deliveryPct']:.1f}%</div>
        <div class="stat-sub">{d['deliveryTrend']}</div>
      </div>
      <div class="stat-box">
        <div class="stat-label">ATM IV</div>
        <div class="stat-value" style="color:#58a6ff">{d['atm_iv']:.1f}%</div>
        <div class="stat-sub">Implied Vol</div>
      </div>
    </div>

    <!-- OI BAR -->
    <div class="oi-section">
      <div style="display:flex;justify-content:space-between;margin-bottom:6px">
        <span style="color:#8b949e;font-size:11px;text-transform:uppercase;letter-spacing:1px">Open Interest Distribution</span>
        <span style="color:#8b949e;font-size:11px">OI Change: <span style="color:{oi_chg_c}">{'+' if d['oiChange']>=0 else ''}{d['oiChange']:,}</span></span>
      </div>
      <div class="oi-labels">
        <span style="color:#00e676">CALL {d['totalCallOI']:,}</span>
        <span style="color:#ff5252">PUT {d['totalPutOI']:,}</span>
      </div>
      <div class="oi-bar-track"><div class="oi-bar-fill" style="width:{call_pct}%"></div></div>
    </div>

    <!-- GREEKS -->
    <div>
      <div class="section-title">ATM Option Greeks</div>
      <div class="greeks-row">{greeks_html}</div>
    </div>

    <!-- S/R -->
    <div>
      <div class="section-title">Support &amp; Resistance (by OI)</div>
      <div class="sr-row">{sr_html}</div>
    </div>

    <!-- OPTION CHAIN -->
    <div>
      <div class="section-title">Option Chain — OI · OI Change · IV · LTP</div>
      <div class="chain-wrap">
        <table>
          <thead>
            <tr>
              <th style="text-align:right">C.OI</th>
              <th style="text-align:right">C.ΔOI</th>
              <th style="text-align:right">C.IV%</th>
              <th style="text-align:right">C.LTP</th>
              <th style="text-align:center;color:#ffd740">STRIKE</th>
              <th>P.LTP</th><th>P.IV%</th><th>P.ΔOI</th><th>P.OI</th>
            </tr>
          </thead>
          <tbody>{rows}</tbody>
        </table>
      </div>
    </div>

    <!-- SPREADS -->
    <div>
      <div class="section-title">Recommended Option Spreads</div>
      {spread_html}
    </div>

  </div>
</div>"""


# ─────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────
def main():
    print("\n" + "═"*50)
    print("  NSE Options Dashboard Generator")
    print("  Run after 3:45 PM IST on trading days")
    print("═"*50 + "\n")

    session  = get_nse_session()
    all_data = []

    for symbol in STOCKS:
        print(f"\n📊 Fetching {symbol} ({STOCKS[symbol]['name']})...")
        oc_data    = fetch_option_chain(session, symbol)
        time.sleep(1.5)
        quote_data = fetch_equity_quote(session, symbol)
        time.sleep(1)
        parsed     = parse_data(oc_data, quote_data, symbol)
        if parsed:
            print(f"  ✓ Spot: ₹{parsed['spotPrice']:,}  PCR: {parsed['pcr']}  Signal: {parsed['signal']}")
            all_data.append(parsed)
        else:
            print(f"  ✗ Failed for {symbol}")

    if not all_data:
        print("\n⚠  No data fetched. NSE may be closed or blocking requests.")
        print("   Try running after 3:45 PM IST on a weekday.\n")
        return

    print(f"\n🖥  Generating dashboard...")
    html     = generate_html(all_data)
    out_path = Path(OUTPUT_FILE)
    out_path.write_text(html, encoding="utf-8")

    print(f"  ✓ Saved: {out_path.resolve()}")
    print(f"\n🌐 Opening in browser...")
    webbrowser.open(f"file://{out_path.resolve()}")
    print("\n✅ Done!\n")


if __name__ == "__main__":
    main()
