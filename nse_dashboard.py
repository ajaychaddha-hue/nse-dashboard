"""
NSE Options Dashboard Generator v2
====================================
- Full error messages with reasons
- Log file saved after every run
- Handles NSE blocks, timeouts, empty data
Run after 3:45 PM IST on trading days.
"""

import requests, json, os, math, time, webbrowser, traceback
from datetime import datetime, date
from pathlib import Path

STOCKS = {
    "BHARTIARTL": {"name": "Bharti Airtel", "sector": "Telecom", "lot": 475},
    "TATAMOTORS":  {"name": "Tata Motors",   "sector": "Auto",    "lot": 550},
}
OUTPUT_FILE = "nse_options_dashboard.html"
LOG_FILE    = "nse_run_log.txt"

# ── LOGGER ──────────────────────────────────
log_lines = []
def log(msg, level="INFO"):
    ts   = datetime.now().strftime("%H:%M:%S")
    icon = {"OK":"✅","WARN":"⚠️ ","ERROR":"❌","STEP":"▶ "}.get(level,"ℹ ")
    line = f"[{ts}] {icon} {msg}"
    print(line); log_lines.append(line)

def save_log():
    Path(LOG_FILE).write_text("\n".join(log_lines), encoding="utf-8")
    log(f"Log saved → {Path(LOG_FILE).resolve()}")

# ── NSE SESSION ──────────────────────────────
def get_nse_session():
    log("Setting up NSE session...", "STEP")
    s = requests.Session()
    s.headers.update({
        "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                       "AppleWebKit/537.36 (KHTML, like Gecko) "
                       "Chrome/124.0.0.0 Safari/537.36"),
        "Accept":          "application/json, text/plain, */*",
        "Accept-Language": "en-US,en;q=0.9",
        "Referer":         "https://www.nseindia.com/",
        "Origin":          "https://www.nseindia.com",
    })
    try:
        log("Loading NSE homepage (priming cookies)...", "STEP")
        r1 = s.get("https://www.nseindia.com", timeout=15)
        log(f"Homepage → HTTP {r1.status_code}", "OK" if r1.status_code==200 else "WARN")
        time.sleep(2)
        r2 = s.get("https://www.nseindia.com/option-chain", timeout=15)
        log(f"Option chain page → HTTP {r2.status_code}", "OK" if r2.status_code==200 else "WARN")
        time.sleep(2)
        log("Session ready ✓", "OK")
        return s
    except requests.exceptions.ConnectionError:
        log("NETWORK ERROR — Cannot reach nseindia.com", "ERROR")
        log("  Fix: Check your internet / turn off VPN", "ERROR")
    except requests.exceptions.Timeout:
        log("TIMEOUT — NSE took too long. Try again in 5 min.", "ERROR")
    except Exception as e:
        log(f"Session error: {e}", "ERROR")
        log(traceback.format_exc(), "ERROR")
    return None

# ── FETCH OPTION CHAIN ───────────────────────
def fetch_option_chain(session, symbol):
    url = f"https://www.nseindia.com/api/option-chain-equities?symbol={symbol}"
    log(f"Fetching option chain: {symbol}", "STEP")
    try:
        r = session.get(url, timeout=20)
        log(f"HTTP {r.status_code} for {symbol}")
        if r.status_code == 401:
            log("Session expired (401). Retry after restarting script.", "ERROR"); return None
        if r.status_code == 403:
            log("NSE blocked request (403 Forbidden)", "ERROR")
            log("  Fix 1: Wait 10–15 min and retry", "WARN")
            log("  Fix 2: Don't run during market hours (9:15–3:30 PM IST)", "WARN")
            log("  Fix 3: Try from a different network/hotspot", "WARN")
            return None
        if r.status_code != 200:
            log(f"Unexpected HTTP {r.status_code} for {symbol}", "ERROR"); return None
        data = r.json()
        if "records" not in data:
            log(f"Unexpected response format — 'records' key missing", "ERROR")
            log(f"Response preview: {str(data)[:300]}", "ERROR"); return None
        n = len(data["records"].get("data",[]))
        log(f"Got {n} records for {symbol}", "OK")
        return data
    except requests.exceptions.Timeout:
        log(f"Timeout fetching {symbol}. NSE is slow — retry in a few min.", "ERROR")
    except json.JSONDecodeError:
        log(f"JSON parse failed for {symbol} — NSE may be in maintenance.", "ERROR")
    except Exception as e:
        log(f"Error fetching {symbol}: {e}", "ERROR")
        log(traceback.format_exc(), "ERROR")
    return None

# ── FETCH QUOTE (delivery%) ──────────────────
def fetch_equity_quote(session, symbol):
    url = f"https://www.nseindia.com/api/quote-equity?symbol={symbol}&section=trade_info"
    try:
        r = session.get(url, timeout=15)
        if r.status_code == 200:
            log(f"Quote data OK for {symbol}", "OK"); return r.json()
        log(f"Quote HTTP {r.status_code} for {symbol} — delivery% will be 0", "WARN")
    except Exception as e:
        log(f"Quote fetch failed for {symbol} (non-critical): {e}", "WARN")
    return None

# ── PARSE ────────────────────────────────────
def parse_data(oc_data, quote_data, symbol):
    if not oc_data: return None
    log(f"Parsing {symbol}...", "STEP")
    records  = oc_data.get("records", {})
    spot     = records.get("underlyingValue", 0)
    if spot == 0: log(f"Spot price is 0 for {symbol} — check data", "WARN")

    delivery_pct, delivery_trend = 0, "NEUTRAL"
    if quote_data:
        try:
            dp = float(quote_data.get("securityWiseDP",{}).get("deliveryToTradedQuantity",0))
            delivery_pct   = dp
            delivery_trend = "ACCUMULATION" if dp>55 else ("DISTRIBUTION" if dp<35 else "NEUTRAL")
        except: pass

    expiry_dates = records.get("expiryDates", [])
    expiry       = expiry_dates[0] if expiry_dates else "N/A"
    try:
        exp_dt      = datetime.strptime(expiry, "%d-%b-%Y")
        days_to_exp = max(1, (exp_dt.date() - date.today()).days)
        expiry_type = "weekly" if days_to_exp <= 7 else "monthly"
    except:
        days_to_exp, expiry_type = 0, "monthly"

    strike_map = {}
    tc_oi = tp_oi = tc_chg = tp_chg = 0
    for item in records.get("data", []):
        if item.get("expiryDate") != expiry: continue
        s = item["strikePrice"]
        if s not in strike_map: strike_map[s] = {}
        if "CE" in item:
            strike_map[s]["CE"] = item["CE"]
            tc_oi  += item["CE"].get("openInterest", 0)
            tc_chg += item["CE"].get("changeinOpenInterest", 0)
        if "PE" in item:
            strike_map[s]["PE"] = item["PE"]
            tp_oi  += item["PE"].get("openInterest", 0)
            tp_chg += item["PE"].get("changeinOpenInterest", 0)

    if not strike_map:
        log(f"No strikes found for {symbol} expiry {expiry}", "ERROR"); return None

    pcr         = round(tp_oi/tc_oi, 2) if tc_oi else 1.0
    all_strikes = sorted(strike_map.keys())
    atm         = min(all_strikes, key=lambda x: abs(x-spot))
    atm_idx     = all_strikes.index(atm)
    sel         = all_strikes[max(0,atm_idx-3): min(len(all_strikes),atm_idx+4)]

    strikes_out = []
    atm_iv = atm_ce = atm_pe = 0
    for s in sel:
        ce, pe = strike_map.get(s,{}).get("CE",{}), strike_map.get(s,{}).get("PE",{})
        is_atm = (s == atm)
        if is_atm: atm_iv=ce.get("impliedVolatility",0); atm_ce=ce.get("lastPrice",0); atm_pe=pe.get("lastPrice",0)
        strikes_out.append({"strike":s,"isATM":is_atm,
            "callOI":ce.get("openInterest",0),"callOIChg":ce.get("changeinOpenInterest",0),
            "callIV":ce.get("impliedVolatility",0),"callLTP":ce.get("lastPrice",0),
            "putOI":pe.get("openInterest",0),"putOIChg":pe.get("changeinOpenInterest",0),
            "putIV":pe.get("impliedVolatility",0),"putLTP":pe.get("lastPrice",0)})

    mp     = calc_max_pain(strike_map, all_strikes)
    iv_pct = min(99, max(1, round((atm_iv/50)*100))) if atm_iv else 50
    greeks = calc_greeks(spot, atm, atm_iv/100, days_to_exp/365)
    sp_put  = sorted(strike_map.items(), key=lambda x: x[1].get("PE",{}).get("openInterest",0), reverse=True)
    sp_call = sorted(strike_map.items(), key=lambda x: x[1].get("CE",{}).get("openInterest",0), reverse=True)
    support    = sorted([s for s,_ in sp_put[:2]  if s < spot])
    resistance = sorted([s for s,_ in sp_call[:2] if s > spot])
    trend, signal, strength = derive_signal(pcr, iv_pct, delivery_pct, tc_chg, tp_chg, spot, mp)
    spread = calc_spreads(atm, all_strikes, strike_map, atm_ce, atm_pe, STOCKS[symbol]["lot"])
    log(f"✓ {symbol}: Spot=₹{spot:,}  PCR={pcr}  IV={atm_iv}%  Signal={signal}", "OK")
    return {"symbol":symbol,"name":STOCKS[symbol]["name"],"sector":STOCKS[symbol]["sector"],
            "lot":STOCKS[symbol]["lot"],"spotPrice":spot,"expiry":expiry,
            "expiry_type":expiry_type,"daysToExpiry":days_to_exp,
            "deliveryPct":round(delivery_pct,1),"deliveryTrend":delivery_trend,
            "pcr":pcr,"ivPercentile":iv_pct,"maxPain":mp,"trend":trend,
            "signal":signal,"signalStrength":strength,
            "totalCallOI":tc_oi,"totalPutOI":tp_oi,"oiChange":tc_chg-tp_chg,
            "strikes":strikes_out,"greeks":greeks,"support":support[:2],
            "resistance":resistance[:2],"spread":spread,"atm_iv":atm_iv}

def calc_max_pain(sm, strikes):
    best, best_s = float("inf"), strikes[len(strikes)//2]
    for c in strikes:
        pain = sum(v.get("CE",{}).get("openInterest",0)*max(0,c-s) +
                   v.get("PE",{}).get("openInterest",0)*max(0,s-c) for s,v in sm.items())
        if pain < best: best, best_s = pain, c
    return best_s

def calc_greeks(spot, strike, iv, T):
    if T<=0 or iv<=0: return {"delta":0.5,"gamma":0,"theta":0,"vega":0,"iv":round(iv*100,1)}
    r=0.065; erf=math.erf
    d1=(math.log(spot/strike)+(r+.5*iv**2)*T)/(iv*math.sqrt(T)); d2=d1-iv*math.sqrt(T)
    npdf=math.exp(-.5*d1**2)/math.sqrt(2*math.pi); ncdf=lambda x:.5*(1+erf(x/math.sqrt(2)))
    return {"delta":round(ncdf(d1),4),"gamma":round(npdf/(spot*iv*math.sqrt(T)),6),
            "theta":round((-spot*npdf*iv/(2*math.sqrt(T))-r*strike*math.exp(-r*T)*ncdf(d2))/365,4),
            "vega":round(spot*npdf*math.sqrt(T)/100,4),"iv":round(iv*100,1)}

def calc_spreads(atm, strikes, sm, ce_ltp, pe_ltp, lot):
    idx = strikes.index(atm) if atm in strikes else len(strikes)//2
    sp  = {}
    if idx+1 < len(strikes):
        sell=strikes[idx+1]; s_ltp=sm.get(sell,{}).get("CE",{}).get("lastPrice",0)
        net=round(ce_ltp-s_ltp,2)
        sp["bullCall"]={"buyStrike":atm,"sellStrike":sell,"netPremium":net,
                        "breakeven":round(atm+net,2),"maxProfit":round((sell-atm-net)*lot,2),"maxLoss":round(net*lot,2)}
    if idx-1 >= 0:
        sell=strikes[idx-1]; s_ltp=sm.get(sell,{}).get("PE",{}).get("lastPrice",0)
        net=round(pe_ltp-s_ltp,2)
        sp["bearPut"]={"buyStrike":atm,"sellStrike":sell,"netPremium":net,
                       "breakeven":round(atm-net,2),"maxProfit":round((atm-sell-net)*lot,2),"maxLoss":round(net*lot,2)}
    return sp

def derive_signal(pcr, iv_pct, del_pct, cc, pc, spot, mp):
    sc  = (2 if pcr>1.3 else -2 if pcr<0.7 else 1 if pcr>1 else -1)
    sc += (1 if pc>cc else -1)
    sc += (1 if del_pct>55 else -1 if del_pct<35 else 0)
    sc += (1 if spot>mp else -1 if spot<mp else 0)
    if sc>=3:  return "BULLISH","BUY CE",min(10,5+sc)
    if sc<=-3: return "BEARISH","BUY PE",min(10,5+abs(sc))
    if sc>0:   return "BULLISH","BUY CE",4+sc
    if sc<0:   return "BEARISH","BUY PE",4+abs(sc)
    return "NEUTRAL","WAIT",3

# ── HTML ─────────────────────────────────────
TC={"BULLISH":"#00e676","BEARISH":"#ff5252","NEUTRAL":"#ffd740"}
SC={"BUY CE":"#00e676","BUY PE":"#ff5252","SELL CE":"#ff9100","SELL PE":"#ea80fc","WAIT":"#ffd740"}

def pill(label, color):
    return (f'<span style="background:{color}22;color:{color};border:1px solid {color}44;'
            f'border-radius:4px;padding:2px 9px;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase">{label}</span>')

def card_html(d):
    tc=TC.get(d["trend"],"#8b949e"); sc=SC.get(d["signal"],"#8b949e")
    tot=d["totalCallOI"]+d["totalPutOI"]; cpct=round(d["totalCallOI"]/tot*100,1) if tot else 50
    occ="#00e676" if d["oiChange"]>=0 else "#ff5252"
    g=d["greeks"]
    pcrc="#00e676" if d["pcr"]>1.2 else "#ff5252" if d["pcr"]<0.8 else "#ffd740"
    ivc="#ff5252" if d["ivPercentile"]>70 else "#00e676" if d["ivPercentile"]<30 else "#ffd740"
    greeks_html="".join(
        f'<div class="gb"><div class="gl">{k}</div><div class="gv" style="color:{c}">{v}</div></div>'
        for k,v,c in [("Δ Delta",g["delta"],"#58a6ff"),("Γ Gamma",g["gamma"],"#bc8cff"),
                      ("Θ Theta",g["theta"],"#ff5252"),("ν Vega",g["vega"],"#ffd740"),("IV%",f'{g["iv"]}%',"#00e676")])
    sr="".join(f'<span class="srb" style="background:#00e67611;border:1px solid #00e67633">'
               f'<span style="color:#8b949e;font-size:10px">S{i+1} </span>'
               f'<span style="color:#00e676;font-weight:700;font-family:\'JetBrains Mono\',monospace">₹{s:,.1f}</span></span>'
               for i,s in enumerate(d["support"]))
    sr+="".join(f'<span class="srb" style="background:#ff525211;border:1px solid #ff525233">'
                f'<span style="color:#8b949e;font-size:10px">R{i+1} </span>'
                f'<span style="color:#ff5252;font-weight:700;font-family:\'JetBrains Mono\',monospace">₹{r:,.1f}</span></span>'
                for i,r in enumerate(d["resistance"]))
    rows=""
    for s in d["strikes"]:
        ac=' class="arow"' if s["isATM"] else ""
        at='<span style="font-size:9px;color:#ffd740;margin-left:3px">ATM</span>' if s["isATM"] else ""
        coc="#58a6ff" if s["strike"]<d["spotPrice"] else "#e6edf3"
        poc="#58a6ff" if s["strike"]>d["spotPrice"] else "#e6edf3"
        cc2="#00e676" if s["callOIChg"]>=0 else "#ff5252"
        pc2="#00e676" if s["putOIChg"]>=0 else "#ff5252"
        fw="font-weight:700;" if s["isATM"] else ""
        rows+=(f'<tr{ac}>'
               f'<td style="text-align:right;color:{coc}">{s["callOI"]:,}</td>'
               f'<td style="text-align:right;color:{cc2}">{"+" if s["callOIChg"]>=0 else ""}{s["callOIChg"]:,}</td>'
               f'<td style="text-align:right;color:#8b949e">{s["callIV"]:.1f}</td>'
               f'<td style="text-align:right;{fw}">{s["callLTP"]:.2f}</td>'
               f'<td class="sc" style="color:{"#ffd740" if s["isATM"] else "#e6edf3"}">{s["strike"]:,}{at}</td>'
               f'<td style="{fw}">{s["putLTP"]:.2f}</td>'
               f'<td style="color:#8b949e">{s["putIV"]:.1f}</td>'
               f'<td style="color:{pc2}">{"+" if s["putOIChg"]>=0 else ""}{s["putOIChg"]:,}</td>'
               f'<td style="color:{poc}">{s["putOI"]:,}</td></tr>')
    sp=d.get("spread",{}); bull=sp.get("bullCall",{}); bear=sp.get("bearPut",{})
    def scard(title,data,color):
        if not data: return ""
        flds=[("Buy Strike",f'₹{data.get("buyStrike",0):,}'),("Sell Strike",f'₹{data.get("sellStrike",0):,}'),
              ("Net Premium",f'₹{data.get("netPremium",0):.2f}'),("Breakeven",f'₹{data.get("breakeven",0):.2f}'),
              ("Max Profit",f'₹{data.get("maxProfit",0):,.0f}'),("Max Loss",f'₹{data.get("maxLoss",0):,.0f}')]
        inner="".join(f'<div><div style="color:#8b949e;font-size:9px">{k}</div>'
                      f'<div style="font-family:\'JetBrains Mono\',monospace;font-weight:600;font-size:13px">{v}</div></div>'
                      for k,v in flds)
        return (f'<div style="background:#0d1117;border:1px solid {color}33;border-radius:8px;padding:14px">'
                f'<div style="color:{color};font-weight:700;font-size:12px;letter-spacing:1px;margin-bottom:10px">{title}</div>'
                f'<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">{inner}</div></div>')
    return f"""<div class="card">
  <div class="ch">
    <div>
      <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap">
        <span style="font-size:20px;font-weight:800;font-family:'JetBrains Mono',monospace">{d["symbol"]}</span>
        <span style="color:#8b949e;font-size:12px">{d["sector"]}</span>
        {pill(d["trend"],tc)} {pill(d["signal"],sc)}
      </div>
      <div style="color:#8b949e;font-size:12px;margin-top:4px">{d["name"]} · Lot {d["lot"]} · {d["expiry"]} ({d["daysToExpiry"]}d)</div>
    </div>
    <div style="text-align:right">
      <div style="font-size:24px;font-weight:800;font-family:'JetBrains Mono',monospace">₹{d["spotPrice"]:,.2f}</div>
      <div style="color:#8b949e;font-size:11px;margin-top:2px">Signal strength: {d["signalStrength"]}/10</div>
    </div>
  </div>
  <div class="cb">
    <div class="sg">
      <div class="sb"><div class="sl">PCR</div><div class="sv" style="color:{pcrc}">{d["pcr"]:.2f}</div><div class="ss">Put/Call Ratio</div></div>
      <div class="sb"><div class="sl">IV Percentile</div><div class="sv" style="color:{ivc}">{d["ivPercentile"]:.1f}%</div><div class="ss">{'High Vol' if d["ivPercentile"]>50 else 'Low Vol'}</div></div>
      <div class="sb"><div class="sl">Max Pain</div><div class="sv" style="color:#bc8cff">₹{d["maxPain"]:,}</div><div class="ss">Strike</div></div>
      <div class="sb"><div class="sl">Delivery%</div><div class="sv" style="color:{'#00e676' if d['deliveryPct']>50 else '#8b949e'}">{d["deliveryPct"]:.1f}%</div><div class="ss">{d["deliveryTrend"]}</div></div>
      <div class="sb"><div class="sl">ATM IV</div><div class="sv" style="color:#58a6ff">{d["atm_iv"]:.1f}%</div><div class="ss">Implied Vol</div></div>
      <div class="sb"><div class="sl">OI Change</div><div class="sv" style="color:{occ}">{"+" if d["oiChange"]>=0 else ""}{d["oiChange"]:,}</div><div class="ss">Net shift</div></div>
    </div>
    <div class="ois">
      <div style="display:flex;justify-content:space-between;margin-bottom:5px">
        <span style="color:#8b949e;font-size:11px;text-transform:uppercase;letter-spacing:1px">Open Interest Split</span>
      </div>
      <div style="display:flex;justify-content:space-between;font-size:11px;font-family:'JetBrains Mono',monospace;margin-bottom:5px">
        <span style="color:#00e676">CALL {d["totalCallOI"]:,}</span>
        <span style="color:#ff5252">PUT {d["totalPutOI"]:,}</span>
      </div>
      <div style="height:7px;background:#ff525244;border-radius:4px;overflow:hidden">
        <div style="height:100%;width:{cpct}%;background:#00e676;border-radius:4px"></div>
      </div>
    </div>
    <div><div class="st">ATM Greeks</div><div style="display:flex;flex-wrap:wrap;gap:10px">{greeks_html}</div></div>
    <div><div class="st">Support &amp; Resistance (by OI)</div><div style="display:flex;flex-wrap:wrap;gap:8px">{sr}</div></div>
    <div>
      <div class="st">Option Chain</div>
      <div style="background:#0d1117;border:1px solid #21262d;border-radius:8px;overflow-x:auto;padding:10px">
        <table><thead><tr>
          <th style="text-align:right">C.OI</th><th style="text-align:right">C.ΔOI</th>
          <th style="text-align:right">C.IV%</th><th style="text-align:right">C.LTP</th>
          <th style="text-align:center;color:#ffd740">STRIKE</th>
          <th>P.LTP</th><th>P.IV%</th><th>P.ΔOI</th><th>P.OI</th>
        </tr></thead><tbody>{rows}</tbody></table>
      </div>
    </div>
    <div>
      <div class="st">Option Spreads</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
        {scard("Bull Call Spread",bull,"#00e676")}
        {scard("Bear Put Spread",bear,"#ff5252")}
      </div>
    </div>
  </div>
</div>"""

def generate_html(all_data):
    now=datetime.now().strftime("%d %b %Y, %I:%M %p")
    cards="".join(card_html(d) for d in all_data if d)
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>NSE Options — {now}</title>
<link href="https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=JetBrains+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{background:#010409;color:#e6edf3;font-family:Syne,sans-serif}}
::-webkit-scrollbar{{width:5px;height:5px}}::-webkit-scrollbar-thumb{{background:#30363d;border-radius:3px}}
.topbar{{background:#0d1117;border-bottom:1px solid #21262d;padding:14px 20px;
  display:flex;justify-content:space-between;align-items:center;
  position:sticky;top:0;z-index:100;flex-wrap:wrap;gap:8px}}
.logo{{font-weight:800;font-size:18px}}
.ts{{color:#8b949e;font-size:11px;font-family:'JetBrains Mono',monospace}}
.disc{{background:#ffd74011;border:1px solid #ffd74033;margin:14px 20px 0;
  border-radius:8px;padding:8px 14px;font-size:11px;color:#ffd740}}
.wrap{{padding:16px 20px 60px;display:flex;flex-direction:column;gap:20px}}
.card{{background:#161b22;border:1px solid #30363d;border-radius:12px;overflow:hidden}}
.ch{{background:#0d1117;border-bottom:1px solid #21262d;padding:14px 20px;
  display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px}}
.cb{{padding:16px 20px;display:flex;flex-direction:column;gap:16px}}
.sg{{display:flex;flex-wrap:wrap;gap:10px}}
.sb{{background:#0d1117;border:1px solid #21262d;border-radius:8px;padding:10px 14px;min-width:100px;flex:1}}
.sl{{color:#8b949e;font-size:10px;letter-spacing:1px;text-transform:uppercase;margin-bottom:3px}}
.sv{{font-size:17px;font-weight:700;font-family:'JetBrains Mono',monospace}}
.ss{{color:#8b949e;font-size:11px;margin-top:2px}}
.ois{{background:#0d1117;border:1px solid #21262d;border-radius:8px;padding:12px 14px}}
.st{{color:#8b949e;font-size:11px;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px}}
.gb{{background:#0d1117;border:1px solid #21262d;border-radius:6px;padding:8px 12px;text-align:center;min-width:70px}}
.gl{{color:#8b949e;font-size:9px;letter-spacing:.5px;margin-bottom:3px}}
.gv{{font-size:15px;font-weight:700;font-family:'JetBrains Mono',monospace}}
.srb{{border-radius:6px;padding:6px 12px;font-size:12px}}
table{{width:100%;border-collapse:collapse;font-size:11px;font-family:'JetBrains Mono',monospace}}
thead tr{{border-bottom:1px solid #21262d}}
th{{padding:6px 7px;color:#8b949e;font-weight:600}}
td{{padding:5px 7px}}
tr.arow{{background:#ffd74011}}
.sc{{background:#ffd74011;border-left:1px solid #21262d;border-right:1px solid #21262d;font-weight:700;text-align:center;padding:5px 10px}}
footer{{text-align:center;color:#30363d;font-size:11px;padding:8px}}
@media(max-width:600px){{
  .sb{{min-width:calc(50% - 5px)}}
  div[style*="grid-template-columns:1fr 1fr"]{{grid-template-columns:1fr!important}}
}}
</style>
</head>
<body>
<div class="topbar">
  <div><div class="logo"><span style="color:#58a6ff">NSE</span> Options Dashboard</div>
  <div class="ts">BHARTIARTL · TATAMOTORS · EOD Data</div></div>
  <div class="ts">{now} IST</div>
</div>
<div class="disc">⚠ EOD data from NSE India. Not SEBI-registered investment advice. Verify before trading.</div>
<div class="wrap">{cards}</div>
<footer>NSE Options Dashboard · {now}</footer>
</body></html>"""

# ── MAIN ─────────────────────────────────────
def main():
    print("\n" + "═"*52)
    print("   NSE Options Dashboard  v2")
    print("   Run after 3:45 PM IST, Mon–Fri")
    print("═"*52 + "\n")

    session = get_nse_session()
    if not session:
        save_log(); input("\nPress Enter to close..."); return

    all_data = []
    for symbol in STOCKS:
        log(f"\n── {symbol} ──", "STEP")
        oc    = fetch_option_chain(session, symbol); time.sleep(2)
        quote = fetch_equity_quote(session, symbol); time.sleep(1)
        data  = parse_data(oc, quote, symbol)
        if data: all_data.append(data)
        else: log(f"Skipping {symbol} — no usable data", "WARN")

    if not all_data:
        log("\nNo data fetched. Possible reasons:", "ERROR")
        log("  1. Market closed (weekend / holiday)", "WARN")
        log("  2. NSE blocking requests (run after 3:45 PM IST)", "WARN")
        log("  3. Internet or firewall issue", "WARN")
        log("  4. Try from mobile hotspot if office network blocks NSE", "WARN")
        save_log(); input("\nPress Enter to close..."); return

    log("\nGenerating dashboard...", "STEP")
    html = generate_html(all_data)
    Path(OUTPUT_FILE).write_text(html, encoding="utf-8")
    log(f"Saved: {Path(OUTPUT_FILE).resolve()}", "OK")
    save_log()
    webbrowser.open(f"file://{Path(OUTPUT_FILE).resolve()}")
    print("\n" + "═"*52)
    print("   ✅  Done! Dashboard opened in browser.")
    print(f"   📄  Log: {Path(LOG_FILE).resolve()}")
    print("═"*52 + "\n")
    input("Press Enter to close...")

if __name__ == "__main__":
    try: main()
    except KeyboardInterrupt: print("\nCancelled.")
    except Exception as e:
        log(f"FATAL: {e}", "ERROR"); log(traceback.format_exc(), "ERROR")
        save_log(); input("\nPress Enter to close...")
