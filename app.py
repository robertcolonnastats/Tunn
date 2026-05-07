"""
Tunneling+ Streamlit App — v18
Built on the V18 pipeline (Statcast Edition).
Identical model to V17. hb sign is hand-dependent (LHP fix applied).
No windup filter. Statcast-native release_side.
"""

import streamlit as st
import pandas as pd
import numpy as np
import json
import os
import sys
import asyncio
import subprocess
import tempfile
from datetime import date, datetime, timezone, timedelta
from itertools import combinations

sys.path.insert(0, os.path.dirname(__file__))

# ── Set Playwright browser path before any playwright import ──────────────────
# Streamlit Cloud wipes ~/.cache on each cold boot. We redirect the browser
# binary to /tmp which persists within a session and is writable by appuser.
import tempfile as _tmpfile
_BROWSERS_PATH = os.path.join(_tmpfile.gettempdir(), 'ms-playwright-browsers')
os.environ['PLAYWRIGHT_BROWSERS_PATH'] = _BROWSERS_PATH
os.makedirs(_BROWSERS_PATH, exist_ok=True)

# ── Import V18 pipeline ───────────────────────────────────────────────────────
try:
    from tunneling_plus_v18_pipeline_2 import (
        load_statcast, run_model, normalize, build_excel,
        TUNNEL_FT, TD_TUNNEL, MIN_RATIO, PD_SPEED, VD_SPEED, LG_AVG_PD
    )
except ImportError as _e:
    st.error(
        f"Missing dependency: {_e}\n\n"
        "Run: `pip install pybaseball pandas numpy openpyxl streamlit`"
    )
    st.stop()

# ── Card constants ────────────────────────────────────────────────────────────
PITCH_COLORS = {
    'FF': '#378ADD', 'SI': '#EF9F27', 'FC': '#534AB7', 'FS': '#534AB7',
    'ST': '#E24B4A', 'SL': '#E24B4A', 'CU': '#D85A30', 'KC': '#D85A30',
    'CH': '#1D9E75', 'SC': '#1D9E75', 'KN': '#888888', 'EP': '#888888',
}
def get_color(pt): return PITCH_COLORS.get(pt, '#888888')

TEAM_FULL = {
    'NYM':'New York Mets','MIL':'Milwaukee Brewers','PIT':'Pittsburgh Pirates',
    'ATL':'Atlanta Braves','MIN':'Minnesota Twins','LAA':'Los Angeles Angels',
    'TEX':'Texas Rangers','NYY':'New York Yankees','BOS':'Boston Red Sox',
    'LAD':'Los Angeles Dodgers','SF':'San Francisco Giants','CHC':'Chicago Cubs',
    'CWS':'Chicago White Sox','CLE':'Cleveland Guardians','DET':'Detroit Tigers',
    'HOU':'Houston Astros','KC':'Kansas City Royals','OAK':'Oakland Athletics',
    'SEA':'Seattle Mariners','TB':'Tampa Bay Rays','TOR':'Toronto Blue Jays',
    'BAL':'Baltimore Orioles','WAS':'Washington Nationals','PHI':'Philadelphia Phillies',
    'MIA':'Miami Marlins','ARI':'Arizona Diamondbacks','COL':'Colorado Rockies',
    'SD':'San Diego Padres','STL':'St. Louis Cardinals','CIN':'Cincinnati Reds',
}
HAND_STR = {'R': 'RHP', 'L': 'LHP'}

CARD_CSS = """
*{box-sizing:border-box;margin:0;padding:0}
body{background:#f0f0f0;padding:20px;font-family:Arial,sans-serif}
.card{background:#fff;border:0.5px solid #ddd;border-radius:12px;overflow:hidden;width:680px}
.hdr{background:#1a2744;padding:1.1rem 1.4rem}
.hdr-top{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:.6rem}
.hdr-name{font-size:22px;font-weight:500;color:#fff;line-height:1.1}
.hdr-sub{font-size:12px;color:rgba(255,255,255,0.5);margin-top:3px}
.hdr-pct{background:rgba(255,255,255,0.12);border-radius:10px;padding:.35rem .85rem;text-align:center;flex-shrink:0}
.hdr-pct-num{font-size:22px;font-weight:500;color:#fff;line-height:1}
.hdr-pct-lbl{font-size:11px;color:rgba(255,255,255,0.45);margin-top:1px}
.hdr-center{text-align:center}
.tp-num{font-size:40px;font-weight:500;color:#60a8f0;line-height:1}
.tp-lbl{font-size:11px;color:rgba(255,255,255,0.45);text-transform:uppercase;letter-spacing:.07em;margin-top:2px}
.tp-rank{font-size:12px;color:rgba(255,255,255,0.45);margin-top:3px}
.hdr-watermark{text-align:right;font-size:10px;color:rgba(255,255,255,0.4);margin-top:6px;font-style:italic}
.body{padding:1.1rem 1.4rem;background:#fff}
.sec{font-size:12px;font-weight:500;text-transform:uppercase;letter-spacing:.07em;color:#888;margin:1rem 0 .55rem}
.sec:first-child{margin-top:0}
.divider{height:0.5px;background:#eee;margin:.9rem 0}
.top-stats{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}
.stat-card{background:#f7f7f5;border-radius:8px;padding:.7rem .9rem}
.stat-pct{font-size:28px;font-weight:500;line-height:1}
.stat-lbl{font-size:13px;color:#333;margin-top:4px;font-weight:500}
.stat-raw{font-size:12px;color:#777;margin-top:3px}
.stat-sub{font-size:11px;color:#999;margin-top:4px;line-height:1.4}
.arsenal-wrap{display:flex;flex-wrap:wrap;gap:7px;justify-content:center}
.ap{display:flex;align-items:center;gap:6px;background:#f7f7f5;border:0.5px solid #e0e0e0;border-radius:20px;padding:5px 12px 5px 9px;font-size:12px}
.ap-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0}
.ap-name{font-weight:500;color:#1a1a1a}
.ap-val{color:#888}
.plots-wrap{border:0.5px solid #ddd;border-radius:8px;overflow:hidden;margin-bottom:8px}
.plots-row{display:grid;grid-template-columns:1fr 1fr}
.plot-block{background:#f7f7f5;padding:8px 8px 4px}
.plot-block:first-child{border-right:0.5px solid #ddd}
.plot-title{font-size:11px;font-weight:500;text-transform:uppercase;letter-spacing:.05em;color:#888;text-align:center;margin-bottom:4px}
.leg-row{display:flex;flex-wrap:wrap;gap:10px;justify-content:center;padding:.5rem 0 .15rem}
.li{display:flex;align-items:center;gap:4px;font-size:11px;color:#666}
.li-dot{width:9px;height:9px;border-radius:50%;flex-shrink:0}
.pairs-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px}
.pairs-grid-2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.pair-card{background:#f7f7f5;border:0.5px solid #e0e0e0;border-radius:8px;padding:.65rem .85rem}
.pair-names{font-size:14px;font-weight:500;margin-bottom:3px}
.pair-sep{color:#ccc;font-weight:400;font-size:13px}
.pair-meta{font-size:11px;color:#999;margin-bottom:7px;line-height:1.5}
.bar-row{display:flex;align-items:center;gap:8px}
.bar-track{flex:1;height:6px;background:#e0e0e0;border-radius:3px;overflow:hidden}
.bar-fill{height:6px;border-radius:3px}
.bar-pct{font-size:20px;font-weight:500;min-width:44px;text-align:right;line-height:1}
.pair-ratio{font-size:11px;color:#999;margin-top:5px}
.spd-card{background:#edf9f4;border:0.5px solid #1D9E75;border-radius:8px;padding:.8rem 1rem;margin-top:.9rem;display:grid;grid-template-columns:1fr auto;gap:16px;align-items:center}
.spd-lbl{font-size:12px;font-weight:500;color:#0F6E56;margin-bottom:3px}
.spd-names{font-size:15px;font-weight:500;margin-bottom:5px}
.spd-sep{color:#bbb;font-weight:400}
.spd-meta{font-size:12px;color:#444;line-height:1.55;margin-bottom:8px}
.spd-bar-track{height:6px;background:rgba(15,110,86,0.18);border-radius:3px;overflow:hidden}
.spd-bar-fill{height:6px;background:#0F6E56;border-radius:3px}
.spd-right{text-align:right;flex-shrink:0}
.spd-pct{font-size:32px;font-weight:500;color:#0F6E56;line-height:1}
.spd-pct-lbl{font-size:11px;color:#085041;opacity:.8;margin-top:2px}
.spd-score{font-size:12px;color:#085041;margin-top:5px}
.no-spd{background:#f7f7f5;border:0.5px solid #e0e0e0;border-radius:8px;padding:.75rem 1rem;margin-top:.9rem;text-align:center}
.no-spd-title{font-size:13px;font-weight:500;color:#aaa;margin-bottom:3px}
.no-spd-sub{font-size:11px;color:#bbb}
.fallback-note{background:#FFF8E8;border:0.5px solid #EF9F27;border-radius:8px;padding:.75rem 1rem;margin-top:.9rem;font-size:11px;color:#5C3A00;line-height:1.6}
.fallback-note strong{color:#3D2800}
"""

PLOT_JS = """
const P={P_JSON};
const PLT_MULT={PLT_MULT};
function drawTunnel(id){
  const cv=document.getElementById(id);
  const W=cv.width,H=cv.height,pad=16,maxR=24;
  const ctx=cv.getContext('2d');
  const xs=P.map(p=>p.tx),zs=P.map(p=>p.tz);
  const xR=(Math.max(...xs)-Math.min(...xs))+maxR*2.5;
  const zR=(Math.max(...zs)-Math.min(...zs))+maxR*2.5;
  const sc=Math.min((W-2*pad)/xR,(H-2*pad)/zR);
  const xM=(Math.max(...xs)+Math.min(...xs))/2,zM=(Math.max(...zs)+Math.min(...zs))/2;
  const toC=(x,z)=>({cx:W/2+(x-xM)*sc,cy:H/2-(z-zM)*sc});
  ctx.clearRect(0,0,W,H);
  P.forEach(p=>{
    const {cx,cy}=toC(p.tx,p.tz);
    const r=Math.max(12,Math.min(maxR,p.frac*0.62));
    ctx.beginPath();ctx.arc(cx,cy,r,0,Math.PI*2);
    ctx.fillStyle=p.c+'ee';ctx.fill();ctx.strokeStyle=p.c;ctx.lineWidth=1.5;ctx.stroke();
    ctx.font=`500 ${Math.max(9,Math.min(12,r))}px Arial`;
    ctx.fillStyle='#fff';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(p.t,cx,cy);
  });
  ctx.font='9px Arial';ctx.fillStyle='rgba(0,0,0,0.28)';
  ctx.textAlign='left';ctx.textBaseline='bottom';ctx.fillText('← glove',pad+2,H-2);
  ctx.textAlign='right';ctx.fillText('arm →',W-pad-2,H-2);
}
function drawPlate(id){
  const cv=document.getElementById(id);
  const W=cv.width,H=cv.height,pad=10;
  const ctx=cv.getContext('2d');
  const xs=P.map(p=>p.px),zs=P.map(p=>p.pz);
  const xMin=Math.min(...xs),xMax=Math.max(...xs),zMin=Math.min(...zs),zMax=Math.max(...zs);
  const effX=(xMax-xMin)/PLT_MULT;
  const effZ=(zMax-zMin)/PLT_MULT;
  const sc=Math.min((W-2*pad)/effX,(H-2*pad)/effZ)*0.88;
  const xM=(xMax+xMin)/2,zM=(zMax+zMin)/2;
  const toC=(x,z)=>({cx:W/2+(x-xM)*sc,cy:H/2-(z-zM)*sc});
  ctx.clearRect(0,0,W,H);
  const tl=toC(-8.5,44),br=toC(8.5,18);
  if(br.cx>tl.cx&&br.cy>tl.cy){ctx.strokeStyle='rgba(0,0,0,0.13)';ctx.setLineDash([3,3]);ctx.lineWidth=0.5;ctx.strokeRect(tl.cx,tl.cy,br.cx-tl.cx,br.cy-tl.cy);ctx.setLineDash([]);}
  P.forEach(p=>{
    const {cx,cy}=toC(p.px,p.pz);
    const r=Math.max(14,Math.min(22,p.frac*0.50));
    ctx.beginPath();ctx.arc(cx,cy,r,0,Math.PI*2);
    ctx.fillStyle=p.c+'ee';ctx.fill();ctx.strokeStyle=p.c;ctx.lineWidth=1.5;ctx.stroke();
    ctx.font=`500 ${Math.max(9,Math.min(12,r))}px Arial`;
    ctx.fillStyle='#fff';ctx.textAlign='center';ctx.textBaseline='middle';ctx.fillText(p.t,cx,cy);
  });
  ctx.font='9px Arial';ctx.fillStyle='rgba(0,0,0,0.28)';
  ctx.textAlign='left';ctx.textBaseline='bottom';ctx.fillText('← glove',pad+2,H-2);
  ctx.textAlign='right';ctx.fillText('arm →',W-pad-2,H-2);
}
drawTunnel('cTun');drawPlate('cPlt');
"""

# ── Card helpers ──────────────────────────────────────────────────────────────
def sfx(n):
    if n is None: return ''
    if n == 1:  return 'st'
    if n == 2:  return 'nd'
    if n == 3:  return 'rd'
    return 'th'

def stat_color(p):
    if p is None: return '#aaa'
    if p >= 80:  return '#185FA5'
    if p >= 60:  return '#0F6E56'
    if p >= 35:  return '#5F5E5A'
    return '#BA7517'

def bar_color(p):
    if p >= 80:  return '#185FA5'
    if p >= 60:  return '#1D9E75'
    if p >= 35:  return '#888780'
    return '#BA7517'

def pct_color(p):
    if p >= 80:  return '#185FA5'
    if p >= 60:  return '#0F6E56'
    if p >= 35:  return '#5F5E5A'
    return '#BA7517'


def build_league_pools(pools_df):
    all_ratios = []; all_temporal = []
    all_tr_pp  = []; all_rc_pp   = []; all_tm_pp = []
    for pid, group in pools_df.groupby('pitcher_id'):
        if len(group) < 2: continue
        group = group[group['pitches'] >= 10].copy()
        total = group['pitches'].sum()
        group['pitch_frac'] = group['pitches'] / total
        rows = group.to_dict('records')
        tr_w = []; rc_w = []; tm_w = []
        for r1, r2 in combinations(rows, 2):
            td  = np.sqrt((r1['tunnel_x']-r2['tunnel_x'])**2+(r1['tunnel_z']-r2['tunnel_z'])**2)
            pd_ = np.sqrt((r1['plate_x'] -r2['plate_x']) **2+(r1['plate_z'] -r2['plate_z']) **2)
            vd  = abs(r1['velo'] - r2['velo'])
            rd  = np.sqrt((r1['rel_x']-r2['rel_x'])**2+(r1['rel_z']-r2['rel_z'])**2)
            if td < 0.5: continue
            ratio = pd_ / td
            uw    = np.sqrt(r1['pitch_frac'] * r2['pitch_frac'])
            if td < TD_TUNNEL and ratio > MIN_RATIO:
                all_ratios.append(ratio)
                tr_w.append((ratio, uw)); rc_w.append((rd, uw))
            if pd_ < PD_SPEED and vd > VD_SPEED:
                all_temporal.append(vd * (1/(1+pd_)))
                tm_w.append((vd*(1/(1+pd_)), uw))
        if tr_w and total >= 50:
            tw  = sum(x[1] for x in tr_w)
            tw2 = sum(x[1] for x in rc_w)
            all_tr_pp.append((pid, sum(x[0]*x[1] for x in tr_w)/tw))
            all_rc_pp.append((pid, sum(x[0]*x[1] for x in rc_w)/tw2))
        if tm_w and total >= 50:
            tw = sum(x[1] for x in tm_w)
            all_tm_pp.append((pid, sum(x[0]*x[1] for x in tm_w)/tw))
    return {
        'ratios': all_ratios, 'temporal': all_temporal,
        'tr_pp': all_tr_pp, 'rc_pp': all_rc_pp, 'tm_pp': all_tm_pp,
    }


@st.cache_data(ttl=21600, show_spinner=False)
def get_league_pools(start: str, end: str):
    """Cached wrapper: builds league pool stats for player card component scores.
    Pulls from the already-loaded store so no extra Statcast fetch is needed."""
    import sys as _s
    store = _s.modules.get('__tplus_data_store__', {})
    key   = f'data_{start}_{end}'
    if key in store:
        _, pools_df = store[key]
    else:
        _, pools_df = load_season_data(start, end)
    return build_league_pools(pools_df)


def get_pitcher_card_info(name, lb_row, pools_df, league_pools):
    p_match = pools_df[pools_df['pitcher_name'] == name]
    if p_match.empty: return None
    pid = int(p_match['pitcher_id'].iloc[0])

    tr_vals = [v for _, v in league_pools['tr_pp']]
    rc_vals = [v for _, v in league_pools['rc_pp']]
    tm_vals = [v for _, v in league_pools['tm_pp']]

    tr_val = next((v for p, v in league_pools['tr_pp'] if p == pid), None)
    tm_val = next((v for p, v in league_pools['tm_pp'] if p == pid), None)
    rc_val = next((v for p, v in league_pools['rc_pp'] if p == pid), None)

    tr_pct = int(np.mean([v < tr_val for v in tr_vals]) * 100) if tr_val and tr_vals else None
    tm_pct = int(np.mean([v < tm_val for v in tm_vals]) * 100) if tm_val and tm_vals else None
    rc_pct = int((1 - np.mean([v < rc_val for v in rc_vals])) * 100) if rc_val and rc_vals else None

    tp_pct = int(lb_row['tp_pct'])

    details = lb_row.get('pitch_details', [])
    if isinstance(details, str): details = json.loads(details)

    pitches_out = []
    for pt in sorted(details, key=lambda x: -x['frac']):
        r = p_match[p_match['pitch_type'] == pt['type']]
        if r.empty: continue
        r = r.iloc[0]
        pitches_out.append({
            't': pt['type'], 'frac': pt['frac'], 'velo': pt['velo'],
            'ivb': pt['ivb'], 'c': get_color(pt['type']),
            'tx': round(float(r['tunnel_x']), 2), 'tz': round(float(r['tunnel_z']), 2),
            'px': round(float(r['plate_x']),  2), 'pz': round(float(r['plate_z']),  2),
        })

    p_match2 = p_match[p_match['pitches'] >= 10].copy()
    total = p_match2['pitches'].sum()
    p_match2['pitch_frac'] = p_match2['pitches'] / total
    rows = p_match2.to_dict('records')
    tp_pairs = []; sp_pairs = []
    for r1, r2 in combinations(rows, 2):
        td  = np.sqrt((r1['tunnel_x']-r2['tunnel_x'])**2+(r1['tunnel_z']-r2['tunnel_z'])**2)
        pd_ = np.sqrt((r1['plate_x'] -r2['plate_x']) **2+(r1['plate_z'] -r2['plate_z']) **2)
        vd  = abs(r1['velo'] - r2['velo'])
        if td < 0.5: continue
        ratio = pd_ / td
        is_t  = td < TD_TUNNEL and ratio > MIN_RATIO
        is_s  = pd_ < PD_SPEED and vd > VD_SPEED
        if is_t:
            pct = int(np.mean([v < ratio for v in league_pools['ratios']]) * 100) if league_pools['ratios'] else 50
            tp_pairs.append({'p1': r1['pitch_type'], 'p2': r2['pitch_type'],
                             'ratio': ratio, 'td': td, 'pd': pd_, 'pct': pct})
        if is_s:
            score = vd * (1/(1+pd_))
            pct = int(np.mean([v < score for v in league_pools['temporal']]) * 100) if league_pools['temporal'] else 50
            sp_pairs.append({'p1': r1['pitch_type'], 'p2': r2['pitch_type'],
                             'vd': vd, 'pd': pd_, 'score': score, 'pct': pct})

    tp_sorted = sorted(tp_pairs, key=lambda x: -x['ratio'])
    sp_sorted = sorted(sp_pairs, key=lambda x: -x['score'])
    total_hand = int(lb_row.get('total_hand', 0))

    return {
        'name': name,
        'hand': lb_row['hand'],
        'team': lb_row['team'],
        'pitches': lb_row['pitches'],
        'tplus': lb_row['tunneling_plus'],
        'rank': lb_row['rank'],
        'total': total_hand,
        'tp_pct': tp_pct,
        'tr_pct': tr_pct, 'tr_val': round(tr_val, 3) if tr_val else None,
        'tm_pct': tm_pct, 'tm_val': round(tm_val, 4) if tm_val else None,
        'rc_pct': rc_pct, 'rc_val': round(rc_val, 2)  if rc_val else None,
        'n_tp_total': len(tp_pairs),
        'n_sp_total': len(sp_pairs),
        'pitches_out': pitches_out,
        'tp_top3': tp_sorted[:3],
        'sp_best': sp_sorted[0] if sp_sorted else None,
        'fallback': lb_row.get('fallback', False),
    }


def make_card_html(info):
    P   = info['pitches_out']
    col = {p['t']: p['c'] for p in P}
    pct = info['tp_pct']
    mult = round(0.35 + (pct / 100) * 0.65, 4)

    arsenal = ''.join([
        f'<div class="ap"><span class="ap-dot" style="background:{p["c"]}"></span>'
        f'<span class="ap-name">{p["t"]}</span>'
        f'<span class="ap-val">{p["frac"]}% · {p["velo"]} mph</span></div>'
        for p in P])

    legend = ''.join([
        f'<span class="li"><span class="li-dot" style="background:{p["c"]}"></span>'
        f'{p["t"]} IVB {("+" if p["ivb"]>=0 else "")}{p["ivb"]}"</span>'
        for p in P])

    tp    = info['tp_top3']
    top_n = min(3, len(tp))
    grid  = 'pairs-grid' if top_n == 3 else 'pairs-grid-2'
    tp_html = ''.join([
        f'<div class="pair-card">'
        f'<div class="pair-names">'
        f'<span style="color:{col.get(p["p1"],"#888")}">{p["p1"]}</span>'
        f'<span class="pair-sep"> / </span>'
        f'<span style="color:{col.get(p["p2"],"#888")}">{p["p2"]}</span></div>'
        f'<div class="pair-meta">{p["td"]:.1f}" apart at tunnel<br>{p["pd"]:.1f}" apart at plate</div>'
        f'<div class="bar-row"><div class="bar-track"><div class="bar-fill" '
        f'style="width:{p["pct"]}%;background:{bar_color(p["pct"])}"></div></div>'
        f'<span class="bar-pct" style="color:{pct_color(p["pct"])}">{p["pct"]}th</span></div>'
        f'<div class="pair-ratio">Tunnel ratio {p["ratio"]:.2f}x</div></div>'
        for p in tp[:top_n]])

    sp = info['sp_best']
    if sp:
        c1 = col.get(sp['p1'], '#888'); c2 = col.get(sp['p2'], '#888')
        spd_html = (
            f'<div class="spd-card"><div>'
            f'<div class="spd-lbl">Best speed-change pair · {sp["p1"]} / {sp["p2"]}</div>'
            f'<div class="spd-names"><span style="color:{c1}">{sp["p1"]}</span>'
            f'<span class="spd-sep"> / </span><span style="color:{c2}">{sp["p2"]}</span></div>'
            f'<div class="spd-meta">Both pitches arrive <strong>{sp["pd"]:.1f}&quot; apart</strong>'
            f' — similar location — but with a <strong>{sp["vd"]:.1f} mph speed gap</strong>.'
            f' The hitter commits at 23.8ft before the difference is detectable.</div>'
            f'<div class="spd-bar-track"><div class="spd-bar-fill" style="width:{sp["pct"]}%"></div></div>'
            f'</div><div class="spd-right">'
            f'<div class="spd-pct">{sp["pct"]}th</div>'
            f'<div class="spd-pct-lbl">percentile</div>'
            f'<div class="spd-score">score {sp["score"]:.2f}</div>'
            f'</div></div>')
    else:
        spd_html = ('<div class="no-spd">'
                    '<div class="no-spd-title">No speed-change pairs</div>'
                    '<div class="no-spd-sub">No pitch pairs with similar plate location '
                    'and meaningful velocity gap</div></div>')

    def stat_card(pct_val, lbl, raw, sub, invert=False):
        if pct_val is None:
            return (f'<div class="stat-card">'
                    f'<div class="stat-pct" style="color:#aaa">—</div>'
                    f'<div class="stat-lbl">{lbl}</div>'
                    f'<div class="stat-raw">No qualifying pairs</div>'
                    f'<div class="stat-sub">{sub}</div></div>')
        return (f'<div class="stat-card">'
                f'<div class="stat-pct" style="color:{stat_color(pct_val)}">'
                f'{pct_val}{sfx(pct_val)}</div>'
                f'<div class="stat-lbl">{lbl}</div>'
                f'<div class="stat-raw">{raw} · {"lower" if invert else "higher"} is better</div>'
                f'<div class="stat-sub">{sub}</div></div>')

    stat1 = stat_card(info['tr_pct'], 'Tunnel ratio',
                      f'{info["tr_val"]}x' if info['tr_val'] else '—',
                      'Plate spread relative to tunnel tightness — bigger late break after the commit point')
    stat2 = stat_card(info['tm_pct'], 'Speed-change deception',
                      str(info['tm_val']) if info['tm_val'] else '—',
                      'Same plate location, different speed — hitter commits before detecting the difference')
    stat3 = stat_card(info['rc_pct'], 'Release consistency',
                      f'{info["rc_val"]}"' if info['rc_val'] else '—',
                      'How similar his release looks across pitch types — tighter means less tip-off',
                      invert=True)

    fallback_html = ('<div class="fallback-note"><strong>Note:</strong> No pairs met the '
                     'tunnel or speed-change classification thresholds. '
                     'Scored using unclassified pair metrics.</div>') if info.get('fallback') else ''

    hand_str  = HAND_STR.get(info['hand'], 'RHP')
    team_full = TEAM_FULL.get(info['team'], info['team'])
    total_str = f" of {info['total']}" if info['total'] else ''
    tp_sec    = (f'Top {top_n} tunnel pair{"s" if top_n>1 else ""} '
                 f'<span style="font-weight:400;letter-spacing:0;text-transform:none;'
                 f'font-size:12px;color:#aaa">— best of {info["n_tp_total"]} qualifying. '
                 f'Bar = percentile vs all MLB tunnel pairs.</span>'
                 ) if top_n > 0 else 'No tunnel pairs'
    tp_body = (f'<div class="{grid}">{tp_html}</div>' if top_n > 0 else
               '<div class="no-spd"><div class="no-spd-title">No qualifying tunnel pairs</div>'
               '<div class="no-spd-sub">No pitch pairs met the tunnel classification threshold</div></div>')

    P_json = json.dumps(P)
    js = PLOT_JS.replace('{P_JSON}', P_json).replace('{PLT_MULT}', str(mult))

    return f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><style>{CARD_CSS}</style></head><body>
<div class="card"><div class="hdr">
  <div class="hdr-top">
    <div><div class="hdr-name">{info['name']}</div>
    <div class="hdr-sub">{hand_str} · {team_full} · 2026 · {info['pitches']} pitches</div></div>
    <div class="hdr-pct"><div class="hdr-pct-num">{pct}th</div>
    <div class="hdr-pct-lbl">percentile</div></div>
  </div>
  <div class="hdr-center">
    <div class="tp-num">{info['tplus']}</div>
    <div class="tp-lbl">Tunneling+</div>
    <div class="tp-rank">Rank #{info['rank']}{total_str} qualified pitchers</div>
  </div>
  <div class="hdr-watermark">By Robert Colonna</div>
</div>
<div class="body">
  <div class="sec">Component scores</div>
  <div class="top-stats">{stat1}{stat2}{stat3}</div>
  <div class="divider"></div>
  <div class="sec">Pitch arsenal</div>
  <div class="arsenal-wrap">{arsenal}</div>
  <div class="divider"></div>
  <div class="sec">Pitch location — catcher's view</div>
  <div class="plots-wrap"><div class="plots-row">
    <div class="plot-block">
      <div class="plot-title">At tunnel point · 23.8ft — pitches converge</div>
      <canvas id="cTun" width="310" height="215"></canvas>
    </div>
    <div class="plot-block">
      <div class="plot-title">At the plate — pitches diverge</div>
      <canvas id="cPlt" width="310" height="215"></canvas>
    </div>
  </div><div class="leg-row">{legend}</div></div>
  <div class="divider"></div>
  <div class="sec">{tp_sec}</div>
  {tp_body}
  {spd_html}
  {fallback_html}
</div></div>
<script>{js}</script></body></html>"""


def find_chromium_executable():
    """Check for a system Chromium before using Playwright's bundled one."""
    import shutil as _shutil
    for name in ['chromium', 'chromium-browser', 'google-chrome-stable', 'google-chrome']:
        path = _shutil.which(name)
        if path:
            return path
    return None


def install_chromium():
    """Install Playwright Chromium into PLAYWRIGHT_BROWSERS_PATH."""
    # PLAYWRIGHT_BROWSERS_PATH is already set at module load time.
    # Just run the install — it will use the env var automatically.
    proc = subprocess.run(
        [sys.executable, '-m', 'playwright', 'install', 'chromium'],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=os.environ.copy(),
        text=True,
    )
    if proc.returncode != 0:
        raise RuntimeError(
            'Failed to install Chromium automatically.\n'
            f'{proc.stdout}\n{proc.stderr}'
        )


async def _render_jpg(html: str) -> bytes:
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        chromium_exec = find_chromium_executable()
        launch_kwargs = {
            'headless': True,
            'args': ['--no-sandbox', '--disable-dev-shm-usage', '--disable-gpu'],
        }
        if chromium_exec:
            launch_kwargs['executable_path'] = chromium_exec
        browser = await p.chromium.launch(**launch_kwargs)
        page    = await browser.new_page(viewport={'width': 760, 'height': 2000})
        await page.set_content(html)
        await page.wait_for_timeout(800)
        card      = await page.query_selector('.card')
        img_bytes = await card.screenshot(type='jpeg', quality=95)
        await browser.close()
    return img_bytes


def render_card(html: str) -> bytes:
    # PLAYWRIGHT_BROWSERS_PATH is set at module load — no need to repeat here.
    try:
        return asyncio.run(_render_jpg(html))
    except Exception as e:
        err_text = str(e).lower()
        if any(k in err_text for k in ('chromium', 'browser', 'playwright',
                                        "doesn't exist", 'executable')):
            try:
                install_chromium()
                return asyncio.run(_render_jpg(html))
            except Exception as install_err:
                raise RuntimeError(
                    'Playwright is installed but Chromium could not be installed. '
                    'Please ensure the environment allows `playwright install chromium`.\n'
                    f'{install_err}'
                ) from install_err
        raise RuntimeError(
            'Playwright render failed. Ensure Playwright and Chromium are installed: '
            '`pip install playwright && playwright install chromium`'
        ) from e




# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title='Tunneling+',
    page_icon='⚾',
    layout='wide',
    initial_sidebar_state='expanded'
)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    _logo_b64 = "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XJv/AIY/zXIiie01ec9/3awxf7XjL1OO/wAFaf8A9cm/+GP81f4KU3/rs3/wx/muQKhPaavOP7t4Z/4vGXqceGlKfrXSj/8AKH+aSaSg9HmK4v3+50HD4FciAOFU9qq85R6L4Y/9vxl6nF/4JO3T/wAox5xwBiP+a/Bukq4/+dUY83u/8K5e7gFpz4r0ryqWJ6JYa+DXb65nC5NMXVjy1scUoBwHNlGD78L85dO3lgz8gkd+I5r/ALCVzjC1N4cl6V7PijFnoXZv4ZyXc/odcyW24RgmShqmAcy6Jw/cvlx5LtHfeOT3D2qSffGbkwZK3Od17A4e4q4r/niYdTQf/wAdbvX3OryMKLsaS3WyRu5JbaUj8Fm4fe3C+SfT1mlGGw1EB745c/rZVyN7Te8i62h1/D4HGXb6o4Ii5dU6Tp3Fxp7g9vDg2WLPHzB/ctvm0rdGY9F8nn/u5cY/Swr0binLcyLrYBiNH4qT7NvlmbCi+2ptVxpxmainY3Gc7hIx5r4yParqae4ip0503lNZPpCivFRVPAREQBERAEREARVRAEREAREQBERAEVRARERAEREBVERAEREAREQBERAEREBVERAEREAREQBERAEREAREQBXooiAIiIAiIgKoiIAiIgCvBREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBFcqIAiIgCIiAIiIAivRRAEREARFcEoAovtobXX1ozTUssg+tjDfeeC3yj0lJwNZWRsGR6sQ3yR148gferc6sIfEzPtMLu7v5NNtc/Dv3HF8FfrTU09Q7cghkld3MaT9i53S2K0Uo4UhncMjendn4DgtxBDW7sbWxt6NY0NHwWLO+gvhWZs1poVcz215qPVtf0XicKpdL3SUZlZFTjj/KvwfcMrc6fSdK3jUVssh4cImBuO/icrkCqxp3lR7thsVtolh9H405vpfpkfFT2a1U+DHQseeHGUl54efBffGdxm5GGxtHABgAGFACqGkngseU5S3snreytrf5VNR6kjS7icnj5qFfoGOc4Mb6zieDRxJ9i5HY9A60ve79y9K3ipDuTxTOa0/nOwF5Szewu1a1OktackkcYRdwWbs47T7huumttDbWnmaurGR7Ggrmtm7KV0duuvGraODv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a5JHyPL3uc554lzjkn2rQivxio7kQda5rV3rVZOT6XmXJURF6LAREQBERAEREAVCiIAiIgCIiAIiIAqVEQBERAEREAREQBEVQEREQBERAEREAREQBERAEREBVERAEREAREQBERAERXogIidEQBERAEREARXqogCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAqoiAIiIAiIgCIiAIiIAqoiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAqQoiAIqogCIiAIiIAiIgCIiAIiIAiIgKoiIC7xQOIORwI5EKIgzPsiuVfE8PZW1LXDkRK7/NfRHf7swYFdKeOfWw77QtrReXCL3oyad5cU/gm11Nm9nU92Lt50sR8PQt/cv3/AIWV5aB8loie/cd/4lx1F4dCm/2oyo41iEd1aXezkkGq6gPBmoqd7eoaXNPvyV+/8LR0trP/AIx/yXFEXn2alzGRDSLEo7qr8H5o5W3Vrc+tbm48Jj/ktR1ZD0tzv/jf/JcTUVPZaXMXVpRimXzPCPocr/hZH/7PP/xv/wClP4WR/wDs7/8Ae/8A6VxRE9lpcxT+8+J/+Twj6HKxqyP/ANm5/wDzz/4VXatj3SBbGg44Ezk4PuXE0T2WlzFHpNib/wB3wXociGq6wPz8jpCO4h3+a0S6quLhhsVLH4tjJ+0lbAi9KhTX7TGnjmIS31pd+Xkby/U12Jbuzsjx9SJvH3hfhJfLrI1zXV84Due67d+xbai9qnBbkY88Ru6nxVZPtZ9E1bVzP35aqeR3e6Qk/avwyoi9ZGJKTk82y5KiIqnkIqogCIiAIiIAiIgCIiAIiIAiKoCIiIAiIgCIiAIiIAiIgCIqgIiIgCIiAqiIgCIiAIiIAqoqgIiIgCIiAIiIAiIgCoURAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBEVQEREQBERAUKIiAIiIAqoiAIiIAiIgCIiAIiIAiIgCIiAKqIgCIiAIiIAiIgCqiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAIiIAiIgCIiAKqKjmgIiFEAREQBERAEREAREQBERAEREAREQBVREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBVREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREAREQBERAEREB//2Q=="
    st.markdown(
        f'''<div style="display:flex;align-items:center;gap:12px;margin-bottom:4px">
        <img src="data:image/png;base64,{_logo_b64}" width="52" style="border-radius:50%"/>
        <div><div style="font-size:1.25rem;font-weight:700">Tunneling+</div>
        <div style="font-size:0.8rem;color:#888">By Robert Colonna</div></div>
        </div>''',
        unsafe_allow_html=True
    )
    st.markdown('---')

    season = 2026  # locked to 2026 — multi-season switching caused OOM crashes

    st.markdown('**Date range**')
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        start_date = st.date_input(
            'From',
            value=date(season, 3, 27),
            min_value=date(2021, 1, 1),
            max_value=date.today(),
        )
    with col_d2:
        end_date = st.date_input(
            'Through',
            value=date.today(),
            min_value=date(2021, 1, 1),
            max_value=date.today(),
        )
    start_str = start_date.strftime('%Y-%m-%d')
    end_str   = end_date.strftime('%Y-%m-%d')

    st.markdown('---')
    min_pitches = st.slider('Min pitches (total)', 25, 200, 50, step=25)

    st.markdown('---')
    if st.button('🔄 Force Refresh'):
        import sys as _sys2
        if '__tplus_data_store__' in _sys2.modules:
            _sys2.modules['__tplus_data_store__'].clear()
        st.cache_data.clear()
        st.rerun()

# ── Data loading ──────────────────────────────────────────────────────────────
# NOTE: _load_season_data_direct is the raw function used by the background
# thread. st.cache_data decorators cannot be called from non-Streamlit threads
# (they try to access ScriptRunContext and fail). The thread calls this directly;
# the result is stored in sys.modules and the main thread reads it from there.
def _compute_outcomes(df_raw):
    """Derive per-pitcher outcome stats from raw Statcast pitch-level data.
    Called before df_raw is trimmed so we still have all columns."""
    import pandas as _pd2
    import numpy as _np2

    # Column name map — savant uses 'pitcher' for ID, 'player_name' for name
    id_col   = 'pitcher'      if 'pitcher'      in df_raw.columns else None
    name_col = 'player_name'  if 'player_name'  in df_raw.columns else None
    if id_col is None or name_col is None:
        return _pd2.DataFrame()

    rows = []
    for pid, grp in df_raw.groupby(id_col):
        pname = grp[name_col].iloc[0] if name_col else str(pid)
        n_pitches = len(grp)
        if n_pitches < 50:
            continue

        # Whiff% = swinging strikes / total pitches
        desc = grp['description'] if 'description' in grp.columns else _pd2.Series(dtype=str)
        swings   = desc.isin(['swinging_strike', 'swinging_strike_blocked', 'foul', 'foul_tip', 'hit_into_play', 'hit_into_play_no_out', 'hit_into_play_score'])
        whiffs   = desc.isin(['swinging_strike', 'swinging_strike_blocked'])
        whiff_pct = float(whiffs.sum() / n_pitches * 100) if n_pitches > 0 else None

        # Called strike + whiff % (CSW)
        called_strikes = desc.isin(['called_strike'])
        csw_pct = float((whiffs.sum() + called_strikes.sum()) / n_pitches * 100) if n_pitches > 0 else None

        # GB% — of balls in play
        bb_type = grp['bb_type'] if 'bb_type' in grp.columns else _pd2.Series(dtype=str)
        bip = bb_type.notna() & (bb_type != '')
        gb  = bb_type == 'ground_ball'
        gb_pct = float(gb.sum() / bip.sum() * 100) if bip.sum() > 0 else None
        fb_pct = float((bb_type == 'fly_ball').sum() / bip.sum() * 100) if bip.sum() > 0 else None

        # K% and BB% — from plate appearances (events column)
        events = grp['events'] if 'events' in grp.columns else _pd2.Series(dtype=str)
        pa_events = events.notna() & (events != '')
        n_pa = pa_events.sum()
        k_pct  = float(events.isin(['strikeout','strikeout_double_play']).sum() / n_pa * 100) if n_pa > 0 else None
        bb_pct = float(events.isin(['walk','intent_walk']).sum() / n_pa * 100) if n_pa > 0 else None

        # xwOBA — mean of estimated_woba_using_speedangle where available
        xwoba_col = 'estimated_woba_using_speedangle'
        xwoba = float(grp[xwoba_col].dropna().mean()) if xwoba_col in grp.columns else None

        # Barrel% — launch_speed >= 98 and launch_angle 26-30
        if 'launch_speed' in grp.columns and 'launch_angle' in grp.columns:
            bip_df  = grp[bip].copy() if bip.sum() > 0 else grp.iloc[0:0]
            barrels = bip_df[
                (bip_df['launch_speed'] >= 98) &
                (bip_df['launch_angle'].between(26, 30))
            ]
            barrel_pct = float(len(barrels) / bip.sum() * 100) if bip.sum() > 0 else None
        else:
            barrel_pct = None

        rows.append({
            'pitcher_id': int(pid),
            'sv_name':    pname,
            'n_pitches':  n_pitches,
            'whiff_pct':  round(whiff_pct, 2)  if whiff_pct  is not None else None,
            'csw_pct':    round(csw_pct, 2)    if csw_pct    is not None else None,
            'gb_pct':     round(gb_pct, 2)     if gb_pct     is not None else None,
            'fb_pct':     round(fb_pct, 2)     if fb_pct     is not None else None,
            'k_pct':      round(k_pct, 2)      if k_pct      is not None else None,
            'bb_pct':     round(bb_pct, 2)     if bb_pct     is not None else None,
            'xwoba':      round(xwoba, 4)      if xwoba      is not None else None,
            'barrel_pct': round(barrel_pct, 2) if barrel_pct is not None else None,
        })

    return _pd2.DataFrame(rows)


def _load_season_data_direct(start: str, end: str):
    df_raw = load_statcast(start, end, verbose=False)
    c, f   = run_model(df_raw)
    q      = normalize(c, f)

    # Trim df_raw to only the columns downstream code actually needs.
    # This dramatically reduces RAM — the full df_raw may have dozens of extra
    # Statcast columns never used after this point. Freeing it before storing
    # prevents OOM crashes when switching seasons.
    _POOLS_COLS = [
        'pitcher_id', 'pitcher_name', 'pitcher_team', 'hand',
        'pitch_type', 'pitches',
        'velo', 'ivb', 'hb',
        'tunnel_x', 'tunnel_z', 'plate_x', 'plate_z',
        'extension', 'release_height', 'release_side', 'rel_x', 'rel_z',
    ]
    _keep = [col for col in _POOLS_COLS if col in df_raw.columns]
    pools = df_raw[_keep].copy()
    del df_raw  # free the full table immediately
    return q, pools

# Cached wrapper for calls made from the main Streamlit thread (e.g. rerenders)
@st.cache_data(ttl=21600, show_spinner=False)
def load_season_data(start: str, end: str):
    return _load_season_data_direct(start, end)


@st.cache_data(ttl=3600, show_spinner=False)
def load_outcome_data(season: int):
    """Pull outcome stats from Baseball Savant via pybaseball (same domain as main load).
    Three sources merged on pitcher MLBAM id:
      - statcast_pitcher_exitvelo_barrels : K%, BB%, Whiff%, Barrel%, Hard Hit%
      - statcast_pitcher_expected_stats   : xwOBA, xERA, xBA
      - statcast_pitcher_pitch_arsenal    : Chase%, GB%, O-Swing%, Z-Contact%
    """
    try:
        from pybaseball import (
            statcast_pitcher_exitvelo_barrels as _evb,
            statcast_pitcher_expected_stats   as _exp,
            statcast_pitcher_pitch_arsenal    as _ars,
        )

        def _prep(df):
            df = df.copy()
            df.columns = [c.lower().strip() for c in df.columns]
            for id_col in ('pitcher', 'player_id', 'mlbam_id'):
                if id_col in df.columns:
                    df.rename(columns={id_col: 'pitcher_id'}, inplace=True)
                    break
            df['pitcher_id'] = pd.to_numeric(df['pitcher_id'], errors='coerce')
            return df

        ev = _prep(_evb(season, minBBE=25))
        ex = _prep(_exp(season, minPA=25))
        # pitch_arsenal has one row per pitcher per pitch type — aggregate to pitcher level
        ar_raw = _prep(_ars(season, minP=50, arsenal_type='n_'))
        # columns of interest from arsenal: gb_percent, o_swing_percent, z_contact_percent, whiff_percent
        # aggregate by pitcher_id (mean across pitch types, weighted by pitch count if available)
        ar_cols = [c for c in ar_raw.columns if c in (
            'gb_percent','o_swing_percent','z_contact_percent',
            'whiff_percent','chase_percent','csw_percent',
        )]
        if ar_cols and 'pitcher_id' in ar_raw.columns:
            ar = ar_raw.groupby('pitcher_id')[ar_cols].mean().reset_index()
        else:
            ar = pd.DataFrame(columns=['pitcher_id'])

        # Column rename maps
        EV_MAP = {
            'k_percent':         'k_pct',
            'bb_percent':        'bb_pct',
            'whiff_percent':     'whiff_pct',
            'barrel_batted_rate':'barrel_pct',
            'hard_hit_percent':  'hard_hit_pct',
        }
        EX_MAP = {
            'est_woba': 'xwoba',
            'est_era':  'xera',
            'est_ba':   'xba',
        }
        AR_MAP = {
            'gb_percent':         'gb_pct',
            'o_swing_percent':    'o_swing_pct',
            'z_contact_percent':  'z_contact_pct',
            'whiff_percent':      'whiff_pct_ar',  # deduplicate if both sources have it
            'chase_percent':      'chase_pct',
            'csw_percent':        'csw_pct',
        }

        def _slim(df, col_map):
            keep = ['pitcher_id'] + [c for c in col_map if c in df.columns]
            return df[keep].rename(columns=col_map)

        ev_c = _slim(ev, EV_MAP)
        ex_c = _slim(ex, EX_MAP)
        ar_c = _slim(ar, AR_MAP)

        out = ev_c.merge(ex_c, on='pitcher_id', how='outer')                   .merge(ar_c, on='pitcher_id', how='outer')

        # Prefer ev whiff_pct; if ar also has it, drop the duplicate
        if 'whiff_pct' in out.columns and 'whiff_pct_ar' in out.columns:
            out['whiff_pct'] = out['whiff_pct'].fillna(out['whiff_pct_ar'])
            out.drop(columns=['whiff_pct_ar'], inplace=True)

        for col in out.columns:
            if col != 'pitcher_id':
                out[col] = pd.to_numeric(
                    out[col].astype(str).str.replace('%','').str.strip(),
                    errors='coerce'
                )
        return out

    except Exception as _e:
        import traceback
        return pd.DataFrame({'_error': [str(_e)], '_tb': [traceback.format_exc()]})




# Use sys.modules to store results — truly process-global, survives reruns.
import threading, time as _time, sys as _sys

_STORE_KEY = '__tplus_data_store__'
if _STORE_KEY not in _sys.modules:
    _sys.modules[_STORE_KEY] = {}
_tplus_store = _sys.modules[_STORE_KEY]

_cache_key  = f'data_{start_str}_{end_str}'
_err_key    = f'err_{start_str}_{end_str}'
_thread_key = f'thread_{start_str}_{end_str}'
_t0_key     = f't0_{start_str}_{end_str}'

# Evict any OTHER season's data immediately so we never hold two seasons in RAM
for _stale_key in [k for k in list(_tplus_store.keys())
                   if k.startswith('data_') and k != _cache_key]:
    del _tplus_store[_stale_key]
for _stale_k in [k for k in list(_tplus_store.keys())
                 if (k.startswith('thread_') or k.startswith('err_') or k.startswith('t0_'))
                 and not k.endswith(f'_{start_str}_{end_str}')]:
    _tplus_store.pop(_stale_k, None)

# Start background thread if result not yet available and no thread running
if _cache_key not in _tplus_store:
    _existing_thread = _tplus_store.get(_thread_key)
    _thread_running  = _existing_thread is not None and _existing_thread.is_alive()
    if not _thread_running and _err_key not in _tplus_store:
        def _load_in_background():
            try:
                _tplus_store[_cache_key] = _load_season_data_direct(start_str, end_str)
            except Exception as exc:
                _tplus_store[_err_key] = exc

        _t = threading.Thread(target=_load_in_background, daemon=True)
        _tplus_store[_thread_key] = _t
        _tplus_store[_t0_key]     = _time.time()
        _t.start()

# Poll: show status and rerun
if _cache_key not in _tplus_store:
    _thread_done = (
        _tplus_store.get(_thread_key) is not None
        and not _tplus_store[_thread_key].is_alive()
    )
    _elapsed = int(_time.time() - _tplus_store.get(_t0_key, _time.time()))
    if _err_key in _tplus_store:
        _err = _tplus_store[_err_key]
        st.error(f'Failed to load data: {_err}')
        st.exception(_err)
        if st.button('🔄 Retry'):
            for _k in [_cache_key, _err_key, _thread_key, _t0_key]:
                _tplus_store.pop(_k, None)
            st.rerun()
        st.stop()
    if _elapsed > 480:
        st.error(f'Data load timed out after {_elapsed}s. Try clicking Retry.')
        if st.button('🔄 Retry'):
            for _k in [_cache_key, _err_key, _thread_key, _t0_key]:
                _tplus_store.pop(_k, None)
            st.rerun()
        st.stop()
    if _thread_done:
        st.error(
            'Data load thread finished without a result. '
            f'Store keys: {list(_tplus_store.keys())}.'
        )
        if st.button('🔄 Retry'):
            for _k in [_cache_key, _err_key, _thread_key, _t0_key]:
                _tplus_store.pop(_k, None)
            st.rerun()
        st.stop()
    st.info(f'⏳ Loading Statcast data... ({_elapsed}s — full season pull takes 2–3 min)')
    _time.sleep(2)
    st.rerun()

if _err_key in _tplus_store:
    st.error(f'Failed to load data: {_tplus_store[_err_key]}')
    st.exception(_tplus_store[_err_key])
    st.stop()

lb, pools = _tplus_store[_cache_key]

# Filter and rerank
lb_filtered = lb[lb['pitches'] >= min_pitches].copy()
lb_filtered['rank'] = lb_filtered['tunneling_plus'].rank(
    ascending=False, method='min').astype(int)
lb_filtered['tp_pct'] = lb_filtered['tunneling_plus'].rank(
    pct=True).mul(100).round(0).astype(int)

# Percentile columns for leaderboard
lb_filtered['tr_pct'] = lb_filtered['avg_tunnel_ratio'].rank(
    pct=True).mul(100).round(0).astype(int)
lb_filtered['spd_pct'] = lb_filtered['n_speed_pairs'].rank(
    pct=True).mul(100).round(0).astype(int)

def _compute_rc(pools_df, pitcher_ids):
    from itertools import combinations as _comb
    rc_map = {}
    for pid, grp in pools_df.groupby('pitcher_id'):
        if pid not in pitcher_ids:
            continue
        grp = grp[grp['pitches'] >= 10].copy()
        if len(grp) < 2:
            continue
        total = grp['pitches'].sum()
        grp['pitch_frac'] = grp['pitches'] / total
        rc_w = []
        for (_, r1), (_, r2) in _comb(grp.iterrows(), 2):
            if 'rel_x' not in r1 or 'rel_x' not in r2:
                continue
            rd = float(np.sqrt((r1['rel_x']-r2['rel_x'])**2+(r1['rel_z']-r2['rel_z'])**2))
            uw = float(r1['pitch_frac'] * r2['pitch_frac'])
            rc_w.append((rd, uw))
        if rc_w:
            tw = sum(x[1] for x in rc_w)
            rc_map[pid] = sum(x[0]*x[1] for x in rc_w) / tw
    return rc_map

_rc_map = _compute_rc(pools, set(lb_filtered['pitcher_id'].tolist()))
lb_filtered['rc_val'] = lb_filtered['pitcher_id'].map(_rc_map)
lb_filtered['rc_pct'] = (
    lb_filtered['rc_val'].rank(pct=True, ascending=False)
    .mul(100).round(0).fillna(50).astype(int)
)

for hand in ['R', 'L']:
    mask = lb_filtered['hand'] == hand
    lb_filtered.loc[mask, 'total_hand'] = int(mask.sum())

total_q = len(lb_filtered)
st.success(
    f'✓ {total_q} pitchers · {start_str} → {end_str} · '
    f'Loaded {datetime.now(timezone(timedelta(hours=-5))).strftime("%b %d, %I:%M %p")} EST'
)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs(
    ['📊 Leaderboard', '🃏 Player Card', '📥 Export', '🔬 Diagnostic', '📖 Methodology', '🔍 Tunnel Comps', '📈 Correlations']
)

# ─── Tab 1: Leaderboard ───────────────────────────────────────────────────────
with tab1:
    st.caption(f'Season {season} · {start_str} → {end_str} · Model v18 (Statcast)')

    fc1, fc2, fc3 = st.columns([3, 1, 1])
    with fc1:
        all_names   = sorted(lb_filtered['name'].dropna().unique().tolist())
        name_filter = st.multiselect('Search pitchers', all_names, placeholder='Type to search…')
    with fc2:
        teams    = ['All'] + sorted(lb_filtered['team'].dropna().unique().tolist())
        team_sel = st.selectbox('Team', teams)
    with fc3:
        hand_sel = st.selectbox('Hand', ['All', 'R', 'L'])

    display = lb_filtered.copy()
    if name_filter:
        display = display[display['name'].isin(name_filter)]
    if team_sel != 'All':
        display = display[display['team'] == team_sel]
    if hand_sel != 'All':
        display = display[display['hand'] == hand_sel]

    show_cols = ['rank', 'tunneling_plus', 'tp_pct', 'name', 'team', 'hand',
                 'pitches', 'n_types', 'n_tunnel_pairs', 'avg_tunnel_ratio',
                 'n_speed_pairs', 'temporal']
    col_names = {
        'rank': 'Rank', 'tunneling_plus': 'T+', 'tp_pct': 'Pct',
        'name': 'Pitcher', 'team': 'Team', 'hand': 'Hand',
        'pitches': 'Pitches', 'n_types': 'Types',
        'n_tunnel_pairs': 'Tun Pairs', 'avg_tunnel_ratio': 'Avg Ratio',
        'n_speed_pairs': 'Spd Pairs', 'temporal': 'Temporal'
    }
    disp = display[show_cols].rename(columns=col_names)
    st.dataframe(
        disp,
        width="stretch",
        hide_index=True,
        column_config={
            'Rank':      st.column_config.NumberColumn(width='small'),
            'T+':        st.column_config.NumberColumn(format='%.1f', width='small'),
            'Pct':       st.column_config.NumberColumn(format='%d', width='small'),
            'Pitcher':   st.column_config.TextColumn(width='medium'),
            'Team':      st.column_config.TextColumn(width='small'),
            'Hand':      st.column_config.TextColumn(width='small'),
            'Pitches':   st.column_config.NumberColumn(width='small'),
            'Types':     st.column_config.NumberColumn(width='small'),
            'Tun Pairs': st.column_config.NumberColumn(width='small'),
            'Avg Ratio': st.column_config.NumberColumn(format='%.3f', width='medium'),
            'Spd Pairs': st.column_config.NumberColumn(width='small'),
            'Temporal':  st.column_config.NumberColumn(format='%.4f', width='medium'),
        }
    )
    st.caption(f'Showing {len(disp)} of {total_q} qualified pitchers')

    # Quick-nav to Player Card
    if len(display) <= 20:
        nav_names = display['name'].tolist()
        if nav_names:
            st.markdown('**→ Player Card:**')
            nav_cols = st.columns(min(len(nav_names), 5))
            for i, pname in enumerate(nav_names[:10]):
                with nav_cols[i % 5]:
                    if st.button(pname, key=f'nav_{pname}'):
                        st.session_state['selected_pitcher'] = pname
                        st.session_state['active_tab'] = 1

# ─── Tab 2: Player Card ───────────────────────────────────────────────────────
with tab2:
    pitcher_names   = lb_filtered['name'].sort_values().tolist()
    _default_pitcher = st.session_state.get('selected_pitcher', pitcher_names[0] if pitcher_names else None)
    _default_idx    = pitcher_names.index(_default_pitcher) if _default_pitcher in pitcher_names else 0
    selected = st.selectbox('Select pitcher', pitcher_names, index=_default_idx)
    # Clear nav state after use
    if 'selected_pitcher' in st.session_state:
        del st.session_state['selected_pitcher']

    playwright_ok = True
    try:
        from playwright.async_api import async_playwright  # noqa
    except ImportError:
        playwright_ok = False

    if selected:
        lb_row = lb_filtered[lb_filtered['name'] == selected].iloc[0]

        if not playwright_ok:
            st.warning(
                'Playwright or Chromium not available — showing text card only. '
                'Run `pip install playwright && playwright install chromium` for JPG cards.'
            )
            tp  = lb_row['tunneling_plus']
            c1, c2, c3, c4 = st.columns(4)
            c1.metric('T+', f'{tp:.1f}')
            c2.metric('Percentile', f'{int(lb_row["tp_pct"])}th')
            c3.metric('Tunnel Pairs', int(lb_row['n_tunnel_pairs']))
            c4.metric('Avg Ratio', f'{lb_row["avg_tunnel_ratio"]:.3f}x')
            details = lb_row.get('pitch_details', [])
            if isinstance(details, str): details = json.loads(details)
            if details:
                adf = pd.DataFrame(sorted(details, key=lambda x: -x['frac']))
                adf = adf.rename(columns={'type':'Pitch','pitches':'N','velo':'Velo',
                                          'ivb':'IVB','hb':'HB','frac':'Usage%'})
                cols = [c for c in ['Pitch','N','Usage%','Velo','IVB','HB'] if c in adf.columns]
                st.dataframe(adf[cols], width="stretch", hide_index=True)
        else:
            with st.spinner(f'Rendering card for {selected}…'):
                try:
                    league_pools = get_league_pools(start_str, end_str)
                    info         = get_pitcher_card_info(selected, lb_row, pools, league_pools)
                    if info is None:
                        st.error(f'No pitch-level data found for {selected}.')
                    else:
                        html      = make_card_html(info)
                        jpg_bytes = render_card(html)
                        st.image(jpg_bytes, width=720)
                        slug = selected.lower().replace(' ', '_').replace('.', '')
                        st.download_button(
                            label='⬇️ Download card JPG',
                            data=jpg_bytes,
                            file_name=f'{slug}_tunneling_card.jpg',
                            mime='image/jpeg'
                        )
                except Exception as e:
                    st.error(f'Card render failed: {e}')
                    st.exception(e)

# ─── Tab 3: Export ────────────────────────────────────────────────────────────
with tab3:
    st.markdown('### Download Leaderboard')
    st.markdown(
        f'Exports the full **{total_q}-pitcher** leaderboard '
        f'({start_str} → {end_str}) in V17-compatible Excel format.'
    )
    if st.button('📥 Generate Excel'):
        with st.spinner('Building Excel…'):
            try:
                tmp = '/tmp/tunneling_export.xlsx'
                build_excel(lb_filtered, tmp, start_str, end_str)
                with open(tmp, 'rb') as f:
                    st.download_button(
                        label='⬇️ Download Tunneling+.xlsx',
                        data=f.read(),
                        file_name=f'tunneling_plus_{season}_{end_str}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
            except ImportError:
                st.error('openpyxl not installed. Run: `pip install openpyxl`')

# ─── Tab 4: Diagnostic ───────────────────────────────────────────────────────
with tab4:
    st.markdown('### Statcast Aggregated Data')
    st.markdown(
        'Per pitcher per pitch type after all calibration fixes. '
        'All pitches — no windup filter. '
        'hb sign is hand-dependent (LHP fix applied). '
        'All pitches — no windup or runner filter. hb sign is hand-dependent.'
    )

    search_diag = st.text_input('Filter pitcher', '', key='diag_search')

    diag_cols = ['pitcher_id', 'pitcher_name', 'hand', 'pitcher_team', 'pitch_type',
                 'pitches', 'velo', 'ivb', 'hb', 'extension', 'release_height',
                 'release_side', 'tunnel_x', 'tunnel_z', 'plate_x', 'plate_z']
    diag_df = pools[[c for c in diag_cols if c in pools.columns]].copy()

    if search_diag:
        try:
            diag_df = diag_df[diag_df['pitcher_name'].str.contains(
                search_diag, case=False, na=False)]
        except Exception:
            pass

    st.dataframe(diag_df.round(4), width="stretch", hide_index=True)

    csv_bytes = diag_df.round(4).to_csv(index=False).encode()
    st.download_button(
        label='⬇️ Download aggregated Statcast CSV',
        data=csv_bytes,
        file_name=f'statcast_aggregated_{start_str}_{end_str}.csv',
        mime='text/csv'
    )
    st.caption(
        'release_side = release_pos_x + (0.3655×ext − 2.4608). '
        'hb: RHP = pfx_x×12, LHP = pfx_x×12×−1. All pitches, no windup filter.'
    )

# ─── Tab 5: Methodology ───────────────────────────────────────────────────────
with tab5:
    st.markdown("## Tunneling+ — Model Methodology")
    st.caption(f"v18 · Statcast Edition · {start_str} to {end_str}")

    st.markdown("---")
    st.markdown("### What is Tunneling+?")
    st.markdown(
        "Tunneling+ measures how well a pitcher disguises his pitches through the "
        "**tunnel point** — the moment roughly 23.8 ft from the plate where a hitter "
        "must commit to swing or take. Pitches that look identical at the tunnel point "
        "but diverge dramatically by the plate are the hardest to read. "
        "Normalized within handedness: 100 = league average, 1 SD approx 15 points, "
        "winsorized to [60, 160]."
    )

    st.markdown("---")
    st.markdown("### Pair Classification")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.markdown("**Tunnel pair**")
        st.markdown(
            "Tunnel distance < 6.6 inches (league p50) "
            "AND plate divergence ratio > 1.5x. "
            "Pitches that converge at the commit point then break sharply apart."
        )
    with col_b:
        st.markdown("**Speed-change pair**")
        st.markdown(
            "Plate distance < 14.9 inches (league p33) "
            "AND velocity gap > 6.8 mph (league p50). "
            "Same location, different speed — hitter commits before detecting the difference."
        )
    with col_c:
        st.markdown("**Irrelevant pair**")
        st.markdown(
            "Neither condition met. Excluded for pitchers with at least one qualifying pair. "
            "Pitchers with only irrelevant pairs are scored on raw metrics and "
            "flagged 'Unclassified pairs.'"
        )

    st.markdown("---")
    st.markdown("### Tunnel Composite Score")
    st.markdown(
        "Qualifying tunnel pairs are weighted by **usage x tunnel ratio** — "
        "higher weight to elite, frequently-thrown pairs. "
        "The composite is a weighted average of five components:"
    )
    components = [
        ("35%", "Tunnel Ratio (TR)",
         "Plate spread divided by tunnel distance. Primary signal — how much the pitch "
         "breaks after the commit point relative to how tight it looks at the tunnel point."),
        ("20%", "Break:Tunnel Ratio (BTR)",
         "Additional plate spread beyond the tunnel distance. "
         "Rewards pitches that break more than they tunnel."),
        ("15%", "Interaction Index (IX)",
         "Axis alignment combined with plate divergence. Rewards pitches that tunnel "
         "on-axis but diverge sharply at the plate."),
        ("15%", "Release:Tunnel Ratio (RTR) — inverted",
         "Release spread relative to tunnel distance. Lower is better — a consistent "
         "release slot hiding different pitches scores higher."),
        ("10%", "Late Break Multiplier (LBM)",
         "Rewards pitches with break concentrated after the tunnel point "
         "rather than early in flight."),
        ("5%",  "Velocity Differential (VD)",
         "Absolute velocity gap between paired pitches. Modest contribution — "
         "speed difference adds deception on top of movement."),
    ]
    for pct, name, desc in components:
        st.markdown(f"**{pct} — {name}:** {desc}")
        st.markdown("")

    st.markdown("---")
    st.markdown("### Final Composite and Normalization")
    st.markdown(
        "Final score = **80% tunnel composite + 20% temporal** (speed-change deception). "
        "Both components are z-scored within handedness before combining, so RHP and LHP "
        "are evaluated against their own peers. Converted to a 100-point scale "
        "(mean=100, std approx 15), winsorized to [60, 160]. "
        "Minimum 10 pitches per pitch type and 50 total pitches to qualify."
    )

    st.markdown("---")
    st.markdown("### Data and Calibration")
    cc1, cc2 = st.columns(2)
    with cc1:
        st.markdown("**Source**")
        st.markdown(
            "Baseball Savant (Statcast) via pybaseball. "
            "All pitches included — no runner or windup filter. "
            "Data cached for 1 hour per date range."
        )
        st.markdown("**hb sign convention**")
        st.markdown(
            "Positive hb = arm side for that pitcher. "
            "RHP: pfx_x x 12 x -1. "
            "LHP: pfx_x x 12 (no flip). "
            "Converts Statcast catcher-relative coordinates to pitcher-relative arm-side convention."
        )
    with cc2:
        st.markdown("**release_side correction**")
        st.markdown(
            "release_side = release_pos_x + (0.3655 x extension - 2.4608). "
            "Extension-based correction fitted from diagnostic analysis of "
            "Statcast release_pos_x. "
            "Reduces mean offset to under 0.03 ft per pitch type."
        )
        st.markdown("**Known limitation**")
        st.markdown(
            "Statcast release_pos_x has less within-pitcher spread across pitch types "
            "than some reference implementations (avg 0.24 ft vs 0.33 ft). "
            "This produces a Statcast-native calibration."
        )

    st.markdown("---")
    st.markdown("### Model Constants")
    const_data = {
        "Tunnel point": "23.8 ft from plate",
        "Tunnel threshold": "6.6 in (league p50 tunnel distance)",
        "Min tunnel ratio": "1.5x",
        "Speed-change plate threshold": "14.9 in (league p33)",
        "Speed-change velo threshold": "6.8 mph (league p50)",
        "League avg plate distance": "20.07 in",
        "Min pitches per pitch type": "10",
        "Min total pitches to qualify": "50",
        "Score range": "60 to 160 (winsorized)",
        "Normalization": "Within handedness, mean=100, std approx 15",
    }
    for k, v in const_data.items():
        st.markdown(f"- **{k}:** {v}")

    st.markdown("---")
    st.caption("Tunneling+ v18 · By Robert Colonna · Built on Statcast via pybaseball")


# ─── Tab 6: Tunnel Comps ──────────────────────────────────────────────────────

# ─── Tab 6: Tunnel Comps ──────────────────────────────────────────────────────
with tab6:
    st.markdown("### Pitcher Tunnel Comps")
    st.caption(
        "Finds the three pitchers who tunnel most like the selected pitcher. "
        "Similarity is based on pitch mix (50%) and tunneling metrics (50%), "
        "within same handedness."
    )

    import numpy as np

    comp_names = lb_filtered['name'].sort_values().tolist()
    _comp_default = st.session_state.get('comp_target', comp_names[0] if comp_names else None)
    _comp_idx = comp_names.index(_comp_default) if _comp_default in comp_names else 0
    comp_target = st.selectbox('Select pitcher to find comps for', comp_names, index=_comp_idx, key='comp_select')
    if 'comp_target' in st.session_state:
        del st.session_state['comp_target']

    if comp_target:
        # ── Build pitch mix vectors ────────────────────────────────────────
        # `pools` is the aggregated per-pitcher/per-pitch-type df (already in scope)
        # It has columns: pitcher_name, pitch_type, pitches, etc.

        ALL_TYPES = ['FF', 'SI', 'FC', 'SL', 'ST', 'CU', 'KC', 'CH', 'FS', 'KN', 'EP', 'SC']

        def _pitch_mix_vec(pitcher_name):
            rows = pools[pools['pitcher_name'] == pitcher_name]
            rows = rows[rows['pitches'] >= 10]
            total = rows['pitches'].sum()
            if total == 0:
                return None
            vec = []
            for pt in ALL_TYPES:
                match = rows[rows['pitch_type'] == pt]
                vec.append(float(match['pitches'].sum()) / total if len(match) > 0 else 0.0)
            return np.array(vec)

        def _cosine_sim(a, b):
            na, nb = np.linalg.norm(a), np.linalg.norm(b)
            if na == 0 or nb == 0:
                return 0.0
            return float(np.dot(a, b) / (na * nb))

        # ── Build metric vectors ───────────────────────────────────────────
        METRIC_COLS = ['tunneling_plus', 'avg_tunnel_ratio', 'n_tunnel_pairs',
                       'n_speed_pairs', 'temporal']

        lb_metrics = lb_filtered[['name', 'hand'] + METRIC_COLS].copy().dropna()

        # Z-score normalize metrics
        for col in METRIC_COLS:
            mu = lb_metrics[col].mean()
            sd = lb_metrics[col].std()
            lb_metrics[col + '_z'] = (lb_metrics[col] - mu) / sd if sd > 0 else 0.0

        z_cols = [c + '_z' for c in METRIC_COLS]

        # ── Target pitcher ─────────────────────────────────────────────────
        target_row = lb_filtered[lb_filtered['name'] == comp_target]
        if target_row.empty:
            st.warning("Pitcher not found in leaderboard.")
        else:
            target_hand = target_row.iloc[0]['hand']
            target_mix  = _pitch_mix_vec(comp_target)

            target_metrics_row = lb_metrics[lb_metrics['name'] == comp_target]
            if target_metrics_row.empty or target_mix is None:
                st.warning("Not enough data for this pitcher.")
            else:
                target_z = target_metrics_row[z_cols].values[0]

                # ── Score all same-handed pitchers ─────────────────────────────────
                same_hand = lb_metrics[
                    (lb_metrics['hand'] == target_hand) &
                    (lb_metrics['name'] != comp_target)
                ]

                scores = []
                for _, row in same_hand.iterrows():
                    cand_name = row['name']
                    cand_mix  = _pitch_mix_vec(cand_name)
                    if cand_mix is None:
                        continue
                    mix_sim    = _cosine_sim(target_mix, cand_mix)
                    cand_z     = row[z_cols].values
                    metric_dist = np.linalg.norm(target_z - cand_z)
                    metric_sim  = 1.0 / (1.0 + metric_dist)
                    combined    = 0.50 * mix_sim + 0.50 * metric_sim
                    scores.append((cand_name, combined, mix_sim, metric_sim))

                scores.sort(key=lambda x: -x[1])
                top3 = scores[:3]

                if not top3:
                    st.warning("Not enough same-handed pitchers for comparison.")
                else:
                    # ── Display ────────────────────────────────────────────────────
                    target_lb = lb_filtered[lb_filtered['name'] == comp_target].iloc[0]
                    target_tp = target_lb['tunneling_plus']
                    target_pct = int(target_lb['tp_pct'])

                    st.markdown(f"**{comp_target}** ({target_hand}HP) · T+ **{target_tp:.1f}** · {target_pct}th pct")
                    st.markdown("---")
                    st.markdown("**Top 3 Tunnel Comps:**")

                    for rank_i, (cname, combined, mix_s, met_s) in enumerate(top3, 1):
                        comp_lb = lb_filtered[lb_filtered['name'] == cname].iloc[0]
                        comp_tp  = comp_lb['tunneling_plus']
                        comp_pct = int(comp_lb['tp_pct'])
                        comp_team = comp_lb.get('team', '')

                        # Pitch mix comparison
                        t_mix = _pitch_mix_vec(comp_target)
                        c_mix = _pitch_mix_vec(cname)
                        shared_types = [pt for i, pt in enumerate(ALL_TYPES)
                                        if (t_mix[i] > 0.03 or c_mix[i] > 0.03)]

                        with st.expander(
                            f"#{rank_i} **{cname}** ({comp_team}) · T+ {comp_tp:.1f} · "
                            f"{comp_pct}th pct · Similarity {combined*100:.0f}%",
                            expanded=True
                        ):
                            mc1, mc2, mc3 = st.columns(3)
                            mc1.metric("Overall Similarity", f"{combined*100:.0f}%")
                            mc2.metric("Pitch Mix Similarity", f"{mix_s*100:.0f}%")
                            mc3.metric("Metric Similarity", f"{met_s*100:.0f}%")

                            # Pitch mix side-by-side
                            if shared_types:
                                st.markdown("**Pitch Mix Comparison**")
                                mix_data = []
                                for pt in shared_types:
                                    idx = ALL_TYPES.index(pt)
                                    mix_data.append({
                                        'Pitch': pt,
                                        comp_target: f"{t_mix[idx]*100:.1f}%",
                                        cname: f"{c_mix[idx]*100:.1f}%",
                                    })
                                mix_df = pd.DataFrame(mix_data)
                                st.dataframe(mix_df, hide_index=True, width="content")

                            # Metric comparison
                            st.markdown("**Metric Comparison**")
                            metric_labels = {
                                'tunneling_plus': 'T+',
                                'avg_tunnel_ratio': 'Avg Tunnel Ratio',
                                'n_tunnel_pairs': 'Tunnel Pairs',
                                'n_speed_pairs': 'Speed Pairs',
                                'temporal': 'Temporal Score',
                            }
                            met_data = []
                            for col, label in metric_labels.items():
                                t_val = target_lb[col]
                                c_val = comp_lb[col]
                                fmt = '.1f' if col == 'tunneling_plus' else '.3f' if col in ('avg_tunnel_ratio', 'temporal') else '.0f'
                                met_data.append({
                                    'Metric': label,
                                    comp_target: f'{t_val:{fmt}}',
                                    cname: f'{c_val:{fmt}}',
                                })
                            met_df = pd.DataFrame(met_data)
                            st.dataframe(met_df, hide_index=True, width="content")

                            if st.button(f"Open {cname}'s Player Card", key=f'comp_nav_{rank_i}_{cname}'):
                                st.session_state['selected_pitcher'] = cname
                                st.session_state['comp_target'] = comp_target


# ─── Tab 7: Metric Correlations ───────────────────────────────────────────────
with tab7:
    st.markdown("### Metric Correlations")
    st.caption("How do tunneling metrics relate to real-world outcomes? Outcome data pulled automatically from Baseball Savant.")

    with st.spinner("Loading outcome data from Baseball Savant..."):
        outcomes = load_outcome_data(season)

    if outcomes.empty or '_error' in outcomes.columns:
        err_msg = outcomes['_error'].iloc[0] if '_error' in outcomes.columns else "Empty response"
        tb_msg  = outcomes['_tb'].iloc[0]    if '_tb'    in outcomes.columns else ""
        st.error(f"Could not load outcome data: {err_msg}")
        if tb_msg:
            with st.expander("Full traceback"):
                st.code(tb_msg)
    else:
        # ── Join on pitcher_id ─────────────────────────────────────────────────
        lb_corr = lb_filtered[
            ['pitcher_id', 'name', 'hand', 'team', 'pitches',
             'tunneling_plus', 'tp_pct',
             'avg_tunnel_ratio', 'tr_pct',
             'n_speed_pairs', 'spd_pct',
             'rc_val', 'rc_pct',
             'temporal', 'n_tunnel_pairs']
        ].copy()
        lb_corr['pitcher_id'] = pd.to_numeric(lb_corr['pitcher_id'], errors='coerce')
        merged = lb_corr.merge(outcomes, on='pitcher_id', how='inner')

        if len(merged) < 5:
            st.warning(f"Only {len(merged)} pitchers matched. lb has {len(lb_corr)}, outcomes has {len(outcomes)}.")
        else:
            st.success(f"Matched {len(merged)} pitchers.")

            # ── Column definitions ─────────────────────────────────────────────
            OUTCOME_LABELS = {
                'whiff_pct':    'Whiff%',
                'k_pct':        'K%',
                'bb_pct':       'BB%',
                'chase_pct':    'Chase%',
                'o_swing_pct':  'O-Swing%',
                'gb_pct':       'GB%',
                'barrel_pct':   'Barrel%',
                'hard_hit_pct': 'Hard Hit%',
                'csw_pct':      'CSW%',
                'z_contact_pct':'Z-Contact%',
                'xwoba':        'xwOBA',
                'xera':         'xERA',
                'xba':          'xBA',
            }
            TUNNEL_COLS = {
                'tunneling_plus':   'T+',
                'avg_tunnel_ratio': 'Tunnel Ratio',
                'n_speed_pairs':    'Speed Pairs',
                'rc_val':           'Release Cons.',
                'temporal':         'Temporal',
                'n_tunnel_pairs':   'Tunnel Pairs',
            }
            avail_outcomes = {k: v for k, v in OUTCOME_LABELS.items()
                              if k in merged.columns and merged[k].notna().sum() >= 5}
            avail_tunnel   = {k: v for k, v in TUNNEL_COLS.items()
                              if k in merged.columns}

            # ── Shared axis selectors (drive both scatter and league table) ────
            st.markdown("---")
            _ax1, _ax2, _ax3 = st.columns([2, 2, 2])
            with _ax1:
                x_label = st.selectbox('Tunneling metric (X)', list(avail_tunnel.values()), key='corr_x')
            with _ax2:
                y_label = st.selectbox('Outcome (Y)', list(avail_outcomes.values()), key='corr_y')
            with _ax3:
                all_names_corr = sorted(merged['name'].dropna().unique().tolist())
                _def = st.session_state.get('selected_pitcher', all_names_corr[0] if all_names_corr else None)
                _idx = all_names_corr.index(_def) if _def in all_names_corr else 0
                hl_pitcher = st.selectbox('Highlight pitcher', ['(none)'] + all_names_corr, index=_idx + 1, key='corr_hl')

            x_col = [k for k, v in avail_tunnel.items()   if v == x_label][0]
            y_col = [k for k, v in avail_outcomes.items() if v == y_label][0]
            scatter_df = merged[['name', 'team', 'hand', x_col, y_col]].dropna()

            # ══════════════════════════════════════════════════════════════════
            # SECTION 1: Scatter Explorer
            # ══════════════════════════════════════════════════════════════════
            st.markdown("---")
            st.markdown("#### Scatter Explorer")
            st.markdown(
                "Each dot is one pitcher. The dashed line is the best-fit trend — "
                "a steep slope means that tunneling metric is a strong predictor of that outcome. "
                "Hover to see any pitcher. Use **Highlight pitcher** above to call one out specifically."
            )

            if len(scatter_df) >= 5:
                import json as _json
                _x  = scatter_df[x_col].values.astype(float)
                _y  = scatter_df[y_col].values.astype(float)
                _r  = float(pd.Series(_x).corr(pd.Series(_y)))
                _c  = np.polyfit(_x, _y, 1)
                xmin, xmax = float(_x.min()), float(_x.max())
                ymin, ymax = float(_y.min()), float(_y.max())
                xpad = (xmax - xmin) * 0.1 or 1
                ypad = (ymax - ymin) * 0.1 or 1
                tlx = [xmin - xpad, xmax + xpad]
                tly = [float(_c[0]*v + _c[1]) for v in tlx]

                # Mark the highlighted pitcher
                hl_row = None
                if hl_pitcher != '(none)':
                    _hl = scatter_df[scatter_df['name'] == hl_pitcher]
                    if not _hl.empty:
                        hl_row = {'name': hl_pitcher,
                                  'x': float(_hl[x_col].iloc[0]),
                                  'y': float(_hl[y_col].iloc[0]),
                                  'team': str(_hl['team'].iloc[0]),
                                  'hand': str(_hl['hand'].iloc[0])}

                chart_data = [{'name': row['name'], 'team': row['team'], 'hand': row['hand'],
                               'x': float(row[x_col]), 'y': float(row[y_col]),
                               'hl': row['name'] == hl_pitcher}
                              for _, row in scatter_df.iterrows()]

                scatter_html = f"""
<div style="font-family:sans-serif">
<div style="margin-bottom:6px;color:#555;font-size:13px">r = <strong>{_r:+.3f}</strong> &nbsp;·&nbsp; n = {len(scatter_df)}{f' &nbsp;·&nbsp; <span style="color:#e07b00;font-weight:600">▶ {hl_pitcher}</span>' if hl_pitcher != '(none)' else ''}</div>
<div style="position:relative;width:100%;height:460px;background:#fafafa;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden">
<canvas id="scMain" width="820" height="460" style="width:100%;height:100%"></canvas>
<div id="ttMain" style="position:absolute;display:none;background:#fff;border:1px solid #ccc;border-radius:6px;padding:6px 10px;font-size:12px;pointer-events:none;box-shadow:0 2px 8px rgba(0,0,0,.15)"></div>
</div></div>
<script>
(function(){{
  const pts={_json.dumps(chart_data)},tlX={_json.dumps(tlx)},tlY={_json.dumps(tly)};
  const xL={_json.dumps(x_label)},yL={_json.dumps(y_label)},hlName={_json.dumps(hl_pitcher)};
  const xMn={xmin-xpad},xMx={xmax+xpad},yMn={ymin-ypad},yMx={ymax+ypad};
  const cv=document.getElementById('scMain'),ctx=cv.getContext('2d'),W=cv.width,H=cv.height;
  const P={{l:64,r:24,t:24,b:54}};
  function tx(v){{return P.l+(v-xMn)/(xMx-xMn)*(W-P.l-P.r);}}
  function ty(v){{return H-P.b-(v-yMn)/(yMx-yMn)*(H-P.t-P.b);}}
  // grid
  ctx.strokeStyle='#ececec';ctx.lineWidth=1;
  for(let i=0;i<=5;i++){{
    const y=P.t+i*(H-P.t-P.b)/5;ctx.beginPath();ctx.moveTo(P.l,y);ctx.lineTo(W-P.r,y);ctx.stroke();
    const x=P.l+i*(W-P.l-P.r)/5;ctx.beginPath();ctx.moveTo(x,P.t);ctx.lineTo(x,H-P.b);ctx.stroke();
  }}
  // axes
  ctx.strokeStyle='#bbb';ctx.lineWidth=1.5;
  ctx.beginPath();ctx.moveTo(P.l,P.t);ctx.lineTo(P.l,H-P.b);ctx.lineTo(W-P.r,H-P.b);ctx.stroke();
  // axis labels
  ctx.fillStyle='#555';ctx.font='13px sans-serif';ctx.textAlign='center';
  ctx.fillText(xL,W/2,H-6);
  ctx.save();ctx.translate(13,H/2);ctx.rotate(-Math.PI/2);ctx.fillText(yL,0,0);ctx.restore();
  // ticks
  ctx.font='10px sans-serif';ctx.fillStyle='#aaa';
  for(let i=0;i<=5;i++){{
    const xv=xMn+i*(xMx-xMn)/5;ctx.textAlign='center';ctx.fillText(xv.toFixed(2),tx(xv),H-P.b+14);
    const yv=yMn+i*(yMx-yMn)/5;ctx.textAlign='right';ctx.fillText(yv.toFixed(2),P.l-5,ty(yv)+4);
  }}
  // trendline
  ctx.strokeStyle='rgba(200,80,80,.5)';ctx.lineWidth=1.5;ctx.setLineDash([6,4]);
  ctx.beginPath();ctx.moveTo(tx(tlX[0]),ty(tlY[0]));ctx.lineTo(tx(tlX[1]),ty(tlY[1]));ctx.stroke();
  ctx.setLineDash([]);
  // draw non-highlighted first, then highlighted on top
  pts.filter(p=>!p.hl).forEach(p=>{{
    ctx.beginPath();ctx.arc(tx(p.x),ty(p.y),4.5,0,Math.PI*2);
    ctx.fillStyle=p.hand==='L'?'rgba(59,130,246,.55)':'rgba(16,185,129,.55)';
    ctx.strokeStyle=p.hand==='L'?'#3b82f6':'#10b981';ctx.lineWidth=.8;
    ctx.fill();ctx.stroke();
  }});
  pts.filter(p=>p.hl).forEach(p=>{{
    // halo
    ctx.beginPath();ctx.arc(tx(p.x),ty(p.y),11,0,Math.PI*2);
    ctx.fillStyle='rgba(230,120,0,.18)';ctx.fill();
    // dot
    ctx.beginPath();ctx.arc(tx(p.x),ty(p.y),7,0,Math.PI*2);
    ctx.fillStyle='#e07b00';ctx.strokeStyle='#a05000';ctx.lineWidth=1.5;
    ctx.fill();ctx.stroke();
    // label
    ctx.fillStyle='#a05000';ctx.font='bold 11px sans-serif';ctx.textAlign='left';
    const lx=tx(p.x)+10, ly=ty(p.y)-6;
    ctx.fillText(p.name,lx,ly);
  }});
  // tooltip
  const tip=document.getElementById('ttMain');
  cv.addEventListener('mousemove',function(e){{
    const rc=cv.getBoundingClientRect(),sx=cv.width/rc.width,sy=cv.height/rc.height;
    const mx=(e.clientX-rc.left)*sx,my=(e.clientY-rc.top)*sy;
    let hit=null,bestD=12;
    pts.forEach(p=>{{const dx=tx(p.x)-mx,dy=ty(p.y)-my,d=Math.sqrt(dx*dx+dy*dy);if(d<bestD){{bestD=d;hit=p;}}}});
    if(hit){{
      tip.style.display='block';
      tip.style.left=(e.clientX-rc.left+14)+'px';tip.style.top=(e.clientY-rc.top-24)+'px';
      tip.innerHTML=`<strong>${{hit.name}}</strong> (${{hit.team}} · ${{hit.hand}}HP)<br>${{xL}}: ${{hit.x.toFixed(3)}}<br>${{yL}}: ${{hit.y.toFixed(3)}}`;
    }}else tip.style.display='none';
  }});
  cv.addEventListener('mouseleave',()=>tip.style.display='none');
}})();
</script>"""
                st.components.v1.html(scatter_html, height=500)
                st.caption("🔵 LHP &nbsp;·&nbsp; 🟢 RHP &nbsp;·&nbsp; 🟠 Highlighted pitcher &nbsp;·&nbsp; Hover for details")

            # ══════════════════════════════════════════════════════════════════
            # SECTION 2: League correlation table (same outcome as scatter)
            # ══════════════════════════════════════════════════════════════════
            st.markdown("---")
            st.markdown("#### League-Wide Correlations")
            st.markdown(
                f"How strongly each tunneling metric correlates with **{y_label}** across all {len(merged)} pitchers. "
                "Pearson r ranges from -1 to +1. Values near 0 mean little relationship. "
                f"{'A *negative* r is the good direction here — lower '+y_label+' is better for the pitcher.' if y_col in ('xwoba','xera','barrel_pct','bb_pct') else 'A *positive* r means the metric predicts better outcomes.'}"
            )

            # Build one row per tunneling metric for the selected outcome only
            corr_rows = []
            for tcol, tlabel in avail_tunnel.items():
                pair = merged[[tcol, y_col]].dropna()
                if len(pair) >= 8:
                    r = round(float(pair[tcol].corr(pair[y_col])), 3)
                    # Also compute r for all outcomes for completeness
                    corr_rows.append({'Tunneling Metric': tlabel, f'r vs {y_label}': r, 'n': len(pair)})

            if corr_rows:
                corr_tbl = pd.DataFrame(corr_rows).sort_values(f'r vs {y_label}', key=abs, ascending=False)
                st.dataframe(corr_tbl, hide_index=True, width='stretch',
                    column_config={{
                        f'r vs {y_label}': st.column_config.NumberColumn(format='%.3f'),
                        'n': st.column_config.NumberColumn('Sample', width='small'),
                    }})

            # ══════════════════════════════════════════════════════════════════
            # SECTION 3: Individual pitcher breakdown (same axes)
            # ══════════════════════════════════════════════════════════════════
            st.markdown("---")
            st.markdown("#### Individual Pitcher Breakdown")
            st.markdown(
                f"What does **{hl_pitcher if hl_pitcher != '(none)' else 'the selected pitcher'}**'s tunneling profile predict for their outcomes, "
                "and how do their actual results compare? "
                "**Predicted** is from linear regression of each metric vs that outcome league-wide. "
                "**Exp Range** is ±1 SD — where ~68% of similar pitchers land. "
                "✅ outperforming · ⚠️ underperforming · ➖ within expected range."
            )

            sel = hl_pitcher if hl_pitcher != '(none)' else (all_names_corr[0] if all_names_corr else None)

            if sel and sel in merged['name'].values:
                pr = merged[merged['name'] == sel].iloc[0]
                st.markdown(f"**{sel}** &nbsp;·&nbsp; {pr.get('team','')} &nbsp;·&nbsp; {pr.get('hand','')}HP")

                rows_bd = []
                for tcol, tlabel in avail_tunnel.items():
                    for ocol, olabel in avail_outcomes.items():
                        pair = merged[[tcol, ocol]].dropna()
                        if len(pair) < 8: continue
                        r = float(pair[tcol].corr(pair[ocol]))
                        if abs(r) < 0.10: continue
                        _xv = pair[tcol].values.astype(float)
                        _yv = pair[ocol].values.astype(float)
                        _co = np.polyfit(_xv, _yv, 1)
                        _m, _b = float(_co[0]), float(_co[1])
                        pred   = _m * float(pr[tcol]) + _b
                        sd     = float(np.std(_yv - (_m * _xv + _b)))
                        actual = pr.get(ocol, None)
                        if pd.isna(actual): continue
                        actual = float(actual)
                        _low = ocol in ('xwoba','xera','barrel_pct','bb_pct','hard_hit_pct')
                        if _low:
                            verdict = '✅ Better' if actual < pred-sd else ('⚠️ Worse' if actual > pred+sd else '➖ In range')
                        else:
                            verdict = '✅ Better' if actual > pred+sd else ('⚠️ Worse' if actual < pred-sd else '➖ In range')
                        rows_bd.append({
                            'Tunnel Metric': tlabel, 'Outcome': olabel,
                            'r': round(r,3), 'Predicted': round(pred,3),
                            'Exp Range': f'{pred-sd:.2f} – {pred+sd:.2f}',
                            'Actual': round(actual,3), 'vs Expected': verdict,
                        })

                if rows_bd:
                    bd_df = pd.DataFrame(rows_bd).sort_values(['Outcome','Tunnel Metric'])
                    st.dataframe(bd_df, hide_index=True, width='stretch',
                        column_config={{
                            'r':         st.column_config.NumberColumn('Pearson r', format='%.3f', width='small'),
                            'Predicted': st.column_config.NumberColumn(format='%.3f', width='small'),
                            'Actual':    st.column_config.NumberColumn(format='%.3f', width='small'),
                        }})
                else:
                    st.info("No correlations strong enough (|r| ≥ 0.10) for this pitcher.")
            else:
                st.info("Select a pitcher in the Highlight pitcher dropdown above to see their breakdown.")
