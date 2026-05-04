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
from datetime import date, datetime
from itertools import combinations

sys.path.insert(0, os.path.dirname(__file__))

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


async def _render_jpg(html: str) -> bytes:
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch()
        page    = await browser.new_page(viewport={'width': 760, 'height': 2000})
        await page.set_content(html)
        await page.wait_for_timeout(800)
        card      = await page.query_selector('.card')
        img_bytes = await card.screenshot(type='jpeg', quality=95)
        await browser.close()
    return img_bytes


def install_chromium():
    try:
        subprocess.run(
            [sys.executable, '-m', 'playwright', 'install', 'chromium'],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        return True
    except Exception as exc:
        raise RuntimeError(
            'Failed to install Chromium automatically. '
            'Please install it manually with `playwright install chromium`.'
        ) from exc


def render_card(html: str) -> bytes:
    try:
        return asyncio.run(_render_jpg(html))
    except Exception as e:
        err_text = str(e).lower()
        if 'chromium' in err_text or 'browser' in err_text or 'playwright' in err_text:
            try:
                install_chromium()
                return asyncio.run(_render_jpg(html))
            except Exception as install_err:
                raise RuntimeError(
                    'Playwright is installed but Chromium could not be installed. '
                    'Please ensure the environment allows `playwright install chromium`.'
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
    st.markdown('## ⚾ Tunneling+')
    st.markdown('*By Robert Colonna*')
    st.markdown('---')

    season = st.selectbox('Season', [2026, 2025, 2024, 2023, 2022, 2021], index=0)

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
        st.cache_data.clear()
        st.rerun()

# ── Data loading ──────────────────────────────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner=False)
def load_season_data(start: str, end: str):
    df_raw = load_statcast(start, end, verbose=False)
    c, f   = run_model(df_raw)
    q      = normalize(c, f)
    return q, df_raw

@st.cache_data(ttl=3600, show_spinner=False)
def get_league_pools(start: str, end: str):
    _, df_raw = load_season_data(start, end)
    return build_league_pools(df_raw)

with st.spinner('Loading Statcast data… (first load ~2–3 min)'):
    try:
        lb, pools = load_season_data(start_str, end_str)
    except Exception as e:
        st.error(f'Failed to load data: {e}')
        st.stop()

# Filter and rerank
lb_filtered = lb[lb['pitches'] >= min_pitches].copy()
lb_filtered['rank'] = lb_filtered['tunneling_plus'].rank(
    ascending=False, method='min').astype(int)
lb_filtered['tp_pct'] = lb_filtered['tunneling_plus'].rank(
    pct=True).mul(100).round(0).astype(int)

for hand in ['R', 'L']:
    mask = lb_filtered['hand'] == hand
    lb_filtered.loc[mask, 'total_hand'] = int(mask.sum())

total_q = len(lb_filtered)
st.success(
    f'✓ {total_q} pitchers · {start_str} → {end_str} · '
    f'Loaded {datetime.now().strftime("%b %d, %I:%M %p")}'
)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs(
    ['📊 Leaderboard', '🃏 Player Card', '📥 Export', '🔬 Diagnostic']
)

# ─── Tab 1: Leaderboard ───────────────────────────────────────────────────────
with tab1:
    st.caption(f'Season {season} · {start_str} → {end_str} · Model v18 (Statcast)')

    fc1, fc2, fc3, _ = st.columns([2, 1, 1, 3])
    with fc1:
        search = st.text_input('Search pitcher', '')
    with fc2:
        teams    = ['All'] + sorted(lb_filtered['team'].dropna().unique().tolist())
        team_sel = st.selectbox('Team', teams)
    with fc3:
        hand_sel = st.selectbox('Hand', ['All', 'R', 'L'])

    display = lb_filtered.copy()
    if search:
        try:
            display = display[display['name'].str.contains(search, case=False, na=False)]
        except Exception:
            display = display.iloc[0:0]
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
        width='stretch',
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

# ─── Tab 2: Player Card ───────────────────────────────────────────────────────
with tab2:
    pitcher_names = lb_filtered['name'].sort_values().tolist()
    selected      = st.selectbox('Select pitcher', pitcher_names)

    playwright_ok = True
    try:
        from playwright.async_api import async_playwright  # noqa
    except ImportError:
        playwright_ok = False

    if selected:
        lb_row = lb_filtered[lb_filtered['name'] == selected].iloc[0]

        if not playwright_ok:
            st.warning(
                'Playwright not installed — showing text card only. '
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
                st.dataframe(adf[cols], width='stretch', hide_index=True)
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
                    if 'chromium' in str(e).lower() or 'playwright' in str(e).lower():
                        st.warning(
                            'Playwright is installed but Chromium is not available. '
                            'Run `playwright install chromium` for JPG cards.'
                        )
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
        'Use this to validate against TJ Stats.'
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

    st.dataframe(diag_df.round(4), width='stretch', hide_index=True)

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
