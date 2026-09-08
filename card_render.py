"""
Pillow-based Tunneling+ card renderer.
Replaces the Playwright/Chromium HTML-screenshot pipeline entirely — no
headless browser, no system libs, no packages.txt dependency.

Draws the exact same visual layout as the old CARD_CSS/PLOT_JS template,
directly onto a Pillow canvas.
"""

import io
import math
from PIL import Image, ImageDraw, ImageFont
from matplotlib import font_manager

# ── Fonts (bundled with matplotlib — no extra dependency, no network) ────────
_FONT_REG_PATH  = font_manager.findfont(font_manager.FontProperties(family='DejaVu Sans', weight='normal'))
_FONT_BOLD_PATH = font_manager.findfont(font_manager.FontProperties(family='DejaVu Sans', weight='bold'))

SCALE = 4  # supersample at 4x, then downsample to 2x at the end for crisp anti-aliased edges/text
FINAL_SCALE = 2

_font_cache = {}
def F(size, bold=False):
    key = (round(size * SCALE), bold)
    if key not in _font_cache:
        path = _FONT_BOLD_PATH if bold else _FONT_REG_PATH
        _font_cache[key] = ImageFont.truetype(path, key[0])
    return _font_cache[key]

def S(px):
    """Scale a CSS-px value to actual pixels."""
    return round(px * SCALE)


# ── Small drawing helpers ─────────────────────────────────────────────────────
def hex_rgb(h, alpha=None):
    h = h.lstrip('#')
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (r, g, b, alpha) if alpha is not None else (r, g, b)

def text_w(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0]

def draw_text(draw, xy, text, font, fill, anchor='la'):
    draw.text(xy, text, font=font, fill=fill, anchor=anchor)

def rounded_rect(draw, box, radius, fill=None, outline=None, width=1):
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=width)


# ── Color helpers (ported 1:1 from app.py) ────────────────────────────────────
def sfx(n):
    if n is None: return ''
    if n == 1: return 'st'
    if n == 2: return 'nd'
    if n == 3: return 'rd'
    return 'th'

def stat_color(p):
    if p is None: return '#aaaaaa'
    if p >= 80: return '#185FA5'
    if p >= 60: return '#0F6E56'
    if p >= 35: return '#5F5E5A'
    return '#BA7517'

def bar_color(p):
    if p >= 80: return '#185FA5'
    if p >= 60: return '#1D9E75'
    if p >= 35: return '#888780'
    return '#BA7517'

pct_color = stat_color  # identical logic in the original


# ── Layout constants (ported from CARD_CSS) ───────────────────────────────────
CARD_W       = 680
PAD_H        = 22   # ~1.4rem
PAD_V        = 18   # ~1.1rem
BG_WHITE     = '#ffffff'
BG_SOFT      = '#f7f7f5'
BORDER_SOFT  = '#e0e0e0'
BORDER_LT    = '#dddddd'
TEXT_DARK    = '#1a1a1a'
TEXT_GRAY    = '#888888'
TEXT_LGRAY   = '#999999'
NAVY         = '#1a2744'
BLUE_ACCENT  = '#60a8f0'
GREEN_BG     = '#edf9f4'
GREEN_BORDER = '#1D9E75'
GREEN_DARK   = '#0F6E56'
GREEN_DARKER = '#085041'
AMBER_BG     = '#FFF8E8'
AMBER_BORDER = '#EF9F27'

# Pre-blended solid equivalents of the original's translucent-white-on-navy
# and translucent-black-on-light-gray elements. PIL's ImageDraw silently
# drops the alpha byte when drawing on an 'RGB' canvas (no error, no effect),
# so anything relying on alpha must be pre-composited to a flat color instead.
NAVY_MUTED_1 = '#8c93a1'   # white @ ~50% over navy   (header subtitle)
NAVY_MUTED_2 = '#7e869c'   # white @ ~45% over navy   (percentile/rank labels)
NAVY_MUTED_3 = '#767f8f'   # white @ ~40% over navy   (watermark)
PCT_BOX_FILL = '#36415a'   # white @ 12% over navy    (percentile box bg)
PLOT_LABEL   = '#b2b2b0'   # black @ ~28% over #f7f7f5 (glove/arm plot labels)
STRIKE_ZONE  = '#d7d7d5'   # black @ ~13% over #f7f7f5 (plate strike-zone box)


class Canvas:
    """Grows a card image downward as sections are added."""
    def __init__(self, width_css):
        self.w = S(width_css)
        self.y = 0                       # current CSS-px cursor (unscaled)
        self.img = Image.new('RGB', (self.w, S(3000)), BG_WHITE)
        self.draw = ImageDraw.Draw(self.img)

    def ensure_height(self, extra_css_px):
        needed = S(self.y + extra_css_px)
        if needed > self.img.height:
            grown = Image.new('RGB', (self.w, needed + S(200)), BG_WHITE)
            grown.paste(self.img, (0, 0))
            self.img = grown
            self.draw = ImageDraw.Draw(self.img)

    def advance(self, css_px):
        self.ensure_height(css_px)
        self.y += css_px

    def finish(self):
        return self.img.crop((0, 0, self.w, S(self.y)))


def wrap_text(draw, text, font, max_width):
    words = text.split(' ')
    lines, cur = [], ''
    for w in words:
        trial = (cur + ' ' + w).strip()
        if text_w(draw, trial, font) <= max_width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def draw_stat_card(c, x, y, w, h, pct_val, lbl, raw, sub, invert=False):
    rounded_rect(c.draw, [S(x), S(y), S(x + w), S(y + h)], S(8), fill=hex_rgb(BG_SOFT))
    px, py = x + 14, y + 11
    if pct_val is None:
        draw_text(c.draw, (S(px), S(py)), '\u2014', F(24, True), hex_rgb('#aaaaaa'))
        py += 30
        draw_text(c.draw, (S(px), S(py)), lbl, F(13, True), hex_rgb(TEXT_DARK))
        py += 17
        draw_text(c.draw, (S(px), S(py)), 'No qualifying pairs', F(12), hex_rgb('#777777'))
        py += 15
    else:
        draw_text(c.draw, (S(px), S(py)), f'{pct_val}{sfx(pct_val)}', F(24, True), hex_rgb(stat_color(pct_val)))
        py += 30
        draw_text(c.draw, (S(px), S(py)), lbl, F(13, True), hex_rgb(TEXT_DARK))
        py += 17
        better = 'lower is better' if invert else 'higher is better'
        draw_text(c.draw, (S(px), S(py)), f'{raw} \u00b7 {better}', F(12), hex_rgb('#777777'))
        py += 15
    for line in wrap_text(c.draw, sub, F(11), S(w - 20) / 1.0):
        draw_text(c.draw, (S(px), S(py)), line, F(11), hex_rgb('#999999'))
        py += 13


def measure_pill_width(draw, name, val_text):
    dot_r = 4.5
    pad_l, pad_r = 9, 12
    gap = 6
    nw = text_w(draw, name, F(12, True)) / SCALE
    vw = text_w(draw, val_text, F(12)) / SCALE
    return pad_l + (dot_r * 2) + gap + nw + gap + vw + pad_r


def draw_pill(c, x, y, color, name, val_text):
    dot_r = 4.5
    pad_l, pad_r, pad_v = 9, 12, 5
    name_font, val_font = F(12, True), F(12)
    nw = text_w(c.draw, name, name_font) / SCALE
    vw = text_w(c.draw, val_text, val_font) / SCALE
    gap = 6
    content_w = (dot_r * 2) + gap + nw + gap + vw
    h = 22
    rounded_rect(c.draw, [S(x), S(y), S(x + pad_l + content_w + pad_r), S(y + h)],
                 S(h / 2), fill=hex_rgb(BG_SOFT), outline=hex_rgb(BORDER_SOFT), width=1)
    cx = x + pad_l + dot_r
    cy = y + h / 2
    c.draw.ellipse([S(cx - dot_r), S(cy - dot_r), S(cx + dot_r), S(cy + dot_r)], fill=hex_rgb(color))
    tx = x + pad_l + dot_r * 2 + gap
    draw_text(c.draw, (S(tx), S(cy)), name, name_font, hex_rgb(TEXT_DARK), anchor='lm')
    tx += nw + gap
    draw_text(c.draw, (S(tx), S(cy)), val_text, val_font, hex_rgb('#888888'), anchor='lm')
    return pad_l + content_w + pad_r  # total pill width


def draw_legend_dot(c, x, y, color, text):
    r = 4.5
    font = F(11)
    c.draw.ellipse([S(x), S(y - r), S(x + r * 2), S(y + r)], fill=hex_rgb(color))
    tx = x + r * 2 + 4
    draw_text(c.draw, (S(tx), S(y)), text, font, hex_rgb('#666666'), anchor='lm')
    return (r * 2 + 4) + text_w(c.draw, text, font) / SCALE


def draw_scatter_plot(c, x, y, w, h, points, mode, plt_mult=0.65):
    """mode: 'tunnel' uses tx/tz, 'plate' uses px/pz. Ported from PLOT_JS.

    Drawn onto its own sub-image sized exactly (w, h) and pasted onto the
    main canvas afterward. PIL's ImageDraw does not clip to any bounding
    region the way an HTML <canvas> does automatically — without this, a
    shape whose computed coordinates fall outside the intended plot area
    (e.g. the fixed-size strike-zone box against a tightly clustered pitch
    location) gets drawn wherever the math says, bleeding into whatever
    section happens to sit below it on the card.
    """
    sub = Image.new('RGB', (S(w), S(h)), hex_rgb(BG_SOFT))
    sd  = ImageDraw.Draw(sub)

    pad = 16 if mode == 'tunnel' else 10
    max_r = 24
    key_x = 'tx' if mode == 'tunnel' else 'px'
    key_z = 'tz' if mode == 'tunnel' else 'pz'

    xs = [p[key_x] for p in points]
    zs = [p[key_z] for p in points]
    x_min, x_max = min(xs), max(xs)
    z_min, z_max = min(zs), max(zs)

    if mode == 'tunnel':
        x_range = (x_max - x_min) + max_r * 2.5
        z_range = (z_max - z_min) + max_r * 2.5
        scale = min((w - 2 * pad) / x_range, (h - 2 * pad) / z_range)
    else:
        # plt_mult mirrors the original: mult = round(0.35 + (pct/100)*0.65, 4)
        # — a percentile-scaled zoom level, not a fixed constant.
        eff_x = (x_max - x_min) / plt_mult if (x_max - x_min) else 1
        eff_z = (z_max - z_min) / plt_mult if (z_max - z_min) else 1
        scale = min((w - 2 * pad) / eff_x, (h - 2 * pad) / eff_z) * 0.88

    x_mid, z_mid = (x_max + x_min) / 2, (z_max + z_min) / 2

    def to_canvas(px_, pz_):
        # coordinates relative to the sub-image's own (0,0), not the card
        cx = w / 2 + (px_ - x_mid) * scale
        cy = h / 2 - (pz_ - z_mid) * scale
        return cx, cy

    if mode == 'plate':
        tl = to_canvas(-8.5, 44)
        br = to_canvas(8.5, 18)
        if br[0] > tl[0] and br[1] > tl[1]:
            # clamp to the sub-image bounds — belt-and-suspenders alongside
            # the paste-based clipping, in case of extreme data
            rx0 = max(0, min(tl[0], w))
            ry0 = max(0, min(tl[1], h))
            rx1 = max(0, min(br[0], w))
            ry1 = max(0, min(br[1], h))
            if rx1 > rx0 and ry1 > ry0:
                sd.rectangle([S(rx0), S(ry0), S(rx1), S(ry1)],
                             outline=hex_rgb(STRIKE_ZONE), width=max(1, S(0.5)))

    for p in points:
        cx, cy = to_canvas(p[key_x], p[key_z])
        r = max(12, min(max_r, p['frac'] * 0.62)) if mode == 'tunnel' else max(14, min(22, p['frac'] * 0.50))
        col = hex_rgb(p['c'])
        sd.ellipse([S(cx - r), S(cy - r), S(cx + r), S(cy + r)], fill=col, outline=hex_rgb(p['c']), width=S(1.2))
        f = F(max(9, min(12, r)), True)
        draw_text(sd, (S(cx), S(cy)), p['t'], f, (255, 255, 255), anchor='mm')

    fnote = F(9)
    draw_text(sd, (S(pad + 2), S(h - 4)), '\u2190 glove', fnote, hex_rgb(PLOT_LABEL), anchor='ls')
    draw_text(sd, (S(w - pad - 2), S(h - 4)), 'arm \u2192', fnote, hex_rgb(PLOT_LABEL), anchor='rs')

    # Paste the finished sub-plot onto the main canvas — this is what actually
    # clips anything drawn outside (0,0)-(w,h) above, fixing the overflow bug.
    c.img.paste(sub, (S(x), S(y)))


def draw_bar(c, x, y, w, h, pct_val, color):
    rounded_rect(c.draw, [S(x), S(y), S(x + w), S(y + h)], S(h / 2), fill=hex_rgb('#e0e0e0'))
    fill_w = max(h, w * (pct_val / 100))
    rounded_rect(c.draw, [S(x), S(y), S(x + fill_w), S(y + h)], S(h / 2), fill=hex_rgb(color))


# ── Team full names + hand strings (must match app.py's tables) ─────────────
def render_pitcher_card(info, team_full, hand_str):
    """
    Draws the full Tunneling+ player card and returns JPEG bytes.
    `info` is the exact dict produced by get_pitcher_card_info() in app.py.
    `team_full` / `hand_str` are the TEAM_FULL / HAND_STR lookup dicts from app.py.
    """
    c = Canvas(CARD_W)
    P = info['pitches_out']
    col = {p['t']: p['c'] for p in P}
    pct = info['tp_pct']

    # ── Header ────────────────────────────────────────────────────────────────
    # Height matches the original card exactly (measured at 163 CSS px) — no
    # watermark line, which the current live design doesn't include.
    hdr_top_y = c.y
    hdr_h = 163
    c.advance(hdr_h)
    rounded_rect(c.draw, [0, S(hdr_top_y), c.w, S(hdr_top_y + hdr_h)], 0, fill=hex_rgb(NAVY))

    pct_box_w, pct_box_h = 90, 48
    pbx = CARD_W - PAD_H - pct_box_w
    pby = hdr_top_y + PAD_V

    tx, ty = PAD_H, hdr_top_y + PAD_V
    name_max_w = pbx - 16 - tx  # leave a gap before the percentile box
    name_size = 22
    while name_size > 12 and text_w(c.draw, info['name'], F(name_size)) > S(name_max_w):
        name_size -= 1
    draw_text(c.draw, (S(tx), S(ty)), info['name'], F(name_size), (255, 255, 255))
    hand_full = hand_str.get(info['hand'], 'RHP')
    team_name = team_full.get(info['team'], info['team'])
    sub = f"{hand_full} \u00b7 {team_name} \u00b7 2026 \u00b7 {info['pitches']} pitches"
    draw_text(c.draw, (S(tx), S(ty + 27)), sub, F(12), hex_rgb(NAVY_MUTED_1))
    rounded_rect(c.draw, [S(pbx), S(pby), S(pbx + pct_box_w), S(pby + pct_box_h)], S(10), fill=hex_rgb(PCT_BOX_FILL))
    draw_text(c.draw, (S(pbx + pct_box_w / 2), S(pby + 10)), f'{pct}th', F(20), (255, 255, 255), anchor='ma')
    draw_text(c.draw, (S(pbx + pct_box_w / 2), S(pby + 32)), 'percentile', F(11), hex_rgb(NAVY_MUTED_2), anchor='ma')

    center_y = hdr_top_y + PAD_V + 34
    tp_str = f"{info['tplus']:.1f}" if isinstance(info['tplus'], float) else str(info['tplus'])
    draw_text(c.draw, (S(CARD_W / 2), S(center_y)), tp_str, F(38), hex_rgb(BLUE_ACCENT), anchor='ma')
    draw_text(c.draw, (S(CARD_W / 2), S(center_y + 46)), 'TUNNELING+', F(11), hex_rgb(NAVY_MUTED_2), anchor='ma')
    total_str = f" of {info['total']}" if info.get('total') else ''
    draw_text(c.draw, (S(CARD_W / 2), S(center_y + 66)), f"Rank #{info['rank']}{total_str} qualified pitchers",
               F(12), hex_rgb(NAVY_MUTED_2), anchor='ma')

    # ── Body ──────────────────────────────────────────────────────────────────
    body_x = PAD_H
    content_w = CARD_W - 2 * PAD_H
    c.advance(16)  # body top padding

    def section_label(text, subtitle=None):
        draw_text(c.draw, (S(body_x), S(c.y)), text.upper(), F(12, True), hex_rgb('#888888'))
        if subtitle:
            main_w = text_w(c.draw, text.upper(), F(12, True)) / SCALE
            draw_text(c.draw, (S(body_x + main_w + 6), S(c.y + 1)), subtitle, F(11), hex_rgb('#aaaaaa'))
        c.advance(20)

    def divider():
        c.advance(6)
        c.draw.line([S(body_x), S(c.y), S(body_x + content_w), S(c.y)], fill=hex_rgb('#eeeeee'), width=S(0.5))
        c.advance(14)

    # Component scores
    section_label('Component scores')
    gap = 10
    stat_w = (content_w - 2 * gap) / 3
    stat_h = 92
    stat1 = ('Tunnel ratio', info['tr_pct'], f"{info['tr_val']}x" if info['tr_val'] else '\u2014',
             'Plate spread relative to tunnel tightness \u2014 bigger late break after the commit point', False)
    stat2 = ('Speed-change deception', info['tm_pct'], str(info['tm_val']) if info['tm_val'] else '\u2014',
             'Same plate location, different speed \u2014 hitter commits before detecting the difference', False)
    stat3 = ('Release consistency', info['rc_pct'], f"{info['rc_val']}\"" if info['rc_val'] else '\u2014',
             'How similar his release looks across pitch types \u2014 tighter means less tip-off', True)
    for i, (lbl, pv, raw, sub_txt, inv) in enumerate([stat1, stat2, stat3]):
        sx = body_x + i * (stat_w + gap)
        draw_stat_card(c, sx, c.y, stat_w, stat_h, pv, lbl, raw, sub_txt, invert=inv)
    c.advance(stat_h)
    divider()

    # Pitch arsenal
    section_label('Pitch arsenal')
    row_x = body_x
    max_row_w = content_w
    row_gap = 7
    for p in P:
        val_text = f"{p['frac']}% \u00b7 {p['velo']} mph"
        pill_w = measure_pill_width(c.draw, p['t'], val_text)
        if row_x != body_x and row_x + pill_w > body_x + max_row_w:
            row_x = body_x
            c.advance(29)
        draw_pill(c, row_x, c.y, p['c'], p['t'], val_text)
        row_x += pill_w + row_gap
    c.advance(30)
    divider()

    # Pitch location plots
    section_label("Pitch location \u2014 catcher's view")
    plot_w = (content_w - 1) / 2
    plot_h = 130
    box_y = c.y
    rounded_rect(c.draw, [S(body_x), S(box_y), S(body_x + content_w), S(box_y + plot_h + 24)], S(8),
                 outline=hex_rgb(BORDER_LT), width=S(0.5))
    title_font = F(11, True)
    draw_text(c.draw, (S(body_x + plot_w / 2), S(box_y + 8)), 'AT TUNNEL POINT \u00b7 23.8FT \u2014 PITCHES CONVERGE',
               title_font, hex_rgb('#888888'), anchor='ma')
    draw_text(c.draw, (S(body_x + plot_w + plot_w / 2), S(box_y + 8)), 'AT THE PLATE \u2014 PITCHES DIVERGE',
               title_font, hex_rgb('#888888'), anchor='ma')
    plots_top = box_y + 20
    plt_mult = round(0.35 + (pct / 100) * 0.65, 4)
    draw_scatter_plot(c, body_x, plots_top, plot_w, plot_h - 20, P, 'tunnel')
    draw_scatter_plot(c, body_x + plot_w, plots_top, plot_w, plot_h - 20, P, 'plate', plt_mult=plt_mult)
    c.y = box_y + plot_h + 4
    c.ensure_height(20)
    leg_x = body_x
    leg_y = c.y + 10
    for p in P:
        w_used = draw_legend_dot(c, leg_x, leg_y, p['c'], f"{p['t']} IVB {'+' if p['ivb'] >= 0 else ''}{p['ivb']}\"")
        leg_x += w_used + 10
    c.y = box_y + plot_h + 24 + 12
    divider()

    # Tunnel pairs
    tp = info['tp_top3']
    top_n = min(3, len(tp))
    if top_n > 0:
        section_label(f"Top {top_n} tunnel pair{'s' if top_n > 1 else ''}",
                       subtitle=f"\u2014 best of {info['n_tp_total']} qualifying. Bar = percentile vs all MLB tunnel pairs.")
        n_cols = 3 if top_n == 3 else 2
        pair_gap = 8
        pair_w = (content_w - (n_cols - 1) * pair_gap) / n_cols
        pair_h = 92
        for i, p in enumerate(tp[:top_n]):
            px_ = body_x + i * (pair_w + pair_gap)
            py_ = c.y
            rounded_rect(c.draw, [S(px_), S(py_), S(px_ + pair_w), S(py_ + pair_h)], S(8),
                         fill=hex_rgb(BG_SOFT), outline=hex_rgb(BORDER_SOFT), width=S(0.5))
            ipx, ipy = px_ + 12, py_ + 10
            c1 = col.get(p['p1'], '#888888')
            c2 = col.get(p['p2'], '#888888')
            n1w = text_w(c.draw, p['p1'], F(14, True))
            draw_text(c.draw, (S(ipx), S(ipy)), p['p1'], F(14, True), hex_rgb(c1))
            sep_x = ipx + n1w / SCALE + 2
            draw_text(c.draw, (S(sep_x), S(ipy)), ' / ', F(13), hex_rgb('#cccccc'))
            sepw = text_w(c.draw, ' / ', F(13))
            draw_text(c.draw, (S(sep_x + sepw / SCALE), S(ipy)), p['p2'], F(14, True), hex_rgb(c2))
            ipy += 20
            meta = f"{p['td']:.1f}\" apart at tunnel"
            draw_text(c.draw, (S(ipx), S(ipy)), meta, F(11), hex_rgb('#999999'))
            ipy += 13
            draw_text(c.draw, (S(ipx), S(ipy)), f"{p['pd']:.1f}\" apart at plate", F(11), hex_rgb('#999999'))
            ipy += 20
            bar_w = pair_w - 24 - 46
            draw_bar(c, ipx, ipy + 3, bar_w, 6, p['pct'], bar_color(p['pct']))
            draw_text(c.draw, (S(ipx + bar_w + 8), S(ipy - 5)), f"{p['pct']}th", F(15, True),
                       hex_rgb(pct_color(p['pct'])), anchor='la')
            ipy += 18
            draw_text(c.draw, (S(ipx), S(ipy)), f"Tunnel ratio {p['ratio']:.2f}x", F(11), hex_rgb('#999999'))
        c.advance(pair_h)
    else:
        section_label('No tunnel pairs')
        rounded_rect(c.draw, [S(body_x), S(c.y), S(body_x + content_w), S(c.y + 60)], S(8),
                     fill=hex_rgb(BG_SOFT), outline=hex_rgb(BORDER_SOFT), width=S(0.5))
        draw_text(c.draw, (S(CARD_W / 2), S(c.y + 22)), 'No qualifying tunnel pairs', F(13, True),
                   hex_rgb('#aaaaaa'), anchor='ma')
        draw_text(c.draw, (S(CARD_W / 2), S(c.y + 40)), 'No pitch pairs met the tunnel classification threshold',
                   F(11), hex_rgb('#bbbbbb'), anchor='ma')
        c.advance(60)

    # Speed-change pair
    c.advance(14)
    sp = info['sp_best']
    if sp:
        c1 = col.get(sp['p1'], '#888888')
        c2 = col.get(sp['p2'], '#888888')
        text_col_w = content_w - 190  # leave room for the right-side percentile block
        meta1 = f"Both pitches arrive {sp['pd']:.1f}\" apart"
        meta2 = f" \u2014 similar location \u2014 but with a {sp['vd']:.1f} mph speed gap."
        full_meta = meta1 + meta2 + ' The hitter commits at 23.8ft before the difference is detectable.'
        meta_lines = wrap_text(c.draw, full_meta, F(12), S(text_col_w))
        # header(13) + names row(24) + wrapped lines(16 each) + bar gap(10) + bottom pad(14)
        card_h = max(84, 13 + 24 + len(meta_lines) * 16 + 24)

        rounded_rect(c.draw, [S(body_x), S(c.y), S(body_x + content_w), S(c.y + card_h)], S(8),
                     fill=hex_rgb(GREEN_BG), outline=hex_rgb(GREEN_BORDER), width=S(0.5))
        ipx, ipy = body_x + 16, c.y + 13
        draw_text(c.draw, (S(ipx), S(ipy)), f"Best speed-change pair \u00b7 {sp['p1']} / {sp['p2']}",
                   F(12, True), hex_rgb(GREEN_DARK))
        ipy += 18
        n1w = text_w(c.draw, sp['p1'], F(15, True))
        draw_text(c.draw, (S(ipx), S(ipy)), sp['p1'], F(15, True), hex_rgb(c1))
        sep_x = ipx + n1w / SCALE + 2
        draw_text(c.draw, (S(sep_x), S(ipy)), ' / ', F(14), hex_rgb('#bbbbbb'))
        sepw = text_w(c.draw, ' / ', F(14))
        draw_text(c.draw, (S(sep_x + sepw / SCALE), S(ipy)), sp['p2'], F(15, True), hex_rgb(c2))
        ipy += 24
        for line in meta_lines:
            draw_text(c.draw, (S(ipx), S(ipy)), line, F(12), hex_rgb('#444444'))
            ipy += 16
        draw_bar(c, ipx, ipy + 4, text_col_w, 6, sp['pct'], GREEN_DARK)

        rx = body_x + content_w - 16
        draw_text(c.draw, (S(rx), S(c.y + 16)), f"{sp['pct']}th", F(26, True), hex_rgb(GREEN_DARK), anchor='ra')
        draw_text(c.draw, (S(rx), S(c.y + 42)), 'percentile', F(11), hex_rgb(GREEN_DARKER), anchor='ra')
        draw_text(c.draw, (S(rx), S(c.y + 58)), f"score {sp['score']:.2f}", F(12), hex_rgb(GREEN_DARKER), anchor='ra')
        c.advance(card_h)
    else:
        card_h = 54
        rounded_rect(c.draw, [S(body_x), S(c.y), S(body_x + content_w), S(c.y + card_h)], S(8),
                     fill=hex_rgb(BG_SOFT), outline=hex_rgb(BORDER_SOFT), width=S(0.5))
        draw_text(c.draw, (S(CARD_W / 2), S(c.y + 15)), 'No speed-change pairs', F(13, True),
                   hex_rgb('#aaaaaa'), anchor='ma')
        draw_text(c.draw, (S(CARD_W / 2), S(c.y + 32)), 'No pitch pairs with similar plate location and meaningful velocity gap',
                   F(11), hex_rgb('#bbbbbb'), anchor='ma')
        c.advance(card_h)

    if info.get('fallback'):
        c.advance(14)
        note = ('Note: No pairs met the tunnel or speed-change classification thresholds. '
                'Scored using unclassified pair metrics.')
        card_h = 50
        rounded_rect(c.draw, [S(body_x), S(c.y), S(body_x + content_w), S(c.y + card_h)], S(8),
                     fill=hex_rgb(AMBER_BG), outline=hex_rgb(AMBER_BORDER), width=S(0.5))
        ipy = c.y + 12
        for line in wrap_text(c.draw, note, F(11), S(content_w - 32) / 1.0):
            draw_text(c.draw, (S(body_x + 16), S(ipy)), line, F(11), hex_rgb('#5C3A00'))
            ipy += 15
        c.advance(card_h)

    c.advance(16)  # body bottom padding

    final = c.finish()

    # ── Apply rounded outer corners on a light page background ─────────────────
    radius = S(12)
    mask = Image.new('L', final.size, 0)
    mdraw = ImageDraw.Draw(mask)
    mdraw.rounded_rectangle([0, 0, final.width - 1, final.height - 1], radius=radius, fill=255)

    bg = Image.new('RGB', final.size, hex_rgb('#f0f0f0'))
    bg.paste(final, (0, 0), mask)

    # ── Downsample from the 3x supersampled draw to 1.5x final resolution.
    # PIL's font/shape anti-aliasing looks noticeably softer than a browser's
    # hinted text rendering at 1:1; drawing everything larger and downsampling
    # with a high-quality filter sharpens edges the same way a 2x screenshot
    # scaled down looks crisper than one rendered at native size. ──────────────
    target_w = round(bg.width * (FINAL_SCALE / SCALE))
    target_h = round(bg.height * (FINAL_SCALE / SCALE))
    bg = bg.resize((target_w, target_h), Image.LANCZOS)

    buf = io.BytesIO()
    bg.save(buf, format='JPEG', quality=95)
    return buf.getvalue()
