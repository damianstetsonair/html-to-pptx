#!/usr/bin/env python3
"""
HTML-to-PPTX converter.
Converts a presentation rendered as HTML slides into a faithful python-pptx file.

Slide types handled:
  1. Summary   – standalone table + legend
  2. Detail    – 8 section boxes (description, scope, progress, next steps,
                 decisions, risks, trends, budget)
  3. Budget    – budget detail table + project team + link
  4. Planning  – progress bar, milestones list, workload table
  5. Notes     – plain text data notes
"""

import re, sys
from typing import Optional, List, Tuple

from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from lxml import etree, html as lxml_html

# ───────────────────────── constants ──────────────────────────
SLIDE_W_PX, SLIDE_H_PX = 960, 540
SLIDE_W_IN, SLIDE_H_IN = 10.0, 5.625
_SCALE = SLIDE_W_IN / SLIDE_W_PX  # inches per CSS pixel

TEAL     = RGBColor(0x00, 0x62, 0x72)
TEAL_BAR = RGBColor(0x7A, 0x9A, 0x9E)
RED_BAR  = RGBColor(0xC4, 0x1E, 0x3A)
RED_BULL = RGBColor(0xCC, 0x00, 0x00)
WHITE    = RGBColor(0xFF, 0xFF, 0xFF)
BLACK33  = RGBColor(0x33, 0x33, 0x33)
GREY66   = RGBColor(0x66, 0x66, 0x66)
GREY_CC  = RGBColor(0xCC, 0xCC, 0xCC)
GREY_E5  = RGBColor(0xE5, 0xE5, 0xE5)
FONT     = 'Arial'

# ───────────────────────── helpers ────────────────────────────
def E(px: float) -> int:
    """CSS pixels → EMU."""
    return int(round(px * _SCALE * 914400))

def _parse_color(s: str) -> Optional[RGBColor]:
    if not s:
        return None
    s = s.strip().lower()
    _MAP = {
        '#7a9a9e': TEAL_BAR, '#006272': TEAL, '#c41e3a': RED_BAR,
        '#c00': RED_BULL, '#333': BLACK33, '#333333': BLACK33,
        '#666': GREY66, '#666666': GREY66, '#ccc': GREY_CC,
        '#cccccc': GREY_CC, '#fff': WHITE, '#ffffff': WHITE,
        '#f5f5f5': RGBColor(0xF5,0xF5,0xF5), '#f9f9f9': RGBColor(0xF9,0xF9,0xF9),
        '#e5e5e5': GREY_E5, '#22c55e': RGBColor(0x22,0xC5,0x5E),
        '#f59e0b': RGBColor(0xF5,0x9E,0x0B), '#ef4444': RGBColor(0xEF,0x44,0x44),
        'white': WHITE,
    }
    if s in _MAP:
        return _MAP[s]
    m = re.match(r'#([0-9a-f]{6})$', s)
    if m:
        h = m.group(1); return RGBColor(int(h[0:2],16), int(h[2:4],16), int(h[4:6],16))
    m = re.match(r'#([0-9a-f]{3})$', s)
    if m:
        h = m.group(1); return RGBColor(int(h[0]*2,16), int(h[1]*2,16), int(h[2]*2,16))
    m = re.match(r'rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)', s)
    if m:
        return RGBColor(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None

def _sty(el) -> dict:
    """Parse inline style attribute → dict."""
    d = {}
    for part in (el.get('style', '') or '').split(';'):
        if ':' not in part: continue
        k, v = part.split(':', 1)
        d[k.strip().lower()] = v.strip()
    return d

def _px(v: str) -> float:
    m = re.match(r'([\d.]+)', v.strip()); return float(m.group(1)) if m else 0

def _pct(v: str) -> float:
    m = re.match(r'([\d.]+)\s*%', v.strip()); return float(m.group(1)) if m else 0

# ───────────── XML-level table cell helpers ───────────────────
def _cell_border(cell, color=GREY_CC, width=6350, dash='solid'):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for tag_name in ('a:lnT','a:lnB','a:lnL','a:lnR'):
        tag = qn(tag_name)
        for old in tcPr.findall(tag): tcPr.remove(old)
        ln = etree.SubElement(tcPr, tag)
        ln.set('w', str(int(width))); ln.set('cap','flat'); ln.set('cmpd','sng'); ln.set('algn','ctr')
        sf = etree.SubElement(ln, qn('a:solidFill'))
        srgb = etree.SubElement(sf, qn('a:srgbClr'))
        srgb.set('val', '%02X%02X%02X' % (color[0], color[1], color[2]))
        if dash == 'dashed':
            etree.SubElement(ln, qn('a:prstDash')).set('val', 'dash')

def _cell_fill(cell, color: RGBColor):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('a:solidFill')): tcPr.remove(old)
    sf = etree.SubElement(tcPr, qn('a:solidFill'))
    etree.SubElement(sf, qn('a:srgbClr')).set('val', '%02X%02X%02X' % (color[0], color[1], color[2]))
    tcPr.insert(0, sf)

def _nuke_table_theme(shape):
    """Remove built-in table theme so manual cell fills/borders show."""
    tbl = shape._element.find('.//' + qn('a:tbl'))
    if tbl is None: return
    tblPr = tbl.find(qn('a:tblPr'))
    if tblPr is None: return
    for k in ('bandRow','bandCol','firstRow','lastRow','firstCol','lastCol'):
        tblPr.set(k, '0')
    for child_tag in (qn('a:tblStyle'), qn('a:tableStyleId')):
        for ch in tblPr.findall(child_tag): tblPr.remove(ch)

# ───────────── pptx shape factories ──────────────────────────
def _rect(slide, left, top, w, h, fill=None, line_color=None, line_w=None):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, E(left), E(top), E(w), E(h))
    if fill:
        s.fill.solid(); s.fill.fore_color.rgb = fill
    else:
        s.fill.background()
    if line_color:
        s.line.color.rgb = line_color
        if line_w: s.line.width = line_w
    else:
        s.line.fill.background()
    return s

def _oval(slide, left, top, size, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, E(left), E(top), E(size), E(size))
    s.fill.solid(); s.fill.fore_color.rgb = fill; s.line.fill.background()
    return s

def _textbox(slide, left, top, w, h, text='', size=8, bold=False, color=BLACK33,
             align=PP_ALIGN.LEFT, font=FONT, wrap=True, valign='top'):
    tb = slide.shapes.add_textbox(E(left), E(top), E(w), E(h))
    tf = tb.text_frame; tf.word_wrap = wrap; tf.auto_size = None
    # eliminate internal insets so text isn't clipped in small boxes
    bp = tf._txBody.find(qn('a:bodyPr'))
    if bp is not None:
        bp.set('lIns', '0'); bp.set('tIns', '0')
        bp.set('rIns', '0'); bp.set('bIns', '0')
        if valign == 'ctr':
            bp.set('anchor', 'ctr')
    p = tf.paragraphs[0]; p.alignment = align
    p.space_before = Pt(0); p.space_after = Pt(0)
    if text:
        r = p.add_run(); r.text = text
        r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = font
    return tb

def _add_run(paragraph, text, size=8, bold=False, color=BLACK33, font=FONT):
    r = paragraph.add_run(); r.text = text
    r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = color; r.font.name = font
    return r

# ───────────── rich inline text rendering ─────────────────────
def _render_rich(paragraph, el, pt=8, skip_blocks=False):
    """Walk *el* children and emit runs with bold / colour.
    If skip_blocks=True, <ul> and <div> children are ignored (only their tail is kept).
    """
    parts: List[Tuple] = []  # (kind, text [, color])
    if el.text:
        parts.append(('n', el.text))
    for sub in el:
        if skip_blocks and sub.tag in ('ul', 'div', 'table'):
            if sub.tail: parts.append(('n', sub.tail))
            continue
        if sub.tag in ('strong', 'b'):
            parts.append(('b', sub.text_content()))
        elif sub.tag == 'span':
            ss = _sty(sub)
            # coloured circle indicator → render as inline ● with colour
            bg = ss.get('background', '')
            if bg and ('border-radius' in str(ss) or 'display' in ss):
                cc = _parse_color(bg)
                if cc:
                    parts.append(('c', '●', cc))
                if sub.tail: parts.append(('n', sub.tail))
                continue
            txt = sub.text_content().strip()
            if txt == '●':
                c = _parse_color(ss.get('color', ''))
                if c:
                    parts.append(('c', '●', c))
                if sub.tail: parts.append(('n', sub.tail))
                continue
            if txt == '':
                if sub.tail: parts.append(('n', sub.tail))
                continue
            c = _parse_color(ss.get('color', ''))
            parts.append(('c', sub.text_content(), c))
        elif sub.tag == 'a':
            parts.append(('n', sub.text_content()))
        else:
            parts.append(('n', sub.text_content()))
        if sub.tail:
            parts.append(('n', sub.tail))

    if not parts:
        clean = el.text_content().strip()
        if clean:
            _add_run(paragraph, clean, pt)
        return

    for p in parts:
        txt = p[1]
        if skip_blocks:
            txt = txt.strip('\n')
        if not txt.strip():
            continue
        if p[0] == 'b':
            _add_run(paragraph, txt, pt, bold=True)
        elif p[0] == 'c' and len(p) > 2 and p[2]:
            _add_run(paragraph, txt, pt, color=p[2])
        else:
            _add_run(paragraph, txt, pt)

# ───────────── detect circle indicator colour ─────────────────
def _circle_color(el) -> Optional[RGBColor]:
    for span in el.cssselect('span[style]'):
        ss = _sty(span)
        bg = ss.get('background', '')
        if bg: return _parse_color(bg)
        if ss.get('color') and '●' in (span.text or ''):
            return _parse_color(ss.get('color', ''))
    return None

# ═══════════════════════════════════════════════════════════════
#  MAIN RENDERER
# ═══════════════════════════════════════════════════════════════
class SlideRenderer:
    def __init__(self, pptx_slide, html_el):
        self.s = pptx_slide
        self.el = html_el

    def render(self):
        self._chrome()           # top bar, date box, title, footer
        self._positioned_blocks() # all absolute-positioned content
        self._legend()
        self._links()

    # ── chrome ─────────────────────────────────────────────────
    def _chrome(self):
        # top bar
        if self.el.cssselect('.top-bar'):
            _rect(self.s, 0, 0, SLIDE_W_PX, 8, fill=TEAL_BAR)
        # date box
        dbs = self.el.cssselect('.date-box')
        if dbs:
            _rect(self.s, SLIDE_W_PX-100, 8, 100, 50, fill=TEAL_BAR)
            _textbox(self.s, SLIDE_W_PX-100, 8, 100, 50,
                     dbs[0].text_content().strip(), size=10, bold=True,
                     color=WHITE, align=PP_ALIGN.CENTER, valign='ctr')
        # title
        titles = self.el.cssselect('.main-title')
        if titles:
            t = titles[0]; st = _sty(t)
            fs = _px(st.get('font-size', '42'))
            mw = _px(st.get('max-width', '800')) if 'max-width' in st else 800
            _textbox(self.s, 30, 20, mw, fs*1.4,
                     t.text_content().strip(), size=fs*0.75, bold=True, color=TEAL)
        # footer
        fbs = self.el.cssselect('.footer-bar')
        if fbs:
            fb = fbs[0]
            _rect(self.s, 0, SLIDE_H_PX-32, SLIDE_W_PX, 32, fill=RED_BAR)
            pn = fb.cssselect('.page-number')
            if pn:
                _textbox(self.s, 20, SLIDE_H_PX-32, 100, 32,
                         pn[0].text_content().strip(), size=10, color=WHITE, valign='ctr')
            lg = fb.cssselect('.logo')
            if lg:
                _textbox(self.s, SLIDE_W_PX-140, SLIDE_H_PX-32, 120, 32,
                         lg[0].text_content().strip(), size=13, bold=True,
                         color=WHITE, align=PP_ALIGN.RIGHT, valign='ctr')

    # ── content dispatcher ─────────────────────────────────────
    def _positioned_blocks(self):
        for div in self.el.cssselect('div[style]'):
            st = _sty(div)
            if st.get('position') != 'absolute':
                continue
            # skip things handled elsewhere
            if div.cssselect('.footer-bar'):                          continue
            if 'Legend' in div.text_content() and div.cssselect('strong'): continue
            if div.cssselect('a.link-text') and not div.cssselect('.section-header'): continue

            has_section = bool(div.cssselect('.section-header'))
            has_table   = bool(div.cssselect('table'))

            if has_section:
                sec_name = ''
                t = div.cssselect('.section-title')
                if t: sec_name = t[0].text_content().strip()
                if sec_name == 'PLANNING':
                    self._section_chrome(div, st)
                    self._planning(div, st)
                elif sec_name == 'TEAM WORKLOAD IN M/D':
                    self._section_chrome(div, st)
                    self._workload_table(div, st)
                else:
                    self._section_full(div, st)
            elif has_table:
                self._standalone_table(div, st)

    # ── section header + box outline ───────────────────────────
    def _section_chrome(self, div, st):
        top, left, w = _px(st.get('top','0')), _px(st.get('left','0')), _px(st.get('width','420'))
        _rect(self.s, left, top, w, 1, fill=GREY_CC)
        titles = div.cssselect('.section-title')
        if titles:
            _textbox(self.s, left, top+2, w, 16,
                     titles[0].text_content().strip(), size=9, bold=True, color=TEAL)
        box = div.cssselect('.section-box')
        if box:
            bh = _px(_sty(box[0]).get('height','80'))
            _rect(self.s, left, top+20, w, bh, fill=WHITE, line_color=GREY_CC, line_w=Pt(0.75))

    # ── full section (header + box + content) ──────────────────
    def _section_full(self, div, st):
        self._section_chrome(div, st)
        top, left, w = _px(st.get('top','0')), _px(st.get('left','0')), _px(st.get('width','420'))
        box = div.cssselect('.section-box')
        if not box: return
        box = box[0]
        bh = _px(_sty(box).get('height','80'))
        box_top = top + 20

        # if box contains a table → render table, done
        tables = box.cssselect('table')
        if tables:
            self._render_table(tables[0], left+2, box_top+2, w-4,
                               dashed='workload' in (tables[0].get('class','') or ''))
            return

        # trend-box
        trend = box.cssselect('.trend-box')
        if trend:
            x = left + 8
            for item in trend[0].cssselect('.trend-item'):
                _textbox(self.s, x, box_top+10, 80, 20,
                         item.text_content().strip(), size=10, bold=True)
                x += 100
            return

        # render content recursively
        y = box_top + 6
        y = self._render_box_content(box, left, y, w)

    # ── recursive box content renderer ────────────────────────
    _INLINE_TAGS = {'strong', 'b', 'em', 'i', 'span', 'a', 'br', 'sub', 'sup'}

    def _render_box_content(self, parent, left, y, w, indent=0, fs_pt=8):
        """Walk *parent* children and render blocks: div, p, ul/li, etc.
        Inline tags (strong, span, etc.) are skipped — they are handled by _render_rich.
        Respects margin-top, margin-bottom, margin-left, line-height from CSS."""
        x_base = left + 8 + indent
        w_inner = w - 16 - indent
        LINE_H = 13  # default line height in px

        for child in parent:
            if child.tag in self._INLINE_TAGS:
                continue
            tag = child.tag
            cls = child.get('class', '') or ''
            cst = _sty(child)

            # read CSS margins
            mt = _px(cst.get('margin-top', '0'))
            mb = _px(cst.get('margin-bottom', '0'))
            y += mt

            # line-height → per-line spacing (1.6 → 13*1.6 ≈ 21)
            lh_str = cst.get('line-height', '')
            lh = LINE_H
            if lh_str:
                lh_val = _px(lh_str)
                if lh_val > 0 and lh_val < 5:  # unitless multiplier like 1.6
                    lh = int(LINE_H * lh_val)
                elif lh_val >= 5:               # px value
                    lh = int(lh_val)

            # budget-label
            if 'budget-label' in cls:
                _textbox(self.s, x_base, y, w_inner, 14,
                         child.text_content().strip(), size=9, bold=True)
                y += 14 + mb; continue

            # bullet-item
            if 'bullet-item' in cls:
                fs_css = _px(cst.get('font-size','11')) if 'font-size' in cst else 11
                fpt = fs_css * 0.75
                item_mb = _px(cst.get('margin-bottom', '4'))
                _textbox(self.s, x_base, y, 10, 12, '▪', size=7, color=RED_BULL)
                tb = _textbox(self.s, x_base+12, y, w_inner-12, 14)
                _render_rich(tb.text_frame.paragraphs[0], child, fpt)
                y += max(14, int(fpt*1.8)) + item_mb; continue

            # <ul> → render each <li> with bullet, respecting margin-left
            if tag == 'ul':
                ul_margin = _px(cst.get('margin-left', '0'))
                li_x = x_base + ul_margin
                li_w = w_inner - ul_margin
                for li in child:
                    if li.tag != 'li': continue
                    txt = li.text_content().strip()
                    if not txt: continue
                    _textbox(self.s, li_x, y, 10, lh, '\u2022', size=8, color=BLACK33)
                    tb = _textbox(self.s, li_x+12, y, li_w-12, lh)
                    _render_rich(tb.text_frame.paragraphs[0], li, fs_pt)
                    y += lh
                y += mb
                continue

            # <p> → render with rich text
            if tag == 'p':
                txt = child.text_content().strip()
                if txt:
                    tb = _textbox(self.s, x_base, y, w_inner, 14)
                    _render_rich(tb.text_frame.paragraphs[0], child, fs_pt)
                    y += 14
                y += mb
                continue

            # <div> with children → recurse into it
            if tag == 'div':
                has_blocks = any(c.tag in ('p', 'ul', 'div', 'table') for c in child)
                if has_blocks:
                    child_fs = fs_pt
                    fs_str = cst.get('font-size', '')
                    if fs_str:
                        child_fs = _px(fs_str) * 0.75
                    child_indent = indent + _px(cst.get('margin-left', '0'))
                    y = self._render_box_content(child, left, y, w, child_indent, child_fs)
                else:
                    txt = child.text_content().strip()
                    if txt:
                        tb = _textbox(self.s, x_base, y, w_inner, 14)
                        _render_rich(tb.text_frame.paragraphs[0], child, fs_pt)
                        y += 14
                y += mb
                continue

        return y

    # ── planning section ──────────────────────────────────────
    def _planning(self, div, st):
        top, left, w = _px(st.get('top','0')), _px(st.get('left','0')), _px(st.get('width','900'))
        box = div.cssselect('.section-box')
        if not box: return
        box = box[0]
        y = top + 28

        for child in box:
            # progress bar container
            progress_bg = child.cssselect('div[style]')
            drew_bar = False
            for inner in progress_bg:
                iss = _sty(inner)
                if iss.get('background') != '#e5e5e5': continue
                drew_bar = True
                bar_w = w - 24
                # bg bar
                bg = self.s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                        E(left+12), E(y), E(bar_w), E(16))
                bg.fill.solid(); bg.fill.fore_color.rgb = GREY_E5; bg.line.fill.background()
                # fill bar
                for fd in inner.cssselect('div[style]'):
                    fds = _sty(fd)
                    if fds.get('background') == '#006272':
                        pct = _pct(fds.get('width','0'))
                        fw = bar_w * pct / 100
                        if fw > 0:
                            fb = self.s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    E(left+12), E(y), E(fw), E(16))
                            fb.fill.solid(); fb.fill.fore_color.rgb = TEAL; fb.line.fill.background()
                # pct label
                for sp in inner.cssselect('span'):
                    stxt = sp.text_content().strip()
                    if '%' in stxt:
                        _textbox(self.s, left+bar_w-60, y, 72, 16,
                                 stxt, size=8, align=PP_ALIGN.RIGHT)
                y += 24
            if drew_bar: continue

            # text / milestones — use recursive renderer
            txt = child.text_content().strip()
            if not txt or child.tag != 'div': continue

            # check if child has block sub-elements
            has_blocks = any(c.tag in ('p', 'ul', 'div', 'table') for c in child)
            if has_blocks:
                # render direct inline text first (e.g. <strong>Key Milestones:</strong>)
                tb = _textbox(self.s, left+12, y, w-24, 14)
                _render_rich(tb.text_frame.paragraphs[0], child, 8, skip_blocks=True)
                y += 16
                # then render block children
                y = self._render_box_content(child, left, y, w, indent=12, fs_pt=8)
            else:
                tb = _textbox(self.s, left+12, y, w-24, 14)
                _render_rich(tb.text_frame.paragraphs[0], child, 8)
                y += 16

    # ── workload table in its own container ────────────────────
    def _workload_table(self, div, st):
        top, left, w = _px(st.get('top','0')), _px(st.get('left','0')), _px(st.get('width','900'))
        # table is inside a nested div under the section-box OR directly in a sibling div
        box = div.cssselect('.section-box')
        tables = div.cssselect('table')
        if not tables:
            # try sibling container (border div wrapping table)
            for ch in div:
                tables = ch.cssselect('table') if hasattr(ch, 'cssselect') else []
                if tables: break
        if not tables: return
        self._render_table(tables[0], left, top+20, w, dashed=True)

    # ── standalone table (summary slide) ───────────────────────
    def _standalone_table(self, div, st):
        top, left, w = _px(st.get('top','0')), _px(st.get('left','0')), _px(st.get('width','900'))
        tables = div.cssselect('table')
        if tables:
            self._render_table(tables[0], left, top, w)

    # ── generic HTML table → pptx table ───────────────────────
    def _render_table(self, table_el, left, top, width, dashed=False):
        trs = list(table_el.cssselect('tr'))
        if not trs: return
        n_cols = len(trs[0].cssselect('th, td'))
        n_rows = len(trs)
        if not n_cols or not n_rows: return

        # ── table-level font-size ──
        tbl_sty = _sty(table_el)
        tbl_fs_px = _px(tbl_sty.get('font-size', '11'))

        # ── resolve column widths from first row ──
        first_cells = trs[0].cssselect('th, td')
        explicit_w = []
        for c in first_cells:
            cw = _px(_sty(c).get('width', '0'))
            explicit_w.append(cw)
        total_explicit = sum(explicit_w)
        remaining = width - total_explicit
        n_auto = sum(1 for w in explicit_w if w == 0)
        auto_w = remaining / max(n_auto, 1) if n_auto else 0
        col_widths = [w if w > 0 else auto_w for w in explicit_w]

        # ── detect if this is a workload-class table ──
        is_workload = 'workload' in (table_el.get('class', '') or '')

        row_h = 22
        shape = self.s.shapes.add_table(n_rows, n_cols,
                    E(left), E(top), E(width), E(n_rows * row_h + 4))
        tbl = shape.table
        _nuke_table_theme(shape)

        # apply column widths
        for ci, cw in enumerate(col_widths):
            if ci < n_cols:
                tbl.columns[ci].width = E(cw)

        for ri, tr in enumerate(trs):
            cells = tr.cssselect('th, td')
            tr_sty = _sty(tr)
            tr_bg = _parse_color(tr_sty.get('background', ''))
            tr_color = _parse_color(tr_sty.get('color', ''))
            for ci, td in enumerate(cells):
                if ci >= n_cols: break
                cell = tbl.cell(ri, ci)
                ds = _sty(td)
                cls = td.get('class', '') or ''

                # detect circle indicator
                txt = td.text_content().strip()
                cc = _circle_color(td)
                if cc: txt = '●'

                # cell font-size
                cell_fs = _px(ds.get('font-size', '0'))
                fs_pt = (cell_fs if cell_fs > 0 else tbl_fs_px) * 0.75
                if fs_pt < 6: fs_pt = 8

                # text alignment — inline style, then class-based
                align = ds.get('text-align', '')
                if not align:
                    if 'row-label' in cls:
                        align = 'left'
                    elif is_workload:
                        align = 'center'
                    else:
                        align = 'left'

                # font-weight — inline or class-based
                fw = ds.get('font-weight', '')
                if not fw and is_workload:
                    fw = '400'  # workload tables default normal weight

                tf = cell.text_frame; tf.clear()
                p = tf.paragraphs[0]
                p.space_before = Pt(0); p.space_after = Pt(0)
                p.alignment = {'center': PP_ALIGN.CENTER, 'right': PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)

                r = p.add_run(); r.font.name = FONT
                if cc:
                    r.text = '●'; r.font.size = Pt(10); r.font.color.rgb = cc
                else:
                    r.text = txt; r.font.size = Pt(fs_pt)
                    # cell color: inline > tr color > default
                    cell_color = _parse_color(ds.get('color', ''))
                    r.font.color.rgb = cell_color or tr_color or BLACK33

                # bold: explicit font-weight or th with dark bg
                if fw:
                    r.font.bold = fw not in ('400', 'normal')
                elif td.tag == 'th':
                    r.font.bold = True  # default for th

                # fills
                cell_bg = _parse_color(ds.get('background', ''))
                if td.tag == 'th':
                    bg = cell_bg or tr_bg
                    if bg:
                        _cell_fill(cell, bg)
                        # only force white text on dark backgrounds
                        if bg not in (WHITE, RGBColor(0xF5,0xF5,0xF5), RGBColor(0xF9,0xF9,0xF9)):
                            r.font.color.rgb = WHITE
                elif tr_bg:
                    _cell_fill(cell, tr_bg)
                elif cell_bg:
                    _cell_fill(cell, cell_bg)

                # borders
                _cell_border(cell, GREY_CC, 6350, 'dashed' if dashed else 'solid')
                cell.margin_left = E(4); cell.margin_right = E(4)
                cell.margin_top  = E(2); cell.margin_bottom = E(2)

    # ── legend (summary slide) ─────────────────────────────────
    def _legend(self):
        for el in self.el.cssselect('div[style]'):
            st = _sty(el)
            if st.get('position') != 'absolute' or 'bottom' not in st: continue
            txt = el.text_content().strip()
            if 'Legend' not in txt or el.cssselect('.section-header'): continue

            bottom = _px(st.get('bottom','50'))
            lx = _px(st.get('left','30'))
            ty = SLIDE_H_PX - bottom - 20

            _textbox(self.s, lx, ty, 800, 20,
                     txt.replace('●','').strip(), size=8, color=GREY66)
            for c, xo in [
                (RGBColor(0x22,0xC5,0x5E), 68),  (RGBColor(0xF5,0x9E,0x0B), 138),
                (RGBColor(0xEF,0x44,0x44), 230),  (RGBColor(0x22,0xC5,0x5E), 370),
                (RGBColor(0xF5,0x9E,0x0B), 425),  (RGBColor(0xEF,0x44,0x44), 505),
            ]:
                _oval(self.s, lx+xo, ty+3, 8, c)

    # ── links ──────────────────────────────────────────────────
    def _links(self):
        for el in self.el.cssselect('div[style]'):
            st = _sty(el)
            if st.get('position') != 'absolute': continue
            links = el.cssselect('a.link-text')
            if not links or el.cssselect('.section-header'): continue
            bottom = _px(st.get('bottom','60'))
            lx = _px(st.get('left','30'))
            ty = SLIDE_H_PX - bottom - 15
            for a in links:
                tb = _textbox(self.s, lx, ty, 300, 15,
                              a.text_content().strip(), size=9, color=TEAL)
                tb.text_frame.paragraphs[0].runs[0].font.underline = True

# ═══════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════
def convert(html_path: str, output_path: str):
    with open(html_path, 'r', encoding='utf-8') as f:
        doc = lxml_html.fromstring(f.read())
    slide_els = doc.cssselect('div.slide')
    if not slide_els:
        print("No slides found."); return

    prs = Presentation()
    prs.slide_width  = Emu(int(SLIDE_W_IN * 914400))
    prs.slide_height = Emu(int(SLIDE_H_IN * 914400))
    blank = prs.slide_layouts[6]

    for i, el in enumerate(slide_els):
        sl = prs.slides.add_slide(blank)
        SlideRenderer(sl, el).render()
        print(f'  [{i+1}/{len(slide_els)}] rendered')

    prs.save(output_path)
    print(f'\nSaved {output_path}  ({len(slide_els)} slides)')

if __name__ == '__main__':
    src = sys.argv[1] if len(sys.argv) > 1 else 'input.html'
    dst = sys.argv[2] if len(sys.argv) > 2 else 'output.pptx'
    convert(src, dst)
