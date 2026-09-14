# -*- coding: utf-8 -*-
"""
Bài trình bày HĐTN THCS - BẢN SÁNG (Liquid Glass, iOS premium)
Chủ đề: Phát triển mối quan hệ hòa đồng, hợp tác với thầy cô và bạn bè
Người trình bày: học sinh ("em / chúng em")
- Giao diện SÁNG, hiệu ứng kính lỏng (glass cards trong suốt + bóng mềm + ánh sáng)
- Icon Apple SF Symbols PNG tải trực tiếp từ web (không tự vẽ/render icon, không emoji)
- Font có italic (Segoe UI Italic) cho các câu nhấn
- Animation: chuyển slide Fade + nội dung hiện dần khi bấm phím

Chạy:  python3 build_liquid.py
"""
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

# ---------------------------------------------------------------- hình học
SW, SH = 13.333, 7.5
EMU_IN = 914400
ML = 0.66
CONTENT_W = SW - 2 * ML          # 12.01
CW2 = (CONTENT_W - 0.42) / 2     # nửa ngang có rãnh
CW3 = (CONTENT_W - 2 * 0.34) / 3
CW4 = (CONTENT_W - 3 * 0.3) / 4

# ---------------------------------------------------------------- màu (nền sáng)
INK    = "101631"   # chữ chính (navy đậm)
SUB    = "4E5879"   # chữ phụ
FAINT  = "8A93B4"   # chữ mờ
WHITE  = "FFFFFF"
BG_T   = "FAFCFF"
BG_M   = "EEF3FE"
BG_B   = "E2EAF9"

# accent iOS (bản text đậm hơn cho dễ đọc trên nền sáng)
BLUE   = "0A66E0"
INDIGO = "5A54D6"
PURPLE = "A24AD6"
PINK   = "DE1F8B"
RED    = "E23B30"
ORANGE = "D97600"
YELLOW = "B08900"
TEAL   = "0E9E96"
MINT   = "149A4E"
VIOLET = "7C3AED"

# Apple SF Symbols tải trực tiếp (PNG nguyên gốc) đặt trong assets/apple-sf
FONT = "Segoe UI"

# Tên nội bộ của bài trình bày -> file PNG Apple SF Symbols đã tải về.
# Không có bước tự vẽ hay tạo lại icon trong mã nguồn.
APPLE_SF = {
    "message-circle": "bubble-left-fill", "messages-square": "bubble-pair-fill",
    "book-open": "book-fill", "ear": "ear", "users": "person-2-fill",
    "user-plus": "person-badge-plus", "graduation-cap": "person-crop-circle-fill",
    "clipboard-list": "doc-on-clipboard-fill", "clipboard-check": "checkmark-circle-fill",
    "flag": "flag-fill", "target": "scope", "sparkles": "sparkles",
    "party-popper": "star-fill", "mic": "waveform", "flask-conical": "drop-triangle-fill",
    "zap": "bolt-fill", "handshake": "person-2-fill", "heart": "heart-fill",
    "heart-handshake": "heart-circle-fill", "smile": "smiley-fill",
    "lightbulb": "lightbulb-fill", "list-checks": "checkmark-circle-fill",
    "quote": "quote-bubble-fill", "presentation": "rectangle-3-offgrid-fill",
}

A_ = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_ = "http://schemas.openxmlformats.org/presentationml/2006/main"

# ---------------------------------------------------------------- low-level
def _rgb(h):
    return RGBColor.from_string(h.lstrip("#"))

def _alpha(hexstr, pct):
    return ('<a:srgbClr val="%s"><a:alpha val="%d"/></a:srgbClr>'
            % (hexstr, int(round(pct * 1000))))

def _fill_xml(sp, xml):
    spPr = sp._element.spPr
    for tag in ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"):
        for el in list(spPr):
            if el.tag == qn("a:" + tag):
                spPr.remove(el)
    el = etree.fromstring(xml)
    ln = spPr.find(qn("a:ln"))
    if ln is not None:
        ln.addprevious(el)
    else:
        spPr.append(el)

def fill_solid(sp, color, alpha=None):
    if alpha is None:
        sp.fill.solid()
        sp.fill.fore_color.rgb = _rgb(color)
    else:
        _fill_xml(sp, '<a:solidFill xmlns:a="%s">%s</a:solidFill>' % (A_, _alpha(color, alpha)))

def fill_grad(sp, stops, angle=90, radial=False):
    """stops: list (pos0..100000, hex, alpha%)"""
    gs = "".join('<a:gs pos="%d">%s</a:gs>'
                 % (p, _alpha(c, a if a is not None else 100)) for p, c, a in stops)
    if radial:
        path = '<a:path path="circle"><a:fillToRect l="50000" t="40000" r="50000" b="60000"/></a:path>'
    else:
        path = '<a:lin ang="%d" scaled="1"/>' % int(angle * 60000)
    _fill_xml(sp, '<a:gradFill xmlns:a="%s" rotWithShape="1"><a:gsLst>%s</a:gsLst>%s</a:gradFill>'
                  % (A_, gs, path))

def stroke(sp, color=WHITE, weight=1.0, alpha=None):
    spPr = sp._element.spPr
    for el in list(spPr):
        if el.tag == qn("a:ln"):
            spPr.remove(el)
    if alpha is None:
        sp.line.color.rgb = _rgb(color)
        sp.line.width = Pt(weight)
    else:
        xml = ('<a:ln xmlns:a="%s" w="%d" cap="flat" cmpd="sng" algn="ctr">'
               '<a:solidFill>%s</a:solidFill></a:ln>'
               % (A_, int(weight * 12700), _alpha(color, alpha)))
        spPr.append(etree.fromstring(xml))

def no_shadow(sp):
    try:
        sp.shadow.inherit = False
    except Exception:
        pass

def shape(slide, kind, x, y, w, h):
    sp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    no_shadow(sp)
    return sp

def _shadow_xml(sp, blur_pt=8, dist_pt=3, alpha_pct=16, color="33406E"):
    """Đổ bóng mềm (outer shadow) cho hiệu ứng nổi trên nền sáng."""
    spPr = sp._element.spPr
    for el in list(spPr):
        if el.tag == qn("a:effectLst"):
            spPr.remove(el)
    xml = ('<a:effectLst xmlns:a="%s"><a:outerShdw blurRad="%d" dist="%d" dir="5400000" rotWithShape="0">%s</a:outerShdw></a:effectLst>'
           % (A_, int(blur_pt * 12700), int(dist_pt * 12700),
              _alpha(color, alpha_pct)))
    spPr.append(etree.fromstring(xml))

def rounded(slide, x, y, w, h, radius=0.12, fill=None, alpha=None,
            border=None, border_w=1.0, border_alpha=None, grad=None,
            shadow=None, gloss=False):
    sp = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h)
    try:
        sp.adjustments[0] = min(0.5, max(0.02, radius))
    except Exception:
        pass
    if grad:
        fill_grad(sp, grad, 90)
    elif fill is not None:
        fill_solid(sp, fill, alpha)
    else:
        fill_solid(sp, "000000", 0)
    if border is not None:
        stroke(sp, border, border_w, border_alpha)
    if shadow:
        _shadow_xml(sp, *shadow)
    if gloss:
        _gloss(slide, x, y, w, h, radius)
    return sp

def _gloss(slide, x, y, w, h, radius):
    """Vệt sáng nhỏ trên đỉnh card - gợi cảm giác bề mặt kính."""
    gw = w * 0.42
    if gw > 0.6:
        g = shape(slide, MSO_SHAPE.OVAL, x + w / 2 - gw / 2, y + h * 0.06, gw, h * 0.1)
        fill_solid(g, WHITE, 62)
        g.rotation = -12

def oval(slide, cx, cy, d, fill=None, alpha=None, border=None, border_w=1.0,
         border_alpha=None, grad=None, shadow=None):
    sp = shape(slide, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d)
    if grad:
        fill_grad(sp, grad, 135, radial=True)
    elif fill is not None:
        fill_solid(sp, fill, alpha)
    else:
        fill_solid(sp, "000000", 0)
    if border is not None:
        stroke(sp, border, border_w, border_alpha)
    if shadow:
        _shadow_xml(sp, *shadow)
    return sp

def pic(slide, name, x, y, s=None, w=None, h=None):
    """Chèn nguyên file PNG Apple SF Symbols đã tải từ web, không bóp méo tỉ lệ."""
    import os
    from PIL import Image
    root = os.path.dirname(os.path.abspath(__file__))
    try:
        source_name = APPLE_SF[name]
    except KeyError as exc:
        raise KeyError("Chưa có Apple SF Symbol cho: %s" % name) from exc
    path = os.path.join(root, "assets", "apple-sf", source_name + ".png")
    if not os.path.exists(path):
        raise FileNotFoundError("Thiếu icon Apple SF Symbols: " + path)
    if s is not None:
        w = h = s
    # File nguồn có tỉ lệ khác nhau; căn giữa trong khung để hình không dính/chạm chữ.
    with Image.open(path) as im:
        iw, ih = im.size
    scale = min(w / iw, h / ih)
    dw, dh = iw * scale, ih * scale
    return slide.shapes.add_picture(path, Inches(x + (w - dw) / 2), Inches(y + (h - dh) / 2),
                                    Inches(dw), Inches(dh))

# ---------------------------------------------------------------- text
def _run(p, text, size, color, bold=False, italic=False, spc=None, font=FONT):
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    r.font.name = font
    r.font.color.rgb = _rgb(color)
    if spc is not None:
        rPr = r._r.get_or_add_rPr()
        rPr.set("spc", str(int(spc * 100)))
    return r

def tx(slide, x, y, w, h, paras, anchor=MSO_ANCHOR.TOP, wrap=True):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    first = True
    for pa in paras:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = pa.get("align", PP_ALIGN.LEFT)
        if pa.get("line"):
            p.line_spacing = pa["line"]
        if pa.get("before"):
            p.space_before = Pt(pa["before"])
        if pa.get("after"):
            p.space_after = Pt(pa["after"])
        for rr in pa["runs"]:
            _run(p, rr[0], rr[1], rr[2],
                 rr[3] if len(rr) > 3 else False,
                 rr[4] if len(rr) > 4 else False,
                 rr[5] if len(rr) > 5 else None)
    return tb

def one(slide, x, y, w, h, text, size, color, bold=True, italic=False,
        align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spc=None, font=FONT,
        line=None):
    pa = {"runs": [(text, size, color, bold, italic, spc)], "align": align}
    if line:
        pa["line"] = line
    return tx(slide, x, y, w, h, [pa], anchor=anchor)

# ---------------------------------------------------------------- decor & QA font
from PIL import ImageFont
_DEJ = "/usr/local/lib/python3.11/dist-packages/matplotlib/mpl-data/fonts/ttf/DejaVuSans.ttf"
_DEJ_B = "/usr/local/lib/python3.11/dist-packages/matplotlib/mpl-data/fonts/ttf/DejaVuSans-Bold.ttf"

def tw(text, size, bold=False):
    """Bề rộng text (inch) đo bằng DejaVu - an toàn hơn Segoe."""
    f = ImageFont.truetype(_DEJ_B if bold else _DEJ, int(size * 4))
    return f.getlength(text) / 4.0 * 0.98 / 72.0

def pill(slide, x, y, text, size=10.5, h=0.34, padx=0.16, color=BLUE,
         fill_alpha=22, border_alpha=40, bold=True, spc=1.2, italic=False,
         border=WHITE, fill=WHITE):
    w = tw(text, size, bold) + padx * 2
    p = rounded(slide, x, y, w, h, radius=0.5, fill=fill, alpha=fill_alpha,
                border=border, border_w=0.9, border_alpha=border_alpha)
    tf = p.text_frame
    tf.word_wrap = False
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    pa = tf.paragraphs[0]; pa.alignment = PP_ALIGN.CENTER
    _run(pa, text, size, color, bold, italic, spc=spc)
    return p

def chip(slide, x, y, text, color, size=10.5, h=0.32, **kw):
    return pill(slide, x, y, text, size=size, h=h, color=color, **kw)

def glass(slide, x, y, w, h, radius=0.16, alpha=46, shadow=True, gloss=True,
          border_alpha=78):
    return rounded(slide, x, y, w, h, radius=radius, fill=WHITE, alpha=alpha,
                   border=WHITE, border_w=1.1, border_alpha=border_alpha,
                   shadow=(9, 3.5, 14) if shadow else None, gloss=gloss)

def num_badge(slide, cx, cy, d, n, color=BLUE):
    c = oval(slide, cx, cy, d, fill=WHITE, alpha=88, border=WHITE,
             border_w=1, border_alpha=90, shadow=(6, 2, 12))
    tf = c.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _run(p, str(n), int(d * 44), color, True)
    return c

def bullet_lines(slide, x, y, w, h, items, size=12.5, gap=7, dot_color=BLUE,
                 text_color=SUB, line=1.24):
    paras = []
    for it in items:
        paras.append({"runs": [("•  ", size, dot_color, True),
                               (it, size, text_color, False)],
                      "line": line, "after": gap})
    return tx(slide, x, y, w, h, paras)

def notes(slide, text):
    slide.notes_slide.notes_text_frame.text = text

# ---------------------------------------------------------------- nền sáng
def bg(slide, seed=0):
    st = len(slide.shapes)
    r = shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SW, SH)
    fill_grad(r, [(0, BG_T, None), (46, BG_M, None), (100, BG_B, None)], 90)
    blobs = [
        (0.2, 0.4, 5.6, "BFD8FF", 78, 0),
        (12.4, 1.0, 5.0, "D9E8FF", 80, 0),
        (12.9, 7.6, 6.4, "E4D6FF", 66, 0),
        (0.6, 7.7, 5.2, "D5F5E2", 70, 0),
        (6.6, 0.2, 3.4, "FFE7D2", 55, 0),
        (8.0, 7.9, 3.6, "FFE2EC", 50, 0),
    ]
    for (cx, cy, d, col, a0, a1) in blobs:
        oval(slide, cx, cy, d, grad=[(0, col, a0), (100, col, a1)])
    # vài chấm sáng nhỏ
    for (x, y, d, col, a) in [(1.6, 1.2, 0.09, WHITE, 90), (11.3, 5.6, 0.12, WHITE, 80),
                              (2.2, 5.9, 0.08, WHITE, 70), (10.8, 6.9, 0.1, WHITE, 85),
                              (0.9, 3.2, 0.07, WHITE, 60)]:
        oval(slide, x, y, d, fill=col, alpha=a)
    for sp in list(slide.shapes)[st:]:
        sp.name = "~bg" + sp.name

# ---------------------------------------------------------------- header
def header(slide, num, total, kicker=None, kcolor=BLUE, title=None,
           title_size=27, italic_note=None, accent_tile=None):
    st = len(slide.shapes)
    if kicker:
        # viên accent nhỏ kiểu app icon
        if accent_tile:
            pic(slide, accent_tile, ML + 0.02, 0.44, s=0.24)
            kx = ML + 0.36
        else:
            r = rounded(slide, ML + 0.01, 0.47, 0.15, 0.15, radius=0.4,
                        fill=kcolor, alpha=100)
            del r
            kx = ML + 0.27
        one(slide, kx, 0.47, 9.4, 0.26, kicker, 10.5, kcolor, bold=True, spc=2.2)
    if title:
        one(slide, ML, 0.88, CONTENT_W - 2.6, 0.62, title, title_size, INK,
            bold=True, spc=0.2)
        rounded(slide, ML + 0.005, 0.88 + 0.66, 0.72, 0.05, radius=0.5,
                fill=kcolor, alpha=100)
    if italic_note:
        one(slide, SW - ML - 4.1, 0.52, 4.1, 0.3, italic_note, 11.5, FAINT,
            bold=False, italic=True, align=PP_ALIGN.RIGHT)
    # footer
    for i, (col) in enumerate(["0A84FF", "00C7BE", "FF9F0A", "FF2D92"]):
        r = rounded(slide, ML + i * 0.24, SH - 0.46, 0.16, 0.16, radius=0.34,
                    fill=col, alpha=100, shadow=(4, 1, 10))
    one(slide, ML + 1.06, SH - 0.445, 6.2, 0.24,
        "HĐTN · HÒA ĐỒNG & HỢP TÁC — BÀI TRÌNH BÀY CỦA HỌC SINH", 7.5, FAINT,
        bold=True, spc=1.6, anchor=MSO_ANCHOR.MIDDLE)
    c = oval(slide, SW - ML - 0.5, SH - 0.37, 0.34, fill=WHITE, alpha=75,
             border=WHITE, border_w=0.8, border_alpha=80)
    tf = c.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    _run(p, "%02d" % num, 8.5, FAINT, True)
    for sp in list(slide.shapes)[st:]:
        sp.name = "~hdr" + sp.name

# ---------------------------------------------------------------- animation XML
def add_timing(slide, groups):
    """groups: list các nhóm; mỗi nhóm là list(spid) hiện cùng lúc.
    Nhấn phím/chuột để hiện từng nhóm (Fade)."""
    ids = [1000]
    def nid():
        ids[0] += 1
        return ids[0]

    def eff_xml(spid, node_type):
        i1 = nid(); i2 = nid(); i3 = nid()
        return (
            '<p:par xmlns:p="%s"><p:cTn id="%d" presetID="10" presetClass="entr" '
            'presetSubtype="0" fill="hold" grpId="0" nodeType="%s">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            '<p:set><p:cBhvr><p:cTn id="%d" dur="1" fill="hold">'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
            '<p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
            '</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
            '<p:animEffect transition="in" filter="fade"><p:cBhvr>'
            '<p:cTn id="%d" dur="700"/><p:tgtEl><p:spTgt spid="%d"/></p:tgtEl>'
            '</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>'
            % (P_, i1, node_type, i2, spid, i3, spid))

    parts = []
    for gi, grp in enumerate(groups):
        if not grp:
            continue
        gid = nid()
        inner = ""
        for si, spid in enumerate(grp):
            node = "clickEffect" if si == 0 else "withEffect"
            inner += eff_xml(spid, node)
        parts.append(
            '<p:par><p:cTn id="%d" fill="hold">'
            '<p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>'
            '<p:childTnLst>%s</p:childTnLst></p:cTn></p:par>' % (gid, inner))
    i1 = nid(); i2 = nid(); i3 = nid()
    timing = (
        '<p:timing xmlns:p="%s"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" '
        'restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" '
        'nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        '<p:childTnLst>%s</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/>'
        '</p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
        '</p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>'
        % (P_, "".join(parts)))
    slide._element.append(etree.fromstring(timing))

def add_transition(slide, spd="med"):
    el = etree.fromstring(
        '<p:transition xmlns:p="%s" spd="%s"><p:fade/></p:transition>'
        % (P_, spd))
    sld = slide._element
    timing = sld.find(qn("p:timing"))
    if timing is not None:
        timing.addprevious(el)
    else:
        sld.append(el)

# ---------------------------------------------------------------- slide đếm
TOTAL = 18

def _ids(*shapes):
    return [s.shape_id for s in shapes if s is not None]

# ================================================================ SLIDES
def s1_cover(prs):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 1)
    anim = [[], [], []]
    # Kicker: một SF Symbol và phần chữ tách hẳn nhau, không chồng icon lên chữ.
    cover_icon = pic(s, "heart-handshake", ML, 0.78, s=0.4)
    k = chip(s, ML + 0.56, 0.78, "HOẠT ĐỘNG TRẢI NGHIỆM · HƯỚNG NGHIỆP — THCS",
             BLUE, size=10.5, h=0.4, padx=0.22, spc=1.6)
    anim[0] += _ids(k, cover_icon)
    # tiêu đề
    one(s, ML, 1.5, 7.6, 0.6, "Phát triển mối quan hệ", 37, INK, bold=True)
    one(s, ML, 2.14, 7.9, 0.6, "hòa đồng, hợp tác", 37, INDIGO, bold=True)
    one(s, ML, 2.78, 7.9, 0.6, "với thầy cô và bạn bè", 37, INK, bold=True)
    one(s, ML, 3.58, 7.3, 0.4,
        "Kết nối thật nhiều, để lớp mình gần nhau hơn mỗi ngày.",
        14.5, SUB, bold=False, italic=True)
    t1 = pill(s, ML, 4.16, "TIẾT 1 · KẾT NỐI VÀ CHIA SẺ", size=11, h=0.42,
              padx=0.2, color=TEAL, fill_alpha=30, spc=1.3)
    w1 = t1.width / EMU_IN
    t2 = pill(s, ML + w1 + 0.2, 4.16, "TIẾT 2 · CÙNG THỰC HÀNH", size=11,
              h=0.42, padx=0.2, color=PINK, fill_alpha=30, spc=1.3)
    anim[0] += [t1.shape_id, t2.shape_id]
    # người trình bày
    meta = glass(s, ML, 5.0, 6.7, 1.32, radius=0.18, alpha=40)
    tx(s, ML + 0.34, 5.22, 6.2, 0.95, [
        {"runs": [("Người trình bày:  ", 12.5, SUB, False),
                  ("[Tên em / nhóm em]", 12.5, INK, True)], "after": 5},
        {"runs": [("Lớp – Trường:  ", 12.5, SUB, False),
                  ("[Lớp … – Tên trường]", 12.5, INK, True)], "after": 5},
        {"runs": [("Ngày trình bày:  ", 12.5, SUB, False),
                  ("[… / … / 2026]", 12.5, INK, True)]},
    ])
    anim[1] += _ids(meta)
    # hero phải: vòng kính + tiles nổi
    for cx, cy, d, a in [(10.35, 3.5, 3.3, 26), (10.55, 3.3, 2.6, 22),
                         (10.1, 3.7, 1.95, 18)]:
        oval(s, cx, cy, d, fill=WHITE, alpha=a, border=WHITE, border_w=1.4,
             border_alpha=80)
    h1 = pic(s, "heart-handshake", 9.35, 2.55, s=0.72)
    h2 = pic(s, "message-circle", 11.15, 3.05, s=0.62)
    h3 = pic(s, "users", 9.7, 4.25, s=0.58)
    h4 = pic(s, "sparkles", 11.35, 4.75, s=0.5)
    h1.rotation = -8; h3.rotation = 8
    anim[1] += [h1.shape_id, h2.shape_id, h3.shape_id, h4.shape_id]
    cap1 = pill(s, 8.35, 5.62, "Hòa đồng", size=10, h=0.32, padx=0.14,
                color=TEAL, fill_alpha=26, bold=True)
    cap2 = pill(s, 11.0, 5.86, "Hợp tác", size=10, h=0.32, padx=0.14,
                color=PINK, fill_alpha=26, bold=True)
    anim[2] += [cap1.shape_id, cap2.shape_id]
    notes(s, "Chào thầy cô và các bạn! Em là [tên], học sinh lớp [..]. Hôm nay "
             "chúng em xin trình bày chủ đề: phát triển mối quan hệ hòa đồng, hợp "
             "tác với thầy cô và bạn bè. Chúng em chia làm 2 phần: Tiết 1 kết nối "
             "và chia sẻ, Tiết 2 cùng nhau thực hành xử lí tình huống. Rất mong cả "
             "lớp tham gia nhiệt tình!")
    return s, anim

def s2_roadmap(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 2)
    header(s, num, TOTAL, kicker="NỘI DUNG BUỔI TRÌNH BÀY", kcolor=BLUE,
           title="Chúng em sẽ chia sẻ những gì?", title_size=30,
           accent_tile="presentation")
    y0, hh = 1.9, 4.0
    groups = [[], []]
    # tiết 1
    c1 = glass(s, ML, y0, CW2, hh)
    i1 = pic(s, "sparkles", ML + 0.35, y0 + 0.35, s=0.62)
    one(s, ML + 1.18, y0 + 0.45, CW2 - 1.4, 0.45, "Tiết 1 · Khám phá & chia sẻ",
        17, INK, bold=True)
    chip(s, ML + 0.35, y0 + 1.3, "CHÚNG EM CÙNG TÌM HIỂU", TEAL, size=9.5,
         h=0.3, spc=1.4)
    b1 = bullet_lines(s, ML + 0.4, y0 + 1.9, CW2 - 0.8, hh - 2.05, [
        "Hiểu hòa đồng, hợp tác là gì và vì sao cần có.",
        "4 “chìa khóa” kết nối + cách làm việc nhóm.",
        "Hợp tác với thầy cô: lắng nghe, mạnh dạn hỏi.",
        "Kể câu chuyện kỉ niệm về tinh thần hợp tác.",
    ], size=12, gap=7)
    groups[0] += [c1.shape_id, i1.shape_id, b1.shape_id]
    # tiết 2
    x2 = ML + CW2 + 0.42
    c2 = glass(s, x2, y0, CW2, hh)
    i2 = pic(s, "party-popper", x2 + 0.35, y0 + 0.35, s=0.62)
    one(s, x2 + 1.18, y0 + 0.45, CW2 - 1.4, 0.45, "Tiết 2 · Cùng thực hành", 17,
        INK, bold=True)
    chip(s, x2 + 0.35, y0 + 1.3, "CẢ LỚP CÙNG THAM GIA", PINK, size=9.5,
         h=0.3, spc=1.4)
    b2 = bullet_lines(s, x2 + 0.4, y0 + 1.9, CW2 - 0.8, hh - 2.05, [
        "3 tình huống SGK — các nhóm cùng phân vai xử lí.",
        "Diễn hoặc thuyết trình cách giải quyết.",
        "Cả lớp nhận xét và góp thêm ý hay.",
        "Chọn một hành động hợp tác cho tuần tới.",
    ], size=12, gap=7)
    groups[1] += [c2.shape_id, i2.shape_id, b2.shape_id]
    ph = pill(s, 0, 6.35, "Mong muốn của chúng em: cả lớp tham gia thật sôi nổi nhé!",
              size=12, h=0.46, padx=0.26, color=INK, fill_alpha=30, spc=0.2)
    ph.left = Emu(int(SW * EMU_IN / 2 - ph.width / 2))
    groups[1].append(ph.shape_id)
    notes(s, "Buổi trình bày gồm 2 tiết như trên màn hình. Các bạn nhìn theo dõi "
             "và giúp chúng em bằng cách tham gia thật nhiệt tình nhé!")
    return s, groups

def s3_goals(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 3)
    header(s, num, TOTAL, kicker="ĐIỀU CHÚNG EM MONG MUỐN", kcolor=PURPLE,
           title="Sau buổi hôm nay, lớp mình sẽ…", title_size=30,
           accent_tile="target")
    cards = [
        ("handshake", "Hiểu đúng hơn", "Hòa đồng, hợp tác là gì và vì sao nó quan trọng."),
        ("message-circle", "Gần nhau hơn", "Biết thêm nhiều cách nhỏ để kết nối với bạn bè."),
        ("users", "Phối hợp ăn ý", "Làm việc nhóm vui và hiệu quả, ai cũng được tham gia."),
        ("target", "Tự tin ứng xử", "Xử lí khéo léo khi gặp tình huống khó ở lớp."),
    ]
    n = 4; gap = 0.3
    w = CW4
    y0, hh = 2.1, 3.3
    groups = [[]]
    for i, (tile, t, d) in enumerate(cards):
        x = ML + i * (w + gap)
        c = glass(s, x, y0, w, hh, radius=0.18)
        ic = pic(s, tile, x + w / 2 - 0.33, y0 + 0.32, s=0.66)
        one(s, x + 0.12, y0 + 1.16, w - 0.24, 0.42, t, 17.5, INK, bold=True,
            align=PP_ALIGN.CENTER)
        tx(s, x + 0.22, y0 + 1.72, w - 0.44, hh - 2.0, [
            {"runs": [(d, 12, SUB, False)], "align": PP_ALIGN.CENTER,
             "line": 1.26}])
        groups[0] += [c.shape_id, ic.shape_id]
    ph = pill(s, 0, 5.95, "Các bạn giúp chúng em bằng cách lắng nghe và góp ý chân thành!",
              size=11.5, h=0.42, padx=0.24, color=SUB, bold=False,
              fill_alpha=20, spc=0.2)
    ph.left = Emu(int(SW * EMU_IN / 2 - ph.width / 2))
    notes(s, "Chúng em mong sau buổi này cả lớp hiểu hơn và thân nhau hơn. Nên rất "
             "cần các bạn cùng lắng nghe và tham gia.")
    return s, groups

def s4_warmup(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 4)
    header(s, num, TOTAL, kicker="KHỞI ĐỘNG", kcolor=ORANGE,
           title="Các bạn thử nhớ lại xem…", title_size=29, accent_tile="smile")
    # Chừa một hàng riêng cho câu nhấn cuối card, tránh dính vào icon/câu hỏi 3.
    y0, hh = 1.88, 4.65
    c = glass(s, ML, y0, CONTENT_W, hh, radius=0.16)
    tx(s, ML + 0.55, y0 + 0.25, CONTENT_W - 1.1, 0.4, [
        {"runs": [("Quay sang bạn kế bên, chia sẻ nhanh trong 1 phút nhé:", 14.5,
                   INK, True)]}])
    qs = [
        ("smile", "Bạn nào trong lớp mình luôn vui vẻ, thân thiện với tất cả mọi "
                  "người? Vì sao bạn ấy được yêu quý?"),
        ("handshake", "Các bạn đã cùng nhau hoàn thành một việc khó bao giờ chưa? "
                      "Lúc đó cảm thấy thế nào?"),
        ("graduation-cap", "Các bạn từng mạnh dạn nhờ thầy cô giúp đỡ chưa? Điều gì "
                           "khiến mình dám hỏi?"),
    ]
    yq = y0 + 0.92
    for tile, q in qs:
        ic = pic(s, tile, ML + 0.5, yq, s=0.5)
        tx(s, ML + 1.25, yq + 0.01, CONTENT_W - 2.1, 0.76, [
            {"runs": [(q, 13.5, SUB, False)], "line": 1.20}])
        yq += 0.94
    groups = [[c.shape_id]]
    ph = pill(s, ML + 0.5, y0 + hh - 0.58,
              "Mỗi câu trả lời đều hé lộ một kĩ năng mà hôm nay chúng ta cùng rèn!",
              size=11.5, h=0.4, padx=0.2, color=ORANGE, fill_alpha=20, spc=0.1)
    notes(s, "Chúng ta chơi nhỏ: quay sang bạn bên cạnh chia sẻ 1 phút theo 3 câu "
             "hỏi. Bạn nào xung phong nói với cả lớp? Những điều các bạn vừa kể - "
             "như trò chuyện, giúp đỡ, dám hỏi - chính là kĩ năng hòa đồng và hợp "
             "tác mà hôm nay chúng em muốn cả lớp cùng rèn.")
    return s, groups

def s5_concept(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 5)
    header(s, num, TOTAL, kicker="TIẾT 1 · CHÚNG EM HIỂU", kcolor=TEAL,
           title="Hòa đồng và hợp tác là gì?", title_size=29, accent_tile="quote")
    y0, hh = 1.95, 2.55
    groups = [[], []]
    c1 = glass(s, ML, y0, CW2, hh)
    ic = pic(s, "heart-handshake", ML + 0.38, y0 + 0.35, s=0.6)
    one(s, ML + 1.2, y0 + 0.42, CW2 - 1.4, 0.45, "Hòa đồng là…", 18, INK, bold=True)
    tx(s, ML + 0.4, y0 + 1.08, CW2 - 0.8, hh - 1.3, [
        {"runs": [("Cởi mở, thân thiện, vui vẻ với mọi người; tôn trọng sự khác "
                   "biệt; sẵn sàng trò chuyện, chia sẻ và giúp đỡ nhau.", 13,
                   SUB, False)], "line": 1.3}])
    groups[0] += [c1.shape_id, ic.shape_id]
    x2 = ML + CW2 + 0.42
    c2 = glass(s, x2, y0, CW2, hh)
    ic2 = pic(s, "users", x2 + 0.38, y0 + 0.35, s=0.6)
    one(s, x2 + 1.2, y0 + 0.42, CW2 - 1.4, 0.45, "Hợp tác là…", 18, INK, bold=True)
    tx(s, x2 + 0.4, y0 + 1.08, CW2 - 0.8, hh - 1.3, [
        {"runs": [("Cùng bàn bạc, chung sức làm một việc chung và cùng chịu trách "
                   "nhiệm về kết quả — không phải làm hộ cho nhau.", 13, SUB,
                   False)], "line": 1.3}])
    groups[1] += [c2.shape_id, ic2.shape_id]
    one(s, ML, 4.85, 9, 0.42, "Vì sao lớp mình nên hòa đồng và hợp tác?", 16.5,
        INK, bold=True)
    ben = [
        ("party-popper", "Lớp vui hơn", "Ai cũng thấy mình được kết nối."),
        ("book-open", "Học tốt hơn", "Kèm cặp, giúp nhau cùng tiến bộ."),
        ("zap", "Mạnh hơn", "Khó khăn nào cũng có tập thể chung vai."),
    ]
    w3 = CW3; gap3 = 0.34
    yb = 5.35
    for i, (tile, t, d) in enumerate(ben):
        x = ML + i * (w3 + gap3)
        c = glass(s, x, yb, w3, 1.35, radius=0.16)
        ic = pic(s, tile, x + 0.28, yb + 0.22, s=0.44)
        one(s, x + 0.88, yb + 0.22, w3 - 1.05, 0.32, t, 14, INK, bold=True)
        tx(s, x + 0.88, yb + 0.56, w3 - 1.1, 0.7, [
            {"runs": [(d, 11, SUB, False)], "line": 1.15}])
        groups[1].append(c.shape_id)
    notes(s, "Theo cách hiểu của chúng em: hòa đồng là cởi mở, thân thiện; hợp tác "
             "là cùng chung sức và cùng chịu trách nhiệm. Lớp nào hòa đồng, hợp "
             "tác thì ai cũng vui và học tốt hơn. Các bạn có đồng ý không?")
    return s, groups

def s6_keys(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 6)
    header(s, num, TOTAL, kicker="TIẾT 1 · KẾT NỐI VỚI BẠN BÈ", kcolor=INDIGO,
           title="Bốn “chìa khóa” kết nối với bạn bè", title_size=29,
           accent_tile="message-circle")
    cards = [
        ("message-circle", "Trò chuyện, tâm sự",
         "Chủ động hỏi thăm bạn; chia sẻ niềm vui, nỗi buồn và khó khăn của mình."),
        ("book-open", "Cùng học",
         "Học nhóm, giảng bài cho nhau, cùng chinh phục những bài khó."),
        ("party-popper", "Cùng tham gia hoạt động",
         "Thể thao, văn nghệ, trực nhật, câu lạc bộ… vui gấp đôi khi có bạn cùng chơi."),
        ("ear", "Lắng nghe – tôn trọng",
         "Lắng nghe để thấu hiểu; tôn trọng ý kiến và sự khác biệt của nhau."),
    ]
    gapx, gapy = 0.32, 0.34
    w = (CONTENT_W - gapx) / 2
    hh = (4.3 - gapy) / 2
    y0 = 1.95
    groups = []
    for i, (tile, t, d) in enumerate(cards):
        r_, c = divmod(i, 2)
        x = ML + c * (w + gapx)
        y = y0 + r_ * (hh + gapy)
        cd = glass(s, x, y, w, hh, radius=0.15)
        ic = pic(s, tile, x + 0.3, y + 0.28, s=0.56)
        one(s, x + 1.05, y + 0.34, w - 1.3, 0.4, t, 15.5, INK, bold=True)
        tx(s, x + 0.32, y + 1.08, w - 0.64, hh - 1.35, [
            {"runs": [(d, 12, SUB, False)], "line": 1.24}])
        groups.append([cd.shape_id, ic.shape_id])
    ph = pill(s, ML, 6.55,
              "Bí quyết nhỏ: một nụ cười và lời chào buổi sáng là chìa khóa vạn năng!",
              size=11.5, h=0.4, padx=0.2, color=INDIGO, fill_alpha=18, spc=0.1)
    notes(s, "Chúng em thấy chỉ cần 4 việc rất gần gũi là bạn bè thân nhau hơn: trò "
             "chuyện, cùng học, cùng chơi, lắng nghe nhau. Các bạn còn cách nào nữa "
             "không? Mọi người góp thêm nhé - ví dụ nhắn tin hỏi thăm bạn ốm, khen "
             "bạn làm bài tốt, rủ bạn ăn sáng cùng…")
    return s, groups

def s7_teamwork(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 7)
    header(s, num, TOTAL, kicker="TIẾT 1 · KINH NGHIỆM LÀM VIỆC NHÓM",
           kcolor=TEAL, title="Bốn bước nhỏ để nhóm nào cũng ăn ý", title_size=29,
           accent_tile="users")
    steps = [
        ("clipboard-list", "Lên kế hoạch", "Xác định mục tiêu, việc cần làm và thời hạn."),
        ("users", "Phân công việc", "Chia việc hợp khả năng từng bạn, ai cũng có việc."),
        ("messages-square", "Lắng nghe nhau", "Tiếp thu ý kiến, bàn bạc rồi thống nhất."),
        ("flag", "Giúp nhau về đích", "Hỗ trợ lúc khó; cùng kiểm tra kết quả."),
    ]
    n = 4; gap = 0.2
    w = (CONTENT_W - gap * 3) / 4
    y0, hh = 2.0, 4.0
    groups = []
    for i, (tile, t, d) in enumerate(steps):
        x = ML + i * (w + gap)
        c = glass(s, x, y0, w, hh, radius=0.16)
        nb = num_badge(s, x + w / 2, y0 + 0.5, 0.5, i + 1,
                       [BLUE, TEAL, PURPLE, ORANGE][i])
        ic = pic(s, tile, x + w / 2 - 0.3, y0 + 1.05, s=0.6)
        one(s, x + 0.1, y0 + 1.78, w - 0.2, 0.55, t, 15, INK, bold=True,
            align=PP_ALIGN.CENTER)
        tx(s, x + 0.2, y0 + 2.4, w - 0.4, hh - 2.7, [
            {"runs": [(d, 11, SUB, False)], "align": PP_ALIGN.CENTER,
             "line": 1.22}])
        if i < n - 1:
            ar = shape(s, MSO_SHAPE.CHEVRON, x + w - 0.06, y0 + hh / 2 - 0.12,
                       0.16, 0.24)
            fill_solid(ar, BLUE, 42)
        groups.append([c.shape_id, ic.shape_id])
    ph = pill(s, ML, 6.3,
              "Nhóm mạnh nhất không phải nhóm toàn ngôi sao, mà là nhóm biết phối hợp!",
              size=12, h=0.44, padx=0.24, color=INK, fill_alpha=26, spc=0.2)
    notes(s, "Kinh nghiệm nhóm của chúng em gói trong 4 bước: lên kế hoạch, phân "
             "công, lắng nghe, giúp nhau về đích. Thử nghĩ: nếu thiếu bước phân công "
             "thì sao? Sẽ có bạn làm, bạn chơi! Các bạn thường vướng bước nào nhất?")
    return s, groups

def s8_teachers(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 8)
    header(s, num, TOTAL, kicker="TIẾT 1 · HỢP TÁC VỚI THẦY CÔ", kcolor=PURPLE,
           title="Thầy cô — người đồng hành của chúng em", title_size=26,
           accent_tile="graduation-cap")
    cards = [
        ("ear", "Lắng nghe và làm theo hướng dẫn",
         "Chú ý nghe thầy cô giao việc; chưa rõ thì hỏi lại; làm bài đúng và đầy đủ."),
        ("message-circle", "Mạnh dạn xin ý kiến",
         "Câu hỏi “ngại” nhất thường là câu nhiều bạn cùng thắc mắc. Cứ hỏi, thầy cô luôn sẵn lòng góp ý!"),
        ("clipboard-check", "Lễ phép, có trách nhiệm",
         "Chào hỏi lễ phép, giữ lời hứa, hoàn thành tốt việc thầy cô giao."),
    ]
    w = CW3; gap = 0.34
    y0, hh = 2.0, 3.55
    groups = []
    for i, (tile, t, d) in enumerate(cards):
        x = ML + i * (w + gap)
        c = glass(s, x, y0, w, hh, radius=0.17)
        ic = pic(s, tile, x + w / 2 - 0.31, y0 + 0.3, s=0.62)
        one(s, x + 0.14, y0 + 1.12, w - 0.28, 0.72, t, 14.5, INK, bold=True,
            align=PP_ALIGN.CENTER)
        tx(s, x + 0.24, y0 + 1.95, w - 0.48, hh - 2.15, [
            {"runs": [(d, 11.5, SUB, False)], "align": PP_ALIGN.CENTER,
             "line": 1.26}])
        groups.append([c.shape_id, ic.shape_id])
    tx(s, ML, 5.95, CONTENT_W, 0.5, [
        {"runs": [("Đừng ngại nhé — ", 15, INK, False, True),
                  ("thầy cô luôn lắng nghe chúng em!", 15, INK, True)],
         "align": PP_ALIGN.CENTER}])
    notes(s, "Nhiều bạn ngại nói chuyện với thầy cô - chúng em hiểu cảm giác đó. "
             "Nhưng khi mạnh dạn hỏi, chúng em nhận ra thầy cô luôn sẵn sàng giúp "
             "đỡ. Chỉ cần lễ phép, lắng nghe và thành thật chia sẻ là mọi chuyện "
             "dễ hơn nhiều!")
    return s, groups

def s9_story(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 9)
    header(s, num, TOTAL, kicker="TIẾT 1 · PHẦN CHUẨN BỊ CỦA CHÚNG EM",
           kcolor=BLUE, title="Câu chuyện của em về một lần hợp tác",
           title_size=27, accent_tile="book-open")
    y0, hh = 1.95, 4.35
    groups = [[], []]
    c1 = glass(s, ML, y0, CW2, hh)
    one(s, ML + 0.35, y0 + 0.25, CW2 - 0.7, 0.4, "Kể theo ba phần cho dễ nhớ",
        15.5, INK, bold=True)
    frame = [
        ("1", "Hoàn cảnh", "Chuyện xảy ra lúc nào, ở đâu, có những ai?"),
        ("2", "Chúng em đã làm gì", "Em và các bạn đã phối hợp với nhau ra sao?"),
        ("3", "Kết quả và bài học", "Việc kết thúc thế nào? Em rút ra được điều gì?"),
    ]
    yf = y0 + 0.95
    for n_, t, d in frame:
        num_badge(s, ML + 0.55, yf + 0.18, 0.4, n_, BLUE)
        one(s, ML + 1.05, yf - 0.03, CW2 - 1.4, 0.32, t, 13.5, INK, bold=True)
        tx(s, ML + 1.05, yf + 0.3, CW2 - 1.5, 0.8, [
            {"runs": [(d, 11.5, SUB, False)], "line": 1.18}])
        yf += 1.13
    groups[0] += [c1.shape_id]
    x2 = ML + CW2 + 0.42
    c2 = glass(s, x2, y0, CW2, hh)
    ic = pic(s, "quote", x2 + 0.35, y0 + 0.25, s=0.46)
    one(s, x2 + 0.95, y0 + 0.3, CW2 - 1.3, 0.4, "Câu chuyện của em (kể mẫu)",
        15.5, INK, bold=True)
    story_txt = ("Tuần trước, nhóm em làm mô hình “Ngôi trường xanh”. Bạn Hùng mải "
                 "chơi quên phần việc của mình, cả nhóm lo lắng vì sắp đến hạn nộp. "
                 "Em góp ý nhẹ nhàng, rủ Hùng cùng làm vào giờ ra chơi và nhận giúp "
                 "một phần để bạn bắt kịp nhóm. Mô hình hoàn thành đúng hạn, cả nhóm "
                 "vui lắm! Từ đó Hùng có trách nhiệm hơn. Em hiểu ra: hợp tác là giúp "
                 "nhau cùng tiến bộ, chứ không phải trách móc nhau.")
    tx(s, x2 + 0.35, y0 + 0.95, CW2 - 0.7, hh - 1.2, [
        {"runs": [(story_txt, 11.8, SUB, False)], "line": 1.38}])
    groups[1] += [c2.shape_id, ic.shape_id]
    ph = pill(s, ML, 6.55,
              "Các bạn hãy thay bằng kỉ niệm thật của mình — kể tự nhiên sẽ hay hơn!",
              size=11.5, h=0.4, padx=0.2, color=BLUE, fill_alpha=18, spc=0.1)
    notes(s, "Em xin kể một kỉ niệm nhỏ của em: ... (kể theo 3 phần). Bạn nào cũng "
             "có một kỉ niệm như vậy - hãy thay câu chuyện mẫu bằng chuyện thật của "
             "mình khi trình bày nhé!")
    return s, groups

def s10_divider(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 10)
    chip(s, ML, 0.85, "BƯỚC SANG TIẾT 2 · CÙNG THỰC HÀNH", PINK, size=11,
         h=0.4, padx=0.24, spc=2.0)
    one(s, ML, 1.45, 6.9, 1.6, "Chúng em mời cả lớp\ncùng xử lí 3 tình huống",
        35, INK, bold=True)
    one(s, ML, 3.55, 6.4, 0.75,
        "Các nhóm sẽ phân vai, diễn hoặc thuyết trình. Cả lớp theo dõi và góp ý.",
        13.5, SUB, bold=False, italic=True, line=1.3)
    sit = [("flask-conical", "TH1", "Việc riêng trong giờ thực hành", MINT),
           ("mic", "TH2", "Bạn ốm trước buổi biểu diễn", PINK),
           ("user-plus", "TH3", "Bạn mới chuyển đến nhút nhát", BLUE)]
    groups = [[], []]
    x1 = 7.75
    for i, (tile, tag, t, col) in enumerate(sit):
        y = 1.55 + i * 1.0
        c = glass(s, x1, y, 4.9, 0.8, radius=0.22, alpha=52)
        ic = pic(s, tile, x1 + 0.16, y + 0.17, s=0.46)
        one(s, x1 + 0.78, y + 0.08, 4.0, 0.3, tag, 12, col, bold=True, spc=1.5)
        one(s, x1 + 0.78, y + 0.4, 4.0, 0.34, t, 13, INK, bold=True)
        groups[0].append(c.shape_id)
    c2 = glass(s, ML, 5.1, 6.5, 1.3, radius=0.2)
    tx(s, ML + 0.35, 5.35, 5.9, 0.9, [
        {"runs": [("Cách chúng em tổ chức:  ", 13.5, INK, True),
                  ("chia nhóm → bốc thăm tình huống → chuẩn bị 5 phút → diễn hoặc "
                   "trình bày.", 12.5, SUB, False)], "line": 1.3}])
    groups[1] += [c2.shape_id]
    notes(s, "Bây giờ tới phần vui nhất! Chúng em chia lớp thành các nhóm; mỗi nhóm "
             "bốc thăm 1 trong 3 tình huống, chuẩn bị 5 phút rồi lên phân vai diễn "
             "hoặc thuyết trình. Cả lớp theo dõi và nhận xét. Mời thầy cô cùng cổ vũ!")
    return s, groups

def situation_slide(prs, num, tag, tile, title, story, steps, disc, skill):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, num)
    header(s, num, TOTAL, kicker="TIẾT 2 · TÌNH HUỐNG %s — MỜI CẢ LỚP CÙNG XỬ LÍ" % tag,
           kcolor=PINK if tag == "2" else (MINT if tag == "1" else BLUE),
           title=title, title_size=24, accent_tile=tile)
    groups = []
    c0 = glass(s, ML, 1.75, CONTENT_W, 1.55, radius=0.17)
    ic = pic(s, tile, ML + 0.28, 1.75 + 0.44, s=0.66)
    tx(s, ML + 1.2, 1.75 + 0.24, CONTENT_W - 1.6, 1.05, [
        {"runs": [("Tình huống:  ", 13.5, BLUE, True),
                  (story, 13.5, INK, False)], "line": 1.34}])
    groups.append([c0.shape_id])
    y0 = 3.55
    c1 = glass(s, ML, y0, 7.95, 2.95)
    one(s, ML + 0.3, y0 + 0.2, 7.4, 0.4,
        "Cách nhóm em định xử lí — các bạn bổ sung nhé!", 14.5, INK, bold=True)
    y = y0 + 0.78
    for i, st in enumerate(steps):
        num_badge(s, ML + 0.45, y + 0.06, 0.36, i + 1, BLUE)
        tx(s, ML + 1.0, y - 0.02, 6.8, 0.78, [
            {"runs": [(st, 12, SUB, False)], "line": 1.22}])
        y += 0.7
    groups.append([c1.shape_id])
    x2 = ML + 8.2
    c2 = glass(s, x2, y0, 3.81, 1.42, radius=0.18)
    ic2 = pic(s, "messages-square", x2 + 0.22, y0 + 0.2, s=0.42)
    tx(s, x2 + 0.78, y0 + 0.2, 2.9, 1.1, [
        {"runs": [("Bàn cùng cả lớp", 12.5, INK, True)], "after": 4},
        {"runs": [(disc, 10.5, SUB, False)], "line": 1.2}])
    c3 = glass(s, x2, y0 + 1.58, 3.81, 1.37, radius=0.18)
    ic3 = pic(s, "zap", x2 + 0.22, y0 + 1.78, s=0.42)
    tx(s, x2 + 0.78, y0 + 1.76, 2.9, 1.05, [
        {"runs": [("Chúng em vận dụng", 12.5, INK, True)], "after": 4},
        {"runs": [(skill, 10.5, SUB, False)], "line": 1.2}])
    groups.append([c2.shape_id, c3.shape_id])
    notes(s, "Mời các bạn xem nhóm được phân công diễn hoặc trình bày tình huống %s. "
             "Sau đó cả lớp cho ý kiến: cách xử lí nào khéo léo, tôn trọng và giúp "
             "bạn tốt nhất? Chúng em ghi nhận mọi ý kiến ạ." % tag)
    return s, groups

def s11(prs, num):
    return situation_slide(
        prs, num, "1", "flask-conical",
        "Tình huống 1: Việc riêng trong giờ thực hành",
        "Trong giờ thực hành môn Khoa học tự nhiên, các bạn cùng nhóm với Thanh "
        "đang làm thí nghiệm thì Thanh lấy bài tập Toán ra làm, không tham gia "
        "cùng nhóm.",
        ["Nhắc nhở nhẹ nhàng, thân thiện: “Thanh ơi, giờ đang thực hành, cậu để "
         "bài Toán sang giờ ra chơi làm nhé!”",
         "Mời Thanh tham gia bằng một việc phù hợp: ghi số liệu, quan sát hiện "
         "tượng, thao tác thí nghiệm…",
         "Nếu Thanh vẫn chưa hợp tác: bình tĩnh nhờ thầy cô góp ý để cả nhóm "
         "hoàn thành nhiệm vụ chung."],
        "Vì sao nên nhắc khéo trước, và nhờ thầy cô chỉ là giải pháp cuối?",
        "giao tiếp khéo léo · kiên nhẫn · tôn trọng bạn")

def s12(prs, num):
    return situation_slide(
        prs, num, "2", "mic",
        "Tình huống 2: Bạn ốm trước buổi biểu diễn",
        "Nhóm em đang tập tiết mục văn nghệ chào mừng ngày Nhà giáo Việt Nam. "
        "Chỉ còn hai ngày nữa là biểu diễn thì Mai — người hát chính — bị ốm, "
        "phải nghỉ học.",
        ["Thăm hỏi, động viên Mai mau khỏi ốm; trấn an để bạn không phải lo "
         "lắng hay thấy có lỗi với nhóm.",
         "Họp khẩn cả nhóm: rà soát phần Mai đảm nhận, điều chỉnh tiết mục cho "
         "phù hợp (đổi người hát hoặc đổi tiết mục).",
         "Tập luyện lại với đội hình mới, nhờ thầy cô góp ý nếu cần; sau buổi "
         "diễn cử bạn thăm và báo kết quả cho Mai."],
        "Tinh thần quan trọng nhất của cả nhóm lúc này là gì? Làm sao để Mai "
        "không thấy áy náy?",
        "quan tâm, chia sẻ · linh hoạt · trách nhiệm với nhau")

def s13(prs, num):
    return situation_slide(
        prs, num, "3", "user-plus",
        "Tình huống 3: Bạn mới chuyển đến lớp",
        "Có bạn mới chuyển đến lớp em. Bạn còn nhút nhát, ít nói, ngại tham gia "
        "các hoạt động chung cùng các bạn.",
        ["Chủ động làm quen: chào hỏi, giới thiệu bản thân, rủ bạn chơi trong "
         "giờ giải lao.",
         "Giúp bạn bắt nhịp việc học: hỏi thăm, cho mượn vở, rủ bạn học nhóm; "
         "không trêu chọc, không để bạn đơn độc.",
         "Rủ bạn tham gia từ những việc nhỏ: trực nhật, thể thao, tập tiết mục… "
         "để bạn tự tin dần."],
        "Nếu mình là bạn mới đó, mình muốn được đón nhận như thế nào?",
        "cởi mở, kiên nhẫn · đồng cảm · tạo cảm giác an toàn")

def s14_script(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 14)
    header(s, num, TOTAL, kicker="TIẾT 2 · KỊCH BẢN CỦA NHÓM EM", kcolor=INDIGO,
           title="Kịch bản mẫu — “Giờ thực hành nhóm 3”", title_size=25,
           accent_tile="presentation")
    roles = ["Dẫn chuyện", "Lan — nhóm trưởng", "Thanh", "Minh"]
    xr = ML
    for rl in roles:
        p = pill(s, xr, 1.72, rl, size=10, h=0.34, padx=0.14,
                 color=INDIGO, fill_alpha=16, bold=True, spc=0.3)
        xr = (p.left + p.width) / EMU_IN + 0.16
    c = glass(s, ML, 2.25, CONTENT_W, 4.15, radius=0.14, alpha=44)
    script = [
        ("Dẫn chuyện", INDIGO, "Đã đến giờ thực hành Khoa học tự nhiên. Nhóm Thanh quây quần bên bàn thí nghiệm."),
        ("Lan (nhóm trưởng)", MINT, "Thanh ơi, đang giờ thực hành, cậu cất vở Toán đã nhé!"),
        ("Thanh", INK, "Nhưng bài Toán khó quá, mai phải nộp rồi…"),
        ("Minh", PURPLE, "Hay Thanh ghi số liệu thí nghiệm cho bọn tớ — vừa giúp nhóm, xong sớm lại về làm Toán!"),
        ("Thanh", INK, "Ừ nhỉ, để tớ thử! … Nhóm mình phối hợp ăn ý thật đấy!"),
        ("Lan (nhóm trưởng)", MINT, "Cảm ơn các cậu! Có hợp tác, việc gì cũng xong."),
    ]
    y = 2.62
    for who, col, line in script:
        rounded(s, ML + 0.32, y, 0.07, 0.36, radius=0.5, fill=col, alpha=100)
        one(s, ML + 0.58, y - 0.06, 3.1, 0.3, who, 11.5, col, bold=True)
        tx(s, ML + 3.5, y - 0.05, CONTENT_W - 3.95, 0.65, [
            {"runs": [(line, 12, SUB if who != "Thanh" else INK, False)],
             "line": 1.18}])
        y += 0.62
    groups = [[c.shape_id]]
    ph = pill(s, ML + 0.32, 6.62,
              "Mỗi nhóm tự viết lời thoại theo cách nói của mình — diễn tự nhiên là hay nhất!",
              size=11.5, h=0.4, padx=0.2, color=INDIGO, fill_alpha=16, spc=0.1)
    notes(s, "Đây là kịch bản mẫu nhóm em chuẩn bị cho tình huống 1. Các nhóm có thể "
             "viết lời thoại theo ngôn ngữ của mình cho tự nhiên. Khi diễn: nói to, "
             "rõ, thể hiện cảm xúc; người dẫn chuyện đứng giữa giới thiệu.")
    return s, groups

def s15_lessons(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 15)
    header(s, num, TOTAL, kicker="TIẾT 2 · SAU KHI CÙNG THỰC HÀNH", kcolor=MINT,
           title="Điều chúng em học được hôm nay", title_size=29,
           accent_tile="list-checks")
    cards = [
        ("handshake", "Nhắc khéo, không trách móc",
         "Một lời nhắc nhẹ nhàng và tôn trọng sẽ dễ nghe hơn nhiều."),
        ("heart", "Quan tâm đúng lúc",
         "Bạn ốm, bạn buồn… hãy chia sẻ và động viên trước tiên."),
        ("user-plus", "Đón bạn mới cởi mở",
         "Giúp bạn hòa nhập từ việc nhỏ: chào hỏi, rủ chơi, cùng học."),
        ("users", "Từ “tôi” đến “chúng ta”",
         "Việc chung là của chung: cùng làm, cùng chịu trách nhiệm."),
    ]
    w = CW4; gap = 0.3
    y0, hh = 2.0, 3.35
    groups = []
    for i, (tile, t, d) in enumerate(cards):
        x = ML + i * (w + gap)
        c = glass(s, x, y0, w, hh, radius=0.18)
        ic = pic(s, tile, x + w / 2 - 0.3, y0 + 0.3, s=0.6)
        one(s, x + 0.12, y0 + 1.08, w - 0.24, 0.7, t, 14, INK, bold=True,
            align=PP_ALIGN.CENTER)
        tx(s, x + 0.2, y0 + 1.85, w - 0.4, hh - 2.05, [
            {"runs": [(d, 11, SUB, False)], "align": PP_ALIGN.CENTER,
             "line": 1.22}])
        groups.append([c.shape_id, ic.shape_id])
    tx(s, ML, 5.8, CONTENT_W, 0.5, [
        {"runs": [("Ứng xử đẹp không chỉ nằm ở lời nói, mà ở hành động mỗi ngày.",
                   14.5, INK, True)], "align": PP_ALIGN.CENTER}])
    notes(s, "Từ ba tình huống vừa rồi, chúng em rút ra bốn điều: nhắc khéo thay vì "
             "trách móc; quan tâm bạn đúng lúc; đón bạn mới cởi mở; và chuyển từ "
             "việc của tôi thành việc của chúng ta. Các bạn thấy điều nào gần gũi "
             "với mình nhất?")
    return s, groups

def s16_message(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 16)
    header(s, num, TOTAL, kicker="THÔNG ĐIỆP CỦA CHỦ ĐỀ", kcolor=PINK,
           title="Mỗi bạn một thế mạnh — lớp học gắn kết", title_size=27,
           accent_tile="heart-handshake")
    groups = [[], [], []]
    # trái: vòng kính + tile trung tâm
    for cx, cy, d, a in [(3.05, 3.9, 3.4, 24), (3.05, 3.9, 2.7, 20),
                         (3.05, 3.9, 2.0, 16)]:
        oval(s, cx, cy, d, fill=WHITE, alpha=a, border=WHITE, border_w=1.3,
             border_alpha=75)
    # Giữ SF Symbol ở kích thước nét, đủ khoảng thở trong vòng kính.
    cen = pic(s, "heart-handshake", 3.05 - 0.45, 3.9 - 0.45, s=0.9)
    # 4 tile quanh vòng
    orbit = [("message-circle", 0.95, 3.15), ("ear", 4.75, 2.5),
             ("users", 5.15, 4.6), ("sparkles", 0.95, 5.35)]
    ids = [cen.shape_id]
    for t, x, y in orbit:
        it = pic(s, t, x, y, s=0.55)
        ids.append(it.shape_id)
    groups[0] += ids
    caps = [("Trò chuyện", 0.7, 3.7), ("Lắng nghe", 4.35, 3.15),
            ("Cùng chơi", 4.55, 5.3), ("Cổ vũ nhau", 0.7, 5.95)]
    for t, x, y in caps:
        chip(s, x, y, t, SUB, size=9.5, h=0.3, padx=0.12, fill_alpha=26)
    # phải: nội dung
    x2 = 6.85
    tx(s, x2, 2.0, 5.85, 3.3, [
        {"runs": [("Cũng như một chiếc điện thoại cao cấp, mỗi phần đều có vai trò "
                   "riêng của nó.", 15, INK, True)], "after": 12, "line": 1.28},
        {"runs": [("Lớp mình cũng vậy: bạn giỏi lên kế hoạch, bạn khéo lắng nghe, "
                   "bạn nhiệt tình giúp đỡ… Mỗi người một điểm mạnh.", 13.5, SUB,
                   False)], "after": 12, "line": 1.32},
        {"runs": [("Khi ai cũng được phát huy thế mạnh và sẵn sàng hỗ trợ nhau, "
                   "việc gì khó đến mấy chúng ta cũng làm được.", 13.5, SUB, False)],
         "line": 1.32},
    ])
    groups[1] = [0]
    cb = glass(s, x2, 5.5, 5.85, 1.35, radius=0.2)
    tx(s, x2 + 0.28, 5.72, 5.3, 0.95, [
        {"runs": [("Tuần này, mỗi bạn chọn một điểm mạnh của mình để góp cho lớp "
                   "nhé?", 13.5, INK, True)], "line": 1.25}])
    groups[2] = [cb.shape_id]
    notes(s, "Chúng em ví lớp như một cỗ máy hoàn hảo: mỗi bạn một vai trò. Hãy chọn "
             "điều mình làm tốt nhất để giúp lớp - chẳng hạn lắng nghe, rủ bạn chơi, "
             "hay động viên mọi người. Bạn nào giơ tay kể điểm mạnh của mình nào?")
    return s, groups

def s17_pledge(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 17)
    header(s, num, TOTAL, kicker="LỜI HỨA NHỎ · VIỆC LÀM TO", kcolor=BLUE,
           title="Tuần này, lớp mình cùng cam kết nhé!", title_size=29,
           accent_tile="list-checks")
    y0, hh = 2.0, 3.8
    c = glass(s, ML, y0, CONTENT_W, hh, radius=0.16)
    acts = [
        "Chủ động chào, mỉm cười và hỏi thăm ít nhất một bạn mỗi ngày.",
        "Sẵn sàng giúp đỡ bạn trong học tập hoặc khi bạn gặp khó khăn.",
        "Khi làm việc nhóm: nhắc khéo, lắng nghe, không để ai bị bỏ rơi.",
        "Nếu băn khoăn điều gì, mạnh dạn trao đổi với thầy cô.",
    ]
    y = y0 + 0.6
    for i, a in enumerate(acts):
        num_badge(s, ML + 0.6, y + 0.14, 0.42, i + 1, [BLUE, TEAL, PURPLE, ORANGE][i])
        # Chừa cột phải cho icon checklist; chữ và icon không bao giờ giao nhau.
        tx(s, ML + 1.3, y + 0.02, CONTENT_W - 2.95, 0.7, [
            {"runs": [(a, 13.5, SUB, False)], "line": 1.22}])
        y += 0.74
    groups = [[c.shape_id]]
    pic(s, "list-checks", ML + CONTENT_W - 0.95, y0 + 0.46, s=0.58)
    ph = pill(s, 0, 6.35, "Chọn một việc thôi cũng được — tuần sau chúng em sẽ hỏi kết quả nha!",
              size=12.5, h=0.48, padx=0.26, color=INK, fill_alpha=26, spc=0.2)
    ph.left = Emu(int(SW * EMU_IN / 2 - ph.width / 2))
    notes(s, "Để lời nói đi đôi với việc làm, chúng em đề nghị mỗi bạn chọn một trong "
             "bốn hành động và thực hiện trong tuần này. Tuần sau ai cũng có thể chia "
             "sẻ kết quả của mình. Chúng em tin lớp mình làm được!")
    return s, groups

def s18_thanks(prs, num):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bg(s, 18)
    # trung tâm: vòng kính + tile
    for cx, cy, d, a in [(SW / 2, 1.95, 2.2, 26), (SW / 2, 1.95, 1.7, 22)]:
        oval(s, cx, cy, d, fill=WHITE, alpha=a, border=WHITE, border_w=1.3,
             border_alpha=78)
    pic(s, "heart-handshake", SW / 2 - 0.4, 1.55, s=0.8)
    pic(s, "sparkles", SW / 2 + 0.85, 1.1, s=0.42)
    pic(s, "party-popper", SW / 2 - 1.35, 0.95, s=0.4)
    one(s, 1.2, 3.44, 10.93, 0.6,
        "Cảm ơn thầy cô và các bạn đã lắng nghe!", 27, INK, bold=True,
        align=PP_ALIGN.CENTER)
    one(s, 1.6, 4.35, 10.13, 0.5,
        "Chúc lớp mình luôn hòa đồng, hợp tác và cùng nhau tiến bộ mỗi ngày.",
        15, SUB, bold=False, italic=True, align=PP_ALIGN.CENTER)
    groups = [[], [0], [0]]
    # chips thông tin
    names = ["[Tên em / nhóm em]", "Lớp [..]", "Trường [..]"]
    widths = [tw(t, 11.5) + 0.34 for t in names]
    totalw = sum(widths) + 0.26 * 2
    xc = SW / 2 - totalw / 2
    for t, wt in zip(names, widths):
        p = pill(s, xc, 5.15, t, size=11.5, h=0.44, padx=0.17, color=INK,
                 bold=False, fill_alpha=24, spc=0.2)
        xc += p.width / EMU_IN + 0.26
    one(s, 1.6, 6.1, 10.13, 0.4,
        "Các bạn có câu hỏi gì cho nhóm em không? Chúng em sẵn sàng trao đổi ạ!",
        13, FAINT, bold=False, align=PP_ALIGN.CENTER)
    groups[0] = []
    notes(s, "Cảm ơn thầy cô và các bạn đã lắng nghe và tham gia nhiệt tình! Nếu ai "
             "có câu hỏi hay góp ý, chúng em rất vui được trao đổi. Đừng quên cam kết "
             "nhỏ của cả lớp trong tuần này nhé!")
    return s, [[], [], []]

# ================================================================ main
def main():
    prs = Presentation()
    prs.slide_width = Emu(int(SW * EMU_IN))
    prs.slide_height = Emu(int(SH * EMU_IN))

    builders = [
        lambda: s1_cover(prs),
        lambda: s2_roadmap(prs, 2),
        lambda: s3_goals(prs, 3),
        lambda: s4_warmup(prs, 4),
        lambda: s5_concept(prs, 5),
        lambda: s6_keys(prs, 6),
        lambda: s7_teamwork(prs, 7),
        lambda: s8_teachers(prs, 8),
        lambda: s9_story(prs, 9),
        lambda: s10_divider(prs, 10),
        lambda: s11(prs, 11),
        lambda: s12(prs, 12),
        lambda: s13(prs, 13),
        lambda: s14_script(prs, 14),
        lambda: s15_lessons(prs, 15),
        lambda: s16_message(prs, 16),
        lambda: s17_pledge(prs, 17),
        lambda: s18_thanks(prs, 18),
    ]
    all_slides = []
    for b in builders:
        sl, _anim = b()
        all_slides.append(sl)
        add_transition(sl, "med")
        # group: mọi shape không nằm trong nền/header (tên ~bg/~hdr)
        body = [sp.shape_id for sp in sl.shapes
                if not sp.name.startswith("~")]
        if body:
            add_timing(sl, [body])

    props = prs.core_properties
    props.title = "Phát triển mối quan hệ hòa đồng, hợp tác với thầy cô và bạn bè (bản sáng)"
    props.author = "Học sinh trình bày – HĐTN THCS"
    props.subject = "Hoạt động trải nghiệm, hướng nghiệp"
    props.language = "vi-VN"

    out = "Phát triển mối quan hệ hòa đồng, hợp tác - HĐTN lớp 7 (học sinh thuyết trình).pptx"
    prs.save(out)
    print("saved:", out, "| slides:", len(prs.slides._sldIdLst))
    # ----- QA đo tràn chữ -----
    import os
    from PIL import ImageFont as IF
    reg = IF.truetype(_DEJ, 100)
    bold = IF.truetype(_DEJ_B, 100)
    issues = 0
    for si, sl in enumerate(all_slides, 1):
        for sh in sl.shapes:
            if not sh.has_text_frame or not sh.text_frame.text.strip():
                continue
            tf = sh.text_frame
            bw = (sh.width - tf.margin_left - tf.margin_right) / 12700.0
            bh = (sh.height - tf.margin_top - tf.margin_bottom) / 12700.0
            th = 0.0
            any_wrap = False
            for p in tf.paragraphs:
                t = "".join(r.text for r in p.runs)
                if not t.strip():
                    continue
                sz = max([r.font.size.pt if r.font.size else 12 for r in p.runs])
                lh = sz * (p.line_spacing if p.line_spacing else 1.0)
                nl = 1; cur = 0.0
                for r in p.runs:
                    f = bold if r.font.bold else reg
                    for ch in r.text:
                        if ch in "\n\v\t":
                            continue
                        wp = f.getlength(ch) / 100.0 * sz * 1.04
                        if cur + wp > bw and cur > 0:
                            nl += 1; cur = wp
                        else:
                            cur += wp
                if nl > 1:
                    any_wrap = True
                th += nl * lh * 1.08 + (p.space_after.pt if p.space_after else 0)
            # chỉ cảnh báo khi thực sự xuống dòng (dòng đơn nằm ngang luôn vẽ được)
            if any_wrap and th > bh * 1.08:
                issues += 1
                sn = tf.text.strip().replace("\n", " ")[:40]
                print("  [tràn?] slide %d '%s' cần %.0fpt / có %.0fpt" % (si, sn, th, bh))
    print("QA tràn chữ:", issues)
    return out

if __name__ == "__main__":
    main()
