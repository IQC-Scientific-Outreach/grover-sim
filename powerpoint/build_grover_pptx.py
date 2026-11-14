"""
Build a PowerPoint deck that steps through Grover's search with the Morph
transition: 3 qubits, marked state |101>, one slide per oracle or diffusion.

Each click advances one slide; Morph rotates the state vector, turns the
mirror line and resizes the probability bars between consecutive slides.
The drawing sits in the left half of a 16:9 slide; the right half is blank.

Requires python-pptx (pip install python-pptx).  Run:
    python3 build_grover_pptx.py [output.pptx]

Morph needs PowerPoint 2019, Microsoft 365 or PowerPoint for the web;
other viewers fall back to a fade between slides.
"""
import math
import sys
from copy import deepcopy

from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE_DASH_STYLE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- the maths
N_QUBITS = 3
N = 2 ** N_QUBITS
MARKED = 0b101
THETA = math.asin(1 / math.sqrt(N))          # sin(theta) = 1/sqrt(N)

# The state always sits at angle m*theta (m odd).  Oracle: m -> -m.
# Diffusion: m -> 2 - m.  One oracle + one diffusion is one iteration.
STEPS = [("start", 1)]
m = 1
for _ in range(3):
    m = -m
    STEPS.append(("oracle", m))
    m = 2 - m
    STEPS.append(("diffusion", m))


def prob_marked(mm):
    return math.sin(mm * THETA) ** 2


def short_delta(a_from, a_to):
    """Signed angle (radians) the short way round, as the web lab animates."""
    d = (a_to - a_from) % (2 * math.pi)
    if d > math.pi:
        d -= 2 * math.pi
    return d


# ------------------------------------------------------------ the look
CRIMSON = RGBColor(0x96, 0x17, 0x2E)
BLUE = RGBColor(0x00, 0x5D, 0x7E)
INK = RGBColor(0x00, 0x00, 0x05)
INK_SOFT = RGBColor(0x5B, 0x56, 0x4E)
AXIS = RGBColor(0x8A, 0x84, 0x78)
GRID = RGBColor(0xE6, 0xE0, 0xD2)
ORANGE = RGBColor(0xD3, 0x54, 0x00)
ORANGE_TEXT = RGBColor(0xA8, 0x43, 0x00)

MATH_FONT = "Cambria Math"   # ships with Office on Windows and Mac; has the ket brackets
TEXT_FONT = "Verdana"

# Geometry, in inches.  Slide is 13.333 x 7.5; the drawing uses the left half.
SLIDE_W, SLIDE_H = 13.333, 7.5
CX, CY, R = 3.25, 2.55, 1.6                  # Grover-plane origin and circle radius
CHART_L, CHART_R = 0.95, 6.3                 # bar chart plot area
CHART_TOP, CHART_BASE = 5.05, 6.65


def emu(v):
    return Emu(int(round(v * 914400)))


def plane_xy(angle, r):
    """Point at math angle (radians, counter-clockwise from +x) and radius r, in slide inches."""
    return CX + r * math.cos(angle), CY - r * math.sin(angle)


def ppt_rotation(angle):
    """PowerPoint rotates clockwise in degrees; our angles are counter-clockwise radians."""
    return (-math.degrees(angle)) % 360


A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def set_custom_geometry(shape, paths_xml):
    """Swap a shape's preset geometry for custom paths (drawn inside its own box)."""
    sp_pr = shape._element.spPr
    prst = sp_pr.find(qn("a:prstGeom"))
    geom = etree.fromstring(
        '<a:custGeom xmlns:a="%s"><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/>'
        '<a:rect l="0" t="0" r="r" b="b"/><a:pathLst>%s</a:pathLst></a:custGeom>' % (A_NS, paths_xml)
    )
    prst.addprevious(geom)
    sp_pr.remove(prst)


def name(shape, nm):
    shape.name = nm
    return shape


def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, nm=None):
    """runs: list of (text, dict(size, font, color, bold, italic, sup))."""
    tb = slide.shapes.add_textbox(emu(x), emu(y), emu(w), emu(h))
    tf = tb.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]
    p.alignment = align
    for text, st in runs:
        r = p.add_run()
        r.text = text
        f = r.font
        f.size = Pt(st.get("size", 12))
        f.name = st.get("font", TEXT_FONT)
        f.bold = st.get("bold", False)
        f.italic = st.get("italic", False)
        f.color.rgb = st.get("color", INK)
        if st.get("sup"):
            r._r.get_or_add_rPr().set("baseline", "30000")
    if nm:
        tb.name = nm
    return tb


def add_line(slide, x1, y1, x2, y2, color, width_pt, dash=None, arrow_end=False, nm=None):
    ln_shape = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, emu(x1), emu(y1), emu(x2), emu(y2))
    ln_shape.line.color.rgb = color
    ln_shape.line.width = Pt(width_pt)
    if dash:
        ln_shape.line.dash_style = dash
    if arrow_end:
        ln = ln_shape.line._get_or_add_ln()
        tail = etree.SubElement(ln, qn("a:tailEnd"))
        tail.set("type", "triangle")
        tail.set("w", "med")
        tail.set("len", "med")
    ln_shape.shadow.inherit = False
    if nm:
        ln_shape.name = nm
    return ln_shape


# ------------------------------------------------------------ static drawing
def draw_plane_static(slide):
    # unit circle
    c = slide.shapes.add_shape(MSO_SHAPE.OVAL, emu(CX - R), emu(CY - R), emu(2 * R), emu(2 * R))
    c.fill.background()
    c.line.color.rgb = INK_SOFT
    c.line.width = Pt(2.25)
    c.shadow.inherit = False
    name(c, "!!circle")
    # axes with arrowheads
    add_line(slide, CX - R - 0.22, CY, CX + R + 0.3, CY, AXIS, 1.75, arrow_end=True, nm="!!axis-x")
    add_line(slide, CX, CY + R + 0.22, CX, CY - R - 0.3, AXIS, 1.75, arrow_end=True, nm="!!axis-y")
    # axis labels
    add_text(slide, CX - 1.0, CY - R - 0.88, 2.0, 0.4,
             [("|101⟩", dict(size=22, font=MATH_FONT, color=CRIMSON, bold=True))], PP_ALIGN.CENTER, nm="!!lbl-w")
    add_text(slide, CX + 0.12, CY - R - 0.47, 1.2, 0.22,
             [("marked", dict(size=10, color=INK_SOFT))], nm="!!lbl-marked")
    add_text(slide, CX + R + 0.36, CY - 0.22, 1.0, 0.4,
             [("|w", dict(size=20, font=MATH_FONT)), ("⊥", dict(size=20, font=MATH_FONT, sup=True)),
              ("⟩", dict(size=20, font=MATH_FONT))], nm="!!lbl-wperp")
    add_text(slide, CX + R + 0.36, CY + 0.2, 1.0, 0.22,
             [("7 others", dict(size=10, color=INK_SOFT))], nm="!!lbl-others")
    # fixed |+++> reference line at theta
    x2, y2 = plane_xy(THETA, R)
    add_line(slide, CX, CY, x2, y2, INK_SOFT, 1.5, dash=MSO_LINE_DASH_STYLE.DASH, nm="!!ref-line")
    lx, ly = plane_xy(THETA + 0.05, R + 0.28)
    add_text(slide, lx - 0.4, ly - 0.17, 0.8, 0.34,
             [("|+++⟩", dict(size=14, font=MATH_FONT, color=INK_SOFT))], PP_ALIGN.CENTER, nm="!!lbl-plus")


def draw_chart_static(slide):
    ph = CHART_BASE - CHART_TOP
    for i in range(6):
        v = i / 5
        y = CHART_BASE - v * ph
        add_line(slide, CHART_L, y, CHART_R, y, GRID, 0.75, nm="!!grid%d" % i)
        add_text(slide, CHART_L - 0.62, y - 0.1, 0.55, 0.2,
                 [("%d%%" % round(v * 100), dict(size=9, color=INK_SOFT))], PP_ALIGN.RIGHT, nm="!!ytick%d" % i)
    add_line(slide, CHART_L, CHART_TOP, CHART_L, CHART_BASE, AXIS, 0.75, nm="!!y-axis")
    # uniform 1/N reference
    yr = CHART_BASE - (1 / N) * ph
    add_line(slide, CHART_L, yr, CHART_R, yr, INK_SOFT, 1.0, dash=MSO_LINE_DASH_STYLE.DASH, nm="!!uniform")
    add_text(slide, CHART_R - 0.5, yr - 0.22, 0.5, 0.2,
             [("1/N", dict(size=9, color=INK_SOFT, bold=True))], PP_ALIGN.RIGHT, nm="!!lbl-uniform")
    # baseline and outcome labels
    add_line(slide, CHART_L, CHART_BASE, CHART_R, CHART_BASE, INK, 1.25, nm="!!x-axis")
    slot = (CHART_R - CHART_L) / N
    for i in range(N):
        xc = CHART_L + (i + 0.5) * slot
        add_text(slide, xc - slot / 2, CHART_BASE + 0.05, slot, 0.25,
                 [("|%s⟩" % format(i, "03b"),
                   dict(size=11, font=MATH_FONT, color=CRIMSON if i == MARKED else INK, bold=(i == MARKED)))],
                 PP_ALIGN.CENTER, nm="!!xlbl%d" % i)
    add_text(slide, CHART_L, CHART_BASE + 0.33, CHART_R - CHART_L, 0.22,
             [("Probability of each measurement outcome", dict(size=10, color=INK_SOFT))], PP_ALIGN.CENTER, nm="!!x-title")


# ------------------------------------------------------------ moving parts
def draw_vector(slide, angle):
    """Orange state vector.  Its box is centred on the origin so rotation pivots there;
    only the right half of the box is drawn."""
    head_len, box_h = 0.2, 0.17
    w = 2 * R
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(CX - R), emu(CY - box_h / 2), emu(w), emu(box_h))
    W, H, HL = int(emu(w)), int(emu(box_h)), int(emu(head_len))
    set_custom_geometry(shp, (
        '<a:path w="{W}" h="{H}" fill="none"><a:moveTo><a:pt x="{c}" y="{m}"/></a:moveTo>'
        '<a:lnTo><a:pt x="{s}" y="{m}"/></a:lnTo></a:path>'
        '<a:path w="{W}" h="{H}" stroke="0"><a:moveTo><a:pt x="{W}" y="{m}"/></a:moveTo>'
        '<a:lnTo><a:pt x="{b}" y="0"/></a:lnTo><a:lnTo><a:pt x="{b}" y="{H}"/></a:lnTo><a:close/></a:path>'
    ).format(W=W, H=H, c=W // 2, m=H // 2, s=W - HL + 2000, b=W - HL))
    shp.fill.solid()
    shp.fill.fore_color.rgb = ORANGE
    shp.line.color.rgb = ORANGE
    shp.line.width = Pt(4.5)
    shp.rotation = ppt_rotation(angle)
    shp.shadow.inherit = False
    return name(shp, "!!state-vector")


def draw_mirror(slide, angle):
    """Black dashed line through the origin: the axis the current step reflects about."""
    w, box_h = 2 * (R + 0.15), 0.1
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(CX - w / 2), emu(CY - box_h / 2), emu(w), emu(box_h))
    W, H = int(emu(w)), int(emu(box_h))
    set_custom_geometry(shp, (
        '<a:path w="{W}" h="{H}" fill="none"><a:moveTo><a:pt x="0" y="{m}"/></a:moveTo>'
        '<a:lnTo><a:pt x="{W}" y="{m}"/></a:lnTo></a:path>').format(W=W, H=H, m=H // 2))
    shp.fill.background()
    shp.line.color.rgb = INK
    shp.line.width = Pt(2.25)
    shp.line.dash_style = MSO_LINE_DASH_STYLE.DASH
    shp.rotation = ppt_rotation(angle)
    shp.shadow.inherit = False
    return name(shp, "!!mirror")


def draw_sweep(slide, a_from, a_to, tag):
    """Dashed arc with an arrowhead showing which way the last step turned the state."""
    r = 0.42 * R
    d = a_to - a_from
    n = max(8, int(abs(math.degrees(d)) / 4))
    pts = [plane_xy(a_from + d * k / n, r) for k in range(n + 1)]
    fb = slide.shapes.build_freeform(int(emu(pts[0][0])), int(emu(pts[0][1])), scale=1.0)
    fb.add_line_segments([(int(emu(x)), int(emu(y))) for x, y in pts[1:-1]] +
                         [(int(emu(pts[-1][0] - 0.0)), int(emu(pts[-1][1])))], close=False)
    arc = fb.convert_to_shape()
    arc.fill.background()
    arc.line.color.rgb = INK
    arc.line.width = Pt(2.25)
    arc.line.dash_style = MSO_LINE_DASH_STYLE.SQUARE_DOT
    arc.shadow.inherit = False
    name(arc, "sweep-arc-" + tag)
    # arrowhead at the end of the arc, pointing along the direction of travel
    ex, ey = pts[-1]
    sgn = 1 if d > 0 else -1
    tx, ty = -math.sin(a_to) * sgn, -math.cos(a_to) * sgn       # tangent in slide coords (y down)
    nx, ny = -ty, tx
    L, Wd = 0.13, 0.075
    bx, by = ex - tx * L, ey - ty * L
    tri = [(ex, ey), (bx + nx * Wd, by + ny * Wd), (bx - nx * Wd, by - ny * Wd)]
    fb2 = slide.shapes.build_freeform(int(emu(tri[0][0])), int(emu(tri[0][1])), scale=1.0)
    fb2.add_line_segments([(int(emu(x)), int(emu(y))) for x, y in tri[1:]], close=True)
    head = fb2.convert_to_shape()
    head.fill.solid()
    head.fill.fore_color.rgb = INK
    head.line.fill.background()
    head.shadow.inherit = False
    name(head, "sweep-head-" + tag)


def draw_bars(slide, mm):
    pm = prob_marked(mm)
    pu = (1 - pm) / (N - 1)
    ph = CHART_BASE - CHART_TOP
    slot = (CHART_R - CHART_L) / N
    bw = slot * 0.72
    for i in range(N):
        p = pm if i == MARKED else pu
        h = max(p * ph, 0.004)
        xc = CHART_L + (i + 0.5) * slot
        bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, emu(xc - bw / 2), emu(CHART_BASE - h), emu(bw), emu(h))
        bar.fill.solid()
        bar.fill.fore_color.rgb = CRIMSON if i == MARKED else BLUE
        bar.line.fill.background()
        bar.shadow.inherit = False
        name(bar, "!!bar%d" % i)


def draw_caption(slide, kind, mm):
    k = (abs(mm) - 1) / 2 + (0.5 if mm < 0 else 0)
    k_txt = ("%g" % k)
    p_txt = "%.1f%%" % (prob_marked(mm) * 100)
    tag = {"start": "START", "oracle": "ORACLE", "diffusion": "DIFFUSION"}[kind]
    desc = {
        "start": "Equal superposition |+++⟩",
        "oracle": "Flip the sign of |101⟩: reflect about |w⊥⟩",
        "diffusion": "Inversion about the mean: reflect about |+++⟩",
    }[kind]
    add_text(slide, 0.45, 4.42, 6.0, 0.28,
             [(tag + "   ", dict(size=12, bold=True)), (desc, dict(size=11, color=INK_SOFT))],
             PP_ALIGN.LEFT, nm="caption-" + kind + str(mm))
    add_text(slide, 0.45, 4.7, 6.0, 0.24,
             [("k = ", dict(size=11, color=INK_SOFT)), (k_txt, dict(size=11, bold=True, color=BLUE)),
              ("     angle = ", dict(size=11, color=INK_SOFT)),
              ("%.1f°" % (((math.degrees(mm * THETA) + 180) % 360) - 180), dict(size=11, bold=True, color=ORANGE_TEXT)),
              ("     P(marked) = ", dict(size=11, color=INK_SOFT)), (p_txt, dict(size=11, bold=True, color=CRIMSON))],
             PP_ALIGN.LEFT, nm="readout-" + kind + str(mm))
    return k, p_txt


# ------------------------------------------------------------ transitions
MORPH_XML = (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p159">'
    '<p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="slow" p14:dur="1000">'
    '<p159:morph option="byObject"/></p:transition></mc:Choice>'
    '<mc:Fallback><p:transition xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" spd="med">'
    '<p:fade/></p:transition></mc:Fallback></mc:AlternateContent>'
)


def add_morph(slide):
    sld = slide._element
    anchor = sld.find(qn("p:clrMapOvr"))
    if anchor is None:
        anchor = sld.find(qn("p:cSld"))
    anchor.addnext(etree.fromstring(MORPH_XML))


# ------------------------------------------------------------ build
def build(path):
    prs = Presentation()
    prs.slide_width, prs.slide_height = emu(SLIDE_W), emu(SLIDE_H)
    blank = prs.slide_layouts[6]
    prev_angle = None
    for idx, (kind, mm) in enumerate(STEPS):
        s = prs.slides.add_slide(blank)
        angle = mm * THETA
        draw_plane_static(s)
        if kind != "start":
            draw_mirror(s, 0.0 if kind == "oracle" else THETA)
            draw_sweep(s, prev_angle, prev_angle + short_delta(prev_angle, angle), "%d" % idx)
        draw_vector(s, angle if prev_angle is None else prev_angle + short_delta(prev_angle, angle))
        draw_chart_static(s)
        draw_bars(s, mm)
        k, p_txt = draw_caption(s, kind, mm)
        notes = {
            "start": "Start: the equal superposition |+++>. Every outcome, including the marked |101>, has probability 1/8 = 12.5%.",
            "oracle": "Oracle: flips the sign of the |101> amplitude, a reflection about the horizontal axis. Probabilities do not change.",
            "diffusion": "Diffusion: reflection about |+++>, i.e. inversion about the mean. Completes iteration %g." % k,
        }[kind]
        if kind == "diffusion" and mm == 7:
            notes += " This is one iteration too many: the probability falls from 94.5% to 33.0% (overshoot)."
        s.notes_slide.notes_text_frame.text = notes + " P(marked) = " + p_txt + "."
        # Every colour and line is set explicitly, so drop the theme style references
        # (they would otherwise add the default theme shadow in some viewers).
        for st in list(s._element.iter(qn("p:style"))):
            st.getparent().remove(st)
        if idx > 0:
            add_morph(s)
        prev_angle = angle if prev_angle is None else prev_angle + short_delta(prev_angle, angle)
    prs.save(path)
    return len(STEPS)


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "grover-3-qubit-morph.pptx"
    n = build(out)
    print("wrote %s (%d slides)" % (out, n))
