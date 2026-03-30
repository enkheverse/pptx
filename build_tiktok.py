from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY   = RGBColor(0x1B, 0x2A, 0x4A)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GOLD   = RGBColor(0xC9, 0xA8, 0x4C)
GREY   = RGBColor(0x8A, 0x93, 0xA6)
LGREY  = RGBColor(0xF2, 0xF4, 0xF7)   # very light bg tint
DKGREY = RGBColor(0x2C, 0x3E, 0x50)   # body text on white

SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# ── xml helpers ───────────────────────────────────────────────────────────────
def fill(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color

def no_border(shape):
    from pptx.oxml.ns import qn
    from lxml import etree
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    ln = spPr.find(qn('a:ln'))
    if ln is None:
        ln = etree.SubElement(spPr, qn('a:ln'))
    noFill = ln.find(qn('a:noFill'))
    if noFill is None:
        etree.SubElement(ln, qn('a:noFill'))

def txbox(slide, l, t, w, h):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf  = box.text_frame
    tf.word_wrap = True
    return box, tf

def run(para, text, size, bold=False, italic=False, color=DKGREY, align=None):
    if align is not None:
        para.alignment = align
    r = para.add_run()
    r.text = text
    r.font.name  = "Calibri"
    r.font.size  = Pt(size)
    r.font.bold  = bold
    r.font.italic= italic
    r.font.color.rgb = color
    return r

def solo(para, text, size, bold=False, italic=False, color=DKGREY, align=None):
    """Set full paragraph to a single run (clears first)."""
    para.text = ""
    return run(para, text, size, bold=bold, italic=italic, color=color, align=align)

# ── Slide data ────────────────────────────────────────────────────────────────
SLIDES = [
    {
        "num":   1,
        "topic": "Foundations of Business Communication",
        "title": "When Silence Becomes a Strategy",
        "icon":  "📡",
        "bullets": [
            "TikTok operates in 150+ countries — 1B+ active users",
            "Channels: congressional testimony, press releases, in-app notifications, CEO interviews",
            "Key barrier: trust deficit — US regulators vs ByteDance ownership",
            "Goal: reassure Western stakeholders without alienating Chinese parent",
            "Result: mixed messaging that satisfied no one",
        ],
        "ethical": "Strategic ambiguity is not neutral — it erodes trust",
        "ref":    "US Congress Hearing, Shou Zi Chew Testimony (2023)",
    },
    {
        "num":   2,
        "topic": "Interpersonal & Cross-Cultural Communication",
        "title": "Lost in Translation — East vs West",
        "icon":  "🌐",
        "bullets": [
            "Stakeholders: US Senate, ByteDance (Beijing), 170M US users, advertisers",
            "Cultural gap: collectivist/state-aligned (China) vs individual privacy/free market (USA)",
            "Chew's testimony: calm and formal — misread by senators as evasive",
            "High-context (China) vs low-context (USA) communication clash",
            "No cultural bridge strategy was in place",
        ],
        "ethical": "Cross-cultural communication failures have real policy consequences",
        "ref":    "Hofstede, G. (2001) Culture's Consequences, Sage",
    },
    {
        "num":   3,
        "topic": "Written Communication",
        "title": "The Letter That Didn't Convince Anyone",
        "icon":  "✉️",
        "bullets": [
            "TikTok submitted formal written responses to Congress pre-hearing",
            "Tone: defensive, technical, overly legal — missing empathy and plain language",
            "ACCA standard: clarity, structure, audience-appropriate tone",
            "BETTER MEMO | Subject: TikTok's Commitment to US Data Security | US user data is ring-fenced on domestic servers under Project Texas. No foreign entity has access. Audit trail enclosed.",
        ],
        "ethical": "Burying key facts in legal language = obscuring accountability",
        "ref":    "Guffey, M. (2016) Business Communication, Cengage; TikTok Congressional Brief (2023)",
    },
    {
        "num":   4,
        "topic": "Persuasive & Strategic Communication",
        "title": "Project Texas — Persuasion or PR?",
        "icon":  "🎯",
        "bullets": [
            "Project Texas: $1.5B plan to store all US data on domestic servers",
            "Ethos: 'We've invested $1.5B to prove our commitment'",
            "Pathos: '150M Americans use TikTok — a ban hurts creators'",
            "Logos: independent audits and Oracle partnership as evidence",
            "Goal: delay or prevent ban through credibility-building",
        ],
        "ethical": "Persuasion built on incomplete disclosure is manipulation dressed as transparency",
        "ref":    "Aristotle (350 BC) Rhetoric; Forbes (2023) 'Project Texas Explained'",
    },
    {
        "num":   5,
        "topic": "Corporate Governance & CSR",
        "title": "Who Actually Controls TikTok?",
        "icon":  "🏛️",
        "bullets": [
            "ByteDance holds majority ownership — raises governance independence questions",
            "ESG risk: E — data centre energy | S — youth mental health | G — foreign ownership opacity",
            "CSR claim: digital literacy, creator funds, transparency reports",
            "Reality: governance structure prevents full independence from Chinese law",
            "EU fined TikTok €5.4M (2023) for child data violations",
        ],
        "ethical": "CSR without governance independence is performative",
        "ref":    "EU Data Protection Board (2023); Freeman, R.E. (1984) Stakeholder Theory",
    },
    {
        "num":   6,
        "topic": "Ethics in Communication",
        "title": "Privacy, Power & the Algorithm",
        "icon":  "🔒",
        "bullets": [
            "Ethical risks: data harvesting, surveillance capitalism, targeting minors",
            "ACCA principles violated: Objectivity, Professional Behaviour, Confidentiality",
            "ByteDance engineers accessed US journalist location data (2022 — confirmed)",
            "Response: employees fired, incident labelled isolated",
            "Structural problem: platform incentivises engagement over user wellbeing",
        ],
        "ethical": "'We fired the individuals' is not a systemic ethical fix",
        "ref":    "ACCA Code of Ethics (2023); Forbes (2022) 'ByteDance Tracked Journalists'",
    },
    {
        "num":   7,
        "topic": "Crisis Communication",
        "title": "Ban Threat — 72 Hours to Respond",
        "icon":  "🚨",
        "bullets": [
            "Crisis: April 2024 — US House passed bill forcing divestiture or ban",
            "Response: CEO user video, in-app notifications, legal challenge filed",
            "What worked: direct creator mobilisation — users called Congress",
            "What failed: no clear timeline, no divestiture plan communicated",
            "MOCK STATEMENT | 'TikTok supports independent oversight and will pursue every legal avenue to continue serving 170M American users.'",
        ],
        "ethical": "Mobilising users as political shields raises questions about platform power",
        "ref":    "Coombs, W.T. (2015) Ongoing Crisis Communication, Sage; BBC News (2024)",
    },
    {
        "num":   8,
        "topic": "Digital Communication & Ethics",
        "title": "The Irony — A Digital Giant's Digital Failure",
        "icon":  "📱",
        "bullets": [
            "#SaveTikTok vs #BanTikTok — TikTok's own platform became the battleground",
            "AI algorithm risk: recommendation engine amplified ban panic",
            "Privacy paradox: app built on data collection fighting a battle about data collection",
            "Post-crisis move: transparency centre launched, open-source algorithm proposed",
        ],
        "ethical": "You cannot credibly champion digital ethics while your business model contradicts it",
        "ref":    "Wired (2024) 'TikTok's Fight for Survival'; Zuboff, S. (2019) Surveillance Capitalism",
    },
]

# ────────────────────────────────────────────────────────────────────────────
def build_cover(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # Full navy background
    bg = sl.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    fill(bg, NAVY); no_border(bg)

    # Thin gold left accent bar
    ab = sl.shapes.add_shape(1, 0, 0, Inches(0.07), SLIDE_H)
    fill(ab, GOLD); no_border(ab)

    # White content panel (right 55%)
    wp = sl.shapes.add_shape(1, Inches(6.0), 0, Inches(7.33), SLIDE_H)
    fill(wp, WHITE); no_border(wp)

    # Gold accent on panel left edge
    pa = sl.shapes.add_shape(1, Inches(6.0), 0, Inches(0.05), SLIDE_H)
    fill(pa, GOLD); no_border(pa)

    # ── Left: course tag
    _, tf = txbox(sl, Inches(0.5), Inches(1.6), Inches(5.0), Inches(0.45))
    solo(tf.paragraphs[0], "BCE222  |  FINAL ASSIGNMENT", 10, bold=True, color=GOLD)

    # ── Left: main title
    _, tf = txbox(sl, Inches(0.5), Inches(2.1), Inches(5.1), Inches(1.5))
    tf.word_wrap = True
    p = tf.paragraphs[0]
    solo(p, "Business Communication", 30, bold=True, color=WHITE)
    p2 = tf.add_paragraph()
    solo(p2, "Ethics & Strategy", 30, bold=True, color=WHITE)

    # ── Left: subtitle
    _, tf = txbox(sl, Inches(0.5), Inches(3.75), Inches(5.1), Inches(0.55))
    solo(tf.paragraphs[0], "TikTok US Ban — A Case Study", 13, color=GREY)

    # ── Left: gold divider
    d = sl.shapes.add_shape(1, Inches(0.5), Inches(4.42), Inches(4.5), Pt(1.5))
    fill(d, GOLD); no_border(d)

    # ── Right panel: topic list
    topics = [
        "01  Foundations of Business Communication",
        "02  Interpersonal & Cross-Cultural Communication",
        "03  Written Communication",
        "04  Persuasive & Strategic Communication",
        "05  Corporate Governance & CSR",
        "06  Ethics in Communication",
        "07  Crisis Communication",
        "08  Digital Communication & Ethics",
    ]
    _, tf = txbox(sl, Inches(6.4), Inches(1.4), Inches(6.5), Inches(5.0))
    tf.word_wrap = True
    first = True
    for t in topics:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(6)
        num, rest = t.split("  ", 1)
        run(p, num + "  ", 11, bold=True, color=GOLD)
        run(p, rest, 11, color=NAVY)

    # ── Footer strip
    fs = sl.shapes.add_shape(1, 0, SLIDE_H - Inches(0.32), SLIDE_W, Inches(0.32))
    fill(fs, RGBColor(0x0D, 0x16, 0x2B)); no_border(fs)
    _, tf = txbox(sl, Inches(0.5), SLIDE_H - Inches(0.30), Inches(10), Inches(0.26))
    solo(tf.paragraphs[0], "BCE222 Final Assignment  |  TikTok US Ban Case Study  |  Confidential", 7, color=GREY)


# ────────────────────────────────────────────────────────────────────────────
def build_slide(prs, data):
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # ── Full white background
    bg = sl.shapes.add_shape(1, 0, 0, SLIDE_W, SLIDE_H)
    fill(bg, WHITE); no_border(bg)

    # ── Navy header band
    hh = Inches(1.55)
    hb = sl.shapes.add_shape(1, 0, 0, SLIDE_W, hh)
    fill(hb, NAVY); no_border(hb)

    # ── Gold left accent bar (full height)
    ab = sl.shapes.add_shape(1, 0, 0, Inches(0.07), SLIDE_H)
    fill(ab, GOLD); no_border(ab)

    # ── Slide number pill (navy header, top-left)
    circ_sz = Inches(0.44)
    circ = sl.shapes.add_shape(9, Inches(0.22), Inches(0.55), circ_sz, circ_sz)
    fill(circ, GOLD); no_border(circ)
    circ.text_frame.paragraphs[0].text = str(data["num"])
    cp = circ.text_frame.paragraphs[0]
    cp.alignment = PP_ALIGN.CENTER
    cr = cp.runs[0]
    cr.font.name = "Calibri"; cr.font.size = Pt(13)
    cr.font.bold = True; cr.font.color.rgb = NAVY

    # ── Topic label (small caps style, header band)
    _, tf = txbox(sl, Inches(0.82), Inches(0.20), Inches(10.5), Inches(0.38))
    solo(tf.paragraphs[0], data["topic"].upper(), 8, bold=True, color=GOLD)

    # ── Slide title (header band)
    _, tf = txbox(sl, Inches(0.82), Inches(0.55), Inches(10.2), Inches(0.82))
    tf.word_wrap = True
    solo(tf.paragraphs[0], data["title"], 26, bold=True, color=WHITE)

    # ── Icon (top-right, inside header)
    _, tf = txbox(sl, Inches(11.6), Inches(0.30), Inches(1.5), Inches(0.9))
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    solo(p, data["icon"], 38, color=WHITE, align=PP_ALIGN.RIGHT)

    # ── Gold divider under header
    gd = sl.shapes.add_shape(1, Inches(0.82), hh, Inches(11.69), Pt(1.5))
    fill(gd, GOLD); no_border(gd)

    # ── Bullet section ────────────────────────────────────────────────────────
    bul_top  = hh + Inches(0.18)
    bul_area_h = Inches(3.0)

    _, tf = txbox(sl, Inches(0.82), bul_top, Inches(11.69), bul_area_h)
    tf.word_wrap = True

    first = True
    for b in data["bullets"]:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_before = Pt(5)
        p.space_after  = Pt(2)

        if "|" in b:
            # Block quote style (memo/statement excerpt)
            parts = [x.strip() for x in b.split("|")]
            label = parts[0]
            lines = parts[1:]

            run(p, "▸  ", 11, bold=True, color=GOLD)
            run(p, label, 11, bold=True, color=GOLD)
            for ln_text in lines:
                lp = tf.add_paragraph()
                lp.space_before = Pt(1)
                run(lp, "      " + ln_text, 9.5, italic=True,
                    color=RGBColor(0x4A, 0x5E, 0x78))
        else:
            run(p, "▸  ", 11, bold=True, color=GOLD)
            run(p, b, 11, color=DKGREY)

    # ── Gold thin divider before ethical angle
    eth_top = Inches(4.65)
    ed = sl.shapes.add_shape(1, Inches(0.82), eth_top, Inches(11.69), Pt(1.0))
    fill(ed, GOLD); no_border(ed)

    # ── Ethical angle line
    _, tf = txbox(sl, Inches(0.82), eth_top + Inches(0.10), Inches(11.69), Inches(0.52))
    tf.word_wrap = True
    ep = tf.paragraphs[0]
    run(ep, "Ethical Angle:  ", 11, bold=True, italic=True, color=GOLD)
    run(ep, data["ethical"], 11, italic=True,
        color=RGBColor(0x2C, 0x3E, 0x50))

    # ── Light grey divider before footnote
    fd = sl.shapes.add_shape(1, Inches(0.82), Inches(5.28), Inches(11.69), Pt(0.75))
    fill(fd, RGBColor(0xD0, 0xD6, 0xE0)); no_border(fd)

    # ── Reference footnote
    _, tf = txbox(sl, Inches(0.82), Inches(5.34), Inches(11.69), Inches(0.45))
    tf.word_wrap = True
    rp = tf.paragraphs[0]
    run(rp, "References:  ", 8, bold=True, color=GREY)
    run(rp, data["ref"], 8, italic=True, color=GREY)

    # ── Footer strip (navy, bottom)
    ft_top = SLIDE_H - Inches(0.32)
    ft = sl.shapes.add_shape(1, 0, ft_top, SLIDE_W, Inches(0.32))
    fill(ft, NAVY); no_border(ft)

    _, tf = txbox(sl, Inches(0.5), ft_top + Inches(0.04), Inches(10.5), Inches(0.25))
    solo(tf.paragraphs[0],
         "BCE222 Final Assignment  |  TikTok US Ban Case Study  |  Confidential",
         7, color=GREY)

    _, tf = txbox(sl, SLIDE_W - Inches(1.3), ft_top + Inches(0.04), Inches(1.1), Inches(0.25))
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    solo(p, f"{data['num']} / 8", 7, color=GREY, align=PP_ALIGN.RIGHT)


# ────────────────────────────────────────────────────────────────────────────
def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    build_cover(prs)
    for data in SLIDES:
        build_slide(prs, data)

    out = "BCE222_TikTok_Final.pptx"
    prs.save(out)
    print(f"Saved: {out}")

if __name__ == "__main__":
    main()
