from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
import copy

# ── Colour palette ──────────────────────────────────────────────────────────
NAVY   = RGBColor(0x1B, 0x2A, 0x4A)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
GOLD   = RGBColor(0xC9, 0xA0, 0x2C)   # accent line / ethical angle label
LIGHT  = RGBColor(0xE8, 0xED, 0xF5)   # subtle bullet bg tint (unused but kept)
RED    = RGBColor(0xC0, 0x39, 0x2B)   # ethical angle text

# ── Slide dimensions (widescreen 16:9) ──────────────────────────────────────
SLIDE_W = Inches(13.33)
SLIDE_H = Inches(7.5)

# ── Helper: set paragraph font ───────────────────────────────────────────────
def fmt(para, text, size, bold=False, color=WHITE, italic=False, align=PP_ALIGN.LEFT):
    para.text = ""
    run = para.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    para.alignment = align
    return run

def add_run(para, text, size, bold=False, color=WHITE, italic=False):
    run = para.add_run()
    run.text = text
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    return run

# ── Helper: solid-fill a shape ───────────────────────────────────────────────
def fill_shape(shape, color):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color

# ── Helper: remove shape border ──────────────────────────────────────────────
def no_border(shape):
    from pptx.oxml.ns import qn
    sp = shape._element
    spPr = sp.find(qn('p:spPr'))
    if spPr is None:
        return
    ln = spPr.find(qn('a:ln'))
    if ln is None:
        from lxml import etree
        ln = etree.SubElement(spPr, qn('a:ln'))
    from lxml import etree
    noFill = etree.SubElement(ln, qn('a:noFill'))

# ── Helper: add a text-box ───────────────────────────────────────────────────
def add_textbox(slide, left, top, width, height):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    return txBox, tf

# ── Per-slide data ────────────────────────────────────────────────────────────
slides_data = [
    {
        "title": "Communication at the Core — Volkswagen AG",
        "subtitle": "Foundations of Business Communication",
        "bullets": [
            "VW operates in 150+ countries requiring multi-channel communication strategies",
            "Channels used: internal memos, press releases, shareholder reports, digital media",
            "Key barrier identified: information silos between engineering and management — emissions data was known internally but never escalated",
            "Communication breakdown directly enabled the Dieselgate scandal (2015)",
            "Lesson: transparent vertical communication is not optional, it is a governance requirement",
        ],
        "ethical": "Deliberate withholding of information = violation of integrity and transparency principles",
        "ref": "Volkswagen AG Annual Report (2015); Argenti, P. (2015) Corporate Communication, McGraw-Hill",
    },
    {
        "title": "Global Stakeholders, Local Failures",
        "subtitle": "Interpersonal & Cross-Cultural Communication",
        "bullets": [
            "Stakeholders: US regulators (EPA), German government, 11 million car owners globally, VW employees",
            "Cultural clash: German engineering culture prioritised performance over compliance with US environmental standards",
            "High-context (Germany) vs low-context (USA) communication styles caused misalignment",
            "VW's response messaging was adapted differently per market — inconsistent and damaging",
            "Cross-cultural competence was absent at executive level",
        ],
        "ethical": "Adapting message per audience to minimise accountability = manipulative communication",
        "ref": "Hofstede, G. (2001) Culture's Consequences, Sage; BBC News (2015) VW Emissions Scandal",
    },
    {
        "title": "The Memo That Should Have Been Sent",
        "subtitle": "Written Communication",
        "bullets": [
            "Example: Internal memo from compliance officer to CEO (fictional but realistic)",
            "Tone: formal, factual, urgent — following ACCA professional writing standards",
            "Structure: Issue → Evidence → Risk → Recommended Action",
            "ACCA standard: clear subject line, no ambiguity, action-oriented closing",
            "MOCK MEMO EXCERPT | To: CEO, Volkswagen AG | From: Compliance Department | Subject: Emissions Software — Regulatory Risk | The defeat device in 11M vehicles violates EPA Clean Air Act standards. Immediate disclosure is recommended to avoid criminal liability.",
            "This memo was never sent — absence of written escalation enabled the cover-up",
        ],
        "ethical": "This memo was never sent — absence of written escalation enabled the cover-up",
        "ref": "ACCA (2023) Professional Ethics Module; Guffey, M. (2016) Business Communication, Cengage",
    },
    {
        "title": "Selling Clean Diesel — A False Narrative",
        "subtitle": "Persuasive & Strategic Communication",
        "bullets": [
            "VW's 'Think Blue' campaign persuaded consumers and regulators that diesel was eco-friendly",
            "Ethos: leveraged VW's 60-year brand reputation for credibility",
            "Pathos: environmental messaging appealed directly to eco-conscious buyers",
            "Logos: manipulated emissions test data was presented as factual scientific evidence",
            "Strategic goal: dominate US market by positioning diesel as a sustainable alternative",
        ],
        "ethical": "All three persuasion pillars (ethos, pathos, logos) were weaponised to deceive — textbook unethical persuasion",
        "ref": "Aristotle (350 BC) Rhetoric; New York Times (2015) 'VW's Clean Diesel Campaign'",
    },
    {
        "title": "Governance Failure at Every Level",
        "subtitle": "Corporate Governance & CSR",
        "bullets": [
            "VW's supervisory board failed to challenge executive decisions on emissions targets",
            "ESG failure: Environmental (fake emissions), Social (consumer deception), Governance (absent board oversight)",
            "CSR contradiction: VW published sustainability reports while knowingly violating emissions law",
            "$33 billion paid in fines, settlements, and recalls globally",
            "Post-scandal: VW shifted to EV strategy (ID.4, ID.3) as genuine CSR pivot",
        ],
        "ethical": "CSR as marketing vs CSR as genuine commitment — VW was the former until forced to change",
        "ref": "VW Sustainability Report (2016); Freeman, R.E. (1984) Strategic Management: A Stakeholder Approach",
    },
    {
        "title": "Lies, Omissions and Accountability",
        "subtitle": "Ethics in Communication",
        "bullets": [
            "Ethical risks present: misleading advertising, data manipulation, confidentiality breaches, reporting bias",
            "ACCA ethical principles violated: Integrity, Objectivity, Professional Behaviour",
            "Defeat device = deliberate misrepresentation to regulators — a premeditated ethical breach",
            "Employees who knew stayed silent — ethical pressure and fear of retaliation",
            "Whistleblower protection was absent in VW's corporate culture",
        ],
        "ethical": "Organisational silence is an active ethical failure — not speaking up = complicity",
        "ref": "ACCA (2023) Code of Ethics and Conduct; Transparency International (2016) VW Case Study",
    },
    {
        "title": "Too Late, Too Little",
        "subtitle": "Crisis Communication",
        "bullets": [
            "Crisis: September 2015 — EPA publicly announced VW emissions violation",
            "VW's response: denial → partial admission → CEO resignation (Martin Winterkorn)",
            "Failures: delayed response, inconsistent messaging, no single designated spokesperson",
            "Correct framework (not used): Acknowledge → Accept Responsibility → Action Plan → Apology",
            "WHAT THEY SHOULD HAVE SAID | 'Volkswagen acknowledges a serious compliance failure affecting emissions software. We take full responsibility, are cooperating with all regulators, and commit to full transparency throughout this process.'",
        ],
        "ethical": "Crisis communication is a moral act — delayed truth causes greater harm to all stakeholders",
        "ref": "Coombs, W.T. (2015) Ongoing Crisis Communication, Sage; Reuters (2015) VW Crisis Timeline",
    },
    {
        "title": "Scandal in the Age of Social Media",
        "subtitle": "Digital Communication & Ethics",
        "bullets": [
            "#VWScandal trended globally within hours of the EPA announcement",
            "VW's social media team had no crisis protocol — accounts went silent for days",
            "AI & digital risk: defeat device itself was software-based deception — a technology ethics violation",
            "Misinformation spread rapidly; VW lost complete control of the narrative",
            "Post-crisis: VW invested in digital transparency tools and live emissions tracking dashboards",
        ],
        "ethical": "Digital silence during a crisis = abandonment of stakeholder communication duty",
        "ref": "Kaplan, A. & Haenlein, M. (2010) Users of the world, unite! Business Horizons; Wired (2015) 'How VW's Cheating Software Worked'",
    },
]

# ────────────────────────────────────────────────────────────────────────────
def build_slide(prs, data, slide_num):
    blank_layout = prs.slide_layouts[6]   # completely blank
    slide = prs.slides.add_slide(blank_layout)

    W = SLIDE_W
    H = SLIDE_H

    # ── Full background ──────────────────────────────────────────────────────
    bg = slide.shapes.add_shape(
        1,  # MSO_SHAPE_TYPE.RECTANGLE
        0, 0, W, H
    )
    fill_shape(bg, NAVY)
    no_border(bg)
    bg.zorder = 0

    # ── Gold accent bar (left edge, full height) ────────────────────────────
    accent = slide.shapes.add_shape(1, 0, 0, Inches(0.07), H)
    fill_shape(accent, GOLD)
    no_border(accent)

    # ── Top header band ──────────────────────────────────────────────────────
    header_h = Inches(1.4)
    header = slide.shapes.add_shape(1, 0, 0, W, header_h)
    from pptx.dml.color import RGBColor as RGB
    header.fill.solid()
    header.fill.fore_color.rgb = RGB(0x12, 0x1E, 0x36)
    no_border(header)

    # ── Slide number circle ──────────────────────────────────────────────────
    circle_size = Inches(0.48)
    circ = slide.shapes.add_shape(
        9,  # OVAL
        Inches(0.18), Inches(0.46),
        circle_size, circle_size
    )
    fill_shape(circ, GOLD)
    no_border(circ)
    circ.text_frame.text = str(slide_num)
    circ_para = circ.text_frame.paragraphs[0]
    circ_para.alignment = PP_ALIGN.CENTER
    circ_run = circ_para.runs[0]
    circ_run.font.name = "Calibri"
    circ_run.font.size = Pt(14)
    circ_run.font.bold = True
    circ_run.font.color.rgb = NAVY

    # ── Subtitle (topic name) ────────────────────────────────────────────────
    sub_box, sub_tf = add_textbox(
        slide,
        Inches(0.82), Inches(0.18),
        Inches(11.8), Inches(0.38)
    )
    sub_p = sub_tf.paragraphs[0]
    fmt(sub_p, data["subtitle"].upper(), 9, bold=True, color=GOLD)

    # ── Main title ───────────────────────────────────────────────────────────
    title_box, title_tf = add_textbox(
        slide,
        Inches(0.82), Inches(0.52),
        Inches(11.8), Inches(0.75)
    )
    t_p = title_tf.paragraphs[0]
    fmt(t_p, data["title"], 24, bold=True, color=WHITE)

    # ── Gold divider line ────────────────────────────────────────────────────
    from pptx.oxml.ns import qn
    from lxml import etree

    line = slide.shapes.add_shape(
        1, Inches(0.82), Inches(1.32), Inches(11.7), Pt(1.5)
    )
    fill_shape(line, GOLD)
    no_border(line)

    # ── Bullet points ────────────────────────────────────────────────────────
    # Decide how many bullets vs block bullets
    normal_bullets = []
    block_bullet = None
    for b in data["bullets"]:
        if "|" in b:
            block_bullet = b
        else:
            normal_bullets.append(b)

    bullet_top = Inches(1.48)
    bullet_h   = Inches(3.05)

    bul_box, bul_tf = add_textbox(
        slide,
        Inches(0.82), bullet_top,
        Inches(11.7), bullet_h
    )
    bul_tf.word_wrap = True

    first = True
    for b in normal_bullets:
        if first:
            p = bul_tf.paragraphs[0]
            first = False
        else:
            p = bul_tf.add_paragraph()
        p.space_before = Pt(3)
        p.space_after  = Pt(3)
        add_run(p, "▸  ", 11, bold=True, color=GOLD)
        add_run(p, b, 11, color=WHITE)

    if block_bullet:
        p = bul_tf.add_paragraph() if not first else bul_tf.paragraphs[0]
        first = False
        p.space_before = Pt(6)
        parts = [x.strip() for x in block_bullet.split("|")]
        label = parts[0]
        rest  = parts[1:]

        # Label line
        add_run(p, "▸  ", 11, bold=True, color=GOLD)
        add_run(p, label, 11, bold=True, color=GOLD)

        for line_text in rest:
            lp = bul_tf.add_paragraph()
            lp.space_before = Pt(1)
            add_run(lp, "      " + line_text, 10, color=RGBColor(0xB0, 0xC4, 0xDE), italic=True)

    # ── Ethical Angle box ────────────────────────────────────────────────────
    eth_top = Inches(4.62)
    eth_h   = Inches(0.72)

    eth_bg = slide.shapes.add_shape(
        1, Inches(0.82), eth_top, Inches(11.7), eth_h
    )
    eth_bg.fill.solid()
    eth_bg.fill.fore_color.rgb = RGBColor(0x0D, 0x16, 0x2B)
    no_border(eth_bg)

    # left accent stripe
    eth_stripe = slide.shapes.add_shape(
        1, Inches(0.82), eth_top, Inches(0.06), eth_h
    )
    fill_shape(eth_stripe, RED)
    no_border(eth_stripe)

    eth_box, eth_tf = add_textbox(
        slide,
        Inches(1.02), eth_top + Inches(0.06),
        Inches(11.4), eth_h - Inches(0.12)
    )
    eth_tf.word_wrap = True
    ep = eth_tf.paragraphs[0]
    add_run(ep, "ETHICAL ANGLE:  ", 10, bold=True, color=RED)
    add_run(ep, data["ethical"], 10, color=WHITE, italic=True)

    # ── Reference footnote ───────────────────────────────────────────────────
    ref_top = Inches(5.42)

    ref_divider = slide.shapes.add_shape(
        1, Inches(0.82), ref_top, Inches(11.7), Pt(0.75)
    )
    fill_shape(ref_divider, RGBColor(0x3A, 0x4E, 0x70))
    no_border(ref_divider)

    ref_box, ref_tf = add_textbox(
        slide,
        Inches(0.82), ref_top + Inches(0.06),
        Inches(11.7), Inches(0.52)
    )
    ref_tf.word_wrap = True
    rp = ref_tf.paragraphs[0]
    add_run(rp, "References:  ", 8, bold=True, color=GOLD)
    add_run(rp, data["ref"], 8, color=RGBColor(0xB0, 0xC4, 0xDE), italic=True)

    # ── Bottom footer strip ──────────────────────────────────────────────────
    footer_top = H - Inches(0.35)
    footer = slide.shapes.add_shape(1, 0, footer_top, W, Inches(0.35))
    footer.fill.solid()
    footer.fill.fore_color.rgb = RGBColor(0x0D, 0x16, 0x2B)
    no_border(footer)

    footer_box, footer_tf = add_textbox(
        slide,
        Inches(0.82), footer_top + Inches(0.04),
        Inches(10), Inches(0.28)
    )
    fp = footer_tf.paragraphs[0]
    add_run(fp, "BCE222 Final Assignment  |  Volkswagen Dieselgate Case Study  |  Confidential", 7,
            color=RGBColor(0x6A, 0x80, 0xA8))

    # slide number (right)
    snum_box, snum_tf = add_textbox(
        slide,
        W - Inches(1.2), footer_top + Inches(0.04),
        Inches(1.0), Inches(0.28)
    )
    snp = snum_tf.paragraphs[0]
    snp.alignment = PP_ALIGN.RIGHT
    add_run(snp, f"{slide_num} / 8", 7, color=RGBColor(0x6A, 0x80, 0xA8))


# ── Title / Cover slide ───────────────────────────────────────────────────────
def build_cover(prs):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    W, H = SLIDE_W, SLIDE_H

    # Background
    bg = slide.shapes.add_shape(1, 0, 0, W, H)
    fill_shape(bg, NAVY)
    no_border(bg)

    # Dark panel right side
    panel = slide.shapes.add_shape(1, Inches(7.8), 0, Inches(5.53), H)
    panel.fill.solid()
    panel.fill.fore_color.rgb = RGBColor(0x0D, 0x16, 0x2B)
    no_border(panel)

    # Gold accent bar
    accent = slide.shapes.add_shape(1, 0, 0, Inches(0.07), H)
    fill_shape(accent, GOLD)
    no_border(accent)

    # Horizontal gold line
    hline = slide.shapes.add_shape(1, Inches(0.6), Inches(3.1), Inches(6.8), Pt(2))
    fill_shape(hline, GOLD)
    no_border(hline)

    # Course code
    cc_box, cc_tf = add_textbox(slide, Inches(0.6), Inches(1.5), Inches(7), Inches(0.5))
    cc_p = cc_tf.paragraphs[0]
    fmt(cc_p, "BCE222  |  FINAL ASSIGNMENT", 13, bold=True, color=GOLD)

    # Main title
    t_box, t_tf = add_textbox(slide, Inches(0.6), Inches(1.95), Inches(7.0), Inches(1.1))
    t_tf.word_wrap = True
    tp = t_tf.paragraphs[0]
    fmt(tp, "Business Communication", 34, bold=True, color=WHITE)
    tp2 = t_tf.add_paragraph()
    fmt(tp2, "Ethics & Strategy", 34, bold=True, color=WHITE)

    # Subtitle
    s_box, s_tf = add_textbox(slide, Inches(0.6), Inches(3.22), Inches(7.0), Inches(0.6))
    s_tf.word_wrap = True
    sp = s_tf.paragraphs[0]
    fmt(sp, "A Case Study on the Volkswagen Dieselgate Scandal", 14, color=RGBColor(0xB0, 0xC4, 0xDE))

    # Right panel content
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
    tp_box, tp_tf = add_textbox(slide, Inches(8.1), Inches(1.2), Inches(4.8), Inches(5.0))
    tp_tf.word_wrap = True
    first = True
    for t in topics:
        p = tp_tf.paragraphs[0] if first else tp_tf.add_paragraph()
        first = False
        p.space_before = Pt(5)
        num, rest = t.split("  ", 1)
        add_run(p, num + "  ", 10, bold=True, color=GOLD)
        add_run(p, rest, 10, color=WHITE)

    # Bottom footer
    footer_top = H - Inches(0.35)
    footer = slide.shapes.add_shape(1, 0, footer_top, W, Inches(0.35))
    footer.fill.solid()
    footer.fill.fore_color.rgb = RGBColor(0x0D, 0x16, 0x2B)
    no_border(footer)
    fb, ftf = add_textbox(slide, Inches(0.6), footer_top + Inches(0.04), Inches(12), Inches(0.28))
    fp = ftf.paragraphs[0]
    add_run(fp, "BCE222 Final Assignment  |  Volkswagen Dieselgate Case Study  |  Confidential", 7,
            color=RGBColor(0x6A, 0x80, 0xA8))


# ── Assemble presentation ─────────────────────────────────────────────────────
def main():
    prs = Presentation()
    prs.slide_width  = SLIDE_W
    prs.slide_height = SLIDE_H

    build_cover(prs)

    for i, data in enumerate(slides_data, start=1):
        build_slide(prs, data, i)

    out = "BCE222_Final_VW.pptx"
    prs.save(out)
    print(f"Saved: {out}")

if __name__ == "__main__":
    main()
