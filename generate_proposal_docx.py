#!/usr/bin/env python3
"""Generate StacksPot Growth Acceleration Proposal as a Word document."""

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor, Twips


NAVY = RGBColor(0x1A, 0x36, 0x5D)
ACCENT = RGBColor(0x2B, 0x6C, 0xB0)
DARK = RGBColor(0x1F, 0x29, 0x37)
GRAY = RGBColor(0x4A, 0x55, 0x68)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_FILL = "EEF2F7"
HEADER_FILL = "1A365D"
ALT_ROW = "F7FAFC"


def set_run_font(run, name="Calibri", size=11, bold=False, italic=False, color=DARK):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    run.font.color.rgb = color


def add_text(paragraph, text, **kwargs):
    run = paragraph.add_run(text)
    set_run_font(run, **kwargs)
    return run


def set_paragraph_spacing(paragraph, before=0, after=8, line=1.15, space_after=None):
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(before)
    fmt.space_after = Pt(after if space_after is None else space_after)
    fmt.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    fmt.line_spacing = line


def shade_cell(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), hex_color)
    shd.set(qn("w:val"), "clear")
    tc_pr.append(shd)


def set_cell_borders(cell, color="CBD5E1", size="4"):
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_borders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), size)
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), color)
        tc_borders.append(border)
    tc_pr.append(tc_borders)


def set_cell_text(cell, text, bold=False, size=10, color=DARK, align=WD_ALIGN_PARAGRAPH.LEFT, font="Calibri"):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    set_paragraph_spacing(p, before=2, after=2, line=1.0)
    add_text(p, text, size=size, bold=bold, color=color, name=font)
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_before = Pt(2)
        paragraph.paragraph_format.space_after = Pt(2)


def add_heading_styled(doc, text, level=1):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=18 if level == 1 else 12, after=6, line=1.15)
    if level == 1:
        add_text(p, text, name="Calibri", size=16, bold=True, color=NAVY)
    else:
        add_text(p, text, name="Calibri", size=13, bold=True, color=ACCENT)
    return p


def add_body(doc, text, after=8):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=0, after=after, line=1.15)
    add_text(p, text, size=11, color=DARK)
    return p


def add_rich_body(doc, parts, after=8):
    """parts: list of (text, kwargs)"""
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=0, after=after, line=1.15)
    for text, kwargs in parts:
        add_text(p, text, **kwargs)
    return p


def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style="List Bullet")
    set_paragraph_spacing(p, before=0, after=3, line=1.1)
    if bold_prefix:
        add_text(p, bold_prefix, size=11, bold=True, color=DARK)
        add_text(p, text, size=11, color=DARK)
    else:
        # Clear default run if any
        if p.runs:
            p.runs[0].text = ""
        add_text(p, text, size=11, color=DARK)
    return p


def add_numbered(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style="List Number")
    set_paragraph_spacing(p, before=0, after=3, line=1.1)
    if bold_prefix:
        add_text(p, bold_prefix, size=11, bold=True, color=DARK)
        add_text(p, text, size=11, color=DARK)
    else:
        if p.runs:
            p.runs[0].text = ""
        add_text(p, text, size=11, color=DARK)
    return p


def make_table(doc, headers, rows, col_widths=None, header=True):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_text(cell, h, bold=True, size=10, color=WHITE, align=WD_ALIGN_PARAGRAPH.CENTER)
        shade_cell(cell, HEADER_FILL)
        set_cell_borders(cell, color="1A365D", size="4")

    for r_idx, row in enumerate(rows):
        for c_idx, val in enumerate(row):
            cell = table.rows[r_idx + 1].cells[c_idx]
            is_total = r_idx == len(rows) - 1 and any(
                str(row[0]).lower().startswith(("8", "program")) or "total" in str(x).lower() for x in row
            )
            # Detect totals more simply
            bold = False
            if isinstance(val, str) and ("300,000" in val or val == "8" or "Program total" in val or "8 sponsored" in val):
                bold = True
            # Bold last total row
            if r_idx == len(rows) - 1 and ("300,000 STX" in str(row[-1]) or "Program total" in str(row)):
                bold = True
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT
            if c_idx >= len(headers) - 3:
                align = WD_ALIGN_PARAGRAPH.RIGHT if c_idx > 0 else align
            # For descriptive tables keep left
            if len(headers) == 2:
                align = WD_ALIGN_PARAGRAPH.LEFT
            set_cell_text(cell, str(val), bold=bold, size=10, color=DARK, align=align)
            set_cell_borders(cell)
            if r_idx % 2 == 1:
                shade_cell(cell, ALT_ROW)
            if r_idx == len(rows) - 1 and ("300,000 STX" in str(row[-1]) or "Program total" in str(row) or (row[0] == "8")):
                shade_cell(cell, LIGHT_FILL)
                set_cell_text(cell, str(val), bold=True, size=10, color=NAVY, align=align)

    if col_widths:
        for row in table.rows:
            for idx, width in enumerate(col_widths):
                row.cells[idx].width = Inches(width)

    # Spacer after table
    spacer = doc.add_paragraph()
    set_paragraph_spacing(spacer, before=0, after=6, line=1.0)
    return table


def add_horizontal_line(doc):
    p = doc.add_paragraph()
    set_paragraph_spacing(p, before=4, after=4, line=1.0)
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "2B6CB0")
    pBdr.append(bottom)
    pPr.append(pBdr)
    return p


def build():
    doc = Document()

    # Page setup
    section = doc.sections[0]
    section.top_margin = Inches(0.85)
    section.bottom_margin = Inches(0.85)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)

    # Default style
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    style.font.color.rgb = DARK

    # ---- Cover / Title ----
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(title, before=24, after=4, line=1.1)
    add_text(title, "StacksPot Growth Acceleration Proposal", name="Calibri", size=26, bold=True, color=NAVY)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    set_paragraph_spacing(subtitle, before=0, after=14, line=1.1)
    add_text(
        subtitle,
        "A Sponsored Yield Boosting Initiative for Ecosystem Growth",
        name="Calibri",
        size=13,
        italic=True,
        color=ACCENT,
    )

    add_horizontal_line(doc)

    # Meta table
    meta = [
        ("Submitted By", "StacksPot Team"),
        ("Submitted To", "Stacks Endowment Team"),
        ("Program Duration", "5 months"),
        ("Requested Allocation", "~300,000 STX"),
        ("Date", "July 2026"),
    ]
    meta_table = doc.add_table(rows=len(meta), cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (label, value) in enumerate(meta):
        c0, c1 = meta_table.rows[i].cells
        set_cell_text(c0, label, bold=True, size=11, color=WHITE)
        shade_cell(c0, HEADER_FILL)
        set_cell_borders(c0, color="1A365D")
        set_cell_text(c1, value, bold=False, size=11, color=DARK)
        set_cell_borders(c1)
        shade_cell(c1, ALT_ROW if i % 2 else "FFFFFF")
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)

    spacer = doc.add_paragraph()
    set_paragraph_spacing(spacer, before=6, after=6)

    # ---- Proposal Objective ----
    add_heading_styled(doc, "Proposal Objective", level=1)
    add_rich_body(
        doc,
        [
            ("Accelerate user adoption through sponsored sBTC reward campaigns, boosted community pots, and ecosystem participation incentives over a structured ", {"size": 11}),
            ("5-month", {"size": 11, "bold": True}),
            (" Growth Accelerator Program.", {"size": 11}),
        ],
    )

    # ---- 1. Executive Summary ----
    add_heading_styled(doc, "1. Executive Summary", level=1)
    add_body(
        doc,
        "StacksPot is a decentralized, yield-powered community pot protocol built on Stacks. It enables users to participate in transparent, sBTC-based reward pools while benefiting from Bitcoin-secured DeFi yield.",
    )
    add_body(
        doc,
        "Since launching, StacksPot has demonstrated the potential for community-driven reward mechanisms within the Stacks ecosystem. The primary challenge limiting broader adoption is the perceived attractiveness of current pot rewards.",
    )
    add_body(
        doc,
        "While StacksPot offers a unique model—yield-generating pools with principal preservation—many users evaluate participation based on expected reward versus waiting period. At current yield levels, some users do not find the opportunity compelling enough to join.",
    )
    add_rich_body(
        doc,
        [
            ("To address this, StacksPot proposes a strategic partnership with the Stacks Endowment through a ", {"size": 11}),
            ("Sponsored Pot Growth Accelerator Program", {"size": 11, "bold": True}),
            ("—a focused, ", {"size": 11}),
            ("5-month", {"size": 11, "bold": True}),
            (" campaign cycle designed to boost community pots with additional sBTC rewards, strengthen participation incentives, encourage more pot deployments, and increase overall ecosystem activity.", {"size": 11}),
        ],
    )
    add_rich_body(
        doc,
        [
            ("We are seeking consideration for an ecosystem growth allocation of approximately ", {"size": 11}),
            ("300,000 STX", {"size": 11, "bold": True}),
            (" to fund sponsored pot campaigns across this 5-month period, with the goal of accelerating adoption and building sustainable user growth.", {"size": 11}),
        ],
    )

    # ---- 2. Current Challenge ----
    add_heading_styled(doc, "2. Current Challenge", level=1)
    add_heading_styled(doc, "The User Adoption Barrier", level=2)
    add_rich_body(
        doc,
        [
            ("StacksPot’s growth constraint is not technology or product experience—it is ", {"size": 11}),
            ("reward perception", {"size": 11, "bold": True}),
            (" among potential participants.", {"size": 11}),
        ],
    )
    add_body(doc, "Users naturally weigh:", after=4)
    for item in [
        "The amount of STX committed",
        "The duration of participation",
        "The expected reward outcome",
        "Alternative opportunities available",
    ]:
        add_bullet(doc, item)

    add_body(
        doc,
        "Although StacksPot provides a transparent and innovative yield model, early-stage adoption requires stronger incentives so users can experience the platform and understand its long-term value.",
    )
    add_rich_body(
        doc,
        [
            ("To establish network effects, StacksPot needs an initial growth catalyst that increases reward attractiveness while introducing more users to the protocol. A ", {"size": 11}),
            ("time-bound, 5-month sponsored program", {"size": 11, "bold": True}),
            (" provides that catalyst with clear milestones, measurable outcomes, and a defined end state for evaluation.", {"size": 11}),
        ],
    )

    # ---- 3. Proposed Solution ----
    add_heading_styled(doc, "3. Proposed Solution", level=1)
    add_heading_styled(doc, "StacksPot Growth Accelerator Program", level=2)
    add_rich_body(
        doc,
        [
            ("With StacksPot’s updated smart contracts, ", {"size": 11}),
            ("pot sponsorship is now enabled", {"size": 11, "bold": True}),
            (". Ecosystem partners, communities, projects, and organizations can contribute additional STX directly to active pots.", {"size": 11}),
        ],
    )
    add_body(
        doc,
        "Through Stacks Endowment sponsorship, selected pots can receive reward boosts—making campaigns more attractive and transforming StacksPot from a passive yield product into an active ecosystem engagement tool.",
    )

    add_heading_styled(doc, "Illustrative Example", level=2)
    make_table(
        doc,
        ["Component", "Amount"],
        [
            ("Organic community pooling", "500 STX"),
            ("Stacks Endowment sponsorship boost", "+5,000 STX"),
            ("Final reward pool (delegated)", "5,500 STX"),
        ],
        col_widths=[4.5, 2.0],
    )

    # ---- 4. Program Duration ----
    add_heading_styled(doc, "4. Program Duration & Operating Context", level=1)
    add_heading_styled(doc, "5-Month Campaign Window", level=2)
    add_rich_body(
        doc,
        [
            ("This proposal requests funding for a ", {"size": 11}),
            ("5-month Growth Accelerator Program", {"size": 11, "bold": True}),
            (", structured to:", {"size": 11}),
        ],
        after=4,
    )
    add_numbered(doc, " with smaller, accessible pots early in the program", bold_prefix="Ramp adoption")
    add_numbered(doc, " with mid-tier campaigns through the middle months", bold_prefix="Sustain momentum")
    add_numbered(doc, " with larger flagship pots toward the end of the cycle", bold_prefix="Concentrate demand")
    add_numbered(doc, " across a long enough window to distinguish campaign spikes from lasting engagement", bold_prefix="Measure retention")

    add_body(doc, "A 5-month runway is long enough to:", after=4)
    for item in [
        "Run multiple campaign tiers without overcrowding the market",
        "Allow communities and builders to plan and deploy their own pots",
        "Collect meaningful retention and repeat-participation data",
        "Provide the Endowment with clear monthly visibility into impact",
    ]:
        add_bullet(doc, item)

    add_body(
        doc,
        "At the end of Month 5, StacksPot will deliver a full program retrospective covering participation, STX utilization, pot creation, sponsored reward distribution, and retention—supporting a data-driven decision on any future continuation.",
    )

    # ---- 5. Campaign Structure ----
    add_heading_styled(doc, "5. Proposed Campaign Structure", level=1)
    add_heading_styled(doc, "Campaign Name", level=2)
    add_rich_body(
        doc,
        [
            ("StacksPot Growth Accelerator", {"size": 11, "bold": True}),
        ],
        after=4,
    )
    add_body(
        doc,
        "A funded, 5-month initiative to increase participation, reward users, and encourage individual and community-driven pot creation.",
    )

    add_heading_styled(doc, "Proposed Allocation", level=2)
    add_rich_body(
        doc,
        [
            ("The ", {"size": 11}),
            ("~300,000 STX", {"size": 11, "bold": True}),
            (" sponsorship budget will be distributed across StacksPot campaigns of varying entry sizes to maximize participation, accelerate pot growth, and generate sustainable sBTC rewards.", {"size": 11}),
        ],
    )

    make_table(
        doc,
        ["Qty (×)", "Entry Min.", "Min. Participants", "Pot Target", "Awarded Boost", "Est. Rewards"],
        [
            ("1", "25 STX", "10", "250 STX", "10,000 STX", "~$4.92 sBTC"),
            ("1", "30 STX", "15", "450 STX", "15,000 STX", "~$7.41 sBTC"),
            ("1", "35 STX", "20", "700 STX", "20,000 STX", "~$9.93 sBTC"),
            ("1", "40 STX", "25", "1,000 STX", "25,000 STX", "~$12.50 sBTC"),
            ("1", "45 STX", "30", "1,350 STX", "30,000 STX", "~$15.00 sBTC"),
            ("2", "50 STX", "50", "2,500 STX", "50,000 STX", "~$25.20 sBTC"),
            ("1", "100 STX", "100", "10,000 STX", "100,000 STX", "~$52.80 sBTC"),
            ("8", "—", "—", "—", "300,000 STX", "—"),
        ],
        col_widths=[0.7, 1.0, 1.3, 1.1, 1.2, 1.2],
    )

    note = doc.add_paragraph()
    set_paragraph_spacing(note, before=0, after=10, line=1.1)
    add_text(note, "Note: ", size=10, italic=True, bold=True, color=GRAY)
    add_text(
        note,
        "Estimated sBTC reward figures are approximate and subject to market conditions and yield performance at campaign time.",
        size=10,
        italic=True,
        color=GRAY,
    )

    add_heading_styled(doc, "Indicative 5-Month Rollout", level=2)
    add_body(
        doc,
        "Campaigns will be sequenced to build familiarity, then scale commitment:",
        after=6,
    )
    make_table(
        doc,
        ["Month", "Focus", "Campaign Tiers (Entry Min.)", "Indicative Boost"],
        [
            ("1", "Launch & discovery — lower barriers, introduce sponsorship", "25 / 30 / 35 STX", "45,000 STX"),
            ("2", "Expand mid-tier participation", "40 / 45 STX", "55,000 STX"),
            ("3", "Deepen community pots", "50 STX (1 of 2)", "50,000 STX"),
            ("4", "Sustain engagement & retention", "50 STX (2 of 2)", "50,000 STX"),
            ("5", "Flagship campaign & program close", "100 STX", "100,000 STX"),
            ("", "Program total", "8 sponsored pots", "300,000 STX"),
        ],
        col_widths=[0.7, 3.0, 1.8, 1.3],
    )
    add_body(
        doc,
        "Timing may be adjusted slightly based on fill rates, community readiness, and observed demand, while remaining within the 5-month window and total allocation.",
    )

    # ---- 6. Benefits ----
    add_heading_styled(doc, "6. Why This Benefits the Stacks Ecosystem", level=1)

    add_heading_styled(doc, "Increasing User Adoption", level=2)
    add_body(
        doc,
        "Sponsored rewards lower the initial barrier for users and encourage more participants to interact with Stacks DeFi.",
    )

    add_heading_styled(doc, "Growing STX Activity", level=2)
    add_body(doc, "More participants means:", after=4)
    for item in [
        "Increased STX deposits",
        "More wallet activity",
        "Greater ecosystem engagement",
    ]:
        add_bullet(doc, item)

    add_heading_styled(doc, "Supporting Builders and Communities", level=2)
    add_body(
        doc,
        "StacksPot becomes infrastructure that lets communities create their own incentive campaigns—not only during the Endowment-sponsored window, but as an ongoing pattern afterward.",
    )

    add_heading_styled(doc, "Expanding Bitcoin DeFi Awareness", level=2)
    add_body(
        doc,
        "Each sponsored pot introduces new users to Bitcoin-secured yield opportunities available through Stacks.",
    )

    # ---- 7. Expected Outcomes ----
    add_heading_styled(doc, "7. Expected Outcomes", level=1)
    add_body(
        doc,
        "With Stacks Endowment support over the 5-month program, StacksPot expects to achieve:",
        after=6,
    )
    make_table(
        doc,
        ["Outcome", "Description"],
        [
            ("User growth", "Increase active participants through attractive reward campaigns"),
            ("More pot deployments", "Encourage community members and projects to create their own reward pools"),
            ("Higher ecosystem engagement", "Increase STX utilization and wallet interactions"),
            (
                "Long-term adoption",
                "Convert campaign participants into recurring StacksPot users, measurable across the full 5-month window",
            ),
        ],
        col_widths=[2.2, 4.3],
    )

    # ---- 8. Transparency ----
    add_heading_styled(doc, "8. Transparency & Reporting", level=1)
    add_body(
        doc,
        "All sponsored campaigns will include transparent reporting. Given the 5-month structure, reporting will operate on two cadences:",
    )

    add_heading_styled(doc, "Monthly Progress Reports", level=2)
    add_body(doc, "Delivered at the end of each program month, covering:", after=4)
    for item in [
        "Number of participants",
        "Total STX deposited",
        "Number of pots created",
        "Sponsored reward distribution for that month",
        "Early retention / repeat participation signals",
    ]:
        add_bullet(doc, item)

    add_heading_styled(doc, "Final Program Report (End of Month 5)", level=2)
    add_body(doc, "A comprehensive retrospective including:", after=4)
    for item in [
        "Full campaign performance analysis",
        "User retention statistics across the 5-month period",
        "Cumulative STX utilization and wallet activity",
        "Lessons learned and recommendations for any future ecosystem incentive programs",
    ]:
        add_bullet(doc, item)

    add_body(
        doc,
        "Stacks Endowment will have clear visibility into the impact generated by the sponsorship throughout and at the close of the program.",
    )

    # ---- 9. Conclusion ----
    add_heading_styled(doc, "9. Conclusion", level=1)
    add_body(
        doc,
        "StacksPot has established the foundation for community-driven yield campaigns on Stacks. The next phase is accelerating adoption by creating stronger participation incentives.",
    )
    add_rich_body(
        doc,
        [
            ("Through collaboration with the Stacks Endowment, a ", {"size": 11}),
            ("5-month Sponsored Pot Growth Accelerator", {"size": 11, "bold": True}),
            (" can become a powerful ecosystem growth mechanism—attracting users, supporting builders, increasing STX activity, and introducing more participants to Bitcoin DeFi.", {"size": 11}),
        ],
    )
    add_body(
        doc,
        "This partnership is an opportunity to create a scalable incentive layer that benefits not only StacksPot, but the wider Stacks ecosystem.",
    )
    add_body(
        doc,
        "Together, we can transform every sponsored pot into a gateway for new users to experience Bitcoin-powered DeFi on Stacks.",
    )

    add_horizontal_line(doc)

    # Closing meta
    close = [
        ("Prepared by", "StacksPot Team"),
        ("Date", "July 2026"),
        ("Program length", "5 months"),
        ("Requested allocation", "~300,000 STX"),
    ]
    close_table = doc.add_table(rows=len(close), cols=2)
    for i, (label, value) in enumerate(close):
        c0, c1 = close_table.rows[i].cells
        set_cell_text(c0, label, bold=True, size=11, color=WHITE)
        shade_cell(c0, HEADER_FILL)
        set_cell_borders(c0, color="1A365D")
        set_cell_text(c1, value, size=11, color=DARK)
        set_cell_borders(c1)
        shade_cell(c1, ALT_ROW if i % 2 else "FFFFFF")
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)

    # Footer
    footer = section.footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_text(
        fp,
        "StacksPot Growth Acceleration Proposal  |  Confidential — For Stacks Endowment Review",
        size=8,
        color=GRAY,
        italic=True,
    )

    out = "/workspace/StacksPot-Growth-Acceleration-Proposal.docx"
    doc.save(out)
    print(out)


if __name__ == "__main__":
    build()
