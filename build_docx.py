# -*- coding: utf-8 -*-
"""
Build a professionally typeset .docx of "The Church Fathers"
Compatible with Google Docs.
"""
import xml.etree.ElementTree as ET
from PIL import Image as PILImage
from docx import Document
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

# ---------------------------------------------------------------- config
BODY_FONT = "EB Garamond"
INK_HEX   = "1A1712"
INK       = RGBColor(0x1A, 0x17, 0x12)
MUTED     = RGBColor(0x55, 0x50, 0x48)
CREAM_PAGE = "FAF4E6"     # page background (best-effort)
CALLBOX    = "F1EADB"     # callout shading (slightly darker than cream)
RULE_HEX   = "8A8174"

IMG_DIR = "images/final"
XML_PATH = "the-church-fathers-complete-book.xml"
OUT_PATH = "The_Church_Fathers.docx"

# ---------------------------------------------------------------- helpers
def _set_rfonts(rpr, name):
    rFonts = rpr.get_or_add_rFonts()
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), name)

def set_run(run, size=None, bold=None, italic=None, color=None,
            small_caps=None, spacing=None, name=BODY_FONT):
    rpr = run._element.get_or_add_rPr()
    _set_rfonts(rpr, name)
    if size is not None:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    if small_caps is not None:
        run.font.small_caps = small_caps
    if spacing is not None:
        sp = OxmlElement("w:spacing")
        sp.set(qn("w:val"), str(int(spacing * 20)))
        for tag in ("w:w", "w:kern", "w:position", "w:sz", "w:szCs"):
            nxt = rpr.find(qn(tag))
            if nxt is not None:
                nxt.addprevious(sp)
                break
        else:
            rpr.append(sp)

def shade_paragraph(p, fill):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)

def border_paragraph(p, edge="left", color=INK_HEX, sz=14, space=6):
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    el = OxmlElement("w:" + edge)
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    pBdr.append(el)

def add_field(p, instr, cached, size=9, color=MUTED):
    def fldchar(t):
        e = OxmlElement("w:fldChar")
        e.set(qn("w:fldCharType"), t)
        return e
    r1 = p.add_run(); r1._r.append(fldchar("begin"))
    r2 = p.add_run()
    it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
    r2._r.append(it)
    r3 = p.add_run(); r3._r.append(fldchar("separate"))
    r4 = p.add_run(); r4.text = cached
    r5 = p.add_run(); r5._r.append(fldchar("end"))
    for r in (r1, r2, r3, r4, r5):
        set_run(r, size=size, color=color)

def roman(n):
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
            (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
            (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out

# ---------------------------------------------------------------- captions
CAPTIONS = {
"ch01_fig01": "In the closing decades of the first century, Christian communities such as the one remembered at Antioch gathered in private homes, where a shared meal of blessed bread and wine was followed by the reading aloud of letters and remembered sayings. No fixed New Testament yet existed; what these believers possessed was a patchwork of writings, oral memory, and the urgent, unresolved questions that would drive the theology this book traces.",
"ch01_fig02": "Christianity began as a movement within Judaism, and its earliest followers continued to read the Hebrew Scriptures as the authoritative word of God and to understand themselves within Israel's covenant history. From this inheritance came a strict monotheism, messianic expectation, and the prophetic vocabulary through which the first believers interpreted the life, death, and resurrection of Jesus.",
"ch01_fig03": "The letters of Paul, composed roughly between the late 40s and early 60s of the first century, are the earliest surviving Christian writings and already display substantial theological reflection on Christ's death, faith, sin, and the unity of the Church. Though Paul belonged to the apostolic generation rather than to the Fathers who came after, his letters began a conversation the later Fathers would continue.",
"ch01_fig04": "In the first and second centuries no single, fixed collection of Christian writings existed, and different communities possessed different combinations of texts, copied by hand and circulating slowly across a wide and thinly connected world. The gradual consolidation of a recognized canon was itself one of the great unfinished tasks the Church Fathers inherited and helped to shape.",

"ch02_fig01": "Around the early second century, Ignatius, bishop of Antioch, was transported under military guard along the roads of Asia Minor toward execution in Rome, and he composed seven letters along the way. Those letters survive as one of the clearest windows into Christian thought in the decades immediately after the apostolic generation.",
"ch02_fig02": "Ignatius insisted on the authority of a single bishop in each local church and connected the unity of the community directly to the bread and wine of the Eucharist, which he understood as genuinely joined to the flesh and blood of a Christ who truly suffered. His urgency reflects how far this pattern of leadership was still from being a settled norm.",
"ch02_fig03": "Polycarp of Smyrna was remembered by Irenaeus as a personal link to the apostle John, and his execution became the subject of one of the earliest surviving detailed martyrdom accounts. The Martyrdom of Polycarp helped establish the literary conventions of courtroom dialogue and final testimony that later martyr narratives would imitate.",
"ch02_fig04": "Later tradition, transmitted through Irenaeus, claimed that Polycarp had been personally instructed by the apostle John and appointed to office by apostles themselves, situating him as a living bridge between the apostolic generation and the emerging patristic tradition. Historians weigh such claims carefully, but they carried enormous authority for how the early Church understood its own continuity.",

"ch03_fig01": "In mid-second-century Rome, Justin presented himself as a professional philosopher who argued that Christianity was not the enemy of philosophy but its truest fulfillment. With him, the recurring tension between Christian faith and Greek thought moves from background context into direct, sustained argument.",
"ch03_fig02": "Justin's Dialogue with Trypho describes his encounter with an elderly man, perhaps by the seashore, who argued that true knowledge of God required divine revelation rather than the soul's own contemplative effort, and who directed him to the Hebrew prophets. Whatever its precise historical accuracy, the narrative presents Christianity as the completion of serious philosophical searching.",
"ch03_fig03": "According to the Acts of Justin and his Companions, Justin was brought before the Roman prefect Rusticus around the year 165 and, refusing to sacrifice to the traditional gods, was sentenced to be scourged and beheaded. As with other martyrdom accounts, historians distinguish the underlying event from the literary shaping of the courtroom dialogue.",
"ch03_fig04": "Before his conversion, Justin had moved through the Stoic, Peripatetic, and Pythagorean schools before finding in Middle Platonism an account of transcendent reality that briefly satisfied him. His later theory of the logos spermatikos, the seed of the divine Word scattered throughout history, grew directly out of that philosophical formation.",

"ch04_fig01": "In the year 177, a violent persecution broke out against the Christians of Lyon in Roman Gaul, claiming the elderly bishop Pothinus among its victims. Irenaeus, then a presbyter away on a mission to Rome, survived the crisis and returned to succeed Pothinus as bishop, a position from which he composed his great refutation of rival Christian teachers.",
"ch04_fig02": "Irenaeus described at length the teaching associated with Valentinus, in which a realm of paired divine emanations called Aeons surrounded the unknowable divine source, and a crisis surrounding the Aeon Sophia was said to account for the flawed material world. Modern readers should remember that his polemical presentation may flatten real differences among the teachers he discusses.",
"ch04_fig03": "Irenaeus presented himself not as an innovator but as a guardian of teaching transmitted publicly and continuously from the apostles, and his major work, Against Heresies, was written to overthrow what he called falsely named knowledge. His insistence that there is one God, both creator and redeemer, remains among his most influential arguments.",
"ch04_fig04": "For centuries, Irenaeus's hostile summaries were historians' primary access to the teachers he opposed, because many of their original texts had not survived. The 1945 discovery of a substantial collection of ancient texts near Nag Hammadi allowed scholars to compare his account against original sources, confirming its broad outlines while revealing places where polemical simplification had occurred.",

"ch05_fig01": "Alexandria, home of one of the ancient world's most famous libraries, was the intellectual center where Christian teachers established a catechetical school for instructing converts and inquirers. It was to this school, after Clement of Alexandria's departure, that the young Origen succeeded, becoming the most theologically ambitious Christian writer of the pre-Nicene period.",
"ch05_fig02": "According to Eusebius, a wealthy patron named Ambrose supported Origen's work with a team of shorthand writers, copyists, and calligraphers, allowing him to dictate his vast output to multiple scribes working in shifts. This support helps explain how a single scholar produced commentaries, homilies, and treatises on a staggering scale.",
"ch05_fig03": "Origen's Hexapla arranged the Old Testament in six parallel columns, placing the Hebrew text, its Greek transliteration, and four Greek translations side by side for comparison. It was one of antiquity's most ambitious works of textual scholarship, though it survives today only in fragments preserved by later writers.",
"ch05_fig04": "After a rupture with Bishop Demetrius of Alexandria over his ordination in Caesarea Maritima, Origen relocated permanently to Caesarea, where he established a new school and continued to attract students from across the region. His relationship with episcopal authority illustrates the developing tensions around jurisdiction that could affect even the most brilliant scholars of the age.",

"ch06_fig01": "The great North African port of Carthage was the center of a Latin-speaking Christian tradition that emerged distinctly from the Greek-speaking East. It was from this environment that Christian theology acquired much of its most enduring Latin technical vocabulary.",
"ch06_fig02": "Tertullian, trained in rhetoric and law before his conversion, became the most prolific Christian Latin writer of his generation and the first known author to use the term trinitas. His distinctions of persona and substantia gave later theology essential tools for describing how Father, Son, and Spirit could be three and yet one.",
"ch06_fig03": "During the empire-wide persecution under Decius around the year 250, Cyprian of Carthage withdrew into hiding outside the city and continued to direct his congregation through a steady stream of letters. His decision drew sharp criticism from some contemporaries, illustrating how seriously faithful Christian leaders could disagree about whether flight or confrontation was the more responsible course.",
"ch06_fig04": "Cyprian's insistence on church unity led to a sharp dispute with Stephen of Rome over the validity of baptism performed by heretical or schismatic groups. Cyprian argued that such baptism required repetition upon reconciliation with the Church, while Stephen's opposing view would eventually prevail in later Western practice.",

"ch07_fig01": "In the early summer of 325, several hundred bishops gathered at Nicaea at imperial expense to resolve the dispute over Arius's teaching that the Son was a created being. The council adopted the term homoousios, 'of one substance,' to exclude any reading that placed the Son among created things.",
"ch07_fig02": "Constantine's involvement marked a new development: an emperor convening a gathering of bishops out of concern for imperial and ecclesiastical unity. Yet the emperor did not invent Christian doctrine or decide the council's theology; the substantive argument and decision rested with the assembled bishops themselves.",
"ch07_fig03": "Athanasius spent nearly five decades defending the Nicene formula, during which successive emperors of differing theological sympathies drove him into exile from Alexandria on five separate occasions. His endurance earned him the later epithet Athanasius contra mundum, 'Athanasius against the world.'",
"ch07_fig04": "In On the Incarnation, Athanasius argued that only a Son truly and fully divine could accomplish salvation, since a merely exalted creature could not bridge the gap between humanity and the uncreated God. This argument connected the technical debate over homoousios directly to the practical question of how human salvation actually works.",

"ch08_fig01": "The rugged volcanic landscape of Cappadocia in central Asia Minor was home, by the fourth century, to a flourishing tradition of monastic settlement, and it produced three of the most theologically consequential figures examined in this book: Basil of Caesarea, his brother Gregory of Nyssa, and their friend Gregory of Nazianzus.",
"ch08_fig02": "Basil and Gregory of Nazianzus formed a close and lasting friendship while studying rhetoric and philosophy together in Athens, still a great center of classical learning. The friendship would shape both men's careers, and Gregory would later describe it as one of the defining relationships of his life.",
"ch08_fig03": "As bishop of Caesarea, Basil constructed on the outskirts of the city an extensive charitable complex known as the Basiliad, combining a hospital, a hospice for travelers, and a facility for the poor and sick. The project reflected a conviction that theological orthodoxy and practical Christian charity belonged inseparably together.",
"ch08_fig04": "Gregory of Nyssa's dialogue On the Soul and Resurrection presents his sister Macrina, during her final illness, as a genuine philosophical teacher guiding him through arguments about the soul, death, and resurrection. Her portrait reaches modern readers through a devoted brother's literary shaping, but it offers a valuable glimpse into the role women could play in shaping Christian intellectual life.",

"ch09_fig01": "In the late summer of 386, a thirty-one-year-old professor of rhetoric wept beneath a fig tree in a Milanese garden, torn between years of searching and the faith his mother Monica had long prayed he would accept. His account of hearing 'take up and read' and opening Paul's letters marks the climactic moment of his conversion, though the Confessions was written about a decade later and shaped by his mature convictions.",
"ch09_fig02": "Augustine's movement toward conversion accelerated in Milan, where the eloquent bishop Ambrose's sophisticated allegorical interpretation of the Old Testament dissolved intellectual objections Augustine had carried since his Manichaean years. Ambrose's preaching proved a decisive influence on the future bishop of Hippo.",
"ch09_fig03": "Ordained against his initial wishes after being effectively drafted by the congregation of Hippo Regius, Augustine became bishop in the mid-390s and produced an extraordinary body of writing on nearly every major theological controversy of his era. His work would shape Western Christian thought more profoundly than any other figure in this book.",
"ch09_fig04": "The sack of Rome by Alaric and his Gothic forces in the year 410 sent shockwaves throughout the Roman world and prompted pagan critics to blame Christianity for the catastrophe. Augustine responded over more than a decade with the City of God, a philosophy of history distinguishing the earthly city from the heavenly city.",

"ch10_fig01": "Centuries after Augustine, a monastic copyist bent over a desk surrounded by the letters of Ignatius, the arguments of Irenaeus, fragments of Origen, the creeds of Nicaea and Constantinople, and a worn copy of the Confessions. This scene captures how the individual, contested arguments of the Fathers gradually became a shared, authoritative tradition.",
"ch10_fig02": "The practical mechanisms by which certain writings achieved lasting authority included the compilation of catena literature, chains of quotations from respected authorities, and the selective copying and preservation of manuscripts by communities with their own theological commitments. The surviving textual record thus already embodies centuries of judgment about what was worth preserving.",
"ch10_fig03": "Across the arc of this book, the Greek East and the Latin West developed distinct theological temperaments, the East centered on trinitarian precision and theosis, the West on sin, grace, and ecclesiastical unity. This divergence, shaped by language, history, and institutions, would eventually contribute to the formal division between Eastern Orthodox and Western Catholic Christianity.",
"ch10_fig04": "Understanding how Christian doctrine actually took shape, through argument, political circumstance, and the unpredictable judgments of later generations, offers a richer appreciation of beliefs that can otherwise seem always to have existed in their present form. The long, contested, and deeply human process traced in this book remains, in the end, its most enduring lesson.",
}

# figure -> (anchor_section_or_None_for_intro)
ANCHORS = {
    1: [("ch01_fig01", None), ("ch01_fig02", "The Jewish Roots"),
        ("ch01_fig03", "Paul and Early Christian Theology"), ("ch01_fig04", "The Canon Was Not Yet Fixed")],
    2: [("ch02_fig01", "A Bishop in Chains"), ("ch02_fig02", "The Body and Blood of Christ"),
        ("ch02_fig04", "Polycarp of Smyrna"), ("ch02_fig03", "The Martyrdom of Polycarp")],
    3: [("ch03_fig01", None), ("ch03_fig04", "A Search Among the Schools"),
        ("ch03_fig02", "The Old Man by the Sea"), ("ch03_fig03", "Philosopher and Martyr")],
    4: [("ch04_fig01", None), ("ch04_fig03", "Falsely-Named Knowledge"),
        ("ch04_fig02", "The World of Valentinus"), ("ch04_fig04", "Reading Irenaeus Sideways")],
    5: [("ch05_fig01", None), ("ch05_fig02", "A Man of Steel"),
        ("ch05_fig03", "The Hexapla"), ("ch05_fig04", "A Priest Without a Church's Blessing")],
    6: [("ch06_fig01", None), ("ch06_fig02", "A Lawyer's Mind for the Faith"),
        ("ch06_fig03", "A Bishop in Hiding"), ("ch06_fig04", "The Baptismal Controversy")],
    7: [("ch07_fig01", None), ("ch07_fig02", "An Emperor's Concern for Unity"),
        ("ch07_fig03", "Athanasius Against the World"), ("ch07_fig04", "Why the Word Became Flesh")],
    8: [("ch08_fig01", None), ("ch08_fig02", "Basil and Gregory: Friends from Athens"),
        ("ch08_fig03", "Basil the Bishop"), ("ch08_fig04", "A Sister's Philosophical Deathbed")],
    9: [("ch09_fig01", None), ("ch09_fig02", "Ambrose and the Milan Circle"),
        ("ch09_fig03", "Bishop of Hippo"), ("ch09_fig04", "The City of God")],
    10: [("ch10_fig01", None), ("ch10_fig02", "From Argument to Authority"),
         ("ch10_fig03", "Diverging Paths: East and West"), ("ch10_fig04", "Why This History Still Matters")],
}

# ---------------------------------------------------------------- parse
def parse_blocks(text):
    lines = text.splitlines()
    blocks, i, n = [], 0, len(lines)
    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if s.startswith("**") and s.endswith("**") and len(s) > 4:
            blocks.append(("heading", s[2:-2].strip()))
            i += 1
        elif s.startswith(">"):
            if s.startswith("> **"):
                title = s[2:].strip().strip("*").strip()
                body = []
            else:
                title = None
                body = [s[1:].strip()]
            i += 1
            while i < n and lines[i].strip().startswith(">"):
                body.append(lines[i].strip()[1:].strip())
                i += 1
            blocks.append(("callout", title, body))
        else:
            blocks.append(("para", s))
            i += 1
    return blocks

# ---------------------------------------------------------------- document
doc = Document()

sec = doc.sections[0]
sec.page_width = Cm(15.24)
sec.page_height = Cm(22.90)
sec.top_margin = Cm(1.5)
sec.bottom_margin = Cm(1.0)
sec.left_margin = Cm(1.9)    # inside (binding)
sec.right_margin = Cm(1.6)   # outside
sec.gutter = Cm(0)
sec.footer_distance = Cm(0.7)

# Mirrored (facing-page) margins: left/right swap on even pages.
sectPr = sec._sectPr
pgMar = sectPr.find(qn("w:pgMar"))
pgMar.set(qn("w:mirrorMargins"), "1")

# Page background color lives in document.xml as the first child of w:document,
# before w:body (Word's Design > Page Color writes it there).
bg = OxmlElement("w:background")
bg.set(qn("w:color"), CREAM_PAGE)
doc.element.insert(0, bg)

normal = doc.styles["Normal"]
normal.font.name = BODY_FONT
normal.font.size = Pt(11)
normal.font.color.rgb = INK
_set_rfonts(normal.element.get_or_add_rPr(), BODY_FONT)
npf = normal.paragraph_format
npf.space_after = Pt(0)
npf.line_spacing = 1.28

h1 = doc.styles["Heading 1"]
h1.font.name = BODY_FONT
h1.font.size = Pt(22)
h1.font.bold = True
h1.font.color.rgb = INK
_set_rfonts(h1.element.get_or_add_rPr(), BODY_FONT)

h2 = doc.styles["Heading 2"]
h2.font.name = BODY_FONT
h2.font.size = Pt(13.5)
h2.font.bold = True
h2.font.color.rgb = INK
_set_rfonts(h2.element.get_or_add_rPr(), BODY_FONT)

footer = sec.footer
footer.is_linked_to_previous = False
fp = footer.paragraphs[0]
for r in list(fp.runs):
    r._element.getparent().remove(r._element)
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_field(fp, " PAGE ", "1", size=9, color=MUTED)

# ---------------------------------------------------------------- builders
def add_body(text, first=False, italic=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.line_spacing = 1.28
    pf.space_after = Pt(0)
    if not first:
        pf.first_line_indent = Cm(0.5)
    r = p.add_run(text)
    set_run(r, size=11, italic=italic)
    return p

def add_section_heading(text):
    p = doc.add_paragraph(style="Heading 2")
    pf = p.paragraph_format
    pf.space_before = Pt(16)
    pf.space_after = Pt(6)
    pf.keep_with_next = True
    r = p.add_run(text)
    set_run(r, size=13.5, bold=True)
    return p

def add_callout(title, body):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(10)
    pf.space_after = Pt(2)
    pf.keep_with_next = True
    shade_paragraph(p, CALLBOX)
    border_paragraph(p, "left", INK_HEX, 14, 8)
    r = p.add_run(title)
    set_run(r, size=11, bold=True, small_caps=True, spacing=0.8)
    for line in body:
        if line.startswith("- "):
            bp = doc.add_paragraph()
            bpf = bp.paragraph_format
            bpf.left_indent = Cm(0.9)
            bpf.first_line_indent = Cm(-0.35)
            bpf.space_after = Pt(2)
            bpf.line_spacing = 1.22
            shade_paragraph(bp, CALLBOX)
            border_paragraph(bp, "left", INK_HEX, 14, 8)
            rb = bp.add_run("\u2013  ")
            set_run(rb, size=10.5)
            rt = bp.add_run(line[2:].strip())
            set_run(rt, size=10.5)
        else:
            tp = doc.add_paragraph()
            tpf = tp.paragraph_format
            tpf.left_indent = Cm(0.3)
            tpf.space_after = Pt(2)
            tpf.line_spacing = 1.22
            shade_paragraph(tp, CALLBOX)
            border_paragraph(tp, "left", INK_HEX, 14, 8)
            if title == "Who's Who" and " \u2014 " in line:
                name, rest = line.split(" \u2014 ", 1)
                rn = tp.add_run(name)
                set_run(rn, size=10.5, bold=True)
                rr = tp.add_run(" \u2014 " + rest)
                set_run(rr, size=10.5, italic=True)
            else:
                rt = tp.add_run(line)
                set_run(rt, size=10.5, italic=True)
    endp = doc.add_paragraph()
    endp.paragraph_format.space_after = Pt(6)
    endp.paragraph_format.line_spacing = 1.0

def add_image(fig, caption):
    path = f"{IMG_DIR}/{fig}.jpg"
    with PILImage.open(path) as im:
        w, h = im.size
    width = Inches(4.6) if w >= h else Inches(2.7)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=width)
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cpf = cp.paragraph_format
    cpf.left_indent = Cm(0.7)
    cpf.right_indent = Cm(0.7)
    cpf.space_after = Pt(12)
    cpf.line_spacing = 1.15
    cr = cp.add_run(caption)
    set_run(cr, size=10.5, italic=True, color=MUTED)

def add_rule(centered=True, left=1.6, right=1.6, sz=6, space=8):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if centered else WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    if centered:
        pf.left_indent = Cm(left)
        pf.right_indent = Cm(right)
    pf.space_before = Pt(4)
    pf.space_after = Pt(14)
    border_paragraph(p, "bottom", RULE_HEX, sz, space)
    return p

def add_chapter_opener(num, title):
    if num:
        p1 = doc.add_paragraph()
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(4)
        p1.paragraph_format.page_break_before = True
        r1 = p1.add_run("CHAPTER " + num)
        set_run(r1, size=11, small_caps=True, spacing=2.2, color=MUTED)
        p2 = doc.add_paragraph(style="Heading 1")
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(6)
        p2.paragraph_format.keep_with_next = True
        r2 = p2.add_run(title)
        set_run(r2, size=22, bold=True)
    else:
        p2 = doc.add_paragraph(style="Heading 1")
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(6)
        p2.paragraph_format.page_break_before = True
        p2.paragraph_format.keep_with_next = True
        r2 = p2.add_run(title)
        set_run(r2, size=22, bold=True)
    add_rule()

def add_part_title(text):
    p = doc.add_paragraph(style="Heading 1")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.page_break_before = True
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(8)
    r = p.add_run(text)
    set_run(r, size=22, bold=True)
    add_rule()

# ---------------------------------------------------------------- front matter / title
title_full = "The Church Fathers: The Thinkers Who Shaped Christian Doctrine"
main_title, subtitle = title_full.split(": ", 1)

for _ in range(5):
    sp = doc.add_paragraph()
    sp.paragraph_format.space_after = Pt(0)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("\u00b7 \u00b7 \u00b7"); set_run(r, size=14, spacing=8, color=MUTED)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_before = Pt(34); p.paragraph_format.space_after = Pt(14)
r = p.add_run(main_title); set_run(r, size=34, bold=True)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.paragraph_format.space_after = Pt(30)
r = p.add_run(subtitle); set_run(r, size=15, italic=True, color=MUTED)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("\u00b7 \u00b7 \u00b7"); set_run(r, size=14, spacing=8, color=MUTED)

# --- blank page (page 2)
pb = doc.add_paragraph()
pb.add_run().add_break(WD_BREAK.PAGE)
blank = doc.add_paragraph()

# --- table of contents (page 3)
toc_head = doc.add_paragraph(style="Heading 1")
toc_head.alignment = WD_ALIGN_PARAGRAPH.CENTER
toc_head.paragraph_format.page_break_before = True
toc_head.paragraph_format.space_before = Pt(0)
toc_head.paragraph_format.space_after = Pt(6)
tr = toc_head.add_run("Contents"); set_run(tr, size=22, bold=True)
add_rule()

def toc_part(label):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(label)
    set_run(r, size=11, small_caps=True, spacing=1.6, color=MUTED)

def toc_item(label, title=None):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    if title is None:
        r = p.add_run(label)
        set_run(r, size=11.5)
    else:
        rn = p.add_run(label + "  ")
        set_run(rn, size=11.5, bold=True)
        rt = p.add_run(title)
        set_run(rt, size=11.5)

toc_part("Front Matter")
toc_item("Dedication")
toc_item("Prologue")

tree = ET.parse(XML_PATH)
root = tree.getroot()
chapters = root.find("chapters").findall("chapter")
chapter_titles = [ch.find("chapter_title").text.strip() for ch in chapters]

toc_part("The Chapters")
for i, t in enumerate(chapter_titles, 1):
    toc_item(roman(i), t)
toc_part("Closing")
bonus_title = root.find("bonus_chapter").find("chapter_title").text.strip()
toc_item(bonus_title)

# --- dedication
fm = root.find("front_matter")
ded = fm.find("dedication").text.strip()
add_part_title("Dedication")
for para in [x for x in ded.splitlines() if x.strip()]:
    add_body(para.strip(), italic=True)

# --- prologue
pro = fm.find("prologue").text.strip()
add_part_title("Prologue")
pro_paras = [x for x in pro.splitlines() if x.strip()]
for j, para in enumerate(pro_paras):
    add_body(para.strip(), first=(j == 0))

# --- chapters
for idx, ch in enumerate(chapters, 1):
    title = ch.find("chapter_title").text.strip()
    content = ch.find("content").text
    add_chapter_opener(roman(idx), title)

    placed = set()
    current_section = None
    last_was_text = False

    def flush(section):
        for fig, anchor in ANCHORS[idx]:
            if fig not in placed and anchor == section:
                add_image(fig, CAPTIONS[fig])
                placed.add(fig)

    for block in parse_blocks(content):
        if block[0] == "heading":
            flush(current_section)
            current_section = block[1]
            add_section_heading(block[1])
            last_was_text = False
        elif block[0] == "para":
            add_body(block[1], first=not last_was_text)
            last_was_text = True
        else:
            add_callout(block[1], block[2])
            last_was_text = False

    flush(current_section)
    for fig, anchor in ANCHORS[idx]:
        if fig not in placed:
            add_image(fig, CAPTIONS[fig])
            placed.add(fig)

# --- bonus chapter
bonus = root.find("bonus_chapter")
add_chapter_opener("", bonus.find("chapter_title").text.strip())
bcontent = bonus.find("content").text
b_first = True
for block in parse_blocks(bcontent):
    if block[0] == "heading":
        add_section_heading(block[1])
        b_first = False
    elif block[0] == "para":
        add_body(block[1], first=b_first)
        b_first = False
    else:
        add_callout(block[1], block[2])
        b_first = False

doc.core_properties.title = title_full
doc.core_properties.subject = "Church Fathers"

doc.save(OUT_PATH)
print("Saved", OUT_PATH)
