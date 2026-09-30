"""Grade 7 English Adventure — 60 min, Past Perfect + The Secret Door in the Library."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

NAVY = RGBColor(0x0D, 0x47, 0xA1)
SKY = RGBColor(0x29, 0xB6, 0xF6)
TEAL = RGBColor(0x00, 0x89, 0x7B)
AMBER = RGBColor(0xFF, 0xB3, 0x00)
CORAL = RGBColor(0xFF, 0x70, 0x43)
PURPLE = RGBColor(0x7E, 0x57, 0xC2)
PINK = RGBColor(0xEC, 0x40, 0x7A)
CREAM = RGBColor(0xFF, 0xFB, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1A, 0x23, 0x37)
SOFT = RGBColor(0x64, 0x74, 0x8B)
LIGHT_SKY = RGBColor(0xE1, 0xF5, 0xFE)
LIGHT_TEAL = RGBColor(0xE0, 0xF2, 0xF1)
LIGHT_AMBER = RGBColor(0xFF, 0xF8, 0xE1)
LIGHT_CORAL = RGBColor(0xFF, 0xEB, 0xEE)
LIGHT_PURPLE = RGBColor(0xF3, 0xE5, 0xF5)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
TOTAL = 22

PASSAGE = (
    "Riley had never noticed the small brass handle behind the biography shelf at Westbrook Middle "
    "School Library—not until Tuesday afternoon, when a loose book revealed a narrow door painted "
    "the color of old parchment.\n\n"
    "Earlier that day, Riley had finished a science quiz and had hurried to return a novel before "
    "study hall. The librarian, Ms. Delgado, had already left for a meeting, so the room felt unusually "
    "quiet. Riley had whispered, \"Just one quick look,\" and had tugged the handle.\n\n"
    "The door creaked open. Inside stood a spiral staircase lit by glowing blue moss. Riley had "
    "clutched a flashlight from a nearby cart and had stepped down carefully. At the bottom, shelves "
    "held books that hummed softly, as if they remembered every story ever told.\n\n"
    "A note on a desk read: \"If you have read Chapter One, you may enter Chapter Two.\" Riley had "
    "realized the library was larger than anyone had ever guessed. Footsteps echoed above—Jordan, "
    "Riley's best friend, had followed after seeing the open door.\n\n"
    "\"We probably should have told an adult,\" Jordan murmured. Riley had nodded, but curiosity had "
    "already won. Together they had opened a book titled The Map of Unfinished Adventures. The pages "
    "had turned by themselves, showing a hallway that did not exist on any school floor plan.\n\n"
    "Then the lights flickered. A voice from the staircase called Riley's name. Had someone discovered "
    "the secret first? Riley and Jordan exchanged a glance, hearts racing, as the humming grew louder "
    "and the next chapter waited—just out of reach."
)

VOCAB = [
    ("biography", "bio-GRAH-fee", "A book about the life of a real person", "📖", "Riley grabbed a biography about an astronaut.", "life story"),
    ("parchment", "PAR-chmunt", "Thick paper used for old documents", "📜", "The door matched the color of parchment.", "manuscript"),
    ("creaked", "KREEKT", "Made a long squeaky sound", "🚪", "The door creaked when Riley pulled it.", "groaned"),
    ("moss", "MAWS", "Small green plant that grows on stones", "🌿", "Blue moss lit the staircase.", "lichen"),
    ("curiosity", "kyoor-ee-AH-suh-tee", "A strong wish to learn or know", "🔍", "Curiosity led Riley downward.", "interest"),
    ("echoed", "EK-ohd", "Bounced back as a repeating sound", "🔊", "Footsteps echoed above them.", "repeated"),
    ("flickered", "FLIK-erd", "Flashed on and off quickly", "💡", "The lights flickered suddenly.", "blinked"),
    ("unusual", "un-YOO-zhoo-ul", "Not common; different", "✨", "The library felt unusual and quiet.", "strange"),
    ("realized", "REE-uh-lyzed", "Understood clearly", "💭", "Riley realized the library was huge.", "discovered"),
    ("adventure", "ad-VEN-chur", "An exciting experience", "🗺️", "The map showed unfinished adventures.", "quest"),
]


def set_run(run, size=18, bold=False, color=DARK, font="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_bg(slide, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    sp = s._element
    spTree = slide.shapes._spTree
    spTree.remove(sp)
    spTree.insert(2, sp)


def add_rect(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s


def add_round(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s


def add_oval(slide, l, t, w, h, fill):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = fill
    s.line.fill.background()
    return s


def tb(slide, l, t, w, h, text, size=18, bold=False, color=DARK, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, color, font)
    return box


def bullets(slide, l, t, w, h, items, size=15, color=DARK, sp=8):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(sp)
        r = p.add_run()
        r.text = "•  " + item
        set_run(r, size, False, color)


def teacher_notes(slide, say, activity, expected, assess, mistakes="", extend="", timing=""):
    parts = [
        "WHAT TO SAY:\n" + say,
        "\nSTUDENT ACTIVITY:\n" + activity,
        "\nEXPECTED ANSWERS / RESPONSES:\n" + expected,
    ]
    if mistakes:
        parts.append("\nCOMMON MISTAKES:\n" + mistakes)
    parts.append("\nASSESSMENT TIPS:\n" + assess)
    if extend:
        parts.append("\nEXTENSION QUESTIONS:\n" + extend)
    if timing:
        parts.append("\nTIMING: " + timing)
    slide.notes_slide.notes_text_frame.text = "\n".join(parts)


def fade(slide):
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    etree.SubElement(tr, qn("p:fade"))


def footer(slide, n, timing=""):
    add_rect(slide, Inches(0), Inches(7.12), prs.slide_width, Inches(0.38), NAVY)
    msg = "Grade 7 English Adventure  |  60 min  |  Past Perfect"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.15), Inches(11.5), Inches(0.3), msg, size=10, color=WHITE)
    tb(slide, Inches(12.0), Inches(7.15), Inches(1.1), Inches(0.3), f"{n}/{TOTAL}", size=10,
       color=WHITE, align=PP_ALIGN.RIGHT)


def chip(slide, text, color=SKY, w=Inches(2.8)):
    add_round(slide, Inches(0.38), Inches(0.28), w, Inches(0.36), color)
    tb(slide, Inches(0.38), Inches(0.29), w, Inches(0.34), text, size=11, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def time_chip(slide, text):
    add_round(slide, Inches(10.45), Inches(0.28), Inches(2.5), Inches(0.36), CORAL)
    tb(slide, Inches(10.45), Inches(0.29), Inches(2.5), Inches(0.34), text, size=11, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def header(slide, title, tag="ACTIVITY", timing="", page=1):
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.14), prs.slide_height, SKY)
    chip(slide, tag)
    if timing:
        time_chip(slide, timing)
    tb(slide, Inches(0.38), Inches(0.78), Inches(12.5), Inches(0.55), title, size=26, bold=True,
       color=NAVY, font="Georgia")
    footer(slide, page, timing)
    fade(slide)


def btn(slide, label="Continue Adventure →"):
    add_round(slide, Inches(10.15), Inches(6.55), Inches(2.95), Inches(0.45), AMBER)
    tb(slide, Inches(10.15), Inches(6.58), Inches(2.95), Inches(0.4), label, size=11, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER)


# --- Slides ---

def s01():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), AMBER)
    for x, y, c in [(0.6, 1.0, TEAL), (11.8, 1.3, PINK), (1.2, 5.6, SKY)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.55), Inches(0.55), c)
    tb(slide, Inches(0.7), Inches(2.2), Inches(12), Inches(1.0),
       "Welcome to Today's\nEnglish Adventure!", size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(3.8), Inches(12), Inches(0.45),
       "Grade 7  •  United States  •  60 Minutes  •  Mystery + Grammar Quest", size=17,
       color=RGBColor(0xBB, 0xDE, 0xFB), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.2), Inches(4.6), Inches(6.9), Inches(1.1), TEAL)
    tb(slide, Inches(3.4), Inches(4.85), Inches(6.5), Inches(0.6),
       "Past Perfect  •  Reading  •  Vocabulary  •  Storytelling\n(An adventure—not a lecture!)", size=15, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, 1, "Welcome 5 min")
    fade(slide)
    teacher_notes(slide,
        "Welcome warmly. Set playful tone: 'We are detectives and storytellers today.'",
        "Share one word you hope to learn today.",
        "Student names a skill/word; shows mood.",
        "Confidence, engagement baseline.",
        "Saying 'I don't know'—normalize trying.",
        "What kind of adventure stories do you like?",
        "5 min")


def s02():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Fun Icebreaker: This or That 🎯", "GAME", "Icebreaker 5 min", 2)
    pairs = [
        ("Mystery novels", "Science fiction"),
        ("Read alone", "Read with a friend"),
        ("Library at night", "Library at lunch"),
        ("Hidden doors", "Hidden maps"),
        ("Write stories", "Act stories out"),
        ("Quiet puzzles", "Fast quizzes"),
    ]
    for i, (a, b) in enumerate(pairs):
        col, row = i % 3, i // 3
        left = Inches(0.38 + col * 4.3)
        top = Inches(1.45 + row * 2.45)
        add_round(slide, left, top, Inches(4.05), Inches(2.2), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(0.35), Inches(3.75), Inches(0.55), a, size=15, bold=True,
           color=TEAL, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.0), Inches(3.75), Inches(0.35), "OR", size=12, bold=True,
           color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.45), Inches(3.75), Inches(0.55), b, size=15, bold=True,
           color=CORAL, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.38), Inches(6.35), Inches(12.5), Inches(0.35),
       "Pick one and explain why in a full sentence.", size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Do 4–6 pairs quickly.",
        "Choose + because…",
        "Full sentences with reasons.",
        "Speaking fluency, preferences.",
        "One-word answers.",
        "Would your choice change on weekends?",
        "5 min icebreaker")


def s03():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Brain Warm-Up: Guess the Emoji 🧠", "GAME", "5 min", 3)
    emojis = ["📚", "🗝️", "🚪", "👣", "💡", "🗺️"]
    hints = ["You read here", "Opens locks", "Secret entrance", "Sound repeats", "Idea moment", "Shows a route"]
    for i, (e, h) in enumerate(zip(emojis, hints)):
        left = Inches(0.38 + (i % 3) * 4.3)
        top = Inches(1.45 + (i // 3) * 2.55)
        add_round(slide, left, top, Inches(4.05), Inches(2.3), LIGHT_SKY)
        tb(slide, left, top + Inches(0.35), Inches(4.05), Inches(0.7), e, size=40, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.35), Inches(3.75), Inches(0.7), h, size=13, color=SOFT,
           align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Cover hints; student guesses word/idea.",
        "Match emoji to hint; say a sentence.",
        "library, key, door, echo, realize, map (accept close).",
        "Quick listening + vocabulary warm-up.",
        "Random guesses—give second clue.",
        "Link to today's library story.",
        "5 min")


def s04():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Today's Mission Objectives 🎯", "GOALS", "1 min", 4)
    goals = [
        (SKY, "Past Perfect", "Use had + past participle correctly"),
        (TEAL, "Reading", "Understand mystery text + infer"),
        (PURPLE, "Vocabulary", "Learn 10 story words"),
        (CORAL, "Speaking", "Describe, interview, create stories"),
        (AMBER, "Games", "Stay engaged every 5–7 minutes"),
    ]
    for i, (c, t, d) in enumerate(goals):
        top = Inches(1.45 + i * 1.05)
        add_round(slide, Inches(0.38), top, Inches(12.5), Inches(0.95), WHITE)
        add_round(slide, Inches(0.58), top + Inches(0.22), Inches(0.55), Inches(0.55), c)
        tb(slide, Inches(0.58), top + Inches(0.3), Inches(0.55), Inches(0.4), str(i + 1), size=16, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.35), top + Inches(0.18), Inches(11.3), Inches(0.35), t, size=18, bold=True, color=NAVY)
        tb(slide, Inches(1.35), top + Inches(0.55), Inches(11.3), Inches(0.35), d, size=14, color=SOFT)
    btn(slide)
    teacher_notes(slide,
        "Quick tour—student repeats one goal in own words.",
        "Which goal excites you most?",
        "Student picks goal; shows motivation.",
        "Goal-setting buy-in.",
        "Disengagement—connect to games.",
        "How will you know you improved today?",
        "1 min")


def s05():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Past Perfect Tense ⏳", "GRAMMAR", "Grammar 15 min", 5)
    add_round(slide, Inches(0.38), Inches(1.4), Inches(12.5), Inches(2.0), WHITE)
    tb(slide, Inches(0.6), Inches(1.55), Inches(12.0), Inches(1.7),
       "Definition: Shows an action that happened BEFORE another action in the past.\n\n"
       "Structure:  had + past participle  (had opened, had finished, had realized)\n\n"
       "Signal words: before, after, already, just, never, by the time, until",
       size=15, color=DARK)
    cols = [
        ("When to use", "Earlier past action\nTwo past events\nExperience before a moment", TEAL),
        ("Common mistakes", "Using simple past for BOTH events\nForgetting 'had'\nWrong participle (had went)", CORAL),
        ("vs Simple Past", "Past Perfect = earlier\nSimple Past = main event\nI had eaten before the bell rang.", SKY),
    ]
    for i, (t, d, c) in enumerate(cols):
        left = Inches(0.38 + i * 4.25)
        add_round(slide, left, Inches(3.55), Inches(4.05), Inches(2.85), LIGHT_TEAL if i == 0 else LIGHT_AMBER if i == 1 else LIGHT_SKY)
        tb(slide, left + Inches(0.15), Inches(3.75), Inches(3.75), Inches(0.4), t, size=15, bold=True, color=c)
        tb(slide, left + Inches(0.15), Inches(4.25), Inches(3.75), Inches(1.8), d, size=13, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Teach with timeline gesture: earlier ← had ___, later ← simple past.",
        "Repeat structure; create 1 sentence about school day.",
        "Sample: I had finished homework before dinner.",
        "Conceptual grasp of sequence in past.",
        "had + past tense (had went).",
        "What happened first in: 'After I had locked the door, I left'?",
        "Grammar 15 min start")


def s06():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Past Perfect Timeline 📊", "GRAMMAR", "Grammar", 6)
    events = [
        ("1️⃣ Earlier", "Riley had returned the novel.", TEAL),
        ("2️⃣ Earlier", "Riley had opened the secret door.", SKY),
        ("3️⃣ Later", "Jordan followed Riley.", CORAL),
        ("4️⃣ Later", "The lights flickered.", PURPLE),
    ]
    add_rect(slide, Inches(0.8), Inches(3.2), Inches(11.7), Inches(0.12), NAVY)
    for i, (lab, sent, c) in enumerate(events):
        left = Inches(0.55 + i * 3.15)
        add_oval(slide, left + Inches(1.0), Inches(3.05), Inches(0.35), Inches(0.35), c)
        add_round(slide, left, Inches(1.55), Inches(3.0), Inches(1.35), WHITE)
        tb(slide, left + Inches(0.1), Inches(1.7), Inches(2.8), Inches(0.35), lab, size=13, bold=True, color=c,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(2.15), Inches(2.8), Inches(0.65), sent, size=12, color=DARK,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.38), Inches(4.0), Inches(12.5), Inches(0.35),
       "Past Perfect = steps 1–2 (earlier)  |  Simple Past = steps 3–4 (later story moments)", size=14, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.38), Inches(4.5), Inches(12.5), Inches(1.8),
       "Real life:\n• I had studied before the quiz started.\n• She had never seen snow until last winter.\n"
       "• By the time the bus arrived, we had already left.",
       size=14, color=DARK, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Point to timeline left-to-right.",
        "Label which sentences need Past Perfect.",
        "1–2 Past Perfect; 3–4 simple past.",
        "Sequence understanding.",
        "Marking later events as past perfect.",
        "Rewrite #3 using past perfect for earlier action.",
        "Grammar block")


def s07():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Grammar Practice ✏️", "GRAMMAR", "Grammar", 7)
    items = [
        "By the time study hall began, Riley ___ (finish) the quiz.",
        "Jordan ___ (never / see) the secret door before Tuesday.",
        "After Riley ___ (open) the door, footsteps echoed.",
        "Ms. Delgado ___ (already / leave) when Riley whispered.",
        "They ___ (realize) the library was bigger than they thought.",
    ]
    for i, it in enumerate(items):
        top = Inches(1.42 + i * 1.02)
        add_round(slide, Inches(0.38), top, Inches(12.5), Inches(0.92), WHITE if i % 2 == 0 else LIGHT_TEAL)
        tb(slide, Inches(0.62), top + Inches(0.25), Inches(12.0), Inches(0.45), f"{i + 1}.  {it}", size=15, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Think aloud #1.",
        "Say answers aloud.",
        "had finished; had never seen; had opened; had already left; realized (or had realized if before another past event—accept had realized)",
        "Form accuracy.",
        "For #5 accept simple past realized as main event.",
        "Write your own sentence using 'before'.",
        "Grammar")


def s08():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Grammar Game: Find the Mistake 🕵️", "GAME", "Grammar game 5 min", 8)
    lines = [
        "Riley had went down the staircase.",
        "Jordan had never heard the humming books before.",
        "After they had opened the map, the lights flicker.",
        "By the time Riley had tugged the handle, the door creaked.",
        "Ms. Delgado had leave for a meeting.",
    ]
    for i, ln in enumerate(lines):
        top = Inches(1.42 + i * 1.02)
        add_round(slide, Inches(0.38), top, Inches(12.5), Inches(0.92), LIGHT_CORAL if i % 2 else WHITE)
        tb(slide, Inches(0.62), top + Inches(0.25), Inches(12.0), Inches(0.45), f"{i + 1}.  {ln}", size=15, color=DARK)
    tb(slide, Inches(0.38), Inches(6.35), Inches(12.5), Inches(0.35),
       "Fix it! (#2 and #4 may be correct ✅)", size=13, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Lightning detective game.",
        "Fix or say 'Correct!'",
        "1 had gone 2 OK 3 flickered 4 OK 5 had left",
        "Error detection.",
        "Changing correct sentences.",
        "Create one mistake for teacher.",
        "5 min grammar game")


def s09():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "📖 The Secret Door in the Library", "READING", "Reading 10 min", 9)
    add_round(slide, Inches(0.38), Inches(1.35), Inches(12.5), Inches(5.45), WHITE)
    add_round(slide, Inches(0.55), Inches(1.5), Inches(2.4), Inches(5.15), LIGHT_PURPLE)
    tb(slide, Inches(0.65), Inches(2.0), Inches(2.2), Inches(0.6), "🚪📚", size=36, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.65), Inches(3.0), Inches(2.2), Inches(1.8),
       "mystery\nfriendship\ncliffhanger", size=13, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(3.15), Inches(1.5), Inches(9.5), Inches(5.1), PASSAGE, size=12, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Read with suspense. Highlight past perfect forms.",
        "Listen; then student reads one paragraph.",
        "Follows plot; notices had + participle.",
        "Fluency + engagement.",
        "Losing track—pause to summarize.",
        "Predict the next chapter.",
        "Reading ~6 min")


def s10():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Reading Together 🔊", "READING", "Reading", 10)
    bullets(slide, Inches(0.5), Inches(1.45), Inches(6.0), Inches(4.8), [
        "Circle Past Perfect verbs (had + V3)",
        "Underline mystery clues",
        "Star friendship moments",
        "Question mark confusing lines",
    ], size=17, sp=14)
    add_round(slide, Inches(6.8), Inches(1.45), Inches(6.1), Inches(4.8), LIGHT_AMBER)
    tb(slide, Inches(7.05), Inches(1.65), Inches(5.6), Inches(0.35), "Tricky Words", size=16, bold=True, color=NAVY)
    bullets(slide, Inches(7.05), Inches(2.15), Inches(5.6), Inches(3.8), [
        "biography — bio-GRAH-fee",
        "parchment — PAR-chmunt",
        "creaked — KREEKT",
        "curiosity — kyoor-ee-AH-suh-tee",
        "realized — REE-uh-lyzed",
    ], size=14, sp=10)
    btn(slide)
    teacher_notes(slide,
        "Partner read or echo read.",
        "Annotate + pronounce 5 words.",
        "Find 4+ had ___ forms.",
        "Close reading habits.",
        "Missing past participles.",
        "Which clue was most important?",
        "Reading block")


def s11():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Reading Comprehension (10 Questions)", "READING", "Reading", 11)
    qs = [
        "1. MCQ: School name? A) Westbrook Middle B) Oakwood C) Riverview",
        "2. MCQ: Who followed Riley? A) Ms. Delgado B) Jordan C) A stranger",
        "3. T/F: Riley had noticed the door long ago.",
        "4. T/F: The moss on the stairs was glowing blue.",
        "5. Short: What did the note on the desk say?",
        "6. Main idea: What is this story mostly about?",
        "7. Inference: Why did Jordan say they should tell an adult?",
        "8. Vocab: 'Humming' books suggests they were… A) silent B) alive/magical C) broken",
        "9. Sequence: Order: open map → tug handle → flickering lights (fix order)",
        "10. Opinion: Would you enter the secret door? Why?",
    ]
    for i, q in enumerate(qs):
        col = i // 5
        row = i % 5
        left = Inches(0.35 + col * 6.55)
        top = Inches(1.32 + row * 1.02)
        add_round(slide, left, top, Inches(6.35), Inches(0.92), WHITE if i % 2 == 0 else LIGHT_SKY)
        tb(slide, left + Inches(0.12), top + Inches(0.18), Inches(6.1), Inches(0.65), q, size=10, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "One at a time; cite evidence.",
        "Answer all 10.",
        "1A 2B 3F 4T 5 Chapter One/Two note 6 secret library discovery 7 safety 8B 9 handle→map→lights 10 reasoned opinion",
        "Comprehension depth.",
        "Guessing.",
        "Which past perfect sentence best supports #6?",
        "Reading 10 min total")


def s12():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Vocabulary Builder (10 Words) 🌟", "VOCAB", "Vocab 5 min", 12)
    for i, (w, pron, mean, pic, sent, syn) in enumerate(VOCAB):
        col, row = i % 5, i // 5
        left = Inches(0.3 + col * 2.55)
        top = Inches(1.35 + row * 2.75)
        add_round(slide, left, top, Inches(2.4), Inches(2.55), WHITE)
        tb(slide, left + Inches(0.05), top + Inches(0.08), Inches(2.3), Inches(0.35),
           f"{pic} {w}", size=11, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.05), top + Inches(0.45), Inches(2.3), Inches(1.95),
           f"{pron}\n{mean}\nSyn: {syn}\n{sent}", size=8, color=DARK, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Teach 6–10 words with repeat-back.",
        "Say word + own sentence for 3 words.",
        "Correct usage in sentences.",
        "Pronunciation + meaning.",
        "Synonym confusion.",
        "Which word describes Riley's feeling?",
        "5 min vocab")


def s13():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Vocabulary Matching 🎯 (Word Detective)", "GAME", "Vocab", 13)
    add_round(slide, Inches(0.38), Inches(1.4), Inches(5.9), Inches(5.2), WHITE)
    tb(slide, Inches(0.55), Inches(1.55), Inches(5.5), Inches(0.35), "Words", size=14, bold=True, color=TEAL)
    bullets(slide, Inches(0.55), Inches(2.0), Inches(5.5), Inches(4.3),
            [f"{i + 1}. {w}" for i, (w, *_) in enumerate(VOCAB[:6])], size=14, sp=8)
    add_round(slide, Inches(6.55), Inches(1.4), Inches(6.35), Inches(5.2), LIGHT_AMBER)
    tb(slide, Inches(6.75), Inches(1.55), Inches(6.0), Inches(0.35), "Meanings (shuffled)", size=14, bold=True, color=AMBER)
    bullets(slide, Inches(6.75), Inches(2.0), Inches(6.0), Inches(4.3), [
        "A. bounced-back sound",
        "B. wish to know",
        "C. life-story book",
        "D. flashed on/off",
        "E. old thick paper",
        "F. exciting experience",
    ], size=14, sp=8)
    btn(slide)
    teacher_notes(slide,
        "Match 1–6 to letters.",
        "Say matches aloud.",
        "1-C biography  2-E parchment  3-A echoed/creaked  4-B curiosity  5-D flickered  6-F adventure",
        "Matching skill.",
        "Shuffled confusion.",
        "Match remaining 4 words from slide 12.",
        "Vocab game")


def s14():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Word Search + Unscramble 🔤", "GAME", "Vocab", 14)
    grid = "SECRETXX\nDOORXXXX\nLIBRARYX\nMOSSXXXX\nMAPXXXXX\nECHOXXXX"
    add_round(slide, Inches(0.38), Inches(1.4), Inches(7.2), Inches(4.5), WHITE)
    tb(slide, Inches(0.55), Inches(1.8), Inches(6.8), Inches(3.5), grid, size=24, bold=True, color=NAVY,
       font="Consolas", align=PP_ALIGN.CENTER)
    add_round(slide, Inches(7.85), Inches(1.4), Inches(5.05), Inches(4.5), LIGHT_TEAL)
    tb(slide, Inches(8.05), Inches(1.6), Inches(4.7), Inches(0.35), "Unscramble", size=14, bold=True, color=NAVY)
    bullets(slide, Inches(8.05), Inches(2.05), Inches(4.7), Inches(3.5), [
        "dekaerc → ?",
        "yrotarobbi → ?",
        "ytisoiruc → ?",
        "dezilera → ?",
    ], size=14, sp=10)
    tb(slide, Inches(8.05), Inches(4.35), Inches(4.7), Inches(1.0),
       "Key: creaked, biography,\ncuriosity, realized", size=12, color=TEAL)
    btn(slide)
    teacher_notes(slide,
        "3 min word search; 2 min unscramble.",
        "Find SECRET, DOOR, LIBRARY, MOSS, MAP, ECHO; unscramble words.",
        "See slide key.",
        "Word recognition.",
        "Spelling.",
        "Use ECHO in past perfect sentence.",
        "Vocab games")


def s15():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Speaking Activities 🎤", "SPEAKING", "Storytelling 10 min", 15)
    acts = [
        ("Describe the Picture", "Mystery library door scene"),
        ("Finish the Story", "Continue after cliffhanger"),
        ("If I Found a Secret Door…", "What would you do first?"),
        ("Interview the Character", "Ask Riley 3 questions"),
    ]
    for i, (t, d) in enumerate(acts):
        left = Inches(0.38 + (i % 2) * 6.45)
        top = Inches(1.45 + (i // 2) * 2.45)
        add_round(slide, left, top, Inches(6.2), Inches(2.25), WHITE)
        add_rect(slide, left, top, Inches(6.2), Inches(0.5), [SKY, TEAL, CORAL, PURPLE][i])
        tb(slide, left + Inches(0.2), top + Inches(0.08), Inches(5.8), Inches(0.35), t, size=15, bold=True, color=WHITE)
        tb(slide, left + Inches(0.2), top + Inches(0.75), Inches(5.8), Inches(1.2), d + "\n(45–60 sec each)", size=14, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Student picks 2 activities.",
        "Extended speaking with past perfect where possible.",
        "Clear ideas; uses story vocabulary.",
        "Confidence + fluency.",
        "Short answers—prompt 'tell me more'.",
        "Use one past perfect sentence in interview.",
        "Speaking 10 min block")


def s16():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Storytelling Picture Prompt 🖼️", "STORY", "Storytelling", 16)
    add_round(slide, Inches(0.38), Inches(1.4), Inches(7.5), Inches(5.2), LIGHT_SKY)
    tb(slide, Inches(0.55), Inches(2.0), Inches(7.2), Inches(3.5),
       "🏫 Old library\n🚪 Hidden door\n🌟 Glowing stairs\n👫 Two friends\n❓ Unfinished map",
       size=22, color=NAVY, align=PP_ALIGN.CENTER)
    guides = [
        "Who is the main character?",
        "Where does the story happen?",
        "What problem occurs?",
        "How is the problem solved?",
        "What lesson is learned?",
    ]
    for i, g in enumerate(guides):
        top = Inches(1.45 + i * 1.0)
        add_round(slide, Inches(8.15), top, Inches(4.75), Inches(0.88), WHITE)
        tb(slide, Inches(8.35), top + Inches(0.25), Inches(4.35), Inches(0.45), g, size=14, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Plan 60 sec; then tell story.",
        "5-part story from prompt.",
        "Beginning/middle/end + lesson.",
        "Creative organization.",
        "Skipping problem/solution.",
        "Add a cliffhanger ending.",
        "Storytelling")


def s17():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Creative Story Challenge 🎲 Roll the Dice", "GAME", "Storytelling", 17)
    tb(slide, Inches(0.38), Inches(1.4), Inches(12.5), Inches(0.35),
       "Roll 1–6 for each column → build a paragraph using Past Perfect at least twice.", size=14, color=SOFT)
    cols = [
        ("Character", "Riley", "Jordan", "Librarian", "Stranger", "You", "Detective"),
        ("Setting", "Library", "Stairwell", "Map room", "Hallway", "Roof", "Garden"),
        ("Problem", "Lost book", "Locked door", "Strange sound", "Missing key", "Darkness", "Note clue"),
    ]
    for j, (title, *opts) in enumerate(cols):
        left = Inches(0.38 + j * 4.25)
        add_round(slide, left, Inches(1.85), Inches(4.05), Inches(4.5), WHITE)
        tb(slide, left, Inches(2.0), Inches(4.05), Inches(0.35), title, size=15, bold=True, color=TEAL,
           align=PP_ALIGN.CENTER)
        for k, o in enumerate(opts):
            tb(slide, left + Inches(0.2), Inches(2.45 + k * 0.55), Inches(3.65), Inches(0.45),
               f"{k + 1}. {o}", size=13, color=DARK, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Roll dice or pick numbers.",
        "Write/say 5–7 sentence story.",
        "Includes 2+ past perfect verbs.",
        "Creative writing + grammar.",
        "Only simple past.",
        "Peer question: What happened first?",
        "Storytelling 10 min")


def s18():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Grammar Quiz ⚡ Lightning Round", "QUIZ", "Games", 18)
    qs = [
        "Past Perfect form? had + ___",
        "Signal word? before / quickly / very",
        "Fix: She had ate lunch.",
        "Earlier event? I had finished / I finished first?",
        "Choose: By the time he arrived, we (leave / had left).",
    ]
    for i, q in enumerate(qs):
        top = Inches(1.42 + i * 1.02)
        add_round(slide, Inches(0.38), top, Inches(12.5), Inches(0.92), [SKY, TEAL, CORAL, PURPLE, AMBER][i])
        tb(slide, Inches(0.62), top + Inches(0.25), Inches(12.0), Inches(0.45), f"Q{i + 1}: {q}", size=15, bold=True,
           color=WHITE)
    btn(slide)
    teacher_notes(slide,
        "Fast quiz—celebrate streaks.",
        "Verbal answers.",
        "participle; before; had eaten; had finished; had left",
        "Quick recall.",
        "Panicking—slow down.",
        "Make Q6.",
        "Games")


def s19():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Escape Room Challenge 🔐", "GAME", "Games", 19)
    puzzles = [
        "Clue 1 Unscramble: tsap erfectf → past perfect",
        "Clue 2 Fix: had went → had gone",
        "Clue 3 Vocab: synonym for adventure → quest/journey",
        "Clue 4 Past Perfect: I ___ (finish) before the bell.",
        "Clue 5 Reading: Who is Riley's friend?",
    ]
    for i, p in enumerate(puzzles):
        top = Inches(1.42 + i * 1.02)
        add_round(slide, Inches(0.38), top, Inches(12.5), Inches(0.92), WHITE if i % 2 == 0 else LIGHT_PURPLE)
        tb(slide, Inches(0.62), top + Inches(0.25), Inches(12.0), Inches(0.45), p, size=14, color=DARK)
    tb(slide, Inches(0.38), Inches(6.35), Inches(12.5), Inches(0.35),
       "Unlock the treasure when all 5 are solved! 🎁", size=14, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Timer 5 min escape.",
        "Solve all clues aloud.",
        "Clue 1 past perfect; 2 gone; 3 quest; 4 had finished; 5 Jordan",
        "Integrated skills.",
        "Stuck—offer hint.",
        "Design a Clue 6.",
        "Games / recap transition")


def s20():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Recap: You Teach the Teacher 🎓", "RECAP", "Recap 5 min", 20)
    bullets(slide, Inches(0.45), Inches(1.45), Inches(12.3), Inches(4.8), [
        "Explain Past Perfect in one sentence",
        "Give the structure (had + ___)",
        "Name 2 signal words",
        "Summarize The Secret Door in 2 sentences",
        "Share your favorite game today",
        "What did you learn about yourself as a reader/writer?",
    ], size=17, sp=12)
    btn(slide)
    teacher_notes(slide,
        "Student-led recap 2 minutes.",
        "Teach back key ideas.",
        "Accurate mini-lesson.",
        "Metacognition + retention.",
        "Minimal response—prompt examples.",
        "One goal for next week?",
        "5 min recap")


def s21():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    header(slide, "Homework Mission 📝", "HOMEWORK", "After class", 21)
    add_round(slide, Inches(1.0), Inches(1.45), Inches(11.3), Inches(5.0), WHITE)
    tb(slide, Inches(1.25), Inches(1.7), Inches(10.8), Inches(0.45),
       "Write a short story (120+ words)", size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    bullets(slide, Inches(1.25), Inches(2.35), Inches(10.8), Inches(3.8), [
        "Include at least 5 Past Perfect sentences",
        "Use at least 5 vocabulary words from today",
        "Include beginning, middle, ending",
        "Optional: draw your secret door",
        "Read aloud to someone at home",
    ], size=16, sp=12)
    btn(slide)
    teacher_notes(slide,
        "Explain clearly; show example starter.",
        "Plan 3 bullet points before writing.",
        "Meets length + grammar + vocab requirements.",
        "Transfer to independent work.",
        "All simple past.",
        "Email/share next class optional.",
        "Homework")


def s22():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), AMBER)
    add_round(slide, Inches(2.0), Inches(1.15), Inches(9.3), Inches(5.2), WHITE)
    tb(slide, Inches(2.3), Inches(1.7), Inches(8.7), Inches(0.5),
       "🎉 Excellent Work! 🎉", size=34, bold=True, color=NAVY, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(2.3), Inches(2.6), Inches(8.7), Inches(1.5),
       "You survived the grammar puzzles,\nsolved the library mystery,\nand leveled up your English adventure skills!",
       size=18, color=DARK, align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.3), Inches(4.4), Inches(8.7), Inches(0.45),
       "Grade 7 English Adventure — Complete", size=15, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.3), Inches(5.0), Inches(8.7), Inches(0.4),
       "See you on the next quest! 📚🗝️", size=16, color=CORAL, align=PP_ALIGN.CENTER)
    footer(slide, 22, "Celebrate!")
    fade(slide)
    teacher_notes(slide,
        "Specific praise (2 strengths).",
        "Student shares proudest moment.",
        "Positive closure.",
        "End energy high for parents.",
        "Generic praise only.",
        "Preview next topic.",
        "End")


for fn in [s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20, s21, s22]:
    fn()

out = r"C:\Users\bhushaja\Downloads\shaip\Grade7_Past_Perfect_English_Adventure_60min.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
