"""Grade 5 English Demo Class - 60 minutes, 37 slides.

Reading + Tenses + Vocabulary + Speaking + Storytelling + Assessment.

SOURCE NOTE: No source PDF was available when this deck was generated, so the
reading passage and all exercises below are ORIGINAL teacher-created material.
Swap PASSAGE_* and VOCAB if source material is supplied later.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

NAVY = RGBColor(0x0A, 0x2A, 0x6E)
SKY = RGBColor(0x29, 0xB6, 0xF6)
TEAL = RGBColor(0x00, 0x96, 0x88)
YELLOW = RGBColor(0xFF, 0xD5, 0x4F)
GOLD = RGBColor(0xFF, 0xB3, 0x00)
CORAL = RGBColor(0xFF, 0x6F, 0x61)
PINK = RGBColor(0xF0, 0x62, 0x92)
PURPLE = RGBColor(0x7E, 0x57, 0xC2)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
CREAM = RGBColor(0xFF, 0xFB, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x29, 0x37)
SOFT = RGBColor(0x64, 0x74, 0x8B)
LIGHT_SKY = RGBColor(0xE1, 0xF5, 0xFE)
LIGHT_TEAL = RGBColor(0xE0, 0xF7, 0xFA)
LIGHT_YELLOW = RGBColor(0xFF, 0xF9, 0xC4)
LIGHT_CORAL = RGBColor(0xFF, 0xEB, 0xEE)
LIGHT_PURPLE = RGBColor(0xF3, 0xE5, 0xF5)
LIGHT_GREEN = RGBColor(0xE8, 0xF5, 0xE9)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 37
STUDENT = "Student"
TEACHER_NAME = "Your Teacher Name"

_counter = {"n": 0}

# ---------------------------------------------------------------- content

PASSAGE_1 = (
    "Maya lived in a small beach town in Florida. Every August, something amazing "
    "happened there. Baby sea turtles hatched in the warm sand and raced toward the ocean.\n\n"
    "One quiet night, Maya and her older brother Diego walked along the shore with a red "
    "flashlight. Volunteers always used red light, because bright white light confuses "
    "baby turtles and sends them the wrong way.\n\n"
    "\"Look!\" Diego whispered. The sand above a marked nest was moving. Tiny flippers "
    "pushed up through the surface."
)

PASSAGE_2 = (
    "Suddenly, Maya noticed a problem. A deep tire track crossed the beach like a canyon. "
    "Three hatchlings had tumbled in and could not climb out. Their flippers kept slipping "
    "on the loose sand.\n\n"
    "Maya remembered her training. Volunteers were never allowed to pick up the turtles and "
    "carry them to the water. So instead, she knelt down and gently smoothed the sand into "
    "a small ramp.\n\n"
    "\"Come on, little ones,\" she whispered. \"You can do it.\"\n\n"
    "One by one, the hatchlings scrambled up the ramp and continued their journey."
)

PASSAGE_3 = (
    "By midnight, ninety-two turtles had reached the waves. Maya was exhausted, but she "
    "could not stop smiling.\n\n"
    "Dr. Reyes, the biologist, wrote the number in her notebook. \"Only about one in a "
    "thousand will survive to become an adult,\" she said. \"But tonight, you gave these "
    "ninety-two a real chance.\"\n\n"
    "Maya looked out at the dark water. Somewhere out there, a tiny turtle was swimming "
    "toward a life she would never see.\n\n"
    "\"Next August,\" Maya said, \"I'm coming back.\"\n\n"
    "Diego grinned. \"I knew you'd say that.\""
)

VOCAB = [
    ("🥚", "hatched", "/HACHT/", "came out of an egg", "emerged"),
    ("🙋", "volunteer", "/VOL-un-teer/", "someone who helps without being paid", "helper"),
    ("😕", "confuse", "/kun-FYOOZ/", "to mix up someone's thinking", "puzzle"),
    ("🐢", "hatchling", "/HACH-ling/", "a baby animal just out of its egg", "newborn"),
    ("🏃", "scramble", "/SKRAM-bul/", "to move fast using hands and feet", "clamber"),
    ("😴", "exhausted", "/eg-ZAWS-ted/", "very, very tired", "worn out"),
    ("🔬", "biologist", "/by-OL-uh-jist/", "a scientist who studies living things", "scientist"),
    ("💪", "survive", "/sur-VYVE/", "to stay alive", "live on"),
    ("🗺️", "journey", "/JUR-nee/", "a long trip from one place to another", "voyage"),
    ("🤲", "gently", "/JENT-lee/", "softly and carefully", "carefully"),
    ("🏔️", "canyon", "/KAN-yun/", "a deep, narrow valley", "gorge"),
    ("🚩", "marked", "/MARKT/", "showed with a sign so people know", "labeled"),
]

TENSES = [
    dict(
        name="Simple Present", icon="🔁", color=TEAL, light=LIGHT_TEAL, mark="NOW",
        when="Things that happen again and again, and facts that are always true.",
        structure="Subject  +  verb  (+ s for he / she / it)",
        pos="Maya helps the turtles every August.",
        neg="Maya does not help on rainy nights.",
        que="Does Maya help the turtles?",
        signals="every day  •  always  •  usually  •  often  •  never  •  on Mondays",
        real="I brush my teeth every morning.",
    ),
    dict(
        name="Present Continuous", icon="⏳", color=SKY, light=LIGHT_SKY, mark="NOW",
        when="Something that is happening right now, while we speak.",
        structure="Subject  +  am / is / are  +  verb-ing",
        pos="The hatchlings are crawling to the sea.",
        neg="They are not stopping to rest.",
        que="Are the hatchlings crawling to the sea?",
        signals="now  •  right now  •  at the moment  •  Look!  •  Listen!",
        real="I am learning English right now.",
    ),
    dict(
        name="Simple Past", icon="⏪", color=CORAL, light=LIGHT_CORAL, mark="PAST",
        when="Something that started AND finished in the past.",
        structure="Subject  +  past verb  (walked, went, saw)",
        pos="Maya built a sand ramp last night.",
        neg="Maya did not carry the turtles.",
        que="Did Maya build a sand ramp?",
        signals="yesterday  •  last night  •  ago  •  in 2019  •  when I was little",
        real="I watched a movie last night.",
    ),
    dict(
        name="Past Continuous", icon="🎬", color=PURPLE, light=LIGHT_PURPLE, mark="PAST",
        when="Something that was already happening at a moment in the past.",
        structure="Subject  +  was / were  +  verb-ing",
        pos="At midnight, Maya was counting turtles.",
        neg="Diego was not sleeping.",
        que="Was Maya counting turtles at midnight?",
        signals="while  •  when  •  at 8 o'clock last night  •  all evening",
        real="I was sleeping when the phone rang.",
    ),
    dict(
        name="Simple Future", icon="🚀", color=GOLD, light=LIGHT_YELLOW, mark="FUTURE",
        when="Something that has not happened yet, but will happen later.",
        structure="Subject  +  will  +  verb",
        pos="Maya will come back next August.",
        neg="She will not forget the turtles.",
        que="Will Maya come back next August?",
        signals="tomorrow  •  next week  •  soon  •  later  •  in 2030",
        real="I will visit my grandma next weekend.",
    ),
]

# ---------------------------------------------------------------- helpers


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
    tree = slide.shapes._spTree
    el = s._element
    tree.remove(el)
    tree.insert(2, el)


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


def tb(slide, l, t, w, h, text, size=18, bold=False, color=DARK,
       align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, color, font)
    return box


def bullets(slide, l, t, w, h, items, size=15, color=DARK, sp=8, bullet="•  "):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(sp)
        r = p.add_run()
        r.text = bullet + item
        set_run(r, size, False, color)
    return box


def fade(slide):
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    etree.SubElement(tr, qn("p:fade"))


def progress(slide, n):
    """Thin progress bar just above the footer."""
    add_rect(slide, Inches(0), Inches(7.04), prs.slide_width, Inches(0.08), RGBColor(0xDD, 0xE3, 0xEA))
    width = prs.slide_width * n / TOTAL
    add_rect(slide, Inches(0), Inches(7.04), int(width), Inches(0.08), GOLD)


def footer(slide, n, timing="", step=""):
    progress(slide, n)
    add_rect(slide, Inches(0), Inches(7.12), prs.slide_width, Inches(0.38), NAVY)
    msg = f"Grade 5 English Adventure  |  60 min"
    if step:
        msg += f"  |  {step}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.15), Inches(11.4), Inches(0.3), msg, size=10, color=WHITE)
    tb(slide, Inches(11.9), Inches(7.15), Inches(1.15), Inches(0.3), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def chip(slide, text, color=SKY, w=Inches(2.9)):
    add_round(slide, Inches(0.38), Inches(0.26), w, Inches(0.38), color)
    tb(slide, Inches(0.38), Inches(0.29), w, Inches(0.34), text, size=11, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def time_chip(slide, text):
    add_round(slide, Inches(10.45), Inches(0.26), Inches(2.5), Inches(0.38), CORAL)
    tb(slide, Inches(10.45), Inches(0.29), Inches(2.5), Inches(0.34), text, size=11, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def star_badge(slide, text="+1 Star ⭐"):
    add_round(slide, Inches(10.55), Inches(6.45), Inches(2.4), Inches(0.45), YELLOW)
    tb(slide, Inches(10.55), Inches(6.5), Inches(2.4), Inches(0.4), text, size=13, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER)


def new_slide(title, tag="ACTIVITY", timing="", step="", accent=SKY, bg=CREAM):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.14), prs.slide_height, accent)
    chip(slide, tag, accent)
    if timing:
        time_chip(slide, timing)
    tb(slide, Inches(0.38), Inches(0.76), Inches(12.4), Inches(0.55), title, size=26, bold=True,
       color=NAVY, font="Georgia")
    footer(slide, n, timing, step)
    fade(slide)
    return slide, n


def notes(slide, say, activity, expected, mistakes, support, bonus, praise, timing):
    tf = slide.notes_slide.notes_text_frame
    tf.text = (
        f"⏱ TIMING: {timing}\n\n"
        f"SAY: {say}\n\n"
        f"HOW TO RUN IT: {activity}\n\n"
        f"EXPECTED ANSWERS: {expected}\n\n"
        f"COMMON MISTAKES: {mistakes}\n\n"
        f"IF THE STUDENT STRUGGLES (easier step): {support}\n\n"
        f"IF THE STUDENT FINISHES FAST (bonus): {bonus}\n\n"
        f"ENCOURAGE: {praise}"
    )


def timeline_strip(slide, top, highlight):
    """PAST ---- NOW ---- FUTURE strip. highlight in {PAST, NOW, FUTURE}."""
    add_rect(slide, Inches(1.3), top + Inches(0.29), Inches(10.6), Inches(0.05), SOFT)
    for label, x in [("PAST", 1.55), ("NOW", 6.6), ("FUTURE", 11.55)]:
        on = label == highlight
        dot = CORAL if on else WHITE
        size = 0.66 if on else 0.5
        off = (0.66 - size) / 2 + 0.0
        add_oval(slide, Inches(x - size / 2), top + Inches(0.31 - size / 2 + off),
                 Inches(size), Inches(size), dot)
        tb(slide, Inches(x - 0.85), top + Inches(0.72), Inches(1.7), Inches(0.3), label,
           size=12, bold=True, color=NAVY if on else SOFT, align=PP_ALIGN.CENTER)


def qa_card(slide, left, top, w, h, tagtext, tagcolor, question, hint=""):
    add_round(slide, left, top, w, h, WHITE)
    add_round(slide, left + Inches(0.15), top + Inches(0.14), Inches(1.5), Inches(0.32), tagcolor)
    tb(slide, left + Inches(0.15), top + Inches(0.16), Inches(1.5), Inches(0.3), tagtext,
       size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, left + Inches(0.15), top + Inches(0.58), w - Inches(0.3), h - Inches(0.9),
       question, size=14, color=DARK)
    if hint:
        tb(slide, left + Inches(0.15), top + h - Inches(0.42), w - Inches(0.3), Inches(0.32),
           hint, size=11, bold=True, color=TEAL)


# ---------------------------------------------------------------- slides


def s01_welcome():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.24), YELLOW)
    for x, y, c in [(0.6, 0.6, SKY), (12.0, 0.65, PINK), (0.75, 5.95, TEAL), (11.95, 5.85, GOLD)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(1.85), Inches(12), Inches(0.75),
       "🌟 Welcome to Your English Adventure! 🌟", size=38, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.8), Inches(12), Inches(0.45),
       "Grade 5  •  United States  •  60-Minute Live Demo Class",
       size=18, color=LIGHT_SKY, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.3), Inches(3.6), Inches(6.7), Inches(1.35), TEAL)
    tb(slide, Inches(3.5), Inches(3.85), Inches(6.3), Inches(0.9),
       "Today we PLAY, READ, and become\nTENSE DETECTIVES 🕵️\n(This is NOT a test!)",
       size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(5.35), Inches(12), Inches(0.45),
       "📖  🎮  🗣️  🧩  ⏰  🎯  —  Ready? Let's go!",
       size=22, color=YELLOW, align=PP_ALIGN.CENTER)
    footer(slide, n, "0–5 min", "Welcome")
    fade(slide)
    notes(slide,
          "Hi! I am so happy to meet you. Today we are going on an English adventure. "
          "We will read a true-to-life turtle rescue story, play detective games, and you will "
          "earn stars along the way. There is no test today — only fun.",
          "Smile, wave, say your name. Ask the student to say their name and one thing they like. "
          "Keep this under 60 seconds so the energy stays high.",
          "Student says their name and one interest (soccer, drawing, video games, animals).",
          "Very short one-word answers; nervous silence at the start.",
          "Give a choice instead of an open question: 'Do you like animals or sports more?'",
          "Ask: 'Can you tell me ONE reason you like it?'",
          "You're already speaking English — great start! ⭐",
          "0–5 min")


def s02_mission():
    slide, n = new_slide("🎯 Today's Mission", "AGENDA", "0–5 min", "Mission", GOLD)
    steps = [
        ("🎮", "Vocabulary Games", "5–10 min", TEAL),
        ("📖", "Reading Adventure", "10–20 min", SKY),
        ("🧩", "Reading Games", "20–25 min", PURPLE),
        ("⏰", "Tenses Lesson", "25–35 min", CORAL),
        ("🕵️", "Tense Detective", "35–47 min", PINK),
        ("🗣️", "Story + Speaking", "47–55 min", GREEN),
        ("⚡", "Quick Challenge", "55–58 min", GOLD),
        ("⭐", "Recap + Reward", "58–60 min", YELLOW),
    ]
    for i, (icon, label, when, color) in enumerate(steps):
        col, row = i % 4, i // 4
        left = Inches(0.42 + col * 3.18)
        top = Inches(1.55 + row * 2.45)
        add_round(slide, left, top, Inches(3.0), Inches(2.2), WHITE)
        add_oval(slide, left + Inches(1.15), top + Inches(0.2), Inches(0.7), Inches(0.7), color)
        tb(slide, left + Inches(1.15), top + Inches(0.32), Inches(0.7), Inches(0.5), icon,
           size=22, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.12), top + Inches(1.05), Inches(2.76), Inches(0.7), label,
           size=15, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.12), top + Inches(1.68), Inches(2.76), Inches(0.35), when,
           size=12, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.42), Inches(6.5), Inches(9.5), Inches(0.4),
       "Collect ⭐ stars at every stop. Can you get all 8?",
       size=15, bold=True, color=CORAL)
    notes(slide,
          "Here is our mission map. Eight stops, and you can earn a star at every single one. "
          "Nothing lasts long — if something feels tricky, we move on quickly.",
          "Point to each stop as you read it. Ask the student which stop looks most exciting. "
          "Keep it to about 45 seconds.",
          "Student picks a stop, usually the games or storytelling.",
          "Student may worry about the word 'challenge' or 'quiz'.",
          "Reassure: 'The quiz is just a game — I help you the whole time.'",
          "Ask them to predict how many stars they will collect today.",
          "Great choice! That one is really fun. ⭐",
          "0–5 min")


def s03_icebreaker():
    slide, n = new_slide("😄 Icebreaker: Would You Rather?", "GAME", "0–5 min", "Icebreaker", PINK)
    pairs = [
        ("Swim with sea turtles", "Fly with eagles"),
        ("Explore a dark cave", "Explore outer space"),
        ("Have a pet dragon", "Have a robot helper"),
        ("Be super fast", "Be super strong"),
        ("Always summer", "Always snow"),
        ("Talk to animals", "Speak every language"),
    ]
    for i, (a, b) in enumerate(pairs):
        col, row = i % 3, i // 3
        left = Inches(0.45 + col * 4.2)
        top = Inches(1.55 + row * 2.5)
        add_round(slide, left, top, Inches(3.95), Inches(2.25), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(0.2), Inches(3.65), Inches(0.62), a,
           size=14, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.93), Inches(3.65), Inches(0.34), "— OR —",
           size=12, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.37), Inches(3.65), Inches(0.62), b,
           size=14, bold=True, color=CORAL, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.45), Inches(6.4), Inches(9.4), Inches(0.5), LIGHT_YELLOW)
    tb(slide, Inches(0.6), Inches(6.48), Inches(9.1), Inches(0.36),
       "💬 Say the full sentence:  \"I would rather ______ because ______.\"",
       size=14, bold=True, color=NAVY)
    star_badge(slide)
    notes(slide,
          "Let's warm up our talking muscles. I will read two choices and you pick one — "
          "but you have to tell me WHY. I'll go first: I would rather swim with sea turtles "
          "because I love the ocean.",
          "Model the first one yourself, then do 3–4 pairs with the student. Do not do all six "
          "unless the student is enjoying it. Stop while it is still fun.",
          "Full sentence with 'because' — e.g. 'I would rather fly with eagles because I want to "
          "see the mountains.'",
          "Answering with just one word ('dragon') and leaving out 'because'.",
          "Give a sentence frame out loud and let them finish it: 'I would rather ___ because ___.'",
          "Ask a follow-up: 'What is one BAD thing about your choice?'",
          "I love that reason — full sentence too! +1 Star ⭐",
          "0–5 min")


def _vocab_card(slide, left, top, emoji, word, pron, meaning, syn):
    w, h = Inches(4.0), Inches(2.05)
    add_round(slide, left, top, w, h, WHITE)
    add_oval(slide, left + Inches(0.16), top + Inches(0.18), Inches(0.82), Inches(0.82), LIGHT_SKY)
    tb(slide, left + Inches(0.16), top + Inches(0.34), Inches(0.82), Inches(0.55), emoji,
       size=24, align=PP_ALIGN.CENTER)
    tb(slide, left + Inches(1.1), top + Inches(0.18), w - Inches(1.25), Inches(0.4), word,
       size=18, bold=True, color=NAVY)
    tb(slide, left + Inches(1.1), top + Inches(0.58), w - Inches(1.25), Inches(0.32), pron,
       size=11, color=SOFT)
    tb(slide, left + Inches(0.16), top + Inches(1.08), w - Inches(0.32), Inches(0.58), meaning,
       size=13, color=DARK)
    tb(slide, left + Inches(0.16), top + Inches(1.66), w - Inches(0.32), Inches(0.3),
       f"≈ same as: {syn}", size=11, bold=True, color=TEAL)


def s04_vocab_a():
    slide, n = new_slide("📚 Vocabulary Warm-Up — Part 1", "VOCAB", "5–10 min", "Vocab", TEAL)
    for i, (emoji, word, pron, meaning, syn) in enumerate(VOCAB[:6]):
        col, row = i % 3, i // 3
        _vocab_card(slide, Inches(0.42 + col * 4.2), Inches(1.5 + row * 2.28),
                    emoji, word, pron, meaning, syn)
    tb(slide, Inches(0.42), Inches(6.35), Inches(9.6), Inches(0.4),
       "🔊 Say each word after me — then use ONE of them in your own sentence.",
       size=14, bold=True, color=CORAL)
    notes(slide,
          "These six words come from our story. I will say each word, you say it back, "
          "then I'll tell you what it means. Repeat after me: hatched... volunteer...",
          "Choral repeat each word once. Read the meaning, not the whole card. Spend about "
          "20 seconds per word. Then ask for one student sentence using any word.",
          "Clear repetition; one original sentence such as 'The bird hatched from the egg.'",
          "Stress on the wrong syllable in 'biologist' and 'exhausted'; confusing "
          "'hatched' with 'hatchling'.",
          "Break the word into parts and clap the syllables: ex-HAUS-ted.",
          "Ask them to use TWO words in one sentence.",
          "Your pronunciation is getting better every time! ⭐",
          "5–10 min")


def s05_vocab_b():
    slide, n = new_slide("📚 Vocabulary Warm-Up — Part 2", "VOCAB", "5–10 min", "Vocab", TEAL)
    for i, (emoji, word, pron, meaning, syn) in enumerate(VOCAB[6:]):
        col, row = i % 3, i // 3
        _vocab_card(slide, Inches(0.42 + col * 4.2), Inches(1.5 + row * 2.28),
                    emoji, word, pron, meaning, syn)
    tb(slide, Inches(0.42), Inches(6.35), Inches(9.6), Inches(0.4),
       "🔊 Which word is your favorite? Say it three times, loud and proud!",
       size=14, bold=True, color=CORAL)
    notes(slide,
          "Six more words. These are the tricky ones — but you are going to own them. "
          "Repeat after me: biologist... survive... journey...",
          "Same routine as the previous slide. Keep it brisk. If the student is fading, "
          "do only four words and move to the game.",
          "Accurate repetition; student picks a favorite word and says it three times.",
          "'Journey' pronounced with a hard J-O sound; 'gently' said as 'gentle'.",
          "Say the word slowly in syllables and let them copy one syllable at a time.",
          "Ask: 'Which word could describe YOU today?'",
          "Excellent! You said the hardest word on the page. ⭐",
          "5–10 min")


def s06_guess_word():
    slide, n = new_slide("🎮 Game 1: Guess the Word!", "GAME", "5–10 min", "Vocab game", PURPLE)
    clues = [
        ("🥚➡️🐢", "It just came out of an egg.", "hatchling"),
        ("😴💤", "You feel this after a very long day.", "exhausted"),
        ("🙋❤️", "This person helps but gets no money.", "volunteer"),
        ("🗺️➡️🌊", "A long trip to somewhere far away.", "journey"),
    ]
    for i, (pic, clue, _ans) in enumerate(clues):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.3)
        top = Inches(1.5 + row * 2.5)
        add_round(slide, left, top, Inches(6.0), Inches(2.25), WHITE)
        add_round(slide, left + Inches(0.18), top + Inches(0.2), Inches(2.2), Inches(1.85), LIGHT_PURPLE)
        tb(slide, left + Inches(0.18), top + Inches(0.72), Inches(2.2), Inches(0.7), pic,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(2.55), top + Inches(0.42), Inches(3.3), Inches(0.9), clue,
           size=15, bold=True, color=NAVY)
        tb(slide, left + Inches(2.55), top + Inches(1.45), Inches(3.3), Inches(0.45),
           "Your word:  _____________", size=13, color=SOFT)
    star_badge(slide, "+1 Star each ⭐")
    notes(slide,
          "Detective time! I'll show you a picture clue and a hint. You guess the word. "
          "You get a star for every word you catch.",
          "Reveal one card at a time by pointing. Give 5–8 seconds of thinking time before "
          "helping. Answers: hatchling, exhausted, volunteer, journey.",
          "hatchling  •  exhausted  •  volunteer  •  journey",
          "Saying 'hatched' instead of 'hatchling'; saying 'tired' instead of 'exhausted'.",
          "Give the first sound: 'It starts with /h/...' then the first two letters.",
          "Ask them to make a sentence with the word they just guessed.",
          "You got it! Real detective thinking. +1 Star ⭐",
          "5–10 min")


def s07_word_detective():
    slide, n = new_slide("🕵️ Game 2: Word Detective", "GAME", "5–10 min", "Vocab game", PURPLE)
    rounds = [
        ("Which word means VERY TIRED?", ["excited", "exhausted", "expensive"], "exhausted"),
        ("Which word means A LONG TRIP?", ["journey", "jungle", "juggle"], "journey"),
        ("Which word means A DEEP VALLEY?", ["cannon", "canyon", "candle"], "canyon"),
    ]
    for i, (q, opts, _a) in enumerate(rounds):
        top = Inches(1.5 + i * 1.72)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(1.5), WHITE)
        tb(slide, Inches(0.7), top + Inches(0.22), Inches(4.6), Inches(1.0), q,
           size=15, bold=True, color=NAVY)
        for j, opt in enumerate(opts):
            oleft = Inches(5.5 + j * 2.45)
            add_round(slide, oleft, top + Inches(0.42), Inches(2.25), Inches(0.66), LIGHT_TEAL)
            tb(slide, oleft, top + Inches(0.56), Inches(2.25), Inches(0.4), opt,
               size=14, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.45), Inches(6.72), Inches(9.5), Inches(0.32),
       "🔍 Careful — the wrong answers LOOK similar. Read all three first!",
       size=13, bold=True, color=CORAL)
    notes(slide,
          "The imposters are here! Each row has one real word and two sneaky look-alikes. "
          "Read all three before you choose.",
          "Read the question aloud, then read the three options slowly. Let the student point "
          "or say the answer. Answers: exhausted, journey, canyon.",
          "exhausted  •  journey  •  canyon",
          "Choosing by first letter only — picking 'excited' or 'cannon'.",
          "Cover one wrong option with your hand to make it a 50/50 choice.",
          "Ask what the two imposter words actually mean.",
          "Sharp eyes! You didn't fall for the trick. ⭐",
          "5–10 min")


def s08_reading_intro():
    slide, n = new_slide("📖 Reading Adventure: Let's Predict!", "READING", "10–20 min",
                         "Reading", SKY)
    add_round(slide, Inches(0.45), Inches(1.45), Inches(6.1), Inches(5.2), LIGHT_SKY)
    tb(slide, Inches(0.75), Inches(1.75), Inches(5.5), Inches(0.45), "Our Story Today",
       size=18, bold=True, color=NAVY)
    tb(slide, Inches(0.75), Inches(2.3), Inches(5.5), Inches(0.9),
       "\"The Midnight Turtle Rescue\"", size=25, bold=True, color=TEAL, font="Georgia")
    tb(slide, Inches(0.75), Inches(3.4), Inches(5.5), Inches(1.2), "🐢  🌊  🔦  🌙",
       size=42, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.75), Inches(5.0), Inches(5.5), Inches(1.2),
       "A true-to-life adventure about a girl,\na dark beach, and 92 baby turtles.",
       size=15, color=DARK)
    add_round(slide, Inches(6.8), Inches(1.45), Inches(6.05), Inches(5.2), WHITE)
    tb(slide, Inches(7.1), Inches(1.75), Inches(5.45), Inches(0.45),
       "🤔 Before we read — guess!", size=18, bold=True, color=CORAL)
    bullets(slide, Inches(7.1), Inches(2.35), Inches(5.45), Inches(3.6), [
        "Who do you think Maya is?",
        "Why would turtles need rescuing?",
        "Why is it happening at MIDNIGHT?",
        "Do you think the story ends happy or sad?",
    ], size=16, sp=18)
    tb(slide, Inches(7.1), Inches(5.95), Inches(5.45), Inches(0.5),
       "There are no wrong guesses here! 💛", size=14, bold=True, color=TEAL)
    notes(slide,
          "Before we read a single word, good readers make predictions. Look at the title and "
          "the pictures. What do you think is going to happen?",
          "Ask two or three of the questions, not all four. Accept every guess warmly and write "
          "one prediction down so you can check it after Part 3.",
          "Any reasonable guess: Maya is a girl / a helper; turtles are lost; midnight because "
          "it is safer or cooler.",
          "Saying 'I don't know' because they are afraid of being wrong.",
          "Offer two options: 'Do you think Maya is a kid or a scientist?'",
          "Ask them to predict the LAST sentence of the story.",
          "That is exactly what real readers do — great thinking! ⭐",
          "10–20 min")


def _reading_slide(part_no, title, text, read_mode, tip, timing, color, light):
    slide, n = new_slide(title, "READING", timing, f"Reading {part_no}", color)
    add_round(slide, Inches(0.45), Inches(1.45), Inches(8.5), Inches(5.2), WHITE)
    tb(slide, Inches(0.75), Inches(1.7), Inches(7.9), Inches(4.7), text, size=15, color=DARK)
    add_round(slide, Inches(9.2), Inches(1.45), Inches(3.65), Inches(2.35), light)
    tb(slide, Inches(9.42), Inches(1.65), Inches(3.2), Inches(0.4), "How we read this",
       size=14, bold=True, color=NAVY)
    tb(slide, Inches(9.42), Inches(2.12), Inches(3.2), Inches(1.5), read_mode,
       size=14, bold=True, color=color)
    add_round(slide, Inches(9.2), Inches(3.95), Inches(3.65), Inches(2.7), LIGHT_YELLOW)
    tb(slide, Inches(9.42), Inches(4.15), Inches(3.2), Inches(0.4), "💡 Word help",
       size=14, bold=True, color=NAVY)
    tb(slide, Inches(9.42), Inches(4.62), Inches(3.2), Inches(1.9), tip, size=13, color=DARK)
    return slide, n


def s09_reading_1():
    slide, n = _reading_slide(
        "Part 1", "📖 Part 1 — A Night on the Beach",
        PASSAGE_1,
        "👩‍🏫 TEACHER reads first.\n\nYou follow along with\nyour finger.",
        "hatched = came out of an egg\n\nconfuses = mixes up\n\nmarked = has a sign on it",
        "10–20 min", SKY, LIGHT_SKY)
    notes(slide,
          "I will read Part 1 out loud first. Your job is easy — follow along with your finger "
          "and listen to how I pause at the periods.",
          "Read with expression at a slightly slow pace. Then ask one quick check question so "
          "the student stays active. Do not ask them to read yet.",
          "Check question: 'Why do the volunteers use a RED flashlight?' → Because white light "
          "confuses the baby turtles.",
          "Losing their place in the text; missing the red-light detail.",
          "Re-read just the sentence about the red flashlight, then ask again.",
          "Ask: 'What time of year does this happen?' (August)",
          "You followed every line — nice focus! ⭐",
          "10–20 min")


def s10_reading_2():
    slide, n = _reading_slide(
        "Part 2", "📖 Part 2 — A Problem in the Sand",
        PASSAGE_2,
        "🤝 WE read together.\n\nSame speed, same time —\nlike a team!",
        "canyon = deep, narrow valley\n\nscrambled = climbed fast\n\ngently = softly, carefully",
        "10–20 min", TEAL, LIGHT_TEAL)
    notes(slide,
          "Now we read together, at the same time, like a team. If I go too fast, squeeze the "
          "table and I'll slow down.",
          "Choral read at a comfortable pace. Let your voice drop slightly so the student's "
          "voice leads by the second paragraph.",
          "Check question: 'What was the problem?' → Three hatchlings fell into a deep tire "
          "track and could not climb out.",
          "Rushing past punctuation; mumbling on 'scrambled'.",
          "Read one sentence, then have them echo it back before moving on.",
          "Ask: 'Why was Maya NOT allowed to pick up the turtles?'",
          "Our voices matched perfectly — that's real fluency! ⭐",
          "10–20 min")


def s11_reading_3():
    slide, n = _reading_slide(
        "Part 3", "📖 Part 3 — Ninety-Two Chances",
        PASSAGE_3,
        "🌟 YOU read this part\nby yourself.\n\nI'm right here if you\nneed a word.",
        "exhausted = very tired\n\nbiologist = living-things\nscientist\n\nsurvive = stay alive",
        "10–20 min", CORAL, LIGHT_CORAL)
    notes(slide,
          "This is your part. You read it on your own, out loud. Take your time — if a word "
          "looks hard, pause and I will help. There is no rush at all.",
          "Stay quiet unless they stall for more than 3 seconds. Do not correct small errors "
          "mid-sentence; note them and praise first at the end.",
          "Reads through with some hesitation on 'biologist' and 'ninety-two'.",
          "Skipping the quotation marks and reading dialogue flatly; stumbling on 'biologist'.",
          "Read the hard word for them and have them repeat it, then continue.",
          "Ask them to re-read Maya's line with real emotion in the voice.",
          "You read that whole page by yourself. That is Grade 5 reading power! ⭐",
          "10–20 min")


def s12_reading_challenge():
    slide, n = new_slide("⚡ Reading Challenge: Beat the Clock!", "CHALLENGE", "20–25 min",
                         "Reading game", GOLD)
    add_round(slide, Inches(0.45), Inches(1.45), Inches(7.6), Inches(3.1), WHITE)
    tb(slide, Inches(0.75), Inches(1.7), Inches(7.0), Inches(0.4), "Read this out loud — smoothly!",
       size=16, bold=True, color=NAVY)
    tb(slide, Inches(0.75), Inches(2.25), Inches(7.0), Inches(2.1),
       "\"By midnight, ninety-two turtles had reached the waves.\n"
       "Maya was exhausted, but she could not stop smiling.\"",
       size=19, bold=True, color=TEAL, font="Georgia")
    add_round(slide, Inches(8.35), Inches(1.45), Inches(4.5), Inches(3.1), LIGHT_YELLOW)
    tb(slide, Inches(8.6), Inches(1.7), Inches(4.0), Inches(0.4), "⏱️ Your score", size=15,
       bold=True, color=NAVY)
    bullets(slide, Inches(8.6), Inches(2.2), Inches(4.0), Inches(2.2), [
        "Try 1:  ______ seconds",
        "Try 2:  ______ seconds",
        "Beat your own record! 🔥",
    ], size=15, sp=14)
    add_round(slide, Inches(0.45), Inches(4.75), Inches(12.4), Inches(1.85), LIGHT_SKY)
    tb(slide, Inches(0.75), Inches(4.98), Inches(11.8), Inches(0.4),
       "🎯 Fluency Goals — check them off:", size=15, bold=True, color=NAVY)
    goals = ["👀 Look up once", "🛑 Stop at periods", "🔊 Big clear voice", "😀 Add feeling"]
    for i, g in enumerate(goals):
        gleft = Inches(0.75 + i * 3.0)
        add_round(slide, gleft, Inches(5.5), Inches(2.8), Inches(0.75), WHITE)
        tb(slide, gleft, Inches(5.68), Inches(2.8), Inches(0.4), g, size=13, bold=True,
           color=DARK, align=PP_ALIGN.CENTER)
    notes(slide,
          "Quick challenge! Read these two sentences out loud as smoothly as you can. I'll time "
          "you. Then we try again and see if you can beat your own record.",
          "Time attempt one, celebrate it, then attempt two. Emphasize SMOOTH, not fast — praise "
          "expression over speed. Two attempts only, then move on.",
          "Second attempt is usually smoother and 1–3 seconds faster.",
          "Racing so fast the words blur; ignoring the comma and period pauses.",
          "Read it together once first, then let them try solo.",
          "Add a third goal: read it as if you are telling a secret.",
          "You beat your own record! That's a streak 🔥 +1 Star ⭐",
          "20–25 min")


def s13_comprehension_a():
    slide, n = new_slide("🧩 Story Detective — Round 1", "GAME", "20–25 min", "Comprehension", PURPLE)
    qa_card(slide, Inches(0.45), Inches(1.45), Inches(6.1), Inches(2.4), "MULTIPLE CHOICE", SKY,
            "Why do the volunteers use a RED flashlight?\n\n"
            "a)  Red is prettier at night\n"
            "b)  White light confuses baby turtles\n"
            "c)  Red flashlights are cheaper",
            "Point to your answer!")
    qa_card(slide, Inches(6.75), Inches(1.45), Inches(6.1), Inches(2.4), "TRUE OR FALSE", CORAL,
            "\"Maya picked up the turtles and\ncarried them to the ocean.\"\n\n"
            "👍 TRUE          👎 FALSE",
            "Prove it — find the line in Part 2!")
    qa_card(slide, Inches(0.45), Inches(4.05), Inches(6.1), Inches(2.4), "FILL THE BLANK", TEAL,
            "\"By midnight, ______________ turtles\nhad reached the waves.\"\n\n"
            "Hint: it is a number 🔢",
            "Look back at Part 3 if you need to.")
    qa_card(slide, Inches(6.75), Inches(4.05), Inches(6.1), Inches(2.4), "PUT IT IN ORDER", PINK,
            "Number these 1–4:\n\n"
            "___  Maya smoothed a sand ramp\n"
            "___  Diego saw the sand moving\n"
            "___  Dr. Reyes wrote in her notebook\n"
            "___  Three hatchlings fell in a track",
            "Which happened FIRST?")
    notes(slide,
          "Time to be story detectives. These are not test questions — they are clues, and you "
          "are allowed to look back at the story any time you want.",
          "Do the cards one at a time, clockwise. Let the student look back at the passage — "
          "that is a reading skill, not cheating. About 60 seconds per card.",
          "b) White light confuses baby turtles  •  FALSE (she built a ramp instead)  •  "
          "ninety-two  •  Order: 2, 1, 4, 3",
          "On the sequence card, mixing up the ramp and the fall; guessing on true/false without "
          "checking the text.",
          "Narrow it down: 'Was it 9, 92, or 900?' Then re-read that one sentence together.",
          "Ask: 'How do you KNOW that answer? Read me the line that proves it.'",
          "You found the proof in the text — that's expert reading! ⭐",
          "20–25 min")


def s14_comprehension_b():
    slide, n = new_slide("🧩 Story Detective — Round 2", "GAME", "20–25 min", "Comprehension", PURPLE)
    qa_card(slide, Inches(0.45), Inches(1.45), Inches(6.1), Inches(2.4), "SHORT ANSWER", SKY,
            "What problem did Maya find\non the beach?\n\n"
            "Answer in ONE full sentence.",
            "Start with: \"The problem was...\"")
    qa_card(slide, Inches(6.75), Inches(1.45), Inches(6.1), Inches(2.4), "MAIN IDEA", TEAL,
            "What is this story MOSTLY about?\n\n"
            "a)  How to use a flashlight\n"
            "b)  A girl who helps turtles reach the sea\n"
            "c)  A scientist writing a notebook",
            "The main idea covers the WHOLE story.")
    qa_card(slide, Inches(0.45), Inches(4.05), Inches(6.1), Inches(2.4), "INFERENCE 🔍", CORAL,
            "Diego said, \"I knew you'd say that.\"\n\n"
            "What does that tell us about Maya?",
            "The answer is NOT written — you infer it!")
    qa_card(slide, Inches(6.75), Inches(4.05), Inches(6.1), Inches(2.4), "YOUR OPINION 💭", PINK,
            "Would YOU volunteer on a dark beach\nat midnight?\n\n"
            "Why or why not?",
            "There is no wrong answer here!")
    notes(slide,
          "Round two. These questions need a little more thinking — especially the detective one "
          "at the bottom left, where the answer is hidden between the lines.",
          "Let the student answer in full sentences. For the inference card, give real thinking "
          "time — count to five silently before you help.",
          "Problem: three hatchlings were trapped in a deep tire track  •  b)  •  Maya really "
          "cares about turtles / Diego knows her well / she always comes back  •  Opinion: any "
          "answer with a reason.",
          "Retelling the whole story instead of naming the main idea; treating the inference "
          "question as if the answer is written in the text.",
          "For the inference: ask 'Does Diego look surprised? Why not?'",
          "Ask them to invent one more question to ask YOU about the story.",
          "You just read BETWEEN the lines. That's a big-kid reading skill! ⭐",
          "20–25 min")


def s15_vocab_context():
    slide, n = new_slide("🔗 Game 3: Match It — Word ➡️ Meaning", "GAME", "20–25 min",
                         "Vocab game", TEAL)
    left_words = ["exhausted", "biologist", "survive", "gently", "scramble"]
    right_meanings = ["to stay alive", "softly and carefully", "very, very tired",
                      "to climb fast with hands and feet", "a scientist of living things"]
    tb(slide, Inches(0.7), Inches(1.4), Inches(5.2), Inches(0.4), "WORDS", size=15, bold=True,
       color=NAVY)
    tb(slide, Inches(7.2), Inches(1.4), Inches(5.5), Inches(0.4), "MEANINGS", size=15, bold=True,
       color=NAVY)
    for i, word in enumerate(left_words):
        top = Inches(1.88 + i * 1.0)
        add_round(slide, Inches(0.7), top, Inches(4.6), Inches(0.8), LIGHT_TEAL)
        tb(slide, Inches(0.9), top + Inches(0.2), Inches(4.2), Inches(0.42), f"{i + 1}.  {word}",
           size=16, bold=True, color=NAVY)
    for i, meaning in enumerate(right_meanings):
        top = Inches(1.88 + i * 1.0)
        add_round(slide, Inches(7.2), top, Inches(5.6), Inches(0.8), WHITE)
        letter = "ABCDE"[i]
        tb(slide, Inches(7.4), top + Inches(0.2), Inches(5.2), Inches(0.42),
           f"{letter}.  {meaning}", size=14, color=DARK)
    for i in range(5):
        top = Inches(2.13 + i * 1.0)
        tb(slide, Inches(5.45), top, Inches(1.6), Inches(0.4), "____", size=16, bold=True,
           color=CORAL, align=PP_ALIGN.CENTER)
    star_badge(slide)
    notes(slide,
          "Match the word on the left to its meaning on the right. Draw a line in the air with "
          "your finger, or just say 'one goes with C'.",
          "Do the first one together as a model, then let the student finish. If they hesitate, "
          "read the word and the two most likely meanings only.",
          "1–C  •  2–E  •  3–A  •  4–B  •  5–D",
          "Matching 'scramble' with 'survive' because both sound active; guessing by position "
          "instead of meaning.",
          "Cross out the matches already used so fewer options remain each time.",
          "Ask them to use two of the matched words in a single sentence about the story.",
          "Five out of five! Your vocabulary is locked in. ⭐",
          "20–25 min")


def s16_use_the_word():
    slide, n = new_slide("✏️ Game 4: Use the Word!", "SPEAKING", "20–25 min", "Vocab game", GREEN)
    prompts = [
        ("😴", "exhausted", "Tell me about a day YOU felt exhausted."),
        ("🗺️", "journey", "Describe a journey you have taken."),
        ("🤲", "gently", "What do you hold gently?"),
        ("🙋", "volunteer", "How could you volunteer at your school?"),
    ]
    for i, (emoji, word, prompt) in enumerate(prompts):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.3)
        top = Inches(1.5 + row * 2.5)
        add_round(slide, left, top, Inches(6.0), Inches(2.25), WHITE)
        add_oval(slide, left + Inches(0.25), top + Inches(0.6), Inches(1.0), Inches(1.0), LIGHT_GREEN)
        tb(slide, left + Inches(0.25), top + Inches(0.8), Inches(1.0), Inches(0.6), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.5), top + Inches(0.35), Inches(4.3), Inches(0.45), word,
           size=19, bold=True, color=GREEN)
        tb(slide, left + Inches(1.5), top + Inches(0.95), Inches(4.3), Inches(1.1), prompt,
           size=14, color=DARK)
    tb(slide, Inches(0.5), Inches(6.6), Inches(9.6), Inches(0.35),
       "🌟 Full sentences only — I want to hear that word doing its job!",
       size=14, bold=True, color=CORAL)
    notes(slide,
          "Now you own these words. I'll give you a word and a question, and you answer with a "
          "full sentence that uses the word.",
          "Do two or three prompts, not all four. Let the student finish completely before you "
          "say anything. Praise first, then fix at most one thing.",
          "Full sentences, e.g. 'I was exhausted after my soccer game on Saturday.'",
          "Using the word in the wrong form ('I was exhaust'); answering without the target word.",
          "Give them the sentence opening: 'I felt exhausted when...'",
          "Ask them to use the word in a question addressed to you.",
          "Perfect use of the word — and a full sentence! +1 Star ⭐",
          "20–25 min")


def s17_tenses_intro():
    slide, n = new_slide("⏰ Tenses: Your Time Machine", "GRAMMAR", "25–35 min", "Tenses", CORAL)
    tb(slide, Inches(0.45), Inches(1.35), Inches(12.4), Inches(0.45),
       "Every sentence tells us WHEN something happens. That's a tense!",
       size=17, bold=True, color=NAVY)
    timeline_strip(slide, Inches(2.05), "NOW")
    cards = [
        ("⏪", "PAST", "It already\nhappened.", "\"I played soccer\nyesterday.\"", CORAL, LIGHT_CORAL),
        ("⏺️", "NOW", "It is happening\nright now.", "\"I am playing\nsoccer.\"", TEAL, LIGHT_TEAL),
        ("⏩", "FUTURE", "It has not\nhappened yet.", "\"I will play soccer\ntomorrow.\"", GOLD, LIGHT_YELLOW),
    ]
    for i, (icon, label, desc, ex, color, light) in enumerate(cards):
        left = Inches(0.55 + i * 4.15)
        top = Inches(3.5)
        add_round(slide, left, top, Inches(3.95), Inches(2.9), light)
        add_oval(slide, left + Inches(1.5), top + Inches(0.2), Inches(0.95), Inches(0.95), color)
        tb(slide, left + Inches(1.5), top + Inches(0.36), Inches(0.95), Inches(0.6), icon,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.25), Inches(3.65), Inches(0.4), label,
           size=17, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.7), Inches(3.65), Inches(0.6), desc,
           size=13, color=DARK, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(2.3), Inches(3.65), Inches(0.5), ex,
           size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    notes(slide,
          "Here is the big secret of tenses: every sentence is standing somewhere on this "
          "timeline. Past, now, or future. That's it. Once you can hear WHEN, you can name the tense.",
          "Walk left to right along the timeline with your hand. Say all three example sentences "
          "and have the student repeat them while pointing to the right spot.",
          "Student repeats each sentence and points to the matching part of the timeline.",
          "Thinking 'tense' means 'nervous'; assuming every past sentence ends in -ed.",
          "Use their own life: 'Yesterday I ___', 'Right now I ___', 'Tomorrow I ___'.",
          "Ask them to give one sentence for each spot about their weekend.",
          "You just learned what tenses really are in 60 seconds! ⭐",
          "25–35 min")


def _tense_slide(t, timing):
    slide, n = new_slide(f"{t['icon']} {t['name']}", "GRAMMAR", timing, "Tenses", t["color"])
    timeline_strip(slide, Inches(1.28), t["mark"])
    add_round(slide, Inches(0.45), Inches(2.5), Inches(6.1), Inches(4.15), t["light"])
    tb(slide, Inches(0.72), Inches(2.7), Inches(5.6), Inches(0.35), "WHEN DO WE USE IT?",
       size=13, bold=True, color=t["color"])
    tb(slide, Inches(0.72), Inches(3.1), Inches(5.6), Inches(0.85), t["when"], size=15, color=DARK)
    tb(slide, Inches(0.72), Inches(4.05), Inches(5.6), Inches(0.35), "SENTENCE BUILDER",
       size=13, bold=True, color=t["color"])
    add_round(slide, Inches(0.72), Inches(4.45), Inches(5.55), Inches(0.65), WHITE)
    tb(slide, Inches(0.85), Inches(4.6), Inches(5.3), Inches(0.4), t["structure"], size=14,
       bold=True, color=NAVY)
    tb(slide, Inches(0.72), Inches(5.3), Inches(5.6), Inches(0.35), "SIGNAL WORDS 🚦",
       size=13, bold=True, color=t["color"])
    tb(slide, Inches(0.72), Inches(5.7), Inches(5.6), Inches(0.75), t["signals"], size=13,
       color=DARK)
    rows = [("✅  POSITIVE", t["pos"], GREEN), ("❌  NEGATIVE", t["neg"], CORAL),
            ("❓  QUESTION", t["que"], PURPLE)]
    for i, (label, text, color) in enumerate(rows):
        top = Inches(2.5 + i * 1.12)
        add_round(slide, Inches(6.75), top, Inches(6.1), Inches(0.98), WHITE)
        tb(slide, Inches(6.95), top + Inches(0.1), Inches(2.0), Inches(0.32), label, size=11,
           bold=True, color=color)
        tb(slide, Inches(6.95), top + Inches(0.46), Inches(5.7), Inches(0.42), text, size=15,
           bold=True, color=NAVY)
    add_round(slide, Inches(6.75), Inches(5.9), Inches(6.1), Inches(0.75), LIGHT_YELLOW)
    tb(slide, Inches(6.95), Inches(6.0), Inches(5.7), Inches(0.3), "🏠 REAL LIFE", size=11,
       bold=True, color=GOLD)
    tb(slide, Inches(6.95), Inches(6.3), Inches(5.7), Inches(0.35), t["real"], size=14,
       bold=True, color=DARK)
    return slide, n


def s18_simple_present():
    slide, n = _tense_slide(TENSES[0], "25–35 min")
    notes(slide,
          "Simple Present is for things that happen again and again — habits and facts. "
          "Listen: Maya helps the turtles every August. Every August — it repeats.",
          "Read the three example rows out loud. Point at the -s in 'helps' and explain it is "
          "only for he, she, or it. Keep the whole slide under two minutes.",
          "Student makes a sentence about their own routine, e.g. 'I walk to school every day.'",
          "Forgetting the -s for he/she/it ('He play soccer'); using 'do' instead of 'does' "
          "in questions.",
          "Give a fill-in frame: 'Every day I ______.' Let them add just the verb.",
          "Ask them to turn their sentence into a question and a negative.",
          "Great — and you remembered the -s! ⭐",
          "25–35 min")


def s19_present_continuous():
    slide, n = _tense_slide(TENSES[1], "25–35 min")
    notes(slide,
          "Present Continuous is happening RIGHT NOW, this second. The clue is the -ing ending "
          "plus a helper word: am, is, or are. Look — I am talking to you right now.",
          "Act it out. Stand up and say 'I am standing.' Have the student describe what they "
          "are doing at this moment. Physical actions make this tense stick.",
          "'I am sitting at my desk.' / 'I am listening to my teacher.'",
          "Dropping the helper verb ('I going'); using -ing with feeling verbs "
          "('I am knowing the answer').",
          "Point at something you are physically doing and ask 'What am I doing?'",
          "Ask what someone else in their house is doing right now.",
          "Perfect — you used 'am' AND '-ing'. That's the full formula! ⭐",
          "25–35 min")


def s20_simple_past():
    slide, n = _tense_slide(TENSES[2], "25–35 min")
    notes(slide,
          "Simple Past means it started and finished — it is completely over. Most verbs just "
          "add -ed, but the sneaky ones change completely: go becomes went, see becomes saw.",
          "Write three irregular pairs where the student can see them: go/went, eat/ate, "
          "see/saw. Ask for one true sentence about their yesterday.",
          "'I ate pizza last night.' / 'I went to my friend's house.'",
          "Adding -ed to irregular verbs ('goed', 'eated'); keeping the past verb after 'did' "
          "('I didn't went').",
          "Ask a yes/no question first: 'Did you eat breakfast?' Then build the full sentence.",
          "Challenge them to say three past sentences in a row without stopping.",
          "You used an irregular verb correctly — those are the hard ones! ⭐",
          "25–35 min")


def s21_past_continuous():
    slide, n = _tense_slide(TENSES[3], "25–35 min")
    notes(slide,
          "Past Continuous is the movie-scene tense. It was already happening at a moment in "
          "the past. At midnight, Maya was counting turtles — the counting was in progress.",
          "Use the classic interruption pattern: 'I WAS SLEEPING when the phone RANG.' Show the "
          "long action and the short action with your hands.",
          "'I was watching TV when my mom called me.'",
          "Using 'was' with plural subjects ('they was playing'); confusing this with simple past.",
          "Give the frame: 'I was ______ing when ______.' They fill in two blanks only.",
          "Ask them to describe what they were doing at 8 o'clock last night.",
          "You built a two-part sentence — that's advanced! ⭐",
          "25–35 min")


def s22_simple_future():
    slide, n = _tense_slide(TENSES[4], "25–35 min")
    notes(slide,
          "Simple Future is the easiest one of all. Just add the magic word WILL, and the verb "
          "never changes. I will eat. She will eat. They will eat. Same every time!",
          "Emphasize that 'will' works with every subject — no changes at all. Then ask about "
          "their plans for the weekend.",
          "'I will play video games this weekend.' / 'We will visit my grandma.'",
          "Adding -s after will ('He will plays'); using present tense for future plans.",
          "Ask a simple question: 'What will you eat for dinner tonight?'",
          "Ask them to predict three things about next year.",
          "'Will' never changes — and you got that right away! ⭐",
          "25–35 min")


def s23_tense_timeline():
    slide, n = new_slide("🗺️ All Five Tenses on One Map", "GRAMMAR", "25–35 min", "Tenses", NAVY)
    timeline_strip(slide, Inches(1.3), "NOW")
    boxes = [
        ("Past Continuous", "was / were + -ing", "I was reading.", PURPLE, LIGHT_PURPLE, 0.45),
        ("Simple Past", "past verb", "I read it.", CORAL, LIGHT_CORAL, 3.0),
        ("Simple Present", "verb (+s)", "I read daily.", TEAL, LIGHT_TEAL, 5.55),
        ("Present Continuous", "am/is/are + -ing", "I am reading.", SKY, LIGHT_SKY, 8.1),
        ("Simple Future", "will + verb", "I will read.", GOLD, LIGHT_YELLOW, 10.65),
    ]
    for label, formula, example, color, light, x in boxes:
        left = Inches(x)
        top = Inches(2.65)
        add_round(slide, left, top, Inches(2.4), Inches(2.6), light)
        add_rect(slide, left, top, Inches(2.4), Inches(0.1), color)
        tb(slide, left + Inches(0.12), top + Inches(0.28), Inches(2.16), Inches(0.7), label,
           size=13, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.12), top + Inches(1.05), Inches(2.16), Inches(0.6), WHITE)
        tb(slide, left + Inches(0.12), top + Inches(1.18), Inches(2.16), Inches(0.4), formula,
           size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.12), top + Inches(1.8), Inches(2.16), Inches(0.6), example,
           size=13, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.45), Inches(5.55), Inches(12.4), Inches(1.0), LIGHT_YELLOW)
    tb(slide, Inches(0.75), Inches(5.72), Inches(11.8), Inches(0.7),
       "🔑 Shortcut: -ing means it's IN PROGRESS.   WILL means it hasn't happened yet.   "
       "-ed or a changed verb means it's DONE.",
       size=15, bold=True, color=NAVY)
    notes(slide,
          "Here they all are on one map. You do not have to memorize this — just remember the "
          "three shortcuts at the bottom. Those three clues will get you the right answer almost "
          "every time.",
          "Read the shortcut line slowly and let the student repeat it. Then point to a random "
          "box and ask them to say the example sentence.",
          "Student reads any example sentence and can explain one shortcut in their own words.",
          "Feeling overwhelmed by five boxes at once.",
          "Cover three boxes with your hand and work with only two at a time.",
          "Point at a box and ask them to make a NEW sentence in that tense.",
          "Five tenses, one map, and you can read it. Amazing! ⭐",
          "25–35 min")


def s24_tense_detective():
    slide, n = new_slide("🕵️ Tense Game 1: Tense Detective", "GAME", "35–47 min",
                         "Tense game", PINK)
    tb(slide, Inches(0.45), Inches(1.32), Inches(12.4), Inches(0.4),
       "Read each sentence. Which tense is hiding inside? Look for the signal words!",
       size=15, bold=True, color=CORAL)
    sentences = [
        ("Sarah played tennis yesterday.", "yesterday 🚦"),
        ("The turtles are crawling to the sea.", "-ing right now 🚦"),
        ("I will call you tonight.", "will + tonight 🚦"),
        ("Diego walks to school every day.", "every day 🚦"),
        ("We were watching the nest at midnight.", "was/were + -ing 🚦"),
        ("Maya does not like bright lights.", "does not 🚦"),
    ]
    for i, (sentence, hint) in enumerate(sentences):
        col, row = i % 2, i // 2
        left = Inches(0.45 + col * 6.35)
        top = Inches(1.85 + row * 1.62)
        add_round(slide, left, top, Inches(6.05), Inches(1.42), WHITE)
        add_oval(slide, left + Inches(0.18), top + Inches(0.42), Inches(0.55), Inches(0.55), PINK)
        tb(slide, left + Inches(0.18), top + Inches(0.5), Inches(0.55), Inches(0.4), str(i + 1),
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.88), top + Inches(0.2), Inches(5.0), Inches(0.5), sentence,
           size=15, bold=True, color=NAVY)
        tb(slide, left + Inches(0.88), top + Inches(0.72), Inches(3.0), Inches(0.32), hint,
           size=11, color=SOFT)
        tb(slide, left + Inches(4.1), top + Inches(0.72), Inches(1.8), Inches(0.32),
           "Tense: ______", size=12, bold=True, color=TEAL)
    notes(slide,
          "Detective badge on! Every sentence has a signal word that gives away the tense. Find "
          "the signal first, then name the tense. I put the clue in small gray letters to help you.",
          "Do sentence 1 together as a model. Then let the student work through the rest. Move "
          "fast — about 15 seconds each. Celebrate every catch.",
          "1 Simple Past  •  2 Present Continuous  •  3 Simple Future  •  4 Simple Present  •  "
          "5 Past Continuous  •  6 Simple Present",
          "Calling number 5 simple past because it happened in the past; missing that number 6 "
          "is present because of 'does'.",
          "Ask only 'Past, present, or future?' first, then narrow to the exact name.",
          "Ask them to change sentence 1 into all five tenses.",
          "Six for six! You are officially a Tense Detective 🕵️ ⭐",
          "35–47 min")


def s25_fix_sentence():
    slide, n = new_slide("🔧 Tense Game 2: Fix the Sentence!", "GAME", "35–47 min",
                         "Tense game", CORAL)
    tb(slide, Inches(0.45), Inches(1.32), Inches(12.4), Inches(0.4),
       "Each sentence has ONE broken word. Can you repair it? 🛠️",
       size=15, bold=True, color=CORAL)
    broken = [
        ("Yesterday, I go to the park.", "went"),
        ("She are reading a book.", "is reading"),
        ("Tomorrow we goes to the beach.", "will go"),
        ("He didn't went home.", "didn't go"),
        ("I am knowing the answer.", "I know"),
        ("While I slept, the phone was ring.", "was ringing"),
    ]
    for i, (wrong, _fix) in enumerate(broken):
        col, row = i % 2, i // 2
        left = Inches(0.45 + col * 6.35)
        top = Inches(1.85 + row * 1.62)
        add_round(slide, left, top, Inches(6.05), Inches(1.42), WHITE)
        add_round(slide, left + Inches(0.15), top + Inches(0.15), Inches(0.75), Inches(0.32),
                  LIGHT_CORAL)
        tb(slide, left + Inches(0.15), top + Inches(0.17), Inches(0.75), Inches(0.3), "❌ BROKEN",
           size=9, bold=True, color=CORAL, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.0), top + Inches(0.14), Inches(4.9), Inches(0.45), wrong,
           size=15, bold=True, color=DARK)
        add_round(slide, left + Inches(0.15), top + Inches(0.72), Inches(0.75), Inches(0.32),
                  LIGHT_GREEN)
        tb(slide, left + Inches(0.15), top + Inches(0.74), Inches(0.75), Inches(0.3), "✅ FIXED",
           size=9, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.0), top + Inches(0.72), Inches(4.9), Inches(0.45),
           "______________________________", size=14, color=SOFT)
    notes(slide,
          "These sentences are broken, and you are the repair crew. Each one has exactly one "
          "mistake. Read it out loud first — your ear will often catch it before your eyes do.",
          "Have the student read the broken sentence aloud, then say the fixed version as a "
          "complete sentence. Do not just accept the single word.",
          "1 went  •  2 is reading  •  3 will go  •  4 didn't go  •  5 I know  •  6 was ringing",
          "Fixing only part of it ('goes' → 'go' without adding 'will'); missing that 'know' "
          "cannot take -ing.",
          "Read both versions aloud yourself and ask 'Which one sounds right?'",
          "Ask them to explain WHY the fix works, using the signal word.",
          "You heard the mistake by yourself — that's a real English ear! ⭐",
          "35–47 min")


def s26_time_travel():
    slide, n = new_slide("🚀 Tense Games 3 & 4: Time Travel + Sorting", "GAME", "35–47 min",
                         "Tense game", PURPLE)
    add_round(slide, Inches(0.45), Inches(1.35), Inches(6.1), Inches(5.3), LIGHT_PURPLE)
    tb(slide, Inches(0.7), Inches(1.55), Inches(5.6), Inches(0.4), "🚀 GAME 3: Time Travel",
       size=16, bold=True, color=PURPLE)
    tb(slide, Inches(0.7), Inches(2.0), Inches(5.6), Inches(0.4),
       "Base sentence:  \"Maya helps the turtles.\"", size=14, bold=True, color=NAVY)
    travel = [("YESTERDAY", "Maya ______ the turtles.", CORAL),
              ("RIGHT NOW", "Maya ______ the turtles.", TEAL),
              ("TOMORROW", "Maya ______ the turtles.", GOLD)]
    for i, (when, frame, color) in enumerate(travel):
        top = Inches(2.55 + i * 1.32)
        add_round(slide, Inches(0.7), top, Inches(5.55), Inches(1.12), WHITE)
        add_round(slide, Inches(0.88), top + Inches(0.16), Inches(1.6), Inches(0.34), color)
        tb(slide, Inches(0.88), top + Inches(0.18), Inches(1.6), Inches(0.3), when, size=10,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(0.88), top + Inches(0.6), Inches(5.2), Inches(0.4), frame, size=15,
           bold=True, color=DARK)
    add_round(slide, Inches(6.75), Inches(1.35), Inches(6.1), Inches(5.3), LIGHT_SKY)
    tb(slide, Inches(7.0), Inches(1.55), Inches(5.6), Inches(0.4), "🗂️ GAME 4: Tense Sorting",
       size=16, bold=True, color=SKY)
    tb(slide, Inches(7.0), Inches(2.0), Inches(5.6), Inches(0.35),
       "Put each sentence in the right box:", size=13, color=DARK)
    sort_items = ["We will bake a cake.", "She danced all night.",
                  "He is drawing a turtle.", "They visit us every summer.",
                  "I was reading at 9 p.m.", "The bus arrives at eight."]
    for i, item in enumerate(sort_items):
        top = Inches(2.45 + i * 0.62)
        add_round(slide, Inches(7.0), top, Inches(5.6), Inches(0.52), WHITE)
        tb(slide, Inches(7.2), top + Inches(0.1), Inches(4.0), Inches(0.35), item, size=13,
           color=DARK)
        tb(slide, Inches(11.2), top + Inches(0.1), Inches(1.3), Inches(0.35), "____",
           size=13, bold=True, color=CORAL)
    tb(slide, Inches(7.0), Inches(6.2), Inches(5.6), Inches(0.35),
       "Write P (past), N (now), or F (future)", size=12, bold=True, color=NAVY)
    notes(slide,
          "Two quick games side by side. On the left, you are a time traveler — take one "
          "sentence and move it through yesterday, today, and tomorrow. On the right, sort each "
          "sentence into past, now, or future.",
          "Do the left game first, out loud, all three versions. Then switch to sorting. Keep "
          "each game to about three minutes so it stays fresh.",
          "Time Travel: helped / is helping / will help.  Sorting: F, P, N, N, P, N",
          "Saying 'Maya will helps'; sorting 'I was reading' as present because of -ing.",
          "For sorting, ask only 'Did it already happen? Yes or no?' first.",
          "Ask them to invent a fourth version: 'What was Maya doing at midnight?'",
          "You moved that sentence through time perfectly! 🔥 ⭐",
          "35–47 min")


def s27_complete_challenge():
    slide, n = new_slide("⚡ Tense Games 5 & 6: Complete It + 60-Second Sprint", "GAME",
                         "35–47 min", "Tense game", GOLD)
    add_round(slide, Inches(0.45), Inches(1.35), Inches(6.1), Inches(5.3), WHITE)
    tb(slide, Inches(0.7), Inches(1.55), Inches(5.6), Inches(0.4), "✅ GAME 5: Complete the Sentence",
       size=16, bold=True, color=TEAL)
    mcqs = [
        ("Tomorrow I ______ to the zoo.", ["went", "go", "will go"]),
        ("Right now she ______ her homework.", ["does", "is doing", "did"]),
        ("Last night we ______ pizza.", ["eat", "will eat", "ate"]),
        ("Every Monday he ______ soccer.", ["plays", "is playing", "played"]),
    ]
    for i, (stem, opts) in enumerate(mcqs):
        top = Inches(2.05 + i * 1.18)
        tb(slide, Inches(0.7), top, Inches(5.6), Inches(0.35), f"{i + 1}.  {stem}", size=14,
           bold=True, color=NAVY)
        for j, opt in enumerate(opts):
            oleft = Inches(0.8 + j * 1.85)
            add_round(slide, oleft, top + Inches(0.42), Inches(1.7), Inches(0.5), LIGHT_TEAL)
            tb(slide, oleft, top + Inches(0.53), Inches(1.7), Inches(0.35),
               f"{'abc'[j]}) {opt}", size=11, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(6.75), Inches(1.35), Inches(6.1), Inches(5.3), LIGHT_YELLOW)
    tb(slide, Inches(7.0), Inches(1.55), Inches(5.6), Inches(0.4), "⏱️ GAME 6: 60-Second Sprint",
       size=16, bold=True, color=GOLD)
    tb(slide, Inches(7.0), Inches(2.0), Inches(5.6), Inches(0.35),
       "Name the tense as fast as you can!", size=13, color=DARK)
    sprint = ["They ate lunch.", "I am singing.", "She will win.",
              "We swim on Fridays.", "He was cooking.", "It rains a lot.",
              "You will see.", "I was dreaming."]
    for i, item in enumerate(sprint):
        col, row = i % 2, i // 2
        left = Inches(7.0 + col * 2.85)
        top = Inches(2.5 + row * 0.78)
        add_round(slide, left, top, Inches(2.7), Inches(0.62), WHITE)
        tb(slide, left + Inches(0.1), top + Inches(0.15), Inches(2.5), Inches(0.35), item,
           size=12, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(7.0), Inches(5.75), Inches(5.6), Inches(0.7), GOLD)
    tb(slide, Inches(7.0), Inches(5.9), Inches(5.6), Inches(0.4), "🔥 My score:  ____ / 8",
       size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    notes(slide,
          "Left side: pick the right word. Right side: the sprint! I will start the timer and "
          "you name as many tenses as you can in 60 seconds. Ready?",
          "Do the four multiple-choice items calmly first. Then make the sprint genuinely "
          "exciting — count down, keep score out loud, and cheer.",
          "Complete: 1c will go  •  2b is doing  •  3c ate  •  4a plays.  Sprint: past, present "
          "continuous, future, present, past continuous, present, future, past continuous.",
          "In the sprint, rushing and saying 'past' for anything with was/were.",
          "Slow the sprint down and do only four sentences with no timer.",
          "Run the sprint a second time and try to beat the first score.",
          "You beat the clock! That's a streak 🔥 +1 Star ⭐",
          "35–47 min")


def s28_storytelling():
    slide, n = new_slide("📚 Storytelling: Build Your Own Adventure", "STORY", "47–55 min",
                         "Storytelling", PINK)
    scenes = [("🌩️", "1. A storm"), ("🐕", "2. A lost dog"), ("🗺️", "3. An old map"),
              ("🏠", "4. Safe at home")]
    for i, (emoji, label) in enumerate(scenes):
        left = Inches(0.5 + i * 3.2)
        top = Inches(1.4)
        add_round(slide, left, top, Inches(3.0), Inches(2.3), WHITE)
        add_oval(slide, left + Inches(0.9), top + Inches(0.32), Inches(1.2), Inches(1.2), LIGHT_PURPLE)
        tb(slide, left + Inches(0.9), top + Inches(0.52), Inches(1.2), Inches(0.8), emoji,
           size=32, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.72), Inches(2.7), Inches(0.4), label,
           size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(3.95), Inches(6.1), Inches(2.65), LIGHT_YELLOW)
    tb(slide, Inches(0.75), Inches(4.15), Inches(5.6), Inches(0.35), "🎤 Story Prompts",
       size=15, bold=True, color=NAVY)
    bullets(slide, Inches(0.75), Inches(4.6), Inches(5.6), Inches(1.9), [
        "WHO is in your story?",
        "WHERE does it happen?",
        "WHAT HAPPENED?",
        "What was the PROBLEM?",
        "HOW did it END?",
    ], size=14, sp=6)
    add_round(slide, Inches(6.8), Inches(3.95), Inches(6.05), Inches(2.65), LIGHT_TEAL)
    tb(slide, Inches(7.05), Inches(4.15), Inches(5.55), Inches(0.35),
       "⏰ Use all three time zones!", size=15, bold=True, color=TEAL)
    examples = [("PAST", "\"Yesterday, Maya found an old map.\"", CORAL),
                ("NOW", "\"She is opening the map right now.\"", TEAL),
                ("FUTURE", "\"She will find the treasure!\"", GOLD)]
    for i, (label, ex, color) in enumerate(examples):
        top = Inches(4.62 + i * 0.66)
        add_round(slide, Inches(7.05), top, Inches(1.15), Inches(0.4), color)
        tb(slide, Inches(7.05), top + Inches(0.04), Inches(1.15), Inches(0.32), label, size=10,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(8.35), top + Inches(0.02), Inches(4.3), Inches(0.36), ex, size=12,
           color=DARK)
    notes(slide,
          "Now YOU are the author. Look at these four pictures and build a story. Use the five "
          "prompts to keep it organized, and try to use a past, a present, and a future sentence "
          "somewhere in your story.",
          "Let the student talk for 2–3 minutes without interruption. Take quick notes on tense "
          "use. Do not correct while they are speaking — it breaks their flow completely.",
          "A 5–8 sentence story that moves through the four pictures with at least two tenses used.",
          "Telling the whole story in present tense; describing pictures separately instead of "
          "connecting them into one story.",
          "Point to picture 1 and ask just 'Who is this and where are they?' Build it one "
          "picture at a time.",
          "Ask them to add a surprise twist to the ending.",
          "What a storyteller! I could picture the whole thing. +1 Star ⭐",
          "47–55 min")


def s29_speaking():
    slide, n = new_slide("🗣️ Speaking Challenge: Talk Time!", "SPEAKING", "47–55 min",
                         "Speaking", GREEN)
    questions = [
        ("⏪", "What did you do yesterday?", "Simple Past", CORAL),
        ("⏺️", "What are you doing today?", "Present Continuous", TEAL),
        ("🔁", "What do you do every weekend?", "Simple Present", SKY),
        ("⏩", "What will you do this weekend?", "Simple Future", GOLD),
    ]
    for i, (icon, q, tense, color) in enumerate(questions):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.3)
        top = Inches(1.45 + row * 2.4)
        add_round(slide, left, top, Inches(6.0), Inches(2.15), WHITE)
        add_oval(slide, left + Inches(0.25), top + Inches(0.55), Inches(1.0), Inches(1.0), color)
        tb(slide, left + Inches(0.25), top + Inches(0.75), Inches(1.0), Inches(0.6), icon,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.5), top + Inches(0.45), Inches(4.3), Inches(0.85), q,
           size=16, bold=True, color=NAVY)
        add_round(slide, left + Inches(1.5), top + Inches(1.35), Inches(2.6), Inches(0.42), color)
        tb(slide, left + Inches(1.5), top + Inches(1.42), Inches(2.6), Inches(0.32), tense,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.3), Inches(9.3), Inches(0.6), LIGHT_GREEN)
    tb(slide, Inches(0.75), Inches(6.42), Inches(9.0), Inches(0.4),
       "💚 Talk as long as you like. I will listen first, and help after — never in the middle.",
       size=14, bold=True, color=GREEN)
    notes(slide,
          "Four questions, and each one secretly needs a different tense. Answer in as many "
          "sentences as you like. I am going to listen the whole way through and I will not "
          "interrupt you.",
          "Ask all four. Let the student finish completely each time. Then: praise something "
          "specific first, correct at most two things, and have them repeat the corrected "
          "sentence once.",
          "Yesterday → past forms; today → am/is/are + -ing; weekends → present; this weekend "
          "→ will.",
          "Using present tense for the 'yesterday' question; forgetting 'will' for the last one.",
          "Answer the question yourself first as a model, then ask them again.",
          "Ask a follow-up: 'Why?' or 'Who will be with you?'",
          "You used four different tenses without me telling you. That's real English! ⭐",
          "47–55 min")


def s30_quiz():
    slide, n = new_slide("⚡ Quick Challenge — Final Round!", "QUIZ", "55–58 min",
                         "Assessment", GOLD)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.4),
       "Six quick ones. This is a GAME, not a test — I help whenever you want. 💛",
       size=15, bold=True, color=CORAL)
    items = [
        ("1", "\"She ______ TV last night.\"", "watch  /  watched  /  will watch"),
        ("2", "Which is Present Continuous?", "I run  /  I am running  /  I ran"),
        ("3", "\"exhausted\" means...", "excited  /  very tired  /  hungry"),
        ("4", "Fix it: \"He don't like fish.\"", "______________________"),
        ("5", "Why did Maya build a ramp?", "______________________"),
        ("6", "Make a FUTURE sentence about you.", "______________________"),
    ]
    for i, (num, q, opts) in enumerate(items):
        col, row = i % 2, i // 2
        left = Inches(0.45 + col * 6.35)
        top = Inches(1.85 + row * 1.62)
        add_round(slide, left, top, Inches(6.05), Inches(1.42), WHITE)
        add_oval(slide, left + Inches(0.18), top + Inches(0.42), Inches(0.55), Inches(0.55), GOLD)
        tb(slide, left + Inches(0.18), top + Inches(0.5), Inches(0.55), Inches(0.4), num,
           size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.88), top + Inches(0.2), Inches(5.0), Inches(0.5), q,
           size=14, bold=True, color=NAVY)
        tb(slide, left + Inches(0.88), top + Inches(0.75), Inches(5.0), Inches(0.4), opts,
           size=12, color=SOFT)
    notes(slide,
          "Last challenge of the day! Six quick questions that cover everything we did. Remember "
          "— this is a game. If you want a hint, just ask and you still get the star.",
          "Move quickly, roughly 20 seconds per item. Give hints freely. The goal is a confident "
          "finish, not a hard assessment.",
          "1 watched  •  2 I am running  •  3 very tired  •  4 He doesn't like fish  •  5 So the "
          "trapped hatchlings could climb out by themselves  •  6 any correct 'will' sentence.",
          "Rushing question 4 and only fixing 'don't' to 'not'; forgetting 'will' in question 6.",
          "Turn any item into a two-choice question instead of three.",
          "Ask them to write one more question to quiz YOU with.",
          "Look at that — you finished the whole challenge! 🏆 ⭐",
          "55–58 min")


def s31_recap():
    slide, n = new_slide("🎓 What Did We Learn Today?", "RECAP", "58–60 min", "Recap", TEAL)
    wins = [
        ("📖", "Reading", "You read a whole story\nby yourself!"),
        ("📚", "Vocabulary", "12 brand-new words\nin your toolbox."),
        ("⏰", "Tenses", "All 5 tenses on\none timeline."),
        ("🕵️", "Detective", "You found tenses AND\nfixed broken sentences."),
        ("📚", "Storytelling", "You invented your own\nadventure story."),
        ("🗣️", "Speaking", "You spoke in past,\npresent, AND future."),
    ]
    for i, (icon, title, desc) in enumerate(wins):
        col, row = i % 3, i // 3
        left = Inches(0.45 + col * 4.2)
        top = Inches(1.5 + row * 2.5)
        add_round(slide, left, top, Inches(3.95), Inches(2.25), WHITE)
        add_oval(slide, left + Inches(1.5), top + Inches(0.22), Inches(0.9), Inches(0.9), LIGHT_TEAL)
        tb(slide, left + Inches(1.5), top + Inches(0.38), Inches(0.9), Inches(0.6), icon,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.2), Inches(3.65), Inches(0.4), title,
           size=16, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.62), Inches(3.65), Inches(0.6), desc,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.45), Inches(6.5), Inches(9.5), Inches(0.4),
       "🏆 Six skills in sixty minutes. That is a LOT.",
       size=15, bold=True, color=CORAL)
    notes(slide,
          "Look at everything you did in one hour. Six different skills. When you started you "
          "had not read the story yet, and now you can tell me all five tenses.",
          "Point to each card and let the student say what they remember about it. Keep it "
          "celebratory and quick — about 45 seconds total.",
          "Student recalls the story name, a few vocabulary words, and at least two tenses.",
          "Being modest or saying they don't remember.",
          "Prompt with the first sound: 'The story was about a t...?'",
          "Ask which skill they want to practice more next time.",
          "You should be really proud of this. I am! 🏆",
          "58–60 min")


def s32_reflection():
    slide, n = new_slide("💭 How Do YOU Feel?", "REFLECT", "58–60 min", "Reflection", PURPLE)
    tb(slide, Inches(0.45), Inches(1.32), Inches(12.4), Inches(0.4),
       "Point to the face that matches how you feel. There is no wrong answer!",
       size=15, bold=True, color=NAVY)
    skills = ["Reading the story", "New words", "Tenses", "Speaking out loud"]
    faces = ["😐 Still tricky", "🙂 Getting there", "😃 I've got this!"]
    for i, skill in enumerate(skills):
        top = Inches(1.9 + i * 1.15)
        add_round(slide, Inches(0.45), top, Inches(4.2), Inches(0.95), LIGHT_PURPLE)
        tb(slide, Inches(0.7), top + Inches(0.28), Inches(3.8), Inches(0.42), skill, size=15,
           bold=True, color=NAVY)
        for j, face in enumerate(faces):
            fleft = Inches(4.9 + j * 2.75)
            add_round(slide, fleft, top, Inches(2.6), Inches(0.95), WHITE)
            tb(slide, fleft, top + Inches(0.28), Inches(2.6), Inches(0.42), face, size=13,
               bold=True, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.45), Inches(6.6), Inches(12.4), Inches(0.34), LIGHT_YELLOW)
    tb(slide, Inches(0.7), Inches(6.62), Inches(11.9), Inches(0.3),
       "⭐ My favorite part of today was: ________________________________________",
       size=13, bold=True, color=NAVY)
    notes(slide,
          "Last thing before your reward. Tell me honestly how each part felt. If something is "
          "still tricky, that is really useful for me to know — it helps me teach you better.",
          "Go through all four rows quickly. Genuinely thank them for any honest 'still tricky' "
          "answer so honesty feels safe.",
          "Honest self-ratings plus one favorite activity named.",
          "Rating everything as 'I've got this' to please the teacher.",
          "Ask instead: 'Which one would you like to play again next time?'",
          "Ask what they would like to learn in the next class.",
          "Thank you for being honest — that is how we get better. 💛",
          "58–60 min")


def s33_homework():
    slide, n = new_slide("🏡 Fun Practice (Only If You Want!)", "HOMEWORK", "58–60 min",
                         "Homework", SKY)
    tasks = [
        ("🗣️", "Tell a family member the turtle story", "5 minutes"),
        ("✍️", "Write 3 sentences: yesterday, today, tomorrow", "5 minutes"),
        ("🔍", "Spot 5 past-tense verbs in any book", "10 minutes"),
        ("🎨", "Draw your favorite scene and label it", "10 minutes"),
    ]
    for i, (icon, task, mins) in enumerate(tasks):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.3)
        top = Inches(1.5 + row * 2.45)
        add_round(slide, left, top, Inches(6.0), Inches(2.2), WHITE)
        add_oval(slide, left + Inches(0.28), top + Inches(0.6), Inches(1.0), Inches(1.0), LIGHT_SKY)
        tb(slide, left + Inches(0.28), top + Inches(0.8), Inches(1.0), Inches(0.6), icon,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.55), top + Inches(0.55), Inches(4.2), Inches(0.95), task,
           size=15, bold=True, color=NAVY)
        add_round(slide, left + Inches(1.55), top + Inches(1.5), Inches(1.7), Inches(0.4), LIGHT_YELLOW)
        tb(slide, left + Inches(1.55), top + Inches(1.56), Inches(1.7), Inches(0.32), mins,
           size=11, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(6.45), Inches(9.5), Inches(0.4),
       "🎁 Pick just ONE. Doing one thing well beats doing four in a rush.",
       size=14, bold=True, color=CORAL)
    notes(slide,
          "This is optional — pick just one that sounds fun. My favorite is telling the turtle "
          "story to someone at home, because teaching someone else is the fastest way to "
          "remember something.",
          "Let the student choose out loud so they own the decision. Never assign all four.",
          "Student picks one task and says when they will do it.",
          "Feeling pressured to do all four, or seeing this as real homework.",
          "Say clearly: 'Zero is also fine. This is a bonus, not a rule.'",
          "Invite them to bring their drawing or sentences to the next class.",
          "Great pick — I can't wait to hear about it! 🎁",
          "58–60 min")


def s34_congrats():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.24), YELLOW)
    for x, y, c in [(1.0, 1.3, SKY), (11.6, 1.6, PINK), (1.8, 5.6, TEAL), (11.0, 5.4, GOLD)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(1.15), Inches(12), Inches(0.7),
       "🎉 CONGRATULATIONS! 🎉", size=40, bold=True, color=YELLOW,
       align=PP_ALIGN.CENTER, font="Georgia")
    add_oval(slide, Inches(5.42), Inches(2.15), Inches(2.5), Inches(2.5), GOLD)
    tb(slide, Inches(5.42), Inches(2.75), Inches(2.5), Inches(1.3), "🏆", size=60,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.95), Inches(12), Inches(0.5),
       "You are officially an ENGLISH ADVENTURER!", size=26, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(5.6), Inches(12), Inches(0.9),
       "You read a real story  •  Learned 12 new words  •  Mastered 5 tenses\n"
       "Told your own adventure  •  Spoke with confidence",
       size=16, color=LIGHT_SKY, align=PP_ALIGN.CENTER)
    footer(slide, n, "58–60 min", "Reward")
    fade(slide)
    notes(slide,
          "You did it! Look at that trophy — you earned every bit of it. You read, you played, "
          "you told your own story, and you spoke English for a whole hour. I am really proud "
          "of you, and I can't wait for our next adventure.",
          "End on maximum energy. Applaud, give a virtual high five, and say one specific thing "
          "they did well today so the praise feels real and personal.",
          "Student smiles, says thank you, and often asks when the next class is.",
          "Rushing off before the celebration lands.",
          "If they seem shy about praise, just name one concrete fact: 'You read 300 words out loud.'",
          "Ask what adventure they want next time — space, animals, or mystery?",
          "You were fantastic today. See you next time, adventurer! 🏆⭐",
          "58–60 min")


def s35_rubric():
    slide, n = new_slide("📋 Teacher Assessment Rubric", "TEACHER ONLY", "", "Rubric", NAVY)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.35),
       "Fill in after class. Do NOT show this slide to the student.",
       size=13, bold=True, color=CORAL)
    header_y = Inches(1.78)
    add_round(slide, Inches(0.45), header_y, Inches(4.2), Inches(0.5), NAVY)
    tb(slide, Inches(0.65), header_y + Inches(0.09), Inches(3.9), Inches(0.35), "SKILL",
       size=13, bold=True, color=WHITE)
    add_round(slide, Inches(4.8), header_y, Inches(5.6), Inches(0.5), NAVY)
    tb(slide, Inches(5.0), header_y + Inches(0.11), Inches(5.2), Inches(0.32),
       "⭐  1 Beginning  ·  2 Developing  ·  3 Good  ·  4 Very Good  ·  5 Excellent",
       size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(10.55), header_y, Inches(2.3), Inches(0.5), NAVY)
    tb(slide, Inches(10.75), header_y + Inches(0.09), Inches(2.0), Inches(0.35), "NOTES",
       size=13, bold=True, color=WHITE)
    skills = ["Reading", "Vocabulary", "Grammar (Tenses)", "Speaking", "Pronunciation",
              "Confidence", "Participation"]
    for i, skill in enumerate(skills):
        top = header_y + Inches(0.6 + i * 0.68)
        band = WHITE if i % 2 == 0 else LIGHT_SKY
        add_round(slide, Inches(0.45), top, Inches(4.2), Inches(0.58), band)
        tb(slide, Inches(0.65), top + Inches(0.13), Inches(3.9), Inches(0.35), skill, size=14,
           bold=True, color=NAVY)
        add_round(slide, Inches(4.8), top, Inches(5.6), Inches(0.58), band)
        tb(slide, Inches(5.0), top + Inches(0.13), Inches(5.2), Inches(0.35),
           "☆    ☆    ☆    ☆    ☆", size=15, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(10.55), top, Inches(2.3), Inches(0.58), band)
    notes(slide,
          "TEACHER SLIDE — do not display to the student.",
          "Complete this within five minutes of finishing the class, while details are fresh. "
          "Circle one star rating per skill and add one short note.",
          "Most first-demo Grade 5 students land at 3 stars for grammar and 3–4 for "
          "participation.",
          "Rating too harshly on a first demo, which discourages both student and parent.",
          "If unsure between two levels, choose the higher one and note what to watch next time.",
          "Add one specific next-step goal per skill for the follow-up lesson.",
          "Frame all feedback to parents as strengths first, then one growth area.",
          "After class")


def s36_answer_key_a():
    slide, n = new_slide("🔑 Answer Key — Reading & Vocabulary", "TEACHER ONLY", "",
                         "Answer key", NAVY)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.35),
       "TEACHER ONLY — hide this slide before presenting.", size=13, bold=True, color=CORAL)
    add_round(slide, Inches(0.45), Inches(1.78), Inches(6.1), Inches(4.9), WHITE)
    tb(slide, Inches(0.7), Inches(1.98), Inches(5.6), Inches(0.4),
       "📖 Reading — Story Detective", size=16, bold=True, color=SKY)
    bullets(slide, Inches(0.7), Inches(2.45), Inches(5.6), Inches(4.1), [
        "MC: b) White light confuses baby turtles",
        "T/F: FALSE — she built a sand ramp",
        "Blank: ninety-two",
        "Order: 2, 1, 4, 3",
        "Problem: 3 hatchlings trapped in a tire track",
        "Main idea: b) A girl helps turtles reach the sea",
        "Inference: Maya always helps / Diego knows her well",
        "Opinion: any answer with a reason",
    ], size=12, sp=7)
    add_round(slide, Inches(6.75), Inches(1.78), Inches(6.1), Inches(4.9), WHITE)
    tb(slide, Inches(7.0), Inches(1.98), Inches(5.6), Inches(0.4),
       "📚 Vocabulary Games", size=16, bold=True, color=TEAL)
    bullets(slide, Inches(7.0), Inches(2.45), Inches(5.6), Inches(4.1), [
        "Guess the Word: hatchling, exhausted, volunteer, journey",
        "Word Detective: exhausted, journey, canyon",
        "Match It: 1–C, 2–E, 3–A, 4–B, 5–D",
        "Use the Word: any correct full sentence",
        "Vocabulary in context: exhausted = very tired",
    ], size=12, sp=10)
    notes(slide,
          "TEACHER SLIDE — answer key for reading and vocabulary.",
          "Hide this slide in PowerPoint (right-click the thumbnail → Hide Slide) before you "
          "present, or keep it open on a second screen.",
          "See the slide content for all answers.",
          "Accidentally presenting this slide to the student.",
          "Accept any reasonable wording for short-answer and opinion items.",
          "For fast finishers, ask them to justify each answer with a line from the text.",
          "Reference only.",
          "Reference")


def s37_answer_key_b():
    slide, n = new_slide("🔑 Answer Key — Grammar & Tense Games", "TEACHER ONLY", "",
                         "Answer key", NAVY)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.35),
       "TEACHER ONLY — hide this slide before presenting.", size=13, bold=True, color=CORAL)
    add_round(slide, Inches(0.45), Inches(1.78), Inches(6.1), Inches(4.9), WHITE)
    tb(slide, Inches(0.7), Inches(1.98), Inches(5.6), Inches(0.4),
       "🕵️ Tense Detective & Fix the Sentence", size=15, bold=True, color=PINK)
    bullets(slide, Inches(0.7), Inches(2.45), Inches(5.6), Inches(4.1), [
        "Detective: 1 Past, 2 Present Cont., 3 Future,",
        "     4 Present, 5 Past Cont., 6 Present",
        "Fix 1: Yesterday, I went to the park.",
        "Fix 2: She is reading a book.",
        "Fix 3: Tomorrow we will go to the beach.",
        "Fix 4: He didn't go home.",
        "Fix 5: I know the answer.",
        "Fix 6: While I slept, the phone was ringing.",
    ], size=12, sp=6)
    add_round(slide, Inches(6.75), Inches(1.78), Inches(6.1), Inches(4.9), WHITE)
    tb(slide, Inches(7.0), Inches(1.98), Inches(5.6), Inches(0.4),
       "🚀 Time Travel, Sorting, Sprint & Quiz", size=15, bold=True, color=GOLD)
    bullets(slide, Inches(7.0), Inches(2.45), Inches(5.6), Inches(4.1), [
        "Time Travel: helped / is helping / will help",
        "Sorting: F, P, N, N, P, N",
        "Complete: 1c, 2b, 3c, 4a",
        "Sprint: past, pres. cont., future, present,",
        "     past cont., present, future, past cont.",
        "Quiz: 1 watched, 2 I am running, 3 very tired,",
        "     4 He doesn't like fish,",
        "     5 so the hatchlings could climb out,",
        "     6 any correct 'will' sentence",
    ], size=12, sp=6)
    notes(slide,
          "TEACHER SLIDE — answer key for all grammar and tense games.",
          "Hide this slide before presenting. Keep it handy during the tense games so you can "
          "confirm answers instantly without losing pace.",
          "See the slide content for all answers.",
          "Correcting a student answer that is actually acceptable — several items allow "
          "more than one correct phrasing.",
          "If the student produces a different but grammatically correct sentence, accept it "
          "and praise the creativity.",
          "Use the sprint list again as a warm-up at the start of the next lesson.",
          "Reference only.",
          "Reference")


BUILDERS = [
    s01_welcome, s02_mission, s03_icebreaker, s04_vocab_a, s05_vocab_b,
    s06_guess_word, s07_word_detective, s08_reading_intro, s09_reading_1,
    s10_reading_2, s11_reading_3, s12_reading_challenge, s13_comprehension_a,
    s14_comprehension_b, s15_vocab_context, s16_use_the_word, s17_tenses_intro,
    s18_simple_present, s19_present_continuous, s20_simple_past,
    s21_past_continuous, s22_simple_future, s23_tense_timeline,
    s24_tense_detective, s25_fix_sentence, s26_time_travel,
    s27_complete_challenge, s28_storytelling, s29_speaking, s30_quiz,
    s31_recap, s32_reflection, s33_homework, s34_congrats, s35_rubric,
    s36_answer_key_a, s37_answer_key_b,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

out = r"C:\Users\bhushaja\Downloads\shaip\Grade5_English_Adventure_60min_Demo.pptx"
try:
    prs.save(out)
except PermissionError:
    out = out.replace(".pptx", "_new.pptx")
    prs.save(out)

missing = [i + 1 for i, s in enumerate(prs.slides)
           if not s.has_notes_slide or not s.notes_slide.notes_text_frame.text.strip()]
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
print(f"Slides missing teacher notes: {missing if missing else 'none'}")
