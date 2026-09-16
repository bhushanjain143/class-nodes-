"""Grade 2 reading intervention - 90 minutes, 39 slides, no speaker notes.

"The Space Reading Mission" - the student is the Reading Astronaut and flies from
Earth to the Moon by completing reading missions. Built for a struggling reader:
the progression is sound -> syllable -> word -> sight word -> phrase -> sentence
-> short story -> comprehension, and every teaching step is followed by a game.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

SPACE = RGBColor(0x0B, 0x14, 0x37)
INDIGO = RGBColor(0x1E, 0x2A, 0x5A)
SLATE = RGBColor(0x2E, 0x3A, 0x5C)
RED = RGBColor(0xE2, 0x48, 0x3C)
GOLD = RGBColor(0xFF, 0xC5, 0x30)
AMBER = RGBColor(0xEE, 0xA2, 0x1A)
CYAN = RGBColor(0x00, 0xA8, 0xB5)
GREEN = RGBColor(0x2E, 0x9E, 0x5B)
PURPLE = RGBColor(0x7B, 0x5C, 0xD6)
BLUE = RGBColor(0x2E, 0x6F, 0xD9)
SILVER = RGBColor(0xC9, 0xD3, 0xE8)
CREAM = RGBColor(0xFC, 0xFC, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x22, 0x2A, 0x36)
SOFT = RGBColor(0x69, 0x76, 0x88)
L_RED = RGBColor(0xFD, 0xE9, 0xE7)
L_GOLD = RGBColor(0xFF, 0xF3, 0xD8)
L_CYAN = RGBColor(0xDD, 0xF3, 0xF5)
L_GREEN = RGBColor(0xE3, 0xF4, 0xE9)
L_PURPLE = RGBColor(0xEE, 0xE9, 0xFB)
L_BLUE = RGBColor(0xE4, 0xEE, 0xFC)
L_GREY = RGBColor(0xF0, 0xF3, 0xF7)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 39
_counter = {"n": 0}

# ------------------------------------------------------------------ content

MISSION_MAP = [
    ("🌎", "Earth", "Sounds", "0–17 min", GREEN, L_GREEN),
    ("🚀", "Rocket", "Words", "17–41 min", RED, L_RED),
    ("🌙", "Moon", "Sight words", "41–63 min", BLUE, L_BLUE),
    ("⭐", "Space Station", "Sentences", "63–75 min", AMBER, L_GOLD),
    ("🪐", "Planet", "Story", "75–87 min", PURPLE, L_PURPLE),
    ("🏆", "Champion", "You did it!", "87–90 min", CYAN, L_CYAN),
]

SCENE_ITEMS = [("🚀", 1.3, 2.3, 64), ("🌙", 4.3, 2.0, 56), ("⭐", 7.2, 2.25, 46),
               ("🪐", 9.85, 2.05, 54), ("👨‍🚀", 5.8, 3.95, 58), ("☀️", 11.2, 3.7, 46),
               ("⭐", 2.6, 4.3, 34), ("⭐", 9.4, 4.5, 34)]

TARGET_WORDS = [("🚀", "ROCKET", RED, L_RED), ("🌙", "MOON", BLUE, L_BLUE),
                ("⭐", "STAR", GOLD, L_GOLD)]

CONSONANTS = ["m", "s", "t", "p", "b", "d", "c", "f", "r", "n"]
VOWELS = ["a", "e", "i", "o", "u"]

SOUND_BLAST_A = [
    ("🚀", "ROCKET", ["/r/", "/m/", "/s/"]),
    ("🌙", "MOON", ["/s/", "/m/", "/t/"]),
    ("⭐", "STAR", ["/p/", "/d/", "/s/"]),
    ("☀️", "SUN", ["/s/", "/b/", "/f/"]),
]
SOUND_BLAST_B = [
    ("🗺️", "MAP", ["/n/", "/m/", "/r/"]),
    ("🤖", "ROBOT", ["/r/", "/t/", "/c/"]),
    ("🪐", "PLANET", ["/b/", "/p/", "/d/"]),
    ("🚩", "FLAG", ["/f/", "/n/", "/t/"]),
]

SYLLABLES = [("🚀", "ROCKET", ["ROCK", "ET"]), ("🪐", "PLANET", ["PLAN", "ET"]),
             ("🤖", "ROBOT", ["RO", "BOT"]), ("👨‍🚀", "ASTRONAUT", ["AS", "TRO", "NAUT"]),
             ("🪖", "HELMET", ["HEL", "MET"]), ("🌙", "MOON", ["MOON"])]

LAB_SHORT = [("☀️", "sun"), ("🗺️", "map"), ("✈️", "jet"), ("🔴", "red"),
             ("🏃", "run"), ("🐰", "hop")]
LAB_LONGER = [("⭐", "star"), ("🌙", "moon"), ("🚀", "rocket"), ("🤖", "robot")]

BUILD_A = [("☀️", "SUN", ["S", "U", "N"], ["/s/", "/u/", "/n/"], AMBER, L_GOLD),
           ("🗺️", "MAP", ["M", "A", "P"], ["/m/", "/a/", "/p/"], GREEN, L_GREEN),
           ("✈️", "JET", ["J", "E", "T"], ["/j/", "/e/", "/t/"], BLUE, L_BLUE)]
BUILD_B = [("⭐", "STAR", ["S", "T", "A", "R"], ["/s/", "/t/", "/ar/"], GOLD, L_GOLD),
           ("🌙", "MOON", ["M", "OO", "N"], ["/m/", "/oo/", "/n/"], BLUE, L_BLUE),
           ("🤖", "ROBOT", ["RO", "BOT"], ["/ro/", "/bot/"], PURPLE, L_PURPLE)]

HUNT_WORDS = [("STAR", 1.35, 2.35, GOLD), ("MOON", 5.15, 4.55, BLUE),
              ("SUN", 9.35, 2.3, AMBER), ("ROCKET", 2.55, 5.05, RED),
              ("MAP", 10.6, 4.85, GREEN)]
HUNT_DECOR = [("⭐", 3.6, 2.2, 36), ("🌙", 7.0, 2.25, 40), ("🚀", 6.1, 3.15, 38),
              ("☀️", 11.3, 3.25, 34), ("🗺️", 8.5, 4.6, 32), ("⭐", 4.5, 3.4, 26),
              ("⭐", 9.9, 3.6, 26), ("👨‍🚀", 1.5, 3.5, 34)]

HUNT_STEPS = [("1", "Find MOON", BLUE), ("2", "Find STAR", GOLD), ("3", "Find SUN", AMBER),
              ("4", "Find ROCKET", RED), ("5", "Find MAP", GREEN)]

ASTRONAUT_SAYS = [
    ("🙋", "Astronaut says touch your head."),
    ("🚀", "Astronaut says make a rocket sound!"),
    ("👉", "Astronaut says point to the word STAR."),
    ("📖", "Astronaut says read MOON."),
    ("👏", "Astronaut says clap when you hear /s/."),
    ("🦘", "Astronaut says jump like an astronaut!"),
]

SIGHT_A = [("I", "I can go."), ("am", "I am here."), ("the", "the moon"),
           ("a", "a star"), ("is", "It is red.")]
SIGHT_B = [("my", "my rocket"), ("can", "I can see."), ("see", "I see a star."),
           ("we", "we go"), ("go", "Let's go!")]

PASSWORDS = [("🛰️", "THE", "the moon", BLUE), ("🛰️", "CAN", "I can go.", GREEN),
             ("🛰️", "SEE", "I see a star.", GOLD)]

SENTENCES_A = [("I see a star.", "⭐", GOLD, L_GOLD), ("I see the moon.", "🌙", BLUE, L_BLUE),
               ("The sun is hot.", "☀️", AMBER, L_GOLD)]
SENTENCES_B = [("The rocket is red.", "🚀", RED, L_RED), ("I can run.", "🏃", GREEN, L_GREEN),
               ("We can go.", "🚀", CYAN, L_CYAN), ("The robot is big.", "🤖", PURPLE, L_PURPLE)]

READ_ROUNDS = [("ROUND 1", "I read", "Teacher reads. Student listens.", RED),
               ("ROUND 2", "We read", "Teacher and student read together.", GOLD),
               ("ROUND 3", "You read", "Student reads one sentence.", GREEN)]

PUZZLES_A = [(["I", "see", "the", "moon"], "I see the moon.", "🌙", BLUE, L_BLUE),
             (["The", "rocket", "is", "red"], "The rocket is red.", "🚀", RED, L_RED),
             (["I", "can", "see", "a", "star"], "I can see a star.", "⭐", GOLD, L_GOLD)]
PUZZLES_B = [(["The", "sun", "is", "big"], "The sun is big.", "☀️", AMBER, L_GOLD),
             (["We", "can", "go", "up"], "We can go up.", "🚀", CYAN, L_CYAN)]

VOCAB = [("👨‍🚀", "ASTRONAUT", "AS · TRO · NAUT", "A person who goes to space."),
         ("🚀", "ROCKET", "ROCK · ET", "It flies up to space."),
         ("🪐", "PLANET", "PLAN · ET", "A big round world in space."),
         ("🪖", "HELMET", "HEL · MET", "It keeps your head safe."),
         ("🤖", "ROBOT", "RO · BOT", "A machine that can move."),
         ("🌌", "SPACE", "SPACE", "The dark place past the sky.")]

VOCAB_HUNT = [("👨‍🚀", "The ______ is in space.", "ASTRONAUT"),
              ("🚀", "The ______ is red.", "ROCKET"),
              ("🤖", "The ______ is big.", "ROBOT"),
              ("🪖", "I see my ______.", "HELMET")]

STORY = [
    ("Part 1", "🚀", ["Leo is a boy.", "Leo has a red rocket.", "The rocket is big.",
                      "Leo puts on his helmet.", "Leo gets in the rocket."], RED, L_RED),
    ("Part 2", "☀️", ["The rocket goes up.", "Up, up, up!", "Leo can see the sun.",
                      "The sun is hot.", "Leo can see stars."], AMBER, L_GOLD),
    ("Part 3", "🌙", ["Leo sees the moon.", "The moon is big.", "He sees a bright star.",
                      "Leo sees a small planet.", "Leo is happy."], BLUE, L_BLUE),
    ("Part 4", "🤖", ["A robot waves at Leo.", "Leo waves back.", "Leo flies back home.",
                      "He tells his mom about space.", "It was a great space trip!"],
     PURPLE, L_PURPLE),
]

STORY_PASSES = [("PASS 1", "I read", "Teacher reads. Student follows with a finger.", RED),
                ("PASS 2", "Echo", "Teacher reads one line. Student echoes it.", AMBER),
                ("PASS 3", "We read", "Both read the page out loud together.", BLUE),
                ("PASS 4", "You read", "Student reads the 2 easiest lines alone.", GREEN)]

DETECTIVE_A = [("🗺️", "Where does Leo go?", ["The park", "Space", "School"]),
               ("🎨", "What color is the rocket?", ["Red", "Blue", "Green"])]
DETECTIVE_B = [("👀", "What does Leo see?", ["A moon", "A train", "A dog"]),
               ("😀", "How does Leo feel?", ["Happy", "Angry", "Sleepy"])]

FINAL_WORDS = [("🚀", "ROCKET", RED, L_RED), ("🌙", "MOON", BLUE, L_BLUE),
               ("⭐", "STAR", GOLD, L_GOLD)]
FINAL_SENTENCES = [("I see the moon.", "🌙"), ("The rocket is red.", "🚀")]

LEARNED = [("🔊", "Sounds", "m s t p b d c f r n"), ("👏", "Word parts", "ROCK · ET"),
           ("🧩", "New words", "sun · map · star · moon"),
           ("🛰️", "Sight words", "I  am  the  a  is  my  can  see  we  go"),
           ("📕", "Sentences", "I see the moon."), ("📖", "A story", "Leo's Space Trip"),
           ("🕵️", "Questions", "I answered them!")]

LEVELS = [("LEVEL 1", "FULL SUPPORT", "Teacher reads it. Student repeats.",
           "\"This word is moon. Your turn.\"", RED),
          ("LEVEL 2", "PARTIAL SUPPORT", "Give the first sound, or two choices.",
           "\"It starts with /m/... moon or map?\"", AMBER),
          ("LEVEL 3", "INDEPENDENT", "Student reads or builds it alone.",
           "\"Your turn — read it by yourself.\"", GREEN)]

FIX_IT = [("1", "BREAK THE WORD", "Cover part of it. Show one chunk."),
          ("2", "SOUND IT OUT", "Point under each letter. /m/ /oo/ /n/"),
          ("3", "BLEND", "Slide your finger and say it fast."),
          ("4", "READ TOGETHER", "Say it with him, then fade your voice."),
          ("5", "TRY AGAIN", "Let him say the whole word alone.")]

PRAISE = ["\"Let's solve it together.\"", "\"Take your time.\"",
          "\"Let's look at the first sound.\"", "\"You got the first part!\"",
          "\"Great try!\"", "\"You're becoming a stronger reader.\""]

RUBRIC_SKILLS = ["Letter sounds", "Blending", "CVC words", "Longer words", "Sight words",
                 "Sentence reading", "Story reading", "Comprehension", "Vocabulary",
                 "Confidence"]
RUBRIC_NOTES = ["Words the student read independently", "Words requiring support",
                "Main difficulty observed", "Recommended next lesson"]

# ------------------------------------------------------------------ helpers


def set_run(run, size=20, bold=False, color=DARK, font="Calibri"):
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


def tb(slide, l, t, w, h, text, size=20, bold=False, color=DARK,
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


def bullets(slide, l, t, w, h, items, size=12, color=DARK, sp=5):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(sp)
        r = p.add_run()
        # An item starting with spaces is a wrapped continuation, not a new bullet.
        r.text = ("      " + item.strip()) if item.startswith(" ") else ("•  " + item)
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


def starfield(slide, dots):
    for x, y, d in dots:
        add_oval(slide, Inches(x), Inches(y), Inches(d), Inches(d), SILVER)


def footer(slide, n, timing="", mission=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE4, 0xEA, 0xF0))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL), Inches(0.08),
             GOLD)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), SPACE)
    msg = "🚀 The Space Reading Mission  |  Grade 2  |  90 min"
    if mission:
        msg += f"  |  {mission}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=10,
       color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def chip(slide, text, color):
    add_round(slide, Inches(0.4), Inches(0.26), Inches(3.4), Inches(0.42), color)
    tb(slide, Inches(0.4), Inches(0.31), Inches(3.4), Inches(0.34), text, size=12, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def time_chip(slide, text):
    add_round(slide, Inches(10.3), Inches(0.26), Inches(2.65), Inches(0.42), RED)
    tb(slide, Inches(10.3), Inches(0.31), Inches(2.65), Inches(0.34), text, size=12,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def new_slide(title, tag, timing, mission, accent=CYAN, bg=CREAM):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, accent)
    chip(slide, tag, accent)
    if timing:
        time_chip(slide, timing)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title, size=28, bold=True,
       color=INDIGO, font="Georgia")
    footer(slide, n, timing, mission)
    fade(slide)
    return slide, n


def task_banner(slide, text, color=RED, top=1.34):
    """One primary task per slide, stated in the child's words."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text, size=16,
              bold=True, color=color)


def support_strip(slide, top=6.4, text="", color=L_CYAN, ink=INDIGO):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), color)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(11.7), Inches(0.34), text, size=14,
       bold=True, color=ink)


def ido_wedo_youdo(slide, top=6.4):
    steps = [("I READ", "Teacher", RED), ("WE READ", "Together", AMBER),
             ("YOU READ", "Student", GREEN)]
    for i, (label, who, color) in enumerate(steps):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(top), Inches(3.9), Inches(0.5), color)
        tb(slide, left, Inches(top + 0.07), Inches(3.9), Inches(0.36), f"{label}  ·  {who}",
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def sound_rows(slide, items, start_index, accent, light):
    for i, (emoji, word, options) in enumerate(items):
        top = Inches(1.9 + i * 1.15)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.02), WHITE)
        add_oval(slide, Inches(0.72), top + Inches(0.3), Inches(0.5), Inches(0.5), accent)
        tb(slide, Inches(0.72), top + Inches(0.36), Inches(0.5), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.4), top + Inches(0.24), Inches(0.9), Inches(0.66), emoji, size=28,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.45), top + Inches(0.26), Inches(2.5), Inches(0.6), word, size=26,
           bold=True, color=INDIGO, font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(5.3 + j * 2.55)
            add_round(slide, left, top + Inches(0.23), Inches(2.3), Inches(0.6), light)
            tb(slide, left, top + Inches(0.31), Inches(2.3), Inches(0.44),
               f"{chr(65 + j)}.  {opt}", size=18, bold=True, color=INDIGO,
               align=PP_ALIGN.CENTER)


def build_rows(slide, items, top_start=1.9, gap=1.55):
    for i, (emoji, word, parts, sounds, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.38), light)
        tb(slide, Inches(0.75), top + Inches(0.36), Inches(0.9), Inches(0.68), emoji,
           size=28, align=PP_ALIGN.CENTER)
        for j, part in enumerate(parts):
            left = Inches(1.9 + j * 1.15)
            add_round(slide, left, top + Inches(0.32), Inches(1.0), Inches(0.75), WHITE)
            tb(slide, left, top + Inches(0.42), Inches(1.0), Inches(0.55), part, size=24,
               bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
            if j < len(parts) - 1:
                tb(slide, left + Inches(1.0), top + Inches(0.46), Inches(0.16),
                   Inches(0.45), "+", size=15, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.8), top + Inches(0.44), Inches(2.3), Inches(0.5),
           "  ".join(sounds), size=15, bold=True, color=SOFT)
        tb(slide, Inches(9.1), top + Inches(0.42), Inches(0.5), Inches(0.5), "→", size=20,
           bold=True, color=color)
        add_round(slide, Inches(9.7), top + Inches(0.3), Inches(2.9), Inches(0.78), WHITE)
        tb(slide, Inches(9.7), top + Inches(0.4), Inches(2.9), Inches(0.58), word, size=26,
           bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")


def choice_rows(slide, items, start_index, accent, light, top_start=1.95, gap=1.72,
                height=1.5):
    for i, (emoji, question, options) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.48), Inches(0.52), Inches(0.52), accent)
        tb(slide, Inches(0.75), top + Inches(0.54), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.45), top + Inches(0.44), Inches(0.9), Inches(0.66), emoji,
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.5), top + Inches(0.48), Inches(3.3), Inches(0.6), question,
           size=20, bold=True, color=INDIGO)
        for j, opt in enumerate(options):
            left = Inches(6.0 + j * 2.3)
            add_round(slide, left, top + Inches(0.36), Inches(2.1), Inches(0.78), light)
            tb(slide, left, top + Inches(0.47), Inches(2.1), Inches(0.56),
               f"{chr(65 + j)}. {opt}", size=17, bold=True, color=INDIGO,
               align=PP_ALIGN.CENTER)


def sentence_rows(slide, items, top_start=1.9, gap=1.44, height=1.26, size=30):
    for i, (text, emoji, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), light)
        tb(slide, Inches(0.8), top + Inches(0.28), Inches(1.0), Inches(0.72), emoji,
           size=32, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.1), top + Inches(0.3), Inches(8.0), Inches(0.7), text, size=size,
           bold=True, color=INDIGO)
        add_round(slide, Inches(10.4), top + Inches(0.35), Inches(2.2), Inches(0.56), WHITE)
        tb(slide, Inches(10.4), top + Inches(0.43), Inches(2.2), Inches(0.4),
           "I · WE · YOU", size=13, bold=True, color=color, align=PP_ALIGN.CENTER)


def puzzle_rows(slide, items, start_index, top_start=1.95, gap=1.5, height=1.3):
    for i, (cards, answer, emoji, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), WHITE)
        add_oval(slide, Inches(0.72), top + Inches(0.4), Inches(0.5), Inches(0.5), color)
        tb(slide, Inches(0.72), top + Inches(0.46), Inches(0.5), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.35), top + Inches(0.36), Inches(0.85), Inches(0.62), emoji,
           size=26, align=PP_ALIGN.CENTER)
        for j, card in enumerate(cards):
            left = Inches(2.35 + j * 1.5)
            add_round(slide, left, top + Inches(0.32), Inches(1.35), Inches(0.66), light)
            tb(slide, left, top + Inches(0.41), Inches(1.35), Inches(0.5), card, size=18,
               bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.4), top + Inches(1.0), Inches(9.6), Inches(0.26),
           "Put them in order, then read it out loud.", size=11, color=SOFT)


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, SPACE)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    starfield(slide, [(1.1, 1.2, 0.09), (2.4, 2.6, 0.06), (11.4, 1.5, 0.08),
                      (12.3, 3.4, 0.06), (1.6, 4.9, 0.07), (10.8, 5.4, 0.09),
                      (3.4, 1.0, 0.05), (9.6, 0.9, 0.06)])
    add_oval(slide, Inches(0.5), Inches(3.6), Inches(1.0), Inches(1.0), BLUE)
    add_oval(slide, Inches(11.7), Inches(0.95), Inches(1.05), Inches(1.05), PURPLE)
    tb(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(1.1), "🚀", size=58,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(2.05), Inches(12), Inches(0.9),
       "The Space Reading Mission", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Georgia")
    tb(slide, Inches(0.7), Inches(3.0), Inches(12), Inches(0.5),
       "Grade 2  •  90 Minutes  •  Earth 🌎 → Moon 🌙", size=18, color=SILVER,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.4), Inches(3.75), Inches(6.5), Inches(1.2), CYAN)
    tb(slide, Inches(3.6), Inches(4.02), Inches(6.1), Inches(0.75),
       "You are the READING ASTRONAUT.\nReading is how we fly!", size=19, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)
    for i, (emoji, place, _job, _t, color, _l) in enumerate(MISSION_MAP):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(5.35), Inches(1.65), Inches(0.9), color)
        tb(slide, left, Inches(5.44), Inches(1.65), Inches(0.42), emoji, size=17,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(5.88), Inches(1.65), Inches(0.32), place, size=10, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Mission 1")
    fade(slide)


def s02_meet_astronaut():
    slide, n = new_slide("👨‍🚀 Meet Captain Nova", "MEET", "0–7 min", "Mission 1", CYAN)
    task_banner(slide, "Say hello to your crew!", CYAN)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.6), SPACE)
    starfield(slide, [(1.1, 2.3, 0.07), (5.1, 2.6, 0.06), (1.4, 5.7, 0.06),
                      (5.3, 5.5, 0.07), (3.0, 2.1, 0.05)])
    tb(slide, Inches(0.5), Inches(2.55), Inches(5.4), Inches(1.7), "👨‍🚀", size=92,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.4), Inches(0.7), "CAPTAIN NOVA", size=30,
       bold=True, color=GOLD, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.8), Inches(0.6),
       "Your space buddy for today", size=15, color=SILVER, align=PP_ALIGN.CENTER)
    bubbles = [("💬", "\"Hi! I need your help today.\""),
               ("🚀", "\"My rocket only flies when someone READS.\""),
               ("🎯", "\"Every mission you finish takes us higher.\""),
               ("🤝", "\"We read together. You are never alone.\"")]
    for i, (icon, line) in enumerate(bubbles):
        top = Inches(1.9 + i * 1.18)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(1.02), L_CYAN)
        add_oval(slide, Inches(6.5), top + Inches(0.22), Inches(0.58), Inches(0.58), WHITE)
        tb(slide, Inches(6.5), top + Inches(0.28), Inches(0.58), Inches(0.45), icon,
           size=16, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.3), top + Inches(0.28), Inches(5.3), Inches(0.5), line, size=16,
           bold=True, color=INDIGO)


def s03_mission_map():
    slide, n = new_slide("🗺️ Mission Map — Earth to Champion", "MAP", "0–7 min",
                         "Mission 1", GOLD)
    task_banner(slide, "Six stops. One reading mission at each stop.", AMBER)
    for i, (emoji, place, job, when, color, light) in enumerate(MISSION_MAP):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.4)
        add_round(slide, left, top, Inches(3.9), Inches(2.2), light)
        tb(slide, left, top + Inches(0.18), Inches(3.9), Inches(0.72), emoji, size=34,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.95), Inches(3.6), Inches(0.5), place,
           size=21, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), top + Inches(1.42), Inches(3.4), Inches(0.4), job,
           size=14, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.1), top + Inches(1.78), Inches(1.7), Inches(0.36),
                  WHITE)
        tb(slide, left + Inches(1.1), top + Inches(1.81), Inches(1.7), Inches(0.3), when,
           size=11, bold=True, color=SOFT, align=PP_ALIGN.CENTER)


def s04_warmup_scene():
    slide, n = new_slide("👀 Warm-Up: What Do You See?", "TALK", "0–7 min", "Mission 1", CYAN)
    task_banner(slide, "Just look and talk. No reading yet!", CYAN)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(4.05), SPACE)
    starfield(slide, [(1.0, 2.2, 0.08), (3.9, 3.3, 0.06), (8.5, 2.1, 0.07),
                      (12.0, 3.0, 0.06), (2.2, 5.4, 0.07), (11.2, 5.3, 0.06),
                      (5.6, 5.5, 0.05), (7.2, 3.6, 0.05)])
    for emoji, x, y, size in SCENE_ITEMS:
        tb(slide, Inches(x), Inches(y), Inches(1.6), Inches(1.2), emoji, size=size,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.05), Inches(12.35), Inches(0.62), L_GOLD)
    tb(slide, Inches(0.8), Inches(6.18), Inches(11.7), Inches(0.42),
       "🎤 Ask: \"What do you see?\"  ·  Let him name anything. Every answer counts.",
       size=15, bold=True, color=INDIGO)


def s05_three_words():
    slide, n = new_slide("📢 Our First 3 Space Words", "WORDS", "0–7 min", "Mission 1", RED)
    task_banner(slide, "I read it. You say it after me. That's all for now.", RED)
    for i, (emoji, word, color, light) in enumerate(TARGET_WORDS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(3.9), light)
        tb(slide, left, Inches(2.15), Inches(3.9), Inches(1.3), emoji, size=62,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(3.6), Inches(3.3), Inches(1.15), WHITE)
        tb(slide, left + Inches(0.3), Inches(3.82), Inches(3.3), Inches(0.75), word,
           size=34, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.9), Inches(4.95), Inches(2.1), Inches(0.5), color)
        tb(slide, left + Inches(0.9), Inches(5.03), Inches(2.1), Inches(0.38), "Say it! 🔁",
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    ido_wedo_youdo(slide, 6.1)
    support_strip(slide, 6.72, "These 3 words come back in every mission today.", L_GOLD)


def s06_sound_blast_a():
    slide, n = new_slide("🔊 Game: Rocket Sound Blast", "GAME", "7–17 min", "Mission 2", AMBER)
    task_banner(slide, "What sound does the word START with?", AMBER)
    sound_rows(slide, SOUND_BLAST_A, 1, AMBER, L_GOLD)
    support_strip(slide, 6.45, "Stuck? Say the word slowly and stretch the first sound.",
                  L_GOLD)


def s07_sound_blast_b():
    slide, n = new_slide("🔊 Sound Challenge — Round 2", "GAME", "7–17 min", "Mission 2",
                         AMBER)
    task_banner(slide, "Four more. Listen to the very first sound.", AMBER)
    sound_rows(slide, SOUND_BLAST_B, 5, AMBER, L_GOLD)
    support_strip(slide, 6.45,
                  "Our sounds today:  m  s  t  p  b  d  c  f  r  n   ·   a  e  i  o  u",
                  L_GOLD)


def s08_syllable_mission():
    slide, n = new_slide("👏 Syllable Mission — Big Words Break Apart", "LEARN", "17–24 min",
                         "Mission 3", GREEN)
    task_banner(slide, "A long word is just small parts stuck together.", GREEN)
    for i, (emoji, word, parts) in enumerate(SYLLABLES[:4]):
        top = Inches(1.9 + i * 1.15)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.02), L_GREEN)
        tb(slide, Inches(0.8), top + Inches(0.22), Inches(0.9), Inches(0.66), emoji,
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.85), top + Inches(0.26), Inches(3.3), Inches(0.6), word, size=26,
           bold=True, color=INDIGO, font="Arial Black")
        tb(slide, Inches(5.3), top + Inches(0.3), Inches(0.5), Inches(0.5), "→", size=20,
           bold=True, color=GREEN)
        for j, part in enumerate(parts):
            left = Inches(5.95 + j * 1.85)
            add_round(slide, left, top + Inches(0.22), Inches(1.7), Inches(0.62), WHITE)
            tb(slide, left, top + Inches(0.31), Inches(1.7), Inches(0.46), part, size=20,
               bold=True, color=GREEN, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(11.5), top + Inches(0.25), Inches(1.1), Inches(0.56), GREEN)
        tb(slide, Inches(11.5), top + Inches(0.33), Inches(1.1), Inches(0.4),
           "👏" * len(parts), size=13, color=WHITE, align=PP_ALIGN.CENTER)
    support_strip(slide, 6.45, "Say it with him and clap on each part. Hands make it stick.",
                  L_GREEN)


def s09_clap_the_word():
    slide, n = new_slide("👏 Game: Clap the Word", "GAME", "17–24 min", "Mission 3", GREEN)
    task_banner(slide, "How many claps? Clap it, then say it.", GREEN)
    for i, (emoji, word, parts) in enumerate(SYLLABLES + [("🌟", "SPACE", ["SPACE"])]):
        col, row = i % 4, i // 4
        left = Inches(0.5 + col * 3.13)
        top = Inches(1.95 + row * 2.35)
        add_round(slide, left, top, Inches(2.9), Inches(2.15), L_GREEN)
        tb(slide, left, top + Inches(0.18), Inches(2.9), Inches(0.72), emoji, size=32,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(0.95), Inches(2.7), Inches(0.5), word,
           size=19, bold=True, color=INDIGO, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.35), top + Inches(1.5), Inches(2.2), Inches(0.48),
                  WHITE)
        tb(slide, left + Inches(0.35), top + Inches(1.57), Inches(2.2), Inches(0.36),
           "👏 " * len(parts), size=14, align=PP_ALIGN.CENTER)


def s10_word_lab():
    slide, n = new_slide("🧪 Word-Building Lab", "LEARN", "24–34 min", "Mission 4", BLUE)
    task_banner(slide, "Short words first. Then we grow them.", BLUE)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(0.45), GREEN)
    tb(slide, Inches(0.5), Inches(1.92), Inches(12.35), Inches(0.34),
       "START HERE — 3 letters, 3 sounds", size=13, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    for i, (emoji, word) in enumerate(LAB_SHORT):
        left = Inches(0.5 + i * 2.08)
        add_round(slide, left, Inches(2.42), Inches(1.9), Inches(1.7), L_GREEN)
        tb(slide, left, Inches(2.56), Inches(1.9), Inches(0.6), emoji, size=26,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(3.25), Inches(1.9), Inches(0.62), word.upper(), size=24,
           bold=True, color=GREEN, align=PP_ALIGN.CENTER, font="Arial Black")
    add_round(slide, Inches(0.5), Inches(4.3), Inches(12.35), Inches(0.45), PURPLE)
    tb(slide, Inches(0.5), Inches(4.37), Inches(12.35), Inches(0.34),
       "THEN THESE — a little longer, we break them into parts", size=13, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)
    for i, (emoji, word) in enumerate(LAB_LONGER):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(4.87), Inches(2.9), Inches(1.65), L_PURPLE)
        tb(slide, left, Inches(5.0), Inches(2.9), Inches(0.6), emoji, size=26,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(5.68), Inches(2.9), Inches(0.62), word.upper(), size=24,
           bold=True, color=PURPLE, align=PP_ALIGN.CENTER, font="Arial Black")


def s11_build_a():
    slide, n = new_slide("🧩 Game: Build the Space Word", "GAME", "24–34 min", "Mission 4",
                         BLUE)
    task_banner(slide, "Say each sound. Then push them together fast.", BLUE)
    build_rows(slide, BUILD_A)
    support_strip(slide, 6.6, "Slide your finger under the letters as he blends.", L_BLUE)


def s12_build_b():
    slide, n = new_slide("🧩 Word Challenge — Harder Words", "GAME", "24–34 min",
                         "Mission 4", PURPLE)
    task_banner(slide, "These have a two-letter team. I'll help more here.", PURPLE)
    build_rows(slide, BUILD_B)
    support_strip(slide, 6.6,
                  "OO says /oo/ · AR says /ar/ · ROBOT is 2 chunks: RO + BOT", L_PURPLE)


def s13_word_hunt():
    slide, n = new_slide("🔎 Game: Find the Word", "GAME", "34–41 min", "Mission 5", GOLD)
    task_banner(slide, "I call a word. You point to it. Then you read it.", AMBER)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(4.45), SPACE)
    starfield(slide, [(0.95, 2.1, 0.08), (4.2, 2.0, 0.06), (7.9, 5.6, 0.07),
                      (12.1, 2.6, 0.06), (2.0, 6.0, 0.06), (11.0, 5.9, 0.07),
                      (6.5, 2.0, 0.05), (3.2, 3.0, 0.05)])
    for emoji, x, y, size in HUNT_DECOR:
        tb(slide, Inches(x), Inches(y), Inches(1.0), Inches(0.85), emoji, size=size,
           align=PP_ALIGN.CENTER)
    for word, x, y, color in HUNT_WORDS:
        add_round(slide, Inches(x), Inches(y), Inches(1.95), Inches(0.62), color)
        tb(slide, Inches(x), Inches(y + 0.1), Inches(1.95), Inches(0.44), word, size=18,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
    support_strip(slide, 6.45,
                  "👀 Five words are hiding: STAR · MOON · SUN · ROCKET · MAP", L_GOLD)


def s14_hidden_word_challenge():
    slide, n = new_slide("🔎 Hidden Word Challenge", "GAME", "34–41 min", "Mission 5", GOLD)
    task_banner(slide, "Find it → point to it → read it. One word at a time.", AMBER)
    for i, (num, label, color) in enumerate(HUNT_STEPS):
        top = Inches(1.95 + i * 0.86)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.74), WHITE)
        add_oval(slide, Inches(0.72), top + Inches(0.17), Inches(0.42), Inches(0.42), color)
        tb(slide, Inches(0.72), top + Inches(0.21), Inches(0.42), Inches(0.34), num,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.3), top + Inches(0.14), Inches(3.4), Inches(0.48), label,
           size=20, bold=True, color=INDIGO)
        for j, step in enumerate(["👉 Point to it", "🔊 Sound it out", "📖 Read it"]):
            left = Inches(4.9 + j * 2.6)
            add_round(slide, left, top + Inches(0.12), Inches(2.4), Inches(0.5), L_GREY)
            tb(slide, left, top + Inches(0.18), Inches(2.4), Inches(0.38), step, size=13,
               bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    ido_wedo_youdo(slide, 6.45)


def s15_astronaut_says():
    slide, n = new_slide("🚀 Brain Break: Astronaut Says", "BREAK", "41–45 min",
                         "Mission 6", CYAN)
    task_banner(slide, "Stand up! Only move if I say \"Astronaut says.\"", CYAN)
    for i, (icon, command) in enumerate(ASTRONAUT_SAYS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.4)
        add_round(slide, left, top, Inches(5.95), Inches(1.2), L_CYAN)
        add_oval(slide, left + Inches(0.28), top + Inches(0.3), Inches(0.6), Inches(0.6),
                 WHITE)
        tb(slide, left + Inches(0.28), top + Inches(0.36), Inches(0.6), Inches(0.46), icon,
           size=17, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.36), Inches(4.6), Inches(0.55),
           command, size=16, bold=True, color=INDIGO)
    support_strip(slide, 6.35,
                  "Two of these are reading tasks — that's the point. Keep it fast and silly.",
                  L_CYAN)


def s16_sight_words_intro():
    slide, n = new_slide("🛰️ Sight Words — Words We Just Know", "LEARN", "45–54 min",
                         "Mission 7", BLUE)
    task_banner(slide, "We don't sound these out. We just know them. Five first.", BLUE)
    for i, (word, example) in enumerate(SIGHT_A):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(2.5), L_BLUE)
        add_round(slide, left + Inches(0.2), Inches(2.25), Inches(1.9), Inches(1.0), WHITE)
        tb(slide, left + Inches(0.2), Inches(2.45), Inches(1.9), Inches(0.66), word,
           size=30, bold=True, color=BLUE, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.15), Inches(3.45), Inches(2.0), Inches(0.8), example,
           size=15, bold=True, color=INDIGO, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(4.75), Inches(12.35), Inches(1.5), L_GREY)
    tb(slide, Inches(0.8), Inches(4.9), Inches(11.7), Inches(0.42),
       "How we learn a sight word:", size=16, bold=True, color=INDIGO)
    for i, step in enumerate(["👀 Look at it", "👂 I say it", "🗣️ You say it",
                              "✍️ Trace it in the air", "⚡ Read it fast"]):
        left = Inches(0.8 + i * 2.4)
        add_round(slide, left, Inches(5.42), Inches(2.2), Inches(0.58), WHITE)
        tb(slide, left, Inches(5.53), Inches(2.2), Inches(0.4), step, size=13, bold=True,
           color=BLUE, align=PP_ALIGN.CENTER)


def s17_password():
    slide, n = new_slide("🛰️ Game: Space Station Password", "GAME", "45–54 min",
                         "Mission 7", BLUE)
    task_banner(slide, "The door won't open until you read the password!", BLUE)
    for i, (icon, word, example, color) in enumerate(PASSWORDS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(3.9), SPACE)
        tb(slide, left, Inches(2.15), Inches(3.9), Inches(0.62), icon, size=26,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(2.8), Inches(3.6), Inches(0.42), "PASSWORD:",
           size=14, bold=True, color=SILVER, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.5), Inches(3.3), Inches(2.9), Inches(1.15), color)
        tb(slide, left + Inches(0.5), Inches(3.55), Inches(2.9), Inches(0.75), word,
           size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.4), Inches(4.65), Inches(3.1), Inches(0.62), WHITE)
        tb(slide, left + Inches(0.4), Inches(4.78), Inches(3.1), Inches(0.44), example,
           size=15, bold=True, color=INDIGO, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(5.4), Inches(3.6), Inches(0.4),
           "🔓 Read it → door opens!", size=13, bold=True, color=GOLD,
           align=PP_ALIGN.CENTER)
    support_strip(slide, 6.15,
                  "Too hard? Say \"This word is the.\" He repeats. Ask again in 2 minutes.",
                  L_BLUE)


def s18_sight_word_challenge():
    slide, n = new_slide("⚡ Sight Word Challenge — 5 More", "GAME", "45–54 min",
                         "Mission 7", BLUE)
    task_banner(slide, "Read them fast. Then read all ten in a row.", BLUE)
    for i, (word, example) in enumerate(SIGHT_B):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(2.5), L_BLUE)
        add_round(slide, left + Inches(0.2), Inches(2.25), Inches(1.9), Inches(1.0), WHITE)
        tb(slide, left + Inches(0.2), Inches(2.45), Inches(1.9), Inches(0.66), word,
           size=30, bold=True, color=BLUE, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.15), Inches(3.45), Inches(2.0), Inches(0.8), example,
           size=15, bold=True, color=INDIGO, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(4.75), Inches(12.35), Inches(1.45), SPACE)
    tb(slide, Inches(0.8), Inches(4.88), Inches(11.7), Inches(0.4),
       "⚡ SPEED ROUND — all ten:", size=14, bold=True, color=GOLD)
    all_words = [w for w, _e in SIGHT_A] + [w for w, _e in SIGHT_B]
    for i, word in enumerate(all_words):
        left = Inches(0.8 + i * 1.19)
        add_round(slide, left, Inches(5.36), Inches(1.05), Inches(0.62), WHITE)
        tb(slide, left, Inches(5.48), Inches(1.05), Inches(0.44), word, size=16, bold=True,
           color=BLUE, align=PP_ALIGN.CENTER)
    support_strip(slide, 6.3, "Missed one? Put it back in the deck and ask again at the end.",
                  L_BLUE)


def s19_sentence_rocket():
    slide, n = new_slide("🚀 Sentence Rocket — How We Read Together", "LEARN", "54–63 min",
                         "Mission 8", RED)
    task_banner(slide, "You never read a new sentence alone the first time.", RED)
    for i, (label, who, what, color) in enumerate(READ_ROUNDS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(2.85), L_RED if i == 0
                  else (L_GOLD if i == 1 else L_GREEN))
        add_round(slide, left + Inches(1.1), Inches(2.2), Inches(1.7), Inches(0.52), color)
        tb(slide, left + Inches(1.1), Inches(2.28), Inches(1.7), Inches(0.38), label,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(2.9), Inches(3.5), Inches(0.62), who, size=26,
           bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.3), Inches(3.62), Inches(3.3), Inches(0.98), WHITE)
        tb(slide, left + Inches(0.45), Inches(3.8), Inches(3.0), Inches(0.68), what,
           size=14, color=INDIGO, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.05), Inches(12.35), Inches(1.25), SPACE)
    tb(slide, Inches(0.8), Inches(5.2), Inches(11.7), Inches(0.4),
       "If a word is hard, we do this — never skip it, never rush it:", size=14, bold=True,
       color=GOLD)
    for i, (num, step, _how) in enumerate(FIX_IT):
        left = Inches(0.8 + i * 2.4)
        add_round(slide, left, Inches(5.66), Inches(2.2), Inches(0.5), WHITE)
        tb(slide, left, Inches(5.74), Inches(2.2), Inches(0.36), f"{num}. {step}", size=11,
           bold=True, color=INDIGO, align=PP_ALIGN.CENTER)


def s20_echo_reading():
    slide, n = new_slide("📕 Echo Reading — I Read, You Echo", "READ", "54–63 min",
                         "Mission 8", RED)
    task_banner(slide, "I read the line. You say it right back to me.", RED)
    sentence_rows(slide, SENTENCES_A, top_start=2.0, gap=1.5, height=1.3)
    support_strip(slide, 6.5, "Point to each word as you read. His eyes should follow you.",
                  L_RED)


def s21_sentence_reading():
    slide, n = new_slide("📗 Sentence Reading — Your Turn Grows", "READ", "54–63 min",
                         "Mission 8", GREEN)
    task_banner(slide, "Pick the one you want to read by yourself!", GREEN)
    sentence_rows(slide, SENTENCES_B, top_start=1.9, gap=1.2, height=1.05, size=26)
    support_strip(slide, 6.5, "Let him choose. Choosing is what makes him brave.", L_GREEN)


def s22_puzzle_a():
    slide, n = new_slide("🧩 Game: Build the Rocket Sentence", "GAME", "63–70 min",
                         "Mission 9", GOLD)
    task_banner(slide, "The words fell out of order. Put them back!", AMBER)
    puzzle_rows(slide, PUZZLES_A, 1)
    support_strip(slide, 6.6, "Hint: the first word always wears a CAPITAL letter.", L_GOLD)


def s23_puzzle_b():
    slide, n = new_slide("🧩 Sentence Challenge", "GAME", "63–70 min", "Mission 9", GOLD)
    task_banner(slide, "Two more — then you build one of your own.", AMBER)
    puzzle_rows(slide, PUZZLES_B, 4)
    add_round(slide, Inches(0.5), Inches(5.1), Inches(12.35), Inches(1.6), L_GREEN)
    tb(slide, Inches(0.8), Inches(5.25), Inches(11.7), Inches(0.45),
       "⭐ YOUR OWN SENTENCE — pick any words you like:", size=17, bold=True, color=GREEN)
    add_round(slide, Inches(0.8), Inches(5.8), Inches(11.75), Inches(0.68), WHITE)
    tb(slide, Inches(1.1), Inches(5.92), Inches(11.2), Inches(0.5),
       "I  see  ______________ .", size=24, bold=True, color=INDIGO)


def s24_vocabulary():
    slide, n = new_slide("🌌 Space Vocabulary — Picture, Word, Parts", "LEARN", "70–75 min",
                         "Mission 10", PURPLE)
    task_banner(slide, "Six words that show up in our story next.", PURPLE)
    for i, (emoji, word, chunks, meaning) in enumerate(VOCAB):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.4)
        add_round(slide, left, top, Inches(3.9), Inches(2.2), L_PURPLE)
        tb(slide, left + Inches(0.15), top + Inches(0.18), Inches(1.0), Inches(0.72), emoji,
           size=30, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.25), top + Inches(0.2), Inches(2.5), Inches(0.5), word,
           size=19, bold=True, color=PURPLE, font="Arial Black")
        tb(slide, left + Inches(1.25), top + Inches(0.68), Inches(2.5), Inches(0.36),
           chunks, size=13, bold=True, color=SOFT)
        add_round(slide, left + Inches(0.2), top + Inches(1.1), Inches(3.5), Inches(0.9),
                  WHITE)
        tb(slide, left + Inches(0.35), top + Inches(1.24), Inches(3.2), Inches(0.66),
           meaning, size=13, color=INDIGO, align=PP_ALIGN.CENTER)


def s25_vocab_hunt():
    slide, n = new_slide("🌎 Game: Space Vocabulary Hunt", "GAME", "70–75 min",
                         "Mission 10", PURPLE)
    task_banner(slide, "Which word finishes the sentence? Say the whole sentence.", PURPLE)
    for i, (emoji, sentence, _answer) in enumerate(VOCAB_HUNT):
        top = Inches(1.95 + i * 1.08)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), WHITE)
        add_oval(slide, Inches(0.72), top + Inches(0.28), Inches(0.5), Inches(0.5), PURPLE)
        tb(slide, Inches(0.72), top + Inches(0.34), Inches(0.5), Inches(0.4), str(i + 1),
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.4), top + Inches(0.22), Inches(0.9), Inches(0.66), emoji,
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.45), top + Inches(0.28), Inches(9.9), Inches(0.55), sentence,
           size=24, bold=True, color=INDIGO)
    word_bank = ["ASTRONAUT", "ROCKET", "ROBOT", "HELMET"]
    for i, word in enumerate(word_bank):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(6.4), Inches(2.9), Inches(0.52), L_PURPLE)
        tb(slide, left, Inches(6.49), Inches(2.9), Inches(0.38), word, size=15, bold=True,
           color=PURPLE, align=PP_ALIGN.CENTER)


def s26_story_intro():
    slide, n = new_slide("📖 Story Time: Leo's Space Trip", "STORY", "75–83 min",
                         "Mission 11", AMBER)
    task_banner(slide, "Four short pages. We read each one four times.", AMBER)
    add_round(slide, Inches(0.5), Inches(1.9), Inches(5.6), Inches(4.4), SPACE)
    starfield(slide, [(1.0, 2.3, 0.07), (5.3, 2.5, 0.06), (1.3, 5.6, 0.06),
                      (5.1, 5.7, 0.07)])
    tb(slide, Inches(0.5), Inches(2.55), Inches(5.6), Inches(1.5), "🚀", size=78,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.25), Inches(5.6), Inches(0.7), "LEO", size=34,
       bold=True, color=GOLD, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.8), Inches(5.0), Inches(5.0), Inches(0.9),
       "A boy with a red rocket.", size=16, color=SILVER, align=PP_ALIGN.CENTER)
    for i, (label, who, what, color) in enumerate(STORY_PASSES):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.45), top, Inches(6.4), Inches(0.96), L_GREY)
        add_round(slide, Inches(6.7), top + Inches(0.24), Inches(1.4), Inches(0.48), color)
        tb(slide, Inches(6.7), top + Inches(0.31), Inches(1.4), Inches(0.36), label,
           size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(8.3), top + Inches(0.14), Inches(4.3), Inches(0.42), who, size=17,
           bold=True, color=color)
        tb(slide, Inches(8.3), top + Inches(0.55), Inches(4.3), Inches(0.36), what,
           size=12, color=SOFT)


def _story_slide(index, timing):
    part, emoji, lines, color, light = STORY[index]
    slide, n = new_slide(f"📖 Leo's Space Trip — {part}", "STORY", timing, "Mission 11",
                         color)
    add_round(slide, Inches(0.5), Inches(1.5), Inches(4.5), Inches(4.7), light)
    tb(slide, Inches(0.5), Inches(3.05), Inches(4.5), Inches(1.7), emoji, size=88,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(5.25), Inches(1.5), Inches(7.6), Inches(4.7), WHITE)
    gap, first, size = (0.86, 1.85, 27) if len(lines) <= 4 else (0.78, 1.72, 25)
    for i, line in enumerate(lines):
        tb(slide, Inches(5.6), Inches(first + i * gap), Inches(7.0), Inches(0.7), line,
           size=size, bold=True, color=INDIGO)
    ido_wedo_youdo(slide, 6.4)
    return slide, n


def s27_story_1():
    _story_slide(0, "75–83 min")


def s28_story_2():
    _story_slide(1, "75–83 min")


def s29_story_3():
    _story_slide(2, "75–83 min")


def s30_story_4():
    _story_slide(3, "75–83 min")


def s31_story_challenge():
    slide, n = new_slide("⭐ Story Reading Challenge", "READ", "75–83 min", "Mission 11",
                         GREEN)
    task_banner(slide, "Pick TWO lines you want to read all by yourself.", GREEN)
    picks = ["Leo is a boy.", "The rocket is big.", "Up, up, up!", "The sun is hot.",
             "The moon is big.", "Leo waves back."]
    for i, line in enumerate(picks):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.42)
        add_round(slide, left, top, Inches(5.95), Inches(1.22), L_GREEN)
        add_oval(slide, left + Inches(0.3), top + Inches(0.34), Inches(0.55), Inches(0.55),
                 WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(0.55), Inches(0.42),
           "⭐", size=14, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.34), Inches(4.6), Inches(0.6), line,
           size=22, bold=True, color=INDIGO)
    support_strip(slide, 6.35,
                  "He chooses. If he freezes, read the first word for him and wait.",
                  L_GREEN)


def s32_detective_a():
    slide, n = new_slide("🕵️ Game: Story Detective", "GAME", "83–87 min", "Mission 12",
                         CYAN)
    task_banner(slide, "Think about Leo's trip. Which one is right?", CYAN)
    choice_rows(slide, DETECTIVE_A, 1, CYAN, L_CYAN, top_start=2.0, gap=2.1, height=1.85)
    support_strip(slide, 6.35, "Can't remember? Flip back a page and read it together.",
                  L_CYAN)


def s33_detective_b():
    slide, n = new_slide("🕵️ Comprehension Challenge", "GAME", "83–87 min", "Mission 12",
                         CYAN)
    task_banner(slide, "Two more. Answer in a whole sentence if you can.", CYAN)
    choice_rows(slide, DETECTIVE_B, 3, CYAN, L_CYAN, top_start=2.0, gap=2.1, height=1.85)
    support_strip(slide, 6.35,
                  "Bonus: \"How do YOU feel about going to space?\" Any answer is right.",
                  L_CYAN)


def s34_moon_landing():
    slide, n = new_slide("🌙 Final: Moon Landing Challenge", "CHALLENGE", "87–89 min",
                         "Moon Mission", GOLD)
    task_banner(slide, "Read these to land the rocket. I'm right here with you.", AMBER)
    for i, (emoji, word, color, light) in enumerate(FINAL_WORDS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(2.1), light)
        tb(slide, left + Inches(0.2), Inches(2.3), Inches(1.1), Inches(0.9), emoji,
           size=34, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.4), Inches(2.4), Inches(2.3), Inches(0.9), WHITE)
        tb(slide, left + Inches(1.4), Inches(2.56), Inches(2.3), Inches(0.62), word,
           size=24, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.5), Inches(3.45), Inches(2.9), Inches(0.42), color)
        tb(slide, left + Inches(0.5), Inches(3.5), Inches(2.9), Inches(0.34),
           "Read it! ⭐", size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, (sentence, emoji) in enumerate(FINAL_SENTENCES):
        top = Inches(4.3 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), L_GOLD)
        tb(slide, Inches(0.85), top + Inches(0.18), Inches(0.9), Inches(0.62), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.0), top + Inches(0.2), Inches(8.3), Inches(0.6), sentence,
           size=28, bold=True, color=INDIGO)
        add_round(slide, Inches(10.5), top + Inches(0.24), Inches(2.1), Inches(0.5), GOLD)
        tb(slide, Inches(10.5), top + Inches(0.31), Inches(2.1), Inches(0.38),
           "🚀 Landing!", size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def s35_what_can_i_read():
    slide, n = new_slide("✅ What Can I Read Now?", "RECAP", "89–90 min", "Mission 14",
                         GREEN)
    task_banner(slide, "Look how much you did today. Read this list with me.", GREEN)
    for i, (icon, label, detail) in enumerate(LEARNED):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 1.2)
        width = Inches(5.95) if i < 6 else Inches(12.35)
        add_round(slide, left, top, width, Inches(1.02), L_GREEN)
        add_oval(slide, left + Inches(0.28), top + Inches(0.22), Inches(0.58), Inches(0.58),
                 WHITE)
        tb(slide, left + Inches(0.28), top + Inches(0.28), Inches(0.58), Inches(0.45), icon,
           size=16, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.05), top + Inches(0.16), Inches(4.6), Inches(0.44), label,
           size=17, bold=True, color=GREEN)
        tb(slide, left + Inches(1.05), top + Inches(0.58), Inches(10.9 if i >= 6 else 4.6),
           Inches(0.38), detail, size=13, color=INDIGO)


def s36_badge():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, SPACE)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    starfield(slide, [(1.2, 1.4, 0.09), (11.6, 1.3, 0.08), (2.0, 5.6, 0.07),
                      (11.0, 5.8, 0.09), (3.6, 1.0, 0.05), (9.4, 1.1, 0.06),
                      (0.8, 3.4, 0.06), (12.4, 3.6, 0.06)])
    tb(slide, Inches(0.7), Inches(0.85), Inches(12), Inches(0.8),
       "🏆 READING ASTRONAUT BADGE 🏆", size=34, bold=True, color=GOLD,
       align=PP_ALIGN.CENTER, font="Georgia")
    add_oval(slide, Inches(5.42), Inches(1.75), Inches(2.5), Inches(2.5), GOLD)
    tb(slide, Inches(5.42), Inches(2.3), Inches(2.5), Inches(1.4), "👨‍🚀", size=52,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.5), Inches(12), Inches(0.7), "MISSION COMPLETE!",
       size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(5.25), Inches(12), Inches(0.5),
       "You are a READING ASTRONAUT. 🌎 → 🌙", size=21, color=SILVER,
       align=PP_ALIGN.CENTER)
    for i, (emoji, place, _job, _t, color, _l) in enumerate(MISSION_MAP):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(5.95), Inches(1.65), Inches(0.62), color)
        tb(slide, left, Inches(6.07), Inches(1.65), Inches(0.42), f"{emoji} ✓", size=13,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "89–90 min", "Mission Complete")
    fade(slide)


def s37_support_levels():
    slide, n = new_slide("🧰 Support Levels & Word Rescue", "TEACHER ONLY", "", "Support",
                         SLATE)
    tb(slide, Inches(0.45), Inches(1.34), Inches(12.4), Inches(0.35),
       "TEACHER ONLY — move between levels based on his response, not the clock.",
       size=13, bold=True, color=RED)
    for i, (level, name, what, script, color) in enumerate(LEVELS):
        left = Inches(0.45 + i * 4.2)
        add_round(slide, left, Inches(1.82), Inches(4.0), Inches(2.1), WHITE)
        add_round(slide, left + Inches(0.22), Inches(2.02), Inches(1.3), Inches(0.45),
                  color)
        tb(slide, left + Inches(0.22), Inches(2.09), Inches(1.3), Inches(0.34), level,
           size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.7), Inches(2.06), Inches(2.1), Inches(0.4), name,
           size=13, bold=True, color=color)
        tb(slide, left + Inches(0.22), Inches(2.58), Inches(3.55), Inches(0.6), what,
           size=13, color=INDIGO)
        add_round(slide, left + Inches(0.22), Inches(3.24), Inches(3.55), Inches(0.52),
                  L_GREY)
        tb(slide, left + Inches(0.32), Inches(3.34), Inches(3.35), Inches(0.38), script,
           size=11, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.45), Inches(4.1), Inches(12.4), Inches(0.4),
       "When a word stops him — run these five steps in order:", size=16, bold=True,
       color=INDIGO)
    for i, (num, step, how) in enumerate(FIX_IT):
        left = Inches(0.45 + i * 2.52)
        add_round(slide, left, Inches(4.62), Inches(2.35), Inches(1.28), L_CYAN)
        tb(slide, left + Inches(0.15), Inches(4.74), Inches(2.05), Inches(0.38),
           f"{num}. {step}", size=12, bold=True, color=CYAN)
        tb(slide, left + Inches(0.15), Inches(5.14), Inches(2.05), Inches(0.68), how,
           size=11, color=INDIGO)
    add_round(slide, Inches(0.45), Inches(6.02), Inches(12.4), Inches(0.94), L_GOLD)
    tb(slide, Inches(0.7), Inches(6.08), Inches(11.9), Inches(0.3),
       "Say these out loud — never \"no\", never \"wrong\":", size=12, bold=True,
       color=INDIGO)
    for i, line in enumerate(PRAISE):
        col, row = i % 3, i // 3
        tb(slide, Inches(0.9 + col * 4.05), Inches(6.38 + row * 0.26), Inches(3.9),
           Inches(0.24), line, size=11, color=SOFT)


def s38_answer_key():
    slide, n = new_slide("🔒 TEACHER ONLY — Answer Key", "TEACHER ONLY", "", "Answer key",
                         SLATE)
    tb(slide, Inches(0.45), Inches(1.34), Inches(12.4), Inches(0.35),
       "Hide this slide before presenting, or keep it on a second screen.",
       size=13, bold=True, color=RED)
    add_round(slide, Inches(0.45), Inches(1.78), Inches(6.1), Inches(5.05), WHITE)
    tb(slide, Inches(0.7), Inches(1.94), Inches(5.6), Inches(0.38),
       "🔊 Sounds · 👏 Syllables · 🧩 Words", size=14, bold=True, color=AMBER)
    bullets(slide, Inches(0.7), Inches(2.38), Inches(5.6), Inches(4.3), [
        "Sound Blast 1–4: A /r/, B /m/, C /s/, A /s/",
        "Sound Blast 5–8: B /m/, A /r/, B /p/, A /f/",
        "Syllables: ROCK·ET, PLAN·ET, RO·BOT,",
        "     AS·TRO·NAUT, HEL·MET, MOON",
        "Claps: 2, 2, 2, 3, 2, 1  (SPACE = 1)",
        "Build A: S-U-N, M-A-P, J-E-T",
        "Build B: S-T-AR, M-OO-N, RO-BOT",
        "Word Hunt: STAR top-left, MOON center,",
        "     SUN top-right, ROCKET lower-left,",
        "     MAP lower-right",
        "Sight words: I am the a is my can see we go",
    ], size=12, sp=7)
    add_round(slide, Inches(6.75), Inches(1.78), Inches(6.1), Inches(5.05), WHITE)
    tb(slide, Inches(7.0), Inches(1.94), Inches(5.6), Inches(0.38),
       "🧩 Sentences · 🌌 Vocabulary · 📖 Story", size=14, bold=True, color=PURPLE)
    bullets(slide, Inches(7.0), Inches(2.38), Inches(5.6), Inches(4.3), [
        "Puzzle 1: I see the moon.",
        "Puzzle 2: The rocket is red.",
        "Puzzle 3: I can see a star.",
        "Puzzle 4: The sun is big.",
        "Puzzle 5: We can go up.",
        "Vocab Hunt: ASTRONAUT, ROCKET, ROBOT,",
        "     HELMET",
        "Story Detective 1–2: B Space, A Red",
        "Story Detective 3–4: A A moon, A Happy",
        "Final words: ROCKET, MOON, STAR",
        "Final sentences: I see the moon. /",
        "     The rocket is red.",
    ], size=12, sp=7)


def s39_assessment():
    slide, n = new_slide("📋 TEACHER ONLY — Reading Assessment", "TEACHER ONLY", "",
                         "Assessment", SLATE)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Record where he is today, not where he should be. Fill this in right after class.",
       size=12, bold=True, color=RED)
    cols = [(0.45, 4.3, "READING SKILL"), (4.95, 2.6, "Needs Significant Support"),
            (7.75, 2.4, "Developing"), (10.35, 2.5, "Independent")]
    header_y = Inches(1.7)
    for left, width, label in cols:
        add_round(slide, Inches(left), header_y, Inches(width), Inches(0.42), INDIGO)
        tb(slide, Inches(left), header_y + Inches(0.07), Inches(width), Inches(0.3), label,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = header_y + Inches(0.48 + i * 0.36)
        band = WHITE if i % 2 == 0 else L_GREY
        add_round(slide, Inches(0.45), top, Inches(4.3), Inches(0.34), band)
        tb(slide, Inches(0.65), top + Inches(0.03), Inches(3.9), Inches(0.28), skill,
           size=12, bold=True, color=INDIGO)
        for left, width, _l in cols[1:]:
            add_round(slide, Inches(left), top, Inches(width), Inches(0.34), band)
            tb(slide, Inches(left), top + Inches(0.01), Inches(width), Inches(0.3), "☐",
               size=13, color=SOFT, align=PP_ALIGN.CENTER)
    for i, label in enumerate(RUBRIC_NOTES):
        top = Inches(5.94 + i * 0.27)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.24), L_GOLD)
        tb(slide, Inches(0.65), top + Inches(0.01), Inches(12.0), Inches(0.22),
           f"{label}:", size=11, bold=True, color=INDIGO)


BUILDERS = [
    s01_title, s02_meet_astronaut, s03_mission_map, s04_warmup_scene, s05_three_words,
    s06_sound_blast_a, s07_sound_blast_b, s08_syllable_mission, s09_clap_the_word,
    s10_word_lab, s11_build_a, s12_build_b, s13_word_hunt, s14_hidden_word_challenge,
    s15_astronaut_says, s16_sight_words_intro, s17_password, s18_sight_word_challenge,
    s19_sentence_rocket, s20_echo_reading, s21_sentence_reading, s22_puzzle_a,
    s23_puzzle_b, s24_vocabulary, s25_vocab_hunt, s26_story_intro, s27_story_1,
    s28_story_2, s29_story_3, s30_story_4, s31_story_challenge, s32_detective_a,
    s33_detective_b, s34_moon_landing, s35_what_can_i_read, s36_badge,
    s37_support_levels, s38_answer_key, s39_assessment,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

# This deck ships without speaker notes.
for slide in prs.slides:
    if slide.has_notes_slide:
        for rid, rel in list(slide.part.rels.items()):
            if rel.reltype.endswith("/notesSlide"):
                slide.part.drop_rel(rid)

out = r"C:\Users\bhushaja\Downloads\shaip\Grade2_Space_Reading_Mission_90min.pptx"
try:
    prs.save(out)
except PermissionError:
    out = out.replace(".pptx", "_new.pptx")
    prs.save(out)

story_words = sum(len(line.split()) for _p, _e, lines, _c, _l in STORY for line in lines)
with_notes = [i + 1 for i, s in enumerate(prs.slides) if s.has_notes_slide]
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
print(f"Story word count: {story_words}")
print(f"Slides carrying notes: {with_notes if with_notes else 'none'}")
