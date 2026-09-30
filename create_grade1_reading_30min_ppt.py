"""Grade 1 Reading Lesson - 30 minutes, 27 slides, seated games.

SOURCE NOTE: No source PDF was available when this deck was generated, so the
words, sentences and story below are ORIGINAL teacher-created Grade 1 material.
To rebuild from a real PDF, replace WORDS / SENTENCES / STORY / HUNT_GRID /
SCRAMBLES / PICK / READ_OR_SKIP below and re-run. Everything else adapts.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

NAVY = RGBColor(0x14, 0x3A, 0x7B)
SKY = RGBColor(0x35, 0xB8, 0xF5)
TEAL = RGBColor(0x00, 0xA6, 0x93)
YELLOW = RGBColor(0xFF, 0xD5, 0x3D)
GOLD = RGBColor(0xFF, 0xB0, 0x00)
ORANGE = RGBColor(0xFF, 0x8A, 0x33)
CORAL = RGBColor(0xFF, 0x6B, 0x5E)
PINK = RGBColor(0xF5, 0x5D, 0x9A)
PURPLE = RGBColor(0x8A, 0x5C, 0xD6)
GREEN = RGBColor(0x2F, 0xA5, 0x4A)
CREAM = RGBColor(0xFF, 0xFC, 0xF3)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x22, 0x2B, 0x38)
SOFT = RGBColor(0x6B, 0x7A, 0x8C)
L_SKY = RGBColor(0xE3, 0xF6, 0xFF)
L_TEAL = RGBColor(0xDD, 0xF7, 0xF3)
L_YELLOW = RGBColor(0xFF, 0xF7, 0xD1)
L_CORAL = RGBColor(0xFF, 0xEA, 0xE7)
L_PURPLE = RGBColor(0xF1, 0xE9, 0xFF)
L_GREEN = RGBColor(0xE6, 0xF6, 0xE9)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 27
_counter = {"n": 0}

# ------------------------------------------------------------------ content

WORDS = [
    ("🐱", "CAT", ["C", "A", "T"], SKY, L_SKY),
    ("⚽", "BALL", ["B", "A", "LL"], CORAL, L_CORAL),
    ("🔴", "RED", ["R", "E", "D"], PINK, L_PURPLE),
    ("🐘", "BIG", ["B", "I", "G"], PURPLE, L_PURPLE),
    ("🏃", "RUN", ["R", "U", "N"], GREEN, L_GREEN),
    ("☀️", "SUN", ["S", "U", "N"], GOLD, L_YELLOW),
]

HUNT_GRID = [
    ["CAT", "DOG", "SUN", "CAT", "BALL"],
    ["RED", "BALL", "BIG", "RUN", "RED"],
    ["SUN", "BIG", "CAT", "DOG", "RUN"],
]
HUNT_ROUNDS = [("CAT", 3), ("BALL", 2), ("SUN", 2), ("RED", 2), ("BIG", 2), ("RUN", 2)]

SENTENCES = [
    ("I see a cat.", "🐱", SKY, L_SKY),
    ("The ball is red.", "⚽", CORAL, L_CORAL),
    ("The cat can run.", "🐱💨", GREEN, L_GREEN),
]

SCRAMBLES = [
    (["cat", "I", "a", "see"], "I see a cat."),
    (["is", "ball", "red", "The"], "The ball is red."),
    (["run", "can", "cat", "The"], "The cat can run."),
]

STORY = [
    "Sam has a red ball.",
    "Sam plays with the ball.",
    "The ball is big.",
    "The cat runs to Sam.",
]

STORY_QUESTIONS = [
    ("Who has the ball?", "Sam", "🧒"),
    ("What color is the ball?", "Red", "🔴"),
    ("Is the ball big or small?", "Big", "⚽"),
    ("Who runs to Sam?", "The cat", "🐱"),
]

PICK = [
    ("🐱", "I see a ___.", ["CAT", "DOG", "SUN"], "CAT"),
    ("🔴", "The ball is ___.", ["RED", "RUN", "BIG"], "RED"),
    ("🏃", "The cat can ___.", ["SUN", "RUN", "BIG"], "RUN"),
    ("☀️", "The ___ is hot.", ["CAT", "BALL", "SUN"], "SUN"),
]

READ_OR_SKIP = [
    ("CAT", True), ("TRAIN", False), ("BALL", True),
    ("ELEPHANT", False), ("RED", True), ("MONKEY", False),
]

CHALLENGE = [
    ("⭐", "EASY", "Read 2 words", "CAT      BALL", TEAL, L_TEAL),
    ("⭐⭐", "MEDIUM", "Read 1 sentence", "The ball is red.", ORANGE, L_YELLOW),
    ("⭐⭐⭐", "SUPER READER", "Read 2 sentences",
     "Sam has a big ball.\nThe cat can run.", PINK, L_PURPLE),
]

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


def fade(slide):
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    etree.SubElement(tr, qn("p:fade"))


def footer(slide, n, timing="", step=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE4, 0xEA, 0xF0))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL), Inches(0.08), GOLD)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), NAVY)
    msg = "Grade 1 Reading  |  30 min"
    if step:
        msg += f"  |  {step}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=11, color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=11, color=WHITE, align=PP_ALIGN.RIGHT)


def chip(slide, text, color):
    add_round(slide, Inches(0.4), Inches(0.26), Inches(3.1), Inches(0.44), color)
    tb(slide, Inches(0.4), Inches(0.32), Inches(3.1), Inches(0.36), text, size=13, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def time_chip(slide, text):
    add_round(slide, Inches(10.35), Inches(0.26), Inches(2.6), Inches(0.44), CORAL)
    tb(slide, Inches(10.35), Inches(0.32), Inches(2.6), Inches(0.36), text, size=13, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def new_slide(title, tag, timing, step, accent=SKY, bg=CREAM):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, accent)
    chip(slide, tag, accent)
    time_chip(slide, timing)
    tb(slide, Inches(0.4), Inches(0.82), Inches(12.4), Inches(0.62), title, size=30, bold=True,
       color=NAVY, font="Georgia")
    footer(slide, n, timing, step)
    fade(slide)
    return slide, n


def notes(slide, say, child_does, expected, if_stuck, duration):
    tf = slide.notes_slide.notes_text_frame
    tf.text = (
        f"⏱ TIME: {duration}\n\n"
        f"TEACHER SAYS: {say}\n\n"
        f"CHILD DOES: {child_does}\n\n"
        f"EXPECTED ANSWER: {expected}\n\n"
        f"IF THE CHILD STRUGGLES: {if_stuck}\n\n"
        "REMEMBER: Say \"Almost! Let's try it together\" — never \"Wrong.\" "
        "Give thinking time. Praise every attempt."
    )


def star_strip(slide, top, text="⭐ Great reading!"):
    add_round(slide, Inches(9.6), top, Inches(3.3), Inches(0.55), YELLOW)
    tb(slide, Inches(9.6), top + Inches(0.09), Inches(3.3), Inches(0.4), text, size=15,
       bold=True, color=NAVY, align=PP_ALIGN.CENTER)


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), YELLOW)
    for x, y, c in [(0.65, 0.7, SKY), (11.9, 0.75, PINK), (0.8, 5.85, TEAL), (11.95, 5.8, ORANGE)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(1.65), Inches(12), Inches(0.95),
       "📖 Let's Read Together! 📖", size=48, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.75), Inches(12), Inches(0.6),
       "Grade 1  •  Reading Class  •  30 Minutes", size=20, color=L_SKY,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.4), Inches(3.6), Inches(6.5), Inches(1.3), TEAL)
    tb(slide, Inches(3.6), Inches(3.9), Inches(6.1), Inches(0.8),
       "You are going to be a\nSUPER READER today! ⭐", size=22, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(5.3), Inches(12), Inches(0.7),
       "👀 LOOK  →  🔊 SOUND  →  📖 READ  →  💡 UNDERSTAND", size=20, bold=True,
       color=YELLOW, align=PP_ALIGN.CENTER)
    footer(slide, n, "0–3 min", "Hello")
    fade(slide)
    notes(slide,
          "Hi! I am so happy to see you. Today we are going to read together and play lots of "
          "games. Are you ready to be a Super Reader?",
          "Says hello and shows they are ready — a wave, a smile, or a yes.",
          "Child greets the teacher and looks at the screen.",
          "If shy, smile and wave first and let them copy you. No talking needed yet.",
          "About 30 seconds")


def s02_warmup_picture():
    slide, n = new_slide("👀 What Do You See?", "WARM-UP", "0–3 min", "Warm-up", SKY)
    pics = [("🐱", "a cat"), ("⚽", "a ball"), ("☀️", "the sun")]
    for i, (emoji, answer) in enumerate(pics):
        left = Inches(0.85 + i * 4.05)
        top = Inches(1.75)
        add_round(slide, left, top, Inches(3.7), Inches(3.9), WHITE)
        add_oval(slide, left + Inches(0.85), top + Inches(0.4), Inches(2.0), Inches(2.0), L_SKY)
        tb(slide, left + Inches(0.85), top + Inches(0.75), Inches(2.0), Inches(1.3), emoji,
           size=66, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), top + Inches(2.75), Inches(3.3), Inches(0.7),
           "What is it?", size=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.85), Inches(6.0), Inches(8.4), Inches(0.6),
       "Say it out loud:  \"I see ______.\"", size=22, bold=True, color=CORAL)
    star_strip(slide, Inches(6.0))
    notes(slide,
          "Look at the picture. What do you see? Tell me in a big voice!",
          "Points at each picture and names it out loud.",
          "a cat  •  a ball  •  the sun. A one-word answer is fine to start.",
          "Give two choices: 'Is it a cat or a dog?' Then say the full sentence together.",
          "About 1 minute")


def s03_warmup_sound():
    slide, n = new_slide("🔊 First Sound Game", "WARM-UP", "0–3 min", "Warm-up", TEAL)
    items = [("🐱", "cat", "C"), ("⚽", "ball", "B"), ("☀️", "sun", "S")]
    for i, (emoji, word, letter) in enumerate(items):
        top = Inches(1.62 + i * 1.46)
        add_round(slide, Inches(0.85), top, Inches(8.4), Inches(1.28), WHITE)
        tb(slide, Inches(1.15), top + Inches(0.2), Inches(1.2), Inches(0.88), emoji,
           size=38, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(0.27), Inches(3.0), Inches(0.72), word,
           size=32, bold=True, color=NAVY)
        tb(slide, Inches(5.7), top + Inches(0.38), Inches(3.3), Inches(0.55),
           "starts with . . . ?", size=18, color=SOFT)
    add_round(slide, Inches(9.7), Inches(1.62), Inches(3.2), Inches(3.3), L_TEAL)
    tb(slide, Inches(9.9), Inches(1.85), Inches(2.8), Inches(0.5), "Pick a letter:", size=18,
       bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    for i, letter in enumerate(["C", "B", "S"]):
        add_round(slide, Inches(10.55), Inches(2.45 + i * 0.8), Inches(1.5), Inches(0.64), WHITE)
        tb(slide, Inches(10.55), Inches(2.52 + i * 0.8), Inches(1.5), Inches(0.52), letter,
           size=26, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.85), Inches(6.15), Inches(11.5), Inches(0.55),
       "✅ Answers:  cat = C      ball = B      sun = S", size=20, bold=True, color=GREEN)
    notes(slide,
          "Listen carefully. Ccc-cat. What sound do you hear FIRST in cat?",
          "Says the first sound, then points to the matching letter.",
          "cat = C  •  ball = B  •  sun = S",
          "Stretch the sound out loud: 'Ccccc-at.' Then say 'C!' and let them repeat it.",
          "About 1 minute 30 seconds")


def _word_slide(emoji, word, sounds, color, light, timing):
    slide, n = new_slide(f"📖 Word: {word}", "NEW WORD", timing, "Word reading", color)
    add_oval(slide, Inches(1.15), Inches(1.7), Inches(2.6), Inches(2.6), light)
    tb(slide, Inches(1.15), Inches(2.15), Inches(2.6), Inches(1.7), emoji, size=76,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(4.3), Inches(1.75), Inches(8.5), Inches(1.8), word, size=110, bold=True,
       color=color, align=PP_ALIGN.CENTER, font="Arial Black")
    box_w, gap = 1.55, 0.32
    total = len(sounds) * box_w + (len(sounds) - 1) * gap
    start = 4.3 + (8.5 - total) / 2
    for i, s in enumerate(sounds):
        left = Inches(start + i * (box_w + gap))
        add_round(slide, left, Inches(3.85), Inches(box_w), Inches(1.05), light)
        tb(slide, left, Inches(4.0), Inches(box_w), Inches(0.8), s, size=38, bold=True,
           color=NAVY, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.15), Inches(5.25), Inches(5.3), Inches(1.15), L_YELLOW)
    tb(slide, Inches(1.35), Inches(5.5), Inches(4.9), Inches(0.7),
       "1️⃣  Teacher reads\n2️⃣  Read together", size=18, bold=True, color=NAVY)
    add_round(slide, Inches(6.85), Inches(5.25), Inches(5.3), Inches(1.15), color)
    tb(slide, Inches(7.05), Inches(5.5), Inches(4.9), Inches(0.7),
       "3️⃣  NOW YOU TRY! 🌟", size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    return slide, n


def make_word_slide(idx, timing):
    emoji, word, sounds, color, light = WORDS[idx]
    slide, n = _word_slide(emoji, word, sounds, color, light, timing)
    blended = " - ".join(sounds)
    notes(slide,
          f"Look at this word. Let's sound it out together: {blended}... {word}! "
          f"Now blend it fast: {word}.",
          f"Looks at the word, says each sound, then blends and reads {word} alone.",
          f"{word}",
          f"Point to each letter as you say the sound. Read {word} first yourself, then say it "
          f"together, then let the child try alone.",
          "About 40 seconds")


def s10_word_hunt():
    slide, n = new_slide("🎯 Game 1: Word Hunt", "GAME", "7–10 min", "Word hunt", ORANGE)
    tb(slide, Inches(0.4), Inches(1.4), Inches(9.0), Inches(0.5),
       "Point to the word I say!", size=22, bold=True, color=CORAL)
    for r, row in enumerate(HUNT_GRID):
        for c, word in enumerate(row):
            left = Inches(0.45 + c * 1.92)
            top = Inches(2.0 + r * 1.25)
            add_round(slide, left, top, Inches(1.78), Inches(1.1), WHITE)
            tb(slide, left, top + Inches(0.28), Inches(1.78), Inches(0.6), word, size=26,
               bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(10.15), Inches(2.0), Inches(2.75), Inches(3.5), L_YELLOW)
    tb(slide, Inches(10.35), Inches(2.2), Inches(2.35), Inches(0.45), "Find these:", size=17,
       bold=True, color=NAVY)
    for i, (word, _count) in enumerate(HUNT_ROUNDS):
        top = Inches(2.72 + i * 0.45)
        tb(slide, Inches(10.45), top, Inches(2.2), Inches(0.4), f"{i + 1}.  {word}", size=17,
           bold=True, color=DARK)
    tb(slide, Inches(0.45), Inches(5.95), Inches(9.4), Inches(0.55),
       "🔍 Can you find it? Point to it, then READ it out loud!", size=20, bold=True, color=TEAL)
    notes(slide,
          "Can you find the word CAT? Point to it! Great — now read it out loud.",
          "Points to the word on the grid and reads it aloud. Six rounds, one word at a time.",
          "CAT appears 3 times; BALL, SUN, RED, BIG and RUN each appear 2 times.",
          "Cover part of the grid with your hand so there are fewer words to search. "
          "Say the first sound as a hint.",
          "About 3 minutes")


def s11_hunt_answers():
    slide, n = new_slide("✅ Word Hunt — Answers", "ANSWERS", "7–10 min", "Answers", GREEN)
    tb(slide, Inches(0.4), Inches(1.4), Inches(11.0), Inches(0.5),
       "How many times does each word appear?", size=20, bold=True, color=SOFT)
    for i, (word, count) in enumerate(HUNT_ROUNDS):
        col, row = i % 3, i // 3
        left = Inches(0.6 + col * 4.1)
        top = Inches(2.1 + row * 2.1)
        add_round(slide, left, top, Inches(3.8), Inches(1.75), L_GREEN)
        tb(slide, left + Inches(0.2), top + Inches(0.3), Inches(2.2), Inches(0.7), word,
           size=32, bold=True, color=NAVY)
        add_oval(slide, left + Inches(2.75), top + Inches(0.42), Inches(0.85), Inches(0.85), GREEN)
        tb(slide, left + Inches(2.75), top + Inches(0.55), Inches(0.85), Inches(0.6), str(count),
           size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), top + Inches(1.12), Inches(3.4), Inches(0.45),
           f"appears {count} times", size=15, color=SOFT)
    notes(slide,
          "Let's check together! CAT is here, here, and here — that's three!",
          "Counts along with the teacher and points to each one.",
          "CAT 3  •  BALL 2  •  SUN 2  •  RED 2  •  BIG 2  •  RUN 2",
          "Count out loud together and touch each word as you count.",
          "About 30 seconds")


def make_sentence_slide(idx, timing):
    sentence, emoji, color, light = SENTENCES[idx]
    slide, n = new_slide(f"📚 Sentence {idx + 1}", "READING", timing, "Sentences", color)
    add_round(slide, Inches(0.6), Inches(1.6), Inches(12.1), Inches(2.35), WHITE)
    tb(slide, Inches(0.9), Inches(1.85), Inches(1.6), Inches(1.8), emoji, size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.7), Inches(2.05), Inches(9.7), Inches(1.5), sentence, size=58, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
    steps = [("1️⃣", "TEACHER reads", "Listen and follow", L_YELLOW, NAVY),
             ("2️⃣", "READ TOGETHER", "Say it with me", L_TEAL, NAVY),
             ("3️⃣", "YOU READ IT!", "Your turn 🌟", color, WHITE)]
    for i, (num, title, sub, bg, fg) in enumerate(steps):
        left = Inches(0.6 + i * 4.1)
        top = Inches(4.3)
        add_round(slide, left, top, Inches(3.85), Inches(1.85), bg)
        tb(slide, left + Inches(0.15), top + Inches(0.22), Inches(3.55), Inches(0.55), num,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.82), Inches(3.55), Inches(0.5), title,
           size=19, bold=True, color=fg, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.3), Inches(3.55), Inches(0.42), sub,
           size=14, color=fg, align=PP_ALIGN.CENTER)
    notes(slide,
          f"Listen first: \"{sentence}\" Now let's read it together. Ready? "
          f"Great — now you read it all by yourself!",
          "Follows the words with a finger, reads with the teacher, then reads alone.",
          sentence,
          "Point under each word as you read so the child tracks left to right. "
          "Read just one word at a time and let them echo it.",
          "About 1 minute 20 seconds")


def s15_build_sentence():
    slide, n = new_slide("🎲 Game 2: Build the Sentence", "GAME", "14–17 min",
                         "Build it", PURPLE)
    tb(slide, Inches(0.4), Inches(1.42), Inches(10.0), Inches(0.5),
       "Put the words in the right order!", size=22, bold=True, color=CORAL)
    for i, (tiles, _answer) in enumerate(SCRAMBLES):
        top = Inches(2.0 + i * 1.65)
        add_round(slide, Inches(0.45), top, Inches(12.45), Inches(1.45), WHITE)
        add_oval(slide, Inches(0.7), top + Inches(0.43), Inches(0.6), Inches(0.6), PURPLE)
        tb(slide, Inches(0.7), top + Inches(0.5), Inches(0.6), Inches(0.45), str(i + 1),
           size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        for j, tile in enumerate(tiles):
            left = Inches(1.6 + j * 2.35)
            add_round(slide, left, top + Inches(0.32), Inches(2.15), Inches(0.8), L_PURPLE)
            tb(slide, left, top + Inches(0.47), Inches(2.15), Inches(0.55), tile, size=24,
               bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, Inches(11.15), top + Inches(0.5), Inches(1.6), Inches(0.5), "➡️  ?", size=22,
           bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.45), Inches(6.9), Inches(9.0), Inches(0.3), "", size=10)
    notes(slide,
          "Uh oh, the words got mixed up! Can you put them in the right order? "
          "Which word comes first?",
          "Says the words in the correct order, then reads the whole sentence aloud.",
          "1. I see a cat.   2. The ball is red.   3. The cat can run.",
          "Ask 'Which word has a capital letter? That one goes first!' Then find the word "
          "with the period — that goes last.",
          "About 3 minutes")


def s16_build_answers():
    slide, n = new_slide("✅ Build the Sentence — Answers", "ANSWERS", "14–17 min",
                         "Answers", GREEN)
    for i, (_tiles, answer) in enumerate(SCRAMBLES):
        top = Inches(1.85 + i * 1.65)
        add_round(slide, Inches(0.7), top, Inches(11.9), Inches(1.4), L_GREEN)
        add_oval(slide, Inches(1.0), top + Inches(0.42), Inches(0.62), Inches(0.62), GREEN)
        tb(slide, Inches(1.0), top + Inches(0.5), Inches(0.62), Inches(0.45), str(i + 1),
           size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.0), top + Inches(0.35), Inches(10.2), Inches(0.75), answer,
           size=38, bold=True, color=NAVY, font="Arial Black")
    tb(slide, Inches(0.7), Inches(6.85), Inches(9.0), Inches(0.3), "", size=10)
    notes(slide,
          "Let's check! Listen: I see a cat. Does that sound right? Yes! Now read it with me.",
          "Reads each correct sentence out loud once.",
          "I see a cat.  •  The ball is red.  •  The cat can run.",
          "Read it first, then have the child echo it back.",
          "About 30 seconds")


def s17_story():
    slide, n = new_slide("🔍 Reading Detective: Sam's Ball", "READING", "17–21 min",
                         "Story", CORAL)
    add_round(slide, Inches(0.55), Inches(1.55), Inches(8.6), Inches(4.9), WHITE)
    for i, line in enumerate(STORY):
        tb(slide, Inches(1.0), Inches(1.95 + i * 1.12), Inches(7.8), Inches(0.9), line,
           size=34, bold=True, color=NAVY)
    add_round(slide, Inches(9.45), Inches(1.55), Inches(3.4), Inches(4.9), L_CORAL)
    tb(slide, Inches(9.45), Inches(2.1), Inches(3.4), Inches(1.5), "🧒⚽", size=64,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(9.65), Inches(3.9), Inches(3.0), Inches(2.2),
       "1️⃣ I read it\n\n2️⃣ We read it\n\n3️⃣ YOU read it! 🌟",
       size=19, bold=True, color=NAVY)
    notes(slide,
          "Here is a little story about Sam. I will read it first. Then we read it together. "
          "Then you can be the reader!",
          "Follows along with a finger, reads with the teacher, then reads the four lines alone.",
          "Sam has a red ball. Sam plays with the ball. The ball is big. The cat runs to Sam.",
          "Read one line at a time and let the child echo each line before moving on. "
          "Cover the lines below with your hand so only one line shows.",
          "About 2 minutes 30 seconds")


def s18_story_questions():
    slide, n = new_slide("💡 Story Questions", "THINK", "17–21 min", "Comprehension", TEAL)
    for i, (question, _answer, emoji) in enumerate(STORY_QUESTIONS):
        col, row = i % 2, i // 2
        left = Inches(0.6 + col * 6.2)
        top = Inches(1.75 + row * 2.45)
        add_round(slide, left, top, Inches(5.9), Inches(2.2), WHITE)
        add_oval(slide, left + Inches(0.3), top + Inches(0.6), Inches(1.05), Inches(1.05), L_TEAL)
        tb(slide, left + Inches(0.3), top + Inches(0.78), Inches(1.05), Inches(0.7), emoji,
           size=30, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.6), top + Inches(0.72), Inches(4.1), Inches(0.9), question,
           size=22, bold=True, color=NAVY)
    tb(slide, Inches(0.6), Inches(6.65), Inches(9.5), Inches(0.35), "", size=10)
    notes(slide,
          "Let's think about the story. Who has the ball?",
          "Answers each question out loud, one at a time.",
          "Sam  •  Red  •  Big  •  The cat",
          "Point to the line in the story that has the answer and read it together again.",
          "About 1 minute 30 seconds")


def s19_story_answers():
    slide, n = new_slide("✅ Story Questions — Answers", "ANSWERS", "17–21 min",
                         "Answers", GREEN)
    for i, (question, answer, emoji) in enumerate(STORY_QUESTIONS):
        top = Inches(1.7 + i * 1.28)
        add_round(slide, Inches(0.6), top, Inches(12.1), Inches(1.1), L_GREEN)
        tb(slide, Inches(0.9), top + Inches(0.24), Inches(0.9), Inches(0.65), emoji, size=26,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.0), top + Inches(0.28), Inches(6.2), Inches(0.6), question, size=20,
           color=DARK)
        add_round(slide, Inches(8.6), top + Inches(0.22), Inches(3.8), Inches(0.66), GREEN)
        tb(slide, Inches(8.6), top + Inches(0.32), Inches(3.8), Inches(0.5), answer, size=22,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    notes(slide,
          "You got it! Sam has the ball, and the ball is red and big.",
          "Listens and repeats each answer in a full sentence if able.",
          "Sam  •  Red  •  Big  •  The cat",
          "Model the full sentence: 'Sam has the ball.' Then let the child repeat it.",
          "About 30 seconds")


def _pick_slide(items, title, timing, start_index):
    slide, n = new_slide(title, "GAME", timing, "Pick the word", PINK)
    for i, (emoji, stem, options, _answer) in enumerate(items):
        top = Inches(1.65 + i * 2.55)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(2.3), WHITE)
        add_oval(slide, Inches(0.85), top + Inches(0.5), Inches(1.3), Inches(1.3), L_PURPLE)
        tb(slide, Inches(0.85), top + Inches(0.72), Inches(1.3), Inches(0.9), emoji, size=38,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.45), top + Inches(0.28), Inches(4.4), Inches(0.8),
           f"{start_index + i}.  {stem}", size=26, bold=True, color=NAVY)
        for j, opt in enumerate(options):
            left = Inches(2.6 + j * 3.35)
            add_round(slide, left, top + Inches(1.25), Inches(3.0), Inches(0.8), L_SKY)
            tb(slide, left, top + Inches(1.42), Inches(3.0), Inches(0.55), opt, size=24,
               bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    return slide, n


def s20_pick_a():
    slide, n = _pick_slide(PICK[:2], "🎮 Game 3: Pick the Right Word", "21–24 min", 1)
    notes(slide,
          "Look at the picture. Which word fits in the sentence? Point to it, then read the "
          "whole sentence out loud.",
          "Chooses one of the three words, then reads the full sentence.",
          "1. CAT — I see a cat.   2. RED — The ball is red.",
          "Cover one wrong word with your hand so there are only two choices. "
          "Say the sentence with each option and ask 'Which one sounds right?'",
          "About 1 minute 30 seconds")


def s21_pick_b():
    slide, n = _pick_slide(PICK[2:], "🎮 Pick the Right Word — More!", "21–24 min", 3)
    notes(slide,
          "Two more! Look at the picture first, then choose the word.",
          "Chooses the correct word and reads the complete sentence aloud.",
          "3. RUN — The cat can run.   4. SUN — The sun is hot.",
          "Sound out each choice together, then try the sentence both ways.",
          "About 1 minute 30 seconds")


def s22_pick_answers():
    slide, n = new_slide("✅ Pick the Right Word — Answers", "ANSWERS", "21–24 min",
                         "Answers", GREEN)
    for i, (emoji, stem, _options, answer) in enumerate(PICK):
        top = Inches(1.7 + i * 1.28)
        add_round(slide, Inches(0.6), top, Inches(12.1), Inches(1.1), L_GREEN)
        tb(slide, Inches(0.9), top + Inches(0.24), Inches(0.9), Inches(0.65), emoji, size=26,
           align=PP_ALIGN.CENTER)
        full = stem.replace("___", answer.lower())
        tb(slide, Inches(2.0), top + Inches(0.26), Inches(7.0), Inches(0.65),
           full[0].upper() + full[1:], size=22, bold=True, color=NAVY)
        add_round(slide, Inches(9.4), top + Inches(0.22), Inches(3.0), Inches(0.66), GREEN)
        tb(slide, Inches(9.4), top + Inches(0.32), Inches(3.0), Inches(0.5), answer, size=22,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    notes(slide,
          "Let's read the finished sentences together. I see a cat!",
          "Reads each completed sentence with the teacher.",
          "CAT  •  RED  •  RUN  •  SUN",
          "Read it first, child echoes. Celebrate each one.",
          "About 30 seconds")


def s23_challenge():
    slide, n = new_slide("⭐ Reading Challenge", "CHALLENGE", "24–27 min", "Challenge", GOLD)
    tb(slide, Inches(0.4), Inches(1.42), Inches(10.5), Inches(0.5),
       "Pick your level. You can try them all!", size=21, bold=True, color=CORAL)
    for i, (stars, label, task, content, color, light) in enumerate(CHALLENGE):
        left = Inches(0.5 + i * 4.2)
        top = Inches(2.0)
        add_round(slide, left, top, Inches(3.95), Inches(4.4), light)
        add_rect(slide, left, top, Inches(3.95), Inches(0.12), color)
        tb(slide, left + Inches(0.15), top + Inches(0.3), Inches(3.65), Inches(0.6), stars,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.0), Inches(3.65), Inches(0.5), label,
           size=19, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.5), Inches(3.65), Inches(0.42), task,
           size=14, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), top + Inches(2.05), Inches(3.55), Inches(2.05), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(2.4), Inches(3.35), Inches(1.5), content,
           size=21, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    notes(slide,
          "Time for your challenge! Which level do you want to try first? "
          "You can do one, two, or all three!",
          "Chooses a level and reads the words or sentences out loud.",
          "Level 1: CAT, BALL.  Level 2: The ball is red.  Level 3: Sam has a big ball. "
          "The cat can run.",
          "Start at one star even if they pick three. Sound out the first word together, "
          "then let them continue. Clap after every level.",
          "About 3 minutes")


def s24_read_or_skip():
    slide, n = new_slide("🧠 Quick Game: Read or Skip", "REVIEW", "27–29 min", "Review", PURPLE)
    tb(slide, Inches(0.4), Inches(1.42), Inches(11.5), Inches(0.5),
       "Is it a word from today?  👍 READ IT      ❌ SKIP IT", size=21, bold=True, color=CORAL)
    for i, (word, _ok) in enumerate(READ_OR_SKIP):
        col, row = i % 3, i // 3
        left = Inches(0.6 + col * 4.15)
        top = Inches(2.15 + row * 2.2)
        add_round(slide, left, top, Inches(3.85), Inches(1.9), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(0.35), Inches(3.55), Inches(0.8), word,
           size=34, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.25), Inches(3.55), Inches(0.5),
           "👍  or  ❌ ?", size=18, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    notes(slide,
          "Fast game! If it is a word we learned today, read it in a big voice. "
          "If it is not, say SKIP!",
          "Looks at each word and either reads it or says 'skip'.",
          "READ: CAT, BALL, RED.  SKIP: TRAIN, ELEPHANT, MONKEY.",
          "Remind them of today's six words first: CAT, BALL, RED, BIG, RUN, SUN.",
          "About 1 minute 30 seconds")


def s25_review_answers():
    slide, n = new_slide("✅ Read or Skip — Answers", "ANSWERS", "27–29 min", "Answers", GREEN)
    for i, (word, ok) in enumerate(READ_OR_SKIP):
        col, row = i % 3, i // 3
        left = Inches(0.6 + col * 4.15)
        top = Inches(2.05 + row * 2.2)
        bg = L_GREEN if ok else L_CORAL
        badge = GREEN if ok else CORAL
        add_round(slide, left, top, Inches(3.85), Inches(1.9), bg)
        tb(slide, left + Inches(0.15), top + Inches(0.3), Inches(3.55), Inches(0.75), word,
           size=32, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.85), top + Inches(1.15), Inches(2.15), Inches(0.6), badge)
        tb(slide, left + Inches(0.85), top + Inches(1.26), Inches(2.15), Inches(0.45),
           "👍 READ" if ok else "❌ SKIP", size=17, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
    notes(slide,
          "Let's check! CAT — yes, that's ours, read it! TRAIN — we did not learn that one, skip!",
          "Confirms each answer and reads the three lesson words once more.",
          "READ: CAT, BALL, RED.  SKIP: TRAIN, ELEPHANT, MONKEY.",
          "Just read the three lesson words together again — that is the real goal here.",
          "About 30 seconds")


def s26_exit_ticket():
    slide, n = new_slide("🎟️ Exit Ticket", "FINISH", "29–30 min", "Exit ticket", SKY)
    tb(slide, Inches(0.4), Inches(1.42), Inches(11.0), Inches(0.5),
       "Before we finish — read these two things for me!", size=21, bold=True, color=CORAL)
    add_round(slide, Inches(0.7), Inches(2.05), Inches(5.8), Inches(4.3), L_SKY)
    tb(slide, Inches(0.9), Inches(2.35), Inches(5.4), Inches(0.5), "1️⃣  Read ONE word",
       size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.4), Inches(3.1), Inches(4.4), Inches(1.85), WHITE)
    tb(slide, Inches(1.4), Inches(3.55), Inches(4.4), Inches(1.0), "BALL", size=64, bold=True,
       color=CORAL, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.9), Inches(5.25), Inches(5.4), Inches(0.5), "B – A – LL", size=22,
       color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(6.85), Inches(2.05), Inches(5.8), Inches(4.3), L_TEAL)
    tb(slide, Inches(7.05), Inches(2.35), Inches(5.4), Inches(0.5), "2️⃣  Read ONE sentence",
       size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(7.25), Inches(3.1), Inches(5.0), Inches(1.85), WHITE)
    tb(slide, Inches(7.35), Inches(3.6), Inches(4.8), Inches(1.0), "The cat can run.",
       size=32, bold=True, color=TEAL, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(7.05), Inches(5.25), Inches(5.4), Inches(0.5), "Take your time 🌟",
       size=20, color=SOFT, align=PP_ALIGN.CENTER)
    notes(slide,
          "Last thing! Read this word for me... and now this sentence. Take your time, "
          "I know you can do it.",
          "Reads BALL, then reads 'The cat can run.' out loud.",
          "BALL  •  The cat can run.",
          "Sound out B-A-LL together first. For the sentence, point under each word as they read.",
          "About 1 minute")


def s27_celebration():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), YELLOW)
    for x, y, c in [(0.7, 0.75, SKY), (11.9, 0.8, PINK), (0.9, 5.8, TEAL), (11.85, 5.75, ORANGE)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(1.15), Inches(12), Inches(0.85),
       "🎉 GREAT READING! 🎉", size=44, bold=True, color=YELLOW, align=PP_ALIGN.CENTER,
       font="Georgia")
    add_oval(slide, Inches(5.42), Inches(2.25), Inches(2.5), Inches(2.5), GOLD)
    tb(slide, Inches(5.42), Inches(2.85), Inches(2.5), Inches(1.3), "⭐", size=64,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(5.05), Inches(12), Inches(0.6),
       "You are a SUPER READER!", size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Georgia")
    tb(slide, Inches(0.7), Inches(5.8), Inches(12), Inches(0.6),
       "You read words  •  You read sentences  •  You read a story",
       size=18, color=L_SKY, align=PP_ALIGN.CENTER)
    footer(slide, n, "29–30 min", "Celebrate")
    fade(slide)
    notes(slide,
          "You did it! You read words, you read sentences, and you read a whole story by "
          "yourself. Here is your star — you are a Super Reader!",
          "Celebrates, says goodbye, and takes the virtual star.",
          "Child smiles and says thank you or goodbye.",
          "Name one specific thing they did well, for example 'You sounded out BALL all by "
          "yourself!' so the praise feels real.",
          "About 30 seconds")


BUILDERS = [
    s01_title, s02_warmup_picture, s03_warmup_sound,
    lambda: make_word_slide(0, "3–7 min"),
    lambda: make_word_slide(1, "3–7 min"),
    lambda: make_word_slide(2, "3–7 min"),
    lambda: make_word_slide(3, "3–7 min"),
    lambda: make_word_slide(4, "3–7 min"),
    lambda: make_word_slide(5, "3–7 min"),
    s10_word_hunt, s11_hunt_answers,
    lambda: make_sentence_slide(0, "10–14 min"),
    lambda: make_sentence_slide(1, "10–14 min"),
    lambda: make_sentence_slide(2, "10–14 min"),
    s15_build_sentence, s16_build_answers,
    s17_story, s18_story_questions, s19_story_answers,
    s20_pick_a, s21_pick_b, s22_pick_answers,
    s23_challenge, s24_read_or_skip, s25_review_answers,
    s26_exit_ticket, s27_celebration,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

out = r"C:\Users\bhushaja\Downloads\shaip\Grade1_Reading_Class_30min.pptx"
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
