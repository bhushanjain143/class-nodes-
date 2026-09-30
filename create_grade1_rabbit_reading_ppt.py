"""Grade 1 Reading Adventure — The Little Rabbit — 30 min, 28 slides (expanded)."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

PINK = RGBColor(0xFF, 0x8A, 0xB5)
SKY = RGBColor(0x4F, 0xC3, 0xF7)
YELLOW = RGBColor(0xFF, 0xD5, 0x4F)
GREEN = RGBColor(0x81, 0xC7, 0x84)
ORANGE = RGBColor(0xFF, 0xB7, 0x4D)
CORAL = RGBColor(0xFF, 0x8A, 0x65)
PURPLE = RGBColor(0xCE, 0x93, 0xD8)
CREAM = RGBColor(0xFF, 0xFD, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
NAVY = RGBColor(0x2D, 0x3A, 0x4A)
SOFT = RGBColor(0x78, 0x90, 0x9C)
LIGHT_PINK = RGBColor(0xFF, 0xE4, 0xEC)
LIGHT_SKY = RGBColor(0xE1, 0xF5, 0xFE)
LIGHT_YELLOW = RGBColor(0xFF, 0xF9, 0xC4)
LIGHT_GREEN = RGBColor(0xE8, 0xF5, 0xE9)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
TOTAL = 29
PAGE = 0


def set_run(run, size=24, bold=False, color=NAVY, font="Calibri"):
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


def tb(slide, l, t, w, h, text, size=24, bold=False, color=NAVY, align=PP_ALIGN.CENTER, font="Calibri"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, color, font)


def bullets_left(slide, l, t, w, h, items, size=20, color=NAVY):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        r = p.add_run()
        r.text = "•  " + item
        set_run(r, size, False, color)


def notes(slide, say, ask, expected, help_str, praise, easier="", harder="", timing=""):
    text = (
        f"WHAT TO SAY:\n{say}\n\nQUESTION TO ASK:\n{ask}\n\n"
        f"EXPECTED ANSWER:\n{expected}\n\nIF CHILD STRUGGLES:\n{help_str}\n\n"
        f"ENCOURAGING PHRASES:\n{praise}\n\n"
    )
    if easier:
        text += f"EASIER OPTION:\n{easier}\n\n"
    if harder:
        text += f"HARDER OPTION:\n{harder}\n\n"
    if timing:
        text += f"TIMING: {timing}"
    slide.notes_slide.notes_text_frame.text = text


def fade(slide):
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    etree.SubElement(tr, qn("p:fade"))


def footer(slide, n, timing="", step=""):
    add_round(slide, Inches(0), Inches(7.05), prs.slide_width, Inches(0.45), NAVY)
    msg = f"Grade 1  |  🐰 Little Rabbit  |  {timing}"
    if step:
        msg += f"  |  {step}"
    tb(slide, Inches(0.3), Inches(7.1), Inches(11.5), Inches(0.35), msg, size=10, color=WHITE, align=PP_ALIGN.LEFT)
    tb(slide, Inches(12.0), Inches(7.1), Inches(1.1), Inches(0.35), f"{n}/{TOTAL}", size=10, color=WHITE)


def your_turn(slide):
    add_round(slide, Inches(0.35), Inches(6.45), Inches(2.4), Inches(0.5), YELLOW)
    tb(slide, Inches(0.35), Inches(6.48), Inches(2.4), Inches(0.45), "Your Turn! ⭐", size=14, bold=True, color=NAVY)


def new_slide():
    global PAGE
    PAGE += 1
    return prs.slides.add_slide(prs.slide_layouts[6]), PAGE


def pic_grid(title, subtitle, pics, bg, timing, step, note_extra=""):
    slide, n = new_slide()
    add_bg(slide, bg)
    tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55), title, size=30, bold=True, color=NAVY, font="Georgia")
    tb(slide, Inches(0.5), Inches(1.05), Inches(12.3), Inches(0.35), subtitle, size=20, color=SOFT)
    cols = min(len(pics), 3)
    for i, (e, word, c) in enumerate(pics):
        col = i % cols
        row = i // cols
        w = 3.9 if cols == 3 else 5.8
        left = Inches(0.55 + col * (w + 0.25))
        top = Inches(1.55 + row * 2.55)
        add_round(slide, left, top, Inches(w), Inches(2.35), WHITE)
        add_oval(slide, left + Inches(w / 2 - 0.75), top + Inches(0.25), Inches(1.5), Inches(1.5), c)
        tb(slide, left + Inches(w / 2 - 0.75), top + Inches(0.45), Inches(1.5), Inches(1.0), e, size=40)
        tb(slide, left + Inches(0.15), top + Inches(1.85), Inches(w - 0.3), Inches(0.45), word, size=20, bold=True, color=NAVY)
    footer(slide, n, timing, step)
    your_turn(slide)
    fade(slide)
    notes(slide,
          f"Look and say each word. {note_extra}",
          "What is this?",
          ", ".join(w for _, w, _ in pics),
          "First sound help. Point to mouth.",
          "Yes! / Say it with me! / Reading star!",
          "Repeat after teacher.",
          "Use in a sentence: I see a ___",
          timing)


def story_slide(part, lines, emoji, timing):
    slide, n = new_slide()
    add_bg(slide, CREAM)
    tb(slide, Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.5),
       f"📖 The Little Rabbit  (Part {part})", size=26, bold=True, color=NAVY, font="Georgia")
    add_round(slide, Inches(0.5), Inches(1.1), Inches(4.5), Inches(5.5), LIGHT_SKY)
    tb(slide, Inches(0.7), Inches(2.5), Inches(4.1), Inches(2.0), emoji, size=72)
    add_round(slide, Inches(5.3), Inches(1.1), Inches(7.5), Inches(5.5), WHITE)
    tb(slide, Inches(5.5), Inches(1.4), Inches(7.1), Inches(4.8), lines, size=28, bold=True, color=NAVY)
    tb(slide, Inches(5.5), Inches(5.8), Inches(7.1), Inches(0.45),
       "Teacher reads → You read → Together! 📖", size=16, color=PINK)
    footer(slide, n, timing, f"Story {part}")
    your_turn(slide)
    fade(slide)
    notes(slide,
          "One sentence at a time. Big praise!",
          "Read one line.",
          "Child reads with finger tracking.",
          "Echo read together.",
          "Excellent! / You read that!",
          "Repeat after me.",
          "Child reads alone.",
          timing)


# ─── BUILD SLIDES ───

# 1 Welcome
slide, n = new_slide()
add_bg(slide, LIGHT_PINK)
add_oval(slide, Inches(5.5), Inches(1.0), Inches(2.3), Inches(2.3), SKY)
tb(slide, Inches(5.5), Inches(1.55), Inches(2.3), Inches(1.2), "🐰", size=64)
tb(slide, Inches(0.5), Inches(3.5), Inches(12.3), Inches(0.9),
   "Hello, Reading Star! ⭐", size=44, bold=True, color=NAVY, font="Georgia")
tb(slide, Inches(0.5), Inches(4.5), Inches(12.3), Inches(0.5),
   "The Little Rabbit's Big Adventure", size=22, color=PINK)
footer(slide, n, "0–2 min", "Welcome")
fade(slide)
notes(slide, "Hi Reading Star! We play, read, and meet Ben the rabbit!",
      "What's your name?", "Child says name.", "Model your name.", "Excellent!", "", "", "0–2 min")

# 2 Feelings
slide, n = new_slide()
add_bg(slide, LIGHT_SKY)
tb(slide, Inches(0.5), Inches(0.5), Inches(12.3), Inches(0.6),
   "How are you feeling?", size=32, bold=True, color=NAVY, font="Georgia")
for i, (e, label, c) in enumerate([("😊", "Happy", SKY), ("😴", "Sleepy", PURPLE),
                                    ("😎", "Excited", GREEN), ("🤩", "Super excited", ORANGE)]):
    left = Inches(0.6 + i * 3.1)
    add_round(slide, left, Inches(1.8), Inches(2.85), Inches(3.5), WHITE)
    add_oval(slide, left + Inches(0.75), Inches(2.1), Inches(1.35), Inches(1.35), c)
    tb(slide, left + Inches(0.75), Inches(2.35), Inches(1.35), Inches(0.9), e, size=40)
    tb(slide, left + Inches(0.2), Inches(3.65), Inches(2.45), Inches(0.45), label, size=18, bold=True, color=NAVY)
footer(slide, n, "2–3 min", "Feelings")
your_turn(slide)
fade(slide)
notes(slide, "Point to one face!", "How do you feel?", "happy/sleepy/excited", "Model first.",
      "Great choice!", "Teacher picks first.", "Why?", "2–3 min")

# 3 Today adventure
slide, n = new_slide()
add_bg(slide, LIGHT_YELLOW)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
   "Today's Adventure! 🗺️", size=30, bold=True, color=NAVY, font="Georgia")
for i, (e, t) in enumerate([("🔍", "Play games"), ("📖", "Read a story"), ("🎤", "Tell a story"), ("⭐", "Earn stars")]):
    left = Inches(0.55 + i * 3.15)
    add_round(slide, left, Inches(1.6), Inches(2.95), Inches(3.5), WHITE)
    tb(slide, left, Inches(2.0), Inches(2.95), Inches(0.9), e, size=44)
    tb(slide, left, Inches(3.2), Inches(2.95), Inches(0.55), t, size=16, bold=True, color=NAVY)
footer(slide, n, "3 min", "Plan")
fade(slide)
notes(slide, "We will do four fun things!", "Which sounds most fun?", "any choice", "Keep it exciting.",
      "Me too!", "Point to one.", "Let's go!", "3 min")

# 4–5 Guess picture (split)
pic_grid("Guess the Picture! 🔍", "What is this?", [
    ("🐰", "Rabbit", PINK), ("🌳", "Tree", GREEN), ("🏠", "House", ORANGE),
], LIGHT_YELLOW, "3–5 min", "Pictures 1", "Repeat each word.")

pic_grid("Guess More Pictures! 🔍", "What is this?", [
    ("☀️", "Sun", YELLOW), ("🍎", "Apple", CORAL), ("🔴", "Ball", CORAL),
], LIGHT_YELLOW, "5–6 min", "Pictures 2", "Ball is in our story!")

# 6 Phonics 1
slide, n = new_slide()
add_bg(slide, LIGHT_GREEN)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55), "Phonics Fun 🔤 (1)", size=30, bold=True, color=NAVY, font="Georgia")
for i, (sounds, word, c) in enumerate([("R - A - T", "RAT", SKY), ("R - A - B - B - I - T", "RABBIT", PINK), ("R - U - N", "RUN", GREEN)]):
    top = Inches(1.6 + i * 1.15)
    add_round(slide, Inches(1.5), top, Inches(10.3), Inches(1.0), WHITE)
    add_round(slide, Inches(1.65), top + Inches(0.2), Inches(4.5), Inches(0.6), c)
    tb(slide, Inches(1.65), top + Inches(0.28), Inches(4.5), Inches(0.45), sounds, size=18, bold=True, color=WHITE)
    tb(slide, Inches(6.3), top + Inches(0.25), Inches(5.2), Inches(0.5), f"→  {word}", size=24, bold=True, color=NAVY)
footer(slide, n, "6–7 min", "Phonics")
your_turn(slide)
fade(slide)
notes(slide, "Clap the sounds!", "Say R-A-B-B-I-T", "rabbit", "Slow stretch.", "Together!", "RRR sound.", "Tap desk.", "6–7 min")

# 7 Phonics 2
slide, n = new_slide()
add_bg(slide, LIGHT_GREEN)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55), "Phonics Fun 🔤 (2)", size=30, bold=True, color=NAVY, font="Georgia")
for i, (sounds, word, c) in enumerate([("T - R - E - E", "TREE", ORANGE), ("S - U - N", "SUN", YELLOW),
                                        ("H - O - M - E", "HOME", SKY), ("B - A - L - L", "BALL", CORAL)]):
    top = Inches(1.45 + i * 1.05)
    add_round(slide, Inches(1.5), top, Inches(10.3), Inches(0.92), WHITE)
    add_round(slide, Inches(1.65), top + Inches(0.18), Inches(4.5), Inches(0.55), c)
    tb(slide, Inches(1.65), top + Inches(0.25), Inches(4.5), Inches(0.4), sounds, size=16, bold=True, color=WHITE)
    tb(slide, Inches(6.3), top + Inches(0.22), Inches(5.2), Inches(0.45), f"→  {word}", size=22, bold=True, color=NAVY)
footer(slide, n, "7–8 min", "Phonics")
your_turn(slide)
fade(slide)
notes(slide, "Story words!", "Say HOME", "home, ball", "Echo.", "Super!", "First letter.", "Sentence.", "7–8 min")

# 8–9 Feed the rabbit (2 slides)
for rnd, words, ans, timing in [
    ("Round 1", ("RUN", "SUN", "CAT"), "RUN", "8–9 min"),
    ("Round 2", ("HOME", "HAT", "HOP"), "HOME", "9 min"),
    ("Round 3", ("BALL", "BELL", "BIG"), "BALL", "9–10 min"),
]:
    slide, n = new_slide()
    add_bg(slide, LIGHT_PINK)
    tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
       f"🐰 Feed the Rabbit! {rnd}", size=28, bold=True, color=NAVY, font="Georgia")
    add_oval(slide, Inches(1.0), Inches(1.5), Inches(3.0), Inches(3.0), PINK)
    tb(slide, Inches(1.0), Inches(2.2), Inches(3.0), Inches(1.5), "🐰", size=64)
    add_round(slide, Inches(4.5), Inches(2.0), Inches(8.3), Inches(3.0), WHITE)
    tb(slide, Inches(4.7), Inches(2.3), Inches(7.9), Inches(0.4), "Which word?", size=18, color=SOFT)
    tb(slide, Inches(4.7), Inches(3.0), Inches(7.9), Inches(1.2),
       f"   {words[0]}          {words[1]}          {words[2]}", size=32, bold=True, color=NAVY)
    footer(slide, n, timing, "Feed")
    your_turn(slide)
    fade(slide)
    notes(slide, "Pick the story word!", "Which one?", ans, "Say each aloud.", "Yum!", "Two choices.", "Why pick it?", timing)

# 10 Match word to picture
slide, n = new_slide()
add_bg(slide, LIGHT_SKY)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
   "Match the Word! 🎯", size=30, bold=True, color=NAVY, font="Georgia")
pairs = [("🐰", "RABBIT"), ("🌳", "TREE"), ("🔴", "BALL"), ("☀️", "SUN")]
for i, (e, w) in enumerate(pairs):
    col, row = i % 2, i // 2
    left = Inches(0.7 + col * 6.2)
    top = Inches(1.5 + row * 2.5)
    add_round(slide, left, top, Inches(5.8), Inches(2.2), WHITE)
    tb(slide, left + Inches(0.3), top + Inches(0.5), Inches(2.0), Inches(1.0), e, size=48)
    tb(slide, left + Inches(2.8), top + Inches(0.7), Inches(2.5), Inches(0.8), w, size=28, bold=True, color=PINK)
footer(slide, n, "10–11 min", "Match")
your_turn(slide)
fade(slide)
notes(slide, "Draw a line with your finger!", "Which word goes with rabbit?", "rabbit-tree-ball-sun",
      "Say both together.", "Perfect match!", "One pair at a time.", "Make a sentence.", "10–11 min")

# 11–15 Story (5 parts — more reading time)
story_slide("1", "Ben is a little rabbit.\n\nBen lives near a big tree.", "🐰🌳", "11–12 min")
story_slide("2", "One sunny morning,\nBen goes outside.", "☀️", "12–13 min")
story_slide("3", "He sees a red ball.\n\nBen runs to the ball.", "🔴🏃", "13–14 min")
story_slide("4", "But the ball rolls away!\n\nBen follows the ball.", "🏃❓", "14–15 min")
story_slide("5", "The ball stops.\n\nBen finds the ball.\n\nBen smiles!", "🐰😊", "15–16 min")

# 16 Read word wall together
slide, n = new_slide()
add_bg(slide, LIGHT_YELLOW)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
   "Read With Me! 📖", size=30, bold=True, color=NAVY, font="Georgia")
words = ["Ben", "rabbit", "tree", "ball", "run", "sun", "home", "red"]
for i, w in enumerate(words):
    col, row = i % 4, i // 4
    left = Inches(0.5 + col * 3.15)
    top = Inches(1.5 + row * 2.2)
    add_round(slide, left, top, Inches(2.95), Inches(1.85), [PINK, GREEN, CORAL, YELLOW, SKY, ORANGE, PURPLE, CORAL][i])
    tb(slide, left, top + Inches(0.55), Inches(2.95), Inches(0.7), w.upper(), size=22, bold=True, color=WHITE)
footer(slide, n, "16–17 min", "Word wall")
your_turn(slide)
fade(slide)
notes(slide, "Point and read each word!", "Read BEN", "all story words", "Together first.", "Star for each!", "3 words only.", "Use in sentence.", "16–17 min")

# 17–18 Find the word (2 rounds)
for title, find_list, timing in [
    ("Find the Word! 🔍 (1)", ["RABBIT", "TREE", "BALL"], "17–18 min"),
    ("Find the Word! 🔍 (2)", ["SUN", "RUN", "HOME", "RED"], "18–19 min"),
]:
    slide, n = new_slide()
    add_bg(slide, LIGHT_YELLOW)
    tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55), title, size=28, bold=True, color=NAVY, font="Georgia")
    all_w = ["RABBIT", "TREE", "BALL", "SUN", "RUN", "HOME", "RED", "BEN"]
    colors = [PINK, GREEN, CORAL, YELLOW, SKY, ORANGE, CORAL, PURPLE]
    for i, (w, c) in enumerate(zip(all_w, colors)):
        col, row = i % 4, i // 4
        left = Inches(0.5 + col * 3.15)
        top = Inches(1.55 + row * 2.15)
        add_round(slide, left, top, Inches(2.95), Inches(1.85), c)
        tb(slide, left, top + Inches(0.55), Inches(2.95), Inches(0.7), w, size=20, bold=True, color=WHITE)
    tb(slide, Inches(0.5), Inches(6.2), Inches(12.3), Inches(0.35),
       "Find: " + "  •  ".join(find_list), size=16, bold=True, color=NAVY)
    footer(slide, n, timing, "Find")
    your_turn(slide)
    fade(slide)
    notes(slide, "Tap the word I say!", f"Find {find_list[0]}", "points correctly", "First letter clue.",
          "You found it!", "One word.", "Say in sentence.", timing)

# 19 Comprehension 1
slide, n = new_slide()
add_bg(slide, LIGHT_SKY)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55), "Story Questions 🧠 (1)", size=28, bold=True, color=NAVY, font="Georgia")
for i, (q, a, b, c) in enumerate([
    ("Who is the story about?", "🐰 Rabbit", "🐱 Cat", "🐶 Dog"),
    ("Where does Ben live?", "🌳 Near a tree", "🌊 In the sea", "🚗 In a car"),
]):
    top = Inches(1.4 + i * 2.2)
    add_round(slide, Inches(0.5), top, Inches(12.3), Inches(2.0), WHITE)
    tb(slide, Inches(0.7), top + Inches(0.2), Inches(12.0), Inches(0.45), q, size=20, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
    tb(slide, Inches(0.7), top + Inches(0.85), Inches(12.0), Inches(0.7), f"{a}     {b}     {c}", size=17, color=SOFT, align=PP_ALIGN.LEFT)
footer(slide, n, "19–20 min", "Q1")
your_turn(slide)
fade(slide)
notes(slide, "Point to your answer!", "Who? Where?", "Rabbit; near a tree", "Picture help.", "Yes!", "Yes/no.", "What color ball?", "19–20 min")

# 20 Comprehension 2
slide, n = new_slide()
add_bg(slide, LIGHT_SKY)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55), "Story Questions 🧠 (2)", size=28, bold=True, color=NAVY, font="Georgia")
for i, (q, a, b, c) in enumerate([
    ("What does Ben see?", "🔴 A ball", "🍎 Apple", "🚗 Car"),
    ("What happens to the ball?", "Rolls away", "Flies up", "Disappears"),
    ("Was Ben happy at the end?", "Yes 😊", "No 😢", "Not sure 🤔"),
]):
    top = Inches(1.3 + i * 1.75)
    add_round(slide, Inches(0.5), top, Inches(12.3), Inches(1.55), WHITE)
    tb(slide, Inches(0.7), top + Inches(0.15), Inches(12.0), Inches(0.4), q, size=17, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
    tb(slide, Inches(0.7), top + Inches(0.65), Inches(12.0), Inches(0.65), f"{a}     {b}     {c}", size=15, color=SOFT, align=PP_ALIGN.LEFT)
footer(slide, n, "20–22 min", "Q2")
your_turn(slide)
fade(slide)
notes(slide, "Point to answer!", "What? Ball? Happy?", "ball; rolls away; yes", "Look at story slides.", "Smart!", "Two choices.", "Tell me more.", "20–22 min")

# 21 Complete the sentence
slide, n = new_slide()
add_bg(slide, LIGHT_PINK)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
   "Complete the Sentence! 🎲", size=28, bold=True, color=NAVY, font="Georgia")
for i, (stem, a, b, c) in enumerate([
    ("The rabbit can ____.", "🏃 Run", "😴 Sleep", "🍎 Eat"),
    ("The ball is ____.", "🔴 Red", "🌳 Green", "☀️ Yellow"),
    ("Ben is a ____.", "🐰 Rabbit", "🐱 Cat", "🐶 Dog"),
]):
    top = Inches(1.35 + i * 1.65)
    add_round(slide, Inches(0.5), top, Inches(12.3), Inches(1.45), LIGHT_GREEN if i % 2 else WHITE)
    tb(slide, Inches(0.7), top + Inches(0.2), Inches(12.0), Inches(0.4), stem, size=22, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
    tb(slide, Inches(0.7), top + Inches(0.75), Inches(12.0), Inches(0.45), f"{a}     {b}     {c}", size=16, color=SOFT, align=PP_ALIGN.LEFT)
footer(slide, n, "22–23 min", "Complete")
your_turn(slide)
fade(slide)
notes(slide, "Say the whole sentence!", "The rabbit can ___?", "run; red; rabbit", "Say together first.",
      "Wonderful!", "Point to answer.", "New sentence.", "22–23 min")

# 22 Rhyme time
slide, n = new_slide()
add_bg(slide, LIGHT_GREEN)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
   "Rhyme Time! 🎵", size=30, bold=True, color=NAVY, font="Georgia")
rhymes = [("run", "☀️ sun"), ("tree", "🐝 bee"), ("ball", "🏠 hall"), ("cat", "🎩 hat")]
for i, (w, r) in enumerate(rhymes):
    left = Inches(0.55 + i * 3.15)
    add_round(slide, left, Inches(1.6), Inches(2.95), Inches(3.5), WHITE)
    tb(slide, left, Inches(2.0), Inches(2.95), Inches(0.6), w.upper(), size=28, bold=True, color=PINK)
    tb(slide, left, Inches(2.8), Inches(2.95), Inches(0.4), "rhymes with", size=14, color=SOFT)
    tb(slide, left, Inches(3.3), Inches(2.95), Inches(0.9), r, size=22, bold=True, color=NAVY)
footer(slide, n, "23–24 min", "Rhyme")
your_turn(slide)
fade(slide)
notes(slide, "Words that sound alike!", "What rhymes with run?", "sun", "Hum the sounds.", "Great ears!", "Clap syllables.", "Find another rhyme.", "23–24 min")

# 23 Brain break (seated)
slide, n = new_slide()
add_bg(slide, LIGHT_PINK)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
   "Mini Brain Break! 🐰", size=32, bold=True, color=NAVY, font="Georgia")
tb(slide, Inches(0.5), Inches(1.5), Inches(12.3), Inches(4.5),
   "1. Wiggle bunny nose 👃\n\n2. Flop bunny ears 🐰\n\n3. Stretch arms to the sun ☀️\n\n4. Take a big smile breath 😊",
   size=26, bold=True, color=NAVY)
footer(slide, n, "24 min", "Break")
fade(slide)
notes(slide, "Quick seated wiggle!", "Can you do all four?", "follows moves", "Do with child.", "Fun!", "One move.", "Favorite move?", "24 min")

# 24 Storytelling
slide, n = new_slide()
add_bg(slide, LIGHT_GREEN)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
   "Tell Me a Story! 🎨", size=32, bold=True, color=NAVY, font="Georgia")
for i, (e, label) in enumerate([("🐰", "WHO?"), ("🌳", "WHERE?"), ("🔴", "WHAT?")]):
    left = Inches(0.7 + i * 4.1)
    add_round(slide, left, Inches(1.5), Inches(3.7), Inches(3.8), WHITE)
    tb(slide, left, Inches(1.9), Inches(3.7), Inches(1.2), e, size=56)
    tb(slide, left, Inches(3.3), Inches(3.7), Inches(0.5), label, size=22, bold=True, color=PINK)
footer(slide, n, "24–26 min", "Storytell")
your_turn(slide)
fade(slide)
notes(slide, "Your mini-story!", "Who? Where? What?", "any simple story", "Praise ideas not grammar.",
      "I love it!", "One sentence.", "What happens next?", "24–26 min")

# 25–26 Reading challenge (split)
for part, sents, timing in [
    ("(1)", ["I see a rabbit.", "The sun is hot.", "The rabbit can run."], "26–27 min"),
    ("(2)", ["I see a red ball.", "The ball is big.", "Ben can run fast."], "27–28 min"),
]:
    slide, n = new_slide()
    add_bg(slide, LIGHT_YELLOW)
    tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55),
       f"Reading Challenge ⚡ {part}", size=28, bold=True, color=NAVY, font="Georgia")
    for i, s in enumerate(sents):
        top = Inches(1.4 + i * 1.35)
        add_round(slide, Inches(1.0), top, Inches(11.3), Inches(1.15), WHITE if i % 2 == 0 else LIGHT_SKY)
        tb(slide, Inches(1.2), top + Inches(0.3), Inches(9.5), Inches(0.55), f"{i + 1}. {s}", size=24, bold=True, color=NAVY, align=PP_ALIGN.LEFT)
        tb(slide, Inches(10.5), top + Inches(0.35), Inches(1.5), Inches(0.45), "⭐" * (i + 1), size=14, color=YELLOW)
    footer(slide, n, timing, "Read")
    your_turn(slide)
    fade(slide)
    notes(slide, "Read and earn stars!", "Read each line.", "reads with help", "Together on hard words.",
          "Star for you!", "One sentence.", "All alone.", timing)

# 27 Recap
slide, n = new_slide()
add_bg(slide, LIGHT_SKY)
tb(slide, Inches(0.5), Inches(0.45), Inches(12.3), Inches(0.55), "Quick Recap 🌟", size=32, bold=True, color=NAVY, font="Georgia")
bullets_left(slide, Inches(1.0), Inches(1.4), Inches(11.3), Inches(4.5), [
    "What did we read today?",
    "Favorite word?",
    "Who was in the story?",
    "What happened to the ball?",
], size=22)
footer(slide, n, "28–29 min", "Recap")
fade(slide)
notes(slide, "Remember our adventure!", "Four questions.", "Ben; ball rolled; etc.", "Choices OK.", "Great memory!", "Nod.", "Favorite part?", "28–29 min")

# 28 Reading Star
slide, n = new_slide()
add_bg(slide, NAVY)
add_round(slide, Inches(2.0), Inches(0.9), Inches(9.3), Inches(5.8), WHITE)
tb(slide, Inches(2.3), Inches(1.3), Inches(8.7), Inches(0.7),
   "🏆 READING STAR! 🏆", size=36, bold=True, color=PINK, font="Georgia")
add_oval(slide, Inches(5.3), Inches(2.2), Inches(2.7), Inches(2.7), YELLOW)
tb(slide, Inches(5.3), Inches(2.65), Inches(2.7), Inches(1.5), "⭐", size=56)
tb(slide, Inches(2.3), Inches(5.0), Inches(8.7), Inches(1.4),
   "Great job!\nYou read words!\nYou read a story!\nYou told a story!",
   size=20, bold=True, color=NAVY)
footer(slide, n, "29–30 min", "Reward")
fade(slide)
notes(slide, "Big celebration!", "Favorite part?", "any happy answer", "End high energy.", "So proud!", "Virtual sticker.", "See you next time!", "29–30 min")

out = r"C:\Users\bhushaja\Downloads\shaip\Grade1_Little_Rabbit_Reading_30min.pptx"
try:
    prs.save(out)
except PermissionError:
    out = r"C:\Users\bhushaja\Downloads\shaip\Grade1_Little_Rabbit_Reading_30min_expanded.pptx"
    prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
