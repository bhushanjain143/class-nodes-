"""Grade 1 reading adventure - 90 minutes, 42 slides, with full teacher notes.

"Bear's Big Reading Adventure" - the child is the Reading Hero helping Bear across
six stops. Progression is letter sounds -> vowels -> blending -> CVC words ->
word families -> sight words -> matching -> sentences -> story -> comprehension,
with a new activity every 5-10 minutes and everything done while seated.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x2B, 0x2A, 0x4A)
BROWN = RGBColor(0x9B, 0x62, 0x3B)
FOREST = RGBColor(0x2E, 0x8B, 0x57)
SKY = RGBColor(0x2E, 0x7F, 0xD9)
SUNNY = RGBColor(0xF2, 0xA3, 0x1B)
GOLD = RGBColor(0xFF, 0xC5, 0x30)
BERRY = RGBColor(0xE0, 0x4F, 0x5F)
GRAPE = RGBColor(0x7B, 0x5C, 0xD6)
TEAL = RGBColor(0x00, 0x9B, 0x95)
CREAM = RGBColor(0xFD, 0xFC, 0xF7)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x22, 0x2A, 0x36)
SOFT = RGBColor(0x6B, 0x77, 0x88)
L_BROWN = RGBColor(0xF6, 0xEC, 0xE3)
L_FOREST = RGBColor(0xE2, 0xF3, 0xE9)
L_SKY = RGBColor(0xE3, 0xEF, 0xFC)
L_SUNNY = RGBColor(0xFF, 0xF2, 0xD9)
L_BERRY = RGBColor(0xFD, 0xEA, 0xEC)
L_GRAPE = RGBColor(0xEE, 0xE9, 0xFB)
L_TEAL = RGBColor(0xDD, 0xF3, 0xF2)
L_GREY = RGBColor(0xF1, 0xF4, 0xF8)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 42
_counter = {"n": 0}

# ------------------------------------------------------------------ content

MAP_STOPS = [
    ("🏠", "Bear's House", "Sounds", "0–25 min", BROWN, L_BROWN),
    ("🌳", "Forest", "Words", "25–43 min", FOREST, L_FOREST),
    ("🌉", "Reading Bridge", "Sight words", "43–63 min", SKY, L_SKY),
    ("🏕️", "Camp", "Sentences", "63–78 min", SUNNY, L_SUNNY),
    ("⭐", "Star Mountain", "Story", "78–88 min", GRAPE, L_GRAPE),
    ("🏆", "Reading Hero", "You did it!", "88–90 min", TEAL, L_TEAL),
]

PICTURE_TALK = [("🐻", "bear"), ("🌳", "tree"), ("☀️", "sun"), ("🏠", "house"), ("⭐", "star")]

SOUND_CARDS = [("🐻", "BEAR", "b", "/b/"), ("🐱", "CAT", "c", "/k/"),
               ("☀️", "SUN", "s", "/s/"), ("🐶", "DOG", "d", "/d/"),
               ("🟫", "MAT", "m", "/m/"), ("🌳", "TREE", "t", "/t/"),
               ("🐷", "PIG", "p", "/p/"), ("🎉", "FUN", "f", "/f/")]

SOUND_HUNT_A = [
    ("/b/", [("🐻", "Bear"), ("☀️", "Sun"), ("🐱", "Cat")]),
    ("/s/", [("🐶", "Dog"), ("☀️", "Sun"), ("🐷", "Pig")]),
    ("/k/", [("🐱", "Cat"), ("🐻", "Bear"), ("🐔", "Hen")]),
]
SOUND_HUNT_B = [
    ("/d/", [("🐷", "Pig"), ("🐶", "Dog"), ("☀️", "Sun")]),
    ("/p/", [("🐔", "Hen"), ("🐱", "Cat"), ("🐷", "Pig")]),
    ("/m/", [("🟫", "Mat"), ("🐻", "Bear"), ("🐶", "Dog")]),
]

VOWEL_CARDS = [("a", "🍎", "apple"), ("e", "🥚", "egg"), ("i", "🐜", "insect"),
               ("o", "🐙", "octopus"), ("u", "☂️", "umbrella")]

VOWEL_BASKET_A = [("🐱", "CAT", "/k/ /a/ /t/", ["A", "E", "I"]),
                  ("🐷", "PIG", "/p/ /i/ /g/", ["A", "I", "O"]),
                  ("☀️", "SUN", "/s/ /u/ /n/", ["E", "U", "O"])]
VOWEL_BASKET_B = [("🐔", "HEN", "/h/ /e/ /n/", ["A", "E", "I"]),
                  ("🐶", "DOG", "/d/ /o/ /g/", ["O", "U", "A"]),
                  ("🟫", "MAT", "/m/ /a/ /t/", ["A", "E", "O"])]

BLEND_DEMO = ("🐱", "CAT", ["C", "A", "T"], ["/k/", "/a/", "/t/"])

BUILD_A = [("🐱", "CAT", ["C", "A", "T"], ["/k/", "/a/", "/t/"], BERRY, L_BERRY),
           ("☀️", "SUN", ["S", "U", "N"], ["/s/", "/u/", "/n/"], SUNNY, L_SUNNY),
           ("🐷", "PIG", ["P", "I", "G"], ["/p/", "/i/", "/g/"], GRAPE, L_GRAPE)]
BUILD_B = [("🐶", "DOG", ["D", "O", "G"], ["/d/", "/o/", "/g/"], FOREST, L_FOREST),
           ("🐔", "HEN", ["H", "E", "N"], ["/h/", "/e/", "/n/"], SKY, L_SKY),
           ("🦇", "BAT", ["B", "A", "T"], ["/b/", "/a/", "/t/"], BROWN, L_BROWN)]

FAMILIES = [("-AT", ["cat", "bat", "hat", "mat"], BERRY, L_BERRY),
            ("-AN", ["can", "man", "fan"], FOREST, L_FOREST),
            ("-IG", ["pig", "big"], SKY, L_SKY),
            ("-UN", ["sun", "run", "fun"], SUNNY, L_SUNNY)]

FAMILY_GAME_A = [("🐱", "CAT"), ("🏃", "RUN"), ("🐘", "BIG"), ("🎩", "HAT")]
FAMILY_GAME_B = [("🪭", "FAN"), ("🦇", "BAT"), ("☀️", "SUN"), ("🐷", "PIG")]

SIGHT_WORDS = [("I", "I can run."), ("a", "a cat"), ("am", "I am Sam."), ("the", "the sun"),
               ("is", "It is big."), ("my", "my dog"), ("can", "I can see."),
               ("see", "I see a cat.")]

TREASURE_A = [("I", ["I", "cat", "sun"]), ("the", ["dog", "the", "run"]),
              ("can", ["can", "pig", "big"])]
TREASURE_B = [("see", ["hat", "see", "bat"]), ("my", ["my", "mat", "fun"]),
              ("am", ["is", "pen", "am"])]

MATCH_A = [("🐶", "DOG"), ("☀️", "SUN"), ("🐱", "CAT")]
MATCH_B = [("🐷", "PIG"), ("🦇", "BAT"), ("🐔", "HEN"), ("⚽", "BALL")]

READ_STEPS = [("STEP 1", "I read", "Teacher reads it out loud.", BERRY),
              ("STEP 2", "We read", "Teacher and child read together.", SUNNY),
              ("STEP 3", "You read", "Child reads it alone — if ready.", FOREST)]

SENTENCES_A = [("I am Sam.", "👦", BERRY, L_BERRY), ("I see a cat.", "🐱", SKY, L_SKY)]
SENTENCES_B = [("I see the sun.", "☀️", SUNNY, L_SUNNY),
               ("The dog is big.", "🐶", FOREST, L_FOREST)]
SENTENCES_C = [("I can run.", "🏃", GRAPE, L_GRAPE), ("My cat is big.", "🐱", BERRY, L_BERRY),
               ("The sun is hot.", "☀️", SUNNY, L_SUNNY)]

PUZZLES_A = [(["I", "see", "a", "cat"], "🐱", SKY, L_SKY),
             (["The", "dog", "is", "big"], "🐶", FOREST, L_FOREST),
             (["I", "can", "run"], "🏃", GRAPE, L_GRAPE)]
PUZZLES_B = [(["My", "cat", "is", "big"], "🐱", BERRY, L_BERRY),
             (["I", "see", "the", "sun"], "☀️", SUNNY, L_SUNNY)]

STORY = [
    ("Part 1", "🐻", ["Bear wakes up.", "Bear sees the sun.", "The sun is big and hot.",
                      "Bear puts on his hat."], SUNNY, L_SUNNY),
    ("Part 2", "🌳", ["Bear goes to the park.", "Bear sees a dog.", "The dog is big.",
                      "The dog can run."], FOREST, L_FOREST),
    ("Part 3", "⚽", ["The dog has a ball.", "Bear and the dog play.", "They run and run.",
                      "It is fun."], SKY, L_SKY),
    ("Part 4", "🏠", ["Bear is happy.", "Bear goes home.", "Bear sees a cat at home.",
                      "Bear had a fun day."], BERRY, L_BERRY),
]

STORY_PASSES = [("PASS 1", "I read", "Teacher reads. Child points along.", BERRY),
                ("PASS 2", "Echo", "Teacher reads a line, child repeats it.", SUNNY),
                ("PASS 3", "We read", "Both read the page together.", SKY),
                ("PASS 4", "You read", "Child reads 2 easy lines alone.", FOREST)]

EASY_LINES = ["Bear wakes up.", "The dog is big.", "It is fun.", "Bear is happy.",
              "Bear goes home.", "The dog can run."]

DETECTIVE_A = [("🌳", "Where does Bear go?", ["Park", "School", "Store"]),
               ("🐶", "What does Bear see?", ["Dog", "Cat", "Bird"])]
DETECTIVE_B = [("⚽", "What does the dog have?", ["Ball", "Book", "Hat"]),
               ("😀", "How does Bear feel?", ["Happy", "Sad", "Angry"])]

FINAL_WORDS = [("🐱", "CAT", BERRY, L_BERRY), ("☀️", "SUN", SUNNY, L_SUNNY),
               ("🐶", "DOG", FOREST, L_FOREST)]
FINAL_SENTENCE = "I SEE A CAT."

CAN_READ = [("🔊", "Letter sounds", "b c d f m p s t"), ("🍎", "Vowels", "a e i o u"),
            ("🧩", "Blending", "/k/ /a/ /t/ → CAT"), ("📦", "Words", "cat dog sun pig hen"),
            ("🏠", "Word families", "-at  -an  -ig  -un"),
            ("💎", "Sight words", "I  a  am  the  is  my  can  see"),
            ("📕", "Sentences", "I see a cat."), ("📖", "A story", "Bear's Sunny Day"),
            ("🕵️", "Questions", "I answered them all!")]

LEVELS = [("LEVEL 1", "TEACHER SUPPORT", "Teacher models the word first.",
           "\"This word is cat. Your turn.\"", BERRY),
          ("LEVEL 2", "SHARED READING", "Teacher and child read it together.",
           "\"Let's say it together: cat.\"", SUNNY),
          ("LEVEL 3", "INDEPENDENT", "Child reads it alone.",
           "\"Your turn — you read it!\"", FOREST)]

SOUND_IT_OUT = [("1", "First sound", "\"Let's look at the first sound. /c/\""),
                ("2", "Next sound", "\"Now the next sound. /a/\""),
                ("3", "Last sound", "\"Now the last sound. /t/\""),
                ("4", "Put together", "\"Let's put them together. C-A-T.\""),
                ("5", "Try again", "\"CAT! Now you try it.\"")]

RUBRIC_SKILLS = ["Letter sounds", "Vowel sounds", "Blending", "CVC words", "Word families",
                 "Sight words", "Sentence reading", "Story reading", "Comprehension",
                 "Confidence"]
RUBRIC_NOTES = ["Words the child read independently", "Words requiring teacher support",
                "Main difficulty noticed", "Recommended next lesson"]

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


def notes(slide, says, ask, expected, struggles, doing_well, praise, timing):
    """Teacher speaker notes, same seven fields on every activity slide."""
    tf = slide.notes_slide.notes_text_frame
    tf.clear()
    blocks = [("TEACHER SAYS", says), ("ASK", ask), ("EXPECTED ANSWER", expected),
              ("IF CHILD STRUGGLES", struggles), ("IF CHILD IS DOING WELL", doing_well),
              ("PRAISE", praise), ("TIME", timing)]
    first = True
    for label, body in blocks:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        r = p.add_run()
        r.text = label
        r.font.bold = True
        r.font.size = Pt(11)
        p2 = tf.add_paragraph()
        r2 = p2.add_run()
        r2.text = body
        r2.font.size = Pt(11)
        tf.add_paragraph()


def footer(slide, n, timing="", stop=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE4, 0xEA, 0xF0))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL), Inches(0.08),
             GOLD)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), INK)
    msg = "🐻 Bear's Big Reading Adventure  |  Grade 1  |  90 min"
    if stop:
        msg += f"  |  {stop}"
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
    add_round(slide, Inches(10.3), Inches(0.26), Inches(2.65), Inches(0.42), BERRY)
    tb(slide, Inches(10.3), Inches(0.31), Inches(2.65), Inches(0.34), text, size=12,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def new_slide(title, tag, timing, stop, accent=TEAL, bg=CREAM):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, accent)
    chip(slide, tag, accent)
    if timing:
        time_chip(slide, timing)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title, size=28, bold=True,
       color=INK, font="Georgia")
    footer(slide, n, timing, stop)
    fade(slide)
    return slide, n


def one_task(slide, text, color=BERRY, top=1.34):
    """Each slide states its single job in words the child understands."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text, size=16,
              bold=True, color=color)


def helper_strip(slide, top=6.42, text="", color=L_TEAL):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), color)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(11.7), Inches(0.34), text, size=14,
       bold=True, color=INK)


def ido_wedo_youdo(slide, top=6.42):
    steps = [("I DO", "Teacher", BERRY), ("WE DO", "Together", SUNNY),
             ("YOU DO", "You!", FOREST)]
    for i, (label, who, color) in enumerate(steps):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(top), Inches(3.9), Inches(0.46), color)
        tb(slide, left, Inches(top + 0.05), Inches(3.9), Inches(0.36), f"{label}  ·  {who}",
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def picture_choice_rows(slide, items, start_index, accent, light, top_start=1.9):
    """One question per row: the sound, then three big picture buttons."""
    for i, (sound, options) in enumerate(items):
        top = Inches(top_start + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.42), Inches(0.5), Inches(0.5), accent)
        tb(slide, Inches(0.75), top + Inches(0.48), Inches(0.5), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.45), top + Inches(0.34), Inches(2.6), Inches(0.68), sound,
           size=34, bold=True, color=accent, font="Arial Black")
        for j, (emoji, label) in enumerate(options):
            left = Inches(4.5 + j * 2.75)
            add_round(slide, left, top + Inches(0.16), Inches(2.5), Inches(1.02), light)
            tb(slide, left, top + Inches(0.2), Inches(2.5), Inches(0.58), emoji, size=26,
               align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.78), Inches(2.5), Inches(0.36), label, size=15,
               bold=True, color=INK, align=PP_ALIGN.CENTER)


def vowel_rows(slide, items, start_index, top_start=1.9):
    for i, (emoji, word, sounds, options) in enumerate(items):
        top = Inches(top_start + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.42), Inches(0.5), Inches(0.5), GRAPE)
        tb(slide, Inches(0.75), top + Inches(0.48), Inches(0.5), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.4), top + Inches(0.32), Inches(0.9), Inches(0.7), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.45), top + Inches(0.2), Inches(2.4), Inches(0.6), word, size=28,
           bold=True, color=INK, font="Arial Black")
        tb(slide, Inches(2.45), top + Inches(0.82), Inches(2.4), Inches(0.4), sounds,
           size=15, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(5.5 + j * 2.4)
            add_round(slide, left, top + Inches(0.3), Inches(2.2), Inches(0.72), L_GRAPE)
            tb(slide, left, top + Inches(0.39), Inches(2.2), Inches(0.55), opt, size=26,
               bold=True, color=GRAPE, align=PP_ALIGN.CENTER, font="Arial Black")


def build_rows(slide, items, top_start=1.9, gap=1.52):
    for i, (emoji, word, letters, sounds, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), light)
        tb(slide, Inches(0.8), top + Inches(0.32), Inches(0.9), Inches(0.68), emoji,
           size=30, align=PP_ALIGN.CENTER)
        for j, letter in enumerate(letters):
            left = Inches(2.0 + j * 1.35)
            add_round(slide, left, top + Inches(0.25), Inches(1.15), Inches(0.85), WHITE)
            tb(slide, left, top + Inches(0.34), Inches(1.15), Inches(0.62), letter, size=30,
               bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
            if j < len(letters) - 1:
                tb(slide, left + Inches(1.15), top + Inches(0.48), Inches(0.18),
                   Inches(0.45), "+", size=16, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.3), top + Inches(0.46), Inches(2.4), Inches(0.5),
           "  ".join(sounds), size=16, bold=True, color=SOFT)
        tb(slide, Inches(8.8), top + Inches(0.42), Inches(0.5), Inches(0.5), "→", size=22,
           bold=True, color=color)
        add_round(slide, Inches(9.5), top + Inches(0.25), Inches(3.1), Inches(0.85), WHITE)
        tb(slide, Inches(9.5), top + Inches(0.34), Inches(3.1), Inches(0.62), word,
           size=30, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")


def house_rows(slide, items, start_index, top_start=1.95):
    for i, (emoji, word) in enumerate(items):
        top = Inches(top_start + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), WHITE)
        add_oval(slide, Inches(0.72), top + Inches(0.24), Inches(0.5), Inches(0.5), FOREST)
        tb(slide, Inches(0.72), top + Inches(0.3), Inches(0.5), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.4), top + Inches(0.18), Inches(0.9), Inches(0.62), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.4), top + Inches(0.16), Inches(2.2), Inches(0.62), word,
           size=28, bold=True, color=INK, font="Arial Black")
        for j, (name, _words, color, light) in enumerate(FAMILIES):
            left = Inches(5.0 + j * 1.98)
            add_round(slide, left, top + Inches(0.2), Inches(1.8), Inches(0.58), light)
            tb(slide, left, top + Inches(0.27), Inches(1.8), Inches(0.44), f"🏠 {name}",
               size=15, bold=True, color=color, align=PP_ALIGN.CENTER)


def treasure_rows(slide, items, start_index, top_start=1.95):
    for i, (target, options) in enumerate(items):
        top = Inches(top_start + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.42), Inches(0.5), Inches(0.5), SKY)
        tb(slide, Inches(0.75), top + Inches(0.48), Inches(0.5), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.4), top + Inches(0.25), Inches(1.5), Inches(0.4), "FIND:",
           size=14, bold=True, color=SOFT)
        add_round(slide, Inches(1.4), top + Inches(0.62), Inches(2.4), Inches(0.6), SKY)
        tb(slide, Inches(1.4), top + Inches(0.7), Inches(2.4), Inches(0.45), target,
           size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(4.3 + j * 2.85)
            add_round(slide, left, top + Inches(0.34), Inches(2.6), Inches(0.68), L_SKY)
            tb(slide, left, top + Inches(0.42), Inches(2.6), Inches(0.52), opt, size=24,
               bold=True, color=INK, align=PP_ALIGN.CENTER, font="Arial Black")


def sentence_rows(slide, items, top_start=1.95, gap=1.72, height=1.5, size=40):
    for i, (text, emoji, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), light)
        tb(slide, Inches(0.85), top + Inches(height / 2 - 0.45), Inches(1.4), Inches(0.9),
           emoji, size=40, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(height / 2 - 0.42), Inches(7.6), Inches(0.85),
           text, size=size, bold=True, color=INK)
        add_round(slide, Inches(10.4), top + Inches(height / 2 - 0.28), Inches(2.2),
                  Inches(0.56), WHITE)
        tb(slide, Inches(10.4), top + Inches(height / 2 - 0.2), Inches(2.2), Inches(0.4),
           "I · WE · YOU", size=13, bold=True, color=color, align=PP_ALIGN.CENTER)


def puzzle_rows(slide, items, start_index, top_start=1.95, gap=1.5):
    for i, (cards, emoji, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3), WHITE)
        add_oval(slide, Inches(0.72), top + Inches(0.4), Inches(0.5), Inches(0.5), color)
        tb(slide, Inches(0.72), top + Inches(0.46), Inches(0.5), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.35), top + Inches(0.34), Inches(0.9), Inches(0.66), emoji,
           size=28, align=PP_ALIGN.CENTER)
        for j, card in enumerate(cards):
            left = Inches(2.45 + j * 1.85)
            add_round(slide, left, top + Inches(0.3), Inches(1.7), Inches(0.7), light)
            tb(slide, left, top + Inches(0.39), Inches(1.7), Inches(0.52), card, size=22,
               bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.5), top + Inches(1.02), Inches(9.6), Inches(0.24),
           "Put them in order, then read it out loud.", size=11, color=SOFT)


def question_rows(slide, items, start_index, top_start=1.95):
    for i, (emoji, question, options) in enumerate(items):
        top = Inches(top_start + i * 2.16)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.92), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.7), Inches(0.52), Inches(0.52), GRAPE)
        tb(slide, Inches(0.75), top + Inches(0.76), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.6), Inches(1.0), Inches(0.72), emoji,
           size=32, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.7), top + Inches(0.64), Inches(3.4), Inches(0.66), question,
           size=22, bold=True, color=INK)
        for j, opt in enumerate(options):
            left = Inches(6.3 + j * 2.2)
            add_round(slide, left, top + Inches(0.5), Inches(2.0), Inches(0.92), L_GRAPE)
            tb(slide, left, top + Inches(0.68), Inches(2.0), Inches(0.6),
               f"{chr(65 + j)}. {opt}", size=19, bold=True, color=INK,
               align=PP_ALIGN.CENTER)


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    for x, y, c in [(0.6, 3.5, FOREST), (11.95, 3.4, SUNNY)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.95), Inches(0.95), c)
    tb(slide, Inches(0.7), Inches(0.95), Inches(12), Inches(1.15), "🐻", size=60,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(2.05), Inches(12), Inches(0.9),
       "Bear's Big Reading Adventure", size=44, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(3.0), Inches(12), Inches(0.5),
       "Grade 1  •  90 Minutes  •  🌈 A Fun Reading Adventure", size=18,
       color=RGBColor(0xCF, 0xD6, 0xEA), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.3), Inches(3.75), Inches(6.7), Inches(1.2), FOREST)
    tb(slide, Inches(3.5), Inches(4.02), Inches(6.3), Inches(0.75),
       "You are the READING HERO.\nBear needs your help!", size=19, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    for i, (emoji, place, _job, _t, color, _l) in enumerate(MAP_STOPS):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(5.35), Inches(1.65), Inches(0.9), color)
        tb(slide, left, Inches(5.44), Inches(1.65), Inches(0.42), emoji, size=17,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(5.88), Inches(1.65), Inches(0.32), place, size=9, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Bear's House")
    fade(slide)
    notes(slide,
          "Welcome! Today you and I are going on an adventure with Bear. Bear needs a "
          "Reading Hero, and that is going to be you.",
          "Are you ready to help Bear?",
          "Yes! (Any excited response is perfect.)",
          "If the child is shy, point at the six stops and say: we only do one at a time, "
          "and I will help you at every single one.",
          "Ask which stop looks the most fun and start building excitement for it.",
          "\"I already know you are going to be a great Reading Hero.\"",
          "1 minute")


def s02_meet_bear():
    slide, n = new_slide("🐻 Meet Bear", "MEET", "0–7 min", "Bear's House", BROWN)
    one_task(slide, "Say hi to Bear!", BROWN)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.6), L_BROWN)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.4), Inches(1.8), "🐻", size=96,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.4), Inches(0.7), "BEAR", size=36, bold=True,
       color=BROWN, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.8), Inches(0.6), "Your reading buddy today",
       size=15, color=SOFT, align=PP_ALIGN.CENTER)
    lines = [("💬", "\"Hi! I am Bear. I love stories.\""),
             ("😟", "\"But some words are tricky for me.\""),
             ("🤝", "\"Will you read them with me?\""),
             ("⭐", "\"You get a star after every mission!\"")]
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.9 + i * 1.18)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(1.02), L_SUNNY)
        add_oval(slide, Inches(6.5), top + Inches(0.22), Inches(0.58), Inches(0.58), WHITE)
        tb(slide, Inches(6.5), top + Inches(0.28), Inches(0.58), Inches(0.45), icon,
           size=16, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.3), top + Inches(0.28), Inches(5.3), Inches(0.5), line, size=16,
           bold=True, color=INK)
    notes(slide,
          "This is Bear. Bear loves stories, but some words are tricky for him. He needs "
          "someone brave to read with him.",
          "Can you say hi to Bear?",
          "Hi Bear! (Or a wave.)",
          "If the child will not speak yet, wave to Bear yourself and let the child copy "
          "the wave. Do not push for words.",
          "Ask what the child thinks Bear's favourite story might be.",
          "\"Bear is already smiling because you are here.\"",
          "1–2 minutes")


def s03_reading_hero():
    slide, n = new_slide("⭐ You Are the Reading Hero", "MISSION", "0–7 min",
                         "Bear's House", GOLD)
    one_task(slide, "This is your job today.", SUNNY)
    add_round(slide, Inches(0.5), Inches(1.9), Inches(12.35), Inches(1.5), L_SUNNY)
    tb(slide, Inches(0.5), Inches(2.1), Inches(12.35), Inches(1.1),
       "⭐  Every mission you finish = one star for Bear  ⭐", size=28, bold=True,
       color=SUNNY, align=PP_ALIGN.CENTER, font="Georgia")
    promises = [("🙋", "You never read alone", "I read it first. Always."),
                ("🔁", "We can try again", "Trying again is not a mistake."),
                ("🧩", "Hard word? We break it", "/c/ /a/ /t/ → CAT"),
                ("🎉", "Every try earns praise", "Even a good guess counts.")]
    for i, (icon, title, detail) in enumerate(promises):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(3.6 + row * 1.5)
        add_round(slide, left, top, Inches(5.95), Inches(1.3), L_FOREST)
        add_oval(slide, left + Inches(0.3), top + Inches(0.35), Inches(0.62), Inches(0.62),
                 WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(0.62), Inches(0.48), icon,
           size=17, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.15), top + Inches(0.22), Inches(4.5), Inches(0.48), title,
           size=18, bold=True, color=FOREST)
        tb(slide, left + Inches(1.15), top + Inches(0.72), Inches(4.5), Inches(0.42),
           detail, size=13, color=INK)
    notes(slide,
          "Here are my four promises to you today. You never read alone, we can always try "
          "again, we break hard words into sounds, and every single try gets praise.",
          "Which promise do you like best?",
          "Any choice. The point is that the child hears there is no failing today.",
          "Read all four promises out loud slowly and point at each icon. Ask the child to "
          "repeat just the first one.",
          "Ask the child to say one of the promises back in his own words.",
          "\"You are already listening like a Reading Hero.\"",
          "1–2 minutes")


def s04_map():
    slide, n = new_slide("🗺️ Today's Adventure Map", "MAP", "0–7 min", "Bear's House", FOREST)
    one_task(slide, "Six stops. One mission at each stop.", FOREST)
    for i, (emoji, place, job, when, color, light) in enumerate(MAP_STOPS):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.4)
        add_round(slide, left, top, Inches(3.9), Inches(2.2), light)
        tb(slide, left, top + Inches(0.18), Inches(3.9), Inches(0.72), emoji, size=34,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.95), Inches(3.6), Inches(0.5), place,
           size=20, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), top + Inches(1.42), Inches(3.4), Inches(0.4), job,
           size=14, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.1), top + Inches(1.78), Inches(1.7), Inches(0.36),
                  WHITE)
        tb(slide, left + Inches(1.1), top + Inches(1.81), Inches(1.7), Inches(0.3), when,
           size=11, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    notes(slide,
          "Look at our map. We start at Bear's House, then the Forest, the Reading Bridge, "
          "Camp, Star Mountain, and at the very end you become a Reading Hero.",
          "Which stop do you want to get to the most?",
          "Any stop. Usually Star Mountain or Reading Hero.",
          "If six stops feel like a lot, cover the last three with your hand and say we "
          "only need to think about the first three right now.",
          "Ask the child to count the stops out loud from one to six.",
          "\"Great looking! You found our whole map already.\"",
          "1–2 minutes")


def s05_picture_talk():
    slide, n = new_slide("👀 Picture Talk — What Do You See?", "TALK", "0–7 min",
                         "Bear's House", TEAL)
    one_task(slide, "Just look and talk. No reading yet!", TEAL)
    for i, (emoji, label) in enumerate(PICTURE_TALK):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(3.0), L_TEAL)
        tb(slide, left, Inches(2.3), Inches(2.3), Inches(1.2), emoji, size=54,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(3.75), Inches(1.8), Inches(0.75), WHITE)
        tb(slide, left + Inches(0.25), Inches(3.9), Inches(1.8), Inches(0.5), label,
           size=20, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    qs = ["🗣️ What do you see?", "⭐ What is your favorite?", "🐻 Can you find the bear?"]
    for i, q in enumerate(qs):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(5.25), Inches(3.9), Inches(0.95), L_SUNNY)
        tb(slide, left + Inches(0.2), Inches(5.48), Inches(3.5), Inches(0.5), q, size=16,
           bold=True, color=INK, align=PP_ALIGN.CENTER)
    notes(slide,
          "Before we read anything, let's just talk about what we can see. There are no "
          "wrong answers here.",
          "What do you see? What is your favorite one? Can you find the bear?",
          "Bear, tree, sun, house, star — in any order, one word or a full sentence.",
          "If the child says nothing, name one picture yourself and ask him to point to a "
          "different one. Pointing counts as an answer.",
          "Ask him to say a whole sentence: \"I see a sun.\" This previews slide 25.",
          "\"You are noticing so much. That is what good readers do.\"",
          "3–4 minutes")


def s06_sound_warmup():
    slide, n = new_slide("🔊 Sound Warm-Up — Our 8 Sounds", "LEARN", "7–17 min",
                         "Bear's House", BERRY)
    one_task(slide, "I say the sound. You say it back to me.", BERRY)
    for i, (emoji, word, letter, sound) in enumerate(SOUND_CARDS):
        col, row = i % 4, i // 4
        left = Inches(0.5 + col * 3.13)
        top = Inches(1.95 + row * 2.35)
        add_round(slide, left, top, Inches(2.9), Inches(2.15), L_BERRY)
        tb(slide, left, top + Inches(0.14), Inches(2.9), Inches(0.68), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(0.84), Inches(2.7), Inches(0.46), word,
           size=18, bold=True, color=INK, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.45), top + Inches(1.34), Inches(2.0), Inches(0.62),
                  WHITE)
        tb(slide, left + Inches(0.45), top + Inches(1.44), Inches(2.0), Inches(0.46),
           f"{letter}  says  {sound}", size=15, bold=True, color=BERRY,
           align=PP_ALIGN.CENTER)
    notes(slide,
          "Here are our eight sounds for today. Bear says: b says /b/, like bear. Cat starts "
          "with /k/. Sun starts with /s/. Go through all eight, pointing at each picture.",
          "Can you say /b/? Can you say /s/?",
          "The child repeats each sound after you.",
          "Say the sound on its own, not the letter name. Stretch it: /mmmm/. Let the child "
          "feel his lips. Do only four sounds if eight is too many.",
          "Ask him to think of another word that starts with /b/.",
          "\"Your mouth is making that sound perfectly.\"",
          "4–5 minutes")


def s07_sound_hunt_a():
    slide, n = new_slide("🔎 Game: Bear's Sound Hunt", "GAME", "7–17 min", "Bear's House",
                         BERRY)
    one_task(slide, "Which picture STARTS with this sound?", BERRY)
    picture_choice_rows(slide, SOUND_HUNT_A, 1, BERRY, L_BERRY)
    helper_strip(slide, 6.45, "Say each picture name out loud first. Stretch the first sound.",
                 L_BERRY)
    notes(slide,
          "Bear is hunting for sounds. I will give you a sound, and you point to the picture "
          "that starts with it. Round one: /b/.",
          "Which picture starts with /b/? /s/? /k/?",
          "1. Bear  2. Sun  3. Cat",
          "Say all three picture names out loud yourself, stretching the first sound: "
          "bbbbear, sssun, cccat. Then ask again. Never say \"wrong\" — say \"let's listen "
          "to that one again.\"",
          "Ask him to make up a fourth picture that would also start with that sound.",
          "\"Your ears caught that sound so fast!\"",
          "4 minutes")


def s08_sound_hunt_b():
    slide, n = new_slide("🔎 Sound Challenge — Round 2", "GAME", "7–17 min", "Bear's House",
                         BERRY)
    one_task(slide, "Three more sounds. You can do it!", BERRY)
    picture_choice_rows(slide, SOUND_HUNT_B, 4, BERRY, L_BERRY)
    helper_strip(slide, 6.45, "Our sounds:  b  c  d  f  m  p  s  t", L_BERRY)
    notes(slide,
          "Three more for Bear. Listen carefully to the very first sound of each word.",
          "Which picture starts with /d/? /p/? /m/?",
          "4. Dog  5. Pig  6. Mat",
          "Cover two of the three pictures with your hand so only two choices remain. "
          "Two choices is much easier than three.",
          "Ask him to say the whole word and then just the first sound on its own.",
          "\"That is six sounds found. Bear gets a star!\"",
          "4 minutes")


def s09_vowel_intro():
    slide, n = new_slide("🍎 The Five Vowels", "LEARN", "17–25 min", "Bear's House", GRAPE)
    one_task(slide, "These five sounds live inside almost every word.", GRAPE)
    for i, (letter, emoji, word) in enumerate(VOWEL_CARDS):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(3.5), L_GRAPE)
        add_round(slide, left + Inches(0.55), Inches(2.2), Inches(1.2), Inches(1.1), GRAPE)
        tb(slide, left + Inches(0.55), Inches(2.34), Inches(1.2), Inches(0.8), letter,
           size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left, Inches(3.5), Inches(2.3), Inches(1.0), emoji, size=40,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(4.6), Inches(1.9), Inches(0.6), WHITE)
        tb(slide, left + Inches(0.2), Inches(4.72), Inches(1.9), Inches(0.42), word,
           size=16, bold=True, color=GRAPE, align=PP_ALIGN.CENTER)
    helper_strip(slide, 5.75,
                 "a · apple    e · egg    i · insect    o · octopus    u · umbrella",
                 L_GRAPE)
    notes(slide,
          "These five letters are special. They are called vowels, and one of them hides "
          "inside almost every word. A says /a/ like apple. E says /e/ like egg.",
          "Can you say /a/? What do you see in the apple picture?",
          "The child repeats each vowel sound.",
          "Do not use the word \"vowel\" if it confuses him. Just call them \"the five "
          "special sounds.\" Teach only a, i and u today if five is too many.",
          "Ask which vowel he can hear in his own name.",
          "\"You said all five! Those are the trickiest sounds in English.\"",
          "3–4 minutes")


def s10_vowel_basket_a():
    slide, n = new_slide("🎯 Game: Vowel Basket", "GAME", "17–25 min", "Bear's House", GRAPE)
    one_task(slide, "What vowel sound do you hear in the MIDDLE?", GRAPE)
    vowel_rows(slide, VOWEL_BASKET_A, 1)
    helper_strip(slide, 6.45, "Say the word slowly and hold the middle sound: c-aaa-t.",
                 L_GRAPE)
    notes(slide,
          "Every word has a sound hiding in the middle. Listen: c-aaa-t. The middle sound "
          "is /a/. Now you find the middle sound.",
          "What vowel do you hear in CAT? In PIG? In SUN?",
          "1. A  2. I  3. U",
          "Say the word very slowly and stretch only the middle sound. Then say just that "
          "sound on its own and ask which letter it matches.",
          "Ask him to say another word with the same middle sound.",
          "\"You heard the sound right in the middle. That is hard!\"",
          "3 minutes")


def s11_vowel_basket_b():
    slide, n = new_slide("🎯 Vowel Challenge — Round 2", "GAME", "17–25 min",
                         "Bear's House", GRAPE)
    one_task(slide, "Three more middle sounds.", GRAPE)
    vowel_rows(slide, VOWEL_BASKET_B, 4)
    helper_strip(slide, 6.45, "Vowels:  a  e  i  o  u", L_GRAPE)
    notes(slide,
          "Bear's basket needs three more vowels. Same game: listen for the middle.",
          "What vowel do you hear in HEN? In DOG? In MAT?",
          "4. E  5. O  6. A",
          "If two choices sound the same to him, model both clearly: /e/ ... /i/. Ask which "
          "one matches the word. Accept a pointed finger, not just a spoken answer.",
          "Ask him to change the vowel: what happens if MAT becomes MET?",
          "\"Bear's basket is full. Another star for you!\"",
          "3 minutes")


def s12_blending_intro():
    slide, n = new_slide("🧩 Blending — Pushing Sounds Together", "LEARN", "25–35 min",
                         "Forest", FOREST)
    one_task(slide, "Three sounds can become one word.", FOREST)
    emoji, word, letters, sounds = BLEND_DEMO
    add_round(slide, Inches(0.5), Inches(1.9), Inches(12.35), Inches(2.0), L_FOREST)
    tb(slide, Inches(0.9), Inches(2.5), Inches(1.2), Inches(0.9), emoji, size=44,
       align=PP_ALIGN.CENTER)
    for j, letter in enumerate(letters):
        left = Inches(2.6 + j * 1.9)
        add_round(slide, left, Inches(2.2), Inches(1.6), Inches(1.4), WHITE)
        tb(slide, left, Inches(2.36), Inches(1.6), Inches(0.88), letter, size=48, bold=True,
           color=FOREST, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left, Inches(3.26), Inches(1.6), Inches(0.32), sounds[j], size=13,
           color=SOFT, align=PP_ALIGN.CENTER)
        if j < 2:
            tb(slide, left + Inches(1.6), Inches(2.62), Inches(0.28), Inches(0.5), "+",
               size=20, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(8.15), Inches(2.6), Inches(0.6), Inches(0.6), "→", size=28, bold=True,
       color=FOREST)
    add_round(slide, Inches(8.95), Inches(2.2), Inches(3.6), Inches(1.4), WHITE)
    tb(slide, Inches(8.95), Inches(2.5), Inches(3.6), Inches(0.9), word, size=48, bold=True,
       color=FOREST, align=PP_ALIGN.CENTER, font="Arial Black")
    steps = [("I DO", "\"/k/  /a/  /t/\"", BERRY, L_BERRY),
             ("WE DO", "\"/k/ /a/ /t/ — CAT\"", SUNNY, L_SUNNY),
             ("YOU DO", "\"CAT!\"", FOREST, L_FOREST)]
    for i, (label, line, color, light) in enumerate(steps):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(4.15), Inches(3.9), Inches(2.1), light)
        add_round(slide, left + Inches(1.1), Inches(4.38), Inches(1.7), Inches(0.5), color)
        tb(slide, left + Inches(1.1), Inches(4.46), Inches(1.7), Inches(0.36), label,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(5.05), Inches(3.3), Inches(1.0), WHITE)
        tb(slide, left + Inches(0.4), Inches(5.3), Inches(3.1), Inches(0.55), line, size=18,
           bold=True, color=color, align=PP_ALIGN.CENTER)
    notes(slide,
          "Watch me. I say each sound on its own: /k/ /a/ /t/. Now I push them together "
          "fast: CAT. Now let's do it together. Now you try it.",
          "What word do /k/ /a/ /t/ make?",
          "CAT",
          "Slide your finger under the letters as you say the sounds, then sweep it fast "
          "for the whole word. The finger sweep is what makes blending click.",
          "Ask him to blend a new word you say out loud: /m/ /a/ /t/.",
          "\"You pushed the sounds together! That is called blending.\"",
          "4 minutes")


def s13_build_a():
    slide, n = new_slide("🧩 Game: Build Bear's Word", "GAME", "25–35 min", "Forest", FOREST)
    one_task(slide, "Say each sound. Then say the whole word fast.", FOREST)
    build_rows(slide, BUILD_A)
    helper_strip(slide, 6.5, "Slide your finger under the letters as he blends.", L_FOREST)
    notes(slide,
          "Bear needs three words. For each one, say the sounds, then push them together.",
          "What sounds do you see? Now what is the word?",
          "1. CAT  2. SUN  3. PIG",
          "Cover the last letter so he only blends two sounds first: /k/ /a/. Then uncover "
          "the /t/ and blend all three.",
          "Ask him to find the vowel in the middle of each word.",
          "\"Three words built! Bear is so proud.\"",
          "5 minutes")


def s14_build_b():
    slide, n = new_slide("🧩 CVC Word Challenge", "GAME", "25–35 min", "Forest", FOREST)
    one_task(slide, "Three more words for Bear.", FOREST)
    build_rows(slide, BUILD_B)
    helper_strip(slide, 6.5, "Stuck? Say the first sound for him, then let him finish.",
                 L_FOREST)
    notes(slide,
          "Three more. These have different vowels in the middle, so listen carefully.",
          "Say the sounds, then say the word.",
          "4. DOG  5. HEN  6. BAT",
          "Give him the first sound out loud and let him do the rest. Handing him the start "
          "is not cheating — it keeps him moving forward.",
          "Ask him to change one letter: what does BAT become if we swap B for C?",
          "\"Six words today. You are really reading now.\"",
          "5 minutes")


def s15_family_intro():
    slide, n = new_slide("🏠 Word Families — Words That Rhyme", "LEARN", "35–43 min",
                         "Forest", SKY)
    one_task(slide, "These words have a part that sounds the same.", SKY)
    for i, (name, words, color, light) in enumerate(FAMILIES):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(4.3), light)
        tb(slide, left, Inches(2.1), Inches(2.9), Inches(0.7), "🏠", size=28,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.55), Inches(2.82), Inches(1.8), Inches(0.72), color)
        tb(slide, left + Inches(0.55), Inches(2.94), Inches(1.8), Inches(0.5), name,
           size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, word in enumerate(words):
            top = Inches(3.72 + j * 0.6)
            add_round(slide, left + Inches(0.35), top, Inches(2.2), Inches(0.5), WHITE)
            tb(slide, left + Inches(0.35), top + Inches(0.06), Inches(2.2), Inches(0.38),
               word, size=19, bold=True, color=color, align=PP_ALIGN.CENTER)
    notes(slide,
          "Look at these houses. Every word in the first house ends the same way: cat, bat, "
          "hat, mat. They rhyme. If you can read one, you can read all of them.",
          "Can you hear the same part in cat and hat?",
          "Yes — the -at sound.",
          "Read the whole family out loud in a rhythm and let him join in on the last word. "
          "Rhyme is easier to hear than it is to explain.",
          "Ask him to make up a silly word for one of the houses, like \"zat.\"",
          "\"You heard the rhyme! That is a big reading trick.\"",
          "4 minutes")


def s16_family_game_a():
    slide, n = new_slide("🏠 Game: Word Family House", "GAME", "35–43 min", "Forest", SKY)
    one_task(slide, "Which house does this word live in?", SKY)
    house_rows(slide, FAMILY_GAME_A, 1)
    helper_strip(slide, 6.5, "Houses:  -AT    -AN    -IG    -UN", L_SKY)
    notes(slide,
          "Each word needs to find its house. Listen to the ending of the word, not the "
          "beginning.",
          "Which house does CAT live in? RUN? BIG? HAT?",
          "1. -AT  2. -UN  3. -IG  4. -AT",
          "Cover the first letter of the word with your finger so only the ending shows. "
          "Then the match is obvious.",
          "Ask him to name another word that lives in the same house.",
          "\"You found every house. Bear did not get lost once!\"",
          "4 minutes")


def s17_family_game_b():
    slide, n = new_slide("🏠 Word Family Challenge", "GAME", "35–43 min", "Forest", SKY)
    one_task(slide, "Four more words need a home.", SKY)
    house_rows(slide, FAMILY_GAME_B, 5)
    helper_strip(slide, 6.5, "Cover the first letter. Now only the ending is left.", L_SKY)
    notes(slide,
          "Four more words are looking for their house. Same trick: listen to the ending.",
          "Which house does FAN live in? BAT? SUN? PIG?",
          "5. -AN  6. -AT  7. -UN  8. -IG",
          "Say the word and an obviously wrong house out loud so he can hear it does not "
          "rhyme. Hearing the mismatch teaches faster than being told.",
          "Ask him to read all four words in a row without help.",
          "\"Eight words housed. You are a rhyming expert.\"",
          "4 minutes")


def s18_bear_says():
    slide, n = new_slide("🐻 Brain Break: Bear Says", "BREAK", "43–47 min",
                         "Reading Bridge", SUNNY)
    one_task(slide, "Stay in your seat. Only do it if I say \"Bear says.\"", SUNNY)
    commands = [("👏", "Bear says clap 2 times."), ("👉", "Bear says point to CAT."),
                ("🔊", "Bear says say the /m/ sound."), ("✋", "Bear says touch the word SUN."),
                ("🔤", "Bear says spell DOG."), ("🤫", "Bear says whisper CAT.")]
    for i, (icon, command) in enumerate(commands):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.4)
        add_round(slide, left, top, Inches(5.95), Inches(1.2), L_SUNNY)
        add_oval(slide, left + Inches(0.28), top + Inches(0.3), Inches(0.6), Inches(0.6),
                 WHITE)
        tb(slide, left + Inches(0.28), top + Inches(0.36), Inches(0.6), Inches(0.46), icon,
           size=17, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.36), Inches(4.6), Inches(0.55),
           command, size=16, bold=True, color=INK)
    helper_strip(slide, 6.35,
                 "Four of these are reading tasks — that is the point. Keep it quick and silly.",
                 L_SUNNY)
    notes(slide,
          "Time for a quick break, but we stay in our seats. If I say \"Bear says,\" you do "
          "it. If I forget to say it, you sit still and catch me out.",
          "Bear says clap two times. Bear says point to CAT.",
          "The child performs each action; he freezes on commands without \"Bear says.\"",
          "Slow the pace right down and use only the clapping and whispering ones. Reading "
          "under time pressure is not the goal here.",
          "Speed up, and try one command without \"Bear says\" to see if he catches it.",
          "\"You caught me! And you read CAT while you were playing.\"",
          "4 minutes")


def s19_sight_intro():
    slide, n = new_slide("💎 Sight Words — Words We Just Know", "LEARN", "47–56 min",
                         "Reading Bridge", SKY)
    one_task(slide, "We do not sound these out. We just know them.", SKY)
    for i, (word, example) in enumerate(SIGHT_WORDS):
        col, row = i % 4, i // 4
        left = Inches(0.5 + col * 3.13)
        top = Inches(1.95 + row * 2.1)
        add_round(slide, left, top, Inches(2.9), Inches(1.9), L_SKY)
        add_round(slide, left + Inches(0.45), top + Inches(0.18), Inches(2.0), Inches(0.85),
                  WHITE)
        tb(slide, left + Inches(0.45), top + Inches(0.32), Inches(2.0), Inches(0.58), word,
           size=28, bold=True, color=SKY, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.1), top + Inches(1.18), Inches(2.7), Inches(0.5), example,
           size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
    helper_strip(slide, 6.25,
                 "👀 Look  →  👂 I say it  →  🗣️ You say it  →  ⚡ Read it fast", L_SKY)
    notes(slide,
          "Some words do not follow the sound rules, so we just learn them by sight. This "
          "one is \"I.\" This one is \"the.\" Show two or three at a time, not all eight.",
          "What is this word? And this one?",
          "The child names each word on sight.",
          "Cover six of the eight cards. Teach only \"I\" and \"the\" properly, then come "
          "back for more later. Two known words beat eight half-known ones.",
          "Put two together and read them as a phrase: \"I am,\" \"the sun.\"",
          "\"You just knew that word straight away. That is real reading.\"",
          "4 minutes")


def s20_treasure_a():
    slide, n = new_slide("💎 Game: Bear's Treasure", "GAME", "47–56 min", "Reading Bridge",
                         SKY)
    one_task(slide, "Bear opens the treasure when you find the word!", SKY)
    treasure_rows(slide, TREASURE_A, 1)
    helper_strip(slide, 6.5, "Say the target word out loud before he looks at the choices.",
                 L_SKY)
    notes(slide,
          "Bear's treasure chest is locked. It opens when you find the right word. First "
          "word to find: \"I.\"",
          "Can you find the word I? The word the? The word can?",
          "1. I  2. the  3. can",
          "Read all three choices out loud yourself and ask which one sounded like the word "
          "you wanted. Listening is easier than looking.",
          "Ask him to use the word in a short sentence out loud.",
          "\"The treasure is open! You found it.\"",
          "4 minutes")


def s21_treasure_b():
    slide, n = new_slide("💎 Sight Word Challenge", "GAME", "47–56 min", "Reading Bridge",
                         SKY)
    one_task(slide, "Three more treasures to unlock.", SKY)
    treasure_rows(slide, TREASURE_B, 4)
    helper_strip(slide, 6.5,
                 "Missed one? Put it back in the pile and ask again before the story.",
                 L_SKY)
    notes(slide,
          "Three more chests. These words are in our story later, so finding them now makes "
          "the story easier.",
          "Can you find see? my? am?",
          "4. see  5. my  6. am",
          "Point to the word on the previous slide's card so he can compare the shapes "
          "side by side. Matching shapes is a valid strategy at this stage.",
          "Ask him to read all eight sight words in a row, quickly.",
          "\"All the treasure is yours. Bear gets another star.\"",
          "4 minutes")


def s22_match_intro():
    slide, n = new_slide("🖼️ Word + Picture Match", "LEARN", "56–63 min", "Reading Bridge",
                         TEAL)
    one_task(slide, "Which word goes with which picture?", TEAL)
    for i, (emoji, _word) in enumerate(MATCH_A):
        left = Inches(1.4 + i * 3.6)
        add_round(slide, left, Inches(1.95), Inches(3.0), Inches(2.0), L_TEAL)
        tb(slide, left, Inches(2.2), Inches(3.0), Inches(1.4), emoji, size=54,
           align=PP_ALIGN.CENTER)
    for i in range(3):
        left = Inches(1.4 + i * 3.6)
        tb(slide, left, Inches(4.05), Inches(3.0), Inches(0.4), "⬇", size=18, bold=True,
           color=SOFT, align=PP_ALIGN.CENTER)
    shuffled = ["SUN", "CAT", "DOG"]
    for i, word in enumerate(shuffled):
        left = Inches(1.4 + i * 3.6)
        add_round(slide, left, Inches(4.55), Inches(3.0), Inches(1.0), WHITE)
        tb(slide, left, Inches(4.72), Inches(3.0), Inches(0.68), word, size=32, bold=True,
           color=TEAL, align=PP_ALIGN.CENTER, font="Arial Black")
    helper_strip(slide, 5.8, "Careful — the words are NOT under the right pictures!", L_TEAL)
    notes(slide,
          "Look carefully. The words got mixed up and they are under the wrong pictures. "
          "Let's put them right.",
          "Which word goes with the dog? With the sun? With the cat?",
          "DOG under the dog, SUN under the sun, CAT under the cat.",
          "Cover two pictures so only one picture and three words remain. Then ask which "
          "word starts with the same sound as the picture.",
          "Ask him to read the word before he matches it, not after.",
          "\"You matched them by reading, not just guessing. Well done.\"",
          "3 minutes")


def s23_match_game():
    slide, n = new_slide("🖼️ Game: Match It!", "GAME", "56–63 min", "Reading Bridge", TEAL)
    one_task(slide, "Four more. The last one has no picture — read it!", TEAL)
    for i, (emoji, word) in enumerate(MATCH_B):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(2.4), L_TEAL)
        show = emoji if i < 3 else "❓"
        tb(slide, left, Inches(2.2), Inches(2.9), Inches(1.3), show, size=48,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.35), Inches(3.55), Inches(2.2), Inches(0.6), WHITE)
        tb(slide, left + Inches(0.35), Inches(3.66), Inches(2.2), Inches(0.44),
           "picture" if i < 3 else "no picture!", size=13, bold=True,
           color=TEAL if i < 3 else BERRY, align=PP_ALIGN.CENTER)
    for i, (_emoji, word) in enumerate(MATCH_B):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(4.6), Inches(2.9), Inches(1.0), WHITE)
        tb(slide, left, Inches(4.78), Inches(2.9), Inches(0.68), word, size=30, bold=True,
           color=INK, align=PP_ALIGN.CENTER, font="Arial Black")
    helper_strip(slide, 5.85,
                 "Now we take the picture away — this is reading the word by itself.", L_TEAL)
    notes(slide,
          "Three of these have a picture to help you. The last one does not. That means you "
          "have to read the word all on your own — and I think you can.",
          "What is this word? And the one with no picture?",
          "PIG, BAT, HEN, BALL",
          "For BALL, cover the -all and show only B. Ask for the first sound, then uncover "
          "the rest and read it together.",
          "Cover a second picture and ask him to read that word unaided too.",
          "\"You read a word with no picture at all. That is huge.\"",
          "4 minutes")


def s24_sentence_intro():
    slide, n = new_slide("📕 Sentences — How We Read Together", "LEARN", "63–72 min",
                         "Camp", SUNNY)
    one_task(slide, "You never read a new sentence first. I do.", SUNNY)
    for i, (label, who, what, color) in enumerate(READ_STEPS):
        left = Inches(0.5 + i * 4.15)
        light = [L_BERRY, L_SUNNY, L_FOREST][i]
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(2.85), light)
        add_round(slide, left + Inches(1.1), Inches(2.2), Inches(1.7), Inches(0.52), color)
        tb(slide, left + Inches(1.1), Inches(2.28), Inches(1.7), Inches(0.38), label,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(2.9), Inches(3.5), Inches(0.62), who, size=26,
           bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.3), Inches(3.62), Inches(3.3), Inches(0.98), WHITE)
        tb(slide, left + Inches(0.45), Inches(3.8), Inches(3.0), Inches(0.68), what,
           size=14, color=INK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.05), Inches(12.35), Inches(1.3), L_GREY)
    tb(slide, Inches(0.8), Inches(5.18), Inches(11.7), Inches(0.4),
       "Hard word inside a sentence? We stop and break it, then go back:", size=14,
       bold=True, color=INK)
    for i, (num, step, _how) in enumerate(SOUND_IT_OUT):
        left = Inches(0.8 + i * 2.4)
        add_round(slide, left, Inches(5.66), Inches(2.2), Inches(0.52), WHITE)
        tb(slide, left, Inches(5.75), Inches(2.2), Inches(0.36), f"{num}. {step}", size=12,
           bold=True, color=SUNNY, align=PP_ALIGN.CENTER)
    notes(slide,
          "Here is how we read every sentence today. First I read it. Then we read it "
          "together. Then, only if you want to, you read it by yourself.",
          "Which step do you want to try first?",
          "Any. Most children pick step two.",
          "Stay on step one and step two for as long as he needs. Step three is optional "
          "every single time.",
          "Ask him to go straight to step three on an easy sentence.",
          "\"You get to choose. That is what heroes do.\"",
          "2 minutes")


def s25_sentence_1():
    slide, n = new_slide("📕 Read With Me", "READ", "63–72 min", "Camp", BERRY)
    one_task(slide, "Point to each word while I read it.", BERRY)
    sentence_rows(slide, SENTENCES_A, top_start=2.1, gap=2.15, height=1.9, size=44)
    ido_wedo_youdo(slide, 6.42)
    notes(slide,
          "I am going to read, and I want your finger to follow along under each word. "
          "Ready? \"I am Sam.\" Now let's read it together.",
          "Can you point to the word \"I\"? Can you read it with me?",
          "The child tracks with a finger and joins in on the second reading.",
          "Read at half speed and tap each word as you say it. If he loses his place, put "
          "your finger over his and move it together.",
          "Ask him to read the second sentence alone.",
          "\"Your finger stayed with every word. That is exactly right.\"",
          "3 minutes")


def s26_sentence_2():
    slide, n = new_slide("📗 Read With Me — More", "READ", "63–72 min", "Camp", SUNNY)
    one_task(slide, "Look at the picture first. It gives you a clue.", SUNNY)
    sentence_rows(slide, SENTENCES_B, top_start=2.1, gap=2.15, height=1.9, size=44)
    ido_wedo_youdo(slide, 6.42)
    notes(slide,
          "Before we read, look at the picture. What do you think this sentence might say? "
          "Good guess — now let's find out.",
          "What do you think it says? Now let's read it and check.",
          "\"I see the sun.\" and \"The dog is big.\"",
          "If a word blocks him, break it: /d/ /o/ /g/ — DOG. Then go back to the start of "
          "the whole sentence and read it again from the beginning.",
          "Ask him to read one sentence with no help at all.",
          "\"You used the picture like a real reader does.\"",
          "3 minutes")


def s27_sentence_3():
    slide, n = new_slide("📘 Read With Me — Your Turn Grows", "READ", "63–72 min", "Camp",
                         FOREST)
    one_task(slide, "Pick the one YOU want to read by yourself.", FOREST)
    sentence_rows(slide, SENTENCES_C, top_start=1.95, gap=1.5, height=1.3, size=34)
    ido_wedo_youdo(slide, 6.42)
    notes(slide,
          "Three sentences here. You get to choose which one you read by yourself. I will "
          "read the other two.",
          "Which one do you want to read?",
          "Any choice. \"I can run.\" is the easiest.",
          "If he will not choose, read all three yourself, then point to the shortest and "
          "say: this one is yours, and I will start it with you.",
          "Ask him to read a second one, or to make up his own sentence using \"my.\"",
          "\"You picked a hard one and you read it. Brilliant.\"",
          "3 minutes")


def s28_puzzle_a():
    slide, n = new_slide("🧩 Game: Build Bear's Sentence", "GAME", "72–78 min", "Camp",
                         SUNNY)
    one_task(slide, "The words fell over. Put them back in order!", SUNNY)
    puzzle_rows(slide, PUZZLES_A, 1)
    helper_strip(slide, 6.5, "Hint: the first word always wears a CAPITAL letter.", L_SUNNY)
    notes(slide,
          "Oh no, Bear knocked the words over. Let's put them back in the right order, then "
          "read the sentence out loud.",
          "Which word goes first? What does the whole sentence say?",
          "1. I see a cat.  2. The dog is big.  3. I can run.",
          "Hand him the first word by pointing at it, then ask only what comes next. One "
          "word at a time, not the whole sentence at once.",
          "Ask him to rebuild it a second time with the cards mixed differently.",
          "\"You built a whole sentence! Bear could not do that alone.\"",
          "4 minutes")


def s29_puzzle_b():
    slide, n = new_slide("🧩 Sentence Challenge", "GAME", "72–78 min", "Camp", SUNNY)
    one_task(slide, "Two more — then you make one of your own.", SUNNY)
    puzzle_rows(slide, PUZZLES_B, 4)
    add_round(slide, Inches(0.5), Inches(5.1), Inches(12.35), Inches(1.55), L_FOREST)
    tb(slide, Inches(0.8), Inches(5.25), Inches(11.7), Inches(0.45),
       "⭐ YOUR OWN SENTENCE — use any word you like:", size=17, bold=True, color=FOREST)
    add_round(slide, Inches(0.8), Inches(5.8), Inches(11.75), Inches(0.68), WHITE)
    tb(slide, Inches(1.1), Inches(5.92), Inches(11.2), Inches(0.5),
       "I  see  ______________ .", size=24, bold=True, color=INK)
    notes(slide,
          "Two more for Bear, and then the best part: you get to build your very own "
          "sentence with any word you want.",
          "What will your sentence be?",
          "4. My cat is big.  5. I see the sun.  Then any child-made sentence.",
          "Offer him two words to choose from for the blank, such as cat or dog, so he is "
          "picking rather than inventing from nothing.",
          "Ask him to write or spell out the word he chose for the blank.",
          "\"That sentence is yours. Nobody else made that one.\"",
          "3 minutes")


def s30_story_intro():
    slide, n = new_slide("📖 Story Time: Bear's Sunny Day", "STORY", "78–85 min",
                         "Star Mountain", GRAPE)
    one_task(slide, "Four short pages. We read each one four times.", GRAPE)
    add_round(slide, Inches(0.5), Inches(1.9), Inches(5.6), Inches(4.4), L_BROWN)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.6), Inches(1.6), "🐻", size=86,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.3), Inches(5.6), Inches(0.7), "BEAR", size=34, bold=True,
       color=BROWN, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.8), Inches(5.05), Inches(5.0), Inches(0.8),
       "Every word in this story is one you already know.", size=15, color=SOFT,
       align=PP_ALIGN.CENTER)
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
    notes(slide,
          "Now the big moment. This story is called Bear's Sunny Day, and every single word "
          "in it is a word you have already read today. Nothing new.",
          "Are you ready to read a whole story?",
          "Yes — often with some nerves.",
          "Say clearly: you will never have to read a page alone. I read first, every time. "
          "That promise removes most of the fear.",
          "Ask him to predict what happens to Bear on a sunny day.",
          "\"You have been getting ready for this all lesson.\"",
          "2 minutes")


def _story_slide(index, timing):
    part, emoji, lines, color, light = STORY[index]
    slide, n = new_slide(f"📖 Bear's Sunny Day — {part}", "STORY", timing, "Star Mountain",
                         color)
    add_round(slide, Inches(0.5), Inches(1.5), Inches(4.5), Inches(4.7), light)
    tb(slide, Inches(0.5), Inches(3.0), Inches(4.5), Inches(1.8), emoji, size=92,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(5.25), Inches(1.5), Inches(7.6), Inches(4.7), WHITE)
    for i, line in enumerate(lines):
        tb(slide, Inches(5.6), Inches(1.95 + i * 1.05), Inches(7.0), Inches(0.8), line,
           size=30, bold=True, color=INK)
    ido_wedo_youdo(slide, 6.42)
    return slide, n


def s31_story_1():
    slide, n = _story_slide(0, "78–85 min")
    notes(slide,
          "Page one. Listen first, then we read it together. \"Bear wakes up. Bear sees the "
          "sun. The sun is big and hot. Bear puts on his hat.\"",
          "What is Bear doing? Can you find the word SUN?",
          "Bear wakes up and sees the sun. SUN is on the second line.",
          "Read one line, then stop and let him echo just that line. One line at a time is "
          "plenty. Point under each word as you go.",
          "Ask him to read the shortest line, \"Bear wakes up,\" by himself.",
          "\"You read a page of a real story.\"",
          "2 minutes")


def s32_story_2():
    slide, n = _story_slide(1, "78–85 min")
    notes(slide,
          "Page two. Bear is going somewhere. \"Bear goes to the park. Bear sees a dog. The "
          "dog is big. The dog can run.\"",
          "Where does Bear go? What does he see there?",
          "The park. He sees a dog.",
          "The word \"park\" is new-ish. Say it for him rather than making him decode it, "
          "and keep the momentum of the story going.",
          "Ask him to find the word DOG twice on this page.",
          "\"You are following the story so well.\"",
          "2 minutes")


def s33_story_3():
    slide, n = _story_slide(2, "78–85 min")
    notes(slide,
          "Page three. Something fun happens. \"The dog has a ball. Bear and the dog play. "
          "They run and run. It is fun.\"",
          "What does the dog have? Do you think Bear is having a good time?",
          "A ball. Yes.",
          "\"It is fun\" is the easiest line in the whole story. If he is tired, let that "
          "be his only independent line on this page.",
          "Ask him to read \"They run and run\" with an excited voice.",
          "\"You made the story sound fun. That is expression!\"",
          "2 minutes")


def s34_story_4():
    slide, n = _story_slide(3, "78–85 min")
    notes(slide,
          "Last page. \"Bear is happy. Bear goes home. Bear sees a cat at home. Bear had a "
          "fun day.\"",
          "How does Bear feel? What does he see at home?",
          "Happy. A cat.",
          "Point out that CAT was the very first word he built today. Coming full circle is "
          "worth saying out loud — it shows him how far he came.",
          "Ask him to read the last line as a big happy ending.",
          "\"You read the whole story, all four pages.\"",
          "2 minutes")


def s35_story_challenge():
    slide, n = new_slide("⭐ Story Reading Challenge", "READ", "78–85 min", "Star Mountain",
                         FOREST)
    one_task(slide, "Pick TWO lines you want to read all by yourself.", FOREST)
    for i, line in enumerate(EASY_LINES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.42)
        add_round(slide, left, top, Inches(5.95), Inches(1.22), L_FOREST)
        add_oval(slide, left + Inches(0.3), top + Inches(0.34), Inches(0.55), Inches(0.55),
                 WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(0.55), Inches(0.42), "⭐",
           size=14, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.32), Inches(4.6), Inches(0.62), line,
           size=24, bold=True, color=INK)
    helper_strip(slide, 6.35,
                 "He chooses. If he freezes, read the first word for him and then wait.",
                 L_FOREST)
    notes(slide,
          "These are the six easiest lines from Bear's story. You choose two and read them "
          "to me by yourself. I will stay right here.",
          "Which two do you want?",
          "Any two. \"It is fun\" and \"Bear is happy\" are the easiest.",
          "Read the first word out loud for him and then go quiet. The silence gives him "
          "room. Do not fill it too quickly.",
          "Ask for a third line, or ask him to read one to a family member later.",
          "\"You just read on your own. Remember that feeling.\"",
          "3 minutes")


def s36_detective_a():
    slide, n = new_slide("🕵️ Game: Bear's Story Detective", "GAME", "85–88 min",
                         "Star Mountain", GRAPE)
    one_task(slide, "Think about Bear's day. Which answer is right?", GRAPE)
    question_rows(slide, DETECTIVE_A, 1)
    helper_strip(slide, 6.35, "Cannot remember? Flip back a page and read it together.",
                 L_GRAPE)
    notes(slide,
          "Now you are a detective. I will ask about Bear's day and you find the answer.",
          "Where does Bear go? What does Bear see?",
          "1. A — Park   2. A — Dog",
          "Go back to story page two and read it together, then ask again. Rereading to "
          "find an answer is a comprehension skill, not a failure.",
          "Ask him to answer in a full sentence: \"Bear goes to the park.\"",
          "\"You remembered the story. That means you understood it.\"",
          "2 minutes")


def s37_detective_b():
    slide, n = new_slide("🕵️ Comprehension Challenge", "GAME", "85–88 min", "Star Mountain",
                         GRAPE)
    one_task(slide, "Two more questions, detective.", GRAPE)
    question_rows(slide, DETECTIVE_B, 3)
    helper_strip(slide, 6.35,
                 "Bonus: \"What would YOU do on a sunny day?\" Any answer is right.",
                 L_GRAPE)
    notes(slide,
          "Two last detective questions, and then something just for you.",
          "What does the dog have? How does Bear feel?",
          "3. A — Ball   4. A — Happy",
          "For the feelings question, ask him to show the face rather than say the word. "
          "Then name it for him: yes, that is happy.",
          "Ask the bonus question about his own sunny day and let him talk freely.",
          "\"Four out of four, detective. Bear's mystery is solved.\"",
          "2 minutes")


def s38_final_challenge():
    slide, n = new_slide("⭐ Final Reading Challenge", "CHALLENGE", "88–89 min",
                         "Star Mountain", GOLD)
    one_task(slide, "Read these to finish the adventure. I am right here.", SUNNY)
    for i, (emoji, word, color, light) in enumerate(FINAL_WORDS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(2.3), light)
        tb(slide, left + Inches(0.25), Inches(2.35), Inches(1.2), Inches(1.0), emoji,
           size=40, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.6), Inches(2.5), Inches(2.1), Inches(1.0), WHITE)
        tb(slide, left + Inches(1.6), Inches(2.7), Inches(2.1), Inches(0.68), word,
           size=28, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.6), Inches(3.62), Inches(2.7), Inches(0.45), color)
        tb(slide, left + Inches(0.6), Inches(3.69), Inches(2.7), Inches(0.36), "Read it! ⭐",
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(4.5), Inches(12.35), Inches(1.75), L_SUNNY)
    tb(slide, Inches(0.9), Inches(4.7), Inches(1.4), Inches(1.0), "🐱", size=44,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.7), Inches(4.85), Inches(8.0), Inches(0.9), FINAL_SENTENCE, size=48,
       bold=True, color=INK, font="Arial Black")
    notes(slide,
          "Last challenge of the whole adventure. Three words and one sentence. This is not "
          "a test — I will help the second you want me to.",
          "Can you read CAT? SUN? DOG? Now the sentence.",
          "CAT, SUN, DOG, and \"I see a cat.\"",
          "Read the sentence together first, then ask him to do it alone. Doing it together "
          "first is not cheating; it is how confidence gets built.",
          "Ask him to make a new sentence by swapping cat for dog or sun.",
          "\"That is the whole adventure, finished by you.\"",
          "1–2 minutes")


def s39_what_i_can_read():
    slide, n = new_slide("✅ What I Can Read Now", "RECAP", "89–90 min", "Reading Hero",
                         FOREST)
    one_task(slide, "Look at everything you did today. Read it with me.", FOREST)
    for i, (icon, label, detail) in enumerate(CAN_READ):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 1.06)
        width = Inches(5.95) if i < 8 else Inches(12.35)
        add_round(slide, left, top, width, Inches(0.92), L_FOREST)
        add_oval(slide, left + Inches(0.26), top + Inches(0.2), Inches(0.52), Inches(0.52),
                 WHITE)
        tb(slide, left + Inches(0.26), top + Inches(0.26), Inches(0.52), Inches(0.42), icon,
           size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.95), top + Inches(0.12), Inches(4.6), Inches(0.42), label,
           size=16, bold=True, color=FOREST)
        tb(slide, left + Inches(0.95), top + Inches(0.52), Inches(10.9 if i >= 8 else 4.6),
           Inches(0.36), detail, size=12, color=INK)
    notes(slide,
          "Look at this list. Every single one of these is something you did today, in 90 "
          "minutes. Sounds, vowels, blending, words, families, sight words, sentences, a "
          "whole story, and the questions afterwards.",
          "Which one are you proudest of?",
          "Any choice. Let him take his time.",
          "Read the list to him and let him just point at his favourite. Pointing is a fine "
          "answer at the end of a long lesson.",
          "Ask him to demonstrate one item again on the spot.",
          "\"Nine things. You did nine reading jobs today.\"",
          "1 minute")


def s40_badge():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    for x, y, c in [(0.65, 1.3, FOREST), (11.95, 1.25, SKY), (0.75, 5.5, BERRY),
                    (11.9, 5.45, SUNNY)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.8), c)
    tb(slide, Inches(0.7), Inches(0.85), Inches(12), Inches(0.8),
       "🏆 READING HERO BADGE 🏆", size=34, bold=True, color=GOLD, align=PP_ALIGN.CENTER,
       font="Georgia")
    add_oval(slide, Inches(5.42), Inches(1.75), Inches(2.5), Inches(2.5), GOLD)
    tb(slide, Inches(5.42), Inches(2.3), Inches(2.5), Inches(1.4), "🐻", size=52,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.5), Inches(12), Inches(0.7), "🎉 MISSION COMPLETE!",
       size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(5.25), Inches(12), Inches(0.5),
       "You are a Reading Hero. Bear says thank you!", size=21,
       color=RGBColor(0xCF, 0xD6, 0xEA), align=PP_ALIGN.CENTER)
    for i, (emoji, _place, _job, _t, color, _l) in enumerate(MAP_STOPS):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(5.95), Inches(1.65), Inches(0.62), color)
        tb(slide, left, Inches(6.07), Inches(1.65), Inches(0.42), f"{emoji} ⭐", size=13,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "89–90 min", "Reading Hero")
    fade(slide)
    notes(slide,
          "You did it. All six stops, every mission. You are officially Bear's Reading Hero, "
          "and this badge is yours.",
          "Can you say it with me? I can read!",
          "\"I can read!\"",
          "If he is too tired to say it, say it for him warmly and let him just smile. The "
          "feeling matters more than the words.",
          "Ask him what he wants to read with Bear next time.",
          "\"I cannot wait to read with you again.\"",
          "1 minute")


def s41_assessment():
    slide, n = new_slide("📋 TEACHER ONLY — Observation Checklist", "TEACHER ONLY", "",
                         "Assessment", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Record where he is today, not where he should be. Fill this in right after class.",
       size=12, bold=True, color=BERRY)
    cols = [(0.45, 4.3, "READING SKILL"), (4.95, 2.6, "Needs Support"),
            (7.75, 2.4, "Developing"), (10.35, 2.5, "Independent")]
    header_y = Inches(1.7)
    for left, width, label in cols:
        add_round(slide, Inches(left), header_y, Inches(width), Inches(0.42), INK)
        tb(slide, Inches(left), header_y + Inches(0.07), Inches(width), Inches(0.3), label,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = header_y + Inches(0.48 + i * 0.36)
        band = WHITE if i % 2 == 0 else L_GREY
        add_round(slide, Inches(0.45), top, Inches(4.3), Inches(0.34), band)
        tb(slide, Inches(0.65), top + Inches(0.03), Inches(3.9), Inches(0.28), skill,
           size=12, bold=True, color=INK)
        for left, width, _l in cols[1:]:
            add_round(slide, Inches(left), top, Inches(width), Inches(0.34), band)
            tb(slide, Inches(left), top + Inches(0.01), Inches(width), Inches(0.3), "☐",
               size=13, color=SOFT, align=PP_ALIGN.CENTER)
    for i, label in enumerate(RUBRIC_NOTES):
        top = Inches(5.86 + i * 0.27)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.24), L_SUNNY)
        tb(slide, Inches(0.65), top + Inches(0.01), Inches(12.0), Inches(0.22),
           f"{label}:", size=11, bold=True, color=INK)
    notes(slide,
          "Teacher slide. Not shown to the child.",
          "—",
          "—",
          "Fill this in within five minutes of the lesson ending, while the detail is still "
          "fresh. Mark what he did today, not what you hoped for.",
          "If most marks land in Independent, move the next lesson to CVCC words and two-"
          "sentence reading.",
          "—",
          "5 minutes after class")


def s42_answer_key():
    slide, n = new_slide("🔒 TEACHER ONLY — Answer Key", "TEACHER ONLY", "", "Answer key",
                         DARK)
    tb(slide, Inches(0.45), Inches(1.34), Inches(12.4), Inches(0.35),
       "Hide this slide before presenting, or keep it on a second screen.",
       size=13, bold=True, color=BERRY)
    add_round(slide, Inches(0.45), Inches(1.78), Inches(6.1), Inches(5.05), WHITE)
    tb(slide, Inches(0.7), Inches(1.94), Inches(5.6), Inches(0.38),
       "🔊 Sounds · 🍎 Vowels · 🧩 Blending · 🏠 Families", size=14, bold=True, color=BERRY)
    bullets(slide, Inches(0.7), Inches(2.38), Inches(5.6), Inches(4.3), [
        "Sound Hunt 1–3: Bear, Sun, Cat",
        "Sound Hunt 4–6: Dog, Pig, Mat",
        "Vowel Basket 1–3: A, I, U",
        "Vowel Basket 4–6: E, O, A",
        "Build 1–3: CAT, SUN, PIG",
        "Build 4–6: DOG, HEN, BAT",
        "Family House 1–4: -AT, -UN, -IG, -AT",
        "Family House 5–8: -AN, -AT, -UN, -IG",
        "Families: -at cat bat hat mat · -an can",
        "     man fan · -ig pig big · -un sun run fun",
    ], size=12, sp=7)
    add_round(slide, Inches(6.75), Inches(1.78), Inches(6.1), Inches(5.05), WHITE)
    tb(slide, Inches(7.0), Inches(1.94), Inches(5.6), Inches(0.38),
       "💎 Sight words · 🖼️ Match · 🧩 Sentences · 📖 Story", size=14, bold=True, color=SKY)
    bullets(slide, Inches(7.0), Inches(2.38), Inches(5.6), Inches(4.3), [
        "Treasure 1–3: I, the, can",
        "Treasure 4–6: see, my, am",
        "Match It: PIG, BAT, HEN, BALL",
        "     (BALL has no picture on purpose)",
        "Puzzle 1–3: I see a cat. / The dog is big. /",
        "     I can run.",
        "Puzzle 4–5: My cat is big. / I see the sun.",
        "Story Detective 1–2: A Park, A Dog",
        "Story Detective 3–4: A Ball, A Happy",
        "Final: CAT, SUN, DOG, \"I see a cat.\"",
    ], size=12, sp=7)
    notes(slide,
          "Teacher slide. Not shown to the child.",
          "—",
          "—",
          "If he missed several answers in one section, that section is the starting point "
          "for the next lesson rather than something to reteach today.",
          "—",
          "—",
          "Reference only")


BUILDERS = [
    s01_title, s02_meet_bear, s03_reading_hero, s04_map, s05_picture_talk,
    s06_sound_warmup, s07_sound_hunt_a, s08_sound_hunt_b, s09_vowel_intro,
    s10_vowel_basket_a, s11_vowel_basket_b, s12_blending_intro, s13_build_a, s14_build_b,
    s15_family_intro, s16_family_game_a, s17_family_game_b, s18_bear_says, s19_sight_intro,
    s20_treasure_a, s21_treasure_b, s22_match_intro, s23_match_game, s24_sentence_intro,
    s25_sentence_1, s26_sentence_2, s27_sentence_3, s28_puzzle_a, s29_puzzle_b,
    s30_story_intro, s31_story_1, s32_story_2, s33_story_3, s34_story_4,
    s35_story_challenge, s36_detective_a, s37_detective_b, s38_final_challenge,
    s39_what_i_can_read, s40_badge, s41_assessment, s42_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

out = r"C:\Users\bhushaja\Downloads\shaip\Grade1_Bear_Big_Reading_Adventure_90min.pptx"
try:
    prs.save(out)
except PermissionError:
    out = out.replace(".pptx", "_new.pptx")
    prs.save(out)

story_words = sum(len(line.split()) for _p, _e, lines, _c, _l in STORY for line in lines)
missing = [i + 1 for i, s in enumerate(prs.slides) if not s.has_notes_slide
           or not s.notes_slide.notes_text_frame.text.strip()]
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
print(f"Story word count: {story_words}")
print(f"Slides missing teacher notes: {missing if missing else 'none'}")
