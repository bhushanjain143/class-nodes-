"""Grade 1 reading lesson - 90 minutes, 75 slides, no speaker notes.

"Build It! Reading Crew" - the child is a Junior Builder Reader working
through the Building Site, the Hard Hat Station, the Tool Box, the Brick
Wall, the Safety Sign Station, the Reading Site and Builder HQ. Reading is
built one step at a time: hear, say, sound, blend, read, understand.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Builder Hint strip, and slides 71-75 hold the optional
game bank, the support system, the assessment table and the answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x1C, 0x22, 0x2B)
DARK = RGBColor(0x2E, 0x38, 0x43)
SOFT = RGBColor(0x7C, 0x85, 0x90)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAGE = RGBColor(0xF8, 0xF8, 0xF5)
SITE = RGBColor(0x2E, 0x5F, 0x8A)
HAT = RGBColor(0xD1, 0x89, 0x06)
TOOL = RGBColor(0x1B, 0x76, 0x68)
BRICK = RGBColor(0xB0, 0x45, 0x2F)
SIGN = RGBColor(0xC4, 0x5A, 0x12)
WOOD = RGBColor(0x8A, 0x5A, 0x2B)
GREEN = RGBColor(0x1B, 0x6E, 0x3A)
L_SITE = RGBColor(0xE7, 0xEE, 0xF5)
L_HAT = RGBColor(0xFC, 0xF1, 0xD9)
L_TOOL = RGBColor(0xDF, 0xF1, 0xEE)
L_BRICK = RGBColor(0xFA, 0xE8, 0xE4)
L_SIGN = RGBColor(0xFC, 0xEB, 0xDD)
L_WOOD = RGBColor(0xF2, 0xEA, 0xE0)
L_GREEN = RGBColor(0xE2, 0xF1, 0xE8)
L_GREY = RGBColor(0xF0, 0xF2, 0xF4)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 75
_counter = {"n": 0}

# ------------------------------------------------------------------ content

STOPS = [("🏗️", "Building Site", "Look and talk", "0–7 min", SITE, L_SITE),
         ("🦺", "Hard Hat", "First sounds", "7–17 min", HAT, L_HAT),
         ("🔤", "Tool Box", "Short vowels", "17–27 min", TOOL, L_TOOL),
         ("🧱", "Brick Wall", "Building words", "27–44 min", BRICK, L_BRICK),
         ("🚧", "Safety Signs", "Words & sentences", "49–67 min", SIGN,
          L_SIGN),
         ("📖", "Reading Site", "The story", "67–85 min", WOOD, L_WOOD),
         ("🏆", "Builder HQ", "Celebrate", "85–90 min", GREEN, L_GREEN)]

MISSION = [("🔤", "Sounds", "Letters and the sounds they make."),
           ("🧩", "Words", "Push the sounds together."),
           ("📖", "Reading", "Short sentences you can read."),
           ("🎮", "Games", "Sorting, matching and puzzles."),
           ("⭐", "Story", "A real story in five parts.")]

SITE_SCENE = [("🏗️", "crane"), ("🧱", "bricks"), ("🚚", "truck"),
              ("🦺", "hard hat"), ("🔨", "tools"), ("🏠", "shed")]

TALK_QS = [("👀", "What do you see?"), ("🚚", "Do you see a truck?"),
           ("🏗️", "Do you see a crane?"), ("🦺", "What colour is the hat?")]
TALK_STARTERS = ["I see a ______.", "It is ______.", "I like the ______."]

VOCAB_CARDS = [
    ("🧱", "BRICK", "a hard block used to build a wall", "I see a brick.",
     "What colour is a brick?", BRICK, L_BRICK),
    ("🏗️", "CRANE", "a tall machine that lifts heavy things",
     "The crane is big.", "What can a crane lift?", SITE, L_SITE),
    ("🔨", "TOOL", "a thing you use to build and fix", "I can use a tool.",
     "Can you name one tool?", TOOL, L_TOOL),
]

WORD_BANK = [("👷", "builder", "a person who builds things"),
             ("🦺", "hard hat", "it keeps your head safe"),
             ("🚚", "truck", "it carries bricks and wood"),
             ("🗺️", "map", "it shows you where to go"),
             ("🧱", "wall", "bricks on top of bricks"),
             ("🚪", "door", "you go in and out of it"),
             ("📦", "box", "you put things inside it"),
             ("🔧", "fix", "make it work again"),
             ("🏠", "build", "make something new")]

SOUNDS = [("B", "/b/", "🧱", "BRICK", "bbb - rick", BRICK, L_BRICK),
          ("C", "/c/", "🏗️", "CRANE", "ccc - rane", SITE, L_SITE),
          ("D", "/d/", "⛏️", "DIG", "ddd - ig", WOOD, L_WOOD),
          ("F", "/f/", "🔧", "FIX", "fff - ix", TOOL, L_TOOL),
          ("M", "/m/", "🗺️", "MAP", "mmm - ap", SIGN, L_SIGN),
          ("T", "/t/", "🔨", "TOOL", "ttt - ool", HAT, L_HAT)]

TOOL_MATCH_A = [("/b/", ["🧱", "🏗️", "🗺️"]),
                ("/c/", ["🚚", "🏗️", "🔨"]),
                ("/m/", ["🧱", "🗺️", "🔧"])]
TOOL_MATCH_B = [("/t/", ["🧱", "🔨", "🗺️"]),
                ("/d/", ["⛏️", "🧱", "🚚"]),
                ("/f/", ["🔧", "🏗️", "📦"])]

DROP_ROUNDS = [("/m/", ["BAG", "MAP", "FIX"]),
               ("/b/", ["TOOL", "DIG", "BAG"]),
               ("/f/", ["FIX", "MAP", "CRANE"])]

VOWELS = [("A", "/a/", "m - aaa - p", [("🗺️", "MAP"), ("🎒", "BAG"),
                                       ("🎩", "HAT")], BRICK, L_BRICK),
          ("E", "/e/", "r - eee - d", [("🟥", "RED"), ("🛏️", "BED"),
                                       ("💧", "WET")], SIGN, L_SIGN),
          ("I", "/i/", "b - iii - g", [("🐘", "BIG"), ("🪑", "SIT"),
                                       ("🥊", "HIT")], TOOL, L_TOOL),
          ("O", "/o/", "f - ooo - x", [("🦊", "FOX"), ("📦", "BOX"),
                                       ("🔥", "HOT")], SITE, L_SITE)]

BRICK_BUCKETS = [("A", BRICK, L_BRICK), ("E", SIGN, L_SIGN),
                 ("I", TOOL, L_TOOL), ("O", SITE, L_SITE)]
BRICK_WORDS = ["MAP", "RED", "BIG", "FOX", "BAG", "HOT"]

MISSING_VOWEL_A = [("B _ G", "🐘", ["A", "E", "I"]),
                   ("M _ P", "🗺️", ["A", "I", "O"]),
                   ("R _ D", "🟥", ["E", "A", "O"])]
MISSING_VOWEL_B = [("H _ T", "🔥", ["A", "E", "O"]),
                   ("F _ X", "🦊", ["I", "O", "A"])]

BLENDS = [("M", "A", "P", "🗺️", "MAP", "I see a map.", SIGN, L_SIGN),
          ("B", "A", "G", "🎒", "BAG", "My bag is red.", BRICK, L_BRICK),
          ("H", "A", "T", "🎩", "HAT", "The hat is big.", HAT, L_HAT),
          ("R", "U", "N", "🏃", "RUN", "I can run.", TOOL, L_TOOL),
          ("F", "O", "X", "🦊", "FOX", "I see a fox.", SITE, L_SITE),
          ("B", "I", "G", "🐘", "BIG", "The truck is big.", WOOD, L_WOOD)]

BUILD_BLOCKS = [(["B", "A", "G"], "🎒"), (["M", "A", "P"], "🗺️"),
                (["F", "O", "X"], "🦊")]

SCRAMBLE = [(["A", "P", "M"], "🗺️"), (["G", "I", "B"], "🐘"),
            (["T", "A", "H"], "🎩"), (["X", "O", "F"], "🦊")]

FAMILIES = [("-AT", "at", [("🐱", "CAT"), ("🎩", "HAT"), ("🦇", "BAT"),
                           ("🟫", "MAT")], BRICK, L_BRICK),
            ("-IG", "ig", [("🐘", "BIG"), ("⛏️", "DIG"), ("🐷", "PIG")],
             TOOL, L_TOOL),
            ("-UN", "un", [("☀️", "SUN"), ("🏃", "RUN"), ("🎉", "FUN")],
             HAT, L_HAT),
            ("-OP", "op", [("🐰", "HOP"), ("⬆️", "TOP"), ("🧹", "MOP")],
             SITE, L_SITE)]

SORT_BUCKETS = [("-AT", BRICK, L_BRICK), ("-UN", HAT, L_HAT)]
SORT_WORDS = ["CAT", "SUN", "HAT", "RUN", "MAT", "FUN"]

ODD_ONE = [["FOX", "BOX", "SUN"], ["CAT", "HAT", "BIG"], ["DIG", "PIG", "MAP"]]

RHYMES = [("CAT", "🐱", ["HAT", "BIG", "SUN"]),
          ("BOX", "📦", ["FOX", "MAP", "RUN"]),
          ("SUN", "☀️", ["FUN", "HAT", "DIG"])]

BUILDER_SAYS = [("🖐️", "Builder says tap the desk.", "two taps"),
                ("🧱", "Builder says point to the brick.", "point"),
                ("🔨", "Builder says pretend to hammer.", "three swings"),
                ("🦺", "Builder says put on your hard hat.", "hands on head")]

BUILDER_READS = [("🗺️", "Builder says read MAP.", "read it out loud"),
                 ("🎩", "Builder says find HAT.", "point to it"),
                 ("📦", "Builder says point to BOX.", "point to it"),
                 ("🐘", "Builder says read BIG.", "read it out loud")]

SIGHT_SLIDES = [(["I"], "I can run.", "You say this word about yourself.",
                 SIGN, L_SIGN),
                (["see"], "I see a brick.", "This word is about your eyes.",
                 SITE, L_SITE),
                (["a", "the"], "I see a truck.  ·  I see the truck.",
                 "Two tiny words. They sit before a naming word.", TOOL,
                 L_TOOL),
                (["my", "is", "can"], "My hat is red.  ·  I can run.",
                 "One word says it belongs to you. One joins. One means "
                 "you are able to.", BRICK, L_BRICK)]

SIGHT_FIND = ["I", "see", "a", "the", "my", "is", "can"]

BUILD_SENTENCES = [("I see a brick.", ["I", "see", "a", "brick"], "🧱",
                    BRICK, L_BRICK),
                   ("I see the truck.", ["I", "see", "the", "truck"], "🚚",
                    SITE, L_SITE),
                   ("My hat is red.", ["My", "hat", "is", "red"], "🎩",
                    SIGN, L_SIGN),
                   ("I can run.", ["I", "can", "run"], "🏃", TOOL, L_TOOL)]

PUZZLES = [(["see", "I", "brick", "a"], "🧱", 4),
           (["red", "hat", "My", "is"], "🎩", 4),
           (["can", "I", "run"], "🏃", 3),
           (["big", "is", "truck", "The"], "🚚", 4)]

SENT_MATCH = [("The truck is big.", ["🚚", "🐱", "📦"]),
              ("I see a brick.", ["🧱", "🔨", "🚪"])]

REAL_SILLY = ["The builder uses a brick.", "The brick eats lunch.",
              "The truck is big.", "The hard hat drives a truck."]

STORY_PREDICT = [("🏠", "a little shed"), ("🚚", "a big truck"),
                 ("🧱", "a tall wall")]

STORY = [
    ("Part 1", "🎒", BRICK, L_BRICK,
     ["Sam has a red bag.",
      "He has a small map.",
      "The map shows the yard.",
      "\"Let's build!\" says Sam."],
     "Sam has a red bag.", "What colour is Sam's bag?"),
    ("Part 2", "📦", SITE, L_SITE,
     ["Sam and Dad go to the yard.",
      "They see a big box.",
      "The box is full of wood.",
      "Dad lifts the box."],
     "They see a big box.", "What do they see in the yard?"),
    ("Part 3", "🔨", HAT, L_HAT,
     ["Sam gets a hammer.",
      "Dad gets some wood.",
      "Sam puts on his hard hat.",
      "Now they can start."],
     "Sam gets a hammer.", "What does Sam get?"),
    ("Part 4", "🧱", TOOL, L_TOOL,
     ["Sam can fix the little shed.",
      "He puts a board on the wall.",
      "Tap, tap, tap goes the hammer.",
      "The wall is strong."],
     "The wall is strong.", "What does Sam put on the wall?"),
    ("Part 5", "🏠", WOOD, L_WOOD,
     ["The shed is small and red.",
      "The door can open and shut.",
      "Sam looks at the shed.",
      "\"I did it!\" says Sam."],
     "Sam looks at the shed.", "How does Sam feel at the end?"),
]

DETECT_A = [("🧒", "Who is in the story?", ["Ben", "Sam", "Mia"], "Part 1"),
            ("🎒", "What colour is Sam's bag?", ["Blue", "Green", "Red"],
             "Part 1")]
DETECT_B = [("🔨", "What does Sam get?", ["A hammer", "A ball", "A cup"],
             "Part 3"),
            ("🏠", "What does Sam help fix?", ["A car", "A shed", "A bike"],
             "Part 4")]

SEQUENCE = [("A", "🔨", "Sam gets a hammer."),
            ("B", "🚪", "Sam goes to the yard."),
            ("C", "🎒", "Sam has a red bag."),
            ("D", "🏠", "Sam helps fix the shed.")]

EVIDENCE = [("What is inside the big box?", "Part 2", SITE, L_SITE),
            ("What does Sam wear on his head?", "Part 3", HAT, L_HAT),
            ("What colour is the shed?", "Part 5", WOOD, L_WOOD)]

FINAL_ROUNDS = [("🔤", "WORD", "MAP  ·  BIG  ·  HAT", "🗺️", BRICK, L_BRICK),
                ("🧩", "BLEND", "B – O – X", "📦", TOOL, L_TOOL),
                ("📖", "SENTENCE", "I see a brick.", "🧱", SIGN, L_SIGN)]

STAR_LINES = [("🔤", "You read new words!"), ("🧩", "You built words!"),
              ("📖", "You read sentences!"), ("📚", "You read a story!"),
              ("🔎", "You answered questions!")]

GAMES = [("🧱", "MISSING LETTER", "Fill the gap in the brick.",
          ["M _ P", "A  ·  E  ·  I"], BRICK, L_BRICK),
         ("🔨", "WORD BUILDER", "Fix the mixed-up word.",
          ["G - A - B", "say the sounds, then swap"], HAT, L_HAT),
         ("🚧", "FIND THE WORD", "Find one word in the sentence.",
          ["I see a big truck.", "Find TRUCK."], SIGN, L_SIGN),
         ("🎯", "RHYME BUILDER", "Which one rhymes?",
          ["CAT", "HAT  ·  FOX  ·  SUN"], TOOL, L_TOOL),
         ("🧩", "SENTENCE PUZZLE", "Put the cards in order.",
          ["big  /  is  /  truck  /  The"], SITE, L_SITE),
         ("🦺", "REAL OR SILLY", "Sensible or silly?",
          ["The builder wears a hard hat.", "The hard hat eats lunch."],
          WOOD, L_WOOD)]

GAME_WHEN = ["He finishes an activity early.",
             "His focus starts to drop.",
             "One sound needs more practice.",
             "You have five spare minutes.",
             "He asks to play one more."]

SUPPORT_LEVELS = [("🟢", "BUILDER HINT", "Needs a picture or a first sound",
                   "Point at the picture, then the first letter.", TOOL,
                   L_TOOL),
                  ("🟡", "BUILDING MISSION", "Reads it with a little help",
                   "Ask one small question, then wait quietly.", HAT, L_HAT),
                  ("⭐", "MASTER BUILDER", "Reads it on his own",
                   "Ask him to build a new sentence with the word.", SITE,
                   L_SITE)]

SUPPORT_LADDER = [("HINT 1", "Look at the first letter.", TOOL),
                  ("HINT 2", "Say each sound.", SITE),
                  ("HINT 3", "Blend the sounds together.", BRICK),
                  ("HINT 4", "Let's read it together.", HAT),
                  ("HINT 5", "I read it, then you read it.", SIGN)]

HARD_WORD = ["Read the whole sentence out loud yourself.",
             "Point to each word as you read it again.",
             "Explain one hard word with a picture.",
             "Read it together, at his speed.",
             "Let him try one short sentence alone.",
             "Praise the effort, not just the answer."]

PRAISE = ["\"Good try!\"", "\"Let's build this word together.\"",
          "\"You found the first sound!\"", "\"Great blending!\"",
          "\"Try it one more time.\"", "\"You solved it!\""]

NEVER_SAY = ["\"You can't read this.\"", "\"That's wrong.\"",
             "\"You should know this by now.\""]

RUBRIC_SKILLS = ["Beginning sounds", "Short vowels", "CVC words", "Blending",
                 "Word families", "Sight words", "Phrase reading",
                 "Sentence reading", "Story reading", "Comprehension"]

REVIEW_NEXT = ["WORDS TO REVIEW", "SIGHT WORDS TO REVIEW", "NEXT READING GOAL"]

# ------------------------------------------------------------------ helpers


def set_run(run, size=20, bold=False, color=DARK, font="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_bg(slide, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width,
                               prs.slide_height)
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
       align=PP_ALIGN.LEFT, font="Calibri", italic=False):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, color, font)
    r.font.italic = italic
    return box


def bullets(slide, l, t, w, h, items, size=12, color=DARK, sp=5):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(sp)
        r = p.add_run()
        # Leading spaces mark a wrapped continuation line rather than a bullet.
        r.text = ("      " + item.strip()) if item.startswith(" ") \
            else ("•  " + item)
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


def footer(slide, n, timing="", stop=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE1, 0xE5, 0xE8))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), HAT)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40),
             INK)
    msg = "🏗️ Build It! Reading Crew  |  Grade 1  |  90 min"
    if stop:
        msg += f"  |  {stop}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg,
       size=10, color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32),
       f"{n} / {TOTAL}", size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, stop, accent=SITE, bg=PAGE):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height,
             accent)
    add_round(slide, Inches(0.4), Inches(0.26), Inches(3.5), Inches(0.42),
              accent)
    tb(slide, Inches(0.4), Inches(0.31), Inches(3.5), Inches(0.34), tag,
       size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if timing:
        add_round(slide, Inches(10.3), Inches(0.26), Inches(2.65),
                  Inches(0.42), INK)
        tb(slide, Inches(10.3), Inches(0.31), Inches(2.65), Inches(0.34),
           timing, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title,
       size=28, bold=True, color=INK, font="Georgia")
    footer(slide, n, timing, stop)
    fade(slide)
    return slide, n


def one_task(slide, text, color=SITE, top=1.34):
    """States the single job of this slide in words the child understands."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4),
              text, size=16, bold=True, color=color)


def hint(slide, text, top=6.42, label="🔧 BUILDER HINT", fill=L_HAT,
         color=HAT):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46),
              fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.7), Inches(0.34),
       label, size=12, bold=True, color=color)
    tb(slide, Inches(3.6), Inches(top + 0.05), Inches(9.0), Inches(0.36),
       text, size=13, bold=True, color=INK)


def i_we_you(slide, top=6.42):
    steps = [("I READ", "Teacher first", BRICK), ("WE READ", "Together", HAT),
             ("YOU READ", "Your turn!", TOOL)]
    for i, (label, who, color) in enumerate(steps):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(top), Inches(3.9), Inches(0.46), color)
        tb(slide, left, Inches(top + 0.05), Inches(3.9), Inches(0.36),
           f"{label}  ·  {who}", size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)


def numbered_rows(slide, count, start_index, accent, top_start=1.9, gap=1.5,
                  height=1.34, fill=WHITE):
    """Shared row scaffold; returns the top edge of each numbered row."""
    tops = []
    for i in range(count):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), fill)
        add_oval(slide, Inches(0.75), top + Inches(height / 2 - 0.25),
                 Inches(0.5), Inches(0.5), accent)
        tb(slide, Inches(0.75), top + Inches(height / 2 - 0.19), Inches(0.5),
           Inches(0.4), str(start_index + i), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tops.append(top)
    return tops


def picture_rows(slide, items, start_index, accent, label="STARTS WITH",
                 top_start=1.95):
    """A target sound, then three pictures with no captions to give it away."""
    tops = numbered_rows(slide, len(items), start_index, accent,
                         top_start=top_start, gap=1.5)
    for (sound, pics), top in zip(items, tops):
        add_round(slide, Inches(1.4), top + Inches(0.28), Inches(2.5),
                  Inches(0.78), accent)
        tb(slide, Inches(1.4), top + Inches(0.33), Inches(2.5), Inches(0.28),
           label, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.4), top + Inches(0.6), Inches(2.5), Inches(0.42),
           sound, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, pic in enumerate(pics):
            left = Inches(4.4 + j * 2.75)
            add_round(slide, left, top + Inches(0.17), Inches(2.5),
                      Inches(1.0), L_GREY)
            tb(slide, left, top + Inches(0.22), Inches(2.5), Inches(0.9), pic,
               size=44, align=PP_ALIGN.CENTER)


def choice_rows(slide, items, start_index, accent, top_start=1.95, gap=1.5,
                height=1.34):
    """A word with a gap, its picture, and letter buttons to choose from."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start,
                         gap, height)
    for (word, pic, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.26), Inches(1.1), Inches(0.82),
           pic, size=40, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.7), top + Inches(0.3), Inches(3.4), Inches(0.72),
           word, size=38, bold=True, color=INK, font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(6.6 + j * 2.1)
            add_round(slide, left, top + Inches(0.3), Inches(1.85),
                      Inches(0.74), L_GREY)
            tb(slide, left, top + Inches(0.38), Inches(1.85), Inches(0.56),
               opt, size=26, bold=True, color=accent, align=PP_ALIGN.CENTER,
               font="Arial Black")


def word_cards(slide, left, top, words, color, light, width=1.65, gap=1.78,
               height=0.8, size=22):
    """A row of big readable word tiles."""
    for i, word in enumerate(words):
        l = Inches(left + i * gap)
        add_round(slide, l, Inches(top), Inches(width), Inches(height), light)
        tb(slide, l, Inches(top + height / 2 - 0.24), Inches(width),
           Inches(0.5), word, size=size, bold=True, color=color,
           align=PP_ALIGN.CENTER, font="Arial Black")


def question_rows(slide, items, start_index, accent, top_start=1.95, gap=2.16,
                  height=1.92):
    """A story question with a picture, three answers and where to look."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start,
                         gap, height)
    for (pic, question, options, where), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.25), Inches(0.9), Inches(0.7),
           pic, size=28, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.4), top + Inches(0.2), Inches(4.2), Inches(0.8),
           question, size=21, bold=True, color=INK)
        add_round(slide, Inches(2.4), top + Inches(1.1), Inches(2.0),
                  Inches(0.44), accent)
        tb(slide, Inches(2.4), top + Inches(1.16), Inches(2.0), Inches(0.34),
           f"look in {where}", size=11, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        for j, opt in enumerate(options):
            left = Inches(6.9 + j * 2.0)
            add_round(slide, left, top + Inches(0.42), Inches(1.85),
                      Inches(1.1), L_GREY)
            tb(slide, left, top + Inches(0.55), Inches(1.85), Inches(0.3),
               chr(65 + j), size=11, bold=True, color=accent,
               align=PP_ALIGN.CENTER)
            tb(slide, left + Inches(0.1), top + Inches(0.86), Inches(1.65),
               Inches(0.56), opt, size=15, bold=True, color=INK,
               align=PP_ALIGN.CENTER)


def badge_slide(title, timing, stop, emoji, badge, sub, lines, next_line,
                color, light):
    """Closes a station: what was learned, and what comes next."""
    slide, n = new_slide(title, "WELL DONE", timing, stop, color)
    one_task(slide, sub, color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(4.6), Inches(4.4),
              light)
    tb(slide, Inches(0.5), Inches(2.25), Inches(4.6), Inches(1.7), emoji,
       size=96, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.1), Inches(4.35), Inches(3.4), Inches(0.85),
              color)
    tb(slide, Inches(1.1), Inches(4.52), Inches(3.4), Inches(0.5), badge,
       size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(5.4), Inches(4.0), Inches(0.6),
       "Give yourself a big clap!", size=16, color=SOFT,
       align=PP_ALIGN.CENTER, italic=True)
    count = len(lines)
    gap, height, size = ((0.95, 0.8, 17) if count <= 4
                         else (0.78, 0.66, 15) if count == 5
                         else (0.64, 0.54, 13))
    dot = min(0.52, height - 0.16)
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.85 + i * gap)
        add_round(slide, Inches(5.45), top, Inches(7.4), Inches(height),
                  WHITE)
        add_oval(slide, Inches(5.75), top + Inches((height - dot) / 2),
                 Inches(dot), Inches(dot), light)
        tb(slide, Inches(5.75), top + Inches(height / 2 - 0.19), Inches(dot),
           Inches(0.4), icon, size=13, align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.55), top + Inches(height / 2 - 0.21), Inches(6.1),
           Inches(0.46), line, size=size, bold=True, color=INK)
    add_round(slide, Inches(5.45), Inches(5.7), Inches(7.4), Inches(0.52),
              color)
    tb(slide, Inches(5.75), Inches(5.8), Inches(6.8), Inches(0.36),
       f"NEXT:  {next_line}", size=13, bold=True, color=WHITE)


def vocab_slide(index):
    """Picture, say it, the word, a simple meaning, then a sentence."""
    pic, word, meaning, sentence, ask, color, light = VOCAB_CARDS[index]
    slide, n = new_slide(f"{pic}  Site Word — {word}", "SITE WORDS",
                         "0–7 min", "Building Site", color)
    one_task(slide, "Look, say it, then read it with me.", color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(4.3), Inches(4.4),
              light)
    tb(slide, Inches(0.5), Inches(2.15), Inches(4.3), Inches(1.9), pic,
       size=100, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.0), Inches(4.3), Inches(3.3), Inches(1.3),
              color)
    tb(slide, Inches(1.0), Inches(4.55), Inches(3.3), Inches(0.85), word,
       size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Arial Black")
    panels = [("SAY IT", f"\"{word.lower()}\"  ·  say it twice", 22, light),
              ("WHAT IT MEANS", meaning, 22, WHITE),
              ("READ THIS SENTENCE", sentence, 26, WHITE)]
    for i, (head, text, size, fill) in enumerate(panels):
        top = Inches(1.85 + i * 1.25)
        add_round(slide, Inches(5.15), top, Inches(7.7), Inches(1.1), fill)
        tb(slide, Inches(5.45), top + Inches(0.1), Inches(7.1), Inches(0.32),
           head, size=11, bold=True, color=color)
        tb(slide, Inches(5.45), top + Inches(0.42), Inches(7.1), Inches(0.58),
           text, size=size, bold=True, color=INK)
    add_round(slide, Inches(5.15), Inches(5.6), Inches(7.7), Inches(0.56),
              L_GREY)
    tb(slide, Inches(5.45), Inches(5.7), Inches(7.1), Inches(0.4),
       f"⭐ YOUR TURN — {ask}", size=15, bold=True, color=INK)
    hint(slide, "If he cannot say it, say it first and let him copy you.",
         6.42)


def sound_slide(index):
    """One letter, its sound, a picture and the site word it starts."""
    letter, sound, pic, word, stretch, color, light = SOUNDS[index]
    slide, n = new_slide(f"{letter}  is for  {word}", "SOUNDS", "7–17 min",
                         "Hard Hat Station", color)
    one_task(slide, f"{letter} makes the {sound} sound. Say it with me.",
             color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(4.3), Inches(4.4),
              light)
    tb(slide, Inches(0.5), Inches(2.1), Inches(4.3), Inches(1.9), pic,
       size=100, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.0), Inches(4.25), Inches(3.3), Inches(1.4),
              WHITE)
    tb(slide, Inches(1.0), Inches(4.5), Inches(3.3), Inches(0.9), word,
       size=40, bold=True, color=color, align=PP_ALIGN.CENTER,
       font="Arial Black")
    add_round(slide, Inches(5.15), Inches(1.85), Inches(7.7), Inches(2.0),
              WHITE)
    tb(slide, Inches(5.15), Inches(1.95), Inches(7.7), Inches(0.4),
       "THE LETTER", size=12, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(5.15), Inches(2.3), Inches(7.7), Inches(1.4),
       f"{letter} {letter.lower()}", size=86, bold=True, color=INK,
       align=PP_ALIGN.CENTER, font="Arial Black")
    add_round(slide, Inches(5.15), Inches(4.0), Inches(3.7), Inches(1.1),
              color)
    tb(slide, Inches(5.15), Inches(4.13), Inches(3.7), Inches(0.36),
       "THE SOUND", size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(5.15), Inches(4.45), Inches(3.7), Inches(0.56), sound,
       size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Arial Black")
    add_round(slide, Inches(9.15), Inches(4.0), Inches(3.7), Inches(1.1),
              L_GREY)
    tb(slide, Inches(9.15), Inches(4.13), Inches(3.7), Inches(0.36),
       "STRETCH IT", size=11, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(9.15), Inches(4.45), Inches(3.7), Inches(0.56), stretch,
       size=24, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(5.15), Inches(5.25), Inches(7.7), Inches(0.9),
              L_GREY)
    tb(slide, Inches(5.4), Inches(5.35), Inches(7.2), Inches(0.36),
       "⭐ YOUR TURN", size=11, bold=True, color=color)
    tb(slide, Inches(5.4), Inches(5.66), Inches(7.2), Inches(0.42),
       f"Say {sound} three times, then say {word}.", size=16, bold=True,
       color=INK)
    hint(slide, f"Put your hand near your mouth and feel the {sound} sound.",
         6.42)


def vowel_slide(index):
    """One short vowel, its stretch sound, and three words that use it."""
    letter, sound, stretch, words, color, light = VOWELS[index]
    slide, n = new_slide(f"🔤 Short {letter}  —  {sound}", "SHORT VOWELS",
                         "17–27 min", "Tool Box", color)
    one_task(slide, f"{letter} in the middle says {sound}.", color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(3.5), Inches(4.4),
              light)
    tb(slide, Inches(0.5), Inches(2.15), Inches(3.5), Inches(1.7), letter,
       size=120, bold=True, color=color, align=PP_ALIGN.CENTER,
       font="Arial Black")
    add_round(slide, Inches(0.9), Inches(4.0), Inches(2.7), Inches(0.8),
              color)
    tb(slide, Inches(0.9), Inches(4.13), Inches(2.7), Inches(0.56), sound,
       size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Arial Black")
    add_round(slide, Inches(0.9), Inches(4.95), Inches(2.7), Inches(0.9),
              WHITE)
    tb(slide, Inches(0.9), Inches(5.05), Inches(2.7), Inches(0.3),
       "STRETCH IT", size=10, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.9), Inches(5.35), Inches(2.7), Inches(0.42), stretch,
       size=20, bold=True, color=INK, align=PP_ALIGN.CENTER)
    for i, (pic, word) in enumerate(words):
        left = Inches(4.35 + i * 2.9)
        add_round(slide, left, Inches(1.85), Inches(2.7), Inches(4.4), WHITE)
        tb(slide, left, Inches(2.25), Inches(2.7), Inches(1.3), pic, size=60,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(3.75), Inches(2.2),
                  Inches(0.95), light)
        tb(slide, left + Inches(0.25), Inches(3.92), Inches(2.2), Inches(0.6),
           word, size=30, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        letters = " - ".join(f"/{ch.lower()}/" for ch in word)
        tb(slide, left + Inches(0.15), Inches(4.9), Inches(2.4), Inches(0.4),
           letters, size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(5.4), Inches(2.4), Inches(0.4),
           "☐ I read it", size=12, color=SOFT, align=PP_ALIGN.CENTER)
    hint(slide, f"Say the three words in a row. The middle sound never "
                f"changes: {sound}.", 6.42)


def blend_slide(index):
    """Three sounds, one arrow, one whole word, then a sentence."""
    a, b, c, pic, word, sentence, color, light = BLENDS[index]
    slide, n = new_slide(f"{a} + {b} + {c}  →  {word}", "BLENDING",
                         "27–37 min", "Brick Wall", color)
    one_task(slide, "Touch each brick, say the sound, then push them "
                    "together.", color)
    for i, letter in enumerate([a, b, c]):
        left = Inches(0.6 + i * 2.55)
        add_round(slide, left, Inches(1.85), Inches(2.3), Inches(2.0), light)
        tb(slide, left, Inches(2.05), Inches(2.3), Inches(1.1), letter,
           size=72, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        add_round(slide, left + Inches(0.4), Inches(3.15), Inches(1.5),
                  Inches(0.5), WHITE)
        tb(slide, left + Inches(0.4), Inches(3.22), Inches(1.5), Inches(0.36),
           f"/{letter.lower()}/", size=18, bold=True, color=INK,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(8.3), Inches(2.3), Inches(1.2), Inches(1.0), "→",
       size=48, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(9.5), Inches(1.85), Inches(3.35), Inches(2.0),
              color)
    tb(slide, Inches(9.5), Inches(1.95), Inches(3.35), Inches(0.7), pic,
       size=30, align=PP_ALIGN.CENTER)
    tb(slide, Inches(9.5), Inches(2.7), Inches(3.35), Inches(0.95), word,
       size=46, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Arial Black")
    add_round(slide, Inches(0.6), Inches(4.05), Inches(12.25), Inches(1.0),
              WHITE)
    tb(slide, Inches(0.9), Inches(4.15), Inches(11.6), Inches(0.32),
       "SLOW, THEN FAST", size=11, bold=True, color=SOFT)
    tb(slide, Inches(0.9), Inches(4.47), Inches(11.6), Inches(0.5),
       f"/{a.lower()}/ … /{b.lower()}/ … /{c.lower()}/      →      "
       f"{a.lower()}{b.lower()}{c.lower()}      →      {word}", size=24,
       bold=True, color=INK)
    add_round(slide, Inches(0.6), Inches(5.2), Inches(12.25), Inches(0.95),
              light)
    tb(slide, Inches(0.9), Inches(5.3), Inches(11.6), Inches(0.32),
       "NOW READ THE WHOLE SENTENCE", size=11, bold=True, color=color)
    tb(slide, Inches(0.9), Inches(5.62), Inches(11.6), Inches(0.46), sentence,
       size=26, bold=True, color=INK, font="Georgia")
    hint(slide, f"Stuck? Blend the first two only: "
                f"{a.lower()}{b.lower()} … then add /{c.lower()}/.", 6.42)


def family_slide(index):
    """One word family, its ending, and the words that rhyme inside it."""
    name, ending, words, color, light = FAMILIES[index]
    slide, n = new_slide(f"🧱 The {name} Family", "WORD FAMILIES", "37–44 min",
                         "Brick Wall", color)
    one_task(slide, f"Every word here ends the same way: {name}. "
                    f"They all rhyme.", color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(3.2), Inches(4.4),
              color)
    tb(slide, Inches(0.5), Inches(2.4), Inches(3.2), Inches(1.5), name,
       size=76, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Arial Black")
    tb(slide, Inches(0.7), Inches(4.1), Inches(2.8), Inches(0.5),
       f"say it: {ending}", size=20, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.8), Inches(4.8), Inches(2.6), Inches(1.0),
              WHITE)
    tb(slide, Inches(0.8), Inches(4.9), Inches(2.6), Inches(0.32),
       "HOW MANY?", size=10, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(5.2), Inches(2.6), Inches(0.5),
       f"{len(words)} words", size=22, bold=True, color=color,
       align=PP_ALIGN.CENTER)
    span = 9.15 / len(words)
    for i, (pic, word) in enumerate(words):
        left = Inches(4.0 + i * span)
        add_round(slide, left, Inches(1.85), Inches(span - 0.2), Inches(4.4),
                  light)
        tb(slide, left, Inches(2.3), Inches(span - 0.2), Inches(1.3), pic,
           size=56, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(3.8),
                  Inches(span - 0.6), Inches(1.0), WHITE)
        tb(slide, left + Inches(0.2), Inches(4.0), Inches(span - 0.6),
           Inches(0.6), word, size=30, bold=True, color=INK,
           align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.1), Inches(5.0), Inches(span - 0.4),
           Inches(0.4), f"{word[0].lower()} + {ending}", size=15, bold=True,
           color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(5.5), Inches(span - 0.4),
           Inches(0.4), "☐ I read it", size=12, color=SOFT,
           align=PP_ALIGN.CENTER)
    hint(slide, "Cover the first letter, read the ending, then uncover it.",
         6.42)


def sight_slide(index):
    """Sight words: see them big, hear what they do, find them in a row."""
    words, sentence, note, color, light = SIGHT_SLIDES[index]
    heading = "  ·  ".join(w.upper() for w in words)
    slide, n = new_slide(f"🚧 Sight Word — {heading}", "SIGHT WORDS",
                         "49–58 min", "Safety Signs", color)
    if len(words) == 1:
        one_task(slide, "This word is a friend. We know it by looking, not "
                        "sounding it out.", color)
        add_round(slide, Inches(0.5), Inches(1.85), Inches(5.0), Inches(2.6),
                  color)
        tb(slide, Inches(0.5), Inches(2.25), Inches(5.0), Inches(1.5),
           words[0], size=96, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        add_round(slide, Inches(0.5), Inches(4.65), Inches(5.0), Inches(1.6),
                  L_GREY)
        tb(slide, Inches(0.5), Inches(4.8), Inches(5.0), Inches(0.36),
           "SAY IT THREE TIMES", size=12, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        for j in range(3):
            add_round(slide, Inches(1.25 + j * 1.2), Inches(5.25),
                      Inches(1.0), Inches(0.8), WHITE)
            tb(slide, Inches(1.25 + j * 1.2), Inches(5.42), Inches(1.0),
               Inches(0.5), "☐", size=22, color=SOFT, align=PP_ALIGN.CENTER)
    else:
        one_task(slide, "These words are friends. We know them by looking, "
                        "not sounding them out.", color)
        two = len(words) == 2
        gap, height = (2.4, 2.0) if two else (1.55, 1.35)
        size, drop = (64, 0.5) if two else (44, 0.32)
        for i, word in enumerate(words):
            top = Inches(1.85 + i * gap)
            add_round(slide, Inches(0.5), top, Inches(5.0), Inches(height),
                      color)
            tb(slide, Inches(0.5), top + Inches(drop), Inches(5.0),
               Inches(height - drop - 0.2), word, size=size, bold=True,
               color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
    add_round(slide, Inches(5.8), Inches(1.85), Inches(7.05), Inches(1.15),
              WHITE)
    tb(slide, Inches(6.1), Inches(1.95), Inches(6.5), Inches(0.32),
       "WHAT IT DOES", size=11, bold=True, color=color)
    tb(slide, Inches(6.1), Inches(2.3), Inches(6.5), Inches(0.6), note,
       size=15, bold=True, color=INK)
    add_round(slide, Inches(5.8), Inches(3.15), Inches(7.05), Inches(1.3),
              light)
    tb(slide, Inches(6.1), Inches(3.28), Inches(6.5), Inches(0.32),
       "READ IT IN A SENTENCE", size=11, bold=True, color=color)
    tb(slide, Inches(6.1), Inches(3.65), Inches(6.5), Inches(0.62), sentence,
       size=23, bold=True, color=INK, font="Georgia")
    add_round(slide, Inches(5.8), Inches(4.6), Inches(7.05), Inches(1.65),
              WHITE)
    tb(slide, Inches(6.1), Inches(4.72), Inches(6.5), Inches(0.32),
       "⭐ FIND IT — point to it in this row", size=11, bold=True,
       color=color)
    word_cards(slide, 6.1, 5.15, SIGHT_FIND, color, light, width=0.88,
               gap=0.94, height=0.72, size=15)
    hint(slide, "Say it, clap it, then hunt for the same shape in the row.",
         6.42)


def build_slide(index, timing="49–58 min", stop="Safety Signs"):
    """Word tiles in order, a picture, then the whole sentence."""
    sentence, words, pic, color, light = BUILD_SENTENCES[index]
    slide, n = new_slide(f"🧩 Build It —  {sentence}", "SENTENCES", timing,
                         stop, color)
    one_task(slide, "Read each card, then read the whole sentence.", color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(3.3), Inches(3.0),
              light)
    tb(slide, Inches(0.5), Inches(2.2), Inches(3.3), Inches(1.9), pic,
       size=90, align=PP_ALIGN.CENTER)
    span = 8.85 / len(words)
    for i, word in enumerate(words):
        left = Inches(4.05 + i * span)
        add_round(slide, left, Inches(1.85), Inches(span - 0.2), Inches(3.0),
                  WHITE)
        add_oval(slide, left + Inches(span / 2 - 0.35), Inches(2.05),
                 Inches(0.6), Inches(0.6), color)
        tb(slide, left + Inches(span / 2 - 0.35), Inches(2.13), Inches(0.6),
           Inches(0.42), str(i + 1), size=14, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.95), Inches(span - 0.2), Inches(0.9), word,
           size=34, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left, Inches(4.05), Inches(span - 0.2), Inches(0.4),
           "☐", size=16, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.05), Inches(12.35), Inches(1.15),
              color)
    tb(slide, Inches(0.5), Inches(5.32), Inches(12.35), Inches(0.65),
       sentence, size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Georgia")
    i_we_you(slide, 6.42)


def story_slide(part, timing="67–77 min"):
    """Picture and part pill on the left, four big story lines on the right."""
    label, pic, color, light, lines, you_line, question = part
    slide, n = new_slide(f"🏠 Sam Builds a Little Shed — {label}", "STORY",
                         timing, "Reading Site", color)
    add_round(slide, Inches(0.5), Inches(1.42), Inches(3.5), Inches(4.8),
              light)
    tb(slide, Inches(0.5), Inches(1.75), Inches(3.5), Inches(1.6), pic,
       size=88, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.15), Inches(3.5), Inches(2.2), Inches(0.5),
              color)
    tb(slide, Inches(1.15), Inches(3.58), Inches(2.2), Inches(0.36),
       label.upper(), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.8), Inches(4.2), Inches(2.9), Inches(1.75),
              WHITE)
    tb(slide, Inches(1.0), Inches(4.32), Inches(2.5), Inches(0.32),
       "💬 QUICK QUESTION", size=10, bold=True, color=color)
    tb(slide, Inches(1.0), Inches(4.66), Inches(2.5), Inches(1.1), question,
       size=15, bold=True, color=INK)
    add_round(slide, Inches(4.25), Inches(1.42), Inches(8.6), Inches(3.45),
              WHITE)
    for i, line in enumerate(lines):
        tb(slide, Inches(4.6), Inches(1.62 + i * 0.82), Inches(8.0),
           Inches(0.7), line, size=25, bold=True, color=INK, font="Georgia")
    add_round(slide, Inches(4.25), Inches(5.05), Inches(8.6), Inches(1.15),
              light)
    tb(slide, Inches(4.6), Inches(5.16), Inches(8.0), Inches(0.32),
       "⭐ YOU READ THIS LINE ON YOUR OWN", size=11, bold=True, color=color)
    tb(slide, Inches(4.6), Inches(5.5), Inches(8.0), Inches(0.5), you_line,
       size=26, bold=True, color=INK, font="Georgia")
    i_we_you(slide, 6.42)


def game_cards(slide, games, start_index, width=3.9, gap=4.22):
    for i, (icon, name, prompt, cards, color, light) in enumerate(games):
        left = Inches(0.5 + i * gap)
        add_round(slide, left, Inches(1.9), Inches(width), Inches(4.3), light)
        add_round(slide, left, Inches(1.9), Inches(width), Inches(0.5), color)
        tb(slide, left, Inches(1.98), Inches(width), Inches(0.36),
           f"GAME {start_index + i} — {name}", size=12, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(width), Inches(0.62), icon,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(3.2), Inches(width - 0.5),
           Inches(0.5), prompt, size=14, bold=True, color=INK,
           align=PP_ALIGN.CENTER)
        for j, card in enumerate(cards):
            add_round(slide, left + Inches(0.4), Inches(3.9 + j * 0.85),
                      Inches(width - 0.8), Inches(0.72), WHITE)
            tb(slide, left + Inches(0.4), Inches(4.05 + j * 0.85),
               Inches(width - 0.8), Inches(0.5), card, size=14, bold=True,
               color=color, align=PP_ALIGN.CENTER)


# ------------------------------------------------------------------- slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), HAT)
    tb(slide, Inches(0.7), Inches(0.75), Inches(12), Inches(1.2), "🏗️",
       size=64, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.95), Inches(12), Inches(1.0),
       "BUILD IT!  READING CREW", size=44, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.95), Inches(12), Inches(0.5),
       "\"Become a Junior Builder Reader!\"", size=20,
       color=RGBColor(0xF0, 0xC5, 0x6B), align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, name) in enumerate(SITE_SCENE):
        left = Inches(1.35 + i * 1.78)
        add_round(slide, left, Inches(3.65), Inches(1.6), Inches(1.5),
                  RGBColor(0x27, 0x2F, 0x39))
        tb(slide, left, Inches(3.85), Inches(1.6), Inches(0.7), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(4.6), Inches(1.6), Inches(0.4), name, size=11,
           color=RGBColor(0xF0, 0xC5, 0x6B), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.1), Inches(5.4), Inches(7.1), Inches(1.0), HAT)
    tb(slide, Inches(3.1), Inches(5.62), Inches(7.1), Inches(0.6),
       "Grade 1  •  90 Minutes  •  Reading", size=24, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Building Site")
    fade(slide)


def s02_mission():
    slide, n = new_slide("🎯 Today's Reading Mission", "WELCOME", "0–7 min",
                         "Building Site", SITE)
    one_task(slide, "Five jobs today. We do them one at a time.", SITE)
    for i, (icon, name, detail) in enumerate(MISSION):
        left = Inches(0.5 + i * 2.47)
        add_round(slide, left, Inches(1.9), Inches(2.35), Inches(4.3),
                  L_SITE)
        add_round(slide, left, Inches(1.9), Inches(2.35), Inches(0.46), SITE)
        tb(slide, left, Inches(1.97), Inches(2.35), Inches(0.36),
           f"JOB {i + 1}", size=11, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        add_oval(slide, left + Inches(0.68), Inches(2.65), Inches(1.0),
                 Inches(1.0), WHITE)
        tb(slide, left + Inches(0.68), Inches(2.82), Inches(1.0),
           Inches(0.7), icon, size=26, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(3.9), Inches(2.35), Inches(0.5), name,
           size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(4.5), Inches(1.95), Inches(1.5),
           detail, size=14, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.35), Inches(12.35), Inches(0.5),
              HAT)
    tb(slide, Inches(0.5), Inches(6.45), Inches(12.35), Inches(0.36),
       "At the end you become a 🏆 READING BUILDER CHAMPION", size=15,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def s03_route():
    slide, n = new_slide("🗺️ Your Building Site Route", "WELCOME", "0–7 min",
                         "Building Site", SITE)
    one_task(slide, "Seven stops. We work at every one.", SITE)
    for i, (emoji, name, focus, timing, color, light) in enumerate(STOPS):
        left = Inches(0.5 + i * 1.78)
        add_round(slide, left, Inches(1.9), Inches(1.6), Inches(3.9), light)
        add_round(slide, left, Inches(1.9), Inches(1.6), Inches(0.44), color)
        tb(slide, left, Inches(1.96), Inches(1.6), Inches(0.34),
           f"STOP {i + 1}", size=10, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(1.6), Inches(0.9), emoji,
           size=34, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.05), Inches(3.5), Inches(1.5), Inches(0.62),
           name, size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.05), Inches(4.2), Inches(1.5), Inches(0.62),
           focus, size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.1), Inches(5.05), Inches(1.4),
                  Inches(0.4), WHITE)
        tb(slide, left + Inches(0.1), Inches(5.11), Inches(1.4), Inches(0.32),
           timing, size=9, bold=True, color=color, align=PP_ALIGN.CENTER)
        if i < len(STOPS) - 1:
            tb(slide, left + Inches(1.55), Inches(3.35), Inches(0.3),
               Inches(0.5), "→", size=18, bold=True, color=SOFT,
               align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.05), Inches(12.35), Inches(0.85),
              L_HAT)
    tb(slide, Inches(0.5), Inches(6.18), Inches(12.35), Inches(0.36),
       "🏆 AT THE FINISH", size=11, bold=True, color=HAT,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(6.5), Inches(12.35), Inches(0.36),
       "You read words, sentences and a whole story — and you become a "
       "Reading Builder Champion.", size=15, bold=True, color=INK,
       align=PP_ALIGN.CENTER)


def s04_look_around():
    slide, n = new_slide("👀 What Do You See?", "TALKING", "0–7 min",
                         "Building Site", SITE)
    one_task(slide, "No reading yet. Just look and tell me what you see.",
             SITE)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.2), Inches(4.4),
              L_SITE)
    tb(slide, Inches(0.7), Inches(2.1), Inches(4.8), Inches(1.5),
       "🏗️  🚚  🧱", size=58, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(3.5), Inches(4.8), Inches(1.5),
       "🦺  🔨  👷", size=58, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.6), Inches(0.6),
       "a busy building site", size=16, color=SOFT, align=PP_ALIGN.CENTER,
       italic=True)
    for i, (icon, question) in enumerate(TALK_QS):
        top = Inches(1.85 + i * 0.92)
        add_round(slide, Inches(6.0), top, Inches(6.85), Inches(0.78), WHITE)
        tb(slide, Inches(6.25), top + Inches(0.16), Inches(0.6), Inches(0.5),
           icon, size=18, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.0), top + Inches(0.15), Inches(5.6), Inches(0.5),
           question, size=20, bold=True, color=INK)
    add_round(slide, Inches(6.0), Inches(5.6), Inches(6.85), Inches(0.65),
              L_HAT)
    tb(slide, Inches(6.25), Inches(5.7), Inches(6.35), Inches(0.26),
       "SAY IT LIKE THIS", size=10, bold=True, color=HAT)
    tb(slide, Inches(6.25), Inches(5.98), Inches(6.35), Inches(0.36),
       "   ·   ".join(TALK_STARTERS), size=14, bold=True, color=INK)
    hint(slide, "Every answer counts here. One word is a good answer too.",
         6.42)


def s05_vocab_brick():
    vocab_slide(0)


def s06_vocab_crane():
    vocab_slide(1)


def s07_vocab_tool():
    vocab_slide(2)


def s08_word_bank():
    slide, n = new_slide("🧰 Site Word Bank", "SITE WORDS", "0–7 min",
                         "Building Site", TOOL)
    one_task(slide, "Nine more words from the site. I say it, then you say "
                    "it.", TOOL)
    for i, (pic, word, meaning) in enumerate(WORD_BANK):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.85 + row * 1.5)
        add_round(slide, left, top, Inches(3.9), Inches(1.35), WHITE)
        add_round(slide, left + Inches(0.15), top + Inches(0.18),
                  Inches(0.95), Inches(0.95), L_TOOL)
        tb(slide, left + Inches(0.15), top + Inches(0.33), Inches(0.95),
           Inches(0.65), pic, size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.25), top + Inches(0.2), Inches(2.5),
           Inches(0.45), word.upper(), size=21, bold=True, color=TOOL,
           font="Arial Black")
        tb(slide, left + Inches(1.25), top + Inches(0.72), Inches(2.5),
           Inches(0.5), meaning, size=12, color=DARK)
    hint(slide, "Do not drill these. Say each one once and move on.", 6.42)


def s09_sound_b():
    sound_slide(0)


def s10_sound_c():
    sound_slide(1)


def s11_sound_d():
    sound_slide(2)


def s12_sound_f():
    sound_slide(3)


def s13_sound_m():
    sound_slide(4)


def s14_sound_t():
    sound_slide(5)


def s15_which_tool_a():
    slide, n = new_slide("🎮 Sound Game — Which Tool?", "GAME", "7–17 min",
                         "Hard Hat Station", HAT)
    one_task(slide, "I say a sound. You point to the picture that starts "
                    "with it.", HAT)
    picture_rows(slide, TOOL_MATCH_A, 1, HAT)
    hint(slide, "Say each picture name out loud first. Listen to the very "
                "first sound.", 6.42)


def s16_which_tool_b():
    slide, n = new_slide("🎮 Which Tool? — Three More", "GAME", "7–17 min",
                         "Hard Hat Station", HAT)
    one_task(slide, "Same game. Listen for the first sound only.", HAT)
    picture_rows(slide, TOOL_MATCH_B, 4, HAT)
    hint(slide, "If he is unsure, name the three pictures for him and say "
                "them again slowly.", 6.42)


def s17_toolbox_drop():
    slide, n = new_slide("🧰 Put It in the Tool Box", "GAME", "7–17 min",
                         "Hard Hat Station", TOOL)
    one_task(slide, "Find the word that starts with the sound. It goes in "
                    "the box.", TOOL)
    tops = numbered_rows(slide, len(DROP_ROUNDS), 1, TOOL, top_start=1.95,
                         gap=1.5)
    for (sound, options), top in zip(DROP_ROUNDS, tops):
        add_round(slide, Inches(1.4), top + Inches(0.28), Inches(1.9),
                  Inches(0.78), TOOL)
        tb(slide, Inches(1.4), top + Inches(0.4), Inches(1.9), Inches(0.5),
           sound, size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(3.8 + j * 2.5)
            add_round(slide, left, top + Inches(0.28), Inches(2.3),
                      Inches(0.78), L_GREY)
            tb(slide, left, top + Inches(0.4), Inches(2.3), Inches(0.5), opt,
               size=24, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")
        add_round(slide, Inches(11.3), top + Inches(0.28), Inches(1.3),
                  Inches(0.78), L_TOOL)
        tb(slide, Inches(11.3), top + Inches(0.42), Inches(1.3), Inches(0.5),
           "🧰", size=22, align=PP_ALIGN.CENTER)
    hint(slide, "Read the three words to him first. He only has to listen "
                "and choose.", 6.42)


def s18_hardhat_done():
    badge_slide("🦺 Hard Hat Station — Done!", "7–17 min",
                "Hard Hat Station", "🦺", "FIRST SOUNDS BADGE",
                "You know six letter sounds now.",
                [("🧱", "B says /b/ — like brick"),
                 ("🏗️", "C says /c/ — like crane"),
                 ("⛏️", "D says /d/ — like dig"),
                 ("🔧", "F says /f/ — like fix"),
                 ("🗺️", "M says /m/ — like map"),
                 ("🔨", "T says /t/ — like tool")],
                "the Tool Box, where the vowels live", HAT, L_HAT)


def s19_vowel_intro():
    slide, n = new_slide("🔤 Open the Tool Box — A  E  I  O", "SHORT VOWELS",
                         "17–27 min", "Tool Box", TOOL)
    one_task(slide, "Four special letters. They go in the middle of a word.",
             TOOL)
    for i, (letter, sound, stretch, words, color, light) in enumerate(VOWELS):
        left = Inches(0.5 + i * 3.12)
        add_round(slide, left, Inches(1.9), Inches(2.9), Inches(4.3), light)
        tb(slide, left, Inches(2.15), Inches(2.9), Inches(1.6), letter,
           size=100, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        add_round(slide, left + Inches(0.55), Inches(3.9), Inches(1.8),
                  Inches(0.75), color)
        tb(slide, left + Inches(0.55), Inches(4.03), Inches(1.8),
           Inches(0.52), sound, size=26, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.3), Inches(4.85), Inches(2.3),
                  Inches(1.1), WHITE)
        tb(slide, left + Inches(0.3), Inches(4.96), Inches(2.3), Inches(0.3),
           "YOU HEAR IT IN", size=9, bold=True, color=SOFT,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(5.28), Inches(2.3), Inches(0.5),
           words[0][1], size=24, bold=True, color=INK,
           align=PP_ALIGN.CENTER, font="Arial Black")
    hint(slide, "Say all four in a row: /a/ /e/ /i/ /o/. Make it a chant.",
         6.42)


def s20_vowel_a():
    vowel_slide(0)


def s21_vowel_e():
    vowel_slide(1)


def s22_vowel_i():
    vowel_slide(2)


def s23_vowel_o():
    vowel_slide(3)


def s24_brick_sort():
    slide, n = new_slide("🧱 Sort the Bricks", "GAME", "17–27 min",
                         "Tool Box", BRICK)
    one_task(slide, "Which wall does each brick belong to? Listen to the "
                    "middle sound.", BRICK)
    for i, (letter, color, light) in enumerate(BRICK_BUCKETS):
        left = Inches(0.5 + i * 3.12)
        add_round(slide, left, Inches(1.9), Inches(2.9), Inches(2.5), light)
        add_round(slide, left, Inches(1.9), Inches(2.9), Inches(0.62), color)
        tb(slide, left, Inches(1.99), Inches(2.9), Inches(0.46),
           f"WALL {letter}", size=17, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        for j in range(2):
            add_round(slide, left + Inches(0.3), Inches(2.75 + j * 0.78),
                      Inches(2.3), Inches(0.66), WHITE)
            tb(slide, left + Inches(0.3), Inches(2.88 + j * 0.78),
               Inches(2.3), Inches(0.44), "____________", size=15,
               color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.65), Inches(12.35), Inches(0.4),
       "THE BRICKS — read each one, then put it on a wall", size=12,
       bold=True, color=BRICK)
    word_cards(slide, 0.5, 5.15, BRICK_WORDS, BRICK, L_BRICK, width=1.85,
               gap=2.08, height=1.0, size=28)
    hint(slide, "Cover the first and last letter. Only the middle letter is "
                "left.", 6.42)


def s25_missing_vowel_a():
    slide, n = new_slide("🎮 Missing Vowel — Fill the Gap", "GAME",
                         "17–27 min", "Tool Box", TOOL)
    one_task(slide, "The picture tells you the word. Pick the middle "
                    "letter.", TOOL)
    choice_rows(slide, MISSING_VOWEL_A, 1, TOOL)
    hint(slide, "Say the word, then say only the middle sound on its own.",
         6.42)


def s26_missing_vowel_b():
    slide, n = new_slide("🎮 Missing Vowel — Two More", "GAME", "17–27 min",
                         "Tool Box", TOOL)
    one_task(slide, "Same job. Say the word first, then choose.", TOOL)
    choice_rows(slide, MISSING_VOWEL_B, 4, TOOL, top_start=2.2, gap=1.7)
    add_round(slide, Inches(0.5), Inches(5.65), Inches(12.35), Inches(0.62),
              L_SITE)
    tb(slide, Inches(0.8), Inches(5.78), Inches(11.7), Inches(0.4),
       "⭐ MASTER BUILDER — read both finished words again, then use one in "
       "a sentence.", size=14, bold=True, color=SITE)
    hint(slide, "Try each letter in the gap and read it out loud. Only one "
                "will sound right.", 6.5)


def s27_toolbox_done():
    badge_slide("🔤 Tool Box — Done!", "17–27 min", "Tool Box", "🔤",
                "VOWEL BADGE", "You know the four short vowel sounds.",
                [("🗺️", "A says /a/ — map, bag, hat"),
                 ("🟥", "E says /e/ — red, bed, wet"),
                 ("🐘", "I says /i/ — big, sit, hit"),
                 ("🦊", "O says /o/ — fox, box, hot"),
                 ("🧱", "You sorted six bricks by their middle sound")],
                "the Brick Wall, where we build whole words", TOOL, L_TOOL)


def s28_blend_map():
    blend_slide(0)


def s29_blend_bag():
    blend_slide(1)


def s30_blend_hat():
    blend_slide(2)


def s31_blend_run():
    blend_slide(3)


def s32_blend_fox():
    blend_slide(4)


def s33_blend_big():
    blend_slide(5)


def s34_build_the_word():
    slide, n = new_slide("🔨 Build the Word", "GAME", "27–37 min",
                         "Brick Wall", BRICK)
    one_task(slide, "Say each block, push them together, then say the whole "
                    "word.", BRICK)
    for i, (letters, pic) in enumerate(BUILD_BLOCKS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(4.3),
                  L_BRICK)
        tb(slide, left, Inches(2.05), Inches(3.9), Inches(0.9), pic,
           size=44, align=PP_ALIGN.CENTER)
        for j, letter in enumerate(letters):
            l = left + Inches(0.3 + j * 1.2)
            add_round(slide, l, Inches(3.1), Inches(0.95), Inches(0.95),
                      BRICK)
            tb(slide, l, Inches(3.26), Inches(0.95), Inches(0.62), letter,
               size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
               font="Arial Black")
            if j < len(letters) - 1:
                tb(slide, l + Inches(0.98), Inches(3.28), Inches(0.19),
                   Inches(0.5), "+", size=16, bold=True, color=SOFT,
                   align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(4.2), Inches(3.9), Inches(0.42), "↓", size=22,
           bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.5), Inches(4.7), Inches(2.9),
                  Inches(1.1), WHITE)
        tb(slide, left + Inches(0.5), Inches(5.05), Inches(2.9), Inches(0.5),
           "say the whole word", size=13, color=SOFT,
           align=PP_ALIGN.CENTER, italic=True)
    hint(slide, "Move your finger under the blocks as you blend. Slow first, "
                "then fast.", 6.42)


def s35_mixed_up():
    slide, n = new_slide("🔤 Mixed-Up Letters", "GAME", "27–37 min",
                         "Brick Wall", SITE)
    one_task(slide, "These blocks fell over. Put them back in the right "
                    "order.", SITE)
    tops = numbered_rows(slide, len(SCRAMBLE), 1, SITE, top_start=1.9,
                         gap=1.12, height=0.98)
    for (letters, pic), top in zip(SCRAMBLE, tops):
        tb(slide, Inches(1.45), top + Inches(0.14), Inches(0.9), Inches(0.7),
           pic, size=28, align=PP_ALIGN.CENTER)
        for j, letter in enumerate(letters):
            l = Inches(2.6 + j * 1.05)
            add_round(slide, l, top + Inches(0.16), Inches(0.9),
                      Inches(0.66), L_SITE)
            tb(slide, l, top + Inches(0.26), Inches(0.9), Inches(0.48),
               letter, size=24, bold=True, color=SITE,
               align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(5.9), top + Inches(0.24), Inches(0.6), Inches(0.5),
           "→", size=20, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        for j in range(3):
            l = Inches(6.7 + j * 1.15)
            add_round(slide, l, top + Inches(0.16), Inches(1.0),
                      Inches(0.66), L_GREY)
            tb(slide, l, top + Inches(0.28), Inches(1.0), Inches(0.44), "?",
               size=20, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(10.4), top + Inches(0.16), Inches(2.2),
                  Inches(0.66), WHITE)
        tb(slide, Inches(10.4), top + Inches(0.29), Inches(2.2),
           Inches(0.44), "☐ I read it", size=13, color=SOFT,
           align=PP_ALIGN.CENTER)
    hint(slide, "Ask which letter he hears first. That block goes at the "
                "front.", 6.42)


def s36_blend_done():
    badge_slide("🧱 Word Builder — Done!", "27–37 min", "Brick Wall", "🔨",
                "BLENDING BADGE", "You built six whole words from sounds.",
                [("🗺️", "MAP  ·  /m/ /a/ /p/"),
                 ("🎒", "BAG  ·  /b/ /a/ /g/"),
                 ("🎩", "HAT  ·  /h/ /a/ /t/"),
                 ("🏃", "RUN  ·  /r/ /u/ /n/"),
                 ("🦊", "FOX  ·  /f/ /o/ /x/"),
                 ("🐘", "BIG  ·  /b/ /i/ /g/")],
                "word families — words that rhyme", BRICK, L_BRICK)


def s37_family_at():
    family_slide(0)


def s38_family_ig():
    family_slide(1)


def s39_family_un():
    family_slide(2)


def s40_family_op():
    family_slide(3)


def s41_family_sort():
    slide, n = new_slide("🧱 Build the Word Family Wall", "GAME", "37–44 min",
                         "Brick Wall", BRICK)
    one_task(slide, "Two walls. Put each word on the wall it rhymes with.",
             BRICK)
    for i, (name, color, light) in enumerate(SORT_BUCKETS):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(1.9), Inches(5.95), Inches(2.6), light)
        add_round(slide, left, Inches(1.9), Inches(5.95), Inches(0.66),
                  color)
        tb(slide, left, Inches(2.0), Inches(5.95), Inches(0.5),
           f"{name} WALL", size=19, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        for j in range(3):
            add_round(slide, left + Inches(0.3 + j * 1.83), Inches(2.85),
                      Inches(1.65), Inches(1.3), WHITE)
            tb(slide, left + Inches(0.3 + j * 1.83), Inches(3.3),
               Inches(1.65), Inches(0.5), "________", size=14, color=SOFT,
               align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.7), Inches(12.35), Inches(0.4),
       "THE WORDS — read each one out loud before you place it", size=12,
       bold=True, color=BRICK)
    word_cards(slide, 0.5, 5.2, SORT_WORDS, BRICK, L_BRICK, width=1.85,
               gap=2.08, height=0.95, size=28)
    hint(slide, "Say the word twice. Listen only to the ending.", 6.42)


def s42_odd_one():
    slide, n = new_slide("🎮 Who Does Not Belong?", "GAME", "37–44 min",
                         "Brick Wall", SIGN)
    one_task(slide, "Two words rhyme. One does not. Find the one that does "
                    "not fit.", SIGN)
    tops = numbered_rows(slide, len(ODD_ONE), 1, SIGN, top_start=2.0,
                         gap=1.45, height=1.25)
    for words, top in zip(ODD_ONE, tops):
        for j, word in enumerate(words):
            left = Inches(1.7 + j * 3.6)
            add_round(slide, left, top + Inches(0.22), Inches(3.2),
                      Inches(0.82), L_SIGN)
            tb(slide, left, top + Inches(0.34), Inches(3.2), Inches(0.58),
               word, size=30, bold=True, color=SIGN, align=PP_ALIGN.CENTER,
               font="Arial Black")
    hint(slide, "Read all three out loud. Two will sound like each other.",
         6.42)


def s43_rhyme():
    slide, n = new_slide("⭐ Rhyme Builder", "GAME", "37–44 min",
                         "Brick Wall", HAT)
    one_task(slide, "Which word rhymes with the big word?", HAT)
    tops = numbered_rows(slide, len(RHYMES), 1, HAT, top_start=2.0, gap=1.45,
                         height=1.25)
    for (word, pic, options), top in zip(RHYMES, tops):
        tb(slide, Inches(1.45), top + Inches(0.3), Inches(0.8), Inches(0.7),
           pic, size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.4), top + Inches(0.22), Inches(2.6),
                  Inches(0.82), HAT)
        tb(slide, Inches(2.4), top + Inches(0.34), Inches(2.6), Inches(0.58),
           word, size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(5.5 + j * 2.45)
            add_round(slide, left, top + Inches(0.22), Inches(2.2),
                      Inches(0.82), L_GREY)
            tb(slide, left, top + Inches(0.34), Inches(2.2), Inches(0.58),
               opt, size=26, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")
    hint(slide, "Rhyming words end the same. Only the first sound changes.",
         6.42)


def s44_builder_says():
    slide, n = new_slide("🦺 Builder Says — Stretch Break", "BREAK",
                         "44–49 min", "Brick Wall", GREEN)
    one_task(slide, "Stay in your seat. Only move when you hear \"Builder "
                    "says\".", GREEN)
    for i, (icon, call, note) in enumerate(BUILDER_SAYS):
        top = Inches(1.95 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98),
                  L_GREEN)
        add_oval(slide, Inches(0.8), top + Inches(0.19), Inches(0.6),
                 Inches(0.6), WHITE)
        tb(slide, Inches(0.8), top + Inches(0.27), Inches(0.6), Inches(0.45),
           icon, size=18, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.7), top + Inches(0.24), Inches(7.5), Inches(0.5),
           call, size=22, bold=True, color=INK)
        add_round(slide, Inches(9.6), top + Inches(0.22), Inches(3.0),
                  Inches(0.54), WHITE)
        tb(slide, Inches(9.6), top + Inches(0.32), Inches(3.0), Inches(0.38),
           note, size=14, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    hint(slide, "Keep it to two minutes. Stop while he is still enjoying "
                "it.", 6.5)


def s45_builder_reads():
    slide, n = new_slide("🦺 Builder Says — Now With Words", "BREAK",
                         "44–49 min", "Brick Wall", GREEN)
    one_task(slide, "Same game, but now the job is a reading job.", GREEN)
    for i, (pic, call, note) in enumerate(BUILDER_READS):
        top = Inches(1.95 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98),
                  L_GREEN)
        tb(slide, Inches(0.8), top + Inches(0.22), Inches(0.7), Inches(0.6),
           pic, size=24, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.7), top + Inches(0.24), Inches(7.5), Inches(0.5),
           call, size=22, bold=True, color=INK)
        add_round(slide, Inches(9.6), top + Inches(0.22), Inches(3.0),
                  Inches(0.54), WHITE)
        tb(slide, Inches(9.6), top + Inches(0.32), Inches(3.0), Inches(0.38),
           note, size=14, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    hint(slide, "Write each word on paper first so he has something to "
                "point at.", 6.5)


def s46_sight_i():
    sight_slide(0)


def s47_sight_see():
    sight_slide(1)


def s48_sight_a_the():
    sight_slide(2)


def s49_read_brick():
    build_slide(0)


def s50_read_truck():
    build_slide(1)


def s51_sight_my_is_can():
    sight_slide(3)


def s52_read_hat_red():
    build_slide(2)


def s53_read_can_run():
    build_slide(3)


def puzzle_row(slide, top, cards, pic, slots, color, light):
    tb(slide, Inches(0.7), Inches(top + 0.6), Inches(1.0), Inches(0.9),
       pic, size=42, align=PP_ALIGN.CENTER)
    tb(slide, Inches(1.85), Inches(top + 0.14), Inches(5.0), Inches(0.3),
       "MIXED-UP CARDS", size=10, bold=True, color=SOFT)
    for i, card in enumerate(cards):
        left = Inches(1.85 + i * 1.32)
        add_round(slide, left, Inches(top + 0.52), Inches(1.2), Inches(1.15),
                  light)
        tb(slide, left, Inches(top + 0.85), Inches(1.2), Inches(0.5), card,
           size=19, bold=True, color=color, align=PP_ALIGN.CENTER)
    tb(slide, Inches(7.3), Inches(top + 0.14), Inches(5.3), Inches(0.3),
       "PUT THEM IN ORDER", size=10, bold=True, color=color)
    for i in range(slots):
        left = Inches(7.3 + i * 1.32)
        add_round(slide, left, Inches(top + 0.52), Inches(1.2), Inches(1.15),
                  WHITE)
        tb(slide, left, Inches(top + 0.9), Inches(1.2), Inches(0.42),
           str(i + 1), size=15, color=SOFT, align=PP_ALIGN.CENTER)


def s54_puzzle_a():
    slide, n = new_slide("🧩 Build the Sentence Wall", "SENTENCES",
                         "58–67 min", "Safety Signs", SIGN)
    one_task(slide, "The word cards fell over. Build the sentence again.",
             SIGN)
    for i, (cards, pic, slots) in enumerate(PUZZLES[:2]):
        top = 1.95 + i * 2.2
        add_round(slide, Inches(0.5), Inches(top), Inches(12.35),
                  Inches(1.95), L_GREY)
        puzzle_row(slide, top, cards, pic, slots, SIGN, L_SIGN)
    add_round(slide, Inches(0.5), Inches(6.35), Inches(12.35), Inches(0.5),
              L_SITE)
    tb(slide, Inches(0.8), Inches(6.44), Inches(11.7), Inches(0.36),
       "🔧 BUILDER HINT — the sentence always starts with the card that has "
       "a capital letter.", size=13, bold=True, color=SITE)


def s55_puzzle_b():
    slide, n = new_slide("🧩 Two More Sentence Walls", "SENTENCES",
                         "58–67 min", "Safety Signs", SIGN)
    one_task(slide, "Same job. Read each card before you place it.", SIGN)
    for i, (cards, pic, slots) in enumerate(PUZZLES[2:]):
        top = 1.95 + i * 2.2
        add_round(slide, Inches(0.5), Inches(top), Inches(12.35),
                  Inches(1.95), L_GREY)
        puzzle_row(slide, top, cards, pic, slots, SIGN, L_SIGN)
    add_round(slide, Inches(0.5), Inches(6.35), Inches(12.35), Inches(0.5),
              L_TOOL)
    tb(slide, Inches(0.8), Inches(6.44), Inches(11.7), Inches(0.36),
       "⭐ MASTER BUILDER — read the finished sentence without stopping "
       "between the words.", size=13, bold=True, color=TOOL)


def s56_sentence_match():
    slide, n = new_slide("🖼️ Sentence and Picture Match", "GAME",
                         "58–67 min", "Safety Signs", TOOL)
    one_task(slide, "Read the sentence, then point to the picture it tells "
                    "about.", TOOL)
    for i, (sentence, pics) in enumerate(SENT_MATCH):
        top = Inches(1.95 + i * 2.35)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(2.1),
                  L_GREY)
        add_round(slide, Inches(0.8), top + Inches(0.25), Inches(5.0),
                  Inches(0.95), WHITE)
        tb(slide, Inches(0.9), top + Inches(0.42), Inches(4.8), Inches(0.62),
           sentence, size=26, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Georgia")
        add_round(slide, Inches(0.8), top + Inches(1.32), Inches(5.0),
                  Inches(0.5), TOOL)
        tb(slide, Inches(0.8), top + Inches(1.4), Inches(5.0), Inches(0.36),
           "I READ  →  WE READ  →  YOU READ", size=12, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        for j, pic in enumerate(pics):
            left = Inches(6.3 + j * 2.25)
            add_round(slide, left, top + Inches(0.25), Inches(2.0),
                      Inches(1.55), WHITE)
            tb(slide, left, top + Inches(0.4), Inches(2.0), Inches(0.85),
               pic, size=42, align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(1.25), Inches(2.0), Inches(0.4),
               "☐", size=15, color=SOFT, align=PP_ALIGN.CENTER)
    hint(slide, "Read the sentence twice. The naming word tells you the "
                "picture.", 6.5)


def s57_real_silly():
    slide, n = new_slide("✅ Real or Silly?", "GAME", "58–67 min",
                         "Safety Signs", WOOD)
    one_task(slide, "Read it, then tell me: could that really happen?",
             WOOD)
    tops = numbered_rows(slide, len(REAL_SILLY), 1, WOOD, top_start=1.95,
                         gap=1.12, height=0.98)
    for sentence, top in zip(REAL_SILLY, tops):
        tb(slide, Inches(1.5), top + Inches(0.22), Inches(6.8), Inches(0.6),
           sentence, size=23, bold=True, color=INK, font="Georgia")
        for j, (mark, label, color, light) in enumerate(
                [("✅", "REAL", GREEN, L_GREEN),
                 ("❌", "SILLY", BRICK, L_BRICK)]):
            left = Inches(8.7 + j * 2.0)
            add_round(slide, left, top + Inches(0.19), Inches(1.85),
                      Inches(0.6), light)
            tb(slide, left, top + Inches(0.3), Inches(1.85), Inches(0.42),
               f"{mark}  {label}", size=15, bold=True, color=color,
               align=PP_ALIGN.CENTER)
    hint(slide, "Silly sentences are meant to be funny. Laugh with him, "
                "then ask why.", 6.5)


def s58_signs_done():
    badge_slide("🚧 Safety Sign Station — Done!", "58–67 min",
                "Safety Signs", "🚧", "SENTENCE BADGE",
                "You read seven sight words and four real sentences.",
                [("👀", "I  ·  see  ·  a  ·  the"),
                 ("🚧", "my  ·  is  ·  can"),
                 ("🧱", "I see a brick."),
                 ("🚚", "I see the truck."),
                 ("🎩", "My hat is red.   ·   I can run.")],
                "the Reading Site — a whole story", SIGN, L_SIGN)


def s59_story_intro():
    slide, n = new_slide("📖 Our Story — Sam Builds a Little Shed", "STORY",
                         "67–77 min", "Reading Site", WOOD)
    one_task(slide, "Before we read: look at the picture and make a guess.",
             WOOD)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.6), Inches(4.4),
              L_WOOD)
    tb(slide, Inches(0.5), Inches(2.2), Inches(5.6), Inches(1.8), "👷🏠",
       size=92, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.1), Inches(4.3), Inches(4.4), Inches(1.0),
              WOOD)
    tb(slide, Inches(1.1), Inches(4.52), Inches(4.4), Inches(0.62),
       "SAM BUILDS A LITTLE SHED", size=19, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(1.1), Inches(5.45), Inches(4.4), Inches(0.5),
       "five short parts", size=15, color=SOFT, align=PP_ALIGN.CENTER,
       italic=True)
    add_round(slide, Inches(6.4), Inches(1.85), Inches(6.45), Inches(1.0),
              WHITE)
    tb(slide, Inches(6.7), Inches(1.98), Inches(5.9), Inches(0.36),
       "QUESTION 1", size=10, bold=True, color=WOOD)
    tb(slide, Inches(6.7), Inches(2.3), Inches(5.9), Inches(0.5),
       "Who do you think Sam is?", size=21, bold=True, color=INK)
    tb(slide, Inches(6.4), Inches(3.0), Inches(6.45), Inches(0.4),
       "QUESTION 2 — What do you think Sam will build?", size=13, bold=True,
       color=WOOD)
    for i, (pic, label) in enumerate(STORY_PREDICT):
        left = Inches(6.4 + i * 2.2)
        add_round(slide, left, Inches(3.5), Inches(1.95), Inches(2.0),
                  L_WOOD)
        tb(slide, left, Inches(3.75), Inches(1.95), Inches(0.9), pic,
           size=40, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.7), Inches(1.75),
           Inches(0.5), label, size=13, bold=True, color=INK,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(5.15), Inches(1.95), Inches(0.36), "☐",
           size=15, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(6.4), Inches(5.7), Inches(6.45), Inches(0.55),
              L_GREEN)
    tb(slide, Inches(6.7), Inches(5.8), Inches(5.9), Inches(0.36),
       "Any guess is a good guess. We check at the end.", size=13,
       bold=True, color=GREEN)
    hint(slide, "Guessing first makes the story easier to follow.", 6.42)


def s60_story_1():
    story_slide(STORY[0])


def s61_story_2():
    story_slide(STORY[1])


def s62_story_3():
    story_slide(STORY[2])


def s63_story_4():
    story_slide(STORY[3], "77–85 min")


def s64_story_5():
    story_slide(STORY[4], "77–85 min")


def s65_detect_a():
    slide, n = new_slide("🔎 Story Builder Detective", "QUESTIONS",
                         "77–85 min", "Reading Site", SITE)
    one_task(slide, "Two questions. The answers are in the story.", SITE)
    question_rows(slide, DETECT_A, 1, SITE)
    hint(slide, "Go back to the part and read it again together before "
                "answering.", 6.5)


def s66_detect_b():
    slide, n = new_slide("🔎 Two More Questions", "QUESTIONS", "77–85 min",
                         "Reading Site", SITE)
    one_task(slide, "Same job. Find the part, then choose.", SITE)
    question_rows(slide, DETECT_B, 3, SITE)
    hint(slide, "If he guesses, do not say no. Say \"let's check\" and read "
                "it again.", 6.5)


def s67_sequence():
    slide, n = new_slide("🧩 Put the Story in Order", "QUESTIONS",
                         "77–85 min", "Reading Site", BRICK)
    one_task(slide, "What happened first? Number the cards 1 to 4.", BRICK)
    for i, (label, pic, line) in enumerate(SEQUENCE):
        left = Inches(0.5 + i * 3.12)
        add_round(slide, left, Inches(1.9), Inches(2.9), Inches(3.5),
                  L_BRICK)
        add_round(slide, left, Inches(1.9), Inches(2.9), Inches(0.5), BRICK)
        tb(slide, left, Inches(1.98), Inches(2.9), Inches(0.36),
           f"CARD {label}", size=12, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.6), Inches(2.9), Inches(0.9), pic, size=40,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.6), Inches(2.5), Inches(1.0),
           line, size=17, bold=True, color=INK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.95), Inches(4.65), Inches(1.0),
                  Inches(0.6), WHITE)
        tb(slide, left + Inches(0.95), Inches(4.78), Inches(1.0),
           Inches(0.42), "☐", size=18, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.6), Inches(12.35), Inches(0.7),
              L_GREY)
    tb(slide, Inches(0.8), Inches(5.83), Inches(3.2), Inches(0.36),
       "WRITE THE ORDER HERE", size=12, bold=True, color=SOFT)
    for i in range(4):
        left = Inches(4.22 + i * 1.3)
        add_round(slide, left, Inches(5.76), Inches(1.0), Inches(0.42),
                  WHITE)
        if i < 3:
            tb(slide, left + Inches(1.04), Inches(5.8), Inches(0.22),
               Inches(0.36), "→", size=15, bold=True, color=SOFT,
               align=PP_ALIGN.CENTER)
    hint(slide, "Ask what Sam had at the very start of the story.", 6.5)


def s68_evidence():
    slide, n = new_slide("🔎 Show Me the Answer", "QUESTIONS", "77–85 min",
                         "Reading Site", TOOL)
    one_task(slide, "Do not just say it. Go back and point to the line.",
             TOOL)
    for i, (question, where, color, light) in enumerate(EVIDENCE):
        top = Inches(1.95 + i * 1.42)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.25),
                  light)
        add_oval(slide, Inches(0.8), top + Inches(0.34), Inches(0.58),
                 Inches(0.58), color)
        tb(slide, Inches(0.8), top + Inches(0.42), Inches(0.58),
           Inches(0.42), str(i + 1), size=14, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.7), top + Inches(0.34), Inches(6.2), Inches(0.6),
           question, size=22, bold=True, color=INK)
        add_round(slide, Inches(8.1), top + Inches(0.3), Inches(2.1),
                  Inches(0.6), color)
        tb(slide, Inches(8.1), top + Inches(0.42), Inches(2.1), Inches(0.4),
           where, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(10.4), top + Inches(0.3), Inches(2.2),
                  Inches(0.6), WHITE)
        tb(slide, Inches(10.4), top + Inches(0.42), Inches(2.2),
           Inches(0.4), "☐ I pointed to it", size=13, color=SOFT,
           align=PP_ALIGN.CENTER)
    hint(slide, "Pointing at the line is the habit we want, more than the "
                "answer itself.", 6.42)


def s69_final_challenge():
    slide, n = new_slide("🏗️ Final Build Challenge", "CHALLENGE",
                         "85–88 min", "Builder HQ", GREEN)
    one_task(slide, "Three last jobs. You already know all of them.", GREEN)
    for i, (icon, label, task, pic, color, light) in enumerate(FINAL_ROUNDS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(4.4), light)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(0.6), color)
        tb(slide, left, Inches(1.96), Inches(3.9), Inches(0.4),
           f"{icon}  {label}", size=15, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.65), Inches(3.9), Inches(1.1), pic,
           size=56, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(3.9), Inches(3.3),
                  Inches(1.3), WHITE)
        tb(slide, left + Inches(0.35), Inches(4.25), Inches(3.2),
           Inches(0.7), task, size=22, bold=True, color=INK,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.05), Inches(5.35), Inches(1.8),
                  Inches(0.6), WHITE)
        tb(slide, left + Inches(1.05), Inches(5.46), Inches(1.8),
           Inches(0.42), "☐ done", size=15, bold=True, color=color,
           align=PP_ALIGN.CENTER)
    hint(slide, "Hints are still allowed here. Finishing is what matters.",
         6.42)


def s70_champion():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), HAT)
    for x, y in [(0.35, 3.9), (12.3, 3.9), (0.6, 5.4), (12.1, 5.4)]:
        tb(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.7), "⭐",
           size=24, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(0.72), Inches(12), Inches(1.1), "🏆",
       size=56, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.82), Inches(12), Inches(0.9),
       "READING BUILDER CHAMPION!", size=40, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.78), Inches(12), Inches(0.5),
       "Look what you built today —", size=18,
       color=RGBColor(0xF0, 0xC5, 0x6B), align=PP_ALIGN.CENTER, italic=True)
    for i, (icon, line) in enumerate(STAR_LINES):
        left = Inches(1.35 + i * 2.15)
        add_round(slide, left, Inches(3.5), Inches(1.95), Inches(1.5),
                  RGBColor(0x27, 0x2F, 0x39))
        tb(slide, left, Inches(3.67), Inches(1.95), Inches(0.6), icon,
           size=20, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.12), Inches(4.27), Inches(1.7),
           Inches(0.62), line, size=11, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.9), Inches(5.2), Inches(9.5), Inches(1.05),
              HAT)
    tb(slide, Inches(1.9), Inches(5.44), Inches(9.5), Inches(0.62),
       "\"I CAN BUILD MY READING SKILLS!\"", size=29, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(1.2), Inches(6.45), Inches(10.9), Inches(0.42),
       "🏗️  📚  ⭐", size=16, bold=True,
       color=RGBColor(0xF0, 0xC5, 0x6B), align=PP_ALIGN.CENTER)
    footer(slide, n, "88–90 min", "Builder HQ")
    fade(slide)


def s71_games_a():
    slide, n = new_slide("🎲 Extra Game Bank — If There Is Time", "OPTIONAL",
                         "", "Builder HQ", TOOL)
    one_task(slide, "Six spare games. Play any of them, in any order.", TOOL)
    game_cards(slide, GAMES[:3], 1)
    hint(slide, "Every game uses today's words only. Nothing new to learn.",
         6.45)


def s72_games_b():
    slide, n = new_slide("🎲 Extra Game Bank — Three More", "OPTIONAL", "",
                         "Builder HQ", SITE)
    one_task(slide, "Three more, plus when to use them.", SITE)
    game_cards(slide, GAMES[3:], 4, width=2.9, gap=3.11)
    add_round(slide, Inches(9.83), Inches(1.9), Inches(3.02), Inches(4.3),
              L_GREY)
    tb(slide, Inches(10.05), Inches(2.1), Inches(2.6), Inches(0.4),
       "WHEN TO USE THESE", size=12, bold=True, color=SITE)
    bullets(slide, Inches(10.05), Inches(2.6), Inches(2.6), Inches(3.4),
            GAME_WHEN, size=12, sp=10)
    hint(slide, "Stop a game while he is still enjoying it, not after.",
         6.45)


def s73_support():
    slide, n = new_slide("🧰 Support System — For the Teacher", "TEACHER", "",
                         "Builder HQ", SITE)
    one_task(slide, "Levels, the hint ladder, and how to handle a hard "
                    "sentence.", SITE)
    for i, (icon, name, who, action, color, light) in enumerate(
            SUPPORT_LEVELS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(2.25),
                  light)
        tb(slide, left + Inches(0.25), Inches(2.02), Inches(0.5),
           Inches(0.5), icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.9), Inches(2.02), Inches(2.9), Inches(0.5),
           name, size=14, bold=True, color=color)
        tb(slide, left + Inches(0.3), Inches(2.6), Inches(3.3), Inches(0.5),
           who, size=12, color=DARK)
        add_round(slide, left + Inches(0.3), Inches(3.15), Inches(3.3),
                  Inches(0.78), WHITE)
        tb(slide, left + Inches(0.45), Inches(3.26), Inches(3.0),
           Inches(0.58), action, size=11, color=INK)
    tb(slide, Inches(0.5), Inches(4.25), Inches(6.0), Inches(0.4),
       "HINT LADDER — never reveal the answer first", size=13, bold=True,
       color=SITE)
    for i, (label, text, color) in enumerate(SUPPORT_LADDER):
        top = Inches(4.66 + i * 0.46)
        add_round(slide, Inches(0.5), top, Inches(6.0), Inches(0.42), L_GREY)
        add_round(slide, Inches(0.62), top + Inches(0.05), Inches(1.2),
                  Inches(0.32), color)
        tb(slide, Inches(0.62), top + Inches(0.07), Inches(1.2),
           Inches(0.28), label, size=9, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.0), top + Inches(0.05), Inches(4.3), Inches(0.32),
           text, size=12, bold=True, color=INK)
    tb(slide, Inches(6.85), Inches(4.25), Inches(3.1), Inches(0.4),
       "A HARD SENTENCE, STEP BY STEP", size=12, bold=True, color=TOOL)
    bullets(slide, Inches(6.85), Inches(4.72), Inches(3.1), Inches(2.1),
            HARD_WORD, size=10, sp=4)
    tb(slide, Inches(10.2), Inches(4.25), Inches(2.65), Inches(0.4),
       "SAY THIS", size=12, bold=True, color=GREEN)
    bullets(slide, Inches(10.2), Inches(4.72), Inches(2.65), Inches(2.1),
            PRAISE, size=10, color=GREEN, sp=4)


def s74_assessment():
    slide, n = new_slide("📋 End-of-Class Assessment", "TEACHER", "",
                         "Builder HQ", HAT)
    one_task(slide, "Tick one box per skill right after the lesson.", HAT)
    heads = ["SKILL", "INDEPENDENT", "WITH HELP", "NEEDS MORE PRACTICE"]
    widths = [5.15, 2.4, 2.4, 2.4]
    lefts = [0.5]
    for w in widths[:-1]:
        lefts.append(lefts[-1] + w)
    add_rect(slide, Inches(0.5), Inches(1.82), Inches(12.35), Inches(0.42),
             HAT)
    for left, w, head in zip(lefts, widths, heads):
        tb(slide, Inches(left), Inches(1.86), Inches(w), Inches(0.34), head,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = Inches(2.3 + i * 0.35)
        add_rect(slide, Inches(0.5), top, Inches(12.35), Inches(0.33),
                 WHITE if i % 2 == 0 else L_GREY)
        tb(slide, Inches(0.7), top + Inches(0.02), Inches(4.8), Inches(0.3),
           skill, size=11, bold=True, color=INK)
        for left, w in zip(lefts[1:], widths[1:]):
            tb(slide, Inches(left), top + Inches(0.01), Inches(w),
               Inches(0.3), "☐", size=13, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(5.95), Inches(12.35), Inches(0.4),
       "NEVER USE THE LABELS  weak  ·  slow  ·  behind  ·  low level  —  "
       "describe the next step instead.", size=12, bold=True, color=BRICK)
    add_round(slide, Inches(0.5), Inches(6.4), Inches(12.35), Inches(0.5),
              L_HAT)
    for i, item in enumerate(REVIEW_NEXT):
        tb(slide, Inches(0.8 + i * 4.1), Inches(6.52), Inches(3.9),
           Inches(0.34), f"{item}:  ____________", size=12, bold=True,
           color=INK)


def s75_answer_key():
    slide, n = new_slide("🔑 Answer Key", "TEACHER", "", "Builder HQ", BRICK)
    one_task(slide, "For the teacher only.", BRICK)
    cols = [
        ("SOUNDS AND VOWELS (15–26)",
         ["Which tool: brick · crane · map ·",
          " tool · dig · fix",
          "Tool box: MAP · BAG · FIX",
          "Brick sort: A map, bag · E red ·",
          " I big · O fox, hot",
          "Missing vowel: I · A · E · O · O"], HAT),
        ("BLENDING AND FAMILIES (34–43)",
         ["Build the word: BAG · MAP · FOX",
          "Mixed up: MAP · BIG · HAT · FOX",
          "Family wall: -AT cat, hat, mat ·",
          " -UN sun, run, fun",
          "Odd one out: SUN · BIG · MAP",
          "Rhymes: HAT · FOX · FUN"], BRICK),
        ("SENTENCES (54–57)",
         ["Puzzles: I see a brick. · My hat is red.",
          " I can run. · The truck is big.",
          "Picture match: the truck · the brick",
          "Real or silly: real · silly · real · silly",
          "Sight words: I, see, a, the, my, is, can"], SIGN),
        ("STORY AND QUESTIONS (59–68)",
         ["Q1 Sam · Q2 Red · Q3 A hammer ·",
          " Q4 A shed",
          "Order: C → B → A → D",
          "Evidence: wood (Part 2) · a hard hat",
          " (Part 3) · red (Part 5)",
          "Story words: 104"], WOOD),
    ]
    for i, (head, lines, color) in enumerate(cols):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.85 + row * 2.4)
        add_round(slide, left, top, Inches(5.95), Inches(2.2), L_GREY)
        add_round(slide, left, top, Inches(5.95), Inches(0.46), color)
        tb(slide, left + Inches(0.25), top + Inches(0.06), Inches(5.45),
           Inches(0.34), head, size=12, bold=True, color=WHITE)
        bullets(slide, left + Inches(0.25), top + Inches(0.58), Inches(5.45),
                Inches(1.5), lines, size=10, sp=3)
    add_round(slide, Inches(0.5), Inches(6.7), Inches(12.35), Inches(0.42),
              L_BRICK)
    tb(slide, Inches(0.8), Inches(6.76), Inches(11.7), Inches(0.32),
       "Accept any sensible answer for the rhyme game, the predictions and "
       "the speaking tasks.", size=11, bold=True, color=BRICK)


BUILDERS = [
    s01_title, s02_mission, s03_route, s04_look_around, s05_vocab_brick,
    s06_vocab_crane, s07_vocab_tool, s08_word_bank, s09_sound_b, s10_sound_c,
    s11_sound_d, s12_sound_f, s13_sound_m, s14_sound_t, s15_which_tool_a,
    s16_which_tool_b, s17_toolbox_drop, s18_hardhat_done, s19_vowel_intro,
    s20_vowel_a, s21_vowel_e, s22_vowel_i, s23_vowel_o, s24_brick_sort,
    s25_missing_vowel_a, s26_missing_vowel_b, s27_toolbox_done,
    s28_blend_map, s29_blend_bag, s30_blend_hat, s31_blend_run,
    s32_blend_fox, s33_blend_big, s34_build_the_word, s35_mixed_up,
    s36_blend_done, s37_family_at, s38_family_ig, s39_family_un,
    s40_family_op, s41_family_sort, s42_odd_one, s43_rhyme,
    s44_builder_says, s45_builder_reads, s46_sight_i, s47_sight_see,
    s48_sight_a_the, s49_read_brick, s50_read_truck, s51_sight_my_is_can,
    s52_read_hat_red, s53_read_can_run, s54_puzzle_a, s55_puzzle_b,
    s56_sentence_match, s57_real_silly, s58_signs_done, s59_story_intro,
    s60_story_1, s61_story_2, s62_story_3, s63_story_4, s64_story_5,
    s65_detect_a, s66_detect_b, s67_sequence, s68_evidence,
    s69_final_challenge, s70_champion, s71_games_a, s72_games_b,
    s73_support, s74_assessment, s75_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL}, built {len(prs.slides)}"

OUT = "Grade1_Build_It_Reading_Crew_90min.pptx"
prs.save(OUT)

story_words = sum(len(line.split()) for part in STORY for line in part[4])
with_notes = sum(1 for s in prs.slides
                 if s.has_notes_slide
                 and s.notes_slide.notes_text_frame.text.strip())
print(f"saved {OUT}")
print(f"slides: {len(prs.slides)}")
print(f"story words: {story_words}")
print(f"slides with notes: {with_notes}")
