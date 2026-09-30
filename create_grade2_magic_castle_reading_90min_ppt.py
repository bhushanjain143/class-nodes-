"""Grade 2 reading lesson - 90 minutes, 73 slides, no speaker notes.

"The Magic Castle Reading Quest" - the child is a Junior Castle Reader moving
through the Castle Gate, the Key Room, the Mirror Hall, the Magic Library, the
Puzzle Room and the King's Reading Room. Reading is built one step at a time:
hear, say, sound, blend, read, understand.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Castle Hint strip, and slides 69-73 hold the optional game
bank, the support system, the assessment table and the answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x19, 0x1F, 0x2B)
DARK = RGBColor(0x2E, 0x37, 0x44)
SOFT = RGBColor(0x7E, 0x87, 0x93)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAGE = RGBColor(0xF8, 0xF7, 0xF3)
STONE = RGBColor(0x4A, 0x5A, 0x73)
GOLD = RGBColor(0xC2, 0x8B, 0x0B)
SAPPHIRE = RGBColor(0x24, 0x5F, 0xA6)
EMERALD = RGBColor(0x1B, 0x7A, 0x5A)
ROYAL = RGBColor(0x6B, 0x3F, 0xA0)
CRIMSON = RGBColor(0xB8, 0x34, 0x2F)
L_STONE = RGBColor(0xEB, 0xEE, 0xF2)
L_GOLD = RGBColor(0xFC, 0xF2, 0xDA)
L_SAPPHIRE = RGBColor(0xE4, 0xED, 0xF8)
L_EMERALD = RGBColor(0xE0, 0xF1, 0xEB)
L_ROYAL = RGBColor(0xF1, 0xEA, 0xF9)
L_CRIMSON = RGBColor(0xFA, 0xE8, 0xE7)
L_GREY = RGBColor(0xF1, 0xF1, 0xEE)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 73
_counter = {"n": 0}

# ------------------------------------------------------------------ content

STOPS = [("🚪", "Castle Gate", "First sounds", "7–17 min", STONE, L_STONE),
         ("🗝️", "Key Room", "Short vowels", "17–27 min", GOLD, L_GOLD),
         ("🪞", "Mirror Hall", "Building words", "27–37 min", SAPPHIRE,
          L_SAPPHIRE),
         ("📚", "Magic Library", "Families & sight words", "37–58 min",
          EMERALD, L_EMERALD),
         ("🧩", "Puzzle Room", "Sentences", "58–67 min", ROYAL, L_ROYAL),
         ("👑", "King's Room", "Story & questions", "67–90 min", CRIMSON,
          L_CRIMSON)]

MISSION = [("📖", "Read", "Words, sentences and a real story."),
           ("🔤", "Build words", "Push the sounds together."),
           ("🎮", "Play games", "Matching, sorting and unlocking doors."),
           ("🧩", "Solve puzzles", "Put mixed-up words back in order.")]

CASTLE_SCENE = [("🏰", "castle"), ("🗝️", "keys"), ("🚪", "doors"),
                ("📚", "books"), ("🪞", "mirrors"), ("👑", "crown")]

TALK_QS = [("👀", "What do you see?"), ("🚪", "Can you find the door?"),
           ("🚩", "Can you find the flag?"),
           ("🎨", "What colour is the castle?")]
TALK_STARTERS = ["I see a ______.", "The castle is ______.",
                 "I like the ______."]

SOUNDS = [("K", "/k/", "🗝️", "KEY", "kkk - ey", GOLD, L_GOLD),
          ("C", "/c/", "👑", "CROWN", "ccc - rown", CRIMSON, L_CRIMSON),
          ("D", "/d/", "🚪", "DOOR", "ddd - oor", STONE, L_STONE),
          ("B", "/b/", "🔔", "BELL", "bbb - ell", SAPPHIRE, L_SAPPHIRE),
          ("F", "/f/", "🚩", "FLAG", "fff - lag", EMERALD, L_EMERALD)]

FIND_SOUND_A = ("/k/", ["🗝️", "🔔", "🚪"])
FIND_SOUND_B = ("/b/", ["🔔", "🗝️", "🚩"])

UNLOCK_A = [("/d/", ["🚪", "👑", "🔔"]), ("/f/", ["🗝️", "🚩", "🚪"]),
            ("/c/", ["🔔", "🚪", "👑"])]
UNLOCK_B = [("/k/", ["🚩", "🗝️", "👑"]), ("/b/", ["🚪", "🔔", "🗝️"])]

VOWELS = [("A", "/a/", "c - aaa - t", [("🐱", "CAT"), ("🎩", "HAT"),
                                       ("🗺️", "MAP")], CRIMSON, L_CRIMSON),
          ("E", "/e/", "r - eee - d", [("🍎", "RED"), ("🛏️", "BED"),
                                       ("🕸️", "WEB")], EMERALD, L_EMERALD),
          ("I", "/i/", "b - iii - g", [("🐘", "BIG"), ("🪑", "SIT"),
                                       ("🥇", "WIN")], SAPPHIRE, L_SAPPHIRE),
          ("O", "/o/", "f - ooo - x", [("🦊", "FOX"), ("📦", "BOX"),
                                       ("🔥", "HOT")], GOLD, L_GOLD)]

VOWEL_DOORS = [("A", CRIMSON, L_CRIMSON), ("E", EMERALD, L_EMERALD),
               ("I", SAPPHIRE, L_SAPPHIRE), ("O", GOLD, L_GOLD)]
VOWEL_SORT_WORDS = ["CAT", "RED", "BIG", "BOX", "HAT", "HOT"]

WHICH_KEY_A = [("B _ G", "🐘", ["A", "E", "I"]),
               ("H _ T", "🔥", ["A", "E", "O"]),
               ("C _ T", "🐱", ["A", "I", "O"])]
WHICH_KEY_B = [("F _ X", "🦊", ["A", "I", "O"]),
               ("B _ D", "🛏️", ["E", "O", "A"])]

BLENDS = [("C", "A", "T", "🐱", "CAT", "The cat is big.", CRIMSON, L_CRIMSON),
          ("H", "A", "T", "🎩", "HAT", "My hat is red.", GOLD, L_GOLD),
          ("B", "I", "G", "🐘", "BIG", "I see a big door.", SAPPHIRE,
           L_SAPPHIRE),
          ("S", "I", "T", "🪑", "SIT", "I can sit here.", EMERALD,
           L_EMERALD),
          ("F", "O", "X", "🦊", "FOX", "The fox can run.", ROYAL, L_ROYAL),
          ("B", "O", "X", "📦", "BOX", "My box is red.", STONE, L_STONE)]

MIRROR_GAME = [("🦊", ["FOX", "FOS", "FIX"]), ("📦", ["BOS", "BOX", "BAX"]),
               ("🐱", ["CAT", "COT", "CUT"])]

MISSING_LETTER = [("F _ X", "🦊", ["A", "I", "O"]),
                  ("C _ T", "🐱", ["A", "E", "I"]),
                  ("B _ LL", "🔔", ["A", "E", "O"])]

SPELLS = [(["T", "A", "H"], "🎩"), (["G", "I", "B"], "🐘"),
          (["X", "O", "B"], "📦"), (["N", "U", "S"], "☀️")]

FAMILIES = [("-AT", "at", [("🐱", "CAT"), ("🎩", "HAT"), ("🦇", "BAT"),
                           ("🟫", "MAT")], CRIMSON, L_CRIMSON),
            ("-IG", "ig", [("🐘", "BIG"), ("⛏️", "DIG"), ("🐷", "PIG")],
             SAPPHIRE, L_SAPPHIRE),
            ("-OX", "ox", [("🦊", "FOX"), ("📦", "BOX")], GOLD, L_GOLD),
            ("-UN", "un", [("☀️", "SUN"), ("🏃", "RUN"), ("🎉", "FUN")],
             EMERALD, L_EMERALD)]

SHELVES = [("-AT", CRIMSON, L_CRIMSON), ("-UN", EMERALD, L_EMERALD)]
SHELF_WORDS = ["CAT", "RUN", "HAT", "SUN", "BAT", "FUN"]

ODD_ONE = [["FOX", "BOX", "SUN"], ["CAT", "HAT", "BIG"],
           ["SUN", "RUN", "MAP"]]

RHYMES = [("CAT", "🐱", ["HAT", "FOX", "BIG"]),
          ("BOX", "📦", ["FOX", "CAT", "SUN"]),
          ("RUN", "🏃", ["SUN", "BED", "MAP"])]

KING_ACTIONS = [("👏", "King says clap twice."),
                ("🚪", "King says point to the door."),
                ("🙋", "King says touch your head."),
                ("👑", "King says make a royal crown."),
                ("🧊", "King says freeze!")]

KING_CARDS = [("🐱", "CAT"), ("🎩", "HAT"), ("📦", "BOX"), ("🐘", "BIG")]
KING_READS = [("📖", "King says read CAT."), ("🔎", "King says find HAT."),
              ("👉", "King says point to BOX."),
              ("📖", "King says read BIG.")]

KING_SILLY = [("🤴", "King says stand like a king.", "royal pose"),
              ("🔔", "Ring a pretend bell.", "listen before you move"),
              ("👑", "King says put on your crown.", "hands on head"),
              ("🧊", "King says freeze!", "hold still and count to three")]

SIGHT_SLIDES = [(["I"], "I can run.", "You say this word about yourself.",
                 CRIMSON, L_CRIMSON),
                (["see"], "I see a cat.", "This word is about your eyes.",
                 SAPPHIRE, L_SAPPHIRE),
                (["a", "the"], "I see a cat.  ·  I see the cat.",
                 "Two small words. They come before a naming word.", EMERALD,
                 L_EMERALD),
                (["is", "my"], "My hat is red.",
                 "One word joins. One word says it belongs to you.", GOLD,
                 L_GOLD)]

SIGHT_FIND = ["I", "see", "a", "the", "is", "my"]

BUILD_SENTENCES = [("I see a cat.", ["I", "see", "a", "cat"], "🐱",
                    CRIMSON, L_CRIMSON),
                   ("The cat is big.", ["The", "cat", "is", "big"], "🐘",
                    SAPPHIRE, L_SAPPHIRE),
                   ("My hat is red.", ["My", "hat", "is", "red"], "🎩",
                    GOLD, L_GOLD)]

DOOR_WORDS = [("I", CRIMSON, L_CRIMSON), ("see", SAPPHIRE, L_SAPPHIRE),
              ("the", EMERALD, L_EMERALD), ("big", GOLD, L_GOLD),
              ("cat", ROYAL, L_ROYAL)]
DOOR_SENTENCE = "I see the big cat."

PUZZLES = [(["see", "I", "cat", "a"], "🐱"), (["big", "is", "cat", "The"],
                                              "🐘"),
           (["red", "My", "hat", "is"], "🎩"), (["can", "I", "run"], "🏃")]

PIC_MATCH = [("I see a big fox.", ["🦊", "🐱", "🐶"]),
             ("The king has a crown.", ["👑", "🗝️", "📦"])]

TRUE_SILLY = ["The king has a crown.", "The crown has a king.",
              "I see a big castle.", "The castle sits on my hat."]

STORY_PREDICT = [("👑", "a crown"), ("🗝️", "a key"), ("📚", "a book")]

STORY = [
    ("Part 1", "👑", GOLD, L_GOLD,
     ["King Ben has a gold crown.",
      "He wears it every day.",
      "One morning, Ben looks for his crown.",
      "It is not on the table. It is not on his bed."],
     "He wears it every day.", "What does King Ben wear every day?"),
    ("Part 2", "🔎", SAPPHIRE, L_SAPPHIRE,
     ["\"Where is my crown?\" says Ben.",
      "He looks in the big room.",
      "He looks near the red chair.",
      "He looks under the rug. The crown is not there."],
     "He looks in the big room.", "Name one place Ben looks."),
    ("Part 3", "🚪", STONE, L_STONE,
     ["Ben asks his helper, Max.",
      "\"Can you help me?\" says Ben.",
      "Max and Ben look near the door.",
      "They look up at the tall tower. They do not see the crown."],
     "Max and Ben look near the door.", "Who helps King Ben?"),
    ("Part 4", "🔔", CRIMSON, L_CRIMSON,
     ["Then they hear a sound. Ring! Ring!",
      "It comes from the big room.",
      "They look behind a box.",
      "The gold crown is there! A little cat sits next to it."],
     "They look behind a box.", "Where is the crown?"),
    ("Part 5", "🎉", EMERALD, L_EMERALD,
     ["Ben smiles. He puts the crown on his head.",
      "\"Thank you, Max!\" he says.",
      "Everyone laughs at the little cat.",
      "The king is happy to find his crown."],
     "Ben smiles.", "How does the king feel at the end?"),
]

DETECT_A = [("🤴", "Who is in the story?", ["King Ben", "Sam", "Mia"],
             "Part 1"),
            ("👑", "What is missing?", ["A crown", "A book", "A key"],
             "Part 1")]
DETECT_B = [("📦", "Where was the crown found?",
             ["Near a box", "Outside", "In a bag"], "Part 4"),
            ("😊", "What did Ben do when he found it?",
             ["He cried", "He put it on", "He hid it"], "Part 5")]

SEQUENCE = [("A", "👑", "Ben puts the crown on his head."),
            ("B", "🔎", "Ben looks for his crown."),
            ("C", "🔔", "Ben hears a sound."),
            ("D", "📦", "Ben finds the crown behind a box.")]

EVIDENCE = [("How do you know where the crown was?", "Part 4", CRIMSON,
             L_CRIMSON),
            ("How do you know who helped Ben?", "Part 3", STONE, L_STONE),
            ("How do you know the king was happy?", "Part 5", EMERALD,
             L_EMERALD)]

KING_CHALLENGE = [("👑", "READ IT", "CROWN", "one word", CRIMSON, L_CRIMSON),
                  ("🐘", "READ IT", "BIG", "one word", SAPPHIRE, L_SAPPHIRE),
                  ("📦", "BLEND IT", "B – O – X", "three sounds", GOLD,
                   L_GOLD),
                  ("🐱", "READ IT", "The cat is big.", "one sentence",
                   EMERALD, L_EMERALD),
                  ("🔎", "ANSWER IT", "What was missing in the story?",
                   "from the story", ROYAL, L_ROYAL)]

CHAMPION_LINES = [("🔤", "You read words!"), ("🧩", "You blended sounds!"),
                  ("✏️", "You built sentences!"), ("📚", "You read a story!"),
                  ("🔎", "You answered questions!")]

GAMES = [("🗝️", "MISSING LETTER", "Which key finishes the word?",
          ["C _ T", "A  ·  E  ·  I"], GOLD, L_GOLD),
         ("👑", "ROYAL WORD MATCH", "Match each word to its picture.",
          ["KING · CROWN · DOOR", "🤴   👑   🚪"], CRIMSON, L_CRIMSON),
         ("🪞", "MIRROR WORD", "Read it, then break it apart.",
          ["BAG", "B – A – G"], SAPPHIRE, L_SAPPHIRE),
         ("🧩", "SENTENCE FIX", "Put the words in order.",
          ["big / is / king / The", "____________________"], ROYAL, L_ROYAL),
         ("📚", "FIND THE WORD", "Point to one word only.",
          ["I see a big castle.", "Find CASTLE."], EMERALD, L_EMERALD),
         ("🎯", "RHYME CASTLE", "Which one rhymes?",
          ["HAT", "CAT · BOX · SUN"], STONE, L_STONE),
         ("👑", "REAL OR SILLY", "Which one makes sense?",
          ["The king has a crown.", "The crown eats lunch."], CRIMSON,
          L_CRIMSON)]

GAME_WHEN = ["He finishes a mission early.",
             "His focus starts to drop.",
             "One sound needs more practice.",
             "You have five spare minutes.",
             "He asks to play one more."]

SUPPORT_LEVELS = [("🟢", "CASTLE HINT", "Needs a picture and a first sound",
                   "Point at the picture, then the first letter.", EMERALD,
                   L_EMERALD),
                  ("🟡", "CASTLE MISSION", "Reads it with a little help",
                   "Ask one small question, then go quiet and wait.", GOLD,
                   L_GOLD),
                  ("⭐", "ROYAL CHALLENGE", "Reads it on his own",
                   "Ask him to build a new sentence with the word.", ROYAL,
                   L_ROYAL)]

SUPPORT_LADDER = [("HINT 1", "Look at the first letter.", EMERALD),
                  ("HINT 2", "Say each sound.", SAPPHIRE),
                  ("HINT 3", "Blend the sounds together.", ROYAL),
                  ("HINT 4", "Let's read it together.", GOLD),
                  ("HINT 5", "I read it, then you read it.", CRIMSON)]

SENTENCE_STEPS = ["Read the whole sentence out loud yourself.",
                  "Point to each word as you read it again.",
                  "Explain one hard word with a picture.",
                  "Read it together, at his speed.",
                  "Let him try one short sentence alone.",
                  "Praise the effort, not just the answer."]

PRAISE = ["\"Good try!\"", "\"Let's unlock this word.\"",
          "\"You found the first sound!\"", "\"Great blending!\"",
          "\"Try it one more time.\"", "\"You solved that clue!\""]

RUBRIC_SKILLS = ["Beginning sounds", "Short vowels", "CVC words", "Blending",
                 "Word families", "Sight words", "Phrase reading",
                 "Sentence reading", "Story reading", "Sequencing",
                 "Comprehension"]

REVIEW_NEXT = ["WORDS TO REVIEW", "SIGHT WORDS TO REVIEW",
               "READING SKILL FOR NEXT CLASS"]

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
        # Leading spaces mark a wrapped continuation line rather than a new bullet.
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


def footer(slide, n, timing="", stop=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE3, 0xE3, 0xDF))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), GOLD)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40),
             INK)
    msg = "🏰 The Magic Castle Reading Quest  |  Grade 2  |  90 min"
    if stop:
        msg += f"  |  {stop}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg,
       size=10, color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32),
       f"{n} / {TOTAL}", size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, stop, accent=STONE, bg=PAGE):
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


def one_task(slide, text, color=STONE, top=1.34):
    """States the single job of this slide in words the child understands."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4),
              text, size=16, bold=True, color=color)


def hint(slide, text, top=6.42, label="🗝️ CASTLE HINT", fill=L_GOLD,
         color=GOLD):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46),
              fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.6), Inches(0.34),
       label, size=12, bold=True, color=color)
    tb(slide, Inches(3.5), Inches(top + 0.05), Inches(9.1), Inches(0.36),
       text, size=13, bold=True, color=INK)


def i_we_you(slide, top=6.42):
    steps = [("I READ", "Teacher first", CRIMSON),
             ("WE READ", "Together", GOLD), ("YOU READ", "Your turn!",
                                             EMERALD)]
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
    """A word with a gap, its picture, and letter keys to choose from."""
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
               f"🗝️ {opt}", size=24, bold=True, color=accent,
               align=PP_ALIGN.CENTER, font="Arial Black")


def word_cards(slide, left, top, words, color, light, width=1.65, gap=1.78,
               height=0.8, size=22):
    """A row of big readable word tiles."""
    for i, word in enumerate(words):
        l = Inches(left + i * gap)
        add_round(slide, l, Inches(top), Inches(width), Inches(height), light)
        tb(slide, l, Inches(top + height / 2 - 0.24), Inches(width),
           Inches(0.5), word, size=size, bold=True, color=color,
           align=PP_ALIGN.CENTER, font="Arial Black")


def sound_slide(index):
    """One letter, its sound, a picture and the castle word it starts."""
    letter, sound, pic, word, stretch, color, light = SOUNDS[index]
    slide, n = new_slide(f"{letter}  is for  {word}", "SOUNDS", "7–17 min",
                         "Castle Gate", color)
    one_task(slide, f"{letter} makes the {sound} sound. Say it with me.",
             color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(4.3), Inches(4.4),
              light)
    tb(slide, Inches(0.5), Inches(2.1), Inches(4.3), Inches(1.9), pic,
       size=100, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.0), Inches(4.25), Inches(3.3), Inches(1.4),
              WHITE)
    tb(slide, Inches(1.0), Inches(4.5), Inches(3.3), Inches(0.9), word,
       size=38, bold=True, color=color, align=PP_ALIGN.CENTER,
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
    hint(slide, f"Say the word slowly and hold the very first sound: "
                f"{stretch}.", 6.42)


def vowel_slide(index):
    """One short vowel, its stretch sound, and three words that use it."""
    letter, sound, stretch, words, color, light = VOWELS[index]
    slide, n = new_slide(f"🗝️ Short {letter}  —  {sound}", "SHORT VOWELS",
                         "17–27 min", "Key Room", color)
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
    hint(slide, f"Read the three words in a row. The middle sound never "
                f"changes: {sound}.", 6.42)


def blend_slide(index):
    """Three sounds, one arrow, one whole word, then a sentence."""
    a, b, c, pic, word, sentence, color, light = BLENDS[index]
    slide, n = new_slide(f"🪞 {a} – {b} – {c}   →   {word}", "BLENDING",
                         "27–37 min", "Mirror Hall", color)
    one_task(slide, "Touch each letter, say the sound, then push them "
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
    tb(slide, Inches(9.5), Inches(1.93), Inches(3.35), Inches(0.75), pic,
       size=34, align=PP_ALIGN.CENTER)
    tb(slide, Inches(9.5), Inches(2.72), Inches(3.35), Inches(0.95), word,
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
    tb(slide, Inches(0.9), Inches(5.62), Inches(11.6), Inches(0.46),
       sentence, size=26, bold=True, color=INK, font="Georgia")
    hint(slide, f"Stuck? Blend the first two only: "
                f"{a.lower()}{b.lower()} … then add /{c.lower()}/.", 6.42)


def family_slide(index):
    """One word family, its ending, and the words that rhyme inside it."""
    name, ending, words, color, light = FAMILIES[index]
    slide, n = new_slide(f"📚 The {name} Family", "WORD FAMILIES",
                         "37–44 min", "Magic Library", color)
    one_task(slide, f"Every word on this shelf ends the same way: {name}. "
                    f"They all rhyme.", color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(3.2), Inches(4.4),
              color)
    tb(slide, Inches(0.5), Inches(2.4), Inches(3.2), Inches(1.5), name,
       size=70, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Arial Black")
    tb(slide, Inches(0.7), Inches(4.1), Inches(2.8), Inches(0.5),
       f"say it: {ending}", size=20, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.8), Inches(4.8), Inches(2.6), Inches(1.0),
              WHITE)
    tb(slide, Inches(0.8), Inches(4.9), Inches(2.6), Inches(0.32),
       "ON THIS SHELF", size=10, bold=True, color=SOFT,
       align=PP_ALIGN.CENTER)
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
    """One or two sight words: see it big, say it, find it in a row."""
    words, sentence, note, color, light = SIGHT_SLIDES[index]
    heading = "  ·  ".join(w.upper() for w in words)
    slide, n = new_slide(f"🚪 Sight Word — {heading}", "SIGHT WORDS",
                         "49–58 min", "Magic Library", color)
    if len(words) == 1:
        one_task(slide, "This word is a key. We know it by looking, not by "
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
        one_task(slide, "These words are keys. We know them by looking, not "
                        "by sounding them out.", color)
        for i, word in enumerate(words):
            top = Inches(1.85 + i * 2.4)
            add_round(slide, Inches(0.5), top, Inches(5.0), Inches(2.0),
                      color)
            tb(slide, Inches(0.5), top + Inches(0.5), Inches(5.0),
               Inches(1.1), word, size=64, bold=True, color=WHITE,
               align=PP_ALIGN.CENTER, font="Arial Black")
    add_round(slide, Inches(5.8), Inches(1.85), Inches(7.05), Inches(1.15),
              WHITE)
    tb(slide, Inches(6.1), Inches(1.95), Inches(6.5), Inches(0.32),
       "WHAT IT DOES", size=11, bold=True, color=color)
    tb(slide, Inches(6.1), Inches(2.3), Inches(6.5), Inches(0.6), note,
       size=17, bold=True, color=INK)
    add_round(slide, Inches(5.8), Inches(3.15), Inches(7.05), Inches(1.3),
              light)
    tb(slide, Inches(6.1), Inches(3.28), Inches(6.5), Inches(0.32),
       "READ IT IN A SENTENCE", size=11, bold=True, color=color)
    tb(slide, Inches(6.1), Inches(3.65), Inches(6.5), Inches(0.62), sentence,
       size=25, bold=True, color=INK, font="Georgia")
    add_round(slide, Inches(5.8), Inches(4.6), Inches(7.05), Inches(1.65),
              WHITE)
    tb(slide, Inches(6.1), Inches(4.72), Inches(6.5), Inches(0.32),
       "⭐ FIND IT — point to it in this row", size=11, bold=True,
       color=color)
    word_cards(slide, 6.1, 5.15, SIGHT_FIND, color, light, width=1.0,
               gap=1.08, height=0.72, size=17)
    hint(slide, "Say it, clap it, then hunt for the same shape in the row.",
         6.42)


def build_slide(index):
    """Word tiles in order, a picture, then the whole sentence."""
    sentence, words, pic, color, light = BUILD_SENTENCES[index]
    slide, n = new_slide(f"🧩 Build It —  {sentence}", "SENTENCES",
                         "49–58 min", "Magic Library", color)
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
        tb(slide, left, Inches(4.05), Inches(span - 0.2), Inches(0.4), "☐",
           size=16, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.05), Inches(12.35), Inches(1.15),
              color)
    tb(slide, Inches(0.5), Inches(5.32), Inches(12.35), Inches(0.65),
       sentence, size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Georgia")
    i_we_you(slide, 6.42)


def story_slide(part):
    """Picture and part pill on the left, four big story lines on the right."""
    label, pic, color, light, lines, you_line, question = part
    slide, n = new_slide(f"👑 The Lost Crown — {label}", "STORY", "67–77 min",
                         "King's Room", color)
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
           Inches(0.7), line, size=23, bold=True, color=INK, font="Georgia")
    add_round(slide, Inches(4.25), Inches(5.05), Inches(8.6), Inches(1.15),
              light)
    tb(slide, Inches(4.6), Inches(5.16), Inches(8.0), Inches(0.32),
       "⭐ YOU READ THIS LINE ON YOUR OWN", size=11, bold=True, color=color)
    tb(slide, Inches(4.6), Inches(5.5), Inches(8.0), Inches(0.5), you_line,
       size=24, bold=True, color=INK, font="Georgia")
    i_we_you(slide, 6.42)


def question_rows(slide, items, start_index, accent, top_start=1.95,
                  gap=2.16, height=1.92):
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
    """Closes a castle room: what was learned, and what comes next."""
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
       "One more room unlocked!", size=16, color=SOFT,
       align=PP_ALIGN.CENTER, italic=True)
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.85 + i * 0.95)
        add_round(slide, Inches(5.45), top, Inches(7.4), Inches(0.8), WHITE)
        add_oval(slide, Inches(5.75), top + Inches(0.14), Inches(0.52),
                 Inches(0.52), light)
        tb(slide, Inches(5.75), top + Inches(0.2), Inches(0.52), Inches(0.4),
           icon, size=14, align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.55), top + Inches(0.18), Inches(6.1), Inches(0.46),
           line, size=17, bold=True, color=INK)
    add_round(slide, Inches(5.45), Inches(5.7), Inches(7.4), Inches(0.52),
              color)
    tb(slide, Inches(5.75), Inches(5.8), Inches(6.8), Inches(0.36),
       f"NEXT:  {next_line}", size=13, bold=True, color=WHITE)


# ------------------------------------------------------------------- slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    tb(slide, Inches(0.7), Inches(0.75), Inches(12), Inches(1.2), "🏰",
       size=64, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.95), Inches(12), Inches(1.0),
       "THE MAGIC CASTLE READING QUEST", size=42, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.95), Inches(12), Inches(0.5),
       "\"Can You Become a Castle Reading Champion?\"", size=20,
       color=RGBColor(0xE6, 0xC5, 0x7A), align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, name) in enumerate(CASTLE_SCENE):
        left = Inches(1.35 + i * 1.78)
        add_round(slide, left, Inches(3.65), Inches(1.6), Inches(1.5),
                  RGBColor(0x24, 0x2B, 0x38))
        tb(slide, left, Inches(3.85), Inches(1.6), Inches(0.7), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(4.6), Inches(1.6), Inches(0.4), name, size=11,
           color=RGBColor(0xE6, 0xC5, 0x7A), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.1), Inches(5.4), Inches(7.1), Inches(1.0), GOLD)
    tb(slide, Inches(3.1), Inches(5.62), Inches(7.1), Inches(0.6),
       "Grade 2  •  90 Minutes  •  Reading", size=24, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Castle Gate")
    fade(slide)


def s02_junior_reader():
    slide, n = new_slide("⭐ Today You Are a Junior Castle Reader!", "WELCOME",
                         "0–7 min", "Castle Gate", ROYAL)
    one_task(slide, "Here is everything you will do inside the castle today.",
             ROYAL)
    for i, (icon, name, detail) in enumerate(MISSION):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 2.2)
        add_round(slide, left, top, Inches(5.95), Inches(2.0), L_ROYAL)
        add_oval(slide, left + Inches(0.35), top + Inches(0.3), Inches(0.9),
                 Inches(0.9), WHITE)
        tb(slide, left + Inches(0.35), top + Inches(0.44), Inches(0.9),
           Inches(0.6), icon, size=24, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.5), top + Inches(0.4), Inches(4.2),
           Inches(0.5), name, size=22, bold=True, color=INK)
        tb(slide, left + Inches(1.5), top + Inches(0.95), Inches(4.2),
           Inches(0.5), detail, size=15, color=DARK)
        add_round(slide, left + Inches(1.5), top + Inches(1.45), Inches(2.6),
                  Inches(0.4), WHITE)
        tb(slide, left + Inches(1.5), top + Inches(1.5), Inches(2.6),
           Inches(0.32), f"MISSION {i + 1}", size=11, bold=True, color=ROYAL,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.35), Inches(12.35), Inches(0.5),
              GOLD)
    tb(slide, Inches(0.5), Inches(6.45), Inches(12.35), Inches(0.36),
       "Finish every room and you become a 🏆 CASTLE READING CHAMPION",
       size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def s03_map():
    slide, n = new_slide("🗺️ Your Castle Mission Map", "WELCOME", "0–7 min",
                         "Castle Gate", STONE)
    one_task(slide, "Six rooms. One reading mission in each one.", STONE)
    for i, (emoji, name, focus, timing, color, light) in enumerate(STOPS):
        left = Inches(0.5 + i * 2.09)
        add_round(slide, left, Inches(1.9), Inches(1.9), Inches(3.9), light)
        add_round(slide, left, Inches(1.9), Inches(1.9), Inches(0.44), color)
        tb(slide, left, Inches(1.96), Inches(1.9), Inches(0.34),
           f"ROOM {i + 1}", size=10, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(1.9), Inches(0.9), emoji,
           size=38, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.5), Inches(1.7), Inches(0.62),
           name, size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.2), Inches(1.7), Inches(0.62),
           focus, size=12, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(5.05), Inches(1.5),
                  Inches(0.4), WHITE)
        tb(slide, left + Inches(0.2), Inches(5.11), Inches(1.5), Inches(0.32),
           timing, size=10, bold=True, color=color, align=PP_ALIGN.CENTER)
        if i < len(STOPS) - 1:
            tb(slide, left + Inches(1.82), Inches(3.35), Inches(0.35),
               Inches(0.5), "→", size=20, bold=True, color=SOFT,
               align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.05), Inches(12.35), Inches(0.85),
              L_GOLD)
    tb(slide, Inches(0.5), Inches(6.18), Inches(12.35), Inches(0.36),
       "🏆 AT THE END OF THE QUEST", size=11, bold=True, color=GOLD,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(6.5), Inches(12.35), Inches(0.36),
       "You read words, sentences and a whole story — and you become a "
       "Castle Reading Champion.", size=15, bold=True, color=INK,
       align=PP_ALIGN.CENTER)


def s04_castle_talk():
    slide, n = new_slide("🏰 Castle Talk — Just Look and Tell Me", "TALKING",
                         "0–7 min", "Castle Gate", EMERALD)
    one_task(slide, "No reading yet. Look at the castle and talk to me.",
             EMERALD)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.2), Inches(4.4),
              L_EMERALD)
    tb(slide, Inches(0.5), Inches(2.2), Inches(5.2), Inches(1.7), "🏰",
       size=110, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(4.5), Inches(4.6), Inches(0.6),
       "Point to what you can see.", size=17, bold=True, color=EMERALD,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(5.2), Inches(4.6), Inches(0.6),
       "There is no wrong answer here.", size=14, color=SOFT,
       align=PP_ALIGN.CENTER, italic=True)
    for i, (icon, question) in enumerate(TALK_QS):
        col, row = i % 2, i // 2
        left = Inches(6.05 + col * 3.45)
        top = Inches(1.9 + row * 1.15)
        add_round(slide, left, top, Inches(3.25), Inches(0.95), WHITE)
        tb(slide, left + Inches(0.2), top + Inches(0.22), Inches(0.5),
           Inches(0.5), icon, size=16, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.8), top + Inches(0.22), Inches(2.3),
           Inches(0.5), question, size=15, bold=True, color=INK)
    tb(slide, Inches(6.05), Inches(4.25), Inches(6.8), Inches(0.4),
       "SAY IT LIKE THIS", size=12, bold=True, color=EMERALD)
    for i, starter in enumerate(TALK_STARTERS):
        top = Inches(4.68 + i * 0.6)
        add_round(slide, Inches(6.05), top, Inches(6.8), Inches(0.5),
                  L_EMERALD)
        tb(slide, Inches(6.35), top + Inches(0.05), Inches(6.2), Inches(0.4),
           starter, size=20, bold=True, color=INK, font="Georgia")
    hint(slide, "Say your own answer first. He will copy the pattern.", 6.5)


def s05_sound_k():
    sound_slide(0)


def s06_sound_c():
    sound_slide(1)


def s07_sound_d():
    sound_slide(2)


def s08_sound_b():
    sound_slide(3)


def s09_sound_f():
    sound_slide(4)


def find_sound_slide(pair, title, accent, hint_text):
    sound, pics = pair
    slide, n = new_slide(title, "LISTENING", "7–17 min", "Castle Gate",
                         accent)
    one_task(slide, "Name each picture out loud, then listen to the first "
                    "sound.", accent)
    add_round(slide, Inches(4.9), Inches(1.8), Inches(3.55), Inches(1.15),
              accent)
    tb(slide, Inches(4.9), Inches(1.9), Inches(3.55), Inches(0.32),
       "FIND THE SOUND", size=11, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(4.9), Inches(2.22), Inches(3.55), Inches(0.62), sound,
       size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Arial Black")
    for i, pic in enumerate(pics):
        left = Inches(0.9 + i * 4.0)
        add_round(slide, left, Inches(3.2), Inches(3.6), Inches(2.9), L_GREY)
        tb(slide, left, Inches(3.5), Inches(3.6), Inches(1.7), pic, size=88,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.3), Inches(5.35), Inches(1.0),
                  Inches(0.55), WHITE)
        tb(slide, left + Inches(1.3), Inches(5.45), Inches(1.0), Inches(0.4),
           "☐", size=18, color=SOFT, align=PP_ALIGN.CENTER)
    hint(slide, hint_text, 6.5)


def s10_find_k():
    find_sound_slide(FIND_SOUND_A, "👂 Which One Starts With /k/?", GOLD,
                     "Say all three names slowly, holding the first sound.")


def s11_find_b():
    find_sound_slide(FIND_SOUND_B, "👂 Which One Starts With /b/?", SAPPHIRE,
                     "Press your lips together for /b/. He will hear it.")


def s12_unlock_a():
    slide, n = new_slide("🎮 Game — Unlock the Door", "GAME", "7–17 min",
                         "Castle Gate", STONE)
    one_task(slide, "Every right sound unlocks one castle door.", STONE)
    picture_rows(slide, UNLOCK_A, 1, STONE)
    hint(slide, "Name the pictures first, then say the sound again.", 6.5)


def s13_unlock_b():
    slide, n = new_slide("🎮 Game — Two Doors Left", "GAME", "7–17 min",
                         "Castle Gate", CRIMSON)
    one_task(slide, "Two more doors and the gate is open.", CRIMSON)
    picture_rows(slide, UNLOCK_B, 4, CRIMSON, top_start=2.1)
    add_round(slide, Inches(0.5), Inches(5.3), Inches(12.35), Inches(0.85),
              L_GOLD)
    tb(slide, Inches(0.85), Inches(5.42), Inches(11.6), Inches(0.32),
       "⭐ ROYAL CHALLENGE", size=11, bold=True, color=GOLD)
    tb(slide, Inches(0.85), Inches(5.74), Inches(11.6), Inches(0.42),
       "Say all five castle sounds in a row:  /k/  /c/  /d/  /b/  /f/",
       size=18, bold=True, color=INK)
    hint(slide, "If he picks the wrong picture, say both names slowly "
                "together.", 6.5)


def s14_gate_done():
    badge_slide("🚪 The Castle Gate Is Open!", "7–17 min", "Castle Gate",
                "🚪", "KEY 1 EARNED",
                "You know five castle sounds now.",
                [("🔤", "K, C, D, B and F."),
                 ("👂", "You heard the first sound in a word."),
                 ("🎮", "Five doors unlocked, all correct."),
                 ("🗣️", "You said every sound out loud.")],
                "🗝️ The Key Room — the sounds in the middle.", STONE,
                L_STONE)


def s15_vowel_a():
    vowel_slide(0)


def s16_vowel_e():
    vowel_slide(1)


def s17_vowel_i():
    vowel_slide(2)


def s18_vowel_o():
    vowel_slide(3)


def s19_vowel_sort():
    slide, n = new_slide("🗝️ Vowel Key Sort — Four Castle Doors", "SORTING",
                         "17–27 min", "Key Room", GOLD)
    one_task(slide, "Read each word, then send it through the right door.",
             GOLD)
    for i, (letter, color, light) in enumerate(VOWEL_DOORS):
        left = Inches(0.5 + i * 3.11)
        add_round(slide, left, Inches(1.85), Inches(2.9), Inches(2.85),
                  light)
        add_round(slide, left, Inches(1.85), Inches(2.9), Inches(0.75),
                  color)
        tb(slide, left, Inches(1.94), Inches(2.9), Inches(0.6),
           f"🚪 {letter}", size=28, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left, Inches(2.75), Inches(2.9), Inches(0.4),
           f"/{letter.lower()}/", size=17, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(3.25), Inches(2.3), Inches(1.3),
           "____________\n\n____________", size=15, color=SOFT,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.85), Inches(12.35), Inches(0.4),
       "WORDS TO SORT — read each one out loud first", size=12, bold=True,
       color=GOLD)
    word_cards(slide, 0.5, 5.3, VOWEL_SORT_WORDS, INK, WHITE, width=1.9,
               gap=2.07, height=0.85, size=26)
    hint(slide, "Say the word slowly and listen for the sound in the middle.",
         6.42)


def s20_which_key_a():
    slide, n = new_slide("🎮 Which Key Opens the Word?", "GAME", "17–27 min",
                         "Key Room", SAPPHIRE)
    one_task(slide, "The picture tells you the word. Pick the middle letter.",
             SAPPHIRE)
    choice_rows(slide, WHICH_KEY_A, 1, SAPPHIRE)
    hint(slide, "Say the word, then say only the middle sound on its own.",
         6.42)


def s21_which_key_b():
    slide, n = new_slide("🎮 Two More Locks", "GAME", "17–27 min", "Key Room",
                         CRIMSON)
    one_task(slide, "Last two keys. Look at the picture, then choose.",
             CRIMSON)
    choice_rows(slide, WHICH_KEY_B, 4, CRIMSON, top_start=2.1, gap=1.6,
                height=1.4)
    add_round(slide, Inches(0.5), Inches(5.3), Inches(12.35), Inches(0.85),
              L_ROYAL)
    tb(slide, Inches(0.85), Inches(5.42), Inches(11.6), Inches(0.32),
       "⭐ ROYAL CHALLENGE", size=11, bold=True, color=ROYAL)
    tb(slide, Inches(0.85), Inches(5.74), Inches(11.6), Inches(0.42),
       "Read all four vowel sounds in a row:  /a/  /e/  /i/  /o/", size=18,
       bold=True, color=INK)
    hint(slide, "If he picks the wrong key, read his word back to him.",
         6.42)


def s22_key_done():
    badge_slide("🗝️ The Key Room Is Yours!", "17–27 min", "Key Room", "🗝️",
                "KEY 2 EARNED",
                "You know the four short vowel sounds.",
                [("🔤", "A, E, I and O in the middle."),
                 ("🚪", "Six words sent through the right door."),
                 ("🎮", "Five missing vowels found."),
                 ("👂", "You heard the middle sound on its own.")],
                "🪞 The Mirror Hall — pushing sounds together.", GOLD,
                L_GOLD)


def s23_blend_cat():
    blend_slide(0)


def s24_blend_hat():
    blend_slide(1)


def s25_blend_big():
    blend_slide(2)


def s26_blend_sit():
    blend_slide(3)


def s27_blend_fox():
    blend_slide(4)


def s28_blend_box():
    blend_slide(5)


def s29_mirror_game():
    slide, n = new_slide("🪞 Magic Mirror Game — Which Word Matches?", "GAME",
                         "27–37 min", "Mirror Hall", SAPPHIRE)
    one_task(slide, "The mirror copied the word twice and changed it. Only "
                    "one matches the picture.", SAPPHIRE)
    for i, (pic, options) in enumerate(MIRROR_GAME):
        top = Inches(1.95 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3),
                  L_SAPPHIRE)
        add_oval(slide, Inches(0.8), top + Inches(0.4), Inches(0.5),
                 Inches(0.5), SAPPHIRE)
        tb(slide, Inches(0.8), top + Inches(0.46), Inches(0.5), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.3), Inches(1.1), Inches(0.72),
           pic, size=36, align=PP_ALIGN.CENTER)
        for j, opt in enumerate(options):
            left = Inches(3.2 + j * 3.15)
            add_round(slide, left, top + Inches(0.28), Inches(2.95),
                      Inches(0.76), WHITE)
            tb(slide, left, top + Inches(0.4), Inches(2.95), Inches(0.54),
               opt, size=28, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")
    hint(slide, "Read all three out loud. Only one will sound like the "
                "picture's name.", 6.42)


def s30_missing_letter():
    slide, n = new_slide("🔤 Missing Magic Letter", "GAME", "27–37 min",
                         "Mirror Hall", ROYAL)
    one_task(slide, "One letter vanished. Which key brings it back?", ROYAL)
    choice_rows(slide, MISSING_LETTER, 1, ROYAL)
    hint(slide, "Try each letter in the gap and say the word. One sounds "
                "right.", 6.42)


def s31_spells():
    slide, n = new_slide("🧩 Mixed-Up Spell — Put the Letters Back", "GAME",
                         "27–37 min", "Mirror Hall", EMERALD)
    one_task(slide, "The spell scrambled the letters. The picture is your "
                    "clue.", EMERALD)
    for i, (letters, pic) in enumerate(SPELLS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(5.95), Inches(2.1), L_GREY)
        add_oval(slide, left + Inches(0.25), top + Inches(0.18), Inches(0.5),
                 Inches(0.5), EMERALD)
        tb(slide, left + Inches(0.25), top + Inches(0.24), Inches(0.5),
           Inches(0.4), str(i + 1), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(4.85), top + Inches(0.3), Inches(0.9),
           Inches(0.9), pic, size=38, align=PP_ALIGN.CENTER)
        for j, ch in enumerate(letters):
            l = Inches(left.inches + 0.9 + j * 1.15)
            add_round(slide, l, top + Inches(0.3), Inches(1.0), Inches(0.9),
                      WHITE)
            tb(slide, l, top + Inches(0.44), Inches(1.0), Inches(0.62), ch,
               size=32, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")
        add_round(slide, left + Inches(0.9), top + Inches(1.35),
                  Inches(4.75), Inches(0.6), WHITE)
        tb(slide, left + Inches(1.15), top + Inches(1.46), Inches(4.2),
           Inches(0.4), "write it here:  ______________", size=15,
           color=SOFT)
    hint(slide, "Ask which letter comes first. The first sound unlocks it.",
         6.6)


def s32_mirror_done():
    badge_slide("🪞 The Mirror Hall Is Clear!", "27–37 min", "Mirror Hall",
                "🪞", "KEY 3 EARNED",
                "You can push three sounds into one word.",
                [("🧩", "Six words blended: cat, hat, big, sit, fox, box."),
                 ("🪞", "Three real words spotted among the fakes."),
                 ("🔤", "Three missing letters found."),
                 ("✨", "Four scrambled spells fixed.")],
                "📚 The Magic Library — words that rhyme.", SAPPHIRE,
                L_SAPPHIRE)


def s33_family_at():
    family_slide(0)


def s34_family_ig():
    family_slide(1)


def s35_family_ox():
    family_slide(2)


def s36_family_un():
    family_slide(3)


def s37_book_sort():
    slide, n = new_slide("📚 Library Book Sort", "SORTING", "37–44 min",
                         "Magic Library", EMERALD)
    one_task(slide, "Read each book, then put it on the right shelf.",
             EMERALD)
    for i, (name, color, light) in enumerate(SHELVES):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(1.85), Inches(5.95), Inches(2.6),
                  light)
        add_round(slide, left, Inches(1.85), Inches(5.95), Inches(0.66),
                  color)
        tb(slide, left, Inches(1.95), Inches(5.95), Inches(0.46),
           f"{name}  shelf", size=20, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.65), Inches(5.95), Inches(0.6), "📚",
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.6), Inches(3.35), Inches(4.75),
           Inches(0.95), "____________      ____________\n"
                          "____________      ____________", size=16,
           color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.6), Inches(12.35), Inches(0.4),
       "BOOKS TO SHELVE", size=12, bold=True, color=EMERALD)
    for i, word in enumerate(SHELF_WORDS):
        left = Inches(0.5 + i * 2.07)
        add_round(slide, left, Inches(5.05), Inches(1.9), Inches(1.1), WHITE)
        tb(slide, left, Inches(5.15), Inches(1.9), Inches(0.4), "📕",
           size=14, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(5.55), Inches(1.9), Inches(0.5), word,
           size=24, bold=True, color=EMERALD, align=PP_ALIGN.CENTER,
           font="Arial Black")
    hint(slide, "Say two words together. If they rhyme, they share a shelf.",
         6.42)


def s38_odd_one():
    slide, n = new_slide("🎮 Who Does Not Belong?", "GAME", "37–44 min",
                         "Magic Library", ROYAL)
    one_task(slide, "Read all three out loud, then cross out the odd one.",
             ROYAL)
    for i, words in enumerate(ODD_ONE):
        top = Inches(1.95 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3),
                  L_GREY)
        add_oval(slide, Inches(0.8), top + Inches(0.4), Inches(0.5),
                 Inches(0.5), ROYAL)
        tb(slide, Inches(0.8), top + Inches(0.46), Inches(0.5), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        for j, word in enumerate(words):
            left = Inches(1.7 + j * 3.0)
            add_round(slide, left, top + Inches(0.26), Inches(2.7),
                      Inches(0.78), WHITE)
            tb(slide, left, top + Inches(0.38), Inches(2.7), Inches(0.56),
               word, size=30, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")
        add_round(slide, Inches(10.9), top + Inches(0.26), Inches(1.7),
                  Inches(0.78), WHITE)
        tb(slide, Inches(10.9), top + Inches(0.42), Inches(1.7), Inches(0.44),
           "✗ odd one", size=12, bold=True, color=SOFT,
           align=PP_ALIGN.CENTER)
    hint(slide, "Two of them rhyme. Say them as a pair and the odd one pops "
                "out.", 6.42)


def s39_rhyme():
    slide, n = new_slide("⭐ Rhyme Challenge", "GAME", "37–44 min",
                         "Magic Library", GOLD)
    one_task(slide, "Which word rhymes? Read all three before you choose.",
             GOLD)
    for i, (word, pic, options) in enumerate(RHYMES):
        top = Inches(1.95 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3), WHITE)
        tb(slide, Inches(0.8), top + Inches(0.3), Inches(0.9), Inches(0.7),
           pic, size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(1.9), top + Inches(0.28), Inches(2.4),
                  Inches(0.74), GOLD)
        tb(slide, Inches(1.9), top + Inches(0.4), Inches(2.4), Inches(0.52),
           word, size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, Inches(4.5), top + Inches(0.42), Inches(1.4), Inches(0.5),
           "rhymes with", size=13, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(6.05 + j * 2.25)
            add_round(slide, left, top + Inches(0.28), Inches(2.1),
                      Inches(0.74), L_GOLD)
            tb(slide, left, top + Inches(0.4), Inches(2.1), Inches(0.52),
               opt, size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "A rhyme sounds the same at the end. Say both words twice.",
         6.42)


def s40_library_done():
    badge_slide("📚 The Magic Library Is Sorted!", "37–44 min",
                "Magic Library", "📚", "KEY 4 EARNED",
                "You can hear which words rhyme.",
                [("📖", "Four families: -at, -ig, -ox, -un."),
                 ("📚", "Six books on the right shelf."),
                 ("🎮", "Three odd words found."),
                 ("⭐", "Three rhymes matched.")],
                "A short break, then the sight-word doors.", EMERALD,
                L_EMERALD)


def s41_king_says():
    slide, n = new_slide("👑 Brain Break — King Says", "BREAK", "44–49 min",
                         "Magic Library", CRIMSON)
    one_task(slide, "Stay beside your chair. Only move if I say \"King "
                    "says\".", CRIMSON)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(3.6), Inches(4.4),
              L_CRIMSON)
    tb(slide, Inches(0.5), Inches(2.5), Inches(3.6), Inches(1.6), "🤴",
       size=88, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.35), Inches(3.6), Inches(0.7),
       "5 MINUTES", size=28, bold=True, color=CRIMSON,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(5.15), Inches(3.2), Inches(0.7),
       "then sight words", size=14, color=SOFT, align=PP_ALIGN.CENTER)
    for i, (icon, call) in enumerate(KING_ACTIONS):
        top = Inches(1.85 + i * 0.9)
        add_round(slide, Inches(4.35), top, Inches(8.5), Inches(0.78), WHITE)
        add_oval(slide, Inches(4.65), top + Inches(0.14), Inches(0.5),
                 Inches(0.5), L_CRIMSON)
        tb(slide, Inches(4.65), top + Inches(0.2), Inches(0.5), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(5.45), top + Inches(0.16), Inches(7.1), Inches(0.5),
           call, size=19, bold=True, color=INK)
    hint(slide, "Mix the order so he has to listen instead of guessing.",
         6.45)


def s42_king_reads():
    slide, n = new_slide("👑 King Says — Now Read One Word", "BREAK",
                         "44–49 min", "Magic Library", SAPPHIRE)
    one_task(slide, "After each move, the king gives one small reading job.",
             SAPPHIRE)
    for i, (pic, word) in enumerate(KING_CARDS):
        left = Inches(0.5 + i * 3.21)
        add_round(slide, left, Inches(1.85), Inches(3.0), Inches(2.4), WHITE)
        tb(slide, left, Inches(2.05), Inches(3.0), Inches(1.0), pic, size=44,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.35), Inches(3.15), Inches(2.3),
                  Inches(0.85), L_SAPPHIRE)
        tb(slide, left + Inches(0.35), Inches(3.32), Inches(2.3),
           Inches(0.56), word, size=30, bold=True, color=SAPPHIRE,
           align=PP_ALIGN.CENTER, font="Arial Black")
    for i, (icon, call) in enumerate(KING_READS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(4.5 + row * 0.9)
        add_round(slide, left, top, Inches(5.95), Inches(0.78), L_GREY)
        tb(slide, left + Inches(0.3), top + Inches(0.16), Inches(0.5),
           Inches(0.5), icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.0), top + Inches(0.16), Inches(4.6),
           Inches(0.5), call, size=19, bold=True, color=INK)
    hint(slide, "Two jobs are pointing and two are reading. Keep it quick.",
         6.45)


def s43_king_freeze():
    slide, n = new_slide("👑 King Says — Royal Freeze", "BREAK", "44–49 min",
                         "Magic Library", ROYAL)
    one_task(slide, "One of these is a trick. Listen for \"King says\".",
             ROYAL)
    for i, (icon, call, note) in enumerate(KING_SILLY):
        top = Inches(1.95 + i * 1.12)
        color, light = ROYAL, L_ROYAL
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98),
                  light)
        add_oval(slide, Inches(0.8), top + Inches(0.22), Inches(0.54),
                 Inches(0.54), WHITE)
        tb(slide, Inches(0.8), top + Inches(0.28), Inches(0.54),
           Inches(0.44), icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.65), top + Inches(0.22), Inches(6.6),
           Inches(0.52), call, size=20, bold=True, color=INK)
        add_round(slide, Inches(8.5), top + Inches(0.24), Inches(4.05),
                  Inches(0.5), WHITE)
        tb(slide, Inches(8.5), top + Inches(0.32), Inches(4.05),
           Inches(0.36), note, size=13, bold=True, color=color,
           align=PP_ALIGN.CENTER)
    hint(slide, "One of these four has no \"King says\". Read them out of "
                "order and do not point it out.", 6.5)


def s44_sight_i():
    sight_slide(0)


def s45_sight_see():
    sight_slide(1)


def s46_sight_a_the():
    sight_slide(2)


def s47_sight_is_my():
    sight_slide(3)


def s48_build_cat():
    build_slide(0)


def s49_build_big():
    build_slide(1)


def s50_build_red():
    build_slide(2)


def s51_magic_door():
    slide, n = new_slide("🚪 Magic Door Game — Read to Open", "GAME",
                         "49–58 min", "Magic Library", ROYAL)
    one_task(slide, "Read the word on each door. Every word opens one door.",
             ROYAL)
    for i, (word, color, light) in enumerate(DOOR_WORDS):
        left = Inches(0.5 + i * 2.52)
        add_round(slide, left, Inches(1.85), Inches(2.3), Inches(3.1), light)
        tb(slide, left, Inches(2.05), Inches(2.3), Inches(0.8), "🚪",
           size=34, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(2.95), Inches(1.9),
                  Inches(0.9), color)
        tb(slide, left + Inches(0.2), Inches(3.12), Inches(1.9),
           Inches(0.58), word, size=26, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.6), Inches(4.05), Inches(1.1),
                  Inches(0.6), WHITE)
        tb(slide, left + Inches(0.6), Inches(4.17), Inches(1.1),
           Inches(0.42), "☐", size=18, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.15), Inches(12.35), Inches(1.05),
              ROYAL)
    tb(slide, Inches(0.5), Inches(5.25), Inches(12.35), Inches(0.3),
       "ALL FIVE DOORS OPEN — NOW READ THE WHOLE SENTENCE", size=11,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(5.58), Inches(12.35), Inches(0.55),
       DOOR_SENTENCE, size=34, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    hint(slide, "If a door sticks, read it together, then let him read it "
                "alone.", 6.42)


def puzzle_row(slide, top, index, words, pic, color, light):
    """Mixed-up words, a picture clue, and a line to say the sentence on."""
    add_round(slide, Inches(0.5), top, Inches(12.35), Inches(2.1), light)
    add_oval(slide, Inches(0.8), top + Inches(0.25), Inches(0.52),
             Inches(0.52), color)
    tb(slide, Inches(0.8), top + Inches(0.31), Inches(0.52), Inches(0.4),
       str(index), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(1.5), top + Inches(0.26), Inches(2.4), Inches(0.5),
       "MIXED-UP WORDS", size=11, bold=True, color=color)
    for j, word in enumerate(words):
        left = Inches(1.5 + j * 2.45)
        add_round(slide, left, top + Inches(0.66), Inches(2.3),
                  Inches(0.78), WHITE)
        tb(slide, left, top + Inches(0.8), Inches(2.3), Inches(0.52), word,
           size=26, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Arial Black")
    tb(slide, Inches(11.3), top + Inches(0.55), Inches(1.3), Inches(1.0),
       pic, size=42, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.5), top + Inches(1.52), Inches(9.4),
              Inches(0.46), WHITE)
    tb(slide, Inches(1.75), top + Inches(1.58), Inches(8.9), Inches(0.36),
       "say it in order:  ______________________________________", size=14,
       color=SOFT)


def s52_puzzles_a():
    slide, n = new_slide("🧩 Castle Sentence Puzzles", "PUZZLE", "58–67 min",
                         "Puzzle Room", ROYAL)
    one_task(slide, "The words are jumbled. Put them in an order that makes "
                    "sense.", ROYAL)
    for i in range(2):
        words, pic = PUZZLES[i]
        puzzle_row(slide, Inches(1.9 + i * 2.3), i + 1, words, pic, ROYAL,
                   L_ROYAL)
    hint(slide, "Which word would you say first? A sentence starts with a "
                "capital letter.", 6.6)


def s53_puzzles_b():
    slide, n = new_slide("🧩 Two More Puzzles", "PUZZLE", "58–67 min",
                         "Puzzle Room", SAPPHIRE)
    one_task(slide, "Last two. Read your sentence back to me when you "
                    "finish.", SAPPHIRE)
    for i in range(2):
        words, pic = PUZZLES[i + 2]
        puzzle_row(slide, Inches(1.9 + i * 2.3), i + 3, words, pic, SAPPHIRE,
                   L_SAPPHIRE)
    hint(slide, "Move one card at a time. Read it out loud after every move.",
         6.6)


def s54_picture_match():
    slide, n = new_slide("🖼️ Sentence and Picture Match", "GAME",
                         "58–67 min", "Puzzle Room", EMERALD)
    one_task(slide, "Read the sentence, then point to the picture it tells "
                    "about.", EMERALD)
    for i, (sentence, pics) in enumerate(PIC_MATCH):
        top = Inches(1.9 + i * 2.35)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(2.15),
                  L_GREY)
        add_round(slide, Inches(0.8), top + Inches(0.25), Inches(5.0),
                  Inches(0.9), WHITE)
        tb(slide, Inches(0.8), top + Inches(0.45), Inches(5.0), Inches(0.52),
           sentence, size=25, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Georgia")
        add_round(slide, Inches(0.8), top + Inches(1.3), Inches(5.0),
                  Inches(0.5), EMERALD)
        tb(slide, Inches(0.8), top + Inches(1.38), Inches(5.0), Inches(0.36),
           "I READ  →  WE READ  →  YOU READ", size=12, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        for j, pic in enumerate(pics):
            left = Inches(6.3 + j * 2.2)
            add_round(slide, left, top + Inches(0.3), Inches(2.0),
                      Inches(1.55), WHITE)
            tb(slide, left, top + Inches(0.5), Inches(2.0), Inches(0.9), pic,
               size=46, align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(1.42), Inches(2.0), Inches(0.36),
               "☐", size=15, color=SOFT, align=PP_ALIGN.CENTER)
    hint(slide, "Read the sentence twice. The last word usually decides it.",
         6.6)


def s55_true_silly():
    slide, n = new_slide("✅ True or Silly?", "GAME", "58–67 min",
                         "Puzzle Room", GOLD)
    one_task(slide, "Read it, then tell me: does that make sense?", GOLD)
    for i, sentence in enumerate(TRUE_SILLY):
        top = Inches(1.95 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.96),
                  WHITE)
        add_oval(slide, Inches(0.8), top + Inches(0.22), Inches(0.5),
                 Inches(0.5), GOLD)
        tb(slide, Inches(0.8), top + Inches(0.28), Inches(0.5), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.6), top + Inches(0.2), Inches(6.4), Inches(0.56),
           sentence, size=25, bold=True, color=INK, font="Georgia")
        for j, (label, color, light) in enumerate(
                [("✅ MAKES SENSE", EMERALD, L_EMERALD),
                 ("❌ SILLY", CRIMSON, L_CRIMSON)]):
            left = Inches(8.4 + j * 2.15)
            add_round(slide, left, top + Inches(0.23), Inches(1.95),
                      Inches(0.5), light)
            tb(slide, left, top + Inches(0.31), Inches(1.95), Inches(0.36),
               label, size=12, bold=True, color=color,
               align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.42), Inches(12.35), Inches(0.46),
              L_GOLD)
    tb(slide, Inches(0.8), Inches(6.48), Inches(11.7), Inches(0.36),
       "🗝️ CASTLE HINT — For a silly one ask: \"Can a crown really eat "
       "lunch?\" Then laugh about it together.", size=13, bold=True,
       color=INK)


def s56_puzzle_done():
    badge_slide("🧩 The Puzzle Room Is Solved!", "58–67 min", "Puzzle Room",
                "🧩", "KEY 5 EARNED",
                "You can read and build whole sentences.",
                [("👀", "Six sight words you know by looking."),
                 ("🧩", "Four jumbled sentences put back in order."),
                 ("🖼️", "Two sentences matched to the right picture."),
                 ("✅", "You spotted the silly sentences.")],
                "👑 The King's Room — a real story.", ROYAL, L_ROYAL)


def s57_story_intro():
    slide, n = new_slide("📖 Story Time — The Lost Crown", "STORY",
                         "67–77 min", "King's Room", CRIMSON)
    one_task(slide, "Before we read: what do you think is missing?", CRIMSON)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.6), Inches(4.4),
              L_CRIMSON)
    tb(slide, Inches(0.5), Inches(2.2), Inches(5.6), Inches(1.6), "👑",
       size=105, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.15), Inches(5.2), Inches(0.9),
       "THE LOST CROWN", size=38, bold=True, color=CRIMSON,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.2), Inches(5.0), Inches(0.8),
       "five short parts  ·  I read first, then we read together", size=14,
       color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(6.4), Inches(1.9), Inches(6.45), Inches(0.5),
       "🤔 WHAT DO YOU THINK IS MISSING?", size=15, bold=True, color=CRIMSON)
    tb(slide, Inches(6.4), Inches(2.4), Inches(6.45), Inches(0.4),
       "Point to your guess. Any guess is a good guess.", size=13,
       color=SOFT)
    for i, (pic, guess) in enumerate(STORY_PREDICT):
        top = Inches(2.95 + i * 1.1)
        add_round(slide, Inches(6.4), top, Inches(6.45), Inches(0.95), WHITE)
        tb(slide, Inches(6.7), top + Inches(0.2), Inches(0.8), Inches(0.6),
           pic, size=24, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.7), top + Inches(0.22), Inches(4.8), Inches(0.5),
           guess, size=22, bold=True, color=INK)
    add_round(slide, Inches(6.4), Inches(6.42), Inches(6.45), Inches(0.46),
              L_GOLD)
    tb(slide, Inches(6.65), Inches(6.49), Inches(5.95), Inches(0.36),
       "Remember your guess. We check it at the end!", size=13, bold=True,
       color=GOLD)


def s58_story_1():
    story_slide(STORY[0])


def s59_story_2():
    story_slide(STORY[1])


def s60_story_3():
    story_slide(STORY[2])


def s61_story_4():
    story_slide(STORY[3])


def s62_story_5():
    story_slide(STORY[4])


def s63_detect_a():
    slide, n = new_slide("🔎 Castle Story Detective — Question 1 and 2",
                         "COMPREHENSION", "77–85 min", "King's Room",
                         CRIMSON)
    one_task(slide, "Point to your answer. You can look back at the story.",
             CRIMSON)
    question_rows(slide, DETECT_A, 1, CRIMSON)
    hint(slide, "Read the question again slowly, then read each choice.",
         6.42)


def s64_detect_b():
    slide, n = new_slide("🔎 Castle Story Detective — Question 3 and 4",
                         "COMPREHENSION", "77–85 min", "King's Room",
                         SAPPHIRE)
    one_task(slide, "Two more. Go back to the part if you need it.",
             SAPPHIRE)
    question_rows(slide, DETECT_B, 3, SAPPHIRE)
    hint(slide, "Cover one wrong choice. Two choices is still good "
                "thinking.", 6.42)


def s65_sequence():
    slide, n = new_slide("🧩 Put the Story in Order", "COMPREHENSION",
                         "77–85 min", "King's Room", ROYAL)
    one_task(slide, "These four things are mixed up. Number them 1 to 4.",
             ROYAL)
    for i, (letter, pic, event) in enumerate(SEQUENCE):
        left = Inches(0.5 + i * 3.21)
        add_round(slide, left, Inches(1.9), Inches(3.0), Inches(3.3),
                  L_ROYAL)
        add_oval(slide, left + Inches(1.25), Inches(2.1), Inches(0.5),
                 Inches(0.5), ROYAL)
        tb(slide, left + Inches(1.25), Inches(2.16), Inches(0.5),
           Inches(0.4), letter, size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.75), Inches(3.0), Inches(0.9), pic,
           size=40, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.72), Inches(2.6),
           Inches(0.76), event, size=16, bold=True, color=INK,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.9), Inches(4.55), Inches(1.2),
                  Inches(0.55), WHITE)
        tb(slide, left + Inches(0.9), Inches(4.65), Inches(1.2), Inches(0.4),
           "___", size=18, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.4), Inches(12.35), Inches(0.85),
              WHITE)
    tb(slide, Inches(0.85), Inches(5.5), Inches(11.6), Inches(0.32),
       "NOW WRITE THE ORDER", size=11, bold=True, color=ROYAL)
    tb(slide, Inches(0.85), Inches(5.82), Inches(11.6), Inches(0.42),
       "____  →  ____  →  ____  →  ____", size=26, bold=True, color=SOFT)
    hint(slide, "Ask what happened at the very start of the story.", 6.42)


def s66_evidence():
    slide, n = new_slide("🔎 Show Me the Evidence", "COMPREHENSION",
                         "77–85 min", "King's Room", EMERALD)
    one_task(slide, "Do not just tell me. Point to the line in the story.",
             EMERALD)
    for i, (question, part, color, light) in enumerate(EVIDENCE):
        top = Inches(1.9 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34),
                  light)
        add_oval(slide, Inches(0.8), top + Inches(0.42), Inches(0.52),
                 Inches(0.52), color)
        tb(slide, Inches(0.8), top + Inches(0.48), Inches(0.52), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.16), Inches(4.7),
           Inches(0.74), question, size=19, bold=True, color=INK)
        add_round(slide, Inches(1.55), top + Inches(0.88), Inches(1.8),
                  Inches(0.4), WHITE)
        tb(slide, Inches(1.55), top + Inches(0.93), Inches(1.8),
           Inches(0.32), f"look in {part}", size=11, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        add_round(slide, Inches(6.4), top + Inches(0.3), Inches(6.15),
                  Inches(0.74), WHITE)
        tb(slide, Inches(6.65), top + Inches(0.4), Inches(5.7), Inches(0.34),
           "🔎  Go back and put your finger on the line.", size=13,
           bold=True, color=color)
        tb(slide, Inches(6.65), top + Inches(0.72), Inches(5.7), Inches(0.3),
           "☐ found it", size=11, color=SOFT)
    add_round(slide, Inches(0.5), Inches(6.46), Inches(12.35), Inches(0.44),
              L_EMERALD)
    tb(slide, Inches(0.8), Inches(6.53), Inches(11.7), Inches(0.34),
       "Finding the line in the story is the real skill here — not "
       "remembering it.", size=12, bold=True, color=EMERALD)


def s67_king_challenge():
    slide, n = new_slide("👑 The King's Reading Challenge", "CHALLENGE",
                         "85–88 min", "King's Room", GOLD)
    one_task(slide, "Five last jobs. You already know all of them.", GOLD)
    for i, (pic, label, task, note, color, light) in enumerate(
            KING_CHALLENGE):
        left = Inches(0.5 + i * 2.52)
        add_round(slide, left, Inches(1.85), Inches(2.3), Inches(4.4),
                  light)
        add_round(slide, left, Inches(1.85), Inches(2.3), Inches(0.5),
                  color)
        tb(slide, left, Inches(1.93), Inches(2.3), Inches(0.36), label,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(2.3), Inches(0.9), pic, size=38,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(3.5), Inches(2.0),
                  Inches(1.5), WHITE)
        short = len(task) <= 6
        size = 30 if len(task) <= 3 else (23 if short else
                                          (20 if len(task) <= 12 else 16))
        top = 3.82 if len(task) <= 12 else 3.7
        tb(slide, left + Inches(0.2), Inches(top), Inches(1.9), Inches(1.1),
           task, size=size, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Arial Black" if short else "Calibri")
        tb(slide, left + Inches(0.15), Inches(5.12), Inches(2.0),
           Inches(0.36), note, size=11, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.6), Inches(5.55), Inches(1.1),
                  Inches(0.55), WHITE)
        tb(slide, left + Inches(0.6), Inches(5.66), Inches(1.1),
           Inches(0.4), "☐", size=17, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Hints are still allowed here. Finishing is what matters.",
         6.45)


def s68_champion():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26),
             GOLD)
    for x, y in [(0.35, 3.9), (12.3, 3.9), (0.6, 5.4), (12.1, 5.4)]:
        tb(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.7), "⭐",
           size=24, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(0.8), Inches(12), Inches(1.1), "🏆",
       size=56, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.9), Inches(12), Inches(0.9),
       "CASTLE READING CHAMPION!", size=42, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.85), Inches(12), Inches(0.5),
       "Look what you did today —", size=18,
       color=RGBColor(0xE6, 0xC5, 0x7A), align=PP_ALIGN.CENTER, italic=True)
    for i, (icon, line) in enumerate(CHAMPION_LINES):
        left = Inches(1.35 + i * 2.15)
        add_round(slide, left, Inches(3.55), Inches(1.95), Inches(1.5),
                  RGBColor(0x24, 0x2B, 0x38))
        tb(slide, left, Inches(3.72), Inches(1.95), Inches(0.6), icon,
           size=20, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.12), Inches(4.32), Inches(1.7),
           Inches(0.62), line, size=11, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(2.9), Inches(5.25), Inches(7.5), Inches(1.1),
              GOLD)
    tb(slide, Inches(2.9), Inches(5.5), Inches(7.5), Inches(0.62),
       "⭐  \"I CAN READ!\"", size=32, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(1.2), Inches(6.5), Inches(10.9), Inches(0.42),
       "READING IS SOMETHING I CAN DO!", size=16, bold=True,
       color=RGBColor(0xE6, 0xC5, 0x7A), align=PP_ALIGN.CENTER)
    footer(slide, n, "88–90 min", "King's Room")
    fade(slide)


def game_cards(slide, games, start_index, width, gap):
    for i, (icon, name, prompt, cards, color, light) in enumerate(games):
        left = Inches(0.5 + i * gap)
        add_round(slide, left, Inches(1.9), Inches(width), Inches(4.3),
                  light)
        add_round(slide, left, Inches(1.9), Inches(width), Inches(0.5),
                  color)
        tb(slide, left, Inches(1.98), Inches(width), Inches(0.36),
           f"GAME {start_index + i} — {name}", size=11, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(width), Inches(0.62), icon,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.2), Inches(width - 0.4),
           Inches(0.6), prompt, size=13, bold=True, color=INK,
           align=PP_ALIGN.CENTER)
        for j, card in enumerate(cards):
            add_round(slide, left + Inches(0.25), Inches(3.95 + j * 0.85),
                      Inches(width - 0.5), Inches(0.72), WHITE)
            tb(slide, left + Inches(0.25), Inches(4.1 + j * 0.85),
               Inches(width - 0.5), Inches(0.5), card, size=14, bold=True,
               color=color, align=PP_ALIGN.CENTER)


def s69_games_a():
    slide, n = new_slide("🎲 Extra Game Bank — If There Is Time", "OPTIONAL",
                         "", "King's Room", SAPPHIRE)
    one_task(slide, "Seven spare games. Play any of them, in any order.",
             SAPPHIRE)
    game_cards(slide, GAMES[:4], 1, width=2.9, gap=3.11)
    hint(slide, "Every game uses today's words only. Nothing new to learn.",
         6.45)


def s70_games_b():
    slide, n = new_slide("🎲 Extra Game Bank — Three More", "OPTIONAL", "",
                         "King's Room", ROYAL)
    one_task(slide, "Three more, plus when to use them.", ROYAL)
    game_cards(slide, GAMES[4:], 5, width=2.9, gap=3.11)
    add_round(slide, Inches(9.83), Inches(1.9), Inches(3.02), Inches(4.3),
              L_GREY)
    tb(slide, Inches(10.05), Inches(2.1), Inches(2.6), Inches(0.4),
       "WHEN TO USE THESE", size=12, bold=True, color=ROYAL)
    bullets(slide, Inches(10.05), Inches(2.6), Inches(2.6), Inches(3.4),
            GAME_WHEN, size=12, sp=10)
    hint(slide, "Stop a game while he is still enjoying it, not after.",
         6.45)


def s71_support():
    slide, n = new_slide("🧰 Support System — For the Teacher", "TEACHER", "",
                         "King's Room", STONE)
    one_task(slide, "Levels, the hint ladder, and how to handle a hard "
                    "sentence.", STONE)
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
       color=STONE)
    for i, (label, text, color) in enumerate(SUPPORT_LADDER):
        top = Inches(4.66 + i * 0.46)
        add_round(slide, Inches(0.5), top, Inches(6.0), Inches(0.42),
                  L_GREY)
        add_round(slide, Inches(0.62), top + Inches(0.05), Inches(1.2),
                  Inches(0.32), color)
        tb(slide, Inches(0.62), top + Inches(0.07), Inches(1.2),
           Inches(0.28), label, size=9, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.0), top + Inches(0.05), Inches(4.3),
           Inches(0.32), text, size=12, bold=True, color=INK)
    tb(slide, Inches(6.85), Inches(4.25), Inches(3.2), Inches(0.4),
       "A HARD SENTENCE, STEP BY STEP", size=12, bold=True, color=SAPPHIRE)
    bullets(slide, Inches(6.85), Inches(4.72), Inches(3.2), Inches(2.1),
            SENTENCE_STEPS, size=10, sp=4)
    tb(slide, Inches(10.2), Inches(4.25), Inches(2.65), Inches(0.4),
       "SAY THIS", size=12, bold=True, color=EMERALD)
    bullets(slide, Inches(10.2), Inches(4.72), Inches(2.65), Inches(2.1),
            PRAISE, size=10, color=EMERALD, sp=4)


def s72_assessment():
    slide, n = new_slide("📋 End-of-Class Assessment", "TEACHER", "",
                         "King's Room", GOLD)
    one_task(slide, "Tick one box per skill right after the lesson.", GOLD)
    heads = ["SKILL", "INDEPENDENT", "WITH SUPPORT", "NEEDS MORE PRACTICE"]
    widths = [5.15, 2.4, 2.4, 2.4]
    lefts = [0.5]
    for w in widths[:-1]:
        lefts.append(lefts[-1] + w)
    add_rect(slide, Inches(0.5), Inches(1.82), Inches(12.35), Inches(0.42),
             GOLD)
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
    add_round(slide, Inches(0.5), Inches(6.25), Inches(12.35), Inches(0.68),
              L_GOLD)
    for i, item in enumerate(REVIEW_NEXT):
        tb(slide, Inches(0.8 + i * 4.1), Inches(6.45), Inches(3.9),
           Inches(0.34), f"{item}:  ____________", size=11, bold=True,
           color=INK)


def s73_answer_key():
    slide, n = new_slide("🔑 Answer Key", "TEACHER", "", "King's Room",
                         CRIMSON)
    one_task(slide, "For the teacher only.", CRIMSON)
    cols = [
        ("SOUNDS AND VOWELS (10–21)",
         ["Starts with /k/: the key · /b/: the bell",
          "Unlock: door · flag · crown · key · bell",
          "Vowel sort: A cat, hat · E red ·",
          " I big · O box, hot",
          "Which key: I · O · A · O · E"], GOLD),
        ("BLENDING AND FAMILIES (29–39)",
         ["Mirror: FOX · BOX · CAT",
          "Missing letter: O · A · E",
          "Spells: HAT · BIG · BOX · SUN",
          "Book sort: -AT cat, hat, bat ·",
          " -UN run, sun, fun",
          "Odd one: SUN · BIG · MAP",
          "Rhymes: HAT · FOX · SUN"], SAPPHIRE),
        ("SENTENCES (51–55)",
         ["Puzzles: I see a cat. · The cat is big.",
          " My hat is red. · I can run.",
          "Picture match: the fox · the crown",
          "True or silly: sense · silly · sense · silly",
          "Sight words: I, see, a, the, is, my"], EMERALD),
        ("STORY AND QUESTIONS (58–67)",
         ["Q1 King Ben · Q2 A crown ·",
          " Q3 Near a box · Q4 He put it on",
          "Order: B → C → D → A",
          "Evidence: behind a box (Part 4) ·",
          " Max (Part 3) · Ben smiles (Part 5)",
          "Story words: 147"], CRIMSON),
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
              L_CRIMSON)
    tb(slide, Inches(0.8), Inches(6.76), Inches(11.7), Inches(0.32),
       "Accept any sensible answer for the rhyme challenge, the castle talk "
       "and the speaking tasks.", size=11, bold=True, color=CRIMSON)


BUILDERS = [
    s01_title, s02_junior_reader, s03_map, s04_castle_talk, s05_sound_k,
    s06_sound_c, s07_sound_d, s08_sound_b, s09_sound_f, s10_find_k,
    s11_find_b, s12_unlock_a, s13_unlock_b, s14_gate_done, s15_vowel_a,
    s16_vowel_e, s17_vowel_i, s18_vowel_o, s19_vowel_sort, s20_which_key_a,
    s21_which_key_b, s22_key_done, s23_blend_cat, s24_blend_hat,
    s25_blend_big, s26_blend_sit, s27_blend_fox, s28_blend_box,
    s29_mirror_game, s30_missing_letter, s31_spells, s32_mirror_done,
    s33_family_at, s34_family_ig, s35_family_ox, s36_family_un, s37_book_sort,
    s38_odd_one, s39_rhyme, s40_library_done, s41_king_says, s42_king_reads,
    s43_king_freeze, s44_sight_i, s45_sight_see, s46_sight_a_the,
    s47_sight_is_my, s48_build_cat, s49_build_big, s50_build_red,
    s51_magic_door, s52_puzzles_a, s53_puzzles_b, s54_picture_match,
    s55_true_silly, s56_puzzle_done, s57_story_intro, s58_story_1,
    s59_story_2, s60_story_3, s61_story_4, s62_story_5, s63_detect_a,
    s64_detect_b, s65_sequence, s66_evidence, s67_king_challenge,
    s68_champion, s69_games_a, s70_games_b, s71_support, s72_assessment,
    s73_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL}, built {len(prs.slides)}"

OUT = "Grade2_Magic_Castle_Reading_Quest_90min.pptx"
prs.save(OUT)

story_words = sum(len(line.split()) for part in STORY for line in part[4])
with_notes = sum(1 for s in prs.slides
                 if s.has_notes_slide
                 and s.notes_slide.notes_text_frame.text.strip())
print(f"saved {OUT}")
print(f"slides: {len(prs.slides)}")
print(f"story words: {story_words}")
print(f"slides with notes: {with_notes}")
