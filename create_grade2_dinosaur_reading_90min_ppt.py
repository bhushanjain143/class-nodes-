"""Grade 2 reading lesson - 90 minutes, 62 slides, no speaker notes.

"Dinosaur Discovery Reading Mission" - the child is a Dinosaur Reading Explorer
travelling from Dinosaur Valley to the Fossil Museum. Reading is built one step
at a time: sound, letter, blend, word, picture, phrase, sentence, story, question.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Dino Hint strip, and slides 60-62 hold the support system,
the assessment checklist and the answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x12, 0x26, 0x1C)
DARK = RGBColor(0x22, 0x30, 0x2A)
SOFT = RGBColor(0x74, 0x85, 0x7C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAGE = RGBColor(0xF5, 0xF8, 0xF3)
LAVA = RGBColor(0xD9, 0x53, 0x2C)
JUNGLE = RGBColor(0x1B, 0x7A, 0x4B)
AMBER = RGBColor(0xD9, 0x9A, 0x18)
STONE = RGBColor(0x4A, 0x5A, 0x6B)
PLUM = RGBColor(0x7B, 0x4F, 0xA8)
FERN = RGBColor(0x12, 0x83, 0x77)
SAND = RGBColor(0xB8, 0x80, 0x3A)
L_LAVA = RGBColor(0xFB, 0xEA, 0xE4)
L_JUNGLE = RGBColor(0xE3, 0xF3, 0xEA)
L_AMBER = RGBColor(0xFC, 0xF2, 0xDA)
L_STONE = RGBColor(0xE9, 0xED, 0xF2)
L_PLUM = RGBColor(0xF0, 0xEA, 0xF9)
L_FERN = RGBColor(0xE0, 0xF1, 0xEF)
L_SAND = RGBColor(0xF7, 0xEE, 0xE1)
L_GREY = RGBColor(0xF1, 0xF4, 0xF1)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 62
_counter = {"n": 0}

# ------------------------------------------------------------------ content

STOPS = [("🌋", "Dinosaur Valley", "Sounds", "0–17 min", LAVA, L_LAVA),
         ("🥚", "Egg Cave", "Short vowels", "17–27 min", PLUM, L_PLUM),
         ("🦖", "Dino Camp", "Blending & matching", "27–49 min", AMBER, L_AMBER),
         ("🌿", "Jungle Trail", "Words & sentences", "49–75 min", JUNGLE, L_JUNGLE),
         ("🏛️", "Fossil Museum", "Story & questions", "75–90 min", STONE, L_STONE)]

PROMISES = [("🦖", "Dino Hint", "A hint is always waiting when a word looks hard."),
            ("🟢", "Dino Help", "We start together. You are never reading alone."),
            ("🟡", "Dino Mission", "Then you try it, and I go quiet."),
            ("⭐", "Dino Challenge", "Finished early? There is always one more.")]

VALLEY_SCENE = [("🦕", "dinosaur"), ("🥚", "egg"), ("🪨", "rock"), ("🌋", "volcano"),
                ("🌿", "jungle"), ("👣", "footprint")]

TALK_QS = [("🗣️", "What do you see?"), ("📏", "Is the dinosaur big or small?"),
           ("🎨", "What color is the egg?"), ("🪨", "Can you find the rock?")]

STARTERS = ["I see a ______.", "The dino is ______."]

SOUND_CARDS = [("🦕", "DINOSAUR", "D", "/d/"), ("🥚", "EGG", "E", "/e/"),
               ("👣", "FOOTPRINT", "F", "/f/"), ("🪨", "ROCK", "R", "/r/"),
               ("🦎", "TAIL", "T", "/t/"), ("🦴", "BONE", "B", "/b/")]

FIRST_A = [("D", [("🐟", "fish"), ("🦕", "dino"), ("🥚", "egg")]),
           ("E", [("🥚", "egg"), ("🪨", "rock"), ("🦴", "bone")]),
           ("R", [("🦎", "tail"), ("🪨", "rock"), ("🦕", "dino")])]
FIRST_B = [("B", [("🥚", "egg"), ("🦴", "bone"), ("👣", "footprint")]),
           ("T", [("🦎", "tail"), ("🪨", "rock"), ("🥚", "egg")]),
           ("F", [("🦕", "dino"), ("👣", "footprint"), ("🦴", "bone")])]
FIRST_C = [("M", [("🟤", "mud"), ("🦴", "bone"), ("☀️", "sun")]),
           ("S", [("🪨", "rock"), ("☀️", "sun"), ("🥚", "egg")])]

ENDING = [("🎒", "BAG", ["/g/", "/t/", "/p/"]), ("🎩", "HAT", ["/p/", "/t/", "/n/"]),
          ("☀️", "SUN", ["/m/", "/n/", "/g/"]), ("🪵", "LOG", ["/g/", "/d/", "/k/"])]

VOWEL_EGGS = [("A", "🗺️", "map", LAVA, L_LAVA), ("E", "🛏️", "bed", JUNGLE, L_JUNGLE),
              ("I", "🪑", "sit", FERN, L_FERN), ("O", "🪵", "log", PLUM, L_PLUM)]

EGG_A = [("BED", "/b/ /e/ /d/", [("🛏️", "bed"), ("🎒", "bag"), ("🪵", "log")]),
         ("MAP", "/m/ /a/ /p/", [("🎩", "hat"), ("🗺️", "map"), ("🥚", "egg")]),
         ("SIT", "/s/ /i/ /t/", [("☀️", "sun"), ("🪨", "rock"), ("🪑", "sit")])]
EGG_B = [("LOG", "/l/ /o/ /g/", [("🪵", "log"), ("🛏️", "bed"), ("🐱", "cat")]),
         ("RED", "/r/ /e/ /d/", [("🔴", "red"), ("🟢", "green"), ("🟤", "mud")]),
         ("HOT", "/h/ /o/ /t/", [("💧", "wet"), ("🔥", "hot"), ("🔴", "red")])]
EGG_BONUS = [("BIG", "/b/ /i/ /g/", [("🐜", "small"), ("🦖", "big"), ("👣", "print")]),
             ("CAT", "/k/ /a/ /t/", [("🐱", "cat"), ("🦇", "bat"), ("🎩", "hat")])]

DINO_WORDS = [("dino", "🦕"), ("egg", "🥚"), ("dig", "⛏️"), ("big", "🦖"),
              ("hot", "🔥"), ("log", "🪵"), ("mud", "🟤"), ("sun", "☀️"),
              ("run", "🏃"), ("red", "🔴"), ("wet", "💧"), ("top", "⬆️"),
              ("map", "🗺️"), ("bag", "🎒"), ("hat", "🎩"), ("sit", "🪑"),
              ("fin", "🦈"), ("hop", "🐸")]

SIX_STEPS = [("1", "I SAY", "Teacher says the sound.", "/d/", LAVA, L_LAVA),
             ("2", "YOU SAY", "You say it back to me.", "/d/", AMBER, L_AMBER),
             ("3", "POINT", "Point to each letter.", "D  I  G", FERN, L_FERN),
             ("4", "BLEND", "Push the sounds together.", "d-i-g", PLUM, L_PLUM),
             ("5", "READ", "Read the whole word.", "DIG", JUNGLE, L_JUNGLE),
             ("6", "USE IT", "Use it in a sentence.", "I can dig.", STONE, L_STONE)]

BLEND_A = [("🐱", "CAT", ["C", "A", "T"], ["/k/", "/a/", "/t/"], LAVA, L_LAVA),
           ("🗺️", "MAP", ["M", "A", "P"], ["/m/", "/a/", "/p/"], AMBER, L_AMBER),
           ("⛏️", "DIG", ["D", "I", "G"], ["/d/", "/i/", "/g/"], FERN, L_FERN)]
BLEND_B = [("🦖", "BIG", ["B", "I", "G"], ["/b/", "/i/", "/g/"], JUNGLE, L_JUNGLE),
           ("🔥", "HOT", ["H", "O", "T"], ["/h/", "/o/", "/t/"], LAVA, L_LAVA),
           ("🪵", "LOG", ["L", "O", "G"], ["/l/", "/o/", "/g/"], SAND, L_SAND)]
BLEND_C = [("🏃", "RUN", ["R", "U", "N"], ["/r/", "/u/", "/n/"], PLUM, L_PLUM),
           ("☀️", "SUN", ["S", "U", "N"], ["/s/", "/u/", "/n/"], AMBER, L_AMBER),
           ("⬆️", "TOP", ["T", "O", "P"], ["/t/", "/o/", "/p/"], STONE, L_STONE)]

TRICK = [("🐱", ["CAT", "CAG", "CAB"]), ("⛏️", ["DIP", "DIG", "DIB"]),
         ("☀️", ["SUN", "SUM", "SUP"]), ("🪵", ["LOD", "LOG", "LOP"])]

NO_SUPPORT = [("DIG", "⛏️"), ("BIG", "🦖"), ("HOT", "🔥"),
              ("LOG", "🪵"), ("RUN", "🏃"), ("SUN", "☀️")]

FOSSIL_HOW = [("1", "🦴", "READ THE FOSSIL", "Sound out the word on the bone.", LAVA),
              ("2", "👀", "LOOK AT THE PICTURES", "Three pictures are waiting.", AMBER),
              ("3", "🤝", "MAKE THE MATCH", "Point to the one that fits.", JUNGLE)]

MATCH_A = [("DIG", [("☀️", "sun"), ("⛏️", "dig"), ("🪵", "log")]),
           ("BIG", [("🦖", "big"), ("🐜", "small"), ("🥚", "egg")]),
           ("SUN", [("🟤", "mud"), ("🪨", "rock"), ("☀️", "sun")])]
MATCH_B = [("EGG", [("🥚", "egg"), ("🦴", "bone"), ("🎒", "bag")]),
           ("LOG", [("🎩", "hat"), ("🪵", "log"), ("🐸", "hop")]),
           ("MUD", [("🟤", "mud"), ("🔴", "red"), ("🗺️", "map")])]

FOSSIL_BONUS = [("🦴", "BONE"), ("👣", "FOOTPRINT"), ("🦎", "TAIL"), ("🐾", "CLAW")]

DINO_SAYS = [("👏", "Dino says clap two times."),
             ("🪑", "Dino says tap your desk."),
             ("👉", "Dino says point to the word BIG."),
             ("🔊", "Dino says say the /d/ sound."),
             ("📖", "Dino says read the word CAT."),
             ("🦖", "Dino says make a dinosaur roar!")]

FAMILIES = [("-AT", ["cat", "hat", "bat", "sat"], LAVA, L_LAVA),
            ("-IG", ["big", "dig", "pig"], FERN, L_FERN),
            ("-UN", ["sun", "run", "fun"], AMBER, L_AMBER),
            ("-OP", ["hop", "top", "mop"], PLUM, L_PLUM)]

FAMILY_A = [("-AT", ["CAT", "BAT", "SUN", "HAT"], LAVA, L_LAVA),
            ("-IG", ["BIG", "HOP", "DIG", "PIG"], FERN, L_FERN)]
FAMILY_B = [("-UN", ["SUN", "RUN", "MAP", "FUN"], AMBER, L_AMBER),
            ("-OP", ["TOP", "HOP", "BAT", "MOP"], PLUM, L_PLUM)]

NEW_WORDS = [("S", "-AT", "SAT", "🪑", LAVA, L_LAVA),
             ("P", "-IG", "PIG", "🐷", FERN, L_FERN),
             ("F", "-UN", "FUN", "🎉", AMBER, L_AMBER),
             ("M", "-OP", "MOP", "🧹", PLUM, L_PLUM)]

SIGHT_1 = [("the", "the dino", "🦕"), ("a", "a big egg", "🥚"), ("is", "It is red.", "🔴")]
SIGHT_2 = [("in", "in the mud", "🟤"), ("my", "my dino", "🧒"), ("can", "I can dig.", "⛏️")]

PATH_A = [("the", ["a", "the", "is"]), ("a", ["can", "a", "my"]),
          ("is", ["is", "in", "my"]), ("in", ["my", "in", "the"])]
PATH_B = [("my", ["my", "can", "a"]), ("can", ["is", "can", "in"]),
          ("see", ["see", "the", "my"]), ("on", ["on", "go", "a"])]

PHRASES = [("the dino", "🦕"), ("a big dino", "🦖"), ("the red egg", "🥚"),
           ("my dino", "🧒"), ("can run", "🏃"), ("in the mud", "🟤")]

READ_STEPS = [("STEP 1", "I read", "Teacher reads it first.", LAVA),
              ("STEP 2", "We read", "Teacher and child together.", AMBER),
              ("STEP 3", "You read", "Child reads it alone — if ready.", JUNGLE)]

SENTENCES_A = [("The dino is big.", "🦖", LAVA, L_LAVA),
               ("The egg is red.", "🥚", PLUM, L_PLUM)]
SENTENCES_B = [("I see a dino.", "👀", FERN, L_FERN),
               ("The dino can run.", "🏃", AMBER, L_AMBER),
               ("The dino can dig.", "⛏️", JUNGLE, L_JUNGLE)]

PUZZLES_A = [(["is", "dino", "The", "big"], "The dino is big.", "🦖", LAVA, L_LAVA),
             (["egg", "The", "red", "is"], "The egg is red.", "🥚", PLUM, L_PLUM)]
PUZZLES_B = [(["a", "see", "I", "dino"], "I see a dino.", "👀", FERN, L_FERN),
             (["can", "dino", "The", "dig"], "The dino can dig.", "⛏️", JUNGLE,
              L_JUNGLE)]

OWN_SENTENCE = [("🦕", "dino"), ("🥚", "egg"), ("🪨", "rock"), ("🟤", "mud")]

STORY_WATCH = [("dino", "🦕"), ("egg", "🥚"), ("mud", "🟤"), ("rock", "🪨"),
               ("log", "🪵"), ("bone", "🦴")]

STORY = [
    ("Part 1", "🦕", LAVA, L_LAVA,
     ["Ben is on a big hill.", "The sun is hot.", "Ben can see a rock.",
      "A little dino sits on the rock.", "The dino is green.", "It is not big."]),
    ("Part 2", "🥚", PLUM, L_PLUM,
     ["Ben sees a red egg.", "The egg is in the mud.", "Ben can see the egg.",
      "The little dino looks at it.", "It can dig in the mud.", "Dig, dig, dig!"]),
    ("Part 3", "🌿", JUNGLE, L_JUNGLE,
     ["The dino can run fast.", "Ben and the dino run to the jungle.",
      "They see a big log.", "A bug is on the log.", "The dino can hop on the log.",
      "Ben can hop too."]),
    ("Part 4", "🦴", STONE, L_STONE,
     ["Ben digs in the mud.", "He finds a small bone.", "It is a dino fossil!",
      "They walk back to the hill.", "Ben and the dino are happy.",
      "\"You are my pal,\" says Ben."]),
]

EASY_LINES = ["The sun is hot.", "The dino is green.", "It is not big.",
              "Dig, dig, dig!", "Ben can hop too.", "He finds a small bone."]

QUESTIONS_A = [("🧒", "Who sees the dino?", ["Ben", "Sam", "Mia"], "Part 1"),
               ("🎨", "What color is the dino?", ["🟢 Green", "🔴 Red", "🔵 Blue"],
                "Part 1")]
QUESTIONS_B = [("🥚", "What does Ben see in the mud?", ["A red egg", "A hat", "A bag"],
                "Part 2"),
               ("🏃", "What can the dino do?", ["Run and dig", "Swim", "Fly"], "Part 3")]
QUESTIONS_C = [("🦴", "What do they find at the end?",
                ["A bone fossil", "A map", "A cap"], "Part 4"),
               ("🌿", "Where do Ben and the dino run?",
                ["To the jungle", "To school", "To the sea"], "Part 3")]

FINAL_A = [("1", "🦴", "READ A WORD", "Read this word:", "BIG", LAVA, L_LAVA),
           ("2", "👣", "BLEND A WORD", "Blend these sounds:", "/d/  /i/  /g/", FERN,
            L_FERN),
           ("3", "🥚", "READ A SIGHT WORD", "Read this word:", "THE", PLUM, L_PLUM)]
FINAL_B = [("4", "📕", "READ A SENTENCE", "Read this out loud:", "The dino can run.",
            JUNGLE, L_JUNGLE),
           ("5", "🔎", "ANSWER A QUESTION", "From the story:",
            "What can the dino do?", AMBER, L_AMBER)]

CAN_READ = [("🔊", "Beginning sounds", "d  e  f  r  t  b"),
            ("🔚", "Ending sounds", "bag  hat  sun  log"),
            ("🥚", "Short vowels", "a  e  i  o"),
            ("👣", "Blending", "/d/ /i/ /g/ → DIG"),
            ("📦", "CVC words", "cat  map  dig  big  hot  log"),
            ("👨‍👩‍👧", "Word families", "-at  -ig  -un  -op"),
            ("🗺️", "Sight words", "the  a  is  in  my  can"),
            ("🔗", "Short phrases", "a big dino"),
            ("📕", "Sentences", "The dino can dig."),
            ("📖", "A whole story", "The Little Dino")]

SUPPORT_LEVELS = [("🟢", "DINO HELP", "Picture + sound + teacher support",
                   "Say the word first, then let him repeat it.", JUNGLE, L_JUNGLE),
                  ("🟡", "DINO MISSION", "He reads with limited support",
                   "Give the first sound only, then go quiet.", AMBER, L_AMBER),
                  ("⭐", "DINO CHALLENGE", "He reads alone or makes a sentence",
                   "Ask for a sentence using the word.", PLUM, L_PLUM)]

HINT_LADDER = [("HINT 1", "Look at the first sound.", LAVA),
               ("HINT 2", "Let's sound it together.", AMBER),
               ("HINT 3", "Let's blend it together.", FERN),
               ("HINT 4", "I read it, then you repeat.", JUNGLE)]

HINT_BANK = ["Look at the first letter.", "Say each sound.",
             "Put the sounds together.", "Look at the picture.", "Try it slowly.",
             "Read it one more time."]

PRAISE = ["\"Good listening!\"", "\"You found the first sound!\"",
          "\"Let's try that one together.\"", "\"You are getting closer!\"",
          "\"Nice blending!\"", "\"You read that all by yourself.\""]

RUBRIC_SKILLS = ["Beginning sounds", "Short vowels", "CVC words", "Blending",
                 "Word families", "Sight words", "Phrase reading", "Sentence reading",
                 "Story reading", "Comprehension"]

TODAY_I_CAN = ["Read some new words", "Blend sounds", "Read a sentence",
               "Understand a short story"]

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
             RGBColor(0xDE, 0xE6, 0xDD))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), JUNGLE)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), INK)
    msg = "🦖 Dinosaur Discovery Reading Mission  |  Grade 2  |  90 min"
    if stop:
        msg += f"  |  {stop}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=10,
       color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, stop, accent=JUNGLE, bg=PAGE):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, accent)
    add_round(slide, Inches(0.4), Inches(0.26), Inches(3.5), Inches(0.42), accent)
    tb(slide, Inches(0.4), Inches(0.31), Inches(3.5), Inches(0.34), tag, size=12,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if timing:
        add_round(slide, Inches(10.3), Inches(0.26), Inches(2.65), Inches(0.42), INK)
        tb(slide, Inches(10.3), Inches(0.31), Inches(2.65), Inches(0.34), timing,
           size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title, size=28,
       bold=True, color=INK, font="Georgia")
    footer(slide, n, timing, stop)
    fade(slide)
    return slide, n


def one_task(slide, text, color=LAVA, top=1.34):
    """States the single job of this slide in words the child understands."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text, size=16,
              bold=True, color=color)


def hint(slide, text, top=6.42, label="🦖 DINO HINT", fill=L_AMBER, color=SAND):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.2), Inches(0.34), label, size=12,
       bold=True, color=color)
    tb(slide, Inches(3.05), Inches(top + 0.05), Inches(9.5), Inches(0.36), text, size=13,
       bold=True, color=INK)


def i_we_you(slide, top=6.42):
    steps = [("I READ", "Teacher", LAVA), ("WE READ", "Together", AMBER),
             ("YOU READ", "You!", JUNGLE)]
    for i, (label, who, color) in enumerate(steps):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(top), Inches(3.9), Inches(0.46), color)
        tb(slide, left, Inches(top + 0.05), Inches(3.9), Inches(0.36),
           f"{label}  ·  {who}", size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def numbered_rows(slide, count, start_index, accent, top_start=1.9, gap=1.5, height=1.34,
                  fill=WHITE):
    """Shared row scaffold; returns the top edge of each numbered row."""
    tops = []
    for i in range(count):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), fill)
        add_oval(slide, Inches(0.75), top + Inches(height / 2 - 0.25), Inches(0.5),
                 Inches(0.5), accent)
        tb(slide, Inches(0.75), top + Inches(height / 2 - 0.19), Inches(0.5),
           Inches(0.4), str(start_index + i), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tops.append(top)
    return tops


def first_sound_rows(slide, items, start_index):
    """Follow the First Sound: a lettered footprint, then three picture choices."""
    tops = numbered_rows(slide, len(items), start_index, LAVA, top_start=1.9, gap=1.5)
    for (letter, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.3), Inches(0.9), Inches(0.72), "👣",
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.4), top + Inches(0.36), Inches(1.7), Inches(0.62),
                  LAVA)
        tb(slide, Inches(2.4), top + Inches(0.44), Inches(1.7), Inches(0.46), letter,
           size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, (emoji, label) in enumerate(options):
            left = Inches(4.45 + j * 2.75)
            add_round(slide, left, top + Inches(0.16), Inches(2.5), Inches(1.02),
                      L_LAVA)
            tb(slide, left, top + Inches(0.2), Inches(2.5), Inches(0.58), emoji, size=26,
               align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.78), Inches(2.5), Inches(0.36), label,
               size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)


def ending_rows(slide, items, start_index):
    """Ending sounds: a word, then three sounds to choose the last one from."""
    tops = numbered_rows(slide, len(items), start_index, FERN, top_start=1.9, gap=1.12,
                         height=1.0)
    for (emoji, word, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.18), Inches(0.9), Inches(0.62), emoji,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.5), top + Inches(0.2), Inches(2.3), Inches(0.6), word,
           size=26, bold=True, color=INK, font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(5.0 + j * 2.62)
            add_round(slide, left, top + Inches(0.19), Inches(2.4), Inches(0.62),
                      L_FERN)
            tb(slide, left, top + Inches(0.27), Inches(2.4), Inches(0.46), opt, size=20,
               bold=True, color=FERN, align=PP_ALIGN.CENTER, font="Arial Black")


def egg_rows(slide, items, start_index):
    """Crack the Egg: read the word inside the egg, then pick the picture."""
    tops = numbered_rows(slide, len(items), start_index, PLUM, top_start=1.9, gap=1.5)
    for (word, sounds, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.3), Inches(0.9), Inches(0.72), "🥚",
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.4), top + Inches(0.18), Inches(2.4), Inches(0.62), word,
           size=28, bold=True, color=PLUM, font="Arial Black")
        tb(slide, Inches(2.4), top + Inches(0.82), Inches(2.4), Inches(0.4), sounds,
           size=14, color=SOFT)
        for j, (emoji, label) in enumerate(options):
            left = Inches(5.0 + j * 2.62)
            add_round(slide, left, top + Inches(0.16), Inches(2.4), Inches(1.02),
                      L_PLUM)
            tb(slide, left, top + Inches(0.2), Inches(2.4), Inches(0.58), emoji, size=26,
               align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.78), Inches(2.4), Inches(0.36), label,
               size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)


def blend_rows(slide, items, top_start=1.9, gap=1.52):
    """Dino Footprint Blend: letters in footprints, then the finished word."""
    for i, (emoji, word, letters, sounds, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), light)
        tb(slide, Inches(0.8), top + Inches(0.32), Inches(0.9), Inches(0.68), emoji,
           size=30, align=PP_ALIGN.CENTER)
        for j, letter in enumerate(letters):
            left = Inches(2.0 + j * 1.35)
            add_round(slide, left, top + Inches(0.25), Inches(1.15), Inches(0.85), WHITE)
            tb(slide, left, top + Inches(0.34), Inches(1.15), Inches(0.62), letter,
               size=30, bold=True, color=color, align=PP_ALIGN.CENTER,
               font="Arial Black")
            if j < len(letters) - 1:
                tb(slide, left + Inches(1.15), top + Inches(0.48), Inches(0.18),
                   Inches(0.45), "+", size=16, bold=True, color=SOFT,
                   align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.3), top + Inches(0.46), Inches(2.4), Inches(0.5),
           "  ".join(sounds), size=16, bold=True, color=SOFT)
        tb(slide, Inches(8.8), top + Inches(0.42), Inches(0.5), Inches(0.5), "→",
           size=22, bold=True, color=color)
        add_round(slide, Inches(9.5), top + Inches(0.25), Inches(3.1), Inches(0.85),
                  WHITE)
        tb(slide, Inches(9.5), top + Inches(0.34), Inches(3.1), Inches(0.62), word,
           size=30, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")


def trick_rows(slide, items, start_index):
    """Trick words: one picture, three near-identical spellings."""
    tops = numbered_rows(slide, len(items), start_index, AMBER, top_start=1.9, gap=1.12,
                         height=1.0)
    for (emoji, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.16), Inches(0.9), Inches(0.66), emoji,
           size=26, align=PP_ALIGN.CENTER)
        for j, opt in enumerate(options):
            left = Inches(3.1 + j * 3.2)
            add_round(slide, left, top + Inches(0.19), Inches(2.95), Inches(0.62),
                      L_AMBER)
            tb(slide, left, top + Inches(0.26), Inches(2.95), Inches(0.5), opt, size=24,
               bold=True, color=INK, align=PP_ALIGN.CENTER, font="Arial Black")


def match_rows(slide, items, start_index):
    """Find the Fossil Match: a word fossil plus three picture fossils."""
    tops = numbered_rows(slide, len(items), start_index, SAND, top_start=1.9, gap=1.5)
    for (word, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.3), Inches(0.8), Inches(0.72), "🦴",
           size=26, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.3), top + Inches(0.3), Inches(1.85), Inches(0.74),
                  SAND)
        tb(slide, Inches(2.3), top + Inches(0.41), Inches(1.85), Inches(0.54), word,
           size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, (emoji, label) in enumerate(options):
            left = Inches(4.5 + j * 2.72)
            add_round(slide, left, top + Inches(0.16), Inches(2.5), Inches(1.02),
                      L_SAND)
            tb(slide, left, top + Inches(0.2), Inches(2.5), Inches(0.58), emoji, size=26,
               align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.78), Inches(2.5), Inches(0.36), label,
               size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)


def family_rows(slide, items, start_index):
    """Who Is in the Dino Family: one family, four candidate words."""
    tops = numbered_rows(slide, len(items), start_index, JUNGLE, top_start=2.0, gap=2.15,
                         height=1.9)
    for (family, words, color, light), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.5), Inches(0.9), Inches(0.7), "🦕",
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.4), top + Inches(0.6), Inches(1.9), Inches(0.7),
                  color)
        tb(slide, Inches(2.4), top + Inches(0.71), Inches(1.9), Inches(0.5), family,
           size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, word in enumerate(words):
            left = Inches(4.65 + j * 2.0)
            add_round(slide, left, top + Inches(0.55), Inches(1.85), Inches(0.8), light)
            tb(slide, left, top + Inches(0.67), Inches(1.85), Inches(0.56), word,
               size=22, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")
        tb(slide, Inches(4.65), top + Inches(1.42), Inches(7.9), Inches(0.3),
           "Circle every word that belongs to this family.", size=11, color=SOFT)


def path_rows(slide, items, start_index, top_start=1.95):
    """Dino Sight-Word Path: find one sight word among three footprints."""
    for i, (target, options) in enumerate(items):
        top = Inches(top_start + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), WHITE)
        add_oval(slide, Inches(0.78), top + Inches(0.22), Inches(0.52), Inches(0.52),
                 STONE)
        tb(slide, Inches(0.78), top + Inches(0.28), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.27), Inches(1.0), Inches(0.44), "FIND:",
           size=13, bold=True, color=SOFT)
        add_round(slide, Inches(2.55), top + Inches(0.2), Inches(1.9), Inches(0.58),
                  STONE)
        tb(slide, Inches(2.55), top + Inches(0.29), Inches(1.9), Inches(0.42), target,
           size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(4.85 + j * 2.6)
            add_round(slide, left, top + Inches(0.14), Inches(2.4), Inches(0.7),
                      L_STONE)
            tb(slide, left + Inches(0.12), top + Inches(0.26), Inches(0.5), Inches(0.44),
               "👣", size=13, align=PP_ALIGN.CENTER)
            tb(slide, left + Inches(0.7), top + Inches(0.24), Inches(1.6), Inches(0.48),
               opt, size=20, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")


def phrase_grid(slide, items, top=1.95):
    for i, (phrase, emoji) in enumerate(items):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        t = Inches(top + row * 1.5)
        add_round(slide, left, t, Inches(5.95), Inches(1.3), L_JUNGLE)
        tb(slide, left + Inches(0.3), t + Inches(0.3), Inches(1.0), Inches(0.72), emoji,
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.5), t + Inches(0.32), Inches(4.2), Inches(0.66),
           phrase, size=26, bold=True, color=INK)


def sentence_rows(slide, items, top_start=2.0, gap=2.15, height=1.9, size=40):
    for i, (text, emoji, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), light)
        tb(slide, Inches(0.85), top + Inches(height / 2 - 0.42), Inches(1.4),
           Inches(0.85), emoji, size=int(size * 0.85), align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(height / 2 - 0.4), Inches(7.6),
           Inches(0.8), text, size=size, bold=True, color=INK)
        add_round(slide, Inches(10.4), top + Inches(height / 2 - 0.28), Inches(2.2),
                  Inches(0.56), WHITE)
        tb(slide, Inches(10.4), top + Inches(height / 2 - 0.2), Inches(2.2),
           Inches(0.4), "I · WE · YOU", size=13, bold=True, color=color,
           align=PP_ALIGN.CENTER)


def puzzle_rows(slide, items, start_index):
    """Build the Dino Sentence: shuffled word cards plus an answer line."""
    tops = numbered_rows(slide, len(items), start_index, AMBER, top_start=2.0, gap=2.15,
                         height=1.9)
    for (cards, _answer, emoji, color, light), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.5), Inches(0.9), Inches(0.72), emoji,
           size=30, align=PP_ALIGN.CENTER)
        for j, card in enumerate(cards):
            left = Inches(2.5 + j * 2.1)
            add_round(slide, left, top + Inches(0.28), Inches(1.95), Inches(0.72),
                      light)
            tb(slide, left, top + Inches(0.38), Inches(1.95), Inches(0.52), card,
               size=22, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.5), top + Inches(1.12), Inches(10.05), Inches(0.6),
                  WHITE)
        tb(slide, Inches(2.75), top + Inches(1.24), Inches(9.6), Inches(0.42),
           "Put them in order, then read it out loud:  ____________________________",
           size=13, color=SOFT)


def question_rows(slide, items, start_index, top_start=1.95, gap=2.16, height=1.92):
    for i, (emoji, question, options, where) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.7), Inches(0.52), Inches(0.52),
                 STONE)
        tb(slide, Inches(0.75), top + Inches(0.76), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.45), top + Inches(0.6), Inches(1.0), Inches(0.72), emoji,
           size=30, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(0.48), Inches(3.5), Inches(0.7), question,
           size=19, bold=True, color=INK)
        tb(slide, Inches(2.6), top + Inches(1.2), Inches(3.5), Inches(0.34),
           f"show me this in {where}", size=11, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(6.3 + j * 2.2)
            add_round(slide, left, top + Inches(0.5), Inches(2.0), Inches(0.92),
                      L_STONE)
            tb(slide, left, top + Inches(0.68), Inches(2.0), Inches(0.6), opt, size=16,
               bold=True, color=INK, align=PP_ALIGN.CENTER)


def story_slide(part, timing):
    """One story page: a big picture panel beside six short lines."""
    label, emoji, color, light = part[0], part[1], part[2], part[3]
    lines = part[4]
    slide, n = new_slide(f"📖 The Little Dino — {label}", "STORY", timing,
                         "Fossil Museum", color)
    add_round(slide, Inches(0.5), Inches(1.4), Inches(4.3), Inches(4.75), light)
    tb(slide, Inches(0.5), Inches(2.9), Inches(4.3), Inches(1.7), emoji, size=88,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(5.1), Inches(1.4), Inches(7.75), Inches(4.75), WHITE)
    for i, line in enumerate(lines):
        tb(slide, Inches(5.5), Inches(1.62 + i * 0.75), Inches(7.1), Inches(0.62), line,
           size=24, bold=True, color=INK)
    i_we_you(slide)
    return slide, n


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), JUNGLE)
    for x, y, c in [(0.6, 5.6, LAVA), (12.05, 5.55, AMBER)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(0.9), Inches(12), Inches(1.0), "🦖", size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.9), Inches(12), Inches(0.9),
       "Dinosaur Discovery Reading Mission", size=42, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.85), Inches(12), Inches(0.5),
       "Dig Into Reading!", size=22, color=RGBColor(0x9E, 0xD6, 0xB4),
       align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, label) in enumerate(VALLEY_SCENE):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(3.55), Inches(1.65), Inches(1.35),
                  RGBColor(0x1C, 0x39, 0x2A))
        tb(slide, left, Inches(3.73), Inches(1.65), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.38), Inches(1.45), Inches(0.4), label,
           size=10, color=RGBColor(0xBA, 0xD8, 0xC4), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.3), Inches(5.2), Inches(6.7), Inches(1.15), JUNGLE)
    tb(slide, Inches(3.5), Inches(5.45), Inches(6.3), Inches(0.7),
       "Grade 2  •  90 Minutes  •  Reading", size=19, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Dinosaur Valley")
    fade(slide)


def s02_explorer():
    slide, n = new_slide("🦕 You Are the Dinosaur Reading Explorer", "WELCOME",
                         "0–7 min", "Dinosaur Valley", JUNGLE)
    one_task(slide, "Today you are a Dinosaur Reading Explorer!", JUNGLE)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.5), L_JUNGLE)
    tb(slide, Inches(0.5), Inches(2.4), Inches(5.4), Inches(1.7), "🦕", size=88,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.35), Inches(5.4), Inches(0.7), "THAT'S YOU!",
       size=32, bold=True, color=JUNGLE, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.8), Inches(5.15), Inches(4.8), Inches(0.8),
       "An explorer who reads", size=15, color=SOFT, align=PP_ALIGN.CENTER)
    lines = [("🗺️", "We travel to five places today."),
             ("📖", "Each place has one reading mission."),
             ("🥚", "Every mission you finish earns a dino egg."),
             ("🏆", "At the end you become a Dino Reading Champion.")]
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.9 + i * 1.18)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(1.02), L_AMBER)
        add_oval(slide, Inches(6.5), top + Inches(0.22), Inches(0.58), Inches(0.58),
                 WHITE)
        tb(slide, Inches(6.5), top + Inches(0.28), Inches(0.58), Inches(0.44), icon,
           size=16, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.3), top + Inches(0.28), Inches(5.3), Inches(0.5), line,
           size=16, bold=True, color=INK)


def s03_map():
    slide, n = new_slide("🗺️ Our Dinosaur Mission Map", "MAP", "0–7 min",
                         "Dinosaur Valley", AMBER)
    one_task(slide, "Five stops. One reading mission at each stop.", AMBER)
    for i, (emoji, place, job, when, color, light) in enumerate(STOPS):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(3.5), light)
        tb(slide, left, Inches(2.2), Inches(2.3), Inches(0.85), emoji, size=36,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.1), Inches(2.1), Inches(0.9), place,
           size=16, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.98), Inches(2.0), Inches(0.5), job,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.65), Inches(1.7), Inches(0.42),
                  WHITE)
        tb(slide, left + Inches(0.3), Inches(4.69), Inches(1.7), Inches(0.34), when,
           size=10, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.65), Inches(12.35), Inches(0.9), L_JUNGLE)
    tb(slide, Inches(0.8), Inches(5.9), Inches(11.7), Inches(0.44),
       "🌋  →  🥚  →  🦖  →  🌿  →  🏛️   Finish all five and you are a DINO READING "
       "CHAMPION!", size=16, bold=True, color=INK, align=PP_ALIGN.CENTER)


def s04_promises():
    slide, n = new_slide("🤝 How I Help You Today", "PROMISE", "0–7 min",
                         "Dinosaur Valley", FERN)
    one_task(slide, "You never read alone. Here is how it works.", FERN)
    for i, (icon, name, detail) in enumerate(PROMISES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(2.0 + row * 2.2)
        add_round(slide, left, top, Inches(5.95), Inches(1.95), L_FERN)
        add_oval(slide, left + Inches(0.32), top + Inches(0.6), Inches(0.75),
                 Inches(0.75), WHITE)
        tb(slide, left + Inches(0.32), top + Inches(0.7), Inches(0.75), Inches(0.55),
           icon, size=20, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.3), top + Inches(0.4), Inches(4.4), Inches(0.55),
           name, size=22, bold=True, color=FERN)
        tb(slide, left + Inches(1.3), top + Inches(1.0), Inches(4.4), Inches(0.75),
           detail, size=14, color=DARK)
    hint(slide, "Read all four out loud. Knowing help is coming makes him braver.",
         6.45)


def s05_valley_scene():
    slide, n = new_slide("🌋 Picture Talk — Dinosaur Valley", "WARM-UP", "0–7 min",
                         "Dinosaur Valley", LAVA)
    one_task(slide, "Just look and talk. No reading yet!", LAVA)
    for i, (emoji, label) in enumerate(VALLEY_SCENE):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_LAVA)
        tb(slide, left, top + Inches(0.22), Inches(3.9), Inches(0.98), emoji, size=48,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.9), top + Inches(1.34), Inches(2.1),
                  Inches(0.58), WHITE)
        tb(slide, left + Inches(0.9), top + Inches(1.45), Inches(2.1), Inches(0.42),
           label, size=16, bold=True, color=LAVA, align=PP_ALIGN.CENTER)
    hint(slide, "Point at whatever he names. Pointing back shows you are listening.",
         6.42, "🗣️ SPEAKING", L_LAVA, LAVA)


def s06_talk():
    slide, n = new_slide("🗣️ Talk About the Valley", "SPEAKING", "0–7 min",
                         "Dinosaur Valley", LAVA)
    one_task(slide, "Answer out loud. Try a whole sentence.", LAVA)
    for i, (icon, question) in enumerate(TALK_QS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.35)
        add_round(slide, left, top, Inches(5.95), Inches(1.15), L_LAVA)
        add_oval(slide, left + Inches(0.3), top + Inches(0.28), Inches(0.58),
                 Inches(0.58), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.34), Inches(0.58), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.33), Inches(4.6), Inches(0.52),
           question, size=18, bold=True, color=INK)
    for i, starter in enumerate(STARTERS):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(4.75), Inches(5.95), Inches(1.1), L_AMBER)
        tb(slide, left + Inches(0.4), Inches(5.02), Inches(5.2), Inches(0.6), starter,
           size=26, bold=True, color=INK)
    hint(slide, "If he answers with one word, say it back as a full sentence first.",
         6.45)


def s07_sounds():
    slide, n = new_slide("🔊 Beginning Sounds — Listen and Say", "LEARN", "7–17 min",
                         "Dinosaur Valley", LAVA)
    one_task(slide, "I say the sound. You say it back to me.", LAVA)
    for i, (emoji, word, letter, sound) in enumerate(SOUND_CARDS):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.95 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_LAVA)
        tb(slide, left, top + Inches(0.14), Inches(3.9), Inches(0.7), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.84), Inches(3.6), Inches(0.48),
           word, size=18, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Arial Black")
        add_round(slide, left + Inches(0.85), top + Inches(1.36), Inches(2.2),
                  Inches(0.6), WHITE)
        tb(slide, left + Inches(0.85), top + Inches(1.46), Inches(2.2), Inches(0.44),
           f"{letter}  says  {sound}", size=15, bold=True, color=LAVA,
           align=PP_ALIGN.CENTER)
    hint(slide, "Say the sound, not the letter name. Stretch it: ddddd.", 6.45)


def s08_first_a():
    slide, n = new_slide("👣 Game: Follow the First Sound", "GAME", "7–17 min",
                         "Dinosaur Valley", LAVA)
    one_task(slide, "Which picture STARTS with the letter on the footprint?", LAVA)
    first_sound_rows(slide, FIRST_A, 1)
    hint(slide, "Say each picture name out loud first. Stretch the first sound.", 6.45)


def s09_first_b():
    slide, n = new_slide("👣 Follow the First Sound — Steps 4 to 6", "GAME", "7–17 min",
                         "Dinosaur Valley", LAVA)
    one_task(slide, "Three more steps. The dino moves with every right answer.", LAVA)
    first_sound_rows(slide, FIRST_B, 4)
    hint(slide, "Too many choices? Cover one with your hand and leave only two.", 6.45)


def s10_first_c():
    slide, n = new_slide("👣 First Sound Challenge", "GAME", "7–17 min",
                         "Dinosaur Valley", LAVA)
    one_task(slide, "Two last steps. Then the dino reaches the valley edge!", LAVA)
    first_sound_rows(slide, FIRST_C, 7)
    add_round(slide, Inches(0.5), Inches(4.95), Inches(12.35), Inches(1.28), L_JUNGLE)
    tb(slide, Inches(0.8), Inches(5.08), Inches(11.7), Inches(0.38),
       "⭐ DINO CHALLENGE", size=13, bold=True, color=JUNGLE)
    tb(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.6),
       "Can you think of one more word that starts with /d/?  How about /b/?", size=18,
       bold=True, color=INK)
    hint(slide, "Any real word counts, even one with nothing to do with dinosaurs.",
         6.42)


def s11_ending():
    slide, n = new_slide("🔚 Ending Sounds — What Do You Hear Last?", "LEARN",
                         "7–17 min", "Dinosaur Valley", FERN)
    one_task(slide, "Say the word. Now listen only to the LAST sound.", FERN)
    ending_rows(slide, ENDING, 1)
    hint(slide, "Whisper the word, then say the last sound loudly: ba-G!", 6.42)


def s12_valley_done():
    slide, n = new_slide("🌋 Sound Check — Dinosaur Valley Complete", "RECAP",
                         "7–17 min", "Dinosaur Valley", JUNGLE)
    one_task(slide, "Read all six sounds with me, fast.", JUNGLE)
    for i, (emoji, word, letter, sound) in enumerate(SOUND_CARDS):
        left = Inches(0.5 + i * 2.08)
        add_round(slide, left, Inches(2.0), Inches(1.9), Inches(2.6), L_JUNGLE)
        tb(slide, left, Inches(2.2), Inches(1.9), Inches(0.62), emoji, size=24,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(2.9), Inches(1.3), Inches(0.72),
                  JUNGLE)
        tb(slide, left + Inches(0.3), Inches(3.02), Inches(1.3), Inches(0.5), sound,
           size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.05), Inches(3.78), Inches(1.8), Inches(0.5), word,
           size=11, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(4.9), Inches(12.35), Inches(1.35), L_AMBER)
    tb(slide, Inches(0.8), Inches(5.05), Inches(11.7), Inches(0.4),
       "🥚 DINO EGG EARNED — Dinosaur Valley", size=16, bold=True, color=SAND)
    tb(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.6),
       "You can hear the first and last sound in a word. Next stop: Egg Cave, where "
       "words hide inside eggs.", size=15, color=INK)
    hint(slide, "Go left to right, two seconds per sound. Keep it quick and fun.", 6.42)


def s13_vowels():
    slide, n = new_slide("🥚 Short Vowels — Four Dino Eggs", "LEARN", "17–27 min",
                         "Egg Cave", PLUM)
    one_task(slide, "Every word has a sound hiding in the middle.", PLUM)
    for i, (letter, emoji, word, color, light) in enumerate(VOWEL_EGGS):
        left = Inches(0.9 + i * 3.0)
        add_round(slide, left, Inches(1.95), Inches(2.75), Inches(4.1), light)
        tb(slide, left, Inches(2.15), Inches(2.75), Inches(0.72), "🥚", size=30,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.7), Inches(2.9), Inches(1.35), Inches(1.15),
                  color)
        tb(slide, left + Inches(0.7), Inches(3.05), Inches(1.35), Inches(0.85), letter,
           size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left, Inches(4.2), Inches(2.75), Inches(0.72), emoji, size=28,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.5), Inches(5.0), Inches(1.75), Inches(0.62),
                  WHITE)
        tb(slide, left + Inches(0.5), Inches(5.11), Inches(1.75), Inches(0.44), word,
           size=18, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Stretch the middle: m-aaa-p. That is the sound he is listening for.",
         6.42)


def s14_egg_a():
    slide, n = new_slide("🥚 Game: What's Inside the Egg?", "GAME", "17–27 min",
                         "Egg Cave", PLUM)
    one_task(slide, "Read the word. Then crack it open — which picture is it?", PLUM)
    egg_rows(slide, EGG_A, 1)
    hint(slide, "Sound out every letter before he looks at the pictures.", 6.42)


def s15_egg_b():
    slide, n = new_slide("🥚 What's Inside the Egg? — Eggs 4 to 6", "GAME", "17–27 min",
                         "Egg Cave", PLUM)
    one_task(slide, "Three more eggs to crack.", PLUM)
    egg_rows(slide, EGG_B, 4)
    hint(slide, "Cover the pictures with your hand until he has read the word.", 6.42)


def s16_egg_bonus():
    slide, n = new_slide("⭐ Egg Cave Bonus — Two Golden Eggs", "BONUS", "17–27 min",
                         "Egg Cave", AMBER)
    one_task(slide, "Bonus round! These two eggs are golden.", AMBER)
    egg_rows(slide, EGG_BONUS, 7)
    hint(slide, "Skip this if he is tired. It is a bonus, not a requirement.", 6.42)


def s17_dino_words():
    slide, n = new_slide("🦖 Dino Words We Will Read", "VOCABULARY", "17–27 min",
                         "Egg Cave", FERN)
    one_task(slide, "Every word today comes from this list. Nothing new later.", FERN)
    for i, (word, emoji) in enumerate(DINO_WORDS):
        col, row = i % 6, i // 6
        left = Inches(0.5 + col * 2.08)
        top = Inches(2.0 + row * 1.45)
        add_round(slide, left, top, Inches(1.9), Inches(1.32), L_FERN)
        tb(slide, left, top + Inches(0.12), Inches(1.9), Inches(0.58), emoji, size=22,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), top + Inches(0.72), Inches(1.6),
                  Inches(0.5), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(0.8), Inches(1.6), Inches(0.36),
           word, size=16, bold=True, color=FERN, align=PP_ALIGN.CENTER)
    hint(slide, "Read them across in a rhythm. He joins in wherever he can.", 6.42)


def s18_cave_done():
    slide, n = new_slide("🥚 Egg Cave Complete", "RECAP", "17–27 min", "Egg Cave",
                         JUNGLE)
    one_task(slide, "Look what was hiding inside those eggs.", JUNGLE)
    done = [("🥚", "A", "map"), ("🥚", "E", "bed"), ("🥚", "I", "sit"),
            ("🥚", "O", "log")]
    for i, (emoji, letter, word) in enumerate(done):
        left = Inches(0.9 + i * 3.0)
        add_round(slide, left, Inches(1.95), Inches(2.75), Inches(2.5), L_JUNGLE)
        tb(slide, left, Inches(2.1), Inches(2.75), Inches(0.62), emoji, size=24,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.7), Inches(2.78), Inches(1.35), Inches(0.8),
                  JUNGLE)
        tb(slide, left + Inches(0.7), Inches(2.9), Inches(1.35), Inches(0.58), letter,
           size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.5), Inches(3.72), Inches(1.75), Inches(0.58),
                  WHITE)
        tb(slide, left + Inches(0.5), Inches(3.82), Inches(1.75), Inches(0.42), word,
           size=18, bold=True, color=JUNGLE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(4.75), Inches(12.35), Inches(1.4), L_AMBER)
    tb(slide, Inches(0.8), Inches(4.9), Inches(11.7), Inches(0.4),
       "🥚 DINO EGG EARNED — Egg Cave", size=16, bold=True, color=SAND)
    tb(slide, Inches(0.8), Inches(5.35), Inches(11.7), Inches(0.65),
       "You read whole words by yourself. Next stop: Dino Camp, where we build words "
       "out of footprints.", size=15, color=INK)
    hint(slide, "Say all four words once more, quickly, before moving on.", 6.42)


def s19_six_steps():
    slide, n = new_slide("🦖 Our Six Steps for a Hard Word", "LEARN", "27–37 min",
                         "Dino Camp", AMBER)
    one_task(slide, "When a word looks hard, we always do these six steps.", AMBER)
    for i, (num, label, detail, example, color, light) in enumerate(SIX_STEPS):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.95 + row * 2.25)
        add_round(slide, left, top, Inches(3.9), Inches(2.05), light)
        add_oval(slide, left + Inches(0.25), top + Inches(0.24), Inches(0.62),
                 Inches(0.62), color)
        tb(slide, left + Inches(0.25), top + Inches(0.33), Inches(0.62), Inches(0.44),
           num, size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.05), top + Inches(0.26), Inches(2.7), Inches(0.5),
           label, size=19, bold=True, color=color, font="Arial Black")
        tb(slide, left + Inches(0.3), top + Inches(0.95), Inches(3.3), Inches(0.45),
           detail, size=12, color=DARK)
        add_round(slide, left + Inches(0.3), top + Inches(1.38), Inches(3.3),
                  Inches(0.55), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(1.49), Inches(3.3), Inches(0.4),
           example, size=16, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Never skip step 6. Using the word is what makes it stick.", 6.45)


def s20_blend_a():
    slide, n = new_slide("👣 Game: Dino Footprint Blend", "GAME", "27–37 min",
                         "Dino Camp", AMBER)
    one_task(slide, "Say each sound. Then say the whole word fast.", AMBER)
    blend_rows(slide, BLEND_A)
    hint(slide, "Cover the last letter so he blends only two sounds first.", 6.5)


def s21_blend_b():
    slide, n = new_slide("👣 Footprint Blend — Three More Steps", "GAME", "27–37 min",
                         "Dino Camp", AMBER)
    one_task(slide, "The dino moves one step for every word you blend.", AMBER)
    blend_rows(slide, BLEND_B)
    hint(slide, "Slide your finger under the letters, then sweep it fast.", 6.5)


def s22_blend_c():
    slide, n = new_slide("👣 Footprint Blend — The Last Three", "GAME", "27–37 min",
                         "Dino Camp", AMBER)
    one_task(slide, "Three left. Then the dino is across the camp!", AMBER)
    blend_rows(slide, BLEND_C)
    hint(slide, "These all have the U and O sounds. Stretch the middle.", 6.5)


def s23_trick():
    slide, n = new_slide("🦖 Game: Which Word Is Right?", "GAME", "27–37 min",
                         "Dino Camp", AMBER)
    one_task(slide, "Only ONE word matches the picture. Read all three first.", AMBER)
    trick_rows(slide, TRICK, 1)
    hint(slide, "The words look almost the same. Check the last letter carefully.",
         6.42)


def s24_no_support():
    slide, n = new_slide("🦕 No Footprints — Just Read It", "PRACTICE", "27–37 min",
                         "Dino Camp", JUNGLE)
    one_task(slide, "This time there are no letter cards. You can do it!", JUNGLE)
    for i, (word, emoji) in enumerate(NO_SUPPORT):
        col, row = i % 3, i // 3
        left = Inches(0.6 + col * 4.1)
        top = Inches(1.95 + row * 2.25)
        add_round(slide, left, top, Inches(3.85), Inches(2.05), L_JUNGLE)
        tb(slide, left, top + Inches(0.18), Inches(3.85), Inches(0.72), emoji, size=28,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.7), top + Inches(1.0), Inches(2.45),
                  Inches(0.82), WHITE)
        tb(slide, left + Inches(0.7), top + Inches(1.12), Inches(2.45), Inches(0.6),
           word, size=32, bold=True, color=JUNGLE, align=PP_ALIGN.CENTER,
           font="Arial Black")
    hint(slide, "If he stalls, put your finger under the first letter. Say nothing.",
         6.45)


def s25_fossil_how():
    slide, n = new_slide("🦴 Fossil Word Match — How It Works", "LEARN", "37–44 min",
                         "Dino Camp", SAND)
    one_task(slide, "A fossil holds a word. Find the picture that goes with it.", SAND)
    for i, (num, emoji, label, detail, color) in enumerate(FOSSIL_HOW):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(3.3), L_SAND)
        add_oval(slide, left + Inches(1.55), Inches(2.2), Inches(0.8), Inches(0.8),
                 color)
        tb(slide, left + Inches(1.55), Inches(2.34), Inches(0.8), Inches(0.52), num,
           size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(3.15), Inches(3.9), Inches(0.7), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.92), Inches(3.5), Inches(0.5), label,
           size=16, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.3), Inches(4.48), Inches(3.3), Inches(0.6), detail,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.45), Inches(12.35), Inches(0.82), L_AMBER)
    tb(slide, Inches(0.8), Inches(5.66), Inches(11.7), Inches(0.44),
       "Reading the word FIRST is the whole game. The picture is the reward, not the "
       "clue.", size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Read the fossil word yourself once, then hand it over to him.", 6.42)


def s26_fossil_a():
    slide, n = new_slide("🦴 Game: Find the Fossil Match", "GAME", "37–44 min",
                         "Dino Camp", SAND)
    one_task(slide, "Read the fossil. Point to the picture that matches.", SAND)
    match_rows(slide, MATCH_A, 1)
    hint(slide, "Read the word, then say it again while looking at the pictures.", 6.42)


def s27_fossil_b():
    slide, n = new_slide("🦴 Find the Fossil Match — Fossils 4 to 6", "GAME",
                         "37–44 min", "Dino Camp", SAND)
    one_task(slide, "Three more fossils for the museum.", SAND)
    match_rows(slide, MATCH_B, 4)
    hint(slide, "Wrong pick? Say \"nearly\" and ask him to read the word once more.",
         6.42)


def s28_fossil_challenge():
    slide, n = new_slide("⭐ Fossil Challenge — Name the Dino Parts", "BONUS",
                         "37–44 min", "Dino Camp", PLUM)
    one_task(slide, "Four museum fossils. Read each one out loud.", PLUM)
    for i, (emoji, word) in enumerate(FOSSIL_BONUS):
        left = Inches(0.6 + i * 3.12)
        add_round(slide, left, Inches(2.0), Inches(2.9), Inches(3.2), L_PLUM)
        tb(slide, left, Inches(2.3), Inches(2.9), Inches(0.9), emoji, size=40,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(3.4), Inches(2.5), Inches(0.75),
                  WHITE)
        tb(slide, left + Inches(0.2), Inches(3.52), Inches(2.5), Inches(0.55), word,
           size=19, bold=True, color=PLUM, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.4), Inches(4.35), Inches(2.1), Inches(0.55),
                  WHITE)
        tb(slide, left + Inches(0.4), Inches(4.45), Inches(2.1), Inches(0.38),
           "say it · use it", size=11, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.45), Inches(12.35), Inches(0.82), L_AMBER)
    tb(slide, Inches(0.8), Inches(5.66), Inches(11.7), Inches(0.44),
       "⭐ DINO CHALLENGE:  Use one of these words in a sentence about your dino.",
       size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "FOOTPRINT is long. Cover half: foot + print. Then join them.", 6.42)


def s29_dino_says_intro():
    slide, n = new_slide("🧠 Brain Break — Dino Says", "BREAK", "44–49 min",
                         "Dino Camp", FERN)
    one_task(slide, "Stay in your seat. Only move when Dino says!", FERN)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.6), Inches(4.4), L_FERN)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.6), Inches(1.7), "🦖", size=96,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.6), Inches(0.7), "DINO SAYS",
       size=34, bold=True, color=FERN, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.8), Inches(5.3), Inches(5.0), Inches(0.6), "5 minutes only",
       size=15, color=SOFT, align=PP_ALIGN.CENTER)
    rules = [("✅", "\"Dino says clap\" → you clap."),
             ("🚫", "Just \"clap\" → you do nothing."),
             ("🪑", "Stay beside your chair the whole time."),
             ("📖", "Some commands are reading commands too.")]
    for i, (icon, line) in enumerate(rules):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.45), top, Inches(6.4), Inches(0.96), L_AMBER)
        add_oval(slide, Inches(6.7), top + Inches(0.2), Inches(0.56), Inches(0.56),
                 WHITE)
        tb(slide, Inches(6.7), top + Inches(0.26), Inches(0.56), Inches(0.42), icon,
           size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.5), top + Inches(0.26), Inches(5.1), Inches(0.48), line,
           size=15, bold=True, color=INK)
    hint(slide, "Let him be the caller for the last two rounds. He will love it.",
         6.45)


def s30_dino_says():
    slide, n = new_slide("🦖 Dino Says — Six Commands", "BREAK", "44–49 min",
                         "Dino Camp", FERN)
    one_task(slide, "Listen carefully. Did Dino really say it?", FERN)
    for i, (icon, line) in enumerate(DINO_SAYS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.5)
        add_round(slide, left, top, Inches(5.95), Inches(1.3), L_FERN)
        add_oval(slide, left + Inches(0.3), top + Inches(0.34), Inches(0.62),
                 Inches(0.62), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(0.62), Inches(0.46),
           icon, size=17, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.15), top + Inches(0.4), Inches(4.6), Inches(0.6),
           line, size=16, bold=True, color=INK)
    hint(slide, "Mix in one command WITHOUT \"Dino says\" to keep him listening.",
         6.45)


def s31_families():
    slide, n = new_slide("👨‍👩‍👧 Dino Word Families", "LEARN", "49–59 min",
                         "Jungle Trail", JUNGLE)
    one_task(slide, "Same ending sound = word family friends.", JUNGLE)
    for i, (family, words, color, light) in enumerate(FAMILIES):
        left = Inches(0.6 + i * 3.12)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(3.9), light)
        add_round(slide, left + Inches(0.5), Inches(2.2), Inches(1.9), Inches(0.8),
                  color)
        tb(slide, left + Inches(0.5), Inches(2.33), Inches(1.9), Inches(0.58), family,
           size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, word in enumerate(words):
            top = Inches(3.18 + j * 0.64)
            add_round(slide, left + Inches(0.45), top, Inches(2.0), Inches(0.54), WHITE)
            tb(slide, left + Inches(0.45), top + Inches(0.07), Inches(2.0), Inches(0.4),
               word, size=19, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Read down each column. The ending never changes — only the first "
                "sound does.", 6.42)


def s32_family_a():
    slide, n = new_slide("🦕 Game: Who Is in the Dino Family?", "GAME", "49–59 min",
                         "Jungle Trail", JUNGLE)
    one_task(slide, "Three of these four belong. Which one does not?", JUNGLE)
    family_rows(slide, FAMILY_A, 1)
    hint(slide, "Cover the first letter. Now only the ending is showing.", 6.42)


def s33_family_b():
    slide, n = new_slide("🦕 Who Is in the Dino Family? — Rounds 3 and 4", "GAME",
                         "49–59 min", "Jungle Trail", JUNGLE)
    one_task(slide, "Two more families. Read every word out loud first.", JUNGLE)
    family_rows(slide, FAMILY_B, 3)
    hint(slide, "Read the family ending out loud before each round: -un, -op.", 6.42)


def s34_new_words():
    slide, n = new_slide("🧩 Make a Brand New Dino Word", "GAME", "49–59 min",
                         "Jungle Trail", AMBER)
    one_task(slide, "Add one letter to the front. What word appears?", AMBER)
    for i, (letter, family, word, emoji, color, light) in enumerate(NEW_WORDS):
        top = Inches(1.9 + i * 1.14)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.02), light)
        add_round(slide, Inches(1.0), top + Inches(0.2), Inches(1.3), Inches(0.68),
                  WHITE)
        tb(slide, Inches(1.0), top + Inches(0.28), Inches(1.3), Inches(0.5), letter,
           size=26, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(2.4), top + Inches(0.3), Inches(0.4), Inches(0.5), "+",
           size=20, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.95), top + Inches(0.2), Inches(1.7), Inches(0.68),
                  WHITE)
        tb(slide, Inches(2.95), top + Inches(0.28), Inches(1.7), Inches(0.5), family,
           size=24, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(4.8), top + Inches(0.26), Inches(0.6), Inches(0.5), "→",
           size=22, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(5.6), top + Inches(0.2), Inches(2.6), Inches(0.68),
                  WHITE)
        tb(slide, Inches(5.6), top + Inches(0.28), Inches(2.6), Inches(0.5), word,
           size=26, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(8.5), top + Inches(0.22), Inches(0.9), Inches(0.66), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, Inches(9.6), top + Inches(0.32), Inches(3.0), Inches(0.44),
           "read it · then use it", size=13, color=SOFT)
    hint(slide, "He does not need to know the word first. Blending will find it.",
         6.42)


def s35_sight_1():
    slide, n = new_slide("🗺️ Sight Words — Group 1", "LEARN", "59–67 min",
                         "Jungle Trail", STONE)
    one_task(slide, "These three words we do NOT sound out. We just know them.", STONE)
    for i, (word, phrase, emoji) in enumerate(SIGHT_1):
        left = Inches(0.7 + i * 4.1)
        add_round(slide, left, Inches(1.95), Inches(3.8), Inches(3.9), L_STONE)
        add_round(slide, left + Inches(0.55), Inches(2.25), Inches(2.7), Inches(1.2),
                  STONE)
        tb(slide, left + Inches(0.55), Inches(2.47), Inches(2.7), Inches(0.8), word,
           size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left, Inches(3.65), Inches(3.8), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.6), Inches(3.2), Inches(0.8),
                  WHITE)
        tb(slide, left + Inches(0.3), Inches(4.78), Inches(3.2), Inches(0.5), phrase,
           size=19, bold=True, color=STONE, align=PP_ALIGN.CENTER)
    hint(slide, "Say the word, he repeats, then he finds it in the phrase below.",
         6.42)


def s36_sight_2():
    slide, n = new_slide("🗺️ Sight Words — Group 2", "LEARN", "59–67 min",
                         "Jungle Trail", STONE)
    one_task(slide, "Three more. Only three at a time — never a long list.", STONE)
    for i, (word, phrase, emoji) in enumerate(SIGHT_2):
        left = Inches(0.7 + i * 4.1)
        add_round(slide, left, Inches(1.95), Inches(3.8), Inches(3.9), L_STONE)
        add_round(slide, left + Inches(0.55), Inches(2.25), Inches(2.7), Inches(1.2),
                  STONE)
        tb(slide, left + Inches(0.55), Inches(2.47), Inches(2.7), Inches(0.8), word,
           size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left, Inches(3.65), Inches(3.8), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.6), Inches(3.2), Inches(0.8),
                  WHITE)
        tb(slide, left + Inches(0.3), Inches(4.78), Inches(3.2), Inches(0.5), phrase,
           size=19, bold=True, color=STONE, align=PP_ALIGN.CENTER)
    hint(slide, "Go back to Group 1 for ten seconds before you start Group 2.", 6.42)


def s37_path_a():
    slide, n = new_slide("👣 Game: Dino Sight-Word Path", "GAME", "59–67 min",
                         "Jungle Trail", STONE)
    one_task(slide, "Read the word to move the dino one footprint forward.", STONE)
    path_rows(slide, PATH_A, 1)
    hint(slide, "Read all three footprints out loud. Listening beats looking.", 6.42)


def s38_path_b():
    slide, n = new_slide("👣 Sight-Word Path — Footprints 5 to 8", "GAME", "59–67 min",
                         "Jungle Trail", STONE)
    one_task(slide, "Four more footprints and the dino reaches the jungle.", STONE)
    path_rows(slide, PATH_B, 5)
    hint(slide, "Seen this word already today? Say so. Repetition is the point.",
         6.42)


def s39_phrases():
    slide, n = new_slide("🔗 Two and Three Words Together", "READ", "59–67 min",
                         "Jungle Trail", JUNGLE)
    one_task(slide, "Not a whole sentence yet. Just a few words joined up.", JUNGLE)
    phrase_grid(slide, PHRASES)
    hint(slide, "Sweep your finger under the whole phrase so it sounds like talking.",
         6.6)


def s40_read_method():
    slide, n = new_slide("📖 How We Read Every Sentence", "LEARN", "67–75 min",
                         "Jungle Trail", LAVA)
    one_task(slide, "Three turns for every sentence. You are always third.", LAVA)
    for i, (step, who, detail, color) in enumerate(READ_STEPS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(3.4),
                  [L_LAVA, L_AMBER, L_JUNGLE][i])
        add_round(slide, left + Inches(1.05), Inches(2.2), Inches(1.8), Inches(0.62),
                  color)
        tb(slide, left + Inches(1.05), Inches(2.3), Inches(1.8), Inches(0.44), step,
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.05), Inches(3.5), Inches(0.75), who,
           size=30, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.35), Inches(3.95), Inches(3.2), Inches(0.75), detail,
           size=13, color=DARK, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(4.75), Inches(3.9), Inches(0.5),
           ["🧑‍🏫", "🧑‍🏫 🧒", "🧒"][i], size=22, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.55), Inches(12.35), Inches(0.72), L_STONE)
    tb(slide, Inches(0.8), Inches(5.72), Inches(11.7), Inches(0.42),
       "If step 3 is hard, go back to step 2. Going back is allowed, always.",
       size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Never let him attempt a sentence cold. Model it every single time.",
         6.42)


def s41_sentences_a():
    slide, n = new_slide("📕 Our First Dino Sentences", "READ", "67–75 min",
                         "Jungle Trail", LAVA)
    one_task(slide, "Four words. Point to each word as you read it.", LAVA)
    sentence_rows(slide, SENTENCES_A)
    i_we_you(slide)


def s42_sentences_b():
    slide, n = new_slide("📕 Three More Dino Sentences", "READ", "67–75 min",
                         "Jungle Trail", FERN)
    one_task(slide, "Same words, new order. You already know every one.", FERN)
    sentence_rows(slide, SENTENCES_B, top_start=1.95, gap=1.5, height=1.34, size=32)
    i_we_you(slide)


def s43_puzzle_a():
    slide, n = new_slide("🧩 Game: Build the Dino Sentence", "GAME", "67–75 min",
                         "Jungle Trail", AMBER)
    one_task(slide, "The words got mixed up. Put them back in order!", AMBER)
    puzzle_rows(slide, PUZZLES_A, 1)
    hint(slide, "Hint: the first word always wears a CAPITAL letter.", 6.42)


def s44_puzzle_b():
    slide, n = new_slide("🧩 Build the Dino Sentence — Puzzles 3 and 4", "GAME",
                         "67–75 min", "Jungle Trail", AMBER)
    one_task(slide, "Two more puzzles. Read it out loud to check it.", AMBER)
    puzzle_rows(slide, PUZZLES_B, 3)
    hint(slide, "Wrong order? Read it aloud. His ear will catch what his eye missed.",
         6.42)


def s45_own_sentence():
    slide, n = new_slide("⭐ Make Your Own Dino Sentence", "BONUS", "67–75 min",
                         "Jungle Trail", PLUM)
    one_task(slide, "Pick a picture. Now say a sentence about it.", PLUM)
    for i, (emoji, word) in enumerate(OWN_SENTENCE):
        left = Inches(0.6 + i * 3.12)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(2.45), L_PLUM)
        tb(slide, left, Inches(2.2), Inches(2.9), Inches(0.9), emoji, size=40,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.5), Inches(3.35), Inches(1.9), Inches(0.68),
                  WHITE)
        tb(slide, left + Inches(0.5), Inches(3.46), Inches(1.9), Inches(0.48), word,
           size=20, bold=True, color=PLUM, align=PP_ALIGN.CENTER)
    for i, starter in enumerate(["I see a ______.", "The dino can ______."]):
        left = Inches(0.6 + i * 6.35)
        add_round(slide, left, Inches(4.7), Inches(6.05), Inches(1.15), L_AMBER)
        tb(slide, left + Inches(0.45), Inches(4.98), Inches(5.3), Inches(0.62),
           starter, size=26, bold=True, color=INK)
    hint(slide, "Write his sentence down and read it back to him. That is powerful.",
         6.45)


def s46_story_intro():
    slide, n = new_slide("📖 Story Time — Six Words to Watch For", "STORY", "75–83 min",
                         "Fossil Museum", STONE)
    one_task(slide, "You already read all six of these today.", STONE)
    for i, (word, emoji) in enumerate(STORY_WATCH):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.0)
        add_round(slide, left, top, Inches(3.9), Inches(1.8), L_STONE)
        tb(slide, left + Inches(0.3), top + Inches(0.45), Inches(1.1), Inches(0.9),
           emoji, size=36, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.6), top + Inches(0.55), Inches(2.0),
                  Inches(0.7), WHITE)
        tb(slide, left + Inches(1.6), top + Inches(0.67), Inches(2.0), Inches(0.5),
           word, size=22, bold=True, color=STONE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.92), Inches(12.35), Inches(0.42), L_AMBER)
    tb(slide, Inches(0.8), Inches(5.97), Inches(11.7), Inches(0.34),
       "\"THE LITTLE DINO\"  ·  four short parts  ·  we read each part three times",
       size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Read the six words together before page one. No cold surprises.",
         6.45)


def s47_story_1():
    story_slide(STORY[0], "75–83 min")


def s48_story_2():
    story_slide(STORY[1], "75–83 min")


def s49_story_3():
    story_slide(STORY[2], "75–83 min")


def s50_story_4():
    story_slide(STORY[3], "75–83 min")


def s51_story_again():
    slide, n = new_slide("🔁 Read These Six Lines Again", "READ", "75–83 min",
                         "Fossil Museum", JUNGLE)
    one_task(slide, "The six easiest lines of the story. All yours this time.", JUNGLE)
    for i, line in enumerate(EASY_LINES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.5)
        add_round(slide, left, top, Inches(5.95), Inches(1.3), L_JUNGLE)
        add_oval(slide, left + Inches(0.3), top + Inches(0.4), Inches(0.5), Inches(0.5),
                 JUNGLE)
        tb(slide, left + Inches(0.3), top + Inches(0.46), Inches(0.5), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.05), top + Inches(0.38), Inches(4.7), Inches(0.6),
           line, size=21, bold=True, color=INK)
    hint(slide, "Time him gently the second time. Faster reading feels like winning.",
         6.6)


def s52_q_a():
    slide, n = new_slide("🔎 Game: Dino Story Detective", "COMPREHENSION", "83–87 min",
                         "Fossil Museum", STONE)
    one_task(slide, "Think about the story. Which answer is right?", STONE)
    question_rows(slide, QUESTIONS_A, 1)
    hint(slide, "Cannot remember? Go back a slide and read it together.", 6.42)


def s53_q_b():
    slide, n = new_slide("🔎 Story Detective — Questions 3 and 4", "COMPREHENSION",
                         "83–87 min", "Fossil Museum", STONE)
    one_task(slide, "Two more. Then show me where you found it.", STONE)
    question_rows(slide, QUESTIONS_B, 3)
    hint(slide, "Answer first, proof second. Both count as reading.", 6.42)


def s54_q_c():
    slide, n = new_slide("🔎 Story Detective — The Last Two", "COMPREHENSION",
                         "83–87 min", "Fossil Museum", STONE)
    one_task(slide, "Last two questions of the whole mission.", STONE)
    question_rows(slide, QUESTIONS_C, 5)
    hint(slide, "Going back to the page is not cheating. It is what good readers do.",
         6.42)


def s55_museum():
    slide, n = new_slide("🏛️ You Reached the Fossil Museum", "FINAL", "87–90 min",
                         "Fossil Museum", AMBER)
    one_task(slide, "Five fossils left. Then you are a champion.", AMBER)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.6), Inches(4.4), L_AMBER)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.6), Inches(1.8), "🏛️", size=96,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.6), Inches(5.6), Inches(0.7), "FOSSIL MUSEUM",
       size=28, bold=True, color=SAND, align=PP_ALIGN.CENTER, font="Arial Black")
    earned = [("🌋", "Dinosaur Valley", "first and last sounds"),
              ("🥚", "Egg Cave", "short vowels and whole words"),
              ("🦖", "Dino Camp", "blending and fossil matching"),
              ("🌿", "Jungle Trail", "families, sight words, sentences")]
    for i, (emoji, place, what) in enumerate(earned):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.45), top, Inches(6.4), Inches(0.96), L_JUNGLE)
        tb(slide, Inches(6.7), top + Inches(0.22), Inches(0.6), Inches(0.5), emoji,
           size=18, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.45), top + Inches(0.12), Inches(5.2), Inches(0.42), place,
           size=16, bold=True, color=JUNGLE)
        tb(slide, Inches(7.45), top + Inches(0.52), Inches(5.2), Inches(0.36), what,
           size=12, color=DARK)
    hint(slide, "Say the four places out loud with him. He will be surprised by the "
                "list.", 6.45)


def s56_final_a():
    slide, n = new_slide("🦴 Final Challenge — Fossils 1 to 3", "FINAL", "87–90 min",
                         "Fossil Museum", LAVA)
    one_task(slide, "No hints unless you ask for one. You've got this.", LAVA)
    for i, (num, emoji, label, prompt, answer, color, light) in enumerate(FINAL_A):
        top = Inches(1.88 + i * 1.55)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.38), light)
        add_oval(slide, Inches(0.85), top + Inches(0.44), Inches(0.58), Inches(0.58),
                 color)
        tb(slide, Inches(0.85), top + Inches(0.5), Inches(0.58), Inches(0.44), num,
           size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.7), top + Inches(0.4), Inches(0.9), Inches(0.68), emoji,
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.8), top + Inches(0.3), Inches(4.0), Inches(0.45), label,
           size=15, bold=True, color=color)
        tb(slide, Inches(2.8), top + Inches(0.78), Inches(4.0), Inches(0.42), prompt,
           size=13, color=SOFT)
        add_round(slide, Inches(7.2), top + Inches(0.3), Inches(5.35), Inches(0.82),
                  WHITE)
        tb(slide, Inches(7.2), top + Inches(0.42), Inches(5.35), Inches(0.6), answer,
           size=30, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
    hint(slide, "Wait a full ten seconds before offering anything. Silence helps.",
         6.45)


def s57_final_b():
    slide, n = new_slide("🦴 Final Challenge — Fossils 4 and 5", "FINAL", "87–90 min",
                         "Fossil Museum", JUNGLE)
    one_task(slide, "The last two fossils in the whole museum.", JUNGLE)
    for i, (num, emoji, label, prompt, answer, color, light) in enumerate(FINAL_B):
        top = Inches(2.0 + i * 2.1)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.85), light)
        add_oval(slide, Inches(0.85), top + Inches(0.64), Inches(0.62), Inches(0.62),
                 color)
        tb(slide, Inches(0.85), top + Inches(0.72), Inches(0.62), Inches(0.46), num,
           size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.75), top + Inches(0.58), Inches(1.0), Inches(0.75), emoji,
           size=32, align=PP_ALIGN.CENTER)
        tb(slide, Inches(3.0), top + Inches(0.45), Inches(3.4), Inches(0.5), label,
           size=16, bold=True, color=color)
        tb(slide, Inches(3.0), top + Inches(0.98), Inches(3.4), Inches(0.45), prompt,
           size=13, color=SOFT)
        add_round(slide, Inches(6.7), top + Inches(0.45), Inches(5.85), Inches(0.95),
                  WHITE)
        tb(slide, Inches(6.7), top + Inches(0.62), Inches(5.85), Inches(0.62), answer,
           size=26, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "For fossil 5 any true answer counts: run, dig, hop.", 6.42)


def s58_badge():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), AMBER)
    tb(slide, Inches(0.7), Inches(0.55), Inches(12), Inches(0.7),
       "🏆 DINO READING CHAMPION! 🏆", size=32, bold=True, color=AMBER,
       align=PP_ALIGN.CENTER, font="Georgia")
    add_oval(slide, Inches(5.42), Inches(1.6), Inches(2.5), Inches(2.5), JUNGLE)
    tb(slide, Inches(5.42), Inches(2.15), Inches(2.5), Inches(1.4), "🦕", size=52,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.35), Inches(12), Inches(0.75), "🎉 YOU DID IT!",
       size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    claims = ["You read words!", "You built sentences!", "You read a story!",
              "You solved the dinosaur mystery!"]
    for i, claim in enumerate(claims):
        left = Inches(0.75 + i * 3.1)
        add_round(slide, left, Inches(5.2), Inches(2.85), Inches(0.62),
                  RGBColor(0x1C, 0x39, 0x2A))
        tb(slide, left, Inches(5.31), Inches(2.85), Inches(0.44), claim, size=13,
           bold=True, color=RGBColor(0xBA, 0xD8, 0xC4), align=PP_ALIGN.CENTER)
    for i, (emoji, place, _j, _t, color, _l) in enumerate(STOPS):
        left = Inches(1.6 + i * 2.15)
        add_round(slide, left, Inches(6.05), Inches(1.95), Inches(0.62), color)
        tb(slide, left, Inches(6.16), Inches(1.95), Inches(0.42), f"{emoji} 🥚", size=13,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "87–90 min", "Fossil Museum")
    fade(slide)


def s59_can_read():
    slide, n = new_slide("✅ What I Can Read Now", "RECAP", "87–90 min",
                         "Fossil Museum", JUNGLE)
    one_task(slide, "Look at everything you did in 90 minutes.", JUNGLE)
    for i, (icon, label, detail) in enumerate(CAN_READ):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 0.97)
        add_round(slide, left, top, Inches(5.95), Inches(0.85), L_JUNGLE)
        add_oval(slide, left + Inches(0.24), top + Inches(0.18), Inches(0.5),
                 Inches(0.5), WHITE)
        tb(slide, left + Inches(0.24), top + Inches(0.24), Inches(0.5), Inches(0.4),
           icon, size=14, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.9), top + Inches(0.1), Inches(4.8), Inches(0.4),
           label, size=15, bold=True, color=JUNGLE)
        tb(slide, left + Inches(0.9), top + Inches(0.47), Inches(4.8), Inches(0.34),
           detail, size=11, color=DARK)


def s60_support():
    slide, n = new_slide("🔒 TEACHER ONLY — Support System", "TEACHER ONLY", "",
                         "Support", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.34),
       "Move between levels freely inside one activity. Never name the level out loud.",
       size=12, bold=True, color=LAVA)
    for i, (icon, name, what, how, color, light) in enumerate(SUPPORT_LEVELS):
        left = Inches(0.45 + i * 4.18)
        add_round(slide, left, Inches(1.72), Inches(3.95), Inches(2.3), light)
        tb(slide, left, Inches(1.86), Inches(3.95), Inches(0.46), icon, size=16,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(2.36), Inches(3.65), Inches(0.46), name,
           size=17, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.25), Inches(2.85), Inches(3.45), Inches(0.42), what,
           size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(3.3), Inches(3.55), Inches(0.55),
                  WHITE)
        tb(slide, left + Inches(0.35), Inches(3.41), Inches(3.25), Inches(0.38), how,
           size=10, color=SOFT, align=PP_ALIGN.CENTER, italic=True)
    add_round(slide, Inches(0.45), Inches(4.16), Inches(12.4), Inches(0.46), L_AMBER)
    tb(slide, Inches(0.7), Inches(4.23), Inches(11.9), Inches(0.34),
       "POSITIVE LABELS ONLY:   🟢 Dino Help   ·   🟡 Dino Mission   ·   "
       "⭐ Dino Challenge   ·   🦖 Dino Hint", size=12, bold=True, color=SAND)
    tb(slide, Inches(0.45), Inches(4.76), Inches(6.1), Inches(0.32),
       "GRADUATED SUPPORT — never reveal the answer first", size=12, bold=True,
       color=INK)
    for i, (step, text, color) in enumerate(HINT_LADDER):
        top = Inches(5.14 + i * 0.44)
        add_round(slide, Inches(0.45), top, Inches(6.1), Inches(0.38), L_GREY)
        add_round(slide, Inches(0.6), top + Inches(0.05), Inches(1.0), Inches(0.28),
                  color)
        tb(slide, Inches(0.6), top + Inches(0.06), Inches(1.0), Inches(0.26), step,
           size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.75), top + Inches(0.04), Inches(4.6), Inches(0.3), text,
           size=11, bold=True, color=INK)
    add_round(slide, Inches(6.75), Inches(4.76), Inches(6.1), Inches(1.0), L_FERN)
    tb(slide, Inches(7.0), Inches(4.84), Inches(5.6), Inches(0.3), "🦖 DINO HINT BANK",
       size=11, bold=True, color=FERN)
    bullets(slide, Inches(7.0), Inches(5.14), Inches(2.85), Inches(0.58), HINT_BANK[:3],
            size=9, sp=1)
    bullets(slide, Inches(9.9), Inches(5.14), Inches(2.9), Inches(0.58), HINT_BANK[3:],
            size=9, sp=1)
    add_round(slide, Inches(6.75), Inches(5.86), Inches(6.1), Inches(0.96), L_JUNGLE)
    tb(slide, Inches(7.0), Inches(5.93), Inches(5.6), Inches(0.3),
       "🗣️ SAY THIS INSTEAD OF \"WRONG\"", size=11, bold=True, color=JUNGLE)
    bullets(slide, Inches(7.0), Inches(6.22), Inches(3.0), Inches(0.58), PRAISE[:3],
            size=9, sp=1)
    bullets(slide, Inches(10.0), Inches(6.22), Inches(2.8), Inches(0.58), PRAISE[3:],
            size=9, sp=1)


def s61_assessment():
    slide, n = new_slide("🔒 TEACHER ONLY — End-of-Class Assessment", "TEACHER ONLY",
                         "", "Assessment", DARK)
    tb(slide, Inches(0.45), Inches(1.28), Inches(12.4), Inches(0.32),
       "Record what he did today, not what you hoped for. Fill this in right after "
       "class.", size=12, bold=True, color=LAVA)
    cols = [(0.45, 4.3, "READING SKILL"), (4.95, 2.6, "Independent"),
            (7.75, 2.4, "With Help"), (10.35, 2.5, "Needs More Practice")]
    header_y = Inches(1.66)
    for left, width, label in cols:
        add_round(slide, Inches(left), header_y, Inches(width), Inches(0.4), INK)
        tb(slide, Inches(left), header_y + Inches(0.06), Inches(width), Inches(0.3),
           label, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = header_y + Inches(0.46 + i * 0.35)
        band = WHITE if i % 2 == 0 else L_GREY
        add_round(slide, Inches(0.45), top, Inches(4.3), Inches(0.33), band)
        tb(slide, Inches(0.65), top + Inches(0.03), Inches(3.9), Inches(0.27), skill,
           size=12, bold=True, color=INK)
        for left, width, _l in cols[1:]:
            add_round(slide, Inches(left), top, Inches(width), Inches(0.33), band)
            tb(slide, Inches(left), top + Inches(0.0), Inches(width), Inches(0.3), "☐",
               size=13, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.45), Inches(5.78), Inches(12.4), Inches(1.05), L_AMBER)
    tb(slide, Inches(0.7), Inches(5.87), Inches(11.9), Inches(0.34),
       "TODAY I CAN...   (read these out loud with him before he leaves)", size=12,
       bold=True, color=SAND)
    for i, item in enumerate(TODAY_I_CAN):
        left = Inches(0.7 + i * 3.05)
        add_round(slide, left, Inches(6.25), Inches(2.85), Inches(0.46), WHITE)
        tb(slide, left + Inches(0.15), Inches(6.33), Inches(2.6), Inches(0.34),
           f"☐  {item}", size=11, bold=True, color=INK)


def s62_answer_key():
    slide, n = new_slide("🔒 TEACHER ONLY — Answer Key", "TEACHER ONLY", "",
                         "Answer key", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Hide this slide before presenting, or keep it on a second screen.", size=12,
       bold=True, color=LAVA)
    add_round(slide, Inches(0.45), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(0.7), Inches(1.86), Inches(5.6), Inches(0.36),
       "🔊 Sounds  ·  🥚 Eggs  ·  👣 Blending  ·  🦴 Fossils", size=13, bold=True,
       color=LAVA)
    bullets(slide, Inches(0.7), Inches(2.28), Inches(5.6), Inches(4.3), [
        "First Sound 1–3: dino, egg, rock",
        "First Sound 4–6: bone, tail, footprint",
        "First Sound 7–8: mud, sun",
        "Ending sounds: /g/, /t/, /n/, /g/",
        "Eggs 1–3: bed, map, sit",
        "Eggs 4–6: log, red, hot",
        "Eggs 7–8: big, cat",
        "Footprint Blend: CAT, MAP, DIG, BIG, HOT,",
        "     LOG, RUN, SUN, TOP",
        "Which Word Is Right: CAT, DIG, SUN, LOG",
        "Fossil Match 1–3: dig, big, sun",
        "Fossil Match 4–6: egg, log, mud",
    ], size=11, sp=4)
    add_round(slide, Inches(6.75), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(7.0), Inches(1.86), Inches(5.6), Inches(0.36),
       "👨‍👩‍👧 Families  ·  🗺️ Sight words  ·  🧩 Sentences  ·  📖 Story", size=13,
       bold=True, color=JUNGLE)
    bullets(slide, Inches(7.0), Inches(2.28), Inches(5.6), Inches(4.3), [
        "Dino Family -AT: cat, bat, hat  (sun does not)",
        "Dino Family -IG: big, dig, pig  (hop does not)",
        "Dino Family -UN: sun, run, fun  (map does not)",
        "Dino Family -OP: top, hop, mop  (bat does not)",
        "New words: SAT, PIG, FUN, MOP",
        "Sight-Word Path 1–4: the, a, is, in",
        "Sight-Word Path 5–8: my, can, see, on",
        "Sentence puzzles: The dino is big. / The egg is red. /",
        "     I see a dino. / The dino can dig.",
        "Story Q1–6: Ben · green · a red egg · run and dig ·",
        "     a bone fossil · to the jungle",
        "Final: BIG · DIG · THE · \"The dino can run.\" ·",
        "     run / dig / hop",
    ], size=11, sp=4)


BUILDERS = [
    s01_title, s02_explorer, s03_map, s04_promises, s05_valley_scene, s06_talk,
    s07_sounds, s08_first_a, s09_first_b, s10_first_c, s11_ending, s12_valley_done,
    s13_vowels, s14_egg_a, s15_egg_b, s16_egg_bonus, s17_dino_words, s18_cave_done,
    s19_six_steps, s20_blend_a, s21_blend_b, s22_blend_c, s23_trick, s24_no_support,
    s25_fossil_how, s26_fossil_a, s27_fossil_b, s28_fossil_challenge,
    s29_dino_says_intro, s30_dino_says, s31_families, s32_family_a, s33_family_b,
    s34_new_words, s35_sight_1, s36_sight_2, s37_path_a, s38_path_b, s39_phrases,
    s40_read_method, s41_sentences_a, s42_sentences_b, s43_puzzle_a, s44_puzzle_b,
    s45_own_sentence, s46_story_intro, s47_story_1, s48_story_2, s49_story_3,
    s50_story_4, s51_story_again, s52_q_a, s53_q_b, s54_q_c, s55_museum, s56_final_a,
    s57_final_b, s58_badge, s59_can_read, s60_support, s61_assessment, s62_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

out = r"C:\Users\bhushaja\Downloads\shaip\Grade2_Dinosaur_Discovery_Reading_90min.pptx"
try:
    prs.save(out)
except PermissionError:
    out = out.replace(".pptx", "_new.pptx")
    prs.save(out)

story_words = sum(len(line.split()) for _p, _e, _c, _l, lines in STORY for line in lines)
with_notes = [i + 1 for i, s in enumerate(prs.slides)
              if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip()]
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
print(f"Story word count: {story_words}")
print(f"Slides carrying speaker notes: {with_notes if with_notes else 'none'}")
