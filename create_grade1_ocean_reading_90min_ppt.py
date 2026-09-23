"""Grade 1 reading lesson - 90 minutes, 58 slides, no speaker notes.

"Under the Sea Reading Adventure" - the child is an Ocean Explorer swimming from
the Coral Reef to Reading Island. Reading is built one level at a time: letter
sounds, beginning sounds, short vowels, CVC blending, word families, sight words,
phrases, sentences, a short story, then comprehension.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Ocean Hint strip, and slides 56-58 hold the support system,
the observation checklist and the answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

DEEP = RGBColor(0x0B, 0x25, 0x45)
DARK = RGBColor(0x1E, 0x2A, 0x38)
SOFT = RGBColor(0x6B, 0x7A, 0x8C)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
FOAM = RGBColor(0xF3, 0xFA, 0xFC)
AQUA = RGBColor(0x11, 0x8F, 0xB0)
CORAL = RGBColor(0xE8, 0x5D, 0x4E)
KELP = RGBColor(0x1F, 0x96, 0x7B)
SUNNY = RGBColor(0xE0, 0x9B, 0x10)
OCTO = RGBColor(0x76, 0x55, 0xC2)
SAND = RGBColor(0xC9, 0x8A, 0x3E)
NAVY = RGBColor(0x2C, 0x4A, 0x74)
L_AQUA = RGBColor(0xDE, 0xF1, 0xF6)
L_CORAL = RGBColor(0xFD, 0xEB, 0xE8)
L_KELP = RGBColor(0xE2, 0xF4, 0xEF)
L_SUNNY = RGBColor(0xFD, 0xF2, 0xD9)
L_OCTO = RGBColor(0xEE, 0xEA, 0xFA)
L_SAND = RGBColor(0xFA, 0xF0, 0xE0)
L_NAVY = RGBColor(0xE8, 0xEE, 0xF7)
L_GREY = RGBColor(0xF2, 0xF5, 0xF8)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 56
_counter = {"n": 0}

# ------------------------------------------------------------------ content

STOPS = [("🐚", "Coral Reef", "Sounds", "0–17 min", CORAL, L_CORAL),
         ("🐠", "Fish Bay", "Building words", "17–35 min", AQUA, L_AQUA),
         ("🐙", "Octopus Garden", "Word games", "35–57 min", OCTO, L_OCTO),
         ("🐳", "Whale Cove", "Sentences", "57–73 min", NAVY, L_NAVY),
         ("🏝️", "Reading Island", "Story & questions", "73–90 min", SUNNY, L_SUNNY)]

PROMISES = [("🌊", "Ocean Hint", "A hint is always waiting when a word is tricky."),
            ("🐠", "Easy Catch", "We start with the easy ones every time."),
            ("🐙", "Try Again", "Trying again is part of reading, not a mistake."),
            ("⭐", "Bonus Challenge", "Finished early? There is always one more.")]

REEF_SCENE = [("🐟", "fish"), ("🐚", "shell"), ("🪸", "coral"), ("☀️", "sun"),
              ("🪨", "rock"), ("🫧", "bubbles")]

TALK_QS = [("🗣️", "What do you see?"), ("🎨", "What color is the fish?"),
           ("🐚", "Can you find the shell?"), ("⭐", "Which one do you like best?")]

STARTERS = ["I see a ______.", "It is ______."]

SOUND_CARDS = [("🐟", "FISH", "f", "/f/"), ("☀️", "SUN", "s", "/s/"),
               ("🐚", "SHELL", "sh", "/sh/"), ("🐱", "CAT", "c", "/k/"),
               ("🥅", "NET", "n", "/n/"), ("🪨", "ROCK", "r", "/r/")]

CATCH_A = [("F", [("🐟", "fish"), ("☀️", "sun"), ("🐱", "cat")]),
           ("S", [("🐱", "cat"), ("☀️", "sun"), ("🥅", "net")]),
           ("C", [("🐚", "shell"), ("🐱", "cat"), ("🐟", "fish")])]
CATCH_B = [("N", [("🥅", "net"), ("🪨", "rock"), ("☀️", "sun")]),
           ("SH", [("🐱", "cat"), ("🐚", "shell"), ("🐟", "fish")]),
           ("R", [("🪨", "rock"), ("🥅", "net"), ("🐚", "shell")])]
CATCH_C = [("W", [("💧", "wet"), ("🐱", "cat"), ("☀️", "sun")]),
           ("H", [("🎩", "hat"), ("🐟", "fish"), ("🪨", "rock")])]

VOWEL_SHELLS = [("A", "🐱", "cat", CORAL, L_CORAL), ("E", "🥅", "net", KELP, L_KELP),
                ("I", "🦈", "fin", AQUA, L_AQUA), ("O", "🐸", "hop", OCTO, L_OCTO)]

SORT_A = [("🐱", "CAT", "/k/ /a/ /t/", ["A", "E", "I"]),
          ("🥅", "NET", "/n/ /e/ /t/", ["A", "E", "O"]),
          ("🦈", "FIN", "/f/ /i/ /n/", ["I", "O", "A"])]
SORT_B = [("🐸", "HOP", "/h/ /o/ /p/", ["E", "I", "O"]),
          ("💧", "WET", "/w/ /e/ /t/", ["A", "E", "I"]),
          ("🪑", "SIT", "/s/ /i/ /t/", ["I", "O", "E"])]
SORT_BONUS = [("🐋", "BIG", "/b/ /i/ /g/", ["A", "I", "O"]),
              ("🪵", "LOG", "/l/ /o/ /g/", ["E", "O", "I"]),
              ("🔴", "RED", "/r/ /e/ /d/", ["E", "A", "O"])]

OCEAN_WORDS = [("fish", "🐟"), ("fin", "🦈"), ("shell", "🐚"), ("sea", "🌊"),
               ("ship", "🚢"), ("net", "🥅"), ("rock", "🪨"), ("log", "🪵"),
               ("swim", "🏊"), ("sun", "☀️"), ("red", "🔴"), ("big", "🐋"),
               ("wet", "💧"), ("cat", "🐱"), ("hop", "🐸"), ("sit", "🪑"),
               ("run", "🏃"), ("shell", "🐚")]

BLEND_STEPS = [("1", "HEAR", "I say the sounds.", "/f/  /i/  /n/", CORAL, L_CORAL),
               ("2", "SAY", "You say them back.", "/f/  /i/  /n/", SUNNY, L_SUNNY),
               ("3", "BLEND", "We push them together.", "f-i-n", AQUA, L_AQUA),
               ("4", "READ", "You read the word.", "FIN", KELP, L_KELP)]

BUILD_A = [("🐱", "CAT", ["C", "A", "T"], ["/k/", "/a/", "/t/"], CORAL, L_CORAL),
           ("🦈", "FIN", ["F", "I", "N"], ["/f/", "/i/", "/n/"], AQUA, L_AQUA),
           ("🐸", "HOP", ["H", "O", "P"], ["/h/", "/o/", "/p/"], OCTO, L_OCTO)]
BUILD_B = [("🥅", "NET", ["N", "E", "T"], ["/n/", "/e/", "/t/"], KELP, L_KELP),
           ("☀️", "SUN", ["S", "U", "N"], ["/s/", "/u/", "/n/"], SUNNY, L_SUNNY),
           ("🪑", "SIT", ["S", "I", "T"], ["/s/", "/i/", "/t/"], AQUA, L_AQUA)]
BUILD_C = [("💧", "WET", ["W", "E", "T"], ["/w/", "/e/", "/t/"], KELP, L_KELP),
           ("🪵", "LOG", ["L", "O", "G"], ["/l/", "/o/", "/g/"], SAND, L_SAND),
           ("🐋", "BIG", ["B", "I", "G"], ["/b/", "/i/", "/g/"], NAVY, L_NAVY)]

REAL_A = [("🐱", ["CAT", "CAG", "CAP"]), ("🦈", ["FIP", "FIN", "FIT"]),
          ("🥅", ["NEP", "NED", "NET"])]
REAL_B = [("🐸", ["HOD", "HOP", "HOB"]), ("☀️", ["SUM", "SUP", "SUN"])]

DOLPHIN = [("🙌", "Dolphin says touch your head."),
           ("👏", "Dolphin says clap two times."),
           ("👉", "Dolphin says point to the word CAT."),
           ("🐟", "Dolphin says make a fish swimming motion."),
           ("🔊", "Dolphin says say the /f/ sound."),
           ("🤫", "Dolphin says whisper the word SEA.")]

FAMILIES = [("-AT", ["cat", "hat", "bat", "sat"], CORAL, L_CORAL),
            ("-IN", ["fin", "pin", "win"], AQUA, L_AQUA),
            ("-OP", ["hop", "top", "pop"], OCTO, L_OCTO),
            ("-ET", ["net", "wet", "pet"], KELP, L_KELP)]

BELONGS_A = [("-AT", ["CAT", "FIN", "BAT", "HAT"], CORAL, L_CORAL),
             ("-IN", ["PIN", "HOP", "WIN", "FIN"], AQUA, L_AQUA)]
BELONGS_B = [("-OP", ["TOP", "NET", "POP", "HOP"], OCTO, L_OCTO),
             ("-ET", ["WET", "SAT", "PET", "NET"], KELP, L_KELP)]

NEW_WORDS = [("M", "-AT", "MAT", "🟫", CORAL, L_CORAL),
             ("T", "-IN", "TIN", "🥫", AQUA, L_AQUA),
             ("M", "-OP", "MOP", "🧹", OCTO, L_OCTO),
             ("G", "-ET", "GET", "🤲", KELP, L_KELP)]

SIGHT_WORDS = [("I", "I can swim."), ("see", "I see a fish."), ("a", "a big shell"),
               ("the", "the red fish"), ("my", "my little fish"), ("can", "I can go."),
               ("go", "go to the sea"), ("to", "to the rock")]

BUBBLES_A = [("see", ["the", "see", "my"]), ("can", ["can", "go", "to"]),
             ("the", ["a", "I", "the"]), ("my", ["my", "see", "can"])]
BUBBLES_B = [("I", ["to", "I", "a"]), ("go", ["go", "the", "my"]),
             ("a", ["can", "a", "see"]), ("to", ["my", "to", "go"])]

PHRASES_A = [("a big fish", "🐟"), ("the red shell", "🐚"), ("my little fish", "🦈"),
             ("in the sea", "🌊")]
PHRASES_B = [("on the rock", "🪨"), ("a wet net", "🥅"), ("the hot sun", "☀️"),
             ("a small ship", "🚢")]

READ_STEPS = [("STEP 1", "I read", "Teacher reads it first.", CORAL),
              ("STEP 2", "We read", "Teacher and child together.", SUNNY),
              ("STEP 3", "You read", "Child reads it alone — if ready.", KELP)]

SENTENCES_A = [("I see a fish.", "🐟", CORAL, L_CORAL),
               ("The fish is red.", "🔴", AQUA, L_AQUA)]
SENTENCES_B = [("My fish is big.", "🐋", OCTO, L_OCTO),
               ("I can swim.", "🏊", KELP, L_KELP)]
SENTENCES_C = [("I see the sun.", "☀️", SUNNY, L_SUNNY),
               ("The shell is on the rock.", "🐚", NAVY, L_NAVY)]

PUZZLES_A = [(["see", "I", "fish", "a"], "I see a fish.", "🐟", CORAL, L_CORAL),
             (["is", "fish", "The", "big"], "The fish is big.", "🐋", AQUA, L_AQUA)]
PUZZLES_B = [(["can", "I", "swim"], "I can swim.", "🏊", KELP, L_KELP),
             (["the", "see", "I", "shell"], "I see the shell.", "🐚", OCTO, L_OCTO)]

OWN_SENTENCE = [("🐟", "fish"), ("🐚", "shell"), ("☀️", "sun"), ("🚢", "ship")]

STORY_WATCH = [("sea", "🌊"), ("fish", "🐟"), ("shell", "🐚"), ("rock", "🪨"),
               ("red", "🔴"), ("big", "🐋")]

STORY = [
    ("Part 1", "🌊", AQUA, L_AQUA,
     ["Mia is at the sea.", "The sun is big and hot.", "Mia sits on a rock.",
      "She can see the water.", "The water is blue."]),
    ("Part 2", "🐟", CORAL, L_CORAL,
     ["Mia sees a little fish.", "The fish is red.", "It can swim very fast.",
      "The fish looks at Mia.", "Mia waves at the fish."]),
    ("Part 3", "🐚", OCTO, L_OCTO,
     ["Mia sees a big shell.", "The shell is wet.", "She puts the shell in the sea.",
      "The little fish swims to the shell.", "It looks in."]),
    ("Part 4", "😊", KELP, L_KELP,
     ["The fish goes in the shell.", "It is a good home.", "Mia is happy.",
      "She waves to the little fish.", "\"Bye, little fish,\" says Mia."]),
]

EASY_LINES = ["Mia is at the sea.", "The fish is red.", "It looks in.",
              "Mia is happy.", "The shell is wet.", "She can see the water."]

QUESTIONS_A = [("👧", "Who is in the story?", ["Mia", "Sam", "Ben"], "Part 1"),
               ("📍", "Where is Mia?", ["At the sea", "At school", "At home"], "Part 1")]
QUESTIONS_B = [("🎨", "What color is the fish?", ["🔴 Red", "🔵 Blue", "🟢 Green"],
                "Part 2"),
               ("🏊", "What does the fish do?", ["Swims", "Runs", "Hops"], "Part 2")]
QUESTION_C = [("🐚", "What does Mia see in Part 3?", ["A shell", "A cat", "A log"],
               "Part 3")]

FINAL_A = [("1", "🐠", "READ A WORD", "Read this CVC word:", "FIN", AQUA, L_AQUA),
           ("2", "🫧", "READ A SIGHT WORD", "Read this word:", "SEE", OCTO, L_OCTO),
           ("3", "🐙", "BLEND A WORD", "Blend these sounds:", "/n/  /e/  /t/", KELP,
            L_KELP)]
FINAL_B = [("4", "📕", "READ A SENTENCE", "Read this out loud:", "I see a big fish.",
            CORAL, L_CORAL),
           ("5", "🔎", "ANSWER A QUESTION", "From the story:",
            "What color is the fish?", SUNNY, L_SUNNY)]

CAN_READ = [("🔊", "Letter sounds", "f  s  sh  c  n  r"),
            ("🐟", "Beginning sounds", "Which one starts with /f/?"),
            ("🐚", "Short vowels", "a  e  i  o"),
            ("🐙", "Blending", "/f/ /i/ /n/ → FIN"),
            ("📦", "CVC words", "cat  fin  hop  net  sun  sit"),
            ("👨‍👩‍👧", "Word families", "-at  -in  -op  -et"),
            ("🫧", "Sight words", "I  see  a  the  my  can  go  to"),
            ("🔗", "Short phrases", "a big fish"),
            ("📕", "Sentences", "I see a fish."),
            ("📖", "A whole story", "Mia and the Little Fish")]

SUPPORT_LEVELS = [("🟢", "SUPPORT", "Picture + sound + teacher help",
                   "Say the word first, then let her repeat it.", CORAL, L_CORAL),
                  ("🟡", "PRACTICE", "Child reads on her own",
                   "Give the first sound only, then go quiet.", SUNNY, L_SUNNY),
                  ("⭐", "BONUS", "Child creates or reads extra",
                   "Ask for a sentence using the word.", KELP, L_KELP)]

HINT_BANK = ["Point to each letter.", "Say the sounds slowly.", "Blend the sounds.",
             "Look at the picture.", "Read it again.", "Look back at the story."]

PRAISE = ["\"Good listening!\"", "\"You found the first sound!\"",
          "\"Let's try that one together.\"", "\"You are getting closer!\"",
          "\"Nice blending!\"", "\"You read that all by yourself.\""]

RUBRIC_SKILLS = ["Beginning sounds", "Short vowels", "CVC words", "Blending",
                 "Word families", "Sight words", "Short phrases", "Simple sentences",
                 "Story reading", "Comprehension"]

RUBRIC_NOTES = ["Words the child read on her own", "Words that needed help",
                "What to practice next"]

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
             RGBColor(0xDF, 0xE7, 0xEE))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), AQUA)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), DEEP)
    msg = "🌊 Under the Sea Reading Adventure  |  Grade 1  |  90 min"
    if stop:
        msg += f"  |  {stop}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=10,
       color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, stop, accent=AQUA, bg=FOAM):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, accent)
    add_round(slide, Inches(0.4), Inches(0.26), Inches(3.5), Inches(0.42), accent)
    tb(slide, Inches(0.4), Inches(0.31), Inches(3.5), Inches(0.34), tag, size=12,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if timing:
        add_round(slide, Inches(10.3), Inches(0.26), Inches(2.65), Inches(0.42), NAVY)
        tb(slide, Inches(10.3), Inches(0.31), Inches(2.65), Inches(0.34), timing,
           size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title, size=28,
       bold=True, color=DEEP, font="Georgia")
    footer(slide, n, timing, stop)
    fade(slide)
    return slide, n


def one_task(slide, text, color=CORAL, top=1.34):
    """States the single job of this slide in words the child understands."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text, size=16,
              bold=True, color=color)


def hint(slide, text, top=6.42, label="🌊 OCEAN HINT", fill=L_AQUA, color=AQUA):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.2), Inches(0.34), label, size=12,
       bold=True, color=color)
    tb(slide, Inches(3.05), Inches(top + 0.05), Inches(9.5), Inches(0.36), text, size=13,
       bold=True, color=DEEP)


def ido_wedo_youdo(slide, top=6.42):
    steps = [("I READ", "Teacher", CORAL), ("WE READ", "Together", SUNNY),
             ("YOU READ", "You!", KELP)]
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


def catch_rows(slide, items, start_index):
    """Fish Sound Catch: a lettered fish, then three big picture buttons."""
    tops = numbered_rows(slide, len(items), start_index, CORAL, top_start=1.9, gap=1.5)
    for (letter, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.3), Inches(0.9), Inches(0.72), "🐟",
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.4), top + Inches(0.36), Inches(1.7), Inches(0.62),
                  CORAL)
        tb(slide, Inches(2.4), top + Inches(0.44), Inches(1.7), Inches(0.46), letter,
           size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, (emoji, label) in enumerate(options):
            left = Inches(4.45 + j * 2.75)
            add_round(slide, left, top + Inches(0.16), Inches(2.5), Inches(1.02),
                      L_CORAL)
            tb(slide, left, top + Inches(0.2), Inches(2.5), Inches(0.58), emoji, size=26,
               align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.78), Inches(2.5), Inches(0.36), label,
               size=15, bold=True, color=DEEP, align=PP_ALIGN.CENTER)


def sort_rows(slide, items, start_index):
    """Shell Sound Sort: word plus its sounds, then vowel shells to choose from."""
    tops = numbered_rows(slide, len(items), start_index, AQUA, top_start=1.9, gap=1.5)
    for (emoji, word, sounds, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.32), Inches(0.9), Inches(0.7), emoji,
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.4), top + Inches(0.2), Inches(2.4), Inches(0.6), word,
           size=28, bold=True, color=DEEP, font="Arial Black")
        tb(slide, Inches(2.4), top + Inches(0.82), Inches(2.4), Inches(0.4), sounds,
           size=15, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(5.45 + j * 2.42)
            add_round(slide, left, top + Inches(0.28), Inches(2.2), Inches(0.76), L_AQUA)
            tb(slide, left, top + Inches(0.3), Inches(2.2), Inches(0.34), "🐚", size=12,
               align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.6), Inches(2.2), Inches(0.42), opt, size=22,
               bold=True, color=AQUA, align=PP_ALIGN.CENTER, font="Arial Black")


def build_rows(slide, items, top_start=1.9, gap=1.52):
    """Octopus Word Builder: letters, sounds, then the finished word."""
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


def real_word_rows(slide, items, start_index):
    """Catch the Real Word: one picture, three near-identical spellings."""
    tops = numbered_rows(slide, len(items), start_index, KELP, top_start=1.95, gap=1.5,
                         height=1.32)
    for (emoji, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.3), Inches(1.0), Inches(0.72), emoji,
           size=30, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.55), top + Inches(0.42), Inches(1.5), Inches(0.5),
           "Catch it!", size=14, bold=True, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(4.3 + j * 2.85)
            add_round(slide, left, top + Inches(0.28), Inches(2.6), Inches(0.76),
                      L_KELP)
            tb(slide, left, top + Inches(0.39), Inches(2.6), Inches(0.55), opt, size=26,
               bold=True, color=DEEP, align=PP_ALIGN.CENTER, font="Arial Black")


def belongs_rows(slide, items, start_index):
    """Who Belongs to the Family: one family, four candidate words."""
    tops = numbered_rows(slide, len(items), start_index, OCTO, top_start=2.0, gap=2.15,
                         height=1.9)
    for (family, words, color, light), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.5), Inches(0.9), Inches(0.7), "🐠",
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.4), top + Inches(0.6), Inches(1.9), Inches(0.7),
                  color)
        tb(slide, Inches(2.4), top + Inches(0.71), Inches(1.9), Inches(0.5), family,
           size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, word in enumerate(words):
            left = Inches(4.65 + j * 2.0)
            add_round(slide, left, top + Inches(0.55), Inches(1.85), Inches(0.8), light)
            tb(slide, left, top + Inches(0.67), Inches(1.85), Inches(0.56), word,
               size=22, bold=True, color=DEEP, align=PP_ALIGN.CENTER,
               font="Arial Black")
        tb(slide, Inches(4.65), top + Inches(1.42), Inches(7.9), Inches(0.3),
           "Circle every word that belongs to this family.", size=11, color=SOFT)


def bubble_rows(slide, items, start_index, top_start=1.95):
    """Whale's Bubbles: find one sight word among three."""
    for i, (target, options) in enumerate(items):
        top = Inches(top_start + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), WHITE)
        add_oval(slide, Inches(0.78), top + Inches(0.22), Inches(0.52), Inches(0.52),
                 NAVY)
        tb(slide, Inches(0.78), top + Inches(0.28), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.27), Inches(1.0), Inches(0.44), "FIND:",
           size=13, bold=True, color=SOFT)
        add_round(slide, Inches(2.55), top + Inches(0.2), Inches(1.9), Inches(0.58),
                  NAVY)
        tb(slide, Inches(2.55), top + Inches(0.29), Inches(1.9), Inches(0.42), target,
           size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(4.85 + j * 2.6)
            add_oval(slide, left, top + Inches(0.14), Inches(2.4), Inches(0.7), L_NAVY)
            tb(slide, left, top + Inches(0.26), Inches(2.4), Inches(0.48), opt, size=20,
               bold=True, color=DEEP, align=PP_ALIGN.CENTER, font="Arial Black")


def phrase_grid(slide, items, top=1.95):
    for i, (phrase, emoji) in enumerate(items):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        t = Inches(top + row * 1.6)
        add_round(slide, left, t, Inches(5.95), Inches(1.4), L_AQUA)
        tb(slide, left + Inches(0.3), t + Inches(0.36), Inches(1.0), Inches(0.72), emoji,
           size=30, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.5), t + Inches(0.36), Inches(4.2), Inches(0.7),
           phrase, size=28, bold=True, color=DEEP)


def sentence_rows(slide, items, top_start=2.0, gap=2.15, height=1.9, size=42):
    for i, (text, emoji, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), light)
        tb(slide, Inches(0.85), top + Inches(height / 2 - 0.45), Inches(1.4),
           Inches(0.9), emoji, size=40, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(height / 2 - 0.42), Inches(7.6),
           Inches(0.85), text, size=size, bold=True, color=DEEP)
        add_round(slide, Inches(10.4), top + Inches(height / 2 - 0.28), Inches(2.2),
                  Inches(0.56), WHITE)
        tb(slide, Inches(10.4), top + Inches(height / 2 - 0.2), Inches(2.2),
           Inches(0.4), "I · WE · YOU", size=13, bold=True, color=color,
           align=PP_ALIGN.CENTER)


def puzzle_rows(slide, items, start_index):
    """Build the Ocean Sentence: shuffled word cards plus an answer line."""
    tops = numbered_rows(slide, len(items), start_index, SUNNY, top_start=2.0, gap=2.15,
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
                 SUNNY)
        tb(slide, Inches(0.75), top + Inches(0.76), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.45), top + Inches(0.6), Inches(1.0), Inches(0.72), emoji,
           size=30, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(0.5), Inches(3.5), Inches(0.66), question,
           size=20, bold=True, color=DEEP)
        tb(slide, Inches(2.6), top + Inches(1.18), Inches(3.5), Inches(0.36),
           f"look in {where}", size=11, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(6.3 + j * 2.2)
            add_round(slide, left, top + Inches(0.5), Inches(2.0), Inches(0.92),
                      L_SUNNY)
            tb(slide, left, top + Inches(0.68), Inches(2.0), Inches(0.6), opt, size=17,
               bold=True, color=DEEP, align=PP_ALIGN.CENTER)


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, DEEP)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), AQUA)
    for x, y, c in [(0.6, 5.6, CORAL), (12.05, 5.55, KELP)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(0.9), Inches(12), Inches(1.0), "🌊", size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.9), Inches(12), Inches(0.9),
       "Under the Sea Reading Adventure", size=44, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.85), Inches(12), Inches(0.5),
       "Dive Into Reading!", size=22, color=RGBColor(0x9E, 0xDC, 0xEC),
       align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, label) in enumerate(REEF_SCENE):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(3.55), Inches(1.65), Inches(1.35),
                  RGBColor(0x17, 0x38, 0x5E))
        tb(slide, left, Inches(3.73), Inches(1.65), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.38), Inches(1.45), Inches(0.4), label,
           size=10, color=RGBColor(0xB9, 0xD4, 0xE6), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.3), Inches(5.2), Inches(6.7), Inches(1.15), AQUA)
    tb(slide, Inches(3.5), Inches(5.45), Inches(6.3), Inches(0.7),
       "Grade 1  •  90 Minutes  •  Reading", size=19, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Coral Reef")
    fade(slide)


def s02_explorer():
    slide, n = new_slide("🤿 You Are the Ocean Explorer", "WELCOME", "0–7 min",
                         "Coral Reef", AQUA)
    one_task(slide, "Today we are Ocean Readers!", AQUA)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.5), L_AQUA)
    tb(slide, Inches(0.5), Inches(2.4), Inches(5.4), Inches(1.7), "🤿", size=88,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.35), Inches(5.4), Inches(0.7), "THAT'S YOU!",
       size=32, bold=True, color=AQUA, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.8), Inches(5.15), Inches(4.8), Inches(0.8),
       "An Ocean Explorer who reads", size=15, color=SOFT, align=PP_ALIGN.CENTER)
    lines = [("🌊", "We swim to five places today."),
             ("📖", "Each place has one reading mission."),
             ("🫧", "Every mission you finish is a bubble earned."),
             ("🏆", "At the end you become an Ocean Reading Explorer.")]
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.9 + i * 1.18)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(1.02), L_SUNNY)
        add_oval(slide, Inches(6.5), top + Inches(0.22), Inches(0.58), Inches(0.58),
                 WHITE)
        tb(slide, Inches(6.5), top + Inches(0.28), Inches(0.58), Inches(0.44), icon,
           size=16, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.3), top + Inches(0.28), Inches(5.3), Inches(0.5), line,
           size=16, bold=True, color=DEEP)


def s03_journey():
    slide, n = new_slide("🗺️ Our Ocean Journey", "MAP", "0–7 min", "Coral Reef", KELP)
    one_task(slide, "Five stops. One mission at each stop.", KELP)
    for i, (emoji, place, job, when, color, light) in enumerate(STOPS):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(3.5), light)
        tb(slide, left, Inches(2.2), Inches(2.3), Inches(0.85), emoji, size=36,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.1), Inches(2.1), Inches(0.9), place,
           size=17, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.98), Inches(2.0), Inches(0.5), job,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.65), Inches(1.7), Inches(0.42),
                  WHITE)
        tb(slide, left + Inches(0.3), Inches(4.69), Inches(1.7), Inches(0.34), when,
           size=10, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.65), Inches(12.35), Inches(0.9), L_AQUA)
    tb(slide, Inches(0.8), Inches(5.9), Inches(11.7), Inches(0.44),
       "🐚  →  🐠  →  🐙  →  🐳  →  🏝️   Finish all five and you are an Ocean Reading "
       "Explorer!", size=16, bold=True, color=DEEP, align=PP_ALIGN.CENTER)


def s04_promises():
    slide, n = new_slide("🌊 How I Help You Today", "PROMISE", "0–7 min", "Coral Reef",
                         SUNNY)
    one_task(slide, "You never read alone. Here is how it works.", SUNNY)
    for i, (icon, name, detail) in enumerate(PROMISES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(2.0 + row * 2.2)
        add_round(slide, left, top, Inches(5.95), Inches(1.95), L_SUNNY)
        add_oval(slide, left + Inches(0.32), top + Inches(0.6), Inches(0.75),
                 Inches(0.75), WHITE)
        tb(slide, left + Inches(0.32), top + Inches(0.7), Inches(0.75), Inches(0.55),
           icon, size=20, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.3), top + Inches(0.4), Inches(4.4), Inches(0.55),
           name, size=22, bold=True, color=SUNNY)
        tb(slide, left + Inches(1.3), top + Inches(1.0), Inches(4.4), Inches(0.75),
           detail, size=14, color=DARK)
    hint(slide, "Read all four out loud. Knowing help is coming makes her braver.", 6.45)


def s05_reef_scene():
    slide, n = new_slide("🪸 Picture Talk — The Coral Reef", "WARM-UP", "0–7 min",
                         "Coral Reef", CORAL)
    one_task(slide, "Just look and talk. No reading yet!", CORAL)
    for i, (emoji, label) in enumerate(REEF_SCENE):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_CORAL)
        tb(slide, left, top + Inches(0.22), Inches(3.9), Inches(0.98), emoji, size=48,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.9), top + Inches(1.34), Inches(2.1),
                  Inches(0.58), WHITE)
        tb(slide, left + Inches(0.9), top + Inches(1.45), Inches(2.1), Inches(0.42),
           label, size=16, bold=True, color=CORAL, align=PP_ALIGN.CENTER)
    hint(slide, "Point at anything she names. Pointing back shows you are listening.",
         6.42, "🗣️ SPEAKING", L_CORAL, CORAL)


def s06_talk():
    slide, n = new_slide("🗣️ Talk About the Reef", "SPEAKING", "0–7 min", "Coral Reef",
                         CORAL)
    one_task(slide, "Answer out loud. Try a whole sentence.", CORAL)
    for i, (icon, question) in enumerate(TALK_QS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.35)
        add_round(slide, left, top, Inches(5.95), Inches(1.15), L_CORAL)
        add_oval(slide, left + Inches(0.3), top + Inches(0.28), Inches(0.58),
                 Inches(0.58), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.34), Inches(0.58), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.33), Inches(4.6), Inches(0.52),
           question, size=18, bold=True, color=DEEP)
    for i, starter in enumerate(STARTERS):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(4.75), Inches(5.95), Inches(1.1), L_SUNNY)
        tb(slide, left + Inches(0.4), Inches(5.02), Inches(5.2), Inches(0.6), starter,
           size=26, bold=True, color=DEEP)
    hint(slide, "If she answers with one word, say it back as a full sentence first.",
         6.45)


def s07_sounds():
    slide, n = new_slide("🔊 Beginning Sounds — Listen and Say", "LEARN", "7–17 min",
                         "Coral Reef", CORAL)
    one_task(slide, "I say the sound. You say it back to me.", CORAL)
    for i, (emoji, word, letter, sound) in enumerate(SOUND_CARDS):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.95 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_CORAL)
        tb(slide, left, top + Inches(0.14), Inches(3.9), Inches(0.7), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.84), Inches(3.6), Inches(0.48),
           word, size=20, bold=True, color=DEEP, align=PP_ALIGN.CENTER,
           font="Arial Black")
        add_round(slide, left + Inches(0.85), top + Inches(1.36), Inches(2.2),
                  Inches(0.6), WHITE)
        tb(slide, left + Inches(0.85), top + Inches(1.46), Inches(2.2), Inches(0.44),
           f"{letter}  says  {sound}", size=15, bold=True, color=CORAL,
           align=PP_ALIGN.CENTER)
    hint(slide, "Say the sound, not the letter name. Stretch it: fffff.", 6.45)


def s08_catch_a():
    slide, n = new_slide("🎣 Game: Fish Sound Catch", "GAME", "7–17 min", "Coral Reef",
                         CORAL)
    one_task(slide, "Which picture STARTS with the letter on the fish?", CORAL)
    catch_rows(slide, CATCH_A, 1)
    hint(slide, "Say each picture name out loud first. Stretch the first sound.", 6.45)


def s09_catch_b():
    slide, n = new_slide("🎣 Fish Sound Catch — Catches 4 to 6", "GAME", "7–17 min",
                         "Coral Reef", CORAL)
    one_task(slide, "Three more fish to catch.", CORAL)
    catch_rows(slide, CATCH_B, 4)
    hint(slide, "Too many choices? Cover one with your hand and leave only two.", 6.45)


def s10_catch_c():
    slide, n = new_slide("🎣 Sound Catch Challenge", "GAME", "7–17 min", "Coral Reef",
                         CORAL)
    one_task(slide, "Two last fish. Then the Coral Reef is done!", CORAL)
    catch_rows(slide, CATCH_C, 7)
    add_round(slide, Inches(0.5), Inches(4.95), Inches(12.35), Inches(1.28), L_KELP)
    tb(slide, Inches(0.8), Inches(5.08), Inches(11.7), Inches(0.38), "⭐ BONUS CHALLENGE",
       size=13, bold=True, color=KELP)
    tb(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.6),
       "Can you think of one more word that starts with /f/?  How about /s/?", size=18,
       bold=True, color=DEEP)
    hint(slide, "Any real word counts, even one that is not in our ocean.", 6.42)


def s11_sound_check():
    slide, n = new_slide("🐚 Sound Check — Coral Reef Complete", "RECAP", "7–17 min",
                         "Coral Reef", KELP)
    one_task(slide, "Read all six sounds with me, fast.", KELP)
    for i, (emoji, word, letter, sound) in enumerate(SOUND_CARDS):
        left = Inches(0.5 + i * 2.08)
        add_round(slide, left, Inches(2.0), Inches(1.9), Inches(2.6), L_KELP)
        tb(slide, left, Inches(2.2), Inches(1.9), Inches(0.62), emoji, size=24,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(2.9), Inches(1.3), Inches(0.72),
                  KELP)
        tb(slide, left + Inches(0.3), Inches(3.02), Inches(1.3), Inches(0.5), sound,
           size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.1), Inches(3.78), Inches(1.7), Inches(0.5), word,
           size=13, bold=True, color=DEEP, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(4.9), Inches(12.35), Inches(1.35), L_AQUA)
    tb(slide, Inches(0.8), Inches(5.05), Inches(11.7), Inches(0.4),
       "🫧 BUBBLE EARNED — Coral Reef", size=16, bold=True, color=AQUA)
    tb(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.6),
       "You can hear the first sound in a word. Next stop: Fish Bay, where we build "
       "whole words.", size=15, color=DEEP)
    hint(slide, "Go left to right, two seconds per sound. Keep it quick and fun.", 6.42)


def s12_vowels():
    slide, n = new_slide("🐚 Short Vowels — Four Shells", "LEARN", "17–27 min",
                         "Fish Bay", AQUA)
    one_task(slide, "Every word has a sound hiding in the middle.", AQUA)
    for i, (letter, emoji, word, color, light) in enumerate(VOWEL_SHELLS):
        left = Inches(0.9 + i * 3.0)
        add_round(slide, left, Inches(1.95), Inches(2.75), Inches(4.1), light)
        tb(slide, left, Inches(2.15), Inches(2.75), Inches(0.72), "🐚", size=30,
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
    hint(slide, "Stretch the middle: c-aaa-t. That is the sound she is listening for.",
         6.42)


def s13_sort_a():
    slide, n = new_slide("🐚 Game: Shell Sound Sort", "GAME", "17–27 min", "Fish Bay",
                         AQUA)
    one_task(slide, "Which shell does this word belong in?", AQUA)
    sort_rows(slide, SORT_A, 1)
    hint(slide, "Say the word slowly and hold the middle sound.", 6.45)


def s14_sort_b():
    slide, n = new_slide("🐚 Shell Sound Sort — Words 4 to 6", "GAME", "17–27 min",
                         "Fish Bay", AQUA)
    one_task(slide, "Three more words to sort.", AQUA)
    sort_rows(slide, SORT_B, 4)
    hint(slide, "Shells:  A  ·  E  ·  I  ·  O.  Say each one before she chooses.", 6.45)


def s15_sort_bonus():
    slide, n = new_slide("⭐ Shell Sort Bonus — Three More", "BONUS", "17–27 min",
                         "Fish Bay", KELP)
    one_task(slide, "Bonus round! These are a little trickier.", KELP)
    sort_rows(slide, SORT_BONUS, 7)
    hint(slide, "Skip this if she is tired. It is a bonus, not a requirement.", 6.45)


def s16_ocean_words():
    slide, n = new_slide("🌊 Ocean Words We Will Read", "VOCABULARY", "27–35 min",
                         "Fish Bay", OCTO)
    one_task(slide, "Every word today comes from this list. Nothing new later.", OCTO)
    seen = []
    words = [w for w in OCEAN_WORDS if not (w[0] in seen or seen.append(w[0]))]
    for i, (word, emoji) in enumerate(words):
        col, row = i % 6, i // 6
        left = Inches(0.5 + col * 2.08)
        top = Inches(2.0 + row * 1.45)
        add_round(slide, left, top, Inches(1.9), Inches(1.32), L_OCTO)
        tb(slide, left, top + Inches(0.12), Inches(1.9), Inches(0.58), emoji, size=22,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), top + Inches(0.72), Inches(1.6),
                  Inches(0.5), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(0.8), Inches(1.6), Inches(0.36),
           word, size=16, bold=True, color=OCTO, align=PP_ALIGN.CENTER)
    hint(slide, "Read them across in a rhythm. She joins in wherever she can.", 6.42)


def s17_blending():
    slide, n = new_slide("🐙 Blending — Hear, Say, Blend, Read", "LEARN", "27–35 min",
                         "Fish Bay", OCTO)
    one_task(slide, "Three sounds can become one word.", OCTO)
    for i, (num, label, detail, example, color, light) in enumerate(BLEND_STEPS):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(3.5), light)
        add_oval(slide, left + Inches(1.1), Inches(2.18), Inches(0.7), Inches(0.7),
                 color)
        tb(slide, left + Inches(1.1), Inches(2.3), Inches(0.7), Inches(0.46), num,
           size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.0), Inches(2.6), Inches(0.55), label,
           size=22, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.2), Inches(3.6), Inches(2.5), Inches(0.62), detail,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(4.35), Inches(2.4), Inches(0.75),
                  WHITE)
        tb(slide, left + Inches(0.25), Inches(4.52), Inches(2.4), Inches(0.5), example,
           size=18, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Slide your finger under the letters, then sweep it fast for the word.",
         6.42)


def s18_build_a():
    slide, n = new_slide("🐙 Game: Octopus Word Builder", "GAME", "27–35 min",
                         "Fish Bay", OCTO)
    one_task(slide, "Say each sound. Then say the whole word fast.", OCTO)
    build_rows(slide, BUILD_A)
    hint(slide, "Cover the last letter so she blends only two sounds first.", 6.5)


def s19_build_b():
    slide, n = new_slide("🐙 Octopus Word Builder — Three More", "GAME", "27–35 min",
                         "Fish Bay", OCTO)
    one_task(slide, "Different vowels this time. Listen carefully.", OCTO)
    build_rows(slide, BUILD_B)
    hint(slide, "Give her the first sound out loud, then let her finish the word.", 6.5)


def s20_build_c():
    slide, n = new_slide("🐙 Word Builder Challenge", "GAME", "27–35 min", "Fish Bay",
                         OCTO)
    one_task(slide, "Three last words. You are building real words now!", OCTO)
    build_rows(slide, BUILD_C)
    hint(slide, "Finished fast? Ask her to swap one letter and read the new word.", 6.5)


def s21_real_a():
    slide, n = new_slide("🎣 Game: Catch the Real Word", "GAME", "35–42 min",
                         "Octopus Garden", KELP)
    one_task(slide, "Only one of these three is a real word. Catch it!", KELP)
    real_word_rows(slide, REAL_A, 1)
    hint(slide, "Look at the last sound. That is where they are different.", 6.55)


def s22_real_b():
    slide, n = new_slide("🎣 Catch the Real Word — Last Two", "GAME", "35–42 min",
                         "Octopus Garden", KELP)
    one_task(slide, "Two more catches.", KELP)
    real_word_rows(slide, REAL_B, 4)
    add_round(slide, Inches(0.5), Inches(5.0), Inches(12.35), Inches(1.28), L_CORAL)
    tb(slide, Inches(0.8), Inches(5.12), Inches(11.7), Inches(0.38), "🐙 TRY AGAIN",
       size=13, bold=True, color=CORAL)
    tb(slide, Inches(0.8), Inches(5.55), Inches(11.7), Inches(0.58),
       "Caught the wrong one? Read all three out loud together and try once more.",
       size=16, bold=True, color=DEEP)
    hint(slide, "Point to each letter as she reads. Slow beats fast here.", 6.45)


def s23_dolphin():
    slide, n = new_slide("🐬 Brain Break: Dolphin Says", "BREAK", "42–47 min",
                         "Octopus Garden", SUNNY)
    one_task(slide, "Stay in your seat. Only move if I say \"Dolphin says.\"", SUNNY)
    for i, (icon, command) in enumerate(DOLPHIN):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.4)
        add_round(slide, left, top, Inches(5.95), Inches(1.2), L_SUNNY)
        add_oval(slide, left + Inches(0.28), top + Inches(0.3), Inches(0.6),
                 Inches(0.6), WHITE)
        tb(slide, left + Inches(0.28), top + Inches(0.36), Inches(0.6), Inches(0.46),
           icon, size=16, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.36), Inches(4.6), Inches(0.55),
           command, size=16, bold=True, color=DEEP)
    hint(slide, "Three of these are reading tasks. Five minutes max, then move on.",
         6.35, "⏱️ TIMING", L_SUNNY, SUNNY)


def s24_families():
    slide, n = new_slide("🐠 Word Families — Four Fish Families", "LEARN", "47–57 min",
                         "Octopus Garden", AQUA)
    one_task(slide, "These words end the same way. They rhyme.", AQUA)
    for i, (name, words, color, light) in enumerate(FAMILIES):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(4.3), light)
        tb(slide, left, Inches(2.1), Inches(2.9), Inches(0.7), "🐠", size=28,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.55), Inches(2.82), Inches(1.8), Inches(0.72),
                  color)
        tb(slide, left + Inches(0.55), Inches(2.94), Inches(1.8), Inches(0.5), name,
           size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, word in enumerate(words):
            top = Inches(3.72 + j * 0.6)
            add_round(slide, left + Inches(0.35), top, Inches(2.2), Inches(0.5), WHITE)
            tb(slide, left + Inches(0.35), top + Inches(0.06), Inches(2.2), Inches(0.38),
               word, size=19, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Read a family in a rhythm and let her say the last word alone.", 6.45)


def s25_belongs_a():
    slide, n = new_slide("🐠 Game: Who Belongs to the Family?", "GAME", "47–57 min",
                         "Octopus Garden", OCTO)
    one_task(slide, "Three of these four belong. Which one does not?", OCTO)
    belongs_rows(slide, BELONGS_A, 1)
    hint(slide, "Cover the first letter. Now only the ending is showing.", 6.45)


def s26_belongs_b():
    slide, n = new_slide("🐠 Who Belongs? — Families 3 & 4", "GAME", "47–57 min",
                         "Octopus Garden", OCTO)
    one_task(slide, "Two more families to sort out.", OCTO)
    belongs_rows(slide, BELONGS_B, 3)
    hint(slide, "Say the odd word next to the family name so she hears the mismatch.",
         6.45)


def s27_new_words():
    slide, n = new_slide("⭐ Make a New Word Together", "BONUS", "47–57 min",
                         "Octopus Garden", KELP)
    one_task(slide, "Add a new letter to the front. What word do you get?", KELP)
    for i, (letter, family, result, emoji, color, light) in enumerate(NEW_WORDS):
        top = Inches(1.95 + i * 1.18)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.02), light)
        add_round(slide, Inches(1.0), top + Inches(0.18), Inches(1.1), Inches(0.66),
                  WHITE)
        tb(slide, Inches(1.0), top + Inches(0.27), Inches(1.1), Inches(0.48), letter,
           size=24, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(2.2), top + Inches(0.29), Inches(0.4), Inches(0.44), "+",
           size=18, bold=True, color=SOFT)
        add_round(slide, Inches(2.75), top + Inches(0.18), Inches(1.6), Inches(0.66),
                  WHITE)
        tb(slide, Inches(2.75), top + Inches(0.27), Inches(1.6), Inches(0.48), family,
           size=22, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(4.5), top + Inches(0.27), Inches(0.5), Inches(0.44), "→",
           size=20, bold=True, color=color)
        add_round(slide, Inches(5.25), top + Inches(0.18), Inches(2.4), Inches(0.66),
                  WHITE)
        tb(slide, Inches(5.25), top + Inches(0.27), Inches(2.4), Inches(0.48), result,
           size=24, bold=True, color=DEEP, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(8.0), top + Inches(0.2), Inches(0.9), Inches(0.62), emoji,
           size=22, align=PP_ALIGN.CENTER)
        tb(slide, Inches(9.2), top + Inches(0.3), Inches(3.3), Inches(0.42),
           "Now read it out loud.", size=13, color=SOFT)
    add_round(slide, Inches(0.5), Inches(6.42), Inches(12.35), Inches(0.46), L_AQUA)
    tb(slide, Inches(0.8), Inches(6.48), Inches(11.7), Inches(0.36),
       "🫧 BUBBLE EARNED — Octopus Garden.  Next stop: Whale Cove.", size=14, bold=True,
       color=AQUA)


def s28_sight_intro():
    slide, n = new_slide("🫧 Sight Words — Words We Just Know", "LEARN", "57–65 min",
                         "Whale Cove", NAVY)
    one_task(slide, "We do not sound these out. We just know them.", NAVY)
    for i, (word, example) in enumerate(SIGHT_WORDS):
        col, row = i % 4, i // 4
        left = Inches(0.5 + col * 3.13)
        top = Inches(1.95 + row * 2.1)
        add_round(slide, left, top, Inches(2.9), Inches(1.9), L_NAVY)
        add_oval(slide, left + Inches(0.45), top + Inches(0.18), Inches(2.0),
                 Inches(0.85), WHITE)
        tb(slide, left + Inches(0.45), top + Inches(0.34), Inches(2.0), Inches(0.58),
           word, size=28, bold=True, color=NAVY, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left + Inches(0.1), top + Inches(1.18), Inches(2.7), Inches(0.5),
           example, size=14, bold=True, color=DEEP, align=PP_ALIGN.CENTER)
    hint(slide, "Cover six and teach only two properly. Two known beats eight guessed.",
         6.25)


def s29_bubbles_a():
    slide, n = new_slide("🐳 Game: Whale's Sight-Word Bubbles", "GAME", "57–65 min",
                         "Whale Cove", NAVY)
    one_task(slide, "Pop the right bubble. Find the word I say.", NAVY)
    bubble_rows(slide, BUBBLES_A, 1)
    hint(slide, "Read all three bubbles out loud. Listening is easier than looking.",
         6.5)


def s30_bubbles_b():
    slide, n = new_slide("🐳 Whale's Bubbles — Four More", "GAME", "57–65 min",
                         "Whale Cove", NAVY)
    one_task(slide, "Four more bubbles. These are all in our story.", NAVY)
    bubble_rows(slide, BUBBLES_B, 5)
    hint(slide, "Missed one? Put it back in the pile and ask again before the story.",
         6.5)


def s31_phrases_a():
    slide, n = new_slide("🔗 Short Phrases — Read Smoothly", "READING", "57–65 min",
                         "Whale Cove", AQUA)
    one_task(slide, "Read each one as one piece. No stopping in the middle.", AQUA)
    phrase_grid(slide, PHRASES_A)
    hint(slide, "A phrase is not a sentence yet. Read it in one breath.", 6.42)


def s32_phrases_b():
    slide, n = new_slide("🔗 Short Phrases — Four More", "READING", "57–65 min",
                         "Whale Cove", AQUA)
    one_task(slide, "Four more. Then we put them into sentences.", AQUA)
    phrase_grid(slide, PHRASES_B)
    hint(slide, "Model it once smoothly, then let her copy your rhythm.", 6.42)


def s33_read_method():
    slide, n = new_slide("📕 How We Read a Sentence", "METHOD", "65–73 min",
                         "Whale Cove", CORAL)
    one_task(slide, "You never read a new sentence first. I do.", CORAL)
    for i, (label, who, what, color) in enumerate(READ_STEPS):
        left = Inches(0.5 + i * 4.15)
        light = [L_CORAL, L_SUNNY, L_KELP][i]
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(2.85), light)
        add_round(slide, left + Inches(1.1), Inches(2.2), Inches(1.7), Inches(0.52),
                  color)
        tb(slide, left + Inches(1.1), Inches(2.28), Inches(1.7), Inches(0.38), label,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(2.9), Inches(3.5), Inches(0.62), who,
           size=26, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.3), Inches(3.62), Inches(3.3), Inches(0.98),
                  WHITE)
        tb(slide, left + Inches(0.45), Inches(3.8), Inches(3.0), Inches(0.68), what,
           size=14, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.05), Inches(12.35), Inches(1.2), L_AQUA)
    tb(slide, Inches(0.8), Inches(5.18), Inches(11.7), Inches(0.38),
       "STUCK ON A WORD INSIDE A SENTENCE?", size=12, bold=True, color=AQUA)
    tb(slide, Inches(0.8), Inches(5.58), Inches(11.7), Inches(0.5),
       "Point → sound it out → blend → then read the whole sentence again from the "
       "start.", size=16, bold=True, color=DEEP)


def s34_sentences_a():
    slide, n = new_slide("📕 Read With Me", "READING", "65–73 min", "Whale Cove", CORAL)
    one_task(slide, "Point to each word while I read it.", CORAL)
    sentence_rows(slide, SENTENCES_A)
    ido_wedo_youdo(slide, 6.42)


def s35_sentences_b():
    slide, n = new_slide("📗 Read With Me — Two More", "READING", "65–73 min",
                         "Whale Cove", KELP)
    one_task(slide, "Look at the picture first. It gives you a clue.", KELP)
    sentence_rows(slide, SENTENCES_B)
    ido_wedo_youdo(slide, 6.42)


def s36_sentences_c():
    slide, n = new_slide("📘 Read With Me — Your Turn Grows", "READING", "65–73 min",
                         "Whale Cove", NAVY)
    one_task(slide, "Pick the one YOU want to read by yourself.", NAVY)
    sentence_rows(slide, SENTENCES_C)
    ido_wedo_youdo(slide, 6.42)


def s37_puzzle_a():
    slide, n = new_slide("🧩 Game: Build the Ocean Sentence", "GAME", "65–73 min",
                         "Whale Cove", SUNNY)
    one_task(slide, "The words got mixed up. Put them back in order!", SUNNY)
    puzzle_rows(slide, PUZZLES_A, 1)
    hint(slide, "Hint: the first word always wears a CAPITAL letter.", 6.45)


def s38_puzzle_b():
    slide, n = new_slide("🧩 Build the Ocean Sentence — Two More", "GAME", "65–73 min",
                         "Whale Cove", SUNNY)
    one_task(slide, "Two more. Arrange, read, then match it to the picture.", SUNNY)
    puzzle_rows(slide, PUZZLES_B, 3)
    hint(slide, "Hand her the first word, then ask only what comes next.", 6.45)


def s39_own_sentence():
    slide, n = new_slide("⭐ Build Your Own Sentence", "BONUS", "65–73 min",
                         "Whale Cove", KELP)
    one_task(slide, "Now you make one. Pick any word you like.", KELP)
    add_round(slide, Inches(0.5), Inches(1.95), Inches(12.35), Inches(1.6), L_KELP)
    tb(slide, Inches(1.0), Inches(2.35), Inches(11.4), Inches(0.8),
       "I  see  a  ______________ .", size=44, bold=True, color=DEEP,
       align=PP_ALIGN.CENTER)
    for i, (emoji, word) in enumerate(OWN_SENTENCE):
        left = Inches(1.15 + i * 2.85)
        add_round(slide, left, Inches(3.85), Inches(2.6), Inches(2.0), WHITE)
        tb(slide, left, Inches(4.05), Inches(2.6), Inches(0.85), emoji, size=38,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.4), Inches(5.0), Inches(1.8), Inches(0.62),
                  L_KELP)
        tb(slide, left + Inches(0.4), Inches(5.11), Inches(1.8), Inches(0.44), word,
           size=20, bold=True, color=KELP, align=PP_ALIGN.CENTER)
    hint(slide, "Offer two words to choose from rather than asking her to invent one.",
         6.42)


def s40_story_intro():
    slide, n = new_slide("📖 Story Time: Mia and the Little Fish", "STORY", "73–82 min",
                         "Reading Island", SUNNY)
    one_task(slide, "Four short parts. Watch for these six words.", SUNNY)
    add_round(slide, Inches(0.5), Inches(1.9), Inches(5.2), Inches(4.4), L_SUNNY)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.2), Inches(1.6), "🐟", size=82,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.3), Inches(4.8), Inches(0.7), "MIA & THE FISH",
       size=26, bold=True, color=SUNNY, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.9), Inches(5.1), Inches(4.4), Inches(0.9),
       "Every word in this story is one you already read today.", size=14, color=SOFT,
       align=PP_ALIGN.CENTER)
    for i, (word, emoji) in enumerate(STORY_WATCH):
        col, row = i % 3, i // 3
        left = Inches(6.0 + col * 2.32)
        top = Inches(1.9 + row * 2.2)
        add_round(slide, left, top, Inches(2.15), Inches(2.0), WHITE)
        tb(slide, left, top + Inches(0.25), Inches(2.15), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), top + Inches(1.15), Inches(1.55),
                  Inches(0.62), L_SUNNY)
        tb(slide, left + Inches(0.3), top + Inches(1.26), Inches(1.55), Inches(0.44),
           word, size=19, bold=True, color=SUNNY, align=PP_ALIGN.CENTER)
    hint(slide, "Read these six now. Meeting them first removes the roadblocks.", 6.45)


def _story_slide(index, timing):
    part, emoji, color, light, lines = STORY[index]
    slide, n = new_slide(f"📖 Mia and the Little Fish — {part}", "STORY", timing,
                         "Reading Island", color)
    add_round(slide, Inches(0.5), Inches(1.45), Inches(4.3), Inches(4.75), light)
    tb(slide, Inches(0.5), Inches(2.9), Inches(4.3), Inches(1.7), emoji, size=92,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(5.05), Inches(1.45), Inches(7.8), Inches(4.75), WHITE)
    for i, line in enumerate(lines):
        tb(slide, Inches(5.45), Inches(1.78 + i * 0.88), Inches(7.1), Inches(0.72),
           line, size=26, bold=True, color=DEEP)
    ido_wedo_youdo(slide, 6.42)
    return slide, n


def s41_story_1():
    _story_slide(0, "73–82 min")


def s42_story_2():
    _story_slide(1, "73–82 min")


def s43_story_3():
    _story_slide(2, "73–82 min")


def s44_story_4():
    _story_slide(3, "73–82 min")


def s45_story_challenge():
    slide, n = new_slide("⭐ Story Reading Challenge", "READING", "73–82 min",
                         "Reading Island", KELP)
    one_task(slide, "Pick TWO lines you want to read all by yourself.", KELP)
    for i, line in enumerate(EASY_LINES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.42)
        add_round(slide, left, top, Inches(5.95), Inches(1.22), L_KELP)
        add_oval(slide, left + Inches(0.3), top + Inches(0.34), Inches(0.55),
                 Inches(0.55), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(0.55), Inches(0.42),
           "🫧", size=13, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.34), Inches(4.6), Inches(0.6),
           line, size=21, bold=True, color=DEEP)
    hint(slide, "She chooses. If she freezes, read the first word and then wait.", 6.35)


def s46_questions_a():
    slide, n = new_slide("🔎 Game: Find the Answer", "COMPREHENSION", "82–87 min",
                         "Reading Island", SUNNY)
    one_task(slide, "Think about Mia's day. Which answer is right?", SUNNY)
    question_rows(slide, QUESTIONS_A, 1)
    hint(slide, "Cannot remember? Go back a slide and read it together.", 6.35)


def s47_questions_b():
    slide, n = new_slide("🔎 Find the Answer — Two More", "COMPREHENSION", "82–87 min",
                         "Reading Island", SUNNY)
    one_task(slide, "Two more questions about the little fish.", SUNNY)
    question_rows(slide, QUESTIONS_B, 3)
    hint(slide, "The picture choices help. Let her point instead of speaking if needed.",
         6.35)


def s48_questions_c():
    slide, n = new_slide("🔎 Show Me Where You Found It", "COMPREHENSION", "82–87 min",
                         "Reading Island", AQUA)
    one_task(slide, "Answer it, then show me the line in the story.", AQUA)
    question_rows(slide, QUESTION_C, 5, top_start=1.9, gap=2.16, height=1.92)
    add_round(slide, Inches(0.5), Inches(4.2), Inches(12.35), Inches(2.0), L_AQUA)
    tb(slide, Inches(0.8), Inches(4.35), Inches(11.7), Inches(0.42),
       "🌊 THE HABIT WE ARE BUILDING", size=14, bold=True, color=AQUA)
    steps = [("1", "Answer it"), ("2", "Go back to the story"),
             ("3", "Point to the line"), ("4", "Read it out loud")]
    for i, (num, step) in enumerate(steps):
        left = Inches(0.8 + i * 3.0)
        add_round(slide, left, Inches(4.9), Inches(2.8), Inches(1.1), WHITE)
        add_oval(slide, left + Inches(1.15), Inches(5.05), Inches(0.5), Inches(0.5),
                 AQUA)
        tb(slide, left + Inches(1.15), Inches(5.12), Inches(0.5), Inches(0.36), num,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(5.6), Inches(2.5), Inches(0.36), step,
           size=13, bold=True, color=DEEP, align=PP_ALIGN.CENTER)
    hint(slide, "Going back to the story is not cheating. It is the whole skill.", 6.42)


def s49_island():
    slide, n = new_slide("🏝️ Welcome to Reading Island", "CHALLENGE", "87–90 min",
                         "Reading Island", SUNNY)
    one_task(slide, "Five quick challenges. Then you get your badge.", SUNNY)
    all_final = FINAL_A + FINAL_B
    for i, (num, emoji, kind, _prompt, _content, color, light) in enumerate(all_final):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(3.3), light)
        tb(slide, left, Inches(2.2), Inches(2.3), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.75), Inches(3.05), Inches(0.8), Inches(0.5),
                  color)
        tb(slide, left + Inches(0.75), Inches(3.13), Inches(0.8), Inches(0.36), num,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.68), Inches(2.1), Inches(0.75), kind,
           size=13, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_oval(slide, left + Inches(0.9), Inches(4.5), Inches(0.5), Inches(0.5),
                 WHITE)
        tb(slide, left + Inches(0.9), Inches(4.58), Inches(0.5), Inches(0.36), "☐",
           size=14, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.45), Inches(12.35), Inches(0.8), L_SUNNY)
    tb(slide, Inches(0.8), Inches(5.62), Inches(11.7), Inches(0.46),
       "Every challenge uses something from today. Nothing here is new.", size=16,
       bold=True, color=DEEP, align=PP_ALIGN.CENTER)
    hint(slide, "Hints are allowed. Finishing matters more than doing it alone.", 6.42)


def s50_final_a():
    slide, n = new_slide("🏆 Final Challenge — 1, 2 & 3", "CHALLENGE", "87–90 min",
                         "Reading Island", SUNNY)
    one_task(slide, "Take your time. I am right here.", SUNNY)
    for i, (num, emoji, kind, prompt, content, color, light) in enumerate(FINAL_A):
        top = Inches(1.95 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3), light)
        tb(slide, Inches(0.85), top + Inches(0.32), Inches(0.8), Inches(0.66), emoji,
           size=26, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(1.85), top + Inches(0.2), Inches(2.3), Inches(0.45),
                  color)
        tb(slide, Inches(1.85), top + Inches(0.27), Inches(2.3), Inches(0.34), kind,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(4.4), top + Inches(0.24), Inches(8.0), Inches(0.4), prompt,
           size=13, color=SOFT)
        tb(slide, Inches(4.4), top + Inches(0.64), Inches(8.0), Inches(0.56), content,
           size=30, bold=True, color=color, font="Arial Black")
    hint(slide, "If a word stalls her, break it: /f/ /i/ /n/. Then try the word again.",
         6.55)


def s51_final_b():
    slide, n = new_slide("🏆 Final Challenge — 4 & 5", "CHALLENGE", "87–90 min",
                         "Reading Island", SUNNY)
    one_task(slide, "The last two. Then the badge is yours.", SUNNY)
    for i, (num, emoji, kind, prompt, content, color, light) in enumerate(FINAL_B):
        top = Inches(2.0 + i * 2.1)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.85), light)
        tb(slide, Inches(0.9), top + Inches(0.55), Inches(0.9), Inches(0.75), emoji,
           size=30, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.0), top + Inches(0.3), Inches(2.5), Inches(0.5),
                  color)
        tb(slide, Inches(2.0), top + Inches(0.38), Inches(2.5), Inches(0.36), kind,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(4.75), top + Inches(0.36), Inches(7.7), Inches(0.4), prompt,
           size=13, color=SOFT)
        tb(slide, Inches(4.75), top + Inches(0.8), Inches(7.7), Inches(0.66), content,
           size=30, bold=True, color=color, font="Arial Black")
    hint(slide, "Read challenge 4 together first, then let her do it alone.", 6.42)


def s52_badge():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, DEEP)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), AQUA)
    for x, y, c in [(0.7, 1.4, CORAL), (11.95, 1.35, KELP), (0.85, 3.7, OCTO),
                    (11.85, 3.65, SUNNY)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.8), c)
    tb(slide, Inches(0.7), Inches(0.85), Inches(12), Inches(0.75),
       "🏅 OCEAN READING EXPLORER 🏅", size=32, bold=True, color=AQUA,
       align=PP_ALIGN.CENTER, font="Georgia")
    add_oval(slide, Inches(5.42), Inches(1.7), Inches(2.5), Inches(2.5), AQUA)
    tb(slide, Inches(5.42), Inches(2.25), Inches(2.5), Inches(1.4), "🤿", size=52,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.45), Inches(12), Inches(0.75), "🎉 YOU DID IT!",
       size=42, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(5.25), Inches(12), Inches(0.5),
       "You are an Ocean Reading Explorer. Reading is fun — and you can do it!",
       size=19, color=RGBColor(0x9E, 0xDC, 0xEC), align=PP_ALIGN.CENTER)
    for i, (emoji, place, _j, _t, color, _l) in enumerate(STOPS):
        left = Inches(1.6 + i * 2.15)
        add_round(slide, left, Inches(5.95), Inches(1.95), Inches(0.62), color)
        tb(slide, left, Inches(6.06), Inches(1.95), Inches(0.42), f"{emoji} 🫧", size=13,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "87–90 min", "Reading Island")
    fade(slide)


def s53_can_read():
    slide, n = new_slide("✅ What I Can Read Now", "RECAP", "87–90 min",
                         "Reading Island", KELP)
    one_task(slide, "Look at everything you did in 90 minutes.", KELP)
    for i, (icon, label, detail) in enumerate(CAN_READ):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 0.97)
        add_round(slide, left, top, Inches(5.95), Inches(0.85), L_KELP)
        add_oval(slide, left + Inches(0.24), top + Inches(0.18), Inches(0.5),
                 Inches(0.5), WHITE)
        tb(slide, left + Inches(0.24), top + Inches(0.24), Inches(0.5), Inches(0.4),
           icon, size=14, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.9), top + Inches(0.1), Inches(4.8), Inches(0.4),
           label, size=15, bold=True, color=KELP)
        tb(slide, left + Inches(0.9), top + Inches(0.47), Inches(4.8), Inches(0.34),
           detail, size=11, color=DARK)


def s54_support():
    slide, n = new_slide("🔒 TEACHER ONLY — Support System", "TEACHER ONLY", "",
                         "Support", DARK)
    tb(slide, Inches(0.45), Inches(1.32), Inches(12.4), Inches(0.34),
       "Move between levels freely inside one activity. Never name the level out loud.",
       size=12, bold=True, color=CORAL)
    for i, (icon, name, what, how, color, light) in enumerate(SUPPORT_LEVELS):
        left = Inches(0.45 + i * 4.18)
        add_round(slide, left, Inches(1.78), Inches(3.95), Inches(2.6), light)
        tb(slide, left, Inches(1.96), Inches(3.95), Inches(0.5), icon, size=18,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(2.52), Inches(3.65), Inches(0.5), name,
           size=20, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.25), Inches(3.08), Inches(3.45), Inches(0.45), what,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(3.62), Inches(3.55), Inches(0.58),
                  WHITE)
        tb(slide, left + Inches(0.35), Inches(3.74), Inches(3.25), Inches(0.4), how,
           size=10, color=SOFT, align=PP_ALIGN.CENTER, italic=True)
    add_round(slide, Inches(0.45), Inches(4.55), Inches(12.4), Inches(0.5), L_SUNNY)
    tb(slide, Inches(0.7), Inches(4.63), Inches(11.9), Inches(0.36),
       "POSITIVE LABELS ONLY:   🌊 Ocean Hint   ·   🐠 Easy Catch   ·   🐙 Try Again   "
       "·   ⭐ Bonus Challenge", size=13, bold=True, color=SUNNY)
    add_round(slide, Inches(0.45), Inches(5.2), Inches(6.1), Inches(1.65), L_AQUA)
    tb(slide, Inches(0.7), Inches(5.3), Inches(5.6), Inches(0.34), "🌊 OCEAN HINT BANK",
       size=12, bold=True, color=AQUA)
    bullets(slide, Inches(0.7), Inches(5.65), Inches(5.6), Inches(1.15), HINT_BANK,
            size=10, sp=2)
    add_round(slide, Inches(6.75), Inches(5.2), Inches(6.1), Inches(1.65), L_KELP)
    tb(slide, Inches(7.0), Inches(5.3), Inches(5.6), Inches(0.34),
       "🗣️ SAY THIS INSTEAD OF \"WRONG\"", size=12, bold=True, color=KELP)
    bullets(slide, Inches(7.0), Inches(5.65), Inches(5.6), Inches(1.15), PRAISE,
            size=10, sp=2)


def s55_assessment():
    slide, n = new_slide("🔒 TEACHER ONLY — Reading Skills Checklist", "TEACHER ONLY",
                         "", "Assessment", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Record what she did today, not what you hoped for. Fill this in right after "
       "class.", size=12, bold=True, color=CORAL)
    cols = [(0.45, 4.3, "READING SKILL"), (4.95, 2.6, "Achieved"),
            (7.75, 2.4, "With Help"), (10.35, 2.5, "Needs Practice")]
    header_y = Inches(1.7)
    for left, width, label in cols:
        add_round(slide, Inches(left), header_y, Inches(width), Inches(0.42), DEEP)
        tb(slide, Inches(left), header_y + Inches(0.07), Inches(width), Inches(0.3),
           label, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = header_y + Inches(0.48 + i * 0.36)
        band = WHITE if i % 2 == 0 else L_GREY
        add_round(slide, Inches(0.45), top, Inches(4.3), Inches(0.34), band)
        tb(slide, Inches(0.65), top + Inches(0.03), Inches(3.9), Inches(0.28), skill,
           size=12, bold=True, color=DEEP)
        for left, width, _l in cols[1:]:
            add_round(slide, Inches(left), top, Inches(width), Inches(0.34), band)
            tb(slide, Inches(left), top + Inches(0.01), Inches(width), Inches(0.3), "☐",
               size=13, color=SOFT, align=PP_ALIGN.CENTER)
    for i, label in enumerate(RUBRIC_NOTES):
        top = Inches(5.9 + i * 0.32)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.28), L_SUNNY)
        tb(slide, Inches(0.65), top + Inches(0.02), Inches(12.0), Inches(0.24),
           f"{label}:", size=11, bold=True, color=DEEP)


def s56_answer_key():
    slide, n = new_slide("🔒 TEACHER ONLY — Answer Key", "TEACHER ONLY", "",
                         "Answer key", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Hide this slide before presenting, or keep it on a second screen.", size=12,
       bold=True, color=CORAL)
    add_round(slide, Inches(0.45), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(0.7), Inches(1.86), Inches(5.6), Inches(0.36),
       "🔊 Sounds  ·  🐚 Vowels  ·  🐙 Blending  ·  🎣 Real words", size=13, bold=True,
       color=CORAL)
    bullets(slide, Inches(0.7), Inches(2.28), Inches(5.6), Inches(4.3), [
        "Sound Catch 1–3: fish, sun, cat",
        "Sound Catch 4–6: net, shell, rock",
        "Sound Catch 7–8: wet, hat",
        "Shell Sort 1–3: A, E, I",
        "Shell Sort 4–6: O, E, I",
        "Shell Sort bonus 7–9: I, O, E",
        "Word Builder: CAT, FIN, HOP, NET, SUN, SIT,",
        "     WET, LOG, BIG",
        "Catch the Real Word 1–3: CAT, FIN, NET",
        "Catch the Real Word 4–5: HOP, SUN",
    ], size=11, sp=5)
    add_round(slide, Inches(6.75), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(7.0), Inches(1.86), Inches(5.6), Inches(0.36),
       "🐠 Families  ·  🫧 Sight words  ·  🧩 Sentences  ·  📖 Story", size=13,
       bold=True, color=NAVY)
    bullets(slide, Inches(7.0), Inches(2.28), Inches(5.6), Inches(4.3), [
        "Who Belongs -AT: cat, bat, hat  (fin does not)",
        "Who Belongs -IN: pin, win, fin  (hop does not)",
        "Who Belongs -OP: top, pop, hop  (net does not)",
        "Who Belongs -ET: wet, pet, net  (sat does not)",
        "New words: MAT, TIN, MOP, GET",
        "Bubbles 1–4: see, can, the, my",
        "Bubbles 5–8: I, go, a, to",
        "Sentence puzzles: I see a fish. / The fish is big. /",
        "     I can swim. / I see the shell.",
        "Story Q1–5: Mia · at the sea · red · swims · a shell",
        "Final: FIN · SEE · NET · \"I see a big fish.\" · red",
    ], size=11, sp=5)


BUILDERS = [
    s01_title, s02_explorer, s03_journey, s04_promises, s05_reef_scene, s06_talk,
    s07_sounds, s08_catch_a, s09_catch_b, s10_catch_c, s11_sound_check, s12_vowels,
    s13_sort_a, s14_sort_b, s15_sort_bonus, s16_ocean_words, s17_blending, s18_build_a,
    s19_build_b, s20_build_c, s21_real_a, s22_real_b, s23_dolphin, s24_families,
    s25_belongs_a, s26_belongs_b, s27_new_words, s28_sight_intro, s29_bubbles_a,
    s30_bubbles_b, s31_phrases_a, s32_phrases_b, s33_read_method, s34_sentences_a,
    s35_sentences_b, s36_sentences_c, s37_puzzle_a, s38_puzzle_b, s39_own_sentence,
    s40_story_intro, s41_story_1, s42_story_2, s43_story_3, s44_story_4,
    s45_story_challenge, s46_questions_a, s47_questions_b, s48_questions_c, s49_island,
    s50_final_a, s51_final_b, s52_badge, s53_can_read, s54_support, s55_assessment,
    s56_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

out = r"C:\Users\bhushaja\Downloads\shaip\Grade1_Ocean_Reading_Adventure_90min.pptx"
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
