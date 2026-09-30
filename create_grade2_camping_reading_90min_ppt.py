"""Grade 2 reading lesson - 90 minutes, 70 slides, no speaker notes.

"The Camping Reading Adventure" - the child is a Junior Camp Explorer moving
through the Camp Entrance, the Tent Station, the Flashlight Word Cave, the
Animal Track Trail, the Campfire and the Morning Picnic. Reading is built one
step at a time: look, say, sound, blend, read, understand.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Camp Hint strip, and slides 68-70 hold the support system,
the assessment table and the answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x16, 0x21, 0x2B)
DARK = RGBColor(0x29, 0x36, 0x40)
SOFT = RGBColor(0x79, 0x86, 0x90)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAGE = RGBColor(0xF7, 0xF6, 0xF1)
FOREST = RGBColor(0x2F, 0x7D, 0x4F)
NIGHT = RGBColor(0x2B, 0x4A, 0x7A)
BEAM = RGBColor(0xDD, 0x9E, 0x0B)
TRAIL = RGBColor(0x8A, 0x5A, 0x32)
EMBER = RGBColor(0xD6, 0x4B, 0x24)
PINE = RGBColor(0x14, 0x7D, 0x74)
PLUM = RGBColor(0x7A, 0x4F, 0xA3)
L_FOREST = RGBColor(0xE6, 0xF3, 0xEB)
L_NIGHT = RGBColor(0xE7, 0xEE, 0xF8)
L_BEAM = RGBColor(0xFD, 0xF3, 0xDC)
L_TRAIL = RGBColor(0xF5, 0xEE, 0xE5)
L_EMBER = RGBColor(0xFC, 0xEB, 0xE5)
L_PINE = RGBColor(0xE0, 0xF2, 0xF0)
L_PLUM = RGBColor(0xF1, 0xEB, 0xF8)
L_GREY = RGBColor(0xF1, 0xF2, 0xF1)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 70
_counter = {"n": 0}

# ------------------------------------------------------------------ content

STOPS = [("🏕️", "Camp Entrance", "First sounds", "0–17 min", FOREST, L_FOREST),
         ("⛺", "Tent Station", "Short vowels", "17–26 min", NIGHT, L_NIGHT),
         ("🔦", "Word Cave", "Building words", "26–40 min", BEAM, L_BEAM),
         ("🐾", "Animal Trail", "Word families", "40–52 min", TRAIL, L_TRAIL),
         ("🔥", "Campfire", "Words & sentences", "52–70 min", EMBER, L_EMBER),
         ("🌅", "Morning Picnic", "Story & questions", "70–90 min", PINE, L_PINE)]

MISSION = [("🔤", "Read new words", "Sounds first, then whole words."),
           ("🧩", "Solve camp clues", "Puzzles, sorting and matching games."),
           ("📖", "Read a real story", "A camp story with five short parts."),
           ("🔎", "Explore the camp", "Six stops, one badge at each one.")]

CAMP_SCENE = [("⛺", "tent"), ("🔥", "campfire"), ("🌲", "trees"), ("🎒", "backpack"),
              ("🔦", "flashlight"), ("🗺️", "map")]

TALK_QS = [("🗣️", "What do you see at the camp?"),
           ("🛏️", "Where would you sleep?"),
           ("🎒", "What would you take camping?"),
           ("⭐", "What would you do first?")]

STARTERS = ["I see a ______.", "I would take ______."]

SOUND_CARDS = [("C", "/k/", "CAMP", "🏕️", ["camp", "cup", "can"], FOREST, L_FOREST),
               ("T", "/t/", "TENT", "⛺", ["tent", "top", "ten"], NIGHT, L_NIGHT),
               ("M", "/m/", "MAP", "🗺️", ["map", "mud", "man"], BEAM, L_BEAM),
               ("B", "/b/", "BAG", "🎒", ["bag", "bed", "bug"], PLUM, L_PLUM),
               ("F", "/f/", "FIRE", "🔥", ["fire", "fox", "fan"], EMBER, L_EMBER)]

WHICH_START = [("/t/", [("⛺", "tent"), ("🗺️", "map"), ("🎒", "bag")]),
               ("/m/", [("🔥", "fire"), ("🗺️", "map"), ("⛺", "tent")]),
               ("/b/", [("🎒", "bag"), ("🏕️", "camp"), ("🔦", "flashlight")])]

BACKPACK_A = [("/b/", [("🎒", "bag"), ("🗺️", "map"), ("🔥", "fire")]),
              ("/m/", [("🥾", "boots"), ("🗺️", "map"), ("🪵", "log")]),
              ("/t/", [("🔦", "flashlight"), ("⛺", "tent"), ("🎒", "bag")])]
BACKPACK_B = [("/k/", [("🏕️", "camp"), ("🔥", "fire"), ("🥾", "boots")]),
              ("/f/", [("🗺️", "map"), ("🔥", "fire"), ("⛺", "tent")]),
              ("/l/", [("🪵", "log"), ("🎒", "bag"), ("🏕️", "camp")])]

VOWELS = [("A", "/a/", "m-aaa-p", [("MAP", "🗺️"), ("BAG", "🎒"), ("CAMP", "🏕️")],
           EMBER, L_EMBER),
          ("E", "/e/", "b-eee-d", [("RED", "🔴"), ("BED", "🛏️"), ("TENT", "⛺")],
           FOREST, L_FOREST),
          ("I", "/i/", "s-iii-t", [("SIT", "🪑"), ("BIG", "🐘"), ("HIT", "🥎")],
           NIGHT, L_NIGHT),
          ("O", "/o/", "h-ooo-t", [("HOT", "🔥"), ("LOG", "🪵"), ("FOX", "🦊")],
           PLUM, L_PLUM)]

SORT = [("🗺️", "MAP", "/m/ /a/ /p/", ["A", "E", "I"]),
        ("🔴", "RED", "/r/ /e/ /d/", ["E", "O", "A"]),
        ("🪑", "SIT", "/s/ /i/ /t/", ["I", "A", "O"]),
        ("🔥", "HOT", "/h/ /o/ /t/", ["O", "E", "I"])]

MISSING_VOWEL = [("🗺️", "M _ P", ["A", "E", "I"], "A"),
                 ("🛏️", "B _ D", ["E", "A", "O"], "E"),
                 ("🪵", "L _ G", ["O", "I", "E"], "O"),
                 ("🏕️", "C _ M P", ["A", "O", "U"], "A")]

CAMP_WORDS = [("tent", "⛺"), ("camp", "🏕️"), ("map", "🗺️"), ("fire", "🔥"),
              ("torch", "🔦"), ("forest", "🌲"), ("trail", "🛤️"), ("boots", "🥾"),
              ("backpack", "🎒"), ("river", "🏞️"), ("animal", "🦊"),
              ("tracks", "🐾"), ("bed", "🛏️"), ("picnic", "🧺")]

SIX_STEPS = [("1", "LOOK", "Look at the picture.", "🗺️", FOREST, L_FOREST),
             ("2", "SAY", "Say the word out loud.", "map", NIGHT, L_NIGHT),
             ("3", "SOUND", "Find each sound.", "/m/ /a/ /p/", BEAM, L_BEAM),
             ("4", "BLEND", "Push them together.", "m-a-p", TRAIL, L_TRAIL),
             ("5", "READ", "Read the whole word.", "MAP", EMBER, L_EMBER),
             ("6", "GET IT", "Match it to a picture.", "🗺️ = MAP", PINE, L_PINE)]

BUILD_A = [("🏕️", "CAMP", ["C", "A", "M", "P"], ["/k/", "/a/", "/m/", "/p/"],
            FOREST, L_FOREST),
           ("⛺", "TENT", ["T", "E", "N", "T"], ["/t/", "/e/", "/n/", "/t/"],
            NIGHT, L_NIGHT)]
BUILD_B = [("🗺️", "MAP", ["M", "A", "P"], ["/m/", "/a/", "/p/"], BEAM, L_BEAM),
           ("🎒", "BAG", ["B", "A", "G"], ["/b/", "/a/", "/g/"], PLUM, L_PLUM)]
BUILD_C = [("🪵", "LOG", ["L", "O", "G"], ["/l/", "/o/", "/g/"], TRAIL, L_TRAIL),
           ("🦊", "FOX", ["F", "O", "X"], ["/f/", "/o/", "/x/"], EMBER, L_EMBER)]

MISSING_LETTER = [("🏕️", "C _ M P", ["A", "E", "I"], "A"),
                  ("🎒", "B _ G", ["A", "O", "U"], "A"),
                  ("🦊", "F _ X", ["O", "A", "E"], "O"),
                  ("☀️", "S _ N", ["U", "A", "I"], "U")]

SCRAMBLE = [("🗺️", ["P", "A", "M"], "MAP"), ("⛺", ["T", "N", "E", "T"], "TENT"),
            ("🪵", ["G", "O", "L"], "LOG"), ("🎒", ["A", "G", "B"], "BAG")]

NO_CARDS = [("CAMP", "🏕️"), ("TENT", "⛺"), ("MAP", "🗺️"),
            ("BAG", "🎒"), ("LOG", "🪵"), ("FOX", "🦊")]

FAMILIES = [("-AT", ["cat", "hat", "bat"], EMBER, L_EMBER),
            ("-OG", ["dog", "log", "fog"], FOREST, L_FOREST),
            ("-OX", ["fox", "box"], NIGHT, L_NIGHT),
            ("-UN", ["sun", "run", "fun"], PLUM, L_PLUM)]

TRACK_MATCH = [("FOX", [("🦊", "fox"), ("🐕", "dog"), ("☀️", "sun")]),
               ("DOG", [("☀️", "sun"), ("🐕", "dog"), ("🦊", "fox")]),
               ("SUN", [("🪵", "log"), ("🦊", "fox"), ("☀️", "sun")]),
               ("LOG", [("🪵", "log"), ("🐈", "cat"), ("🎒", "bag")])]

FAMILY_A = [("-AT", ["CAT", "DOG", "HAT", "BAT"], EMBER, L_EMBER),
            ("-OG", ["DOG", "LOG", "FOX", "FOG"], FOREST, L_FOREST)]
FAMILY_B = [("-OX", ["FOX", "BOX", "BAG", "OX"], NIGHT, L_NIGHT),
            ("-UN", ["SUN", "RUN", "FUN", "HAT"], PLUM, L_PLUM)]

NEW_TRACK = [("CAT", "🐈", ["HAT", "DOG", "SUN"], EMBER, L_EMBER),
             ("LOG", "🪵", ["MAP", "DOG", "TENT"], FOREST, L_FOREST),
             ("SUN", "☀️", ["RUN", "BOX", "BED"], NIGHT, L_NIGHT)]

FREEZE = [("🔥", "Fire!", "Make a fire shape with your hands."),
          ("🏕️", "Camp!", "Make a tent shape over your head."),
          ("🔦", "Flashlight!", "Pretend to hold a flashlight."),
          ("🧊", "Freeze!", "Stop and hold still."),
          ("🐾", "Tracks!", "March in place, two steps."),
          ("🌲", "Tree!", "Stand tall with your arms up.")]

CAMPER_SAYS = [("👏", "Camper says clap twice.", "move"),
               ("👉", "Camper says point to TENT.", "read"),
               ("📖", "Camper says read FOX.", "read"),
               ("🤫", "Camper says whisper MAP.", "read"),
               ("🦶", "Camper says stand on one foot.", "move"),
               ("📣", "Camper says read BIG out loud.", "read")]

CALM = [("🌬️", "Breathe in", "count to three"),
        ("😮‍💨", "Breathe out", "count to three"),
        ("🙌", "Shake your hands", "five seconds"),
        ("🪑", "Sit tall", "ready for the campfire")]

SIGHT_1 = [("I", "I can camp.", "🙋"), ("see", "I see a tent.", "👀"),
           ("a", "a big log", "🪵")]
SIGHT_2 = [("the", "the hot fire", "🔥"), ("my", "my red bag", "🎒"),
           ("can", "I can go.", "🥾")]

LIGHT_WORD = [("I", ["a", "I", "see"]), ("see", ["see", "the", "my"]),
              ("the", ["my", "the", "can"]), ("can", ["I", "can", "go"])]

SENT_A = [("I see a tent.", "⛺", NIGHT, L_NIGHT),
          ("I see a campfire.", "🔥", EMBER, L_EMBER)]
SENT_B = [("My bag is big.", "🎒", PLUM, L_PLUM),
          ("We can go to camp.", "🏕️", FOREST, L_FOREST)]

PHRASES = [("a big tent", "⛺"), ("my red bag", "🎒"), ("the hot fire", "🔥"),
           ("in the forest", "🌲"), ("on the map", "🗺️"), ("I can go", "🥾")]

READ_STEPS = [("STEP 1", "I read", "Teacher reads it first.", EMBER),
              ("STEP 2", "We read", "Teacher and child together.", BEAM),
              ("STEP 3", "You read", "Child reads it alone — if ready.", FOREST)]

PUZZLES_A = [(["see", "I", "tent", "a"], "I see a tent.", "⛺", NIGHT, L_NIGHT),
             (["is", "The", "big", "tent"], "The tent is big.", "🏕️", FOREST,
              L_FOREST)]
PUZZLES_B = [(["can", "I", "run"], "I can run.", "🏃", BEAM, L_BEAM),
             (["my", "is", "bag", "red"], "My bag is red.", "🎒", PLUM, L_PLUM)]

PICTURE_MATCH = [("The fox can run.",
                  [("🦊", "a fox running"), ("🐕", "a sleeping dog"),
                   ("⛺", "a tent")]),
                 ("The tent is big.",
                  [("🔥", "a campfire"), ("⛺", "a big tent"), ("🗺️", "a map")]),
                 ("I see a log.",
                  [("🪵", "a log"), ("🎒", "a bag"), ("☀️", "the sun")])]

SILLY = [("⛺", ["The tent is red.", "The tent is eating."]),
         ("🦊", ["The fox can run.", "The fox can read."]),
         ("🔥", ["The fire is hot.", "The fire is wet."])]

OWN_SENTENCE = [("⛺", "tent"), ("🔥", "fire"), ("🎒", "bag"), ("🗺️", "map")]

STORY_WATCH = [("tent", "⛺"), ("bag", "🎒"), ("log", "🪵"), ("light", "🔦"),
               ("dad", "👨"), ("camp", "🏕️")]

STORY = [
    ("Part 1", "🏕️", FOREST, L_FOREST,
     ["Sam is camping with his dad.", "They have a small green tent.",
      "Sam has a flashlight in his bag.", "The sun goes down.",
      "The camp gets dark.", "Sam is happy to camp."]),
    ("Part 2", "🎒", NIGHT, L_NIGHT,
     ["At night, Sam looks in his bag.", "He wants his flashlight.",
      "The bag is empty.", "He cannot find it.",
      "\"Where can it be?\" says Sam.", "Sam is sad."]),
    ("Part 3", "🔎", BEAM, L_BEAM,
     ["Sam looks near the tent.", "He looks under his bed.",
      "He looks by the big log.", "It is not in the tent.",
      "Then he sees a small light.", "The light is near the log."]),
    ("Part 4", "🔦", EMBER, L_EMBER,
     ["Sam walks to the log.", "There is the flashlight!",
      "Sam smiles a big smile.", "He takes it back to the tent.",
      "Dad says, \"Good looking, Sam!\"", "The flashlight is not lost now."]),
    ("Part 5", "📖", PLUM, L_PLUM,
     ["Sam turns on the light.", "Now the tent is bright.",
      "Sam reads a book in bed.", "Dad reads with him.",
      "It is a good camp night.", "Sam is a good camper."]),
]

EASY_LINES = ["Sam is sad.", "The bag is empty.", "Sam walks to the log.",
              "There is the flashlight!", "Now the tent is bright.",
              "Sam is a good camper."]

QUESTIONS_A = [("👦", "Who is in the story?", ["Sam", "Mia", "Ben"], "Part 1"),
               ("🔦", "What did Sam lose?", ["A hat", "His flashlight", "His bag"],
                "Part 2")]
QUESTIONS_B = [("🪵", "Where did Sam find it?",
                ["Near a log", "In the river", "In the car"], "Part 3"),
               ("⛺", "What did Sam do next?",
                ["Went home", "Took it to the tent", "Threw it away"], "Part 4")]

EVIDENCE = [("Who is camping with Sam?", "\"Sam is camping with his dad.\"",
             "Part 1", FOREST, L_FOREST),
            ("What was in the bag?", "\"The bag is empty.\"", "Part 2", NIGHT,
             L_NIGHT),
            ("Where was the light?", "\"The light is near the log.\"", "Part 3",
             BEAM, L_BEAM)]

FINAL_WORDS = [("CAMP", "🏕️", FOREST, L_FOREST), ("TENT", "⛺", NIGHT, L_NIGHT),
               ("MAP", "🗺️", BEAM, L_BEAM), ("FOX", "🦊", EMBER, L_EMBER),
               ("BAG", "🎒", PLUM, L_PLUM), ("LOG", "🪵", TRAIL, L_TRAIL)]

FINAL_SENTENCES = [("I see a tent.", "⛺", NIGHT, L_NIGHT),
                   ("The fox can run.", "🦊", EMBER, L_EMBER),
                   ("My bag is red.", "🎒", PLUM, L_PLUM)]

CAN_READ = [("👂", "Beginning sounds", "c  t  m  b  f"),
            ("🔦", "Short vowels", "a  e  i  o"),
            ("🧩", "Blending", "/m/ /a/ /p/ → MAP"),
            ("🏕️", "CVC words", "camp  tent  map  bag"),
            ("🐾", "Word families", "-at  -og  -ox  -un"),
            ("🔥", "Sight words", "I  see  a  the  my  can"),
            ("🔗", "Short phrases", "a big tent"),
            ("📕", "Sentences", "The fox can run."),
            ("📖", "A whole story", "The Lost Flashlight"),
            ("🔎", "Finding the answer", "pointing to the proof")]

CHAMPION_LINES = [("🔤", "You read new words!"), ("🧩", "You blended sounds!"),
                  ("📕", "You built sentences!"), ("📖", "You read a story!"),
                  ("🔎", "You answered questions!")]

HINT_LADDER = [("HINT 1", "Look at the first letter.", EMBER),
               ("HINT 2", "Say the sounds.", BEAM),
               ("HINT 3", "Let's blend them.", NIGHT),
               ("HINT 4", "Let's read it together.", FOREST),
               ("HINT 5", "I read it, then you repeat.", PLUM)]

SUPPORT_LEVELS = [("🟢", "CAMP HINT", "Picture + sound + teacher support",
                   "Say the word first, then he repeats it.", FOREST, L_FOREST),
                  ("🟡", "CAMP MISSION", "He reads with limited help",
                   "Give the first sound only, then go quiet.", BEAM, L_BEAM),
                  ("⭐", "TRAIL CHALLENGE", "He reads alone or makes a sentence",
                   "Ask for a sentence using the word.", PLUM, L_PLUM)]

HINT_BANK = ["Look at the first letter.", "Say the sounds slowly.",
             "Blend the sounds together.", "Look at the picture.",
             "Read it one more time.", "Look back at the story."]

PRAISE = ["\"Nice try!\"", "\"Let's look again.\"", "\"You found the first sound!\"",
          "\"Great blending!\"", "\"You figured that one out!\"",
          "\"Let's read it together.\""]

RUBRIC_SKILLS = ["Beginning sounds", "Short vowels", "CVC words", "Blending",
                 "Word families", "Sight words", "Phrase reading",
                 "Sentence reading", "Story reading", "Comprehension"]

NEXT_CLASS = ["Words to practice", "Sight words to practice",
              "Reading skill for next class"]

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
             RGBColor(0xE1, 0xE4, 0xE0))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), FOREST)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), INK)
    msg = "🏕️ The Camping Reading Adventure  |  Grade 2  |  90 min"
    if stop:
        msg += f"  |  {stop}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=10,
       color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, stop, accent=FOREST, bg=PAGE):
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


def one_task(slide, text, color=EMBER, top=1.34):
    """States the single job of this slide in words the child understands."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text,
              size=16, bold=True, color=color)


def hint(slide, text, top=6.42, label="🔦 CAMP HINT", fill=L_BEAM, color=BEAM):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.5), Inches(0.34), label,
       size=12, bold=True, color=color)
    tb(slide, Inches(3.4), Inches(top + 0.05), Inches(9.2), Inches(0.36), text,
       size=13, bold=True, color=INK)


def i_we_you(slide, top=6.42):
    steps = [("I READ", "Teacher", EMBER), ("WE READ", "Together", BEAM),
             ("YOU READ", "You!", FOREST)]
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
        add_oval(slide, Inches(0.75), top + Inches(height / 2 - 0.25), Inches(0.5),
                 Inches(0.5), accent)
        tb(slide, Inches(0.75), top + Inches(height / 2 - 0.19), Inches(0.5),
           Inches(0.4), str(start_index + i), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tops.append(top)
    return tops


def sound_rows(slide, items, start_index, accent, label="STARTS WITH"):
    """A target sound, then three pictures to choose from."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=1.9,
                         gap=1.5)
    for (sound, options), top in zip(items, tops):
        add_round(slide, Inches(1.4), top + Inches(0.3), Inches(2.5), Inches(0.74),
                  accent)
        tb(slide, Inches(1.4), top + Inches(0.34), Inches(2.5), Inches(0.28), label,
           size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.4), top + Inches(0.6), Inches(2.5), Inches(0.42), sound,
           size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, (emoji, name) in enumerate(options):
            left = Inches(4.2 + j * 2.82)
            add_round(slide, left, top + Inches(0.16), Inches(2.6), Inches(1.02),
                      L_GREY)
            tb(slide, left, top + Inches(0.2), Inches(2.6), Inches(0.58), emoji,
               size=26, align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.78), Inches(2.6), Inches(0.36), name,
               size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)


def vowel_rows(slide, items, start_index, accent=BEAM):
    """Flashlight sort: word plus sounds, then three beams to choose from."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=1.9,
                         gap=1.12, height=1.0)
    for (emoji, word, sounds, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.18), Inches(0.9), Inches(0.62), emoji,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.4), top + Inches(0.16), Inches(2.0), Inches(0.5), word,
           size=24, bold=True, color=INK, font="Arial Black")
        tb(slide, Inches(2.4), top + Inches(0.62), Inches(2.0), Inches(0.32),
           sounds, size=12, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(4.8 + j * 2.66)
            add_round(slide, left, top + Inches(0.19), Inches(2.45), Inches(0.62),
                      L_BEAM)
            tb(slide, left + Inches(0.15), top + Inches(0.29), Inches(0.5),
               Inches(0.4), "🔦", size=13, align=PP_ALIGN.CENTER)
            tb(slide, left + Inches(0.75), top + Inches(0.26), Inches(1.5),
               Inches(0.44), opt, size=22, bold=True, color=BEAM,
               align=PP_ALIGN.CENTER, font="Arial Black")


def letter_rows(slide, items, start_index, accent):
    """Missing letter: a gapped word, then three letters to try."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=1.9,
                         gap=1.14, height=1.02)
    for (emoji, pattern, options, _answer), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.18), Inches(0.9), Inches(0.66), emoji,
           size=24, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.5), top + Inches(0.18), Inches(2.6), Inches(0.66),
                  L_GREY)
        tb(slide, Inches(2.5), top + Inches(0.26), Inches(2.6), Inches(0.5),
           pattern, size=26, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(5.6 + j * 2.36)
            add_round(slide, left, top + Inches(0.2), Inches(2.15), Inches(0.62),
                      WHITE)
            tb(slide, left, top + Inches(0.28), Inches(2.15), Inches(0.46), opt,
               size=22, bold=True, color=accent, align=PP_ALIGN.CENTER,
               font="Arial Black")


def build_rows(slide, items, top_start=2.0, gap=2.15, height=1.9):
    """Letter cards, the sounds, then the whole word on a camp sign."""
    for i, (emoji, word, letters, sounds, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), light)
        tb(slide, Inches(0.75), top + Inches(0.55), Inches(1.0), Inches(0.8), emoji,
           size=34, align=PP_ALIGN.CENTER)
        for j, letter in enumerate(letters):
            left = Inches(1.9 + j * 1.25)
            add_round(slide, left, top + Inches(0.5), Inches(1.1), Inches(0.9),
                      WHITE)
            tb(slide, left, top + Inches(0.6), Inches(1.1), Inches(0.66), letter,
               size=30, bold=True, color=color, align=PP_ALIGN.CENTER,
               font="Arial Black")
            if j < len(letters) - 1:
                tb(slide, left + Inches(1.1), top + Inches(0.78), Inches(0.15),
                   Inches(0.42), "+", size=15, bold=True, color=SOFT,
                   align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.9), top + Inches(0.75), Inches(2.3), Inches(0.5),
           " ".join(sounds), size=15, bold=True, color=SOFT)
        tb(slide, Inches(9.25), top + Inches(0.7), Inches(0.5), Inches(0.5), "→",
           size=22, bold=True, color=color)
        add_round(slide, Inches(9.9), top + Inches(0.5), Inches(2.7), Inches(0.9),
                  WHITE)
        tb(slide, Inches(9.9), top + Inches(0.6), Inches(2.7), Inches(0.66), word,
           size=30, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, Inches(1.9), top + Inches(1.48), Inches(10.2), Inches(0.32),
           "Touch each letter and say its sound, then say the whole word fast.",
           size=11, color=SOFT)


def scramble_rows(slide, items, start_index, accent=BEAM):
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=1.9,
                         gap=1.14, height=1.02)
    for (emoji, letters, _answer), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.18), Inches(0.9), Inches(0.66), emoji,
           size=24, align=PP_ALIGN.CENTER)
        for j, letter in enumerate(letters):
            left = Inches(2.5 + j * 1.25)
            add_round(slide, left, top + Inches(0.2), Inches(1.1), Inches(0.62),
                      L_BEAM)
            tb(slide, left, top + Inches(0.28), Inches(1.1), Inches(0.46), letter,
               size=24, bold=True, color=accent, align=PP_ALIGN.CENTER,
               font="Arial Black")
        tb(slide, Inches(7.6), top + Inches(0.28), Inches(0.5), Inches(0.44), "→",
           size=20, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(8.3), top + Inches(0.2), Inches(4.25), Inches(0.62),
                  WHITE)
        tb(slide, Inches(8.55), top + Inches(0.3), Inches(3.8), Inches(0.44),
           "your word:  ______________", size=13, color=SOFT)


def family_rows(slide, items, start_index):
    """One family sign, then four track cards to check."""
    tops = numbered_rows(slide, len(items), start_index, TRAIL, top_start=2.0,
                         gap=2.15, height=1.9)
    for (family, words, color, light), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.5), Inches(0.9), Inches(0.7), "🐾",
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.4), top + Inches(0.6), Inches(1.9), Inches(0.7),
                  color)
        tb(slide, Inches(2.4), top + Inches(0.71), Inches(1.9), Inches(0.5), family,
           size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, word in enumerate(words):
            left = Inches(4.65 + j * 2.0)
            add_round(slide, left, top + Inches(0.55), Inches(1.85), Inches(0.8),
                      light)
            tb(slide, left, top + Inches(0.67), Inches(1.85), Inches(0.56), word,
               size=22, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")
        tb(slide, Inches(4.65), top + Inches(1.42), Inches(7.9), Inches(0.3),
           "Read all four out loud, then cross out the one that does not belong.",
           size=11, color=SOFT)


def match_rows(slide, items, start_index, accent=TRAIL):
    """A written word, then three pictures to match it to."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=1.9,
                         gap=1.2, height=1.08)
    for (word, options), top in zip(items, tops):
        add_round(slide, Inches(1.35), top + Inches(0.22), Inches(2.3),
                  Inches(0.64), accent)
        tb(slide, Inches(1.35), top + Inches(0.31), Inches(2.3), Inches(0.48), word,
           size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, (emoji, _name) in enumerate(options):
            left = Inches(4.0 + j * 2.9)
            add_round(slide, left, top + Inches(0.14), Inches(2.7), Inches(0.8),
                      L_TRAIL)
            tb(slide, left, top + Inches(0.2), Inches(2.7), Inches(0.66), emoji,
               size=28, align=PP_ALIGN.CENTER)


def find_rows(slide, items, start_index, accent=EMBER, top_start=1.95):
    """Light the Word: find one sight word among three."""
    for i, (target, options) in enumerate(items):
        top = Inches(top_start + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), WHITE)
        add_oval(slide, Inches(0.78), top + Inches(0.22), Inches(0.52), Inches(0.52),
                 accent)
        tb(slide, Inches(0.78), top + Inches(0.28), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.27), Inches(1.0), Inches(0.44),
           "FIND:", size=13, bold=True, color=SOFT)
        add_round(slide, Inches(2.55), top + Inches(0.2), Inches(1.9), Inches(0.58),
                  accent)
        tb(slide, Inches(2.55), top + Inches(0.29), Inches(1.9), Inches(0.42),
           target, size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(4.85 + j * 2.6)
            add_round(slide, left, top + Inches(0.14), Inches(2.4), Inches(0.7),
                      L_EMBER)
            tb(slide, left + Inches(0.15), top + Inches(0.26), Inches(0.5),
               Inches(0.44), "🔦", size=13, align=PP_ALIGN.CENTER)
            tb(slide, left + Inches(0.75), top + Inches(0.24), Inches(1.5),
               Inches(0.48), opt, size=20, bold=True, color=INK,
               align=PP_ALIGN.CENTER, font="Arial Black")


def phrase_grid(slide, items, top=1.95):
    for i, (phrase, emoji) in enumerate(items):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        t = Inches(top + row * 1.5)
        add_round(slide, left, t, Inches(5.95), Inches(1.3), L_EMBER)
        tb(slide, left + Inches(0.3), t + Inches(0.3), Inches(1.0), Inches(0.72),
           emoji, size=28, align=PP_ALIGN.CENTER)
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
    """Shuffled word cards plus a line to write the sentence on."""
    tops = numbered_rows(slide, len(items), start_index, BEAM, top_start=2.0,
                         gap=2.15, height=1.9)
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
                 PINE)
        tb(slide, Inches(0.75), top + Inches(0.76), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.45), top + Inches(0.6), Inches(1.0), Inches(0.72), emoji,
           size=30, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(0.5), Inches(3.5), Inches(0.66),
           question, size=19, bold=True, color=INK)
        tb(slide, Inches(2.6), top + Inches(1.18), Inches(3.5), Inches(0.36),
           f"the answer is in {where}", size=11, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(6.3 + j * 2.2)
            add_round(slide, left, top + Inches(0.5), Inches(2.0), Inches(0.92),
                      L_PINE)
            tb(slide, left, top + Inches(0.66), Inches(2.0), Inches(0.62), opt,
               size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)


def story_slide(part, timing):
    """One story page: a big picture panel beside six short lines."""
    label, emoji, color, light = part[0], part[1], part[2], part[3]
    lines = part[4]
    slide, n = new_slide(f"📖 The Lost Flashlight — {label}", "STORY", timing,
                         "Morning Picnic", color)
    add_round(slide, Inches(0.5), Inches(1.4), Inches(4.3), Inches(4.75), light)
    tb(slide, Inches(0.5), Inches(2.6), Inches(4.3), Inches(1.7), emoji, size=88,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.55), Inches(4.5), Inches(2.2), Inches(0.55), WHITE)
    tb(slide, Inches(1.55), Inches(4.61), Inches(2.2), Inches(0.4), label.upper(),
       size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(5.1), Inches(1.4), Inches(7.75), Inches(4.75), WHITE)
    for i, line in enumerate(lines):
        tb(slide, Inches(5.5), Inches(1.62 + i * 0.75), Inches(7.1), Inches(0.62),
           line, size=21, bold=True, color=INK)
    i_we_you(slide)
    return slide, n


def badge_slide(title, timing, stop, emoji, badge_label, sub, lines, next_line,
                accent=FOREST, light=L_FOREST):
    """Stop-complete slide: a big badge beside what was just mastered."""
    slide, n = new_slide(title, "STOP COMPLETE", timing, stop, accent)
    one_task(slide, sub, accent)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.3), Inches(4.4), light)
    tb(slide, Inches(0.5), Inches(2.6), Inches(5.3), Inches(1.7), emoji, size=88,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.3), Inches(0.7), badge_label,
       size=25, bold=True, color=accent, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.7), Inches(0.6),
       "⭐ explorer badge earned", size=15, color=SOFT, align=PP_ALIGN.CENTER)
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.15), top, Inches(6.7), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.45), top + Inches(0.2), Inches(0.56), Inches(0.56),
                 WHITE)
        tb(slide, Inches(6.45), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.25), top + Inches(0.24), Inches(5.3), Inches(0.5), line,
           size=15, bold=True, color=INK)
    hint(slide, next_line, 6.45, "➡️ NEXT STOP", L_BEAM, BEAM)
    return slide, n


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), FOREST)
    for x, y, c in [(0.55, 5.55, EMBER), (12.1, 5.5, BEAM)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(0.85), Inches(12), Inches(1.0), "🏕️", size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.85), Inches(12), Inches(0.9),
       "CAMPING READING ADVENTURE", size=42, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.8), Inches(12), Inches(0.5),
       "\"Become a Junior Camp Explorer!\"", size=21,
       color=RGBColor(0xA6, 0xD6, 0xB8), align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, label) in enumerate(CAMP_SCENE):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(3.5), Inches(1.65), Inches(1.35),
                  RGBColor(0x21, 0x2F, 0x3A))
        tb(slide, left, Inches(3.68), Inches(1.65), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.33), Inches(1.45), Inches(0.4),
           label, size=10, color=RGBColor(0xB2, 0xC3, 0xCF),
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.3), Inches(5.15), Inches(6.7), Inches(1.15), FOREST)
    tb(slide, Inches(3.5), Inches(5.4), Inches(6.3), Inches(0.7),
       "Grade 2  •  90 Minutes  •  Reading", size=19, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "0–8 min", "Camp Entrance")
    fade(slide)


def s02_mission():
    slide, n = new_slide("⭐ Your Mission Today", "WELCOME", "0–8 min",
                         "Camp Entrance", FOREST)
    one_task(slide, "Today you are a Junior Camp Explorer.", FOREST)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4), L_FOREST)
    tb(slide, Inches(0.5), Inches(2.4), Inches(5.4), Inches(1.7), "⭐", size=88,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.3), Inches(5.4), Inches(0.75),
       "JUNIOR CAMP EXPLORER", size=24, bold=True, color=FOREST,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.2), Inches(4.8), Inches(0.7),
       "that is you, all lesson long", size=15, color=SOFT,
       align=PP_ALIGN.CENTER)
    for i, (icon, name, detail) in enumerate(MISSION):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(0.96), L_BEAM)
        add_oval(slide, Inches(6.5), top + Inches(0.2), Inches(0.56), Inches(0.56),
                 WHITE)
        tb(slide, Inches(6.5), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.3), top + Inches(0.1), Inches(5.3), Inches(0.44), name,
           size=17, bold=True, color=INK)
        tb(slide, Inches(7.3), top + Inches(0.53), Inches(5.3), Inches(0.38),
           detail, size=12, color=SOFT)
    hint(slide, "Read the four missions out loud together before you set off.",
         6.45)


def s03_route():
    slide, n = new_slide("🗺️ Our Camp Route", "MAP", "0–8 min", "Camp Entrance",
                         BEAM)
    one_task(slide, "Six stops. One reading job at each one.", BEAM)
    for i, (emoji, place, job, when, color, light) in enumerate(STOPS):
        left = Inches(0.5 + i * 2.08)
        add_round(slide, left, Inches(1.95), Inches(1.9), Inches(3.5), light)
        tb(slide, left, Inches(2.18), Inches(1.9), Inches(0.8), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.05), Inches(3.05), Inches(1.8), Inches(0.9),
           place, size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.95), Inches(1.7), Inches(0.6), job,
           size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(4.7), Inches(1.6),
                  Inches(0.42), WHITE)
        tb(slide, left + Inches(0.15), Inches(4.74), Inches(1.6), Inches(0.34),
           when, size=9, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.65), Inches(12.35), Inches(0.9),
              L_FOREST)
    tb(slide, Inches(0.8), Inches(5.9), Inches(11.7), Inches(0.44),
       "🏕️  →  ⛺  →  🔦  →  🐾  →  🔥  →  🌅    Finish all six and you are a "
       "CAMP READING CHAMPION!", size=15, bold=True, color=INK,
       align=PP_ALIGN.CENTER)


def s04_warmup():
    slide, n = new_slide("🗣️ Warm-Up — Look at the Camp", "SPEAKING", "0–8 min",
                         "Camp Entrance", PINE)
    one_task(slide, "Just look and talk. No reading yet!", PINE)
    for i, (emoji, label) in enumerate(CAMP_SCENE):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.85 + row * 1.6)
        add_round(slide, left, top, Inches(3.9), Inches(1.42), L_PINE)
        tb(slide, left + Inches(0.2), top + Inches(0.3), Inches(1.1), Inches(0.82),
           emoji, size=32, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.5), top + Inches(0.42), Inches(2.1),
                  Inches(0.58), WHITE)
        tb(slide, left + Inches(1.5), top + Inches(0.53), Inches(2.1),
           Inches(0.42), label, size=16, bold=True, color=PINE,
           align=PP_ALIGN.CENTER)
    for i, (icon, question) in enumerate(TALK_QS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(5.1 + row * 0.62)
        tb(slide, left + Inches(0.1), top, Inches(0.4), Inches(0.44), icon,
           size=13, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.65), top, Inches(5.2), Inches(0.44), question,
           size=16, bold=True, color=INK)
    for i, starter in enumerate(STARTERS):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(6.4), Inches(5.95), Inches(0.5), L_BEAM)
        tb(slide, left + Inches(0.3), Inches(6.47), Inches(5.4), Inches(0.38),
           starter, size=18, bold=True, color=INK)


def sound_slide(index):
    letter, sound, word, emoji, examples, color, light = SOUND_CARDS[index]
    slide, n = new_slide(f"🔤 Camp Sound Scout — {letter} says {sound}", "LEARN",
                         "8–17 min", "Camp Entrance", color)
    one_task(slide, "I say the sound. You say it back to me.", color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.3), Inches(4.4), light)
    add_round(slide, Inches(1.85), Inches(2.1), Inches(2.6), Inches(1.9), color)
    tb(slide, Inches(1.85), Inches(2.4), Inches(2.6), Inches(1.35), letter,
       size=80, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.5), Inches(4.2), Inches(5.3), Inches(0.7), f"says {sound}",
       size=28, bold=True, color=color, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.1), Inches(4.7), Inches(0.7),
       "stretch it — do not say the letter name", size=13, color=SOFT,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(6.15), Inches(1.85), Inches(6.7), Inches(2.85), WHITE)
    tb(slide, Inches(6.15), Inches(2.0), Inches(6.7), Inches(1.1), emoji, size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(6.15), Inches(3.25), Inches(6.7), Inches(1.0), word, size=54,
       bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(6.4), Inches(4.85), Inches(6.2), Inches(0.4),
       "MORE WORDS WITH THIS SOUND", size=12, bold=True, color=color)
    for i, ex in enumerate(examples):
        left = Inches(6.15 + i * 2.26)
        add_round(slide, left, Inches(5.3), Inches(2.1), Inches(0.72), L_GREY)
        tb(slide, left, Inches(5.42), Inches(2.1), Inches(0.5), ex, size=22,
           bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, f"Say {sound} three times, then say {word}. He copies you each "
                f"time.", 6.45)


def s05_sound_c():
    sound_slide(0)


def s06_sound_t():
    sound_slide(1)


def s07_sound_m():
    sound_slide(2)


def s08_sound_b():
    sound_slide(3)


def s09_sound_f():
    sound_slide(4)


def s10_which_start():
    slide, n = new_slide("👂 Which Word Starts With That Sound?", "PRACTICE",
                         "8–17 min", "Camp Entrance", FOREST)
    one_task(slide, "Say all three names out loud, then pick one.", FOREST)
    sound_rows(slide, WHICH_START, 1, FOREST)
    hint(slide, "Stretch the first sound of each picture: t-t-tent.", 6.45)


def s11_backpack_a():
    slide, n = new_slide("🎒 Game: Pack the Backpack", "GAME", "8–17 min",
                         "Camp Entrance", FOREST)
    one_task(slide, "Find the thing that starts with my sound. Pack it!", FOREST)
    sound_rows(slide, BACKPACK_A, 1, FOREST, "PACK SOMETHING WITH")
    hint(slide, "Too many choices? Cover one with your hand and leave only two.",
         6.45)


def s12_backpack_b():
    slide, n = new_slide("🎒 Pack the Backpack — Rounds 4 to 6", "GAME", "8–17 min",
                         "Camp Entrance", FOREST)
    one_task(slide, "Three more things and the backpack is full.", FOREST)
    sound_rows(slide, BACKPACK_B, 4, FOREST, "PACK SOMETHING WITH")
    hint(slide, "Ask him to say the word slowly: what sound do you hear first?",
         6.45)


def s13_entrance_done():
    badge_slide("🏕️ Camp Entrance Complete", "8–17 min", "Camp Entrance", "🏕️",
                "BADGE 1 EARNED",
                "You can hear the first sound in a word.",
                [("👂", "C, T, M, B and F sounds."),
                 ("🎒", "You packed the backpack six times."),
                 ("🗣️", "You said every camp word out loud."),
                 ("⭐", "You did it without reading a single sentence.")],
                "Tent Station — the sound hiding in the MIDDLE of a word.")


def vowel_slide(index):
    letter, sound, stretch, words, color, light = VOWELS[index]
    slide, n = new_slide(f"🔦 Flashlight Vowel — {letter} says {sound}", "LEARN",
                         "17–26 min", "Tent Station", color)
    one_task(slide, "Shine the flashlight on the middle sound.", color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(4.3), Inches(4.4), light)
    tb(slide, Inches(0.5), Inches(2.0), Inches(4.3), Inches(0.9), "🔦", size=40,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.45), Inches(2.95), Inches(2.4), Inches(1.75), color)
    tb(slide, Inches(1.45), Inches(3.2), Inches(2.4), Inches(1.3), letter, size=74,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.5), Inches(4.9), Inches(4.3), Inches(0.7), f"says {sound}",
       size=26, bold=True, color=color, align=PP_ALIGN.CENTER, font="Georgia")
    add_round(slide, Inches(1.05), Inches(5.58), Inches(3.2), Inches(0.6), WHITE)
    tb(slide, Inches(1.05), Inches(5.7), Inches(3.2), Inches(0.42),
       f"say it like this:  {stretch}", size=14, bold=True, color=color,
       align=PP_ALIGN.CENTER)
    for i, (word, emoji) in enumerate(words):
        left = Inches(5.15 + i * 2.6)
        add_round(slide, left, Inches(1.85), Inches(2.4), Inches(4.4), WHITE)
        tb(slide, left, Inches(2.25), Inches(2.4), Inches(1.0), emoji, size=44,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(3.5), Inches(2.0),
                  Inches(0.95), light)
        tb(slide, left + Inches(0.2), Inches(3.65), Inches(2.0), Inches(0.66),
           word, size=30, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left, Inches(4.7), Inches(2.4), Inches(0.5),
           f"the {letter} is in the middle", size=11, color=SOFT,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.45), Inches(5.3), Inches(1.5),
                  Inches(0.55), color)
        tb(slide, left + Inches(0.45), Inches(5.41), Inches(1.5), Inches(0.4),
           sound, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    hint(slide, f"Hold the middle sound out loud: {stretch}. That is the sound "
                f"he is hunting for.", 6.45)


def s14_vowel_a():
    vowel_slide(0)


def s15_vowel_e():
    vowel_slide(1)


def s16_vowel_i():
    vowel_slide(2)


def s17_vowel_o():
    vowel_slide(3)


def s18_flashlight_sort():
    slide, n = new_slide("🔦 Game: Flashlight Sort", "GAME", "17–26 min",
                         "Tent Station", BEAM)
    one_task(slide, "Which beam does this word belong under?", BEAM)
    vowel_rows(slide, SORT, 1)
    hint(slide, "Beams are A, E, I and O. Say all four before he chooses.", 6.42)


def s19_quick_vowel():
    slide, n = new_slide("⚡ Quick Challenge — Missing Vowel", "GAME", "17–26 min",
                         "Tent Station", NIGHT)
    one_task(slide, "One letter fell out of the word. Which one goes back?", NIGHT)
    letter_rows(slide, MISSING_VOWEL, 1, NIGHT)
    hint(slide, "Try the word out loud with each letter. Only one sounds right.",
         6.42)


def s20_tent_done():
    badge_slide("⛺ Tent Station Complete", "17–26 min", "Tent Station", "⛺",
                "BADGE 2 EARNED", "You found the middle sound in every word.",
                [("🔦", "Beams A, E, I and O."),
                 ("🗺️", "map · red · sit · hot."),
                 ("⚡", "You put four missing vowels back."),
                 ("👂", "You heard the difference between them.")],
                "Flashlight Word Cave — where we build whole words.",
                NIGHT, L_NIGHT)


def s21_camp_words():
    slide, n = new_slide("🎒 Camp Words We Will Read", "VOCABULARY", "26–40 min",
                         "Word Cave", PINE)
    one_task(slide, "Every word today comes from this list. Nothing new later.",
             PINE)
    for i, (word, emoji) in enumerate(CAMP_WORDS):
        col, row = i % 7, i // 7
        left = Inches(0.5 + col * 1.78)
        top = Inches(2.1 + row * 1.7)
        add_round(slide, left, top, Inches(1.62), Inches(1.52), L_PINE)
        tb(slide, left, top + Inches(0.15), Inches(1.62), Inches(0.62), emoji,
           size=24, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.1), top + Inches(0.85), Inches(1.42),
                  Inches(0.52), WHITE)
        tb(slide, left + Inches(0.1), top + Inches(0.93), Inches(1.42),
           Inches(0.38), word, size=13, bold=True, color=PINE,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.6), Inches(12.35), Inches(0.62), L_BEAM)
    tb(slide, Inches(0.8), Inches(5.72), Inches(11.7), Inches(0.42),
       "For each one:  picture  →  say it  →  what it means  →  read it  →  "
       "use it in a sentence.", size=14, bold=True, color=BEAM,
       align=PP_ALIGN.CENTER)
    hint(slide, "Read across in a rhythm. He joins in wherever he can.", 6.45)


def s22_six_steps():
    slide, n = new_slide("🧭 Our Six Steps for Any Word", "METHOD", "26–40 min",
                         "Word Cave", BEAM)
    one_task(slide, "When a word looks hard, we always do these six steps.", BEAM)
    for i, (num, label, detail, example, color, light) in enumerate(SIX_STEPS):
        left = Inches(0.5 + i * 2.07)
        add_round(slide, left, Inches(1.95), Inches(1.9), Inches(3.7), light)
        add_oval(slide, left + Inches(0.65), Inches(2.18), Inches(0.6),
                 Inches(0.6), color)
        tb(slide, left + Inches(0.65), Inches(2.28), Inches(0.6), Inches(0.42),
           num, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.05), Inches(2.95), Inches(1.8), Inches(0.55),
           label, size=15, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left + Inches(0.12), Inches(3.58), Inches(1.65), Inches(0.9),
           detail, size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(4.55), Inches(1.6),
                  Inches(0.78), WHITE)
        tb(slide, left + Inches(0.15), Inches(4.72), Inches(1.6), Inches(0.5),
           example, size=13, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.85), Inches(12.35), Inches(0.5),
              L_FOREST)
    tb(slide, Inches(0.8), Inches(5.93), Inches(11.7), Inches(0.36),
       "Never skip step 6. Matching the picture is what makes the word mean "
       "something.", size=13, bold=True, color=FOREST)


def s23_build_a():
    slide, n = new_slide("🔦 Build the Camp Word — CAMP and TENT", "GAME",
                         "26–40 min", "Word Cave", BEAM)
    one_task(slide, "Say each sound. Then say the whole word fast.", BEAM)
    build_rows(slide, BUILD_A)
    hint(slide, "Four letters is a lot. Cover the last two and blend the first "
                "two.", 6.42)


def s24_build_b():
    slide, n = new_slide("🔦 Build the Camp Word — MAP and BAG", "GAME",
                         "26–40 min", "Word Cave", BEAM)
    one_task(slide, "Three sounds each. Slide your finger as you blend.", BEAM)
    build_rows(slide, BUILD_B)
    hint(slide, "Slide your finger under the letters, then sweep it fast.", 6.42)


def s25_build_c():
    slide, n = new_slide("🔦 Build the Camp Word — LOG and FOX", "GAME",
                         "26–40 min", "Word Cave", BEAM)
    one_task(slide, "Two more words and the cave is lit up.", BEAM)
    build_rows(slide, BUILD_C)
    hint(slide, "Both words have O in the middle. Stretch it: l-ooo-g.", 6.42)


def s26_missing_letter():
    slide, n = new_slide("🔤 Game: Missing Letter", "GAME", "26–40 min",
                         "Word Cave", PLUM)
    one_task(slide, "A letter dropped off the camp sign. Put it back.", PLUM)
    letter_rows(slide, MISSING_LETTER, 1, PLUM)
    hint(slide, "Say the word out loud with each letter. Only one sounds right.",
         6.42)


def s27_scramble():
    slide, n = new_slide("🧩 Camp Word Builder — Mixed Letters", "GAME",
                         "26–40 min", "Word Cave", BEAM)
    one_task(slide, "The letters got mixed up. Put them back in order.", BEAM)
    scramble_rows(slide, SCRAMBLE, 1)
    hint(slide, "Find the vowel first. It almost always sits in the middle.",
         6.42)


def s28_no_cards():
    slide, n = new_slide("📖 No Letter Cards — Just Read It", "PRACTICE",
                         "26–40 min", "Word Cave", FOREST)
    one_task(slide, "This time there are no letter cards. You can do it!", FOREST)
    for i, (word, emoji) in enumerate(NO_CARDS):
        col, row = i % 3, i // 3
        left = Inches(0.6 + col * 4.1)
        top = Inches(1.95 + row * 2.25)
        add_round(slide, left, top, Inches(3.85), Inches(2.05), L_FOREST)
        tb(slide, left, top + Inches(0.18), Inches(3.85), Inches(0.72), emoji,
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.7), top + Inches(1.0), Inches(2.45),
                  Inches(0.82), WHITE)
        tb(slide, left + Inches(0.7), top + Inches(1.12), Inches(2.45),
           Inches(0.6), word, size=32, bold=True, color=FOREST,
           align=PP_ALIGN.CENTER, font="Arial Black")
    hint(slide, "If he stalls, put your finger under the first letter. Say "
                "nothing.", 6.45)


def s29_cave_done():
    badge_slide("🔦 Word Cave Complete", "26–40 min", "Word Cave", "🔦",
                "BADGE 3 EARNED", "You built six whole words from single letters.",
                [("🧩", "camp · tent · map · bag · log · fox."),
                 ("🔤", "You put four missing letters back."),
                 ("🎒", "You unscrambled four mixed-up words."),
                 ("📖", "And you read six words with no letter cards.")],
                "Animal Track Trail — words that rhyme travel together.",
                BEAM, L_BEAM)


def s30_families():
    slide, n = new_slide("🐾 Animal Track Words", "LEARN", "40–52 min",
                         "Animal Trail", TRAIL)
    one_task(slide, "Same ending sound = same track family.", TRAIL)
    for i, (family, words, color, light) in enumerate(FAMILIES):
        left = Inches(0.6 + i * 3.12)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(3.9), light)
        add_round(slide, left + Inches(0.5), Inches(2.2), Inches(1.9),
                  Inches(0.8), color)
        tb(slide, left + Inches(0.5), Inches(2.33), Inches(1.9), Inches(0.58),
           family, size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, word in enumerate(words):
            top = Inches(3.25 + j * 0.72)
            add_round(slide, left + Inches(0.45), top, Inches(2.0), Inches(0.6),
                      WHITE)
            tb(slide, left + Inches(0.45), top + Inches(0.08), Inches(2.0),
               Inches(0.44), word, size=20, bold=True, color=color,
               align=PP_ALIGN.CENTER)
    hint(slide, "Read down each column. Only the first sound changes.", 6.42)


def s31_track_match():
    slide, n = new_slide("🐾 Game: Track Match", "GAME", "40–52 min",
                         "Animal Trail", TRAIL)
    one_task(slide, "Read the word, then find the picture it belongs to.", TRAIL)
    match_rows(slide, TRACK_MATCH, 1)
    hint(slide, "Read the word first, every time. Guessing from pictures is easy "
                "— reading is the job.", 6.6)


def s32_family_a():
    slide, n = new_slide("🐾 Who Belongs to the Family?", "GAME", "40–52 min",
                         "Animal Trail", TRAIL)
    one_task(slide, "Three of these four belong. Which one does not?", TRAIL)
    family_rows(slide, FAMILY_A, 1)
    hint(slide, "Cover the first letter. Now only the ending is showing.", 6.42)


def s33_family_b():
    slide, n = new_slide("🐾 Who Belongs? — Rounds 3 and 4", "GAME", "40–52 min",
                         "Animal Trail", TRAIL)
    one_task(slide, "Two more families. Read every word out loud first.", TRAIL)
    family_rows(slide, FAMILY_B, 3)
    hint(slide, "Say the family ending out loud before each round: -ox, -un.",
         6.42)


def s34_new_track():
    slide, n = new_slide("🐾 Make a New Track — Rhyming", "GAME", "40–52 min",
                         "Animal Trail", TRAIL)
    one_task(slide, "Which word rhymes with mine?", TRAIL)
    for i, (word, emoji, options, color, light) in enumerate(NEW_TRACK):
        top = Inches(1.9 + i * 1.32)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.18), light)
        tb(slide, Inches(0.9), top + Inches(0.26), Inches(0.9), Inches(0.68),
           emoji, size=26, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.0), top + Inches(0.28), Inches(2.0),
                  Inches(0.64), color)
        tb(slide, Inches(2.0), top + Inches(0.37), Inches(2.0), Inches(0.48),
           word, size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, Inches(4.2), top + Inches(0.35), Inches(0.7), Inches(0.48), "→",
           size=20, bold=True, color=color, align=PP_ALIGN.CENTER)
        for j, opt in enumerate(options):
            left = Inches(5.1 + j * 2.5)
            add_round(slide, left, top + Inches(0.28), Inches(2.3), Inches(0.64),
                      WHITE)
            tb(slide, left, top + Inches(0.37), Inches(2.3), Inches(0.48), opt,
               size=22, bold=True, color=INK, align=PP_ALIGN.CENTER,
               font="Arial Black")
    add_round(slide, Inches(0.5), Inches(5.86), Inches(12.35), Inches(0.44),
              L_PLUM)
    tb(slide, Inches(0.8), Inches(5.93), Inches(11.7), Inches(0.34),
       "⭐ TRAIL CHALLENGE — Can you think of one more word that rhymes with "
       "CAT?  Any real word counts.", size=13, bold=True, color=PLUM)
    hint(slide, "Rhyme means the ending sounds the same. Say both words back to "
                "back.", 6.45)


def s35_break_intro():
    slide, n = new_slide("🧠 Brain Break — Campfire Freeze", "BREAK", "40–52 min",
                         "Animal Trail", PINE)
    one_task(slide, "Stay by your chair. Act out whatever I call.", PINE)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4), L_PINE)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.4), Inches(1.8), "🔥", size=90,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.4), Inches(0.7), "5 MINUTES",
       size=30, bold=True, color=PINE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.8), Inches(0.6),
       "then straight back to the words", size=14, color=SOFT,
       align=PP_ALIGN.CENTER)
    rules = [("🪑", "Stay beside your chair the whole time."),
             ("🧊", "\"Freeze!\" means stop completely."),
             ("👉", "Some calls are reading calls too."),
             ("🔁", "The last two rounds, you call and I act.")]
    for i, (icon, line) in enumerate(rules):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.55), top + Inches(0.2), Inches(0.56),
                 Inches(0.56), WHITE)
        tb(slide, Inches(6.55), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.35), top + Inches(0.24), Inches(5.2), Inches(0.5),
           line, size=15, bold=True, color=INK)
    hint(slide, "Letting him call the commands is the part that resets his focus.",
         6.45)


def s36_freeze():
    slide, n = new_slide("🧊 Campfire Freeze — Six Calls", "BREAK", "40–52 min",
                         "Animal Trail", PINE)
    one_task(slide, "Listen for the call. Then act it out.", PINE)
    for i, (icon, call, action) in enumerate(FREEZE):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_PINE)
        tb(slide, left, top + Inches(0.18), Inches(3.9), Inches(0.7), icon,
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.92), Inches(3.6),
           Inches(0.5), call, size=19, bold=True, color=PINE,
           align=PP_ALIGN.CENTER, font="Georgia")
        tb(slide, left + Inches(0.3), top + Inches(1.46), Inches(3.3), Inches(0.5),
           action, size=12, color=DARK, align=PP_ALIGN.CENTER)
    hint(slide, "Mix the order so he has to listen, not predict.", 6.45)


def s37_camper_says():
    slide, n = new_slide("🗣️ Camper Says — With Reading", "BREAK", "40–52 min",
                         "Animal Trail", TRAIL)
    one_task(slide, "Only do it if I say \"Camper says\" first!", TRAIL)
    for i, (icon, line, kind) in enumerate(CAMPER_SAYS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 1.5)
        color = EMBER if kind == "read" else PINE
        light = L_EMBER if kind == "read" else L_PINE
        add_round(slide, left, top, Inches(5.95), Inches(1.3), light)
        add_oval(slide, left + Inches(0.3), top + Inches(0.36), Inches(0.58),
                 Inches(0.58), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(0.58),
           Inches(0.44), icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.3), Inches(4.5),
           Inches(0.52), line, size=17, bold=True, color=INK)
        tb(slide, left + Inches(1.1), top + Inches(0.83), Inches(4.5),
           Inches(0.36), "reading call" if kind == "read" else "movement call",
           size=11, color=color)
    hint(slide, "Four of these six are reading calls. That keeps the break on "
                "task.", 6.45)


def s38_calm():
    slide, n = new_slide("🌲 Calm Down — Ready for the Campfire", "BREAK",
                         "40–52 min", "Animal Trail", PINE)
    one_task(slide, "Four quiet things, then we read again.", PINE)
    for i, (icon, label, detail) in enumerate(CALM):
        left = Inches(0.6 + i * 3.13)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(3.2), L_PINE)
        tb(slide, left, Inches(2.2), Inches(2.9), Inches(0.95), icon, size=40,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.35), Inches(2.7), Inches(0.6),
           label, size=19, bold=True, color=PINE, align=PP_ALIGN.CENTER,
           font="Georgia")
        add_round(slide, left + Inches(0.35), Inches(4.1), Inches(2.2),
                  Inches(0.6), WHITE)
        tb(slide, left + Inches(0.35), Inches(4.22), Inches(2.2), Inches(0.42),
           detail, size=12, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.4), Inches(12.35), Inches(0.85),
              L_FOREST)
    tb(slide, Inches(0.8), Inches(5.6), Inches(11.7), Inches(0.5),
       "\"You are ready for the next mission!\"", size=24, bold=True,
       color=FOREST, align=PP_ALIGN.CENTER, font="Georgia")
    hint(slide, "Say the line out loud together. It marks the end of the break.",
         6.45)


def sight_slide(words, title, sub, hint_text):
    slide, n = new_slide(title, "LEARN", "52–70 min", "Campfire", EMBER)
    one_task(slide, sub, EMBER)
    for i, (word, phrase, emoji) in enumerate(words):
        left = Inches(0.7 + i * 4.1)
        add_round(slide, left, Inches(1.95), Inches(3.8), Inches(3.9), L_EMBER)
        add_round(slide, left + Inches(0.55), Inches(2.25), Inches(2.7),
                  Inches(1.2), EMBER)
        tb(slide, left + Inches(0.55), Inches(2.47), Inches(2.7), Inches(0.8),
           word, size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left, Inches(3.65), Inches(3.8), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.6), Inches(3.2),
                  Inches(0.8), WHITE)
        tb(slide, left + Inches(0.3), Inches(4.78), Inches(3.2), Inches(0.5),
           phrase, size=19, bold=True, color=EMBER, align=PP_ALIGN.CENTER)
    hint(slide, hint_text, 6.42)


def s39_sight_1():
    sight_slide(SIGHT_1, "🔥 Campfire Sight Words — Group 1",
                "These three words we do NOT sound out. We just know them.",
                "Say the word, he repeats, then he finds it in the phrase below.")


def s40_sight_2():
    sight_slide(SIGHT_2, "🔥 Campfire Sight Words — Group 2",
                "Three more. Only three at a time — never a long list.",
                "Go back to Group 1 for ten seconds before you start Group 2.")


def s41_light_word():
    slide, n = new_slide("🔦 Game: Light the Word", "GAME", "52–70 min",
                         "Campfire", EMBER)
    one_task(slide, "Shine the flashlight on the word I say.", EMBER)
    find_rows(slide, LIGHT_WORD, 1)
    hint(slide, "Read all three out loud first. Listening beats looking.", 6.42)


def s42_sentence_a():
    slide, n = new_slide("📕 Build a Camp Sentence", "READ", "52–70 min",
                         "Campfire", EMBER)
    one_task(slide, "A whole sentence now. I read it first.", EMBER)
    sentence_rows(slide, SENT_A)
    hint(slide, "Point at each word as you read so his eyes track left to right.",
         6.42)


def s43_sentence_b():
    slide, n = new_slide("📕 Build a Camp Sentence — Two More", "READ",
                         "52–70 min", "Campfire", EMBER)
    one_task(slide, "Two more. Read them like you are talking.", EMBER)
    sentence_rows(slide, SENT_B)
    hint(slide, "A stuck word? Give only the first sound and wait five seconds.",
         6.42)


def s44_phrases():
    slide, n = new_slide("🔗 Two and Three Words Together", "READ", "52–70 min",
                         "Campfire", EMBER)
    one_task(slide, "Not a whole sentence yet. Just a few words joined up.",
             EMBER)
    phrase_grid(slide, PHRASES)
    hint(slide, "Sweep your finger under the whole phrase so it sounds like "
                "talking.", 6.6)


def s45_read_steps():
    slide, n = new_slide("📖 How We Read Every Sentence", "METHOD", "52–70 min",
                         "Campfire", BEAM)
    one_task(slide, "Three steps, every single time. You are never first.", BEAM)
    for i, (step, label, detail, color) in enumerate(READ_STEPS):
        left = Inches(0.55 + i * 4.15)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(3.95), L_GREY)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(0.72), color)
        tb(slide, left, Inches(2.08), Inches(3.9), Inches(0.5), step, size=17,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.95), Inches(3.9), Inches(0.8), label, size=34,
           bold=True, color=color, align=PP_ALIGN.CENTER, font="Georgia")
        tb(slide, left + Inches(0.35), Inches(3.95), Inches(3.2), Inches(0.9),
           detail, size=14, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.6), Inches(4.95), Inches(2.7),
                  Inches(0.62), WHITE)
        tb(slide, left + Inches(0.6), Inches(5.06), Inches(2.7), Inches(0.44),
           ["🧑‍🏫  me", "🧑‍🏫🙋  both", "🙋  you"][i], size=15, bold=True,
           color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Only move to step 3 when step 2 sounded smooth. No rush.", 6.45)


def s46_puzzle_a():
    slide, n = new_slide("🧩 Camp Sentence Puzzle", "GAME", "52–70 min",
                         "Campfire", BEAM)
    one_task(slide, "These word cards are mixed up. Put them in order.", BEAM)
    puzzle_rows(slide, PUZZLES_A, 1)
    hint(slide, "Find the capital letter. That word always goes first.", 6.42)


def s47_puzzle_b():
    slide, n = new_slide("🧩 Camp Sentence Puzzle — Rounds 3 and 4", "GAME",
                         "52–70 min", "Campfire", BEAM)
    one_task(slide, "Two more. Read it out loud after you fix the order.", BEAM)
    puzzle_rows(slide, PUZZLES_B, 3)
    hint(slide, "Read his version back exactly as written. Silly is a clue.",
         6.42)


def s48_picture_match():
    slide, n = new_slide("🖼️ Game: Picture Match", "GAME", "52–70 min",
                         "Campfire", FOREST)
    one_task(slide, "Read the sentence, then find the picture it describes.",
             FOREST)
    tops = numbered_rows(slide, len(PICTURE_MATCH), 1, FOREST, top_start=1.9,
                         gap=1.5)
    for (sentence, options), top in zip(PICTURE_MATCH, tops):
        add_round(slide, Inches(1.35), top + Inches(0.34), Inches(3.2),
                  Inches(0.66), L_FOREST)
        tb(slide, Inches(1.45), top + Inches(0.43), Inches(3.0), Inches(0.5),
           sentence, size=18, bold=True, color=INK, align=PP_ALIGN.CENTER)
        for j, (emoji, _name) in enumerate(options):
            left = Inches(4.85 + j * 2.6)
            add_round(slide, left, top + Inches(0.16), Inches(2.4), Inches(1.02),
                      WHITE)
            tb(slide, left, top + Inches(0.26), Inches(2.4), Inches(0.8), emoji,
               size=36, align=PP_ALIGN.CENTER)
    hint(slide, "Read the sentence twice before he points. The second read is "
                "where it clicks.", 6.45)


def s49_silly():
    slide, n = new_slide("🎯 Game: Sentence or Silly?", "GAME", "52–70 min",
                         "Campfire", PLUM)
    one_task(slide, "Look at the picture. Which sentence makes sense?", PLUM)
    for i, (emoji, options) in enumerate(SILLY):
        top = Inches(1.95 + i * 1.52)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), L_PLUM)
        tb(slide, Inches(0.85), top + Inches(0.3), Inches(1.2), Inches(0.78),
           emoji, size=32, align=PP_ALIGN.CENTER)
        for j, opt in enumerate(options):
            left = Inches(2.4 + j * 5.2)
            add_round(slide, left, top + Inches(0.3), Inches(4.9), Inches(0.74),
                      WHITE)
            tb(slide, left + Inches(0.25), top + Inches(0.42), Inches(3.4),
               Inches(0.52), opt, size=19, bold=True, color=INK)
            tb(slide, left + Inches(3.8), top + Inches(0.42), Inches(0.9),
               Inches(0.52), "✅ / ❌", size=14, color=SOFT,
               align=PP_ALIGN.CENTER)
    hint(slide, "Let him laugh at the silly one, then ask why it cannot be true.",
         6.45)


def s50_own_sentence():
    slide, n = new_slide("🗣️ Now You Make the Sentence", "SPEAKING", "52–70 min",
                         "Campfire", PINE)
    one_task(slide, "Pick a picture. Say a whole sentence about it.", PINE)
    for i, (emoji, word) in enumerate(OWN_SENTENCE):
        left = Inches(0.6 + i * 3.15)
        add_round(slide, left, Inches(1.95), Inches(2.95), Inches(2.6), L_PINE)
        tb(slide, left, Inches(2.2), Inches(2.95), Inches(0.95), emoji, size=40,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.45), Inches(3.4), Inches(2.05),
                  Inches(0.68), WHITE)
        tb(slide, left + Inches(0.45), Inches(3.52), Inches(2.05), Inches(0.48),
           word, size=22, bold=True, color=PINE, align=PP_ALIGN.CENTER)
    for i, starter in enumerate(["I see a ______.", "I can see the ______."]):
        left = Inches(0.6 + i * 6.35)
        add_round(slide, left, Inches(4.85), Inches(6.05), Inches(1.15), L_BEAM)
        tb(slide, left + Inches(0.4), Inches(5.12), Inches(5.3), Inches(0.62),
           starter, size=26, bold=True, color=INK)
    hint(slide, "Two sentences is plenty. Speaking warms him up for the story.",
         6.45)


def s51_campfire_done():
    badge_slide("🔥 Campfire Complete", "52–70 min", "Campfire", "🔥",
                "BADGE 4 EARNED", "You read real sentences out loud.",
                [("🔥", "Sight words: I · see · a · the · my · can."),
                 ("🔗", "Six camp phrases."),
                 ("🧩", "Four sentence puzzles put back in order."),
                 ("🗣️", "And you made up sentences of your own.")],
                "Morning Picnic — a whole story, start to finish.",
                EMBER, L_EMBER)


def s52_story_intro():
    slide, n = new_slide("📖 Story Time — The Lost Flashlight", "STORY",
                         "70–82 min", "Morning Picnic", PINE)
    one_task(slide, "Six words to watch for. You already know all six.", PINE)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4), L_PINE)
    tb(slide, Inches(0.5), Inches(2.25), Inches(5.4), Inches(1.5), "🔦", size=76,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(3.95), Inches(5.4), Inches(0.7),
       "THE LOST FLASHLIGHT", size=25, bold=True, color=PINE,
       align=PP_ALIGN.CENTER, font="Georgia")
    add_round(slide, Inches(0.9), Inches(4.8), Inches(4.6), Inches(1.15), WHITE)
    tb(slide, Inches(1.1), Inches(4.92), Inches(4.2), Inches(0.36),
       "🤔 BEFORE WE READ", size=11, bold=True, color=PINE)
    tb(slide, Inches(1.1), Inches(5.28), Inches(4.2), Inches(0.58),
       "What do you think Sam will lose?", size=15, bold=True, color=INK)
    tb(slide, Inches(6.25), Inches(1.9), Inches(6.6), Inches(0.5),
       "WATCH FOR THESE WORDS", size=15, bold=True, color=PINE)
    for i, (word, emoji) in enumerate(STORY_WATCH):
        col, row = i % 3, i // 3
        left = Inches(6.25 + col * 2.25)
        top = Inches(2.55 + row * 1.85)
        add_round(slide, left, top, Inches(2.05), Inches(1.65), L_GREY)
        tb(slide, left, top + Inches(0.18), Inches(2.05), Inches(0.62), emoji,
           size=24, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), top + Inches(0.88), Inches(1.65),
                  Inches(0.55), WHITE)
        tb(slide, left + Inches(0.2), top + Inches(0.98), Inches(1.65),
           Inches(0.4), word, size=17, bold=True, color=PINE,
           align=PP_ALIGN.CENTER)
    hint(slide, "Read the six words together now. Meeting them early removes the "
                "fear.", 6.45)


def s53_story_1():
    story_slide(STORY[0], "70–82 min")


def s54_story_2():
    story_slide(STORY[1], "70–82 min")


def s55_story_3():
    story_slide(STORY[2], "70–82 min")


def s56_story_4():
    story_slide(STORY[3], "70–82 min")


def s57_story_5():
    story_slide(STORY[4], "70–82 min")


def s58_read_again():
    slide, n = new_slide("🔁 Read It Again — Your Turn", "READ", "70–82 min",
                         "Morning Picnic", FOREST)
    one_task(slide, "Six easy lines from the story. Read them on your own.",
             FOREST)
    for i, line in enumerate(EASY_LINES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.5)
        add_round(slide, left, top, Inches(5.95), Inches(1.3), L_FOREST)
        add_oval(slide, left + Inches(0.3), top + Inches(0.38), Inches(0.55),
                 Inches(0.55), FOREST)
        tb(slide, left + Inches(0.3), top + Inches(0.44), Inches(0.55),
           Inches(0.42), str(i + 1), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.05), top + Inches(0.38), Inches(4.6),
           Inches(0.62), line, size=21, bold=True, color=INK)
    hint(slide, "Second reading is always smoother. Say so out loud when it is.",
         6.55)


def s59_q_a():
    slide, n = new_slide("🔎 Story Detective — Q1 and Q2", "QUESTIONS",
                         "82–88 min", "Morning Picnic", PINE)
    one_task(slide, "Answer from the story, not from your memory.", PINE)
    question_rows(slide, QUESTIONS_A, 1)
    hint(slide, "Turn back to the part named under the question. Let him find "
                "it.", 6.42)


def s60_q_b():
    slide, n = new_slide("🔎 Story Detective — Q3 and Q4", "QUESTIONS",
                         "82–88 min", "Morning Picnic", PINE)
    one_task(slide, "Two more. Point to the line that proves it.", PINE)
    question_rows(slide, QUESTIONS_B, 3)
    hint(slide, "Cover one wrong answer. Two choices is still real thinking.",
         6.42)


def s61_evidence():
    slide, n = new_slide("🔦 Find the Evidence", "QUESTIONS", "82–88 min",
                         "Morning Picnic", BEAM)
    one_task(slide, "Show me the sentence that tells us the answer.", BEAM)
    for i, (question, _sentence, part, color, light) in enumerate(EVIDENCE):
        top = Inches(1.9 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), light)
        add_oval(slide, Inches(0.8), top + Inches(0.42), Inches(0.52),
                 Inches(0.52), color)
        tb(slide, Inches(0.8), top + Inches(0.48), Inches(0.52), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.22), Inches(4.3), Inches(0.5),
           question, size=17, bold=True, color=INK)
        add_round(slide, Inches(1.55), top + Inches(0.76), Inches(2.0),
                  Inches(0.44), WHITE)
        tb(slide, Inches(1.55), top + Inches(0.82), Inches(2.0), Inches(0.34),
           f"look in {part}", size=11, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        add_round(slide, Inches(6.1), top + Inches(0.3), Inches(6.45),
                  Inches(0.74), WHITE)
        tb(slide, Inches(6.35), top + Inches(0.4), Inches(6.0), Inches(0.36),
           "🔎  Turn back to the story and point to the line.", size=13,
           bold=True, color=color)
        tb(slide, Inches(6.35), top + Inches(0.72), Inches(6.0), Inches(0.3),
           "Then read it out loud.  ☐ found it", size=11, color=SOFT)
    add_round(slide, Inches(0.5), Inches(6.46), Inches(12.35), Inches(0.44),
              L_FOREST)
    tb(slide, Inches(0.8), Inches(6.53), Inches(11.7), Inches(0.34),
       "Pointing to the proof matters more than getting the answer fast. This "
       "is how real reading works.", size=12, bold=True, color=FOREST)


def s62_final_words():
    slide, n = new_slide("🏆 Final Camp Challenge — Words", "CHALLENGE",
                         "88–90 min", "Morning Picnic", BEAM)
    one_task(slide, "Six words. You have read every one today.", BEAM)
    for i, (word, emoji, color, light) in enumerate(FINAL_WORDS):
        col, row = i % 3, i // 3
        left = Inches(0.6 + col * 4.1)
        top = Inches(1.95 + row * 2.25)
        add_round(slide, left, top, Inches(3.85), Inches(2.05), light)
        tb(slide, left, top + Inches(0.18), Inches(3.85), Inches(0.72), emoji,
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.7), top + Inches(1.0), Inches(2.45),
                  Inches(0.82), WHITE)
        tb(slide, left + Inches(0.7), top + Inches(1.12), Inches(2.45),
           Inches(0.6), word, size=32, bold=True, color=color,
           align=PP_ALIGN.CENTER, font="Arial Black")
    hint(slide, "Hints are still allowed here. Finishing matters, not scoring.",
         6.45)


def s63_final_sentences():
    slide, n = new_slide("🏆 Final Camp Challenge — Sentences", "CHALLENGE",
                         "88–90 min", "Morning Picnic", BEAM)
    one_task(slide, "Three sentences. Then you are a champion.", BEAM)
    sentence_rows(slide, FINAL_SENTENCES, top_start=1.9, gap=1.5, height=1.34,
                  size=30)
    add_round(slide, Inches(0.5), Inches(6.0), Inches(12.35), Inches(0.5),
              L_FOREST)
    tb(slide, Inches(0.8), Inches(6.08), Inches(11.7), Inches(0.36),
       "✅  All three read? Ring the camp bell — the adventure is complete!",
       size=15, bold=True, color=FOREST, align=PP_ALIGN.CENTER)
    hint(slide, "Support is fine on this slide. Reading it at all is the win.",
         6.55)


def s64_champion():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), BEAM)
    for x, y in [(0.35, 3.95), (12.3, 3.95), (0.6, 5.45), (12.1, 5.45)]:
        tb(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.7), "⭐", size=24,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(0.9), Inches(12), Inches(1.1), "🏆", size=58,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(2.0), Inches(12), Inches(0.9),
       "CAMP READING CHAMPION!", size=38, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.95), Inches(12), Inches(0.5),
       "Junior Camp Explorer badge — earned at all six stops.", size=18,
       color=RGBColor(0xEF, 0xD6, 0x95), align=PP_ALIGN.CENTER, italic=True)
    for i, (icon, line) in enumerate(CHAMPION_LINES):
        left = Inches(1.35 + i * 2.15)
        add_round(slide, left, Inches(3.65), Inches(1.95), Inches(1.5),
                  RGBColor(0x21, 0x2F, 0x3A))
        tb(slide, left, Inches(3.82), Inches(1.95), Inches(0.6), icon, size=20,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.12), Inches(4.42), Inches(1.7), Inches(0.62),
           line, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.4), Inches(5.35), Inches(6.5), Inches(1.1), BEAM)
    tb(slide, Inches(3.4), Inches(5.6), Inches(6.5), Inches(0.62),
       "\"I AM A READER!\"", size=30, bold=True, color=INK,
       align=PP_ALIGN.CENTER, font="Georgia")
    footer(slide, n, "88–90 min", "Morning Picnic")
    fade(slide)


def s65_can_read():
    slide, n = new_slide("✅ Today I Can...", "CHECKLIST", "88–90 min",
                         "Morning Picnic", FOREST)
    one_task(slide, "Tick every one you did today. Read them with me.", FOREST)
    for i, (icon, skill, example) in enumerate(CAN_READ):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 0.95)
        add_round(slide, left, top, Inches(5.95), Inches(0.84), L_FOREST)
        tb(slide, left + Inches(0.25), top + Inches(0.18), Inches(0.5),
           Inches(0.5), icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.9), top + Inches(0.16), Inches(2.5),
           Inches(0.5), skill, size=15, bold=True, color=INK)
        tb(slide, left + Inches(3.5), top + Inches(0.2), Inches(1.9),
           Inches(0.44), example, size=12, color=SOFT)
        tb(slide, left + Inches(5.4), top + Inches(0.16), Inches(0.4),
           Inches(0.5), "☐", size=17, color=FOREST)
    hint(slide, "Read the list aloud together. Hearing the whole list is the "
                "point.", 6.6)


def optional_card(slide, left, top, num, name, prompt, cards, answer_note, color,
                  light, width=3.9, height=4.3):
    add_round(slide, left, top, Inches(width), Inches(height), light)
    add_round(slide, left, top, Inches(width), Inches(0.5), color)
    tb(slide, left, top + Inches(0.08), Inches(width), Inches(0.36),
       f"GAME {num} — {name}", size=12, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    tb(slide, left + Inches(0.25), top + Inches(0.65), Inches(width - 0.5),
       Inches(0.5), prompt, size=14, bold=True, color=INK,
       align=PP_ALIGN.CENTER)
    for j, card in enumerate(cards):
        add_round(slide, left + Inches(0.4), top + Inches(1.25 + j * 0.78),
                  Inches(width - 0.8), Inches(0.66), WHITE)
        tb(slide, left + Inches(0.4), top + Inches(1.37 + j * 0.78),
           Inches(width - 0.8), Inches(0.48), card, size=19, bold=True,
           color=color, align=PP_ALIGN.CENTER)
    tb(slide, left + Inches(0.25), top + Inches(height - 0.55),
       Inches(width - 0.5), Inches(0.42), answer_note, size=11, color=SOFT,
       align=PP_ALIGN.CENTER)


def s66_optional_a():
    slide, n = new_slide("🎲 Extra Games — If There Is Time", "OPTIONAL", "",
                         "Morning Picnic", PLUM)
    one_task(slide, "Five spare games. Use any of them, in any order.", PLUM)
    optional_card(slide, Inches(0.5), Inches(1.9), 1, "MYSTERY WORD",
                  "_ A P  —  which letter?", ["M", "T", "B"],
                  "any of the three makes a real word", EMBER, L_EMBER)
    optional_card(slide, Inches(4.72), Inches(1.9), 2, "WORD OR PICTURE?",
                  "Read FOX, then point.", ["🦊", "🐕", "⛺"],
                  "he must read the word before pointing", FOREST, L_FOREST)
    optional_card(slide, Inches(8.94), Inches(1.9), 3, "MISSING WORD",
                  "\"I see a ____.\"", ["tent", "run", "hot"],
                  "only one word fits the sentence", NIGHT, L_NIGHT)
    hint(slide, "These reuse today's words only. Nothing new appears.", 6.45)


def s67_optional_b():
    slide, n = new_slide("🎲 Extra Games — Two More", "OPTIONAL", "",
                         "Morning Picnic", PLUM)
    one_task(slide, "One rhyming game and one speaking game.", PLUM)
    optional_card(slide, Inches(0.5), Inches(1.9), 4, "RHYME CAMP",
                  "CAT rhymes with...", ["HAT", "DOG", "SUN"],
                  "say both words back to back", BEAM, L_BEAM)
    optional_card(slide, Inches(4.72), Inches(1.9), 5, "CAMPER SENTENCE",
                  "Look at the picture: ⛺", ["I see a ____.", "My ____ is big."],
                  "accept any sensible sentence", PINE, L_PINE)
    add_round(slide, Inches(8.94), Inches(1.9), Inches(3.9), Inches(4.3), L_GREY)
    tb(slide, Inches(9.2), Inches(2.1), Inches(3.4), Inches(0.4),
       "WHEN TO USE THESE", size=13, bold=True, color=PLUM)
    bullets(slide, Inches(9.2), Inches(2.6), Inches(3.4), Inches(3.4),
            ["He finishes an activity early.", "Focus is dropping mid-block.",
             "A skill needs one more pass.", "You have five minutes spare.",
             "He asks to play one more."], size=12, sp=10)
    hint(slide, "Stop while he still wants one more. That is the right moment.",
         6.45)


def s68_support():
    slide, n = new_slide("🧰 Support System — For the Teacher", "TEACHER", "",
                         "Morning Picnic", NIGHT)
    one_task(slide, "Levels, hint ladder and the words to say.", NIGHT)
    for i, (icon, name, who, action, color, light) in enumerate(SUPPORT_LEVELS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(2.25), light)
        tb(slide, left + Inches(0.25), Inches(2.02), Inches(0.5), Inches(0.5),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.9), Inches(2.02), Inches(2.9), Inches(0.5),
           name, size=15, bold=True, color=color)
        tb(slide, left + Inches(0.3), Inches(2.6), Inches(3.3), Inches(0.5), who,
           size=12, color=DARK)
        add_round(slide, left + Inches(0.3), Inches(3.15), Inches(3.3),
                  Inches(0.78), WHITE)
        tb(slide, left + Inches(0.45), Inches(3.26), Inches(3.0), Inches(0.58),
           action, size=11, color=INK)
    tb(slide, Inches(0.5), Inches(4.25), Inches(6.0), Inches(0.4),
       "HINT LADDER — never skip a step, never give the answer first", size=13,
       bold=True, color=NIGHT)
    for i, (label, text, color) in enumerate(HINT_LADDER):
        top = Inches(4.66 + i * 0.46)
        add_round(slide, Inches(0.5), top, Inches(6.0), Inches(0.42), L_GREY)
        add_round(slide, Inches(0.62), top + Inches(0.05), Inches(1.0),
                  Inches(0.32), color)
        tb(slide, Inches(0.62), top + Inches(0.07), Inches(1.0), Inches(0.28),
           label, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.8), top + Inches(0.05), Inches(4.5), Inches(0.32),
           text, size=12, bold=True, color=INK)
    tb(slide, Inches(6.85), Inches(4.25), Inches(3.0), Inches(0.4), "HINT BANK",
       size=13, bold=True, color=NIGHT)
    bullets(slide, Inches(6.85), Inches(4.72), Inches(3.0), Inches(2.1),
            HINT_BANK, size=11)
    tb(slide, Inches(10.0), Inches(4.25), Inches(2.85), Inches(0.4),
       "WORDS TO USE", size=13, bold=True, color=FOREST)
    bullets(slide, Inches(10.0), Inches(4.72), Inches(2.85), Inches(2.1), PRAISE,
            size=11, color=FOREST)


def s69_assessment():
    slide, n = new_slide("📋 End-of-Class Assessment", "TEACHER", "",
                         "Morning Picnic", TRAIL)
    one_task(slide, "Tick one box per skill right after the lesson.", TRAIL)
    heads = ["SKILL", "INDEPENDENT", "WITH SUPPORT", "NEEDS MORE PRACTICE"]
    widths = [5.15, 2.4, 2.4, 2.4]
    lefts = [0.5]
    for w in widths[:-1]:
        lefts.append(lefts[-1] + w)
    add_rect(slide, Inches(0.5), Inches(1.82), Inches(12.35), Inches(0.42), TRAIL)
    for left, w, head in zip(lefts, widths, heads):
        tb(slide, Inches(left), Inches(1.86), Inches(w), Inches(0.34), head,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = Inches(2.3 + i * 0.36)
        add_rect(slide, Inches(0.5), top, Inches(12.35), Inches(0.34),
                 WHITE if i % 2 == 0 else L_GREY)
        tb(slide, Inches(0.7), top + Inches(0.02), Inches(4.8), Inches(0.3),
           skill, size=11, bold=True, color=INK)
        for left, w in zip(lefts[1:], widths[1:]):
            tb(slide, Inches(left), top + Inches(0.01), Inches(w), Inches(0.3),
               "☐", size=13, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.02), Inches(12.35), Inches(0.9),
              L_TRAIL)
    tb(slide, Inches(0.8), Inches(6.1), Inches(4.0), Inches(0.34),
       "FOR NEXT CLASS", size=12, bold=True, color=TRAIL)
    for i, item in enumerate(NEXT_CLASS):
        tb(slide, Inches(0.8 + i * 4.1), Inches(6.48), Inches(3.9), Inches(0.34),
           f"{i + 1}. {item}:  ____________", size=11, color=INK)


def s70_answer_key():
    slide, n = new_slide("🔑 Answer Key — Teacher Only", "TEACHER", "",
                         "Morning Picnic", EMBER)
    one_task(slide, "Quick reference for every game in this deck.", EMBER)
    cols = [
        ("SOUNDS & VOWELS (8–26)",
         ["Starts with: tent · map · bag",
          "Backpack: bag · map · tent · camp · fire · log",
          "Flashlight sort: A · E · I · O",
          "Missing vowel: A  E  O  A"], FOREST),
        ("BUILDING WORDS (26–40)",
         ["Blend: camp · tent · map · bag · log · fox",
          "Missing letter: A  A  O  U",
          "Mixed letters: MAP · TENT · LOG · BAG",
          "No cards: all six read as written"], BEAM),
        ("TRAIL & CAMPFIRE (40–70)",
         ["Odd one out: DOG · FOX · BAG · HAT",
          "Rhymes: HAT · DOG · RUN",
          "Light the word: I · see · the · can",
          "Puzzles: I see a tent. / The tent is big. /",
          " I can run. / My bag is red."], TRAIL),
        ("STORY & QUESTIONS (70–90)",
         ["1 Sam  2 his flashlight  3 near a log  4 took it to the tent",
          "Picture match: fox running · big tent · a log",
          "Sentence or silly: is red · can run · is hot",
          "Evidence: \"Sam is camping with his dad.\" /",
          " \"The bag is empty.\" / \"The light is near the log.\""], PINE),
    ]
    for i, (head, lines, color) in enumerate(cols):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(5.95), Inches(2.1), L_GREY)
        add_round(slide, left, top, Inches(5.95), Inches(0.46), color)
        tb(slide, left + Inches(0.25), top + Inches(0.06), Inches(5.45),
           Inches(0.34), head, size=12, bold=True, color=WHITE)
        bullets(slide, left + Inches(0.25), top + Inches(0.6), Inches(5.45),
                Inches(1.4), lines, size=11)
    add_round(slide, Inches(0.5), Inches(6.5), Inches(12.35), Inches(0.42),
              L_FOREST)
    tb(slide, Inches(0.8), Inches(6.56), Inches(11.7), Inches(0.32),
       "Accept any sensible answer for the speaking and rhyming tasks — "
       "fluency matters more than exact wording.", size=11, bold=True,
       color=FOREST)


BUILDERS = [
    s01_title, s02_mission, s03_route, s04_warmup, s05_sound_c, s06_sound_t,
    s07_sound_m, s08_sound_b, s09_sound_f, s10_which_start, s11_backpack_a,
    s12_backpack_b, s13_entrance_done, s14_vowel_a, s15_vowel_e, s16_vowel_i,
    s17_vowel_o, s18_flashlight_sort, s19_quick_vowel, s20_tent_done,
    s21_camp_words, s22_six_steps, s23_build_a, s24_build_b, s25_build_c,
    s26_missing_letter, s27_scramble, s28_no_cards, s29_cave_done, s30_families,
    s31_track_match, s32_family_a, s33_family_b, s34_new_track, s35_break_intro,
    s36_freeze, s37_camper_says, s38_calm, s39_sight_1, s40_sight_2,
    s41_light_word, s42_sentence_a, s43_sentence_b, s44_phrases, s45_read_steps,
    s46_puzzle_a, s47_puzzle_b, s48_picture_match, s49_silly, s50_own_sentence,
    s51_campfire_done, s52_story_intro, s53_story_1, s54_story_2, s55_story_3,
    s56_story_4, s57_story_5, s58_read_again, s59_q_a, s60_q_b, s61_evidence,
    s62_final_words, s63_final_sentences, s64_champion, s65_can_read,
    s66_optional_a, s67_optional_b, s68_support, s69_assessment, s70_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL}, built {len(prs.slides)}"

OUT = "Grade2_Camping_Reading_Adventure_90min.pptx"
prs.save(OUT)

story_words = sum(len(line.split()) for part in STORY for line in part[4])
with_notes = sum(1 for s in prs.slides
                 if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip())
print(f"saved {OUT}")
print(f"slides: {len(prs.slides)}")
print(f"story words: {story_words}")
print(f"slides with notes: {with_notes}")
