"""Grade 1 reading lesson - 90 minutes, 75 slides, no speaker notes.

"The Weather Station Reading Adventure" - the child is a Junior Weather
Reporter moving through the Sunny Station, the Rainbow Station, the Cloud
Station, the Rain Station, the Wind Station and the Weather Desk. Reading is
built one step at a time: hear, say, sound, blend, read, understand.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Weather Hint strip, and slides 71-75 hold the optional game
bank, the support system, the assessment table and the answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x1B, 0x24, 0x30)
DARK = RGBColor(0x2C, 0x39, 0x45)
SOFT = RGBColor(0x7C, 0x89, 0x95)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAGE = RGBColor(0xF6, 0xF9, 0xFB)
SUNNY = RGBColor(0xE2, 0x90, 0x0A)
CLOUD = RGBColor(0x4E, 0x77, 0x96)
RAIN = RGBColor(0x1C, 0x6E, 0xB5)
RAINBOW = RGBColor(0x7B, 0x4B, 0xA8)
WIND = RGBColor(0x13, 0x86, 0x7C)
DESK = RGBColor(0xCE, 0x41, 0x3D)
L_SUNNY = RGBColor(0xFD, 0xF3, 0xDB)
L_CLOUD = RGBColor(0xE9, 0xEF, 0xF4)
L_RAIN = RGBColor(0xE2, 0xED, 0xF8)
L_RAINBOW = RGBColor(0xF2, 0xEB, 0xF9)
L_WIND = RGBColor(0xDF, 0xF2, 0xF0)
L_DESK = RGBColor(0xFB, 0xE9, 0xE8)
L_GREY = RGBColor(0xF0, 0xF2, 0xF4)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 75
_counter = {"n": 0}

# ------------------------------------------------------------------ content

STOPS = [("☀️", "Sunny Station", "First sounds", "7–17 min", SUNNY, L_SUNNY),
         ("🌈", "Rainbow Station", "Short vowels", "17–27 min", RAINBOW,
          L_RAINBOW),
         ("☁️", "Cloud Station", "Building words", "27–37 min", CLOUD,
          L_CLOUD),
         ("🌧️", "Rain Station", "Families & sight words", "37–59 min", RAIN,
          L_RAIN),
         ("🌬️", "Wind Station", "Sentences", "59–68 min", WIND, L_WIND),
         ("📺", "Weather Desk", "Story & report", "68–90 min", DESK, L_DESK)]

MISSION = [("🔤", "Read new words", "Sounds first, then the whole word."),
           ("🎮", "Play weather games", "Matching, sorting and word building."),
           ("🧩", "Build sentences", "Put the words in the right order."),
           ("📖", "Read a real story", "A rainy day story in five short parts.")]

STATION_SCENE = [("☀️", "sun"), ("☁️", "cloud"), ("🌧️", "rain"),
                 ("🌈", "rainbow"), ("🌬️", "wind"), ("📺", "report")]

TALK_QS = [("👀", "What do you see?"), ("☀️", "Is it sunny?"),
           ("🌧️", "Is it rainy?"), ("🌤️", "Is the sky blue?")]
TALK_STARTERS = ["I see ______.", "It is ______.", "I like ______ weather."]

VOCAB_CARDS = [
    ("☀️", "SUN", "the big light in the sky", "The sun is hot.",
     "What colour is the sun?", SUNNY, L_SUNNY),
    ("🌧️", "RAIN", "water that falls from the clouds", "I see the rain.",
     "Is the rain wet or dry?", RAIN, L_RAIN),
    ("☁️", "CLOUD", "a white or grey puff in the sky", "The cloud is big.",
     "Can you point to a cloud?", CLOUD, L_CLOUD),
    ("🌬️", "WIND", "air that moves and pushes things", "The wind is cold.",
     "What can the wind move?", WIND, L_WIND),
]

WORD_BANK = [("❄️", "snow", "soft white flakes"),
             ("⛈️", "storm", "loud rain and wind"),
             ("🔥", "hot", "very, very warm"),
             ("🧊", "cold", "not warm at all"),
             ("💧", "wet", "full of water"),
             ("🌤️", "sky", "high up above us"),
             ("🧥", "coat", "you wear it when it is cold"),
             ("☂️", "umbrella", "it keeps the rain off you"),
             ("🌈", "rainbow", "colours after the rain")]

SOUNDS = [("S", "/s/", "☀️", "SUN", "sss - un", SUNNY, L_SUNNY),
          ("R", "/r/", "🌧️", "RAIN", "rrr - ain", RAIN, L_RAIN),
          ("C", "/c/", "☁️", "CLOUD", "ccc - loud", CLOUD, L_CLOUD),
          ("W", "/w/", "🌬️", "WIND", "www - ind", WIND, L_WIND),
          ("M", "/m/", "🌫️", "MIST", "mmm - ist", RAINBOW, L_RAINBOW),
          ("H", "/h/", "🔥", "HOT", "hhh - ot", DESK, L_DESK)]

WHICH_FIRST = ("/s/", ["☀️", "☁️", "🌧️"])

SOUND_MATCH_A = [("/r/", ["🌧️", "☀️", "☁️"]),
                 ("/c/", ["☀️", "☁️", "🌧️"]),
                 ("/s/", ["☁️", "🌧️", "☀️"])]
SOUND_MATCH_B = [("/w/", ["🌬️", "☀️", "🌧️"]),
                 ("/m/", ["☁️", "🌫️", "🔥"]),
                 ("/h/", ["🌧️", "🔥", "🌬️"])]

VOWELS = [("A", "/a/", "h - aaa - t", [("🎩", "HAT"), ("🗺️", "MAP"),
                                       ("🧢", "CAP")], SUNNY, L_SUNNY),
          ("E", "/e/", "w - eee - t", [("💧", "WET"), ("🍎", "RED"),
                                       ("🛏️", "BED")], RAIN, L_RAIN),
          ("I", "/i/", "s - iii - t", [("🪑", "SIT"), ("🐘", "BIG"),
                                       ("🥊", "HIT")], RAINBOW, L_RAINBOW),
          ("O", "/o/", "h - ooo - t", [("🔥", "HOT"), ("🪵", "LOG"),
                                       ("🌫️", "FOG")], CLOUD, L_CLOUD)]

RAINBOW_BANDS = [("A", SUNNY, L_SUNNY), ("E", RAIN, L_RAIN),
                 ("I", RAINBOW, L_RAINBOW), ("O", CLOUD, L_CLOUD)]
RAINBOW_WORDS = ["HAT", "RED", "BIG", "HOT", "MAP", "FOG"]

WHICH_VOWEL_A = [("H _ T", "🎩", ["A", "E", "O"]),
                 ("B _ G", "🐘", ["A", "I", "U"]),
                 ("W _ T", "💧", ["E", "O", "A"])]
WHICH_VOWEL_B = [("M _ P", "🗺️", ["A", "I", "O"]),
                 ("L _ G", "🪵", ["E", "O", "I"])]

BLENDS = [("S", "U", "N", "☀️", "SUN", "I see the sun.", SUNNY, L_SUNNY),
          ("R", "U", "N", "🏃", "RUN", "I can run.", WIND, L_WIND),
          ("H", "A", "T", "🎩", "HAT", "My hat is wet.", RAINBOW, L_RAINBOW),
          ("W", "E", "T", "💧", "WET", "The dog is wet.", RAIN, L_RAIN),
          ("B", "I", "G", "🐘", "BIG", "I see a big cloud.", CLOUD, L_CLOUD),
          ("F", "O", "G", "🌫️", "FOG", "The fog is grey.", DESK, L_DESK)]

MISSING = [("S _ N", "☀️", ["U", "A", "I"]),
           ("H _ T", "🎩", ["E", "A", "O"]),
           ("L _ G", "🪵", ["O", "I", "E"])]

SCRAMBLE = [(["T", "A", "H"], "🎩"), (["G", "I", "B"], "🐘"),
            (["N", "U", "S"], "☀️"), (["T", "E", "W"], "💧")]

FAMILIES = [("-AT", "at", [("🎩", "HAT"), ("🦇", "BAT"), ("🐱", "CAT"),
                           ("🟫", "MAT")], SUNNY, L_SUNNY),
            ("-OG", "og", [("🐶", "DOG"), ("🪵", "LOG"), ("🌫️", "FOG")],
             CLOUD, L_CLOUD),
            ("-UN", "un", [("☀️", "SUN"), ("🏃", "RUN"), ("🎉", "FUN")],
             RAIN, L_RAIN)]

DROP_BUCKETS = [("-AT", SUNNY, L_SUNNY), ("-UN", RAIN, L_RAIN)]
DROP_WORDS = ["HAT", "SUN", "MAT", "RUN", "CAT", "FUN"]

ODD_ONE = [["CAT", "HAT", "RUN"], ["DOG", "LOG", "SIT"], ["SUN", "FUN", "MAP"]]

RHYMES = [("HAT", "🎩", ["CAT", "DOG", "SUN"]),
          ("LOG", "🪵", ["FOG", "HAT", "RUN"]),
          ("SUN", "☀️", ["FUN", "BIG", "WET"])]

FREEZE = [("☀️", "SUN!", "Make a big sunshine with your arms."),
          ("🌧️", "RAIN!", "Tap your fingers like raindrops."),
          ("🌬️", "WIND!", "Sway your hands side to side."),
          ("❄️", "SNOW!", "Float your fingers down slowly."),
          ("🧊", "FREEZE!", "Stop and hold very still.")]

FREEZE_CARDS = [("☀️", "SUN"), ("🌧️", "RAIN"), ("🎩", "HAT"), ("🐘", "BIG")]
FREEZE_CALLS = [("👉", "Point to SUN."), ("👉", "Point to RAIN."),
                ("📖", "Read HAT."), ("📖", "Read BIG.")]

SIGHT_SLIDES = [(["I"], "I can run.", "You say this word about yourself.",
                 SUNNY, L_SUNNY),
                (["see"], "I see the sun.", "This word is about your eyes.",
                 RAIN, L_RAIN),
                (["a", "the"], "I see a cloud.  ·  I see the cloud.",
                 "Two tiny words. They sit before a naming word.", CLOUD,
                 L_CLOUD),
                (["is", "my"], "My hat is wet.",
                 "One word joins. One word says it belongs to you.", RAINBOW,
                 L_RAINBOW)]

SIGHT_FIND = ["I", "see", "a", "the", "is", "my"]

BUILD_SENTENCES = [("I see the sun.", ["I", "see", "the", "sun"], "☀️",
                    SUNNY, L_SUNNY),
                   ("The sun is hot.", ["The", "sun", "is", "hot"], "🔥",
                    DESK, L_DESK),
                   ("My hat is wet.", ["My", "hat", "is", "wet"], "🎩",
                    RAIN, L_RAIN)]

PUZZLES = [(["see", "I", "sun", "the"], "☀️", 4),
           (["hot", "is", "sun", "The"], "🔥", 4),
           (["wet", "hat", "My", "is"], "🎩", 4),
           (["can", "I", "run"], "🏃", 3)]

SENT_MATCH = [("The hat is wet.", ["🎩", "☀️", "🐶"]),
              ("I see the rain.", ["🌈", "🌧️", "🔥"])]

REAL_SILLY = ["The sun is hot.", "The sun is wet.", "My coat is red.",
              "The rain is hot."]

STORY_PREDICT = [("🌧️", "a rainy day"), ("☀️", "a hot day"),
                 ("❄️", "a snow day")]

STORY = [
    ("Part 1", "🪟", RAIN, L_RAIN,
     ["Sam looks out the window.",
      "It is a rainy day.",
      "Drip, drip, drip. The rain taps on the glass.",
      "Sam does not feel sad. He has a plan."],
     "It is a rainy day.", "What does Sam see out the window?"),
    ("Part 2", "🧥", DESK, L_DESK,
     ["Sam puts on his red coat.",
      "He gets his hat.",
      "He pulls on his big boots.",
      "\"I am ready!\" says Sam."],
     "He gets his hat.", "What does Sam put on first?"),
    ("Part 3", "💧", WIND, L_WIND,
     ["Sam goes out in the rain.",
      "He sees a big puddle.",
      "Sam jumps in the puddle. Splash!",
      "Sam laughs and laughs."],
     "He sees a big puddle.", "What does Sam jump in?"),
    ("Part 4", "☀️", SUNNY, L_SUNNY,
     ["Then the rain stops.",
      "The clouds go away.",
      "Sam looks up at the sky. He sees the sun.",
      "The sun is warm on his face."],
     "He sees the sun.", "What happens to the rain?"),
    ("Part 5", "🌈", RAINBOW, L_RAINBOW,
     ["A rainbow is in the sky.",
      "It is red and gold.",
      "\"Look at that!\" says Sam.",
      "\"I like rainy days,\" he says with a smile."],
     "A rainbow is in the sky.", "How does Sam feel at the end?"),
]

DETECT_A = [("🧒", "Who is in the story?", ["Sam", "Mia", "Ben"], "Part 1"),
            ("🌧️", "What is the weather like?", ["Sunny", "Rainy", "Snowy"],
             "Part 1")]
DETECT_B = [("🧥", "What does Sam put on?",
             ["A red coat", "A blue shirt", "A green hat"], "Part 2"),
            ("🌈", "What does Sam see after the rain?",
             ["A rainbow", "A train", "A bird"], "Part 5")]

SEQUENCE = [("A", "☀️", "Sam sees the sun."), ("B", "🧥", "Sam puts on his coat."),
            ("C", "🌧️", "It rains."), ("D", "🌈", "Sam sees a rainbow.")]

EVIDENCE = [("Where does Sam jump?", "Part 3", WIND, L_WIND),
            ("What colour is Sam's coat?", "Part 2", DESK, L_DESK),
            ("What is in the sky at the end?", "Part 5", RAINBOW, L_RAINBOW)]

REPORT_PICS = [("☀️", "a sunny day", "It is ______.",
                ["sunny", "hot", "bright"], SUNNY, L_SUNNY),
               ("🌧️", "a rainy day", "It is ______.",
                ["rainy", "wet", "grey"], RAIN, L_RAIN)]

REPORT_FRAMES = [("📺", "Today it is ______."), ("👀", "I see ______."),
                 ("💛", "I like ______.")]

FINAL_ROUNDS = [("🔤", "SOUND", "What sound does SUN start with?", "☀️",
                 SUNNY, L_SUNNY),
                ("🧩", "WORD", "Read this word:  RAIN", "🌧️", RAIN, L_RAIN),
                ("📖", "SENTENCE", "Read this:  I see the sun.", "🌈",
                 RAINBOW, L_RAINBOW)]

STAR_LINES = [("🔤", "You read words!"), ("🧩", "You blended sounds!"),
              ("📖", "You read sentences!"), ("📚", "You read a story!"),
              ("🔎", "You answered questions!")]

GAMES = [("☀️", "WEATHER MEMORY", "Match the word to its picture.",
          ["SUN  ·  RAIN  ·  HAT", "☀️   🌧️   🎩"], SUNNY, L_SUNNY),
         ("☁️", "MISSING LETTER CLOUD", "Fill the gap in the cloud.",
          ["S _ N", "A  ·  U  ·  I"], CLOUD, L_CLOUD),
         ("🌧️", "WORD OR PICTURE?", "Read the word, then point.",
          ["RAIN", "🌧️   🔥   🎩"], RAIN, L_RAIN),
         ("🌈", "RHYME RAINBOW", "Which one rhymes?",
          ["HAT", "CAT  ·  DOG  ·  SUN"], RAINBOW, L_RAINBOW),
         ("💨", "FIND THE WORD", "Find one word in the sentence.",
          ["I see the sun.", "Find SUN."], WIND, L_WIND),
         ("🌦️", "WEATHER WORD SORT", "Sort the pictures under the words.",
          ["HOT  ·  COLD  ·  WET", "🔥   🧊   💧"], DESK, L_DESK)]

GAME_WHEN = ["She finishes an activity early.",
             "Her focus starts to drop.",
             "One sound needs more practice.",
             "You have five spare minutes.",
             "She asks to play one more."]

SUPPORT_LEVELS = [("🟢", "WEATHER HINT", "Needs a picture or a first sound",
                   "Point at the picture, then the first letter.", WIND,
                   L_WIND),
                  ("🟡", "WEATHER MISSION", "Solves it with a little help",
                   "Ask one small question, then wait quietly.", SUNNY,
                   L_SUNNY),
                  ("⭐", "WEATHER STAR", "Reads it on her own",
                   "Ask her to build a new sentence with the word.", RAINBOW,
                   L_RAINBOW)]

SUPPORT_LADDER = [("HINT 1", "Look at the first letter.", WIND),
                  ("HINT 2", "Say each sound.", RAIN),
                  ("HINT 3", "Blend the sounds together.", RAINBOW),
                  ("HINT 4", "Let's read it together.", SUNNY),
                  ("HINT 5", "I read it, then you read it.", DESK)]

PRAISE = ["\"Let's solve this word together.\"", "\"Good try.\"",
          "\"Look closely at the first sound.\"", "\"You are getting it.\"",
          "\"Let's try it one more time.\"", "\"You found it yourself!\""]

NEVER_SAY = ["\"You can't read this.\"", "\"That's wrong.\"",
             "\"You should know this by now.\""]

RHYTHM = [("0–7", "Welcome and weather talk"), ("7–17", "First sounds"),
          ("17–27", "Short vowels"), ("27–37", "Blending words"),
          ("37–44", "Word families"), ("44–49", "Weather Freeze break"),
          ("49–59", "Sight words"), ("59–68", "Sentences"),
          ("68–78", "The story"), ("78–85", "Questions"),
          ("85–90", "Report and celebrate")]

RUBRIC_SKILLS = ["Beginning sounds", "Short vowels", "CVC words", "Blending",
                 "Word families", "Sight words", "Phrase reading",
                 "Sentence reading", "Story reading", "Comprehension",
                 "Speaking"]

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
             RGBColor(0xE1, 0xE5, 0xE8))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), RAIN)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), INK)
    msg = "🌦️ The Weather Reading Adventure  |  Grade 1  |  90 min"
    if stop:
        msg += f"  |  {stop}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg,
       size=10, color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32),
       f"{n} / {TOTAL}", size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, stop, accent=RAIN, bg=PAGE):
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


def one_task(slide, text, color=RAIN, top=1.34):
    """States the single job of this slide in words the child understands."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4),
              text, size=16, bold=True, color=color)


def hint(slide, text, top=6.42, label="🔦 WEATHER HINT", fill=L_SUNNY,
         color=SUNNY):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46),
              fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.7), Inches(0.34),
       label, size=12, bold=True, color=color)
    tb(slide, Inches(3.6), Inches(top + 0.05), Inches(9.0), Inches(0.36), text,
       size=13, bold=True, color=INK)


def i_we_you(slide, top=6.42):
    steps = [("I READ", "Teacher first", DESK), ("WE READ", "Together", SUNNY),
             ("YOU READ", "Your turn!", WIND)]
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


def picture_rows(slide, items, start_index, accent, label="STARTS WITH"):
    """A target sound, then three pictures with no captions to give it away."""
    tops = numbered_rows(slide, len(items), start_index, accent,
                         top_start=1.95, gap=1.5)
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


def sound_slide(index):
    """One letter, its sound, a picture and the weather word it starts."""
    letter, sound, pic, word, stretch, color, light = SOUNDS[index]
    slide, n = new_slide(f"{letter}  is for  {word}", "SOUNDS", "7–17 min",
                         "Sunny Station", color)
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


def vocab_slide(index):
    """Picture, say it, the word, a simple meaning, then a sentence."""
    pic, word, meaning, sentence, ask, color, light = VOCAB_CARDS[index]
    slide, n = new_slide(f"{pic}  Weather Word — {word}", "WEATHER WORDS",
                         "0–7 min", "Sunny Station", color)
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
    hint(slide, "If she cannot say it, say it first and let her copy you.",
         6.42)


def blend_slide(index):
    """Three sounds, one arrow, one whole word, then a sentence."""
    a, b, c, pic, word, sentence, color, light = BLENDS[index]
    slide, n = new_slide(f"{a} + {b} + {c}  →  {word}", "BLENDING",
                         "27–37 min", "Cloud Station", color)
    one_task(slide, "Touch each cloud, say the sound, then push them together.",
             color)
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


def vowel_slide(index):
    """One short vowel, its stretch sound, and three words that use it."""
    letter, sound, stretch, words, color, light = VOWELS[index]
    slide, n = new_slide(f"🌈 Short {letter}  —  {sound}", "SHORT VOWELS",
                         "17–27 min", "Rainbow Station", color)
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


def family_slide(index):
    """One word family, its ending, and the words that rhyme inside it."""
    name, ending, words, color, light = FAMILIES[index]
    slide, n = new_slide(f"☔ The {name} Family", "WORD FAMILIES", "37–44 min",
                         "Rain Station", color)
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
    """One or two sight words: see it big, say it, find it in a sentence."""
    words, sentence, note, color, light = SIGHT_SLIDES[index]
    heading = "  ·  ".join(w.upper() for w in words)
    slide, n = new_slide(f"🌧️ Sight Word — {heading}", "SIGHT WORDS",
                         "49–59 min", "Rain Station", color)
    if len(words) == 1:
        one_task(slide, "This word is a friend. We know it by looking, not "
                        "sounding it out.", color)
    else:
        one_task(slide, "These words are friends. We know them by looking, "
                        "not sounding them out.", color)
    if len(words) == 1:
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
       "⭐ FIND IT — point to it in this row", size=11, bold=True, color=color)
    word_cards(slide, 6.1, 5.15, SIGHT_FIND, color, light, width=1.0,
               gap=1.08, height=0.72, size=17)
    hint(slide, "Say it, clap it, then hunt for the same shape in the row.",
         6.42)


def build_slide(index, timing="49–59 min", stop="Rain Station"):
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


def story_slide(part, timing="68–78 min"):
    """Picture and part pill on the left, four big story lines on the right."""
    label, pic, color, light, lines, you_line, question = part
    slide, n = new_slide(f"🌧️ The Rainy Day — {label}", "STORY", timing,
                         "Weather Desk", color)
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
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), RAIN)
    tb(slide, Inches(0.7), Inches(0.75), Inches(12), Inches(1.2), "🌦️",
       size=64, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.95), Inches(12), Inches(1.0),
       "THE WEATHER READING ADVENTURE", size=44, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.95), Inches(12), Inches(0.5),
       "\"Become a Junior Weather Reporter!\"", size=20,
       color=RGBColor(0x9C, 0xC6, 0xEC), align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, name) in enumerate(STATION_SCENE):
        left = Inches(1.35 + i * 1.78)
        add_round(slide, left, Inches(3.65), Inches(1.6), Inches(1.5),
                  RGBColor(0x25, 0x31, 0x3E))
        tb(slide, left, Inches(3.85), Inches(1.6), Inches(0.7), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(4.6), Inches(1.6), Inches(0.4), name, size=11,
           color=RGBColor(0x9C, 0xC6, 0xEC), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.1), Inches(5.4), Inches(7.1), Inches(1.0), RAIN)
    tb(slide, Inches(3.1), Inches(5.62), Inches(7.1), Inches(0.6),
       "Grade 1  •  90 Minutes  •  Reading", size=24, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Sunny Station")
    fade(slide)


def s02_mission():
    slide, n = new_slide("🎯 Today's Mission", "WELCOME", "0–7 min",
                         "Sunny Station", RAIN)
    one_task(slide, "We will read words, play weather games, build sentences "
                    "and read a story.", RAIN)
    for i, (icon, name, detail) in enumerate(MISSION):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 2.2)
        add_round(slide, left, top, Inches(5.95), Inches(2.0), L_RAIN)
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
           Inches(0.32), f"STEP {i + 1}", size=11, bold=True, color=RAIN,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.35), Inches(12.35), Inches(0.5),
              SUNNY)
    tb(slide, Inches(0.5), Inches(6.45), Inches(12.35), Inches(0.36),
       "At the end you become a 🏆 WEATHER READING STAR", size=15, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def s03_map():
    slide, n = new_slide("🗺️ Your Weather Journey", "WELCOME", "0–7 min",
                         "Sunny Station", RAIN)
    one_task(slide, "Six stations. We stop at every one.", RAIN)
    for i, (emoji, name, focus, timing, color, light) in enumerate(STOPS):
        left = Inches(0.5 + i * 2.09)
        add_round(slide, left, Inches(1.9), Inches(1.9), Inches(3.9), light)
        add_round(slide, left, Inches(1.9), Inches(1.9), Inches(0.44), color)
        tb(slide, left, Inches(1.96), Inches(1.9), Inches(0.34),
           f"STOP {i + 1}", size=10, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(1.9), Inches(0.9), emoji, size=38,
           align=PP_ALIGN.CENTER)
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
              L_SUNNY)
    tb(slide, Inches(0.5), Inches(6.18), Inches(12.35), Inches(0.36),
       "🏆 AT THE FINISH", size=11, bold=True, color=SUNNY,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(6.5), Inches(12.35), Inches(0.36),
       "You read words, sentences and a whole story — and you become a "
       "Weather Reading Star.", size=15, bold=True, color=INK,
       align=PP_ALIGN.CENTER)


def s04_weather_talk():
    slide, n = new_slide("🌤️ Weather Talk — Just Look and Tell Me", "TALKING",
                         "0–7 min", "Sunny Station", SUNNY)
    one_task(slide, "No reading yet. Just look at the sky and talk to me.",
             SUNNY)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.2), Inches(4.4),
              L_SUNNY)
    tb(slide, Inches(0.5), Inches(2.2), Inches(5.2), Inches(1.7), "⛅",
       size=110, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(4.5), Inches(4.6), Inches(0.6),
       "Look out of the window too!", size=17, bold=True, color=SUNNY,
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
       "SAY IT LIKE THIS", size=12, bold=True, color=SUNNY)
    for i, starter in enumerate(TALK_STARTERS):
        top = Inches(4.68 + i * 0.6)
        add_round(slide, Inches(6.05), top, Inches(6.8), Inches(0.5), L_SUNNY)
        tb(slide, Inches(6.35), top + Inches(0.05), Inches(6.2), Inches(0.4),
           starter, size=20, bold=True, color=INK, font="Georgia")
    hint(slide, "Say your own answer first. She will copy the pattern.", 6.5)


def s05_sun():
    vocab_slide(0)


def s06_rain():
    vocab_slide(1)


def s07_cloud():
    vocab_slide(2)


def s08_wind():
    vocab_slide(3)


def s09_word_bank():
    slide, n = new_slide("🌦️ More Weather Words", "WEATHER WORDS", "0–7 min",
                         "Sunny Station", WIND)
    one_task(slide, "I say the word, you point to the picture.", WIND)
    for i, (pic, word, meaning) in enumerate(WORD_BANK):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.85 + row * 1.55)
        add_round(slide, left, top, Inches(3.9), Inches(1.4), L_GREY)
        tb(slide, left + Inches(0.15), top + Inches(0.3), Inches(0.9),
           Inches(0.8), pic, size=34, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.2), top + Inches(0.18), Inches(2.6),
           Inches(0.5), word, size=26, bold=True, color=WIND,
           font="Arial Black")
        tb(slide, left + Inches(1.2), top + Inches(0.78), Inches(2.6),
           Inches(0.5), meaning, size=12, color=DARK)
    hint(slide, "Only three or four new words today. The rest are for "
                "listening.", 6.6)


def s10_sound_s():
    sound_slide(0)


def s11_sound_r():
    sound_slide(1)


def s12_sound_c():
    sound_slide(2)


def s13_sound_w():
    sound_slide(3)


def s14_sound_m():
    sound_slide(4)


def s15_sound_h():
    sound_slide(5)


def s16_which_first():
    sound, pics = WHICH_FIRST
    slide, n = new_slide("👂 Which One Starts With /s/?", "LISTENING",
                         "7–17 min", "Sunny Station", SUNNY)
    one_task(slide, "Name each picture out loud, then listen to the first "
                    "sound.", SUNNY)
    add_round(slide, Inches(4.9), Inches(1.8), Inches(3.55), Inches(1.15),
              SUNNY)
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
    hint(slide, "Say all three names slowly, stretching the very first "
                "sound.", 6.5)


def s17_match_a():
    slide, n = new_slide("🎮 Game — Weather Sound Match", "GAME", "7–17 min",
                         "Sunny Station", RAIN)
    one_task(slide, "I say a sound. You point to the picture that starts "
                    "with it.", RAIN)
    picture_rows(slide, SOUND_MATCH_A, 1, RAIN)
    hint(slide, "Name the pictures first, then listen again.", 6.5)


def s18_match_b():
    slide, n = new_slide("🎮 Game — Three More Sounds", "GAME", "7–17 min",
                         "Sunny Station", WIND)
    one_task(slide, "Same game. Listen, then point.", WIND)
    picture_rows(slide, SOUND_MATCH_B, 4, WIND)
    hint(slide, "For /m/ hum with your lips closed. She will hear it.", 6.5)


def s19_sunny_done():
    badge_slide("☀️ Sunny Station Finished!", "7–17 min", "Sunny Station",
                "☀️", "BADGE 1 EARNED",
                "You know six letter sounds now.",
                [("🔤", "S, R, C, W, M and H."),
                 ("👂", "You heard the first sound in a word."),
                 ("🎮", "Six sound matches, all done."),
                 ("🗣️", "You said every sound out loud.")],
                "🌈 Rainbow Station — the middle sounds.", SUNNY, L_SUNNY)


def s20_vowel_a():
    vowel_slide(0)


def s21_vowel_e():
    vowel_slide(1)


def s22_vowel_i():
    vowel_slide(2)


def s23_vowel_o():
    vowel_slide(3)


def s24_rainbow_sort():
    slide, n = new_slide("🌈 Rainbow Vowel Sort", "SORTING", "17–27 min",
                         "Rainbow Station", RAINBOW)
    one_task(slide, "Read each word, then put it on the right rainbow band.",
             RAINBOW)
    for i, (letter, color, light) in enumerate(RAINBOW_BANDS):
        left = Inches(0.5 + i * 3.11)
        add_round(slide, left, Inches(1.85), Inches(2.9), Inches(2.85), light)
        add_round(slide, left, Inches(1.85), Inches(2.9), Inches(0.75), color)
        tb(slide, left, Inches(1.94), Inches(2.9), Inches(0.6), letter,
           size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left, Inches(2.75), Inches(2.9), Inches(0.4),
           f"/{letter.lower()}/", size=17, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(3.25), Inches(2.3), Inches(1.3),
           "____________\n\n____________", size=15, color=SOFT,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.85), Inches(12.35), Inches(0.4),
       "WORDS TO SORT — read each one first", size=12, bold=True,
       color=RAINBOW)
    word_cards(slide, 0.5, 5.3, RAINBOW_WORDS, INK, WHITE, width=1.9,
               gap=2.07, height=0.85, size=26)
    hint(slide, "Say the word slowly and listen for the sound in the middle.",
         6.42)


def s25_which_vowel_a():
    slide, n = new_slide("🎮 Game — Which Vowel Is Missing?", "GAME",
                         "17–27 min", "Rainbow Station", RAINBOW)
    one_task(slide, "The picture tells you the word. Pick the middle letter.",
             RAINBOW)
    choice_rows(slide, WHICH_VOWEL_A, 1, RAINBOW)
    hint(slide, "Say the word, then say only the middle sound on its own.",
         6.42)


def s26_which_vowel_b():
    slide, n = new_slide("🎮 Game — Two More Words", "GAME", "17–27 min",
                         "Rainbow Station", CLOUD)
    one_task(slide, "Last two. Look at the picture, then choose.", CLOUD)
    choice_rows(slide, WHICH_VOWEL_B, 4, CLOUD, top_start=2.1, gap=1.6,
                height=1.4)
    add_round(slide, Inches(0.5), Inches(5.3), Inches(12.35), Inches(0.85),
              L_RAINBOW)
    tb(slide, Inches(0.85), Inches(5.42), Inches(11.6), Inches(0.32),
       "⭐ WEATHER STAR CHALLENGE", size=11, bold=True, color=RAINBOW)
    tb(slide, Inches(0.85), Inches(5.74), Inches(11.6), Inches(0.42),
       "Read all four vowel sounds in a row:  /a/  /e/  /i/  /o/", size=18,
       bold=True, color=INK)
    hint(slide, "If she picks the wrong letter, read her word back to her.",
         6.42)


def s27_rainbow_done():
    badge_slide("🌈 Rainbow Station Finished!", "17–27 min",
                "Rainbow Station", "🌈", "BADGE 2 EARNED",
                "You know the four short vowel sounds.",
                [("🔤", "A, E, I and O in the middle."),
                 ("🌈", "Six words sorted onto the right band."),
                 ("🎮", "Five missing vowels found."),
                 ("👂", "You heard the middle sound on its own.")],
                "☁️ Cloud Station — pushing sounds together.", RAINBOW,
                L_RAINBOW)


def s28_blend_sun():
    blend_slide(0)


def s29_blend_run():
    blend_slide(1)


def s30_blend_hat():
    blend_slide(2)


def s31_blend_wet():
    blend_slide(3)


def s32_blend_big():
    blend_slide(4)


def s33_blend_fog():
    blend_slide(5)


def s34_missing_letter():
    slide, n = new_slide("☁️ Cloud Word Builder — Fill the Gap", "GAME",
                         "27–37 min", "Cloud Station", CLOUD)
    one_task(slide, "One letter blew away. Which one finishes the word?",
             CLOUD)
    choice_rows(slide, MISSING, 1, CLOUD)
    hint(slide, "Try each letter in the gap and say the word. One will sound "
                "right.", 6.42)


def s35_scramble():
    slide, n = new_slide("🎮 Mix the Clouds — Put Them in Order", "GAME",
                         "27–37 min", "Cloud Station", WIND)
    one_task(slide, "The letters are mixed up. The picture is the clue.",
             WIND)
    for i, (letters, pic) in enumerate(SCRAMBLE):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(5.95), Inches(2.1), L_GREY)
        add_oval(slide, left + Inches(0.25), top + Inches(0.18), Inches(0.5),
                 Inches(0.5), WIND)
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
           Inches(0.4), "write it here:  ______________", size=15, color=SOFT)
    hint(slide, "Ask which letter comes first. The first sound unlocks it.",
         6.6)


def s36_cloud_done():
    badge_slide("☁️ Cloud Station Finished!", "27–37 min", "Cloud Station",
                "☁️", "BADGE 3 EARNED",
                "You can push three sounds into one word.",
                [("🧩", "Six words blended: sun, run, hat, wet, big, fog."),
                 ("☁️", "Three missing letters found."),
                 ("🔀", "Four mixed-up words put in order."),
                 ("📖", "You read a sentence for every word.")],
                "🌧️ Rain Station — words that rhyme.", CLOUD, L_CLOUD)


def s37_family_at():
    family_slide(0)


def s38_family_og():
    family_slide(1)


def s39_family_un():
    family_slide(2)


def s40_raindrop_sort():
    slide, n = new_slide("☔ Raindrop Sort", "SORTING", "37–44 min",
                         "Rain Station", RAIN)
    one_task(slide, "Read each raindrop word, then drop it in the right "
                    "bucket.", RAIN)
    for i, (name, color, light) in enumerate(DROP_BUCKETS):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(1.85), Inches(5.95), Inches(2.6), light)
        add_round(slide, left, Inches(1.85), Inches(5.95), Inches(0.66),
                  color)
        tb(slide, left, Inches(1.95), Inches(5.95), Inches(0.46),
           f"{name}  bucket", size=20, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.65), Inches(5.95), Inches(0.6), "🪣",
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.6), Inches(3.35), Inches(4.75),
           Inches(0.95), "____________      ____________\n"
                          "____________      ____________", size=16,
           color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.6), Inches(12.35), Inches(0.4),
       "RAINDROP WORDS", size=12, bold=True, color=RAIN)
    for i, word in enumerate(DROP_WORDS):
        left = Inches(0.5 + i * 2.07)
        add_round(slide, left, Inches(5.05), Inches(1.9), Inches(1.1), WHITE)
        tb(slide, left, Inches(5.15), Inches(1.9), Inches(0.4), "💧", size=14,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(5.55), Inches(1.9), Inches(0.5), word, size=24,
           bold=True, color=RAIN, align=PP_ALIGN.CENTER, font="Arial Black")
    hint(slide, "Say two words together. If they rhyme, they share a bucket.",
         6.42)


def s41_odd_one():
    slide, n = new_slide("🎮 Who Does Not Belong?", "GAME", "37–44 min",
                         "Rain Station", WIND)
    one_task(slide, "Read all three out loud, then cross out the odd one.",
             WIND)
    for i, words in enumerate(ODD_ONE):
        top = Inches(1.95 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3), L_GREY)
        add_oval(slide, Inches(0.8), top + Inches(0.4), Inches(0.5),
                 Inches(0.5), WIND)
        tb(slide, Inches(0.8), top + Inches(0.46), Inches(0.5), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
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


def s42_rhyme():
    slide, n = new_slide("⭐ Make a Rhyme", "SPEAKING", "37–44 min",
                         "Rain Station", RAINBOW)
    one_task(slide, "Say a word that rhymes. You can use the choices if you "
                    "want.", RAINBOW)
    for i, (word, pic, options) in enumerate(RHYMES):
        top = Inches(1.95 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3), WHITE)
        tb(slide, Inches(0.8), top + Inches(0.3), Inches(0.9), Inches(0.7),
           pic, size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(1.9), top + Inches(0.28), Inches(2.4),
                  Inches(0.74), RAINBOW)
        tb(slide, Inches(1.9), top + Inches(0.4), Inches(2.4), Inches(0.52),
           word, size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, Inches(4.5), top + Inches(0.42), Inches(1.4), Inches(0.5),
           "rhymes with", size=13, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(6.05 + j * 2.25)
            add_round(slide, left, top + Inches(0.28), Inches(2.1),
                      Inches(0.74), L_RAINBOW)
            tb(slide, left, top + Inches(0.4), Inches(2.1), Inches(0.52),
               opt, size=22, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "A rhyme sounds the same at the end. Say both words twice.",
         6.42)


def s43_freeze():
    slide, n = new_slide("🌦️ Brain Break — Weather Freeze", "BREAK",
                         "44–49 min", "Rain Station", DESK)
    one_task(slide, "Stay in your chair. Move only your arms and hands.",
             DESK)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(3.6), Inches(4.4),
              L_DESK)
    tb(slide, Inches(0.5), Inches(2.5), Inches(3.6), Inches(1.6), "🙌",
       size=88, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.35), Inches(3.6), Inches(0.7),
       "5 MINUTES", size=28, bold=True, color=DESK, align=PP_ALIGN.CENTER,
       font="Georgia")
    tb(slide, Inches(0.7), Inches(5.15), Inches(3.2), Inches(0.7),
       "then sight words", size=14, color=SOFT, align=PP_ALIGN.CENTER)
    for i, (icon, call, action) in enumerate(FREEZE):
        top = Inches(1.85 + i * 0.9)
        add_round(slide, Inches(4.35), top, Inches(8.5), Inches(0.78), WHITE)
        tb(slide, Inches(4.6), top + Inches(0.14), Inches(0.6), Inches(0.5),
           icon, size=17, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(5.35), top + Inches(0.14), Inches(1.9),
                  Inches(0.5), DESK)
        tb(slide, Inches(5.35), top + Inches(0.22), Inches(1.9), Inches(0.36),
           call, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.5), top + Inches(0.18), Inches(5.1), Inches(0.44),
           action, size=16, bold=True, color=INK)
    hint(slide, "Say the calls quickly and out of order. That is the fun "
                "part.", 6.45)


def s44_freeze_read():
    slide, n = new_slide("🌦️ Freeze — Now Read One Word", "BREAK",
                         "44–49 min", "Rain Station", WIND)
    one_task(slide, "After each freeze, I give you one small reading job.",
             WIND)
    for i, (pic, word) in enumerate(FREEZE_CARDS):
        left = Inches(0.5 + i * 3.21)
        add_round(slide, left, Inches(1.85), Inches(3.0), Inches(2.4), WHITE)
        tb(slide, left, Inches(2.05), Inches(3.0), Inches(1.0), pic, size=44,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.35), Inches(3.15), Inches(2.3),
                  Inches(0.85), L_WIND)
        tb(slide, left + Inches(0.35), Inches(3.32), Inches(2.3),
           Inches(0.56), word, size=30, bold=True, color=WIND,
           align=PP_ALIGN.CENTER, font="Arial Black")
    for i, (icon, call) in enumerate(FREEZE_CALLS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(4.5 + row * 0.9)
        add_round(slide, left, top, Inches(5.95), Inches(0.78), L_GREY)
        tb(slide, left + Inches(0.3), top + Inches(0.16), Inches(0.5),
           Inches(0.5), icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.0), top + Inches(0.16), Inches(4.6),
           Inches(0.5), call, size=19, bold=True, color=INK)
    hint(slide, "Two jobs are pointing and two are reading. Keep it under a "
                "minute.", 6.45)


def s45_sight_i():
    sight_slide(0)


def s46_sight_see():
    sight_slide(1)


def s47_sight_a_the():
    sight_slide(2)


def s48_sight_is_my():
    sight_slide(3)


def s49_build_sun():
    build_slide(0)


def s50_build_hot():
    build_slide(1)


def s51_build_hat():
    build_slide(2)


def s52_rain_done():
    badge_slide("🌧️ Rain Station Finished!", "37–59 min", "Rain Station",
                "🌧️", "BADGE 4 EARNED",
                "You can read whole sentences now.",
                [("☔", "Three word families: -at, -og, -un."),
                 ("👀", "Six sight words you know by looking."),
                 ("🧩", "Three sentences built card by card."),
                 ("📖", "You read every sentence out loud.")],
                "🌬️ Wind Station — sentence puzzles.", RAIN, L_RAIN)


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


def s53_puzzles_a():
    slide, n = new_slide("🧩 Weather Sentence Puzzles", "PUZZLE", "59–68 min",
                         "Wind Station", WIND)
    one_task(slide, "The words are jumbled. Put them in an order that makes "
                    "sense.", WIND)
    for i in range(2):
        words, pic, _count = PUZZLES[i]
        puzzle_row(slide, Inches(1.9 + i * 2.3), i + 1, words, pic, WIND,
                   L_WIND)
    hint(slide, "Which word would you say first? A sentence starts with a "
                "capital letter.", 6.6)


def s54_puzzles_b():
    slide, n = new_slide("🧩 Two More Puzzles", "PUZZLE", "59–68 min",
                         "Wind Station", RAINBOW)
    one_task(slide, "Last two. Read your sentence back to me when you finish.",
             RAINBOW)
    for i in range(2):
        words, pic, _count = PUZZLES[i + 2]
        puzzle_row(slide, Inches(1.9 + i * 2.3), i + 3, words, pic, RAINBOW,
                   L_RAINBOW)
    hint(slide, "Move one card at a time. Read it out loud after every move.",
         6.6)


def s55_sentence_match():
    slide, n = new_slide("🎮 Weather Sentence Match", "GAME", "59–68 min",
                         "Wind Station", RAIN)
    one_task(slide, "Read the sentence, then point to the picture it tells "
                    "about.", RAIN)
    for i, (sentence, pics) in enumerate(SENT_MATCH):
        top = Inches(1.9 + i * 2.35)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(2.15),
                  L_GREY)
        add_round(slide, Inches(0.8), top + Inches(0.25), Inches(5.0),
                  Inches(0.9), WHITE)
        tb(slide, Inches(0.8), top + Inches(0.45), Inches(5.0), Inches(0.52),
           sentence, size=25, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Georgia")
        add_round(slide, Inches(0.8), top + Inches(1.3), Inches(5.0),
                  Inches(0.5), RAIN)
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
    hint(slide, "Read the sentence twice. The important word is the last "
                "one.", 6.6)


def s56_real_silly():
    slide, n = new_slide("🎮 Real or Silly?", "GAME", "59–68 min",
                         "Wind Station", SUNNY)
    one_task(slide, "Read it, then tell me: could that really happen?", SUNNY)
    for i, sentence in enumerate(REAL_SILLY):
        top = Inches(1.95 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.96), WHITE)
        add_oval(slide, Inches(0.8), top + Inches(0.22), Inches(0.5),
                 Inches(0.5), SUNNY)
        tb(slide, Inches(0.8), top + Inches(0.28), Inches(0.5), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.6), top + Inches(0.2), Inches(6.4), Inches(0.56),
           sentence, size=25, bold=True, color=INK, font="Georgia")
        for j, (label, color, light) in enumerate(
                [("✅ REAL", WIND, L_WIND), ("❌ SILLY", DESK, L_DESK)]):
            left = Inches(8.4 + j * 2.15)
            add_round(slide, left, top + Inches(0.23), Inches(1.95),
                      Inches(0.5), light)
            tb(slide, left, top + Inches(0.3), Inches(1.95), Inches(0.38),
               label, size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.42), Inches(12.35), Inches(0.46),
              L_SUNNY)
    tb(slide, Inches(0.8), Inches(6.48), Inches(11.7), Inches(0.36),
       "🔦 WEATHER HINT — For a silly one, ask: \"Can the sun really be "
       "wet?\" Then laugh about it together.", size=13, bold=True, color=INK)


def s57_story_intro():
    slide, n = new_slide("📖 Story Time — The Rainy Day", "STORY",
                         "68–78 min", "Weather Desk", DESK)
    one_task(slide, "Before we read: what do you think this story is about?",
             DESK)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.6), Inches(4.4),
              L_DESK)
    tb(slide, Inches(0.5), Inches(2.2), Inches(5.6), Inches(1.6), "🌧️",
       size=105, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.15), Inches(5.2), Inches(0.9),
       "THE RAINY DAY", size=38, bold=True, color=DESK,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.2), Inches(5.0), Inches(0.8),
       "five short parts  ·  I read first, then we read together", size=14,
       color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(6.4), Inches(1.9), Inches(6.45), Inches(0.5),
       "🤔 WHAT DO YOU THINK?", size=15, bold=True, color=DESK)
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
              L_SUNNY)
    tb(slide, Inches(6.65), Inches(6.49), Inches(5.95), Inches(0.36),
       "Remember your guess. We check it at the end!", size=13, bold=True,
       color=SUNNY)


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
    slide, n = new_slide("🔎 Story Detective — Question 1 and 2",
                         "COMPREHENSION", "78–85 min", "Weather Desk", RAIN)
    one_task(slide, "Point to your answer. You can look back at the story.",
             RAIN)
    question_rows(slide, DETECT_A, 1, RAIN)
    hint(slide, "Read the question again slowly, then read each choice.",
         6.42)


def s64_detect_b():
    slide, n = new_slide("🔎 Story Detective — Question 3 and 4",
                         "COMPREHENSION", "78–85 min", "Weather Desk",
                         RAINBOW)
    one_task(slide, "Two more. Point to the part of the story if you need it.",
             RAINBOW)
    question_rows(slide, DETECT_B, 3, RAINBOW)
    hint(slide, "Cover one wrong choice. Two choices is still good thinking.",
         6.42)


def s65_sequence():
    slide, n = new_slide("🔢 Put the Story in Order", "COMPREHENSION",
                         "78–85 min", "Weather Desk", WIND)
    one_task(slide, "These four things are mixed up. Number them 1 to 4.",
             WIND)
    for i, (letter, pic, event) in enumerate(SEQUENCE):
        left = Inches(0.5 + i * 3.21)
        add_round(slide, left, Inches(1.9), Inches(3.0), Inches(3.3), L_WIND)
        add_oval(slide, left + Inches(1.25), Inches(2.1), Inches(0.5),
                 Inches(0.5), WIND)
        tb(slide, left + Inches(1.25), Inches(2.16), Inches(0.5),
           Inches(0.4), letter, size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.75), Inches(3.0), Inches(0.9), pic, size=40,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.75), Inches(2.6), Inches(0.7),
           event, size=17, bold=True, color=INK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.9), Inches(4.5), Inches(1.2),
                  Inches(0.55), WHITE)
        tb(slide, left + Inches(0.9), Inches(4.6), Inches(1.2), Inches(0.4),
           "___", size=18, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.4), Inches(12.35), Inches(0.85),
              WHITE)
    tb(slide, Inches(0.85), Inches(5.5), Inches(11.6), Inches(0.32),
       "NOW WRITE THE ORDER", size=11, bold=True, color=WIND)
    tb(slide, Inches(0.85), Inches(5.82), Inches(11.6), Inches(0.42),
       "____  →  ____  →  ____  →  ____", size=26, bold=True, color=SOFT)
    hint(slide, "Ask what happened at the very start of the story.", 6.42)


def s66_evidence():
    slide, n = new_slide("🔎 Show Me the Answer", "COMPREHENSION",
                         "78–85 min", "Weather Desk", DESK)
    one_task(slide, "Do not just tell me. Point to the line in the story.",
             DESK)
    for i, (question, part, color, light) in enumerate(EVIDENCE):
        top = Inches(1.9 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), light)
        add_oval(slide, Inches(0.8), top + Inches(0.42), Inches(0.52),
                 Inches(0.52), color)
        tb(slide, Inches(0.8), top + Inches(0.48), Inches(0.52), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.16), Inches(4.7), Inches(0.74),
           question, size=19, bold=True, color=INK)
        add_round(slide, Inches(1.55), top + Inches(0.88), Inches(1.8),
                  Inches(0.4), WHITE)
        tb(slide, Inches(1.55), top + Inches(0.93), Inches(1.8), Inches(0.32),
           f"look in {part}", size=11, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        add_round(slide, Inches(6.4), top + Inches(0.3), Inches(6.15),
                  Inches(0.74), WHITE)
        tb(slide, Inches(6.65), top + Inches(0.4), Inches(5.7), Inches(0.34),
           "🔎  Go back and put your finger on the line.", size=13,
           bold=True, color=color)
        tb(slide, Inches(6.65), top + Inches(0.72), Inches(5.7), Inches(0.3),
           "☐ found it", size=11, color=SOFT)
    add_round(slide, Inches(0.5), Inches(6.46), Inches(12.35), Inches(0.44),
              L_WIND)
    tb(slide, Inches(0.8), Inches(6.53), Inches(11.7), Inches(0.34),
       "Finding the line in the book is the real skill here — not "
       "remembering it.", size=12, bold=True, color=WIND)


def s67_report_pics():
    slide, n = new_slide("📺 Your First Weather Report", "SPEAKING",
                         "85–88 min", "Weather Desk", SUNNY)
    one_task(slide, "Look at the picture, then finish the sentence out loud.",
             SUNNY)
    for i, (pic, caption, starter, words, color, light) in enumerate(
            REPORT_PICS):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(1.85), Inches(5.95), Inches(4.4),
                  light)
        tb(slide, left, Inches(2.1), Inches(5.95), Inches(1.5), pic, size=90,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(3.65), Inches(5.95), Inches(0.5), caption,
           size=17, color=SOFT, align=PP_ALIGN.CENTER, italic=True)
        add_round(slide, left + Inches(0.5), Inches(4.25), Inches(4.95),
                  Inches(0.85), WHITE)
        tb(slide, left + Inches(0.5), Inches(4.45), Inches(4.95),
           Inches(0.5), starter, size=30, bold=True, color=INK,
           align=PP_ALIGN.CENTER, font="Georgia")
        tb(slide, left, Inches(5.2), Inches(5.95), Inches(0.36),
           "words you could use", size=11, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        for j, word in enumerate(words):
            l = left + Inches(0.55 + j * 1.65)
            add_round(slide, l, Inches(5.55), Inches(1.5), Inches(0.6),
                      WHITE)
            tb(slide, l, Inches(5.66), Inches(1.5), Inches(0.42), word,
               size=17, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Say one whole sentence yourself first, then hand it over.",
         6.45)


def s68_be_reporter():
    slide, n = new_slide("🎤 Be the Weather Reporter!", "SPEAKING",
                         "85–88 min", "Weather Desk", DESK)
    one_task(slide, "Say two or three sentences. You are on TV now!", DESK)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(4.4), Inches(4.4),
              L_DESK)
    tb(slide, Inches(0.5), Inches(2.4), Inches(4.4), Inches(1.7), "🎤",
       size=96, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.35), Inches(4.0), Inches(0.9),
       "\"Good morning!\"", size=26, bold=True, color=DESK,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(5.3), Inches(4.0), Inches(0.7),
       "start like this if you get stuck", size=13, color=SOFT,
       align=PP_ALIGN.CENTER)
    for i, (icon, frame) in enumerate(REPORT_FRAMES):
        top = Inches(1.9 + i * 1.5)
        add_round(slide, Inches(5.25), top, Inches(7.6), Inches(1.3), WHITE)
        tb(slide, Inches(5.55), top + Inches(0.38), Inches(0.6),
           Inches(0.5), icon, size=18, align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.35), top + Inches(0.15), Inches(6.2),
           Inches(0.52), frame, size=26, bold=True, color=INK,
           font="Georgia")
        tb(slide, Inches(6.35), top + Inches(0.72), Inches(6.2),
           Inches(0.44), "______________________________", size=16,
           color=SOFT)
    hint(slide, "Let her hold something as a pretend microphone. It helps a "
                "lot.", 6.45)


def s69_final_challenge():
    slide, n = new_slide("🏆 Final Weather Reading Challenge", "CHALLENGE",
                         "88–90 min", "Weather Desk", RAINBOW)
    one_task(slide, "Three last jobs. You already know all of them.",
             RAINBOW)
    for i, (icon, label, task, pic, color, light) in enumerate(FINAL_ROUNDS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(4.4), light)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(0.6), color)
        tb(slide, left, Inches(1.96), Inches(3.9), Inches(0.4),
           f"{icon}  {label}", size=15, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.65), Inches(3.9), Inches(1.1), pic, size=56,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(3.9), Inches(3.3),
                  Inches(1.3), WHITE)
        tb(slide, left + Inches(0.45), Inches(4.15), Inches(3.0),
           Inches(0.85), task, size=19, bold=True, color=INK,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.05), Inches(5.35), Inches(1.8),
                  Inches(0.6), WHITE)
        tb(slide, left + Inches(1.05), Inches(5.46), Inches(1.8),
           Inches(0.42), "☐ done", size=15, bold=True, color=color,
           align=PP_ALIGN.CENTER)
    hint(slide, "Hints are still allowed here. Finishing is what matters.",
         6.45)


def s70_star():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26),
             SUNNY)
    for x, y in [(0.35, 3.9), (12.3, 3.9), (0.6, 5.4), (12.1, 5.4)]:
        tb(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.7), "⭐",
           size=24, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(0.8), Inches(12), Inches(1.1), "🌈",
       size=56, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.9), Inches(12), Inches(0.9),
       "WEATHER READING STAR!", size=42, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.85), Inches(12), Inches(0.5),
       "Look what you did today —", size=18,
       color=RGBColor(0xF3, 0xCE, 0x7E), align=PP_ALIGN.CENTER, italic=True)
    for i, (icon, line) in enumerate(STAR_LINES):
        left = Inches(1.35 + i * 2.15)
        add_round(slide, left, Inches(3.55), Inches(1.95), Inches(1.5),
                  RGBColor(0x26, 0x32, 0x3F))
        tb(slide, left, Inches(3.72), Inches(1.95), Inches(0.6), icon,
           size=20, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.12), Inches(4.32), Inches(1.7),
           Inches(0.62), line, size=11, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(2.4), Inches(5.25), Inches(8.5), Inches(1.1),
              SUNNY)
    tb(slide, Inches(2.4), Inches(5.5), Inches(8.5), Inches(0.62),
       "\"I AM A WEATHER READING STAR!\"", size=30, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(1.2), Inches(6.5), Inches(10.9), Inches(0.42),
       "🌈  📚  ⭐", size=16, bold=True,
       color=RGBColor(0xF3, 0xCE, 0x7E), align=PP_ALIGN.CENTER)
    footer(slide, n, "88–90 min", "Weather Desk")
    fade(slide)


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
               Inches(width - 0.8), Inches(0.5), card, size=16, bold=True,
               color=color, align=PP_ALIGN.CENTER)


def s71_games_a():
    slide, n = new_slide("🎲 Extra Game Bank — If There Is Time", "OPTIONAL",
                         "", "Weather Desk", WIND)
    one_task(slide, "Six spare games. Play any of them, in any order.", WIND)
    game_cards(slide, GAMES[:3], 1)
    hint(slide, "Every game uses today's words only. Nothing new to learn.",
         6.45)


def s72_games_b():
    slide, n = new_slide("🎲 Extra Game Bank — Three More", "OPTIONAL", "",
                         "Weather Desk", RAINBOW)
    one_task(slide, "Three more, plus when to use them.", RAINBOW)
    game_cards(slide, GAMES[3:], 4, width=2.9, gap=3.11)
    add_round(slide, Inches(9.83), Inches(1.9), Inches(3.02), Inches(4.3),
              L_GREY)
    tb(slide, Inches(10.05), Inches(2.1), Inches(2.6), Inches(0.4),
       "WHEN TO USE THESE", size=12, bold=True, color=RAINBOW)
    bullets(slide, Inches(10.05), Inches(2.6), Inches(2.6), Inches(3.4),
            GAME_WHEN, size=12, sp=10)
    hint(slide, "Stop a game while she is still enjoying it, not after.",
         6.45)


def s73_support():
    slide, n = new_slide("🧰 Support System — For the Teacher", "TEACHER", "",
                         "Weather Desk", CLOUD)
    one_task(slide, "Levels, the hint ladder, and the words to use.", CLOUD)
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
       color=CLOUD)
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
    tb(slide, Inches(6.85), Inches(4.25), Inches(3.0), Inches(0.4),
       "SAY THIS", size=13, bold=True, color=WIND)
    bullets(slide, Inches(6.85), Inches(4.72), Inches(3.0), Inches(2.1),
            PRAISE, size=11, color=WIND)
    tb(slide, Inches(10.0), Inches(4.25), Inches(2.85), Inches(0.4),
       "NEVER SAY", size=13, bold=True, color=DESK)
    bullets(slide, Inches(10.0), Inches(4.72), Inches(2.85), Inches(1.5),
            NEVER_SAY, size=11, color=DESK)
    add_round(slide, Inches(10.0), Inches(6.3), Inches(2.85), Inches(0.58),
              L_DESK)
    tb(slide, Inches(10.2), Inches(6.42), Inches(2.5), Inches(0.36),
       "After any hint, let her try again.", size=11, bold=True, color=DESK)


def s74_assessment():
    slide, n = new_slide("📋 End-of-Class Assessment", "TEACHER", "",
                         "Weather Desk", SUNNY)
    one_task(slide, "Tick one box per skill right after the lesson.", SUNNY)
    heads = ["SKILL", "INDEPENDENT", "WITH HELP", "NEEDS MORE PRACTICE"]
    widths = [5.15, 2.4, 2.4, 2.4]
    lefts = [0.5]
    for w in widths[:-1]:
        lefts.append(lefts[-1] + w)
    add_rect(slide, Inches(0.5), Inches(1.82), Inches(12.35), Inches(0.42),
             SUNNY)
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
              L_SUNNY)
    for i, item in enumerate(REVIEW_NEXT):
        tb(slide, Inches(0.8 + i * 4.1), Inches(6.45), Inches(3.9),
           Inches(0.34), f"{item}:  ____________", size=12, bold=True,
           color=INK)


def s75_answer_key():
    slide, n = new_slide("🔑 Answer Key", "TEACHER", "", "Weather Desk", RAIN)
    one_task(slide, "For the teacher only.", RAIN)
    cols = [
        ("SOUNDS AND VOWELS (10–26)",
         ["Starts with /s/: the sun picture",
          "Sound match: rain · cloud · sun ·",
          " wind · mist · hot (fire)",
          "Rainbow sort: A hat, map · E red ·",
          " I big · O hot, fog",
          "Missing vowel: O · I · E · A · O"], SUNNY),
        ("BLENDING AND FAMILIES (28–42)",
         ["Missing letter: SUN · HAT · LOG",
          "Mixed up: HAT · BIG · SUN · WET",
          "Raindrop sort: -AT hat, mat, cat ·",
          " -UN sun, run, fun",
          "Odd one out: RUN · SIT · MAP",
          "Rhymes: CAT · FOG · FUN"], CLOUD),
        ("SENTENCES (49–56)",
         ["Puzzles: I see the sun. · The sun is hot.",
          " My hat is wet. · I can run.",
          "Sentence match: the hat · the rain",
          "Real or silly: real · silly · real · silly",
          "Sight words: I, see, a, the, is, my"], WIND),
        ("STORY AND QUESTIONS (58–66)",
         ["Q1 Sam · Q2 Rainy · Q3 A red coat ·",
          " Q4 A rainbow",
          "Order: C → B → A → D  (3 → 2 → 1 → 4)",
          "Evidence: in the puddle (Part 3) ·",
          " red (Part 2) · a rainbow (Part 5)",
          "Story words: 120"], DESK),
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
              L_RAIN)
    tb(slide, Inches(0.8), Inches(6.76), Inches(11.7), Inches(0.32),
       "Accept any sensible answer for the rhymes, the weather report and "
       "the speaking tasks.", size=11, bold=True, color=RAIN)


BUILDERS = [
    s01_title, s02_mission, s03_map, s04_weather_talk, s05_sun, s06_rain,
    s07_cloud, s08_wind, s09_word_bank, s10_sound_s, s11_sound_r, s12_sound_c,
    s13_sound_w, s14_sound_m, s15_sound_h, s16_which_first, s17_match_a,
    s18_match_b, s19_sunny_done, s20_vowel_a, s21_vowel_e, s22_vowel_i,
    s23_vowel_o, s24_rainbow_sort, s25_which_vowel_a, s26_which_vowel_b,
    s27_rainbow_done, s28_blend_sun, s29_blend_run, s30_blend_hat,
    s31_blend_wet, s32_blend_big, s33_blend_fog, s34_missing_letter,
    s35_scramble, s36_cloud_done, s37_family_at, s38_family_og, s39_family_un,
    s40_raindrop_sort, s41_odd_one, s42_rhyme, s43_freeze, s44_freeze_read,
    s45_sight_i, s46_sight_see, s47_sight_a_the, s48_sight_is_my,
    s49_build_sun, s50_build_hot, s51_build_hat, s52_rain_done,
    s53_puzzles_a, s54_puzzles_b, s55_sentence_match, s56_real_silly,
    s57_story_intro, s58_story_1, s59_story_2, s60_story_3, s61_story_4,
    s62_story_5, s63_detect_a, s64_detect_b, s65_sequence, s66_evidence,
    s67_report_pics, s68_be_reporter, s69_final_challenge, s70_star,
    s71_games_a, s72_games_b, s73_support, s74_assessment, s75_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL}, built {len(prs.slides)}"

OUT = "Grade1_Weather_Reading_Adventure_90min.pptx"
prs.save(OUT)

story_words = sum(len(line.split()) for part in STORY for line in part[4])
with_notes = sum(1 for s in prs.slides
                 if s.has_notes_slide
                 and s.notes_slide.notes_text_frame.text.strip())
print(f"saved {OUT}")
print(f"slides: {len(prs.slides)}")
print(f"story words: {story_words}")
print(f"slides with notes: {with_notes}")
