"""Grade 2 reading lesson - 90 minutes, 64 slides, no speaker notes.

"The Amazing Grocery Store" - the child is the Grocery Store Helper working
through Produce, Dairy, Breakfast, Bakery, the Shopping Cart and the Checkout.
Reading is built one step at a time: hear, say, sound, blend, read, understand.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Shopping Hint strip, and slides 62-64 hold the support
system, the assessment table and the answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x1A, 0x24, 0x30)
DARK = RGBColor(0x2C, 0x37, 0x42)
SOFT = RGBColor(0x7C, 0x88, 0x94)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAGE = RGBColor(0xF8, 0xF7, 0xF3)
TOMATO = RGBColor(0xD6, 0x40, 0x2F)
LEAF = RGBColor(0x2E, 0x8B, 0x4E)
SKY = RGBColor(0x1E, 0x7B, 0xB8)
SUN = RGBColor(0xE2, 0xA0, 0x08)
CRUST = RGBColor(0xA9, 0x70, 0x2F)
BERRY = RGBColor(0x85, 0x46, 0xA8)
MINT = RGBColor(0x12, 0x85, 0x7A)
L_TOMATO = RGBColor(0xFC, 0xEA, 0xE6)
L_LEAF = RGBColor(0xE5, 0xF4, 0xEA)
L_SKY = RGBColor(0xE6, 0xF1, 0xF9)
L_SUN = RGBColor(0xFD, 0xF4, 0xDE)
L_CRUST = RGBColor(0xF8, 0xEE, 0xE1)
L_BERRY = RGBColor(0xF2, 0xEB, 0xF9)
L_MINT = RGBColor(0xE0, 0xF2, 0xF0)
L_GREY = RGBColor(0xF2, 0xF3, 0xF2)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 64
_counter = {"n": 0}

# ------------------------------------------------------------------ content

AISLES = [("🥕", "Produce", "First sounds", "0–17 min", LEAF, L_LEAF),
          ("🥛", "Dairy", "Short vowels", "17–27 min", SKY, L_SKY),
          ("🥣", "Breakfast", "Building words", "27–37 min", SUN, L_SUN),
          ("🍞", "Bakery", "Word families", "37–49 min", CRUST, L_CRUST),
          ("🛒", "Shopping Cart", "Words & sentences", "49–68 min", BERRY, L_BERRY),
          ("💳", "Checkout", "Story & questions", "68–90 min", TOMATO, L_TOMATO)]

PROMISES = [("🟢", "Shopping Hint", "A hint always comes before the answer."),
            ("🟡", "Shopping Mission", "Then you try it, and I go quiet."),
            ("⭐", "Bonus Shopper", "Finished early? There is always one more."),
            ("🛒", "One Cart", "Every word today goes in the same small cart.")]

STORE_SCENE = [("🍎", "apple"), ("🥛", "milk"), ("🍞", "bread"), ("🍌", "banana"),
               ("🥕", "carrot"), ("🛒", "cart")]

TALK_QS = [("🗣️", "What do you see?"), ("🍎", "Do you see an apple?"),
           ("🎨", "What color is the banana?"), ("⭐", "Which food do you like?")]

STARTERS = ["I see a ______.", "I like ______."]

SOUND_CARDS = [("🍎", "APPLE", "A", "/a/"), ("🍌", "BANANA", "B", "/b/"),
               ("🥕", "CARROT", "C", "/k/"), ("🥛", "MILK", "M", "/m/"),
               ("🧃", "JUICE", "J", "/j/"), ("🥚", "EGG", "E", "/e/")]

CART_A = [("M", [("🍎", "apple"), ("🥛", "milk"), ("🍞", "bread")]),
          ("B", [("🥕", "carrot"), ("🍌", "banana"), ("🥚", "egg")]),
          ("C", [("🥛", "milk"), ("🥕", "carrot"), ("🍎", "apple")])]
CART_B = [("A", [("🍎", "apple"), ("🧃", "juice"), ("🍌", "banana")]),
          ("J", [("🧀", "cheese"), ("🧃", "juice"), ("🍞", "bread")]),
          ("E", [("🥚", "egg"), ("🥕", "carrot"), ("🛒", "cart")])]
CART_C = [("BR", [("🍞", "bread"), ("🥛", "milk"), ("🍎", "apple")]),
          ("D", [("🥤", "drink"), ("🍎", "apple"), ("🥚", "egg")])]

ENDING = [("🛍️", "BAG", ["/g/", "/t/", "/p/"]), ("🥛", "MILK", ["/k/", "/t/", "/m/"]),
          ("🥚", "EGG", ["/g/", "/d/", "/n/"]), ("🥤", "CUP", ["/b/", "/p/", "/t/"])]

VOWEL_BASKETS = [("A", "🛍️", "bag", TOMATO, L_TOMATO), ("E", "🔴", "red", LEAF,
                                                        L_LEAF),
                 ("I", "🪑", "sit", SKY, L_SKY), ("O", "📦", "box", BERRY, L_BERRY)]

SORT_A = [("🛍️", "BAG", "/b/ /a/ /g/", ["A", "E", "I"]),
          ("🍯", "JAM", "/j/ /a/ /m/", ["A", "O", "E"]),
          ("🔴", "RED", "/r/ /e/ /d/", ["E", "I", "A"])]
SORT_B = [("🛏️", "BED", "/b/ /e/ /d/", ["A", "E", "O"]),
          ("🪑", "SIT", "/s/ /i/ /t/", ["I", "O", "E"]),
          ("📦", "BOX", "/b/ /o/ /x/", ["O", "A", "I"])]
SORT_BONUS = [("🔥", "HOT", "/h/ /o/ /t/", ["O", "E", "A"]),
              ("🥫", "CAN", "/k/ /a/ /n/", ["A", "I", "O"])]

GROCERY_WORDS = [("bag", "🛍️"), ("box", "📦"), ("jam", "🍯"), ("ham", "🍖"),
                 ("milk", "🥛"), ("bun", "🥐"), ("cup", "🥤"), ("pan", "🍳"),
                 ("egg", "🥚"), ("red", "🔴"), ("hot", "🔥"), ("wet", "💧"),
                 ("can", "🥫"), ("map", "🗺️"), ("sun", "☀️"), ("sit", "🪑"),
                 ("run", "🏃"), ("hop", "🐸")]

FIVE_STEPS = [("1", "TOUCH", "Touch each letter.", "B  A  G", TOMATO, L_TOMATO),
              ("2", "SAY", "Say each sound.", "/b/ /a/ /g/", SUN, L_SUN),
              ("3", "BLEND", "Push them together.", "b-a-g", SKY, L_SKY),
              ("4", "READ", "Read the whole word.", "BAG", LEAF, L_LEAF),
              ("5", "MATCH", "Find the picture.", "🛍️", BERRY, L_BERRY)]

BUILD_A = [("🛍️", "BAG", ["B", "A", "G"], ["/b/", "/a/", "/g/"], TOMATO, L_TOMATO),
           ("🍯", "JAM", ["J", "A", "M"], ["/j/", "/a/", "/m/"], BERRY, L_BERRY),
           ("🍖", "HAM", ["H", "A", "M"], ["/h/", "/a/", "/m/"], CRUST, L_CRUST)]
BUILD_B = [("📦", "BOX", ["B", "O", "X"], ["/b/", "/o/", "/x/"], SUN, L_SUN),
           ("🥤", "CUP", ["C", "U", "P"], ["/k/", "/u/", "/p/"], SKY, L_SKY),
           ("🪑", "SIT", ["S", "I", "T"], ["/s/", "/i/", "/t/"], LEAF, L_LEAF)]
BUILD_C = [("🥫", "CAN", ["C", "A", "N"], ["/k/", "/a/", "/n/"], TOMATO, L_TOMATO),
           ("🍳", "PAN", ["P", "A", "N"], ["/p/", "/a/", "/n/"], MINT, L_MINT),
           ("🗺️", "MAP", ["M", "A", "P"], ["/m/", "/a/", "/p/"], BERRY, L_BERRY)]

MISSING = [("🛍️", "B _ G", ["A", "E", "I"], "BAG"), ("🥤", "C _ P", ["U", "A", "O"],
                                                     "CUP"),
           ("📦", "B _ X", ["O", "E", "I"], "BOX"), ("🍳", "P _ N", ["A", "I", "O"],
                                                     "PAN")]

SCRAMBLE = [("🛍️", ["G", "A", "B"], "BAG"), ("📦", ["X", "O", "B"], "BOX"),
            ("🍯", ["M", "A", "J"], "JAM"), ("🥤", ["P", "U", "C"], "CUP")]

NO_CARDS = [("BAG", "🛍️"), ("JAM", "🍯"), ("BOX", "📦"),
            ("CUP", "🥤"), ("CAN", "🥫"), ("MAP", "🗺️")]

FAMILIES = [("-AT", ["cat", "hat", "bat"], TOMATO, L_TOMATO),
            ("-AN", ["can", "man", "pan"], LEAF, L_LEAF),
            ("-OP", ["hop", "top", "mop"], SKY, L_SKY),
            ("-IG", ["big", "pig", "dig"], BERRY, L_BERRY)]

FAMILY_A = [("-AT", ["CAT", "CAN", "HAT", "BIG"], TOMATO, L_TOMATO),
            ("-AN", ["CAN", "MAN", "HOP", "PAN"], LEAF, L_LEAF)]
FAMILY_B = [("-OP", ["HOP", "TOP", "BAG", "MOP"], SKY, L_SKY),
            ("-IG", ["BIG", "PIG", "DIG", "HAT"], BERRY, L_BERRY)]

NEW_WORDS = [("B", "-AT", "BAT", "🦇", TOMATO, L_TOMATO),
             ("V", "-AN", "VAN", "🚐", LEAF, L_LEAF),
             ("M", "-OP", "MOP", "🧹", SKY, L_SKY),
             ("P", "-IG", "PIG", "🐷", BERRY, L_BERRY)]

FREEZE = [("🍎", "Apple!", "Pretend to pick an apple."),
          ("🍌", "Banana!", "Pretend to peel a banana."),
          ("🛒", "Cart!", "Pretend to push the cart."),
          ("🧊", "Freeze!", "Stop and hold completely still."),
          ("👉", "Point to APPLE!", "Point at the written word."),
          ("📖", "Read BAG!", "Read the word out loud.")]

SIGHT_1 = [("I", "I can shop.", "🙋"), ("see", "I see a bag.", "👀"),
           ("a", "a red apple", "🍎")]
SIGHT_2 = [("the", "the big box", "📦"), ("my", "my bag", "🛍️"),
           ("can", "I can read.", "📖")]

FILL_A = [("I", ["a", "I", "see"]), ("see", ["see", "the", "my"]),
          ("a", ["can", "a", "I"]), ("the", ["my", "the", "see"])]
FILL_B = [("my", ["my", "can", "a"]), ("can", ["I", "can", "the"]),
          ("is", ["see", "is", "my"]), ("in", ["in", "a", "can"])]

PHRASES = [("a red apple", "🍎"), ("my big bag", "🛍️"), ("the cold milk", "🥛"),
           ("in the cart", "🛒"), ("a box of eggs", "🥚"), ("I can shop", "🏪")]

READ_STEPS = [("STEP 1", "I read", "Teacher reads it first.", TOMATO),
              ("STEP 2", "We read", "Teacher and child together.", SUN),
              ("STEP 3", "You read", "Child reads it alone — if ready.", LEAF)]

SENTENCES_A = [("I see an apple.", "🍎", TOMATO, L_TOMATO),
               ("The bag is big.", "🛍️", BERRY, L_BERRY)]
SENTENCES_B = [("I have a red box.", "📦", SUN, L_SUN),
               ("The milk is cold.", "🥛", SKY, L_SKY),
               ("I can shop.", "🏪", LEAF, L_LEAF)]

PUZZLES_A = [(["see", "I", "apple", "an"], "I see an apple.", "🍎", TOMATO,
              L_TOMATO),
             (["big", "bag", "The", "is"], "The bag is big.", "🛍️", BERRY, L_BERRY)]
PUZZLES_B = [(["milk", "The", "cold", "is"], "The milk is cold.", "🥛", SKY, L_SKY),
             (["can", "I", "shop"], "I can shop.", "🏪", LEAF, L_LEAF)]

SILLY = [("🍌", ["I see a banana.", "I see a shoe."]),
         ("🥛", ["The milk is hot.", "The milk is cold."]),
         ("🛒", ["The cart is big.", "The cart can run."])]

OWN_SENTENCE = [("🍎", "apple"), ("🥛", "milk"), ("🍞", "bread"), ("🛒", "cart")]

STORY_WATCH = [("bag", "🛍️"), ("apple", "🍎"), ("milk", "🥛"), ("box", "📦"),
               ("cart", "🛒"), ("red", "🔴")]

STORY = [
    ("Part 1", "🛒", TOMATO, L_TOMATO,
     ["Mia goes to the store with Dad.", "She has a small red bag.",
      "The store is big.", "Mia can see lots of food.", "Dad pushes the cart.",
      "She is happy to help."]),
    ("Part 2", "🍎", LEAF, L_LEAF,
     ["First, Mia gets an apple.", "The apple is red and big.",
      "Then she gets some milk.", "The milk is cold.", "It is good for her.",
      "Mia puts it in the cart."]),
    ("Part 3", "🥣", SUN, L_SUN,
     ["Mia sees a big box.", "It is a box of cereal.", "Mia can read the box.",
      "She puts the box in the cart.", "Dad says, \"Good job, Mia!\"",
      "Mia is a good helper."]),
    ("Part 4", "💳", BERRY, L_BERRY,
     ["At the checkout, Dad asks, \"Did we get everything?\"",
      "Mia looks in the cart.", "She sees an apple, milk, and a box.",
      "\"Yes!\" says Mia.", "They take the food home.", "Mia had a fun day."]),
]

EASY_LINES = ["The store is big.", "The milk is cold.", "Mia can read the box.",
              "\"Yes!\" says Mia.", "She is happy to help.", "Mia had a fun day."]

QUESTIONS_A = [("📍", "Where does Mia go?", ["To the store", "To school", "To bed"],
                "Part 1"),
               ("👨", "Who goes with Mia?", ["Dad", "Mom", "A cat"], "Part 1")]
QUESTIONS_B = [("🎨", "What color is her bag?", ["🔴 Red", "🔵 Blue", "🟢 Green"],
                "Part 1"),
               ("🍎", "What does she get first?", ["An apple", "Bread", "Juice"],
                "Part 2")]
QUESTIONS_C = [("🛒", "What does she put in the cart?",
                ["Milk and a box", "A hat", "A map"], "Part 2 and 3"),
               ("❓", "What does Dad ask?",
                ["\"Did we get everything?\"", "\"Where is the cat?\"",
                 "\"Can you run?\""], "Part 4")]
QUESTIONS_D = [("💬", "What does Mia say at the end?", ["\"Yes!\"", "\"No!\"",
                                                        "\"Stop!\""], "Part 4")]

RECEIPT = [("🍎", "APPLE"), ("🥛", "MILK"), ("📦", "BOX"), ("🛍️", "BAG"),
           ("🥚", "EGG")]

CHECKOUT_SENTENCES = [("I have milk.", "🥛", SKY, L_SKY),
                      ("I see an apple.", "🍎", TOMATO, L_TOMATO),
                      ("The bag is big.", "🛍️", BERRY, L_BERRY)]

FINAL_A = [("1", "🔤", "SOUND", "What sound does APPLE start with?", "🍎  APPLE",
            LEAF, L_LEAF),
           ("2", "📖", "WORD", "Read this word:", "BAG", TOMATO, L_TOMATO),
           ("3", "🧩", "BLEND", "Blend these sounds:", "M - I - L - K", SKY, L_SKY)]
FINAL_B = [("4", "📕", "SENTENCE", "Read this out loud:", "I see an apple.", SUN,
            L_SUN),
           ("5", "🔎", "STORY", "From the story:", "What did Mia buy?", BERRY,
            L_BERRY)]

CAN_READ = [("👂", "Beginning sounds", "a  b  c  m  j  e"),
            ("🔚", "Ending sounds", "bag  milk  egg  cup"),
            ("🧺", "Short vowels", "a  e  i  o"),
            ("🧩", "Blending", "/b/ /a/ /g/ → BAG"),
            ("📦", "CVC words", "bag  box  cup  can"),
            ("🍞", "Word families", "-at  -an  -op  -ig"),
            ("🛒", "Sight words", "I  see  a  the  my  can"),
            ("🔗", "Short phrases", "a red apple"),
            ("📕", "Sentences", "The milk is cold."),
            ("📖", "A whole story", "Mia Goes Shopping")]

EXTRA_GAMES = [("🍎", "WHAT'S IN MY BAG?", "Show a picture, she reads the word.",
                TOMATO),
               ("🛒", "CART SORT", "Sort words into FOOD and NOT FOOD.", LEAF),
               ("🔤", "MISSING LETTER", "B _ G  →  A, E or I?", SKY),
               ("🧩", "WORD PUZZLE", "G-A-B  →  she rebuilds BAG.", SUN),
               ("🗣️", "SAY IT!", "Show a picture: \"I see a ____.\"", BERRY),
               ("🔎", "FIND THE WORD", "Point to MILK inside a sentence.", MINT),
               ("🎯", "CORRECT OR SILLY?", "Which sentence matches the picture?",
                CRUST)]

SUPPORT_LEVELS = [("🟢", "SHOPPING HINT", "Picture + sound + teacher support",
                   "Say the word first, then let her repeat it.", LEAF, L_LEAF),
                  ("🟡", "SHOPPING MISSION", "She reads with limited help",
                   "Give the first sound only, then go quiet.", SUN, L_SUN),
                  ("⭐", "BONUS SHOPPER", "She reads alone or makes a sentence",
                   "Ask for a sentence using the word.", BERRY, L_BERRY)]

HINT_LADDER = [("HINT 1", "Look at the first letter.", TOMATO),
               ("HINT 2", "Say the sounds.", SUN),
               ("HINT 3", "Let's blend it together.", SKY),
               ("HINT 4", "Let's read it together.", LEAF),
               ("HINT 5", "I read it, then you repeat.", BERRY)]

HINT_BANK = ["Look at the first letter.", "Say the sounds slowly.",
             "Blend the sounds together.", "Look at the picture.",
             "Read it one more time.", "Look back at the story."]

PRAISE = ["\"Nice try!\"", "\"Let's look again.\"", "\"You found the first sound!\"",
          "\"Great blending!\"", "\"You figured that one out!\"",
          "\"Let's read it together.\""]

RUBRIC_SKILLS = ["Beginning sounds", "Short vowels", "CVC words", "Blending",
                 "Word families", "Sight words", "Phrase reading",
                 "Sentence reading", "Story reading", "Comprehension"]

NEXT_LESSON = ["Reading skill to continue", "Words needing practice",
               "Sight words needing practice"]

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


def footer(slide, n, timing="", aisle=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE1, 0xE4, 0xE0))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), LEAF)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), INK)
    msg = "🛒 The Amazing Grocery Store  |  Grade 2  |  90 min"
    if aisle:
        msg += f"  |  {aisle}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=10,
       color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, aisle, accent=LEAF, bg=PAGE):
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
    footer(slide, n, timing, aisle)
    fade(slide)
    return slide, n


def one_task(slide, text, color=TOMATO, top=1.34):
    """States the single job of this slide in words the child understands."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text,
              size=16, bold=True, color=color)


def hint(slide, text, top=6.42, label="🟢 SHOPPING HINT", fill=L_LEAF, color=LEAF):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.5), Inches(0.34), label,
       size=12, bold=True, color=color)
    tb(slide, Inches(3.4), Inches(top + 0.05), Inches(9.2), Inches(0.36), text,
       size=13, bold=True, color=INK)


def i_we_you(slide, top=6.42):
    steps = [("I READ", "Teacher", TOMATO), ("WE READ", "Together", SUN),
             ("YOU READ", "You!", LEAF)]
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


def cart_rows(slide, items, start_index):
    """Put It in the Cart: a lettered cart, then three picture choices."""
    tops = numbered_rows(slide, len(items), start_index, LEAF, top_start=1.9, gap=1.5)
    for (letter, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.3), Inches(0.9), Inches(0.72), "🛒",
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.4), top + Inches(0.36), Inches(1.7), Inches(0.62),
                  LEAF)
        tb(slide, Inches(2.4), top + Inches(0.44), Inches(1.7), Inches(0.46), letter,
           size=24, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Arial Black")
        for j, (emoji, label) in enumerate(options):
            left = Inches(4.45 + j * 2.75)
            add_round(slide, left, top + Inches(0.16), Inches(2.5), Inches(1.02),
                      L_LEAF)
            tb(slide, left, top + Inches(0.2), Inches(2.5), Inches(0.58), emoji,
               size=26, align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.78), Inches(2.5), Inches(0.36), label,
               size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)


def ending_rows(slide, items, start_index):
    """Ending sounds: a word, then three sounds to pick the last one from."""
    tops = numbered_rows(slide, len(items), start_index, MINT, top_start=1.9,
                         gap=1.12, height=1.0)
    for (emoji, word, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.18), Inches(0.9), Inches(0.62), emoji,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.5), top + Inches(0.2), Inches(2.3), Inches(0.6), word,
           size=26, bold=True, color=INK, font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(5.0 + j * 2.62)
            add_round(slide, left, top + Inches(0.19), Inches(2.4), Inches(0.62),
                      L_MINT)
            tb(slide, left, top + Inches(0.27), Inches(2.4), Inches(0.46), opt,
               size=20, bold=True, color=MINT, align=PP_ALIGN.CENTER,
               font="Arial Black")


def basket_rows(slide, items, start_index):
    """Sort the Shopping Words: word plus sounds, then vowel baskets."""
    tops = numbered_rows(slide, len(items), start_index, SKY, top_start=1.9, gap=1.5)
    for (emoji, word, sounds, options), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.32), Inches(0.9), Inches(0.7), emoji,
           size=28, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.4), top + Inches(0.2), Inches(2.4), Inches(0.6), word,
           size=28, bold=True, color=INK, font="Arial Black")
        tb(slide, Inches(2.4), top + Inches(0.82), Inches(2.4), Inches(0.4), sounds,
           size=15, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(5.45 + j * 2.42)
            add_round(slide, left, top + Inches(0.28), Inches(2.2), Inches(0.76),
                      L_SKY)
            tb(slide, left, top + Inches(0.3), Inches(2.2), Inches(0.34), "🧺",
               size=12, align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.6), Inches(2.2), Inches(0.42), opt,
               size=22, bold=True, color=SKY, align=PP_ALIGN.CENTER,
               font="Arial Black")


def build_rows(slide, items, top_start=1.9, gap=1.52):
    """Shopping Word Builder: letter cards, sounds, then the word on the shelf."""
    for i, (emoji, word, letters, sounds, color, light) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), light)
        tb(slide, Inches(0.8), top + Inches(0.32), Inches(0.9), Inches(0.68), emoji,
           size=30, align=PP_ALIGN.CENTER)
        for j, letter in enumerate(letters):
            left = Inches(2.0 + j * 1.35)
            add_round(slide, left, top + Inches(0.25), Inches(1.15), Inches(0.85),
                      WHITE)
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
           size=30, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")


def family_rows(slide, items, start_index):
    """Bakery Word Family: one family, four bread cards to sort."""
    tops = numbered_rows(slide, len(items), start_index, CRUST, top_start=2.0,
                         gap=2.15, height=1.9)
    for (family, words, color, light), top in zip(items, tops):
        tb(slide, Inches(1.4), top + Inches(0.5), Inches(0.9), Inches(0.7), "🍞",
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


def fill_rows(slide, items, start_index, top_start=1.95):
    """Fill the Shopping Cart: find one sight word among three."""
    for i, (target, options) in enumerate(items):
        top = Inches(top_start + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), WHITE)
        add_oval(slide, Inches(0.78), top + Inches(0.22), Inches(0.52), Inches(0.52),
                 BERRY)
        tb(slide, Inches(0.78), top + Inches(0.28), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.27), Inches(1.0), Inches(0.44), "READ:",
           size=13, bold=True, color=SOFT)
        add_round(slide, Inches(2.55), top + Inches(0.2), Inches(1.9), Inches(0.58),
                  BERRY)
        tb(slide, Inches(2.55), top + Inches(0.29), Inches(1.9), Inches(0.42), target,
           size=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(4.85 + j * 2.6)
            add_round(slide, left, top + Inches(0.14), Inches(2.4), Inches(0.7),
                      L_BERRY)
            tb(slide, left + Inches(0.12), top + Inches(0.26), Inches(0.5),
               Inches(0.44), "🛒", size=13, align=PP_ALIGN.CENTER)
            tb(slide, left + Inches(0.7), top + Inches(0.24), Inches(1.6),
               Inches(0.48), opt, size=20, bold=True, color=INK,
               align=PP_ALIGN.CENTER, font="Arial Black")


def phrase_grid(slide, items, top=1.95):
    for i, (phrase, emoji) in enumerate(items):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        t = Inches(top + row * 1.5)
        add_round(slide, left, t, Inches(5.95), Inches(1.3), L_BERRY)
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
    """Sentence Shopping: shuffled word cards plus an answer line."""
    tops = numbered_rows(slide, len(items), start_index, SUN, top_start=2.0,
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
                 TOMATO)
        tb(slide, Inches(0.75), top + Inches(0.76), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.45), top + Inches(0.6), Inches(1.0), Inches(0.72), emoji,
           size=30, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(0.5), Inches(3.5), Inches(0.66),
           question, size=19, bold=True, color=INK)
        tb(slide, Inches(2.6), top + Inches(1.18), Inches(3.5), Inches(0.36),
           f"show me this in {where}", size=11, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(6.3 + j * 2.2)
            add_round(slide, left, top + Inches(0.5), Inches(2.0), Inches(0.92),
                      L_TOMATO)
            tb(slide, left, top + Inches(0.66), Inches(2.0), Inches(0.62), opt,
               size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)


def story_slide(part, timing):
    """One story page: a big picture panel beside six short lines."""
    label, emoji, color, light = part[0], part[1], part[2], part[3]
    lines = part[4]
    slide, n = new_slide(f"📖 Mia Goes Shopping — {label}", "STORY", timing,
                         "Checkout", color)
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


def badge_slide(title, timing, aisle, emoji, badge_label, sub, lines, next_line):
    """Aisle-complete slide: a big badge beside what was just mastered."""
    slide, n = new_slide(title, "AISLE COMPLETE", timing, aisle, LEAF)
    one_task(slide, sub, LEAF)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.3), Inches(4.4), L_LEAF)
    tb(slide, Inches(0.5), Inches(2.6), Inches(5.3), Inches(1.7), emoji, size=88,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.3), Inches(0.7), badge_label,
       size=26, bold=True, color=LEAF, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.7), Inches(0.6),
       "⭐ shopping point earned", size=15, color=SOFT, align=PP_ALIGN.CENTER)
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.15), top, Inches(6.7), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.45), top + Inches(0.2), Inches(0.56), Inches(0.56),
                 WHITE)
        tb(slide, Inches(6.45), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.25), top + Inches(0.24), Inches(5.3), Inches(0.5), line,
           size=15, bold=True, color=INK)
    hint(slide, next_line, 6.45, "➡️ NEXT AISLE", L_SUN, SUN)
    return slide, n


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), LEAF)
    for x, y, c in [(0.6, 5.6, TOMATO), (12.05, 5.55, SUN)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(0.9), Inches(12), Inches(1.0), "🛒", size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.9), Inches(12), Inches(0.9),
       "WELCOME, SMART SHOPPER!", size=44, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.85), Inches(12), Inches(0.5),
       "Today you are helping in a grocery store!", size=21,
       color=RGBColor(0xA8, 0xD4, 0xB6), align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, label) in enumerate(STORE_SCENE):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(3.55), Inches(1.65), Inches(1.35),
                  RGBColor(0x24, 0x31, 0x3E))
        tb(slide, left, Inches(3.73), Inches(1.65), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.38), Inches(1.45), Inches(0.4), label,
           size=10, color=RGBColor(0xB4, 0xC4, 0xD0), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.3), Inches(5.2), Inches(6.7), Inches(1.15), LEAF)
    tb(slide, Inches(3.5), Inches(5.45), Inches(6.3), Inches(0.7),
       "Grade 2  •  90 Minutes  •  Reading", size=19, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Produce")
    fade(slide)


def s02_helper():
    slide, n = new_slide("⭐ You Are the Grocery Store Helper", "WELCOME", "0–7 min",
                         "Produce", LEAF)
    one_task(slide, "Today you are the Grocery Store Helper!", LEAF)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.5), L_LEAF)
    tb(slide, Inches(0.5), Inches(2.4), Inches(5.4), Inches(1.7), "🛒", size=88,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.35), Inches(5.4), Inches(0.7), "THAT'S YOU!",
       size=32, bold=True, color=LEAF, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.8), Inches(5.15), Inches(4.8), Inches(0.8),
       "A helper who reads", size=15, color=SOFT, align=PP_ALIGN.CENTER)
    lines = [("🗺️", "We visit six aisles today."),
             ("📖", "Each aisle has one reading job."),
             ("⭐", "Every job you finish earns a shopping point."),
             ("🏆", "At the end you become a Smart Shopper Champion.")]
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.9 + i * 1.18)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(1.02), L_SUN)
        add_oval(slide, Inches(6.5), top + Inches(0.22), Inches(0.58), Inches(0.58),
                 WHITE)
        tb(slide, Inches(6.5), top + Inches(0.28), Inches(0.58), Inches(0.44), icon,
           size=16, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.3), top + Inches(0.28), Inches(5.3), Inches(0.5), line,
           size=16, bold=True, color=INK)


def s03_map():
    slide, n = new_slide("🗺️ Our Store Map", "MAP", "0–7 min", "Produce", SUN)
    one_task(slide, "Six aisles. One reading job at each one.", SUN)
    for i, (emoji, place, job, when, color, light) in enumerate(AISLES):
        left = Inches(0.5 + i * 2.08)
        add_round(slide, left, Inches(1.95), Inches(1.9), Inches(3.5), light)
        tb(slide, left, Inches(2.18), Inches(1.9), Inches(0.8), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.05), Inches(3.05), Inches(1.8), Inches(0.9), place,
           size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.95), Inches(1.7), Inches(0.6), job,
           size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(4.7), Inches(1.6), Inches(0.42),
                  WHITE)
        tb(slide, left + Inches(0.15), Inches(4.74), Inches(1.6), Inches(0.34), when,
           size=9, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.65), Inches(12.35), Inches(0.9), L_LEAF)
    tb(slide, Inches(0.8), Inches(5.9), Inches(11.7), Inches(0.44),
       "🥕  →  🥛  →  🥣  →  🍞  →  🛒  →  💳    Finish all six and you are a SMART "
       "SHOPPER CHAMPION!", size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)


def s04_promises():
    slide, n = new_slide("🤝 How I Help You Today", "PROMISE", "0–7 min", "Produce",
                         MINT)
    one_task(slide, "You never read alone. Here is how it works.", MINT)
    for i, (icon, name, detail) in enumerate(PROMISES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(2.0 + row * 2.2)
        add_round(slide, left, top, Inches(5.95), Inches(1.95), L_MINT)
        add_oval(slide, left + Inches(0.32), top + Inches(0.6), Inches(0.75),
                 Inches(0.75), WHITE)
        tb(slide, left + Inches(0.32), top + Inches(0.7), Inches(0.75), Inches(0.55),
           icon, size=20, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.3), top + Inches(0.4), Inches(4.4), Inches(0.55),
           name, size=22, bold=True, color=MINT)
        tb(slide, left + Inches(1.3), top + Inches(1.0), Inches(4.4), Inches(0.75),
           detail, size=14, color=DARK)
    hint(slide, "Read all four out loud. Knowing help is coming makes her braver.",
         6.45)


def s05_store_scene():
    slide, n = new_slide("🏪 Picture Talk — Inside the Store", "WARM-UP", "0–7 min",
                         "Produce", TOMATO)
    one_task(slide, "Just look and talk. No reading yet!", TOMATO)
    for i, (emoji, label) in enumerate(STORE_SCENE):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_TOMATO)
        tb(slide, left, top + Inches(0.22), Inches(3.9), Inches(0.98), emoji,
           size=48, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.9), top + Inches(1.34), Inches(2.1),
                  Inches(0.58), WHITE)
        tb(slide, left + Inches(0.9), top + Inches(1.45), Inches(2.1), Inches(0.42),
           label, size=16, bold=True, color=TOMATO, align=PP_ALIGN.CENTER)
    hint(slide, "Point at whatever she names. Pointing back shows you are listening.",
         6.42, "🗣️ SPEAKING", L_TOMATO, TOMATO)


def s06_talk():
    slide, n = new_slide("🗣️ Talk About the Store", "SPEAKING", "0–7 min", "Produce",
                         TOMATO)
    one_task(slide, "Answer out loud. Try a whole sentence.", TOMATO)
    for i, (icon, question) in enumerate(TALK_QS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.35)
        add_round(slide, left, top, Inches(5.95), Inches(1.15), L_TOMATO)
        add_oval(slide, left + Inches(0.3), top + Inches(0.28), Inches(0.58),
                 Inches(0.58), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.34), Inches(0.58), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.33), Inches(4.6), Inches(0.52),
           question, size=18, bold=True, color=INK)
    for i, starter in enumerate(STARTERS):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(4.75), Inches(5.95), Inches(1.1), L_SUN)
        tb(slide, left + Inches(0.4), Inches(5.02), Inches(5.2), Inches(0.6),
           starter, size=26, bold=True, color=INK)
    hint(slide, "If she answers with one word, say it back as a full sentence first.",
         6.45)


def s07_first_sounds():
    slide, n = new_slide("👂 First Sounds — Listen and Say", "LEARN", "7–17 min",
                         "Produce", LEAF)
    one_task(slide, "I say the sound. You say it back to me.", LEAF)
    for i, (emoji, word, letter, sound) in enumerate(SOUND_CARDS):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.95 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_LEAF)
        tb(slide, left, top + Inches(0.14), Inches(3.9), Inches(0.7), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.84), Inches(3.6), Inches(0.48),
           word, size=19, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Arial Black")
        add_round(slide, left + Inches(0.85), top + Inches(1.36), Inches(2.2),
                  Inches(0.6), WHITE)
        tb(slide, left + Inches(0.85), top + Inches(1.46), Inches(2.2), Inches(0.44),
           f"{letter}  says  {sound}", size=15, bold=True, color=LEAF,
           align=PP_ALIGN.CENTER)
    hint(slide, "Say the sound, not the letter name. Stretch it: mmmmm.", 6.45)


def s08_cart_a():
    slide, n = new_slide("🛒 Game: Put It in the Cart", "GAME", "7–17 min",
                         "Produce", LEAF)
    one_task(slide, "Which food STARTS with the letter on the cart?", LEAF)
    cart_rows(slide, CART_A, 1)
    hint(slide, "Say each food name out loud first. Stretch the first sound.", 6.45)


def s09_cart_b():
    slide, n = new_slide("🛒 Put It in the Cart — Rounds 4 to 6", "GAME", "7–17 min",
                         "Produce", LEAF)
    one_task(slide, "Three more items for the cart.", LEAF)
    cart_rows(slide, CART_B, 4)
    hint(slide, "Too many choices? Cover one with your hand and leave only two.",
         6.45)


def s10_cart_c():
    slide, n = new_slide("🛒 Cart Challenge — The Last Two", "GAME", "7–17 min",
                         "Produce", LEAF)
    one_task(slide, "Two last items. Then the produce aisle is done!", LEAF)
    cart_rows(slide, CART_C, 7)
    add_round(slide, Inches(0.5), Inches(4.95), Inches(12.35), Inches(1.28), L_BERRY)
    tb(slide, Inches(0.8), Inches(5.08), Inches(11.7), Inches(0.38),
       "⭐ BONUS SHOPPER", size=13, bold=True, color=BERRY)
    tb(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.6),
       "Can you name one more food that starts with /m/?  How about /b/?", size=18,
       bold=True, color=INK)
    hint(slide, "Any real food counts, even one that is not in this store.", 6.42)


def s11_ending():
    slide, n = new_slide("🔚 Ending Sounds — What Do You Hear Last?", "LEARN",
                         "7–17 min", "Produce", MINT)
    one_task(slide, "Say the word. Now listen only to the LAST sound.", MINT)
    ending_rows(slide, ENDING, 1)
    hint(slide, "Whisper the word, then say the last sound loudly: ba-G!", 6.42)


def s12_produce_done():
    badge_slide("🥕 Produce Aisle Complete", "7–17 min", "Produce", "🥕",
                "POINT 1 EARNED",
                "You can hear the first and last sound in a word.",
                [("👂", "A, B, C, M, J and E sounds."),
                 ("🛒", "You filled the cart eight times."),
                 ("🔚", "You heard the last sound too."),
                 ("🗣️", "You said every food name out loud.")],
                "Dairy Aisle — the sound hiding in the MIDDLE of a word.")


def s13_baskets():
    slide, n = new_slide("🧺 Short Vowels — Four Baskets", "LEARN", "17–27 min",
                         "Dairy", SKY)
    one_task(slide, "Every word has a sound hiding in the middle.", SKY)
    for i, (letter, emoji, word, color, light) in enumerate(VOWEL_BASKETS):
        left = Inches(0.9 + i * 3.0)
        add_round(slide, left, Inches(1.95), Inches(2.75), Inches(4.1), light)
        tb(slide, left, Inches(2.15), Inches(2.75), Inches(0.72), "🧺", size=30,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.7), Inches(2.9), Inches(1.35), Inches(1.15),
                  color)
        tb(slide, left + Inches(0.7), Inches(3.05), Inches(1.35), Inches(0.85),
           letter, size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left, Inches(4.2), Inches(2.75), Inches(0.72), emoji, size=28,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.5), Inches(5.0), Inches(1.75), Inches(0.62),
                  WHITE)
        tb(slide, left + Inches(0.5), Inches(5.11), Inches(1.75), Inches(0.44), word,
           size=18, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Stretch the middle: b-aaa-g. That is the sound she is listening for.",
         6.42)


def s14_sort_a():
    slide, n = new_slide("🧺 Game: Sort the Shopping Words", "GAME", "17–27 min",
                         "Dairy", SKY)
    one_task(slide, "Which basket does this word belong in?", SKY)
    basket_rows(slide, SORT_A, 1)
    hint(slide, "Say the word slowly and hold the middle sound.", 6.45)


def s15_sort_b():
    slide, n = new_slide("🧺 Sort the Shopping Words — 4 to 6", "GAME", "17–27 min",
                         "Dairy", SKY)
    one_task(slide, "Three more words to sort.", SKY)
    basket_rows(slide, SORT_B, 4)
    hint(slide, "Baskets:  A  ·  E  ·  I  ·  O.  Say each one before she chooses.",
         6.45)


def s16_sort_bonus():
    slide, n = new_slide("⭐ Basket Bonus — Two More", "BONUS", "17–27 min", "Dairy",
                         BERRY)
    one_task(slide, "Bonus round! These two are a little trickier.", BERRY)
    basket_rows(slide, SORT_BONUS, 7)
    add_round(slide, Inches(0.5), Inches(4.95), Inches(12.35), Inches(1.28), L_SUN)
    tb(slide, Inches(0.8), Inches(5.08), Inches(11.7), Inches(0.38),
       "⭐ BONUS SHOPPER", size=13, bold=True, color=SUN)
    tb(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.6),
       "Pick any basket. Can you think of one more word that goes inside it?",
       size=18, bold=True, color=INK)
    hint(slide, "Skip this if she is tired. It is a bonus, not a requirement.", 6.42)


def s17_grocery_words():
    slide, n = new_slide("🛍️ Shopping Words We Will Read", "VOCABULARY", "17–27 min",
                         "Dairy", BERRY)
    one_task(slide, "Every word today comes from this list. Nothing new later.",
             BERRY)
    for i, (word, emoji) in enumerate(GROCERY_WORDS):
        col, row = i % 6, i // 6
        left = Inches(0.5 + col * 2.08)
        top = Inches(2.0 + row * 1.45)
        add_round(slide, left, top, Inches(1.9), Inches(1.32), L_BERRY)
        tb(slide, left, top + Inches(0.12), Inches(1.9), Inches(0.58), emoji,
           size=22, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), top + Inches(0.72), Inches(1.6),
                  Inches(0.5), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(0.8), Inches(1.6), Inches(0.36),
           word, size=16, bold=True, color=BERRY, align=PP_ALIGN.CENTER)
    hint(slide, "Read them across in a rhythm. She joins in wherever she can.", 6.42)


def s18_dairy_done():
    badge_slide("🥛 Dairy Aisle Complete", "17–27 min", "Dairy", "🥛",
                "POINT 2 EARNED", "You found the middle sound in every word.",
                [("🧺", "Baskets A, E, I and O."),
                 ("📦", "bag · jam · red · bed · sit · box."),
                 ("🔥", "Even the bonus words: hot and can."),
                 ("🛍️", "Eighteen shopping words are now yours.")],
                "Breakfast Aisle — where we build whole words from letters.")


def s19_five_steps():
    slide, n = new_slide("🥣 Our Five Steps for Any Word", "LEARN", "27–37 min",
                         "Breakfast", SUN)
    one_task(slide, "When a word looks hard, we always do these five steps.", SUN)
    for i, (num, label, detail, example, color, light) in enumerate(FIVE_STEPS):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(3.6), light)
        add_oval(slide, left + Inches(0.85), Inches(2.18), Inches(0.6), Inches(0.6),
                 color)
        tb(slide, left + Inches(0.85), Inches(2.28), Inches(0.6), Inches(0.42), num,
           size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(2.95), Inches(2.1), Inches(0.55), label,
           size=17, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left + Inches(0.18), Inches(3.58), Inches(1.95), Inches(0.75),
           detail, size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(4.45), Inches(1.9), Inches(0.75),
                  WHITE)
        tb(slide, left + Inches(0.2), Inches(4.62), Inches(1.9), Inches(0.5),
           example, size=15, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.75), Inches(12.35), Inches(0.5), L_LEAF)
    tb(slide, Inches(0.8), Inches(5.83), Inches(11.7), Inches(0.36),
       "Never skip step 5. Matching the picture is what makes the word mean "
       "something.", size=13, bold=True, color=LEAF)


def s20_build_a():
    slide, n = new_slide("🛒 Game: Shopping Word Builder", "GAME", "27–37 min",
                         "Breakfast", SUN)
    one_task(slide, "Say each sound. Then say the whole word fast.", SUN)
    build_rows(slide, BUILD_A)
    hint(slide, "Cover the last letter so she blends only two sounds first.", 6.5)


def s21_build_b():
    slide, n = new_slide("🛒 Word Builder — Three More Shelves", "GAME", "27–37 min",
                         "Breakfast", SUN)
    one_task(slide, "Every word you build goes onto the shelf.", SUN)
    build_rows(slide, BUILD_B)
    hint(slide, "Slide your finger under the letters, then sweep it fast.", 6.5)


def s22_build_c():
    slide, n = new_slide("🛒 Word Builder — The Last Three", "GAME", "27–37 min",
                         "Breakfast", SUN)
    one_task(slide, "Three left. Then the shelf is full!", SUN)
    build_rows(slide, BUILD_C)
    hint(slide, "These three all share the short A sound. Stretch the middle.", 6.5)


def s23_missing():
    slide, n = new_slide("🔤 Game: Missing Letter", "GAME", "27–37 min", "Breakfast",
                         MINT)
    one_task(slide, "One letter fell off the label. Which one goes back?", MINT)
    tops = numbered_rows(slide, len(MISSING), 1, MINT, top_start=1.9, gap=1.14,
                         height=1.02)
    for (emoji, pattern, options, _answer), top in zip(MISSING, tops):
        tb(slide, Inches(1.4), top + Inches(0.18), Inches(0.9), Inches(0.66), emoji,
           size=24, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.5), top + Inches(0.18), Inches(2.5), Inches(0.66),
                  L_MINT)
        tb(slide, Inches(2.5), top + Inches(0.26), Inches(2.5), Inches(0.5),
           pattern, size=26, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(5.5 + j * 2.4)
            add_round(slide, left, top + Inches(0.2), Inches(2.2), Inches(0.62),
                      WHITE)
            tb(slide, left, top + Inches(0.28), Inches(2.2), Inches(0.46), opt,
               size=22, bold=True, color=MINT, align=PP_ALIGN.CENTER,
               font="Arial Black")
    hint(slide, "Say the word out loud with each letter. Only one sounds right.",
         6.42)


def s24_scramble():
    slide, n = new_slide("🧩 Game: Word Puzzle", "GAME", "27–37 min", "Breakfast",
                         BERRY)
    one_task(slide, "The letters got mixed up. Put them back in order.", BERRY)
    tops = numbered_rows(slide, len(SCRAMBLE), 1, BERRY, top_start=1.9, gap=1.14,
                         height=1.02)
    for (emoji, letters, _answer), top in zip(SCRAMBLE, tops):
        tb(slide, Inches(1.4), top + Inches(0.18), Inches(0.9), Inches(0.66), emoji,
           size=24, align=PP_ALIGN.CENTER)
        for j, letter in enumerate(letters):
            left = Inches(2.5 + j * 1.25)
            add_round(slide, left, top + Inches(0.2), Inches(1.1), Inches(0.62),
                      L_BERRY)
            tb(slide, left, top + Inches(0.28), Inches(1.1), Inches(0.46), letter,
               size=24, bold=True, color=BERRY, align=PP_ALIGN.CENTER,
               font="Arial Black")
        tb(slide, Inches(6.5), top + Inches(0.3), Inches(0.5), Inches(0.44), "→",
           size=20, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(7.3), top + Inches(0.2), Inches(5.2), Inches(0.62),
                  WHITE)
        tb(slide, Inches(7.55), top + Inches(0.3), Inches(4.8), Inches(0.44),
           "your word:  ____________________", size=13, color=SOFT)
    hint(slide, "Find the vowel first. It almost always sits in the middle.", 6.42)


def s25_no_cards():
    slide, n = new_slide("📖 No Letter Cards — Just Read It", "PRACTICE",
                         "27–37 min", "Breakfast", LEAF)
    one_task(slide, "This time there are no letter cards. You can do it!", LEAF)
    for i, (word, emoji) in enumerate(NO_CARDS):
        col, row = i % 3, i // 3
        left = Inches(0.6 + col * 4.1)
        top = Inches(1.95 + row * 2.25)
        add_round(slide, left, top, Inches(3.85), Inches(2.05), L_LEAF)
        tb(slide, left, top + Inches(0.18), Inches(3.85), Inches(0.72), emoji,
           size=28, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.7), top + Inches(1.0), Inches(2.45),
                  Inches(0.82), WHITE)
        tb(slide, left + Inches(0.7), top + Inches(1.12), Inches(2.45), Inches(0.6),
           word, size=32, bold=True, color=LEAF, align=PP_ALIGN.CENTER,
           font="Arial Black")
    hint(slide, "If she stalls, put your finger under the first letter. Say nothing.",
         6.45)


def s26_breakfast_done():
    badge_slide("🥣 Breakfast Aisle Complete", "27–37 min", "Breakfast", "🥣",
                "POINT 3 EARNED", "You built nine whole words from single letters.",
                [("🧩", "bag · jam · ham · box · cup · sit."),
                 ("🥫", "can · pan · map as well."),
                 ("🔤", "You put missing letters back in place."),
                 ("📖", "And you read six words with no letter cards.")],
                "Bakery — words that rhyme and belong to the same family.")


def s27_families():
    slide, n = new_slide("🍞 Bakery Word Families", "LEARN", "37–44 min", "Bakery",
                         CRUST)
    one_task(slide, "Same ending sound = same word family.", CRUST)
    for i, (family, words, color, light) in enumerate(FAMILIES):
        left = Inches(0.6 + i * 3.12)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(3.9), light)
        add_round(slide, left + Inches(0.5), Inches(2.2), Inches(1.9), Inches(0.8),
                  color)
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
    hint(slide, "Read down each column. The ending never changes — only the first "
                "sound does.", 6.42)


def s28_family_a():
    slide, n = new_slide("🍞 Game: Bakery Word Family", "GAME", "37–44 min",
                         "Bakery", CRUST)
    one_task(slide, "Three of these four belong. Which one does not?", CRUST)
    family_rows(slide, FAMILY_A, 1)
    hint(slide, "Cover the first letter. Now only the ending is showing.", 6.42)


def s29_family_b():
    slide, n = new_slide("🍞 Bakery Word Family — Rounds 3 and 4", "GAME",
                         "37–44 min", "Bakery", CRUST)
    one_task(slide, "Two more families. Read every word out loud first.", CRUST)
    family_rows(slide, FAMILY_B, 3)
    hint(slide, "Read the family ending out loud before each round: -op, -ig.",
         6.42)


def s30_new_words():
    slide, n = new_slide("🥖 Bake a Brand New Word", "GAME", "37–44 min", "Bakery",
                         SUN)
    one_task(slide, "Add one letter to the front. What word comes out?", SUN)
    for i, (letter, family, word, emoji, color, light) in enumerate(NEW_WORDS):
        top = Inches(1.9 + i * 1.14)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.02), light)
        add_round(slide, Inches(1.0), top + Inches(0.2), Inches(1.3), Inches(0.68),
                  WHITE)
        tb(slide, Inches(1.0), top + Inches(0.28), Inches(1.3), Inches(0.5), letter,
           size=26, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, Inches(2.4), top + Inches(0.3), Inches(0.4), Inches(0.5), "+",
           size=20, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.95), top + Inches(0.2), Inches(1.7), Inches(0.68),
                  WHITE)
        tb(slide, Inches(2.95), top + Inches(0.28), Inches(1.7), Inches(0.5),
           family, size=24, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, Inches(4.8), top + Inches(0.26), Inches(0.6), Inches(0.5), "→",
           size=22, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(5.6), top + Inches(0.2), Inches(2.6), Inches(0.68),
                  WHITE)
        tb(slide, Inches(5.6), top + Inches(0.28), Inches(2.6), Inches(0.5), word,
           size=26, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, Inches(8.5), top + Inches(0.22), Inches(0.9), Inches(0.66), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, Inches(9.6), top + Inches(0.32), Inches(3.0), Inches(0.44),
           "read it · then use it", size=13, color=SOFT)
    hint(slide, "She does not need to know the word first. Blending will find it.",
         6.42)


def s31_break_intro():
    slide, n = new_slide("🧠 Brain Break — Shopping Freeze", "BREAK", "44–49 min",
                         "Bakery", MINT)
    one_task(slide, "Stay by your chair. Act out whatever I call.", MINT)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4), L_MINT)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.4), Inches(1.8), "🧊", size=90,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.4), Inches(0.7), "5 MINUTES",
       size=30, bold=True, color=MINT, align=PP_ALIGN.CENTER, font="Georgia")
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
        add_oval(slide, Inches(6.55), top + Inches(0.2), Inches(0.56), Inches(0.56),
                 WHITE)
        tb(slide, Inches(6.55), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.35), top + Inches(0.24), Inches(5.2), Inches(0.5), line,
           size=15, bold=True, color=INK)
    hint(slide, "Letting her call the commands is the part that resets her focus.",
         6.45)


def s32_freeze():
    slide, n = new_slide("🧊 Shopping Freeze — Six Calls", "BREAK", "44–49 min",
                         "Bakery", MINT)
    one_task(slide, "Listen for the call. Then act it out.", MINT)
    for i, (icon, call, action) in enumerate(FREEZE):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_MINT)
        tb(slide, left, top + Inches(0.18), Inches(3.9), Inches(0.7), icon, size=28,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.92), Inches(3.6), Inches(0.5),
           call, size=19, bold=True, color=MINT, align=PP_ALIGN.CENTER,
           font="Georgia")
        tb(slide, left + Inches(0.3), top + Inches(1.46), Inches(3.3), Inches(0.5),
           action, size=12, color=DARK, align=PP_ALIGN.CENTER)
    hint(slide, "Mix the order. The two reading calls keep the break on task.",
         6.45)


def s33_sight_1():
    slide, n = new_slide("🛒 Sight Words — Group 1", "LEARN", "49–59 min",
                         "Shopping Cart", BERRY)
    one_task(slide, "These three words we do NOT sound out. We just know them.",
             BERRY)
    for i, (word, phrase, emoji) in enumerate(SIGHT_1):
        left = Inches(0.7 + i * 4.1)
        add_round(slide, left, Inches(1.95), Inches(3.8), Inches(3.9), L_BERRY)
        add_round(slide, left + Inches(0.55), Inches(2.25), Inches(2.7),
                  Inches(1.2), BERRY)
        tb(slide, left + Inches(0.55), Inches(2.47), Inches(2.7), Inches(0.8), word,
           size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left, Inches(3.65), Inches(3.8), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.6), Inches(3.2), Inches(0.8),
                  WHITE)
        tb(slide, left + Inches(0.3), Inches(4.78), Inches(3.2), Inches(0.5),
           phrase, size=19, bold=True, color=BERRY, align=PP_ALIGN.CENTER)
    hint(slide, "Say the word, she repeats, then she finds it in the phrase below.",
         6.42)


def s34_sight_2():
    slide, n = new_slide("🛒 Sight Words — Group 2", "LEARN", "49–59 min",
                         "Shopping Cart", BERRY)
    one_task(slide, "Three more. Only three at a time — never a long list.", BERRY)
    for i, (word, phrase, emoji) in enumerate(SIGHT_2):
        left = Inches(0.7 + i * 4.1)
        add_round(slide, left, Inches(1.95), Inches(3.8), Inches(3.9), L_BERRY)
        add_round(slide, left + Inches(0.55), Inches(2.25), Inches(2.7),
                  Inches(1.2), BERRY)
        tb(slide, left + Inches(0.55), Inches(2.47), Inches(2.7), Inches(0.8), word,
           size=38, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left, Inches(3.65), Inches(3.8), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.6), Inches(3.2), Inches(0.8),
                  WHITE)
        tb(slide, left + Inches(0.3), Inches(4.78), Inches(3.2), Inches(0.5),
           phrase, size=19, bold=True, color=BERRY, align=PP_ALIGN.CENTER)
    hint(slide, "Go back to Group 1 for ten seconds before you start Group 2.",
         6.42)


def s35_fill_a():
    slide, n = new_slide("🛒 Game: Fill the Shopping Cart", "GAME", "49–59 min",
                         "Shopping Cart", BERRY)
    one_task(slide, "Read the word to put one item in the cart.", BERRY)
    fill_rows(slide, FILL_A, 1)
    hint(slide, "Read all three carts out loud. Listening beats looking.", 6.42)


def s36_fill_b():
    slide, n = new_slide("🛒 Fill the Cart — Items 5 to 8", "GAME", "49–59 min",
                         "Shopping Cart", BERRY)
    one_task(slide, "Four more items and the cart is full.", BERRY)
    fill_rows(slide, FILL_B, 5)
    hint(slide, "Seen this word already today? Say so. Repetition is the point.",
         6.42)


def s37_phrases():
    slide, n = new_slide("🔗 Two and Three Words Together", "READ", "49–59 min",
                         "Shopping Cart", BERRY)
    one_task(slide, "Not a whole sentence yet. Just a few words joined up.", BERRY)
    phrase_grid(slide, PHRASES)
    hint(slide, "Sweep your finger under the whole phrase so it sounds like talking.",
         6.6)


def s38_read_steps():
    slide, n = new_slide("📕 How We Read Every Sentence", "METHOD", "59–68 min",
                         "Shopping Cart", SUN)
    one_task(slide, "Three steps, every single time. You are never first.", SUN)
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
           ["👩‍🏫  me", "👩‍🏫🙋  both", "🙋  you"][i], size=15, bold=True, color=color,
           align=PP_ALIGN.CENTER)
    hint(slide, "Only move to step 3 when step 2 sounded smooth. No rush.", 6.45)


def s39_sentences_a():
    slide, n = new_slide("📕 Sentence Shopping — First Two", "READ", "59–68 min",
                         "Shopping Cart", BERRY)
    one_task(slide, "A whole sentence now. I read it first.", BERRY)
    sentence_rows(slide, SENTENCES_A)
    hint(slide, "Point at each word as you read so her eyes follow left to right.",
         6.42)


def s40_sentences_b():
    slide, n = new_slide("📕 Sentence Shopping — Three More", "READ", "59–68 min",
                         "Shopping Cart", BERRY)
    one_task(slide, "Three more sentences. Read them like you are talking.", BERRY)
    sentence_rows(slide, SENTENCES_B, top_start=1.9, gap=1.5, height=1.34, size=30)
    hint(slide, "A stuck word? Give only the first sound and wait five seconds.",
         6.42)


def s41_puzzle_a():
    slide, n = new_slide("🧩 Game: Build the Sentence", "GAME", "59–68 min",
                         "Shopping Cart", SUN)
    one_task(slide, "These word cards are mixed up. Put them in order.", SUN)
    puzzle_rows(slide, PUZZLES_A, 1)
    hint(slide, "Find the capital letter. That word always goes first.", 6.42)


def s42_puzzle_b():
    slide, n = new_slide("🧩 Build the Sentence — Rounds 3 and 4", "GAME",
                         "59–68 min", "Shopping Cart", SUN)
    one_task(slide, "Two more. Read it out loud after you fix the order.", SUN)
    puzzle_rows(slide, PUZZLES_B, 3)
    hint(slide, "Read her version back exactly as written. Silly is a clue.", 6.42)


def s43_silly():
    slide, n = new_slide("🎯 Game: Correct or Silly?", "GAME", "59–68 min",
                         "Shopping Cart", MINT)
    one_task(slide, "Look at the picture. Which sentence is TRUE?", MINT)
    for i, (emoji, options) in enumerate(SILLY):
        top = Inches(1.95 + i * 1.52)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), L_MINT)
        tb(slide, Inches(0.85), top + Inches(0.3), Inches(1.2), Inches(0.78), emoji,
           size=32, align=PP_ALIGN.CENTER)
        for j, opt in enumerate(options):
            left = Inches(2.4 + j * 5.2)
            add_round(slide, left, top + Inches(0.3), Inches(4.9), Inches(0.74),
                      WHITE)
            tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(4.3),
               Inches(0.52), opt, size=20, bold=True, color=INK)
    hint(slide, "Let her laugh at the silly one. Then ask why it cannot be true.",
         6.45)


def s44_own_sentence():
    slide, n = new_slide("🗣️ Now You Make the Sentence", "SPEAKING", "59–68 min",
                         "Shopping Cart", LEAF)
    one_task(slide, "Pick a picture. Say a sentence about it.", LEAF)
    for i, (emoji, word) in enumerate(OWN_SENTENCE):
        left = Inches(0.6 + i * 3.15)
        add_round(slide, left, Inches(1.95), Inches(2.95), Inches(2.6), L_LEAF)
        tb(slide, left, Inches(2.2), Inches(2.95), Inches(0.95), emoji, size=40,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.45), Inches(3.4), Inches(2.05),
                  Inches(0.68), WHITE)
        tb(slide, left + Inches(0.45), Inches(3.52), Inches(2.05), Inches(0.48),
           word, size=22, bold=True, color=LEAF, align=PP_ALIGN.CENTER)
    for i, starter in enumerate(["I see a ______.", "I like the ______."]):
        left = Inches(0.6 + i * 6.35)
        add_round(slide, left, Inches(4.85), Inches(6.05), Inches(1.15), L_SUN)
        tb(slide, left + Inches(0.4), Inches(5.12), Inches(5.3), Inches(0.62),
           starter, size=26, bold=True, color=INK)
    hint(slide, "Two sentences is plenty. Speaking warms her up for the story.",
         6.45)


def s45_story_intro():
    slide, n = new_slide("📖 Story Time — Mia Goes Shopping", "STORY", "68–80 min",
                         "Checkout", TOMATO)
    one_task(slide, "Six words to watch for. You already know all six.", TOMATO)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4), L_TOMATO)
    tb(slide, Inches(0.5), Inches(2.45), Inches(5.4), Inches(1.8), "📖", size=90,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.4), Inches(5.4), Inches(0.7),
       "MIA GOES SHOPPING", size=25, bold=True, color=TOMATO, align=PP_ALIGN.CENTER,
       font="Georgia")
    tb(slide, Inches(0.8), Inches(5.2), Inches(4.8), Inches(0.6),
       "four short parts  ·  we read it twice", size=14, color=SOFT,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(6.25), Inches(1.9), Inches(6.6), Inches(0.5),
       "WATCH FOR THESE WORDS", size=15, bold=True, color=TOMATO)
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
           Inches(0.4), word, size=17, bold=True, color=TOMATO,
           align=PP_ALIGN.CENTER)
    hint(slide, "Read the six words together now. Meeting them early removes the "
                "fear.", 6.45)


def s46_story_1():
    story_slide(STORY[0], "68–80 min")


def s47_story_2():
    story_slide(STORY[1], "68–80 min")


def s48_story_3():
    story_slide(STORY[2], "68–80 min")


def s49_story_4():
    story_slide(STORY[3], "68–80 min")


def s50_read_again():
    slide, n = new_slide("🔁 Read It Again — Your Turn", "READ", "68–80 min",
                         "Checkout", LEAF)
    one_task(slide, "Six easy lines from the story. Read them on your own.", LEAF)
    for i, line in enumerate(EASY_LINES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.5)
        add_round(slide, left, top, Inches(5.95), Inches(1.3), L_LEAF)
        add_oval(slide, left + Inches(0.3), top + Inches(0.38), Inches(0.55),
                 Inches(0.55), LEAF)
        tb(slide, left + Inches(0.3), top + Inches(0.44), Inches(0.55),
           Inches(0.42), str(i + 1), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.05), top + Inches(0.38), Inches(4.6),
           Inches(0.62), line, size=22, bold=True, color=INK)
    hint(slide, "Second reading is always smoother. Say so out loud when it is.",
         6.55)


def s51_q_a():
    slide, n = new_slide("🔎 Shopping Story Detective — Q1 and Q2", "QUESTIONS",
                         "80–85 min", "Checkout", TOMATO)
    one_task(slide, "Answer from the story, not from your memory.", TOMATO)
    question_rows(slide, QUESTIONS_A, 1)
    hint(slide, "Turn back to the part named under the question. Let her find it.",
         6.42)


def s52_q_b():
    slide, n = new_slide("🔎 Story Detective — Q3 and Q4", "QUESTIONS", "80–85 min",
                         "Checkout", TOMATO)
    one_task(slide, "Two more. Point to the line that proves it.", TOMATO)
    question_rows(slide, QUESTIONS_B, 3)
    hint(slide, "Read the question out loud for her. The answer must be her work.",
         6.42)


def s53_q_c():
    slide, n = new_slide("🔎 Story Detective — Q5 and Q6", "QUESTIONS", "80–85 min",
                         "Checkout", TOMATO)
    one_task(slide, "These two need you to remember two parts at once.", TOMATO)
    question_rows(slide, QUESTIONS_C, 5)
    hint(slide, "Cover one wrong answer. Two choices is still real thinking.", 6.42)


def s54_q_d():
    slide, n = new_slide("🔎 Story Detective — Last Question", "QUESTIONS",
                         "80–85 min", "Checkout", TOMATO)
    one_task(slide, "The last question. Then we go to the checkout!", TOMATO)
    question_rows(slide, QUESTIONS_D, 7, top_start=1.95)
    add_round(slide, Inches(0.5), Inches(4.3), Inches(12.35), Inches(1.9), L_LEAF)
    tb(slide, Inches(0.85), Inches(4.5), Inches(11.6), Inches(0.42),
       "⭐ BONUS SHOPPER", size=13, bold=True, color=LEAF)
    tb(slide, Inches(0.85), Inches(4.98), Inches(11.6), Inches(0.6),
       "Tell me the story in your own words. Start with: \"Mia went to...\"",
       size=20, bold=True, color=INK)
    tb(slide, Inches(0.85), Inches(5.6), Inches(11.6), Inches(0.42),
       "Three sentences is a full retell.", size=13, color=SOFT)
    hint(slide, "Retelling shows she understood. It matters more than any answer.",
         6.45)


def s55_receipt():
    slide, n = new_slide("🧾 Checkout Game — Read the Receipt", "GAME", "85–90 min",
                         "Checkout", BERRY)
    one_task(slide, "Read each item on the receipt out loud.", BERRY)
    add_round(slide, Inches(0.5), Inches(1.8), Inches(2.7), Inches(4.5), L_BERRY)
    tb(slide, Inches(0.5), Inches(2.6), Inches(2.7), Inches(1.1), "💳", size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(3.9), Inches(2.3), Inches(1.4),
       "You are at the checkout counter.", size=16, bold=True, color=BERRY,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(10.15), Inches(1.8), Inches(2.7), Inches(4.5), L_GREY)
    tb(slide, Inches(10.4), Inches(2.0), Inches(2.2), Inches(0.4), "HOW TO PLAY",
       size=12, bold=True, color=BERRY)
    bullets(slide, Inches(10.4), Inches(2.5), Inches(2.2), Inches(3.5),
            ["Point at the item.", "Say the first sound.",
             "Blend the whole word.", "Read it out loud.", "Tick the box.",
             "Move to the next line."], size=12, sp=9)
    add_round(slide, Inches(3.4), Inches(1.8), Inches(6.5), Inches(4.5), WHITE)
    add_rect(slide, Inches(3.4), Inches(1.8), Inches(6.5), Inches(0.62), BERRY)
    tb(slide, Inches(3.4), Inches(1.92), Inches(6.5), Inches(0.42),
       "🧾  MIA'S SHOPPING LIST", size=16, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    for i, (emoji, item) in enumerate(RECEIPT):
        top = Inches(2.62 + i * 0.7)
        tb(slide, Inches(3.9), top, Inches(0.7), Inches(0.52), emoji, size=18,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(4.8), top + Inches(0.02), Inches(3.0), Inches(0.5), item,
           size=24, bold=True, color=INK, font="Arial Black")
        tb(slide, Inches(8.2), top + Inches(0.08), Inches(1.3), Inches(0.44), "☐",
           size=20, color=BERRY, align=PP_ALIGN.CENTER)
    tb(slide, Inches(3.4), Inches(6.0), Inches(6.5), Inches(0.36),
       "tick each box as you read it", size=12, color=SOFT, align=PP_ALIGN.CENTER)
    hint(slide, "Five items read on her own is a real win. Say it out loud.", 6.45)


def s56_checkout_sentences():
    slide, n = new_slide("💳 Checkout Counter — Say It to Pay", "READ", "85–90 min",
                         "Checkout", BERRY)
    one_task(slide, "Read the sentence to pay for the item.", BERRY)
    sentence_rows(slide, CHECKOUT_SENTENCES, top_start=1.9, gap=1.5, height=1.34,
                  size=30)
    hint(slide, "Each sentence uses only words she has already read today.", 6.42)


def s57_final_a():
    slide, n = new_slide("🏆 Final Challenge — Smart Shopper (1 of 2)", "CHALLENGE",
                         "85–90 min", "Checkout", SUN)
    one_task(slide, "Five last tasks. You have done every one before.", SUN)
    for i, (num, icon, kind, prompt, item, color, light) in enumerate(FINAL_A):
        top = Inches(1.9 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), light)
        add_oval(slide, Inches(0.8), top + Inches(0.42), Inches(0.52), Inches(0.52),
                 color)
        tb(slide, Inches(0.8), top + Inches(0.48), Inches(0.52), Inches(0.4), num,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.36), Inches(0.8), Inches(0.66), icon,
           size=24, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.5), top + Inches(0.42), Inches(1.7), Inches(0.5),
                  WHITE)
        tb(slide, Inches(2.5), top + Inches(0.5), Inches(1.7), Inches(0.36), kind,
           size=12, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, Inches(4.45), top + Inches(0.42), Inches(3.6), Inches(0.5),
           prompt, size=15, color=DARK)
        add_round(slide, Inches(8.3), top + Inches(0.3), Inches(4.3), Inches(0.74),
                  WHITE)
        tb(slide, Inches(8.3), top + Inches(0.4), Inches(4.3), Inches(0.56), item,
           size=24, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
    hint(slide, "Hints are still allowed here. Finishing matters, not scoring.",
         6.5)


def s58_final_b():
    slide, n = new_slide("🏆 Final Challenge — Smart Shopper (2 of 2)", "CHALLENGE",
                         "85–90 min", "Checkout", SUN)
    one_task(slide, "Two to go. Then you are a champion.", SUN)
    for i, (num, icon, kind, prompt, item, color, light) in enumerate(FINAL_B):
        top = Inches(1.95 + i * 1.72)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.5), light)
        add_oval(slide, Inches(0.8), top + Inches(0.5), Inches(0.52), Inches(0.52),
                 color)
        tb(slide, Inches(0.8), top + Inches(0.56), Inches(0.52), Inches(0.4), num,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.44), Inches(0.8), Inches(0.66), icon,
           size=24, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.5), top + Inches(0.5), Inches(1.7), Inches(0.5),
                  WHITE)
        tb(slide, Inches(2.5), top + Inches(0.58), Inches(1.7), Inches(0.36), kind,
           size=12, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, Inches(4.45), top + Inches(0.5), Inches(3.0), Inches(0.5), prompt,
           size=15, color=DARK)
        add_round(slide, Inches(7.7), top + Inches(0.36), Inches(4.9), Inches(0.8),
                  WHITE)
        tb(slide, Inches(7.7), top + Inches(0.48), Inches(4.9), Inches(0.58), item,
           size=22, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.45), Inches(12.35), Inches(0.8), L_LEAF)
    tb(slide, Inches(0.85), Inches(5.65), Inches(11.6), Inches(0.44),
       "✅  All five done? Ring the bell — the checkout is complete!", size=17,
       bold=True, color=LEAF, align=PP_ALIGN.CENTER)
    hint(slide, "Count out loud with her: one, two, three, four, five.", 6.5)


def s59_champion():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), SUN)
    for x, y in [(0.3, 4.05), (12.3, 4.05), (0.6, 5.5), (12.1, 5.5)]:
        tb(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.7), "⭐", size=24,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(0.95), Inches(12), Inches(1.1), "🏆", size=62,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(2.1), Inches(12), Inches(0.95),
       "SMART SHOPPER READING CHAMPION!", size=34, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(3.1), Inches(12), Inches(0.5),
       "Six aisles. Six points. One very good reader.", size=20,
       color=RGBColor(0xF0, 0xD8, 0x96), align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, place, _job, _when, color, _light) in enumerate(AISLES):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(3.75), Inches(1.65), Inches(1.3), color)
        tb(slide, left, Inches(3.92), Inches(1.65), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.08), Inches(4.55), Inches(1.5), Inches(0.4),
           place, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.4), Inches(5.25), Inches(6.5), Inches(1.1), SUN)
    tb(slide, Inches(3.4), Inches(5.5), Inches(6.5), Inches(0.62), "\"I CAN READ!\"",
       size=30, bold=True, color=INK, align=PP_ALIGN.CENTER, font="Georgia")
    footer(slide, n, "85–90 min", "Checkout")
    fade(slide)


def s60_can_read():
    slide, n = new_slide("✅ Today I Can...", "CHECKLIST", "85–90 min", "Checkout",
                         LEAF)
    one_task(slide, "Tick every one you did today. Read them with me.", LEAF)
    for i, (icon, skill, example) in enumerate(CAN_READ):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 0.95)
        add_round(slide, left, top, Inches(5.95), Inches(0.84), L_LEAF)
        tb(slide, left + Inches(0.25), top + Inches(0.2), Inches(0.5), Inches(0.5),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.9), top + Inches(0.18), Inches(2.5), Inches(0.5),
           skill, size=15, bold=True, color=INK)
        tb(slide, left + Inches(3.5), top + Inches(0.22), Inches(1.9), Inches(0.44),
           example, size=12, color=SOFT)
        tb(slide, left + Inches(5.4), top + Inches(0.18), Inches(0.4), Inches(0.5),
           "☐", size=17, color=LEAF)
    hint(slide, "Read the list aloud together. Hearing the whole list is the point.",
         6.6)


def s61_extra_games():
    slide, n = new_slide("🎲 Extra Games — If There Is Time", "TEACHER", "",
                         "Checkout", MINT)
    one_task(slide, "Seven quick games for a fast finisher or a slow moment.", MINT)
    for i, (icon, name, detail, color) in enumerate(EXTRA_GAMES):
        col, row = i % 4, i // 4
        left = Inches(0.5 + col * 3.15)
        top = Inches(1.95 + row * 2.4)
        add_round(slide, left, top, Inches(2.95), Inches(2.2), L_GREY)
        add_oval(slide, left + Inches(1.1), top + Inches(0.22), Inches(0.75),
                 Inches(0.75), color)
        tb(slide, left + Inches(1.1), top + Inches(0.33), Inches(0.75),
           Inches(0.55), icon, size=18, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.05), Inches(2.65),
           Inches(0.5), name, size=13, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), top + Inches(1.55), Inches(2.55),
           Inches(0.6), detail, size=11, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(9.95), Inches(4.35), Inches(2.9), Inches(1.9), L_SUN)
    tb(slide, Inches(10.2), Inches(4.5), Inches(2.4), Inches(0.36),
       "WHEN TO USE THESE", size=12, bold=True, color=SUN)
    bullets(slide, Inches(10.2), Inches(4.92), Inches(2.4), Inches(1.2),
            ["She finishes an activity early.", "Focus is dropping mid-block.",
             "A skill needs one more pass."], size=12)
    hint(slide, "Every game reuses today's words only. Nothing new appears.", 6.45)


def s62_support():
    slide, n = new_slide("🧰 Support System — For the Teacher", "TEACHER", "",
                         "Checkout", SKY)
    one_task(slide, "Levels, hint ladder and the words to say.", SKY)
    for i, (icon, name, who, action, color, light) in enumerate(SUPPORT_LEVELS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(2.25), light)
        tb(slide, left + Inches(0.25), Inches(2.02), Inches(0.5), Inches(0.5), icon,
           size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.9), Inches(2.02), Inches(2.9), Inches(0.5), name,
           size=15, bold=True, color=color)
        tb(slide, left + Inches(0.3), Inches(2.6), Inches(3.3), Inches(0.5), who,
           size=12, color=DARK)
        add_round(slide, left + Inches(0.3), Inches(3.15), Inches(3.3),
                  Inches(0.78), WHITE)
        tb(slide, left + Inches(0.45), Inches(3.26), Inches(3.0), Inches(0.58),
           action, size=11, color=INK)
    tb(slide, Inches(0.5), Inches(4.25), Inches(6.0), Inches(0.4),
       "HINT LADDER — never skip a step", size=13, bold=True, color=SKY)
    for i, (label, text, color) in enumerate(HINT_LADDER):
        top = Inches(4.66 + i * 0.46)
        add_round(slide, Inches(0.5), top, Inches(6.0), Inches(0.42), L_GREY)
        add_round(slide, Inches(0.62), top + Inches(0.05), Inches(1.0),
                  Inches(0.32), color)
        tb(slide, Inches(0.62), top + Inches(0.07), Inches(1.0), Inches(0.28),
           label, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.8), top + Inches(0.05), Inches(4.5), Inches(0.32), text,
           size=12, bold=True, color=INK)
    tb(slide, Inches(6.85), Inches(4.25), Inches(3.0), Inches(0.4),
       "HINT BANK", size=13, bold=True, color=SKY)
    bullets(slide, Inches(6.85), Inches(4.72), Inches(3.0), Inches(2.1), HINT_BANK,
            size=11)
    tb(slide, Inches(10.0), Inches(4.25), Inches(2.85), Inches(0.4),
       "WORDS TO USE", size=13, bold=True, color=LEAF)
    bullets(slide, Inches(10.0), Inches(4.72), Inches(2.85), Inches(2.1), PRAISE,
            size=11, color=LEAF)


def s63_assessment():
    slide, n = new_slide("📋 End-of-Class Assessment", "TEACHER", "", "Checkout",
                         CRUST)
    one_task(slide, "Tick one box per skill right after the lesson.", CRUST)
    heads = ["SKILL", "🟢 SECURE", "🟡 GROWING", "🔴 NEEDS WORK"]
    widths = [5.15, 2.4, 2.4, 2.4]
    lefts = [0.5]
    for w in widths[:-1]:
        lefts.append(lefts[-1] + w)
    add_rect(slide, Inches(0.5), Inches(1.82), Inches(12.35), Inches(0.42), CRUST)
    for left, w, head in zip(lefts, widths, heads):
        tb(slide, Inches(left), Inches(1.86), Inches(w), Inches(0.34), head,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = Inches(2.3 + i * 0.36)
        add_rect(slide, Inches(0.5), top, Inches(12.35), Inches(0.34),
                 WHITE if i % 2 == 0 else L_GREY)
        tb(slide, Inches(0.7), top + Inches(0.02), Inches(4.8), Inches(0.3), skill,
           size=11, bold=True, color=INK)
        for left, w in zip(lefts[1:], widths[1:]):
            tb(slide, Inches(left), top + Inches(0.01), Inches(w), Inches(0.3), "☐",
               size=13, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.02), Inches(12.35), Inches(0.9), L_CRUST)
    tb(slide, Inches(0.8), Inches(6.1), Inches(3.0), Inches(0.34), "NEXT LESSON",
       size=12, bold=True, color=CRUST)
    for i, item in enumerate(NEXT_LESSON):
        tb(slide, Inches(0.8 + i * 4.1), Inches(6.48), Inches(3.9), Inches(0.34),
           f"{i + 1}. {item}:  ____________", size=11, color=INK)


def s64_answer_key():
    slide, n = new_slide("🔑 Answer Key — Teacher Only", "TEACHER", "", "Checkout",
                         TOMATO)
    one_task(slide, "Quick reference for every game in this deck.", TOMATO)
    cols = [
        ("PUT IT IN THE CART (7–17)",
         ["1 milk  2 banana  3 carrot", "4 apple  5 juice  6 egg",
          "7 bread  8 drink", "Ending sounds: /g/ /k/ /g/ /p/"], LEAF),
        ("BASKETS & BUILDER (17–37)",
         ["Sort: A  A  E  ·  E  I  O  ·  O  A", "Missing letter: A  U  O  A",
          "Word puzzle: BAG BOX JAM CUP",
          "Builder: bag jam ham box cup sit can pan map"], SKY),
        ("BAKERY & CART (37–68)",
         ["Odd one out: BIG · HOP · BAG · HAT", "New words: BAT VAN MOP PIG",
          "Fill the cart: I see a the · my can is in",
          "Sentences: I see an apple. / The bag is big."], BERRY),
        ("STORY & CHECKOUT (68–90)",
         ["1 store  2 Dad  3 red  4 apple",
          "5 milk and a box  6 \"Did we get everything?\"", "7 \"Yes!\"",
          "Correct: banana · cold · big"], TOMATO),
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
    add_round(slide, Inches(0.5), Inches(6.5), Inches(12.35), Inches(0.42), L_LEAF)
    tb(slide, Inches(0.8), Inches(6.56), Inches(11.7), Inches(0.32),
       "Accept any sensible answer for the speaking tasks and the retell — "
       "fluency matters more than exact wording.", size=11, bold=True, color=LEAF)


BUILDERS = [
    s01_title, s02_helper, s03_map, s04_promises, s05_store_scene, s06_talk,
    s07_first_sounds, s08_cart_a, s09_cart_b, s10_cart_c, s11_ending,
    s12_produce_done, s13_baskets, s14_sort_a, s15_sort_b, s16_sort_bonus,
    s17_grocery_words, s18_dairy_done, s19_five_steps, s20_build_a, s21_build_b,
    s22_build_c, s23_missing, s24_scramble, s25_no_cards, s26_breakfast_done,
    s27_families, s28_family_a, s29_family_b, s30_new_words, s31_break_intro,
    s32_freeze, s33_sight_1, s34_sight_2, s35_fill_a, s36_fill_b, s37_phrases,
    s38_read_steps, s39_sentences_a, s40_sentences_b, s41_puzzle_a, s42_puzzle_b,
    s43_silly, s44_own_sentence, s45_story_intro, s46_story_1, s47_story_2,
    s48_story_3, s49_story_4, s50_read_again, s51_q_a, s52_q_b, s53_q_c, s54_q_d,
    s55_receipt, s56_checkout_sentences, s57_final_a, s58_final_b, s59_champion,
    s60_can_read, s61_extra_games, s62_support, s63_assessment, s64_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL}, built {len(prs.slides)}"

OUT = "Grade2_Amazing_Grocery_Store_90min.pptx"
prs.save(OUT)

story_words = sum(len(line.split()) for part in STORY for line in part[4])
with_notes = sum(1 for s in prs.slides
                 if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip())
print(f"saved {OUT}")
print(f"slides: {len(prs.slides)}")
print(f"story words: {story_words}")
print(f"slides with notes: {with_notes}")
