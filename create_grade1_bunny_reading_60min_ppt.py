"""Grade 1 reading lesson - 60 minutes, 30 slides.

"Bunny's Reading Adventure" - the child is a Reading Explorer who helps Bunny
finish reading challenges. Builds letter sounds -> blending -> CVC words ->
sight words -> sentences -> a short story, changing activity every 5-7 minutes.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

NAVY = RGBColor(0x1B, 0x35, 0x60)
GRASS = RGBColor(0x3F, 0x9E, 0x5C)
TEAL = RGBColor(0x00, 0x9C, 0x93)
SKY = RGBColor(0x33, 0xA6, 0xE8)
SUNNY = RGBColor(0xF6, 0xB2, 0x1B)
GOLD = RGBColor(0xFF, 0xC7, 0x33)
CARROT = RGBColor(0xF0, 0x7A, 0x38)
ROSE = RGBColor(0xE8, 0x5C, 0x86)
LILAC = RGBColor(0x8C, 0x62, 0xC6)
CREAM = RGBColor(0xFF, 0xFC, 0xF4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x24, 0x2C, 0x38)
SOFT = RGBColor(0x6B, 0x78, 0x8A)
L_SKY = RGBColor(0xE3, 0xF4, 0xFD)
L_TEAL = RGBColor(0xDD, 0xF5, 0xF3)
L_SUN = RGBColor(0xFF, 0xF3, 0xD8)
L_CARROT = RGBColor(0xFD, 0xEC, 0xE0)
L_ROSE = RGBColor(0xFD, 0xE9, 0xEF)
L_LILAC = RGBColor(0xF0, 0xEA, 0xFB)
L_GRASS = RGBColor(0xE6, 0xF4, 0xEA)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 30
_counter = {"n": 0}

# ------------------------------------------------------------------ content

POINT_AND_SAY = [("🐶", "Dog"), ("🐱", "Cat"), ("☀️", "Sun"), ("🍎", "Apple")]

ICEBREAKERS = [
    ("🙋", "What is your name?"),
    ("🐾", "What is your favorite animal?"),
    ("📚", "Do you like stories?"),
    ("🐰", "Can you make a bunny sound?"),
]

LETTERS = [
    ("m", "/m/", "🗺️", "map", ROSE, L_ROSE),
    ("s", "/s/", "☀️", "sun", SUNNY, L_SUN),
    ("t", "/t/", "🌳", "tree", GRASS, L_GRASS),
    ("p", "/p/", "🐷", "pig", ROSE, L_ROSE),
    ("c", "/k/", "🐱", "cat", SKY, L_SKY),
    ("b", "/b/", "🐻", "bear", CARROT, L_CARROT),
    ("d", "/d/", "🐶", "dog", TEAL, L_TEAL),
    ("r", "/r/", "🐰", "rabbit", LILAC, L_LILAC),
]

SOUND_DETECTIVE = [
    ("🐱", "CAT", ["/m/", "/k/", "/s/"], "/k/"),
    ("🗺️", "MAP", ["/m/", "/p/", "/b/"], "/m/"),
    ("☀️", "SUN", ["/s/", "/p/", "/d/"], "/s/"),
    ("🌳", "TREE", ["/t/", "/s/", "/m/"], "/t/"),
    ("🐷", "PIG", ["/b/", "/p/", "/t/"], "/p/"),
    ("🐶", "DOG", ["/d/", "/k/", "/m/"], "/d/"),
]

BLEND_WORDS = [
    ("MAT", ["M", "A", "T"], ["/m/", "/a/", "/t/"], "🟫", ROSE, L_ROSE),
    ("SUN", ["S", "U", "N"], ["/s/", "/u/", "/n/"], "☀️", SUNNY, L_SUN),
    ("PIG", ["P", "I", "G"], ["/p/", "/i/", "/g/"], "🐷", CARROT, L_CARROT),
    ("DOG", ["D", "O", "G"], ["/d/", "/o/", "/g/"], "🐶", TEAL, L_TEAL),
]

WORD_FAMILIES = [
    ("- at", ["cat", "bat", "mat", "sat"], "🐱", SKY, L_SKY),
    ("- og", ["dog", "log"], "🐶", TEAL, L_TEAL),
    ("- un", ["sun", "run"], "☀️", SUNNY, L_SUN),
    ("- ig", ["pig", "big"], "🐷", ROSE, L_ROSE),
    ("- en", ["hen", "pen"], "🐔", CARROT, L_CARROT),
]

FEED_BUNNY = [
    ("🐱", "CAT", ["CAT", "DOG", "SUN"]),
    ("☀️", "SUN", ["MAT", "SUN", "PIG"]),
    ("🐶", "DOG", ["DOG", "BIG", "HEN"]),
    ("🐷", "PIG", ["RUN", "PIG", "BAT"]),
    ("🐔", "HEN", ["LOG", "HEN", "PEN"]),
]

MATCH_PAIRS = [("🐱", "CAT"), ("☀️", "SUN"), ("🐷", "PIG"), ("🐶", "DOG")]
MATCH_SHUFFLED = ["DOG", "PIG", "CAT", "SUN"]

SIGHT_SETS = [
    ("Set 1", ["I", "am", "a"], TEAL, L_TEAL),
    ("Set 2", ["the", "is", "my"], ROSE, L_ROSE),
    ("Set 3", ["can", "see"], LILAC, L_LILAC),
]

CATCH_WORD = [
    ("I", ["cat", "I", "dog", "sun"]),
    ("am", ["an", "am", "at", "and"]),
    ("the", ["then", "the", "they", "this"]),
    ("can", ["can", "cat", "cap", "car"]),
    ("my", ["me", "my", "may", "man"]),
    ("see", ["sea", "see", "saw", "so"]),
]

BUNNY_SAYS = [
    "Bunny says touch your nose.",
    "Bunny says clap 3 times.",
    "Bunny says hop like a bunny.",
    "Bunny says point to the letter B.",
    "Bunny says make the /m/ sound.",
    "Bunny says read the word CAT.",
]
BUNNY_BOARD = ["B", "M", "CAT", "SUN"]

SENTENCES = [
    ("I am Sam.", "🧒", SKY, L_SKY),
    ("I see a cat.", "🐱", TEAL, L_TEAL),
    ("The dog is big.", "🐶", GRASS, L_GRASS),
    ("I can run.", "🏃", CARROT, L_CARROT),
    ("My cat is red.", "🐱", ROSE, L_ROSE),
    ("I see the sun.", "☀️", SUNNY, L_SUN),
]

SCENE_ITEMS = [("☀️", 0.55, 0.30), ("🌳", 4.45, 2.20), ("🐶", 0.85, 2.50),
               ("🔴", 2.25, 2.75), ("🐱", 3.45, 2.60)]

READ_AND_FIND = [
    ("The dog is big.", "🐶"),
    ("I see the sun.", "☀️"),
    ("The cat is on the mat.", "🐱"),
]

VOCAB = [
    ("🐰", "bunny", "bun - ny", "Bunny can hop."),
    ("🌳", "park", "p - ar - k", "I see the park."),
    ("🐶", "dog", "d - o - g", "The dog is big."),
    ("🔴", "ball", "b - all", "I see a ball."),
    ("🎨", "red", "r - e - d", "My ball is red."),
    ("🏃", "run", "r - u - n", "I can run."),
    ("🤸", "play", "p - l - ay", "We can play."),
    ("😀", "happy", "hap - py", "Bunny is happy."),
]

STORY = [
    ("Part 1", "🐰", ["Bunny goes to the park.", "The sun is big.", "Bunny can hop.",
                      "Bunny sees a dog.", "The dog is big."], GRASS, L_GRASS),
    ("Part 2", "🔴", ["The dog has a red ball.", "Bunny runs to the dog.",
                      "\"Can I play?\" says Bunny.", "\"Yes!\" says the dog.",
                      "The dog can run."], SKY, L_SKY),
    ("Part 3", "😀", ["They play with the ball.", "Bunny can run and hop.",
                      "It is a fun day!", "Bunny is happy.",
                      "The dog is happy too."], SUNNY, L_SUN),
]

COMPREHENSION = [
    ("Where does Bunny go?", [("🏫", "School"), ("🌳", "Park"), ("🏪", "Store")], "Park"),
    ("What does the dog have?", [("🔴", "A ball"), ("📚", "A book"), ("🎩", "A hat")], "A ball"),
    ("How does Bunny feel?", [("😀", "Happy"), ("😢", "Sad"), ("😠", "Angry")], "Happy"),
]

FINAL_WORDS = ["CAT", "SUN", "DOG"]
FINAL_SENTENCE = "I see a dog."

RECAP = [
    ("🔤", "Say letter sounds"), ("🎵", "Blend sounds"), ("🧩", "Read words"),
    ("📖", "Read sentences"), ("📚", "Read a story"), ("💬", "Answer questions"),
]

RUBRIC_SKILLS = ["Letter sounds", "Blending", "CVC words", "Sight words",
                 "Sentence reading", "Comprehension", "Confidence"]

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


def bullets(slide, l, t, w, h, items, size=12, color=DARK, sp=6):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(sp)
        r = p.add_run()
        r.text = "•  " + item
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


def footer(slide, n, timing="", step=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE4, 0xEA, 0xF0))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL), Inches(0.08), GOLD)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), NAVY)
    msg = "🐰 Bunny's Reading Adventure  |  Grade 1  |  60 min"
    if step:
        msg += f"  |  {step}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=10, color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def chip(slide, text, color):
    add_round(slide, Inches(0.4), Inches(0.26), Inches(3.2), Inches(0.42), color)
    tb(slide, Inches(0.4), Inches(0.31), Inches(3.2), Inches(0.34), text, size=12, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def time_chip(slide, text):
    add_round(slide, Inches(10.3), Inches(0.26), Inches(2.65), Inches(0.42), CARROT)
    tb(slide, Inches(10.3), Inches(0.31), Inches(2.65), Inches(0.34), text, size=12, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def star_badge(slide, left=Inches(10.3), top=Inches(6.42), text="⭐ +1 Star"):
    add_round(slide, left, top, Inches(2.65), Inches(0.5), GOLD)
    tb(slide, left, top + Inches(0.08), Inches(2.65), Inches(0.38), text, size=13, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER)


def new_slide(title, tag, timing, step, accent=SKY, bg=CREAM):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, accent)
    chip(slide, tag, accent)
    if timing:
        time_chip(slide, timing)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title, size=28, bold=True,
       color=NAVY, font="Georgia")
    footer(slide, n, timing, step)
    fade(slide)
    return slide, n


def notes(slide, say, ask, expected, help_step, praise, timing, support):
    tf = slide.notes_slide.notes_text_frame
    tf.text = (
        f"⏱ TIME: {timing}     |     READING MODE: {support}\n\n"
        f"WHAT TO SAY: {say}\n\n"
        f"QUESTION TO ASK: {ask}\n\n"
        f"EXPECTED RESPONSE: {expected}\n\n"
        f"IF THE CHILD STRUGGLES: {help_step}\n\n"
        f"PRAISE: {praise}\n\n"
        "NEVER say \"Wrong.\" Say \"Let's try it together,\" break the word into sounds, "
        "blend them, then ask the child to repeat it successfully."
    )


def ido_wedo_youdo(slide, top, accent, labels=("I read it", "We read it", "YOU read it!")):
    steps = [("1️⃣", labels[0], "Listen and follow", L_SUN, NAVY),
             ("2️⃣", labels[1], "Say it with me", L_TEAL, NAVY),
             ("3️⃣", labels[2], "Your turn ⭐", accent, WHITE)]
    for i, (num, title, sub, bg, fg) in enumerate(steps):
        left = Inches(0.55 + i * 4.15)
        add_round(slide, left, top, Inches(3.9), Inches(1.5), bg)
        tb(slide, left + Inches(0.15), top + Inches(0.14), Inches(3.6), Inches(0.45), num,
           size=22, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.62), Inches(3.6), Inches(0.45), title,
           size=18, bold=True, color=fg, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.06), Inches(3.6), Inches(0.36), sub,
           size=13, color=fg, align=PP_ALIGN.CENTER)


# ------------------------------------------------------------------ slides


def s01_welcome():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    for x, y, c in [(0.65, 0.8, SKY), (11.95, 0.85, ROSE), (0.85, 5.8, GRASS), (11.9, 5.75, SUNNY)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.8), c)
    tb(slide, Inches(0.7), Inches(1.2), Inches(12), Inches(1.1), "🐰", size=64,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(2.3), Inches(12), Inches(0.85),
       "Bunny's Reading Adventure", size=42, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(3.25), Inches(12), Inches(0.5),
       "Grade 1  •  60 Minutes  •  You are the Reading Explorer!",
       size=18, color=L_SKY, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.4), Inches(4.0), Inches(6.5), Inches(1.25), GRASS)
    tb(slide, Inches(3.6), Inches(4.25), Inches(6.1), Inches(0.8),
       "Bunny needs a reading buddy today.\nAre you ready? ⭐",
       size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(5.6), Inches(12), Inches(0.5),
       "🔤 Sounds  →  🧩 Words  →  📖 Sentences  →  📚 A Story",
       size=18, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    footer(slide, n, "0–5 min", "Welcome")
    fade(slide)
    notes(slide,
          "Hi! I am so happy to see you today. Look who came to class — it's Bunny! "
          "Bunny loves to read, and Bunny needs a reading buddy. That's going to be you.",
          "Are you ready to be a Reading Explorer with Bunny?",
          "A smile, a nod, or 'yes'. Any response is a win.",
          "If the child is shy, wave to Bunny yourself and say 'You can just wave hello.'",
          "I am so glad you are here. Bunny is excited too!",
          "0–5 min", "No reading required")


def s02_meet_bunny():
    slide, n = new_slide("🐰 Meet Bunny and Friends", "MEET", "0–5 min", "Meet Bunny", LILAC)
    friends = [("🐰", "Bunny", "Loves to hop", LILAC, L_LILAC),
               ("🐻", "Bear", "Loves stories", CARROT, L_CARROT),
               ("🐶", "Puppy", "Loves to run", TEAL, L_TEAL),
               ("🐱", "Cat", "Loves to nap", SKY, L_SKY)]
    for i, (emoji, name, blurb, color, light) in enumerate(friends):
        left = Inches(0.5 + i * 3.16)
        add_round(slide, left, Inches(1.6), Inches(2.96), Inches(3.3), light)
        tb(slide, left, Inches(1.95), Inches(2.96), Inches(1.2), emoji, size=54,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(3.25), Inches(2.96), Inches(0.6), name, size=26, bold=True,
           color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.15), Inches(3.95), Inches(2.66), Inches(0.5), blurb,
           size=14, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.15), Inches(12.35), Inches(1.2), L_SUN)
    tb(slide, Inches(0.85), Inches(5.4), Inches(11.6), Inches(0.7),
       "These four friends will help you all class long. Say hello to each one! 👋",
       size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    notes(slide,
          "Here are our four friends. This is Bunny, this is Bear, this is Puppy, and "
          "this is Cat. They will help us read today. Let's say hi to each one.",
          "Which friend is your favorite? Can you make a bunny sound?",
          "The child names a favorite animal and makes a sound or a hop motion.",
          "Point to one and say 'This is Bunny. Can you say Bunny?' Keep it playful.",
          "Great! Bunny likes you already.",
          "0–5 min", "No reading required")


def s03_point_and_say():
    slide, n = new_slide("👀 Point and Say!", "WARM-UP", "0–5 min", "Point & say", TEAL)
    tb(slide, Inches(0.45), Inches(1.4), Inches(11.5), Inches(0.42),
       "Point to each picture and tell me what it is.", size=17, bold=True, color=CARROT)
    for i, (emoji, label) in enumerate(POINT_AND_SAY):
        left = Inches(0.55 + i * 3.15)
        add_round(slide, left, Inches(1.95), Inches(2.95), Inches(2.35), WHITE)
        tb(slide, left, Inches(2.2), Inches(2.95), Inches(1.1), emoji, size=52,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(3.45), Inches(2.95), Inches(0.6), label, size=26, bold=True,
           color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
    for i, (icon, question) in enumerate(ICEBREAKERS):
        col, row = i % 2, i // 2
        left = Inches(0.55 + col * 6.2)
        top = Inches(4.55 + row * 0.78)
        add_round(slide, left, top, Inches(5.9), Inches(0.65), L_TEAL)
        tb(slide, left + Inches(0.22), top + Inches(0.13), Inches(0.5), Inches(0.42), icon,
           size=16)
        tb(slide, left + Inches(0.85), top + Inches(0.14), Inches(4.9), Inches(0.42), question,
           size=16, bold=True, color=NAVY)
    notes(slide,
          "Let's warm up our talking voices. Point to this one — what is it? Yes, a dog! "
          "Now tell me about you.",
          "What do you see in the picture? What is your name? Do you like stories?",
          "Dog, cat, sun, apple. Then short spoken answers to each question.",
          "Name the picture yourself first, then ask the child to repeat it back to you.",
          "You know all of these! Great talking.",
          "0–5 min", "Speaking only, no reading")


def s04_letter_sounds():
    slide, n = new_slide("🔤 Letter Sound Warm-Up", "PHONICS", "5–12 min", "Letter sounds", ROSE)
    tb(slide, Inches(0.45), Inches(1.38), Inches(11.5), Inches(0.42),
       "I say the sound → You say the sound. Then I point → You tell me!",
       size=16, bold=True, color=CARROT)
    for i, (letter, sound, emoji, word, color, light) in enumerate(LETTERS):
        col, row = i % 4, i // 4
        left = Inches(0.45 + col * 3.18)
        top = Inches(1.92 + row * 2.42)
        add_round(slide, left, top, Inches(3.0), Inches(2.22), light)
        tb(slide, left + Inches(0.12), top + Inches(0.1), Inches(1.3), Inches(0.95), letter,
           size=46, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(1.55), top + Inches(0.32), Inches(1.3), Inches(0.5), sound,
           size=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), top + Inches(1.16), Inches(2.7), Inches(0.88), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(1.31), Inches(0.7), Inches(0.6), emoji,
           size=22, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(1.35), Inches(1.65), Inches(0.5), word,
           size=17, bold=True, color=NAVY)
    notes(slide,
          "Let's wake up our sounds. The letter m says /m/... like map. Your turn: /m/. "
          "Now s says /s/... like sun.",
          "What sound does this letter make?",
          "The child repeats each sound, then names sounds alone when you point.",
          "Say the sound with the picture and stretch it out: 'mmmmap — /m/ /m/ /m/.' "
          "Use only four letters if attention drops.",
          "Your sounds are getting stronger already!",
          "5–12 min", "Listen and repeat, then identify")


def _detective_slide(items, title, timing, start_index):
    slide, n = new_slide(title, "GAME", timing, "Sound detective", LILAC)
    tb(slide, Inches(0.45), Inches(1.38), Inches(11.5), Inches(0.42),
       "What sound do you hear at the BEGINNING?", size=17, bold=True, color=CARROT)
    for i, (emoji, word, options, _answer) in enumerate(items):
        top = Inches(1.95 + i * 1.68)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.48), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.44), Inches(0.6), Inches(0.6), LILAC)
        tb(slide, Inches(0.75), top + Inches(0.52), Inches(0.6), Inches(0.45),
           str(start_index + i), size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.6), top + Inches(0.4), Inches(1.0), Inches(0.7), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.8), top + Inches(0.42), Inches(2.4), Inches(0.6), word, size=30,
           bold=True, color=NAVY, font="Arial Black")
        for j, opt in enumerate(options):
            left = Inches(5.6 + j * 2.45)
            add_round(slide, left, top + Inches(0.34), Inches(2.2), Inches(0.8), L_LILAC)
            tb(slide, left, top + Inches(0.47), Inches(2.2), Inches(0.55),
               f"{chr(65 + j)}.  {opt}", size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    return slide, n


def s05_detective_a():
    slide, n = _detective_slide(SOUND_DETECTIVE[:3], "🎯 Game: Sound Detective", "5–12 min", 1)
    notes(slide,
          "Detective time! Look at the picture. Cccc-at. What sound do you hear at the very "
          "beginning of cat? Is it A, B, or C?",
          "What sound does this word start with?",
          "1. /k/   2. /m/   3. /s/",
          "Stretch the first sound three times: 'ccccc-at.' Then read the three choices aloud "
          "and let the child point.",
          "Great job, Reading Detective!",
          "5–12 min", "Listening, child points to answer")


def s06_detective_b():
    slide, n = _detective_slide(SOUND_DETECTIVE[3:], "🎯 Sound Detective — Round 2", "5–12 min", 4)
    notes(slide,
          "Three more mysteries! Tttt-ree. Which sound is at the beginning?",
          "What sound do you hear first?",
          "4. /t/   5. /p/   6. /d/",
          "Cover one wrong choice with your finger so there are only two options left.",
          "You found it! Nothing gets past you.",
          "5–12 min", "Listening, child points to answer")


def s07_blending_intro():
    slide, n = new_slide("🎵 Sounds Hold Hands: Blending", "PHONICS", "12–20 min",
                         "Blending", GRASS)
    add_round(slide, Inches(0.5), Inches(1.5), Inches(12.35), Inches(2.5), WHITE)
    parts = [("C", "/k/", SKY), ("A", "/a/", ROSE), ("T", "/t/", GRASS)]
    for i, (letter, sound, color) in enumerate(parts):
        left = Inches(1.35 + i * 2.5)
        add_round(slide, left, Inches(1.85), Inches(1.9), Inches(1.8), L_SKY)
        tb(slide, left, Inches(1.98), Inches(1.9), Inches(0.95), letter, size=48, bold=True,
           color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left, Inches(2.95), Inches(1.9), Inches(0.5), sound, size=20, bold=True,
           color=NAVY, align=PP_ALIGN.CENTER)
        if i < 2:
            tb(slide, left + Inches(1.93), Inches(2.3), Inches(0.55), Inches(0.65), "→",
               size=26, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(8.85), Inches(2.3), Inches(0.8), Inches(0.65), "➡", size=26, bold=True,
       color=CARROT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(9.7), Inches(1.85), Inches(2.75), Inches(1.8), L_SUN)
    tb(slide, Inches(9.7), Inches(1.98), Inches(2.75), Inches(0.95), "CAT", size=42, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(9.7), Inches(2.95), Inches(2.75), Inches(0.55), "🐱", size=24,
       align=PP_ALIGN.CENTER)
    ido_wedo_youdo(slide, Inches(4.3), GRASS, ("I blend it", "We blend it", "YOU blend it!"))
    tb(slide, Inches(0.55), Inches(6.05), Inches(9.5), Inches(0.42),
       "🐢 Slow:  /k/ ... /a/ ... /t/        🐇 Fast:  CAT!", size=17, bold=True, color=CARROT)
    notes(slide,
          "Watch this. When sounds hold hands, they make a word. /k/ ... /a/ ... /t/. "
          "Now faster: /k/-/a/-/t/... CAT!",
          "Can you blend them with me? What word did we make?",
          "The child says CAT, first with you and then alone.",
          "Blend just the first two sounds: /k/ /a/ = 'ca'. Then add /t/. Sweep your finger "
          "under the letters so the child sees the left-to-right movement.",
          "You made the sounds turn into a word. That is reading!",
          "12–20 min", "I DO → WE DO → YOU DO")


def s08_build_the_word():
    slide, n = new_slide("🔤 Game: Build the Word", "GAME", "12–20 min", "Build the word", SKY)
    tb(slide, Inches(0.45), Inches(1.38), Inches(11.5), Inches(0.42),
       "Say each sound, then push them together!", size=16, bold=True, color=CARROT)
    for i, (word, letters, sounds, emoji, color, light) in enumerate(BLEND_WORDS):
        top = Inches(1.9 + i * 1.28)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.12), WHITE)
        for j, (letter, sound) in enumerate(zip(letters, sounds)):
            left = Inches(0.8 + j * 1.72)
            add_round(slide, left, top + Inches(0.18), Inches(1.45), Inches(0.76), light)
            tb(slide, left, top + Inches(0.22), Inches(1.45), Inches(0.42), letter, size=22,
               bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
            tb(slide, left, top + Inches(0.62), Inches(1.45), Inches(0.3), sound, size=11,
               bold=True, color=SOFT, align=PP_ALIGN.CENTER)
            if j < len(letters) - 1:
                tb(slide, left + Inches(1.46), top + Inches(0.3), Inches(0.26), Inches(0.45),
                   "→", size=16, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.1), top + Inches(0.28), Inches(0.8), Inches(0.5), "➡", size=20,
           bold=True, color=CARROT, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(7.1), top + Inches(0.18), Inches(2.7), Inches(0.76), L_SUN)
        tb(slide, Inches(7.1), top + Inches(0.26), Inches(2.7), Inches(0.58), word, size=26,
           bold=True, color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(10.1), top + Inches(0.24), Inches(0.9), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        add_round(slide, Inches(11.25), top + Inches(0.28), Inches(1.4), Inches(0.55), L_TEAL)
        tb(slide, Inches(11.25), top + Inches(0.36), Inches(1.4), Inches(0.4), "⭐", size=14,
           align=PP_ALIGN.CENTER)
    notes(slide,
          "Let's build words! Say each sound with me: /m/ /a/ /t/. Now blend them fast... MAT!",
          "What sound is first? What word did we build?",
          "MAT  •  SUN  •  PIG  •  DOG",
          "Cover the last letter with your hand, blend the first two, then reveal the last "
          "sound and blend all three.",
          "You didn't give up — and you built it!",
          "12–20 min", "WE DO first, then YOU DO")


def s09_word_families():
    slide, n = new_slide("🧩 Word Families — Words That Rhyme", "READING", "20–28 min",
                         "Word families", TEAL)
    tb(slide, Inches(0.45), Inches(1.38), Inches(11.5), Inches(0.42),
       "Change the first sound and you get a brand new word!", size=16, bold=True, color=CARROT)
    for i, (family, words, emoji, color, light) in enumerate(WORD_FAMILIES):
        left = Inches(0.5 + i * 2.5)
        add_round(slide, left, Inches(1.95), Inches(2.32), Inches(4.35), light)
        tb(slide, left, Inches(2.15), Inches(2.32), Inches(0.65), emoji, size=26,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.4), Inches(2.85), Inches(1.52), Inches(0.55), color)
        tb(slide, left + Inches(0.4), Inches(2.94), Inches(1.52), Inches(0.4), family, size=17,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        for j, word in enumerate(words):
            top = Inches(3.55 + j * 0.68)
            add_round(slide, left + Inches(0.2), top, Inches(1.92), Inches(0.56), WHITE)
            tb(slide, left + Inches(0.2), top + Inches(0.07), Inches(1.92), Inches(0.42), word,
               size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    notes(slide,
          "Look at this family. Cat, bat, mat, sat — they all end the same way. Only the "
          "first sound changes. Listen: /k/-at, /b/-at, /m/-at.",
          "Can you read the next one in the family?",
          "cat bat mat sat  •  dog log  •  sun run  •  pig big  •  hen pen",
          "Read the first word of the family yourself, then cover the first letter of the "
          "next word and reveal it slowly.",
          "Nice reading! You spotted the pattern.",
          "20–28 min", "WE DO, then child tries one family alone")


def s10_feed_bunny():
    slide, n = new_slide("🐰 Game: Feed the Bunny", "GAME", "20–28 min", "Feed the bunny", CARROT)
    tb(slide, Inches(0.45), Inches(1.32), Inches(11.5), Inches(0.4),
       "Bunny is hungry! Find the right word and feed it to Bunny.  ⭐ 5 stars to win!",
       size=16, bold=True, color=CARROT)
    for i, (emoji, target, options) in enumerate(FEED_BUNNY):
        top = Inches(1.82 + i * 1.05)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.92), WHITE)
        tb(slide, Inches(0.68), top + Inches(0.18), Inches(0.7), Inches(0.55), "🐰", size=20,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.2), Inches(0.7), Inches(0.52), emoji, size=18,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.35), top + Inches(0.24), Inches(2.6), Inches(0.45),
           f"Find  {target}", size=17, bold=True, color=CARROT)
        for j, opt in enumerate(options):
            left = Inches(5.3 + j * 2.55)
            add_round(slide, left, top + Inches(0.14), Inches(2.3), Inches(0.64), L_CARROT)
            tb(slide, left, top + Inches(0.2), Inches(2.3), Inches(0.5), opt, size=21,
               bold=True, color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
    notes(slide,
          "Bunny is hungry and only eats the right word! Look at the picture, then find "
          "the word that matches. Point to it and read it out loud.",
          "Bunny wants the word CAT — can you find it?",
          "CAT  •  SUN  •  DOG  •  PIG  •  HEN",
          "Ask what sound the word starts with, then look for that letter in the choices. "
          "Cover one wrong word so only two remain.",
          "You found it! Bunny says thank you. Munch munch!",
          "20–28 min", "Child reads the word aloud after pointing")


def s11_match_word():
    slide, n = new_slide("🧩 Game: Match the Word to the Picture", "GAME", "20–28 min",
                         "Match it", ROSE)
    tb(slide, Inches(0.45), Inches(1.38), Inches(11.5), Inches(0.42),
       "The words got mixed up! Match each picture to its word.",
       size=16, bold=True, color=CARROT)
    for i, (emoji, _word) in enumerate(MATCH_PAIRS):
        left = Inches(0.75 + i * 3.05)
        add_round(slide, left, Inches(1.95), Inches(2.8), Inches(1.95), L_ROSE)
        tb(slide, left, Inches(2.3), Inches(2.8), Inches(1.1), emoji, size=44,
           align=PP_ALIGN.CENTER)
    for i, word in enumerate(MATCH_SHUFFLED):
        left = Inches(0.75 + i * 3.05)
        add_round(slide, left, Inches(4.35), Inches(2.8), Inches(1.15), WHITE)
        tb(slide, left, Inches(4.6), Inches(2.8), Inches(0.65), word, size=28, bold=True,
           color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.75), Inches(5.75), Inches(9.5), Inches(0.45),
       "Draw a line in the air from each picture to its word! ✏️", size=16, bold=True,
       color=TEAL)
    notes(slide,
          "The pictures are on top and the words are on the bottom, but they are all mixed "
          "up. Point to the cat, then find the word CAT.",
          "Which word goes with this picture?",
          "🐱 = CAT (3rd word)  •  ☀️ = SUN (4th)  •  🐷 = PIG (2nd)  •  🐶 = DOG (1st)",
          "Ask for the first sound of the picture word, then hunt for that letter in the "
          "word cards.",
          "Four out of four! Excellent matching.",
          "20–28 min", "Child reads each word after matching")


def s12_sight_words():
    slide, n = new_slide("👀 Sight Words — Just Know Them!", "SIGHT WORDS", "28–35 min",
                         "Sight words", LILAC)
    tb(slide, Inches(0.45), Inches(1.38), Inches(11.5), Inches(0.42),
       "These words we do NOT sound out. We just know them by sight!",
       size=16, bold=True, color=CARROT)
    top = Inches(1.9)
    for label, words, color, light in SIGHT_SETS:
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.5), light)
        add_round(slide, Inches(0.72), top + Inches(0.45), Inches(1.5), Inches(0.55), color)
        tb(slide, Inches(0.72), top + Inches(0.54), Inches(1.5), Inches(0.4), label, size=13,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        for j, word in enumerate(words):
            left = Inches(2.6 + j * 2.55)
            add_round(slide, left, top + Inches(0.3), Inches(2.3), Inches(0.88), WHITE)
            tb(slide, left, top + Inches(0.45), Inches(2.3), Inches(0.58), word, size=30,
               bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        top = top + Inches(1.62)
    notes(slide,
          "These are special words. We do not sound them out — we just learn them by sight, "
          "like knowing a friend's face. This one says THE. Your turn: THE.",
          "What does this word say?",
          "I, am, a  →  the, is, my  →  can, see",
          "Cover Sets 2 and 3 with a sheet of paper. Teach Set 1 only, three times each, and "
          "reveal the next set only when Set 1 is solid. Two words truly known beats eight "
          "half-known.",
          "You just KNEW that word without sounding it out!",
          "28–35 min", "Look and say, teacher models first")


def _catch_slide(items, title, timing, start_index):
    slide, n = new_slide(title, "GAME", timing, "Catch the word", TEAL)
    tb(slide, Inches(0.45), Inches(1.38), Inches(11.5), Inches(0.42),
       "Careful — the other words look almost the same!", size=16, bold=True, color=CARROT)
    for i, (target, options) in enumerate(items):
        top = Inches(1.95 + i * 1.68)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.48), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.44), Inches(0.6), Inches(0.6), TEAL)
        tb(slide, Inches(0.75), top + Inches(0.52), Inches(0.6), Inches(0.45),
           str(start_index + i), size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.6), top + Inches(0.45), Inches(2.6), Inches(0.6),
           f"Find  {target}", size=22, bold=True, color=TEAL)
        for j, opt in enumerate(options):
            left = Inches(4.5 + j * 2.15)
            add_round(slide, left, top + Inches(0.34), Inches(1.95), Inches(0.8), L_TEAL)
            tb(slide, left, top + Inches(0.48), Inches(1.95), Inches(0.55), opt, size=22,
               bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    return slide, n


def s13_catch_a():
    slide, n = _catch_slide(CATCH_WORD[:3], "👀 Game: Catch the Sight Word", "28–35 min", 1)
    notes(slide,
          "Can you catch the word I? Look at all four words and point to the one that says I.",
          "Can you find the word AM? Where is THE?",
          "1. I (2nd)   2. am (2nd)   3. the (2nd)",
          "Say the word again slowly, then cover two wrong choices with your fingers so "
          "only two remain.",
          "You caught it! Nice reading.",
          "28–35 min", "Child points, then says the word")


def s14_catch_b():
    slide, n = _catch_slide(CATCH_WORD[3:], "👀 Catch the Sight Word — Round 2", "28–35 min", 4)
    notes(slide,
          "These are tricky ones — they all start the same way! Look at the whole word, "
          "not just the first letter. Can you find CAN?",
          "Where is MY? Can you find SEE?",
          "4. can (1st)   5. my (2nd)   6. see (2nd)",
          "Read all four choices aloud yourself, then ask again. Point out the last letter "
          "as the clue that makes them different.",
          "Those tricky words did not fool you!",
          "28–35 min", "Child points, then says the word")


def s15_brain_break():
    slide, n = new_slide("🐰 Brain Break: Bunny Says!", "BREAK", "35–38 min",
                         "Bunny says", GOLD)
    tb(slide, Inches(0.45), Inches(1.38), Inches(11.5), Inches(0.42),
       "Do it only if Bunny says! Stay by your desk. 🎵", size=16, bold=True, color=CARROT)
    for i, cmd in enumerate(BUNNY_SAYS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 4.6)
        top = Inches(1.95 + row * 1.15)
        add_round(slide, left, top, Inches(4.35), Inches(0.95), WHITE)
        add_oval(slide, left + Inches(0.2), top + Inches(0.24), Inches(0.48), Inches(0.48), GOLD)
        tb(slide, left + Inches(0.2), top + Inches(0.3), Inches(0.48), Inches(0.38), str(i + 1),
           size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.85), top + Inches(0.28), Inches(3.35), Inches(0.45), cmd,
           size=14, bold=True, color=NAVY)
    add_round(slide, Inches(9.75), Inches(1.95), Inches(3.1), Inches(3.55), L_SUN)
    tb(slide, Inches(9.95), Inches(2.12), Inches(2.7), Inches(0.4), "Point to these:", size=15,
       bold=True, color=NAVY)
    for i, item in enumerate(BUNNY_BOARD):
        top = Inches(2.62 + i * 0.7)
        add_round(slide, Inches(10.0), top, Inches(2.6), Inches(0.58), WHITE)
        tb(slide, Inches(10.0), top + Inches(0.08), Inches(2.6), Inches(0.42), item, size=20,
           bold=True, color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.5), Inches(5.65), Inches(9.0), Inches(0.45),
       "🎵 Fast and silly — three minutes only, then back to reading!", size=16, bold=True,
       color=TEAL)
    notes(slide,
          "Time to wiggle! But only move if Bunny says. Bunny says touch your nose! "
          "Bunny says make the /m/ sound!",
          "Bunny says point to the letter B. Bunny says read the word CAT.",
          "The child follows the commands and points to the right letter or word.",
          "Do the movement together the first time. If they move when you did not say "
          "'Bunny says', laugh it off — this is a rest, not a test.",
          "You are so quick! Ready for some reading?",
          "35–38 min", "Movement plus quick word pointing")


def s16_sentence_intro():
    slide, n = new_slide("📖 Now We Read Sentences!", "READING", "38–45 min", "Sentences", SKY)
    sentence, emoji, color, light = SENTENCES[0]
    add_round(slide, Inches(0.5), Inches(1.5), Inches(12.35), Inches(2.55), WHITE)
    tb(slide, Inches(0.85), Inches(1.95), Inches(1.5), Inches(1.7), emoji, size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.6), Inches(2.15), Inches(9.9), Inches(1.3), sentence, size=54, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
    ido_wedo_youdo(slide, Inches(4.35), SKY)
    tb(slide, Inches(0.55), Inches(6.1), Inches(9.2), Inches(0.42),
       "👉 Point under each word as you read it.", size=17, bold=True, color=CARROT)
    notes(slide,
          "Now we put words together to make a sentence. Watch my finger. I... am... Sam. "
          "Now let's read it together.",
          "Can you read it with me? Now can you try by yourself?",
          "I am Sam.",
          "Read the first two words and let the child finish the last one. Build up "
          "backwards from the end of the sentence.",
          "You read a whole sentence! That is a big jump from words.",
          "38–45 min", "I DO → WE DO → YOU DO")


def _sentence_pair(idx_a, idx_b, title, timing):
    slide, n = new_slide(title, "READING", timing, "Sentences", TEAL)
    for slot, idx in enumerate((idx_a, idx_b)):
        sentence, emoji, color, light = SENTENCES[idx]
        top = Inches(1.6 + slot * 2.6)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(2.35), light)
        tb(slide, Inches(0.85), top + Inches(0.6), Inches(1.4), Inches(1.2), emoji, size=44,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.5), top + Inches(0.55), Inches(8.0), Inches(1.15), sentence,
           size=40, bold=True, color=NAVY, font="Arial Black")
        add_round(slide, Inches(10.6), top + Inches(0.75), Inches(2.0), Inches(0.75), color)
        tb(slide, Inches(10.6), top + Inches(0.9), Inches(2.0), Inches(0.5), "YOU read ⭐",
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    return slide, n


def s17_read_with_me():
    slide, n = _sentence_pair(1, 2, "🤝 Read With Me", "38–45 min")
    notes(slide,
          "My turn first, then we read together, then your turn. I see a cat. "
          "Now with me: I see a cat.",
          "Which word says CAT? Can you read it with me?",
          "I see a cat.  •  The dog is big.",
          "Cover all but the first word, then reveal one word at a time as the child reads. "
          "Do not correct in the middle — wait until the end of the sentence.",
          "Nice reading! Take your time, there is no rush.",
          "38–45 min", "WE DO, then YOU DO if comfortable")


def s18_sentence_practice():
    slide, n = _sentence_pair(3, 5, "⭐ Your Turn to Read", "38–45 min")
    notes(slide,
          "These two are yours. If a word feels tricky, we will sound it out together — "
          "that is what good readers do.",
          "Can you read this one all by yourself?",
          "I can run.  •  I see the sun.",
          "Break the tricky word into sounds, blend them, then say the word and ask the "
          "child to repeat it successfully. Always finish on a success.",
          "You read that all by yourself! I am so proud of you.",
          "38–45 min", "YOU DO with support ready")


def s19_read_and_find():
    slide, n = new_slide("🔎 Game: Read and Find", "GAME", "45–50 min", "Read & find", GRASS)
    tb(slide, Inches(0.45), Inches(1.34), Inches(11.5), Inches(0.4),
       "Look at the park. Read the sentence, then find it in the picture!",
       size=16, bold=True, color=CARROT)
    add_round(slide, Inches(0.5), Inches(1.82), Inches(6.3), Inches(4.55), L_SKY)
    add_round(slide, Inches(0.5), Inches(4.62), Inches(6.3), Inches(1.75), L_GRASS)
    for emoji, x, y in SCENE_ITEMS:
        tb(slide, Inches(0.5 + x), Inches(1.95 + y), Inches(1.3), Inches(1.1), emoji, size=40,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(5.92), Inches(5.9), Inches(0.4),
       "🌳 The Park — can you find the dog?", size=14, bold=True, color=NAVY)
    for i, (sentence, emoji) in enumerate(READ_AND_FIND):
        top = Inches(1.95 + i * 1.15)
        add_round(slide, Inches(7.05), top, Inches(5.8), Inches(0.98), WHITE)
        tb(slide, Inches(7.3), top + Inches(0.26), Inches(4.2), Inches(0.5), sentence, size=19,
           bold=True, color=NAVY)
        add_round(slide, Inches(11.6), top + Inches(0.18), Inches(1.0), Inches(0.62), L_GRASS)
        tb(slide, Inches(11.6), top + Inches(0.24), Inches(1.0), Inches(0.5), emoji, size=18,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(7.05), Inches(5.45), Inches(5.8), Inches(0.9), L_SUN)
    tb(slide, Inches(7.3), Inches(5.62), Inches(5.4), Inches(0.55),
       "Which sentence tells us about the dog?", size=16, bold=True, color=NAVY)
    notes(slide,
          "Here is a park. Let's see what is in it — a sun, a tree, a dog, a ball, and a cat. "
          "Now read this sentence with me: The dog is big.",
          "Can you find the dog in the picture? Which sentence tells us about the dog?",
          "Sentence 1 is about the dog. Sentence 2 is about the sun. Sentence 3 is about "
          "the cat.",
          "Read the sentence aloud yourself, then ask the child to point to the matching "
          "picture. Pointing is easier than reading and still shows understanding.",
          "You found it! You are reading AND thinking.",
          "45–50 min", "Reading plus picture comprehension")


def s20_story_words():
    slide, n = new_slide("📚 Story Words — Get Ready for Bunny's Story", "VOCAB", "50–56 min",
                         "Story words", LILAC)
    tb(slide, Inches(0.45), Inches(1.36), Inches(11.5), Inches(0.4),
       "These words are in Bunny's story. Let's meet them first!", size=16, bold=True,
       color=CARROT)
    for i, (emoji, word, sounds, sentence) in enumerate(VOCAB):
        col, row = i % 4, i // 4
        left = Inches(0.45 + col * 3.18)
        top = Inches(1.85 + row * 2.5)
        add_round(slide, left, top, Inches(3.0), Inches(2.3), L_LILAC)
        tb(slide, left + Inches(0.1), top + Inches(0.12), Inches(2.8), Inches(0.6), emoji,
           size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(0.75), Inches(2.8), Inches(0.6), word,
           size=28, bold=True, color=LILAC, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.1), top + Inches(1.35), Inches(2.8), Inches(0.35), sounds,
           size=13, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), top + Inches(1.72), Inches(2.7), Inches(0.45), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(1.79), Inches(2.7), Inches(0.35), sentence,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
    notes(slide,
          "Before we read Bunny's story, let's meet the words we will see. This one is "
          "BUNNY. Bunny can hop. Your turn: BUNNY.",
          "What is this word? Can you read the little sentence with me?",
          "The child repeats each word and reads at least four of them.",
          "Cover the word and show only the picture, then reveal the word and blend it "
          "together. Skip 'happy' and 'play' if time is short.",
          "You already know so many of these — the story will be easy for you!",
          "50–56 min", "WE DO, teacher models each word")


def _story_slide(index, timing):
    part, emoji, lines, color, light = STORY[index]
    slide, n = new_slide(f"📚 Bunny Goes to the Park — {part}", "STORY", timing, part, color)
    add_round(slide, Inches(0.5), Inches(1.5), Inches(8.9), Inches(4.95), WHITE)
    gap, first, size = (1.12, 1.95, 32) if len(lines) <= 4 else (0.92, 1.78, 28)
    for i, line in enumerate(lines):
        tb(slide, Inches(0.95), Inches(first + i * gap), Inches(8.1), Inches(0.85), line,
           size=size, bold=True, color=NAVY)
    add_round(slide, Inches(9.65), Inches(1.5), Inches(3.2), Inches(4.95), light)
    tb(slide, Inches(9.65), Inches(2.1), Inches(3.2), Inches(1.4), emoji, size=64,
       align=PP_ALIGN.CENTER)
    steps = ["1️⃣  I read it", "2️⃣  We read it", "3️⃣  You pick ONE line"]
    for i, step in enumerate(steps):
        top = Inches(3.95 + i * 0.72)
        add_round(slide, Inches(9.9), top, Inches(2.7), Inches(0.58), WHITE)
        tb(slide, Inches(9.9), top + Inches(0.11), Inches(2.7), Inches(0.4), step, size=14,
           bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    return slide, n


def s21_story_1():
    slide, n = _story_slide(0, "50–56 min")
    notes(slide,
          "Story time! Bunny is going somewhere fun today. Listen while I read, then we "
          "read together, then you pick just one line to read by yourself.",
          "Where is Bunny going? What does Bunny see?",
          "Bunny goes to the park. The sun is big. Bunny sees a dog. The dog is big.",
          "Offer the shortest line: 'The sun is big.' Letting the child choose the line "
          "removes almost all the worry.",
          "You picked a line and read it. That is being a real reader!",
          "50–56 min", "I DO → WE DO → child picks one line")


def s22_story_2():
    slide, n = _story_slide(1, "50–56 min")
    notes(slide,
          "Look what the dog has — a red ball! Let's find out what Bunny says.",
          "What does the dog have? What do you think Bunny will say?",
          "The dog has a red ball. Bunny runs to the dog. \"Can I play?\" says Bunny. "
          "\"Yes!\" says the dog.",
          "Explain that the marks around the words mean somebody is talking. Read the "
          "talking lines in a fun voice and let the child copy you.",
          "I loved your Bunny voice! That was excellent reading.",
          "50–56 min", "I DO → WE DO → child reads a talking line")


def s23_story_3():
    slide, n = _story_slide(2, "50–56 min")
    notes(slide,
          "And here is the happy ending. Bunny and the dog play together all day.",
          "How does Bunny feel? Can you tell me the story in your own words?",
          "They play with the ball. Bunny can run. Bunny is happy. It is a fun day!",
          "Retell the story together using the three story slides as picture prompts, "
          "one sentence per slide.",
          "You read a whole story today! Not just words — a whole story.",
          "50–56 min", "I DO → WE DO → child reads one or two lines")


def s24_comprehension():
    slide, n = new_slide("🧩 Story Questions", "THINK", "56–58 min", "Comprehension", ROSE)
    for i, (question, options, _answer) in enumerate(COMPREHENSION):
        top = Inches(1.55 + i * 1.78)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.6), WHITE)
        add_oval(slide, Inches(0.72), top + Inches(0.5), Inches(0.6), Inches(0.6), ROSE)
        tb(slide, Inches(0.72), top + Inches(0.58), Inches(0.6), Inches(0.45), str(i + 1),
           size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.52), Inches(3.9), Inches(0.6), question,
           size=18, bold=True, color=NAVY)
        for j, (emoji, label) in enumerate(options):
            left = Inches(5.7 + j * 2.4)
            add_round(slide, left, top + Inches(0.28), Inches(2.2), Inches(1.05), L_ROSE)
            tb(slide, left, top + Inches(0.38), Inches(2.2), Inches(0.45), emoji, size=20,
               align=PP_ALIGN.CENTER)
            tb(slide, left + Inches(0.1), top + Inches(0.86), Inches(2.0), Inches(0.38), label,
               size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    notes(slide,
          "Let's think about Bunny's story. Where does Bunny go? Point to your answer.",
          "Where does Bunny go? What does the dog have? How does Bunny feel?",
          "1. Park   2. A ball   3. Happy",
          "Go back to the story slide and read the line that holds the answer, then ask "
          "again. Offer two choices instead of three if needed.",
          "That is exactly right! You remembered the story so well.",
          "56–58 min", "Listening and pointing, no reading needed")


def s25_final_challenge():
    slide, n = new_slide("⭐ Reading Superhero Challenge!", "CHALLENGE", "58–60 min",
                         "Final challenge", CARROT)
    tb(slide, Inches(0.45), Inches(1.36), Inches(11.5), Inches(0.4),
       "Read as many as you can. Every one counts! 🌟", size=16, bold=True, color=CARROT)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(2.0), L_SUN)
    tb(slide, Inches(0.85), Inches(2.02), Inches(3.0), Inches(0.4), "3 WORDS", size=15,
       bold=True, color=NAVY)
    for i, word in enumerate(FINAL_WORDS):
        left = Inches(0.85 + i * 4.0)
        add_round(slide, left, Inches(2.5), Inches(3.7), Inches(1.1), WHITE)
        tb(slide, left, Inches(2.7), Inches(3.7), Inches(0.75), word, size=36, bold=True,
           color=CARROT, align=PP_ALIGN.CENTER, font="Arial Black")
    add_round(slide, Inches(0.5), Inches(4.05), Inches(12.35), Inches(2.15), L_TEAL)
    tb(slide, Inches(0.85), Inches(4.22), Inches(4.0), Inches(0.4), "1 SENTENCE", size=15,
       bold=True, color=NAVY)
    add_round(slide, Inches(0.85), Inches(4.75), Inches(11.65), Inches(1.2), WHITE)
    tb(slide, Inches(1.1), Inches(5.05), Inches(11.15), Inches(0.65), FINAL_SENTENCE, size=40,
       bold=True, color=NAVY, align=PP_ALIGN.CENTER, font="Arial Black")
    star_badge(slide, Inches(10.3), Inches(6.35), "⭐ Every one counts!")
    notes(slide,
          "Last challenge, Reading Superhero! Read as many as you can. There is no pass or "
          "fail here — every word you read is a win.",
          "Can you read this word? Now can you read the sentence?",
          "CAT, SUN, DOG, and 'I see a dog.'",
          "Sound out the first word together to build momentum, then step back and let the "
          "child lead. Count successes aloud: 'That's two! That's three!'",
          "You read three words! Look how much you learned in one class.",
          "58–60 min", "YOU DO — celebrate, never correct harshly")


def s26_recap():
    slide, n = new_slide("✅ Today I Can...", "RECAP", "58–60 min", "Recap", GRASS)
    for i, (icon, text) in enumerate(RECAP):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.65 + row * 2.4)
        add_round(slide, left, top, Inches(3.9), Inches(2.15), WHITE)
        add_oval(slide, left + Inches(1.5), top + Inches(0.22), Inches(0.9), Inches(0.9),
                 L_GRASS)
        tb(slide, left + Inches(1.5), top + Inches(0.36), Inches(0.9), Inches(0.6), icon,
           size=22, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), top + Inches(1.25), Inches(3.5), Inches(0.7), text,
           size=18, bold=True, color=GRASS, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(6.5), Inches(9.5), Inches(0.42),
       "🗣️ What was your favorite game today? What word do you remember?", size=16,
       bold=True, color=CARROT)
    notes(slide,
          "Look at everything you can do now! Letter sounds, blending, words, sentences, "
          "a whole story, and answering questions. That is a lot for one class.",
          "What was your favorite game? What word do you remember best?",
          "The child names a favorite game and recalls at least one or two words.",
          "Prompt with the first sound: 'We read about a b... bunny!'",
          "You should be proud of yourself today. I certainly am!",
          "58–60 min", "Speaking only")


def s27_badge():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    for x, y, c in [(0.7, 0.85, SKY), (11.9, 0.85, ROSE), (0.9, 5.75, GRASS), (11.85, 5.7, SUNNY)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.8), c)
    tb(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(0.8), "🏅 READING SUPERSTAR 🏅",
       size=38, bold=True, color=GOLD, align=PP_ALIGN.CENTER, font="Georgia")
    add_oval(slide, Inches(5.42), Inches(1.95), Inches(2.5), Inches(2.5), GOLD)
    tb(slide, Inches(5.42), Inches(2.45), Inches(2.5), Inches(1.4), "🐰", size=58,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.7), Inches(12), Inches(0.6), "Great job, Reading Explorer!",
       size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(5.35), Inches(12), Inches(0.55),
       "You helped Bunny finish every reading challenge!", size=20, color=WHITE,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(5.98), Inches(12), Inches(0.5),
       "Sounds ⭐ Words ⭐ Sentences ⭐ A Whole Story", size=16, color=L_SKY,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "58–60 min", "Badge")
    fade(slide)
    notes(slide,
          "This badge is for you. You are officially a Reading Superstar. You read sounds, "
          "words, sentences and a whole story today, and Bunny says thank you!",
          "Do you want to show this badge to your family?",
          "A smile, a thank you, or a question about the next class.",
          "If the child seems shy about praise, name one concrete fact: 'You read three "
          "words all by yourself.'",
          "I am so proud of you. See you next time, Reading Explorer!",
          "58–60 min", "Celebration only")


def s28_assessment():
    slide, n = new_slide("📋 How Did We Do Today?", "TEACHER ONLY", "", "Assessment", NAVY)
    tb(slide, Inches(0.45), Inches(1.34), Inches(12.4), Inches(0.35),
       "TEACHER ONLY — do not show this slide to the child. Record progress, not failure.",
       size=13, bold=True, color=CARROT)
    header_y = Inches(1.8)
    widths = [(0.45, 4.6, "SKILL"), (5.2, 2.35, "Emerging"), (7.7, 2.35, "Developing"),
              (10.2, 2.6, "Good")]
    for left, width, label in widths:
        add_round(slide, Inches(left), header_y, Inches(width), Inches(0.5), NAVY)
        tb(slide, Inches(left), header_y + Inches(0.1), Inches(width), Inches(0.34), label,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = header_y + Inches(0.6 + i * 0.62)
        band = WHITE if i % 2 == 0 else L_SKY
        add_round(slide, Inches(0.45), top, Inches(4.6), Inches(0.54), band)
        tb(slide, Inches(0.7), top + Inches(0.1), Inches(4.2), Inches(0.36), skill, size=14,
           bold=True, color=NAVY)
        for left, width, _label in widths[1:]:
            add_round(slide, Inches(left), top, Inches(width), Inches(0.54), band)
            tb(slide, Inches(left), top + Inches(0.08), Inches(width), Inches(0.4), "☐",
               size=16, color=SOFT, align=PP_ALIGN.CENTER)
    notes(slide,
          "TEACHER SLIDE — do not display to the child.",
          "Fill this in within five minutes of finishing, while the details are fresh.",
          "A Grade 1 child who struggles with reading will often sit at Emerging for "
          "sentence reading and Developing for comprehension, because listening "
          "comprehension usually runs ahead of decoding.",
          "Record how many words and sentences were read independently in the Reading "
          "Superhero Challenge — that single number is the clearest way to show progress "
          "to the parent next lesson.",
          "Frame parent feedback as strengths first, then one growth area, then the plan.",
          "After class", "Teacher record")


def s29_answer_key_a():
    slide, n = new_slide("🔑 Teacher Answer Key — Part 1", "DO NOT SHOW STUDENT", "",
                         "Answer key", NAVY)
    tb(slide, Inches(0.45), Inches(1.34), Inches(12.4), Inches(0.35),
       "TEACHER ANSWER KEY — DO NOT SHOW STUDENT. Hide this slide before presenting.",
       size=13, bold=True, color=CARROT)
    add_round(slide, Inches(0.45), Inches(1.72), Inches(6.1), Inches(4.95), WHITE)
    tb(slide, Inches(0.7), Inches(1.9), Inches(5.6), Inches(0.4), "🎯 Phonics & Blending",
       size=15, bold=True, color=ROSE)
    bullets(slide, Inches(0.7), Inches(2.35), Inches(5.6), Inches(4.2), [
        "Sound Detective 1. CAT = /k/  (B)",
        "Sound Detective 2. MAP = /m/  (A)",
        "Sound Detective 3. SUN = /s/  (A)",
        "Sound Detective 4. TREE = /t/  (A)",
        "Sound Detective 5. PIG = /p/  (B)",
        "Sound Detective 6. DOG = /d/  (A)",
        "Build the Word: MAT, SUN, PIG, DOG",
        "Letter sounds: m s t p c b d r",
    ])
    add_round(slide, Inches(6.75), Inches(1.72), Inches(6.1), Inches(4.95), WHITE)
    tb(slide, Inches(7.0), Inches(1.9), Inches(5.6), Inches(0.4), "🐰 Word & Sight Word Games",
       size=15, bold=True, color=TEAL)
    bullets(slide, Inches(7.0), Inches(2.35), Inches(5.6), Inches(4.2), [
        "Feed the Bunny: CAT, SUN, DOG, PIG, HEN",
        "Match the Word: cat=CAT, sun=SUN,",
        "     pig=PIG, dog=DOG",
        "Catch the Sight Word 1. I (2nd)",
        "Catch the Sight Word 2. am (2nd)",
        "Catch the Sight Word 3. the (2nd)",
        "Catch the Sight Word 4. can (1st)",
        "Catch the Sight Word 5. my (2nd)",
        "Catch the Sight Word 6. see (2nd)",
    ])
    notes(slide,
          "TEACHER SLIDE — answers for the phonics, blending, CVC and sight word games.",
          "Hide this slide in PowerPoint (right-click the thumbnail, then Hide Slide) "
          "before presenting, or keep it open on a second screen.",
          "See the slide content for all answers.",
          "Use the Sound Detective and Catch the Sight Word results to choose which sounds "
          "and words to review at the start of the next lesson.",
          "Reference only.",
          "Reference", "Teacher only")


def s30_answer_key_b():
    slide, n = new_slide("🔑 Teacher Answer Key — Part 2", "DO NOT SHOW STUDENT", "",
                         "Answer key", NAVY)
    tb(slide, Inches(0.45), Inches(1.34), Inches(12.4), Inches(0.35),
       "TEACHER ANSWER KEY — DO NOT SHOW STUDENT.", size=13, bold=True, color=CARROT)
    add_round(slide, Inches(0.45), Inches(1.72), Inches(6.1), Inches(4.95), WHITE)
    tb(slide, Inches(0.7), Inches(1.9), Inches(5.6), Inches(0.4), "📖 Sentences & Read and Find",
       size=15, bold=True, color=SKY)
    bullets(slide, Inches(0.7), Inches(2.35), Inches(5.6), Inches(4.2), [
        "Sentences taught: I am Sam. / I see a cat. /",
        "     The dog is big. / I can run. /",
        "     My cat is red. / I see the sun.",
        "Read & Find scene: sun, tree, dog, ball, cat",
        "\"The dog is big.\" → the dog",
        "\"I see the sun.\" → the sun",
        "\"The cat is on the mat.\" → the cat",
        "Final challenge: CAT, SUN, DOG,",
        "     \"I see a dog.\"",
    ])
    add_round(slide, Inches(6.75), Inches(1.72), Inches(6.1), Inches(4.95), WHITE)
    tb(slide, Inches(7.0), Inches(1.9), Inches(5.6), Inches(0.4), "📚 Story & Comprehension",
       size=15, bold=True, color=GRASS)
    bullets(slide, Inches(7.0), Inches(2.35), Inches(5.6), Inches(4.2), [
        "Story: \"Bunny Goes to the Park\" (61 words)",
        "Q1. Where does Bunny go?  →  B. Park",
        "Q2. What does the dog have?  →  A. A ball",
        "Q3. How does Bunny feel?  →  A. Happy",
        "Story words: bunny, park, dog, ball,",
        "     red, run, play, happy",
        "Accept any reasonable spoken phrasing",
        "     for the comprehension answers.",
    ])
    notes(slide,
          "TEACHER SLIDE — answers for the sentence work, Read and Find, the story and the "
          "comprehension questions.",
          "Hide this slide before presenting.",
          "See the slide content for all answers.",
          "For a child who finishes early, ask them to point to the line in the story that "
          "proves each comprehension answer.",
          "Reference only.",
          "Reference", "Teacher only")


BUILDERS = [
    s01_welcome, s02_meet_bunny, s03_point_and_say, s04_letter_sounds,
    s05_detective_a, s06_detective_b, s07_blending_intro, s08_build_the_word,
    s09_word_families, s10_feed_bunny, s11_match_word, s12_sight_words,
    s13_catch_a, s14_catch_b, s15_brain_break, s16_sentence_intro,
    s17_read_with_me, s18_sentence_practice, s19_read_and_find, s20_story_words,
    s21_story_1, s22_story_2, s23_story_3, s24_comprehension, s25_final_challenge,
    s26_recap, s27_badge, s28_assessment, s29_answer_key_a, s30_answer_key_b,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

out = r"C:\Users\bhushaja\Downloads\shaip\Grade1_Bunny_Reading_Adventure_60min.pptx"
try:
    prs.save(out)
except PermissionError:
    out = out.replace(".pptx", "_new.pptx")
    prs.save(out)

story_words = sum(len(line.split()) for _p, _e, lines, _c, _l in STORY for line in lines)
missing = [i + 1 for i, s in enumerate(prs.slides)
           if not s.has_notes_slide or not s.notes_slide.notes_text_frame.text.strip()]
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
print(f"Story word count: {story_words}")
print(f"Slides missing teacher notes: {missing if missing else 'none'}")
