"""Premium Grade 5 English Demo for Sage — 90 min, 30 slides, game-based lesson."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

NAVY = RGBColor(0x0A, 0x2A, 0x6E)
SKY = RGBColor(0x29, 0xB6, 0xF6)
TEAL = RGBColor(0x00, 0x96, 0x88)
YELLOW = RGBColor(0xFF, 0xD5, 0x4F)
GOLD = RGBColor(0xFF, 0xB3, 0x00)
CORAL = RGBColor(0xFF, 0x6F, 0x61)
PINK = RGBColor(0xF0, 0x62, 0x92)
PURPLE = RGBColor(0x7E, 0x57, 0xC2)
LIME = RGBColor(0xA5, 0xD6, 0xA7)
CREAM = RGBColor(0xFF, 0xFB, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x29, 0x37)
SOFT = RGBColor(0x64, 0x74, 0x8B)
LIGHT_SKY = RGBColor(0xE1, 0xF5, 0xFE)
LIGHT_TEAL = RGBColor(0xE0, 0xF7, 0xFA)
LIGHT_YELLOW = RGBColor(0xFF, 0xF9, 0xC4)
LIGHT_CORAL = RGBColor(0xFF, 0xEB, 0xEE)
LIGHT_PURPLE = RGBColor(0xF3, 0xE5, 0xF5)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
TOTAL = 30
TEACHER_NAME = "Your Teacher Name"
STUDENT = "Sage"

PASSAGE = (
    "On Monday morning, Leo rushed into Oakwood Elementary with his favorite blue backpack "
    "swinging at his side. Inside were his math notebook, a shiny pencil case, and a permission slip "
    "for the field trip to the science museum.\n\n"
    "After lunch, Leo opened his locker and froze. The backpack was gone. Only a crumpled sticky note "
    "remained: \"Check the reading nook.\" Leo's heart beat fast. Had someone taken it as a joke?\n\n"
    "He searched the hallway, then the library reading nook. Under a beanbag chair, he spotted a corner "
    "of blue fabric. His backpack! Beside it sat a note from Ms. Carter: \"Found this near the art room. "
    "Looks like it slid off the cart during delivery. — Signed, the Mystery Helper.\"\n\n"
    "Leo laughed with relief. He thanked Ms. Carter and promised to clip his bag to the cart next time. "
    "That afternoon, his class learned about observation skills. Leo raised his hand and said, "
    "\"Clues help you solve mysteries—even small ones.\"\n\n"
    "By dismissal, Leo had his backpack, his permission slip, and a new habit: double-checking where "
    "he left important things. The mystery was solved, and Leo felt proud."
)

LISTENING = (
    "Last Friday, our school held a reading festival. Students built tiny book tents, shared stories, "
    "and listened to a visiting author. Sage's class created a poster about teamwork. "
    "They won second place and celebrated with fruit popsicles."
)


def set_run(run, size=18, bold=False, color=DARK, font="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_bg(slide, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    spTree = slide.shapes._spTree
    el = s._element
    spTree.remove(el)
    spTree.insert(2, el)


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


def tb(slide, l, t, w, h, text, size=18, bold=False, color=DARK, align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    set_run(r, size, bold, color, font)
    return box


def bullets(slide, l, t, w, h, items, size=15, color=DARK, sp=8):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(sp)
        r = p.add_run()
        r.text = "•  " + item
        set_run(r, size, False, color)


def teacher_notes(slide, say, activity, expected, assess, mistakes="", extend="", timing=""):
    # Speaker notes omitted from deliverable PPT; keep calls for script structure.
    return


def fade_transition(slide):
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    etree.SubElement(tr, qn("p:fade"))


def footer(slide, n, timing=""):
    add_rect(slide, Inches(0), Inches(7.12), prs.slide_width, Inches(0.38), NAVY)
    msg = f"Grade 5 Premium Demo  |  {STUDENT}  |  90 min"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.15), Inches(11.5), Inches(0.3), msg, size=10, color=WHITE)
    tb(slide, Inches(12.0), Inches(7.15), Inches(1.1), Inches(0.3), f"{n}/{TOTAL}", size=10,
       color=WHITE, align=PP_ALIGN.RIGHT)


def chip(slide, text, color=SKY, w=Inches(2.9)):
    add_round(slide, Inches(0.38), Inches(0.28), w, Inches(0.36), color)
    tb(slide, Inches(0.38), Inches(0.29), w, Inches(0.34), text, size=11, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def time_chip(slide, text):
    add_round(slide, Inches(10.45), Inches(0.28), Inches(2.5), Inches(0.36), CORAL)
    tb(slide, Inches(10.45), Inches(0.29), Inches(2.5), Inches(0.34), text, size=11, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)


def btn(slide, text="Next Adventure →"):
    add_round(slide, Inches(10.2), Inches(6.55), Inches(2.9), Inches(0.45), GOLD)
    tb(slide, Inches(10.2), Inches(6.58), Inches(2.9), Inches(0.4), text, size=11, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER)


def deco_stars(slide):
    for x, y, c in [(0.5, 0.9, YELLOW), (12.3, 1.2, PINK), (11.5, 5.8, SKY), (0.8, 5.5, TEAL)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.5), Inches(0.5), c)


def slide_header(slide, title, tag="ACTIVITY", timing="", page=1):
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.14), prs.slide_height, SKY)
    chip(slide, tag)
    if timing:
        time_chip(slide, timing)
    tb(slide, Inches(0.38), Inches(0.78), Inches(12.5), Inches(0.55), title, size=26, bold=True,
       color=NAVY, font="Georgia")
    footer(slide, page, timing)
    fade_transition(slide)


# ---------- 30 SLIDES ----------

def s01():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), YELLOW)
    deco_stars(slide)
    for x, y in [(1.0, 1.2), (11.5, 1.5), (2.5, 5.8)]:
        add_oval(slide, Inches(x), Inches(y), Inches(1.1), Inches(1.1), RGBColor(0x15, 0x65, 0xC0))
    tb(slide, Inches(0.7), Inches(2.0), Inches(12), Inches(0.5), "🌟 Welcome to Your English Adventure 🌟",
       size=36, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(3.0), Inches(12), Inches(0.45),
       f"Student: {STUDENT}  •  Grade 5  •  U.S.  •  90-Minute Game Quest",
       size=18, color=LIGHT_SKY, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.5), Inches(3.8), Inches(6.3), Inches(1.2), TEAL)
    tb(slide, Inches(3.7), Inches(4.05), Inches(5.9), Inches(0.7),
       f"Teacher: {TEACHER_NAME}\nToday = Play + Learn (NOT a test!)",
       size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(5.4), Inches(12), Inches(0.4),
       "📚 ✏️ 💬 🎲 — Let's go!", size=20, color=YELLOW, align=PP_ALIGN.CENTER)
    footer(slide, 1, "Welcome 5 min")
    fade_transition(slide)
    teacher_notes(slide,
        f"Hi {STUDENT}! Welcome to your English adventure. We will play games, read a mystery, and level up your English superpowers.",
        "Wave hello; say one word that means 'fun' to you.",
        "Student greets; may say excited/nervous/happy.",
        "Watch eye contact, volume, first-sentence length.",
        "One-word answers only at first.",
        "If shy: offer choices (cat/dog/sports?).",
        "5 min")


def s02():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Meet Your Teacher 👋", "WELCOME", "5 min", 2)
    add_round(slide, Inches(0.4), Inches(1.45), Inches(6.0), Inches(5.2), WHITE)
    tb(slide, Inches(0.65), Inches(1.65), Inches(5.5), Inches(0.35), "About Me", size=18, bold=True, color=TEAL)
    bullets(slide, Inches(0.65), Inches(2.1), Inches(5.5), Inches(4.0), [
        f"I am {TEACHER_NAME}.",
        "Fun fact: I turn grammar into detective games!",
        "Favorite book: Charlotte's Web",
        "Mission: Help you feel brave in English",
    ], size=16)
    add_round(slide, Inches(6.7), Inches(1.45), Inches(6.2), Inches(5.2), LIGHT_YELLOW)
    tb(slide, Inches(6.95), Inches(1.65), Inches(5.7), Inches(0.35), f"Now You, {STUDENT}! 🌟", size=18, bold=True, color=CORAL)
    bullets(slide, Inches(6.95), Inches(2.2), Inches(5.7), Inches(4.0), [
        "What's your favorite animal?",
        "What's your favorite movie?",
        "If you could have one superpower, what would it be?",
    ], size=17, sp=14)
    btn(slide)
    teacher_notes(slide,
        "Share warmly, then invite Sage to answer one question at a time.",
        "Answer 3 questions in full sentences.",
        "Animal/movie/superpower with short reason.",
        "Fluency, pronunciation, confidence—not perfection.",
        "Listing without because.",
        "Why that animal? What would you do with that power?",
        "5 min")


def s03():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Icebreaker: Emoji Guess Game 😊😎🐶🍕", "GAME", "Icebreaker 10 min", 3)
    emojis = ["😊", "😎", "🐶", "🍕", "⚽", "📚", "🌈", "🚀"]
    for i, e in enumerate(emojis):
        col, row = i % 4, i // 4
        left = Inches(0.45 + col * 3.2)
        top = Inches(1.5 + row * 2.45)
        add_round(slide, left, top, Inches(3.05), Inches(2.2), WHITE)
        add_oval(slide, left + Inches(0.95), top + Inches(0.25), Inches(1.15), Inches(1.15), LIGHT_SKY)
        tb(slide, left + Inches(0.95), top + Inches(0.45), Inches(1.15), Inches(0.7), e, size=34, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.55), Inches(2.75), Inches(0.5),
           "Make 1 full sentence!", size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(6.35), Inches(12.5), Inches(0.35),
       "Example: \"I feel happy when I play with my dog.\"", size=14, color=SOFT, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Pick 4 emojis. Model: 'I feel happy when…'",
        "One sentence per emoji (4–8 total).",
        "Subject + verb + detail.",
        "Sentence completeness, creativity.",
        "Fragment sentences.",
        "Can you add an adjective?",
        "Part of 10 min icebreaker block")


def s04():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Would You Rather? 🎯", "GAME", "Icebreaker", 4)
    qs = [
        ("Fly", "Become invisible"),
        ("Live in a treehouse", "Live on a cloud"),
        ("Talk to animals", "Read every mind"),
        ("Always summer", "Always winter"),
        ("Be a famous singer", "Be a game designer"),
        ("Explore the ocean", "Explore outer space"),
        ("Have a robot helper", "Have a magic backpack"),
        ("Eat only breakfast", "Eat only dinner"),
    ]
    for i, (a, b) in enumerate(qs):
        col, row = i % 4, i // 4
        left = Inches(0.35 + col * 3.25)
        top = Inches(1.45 + row * 2.5)
        add_round(slide, left, top, Inches(3.1), Inches(2.35), WHITE)
        tb(slide, left + Inches(0.1), top + Inches(0.15), Inches(2.9), Inches(0.7), a, size=13, bold=True,
           color=TEAL, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(0.95), Inches(2.9), Inches(0.3), "OR", size=12, bold=True,
           color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(1.35), Inches(2.9), Inches(0.7), b, size=13, bold=True,
           color=CORAL, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(6.35), Inches(12.5), Inches(0.35),
       "Say: \"I would rather ___ because ___\"", size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Do 5–6 questions. Celebrate reasons!",
        "Choose + explain why.",
        "Full sentence with because.",
        "Speaking confidence, reasoning.",
        "No reason given.",
        "What is the downside of your choice?",
        "Icebreaker 10 min total")


def s05():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Vocabulary Warm-Up: Word Hunt 🔍", "GAME", "Vocab 10 min", 5)
    add_round(slide, Inches(0.4), Inches(1.4), Inches(8.0), Inches(5.3), LIGHT_SKY)
    tb(slide, Inches(0.6), Inches(1.55), Inches(7.6), Inches(0.35), "Park Scene — Find the words!", size=16, bold=True, color=NAVY)
    tb(slide, Inches(0.6), Inches(2.0), Inches(7.6), Inches(4.5),
       "🌳 TREE   🪁 KITE   🐿️ SQUIRREL   🌤️ CLOUD\n"
       "🛝 SLIDE   🌸 FLOWER   🚲 BICYCLE   📖 BOOK\n"
       "🎒 BACKPACK   🕵️ CLUE   🔍 MYSTERY   🏫 SCHOOL",
       size=20, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(8.6), Inches(1.4), Inches(4.3), Inches(5.3), WHITE)
    bullets(slide, Inches(8.85), Inches(1.7), Inches(3.9), Inches(4.8), [
        "Point to 6 hidden words",
        "Say each in a sentence",
        "Bonus: use mystery + clue",
    ], size=15, sp=12)
    btn(slide)
    teacher_notes(slide,
        "Treat picture as interactive hunt.",
        "Find words; 6 sentences aloud.",
        "Uses target words correctly.",
        "Vocabulary activation before reading.",
        "Confuses clue/clueless.",
        "Use 'mystery' in a sentence about Leo.",
        "Vocab block")


def s06():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "📖 The Mystery of the Missing Backpack", "READING", "Reading 20 min", 6)
    add_round(slide, Inches(0.4), Inches(1.35), Inches(12.5), Inches(5.35), WHITE)
    add_round(slide, Inches(0.55), Inches(1.5), Inches(2.3), Inches(5.0), LIGHT_YELLOW)
    tb(slide, Inches(0.65), Inches(2.3), Inches(2.1), Inches(0.5), "🎒🔍", size=36, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.65), Inches(3.2), Inches(2.1), Inches(1.5), "mystery\nclues\nobservation", size=13,
       color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(3.05), Inches(1.55), Inches(9.6), Inches(5.0), PASSAGE, size=13, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Read with expression. Pause at mystery moments.",
        "Listen first; then read 1 paragraph.",
        "Follows plot; asks about unknown words.",
        "Decoding, fluency, engagement.",
        "Losing place in text.",
        "Predict: Where is the backpack?",
        "Reading block ~8 min")


def s07():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Read Together + Tricky Words 🔊", "READING", "Reading", 7)
    words = [
        ("permission", "per-MISH-un"),
        ("observation", "ob-ser-VAY-shun"),
        ("crumpled", "Krum-pled"),
        ("delivery", "duh-LIV-er-ee"),
        ("dismissal", "dis-MIS-ul"),
        ("elementary", "el-uh-MEN-tuh-ree"),
    ]
    for i, (w, p) in enumerate(words):
        col, row = i % 3, i // 3
        left = Inches(0.4 + col * 4.3)
        top = Inches(1.5 + row * 2.5)
        add_round(slide, left, top, Inches(4.05), Inches(2.25), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(0.25), Inches(3.75), Inches(0.55), w, size=17, bold=True,
           color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.9), Inches(3.75), Inches(0.5), p, size=14, color=TEAL,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.55), Inches(3.75), Inches(0.45), "Listen → Repeat → Use!", size=12,
           color=SOFT, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Clap syllables; choral repeat.",
        "Repeat each word; one sentence each.",
        "Approximate American pronunciation.",
        "Pronunciation + usage.",
        "Stress on wrong syllable.",
        "Use 'observation' about the story.",
        "Reading 20 min block")


def s08():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Reading Comprehension (10 Questions)", "READING", "Reading", 8)
    qs = [
        "1. MCQ: Where does Leo go to school? A) Oakwood B) Maple C) River",
        "2. MCQ: What color is the backpack? A) Red B) Blue C) Green",
        "3. T/F: Leo lost his backpack after lunch.",
        "4. T/F: Ms. Carter wrote the sticky note in the hallway locker.",
        "5. Short: What was inside the backpack?",
        "6. Short: Where was the backpack found?",
        "7. Inference: Why did Leo feel proud at the end?",
        "8. Opinion: Was the Mystery Helper kind? Why?",
        "9. Vocab: \"Crumpled\" means… A) smooth B) wrinkled C) colorful",
        "10. Sequence: Put in order: found note → searched nook → opened locker",
    ]
    for i, q in enumerate(qs):
        top = Inches(1.35 + (i % 5) * 1.05)
        left = Inches(0.35 if i < 5 else 6.7)
        if i == 5:
            pass
        row = i if i < 5 else i - 5
        top = Inches(1.35 + row * 1.05)
        add_round(slide, left, top, Inches(6.2 if i < 5 else 6.35), Inches(0.95), WHITE if i % 2 == 0 else LIGHT_TEAL)
        tb(slide, left + Inches(0.15), top + Inches(0.2), Inches(6.0), Inches(0.65), q, size=11, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "One question at a time; praise text evidence.",
        "Answer all 10 (or 6 if time tight).",
        "1A 2B 3T 4F 5 notebook/pencil case/slip 6 reading nook 7 solved mystery/habit 8 yes+reason 9B 10 locker→nook→note",
        "Literal, inferential, opinion skills.",
        "Guessing without story.",
        "Which sentence proves your answer?",
        "Reading 20 min")


def s09():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Speaking Challenge: Retell the Mystery 🗣️", "SPEAKING", "Speaking 10 min", 9)
    steps = ["Who?", "What happened?", "Problem?", "Clues?", "Solution?", "Lesson?"]
    for i, st in enumerate(steps):
        left = Inches(0.4 + (i % 3) * 4.25)
        top = Inches(1.5 + (i // 3) * 2.4)
        add_round(slide, left, top, Inches(4.05), Inches(2.15), WHITE)
        add_oval(slide, left + Inches(1.55), top + Inches(0.25), Inches(0.95), Inches(0.95), SKY)
        tb(slide, left + Inches(1.55), top + Inches(0.45), Inches(0.95), Inches(0.55), str(i + 1), size=20, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), top + Inches(1.45), Inches(3.65), Inches(0.55), st, size=16, bold=True, color=NAVY,
           align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Student retells without reading slide.",
        "60–90 second story retell using prompts.",
        "Clear beginning/middle/end; past tense.",
        "Oral organization, confidence.",
        "Copying exact sentences.",
        "Add how Leo felt at each step.",
        "Speaking 10 min")


def s10():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Grammar: Tenses Time Machine ⏰", "GRAMMAR", "Grammar 25 min", 10)
    tenses = [
        ("Present Simple", "Every day", "Leo checks his backpack.", SKY),
        ("Present Continuous", "Right now", "Leo is searching the nook.", TEAL),
        ("Past Simple", "Yesterday", "Leo found his backpack.", CORAL),
        ("Future Simple", "Tomorrow", "Leo will double-check his bag.", PURPLE),
    ]
    for i, (t, when, ex, c) in enumerate(tenses):
        left = Inches(0.35 + i * 3.25)
        add_round(slide, left, Inches(1.45), Inches(3.1), Inches(4.9), WHITE)
        add_rect(slide, left, Inches(1.45), Inches(3.1), Inches(1.0), c)
        tb(slide, left, Inches(1.55), Inches(3.1), Inches(0.35), "🎬", size=22, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(2.05), Inches(2.9), Inches(0.55), t, size=14, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(2.85), Inches(2.8), Inches(0.45), when, size=13, bold=True, color=NAVY,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.45), Inches(2.8), Inches(1.2), ex, size=14, color=DARK,
           align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Act out time words (now/yesterday/tomorrow).",
        "Repeat examples; make one new sentence each tense.",
        "Correct tense markers (-s, is __ing, -ed, will).",
        "Tense concept understanding.",
        "Mixing past and present.",
        "Change 'found' to future.",
        "Grammar 25 min")


def s11():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Grammar Activity: Choose the Correct Tense ✔️", "GRAMMAR", "Grammar", 11)
    items = [
        "Leo (check / checks / checking) his locker every day.",
        "Right now, Ms. Carter (help / is helping / helped) Leo.",
        "Yesterday, Leo (lose / lost / will lose) his backpack.",
        "Tomorrow, he (double-check / double-checked / will double-check).",
        "The class (learn / learned / is learn) about clues today.",
    ]
    for i, it in enumerate(items):
        top = Inches(1.45 + i * 1.05)
        add_round(slide, Inches(0.4), top, Inches(12.5), Inches(0.95), WHITE)
        tb(slide, Inches(0.65), top + Inches(0.28), Inches(12.0), Inches(0.45), f"{i + 1}.  {it}", size=15, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Think aloud for #1.",
        "Say the correct verb form.",
        "checks; is helping; lost; will double-check; learned/is learning (present: learns if daily)",
        "Accuracy under light pressure.",
        "For #5 accept 'learns' if habitual present.",
        "Make your own sentence in past tense.",
        "Grammar block")


def s12():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Grammar Game: Tense Detective 🕵️", "GAME", "Grammar", 12)
    sents = [
        "Leo will found his backpack yesterday.",
        "She are reading the mystery note now.",
        "Tomorrow, I am go to the museum.",
        "Leo checked his bag every day.",
        "Ms. Carter help Leo last Monday.",
    ]
    for i, s in enumerate(sents):
        top = Inches(1.45 + i * 1.05)
        add_round(slide, Inches(0.4), top, Inches(12.5), Inches(0.95), LIGHT_CORAL if i % 2 else WHITE)
        tb(slide, Inches(0.65), top + Inches(0.28), Inches(12.0), Inches(0.45),
           f"Case {i + 1}:  {s}", size=15, color=DARK)
    tb(slide, Inches(0.4), Inches(6.35), Inches(12.5), Inches(0.35),
       "Fix the tense!  (#4 is the innocent sentence ✅)", size=14, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Detective theme—ding when fixed!",
        "Spot error + say corrected sentence.",
        "1 Leo found/will find 2 is reading 3 will go 4 OK 5 helped",
        "Error detection.",
        "Changing correct #4.",
        "Write one 'wrong' sentence for teacher to fix.",
        "Grammar 25 min")


def s13():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Spin the Wheel: Sentence Builder 🎡", "GAME", "Grammar", 13)
    cats = [("Verb", "🏃", CORAL), ("Noun", "📦", SKY), ("Adjective", "🎨", GOLD), ("Adverb", "⚡", PURPLE)]
    for i, (t, e, c) in enumerate(cats):
        left = Inches(0.5 + i * 3.15)
        add_oval(slide, left + Inches(0.55), Inches(1.8), Inches(2.0), Inches(2.0), c)
        tb(slide, left + Inches(0.55), Inches(2.45), Inches(2.0), Inches(0.7), e, size=28, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(4.0), Inches(2.8), Inches(0.45), t, size=18, bold=True, color=NAVY,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(4.8), Inches(12.5), Inches(1.5),
       "Spin (roll dice)! Pick a category → say a word → build a sentence.\n"
       "Example: Adjective=shiny + Noun=backpack + Verb=slides + Adverb=quickly",
       size=16, color=DARK, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Use dice or random number 1–4.",
        "4 rounds of sentence building.",
        "Grammatically sensible sentences.",
        "Parts of speech in production.",
        "Adjective used as verb.",
        "Add a second adjective.",
        "Grammar games")


def s14():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Word Builder: Prefixes • Suffixes • Compounds 🧱", "VOCAB", "Vocab 10 min", 14)
    cols = [
        ("Prefix re-", ["rewrite", "replay", "return"], SKY),
        ("Suffix -ful", ["helpful", "thankful", "colorful"], TEAL),
        ("Compound", ["backpack", "classroom", "sunflower"], GOLD),
    ]
    for i, (title, words, c) in enumerate(cols):
        left = Inches(0.4 + i * 4.25)
        add_round(slide, left, Inches(1.45), Inches(4.05), Inches(4.9), WHITE)
        add_rect(slide, left, Inches(1.45), Inches(4.05), Inches(0.55), c)
        tb(slide, left, Inches(1.55), Inches(4.05), Inches(0.35), title, size=14, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        bullets(slide, left + Inches(0.2), Inches(2.2), Inches(3.65), Inches(3.5), words, size=16, sp=10)
        tb(slide, left + Inches(0.2), Inches(5.5), Inches(3.65), Inches(0.6), "Build 1 new word!", size=13,
           bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Explain re- = again; -ful = full of.",
        "Create 3 new examples aloud.",
        "Reasonable word formations.",
        "Morphology awareness.",
        "Invalid compounds.",
        "Use 'helpful' about Mystery Helper.",
        "Vocab 10 min")


def s15():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Word Search Puzzle 🔤", "GAME", "Vocab", 15)
    grid = (
        "MYSTERYX\n"
        "BACKPACK\n"
        "CLUEXXXO\n"
        "OBSERVEX\n"
        "SCHOOLXX\n"
        "FUTUREXX"
    )
    add_round(slide, Inches(0.4), Inches(1.4), Inches(7.5), Inches(5.3), WHITE)
    tb(slide, Inches(0.6), Inches(1.8), Inches(7.1), Inches(4.5), grid, size=22, bold=True, color=NAVY, font="Consolas",
       align=PP_ALIGN.CENTER)
    bullets(slide, Inches(8.2), Inches(1.7), Inches(4.7), Inches(4.5), [
        "Find: MYSTERY",
        "BACKPACK",
        "CLUE",
        "OBSERVE",
        "SCHOOL",
        "FUTURE",
    ], size=16, sp=8)
    btn(slide)
    teacher_notes(slide,
        "Timer 3 minutes—game feel!",
        "Circle words on screen or paper.",
        "Finds 4+ words.",
        "Word recognition speed.",
        "Spelling confusion.",
        "Use FUTURE in a sentence.",
        "Vocab games")


def s16():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Crossword Clues (Vocabulary) 🧩", "GAME", "Vocab", 16)
    clues = [
        "1 Across: Something you carry books in (8)",
        "2 Down: A hint that helps solve a mystery (4)",
        "3 Across: Watching carefully (9)",
        "4 Down: Time that comes after today (6)",
        "5 Across: Leo's school type (10)",
    ]
    for i, c in enumerate(clues):
        top = Inches(1.45 + i * 1.05)
        add_round(slide, Inches(0.4), top, Inches(12.5), Inches(0.95), LIGHT_PURPLE if i % 2 else WHITE)
        tb(slide, Inches(0.65), top + Inches(0.28), Inches(12.0), Inches(0.45), c, size=15, color=DARK)
    tb(slide, Inches(0.4), Inches(6.35), Inches(12.5), Inches(0.35),
       "Answers: BACKPACK, CLUE, OBSERVATION, FUTURE, ELEMENTARY", size=12, color=TEAL, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Reveal answers after attempt—or hide bottom line during play.",
        "Solve clues aloud.",
        "See bottom line (teacher key).",
        "Spelling + meaning.",
        "Letter count ignored.",
        "Create one new clue.",
        "Vocab")


def s17():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Memory Game 🧠 (12 Pictures)", "GAME", "Brain break", 17)
    pics = "🎒 📚 🕵️ 🔍 🌳 🐶 🍕 ⚽ 🌈 🚀 🎨 🎤"
    add_round(slide, Inches(1.5), Inches(1.5), Inches(10.3), Inches(3.2), LIGHT_YELLOW)
    tb(slide, Inches(1.7), Inches(2.0), Inches(9.9), Inches(2.0), pics, size=36, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(4.9), Inches(12.5), Inches(1.2),
       "Step 1: Look 20 seconds\nStep 2: Hide screen\nStep 3: Name as many as you remember!",
       size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Cover screen after 20 sec.",
        "Recall list; compare.",
        "6+ items strong; 4+ OK.",
        "Working memory, speaking.",
        "Random guessing.",
        "Which item was hardest?",
        "Every 5–7 min activity")


def s18():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Picture Description: Sunny Park 🌳", "SPEAKING", "Speaking", 18)
    add_round(slide, Inches(0.4), Inches(1.4), Inches(7.8), Inches(5.3), LIGHT_TEAL)
    tb(slide, Inches(0.6), Inches(2.2), Inches(7.4), Inches(3.5),
       "☀️ Sunny day\n👧 Kids playing\n🐕 Dog running\n🌸 Flowers blooming\n😊 Happy faces",
       size=22, color=NAVY, align=PP_ALIGN.CENTER)
    boxes = ["People", "Actions", "Weather", "Feelings"]
    for i, b in enumerate(boxes):
        top = Inches(1.5 + i * 1.25)
        add_round(slide, Inches(8.5), top, Inches(4.4), Inches(1.05), WHITE)
        tb(slide, Inches(8.7), top + Inches(0.3), Inches(4.0), Inches(0.45), b, size=16, bold=True, color=CORAL)
    btn(slide)
    teacher_notes(slide,
        "90-second description.",
        "Cover all 4 boxes.",
        "Uses adjectives + present continuous.",
        "Descriptive language.",
        "Only listing nouns.",
        "What might happen next in the park?",
        "Speaking 10 min")


def s19():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Story Cubes 🎲 Roll & Create", "GAME", "Creativity", 19)
    dice = ["🏫 School", "🎒 Backpack", "👻 Surprise", "🌧️ Rain", "🦸 Hero", "🗺️ Map"]
    for i, d in enumerate(dice):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.2)
        top = Inches(1.55 + row * 2.45)
        add_round(slide, left, top, Inches(3.9), Inches(2.2), WHITE)
        tb(slide, left + Inches(0.15), top + Inches(0.7), Inches(3.6), Inches(0.8), d, size=18, bold=True,
           color=NAVY, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(6.35), Inches(12.5), Inches(0.35),
       "Roll 3 dice → 5-sentence story (beginning, middle, end)", size=14, bold=True, color=TEAL,
       align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Roll dice or pick 3 cards.",
        "5-sentence original story.",
        "Coherent plot.",
        "Creativity + grammar in speech.",
        "No ending.",
        "Add dialogue.",
        "Storytelling")


def s20():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Role Play Theater 🎭", "SPEAKING", "Speaking", 20)
    roles = [
        ("🍽️ Restaurant", "Order food politely"),
        ("🦁 Zoo", "Guide describes animals"),
        ("✈️ Airport", "Check-in conversation"),
        ("📚 Library", "Whisper & ask for a book"),
    ]
    for i, (t, d) in enumerate(roles):
        left = Inches(0.4 + (i % 2) * 6.45)
        top = Inches(1.5 + (i // 2) * 2.5)
        add_round(slide, left, top, Inches(6.2), Inches(2.25), WHITE)
        tb(slide, left + Inches(0.25), top + Inches(0.35), Inches(5.7), Inches(0.45), t, size=18, bold=True, color=PURPLE)
        tb(slide, left + Inches(0.25), top + Inches(1.0), Inches(5.7), Inches(0.8), d, size=15, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Sage picks 1 role; teacher plays partner.",
        "45–60 second role play.",
        "Appropriate phrases (May I…, Could you…).",
        "Pragmatics, confidence.",
        "Too quiet in library scene.",
        "Switch roles quickly.",
        "Speaking block")


def s21():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Pronunciation Challenge: Tongue Twisters 👅", "PRONUNCIATION", "Flex", 21)
    tw = [
        "Red lorry, yellow lorry.",
        "Unique New York.",
        "She sells seashells by the seashore.",
        "Toy boat, toy boat, toy boat.",
    ]
    for i, t in enumerate(tw):
        top = Inches(1.5 + i * 1.25)
        add_round(slide, Inches(0.4), top, Inches(12.5), Inches(1.1), WHITE)
        tb(slide, Inches(0.65), top + Inches(0.35), Inches(12.0), Inches(0.45), t, size=17, bold=True, color=NAVY)
    btn(slide)
    teacher_notes(slide,
        "Slow → medium → fast.",
        "Repeat each 3 times.",
        "Clearer consonants over time.",
        "Pronunciation playfulness.",
        "Frustration—keep it fun.",
        "Clap the rhythm.",
        "Flex within 90 min")


def s22():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Listening Activity 👂", "LISTENING", "Listening", 22)
    add_round(slide, Inches(0.4), Inches(1.4), Inches(12.5), Inches(2.8), LIGHT_SKY)
    tb(slide, Inches(0.65), Inches(1.6), Inches(12.0), Inches(2.4),
       "Teacher reads aloud (do not show student the text first):\n\n" + LISTENING, size=14, color=DARK)
    bullets(slide, Inches(0.65), Inches(4.4), Inches(12.0), Inches(2.2), [
        "1. What event was held?",
        "2. What did Sage's class make?",
        "3. What place did they win?",
        "4. What did they eat?",
    ], size=16, sp=10)
    btn(slide)
    teacher_notes(slide,
        "Read twice; second time Sage may take notes.",
        "Answer 4 questions without looking at passage.",
        "Reading festival; poster; second; fruit popsicles.",
        "Listening comprehension.",
        "Copying irrelevant details.",
        "What was the main idea?",
        "Listening skill")


def s23():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Find the Mistakes ✏️", "GAME", "Grammar", 23)
    lines = [
        "Leo go to school every day.",
        "He are looking for his backpack now.",
        "Yesterday, he will lose his slip.",
        "Tomorrow, Ms. Carter helped him.",
        "The mystery were solved quickly.",
    ]
    for i, ln in enumerate(lines):
        top = Inches(1.45 + i * 1.05)
        add_round(slide, Inches(0.4), top, Inches(12.5), Inches(0.95), LIGHT_YELLOW if i % 2 else WHITE)
        tb(slide, Inches(0.65), top + Inches(0.28), Inches(12.0), Inches(0.45), f"{i + 1}.  {ln}", size=15, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Lightning round—10 sec think time each.",
        "Say corrected sentence.",
        "goes; is looking; lost; will help; was solved",
        "Tense + agreement.",
        "Only naming error type.",
        "Write one mistake for teacher.",
        "Grammar games")


def s24():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Fast Quiz ⚡ Kahoot Style", "GAME", "Games 5 min", 24)
    qs = [
        "Q1: Past of 'find' → found / finded",
        "Q2: 'is running' = Past / Present Continuous",
        "Q3: Synonym of happy → sad / glad",
        "Q4: Leo's backpack color → blue / red",
        "Q5: Future helper → will / did",
    ]
    for i, q in enumerate(qs):
        top = Inches(1.45 + i * 1.05)
        add_round(slide, Inches(0.4), top, Inches(12.5), Inches(0.95), [SKY, TEAL, GOLD, CORAL, PURPLE][i % 5])
        tb(slide, Inches(0.65), top + Inches(0.28), Inches(12.0), Inches(0.45), q, size=16, bold=True, color=WHITE)
    btn(slide)
    teacher_notes(slide,
        "Score with fingers/stars—no stress.",
        "Quick verbal answers.",
        "found; Present Continuous; glad; blue; will",
        "Retention check.",
        "Rushing.",
        "Make up Q6.",
        "Games 5 min")


def s25():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Treasure Hunt: Hidden Vocabulary 🎁", "GAME", "Games", 25)
    tb(slide, Inches(0.4), Inches(1.4), Inches(12.5), Inches(1.0),
       "Clues lead to vocabulary words about today's lesson:", size=16, bold=True, color=NAVY)
    clues = [
        "Clue 1: You carry books in it → BACKPACK",
        "Clue 2: It helps you solve a mystery → CLUE",
        "Clue 3: Not past, not present… → FUTURE",
        "Clue 4: Action word for 'found' → VERB (past tense)",
    ]
    for i, c in enumerate(clues):
        top = Inches(2.6 + i * 0.95)
        add_round(slide, Inches(0.4), top, Inches(12.5), Inches(0.85), WHITE)
        tb(slide, Inches(0.65), top + Inches(0.25), Inches(12.0), Inches(0.45), c, size=15, color=DARK)
    btn(slide)
    teacher_notes(slide,
        "Read clues dramatically.",
        "Shout answers; use in sentence.",
        "See clue answers.",
        "Vocab recall.",
        "Random words.",
        "Sage writes own clue.",
        "Games")


def s26():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Escape Room: 5 English Puzzles 🔐", "GAME", "Games", 26)
    puzzles = [
        "Puzzle 1: Unscramble TENSES → ?",
        "Puzzle 2: Fix → He go to school.",
        "Puzzle 3: Antonym of lost → ?",
        "Puzzle 4: Complete → Leo ___ (find) his bag.",
        "Puzzle 5: Riddle: I am full of words but cannot speak. → ?",
    ]
    for i, p in enumerate(puzzles):
        top = Inches(1.45 + i * 1.05)
        add_round(slide, Inches(0.4), top, Inches(12.5), Inches(0.95), LIGHT_TEAL if i % 2 else WHITE)
        tb(slide, Inches(0.65), top + Inches(0.28), Inches(12.0), Inches(0.45), p, size=15, color=DARK)
    tb(slide, Inches(0.4), Inches(6.35), Inches(12.5), Inches(0.35),
       "KEY: SENT; goes; found; found; library/book", size=12, color=CORAL, align=PP_ALIGN.CENTER)
    btn(slide)
    teacher_notes(slide,
        "Team vs clock—celebrate escape!",
        "Solve all 5.",
        "See key line (hide if needed).",
        "Integrated skills.",
        "Giving up—offer hint.",
        "Design puzzle 6.",
        "Games 5 min")


def s27():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Mini Assessment Snapshot ⭐", "TEACHER", "Assessment", 27)
    skills = ["Reading", "Grammar", "Vocabulary", "Speaking", "Confidence", "Pronunciation"]
    add_rect(slide, Inches(0.4), Inches(1.4), Inches(12.5), Inches(0.5), NAVY)
    tb(slide, Inches(0.6), Inches(1.48), Inches(3), Inches(0.35), "Skill", size=12, bold=True, color=WHITE)
    tb(slide, Inches(4.0), Inches(1.48), Inches(8.5), Inches(0.35), "⭐ ⭐ ⭐ ⭐ ⭐  (circle one)", size=12, bold=True,
       color=WHITE)
    for i, sk in enumerate(skills):
        top = Inches(2.0 + i * 0.75)
        add_rect(slide, Inches(0.4), top, Inches(12.5), Inches(0.75), WHITE if i % 2 == 0 else LIGHT_SKY)
        tb(slide, Inches(0.6), top + Inches(0.2), Inches(3), Inches(0.4), sk, size=14, bold=True, color=NAVY)
        tb(slide, Inches(4.0), top + Inches(0.2), Inches(8.5), Inches(0.4), "○ ○ ○ ○ ○", size=14, color=SOFT)
    btn(slide)
    teacher_notes(slide,
        "Complete privately—never make Sage feel graded.",
        "Teacher observes and marks rubric.",
        "1=emerging 5=strong for Grade 5 demo.",
        "Holistic demo assessment for parents.",
        "Sharing scores aloud.",
        "Write 1 strength + 1 next step.",
        "Throughout class")


def s28():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Recap: You Teach the Teacher! 🎓", "RECAP", "Recap 5 min", 28)
    bullets(slide, Inches(0.5), Inches(1.5), Inches(12.3), Inches(4.5), [
        "Teach me one tense rule.",
        "Teach me 3 new words.",
        "Teach me what happened in the backpack mystery.",
        "What did we learn today?",
        "What did you enjoy MOST?",
    ], size=18, sp=14)
    btn(slide)
    teacher_notes(slide,
        "Switch roles—Sage is the expert!",
        "2-minute mini-lesson from student.",
        "Accurate recap.",
        "Metacognition + confidence.",
        "I don't know.",
        "What will you practice tomorrow?",
        "Recap 5 min")


def s29():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide_header(slide, "Homework Mission 📝", "HOMEWORK", "Home", 29)
    add_round(slide, Inches(1.2), Inches(1.45), Inches(10.9), Inches(5.2), WHITE)
    tb(slide, Inches(1.5), Inches(1.75), Inches(10.3), Inches(0.45),
       "Write: \"My Dream Vacation\"", size=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    bullets(slide, Inches(1.5), Inches(2.5), Inches(10.3), Inches(3.8), [
        "Use 5 adjectives",
        "Use 5 verbs",
        "Use Present, Past, AND Future tense",
        "Draw a picture (optional bonus)",
        "Read aloud to a family member",
    ], size=17, sp=12)
    btn(slide)
    teacher_notes(slide,
        "Keep homework fun—15–20 min max.",
        "Plan orally before writing.",
        "Paragraph with tense variety.",
        "Transfer to independent work.",
        "All one tense.",
        "Where would you go first?",
        "After class")


def s30():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), YELLOW)
    add_round(slide, Inches(2.0), Inches(1.1), Inches(9.3), Inches(5.3), WHITE)
    add_rect(slide, Inches(2.0), Inches(1.1), Inches(9.3), Inches(0.25), TEAL)
    add_rect(slide, Inches(2.0), Inches(6.15), Inches(9.3), Inches(0.25), GOLD)
    tb(slide, Inches(2.3), Inches(1.6), Inches(8.7), Inches(0.5),
       "🎉 Congratulations Sage! 🎉", size=32, bold=True, color=NAVY, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(2.3), Inches(2.5), Inches(8.7), Inches(1.5),
       "Great Job!\nYou completed the 90-Minute English Adventure Quest.",
       size=20, color=DARK, align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.3), Inches(4.3), Inches(8.7), Inches(0.5),
       "Reading • Grammar • Vocabulary • Speaking • Listening • Creativity",
       size=15, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.3), Inches(5.0), Inches(8.7), Inches(0.45),
       "Certificate of Awesome Effort — Grade 5 Demo", size=14, color=SOFT, align=PP_ALIGN.CENTER)
    footer(slide, 30, "Celebrate!")
    fade_transition(slide)
    teacher_notes(slide,
        "Big celebration! Specific praise (2 strengths).",
        "Sage reads certificate title aloud.",
        "Smiles, pride, willingness to return.",
        "End on high energy for parents.",
        "Rushing out.",
        "What game do you want next time?",
        "End")


# Build all slides
for fn in [s01, s02, s03, s04, s05, s06, s07, s08, s09, s10, s11, s12, s13, s14, s15, s16, s17, s18, s19, s20,
           s21, s22, s23, s24, s25, s26, s27, s28, s29, s30]:
    fn()

out = r"C:\Users\bhushaja\Downloads\shaip\Grade5_Sage_Premium_English_Demo_90min.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
