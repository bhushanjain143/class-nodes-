"""Grade 5 English Demo Class PPT for Sage — 90 minutes, professional & engaging."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

# Blue–green–yellow palette
NAVY = RGBColor(0x0B, 0x3D, 0x91)
SKY = RGBColor(0x1E, 0x88, 0xE5)
TEAL = RGBColor(0x00, 0x96, 0x88)
MINT = RGBColor(0x26, 0xA6, 0x9A)
YELLOW = RGBColor(0xFF, 0xC1, 0x07)
GOLD = RGBColor(0xFF, 0xB3, 0x00)
CORAL = RGBColor(0xFF, 0x70, 0x43)
PINK = RGBColor(0xEC, 0x40, 0x7A)
PURPLE = RGBColor(0x7E, 0x57, 0xC2)
CREAM = RGBColor(0xFF, 0xFB, 0xF0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1A, 0x23, 0x37)
SOFT = RGBColor(0x54, 0x6E, 0x7A)
LIGHT_SKY = RGBColor(0xE3, 0xF2, 0xFD)
LIGHT_TEAL = RGBColor(0xE0, 0xF2, 0xF1)
LIGHT_YELLOW = RGBColor(0xFF, 0xF8, 0xE1)
LIGHT_CORAL = RGBColor(0xFF, 0xEB, 0xEE)
LIGHT_PURPLE = RGBColor(0xED, 0xE7, 0xF6)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
TOTAL = 25

PASSAGE = (
    "Mia loved walking through Maple Park every Saturday morning. The trees were tall, "
    "the birds sang cheerfully, and the soft grass felt cool under her sneakers.\n\n"
    "One sunny day, Mia heard a tiny whimper near the picnic benches. She stopped and looked around. "
    "Behind a bright red backpack sat a small brown puppy with muddy paws and sad eyes. "
    "A blue collar hung loosely around its neck, but there was no name tag.\n\n"
    "\"Don't worry, little friend,\" Mia whispered. \"I will help you.\" She gently offered her water bottle. "
    "The puppy drank quickly, then wagged its tail. Mia asked nearby families if they had lost a pet. "
    "Nobody knew the puppy.\n\n"
    "Then Mia noticed a poster on the park board: \"Lost Puppy — Brown, Friendly, Answers to Coco.\" "
    "There was a phone number at the bottom. Mia carefully dialed the number with her mom's help. "
    "Ten minutes later, a smiling girl named Ava ran toward them, tears of joy on her face.\n\n"
    "\"You found Coco!\" Ava cried. Mia felt proud and happy. That day she learned that kindness "
    "can turn an ordinary walk into an unforgettable adventure."
)


def set_run(run, size=18, bold=False, color=DARK, font="Calibri"):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font


def add_bg(slide, color):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()
    spTree = slide.shapes._spTree
    sp = shape._element
    spTree.remove(sp)
    spTree.insert(2, sp)


def add_rect(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def add_round(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def add_oval(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def tb(slide, left, top, width, height, text, size=18, bold=False, color=DARK,
       align=PP_ALIGN.LEFT, font="Calibri"):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    set_run(run, size=size, bold=bold, color=color, font=font)
    return box


def multi_tb(slide, left, top, width, height, lines, size=15, color=DARK, spacing=6, bold=False):
    """lines: list of str or (str, bold, color)"""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(spacing)
        if isinstance(line, tuple):
            text, is_bold, col = line
        else:
            text, is_bold, col = line, bold, color
        run = p.add_run()
        run.text = text
        set_run(run, size=size, bold=is_bold, color=col)
    return box


def bullets(slide, left, top, width, height, items, size=15, color=DARK, spacing=8):
    return multi_tb(
        slide, left, top, width, height,
        [("•  " + i, False, color) for i in items],
        size=size, spacing=spacing,
    )


def notes(slide, text):
    # Speaker notes intentionally omitted from the deliverable PPT.
    return


def add_transition(slide, transition="fade"):
    """Add a simple slide transition via XML."""
    nsmap_p = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    sld = slide._element
    # Remove existing transition if any
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    if transition == "fade":
        etree.SubElement(tr, qn("p:fade"))
    elif transition == "push":
        etree.SubElement(tr, qn("p:push"))
    else:
        etree.SubElement(tr, qn("p:fade"))


def footer(slide, page, timing=""):
    add_rect(slide, Inches(0), Inches(7.15), prs.slide_width, Inches(0.35), NAVY)
    label = "Grade 5 English Demo  |  Sage  |  90 Minutes"
    if timing:
        label += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.18), Inches(11.2), Inches(0.3), label, size=11, color=WHITE)
    tb(slide, Inches(12.0), Inches(7.18), Inches(1.1), Inches(0.3),
       f"{page}/{TOTAL}", size=11, color=WHITE, align=PP_ALIGN.RIGHT)


def chip(slide, text, left=Inches(0.4), top=Inches(0.28), color=SKY, w=Inches(2.8)):
    add_round(slide, left, top, w, Inches(0.36), color)
    tb(slide, left, top + Inches(0.01), w, Inches(0.34),
       text, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def time_chip(slide, text):
    chip(slide, text, left=Inches(10.5), top=Inches(0.28), color=CORAL, w=Inches(2.4))


def icon_bubble(slide, left, top, emoji, bg):
    add_oval(slide, left, top, Inches(0.85), Inches(0.85), bg)
    tb(slide, left, top + Inches(0.18), Inches(0.85), Inches(0.55),
       emoji, size=22, align=PP_ALIGN.CENTER)


# ===================== SLIDES =====================

def s01_welcome():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.2), YELLOW)
    add_oval(slide, Inches(10.8), Inches(-1.0), Inches(3.8), Inches(3.8), RGBColor(0x15, 0x4F, 0xB0))
    add_oval(slide, Inches(-1.2), Inches(5.2), Inches(3.2), Inches(3.2), RGBColor(0x12, 0x48, 0xA0))
    # decorative icons via shapes + emoji
    for i, (x, y, e, c) in enumerate([
        (0.7, 1.1, "📚", TEAL), (2.0, 0.9, "✏️", YELLOW), (11.2, 1.3, "💬", CORAL),
        (12.2, 2.4, "⭐", GOLD), (1.2, 5.5, "🌈", PINK),
    ]):
        icon_bubble(slide, Inches(x), Inches(y), e, c)

    tb(slide, Inches(0.8), Inches(2.2), Inches(11.7), Inches(0.4),
       "HELLO, SAGE!", size=18, bold=True, color=YELLOW, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(2.7), Inches(11.7), Inches(1.2),
       "Welcome to Your\nEnglish Adventure!", size=40, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(4.5), Inches(11.7), Inches(0.4),
       "Grade 5 English Demo Class  •  90 Minutes", size=18, color=RGBColor(0xBB, 0xDE, 0xFB),
       align=PP_ALIGN.CENTER)
    add_rect(slide, Inches(0), Inches(5.7), prs.slide_width, Inches(1.8), TEAL)
    tb(slide, Inches(0.8), Inches(6.0), Inches(11.7), Inches(0.4),
       "Reading  •  Speaking  •  Grammar  •  Vocabulary  •  Pronunciation  •  Fun!",
       size=16, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(6.55), Inches(11.7), Inches(0.35),
       "Let's learn, play, and shine together ✨", size=15, color=LIGHT_YELLOW, align=PP_ALIGN.CENTER)
    add_transition(slide)
    notes(slide,
          "SAY: \"Hi Sage! Welcome to your English adventure. I'm so happy you're here today!\"\n"
          "Smile, wave, and keep energy high.\n"
          "EXPECTED: Smile, greeting, maybe shy \"hello.\"\n"
          "FOLLOW-UP: \"Are you ready for some fun English games?\"\n"
          "ASSESS: First impression of confidence and willingness to speak.\n"
          "TIMING: 1–2 minutes.")


def s02_teacher():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, SKY)
    chip(slide, "LET'S MEET!", color=SKY)
    time_chip(slide, "Intro • 5 min")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.45),
       "Teacher Introduction", size=28, bold=True, color=NAVY, font="Georgia")

    # Teacher card
    add_round(slide, Inches(0.4), Inches(1.5), Inches(6.1), Inches(5.1), WHITE)
    icon_bubble(slide, Inches(0.7), Inches(1.8), "👩‍🏫", LIGHT_SKY)
    tb(slide, Inches(1.7), Inches(1.95), Inches(4.5), Inches(0.4),
       "Your Teacher Today", size=18, bold=True, color=NAVY)
    bullets(slide, Inches(0.7), Inches(2.9), Inches(5.5), Inches(3.3), [
        "Name: (Say your name here)",
        "Fun fact: I love turning English into games!",
        "Favorite book: Charlotte's Web",
        "Goal today: Help Sage shine in English",
    ], size=16)

    # Student intro card
    add_round(slide, Inches(6.8), Inches(1.5), Inches(6.1), Inches(5.1), LIGHT_YELLOW)
    icon_bubble(slide, Inches(7.1), Inches(1.8), "🌟", YELLOW)
    tb(slide, Inches(8.1), Inches(1.95), Inches(4.5), Inches(0.4),
       "Now It's Your Turn, Sage!", size=18, bold=True, color=NAVY)
    bullets(slide, Inches(7.1), Inches(2.9), Inches(5.5), Inches(3.3), [
        "What's your name?",
        "Which school do you study in?",
        "What's your favorite subject?",
        "What's your favorite hobby?",
    ], size=16)
    footer(slide, 2, "0–5 min")
    add_transition(slide)
    notes(slide,
          "SAY: Introduce yourself warmly. Share name, fun fact, favorite book.\n"
          "Then point to Sage's card: \"Now tell me about you!\"\n"
          "ASK one question at a time. Wait patiently.\n"
          "EXPECTED: Short answers; encourage full sentences: \"I study at…\"\n"
          "FOLLOW-UP: \"That's awesome! Tell me one more thing about your hobby.\"\n"
          "ASSESS: Pronunciation, fluency, confidence, listening.\n"
          "TIMING: 5 minutes.")


def s03_icebreaker():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, GOLD)
    chip(slide, "ICEBREAKER GAME", color=GOLD, w=Inches(3.2))
    time_chip(slide, "10 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Would You Rather…? 🎲", size=26, bold=True, color=NAVY, font="Georgia")
    tb(slide, Inches(0.4), Inches(1.25), Inches(12), Inches(0.3),
       "Choose one side and speak in a FULL sentence: \"I would rather ___ because ___\"",
       size=14, color=SOFT)

    qs = [
        ("🐶", "Have a pet dragon", "🦄", "Ride a unicorn", TEAL, PURPLE),
        ("📚", "Read all day", "🏞️", "Explore a forest", SKY, TEAL),
        ("🍕", "Eat pizza for breakfast", "🥞", "Eat pancakes for dinner", CORAL, GOLD),
        ("🚀", "Visit the Moon", "🌊", "Live under the sea", NAVY, SKY),
        ("🎤", "Be a singer", "🎨", "Be an artist", PINK, PURPLE),
        ("🦸", "Have super speed", "🧙", "Have magic powers", CORAL, TEAL),
        ("🏫", "Teach your class", "🎬", "Star in a movie", SKY, GOLD),
        ("🌈", "Make it rain candy", "❄️", "Make it snow in summer", YELLOW, SKY),
    ]
    for i, (e1, a, e2, b, c1, c2) in enumerate(qs):
        col, row = i % 4, i // 4
        left = Inches(0.35 + col * 3.25)
        top = Inches(1.7 + row * 2.55)
        add_round(slide, left, top, Inches(3.1), Inches(2.35), WHITE)
        tb(slide, left + Inches(0.1), top + Inches(0.12), Inches(2.9), Inches(0.3),
           f"Q{i + 1}", size=12, bold=True, color=c1, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(0.45), Inches(2.9), Inches(0.7),
           f"{e1} {a}", size=13, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(1.15), Inches(2.9), Inches(0.3),
           "OR", size=12, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(1.5), Inches(2.9), Inches(0.7),
           f"{e2} {b}", size=13, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    footer(slide, 3, "5–15 min")
    add_transition(slide)
    notes(slide,
          "SAY: \"This is Would You Rather! Pick one option and tell me WHY in a full sentence.\"\n"
          "Model first: \"I would rather ride a unicorn because it looks magical!\"\n"
          "Do 5–8 questions depending on energy. Keep it playful.\n"
          "EXPECTED: Full sentences with because.\n"
          "FOLLOW-UP: \"Interesting! Can you add one more reason?\"\n"
          "ASSESS: Speaking fluency, vocabulary, confidence, sentence structure.\n"
          "TIMING: 10 minutes.")


def s04_warmup():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, TEAL)
    chip(slide, "WARM-UP TALK", color=TEAL)
    time_chip(slide, "5 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.45),
       "Warm-up Conversation 💬", size=28, bold=True, color=NAVY, font="Georgia")

    prompts = [
        ("👫", TEAL, "Best Friend", "Tell me about your best friend."),
        ("🗓️", SKY, "Last Weekend", "What did you do last weekend?"),
        ("⚡", PURPLE, "Superpower", "If you had one superpower, what would it be?"),
    ]
    for i, (e, c, t, q) in enumerate(prompts):
        left = Inches(0.4 + i * 4.25)
        add_round(slide, left, Inches(1.6), Inches(4.05), Inches(4.8), WHITE)
        add_rect(slide, left, Inches(1.6), Inches(4.05), Inches(1.1), c)
        tb(slide, left, Inches(1.75), Inches(4.05), Inches(0.4),
           e + "  " + t, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(3.1), Inches(3.55), Inches(2.5),
           q, size=18, color=DARK, align=PP_ALIGN.CENTER)
    footer(slide, 4, "15–20 min")
    add_transition(slide)
    notes(slide,
          "SAY each prompt slowly. Give Sage 20–30 seconds think time.\n"
          "Encourage 3–5 sentences, not one-word answers.\n"
          "EXPECTED: Past tense on weekend; imaginative language on superpower.\n"
          "FOLLOW-UP: \"What happened next?\" \"How did that make you feel?\"\n"
          "ASSESS: Fluency, pronunciation, vocabulary range, confidence.\n"
          "TIMING: 5 minutes (part of intro block).")


def s05_passage():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, SKY)
    chip(slide, "READING", color=SKY)
    time_chip(slide, "Reading • 20 min")
    tb(slide, Inches(0.4), Inches(0.75), Inches(12), Inches(0.4),
       "📖  The Lost Puppy", size=26, bold=True, color=NAVY, font="Georgia")
    add_round(slide, Inches(0.4), Inches(1.25), Inches(12.5), Inches(5.55), WHITE)
    # left illustration panel
    add_round(slide, Inches(0.6), Inches(1.45), Inches(2.4), Inches(5.15), LIGHT_SKY)
    tb(slide, Inches(0.7), Inches(2.2), Inches(2.2), Inches(0.5), "🐶", size=40, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(3.0), Inches(2.2), Inches(0.4), "Coco", size=16, bold=True,
       color=NAVY, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(3.5), Inches(2.2), Inches(1.5),
       "park • puppy\nkindness • adventure", size=13, color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(3.2), Inches(1.5), Inches(9.4), Inches(5.1), PASSAGE, size=13, color=DARK)
    footer(slide, 5, "20–40 min")
    add_transition(slide)
    notes(slide,
          "SAY: \"Today's story is The Lost Puppy. First I'll read; then you can read parts.\"\n"
          "Teacher reads with expression. Then Sage reads 1–2 paragraphs.\n"
          "Help with tricky words gently (cheerfully, whimper, dialed).\n"
          "EXPECTED: Follow along; ask about unknown words.\n"
          "FOLLOW-UP: \"Which part did you like most?\"\n"
          "ASSESS: Decoding, fluency, expression, comprehension while listening.\n"
          "TIMING: ~8–10 minutes of the 20-min reading block.")


def s06_questions():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, SKY)
    chip(slide, "COMPREHENSION")
    time_chip(slide, "~8 minutes")
    tb(slide, Inches(0.4), Inches(0.75), Inches(12), Inches(0.4),
       "Reading Questions — The Lost Puppy", size=24, bold=True, color=NAVY, font="Georgia")

    qs = [
        ("MCQ", TEAL, "1. Where did Mia walk every Saturday?\nA) Beach  B) Maple Park  C) School  D) Zoo"),
        ("MCQ", TEAL, "2. What color was the puppy?\nA) Black  B) White  C) Brown  D) Gray"),
        ("T/F", GOLD, "3. True or False: The puppy had a name tag."),
        ("T/F", GOLD, "4. True or False: Ava was happy to see Coco again."),
        ("Short", SKY, "5. What did Mia offer the puppy first?"),
        ("Short", SKY, "6. How did Mia find the owner's phone number?"),
        ("Infer", PURPLE, "7. Why do you think Mia felt proud at the end?"),
        ("Vocab", CORAL, "8. What does \"whimper\" most likely mean?\nA) Loud bark  B) Soft sad sound  C) Happy laugh"),
    ]
    for i, (tag, c, q) in enumerate(qs):
        col, row = i % 2, i // 2
        left = Inches(0.35 + col * 6.45)
        top = Inches(1.25 + row * 1.35)
        add_round(slide, left, top, Inches(6.25), Inches(1.25), WHITE)
        add_round(slide, left + Inches(0.12), top + Inches(0.15), Inches(0.9), Inches(0.35), c)
        tb(slide, left + Inches(0.12), top + Inches(0.17), Inches(0.9), Inches(0.3),
           tag, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.15), top + Inches(0.2), Inches(4.9), Inches(0.95), q, size=12, color=DARK)
    footer(slide, 6, "20–40 min")
    add_transition(slide)
    notes(slide,
          "Ask questions one by one. Celebrate evidence: \"Where in the story did you find that?\"\n"
          "EXPECTED ANSWERS (preview): B; C; False; True; water; poster; kindness/helping; B.\n"
          "FOLLOW-UP for #7: \"What clue in the text shows her pride?\"\n"
          "ASSESS: Literal vs inferential understanding; vocabulary in context.\n"
          "Full answer key on Slide 23.\n"
          "TIMING: ~8 minutes.")


def s07_speaking():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, PINK)
    chip(slide, "SPEAKING", color=PINK)
    time_chip(slide, "~5 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Retell the Story in Your Own Words 🗣️", size=26, bold=True, color=NAVY, font="Georgia")

    prompts = [
        ("1️⃣", "Beginning", "Who is Mia and where is she?"),
        ("2️⃣", "Middle", "What problem does she find?"),
        ("3️⃣", "Turning Point", "How does she solve it?"),
        ("4️⃣", "Ending", "How does everyone feel?"),
    ]
    for i, (n, t, q) in enumerate(prompts):
        left = Inches(0.4 + i * 3.2)
        add_round(slide, left, Inches(1.6), Inches(3.05), Inches(4.7), WHITE)
        add_rect(slide, left, Inches(1.6), Inches(3.05), Inches(1.2), [TEAL, SKY, PURPLE, CORAL][i])
        tb(slide, left, Inches(1.75), Inches(3.05), Inches(0.4), n, size=20, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.2), Inches(3.05), Inches(0.4), t, size=16, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.3), Inches(2.65), Inches(2.5),
           q, size=15, color=DARK, align=PP_ALIGN.CENTER)
    footer(slide, 7, "20–40 min")
    add_transition(slide)
    notes(slide,
          "SAY: \"Pretend I'm a friend who hasn't heard the story. Retell it using these four steps.\"\n"
          "Allow notes if needed; discourage reading the passage word-for-word.\n"
          "EXPECTED: Clear sequence; past tense; key details (Coco, poster, Ava).\n"
          "FOLLOW-UP: \"Can you add how Mia felt when she first saw the puppy?\"\n"
          "ASSESS: Oral organization, vocabulary, grammar in speech, confidence.\n"
          "TIMING: ~5 minutes (end of reading block).")


def s08_grammar():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, TEAL)
    chip(slide, "GRAMMAR", color=TEAL)
    time_chip(slide, "Grammar • 20 min")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Parts of Speech Power-Up ⚡", size=26, bold=True, color=NAVY, font="Georgia")

    parts = [
        ("Nouns", "📚", SKY, "Naming words\nperson, place, thing", "Mia, park, puppy"),
        ("Verbs", "🏃", TEAL, "Action / being words", "walked, heard, felt"),
        ("Adjectives", "🎨", GOLD, "Describe nouns", "tiny, brown, proud"),
        ("Adverbs", "🚀", PURPLE, "Describe verbs\nhow / when / where", "cheerfully, gently, quickly"),
    ]
    for i, (t, e, c, d, ex) in enumerate(parts):
        left = Inches(0.35 + i * 3.25)
        add_round(slide, left, Inches(1.5), Inches(3.1), Inches(5.0), WHITE)
        add_rect(slide, left, Inches(1.5), Inches(3.1), Inches(1.3), c)
        tb(slide, left, Inches(1.6), Inches(3.1), Inches(0.5), e, size=28, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.2), Inches(3.1), Inches(0.45), t, size=18, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.2), Inches(2.7), Inches(1.5),
           d, size=14, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(5.0), Inches(2.6), Inches(1.1), LIGHT_TEAL if i % 2 == 0 else LIGHT_YELLOW)
        tb(slide, left + Inches(0.3), Inches(5.2), Inches(2.5), Inches(0.8),
           "Ex: " + ex, size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
    footer(slide, 8, "40–60 min")
    add_transition(slide)
    notes(slide,
          "Teach each part with finger gestures (point for noun, act for verb, paint for adjective, speed for adverb).\n"
          "Use story examples so grammar connects to reading.\n"
          "EXPECTED: Sage can give one new example for each.\n"
          "FOLLOW-UP: \"Is 'quickly' an adjective or adverb? Why?\"\n"
          "ASSESS: Concept understanding before practice.\n"
          "TIMING: ~6 minutes.")


def s09_grammar_practice():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, TEAL)
    chip(slide, "PRACTICE", color=TEAL)
    time_chip(slide, "~7 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Grammar Practice — Find It! 🔍", size=26, bold=True, color=NAVY, font="Georgia")
    tb(slide, Inches(0.4), Inches(1.25), Inches(12), Inches(0.3),
       "Teacher-led interactive: Sage points / says the answer (drag-and-drop style).",
       size=13, color=SOFT)

    tasks = [
        ("Circle the VERB", CORAL, "Mia walked through Maple Park."),
        ("Underline the ADJECTIVE", GOLD, "A tiny brown puppy sat nearby."),
        ("Choose the NOUN", SKY, "The birds sang cheerfully.\nNoun = birds / sang / cheerfully?"),
        ("Find the ADVERB", PURPLE, "Mia gently offered water."),
        ("Spot ALL parts", TEAL, "Happy Ava quickly hugged Coco.\n(Adj + Noun + Adv + Verb + Noun)"),
        ("Make one sentence", PINK, "Use: park (noun) + ran (verb)\n+ soft (adjective)"),
    ]
    for i, (t, c, s) in enumerate(tasks):
        col, row = i % 3, i // 3
        left = Inches(0.35 + col * 4.3)
        top = Inches(1.7 + row * 2.5)
        add_round(slide, left, top, Inches(4.1), Inches(2.3), WHITE)
        add_rect(slide, left, top, Inches(4.1), Inches(0.55), c)
        tb(slide, left + Inches(0.15), top + Inches(0.12), Inches(3.8), Inches(0.35),
           t, size=14, bold=True, color=WHITE)
        tb(slide, left + Inches(0.2), top + Inches(0.8), Inches(3.7), Inches(1.3),
           s, size=14, color=DARK)
    footer(slide, 9, "40–60 min")
    add_transition(slide)
    notes(slide,
          "Treat like live drag-and-drop: Sage says or points to the word.\n"
          "ANSWERS: walked; tiny/brown; birds; gently; Happy(adj) Ava(n) quickly(adv) hugged(v) Coco(n).\n"
          "For task 6, praise creativity.\n"
          "FOLLOW-UP: \"Can you change the adjective to make a new meaning?\"\n"
          "ASSESS: Accuracy + reasoning.\n"
          "TIMING: ~7 minutes.")


def s10_grammar_quiz():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, CORAL)
    chip(slide, "QUIZ TIME", color=CORAL)
    time_chip(slide, "~7 minutes")
    tb(slide, Inches(0.4), Inches(0.75), Inches(12), Inches(0.4),
       "Grammar Challenge — 10 Quick Questions 🏁", size=24, bold=True, color=NAVY, font="Georgia")

    quiz = [
        "1. \"Puppy\" is a…  A) Verb  B) Noun  C) Adverb",
        "2. In \"birds sang,\" sang is a…  A) Noun  B) Adjective  C) Verb",
        "3. \"Tiny\" describes…  A) Verb  B) Noun  C) Adverb",
        "4. \"Cheerfully\" is an…  A) Adverb  B) Noun  C) Adjective",
        "5. Best adjective for Coco:  A) run  B) muddy  C) quickly",
        "6. \"Mia\" is a…  A) Proper noun  B) Verb  C) Adverb",
        "7. Which is a verb?  A) park  B) dialed  C) brown",
        "8. \"Proud\" is an…  A) Adjective  B) Verb  C) Noun",
        "9. \"Quickly\" tells us…  A) What  B) How  C) Who",
        "10. \"Adventure\" is a…  A) Noun  B) Adverb  C) Verb",
    ]
    for i, q in enumerate(quiz):
        col, row = i % 2, i // 2
        left = Inches(0.35 + col * 6.45)
        top = Inches(1.25 + row * 1.05)
        add_round(slide, left, top, Inches(6.25), Inches(0.95), WHITE if i % 2 == 0 else LIGHT_YELLOW)
        tb(slide, left + Inches(0.2), top + Inches(0.25), Inches(5.9), Inches(0.5), q, size=13, color=DARK)
    footer(slide, 10, "40–60 min")
    add_transition(slide)
    notes(slide,
          "Click through one question at a time verbally. Celebrate streaks.\n"
          "ANSWERS: 1B 2C 3B 4A 5B 6A 7B 8A 9B 10A (also on Slide 23).\n"
          "If wrong, teach the \"why\" in 10 seconds, then move on.\n"
          "ASSESS: Speed + accuracy under light pressure.\n"
          "TIMING: ~7 minutes.")


def s11_vocab():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, TEAL)
    chip(slide, "VOCABULARY", color=TEAL)
    time_chip(slide, "Vocab • 15 min")
    tb(slide, Inches(0.4), Inches(0.75), Inches(12), Inches(0.35),
       "Nature Word Garden 🌿 — 12 New Words", size=24, bold=True, color=NAVY, font="Georgia")

    words = [
        ("🌳", "forest", "a large area of trees", "We hiked in the forest."),
        ("🌊", "river", "flowing water", "Fish swim in the river."),
        ("☀️", "sunshine", "light from the sun", "Sunshine warms the grass."),
        ("🍃", "breeze", "a soft wind", "A cool breeze felt nice."),
        ("🌼", "bloom", "to open as a flower", "Roses bloom in spring."),
        ("🪨", "boulder", "a very large rock", "A boulder blocked the path."),
        ("🦅", "soar", "to fly high", "Eagles soar above hills."),
        ("🌧️", "drizzle", "light rain", "A drizzle wet our coats."),
        ("🌱", "sprout", "a young plant", "Tiny sprouts pushed up."),
        ("🏞️", "valley", "low land between hills", "Our tent is in the valley."),
        ("🌫️", "mist", "thin fog", "Morning mist hid the lake."),
        ("🧭", "trail", "a path outdoors", "Follow the trail carefully."),
    ]
    for i, (e, w, m, ex) in enumerate(words):
        col, row = i % 4, i // 4
        left = Inches(0.3 + col * 3.25)
        top = Inches(1.2 + row * 1.85)
        add_round(slide, left, top, Inches(3.1), Inches(1.7), WHITE)
        tb(slide, left + Inches(0.1), top + Inches(0.1), Inches(2.9), Inches(0.35),
           f"{e} {w}", size=14, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.5), Inches(2.8), Inches(0.45),
           m, size=12, color=DARK, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.05), Inches(2.8), Inches(0.5),
           ex, size=11, color=SOFT, align=PP_ALIGN.CENTER)
    footer(slide, 11, "60–75 min")
    add_transition(slide)
    notes(slide,
          "Teach 8–12 words based on time. Say → student repeats → meaning → example.\n"
          "Act out soar, breeze, drizzle for memory.\n"
          "EXPECTED: Correct repetition; attempt own sentence for 2–3 words.\n"
          "FOLLOW-UP: \"Which word fits: soft wind?\"\n"
          "ASSESS: Pronunciation + meaning retention.\n"
          "TIMING: ~7 minutes.")


def s12_matching():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, GOLD)
    chip(slide, "MATCHING GAME", color=GOLD, w=Inches(3.2))
    time_chip(slide, "~4 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Vocabulary Matching Game 🎯", size=26, bold=True, color=NAVY, font="Georgia")
    tb(slide, Inches(0.4), Inches(1.25), Inches(12), Inches(0.3),
       "Match the word (left) with the correct meaning (right).", size=14, color=SOFT)

    left_words = ["1. breeze", "2. soar", "3. drizzle", "4. trail", "5. bloom", "6. mist"]
    right_mean = ["A. light rain", "B. soft wind", "C. thin fog", "D. fly high", "E. outdoor path", "F. flower opens"]
    add_round(slide, Inches(0.5), Inches(1.7), Inches(5.8), Inches(4.8), WHITE)
    add_round(slide, Inches(7.0), Inches(1.7), Inches(5.8), Inches(4.8), LIGHT_YELLOW)
    tb(slide, Inches(0.7), Inches(1.9), Inches(5.4), Inches(0.4), "WORDS", size=16, bold=True, color=TEAL)
    tb(slide, Inches(7.2), Inches(1.9), Inches(5.4), Inches(0.4), "MEANINGS (shuffled)", size=16, bold=True, color=GOLD)
    bullets(slide, Inches(0.7), Inches(2.5), Inches(5.4), Inches(3.7), left_words, size=18, spacing=14)
    bullets(slide, Inches(7.2), Inches(2.5), Inches(5.4), Inches(3.7), right_mean, size=18, spacing=14)
    footer(slide, 12, "60–75 min")
    add_transition(slide)
    notes(slide,
          "Sage draws imaginary lines or says \"1 goes with B.\"\n"
          "KEY: 1-B, 2-D, 3-A, 4-E, 5-F, 6-C.\n"
          "Keep score for fun (not pressure).\n"
          "ASSESS: Word-meaning mapping speed.\n"
          "TIMING: ~4 minutes.")


def s13_syn_ant():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, PURPLE)
    chip(slide, "WORD POWER", color=PURPLE)
    time_chip(slide, "~4 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Synonyms & Antonyms Challenge 🔄", size=26, bold=True, color=NAVY, font="Georgia")

    add_round(slide, Inches(0.4), Inches(1.4), Inches(6.2), Inches(5.1), WHITE)
    tb(slide, Inches(0.6), Inches(1.6), Inches(5.8), Inches(0.4),
       "😊 Synonyms (same / similar)", size=16, bold=True, color=TEAL)
    bullets(slide, Inches(0.6), Inches(2.2), Inches(5.8), Inches(4.0), [
        "happy → glad / joyful",
        "tiny → small / little",
        "fast → quick / speedy",
        "brave → courageous",
        "Your turn: big → ?",
    ], size=16, spacing=12)

    add_round(slide, Inches(6.9), Inches(1.4), Inches(6.0), Inches(5.1), WHITE)
    tb(slide, Inches(7.1), Inches(1.6), Inches(5.6), Inches(0.4),
       "⚡ Antonyms (opposites)", size=16, bold=True, color=CORAL)
    bullets(slide, Inches(7.1), Inches(2.2), Inches(5.6), Inches(4.0), [
        "happy → sad",
        "tiny → huge / large",
        "fast → slow",
        "lost → found",
        "Your turn: kind → ?",
    ], size=16, spacing=12)
    footer(slide, 13, "60–75 min")
    add_transition(slide)
    notes(slide,
          "Explain synonym = similar, antonym = opposite with hand gestures.\n"
          "EXPECTED: big→large/huge; kind→mean/unkind/cruel.\n"
          "Connect to passage: tiny/sad/proud.\n"
          "ASSESS: Lexical flexibility.\n"
          "TIMING: ~4 minutes (end vocab block).")


def s14_storytelling():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, CORAL)
    chip(slide, "STORYTELLING", color=CORAL)
    time_chip(slide, "Story • 10 min")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Create Your Own Story ✨", size=26, bold=True, color=NAVY, font="Georgia")

    # Picture scene made of shapes
    add_round(slide, Inches(0.4), Inches(1.4), Inches(7.3), Inches(5.1), LIGHT_SKY)
    add_oval(slide, Inches(5.5), Inches(1.7), Inches(1.3), Inches(1.3), YELLOW)  # sun
    add_rect(slide, Inches(0.4), Inches(4.6), Inches(7.3), Inches(1.9), TEAL)  # ground
    add_oval(slide, Inches(1.2), Inches(3.6), Inches(1.8), Inches(1.2), RGBColor(0x66, 0xBB, 0x6A))
    add_oval(slide, Inches(2.8), Inches(3.3), Inches(2.2), Inches(1.5), RGBColor(0x43, 0xA0, 0x47))
    add_round(slide, Inches(5.0), Inches(3.8), Inches(2.0), Inches(1.5), RGBColor(0x8D, 0x6E, 0x63))
    tb(slide, Inches(5.2), Inches(4.2), Inches(1.6), Inches(0.8), "🏠", size=28, align=PP_ALIGN.CENTER)
    tb(slide, Inches(1.5), Inches(5.1), Inches(5), Inches(0.5),
       "👧  🐶  🌳  Mystery backpack!", size=18, color=WHITE, align=PP_ALIGN.CENTER)

    add_round(slide, Inches(7.95), Inches(1.4), Inches(4.9), Inches(5.1), WHITE)
    tb(slide, Inches(8.15), Inches(1.6), Inches(4.5), Inches(0.4),
       "Story Helpers", size=16, bold=True, color=NAVY)
    bullets(slide, Inches(8.15), Inches(2.2), Inches(4.5), Inches(4.0), [
        "Who is in the story?",
        "Where are they?",
        "What happened?",
        "What was the problem?",
        "How did it end?",
        "How did they feel?",
    ], size=15, spacing=10)
    footer(slide, 14, "75–85 min")
    add_transition(slide)
    notes(slide,
          "SAY: \"Use the picture. Invent a NEW story (not only The Lost Puppy).\"\n"
          "Give 30 seconds planning. Aim for 6–10 sentences.\n"
          "EXPECTED: Clear who/where/problem/solution; creative details.\n"
          "FOLLOW-UP: \"Add one nature vocabulary word from today!\"\n"
          "ASSESS: Creativity, coherence, grammar in extended speech.\n"
          "TIMING: ~5 minutes of storytelling block.")


def s15_picture():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, SKY)
    chip(slide, "DESCRIBE IT", color=SKY)
    time_chip(slide, "~3 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Picture Description 🖼️", size=26, bold=True, color=NAVY, font="Georgia")

    add_round(slide, Inches(0.4), Inches(1.4), Inches(7.5), Inches(5.1), LIGHT_PURPLE)
    tb(slide, Inches(0.6), Inches(2.5), Inches(7.1), Inches(2.5),
       "🏫 Classroom party scene\n🎈 Friends laughing\n📚 Books on a table\n🍕 Snacks & smiles",
       size=22, color=NAVY, align=PP_ALIGN.CENTER)

    boxes = [
        ("People", "Who do you see?"),
        ("Place", "Where is this?"),
        ("Actions", "What are they doing?"),
        ("Feelings", "How do they feel?"),
    ]
    for i, (t, q) in enumerate(boxes):
        top = Inches(1.4 + i * 1.3)
        add_round(slide, Inches(8.2), top, Inches(4.7), Inches(1.15), WHITE)
        tb(slide, Inches(8.4), top + Inches(0.15), Inches(4.3), Inches(0.35),
           t, size=15, bold=True, color=SKY)
        tb(slide, Inches(8.4), top + Inches(0.55), Inches(4.3), Inches(0.4),
           q, size=13, color=DARK)
    footer(slide, 15, "75–85 min")
    add_transition(slide)
    notes(slide,
          "Student describes for 60–90 seconds using all four boxes.\n"
          "Push adjectives: colorful, excited, noisy, delicious.\n"
          "ASSESS: Observation language, present continuous (are laughing).\n"
          "TIMING: ~3 minutes.")


def s16_roleplay():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, GOLD)
    chip(slide, "ROLE PLAY", color=GOLD)
    time_chip(slide, "~2 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Creative Speaking — Pick a Role! 🎭", size=26, bold=True, color=NAVY, font="Georgia")

    roles = [
        ("📰", CORAL, "News Reporter", "Report: Puppy found in Maple Park!"),
        ("🏴‍☠️", NAVY, "Pirate", "Tell your crew about treasure in the forest."),
        ("🦁", TEAL, "Zoo Guide", "Welcome visitors and describe 3 animals."),
    ]
    for i, (e, c, t, d) in enumerate(roles):
        left = Inches(0.4 + i * 4.25)
        add_round(slide, left, Inches(1.6), Inches(4.05), Inches(4.8), WHITE)
        add_rect(slide, left, Inches(1.6), Inches(4.05), Inches(1.4), c)
        tb(slide, left, Inches(1.75), Inches(4.05), Inches(0.5), e, size=28, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.35), Inches(4.05), Inches(0.45), t, size=16, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(3.4), Inches(3.55), Inches(2.4),
           d + "\n\nSpeak for 45–60 seconds!", size=15, color=DARK, align=PP_ALIGN.CENTER)
    footer(slide, 16, "75–85 min")
    add_transition(slide)
    notes(slide,
          "Let Sage choose one role. Teacher can be the audience/interviewer.\n"
          "EXPECTED: Imaginative language; attempt formal tone for reporter.\n"
          "FOLLOW-UP: \"What happens next in your news story?\"\n"
          "ASSESS: Confidence, creativity, pronunciation under performance.\n"
          "TIMING: ~2 minutes (flex if early).")


def s17_pronunciation():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, PINK)
    chip(slide, "PRONUNCIATION", color=PINK, w=Inches(3.2))
    time_chip(slide, "Flex • 3 min")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Pronunciation Practice 🔊", size=26, bold=True, color=NAVY, font="Georgia")
    tb(slide, Inches(0.4), Inches(1.25), Inches(12), Inches(0.3),
       "Listen → Repeat → Use in a tiny sentence", size=14, color=SOFT)

    words = [
        ("adventure", "ad-VEN-chure"),
        ("library", "LY-brer-ee"),
        ("comfortable", "COMFter-bul"),
        ("chocolate", "CHOK-lit"),
        ("vegetable", "VEJ-tuh-bul"),
        ("interesting", "IN-truh-sting"),
        ("temperature", "TEM-pruh-chur"),
        ("Wednesday", "WENZ-day"),
        ("February", "FEB-roo-er-ee"),
        ("squirrel", "SKWUR-ul"),
    ]
    for i, (w, p) in enumerate(words):
        col, row = i % 5, i // 5
        left = Inches(0.35 + col * 2.55)
        top = Inches(1.8 + row * 2.4)
        add_round(slide, left, top, Inches(2.4), Inches(2.15), WHITE)
        tb(slide, left + Inches(0.1), top + Inches(0.45), Inches(2.2), Inches(0.5),
           w, size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), top + Inches(1.15), Inches(2.2), Inches(0.6),
           p, size=12, color=TEAL, align=PP_ALIGN.CENTER)
    footer(slide, 17, "Pronunciation")
    add_transition(slide)
    notes(slide,
          "Model clearly; clap syllables. Do 6–10 words based on time.\n"
          "Common errors: library (not lie-berry), Wednesday (silent d).\n"
          "EXPECTED: Improved approximation after 2 repeats.\n"
          "ASSESS: Segmentals + confidence saying hard words aloud.\n"
          "TIMING: 3 minutes (squeeze into speaking/story block if needed).")


def s18_fun_quiz():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, YELLOW)
    chip(slide, "FUN QUIZ", color=GOLD)
    time_chip(slide, "~4 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Mixed Fun Quiz 🎉", size=26, bold=True, color=NAVY, font="Georgia")

    items = [
        ("Grammar", "\"Gently\" is an…", "A) Noun  B) Adverb  C) Verb"),
        ("Vocab", "A soft wind is a…", "A) boulder  B) breeze  C) trail"),
        ("Reading", "The puppy's name was…", "A) Mia  B) Ava  C) Coco"),
        ("Grammar", "Best noun: ", "A) ran  B) park  C) quickly"),
        ("Vocab", "Opposite of tiny:", "A) huge  B) soft  C) glad"),
        ("Reading", "Mia found help from a…", "A) poster  B) boat  C) teacher only"),
        ("Bonus", "Synonym of happy:", "A) sad  B) joyful  C) slow"),
        ("Bonus", "Soar means…", "A) dig  B) fly high  C) sleep"),
    ]
    for i, (tag, q, opts) in enumerate(items):
        col, row = i % 4, i // 4
        left = Inches(0.3 + col * 3.25)
        top = Inches(1.4 + row * 2.7)
        add_round(slide, left, top, Inches(3.1), Inches(2.5), WHITE)
        add_rect(slide, left, top, Inches(3.1), Inches(0.45), [TEAL, SKY, CORAL, PURPLE][i % 4])
        tb(slide, left, top + Inches(0.08), Inches(3.1), Inches(0.35),
           tag, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.65), Inches(2.8), Inches(0.8),
           q, size=13, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.5), Inches(2.8), Inches(0.8),
           opts, size=12, color=DARK, align=PP_ALIGN.CENTER)
    footer(slide, 18, "Quiz")
    add_transition(slide)
    notes(slide,
          "Rapid-fire quiz. Keep celebratory tone.\n"
          "ANSWERS: B, B, C, B, A, A, B, B.\n"
          "ASSESS: Retention across skills.\n"
          "TIMING: ~4 minutes before assessment/recap.")


def s19_assessment():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, NAVY)
    chip(slide, "MINI ASSESSMENT", color=NAVY, w=Inches(3.3))
    time_chip(slide, "Teacher use")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Sage — Skills Snapshot ⭐", size=26, bold=True, color=NAVY, font="Georgia")

    # Header
    headers = ["Skill", "⭐", "⭐⭐", "⭐⭐⭐", "⭐⭐⭐⭐", "⭐⭐⭐⭐⭐"]
    widths = [2.6, 1.8, 1.8, 1.8, 2.0, 2.2]
    x = 0.4
    add_rect(slide, Inches(0.4), Inches(1.4), Inches(12.5), Inches(0.55), NAVY)
    for h, w in zip(headers, widths):
        tb(slide, Inches(x), Inches(1.5), Inches(w), Inches(0.4), h, size=12, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
        x += w

    skills = ["Reading", "Speaking", "Grammar", "Vocabulary", "Confidence", "Pronunciation"]
    for i, skill in enumerate(skills):
        top = Inches(2.05 + i * 0.75)
        bg = WHITE if i % 2 == 0 else LIGHT_SKY
        add_rect(slide, Inches(0.4), top, Inches(12.5), Inches(0.75), bg)
        x = 0.4
        tb(slide, Inches(x), top + Inches(0.2), Inches(widths[0]), Inches(0.4),
           skill, size=14, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        x += widths[0]
        for w in widths[1:]:
            tb(slide, Inches(x), top + Inches(0.2), Inches(w), Inches(0.4),
               "○", size=16, color=SOFT, align=PP_ALIGN.CENTER)
            x += w
    footer(slide, 19, "Assessment")
    add_transition(slide)
    notes(slide,
          "Complete privately during/after class (do not stress Sage).\n"
          "1=emerging … 5=excellent for Grade 5 demo expectations.\n"
          "Note 1 strength + 1 growth area for parents.\n"
          "TIMING: Ongoing observation; finalize in last 2 minutes.")


def s20_recap():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, TEAL)
    chip(slide, "RECAP", color=TEAL)
    time_chip(slide, "5 minutes")
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "What Did We Learn Today? 🧠", size=26, bold=True, color=NAVY, font="Georgia")

    cards = [
        ("🌿 Words", TEAL, "breeze, soar, drizzle, trail, bloom, mist…"),
        ("📘 Grammar", SKY, "Nouns, verbs, adjectives, adverbs"),
        ("📖 Reading", PURPLE, "Find answers + infer feelings from clues"),
        ("🗣️ Speaking", CORAL, "Full sentences, story retell, role play"),
    ]
    for i, (t, c, d) in enumerate(cards):
        left = Inches(0.4 + (i % 2) * 6.4)
        top = Inches(1.5 + (i // 2) * 2.3)
        add_round(slide, left, top, Inches(6.15), Inches(2.1), WHITE)
        add_rect(slide, left, top, Inches(6.15), Inches(0.55), c)
        tb(slide, left + Inches(0.2), top + Inches(0.12), Inches(5.7), Inches(0.35),
           t, size=16, bold=True, color=WHITE)
        tb(slide, left + Inches(0.25), top + Inches(0.85), Inches(5.7), Inches(1.0),
           d, size=15, color=DARK)
    add_round(slide, Inches(0.4), Inches(6.2), Inches(12.5), Inches(0.7), LIGHT_YELLOW)
    tb(slide, Inches(0.6), Inches(6.35), Inches(12.1), Inches(0.4),
       "Reflection: What did you enjoy the most today, Sage?", size=15, bold=True, color=NAVY)
    footer(slide, 20, "85–90 min")
    add_transition(slide)
    notes(slide,
          "Quick student-led recap. Ask favorite activity.\n"
          "EXPECTED: Mentions a game/story/word.\n"
          "End with specific praise: \"Your evidence answers were excellent!\"\n"
          "TIMING: 5 minutes.")


def s21_homework():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, SKY)
    chip(slide, "HOMEWORK", color=SKY)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Mission for Sage 📝", size=28, bold=True, color=NAVY, font="Georgia")

    add_round(slide, Inches(1.5), Inches(1.6), Inches(10.3), Inches(4.8), WHITE)
    icon_bubble(slide, Inches(2.0), Inches(2.0), "🐾", LIGHT_TEAL)
    tb(slide, Inches(3.0), Inches(2.15), Inches(8), Inches(0.4),
       "Write 5 sentences about your favorite animal", size=20, bold=True, color=NAVY)
    bullets(slide, Inches(2.2), Inches(3.0), Inches(9), Inches(3.0), [
        "Use at least 5 new vocabulary words from today",
        "Underline nouns in blue (or say them aloud)",
        "Circle adjectives in green",
        "Read your sentences aloud to a family member",
        "Bonus: Draw your animal and label 3 nature words",
    ], size=16, spacing=10)
    footer(slide, 21, "Homework")
    add_transition(slide)
    notes(slide,
          "Explain homework clearly; keep it light and doable (15–20 min).\n"
          "Optional share next class.\n"
          "TIMING: 1–2 minutes.")


def s22_thanks():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.2), YELLOW)
    # certificate card
    add_round(slide, Inches(2.2), Inches(1.2), Inches(8.9), Inches(5.0), WHITE)
    add_rect(slide, Inches(2.2), Inches(1.2), Inches(8.9), Inches(0.25), TEAL)
    add_rect(slide, Inches(2.2), Inches(5.95), Inches(8.9), Inches(0.25), GOLD)
    tb(slide, Inches(2.5), Inches(1.7), Inches(8.3), Inches(0.4),
       "🌟 CERTIFICATE OF AWESOME EFFORT 🌟", size=16, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.5), Inches(2.4), Inches(8.3), Inches(0.5),
       "Great Job, Sage!", size=36, bold=True, color=NAVY, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(2.5), Inches(3.2), Inches(8.3), Inches(1.2),
       "You explored reading, speaking, grammar,\nvocabulary, and pronunciation with courage and joy.",
       size=16, color=DARK, align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.5), Inches(4.6), Inches(8.3), Inches(0.4),
       "Keep shining — your English adventure has just begun!",
       size=15, bold=True, color=CORAL, align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.5), Inches(5.2), Inches(8.3), Inches(0.4),
       "Grade 5 English Demo Class", size=13, color=SOFT, align=PP_ALIGN.CENTER)
    add_transition(slide)
    notes(slide,
          "Celebrate specifically: name 2 strengths observed.\n"
          "Thank Sage and parents if present.\n"
          "Invite questions.\n"
          "TIMING: 1–2 minutes closing.")


def s23_answer_key_reading():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, TEAL)
    chip(slide, "ANSWER KEY A", color=TEAL)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Answer Key — Reading & Fun Quiz", size=24, bold=True, color=NAVY, font="Georgia")
    add_round(slide, Inches(0.4), Inches(1.4), Inches(6.2), Inches(5.1), WHITE)
    tb(slide, Inches(0.6), Inches(1.55), Inches(5.8), Inches(0.35),
       "Reading Questions", size=16, bold=True, color=TEAL)
    bullets(slide, Inches(0.6), Inches(2.1), Inches(5.8), Inches(4.2), [
        "1. B) Maple Park",
        "2. C) Brown",
        "3. False (no name tag)",
        "4. True",
        "5. Her water bottle",
        "6. A lost-puppy poster on the board",
        "7. She helped Coco reunite with Ava / showed kindness",
        "8. B) Soft sad sound",
    ], size=14, spacing=6)

    add_round(slide, Inches(6.9), Inches(1.4), Inches(6.0), Inches(5.1), LIGHT_YELLOW)
    tb(slide, Inches(7.1), Inches(1.55), Inches(5.6), Inches(0.35),
       "Fun Quiz", size=16, bold=True, color=GOLD)
    bullets(slide, Inches(7.1), Inches(2.1), Inches(5.6), Inches(4.2), [
        "1. B Adverb",
        "2. B breeze",
        "3. C Coco",
        "4. B park",
        "5. A huge",
        "6. A poster",
        "7. B joyful",
        "8. B fly high",
    ], size=15, spacing=8)
    footer(slide, 23, "Keys")
    add_transition(slide)
    notes(slide, "Teacher reference only — reveal after student attempts.\nUse to confirm scoring and parent feedback examples.")


def s24_answer_key_grammar():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, CORAL)
    chip(slide, "ANSWER KEY B", color=CORAL)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "Answer Key — Grammar & Vocabulary", size=24, bold=True, color=NAVY, font="Georgia")

    add_round(slide, Inches(0.4), Inches(1.35), Inches(6.2), Inches(5.2), WHITE)
    tb(slide, Inches(0.6), Inches(1.5), Inches(5.8), Inches(0.35),
       "Grammar Quiz (1–10)", size=16, bold=True, color=CORAL)
    bullets(slide, Inches(0.6), Inches(2.05), Inches(5.8), Inches(4.2), [
        "1B Noun   2C Verb   3B Noun",
        "4A Adverb   5B muddy   6A Proper noun",
        "7B dialed   8A Adjective",
        "9B How   10A Noun",
        "",
        "Practice: walked | tiny/brown | birds | gently",
        "Happy(adj) Ava(n) quickly(adv) hugged(v) Coco(n)",
    ], size=14, spacing=8)

    add_round(slide, Inches(6.9), Inches(1.35), Inches(6.0), Inches(5.2), LIGHT_TEAL)
    tb(slide, Inches(7.1), Inches(1.5), Inches(5.6), Inches(0.35),
       "Matching + Syn/Ant", size=16, bold=True, color=TEAL)
    bullets(slide, Inches(7.1), Inches(2.05), Inches(5.6), Inches(4.2), [
        "Matching: 1-B, 2-D, 3-A, 4-E, 5-F, 6-C",
        "",
        "Synonyms: big→large/huge/enormous",
        "Antonyms: kind→mean/unkind/cruel",
        "",
        "Tip: Accept reasonable Grade 5 synonyms.",
    ], size=14, spacing=8)
    footer(slide, 24, "Keys")
    add_transition(slide)
    notes(slide, "Teacher key. Reveal answers with celebration animations verbally (\"Ding!\").")


def s25_timing():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, NAVY)
    chip(slide, "FACILITATOR MAP", color=NAVY, w=Inches(3.3))
    tb(slide, Inches(0.4), Inches(0.8), Inches(12), Inches(0.4),
       "90-Minute Timing Guide ⏱️", size=26, bold=True, color=NAVY, font="Georgia")

    rows = [
        ("0–10", "Introduction + Warm-up talk", "Slides 1–4"),
        ("10–20", "Icebreaker: Would You Rather?", "Slide 3 (extend)"),
        ("20–40", "Reading passage + questions + retell", "Slides 5–7"),
        ("40–60", "Grammar teach + practice + quiz", "Slides 8–10"),
        ("60–75", "Vocabulary + matching + syn/ant", "Slides 11–13"),
        ("75–85", "Storytelling + describe + role play", "Slides 14–16"),
        ("85–90", "Recap + homework + celebrate", "Slides 20–22"),
    ]
    add_rect(slide, Inches(0.4), Inches(1.4), Inches(12.5), Inches(0.5), NAVY)
    tb(slide, Inches(0.6), Inches(1.5), Inches(2.2), Inches(0.35), "Time", size=13, bold=True, color=WHITE)
    tb(slide, Inches(3.0), Inches(1.5), Inches(6.5), Inches(0.35), "Block", size=13, bold=True, color=WHITE)
    tb(slide, Inches(9.6), Inches(1.5), Inches(3.0), Inches(0.35), "Slides", size=13, bold=True, color=WHITE)
    for i, (t, b, s) in enumerate(rows):
        top = Inches(2.0 + i * 0.65)
        bg = WHITE if i % 2 == 0 else LIGHT_SKY
        add_rect(slide, Inches(0.4), top, Inches(12.5), Inches(0.65), bg)
        tb(slide, Inches(0.6), top + Inches(0.15), Inches(2.2), Inches(0.35), t, size=13, bold=True, color=TEAL)
        tb(slide, Inches(3.0), top + Inches(0.15), Inches(6.5), Inches(0.35), b, size=13, color=DARK)
        tb(slide, Inches(9.6), top + Inches(0.15), Inches(3.0), Inches(0.35), s, size=13, color=SOFT)
    footer(slide, 25, "Teacher")
    add_transition(slide)
    notes(slide,
          "Use this map to stay on pace. If Sage needs more speaking time, shorten quiz.\n"
          "Pronunciation slide is flexible—insert when energy dips.\n"
          "Parent wow factor: specific praise + clear next step (homework).")


# Build
s01_welcome()
s02_teacher()
s03_icebreaker()
s04_warmup()
s05_passage()
s06_questions()
s07_speaking()
s08_grammar()
s09_grammar_practice()
s10_grammar_quiz()
s11_vocab()
s12_matching()
s13_syn_ant()
s14_storytelling()
s15_picture()
s16_roleplay()
s17_pronunciation()
s18_fun_quiz()
s19_assessment()
s20_recap()
s21_homework()
s22_thanks()
s23_answer_key_reading()
s24_answer_key_grammar()
s25_timing()

out = r"C:\Users\bhushaja\Downloads\shaip\Grade5_Sage_English_Demo_90min.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
