"""Grade 5 English lesson - 90 minutes, 55 slides, no speaker notes.

"The Mystery of the Missing Treasure" - the student is a Young Mystery Detective
collecting five clue badges. Reading, vocabulary, grammar, speaking and short
writing are interleaved so no activity type runs longer than about ten minutes.

Teaching support lives on visible slides rather than in speaker notes: every
activity carries a Detective Hint strip, and slides 53-55 hold the support
ladder, the assessment table and the full answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x1B, 0x1B, 0x32)
DARK = RGBColor(0x23, 0x25, 0x38)
SOFT = RGBColor(0x6E, 0x74, 0x86)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PARCH = RGBColor(0xFB, 0xF7, 0xEC)
GOLD = RGBColor(0xC9, 0x96, 0x0C)
CRIMSON = RGBColor(0xC0, 0x39, 0x2B)
EMERALD = RGBColor(0x1E, 0x84, 0x49)
TEAL = RGBColor(0x11, 0x8F, 0x96)
PURPLE = RGBColor(0x6C, 0x4A, 0xB6)
AMBER = RGBColor(0xD9, 0x7D, 0x0A)
SLATE = RGBColor(0x3D, 0x4A, 0x66)
L_GOLD = RGBColor(0xFA, 0xF0, 0xD5)
L_CRIMSON = RGBColor(0xFB, 0xE9, 0xE7)
L_EMERALD = RGBColor(0xE4, 0xF3, 0xE9)
L_TEAL = RGBColor(0xDF, 0xF2, 0xF3)
L_PURPLE = RGBColor(0xEE, 0xE9, 0xFA)
L_AMBER = RGBColor(0xFD, 0xF0, 0xDC)
L_SLATE = RGBColor(0xEC, 0xEF, 0xF5)
L_GREY = RGBColor(0xF3, 0xF5, 0xF9)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 55
_counter = {"n": 0}

# ------------------------------------------------------------------ content

BADGES = [("🔎", "Reading Clue", "Words & sentences", CRIMSON, L_CRIMSON),
          ("🧩", "Vocabulary Clue", "Mystery words", PURPLE, L_PURPLE),
          ("✏️", "Grammar Clue", "Parts of speech & tenses", EMERALD, L_EMERALD),
          ("📖", "Story Clue", "The old door", TEAL, L_TEAL),
          ("🏆", "Final Clue", "Solve the mystery", GOLD, L_GOLD)]

SCENE = [("🗺️", "an old map"), ("🔑", "a key"), ("💰", "a treasure chest"),
         ("👣", "footprints"), ("🚪", "a door"), ("🔍", "a magnifying glass")]

STARTERS = ["I see ______.", "I think ______.", "Maybe ______.",
            "The clue might mean ______."]

PREDICT = [("🧰", "What might be inside the chest?"),
           ("👣", "Who left those footprints?"),
           ("🚪", "Where do you think the door leads?"),
           ("🗺️", "Why would someone hide a map?")]

VOCAB_A = [("detective", "🕵️", "A person who solves mysteries.",
            "The detective studied every clue.", CRIMSON, L_CRIMSON),
           ("mystery", "❓", "Something strange we do not understand yet.",
            "The mystery kept everyone guessing.", PURPLE, L_PURPLE),
           ("treasure", "💰", "Something valuable that is kept or hidden.",
            "They searched for buried treasure.", GOLD, L_GOLD),
           ("clue", "🔍", "A piece of information that helps solve a problem.",
            "A muddy footprint was the first clue.", TEAL, L_TEAL)]

VOCAB_B = [("map", "🗺️", "A drawing that shows where places are.",
            "The map showed a path through the forest.", EMERALD, L_EMERALD),
           ("secret", "🤫", "Something kept hidden from other people.",
            "She kept the secret for years.", PURPLE, L_PURPLE),
           ("ancient", "🏛️", "Extremely old.",
            "An ancient door stood in the wall.", AMBER, L_AMBER),
           ("discover", "💡", "To find something for the first time.",
            "They discovered a hidden room.", CRIMSON, L_CRIMSON)]

WORD_ROUTINE = [("1", "I DO", "I say the word and its meaning.", CRIMSON),
                ("2", "WE DO", "We say it together, twice.", AMBER),
                ("3", "YOU DO", "You say it, explain it, use it.", EMERALD)]

REVEAL_A = [("DETECTIVE", "de·tec·tive", "🕵️"), ("CLUE", "clue", "🔍"),
            ("TREASURE", "trea·sure", "💰"), ("MAP", "map", "🗺️")]
REVEAL_B = [("KEY", "key", "🔑"), ("SECRET", "se·cret", "🤫"),
            ("SEARCH", "search", "👀"), ("DISCOVER", "dis·cov·er", "💡")]

POS = [("NOUN", "person, place, or thing", "detective, map, door", CRIMSON, L_CRIMSON),
       ("VERB", "the action word", "opened, searched, found", EMERALD, L_EMERALD),
       ("ADJECTIVE", "describes a noun", "clever, heavy, ancient", PURPLE, L_PURPLE),
       ("ADVERB", "describes a verb (often -ly)", "quickly, silently, slowly", TEAL,
        L_TEAL)]

# Each word carries the part of speech it belongs to, or None for plain text.
DEMO_SENTENCE = [("The", None), ("clever", "adj"), ("detective", "noun"),
                 ("quickly", "adv"), ("opened", "verb"), ("the", None),
                 ("heavy", "adj"), ("door.", "noun")]

DEMO_EXTRA = [
    [("A", None), ("strange", "adj"), ("light", "noun"), ("suddenly", "adv"),
     ("appeared", "verb"), ("in", None), ("the", None), ("ancient", "adj"),
     ("castle.", "noun")],
    [("Maya", "noun"), ("quietly", "adv"), ("searched", "verb"), ("the", None),
     ("dusty", "adj"), ("attic.", "noun")],
]

GD_A = [("The brave girl carefully opened the secret box.",
         "girl, box", "opened", "brave, secret", "carefully"),
        ("The young detective silently followed the muddy footprints.",
         "detective, footprints", "followed", "young, muddy", "silently"),
        ("A rusty key suddenly fell from the old shelf.",
         "key, shelf", "fell", "rusty, old", "suddenly")]
GD_B = [("The curious students quickly solved the difficult puzzle.",
         "students, puzzle", "solved", "curious, difficult", "quickly"),
        ("Maya carefully unfolded the ancient treasure map.",
         "Maya, map", "unfolded", "ancient, treasure", "carefully")]

PHRASES = [("a hidden key", "🔑"), ("an old treasure map", "🗺️"),
           ("a secret room", "🚪"), ("the mysterious door", "❓"),
           ("a dangerous forest path", "🌲"), ("the ancient stone wall", "🧱")]

CLUE_OR_NOT = [("KEY", True), ("RUN", False), ("BLUE", False), ("TREASURE", True),
               ("APPLE", False), ("SECRET", True), ("JUMP", False), ("MAP", True)]

WORD_POINTS = [("1 pt", "Read the word out loud", CRIMSON),
               ("2 pts", "Explain what it means", AMBER),
               ("3 pts", "Use it in a full sentence", EMERALD)]

WORD_CHALLENGE = [("ANCIENT", "🏛️"), ("DISCOVER", "💡"), ("DANGEROUS", "⚠️"),
                  ("HIDDEN", "🫥")]

SUBJ_PRED = [("The detective", "found a key."), ("The old door", "opened slowly."),
             ("Maya and her grandmother", "solved the mystery together.")]

LOCK_A = [("The detective", "found a hidden map."),
          ("The old door", "opened slowly."),
          ("The students", "searched the room.")]
LOCK_B = [("A secret passage", "led to the treasure."),
          ("Maya and her grandmother", "solved the mystery together.")]

BREAK_TASKS = [("🔵", "Find something BLUE"), ("⭕", "Find something ROUND"),
               ("🤏", "Find something SMALL"), ("🅱️", "Find something starting with B"),
               ("📏", "Find something LONG"), ("✨", "Find something SHINY")]

TIMELINE = [("PAST", "Already happened", "The detective searched the room.", CRIMSON,
             L_CRIMSON),
            ("NOW", "Happening today", "The detective searches the room.", EMERALD,
             L_EMERALD),
            ("FUTURE", "Has not happened yet", "The detective will search the room.",
             PURPLE, L_PURPLE)]

TENSE_A = [("The detective searches the room.", "PRESENT"),
           ("She followed the map into the forest.", "PAST"),
           ("We will open the box tomorrow.", "FUTURE")]
TENSE_B = [("Maya finds a rusty key.", "PRESENT"),
           ("They discovered a secret room.", "PAST")]

TRANSFORM = [("She finds the key.", "She found the key.", "She will find the key."),
             ("The detective opens the door.", "The detective opened the door.",
              "The detective will open the door."),
             ("They search the attic.", "They searched the attic.",
              "They will search the attic.")]

AGREEMENT_RULE = [("One", "detective", "finds", "singular subject → verb takes -s",
                   CRIMSON, L_CRIMSON),
                  ("Many", "detectives", "find", "plural subject → verb has no -s",
                   EMERALD, L_EMERALD)]

AGREEMENT = [("The detective ______ the clue.", "find / finds", "finds"),
             ("The detectives ______ the clues.", "find / finds", "find"),
             ("Maya ______ the dusty attic.", "search / searches", "searches"),
             ("The clues ______ to the old door.", "point / points", "point"),
             ("The map ______ a path through the forest.", "show / shows", "shows")]

SENTENCES_L3 = [("The detective found a strange key.", "🔑"),
                ("She followed the map through the forest.", "🌲"),
                ("The old door was locked.", "🚪"),
                ("A shining object sat inside the nest.", "✨")]

STORY_VOCAB = [("mysterious", "strange and hard to explain"),
               ("ancient", "extremely old"), ("hidden", "put out of sight"),
               ("discover", "to find for the first time"),
               ("faded", "lost its color over time"), ("ivy", "a climbing green plant")]

STORY = [
    ("Part 1", "📦", CRIMSON, L_CRIMSON, [
        "Maya was eleven years old, and she loved solving mysteries. One rainy Saturday, "
        "she was cleaning her grandmother's attic when she found an old wooden box.",
        "Inside was a folded paper, yellow with age. Maya opened it carefully. It was a "
        "map. A thin red line curved through a forest and stopped at a small drawing of "
        "a door.",
        "Under the door, someone had written four words in faded ink: "
        "The treasure waits inside."]),
    ("Part 2", "🌲", EMERALD, L_EMERALD, [
        "Maya showed the map to her grandmother. \"I drew that when I was your age,\" "
        "Grandma said with a smile. \"But I never found what was behind the door.\"",
        "That was all Maya needed to hear. The next morning she packed a flashlight, a "
        "notebook, and two sandwiches, and followed the red line into the forest behind "
        "the house.",
        "The trees were tall and quiet. After twenty minutes, she saw something strange "
        "between them: a stone wall, covered in ivy."]),
    ("Part 3", "🚪", PURPLE, L_PURPLE, [
        "In the middle of the wall stood an ancient door. It was made of dark wood, and "
        "it was locked. Maya searched everywhere for a key. She looked under rocks. She "
        "looked behind bushes. Nothing.",
        "Then she remembered the map. She unfolded it again and noticed a tiny arrow she "
        "had missed, pointing to the bottom corner. There, drawn in the same faded ink, "
        "was a picture of a bird's nest.",
        "Maya looked up. A nest sat in the ivy above the door, and something metal was "
        "shining inside it."]),
    ("Part 4", "📚", GOLD, L_GOLD, [
        "Maya reached up and pulled out a small iron key. Her hands shook as she turned "
        "it in the lock. The door opened slowly.",
        "Inside was not gold. Inside was a tiny room full of books, letters, and "
        "drawings — hundreds of them, all made by her grandmother when she was young.",
        "Maya sat down on the dusty floor and began to read. She had found the treasure "
        "after all. It was a story, and it belonged to her family."]),
]

EVIDENCE_STEPS = [("1", "Read the question", "What is it actually asking?", CRIMSON),
                  ("2", "Go back to the story", "Do not answer from memory.", AMBER),
                  ("3", "Find the sentence", "Point to the exact line.", EMERALD),
                  ("4", "Say it out loud", "\"My evidence is...\"", TEAL)]

Q_LITERAL = [("Who is the main character?", "Part 1"),
             ("What did Maya find in the attic?", "Part 1"),
             ("Where did the red line on the map lead her?", "Part 2")]
Q_MIXED = [("VOCABULARY", "What does \"ancient\" mean in Part 3?", PURPLE, L_PURPLE),
           ("VOCABULARY", "What does \"faded\" tell you about the ink?", PURPLE,
            L_PURPLE),
           ("INFERENCE", "Why do you think the door mattered to Grandma?", TEAL, L_TEAL),
           ("INFERENCE", "Why did Maya keep searching after she found no key?", TEAL,
            L_TEAL)]
Q_EVIDENCE = [("Where exactly was the key hidden?",
               "Find the sentence in Part 3 that proves it."),
              ("Was Maya disappointed by what she found?",
               "Find the sentence in Part 4 that proves it.")]

ERROR_TYPES = [("🔠", "Capital letters", "Sentences and names start with capitals."),
               ("❗", "Punctuation", "Every sentence needs an end mark."),
               ("⏰", "Verb tense", "Does the time match the rest of the sentence?"),
               ("🤝", "Subject-verb agreement", "One finds. Many find."),
               ("🅰️", "a / an", "Use \"an\" before a vowel sound."),
               ("🔤", "Spelling", "Read it slowly, syllable by syllable.")]

FIX_A = [("the detective find a old key", "The detective found an old key.",
          "capital · tense · a/an"),
         ("she open the door slow", "She opened the door slowly.",
          "capital · tense · adverb · period"),
         ("maya and her friend searchs the forest",
          "Maya and her friend searched the forest.",
          "capitals · agreement · tense")]
FIX_B = [("the map show a secrit room", "The map shows a secret room.",
          "capital · agreement · spelling"),
         ("they will finded the treasure tomorow",
          "They will find the treasure tomorrow.",
          "capital · future form · spelling")]

WRITING_FRAMES = ["Today I discovered ______________________________.",
                  "The most important clue was ______________________________.",
                  "I think the treasure was ______________________________."]

FINAL_CLUES = [("CLUE 1", "🔎", "READ", "Read these three words out loud:",
                "ANCIENT · DISCOVER · MYSTERIOUS", CRIMSON, L_CRIMSON),
               ("CLUE 2", "✏️", "GRAMMAR", "Find the adverb in this sentence:",
                "The clever detective searched quietly.", EMERALD, L_EMERALD),
               ("CLUE 3", "🛠️", "EDIT", "Fix this sentence:",
                "the door was lock", PURPLE, L_PURPLE),
               ("CLUE 4", "📖", "COMPREHEND", "Answer from the story:",
                "Where did Maya find the key?", TEAL, L_TEAL),
               ("CLUE 5", "⭐", "CREATE", "Build your own sentence using:",
                "TREASURE", GOLD, L_GOLD)]

EXIT_TICKET = [("1", "Name one new word you learned today.", CRIMSON),
               ("2", "Change to past tense:  \"The detective opens the door.\"", EMERALD),
               ("3", "What was your favorite activity?", PURPLE),
               ("4", "What is one thing you learned today?", TEAL)]

SUPPORT_LEVELS = [
    ("🟢", "LEVEL 1 — SUPPORTED", "Detective Hint",
     "Give two choices and a hint. Model the answer first, then ask again.",
     "\"Is it 'find' or 'finds'? Let's look at the subject.\"", EMERALD, L_EMERALD),
    ("🟡", "LEVEL 2 — GUIDED", "Challenge Clue",
     "She answers with a nudge. Point at the part of the sentence that matters.",
     "\"Look at the word right before the verb.\"", AMBER, L_AMBER),
    ("🔵", "LEVEL 3 — CHALLENGE", "Bonus Mission",
     "She writes her own example or explains why her answer works.",
     "\"Can you make your own sentence and tell me why it's correct?\"", TEAL, L_TEAL)]

PRAISE = ["\"Let's try that clue again.\"", "\"Good thinking!\"",
          "\"Let's check the evidence.\"", "\"You're getting closer!\"",
          "\"Detective hint!\"", "\"That's exactly how a detective thinks.\""]

HINT_BANK = ["Look at the subject first.", "Ask: who is doing the action?",
             "Look for the action word.", "Check the ending of the verb.",
             "Read the sentence one more time.", "Look back at the story."]

RUBRIC_SKILLS = ["Word Reading", "Reading Fluency", "Vocabulary", "Comprehension",
                 "Parts of Speech", "Verb Tenses", "Subject-Verb Agreement",
                 "Sentence Editing", "Speaking"]

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


def para(slide, l, t, w, h, text, size=15, color=DARK, sp=8, line=None):
    """Justified body copy for story pages."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.space_after = Pt(sp)
    if line:
        p.line_spacing = line
    r = p.add_run()
    r.text = text
    set_run(r, size, False, color)
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


POS_COLOR = {"noun": CRIMSON, "verb": EMERALD, "adj": PURPLE, "adv": TEAL}


def colored_sentence(slide, l, t, w, h, words, size=28, align=PP_ALIGN.LEFT):
    """Renders a sentence with each word tinted by its part of speech."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    for i, (word, kind) in enumerate(words):
        r = p.add_run()
        r.text = word if i == 0 else " " + word
        color = POS_COLOR.get(kind, SLATE)
        set_run(r, size, kind is not None, color)
    return box


def fade(slide):
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    etree.SubElement(tr, qn("p:fade"))


def footer(slide, n, timing="", badge=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE2, 0xE7, 0xEF))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), GOLD)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), INK)
    msg = "🕵️ The Mystery of the Missing Treasure  |  Grade 5  |  90 min"
    if badge:
        msg += f"  |  {badge}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=10,
       color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, badge, accent=GOLD, bg=PARCH):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, accent)
    add_round(slide, Inches(0.4), Inches(0.26), Inches(3.5), Inches(0.42), accent)
    tb(slide, Inches(0.4), Inches(0.31), Inches(3.5), Inches(0.34), tag, size=12,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if timing:
        add_round(slide, Inches(10.3), Inches(0.26), Inches(2.65), Inches(0.42), SLATE)
        tb(slide, Inches(10.3), Inches(0.31), Inches(2.65), Inches(0.34), timing, size=12,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title, size=27,
       bold=True, color=INK, font="Georgia")
    footer(slide, n, timing, badge)
    fade(slide)
    return slide, n


def one_task(slide, text, color=CRIMSON, top=1.34):
    """States the single job of this slide, in plain language."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text, size=15,
              bold=True, color=color)


def hint(slide, text, top=6.42, label="💡 DETECTIVE HINT"):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), L_AMBER)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.3), Inches(0.34), label, size=12,
       bold=True, color=AMBER)
    tb(slide, Inches(3.15), Inches(top + 0.05), Inches(9.4), Inches(0.36), text, size=13,
       bold=True, color=INK)


def ido_wedo_youdo(slide, top=6.42):
    steps = [("I DO", "I model it", CRIMSON), ("WE DO", "We do it together", AMBER),
             ("YOU DO", "You take the lead", EMERALD)]
    for i, (label, who, color) in enumerate(steps):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(top), Inches(3.9), Inches(0.46), color)
        tb(slide, left, Inches(top + 0.05), Inches(3.9), Inches(0.36),
           f"{label}  ·  {who}", size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def numbered_rows(slide, count, start_index, accent, top_start=1.9, gap=1.5, height=1.34,
                  fill=WHITE):
    """Shared row scaffold: returns the top edge of each numbered row."""
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


def vocab_cards(slide, items, top=1.9):
    for i, (word, emoji, meaning, example, color, light) in enumerate(items):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        t = Inches(top + row * 2.3)
        add_round(slide, left, t, Inches(5.95), Inches(2.1), light)
        tb(slide, left + Inches(0.25), t + Inches(0.42), Inches(1.0), Inches(0.8), emoji,
           size=32, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.35), t + Inches(0.18), Inches(4.4), Inches(0.5),
           word.upper(), size=24, bold=True, color=color, font="Arial Black")
        tb(slide, left + Inches(1.35), t + Inches(0.72), Inches(4.4), Inches(0.5),
           meaning, size=13, color=DARK)
        add_round(slide, left + Inches(1.35), t + Inches(1.32), Inches(4.35),
                  Inches(0.58), WHITE)
        tb(slide, left + Inches(1.5), t + Inches(1.44), Inches(4.05), Inches(0.4),
           example, size=12, color=SLATE, italic=True)


def reveal_rows(slide, items, start_index, accent, light):
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=1.95, gap=1.16,
                         height=1.0)
    for (word, syll, emoji), top in zip(items, tops):
        tb(slide, Inches(1.45), top + Inches(0.18), Inches(0.9), Inches(0.66), emoji,
           size=26, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.5), top + Inches(0.16), Inches(4.3), Inches(0.68),
                  light)
        tb(slide, Inches(2.5), top + Inches(0.24), Inches(4.3), Inches(0.52), word,
           size=26, bold=True, color=accent, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, Inches(7.1), Inches(top.inches + 0.3), Inches(2.1), Inches(0.4), syll,
           size=15, color=SOFT)
        add_round(slide, Inches(9.35), top + Inches(0.2), Inches(3.25), Inches(0.6),
                  L_GREY)
        tb(slide, Inches(9.35), top + Inches(0.29), Inches(3.25), Inches(0.42),
           "read → explain → use it", size=12, bold=True, color=SLATE,
           align=PP_ALIGN.CENTER)


def find_rows(slide, items, start_index):
    """Grammar Detective: a sentence plus four labeled slots to fill in."""
    tops = numbered_rows(slide, len(items), start_index, EMERALD, top_start=1.9,
                         gap=1.52, height=1.36)
    labels = [("NOUN", CRIMSON), ("VERB", EMERALD), ("ADJ", PURPLE), ("ADV", TEAL)]
    for (sentence, _n, _v, _a, _d), top in zip(items, tops):
        tb(slide, Inches(1.45), top + Inches(0.14), Inches(11.1), Inches(0.5), sentence,
           size=19, bold=True, color=INK)
        for j, (label, color) in enumerate(labels):
            left = Inches(1.45 + j * 2.8)
            add_round(slide, left, top + Inches(0.72), Inches(2.6), Inches(0.5), L_GREY)
            tb(slide, left + Inches(0.12), top + Inches(0.8), Inches(1.0), Inches(0.36),
               label, size=12, bold=True, color=color)
            tb(slide, left + Inches(1.05), top + Inches(0.8), Inches(1.45), Inches(0.36),
               "____________", size=12, color=SOFT)


def lock_rows(slide, items, start_index):
    tops = numbered_rows(slide, len(items), start_index, PURPLE, top_start=1.95, gap=1.5,
                         height=1.32)
    for (subject, predicate), top in zip(items, tops):
        add_round(slide, Inches(1.45), top + Inches(0.3), Inches(4.6), Inches(0.72),
                  L_CRIMSON)
        tb(slide, Inches(1.6), top + Inches(0.1), Inches(4.3), Inches(0.28), "SUBJECT",
           size=10, bold=True, color=CRIMSON)
        tb(slide, Inches(1.45), top + Inches(0.42), Inches(4.6), Inches(0.5), subject,
           size=20, bold=True, color=CRIMSON, align=PP_ALIGN.CENTER)
        tb(slide, Inches(6.2), top + Inches(0.38), Inches(0.6), Inches(0.5), "🔗",
           size=18, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(6.95), top + Inches(0.3), Inches(5.65), Inches(0.72),
                  L_EMERALD)
        tb(slide, Inches(7.1), top + Inches(0.1), Inches(5.3), Inches(0.28), "PREDICATE",
           size=10, bold=True, color=EMERALD)
        tb(slide, Inches(6.95), top + Inches(0.42), Inches(5.65), Inches(0.5), predicate,
           size=20, bold=True, color=EMERALD, align=PP_ALIGN.CENTER)


def tense_rows(slide, items, start_index):
    tops = numbered_rows(slide, len(items), start_index, PURPLE, top_start=2.0, gap=1.55,
                         height=1.38)
    for (sentence, _answer), top in zip(items, tops):
        tb(slide, Inches(1.45), top + Inches(0.14), Inches(11.1), Inches(0.5), sentence,
           size=21, bold=True, color=INK)
        for j, (label, color, light) in enumerate([("PAST", CRIMSON, L_CRIMSON),
                                                   ("PRESENT", EMERALD, L_EMERALD),
                                                   ("FUTURE", PURPLE, L_PURPLE)]):
            left = Inches(1.45 + j * 3.75)
            add_round(slide, left, top + Inches(0.74), Inches(3.5), Inches(0.5), light)
            tb(slide, left, top + Inches(0.82), Inches(3.5), Inches(0.36), label, size=14,
               bold=True, color=color, align=PP_ALIGN.CENTER)


def fix_rows(slide, items, start_index):
    tops = numbered_rows(slide, len(items), start_index, AMBER, top_start=1.95, gap=1.5,
                         height=1.32)
    for (broken, _fixed, types), top in zip(items, tops):
        add_round(slide, Inches(1.45), top + Inches(0.18), Inches(5.6), Inches(0.66),
                  L_CRIMSON)
        tb(slide, Inches(1.6), top + Inches(0.29), Inches(5.3), Inches(0.46), broken,
           size=17, bold=True, color=CRIMSON)
        tb(slide, Inches(7.2), top + Inches(0.3), Inches(0.5), Inches(0.44), "→",
           size=18, bold=True, color=SOFT)
        add_round(slide, Inches(7.85), top + Inches(0.18), Inches(4.75), Inches(0.66),
                  L_EMERALD)
        tb(slide, Inches(7.85), top + Inches(0.3), Inches(4.75), Inches(0.46),
           "your correction here", size=13, color=EMERALD, align=PP_ALIGN.CENTER,
           italic=True)
        tb(slide, Inches(1.6), top + Inches(0.92), Inches(10.8), Inches(0.3),
           f"Check:  {types}", size=11, color=SOFT)


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    for x, y, c in [(0.6, 5.55, CRIMSON), (12.05, 5.5, TEAL)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(0.9), Inches(12), Inches(1.0), "🕵️", size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.9), Inches(12), Inches(0.9),
       "The Mystery of the Missing Treasure", size=42, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.85), Inches(12), Inches(0.5),
       "Can You Solve the Mystery?", size=21, color=GOLD, align=PP_ALIGN.CENTER,
       italic=True)
    for i, (emoji, label) in enumerate(SCENE):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(3.65), Inches(1.65), Inches(1.4),
                  RGBColor(0x2A, 0x2C, 0x4A))
        tb(slide, left, Inches(3.85), Inches(1.65), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.5), Inches(1.45), Inches(0.42), label,
           size=10, color=RGBColor(0xC8, 0xCF, 0xE2), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.55), Inches(5.35), Inches(6.2), Inches(1.0), CRIMSON)
    tb(slide, Inches(3.55), Inches(5.58), Inches(6.2), Inches(0.6),
       "Grade 5  •  90 Minutes  •  Reading + Grammar", size=17, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Briefing")
    fade(slide)


def s02_you_are_detective():
    slide, n = new_slide("🕵️ Today You Are the Detective", "BRIEFING", "0–7 min",
                         "Briefing", CRIMSON)
    one_task(slide, "Here is your case. Read it with me.", CRIMSON)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(6.0), Inches(4.5), L_CRIMSON)
    tb(slide, Inches(0.5), Inches(2.4), Inches(6.0), Inches(1.4), "🔍", size=76,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.0), Inches(5.6), Inches(0.6), "THE CASE FILE",
       size=26, bold=True, color=CRIMSON, align=PP_ALIGN.CENTER, font="Arial Black")
    tb(slide, Inches(0.9), Inches(4.75), Inches(5.2), Inches(1.2),
       "A treasure has gone missing. The only thing left behind is a very old map — "
       "and it is full of clues.", size=15, color=DARK, align=PP_ALIGN.CENTER)
    facts = [("🎯", "Your goal", "Solve the mystery by the end of class."),
             ("🧠", "Your tools", "Reading, vocabulary, and grammar."),
             ("🤝", "Your partner", "Me. You never work a clue alone."),
             ("⭐", "Your reward", "Five clue badges and the final answer.")]
    for i, (icon, label, detail) in enumerate(facts):
        top = Inches(1.9 + i * 1.15)
        add_round(slide, Inches(6.85), top, Inches(6.0), Inches(1.0), L_SLATE)
        add_oval(slide, Inches(7.1), top + Inches(0.22), Inches(0.56), Inches(0.56),
                 WHITE)
        tb(slide, Inches(7.1), top + Inches(0.28), Inches(0.56), Inches(0.44), icon,
           size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.85), top + Inches(0.12), Inches(4.8), Inches(0.42), label,
           size=16, bold=True, color=SLATE)
        tb(slide, Inches(7.85), top + Inches(0.54), Inches(4.8), Inches(0.38), detail,
           size=13, color=DARK)
    hint(slide, "Read the case file out loud together before moving on.", 6.42,
         "🗣️ SPEAKING")


def s03_mission():
    slide, n = new_slide("🎖️ Your Detective Mission — 5 Clue Badges", "MISSION",
                         "0–7 min", "Briefing", GOLD)
    one_task(slide, "Collect all five badges to unlock the final answer.", GOLD)
    for i, (emoji, name, detail, color, light) in enumerate(BADGES):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.9), Inches(2.3), Inches(3.1), light)
        tb(slide, left, Inches(2.15), Inches(2.3), Inches(0.8), emoji, size=34,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.0), Inches(2.1), Inches(0.9), name,
           size=16, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.85), Inches(2.0), Inches(0.7), detail,
           size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_oval(slide, left + Inches(0.9), Inches(4.5), Inches(0.5), Inches(0.5), WHITE)
        tb(slide, left + Inches(0.9), Inches(4.58), Inches(0.5), Inches(0.36), "☐",
           size=14, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.2), Inches(12.35), Inches(1.1), L_SLATE)
    tb(slide, Inches(0.8), Inches(5.32), Inches(11.7), Inches(0.36),
       "TODAY'S ROUTE", size=12, bold=True, color=SLATE)
    tb(slide, Inches(0.8), Inches(5.68), Inches(11.7), Inches(0.48),
       "READ → PLAY → GRAMMAR → SPEAK → READ → MOVE → GRAMMAR → GAME → STORY → CHALLENGE",
       size=14, bold=True, color=INK)
    hint(slide, "Tick a badge together the moment she earns it. Say it out loud.", 6.45,
         "🏅 REWARD")


def s04_scene():
    slide, n = new_slide("🔍 Warm-Up: The Mystery Scene", "WARM-UP", "0–7 min",
                         "Reading Clue", TEAL)
    one_task(slide, "Look closely. What can you find?", TEAL)
    for i, (emoji, label) in enumerate(SCENE):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_TEAL)
        tb(slide, left, top + Inches(0.25), Inches(3.9), Inches(0.95), emoji, size=46,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.75), top + Inches(1.35), Inches(2.4),
                  Inches(0.58), WHITE)
        tb(slide, left + Inches(0.75), top + Inches(1.46), Inches(2.4), Inches(0.42),
           label, size=15, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    hint(slide, "Ask \"what do you see?\" and wait. Let her name things in any order.",
         6.42, "🗣️ SPEAKING")


def s05_starters():
    slide, n = new_slide("🗣️ Talk Like a Detective", "SPEAKING", "0–7 min",
                         "Reading Clue", TEAL)
    one_task(slide, "Answer in a full sentence. Use one of these starters.", TEAL)
    for i, starter in enumerate(STARTERS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.45)
        add_round(slide, left, top, Inches(5.95), Inches(1.25), L_GOLD)
        tb(slide, left + Inches(0.4), top + Inches(0.32), Inches(5.2), Inches(0.62),
           starter, size=24, bold=True, color=INK)
    add_round(slide, Inches(0.5), Inches(4.9), Inches(12.35), Inches(1.4), L_SLATE)
    tb(slide, Inches(0.8), Inches(5.02), Inches(11.7), Inches(0.38),
       "WHY FULL SENTENCES MATTER", size=12, bold=True, color=SLATE)
    tb(slide, Inches(0.8), Inches(5.42), Inches(11.7), Inches(0.8),
       "A detective explains her thinking. One-word answers hide the reasoning — full "
       "sentences show it.", size=15, color=DARK)
    hint(slide, "If she answers with one word, repeat it back as a full sentence first.",
         6.42)


def s06_predict():
    slide, n = new_slide("🤔 Detective Talk — What Do You Think?", "SPEAKING", "0–7 min",
                         "Reading Clue", TEAL)
    one_task(slide, "There are no wrong answers here. Just good thinking.", TEAL)
    for i, (emoji, question) in enumerate(PREDICT):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), L_TEAL)
        add_oval(slide, Inches(0.8), top + Inches(0.22), Inches(0.58), Inches(0.58),
                 WHITE)
        tb(slide, Inches(0.8), top + Inches(0.28), Inches(0.58), Inches(0.44), emoji,
           size=16, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.65), top + Inches(0.26), Inches(10.5), Inches(0.5), question,
           size=20, bold=True, color=INK)
    add_round(slide, Inches(0.5), Inches(6.42), Inches(12.35), Inches(0.46), L_EMERALD)
    tb(slide, Inches(0.8), Inches(6.47), Inches(4.4), Inches(0.36),
       "🏅 BADGE 1 EARNED — Reading Clue", size=13, bold=True, color=EMERALD)
    tb(slide, Inches(5.3), Inches(6.48), Inches(7.2), Inches(0.34),
       "\"Good thinking!\" works even when the guess turns out to be wrong.", size=12,
       color=INK)


def s07_vocab_a():
    slide, n = new_slide("🧩 Mystery Vocabulary — Part 1", "VOCABULARY", "7–17 min",
                         "Vocabulary Clue", PURPLE)
    one_task(slide, "Four words every detective needs.", PURPLE)
    vocab_cards(slide, VOCAB_A)
    hint(slide, "Read the word, then the meaning, then the example. Always that order.",
         6.5)


def s08_vocab_b():
    slide, n = new_slide("🧩 Mystery Vocabulary — Part 2", "VOCABULARY", "7–17 min",
                         "Vocabulary Clue", PURPLE)
    one_task(slide, "Four more. These all show up in today's story.", PURPLE)
    vocab_cards(slide, VOCAB_B)
    hint(slide, "Ask her to spot which two words she has heard before. Start there.",
         6.5)


def s09_word_routine():
    slide, n = new_slide("🔁 How We Learn Each Word", "METHOD", "7–17 min",
                         "Vocabulary Clue", AMBER)
    one_task(slide, "Same three steps for every single word.", AMBER)
    for i, (num, label, detail, color) in enumerate(WORD_ROUTINE):
        left = Inches(0.5 + i * 4.15)
        light = [L_CRIMSON, L_AMBER, L_EMERALD][i]
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(2.9), light)
        add_oval(slide, left + Inches(1.6), Inches(2.2), Inches(0.7), Inches(0.7), color)
        tb(slide, left + Inches(1.6), Inches(2.32), Inches(0.7), Inches(0.46), num,
           size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.05), Inches(3.5), Inches(0.6), label,
           size=26, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.3), Inches(3.75), Inches(3.3), Inches(0.9),
                  WHITE)
        tb(slide, left + Inches(0.45), Inches(3.95), Inches(3.0), Inches(0.6), detail,
           size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.1), Inches(12.35), Inches(1.2), L_SLATE)
    tb(slide, Inches(0.8), Inches(5.22), Inches(11.7), Inches(0.38),
       "THE FOURTH STEP THAT MATTERS MOST", size=12, bold=True, color=SLATE)
    tb(slide, Inches(0.8), Inches(5.6), Inches(11.7), Inches(0.5),
       "She builds her own sentence with the word. That is when it actually sticks.",
       size=15, color=DARK)
    hint(slide, "Do not skip step 1. Hearing it correctly first prevents mispronunciation.",
         6.45)


def s10_reveal_a():
    slide, n = new_slide("🔎 Game: Read & Reveal", "GAME", "7–17 min", "Reading Clue",
                         CRIMSON)
    one_task(slide, "Read the word correctly and the next clue unlocks.", CRIMSON)
    reveal_rows(slide, REVEAL_A, 1, CRIMSON, L_CRIMSON)
    hint(slide, "Stuck on a long word? Cover all but the first syllable.", 6.6)


def s11_reveal_b():
    slide, n = new_slide("🔎 Read & Reveal — Clues 5 to 8", "GAME", "7–17 min",
                         "Reading Clue", CRIMSON)
    one_task(slide, "Four more words. Then the vocabulary badge is yours.", CRIMSON)
    reveal_rows(slide, REVEAL_B, 5, CRIMSON, L_CRIMSON)
    hint(slide, "For each word ask: what does it mean, and can you use it in a sentence?",
         6.6)


def s12_pos_intro():
    slide, n = new_slide("✏️ The Four Word Detectives", "GRAMMAR", "17–27 min",
                         "Grammar Clue", EMERALD)
    one_task(slide, "Every word in a sentence has a job. Here are four of them.",
             EMERALD)
    for i, (name, job, examples, color, light) in enumerate(POS):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(3.6), light)
        add_round(slide, left + Inches(0.3), Inches(2.2), Inches(2.3), Inches(0.72),
                  color)
        tb(slide, left + Inches(0.3), Inches(2.33), Inches(2.3), Inches(0.5), name,
           size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.1), Inches(2.5), Inches(0.9), job,
           size=14, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(4.1), Inches(2.4), Inches(1.2),
                  WHITE)
        tb(slide, left + Inches(0.35), Inches(4.28), Inches(2.2), Inches(0.9), examples,
           size=13, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Colors stay the same all lesson: red noun, green verb, purple adjective, "
                "teal adverb.", 6.45, "🎨 COLOR CODE")


def s13_pos_demo():
    slide, n = new_slide("🎨 Watch Me Break a Sentence Apart", "GRAMMAR", "17–27 min",
                         "Grammar Clue", EMERALD)
    one_task(slide, "I DO — you just watch this one.", EMERALD)
    add_round(slide, Inches(0.5), Inches(1.9), Inches(12.35), Inches(1.5), WHITE)
    colored_sentence(slide, Inches(0.8), Inches(2.25), Inches(11.75), Inches(0.9),
                     DEMO_SENTENCE, size=32, align=PP_ALIGN.CENTER)
    answers = [("clever", "ADJECTIVE", PURPLE, L_PURPLE),
               ("detective", "NOUN", CRIMSON, L_CRIMSON),
               ("quickly", "ADVERB", TEAL, L_TEAL),
               ("opened", "VERB", EMERALD, L_EMERALD)]
    for i, (word, label, color, light) in enumerate(answers):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(3.6), Inches(2.9), Inches(1.75), light)
        tb(slide, left + Inches(0.1), Inches(3.8), Inches(2.7), Inches(0.55), word,
           size=22, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.38), Inches(2.7), Inches(0.35), "↓",
           size=14, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.4), Inches(4.75), Inches(2.1), Inches(0.45),
                  color)
        tb(slide, left + Inches(0.4), Inches(4.82), Inches(2.1), Inches(0.34), label,
           size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.55), Inches(12.35), Inches(0.72), L_GREY)
    tb(slide, Inches(0.8), Inches(5.7), Inches(11.7), Inches(0.44),
       "Find the verb first. Once you know the action, everything else falls into place.",
       size=15, bold=True, color=SLATE)
    hint(slide, "Adverbs usually end in -ly. That single tip solves most adverb questions.",
         6.45)


def s14_pos_more():
    slide, n = new_slide("🎨 Two More — We Do These Together", "GRAMMAR", "17–27 min",
                         "Grammar Clue", EMERALD)
    one_task(slide, "WE DO — say the colors out loud with me.", EMERALD)
    for i, words in enumerate(DEMO_EXTRA):
        top = Inches(1.95 + i * 1.65)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.4), WHITE)
        colored_sentence(slide, Inches(0.85), top + Inches(0.38), Inches(11.65),
                         Inches(0.75), words, size=27, align=PP_ALIGN.CENTER)
    legend = [("NOUN", CRIMSON), ("VERB", EMERALD), ("ADJECTIVE", PURPLE),
              ("ADVERB", TEAL)]
    for i, (label, color) in enumerate(legend):
        left = Inches(0.5 + i * 3.13)
        add_round(slide, left, Inches(5.35), Inches(2.9), Inches(0.55), color)
        tb(slide, left, Inches(5.44), Inches(2.9), Inches(0.4), label, size=13, bold=True,
           color=WHITE, align=PP_ALIGN.CENTER)
    ido_wedo_youdo(slide, 6.05)
    hint(slide, "Ask her to point at a word and name its job before you say anything.",
         6.62)


def s15_gd_a():
    slide, n = new_slide("🕵️ Game: Grammar Detective", "GAME", "17–27 min",
                         "Grammar Clue", EMERALD)
    one_task(slide, "YOU DO — find all four jobs in each sentence.", EMERALD)
    find_rows(slide, GD_A, 1)
    hint(slide, "Always start with the verb. Ask: what is somebody doing?", 6.6)


def s16_gd_b():
    slide, n = new_slide("🕵️ Grammar Detective — Cases 4 & 5", "GAME", "17–27 min",
                         "Grammar Clue", EMERALD)
    one_task(slide, "Two tougher sentences. Take your time.", EMERALD)
    find_rows(slide, GD_B, 4)
    add_round(slide, Inches(0.5), Inches(5.0), Inches(12.35), Inches(1.28), L_TEAL)
    tb(slide, Inches(0.8), Inches(5.12), Inches(11.7), Inches(0.38),
       "🔵 BONUS MISSION", size=13, bold=True, color=TEAL)
    tb(slide, Inches(0.8), Inches(5.52), Inches(11.7), Inches(0.62),
       "Write your own sentence with one noun, one verb, one adjective and one adverb. "
       "Then color-code it.", size=16, bold=True, color=INK)
    hint(slide, "Two adjectives can sit in one sentence. Find both before you move on.",
         6.45)


def s17_phrases():
    slide, n = new_slide("🔎 Phrase Reading — Level 2", "READING", "27–35 min",
                         "Reading Clue", CRIMSON)
    one_task(slide, "Read each phrase smoothly, as one piece.", CRIMSON)
    for i, (phrase, emoji) in enumerate(PHRASES):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.42)
        add_round(slide, left, top, Inches(5.95), Inches(1.22), L_CRIMSON)
        tb(slide, left + Inches(0.3), top + Inches(0.3), Inches(0.9), Inches(0.66),
           emoji, size=26, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.35), top + Inches(0.3), Inches(4.4), Inches(0.62),
           phrase, size=24, bold=True, color=INK)
    hint(slide, "A phrase is not a sentence. No pause in the middle — read it in one "
                "breath.", 6.42)


def s18_clue_or_not():
    slide, n = new_slide("🧩 Clue or Not a Clue?", "VOCABULARY", "27–35 min",
                         "Vocabulary Clue", PURPLE)
    one_task(slide, "Which of these words belong to our mystery?", PURPLE)
    for i, (word, _is_clue) in enumerate(CLUE_OR_NOT):
        col, row = i % 4, i // 4
        left = Inches(0.5 + col * 3.13)
        top = Inches(2.0 + row * 1.95)
        add_round(slide, left, top, Inches(2.9), Inches(1.7), WHITE)
        tb(slide, left, top + Inches(0.3), Inches(2.9), Inches(0.6), word, size=26,
           bold=True, color=INK, align=PP_ALIGN.CENTER, font="Arial Black")
        add_round(slide, left + Inches(0.45), top + Inches(1.0), Inches(0.9),
                  Inches(0.5), L_EMERALD)
        tb(slide, left + Inches(0.45), top + Inches(1.08), Inches(0.9), Inches(0.36),
           "CLUE", size=11, bold=True, color=EMERALD, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.55), top + Inches(1.0), Inches(0.9),
                  Inches(0.5), L_GREY)
        tb(slide, left + Inches(1.55), top + Inches(1.08), Inches(0.9), Inches(0.36),
           "NOT", size=11, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    hint(slide, "Ask why, not just which. \"Why is APPLE not a clue?\" is the real "
                "question.", 6.42)


def s19_word_game():
    slide, n = new_slide("🧩 Game: Word Clue Game", "GAME", "27–35 min",
                         "Vocabulary Clue", PURPLE)
    one_task(slide, "Score points three ways on every mystery word.", PURPLE)
    for i, (pts, task, color) in enumerate(WORD_POINTS):
        left = Inches(0.5 + i * 4.15)
        light = [L_CRIMSON, L_AMBER, L_EMERALD][i]
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(2.5), light)
        add_round(slide, left + Inches(1.3), Inches(2.2), Inches(1.3), Inches(0.62),
                  color)
        tb(slide, left + Inches(1.3), Inches(2.3), Inches(1.3), Inches(0.44), pts,
           size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(3.05), Inches(3.3), Inches(1.1), task,
           size=17, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(4.7), Inches(12.35), Inches(1.6), L_SLATE)
    tb(slide, Inches(0.8), Inches(4.85), Inches(11.7), Inches(0.38),
       "TODAY'S SCORE SHEET", size=12, bold=True, color=SLATE)
    for i, (word, emoji) in enumerate(WORD_CHALLENGE):
        left = Inches(0.8 + i * 3.05)
        add_round(slide, left, Inches(5.3), Inches(2.85), Inches(0.8), WHITE)
        tb(slide, left + Inches(0.15), Inches(5.46), Inches(0.6), Inches(0.5), emoji,
           size=17, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.8), Inches(5.46), Inches(1.9), Inches(0.5), word,
           size=15, bold=True, color=INK)
    hint(slide, "Six points on one word beats one point on six words. Slow down and go "
                "deep.", 6.45)


def s20_word_challenge():
    slide, n = new_slide("⭐ Word Challenge — Build a Sentence", "SPEAKING", "27–35 min",
                         "Vocabulary Clue", PURPLE)
    one_task(slide, "Say a full sentence using each word. Out loud.", PURPLE)
    for i, (word, emoji) in enumerate(WORD_CHALLENGE):
        top = Inches(1.95 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.96), WHITE)
        tb(slide, Inches(0.85), top + Inches(0.18), Inches(0.8), Inches(0.6), emoji,
           size=22, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(1.85), top + Inches(0.18), Inches(2.9), Inches(0.6),
                  L_PURPLE)
        tb(slide, Inches(1.85), top + Inches(0.27), Inches(2.9), Inches(0.44), word,
           size=20, bold=True, color=PURPLE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(5.1), top + Inches(0.27), Inches(7.4), Inches(0.44),
           "____________________________________________", size=16, color=SOFT)
    add_round(slide, Inches(0.5), Inches(6.5), Inches(12.35), Inches(0.42), L_EMERALD)
    tb(slide, Inches(0.8), Inches(6.55), Inches(11.7), Inches(0.34),
       "🏅 BADGE 2 EARNED — Vocabulary Clue", size=14, bold=True, color=EMERALD)
    hint(slide, "Stuck? Offer a sentence starter: \"The ancient ______ was...\"", 5.95)


def s21_subj_pred():
    slide, n = new_slide("✏️ Subject & Predicate", "GRAMMAR", "35–42 min",
                         "Grammar Clue", CRIMSON)
    one_task(slide, "Every sentence splits into exactly two parts.", CRIMSON)
    halves = [("SUBJECT", "Who or what is the sentence about?", "The detective",
               CRIMSON, L_CRIMSON),
              ("PREDICATE", "What does the subject do, or what is it?", "found a key.",
               EMERALD, L_EMERALD)]
    for i, (name, question, example, color, light) in enumerate(halves):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(1.95), Inches(5.95), Inches(2.9), light)
        add_round(slide, left + Inches(1.6), Inches(2.2), Inches(2.75), Inches(0.62),
                  color)
        tb(slide, left + Inches(1.6), Inches(2.3), Inches(2.75), Inches(0.44), name,
           size=17, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.4), Inches(3.0), Inches(5.15), Inches(0.6), question,
           size=16, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.6), Inches(3.72), Inches(4.75), Inches(0.85),
                  WHITE)
        tb(slide, left + Inches(0.6), Inches(3.9), Inches(4.75), Inches(0.55), example,
           size=24, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.05), Inches(12.35), Inches(1.22), WHITE)
    tb(slide, Inches(0.8), Inches(5.18), Inches(11.7), Inches(0.36),
       "PUT THEM BACK TOGETHER", size=12, bold=True, color=SLATE)
    tb(slide, Inches(0.8), Inches(5.56), Inches(11.7), Inches(0.58),
       "The detective  +  found a key.   →   The detective found a key.", size=22,
       bold=True, color=INK)
    hint(slide, "Ask \"who or what?\" to find the subject. The rest is the predicate.",
         6.45)


def s22_subj_examples():
    slide, n = new_slide("✏️ Guided Examples — We Do", "GRAMMAR", "35–42 min",
                         "Grammar Clue", CRIMSON)
    one_task(slide, "Split each sentence with me. Where does the line go?", CRIMSON)
    for i, (subject, predicate) in enumerate(SUBJ_PRED):
        top = Inches(2.0 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3), WHITE)
        add_round(slide, Inches(0.85), top + Inches(0.3), Inches(4.6), Inches(0.7),
                  L_CRIMSON)
        tb(slide, Inches(1.0), top + Inches(0.1), Inches(4.3), Inches(0.28), "SUBJECT",
           size=10, bold=True, color=CRIMSON)
        tb(slide, Inches(0.85), top + Inches(0.42), Inches(4.6), Inches(0.5), subject,
           size=20, bold=True, color=CRIMSON, align=PP_ALIGN.CENTER)
        tb(slide, Inches(5.62), top + Inches(0.4), Inches(0.45), Inches(0.5), "|",
           size=22, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(6.25), top + Inches(0.3), Inches(6.35), Inches(0.7),
                  L_EMERALD)
        tb(slide, Inches(6.4), top + Inches(0.1), Inches(6.0), Inches(0.28), "PREDICATE",
           size=10, bold=True, color=EMERALD)
        tb(slide, Inches(6.25), top + Inches(0.42), Inches(6.35), Inches(0.5), predicate,
           size=20, bold=True, color=EMERALD, align=PP_ALIGN.CENTER)
    hint(slide, "The subject can be more than one word. \"Maya and her grandmother\" is "
                "one subject.", 6.55)


def s23_lock_a():
    slide, n = new_slide("🔐 Game: Sentence Lock", "GAME", "35–42 min", "Grammar Clue",
                         PURPLE)
    one_task(slide, "Match each subject to the predicate that unlocks it.", PURPLE)
    lock_rows(slide, LOCK_A, 1)
    hint(slide, "Read the subject, then try each predicate out loud. Your ear will "
                "know.", 6.55)


def s24_lock_b():
    slide, n = new_slide("🔐 Sentence Lock — Final Two", "GAME", "35–42 min",
                         "Grammar Clue", PURPLE)
    one_task(slide, "Two left. Then write one of your own.", PURPLE)
    lock_rows(slide, LOCK_B, 4)
    add_round(slide, Inches(0.5), Inches(5.0), Inches(12.35), Inches(1.28), L_TEAL)
    tb(slide, Inches(0.8), Inches(5.12), Inches(11.7), Inches(0.38), "🔵 BONUS MISSION",
       size=13, bold=True, color=TEAL)
    tb(slide, Inches(0.8), Inches(5.55), Inches(11.7), Inches(0.58),
       "Your subject: __________________   +   Your predicate: __________________",
       size=18, bold=True, color=INK)
    hint(slide, "Check her own sentence together: does it have both halves?", 6.45)


def s25_brain_break():
    slide, n = new_slide("🏃 Brain Break: Find the Clue!", "BREAK", "42–47 min",
                         "Movement", AMBER)
    one_task(slide, "Up you get. Find each one, then tell me what you found.", AMBER)
    for i, (icon, task) in enumerate(BREAK_TASKS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.3)
        add_round(slide, left, top, Inches(5.95), Inches(1.1), L_AMBER)
        add_oval(slide, left + Inches(0.28), top + Inches(0.26), Inches(0.58),
                 Inches(0.58), WHITE)
        tb(slide, left + Inches(0.28), top + Inches(0.32), Inches(0.58), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.3), Inches(4.6), Inches(0.52), task,
           size=17, bold=True, color=INK)
    add_round(slide, Inches(0.5), Inches(5.85), Inches(12.35), Inches(0.5), L_EMERALD)
    tb(slide, Inches(0.8), Inches(5.93), Inches(11.7), Inches(0.36),
       "🗣️ Say it in a sentence every time:  \"I found a ______.\"", size=15, bold=True,
       color=EMERALD)
    hint(slide, "Five minutes maximum. Keep the pace fast and end it while it's still "
                "fun.", 6.45, "⏱️ TIMING")


def s26_timeline():
    slide, n = new_slide("⏰ The Time Machine — Past, Now, Future", "GRAMMAR",
                         "47–58 min", "Grammar Clue", PURPLE)
    one_task(slide, "The same action, told at three different times.", PURPLE)
    add_rect(slide, Inches(1.0), Inches(2.08), Inches(11.3), Inches(0.06), SLATE)
    for i, (label, when, example, color, light) in enumerate(TIMELINE):
        left = Inches(0.6 + i * 4.1)
        add_oval(slide, left + Inches(1.65), Inches(1.88), Inches(0.46), Inches(0.46),
                 color)
        add_round(slide, left, Inches(2.55), Inches(3.8), Inches(3.3), light)
        tb(slide, left + Inches(0.2), Inches(2.75), Inches(3.4), Inches(0.6), label,
           size=26, bold=True, color=color, align=PP_ALIGN.CENTER, font="Arial Black")
        tb(slide, left + Inches(0.2), Inches(3.4), Inches(3.4), Inches(0.45), when,
           size=13, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(3.95), Inches(3.3), Inches(1.65),
                  WHITE)
        tb(slide, left + Inches(0.4), Inches(4.2), Inches(3.0), Inches(1.2), example,
           size=15, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Look at the verb ending. -ed means past. \"will\" means future.", 6.45)


def s27_same_sentence():
    slide, n = new_slide("⏰ Same Sentence, Three Times", "GRAMMAR", "47–58 min",
                         "Grammar Clue", PURPLE)
    one_task(slide, "Only the verb changes. Everything else stays.", PURPLE)
    rows = [("PRESENT", "The detective searches the room.", "searches", EMERALD,
             L_EMERALD),
            ("PAST", "The detective searched the room.", "searched", CRIMSON,
             L_CRIMSON),
            ("FUTURE", "The detective will search the room.", "will search", PURPLE,
             L_PURPLE)]
    for i, (label, sentence, verb, color, light) in enumerate(rows):
        top = Inches(2.0 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3), light)
        add_round(slide, Inches(0.85), top + Inches(0.34), Inches(2.0), Inches(0.62),
                  color)
        tb(slide, Inches(0.85), top + Inches(0.44), Inches(2.0), Inches(0.44), label,
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(3.15), top + Inches(0.36), Inches(6.4), Inches(0.6), sentence,
           size=21, bold=True, color=INK)
        add_round(slide, Inches(9.75), top + Inches(0.34), Inches(2.85), Inches(0.62),
                  WHITE)
        tb(slide, Inches(9.75), top + Inches(0.44), Inches(2.85), Inches(0.44), verb,
           size=17, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Cover the verb column and ask her to say each version from memory.",
         6.55)


def s28_tense_a():
    slide, n = new_slide("⏰ Game: Time Travel Detective", "GAME", "47–58 min",
                         "Grammar Clue", PURPLE)
    one_task(slide, "Read the sentence. Which time is it?", PURPLE)
    tense_rows(slide, TENSE_A, 1)
    hint(slide, "Point at the verb and only the verb. That word holds the answer.", 6.7)


def s29_tense_b():
    slide, n = new_slide("⏰ Time Travel — Cases 4 & 5", "GAME", "47–58 min",
                         "Grammar Clue", PURPLE)
    one_task(slide, "Two more, then we start changing them.", PURPLE)
    tense_rows(slide, TENSE_B, 4)
    add_round(slide, Inches(0.5), Inches(5.1), Inches(12.35), Inches(1.2), L_SLATE)
    tb(slide, Inches(0.8), Inches(5.22), Inches(11.7), Inches(0.38),
       "QUICK CHECK — WHAT GIVES IT AWAY?", size=12, bold=True, color=SLATE)
    tb(slide, Inches(0.8), Inches(5.6), Inches(11.7), Inches(0.5),
       "PAST: verb ends in -ed (or changes shape)   ·   PRESENT: verb often ends in -s   "
       "·   FUTURE: the word \"will\"", size=14, bold=True, color=INK)
    hint(slide, "\"Discovered\" and \"found\" are both past, even though they look "
                "different.", 6.45)


def s30_transform():
    slide, n = new_slide("🔄 Transform the Sentence", "GRAMMAR", "47–58 min",
                         "Grammar Clue", TEAL)
    one_task(slide, "I give you the present. You give me the past and the future.", TEAL)
    headers = [("PRESENT", EMERALD), ("PAST", CRIMSON), ("FUTURE", PURPLE)]
    for j, (label, color) in enumerate(headers):
        left = Inches(0.5 + j * 4.15)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(0.5), color)
        tb(slide, left, Inches(1.98), Inches(3.9), Inches(0.36), label, size=13,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, (present, _past, _future) in enumerate(TRANSFORM):
        top = Inches(2.55 + i * 1.3)
        add_round(slide, Inches(0.5), top, Inches(3.9), Inches(1.1), L_EMERALD)
        tb(slide, Inches(0.65), top + Inches(0.2), Inches(3.6), Inches(0.75), present,
           size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
        for j in range(2):
            left = Inches(4.65 + j * 4.15)
            add_round(slide, left, top, Inches(3.9), Inches(1.1), WHITE)
            tb(slide, left, top + Inches(0.34), Inches(3.9), Inches(0.44),
               "________________________", size=14, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.5), Inches(12.35), Inches(0.42), L_EMERALD)
    tb(slide, Inches(0.8), Inches(6.55), Inches(11.7), Inches(0.34),
       "🏅 BADGE 3 EARNED — Grammar Clue", size=14, bold=True, color=EMERALD)
    hint(slide, "Say it out loud before writing. The ear catches \"finded\" faster than "
                "the eye.", 6.0)


def s31_agreement():
    slide, n = new_slide("🤝 Subject-Verb Agreement", "GRAMMAR", "47–58 min",
                         "Grammar Clue", CRIMSON)
    one_task(slide, "One detective finds. Many detectives find. Odd, but true.", CRIMSON)
    for i, (count, subject, verb, rule, color, light) in enumerate(AGREEMENT_RULE):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(1.95), Inches(5.95), Inches(3.0), light)
        add_round(slide, left + Inches(2.1), Inches(2.2), Inches(1.75), Inches(0.58),
                  color)
        tb(slide, left + Inches(2.1), Inches(2.29), Inches(1.75), Inches(0.42), count,
           size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.4), Inches(2.95), Inches(5.15), Inches(0.85),
                  WHITE)
        tb(slide, left + Inches(0.5), Inches(3.14), Inches(2.6), Inches(0.5),
           f"The {subject}", size=19, bold=True, color=SLATE, align=PP_ALIGN.RIGHT)
        tb(slide, left + Inches(3.25), Inches(3.14), Inches(2.2), Inches(0.5), verb,
           size=19, bold=True, color=color)
        tb(slide, left + Inches(0.4), Inches(4.0), Inches(5.15), Inches(0.75), rule,
           size=14, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.15), Inches(12.35), Inches(1.1), L_SLATE)
    tb(slide, Inches(0.8), Inches(5.27), Inches(11.7), Inches(0.38),
       "THE TRICK THAT ALWAYS WORKS", size=12, bold=True, color=SLATE)
    tb(slide, Inches(0.8), Inches(5.65), Inches(11.7), Inches(0.48),
       "If the subject has an -s, the verb usually does not. Only one of them gets the s.",
       size=15, bold=True, color=INK)
    hint(slide, "Find the subject first, every time. The verb just follows it.", 6.4)


def s32_agreement_game():
    slide, n = new_slide("🤝 Game: Agreement Check", "GAME", "47–58 min", "Grammar Clue",
                         CRIMSON)
    one_task(slide, "Pick the verb that agrees. Say the whole sentence out loud.",
             CRIMSON)
    tops = numbered_rows(slide, len(AGREEMENT), 1, CRIMSON, top_start=1.9, gap=0.92,
                         height=0.8)
    for (sentence, options, _answer), top in zip(AGREEMENT, tops):
        tb(slide, Inches(1.45), top + Inches(0.17), Inches(7.3), Inches(0.5), sentence,
           size=18, bold=True, color=INK)
        add_round(slide, Inches(9.0), top + Inches(0.13), Inches(3.6), Inches(0.54),
                  L_CRIMSON)
        tb(slide, Inches(9.0), top + Inches(0.22), Inches(3.6), Inches(0.4), options,
           size=16, bold=True, color=CRIMSON, align=PP_ALIGN.CENTER)
    hint(slide, "\"The clues\" is plural even though \"door\" comes after it. Find the "
                "real subject.", 6.55)


def s33_sentence_reading():
    slide, n = new_slide("🔎 Sentence Reading — Level 3", "READING", "58–68 min",
                         "Reading Clue", CRIMSON)
    one_task(slide, "Last warm-up before the story. Read each one smoothly.", CRIMSON)
    for i, (sentence, emoji) in enumerate(SENTENCES_L3):
        top = Inches(1.95 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.96), WHITE)
        tb(slide, Inches(0.85), top + Inches(0.18), Inches(0.8), Inches(0.6), emoji,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.85), top + Inches(0.2), Inches(8.0), Inches(0.58), sentence,
           size=22, bold=True, color=INK)
        add_round(slide, Inches(10.1), top + Inches(0.2), Inches(2.5), Inches(0.56),
                  L_CRIMSON)
        tb(slide, Inches(10.1), top + Inches(0.29), Inches(2.5), Inches(0.4),
           "I · WE · YOU", size=13, bold=True, color=CRIMSON, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.5), Inches(12.35), Inches(0.42), L_SLATE)
    tb(slide, Inches(0.8), Inches(6.55), Inches(11.7), Inches(0.34),
       "Every one of these sentences comes from the story you are about to read.",
       size=13, bold=True, color=SLATE)
    hint(slide, "If she reads it choppily, model it once smoothly and let her try again.",
         6.0)


def s34_story_intro():
    slide, n = new_slide("📖 The Secret Behind the Old Door", "STORY", "58–68 min",
                         "Story Clue", TEAL)
    one_task(slide, "Four short parts. Watch for these six words as we read.", TEAL)
    add_round(slide, Inches(0.5), Inches(1.9), Inches(4.5), Inches(4.4), L_TEAL)
    tb(slide, Inches(0.5), Inches(2.6), Inches(4.5), Inches(1.5), "🚪", size=82,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.35), Inches(4.1), Inches(0.9),
       "A map. A locked door.\nOne missing key.", size=18, bold=True, color=TEAL,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(5.35), Inches(3.9), Inches(0.7),
       "Roughly 320 words, in four parts.", size=13, color=SOFT, align=PP_ALIGN.CENTER)
    for i, (word, meaning) in enumerate(STORY_VOCAB):
        col, row = i % 2, i // 2
        left = Inches(5.3 + col * 3.85)
        top = Inches(1.9 + row * 1.5)
        add_round(slide, left, top, Inches(3.6), Inches(1.3), WHITE)
        tb(slide, left + Inches(0.2), top + Inches(0.16), Inches(3.2), Inches(0.44),
           word, size=18, bold=True, color=PURPLE)
        tb(slide, left + Inches(0.2), top + Inches(0.62), Inches(3.2), Inches(0.58),
           meaning, size=12, color=DARK)
    hint(slide, "Read these six words now. Meeting them before the story removes the "
                "roadblocks.", 6.45)


def _story_slide(index, timing):
    part, emoji, color, light, paras = STORY[index]
    slide, n = new_slide(f"📖 The Secret Behind the Old Door — {part}", "STORY", timing,
                         "Story Clue", color)
    add_round(slide, Inches(0.5), Inches(1.45), Inches(3.1), Inches(4.75), light)
    tb(slide, Inches(0.5), Inches(2.95), Inches(3.1), Inches(1.6), emoji, size=88,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.85), Inches(1.45), Inches(9.0), Inches(4.75), WHITE)
    tops = [1.75, 3.2, 4.7]
    for text, t in zip(paras, tops):
        para(slide, Inches(4.2), Inches(t), Inches(8.3), Inches(1.4), text, size=17,
             line=1.25)
    ido_wedo_youdo(slide, 6.42)
    return slide, n


def s35_story_1():
    _story_slide(0, "58–68 min")


def s36_story_2():
    _story_slide(1, "58–68 min")


def s37_story_3():
    _story_slide(2, "58–68 min")


def s38_story_4():
    _story_slide(3, "58–68 min")


def s39_evidence_method():
    slide, n = new_slide("🔍 Show Me the Evidence!", "METHOD", "68–76 min", "Story Clue",
                         GOLD)
    one_task(slide, "A real detective never guesses. She proves it.", GOLD)
    for i, (num, label, detail, color) in enumerate(EVIDENCE_STEPS):
        left = Inches(0.5 + i * 3.13)
        light = [L_CRIMSON, L_AMBER, L_EMERALD, L_TEAL][i]
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(3.1), light)
        add_oval(slide, left + Inches(1.1), Inches(2.2), Inches(0.7), Inches(0.7), color)
        tb(slide, left + Inches(1.1), Inches(2.32), Inches(0.7), Inches(0.46), num,
           size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.05), Inches(2.6), Inches(0.9), label,
           size=17, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(3.95), Inches(2.4), Inches(0.9),
                  WHITE)
        tb(slide, left + Inches(0.35), Inches(4.14), Inches(2.2), Inches(0.6), detail,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.25), Inches(12.35), Inches(1.0), L_GOLD)
    tb(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.5),
       "\"My answer is ______, and my evidence is the sentence that says ______.\"",
       size=18, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Going back to the text is not cheating. It is exactly the skill being "
                "taught.", 6.45)


def s40_q_literal():
    slide, n = new_slide("📖 Find the Evidence — What Happened", "COMPREHENSION",
                         "68–76 min", "Story Clue", TEAL)
    one_task(slide, "The answers are all written in the story. Go find them.", TEAL)
    tops = numbered_rows(slide, len(Q_LITERAL), 1, TEAL, top_start=2.0, gap=1.45,
                         height=1.25)
    for (question, where), top in zip(Q_LITERAL, tops):
        tb(slide, Inches(1.45), top + Inches(0.18), Inches(8.4), Inches(0.5), question,
           size=20, bold=True, color=INK)
        tb(slide, Inches(1.45), top + Inches(0.72), Inches(8.4), Inches(0.4),
           "Evidence sentence: ______________________________________", size=13,
           color=SOFT)
        add_round(slide, Inches(10.2), top + Inches(0.38), Inches(2.4), Inches(0.5),
                  L_TEAL)
        tb(slide, Inches(10.2), top + Inches(0.46), Inches(2.4), Inches(0.36),
           f"look in {where}", size=12, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    hint(slide, "She should reread, not remember. Send her back to the part every time.",
         6.55)


def s41_q_mixed():
    slide, n = new_slide("📖 Find the Evidence — Words & Thinking", "COMPREHENSION",
                         "68–76 min", "Story Clue", TEAL)
    one_task(slide, "Two word questions, two thinking questions.", TEAL)
    for i, (kind, question, color, light) in enumerate(Q_MIXED):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.98), light)
        add_round(slide, Inches(0.8), top + Inches(0.26), Inches(1.9), Inches(0.5),
                  color)
        tb(slide, Inches(0.8), top + Inches(0.34), Inches(1.9), Inches(0.36), kind,
           size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(3.0), top + Inches(0.28), Inches(9.4), Inches(0.5), question,
           size=19, bold=True, color=INK)
    add_round(slide, Inches(0.5), Inches(6.42), Inches(12.35), Inches(0.46), L_SLATE)
    tb(slide, Inches(0.8), Inches(6.47), Inches(11.7), Inches(0.36),
       "Inference means the answer is not written down — you work it out from what is. "
       "For vocabulary, cover the word and ask what would still make sense there.",
       size=12, bold=True, color=SLATE)


def s42_q_evidence():
    slide, n = new_slide("📖 Prove It — Point to the Sentence", "COMPREHENSION",
                         "68–76 min", "Story Clue", GOLD)
    one_task(slide, "Answer, then read me the exact line that proves it.", GOLD)
    tops = numbered_rows(slide, len(Q_EVIDENCE), 1, GOLD, top_start=2.0, gap=1.75,
                         height=1.55)
    for (question, task), top in zip(Q_EVIDENCE, tops):
        tb(slide, Inches(1.45), top + Inches(0.18), Inches(11.0), Inches(0.5), question,
           size=21, bold=True, color=INK)
        add_round(slide, Inches(1.45), top + Inches(0.76), Inches(11.05), Inches(0.58),
                  L_GOLD)
        tb(slide, Inches(1.7), top + Inches(0.88), Inches(10.6), Inches(0.4), task,
           size=14, bold=True, color=AMBER)
    add_round(slide, Inches(0.5), Inches(5.65), Inches(12.35), Inches(0.62), L_EMERALD)
    tb(slide, Inches(0.8), Inches(5.77), Inches(11.7), Inches(0.42),
       "🏅 BADGE 4 EARNED — Story Clue", size=15, bold=True, color=EMERALD)
    hint(slide, "If she cannot find the line, narrow it down: \"It's in Part 3 "
                "somewhere.\"", 6.45)


def s43_error_types():
    slide, n = new_slide("✏️ Six Things a Detective Checks", "GRAMMAR", "76–84 min",
                         "Grammar Clue", AMBER)
    one_task(slide, "Before you fix anything, know what to look for.", AMBER)
    for i, (icon, name, detail) in enumerate(ERROR_TYPES):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.95 + row * 2.2)
        add_round(slide, left, top, Inches(3.9), Inches(2.0), L_AMBER)
        add_oval(slide, left + Inches(1.6), top + Inches(0.2), Inches(0.7), Inches(0.7),
                 WHITE)
        tb(slide, left + Inches(1.6), top + Inches(0.32), Inches(0.7), Inches(0.5), icon,
           size=17, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), top + Inches(1.0), Inches(3.5), Inches(0.45), name,
           size=16, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), top + Inches(1.45), Inches(3.3), Inches(0.5),
           detail, size=12, color=DARK, align=PP_ALIGN.CENTER)
    hint(slide, "Work down this list in order. Capitals first, then punctuation, then "
                "verbs.", 6.45)


def s44_fix_a():
    slide, n = new_slide("✏️ Game: Fix the Mystery!", "GAME", "76–84 min",
                         "Grammar Clue", AMBER)
    one_task(slide, "Each sentence has more than one mistake. Find them all.", AMBER)
    fix_rows(slide, FIX_A, 1)
    hint(slide, "Read it out loud exactly as written. The errors become obvious.", 6.55)


def s45_fix_b():
    slide, n = new_slide("✏️ Fix the Mystery — Cases 4 & 5", "GAME", "76–84 min",
                         "Grammar Clue", AMBER)
    one_task(slide, "Two final repairs. Say the fixed version out loud.", AMBER)
    fix_rows(slide, FIX_B, 4)
    add_round(slide, Inches(0.5), Inches(5.0), Inches(12.35), Inches(1.28), L_GOLD)
    tb(slide, Inches(0.8), Inches(5.12), Inches(11.7), Inches(0.38),
       "🎖️ DETECTIVE BADGE", size=13, bold=True, color=AMBER)
    tb(slide, Inches(0.8), Inches(5.52), Inches(11.7), Inches(0.62),
       "One badge per correction. Five corrections means five badges — say each one out "
       "loud as she earns it.", size=15, bold=True, color=INK)
    hint(slide, "After \"will,\" the verb stays plain: will find, never will found.",
         6.45)


def s46_writing():
    slide, n = new_slide("📝 The Detective's Notebook", "WRITING", "76–84 min",
                         "Grammar Clue", SLATE)
    one_task(slide, "Three short lines. Full sentences, correct punctuation.", SLATE)
    for i, frame in enumerate(WRITING_FRAMES):
        top = Inches(2.0 + i * 1.4)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.2), WHITE)
        add_oval(slide, Inches(0.8), top + Inches(0.35), Inches(0.5), Inches(0.5), SLATE)
        tb(slide, Inches(0.8), top + Inches(0.41), Inches(0.5), Inches(0.4), str(i + 1),
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.36), Inches(10.9), Inches(0.5), frame,
           size=19, bold=True, color=INK)
    add_round(slide, Inches(0.5), Inches(6.2), Inches(12.35), Inches(0.5), L_SLATE)
    tb(slide, Inches(0.8), Inches(6.28), Inches(11.7), Inches(0.36),
       "Check each line together: capital letter at the start, period at the end.",
       size=13, bold=True, color=SLATE)
    hint(slide, "She may dictate while you write if handwriting slows her down.", 5.6)


def s47_final_board():
    slide, n = new_slide("🏆 Final Game: Grammar Treasure Challenge", "CHALLENGE",
                         "84–88 min", "Final Clue", GOLD)
    one_task(slide, "Five last clues. Crack them all and the mystery is solved.", GOLD)
    for i, (num, emoji, kind, _prompt, _content, color, light) in enumerate(FINAL_CLUES):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.95), Inches(2.3), Inches(3.3), light)
        tb(slide, left, Inches(2.2), Inches(2.3), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.4), Inches(3.05), Inches(1.5), Inches(0.5),
                  color)
        tb(slide, left + Inches(0.4), Inches(3.13), Inches(1.5), Inches(0.36), num,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.7), Inches(2.1), Inches(0.5), kind,
           size=15, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_oval(slide, left + Inches(0.9), Inches(4.45), Inches(0.5), Inches(0.5),
                 WHITE)
        tb(slide, left + Inches(0.9), Inches(4.53), Inches(0.5), Inches(0.36), "☐",
           size=14, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.45), Inches(12.35), Inches(0.8), L_GOLD)
    tb(slide, Inches(0.8), Inches(5.62), Inches(11.7), Inches(0.46),
       "Every clue uses something from today. Nothing here is new.", size=16, bold=True,
       color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Hints are allowed on the final challenge. Finishing matters more than "
                "solo work.", 6.42)


def s48_final_a():
    slide, n = new_slide("🏆 Challenge Clues 1, 2 & 3", "CHALLENGE", "84–88 min",
                         "Final Clue", GOLD)
    one_task(slide, "Work through them in order.", GOLD)
    for i, (num, emoji, kind, prompt, content, color, light) in enumerate(FINAL_CLUES[:3]):
        top = Inches(1.95 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.3), light)
        tb(slide, Inches(0.85), top + Inches(0.32), Inches(0.8), Inches(0.66), emoji,
           size=26, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(1.85), top + Inches(0.2), Inches(1.55), Inches(0.45),
                  color)
        tb(slide, Inches(1.85), top + Inches(0.27), Inches(1.55), Inches(0.34), num,
           size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(3.6), top + Inches(0.24), Inches(8.8), Inches(0.4), prompt,
           size=13, color=SLATE)
        tb(slide, Inches(3.6), top + Inches(0.66), Inches(8.8), Inches(0.52), content,
           size=22, bold=True, color=color)
    hint(slide, "Clue 3 hides two problems: a capital letter and a verb form.", 6.55)


def s49_final_b():
    slide, n = new_slide("🏆 Challenge Clues 4 & 5", "CHALLENGE", "84–88 min",
                         "Final Clue", GOLD)
    one_task(slide, "The last two. Then we open the chest.", GOLD)
    for i, (num, emoji, kind, prompt, content, color, light) in enumerate(FINAL_CLUES[3:]):
        top = Inches(1.95 + i * 1.9)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.65), light)
        tb(slide, Inches(0.9), top + Inches(0.45), Inches(0.9), Inches(0.75), emoji,
           size=30, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.0), top + Inches(0.26), Inches(1.55), Inches(0.45),
                  color)
        tb(slide, Inches(2.0), top + Inches(0.33), Inches(1.55), Inches(0.34), num,
           size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(3.75), top + Inches(0.3), Inches(8.6), Inches(0.4), prompt,
           size=13, color=SLATE)
        tb(slide, Inches(3.75), top + Inches(0.72), Inches(8.6), Inches(0.55), content,
           size=24, bold=True, color=color)
        add_round(slide, Inches(3.75), top + Inches(1.3), Inches(8.6), Inches(0.24),
                  WHITE)
    hint(slide, "Clue 5 is open-ended on purpose. Any correct sentence wins it.", 6.42)


def s50_solved():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    for x, y, c in [(0.7, 1.4, CRIMSON), (11.95, 1.35, TEAL), (0.85, 3.7, PURPLE),
                    (11.85, 3.65, EMERALD)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.8), c)
    tb(slide, Inches(0.7), Inches(0.85), Inches(12), Inches(1.0), "💰", size=52,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.95), Inches(12), Inches(0.9), "🎉 MYSTERY SOLVED!",
       size=44, bold=True, color=GOLD, align=PP_ALIGN.CENTER, font="Georgia")
    add_round(slide, Inches(2.6), Inches(3.05), Inches(8.1), Inches(1.9), CRIMSON)
    tb(slide, Inches(2.85), Inches(3.35), Inches(7.6), Inches(1.4),
       "The treasure was not gold...\n\nThe real treasure was KNOWLEDGE!", size=24,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, (emoji, name, _d, color, _l) in enumerate(BADGES):
        left = Inches(1.6 + i * 2.15)
        add_round(slide, left, Inches(5.3), Inches(1.95), Inches(0.95), color)
        tb(slide, left, Inches(5.42), Inches(1.95), Inches(0.42), f"{emoji} ⭐", size=14,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(5.85), Inches(1.95), Inches(0.32), name, size=9,
           color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "84–88 min", "Final Clue")
    fade(slide)


def s51_exit_ticket():
    slide, n = new_slide("🎫 Exit Ticket", "WRAP-UP", "88–90 min", "Wrap-up", SLATE)
    one_task(slide, "Four quick questions. Say them or write them.", SLATE)
    for i, (num, question, color) in enumerate(EXIT_TICKET):
        top = Inches(2.0 + i * 1.2)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.02), WHITE)
        add_oval(slide, Inches(0.82), top + Inches(0.26), Inches(0.52), Inches(0.52),
                 color)
        tb(slide, Inches(0.82), top + Inches(0.32), Inches(0.52), Inches(0.4), num,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.6), top + Inches(0.16), Inches(10.8), Inches(0.42), question,
           size=18, bold=True, color=INK)
        tb(slide, Inches(1.6), top + Inches(0.6), Inches(10.8), Inches(0.34),
           "_______________________________________________________________", size=12,
           color=SOFT)
    hint(slide, "Question 2 is the only one with a right answer: \"opened.\"", 6.6)


def s52_great_work():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), GOLD)
    tb(slide, Inches(0.7), Inches(1.0), Inches(12), Inches(1.0), "⭐", size=54,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(2.05), Inches(12), Inches(0.9),
       "GREAT WORK, DETECTIVE!", size=42, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Georgia")
    done = [("🔎", "You read"), ("🧩", "You solved clues"), ("✏️", "You fixed grammar"),
            ("📖", "You understood the story"), ("🏆", "You solved the mystery")]
    for i, (emoji, line) in enumerate(done):
        left = Inches(1.15 + i * 2.25)
        add_round(slide, left, Inches(3.4), Inches(2.05), Inches(1.7),
                  RGBColor(0x2A, 0x2C, 0x4A))
        tb(slide, left, Inches(3.6), Inches(2.05), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(4.25), Inches(1.75), Inches(0.7), line,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.1), Inches(5.4), Inches(7.1), Inches(0.95), GOLD)
    tb(slide, Inches(3.1), Inches(5.62), Inches(7.1), Inches(0.55),
       "Same time next case?", size=20, bold=True, color=INK, align=PP_ALIGN.CENTER)
    footer(slide, n, "88–90 min", "Wrap-up")
    fade(slide)


def s53_support():
    slide, n = new_slide("🔒 TEACHER ONLY — Support Levels & Hint Bank", "TEACHER ONLY",
                         "", "Support", DARK)
    tb(slide, Inches(0.45), Inches(1.32), Inches(12.4), Inches(0.34),
       "Move between levels freely within a single activity. Never announce the level "
       "out loud.", size=12, bold=True, color=CRIMSON)
    for i, (icon, name, label, how, script, color, light) in enumerate(SUPPORT_LEVELS):
        left = Inches(0.45 + i * 4.18)
        add_round(slide, left, Inches(1.78), Inches(3.95), Inches(2.85), light)
        tb(slide, left, Inches(1.95), Inches(3.95), Inches(0.5), icon, size=18,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(2.5), Inches(3.65), Inches(0.42), name,
           size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.1), Inches(2.95), Inches(1.75), Inches(0.38),
                  color)
        tb(slide, left + Inches(1.1), Inches(3.01), Inches(1.75), Inches(0.3), label,
           size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(3.42), Inches(3.45), Inches(0.62), how,
           size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(4.05), Inches(3.55), Inches(0.46),
                  WHITE)
        tb(slide, left + Inches(0.3), Inches(4.13), Inches(3.35), Inches(0.34), script,
           size=9, color=SLATE, align=PP_ALIGN.CENTER, italic=True)
    add_round(slide, Inches(0.45), Inches(4.78), Inches(6.1), Inches(2.05), L_AMBER)
    tb(slide, Inches(0.7), Inches(4.9), Inches(5.6), Inches(0.36), "💡 HINT BANK",
       size=13, bold=True, color=AMBER)
    bullets(slide, Inches(0.7), Inches(5.28), Inches(5.6), Inches(1.45), HINT_BANK,
            size=11, sp=3)
    add_round(slide, Inches(6.75), Inches(4.78), Inches(6.1), Inches(2.05), L_EMERALD)
    tb(slide, Inches(7.0), Inches(4.9), Inches(5.6), Inches(0.36),
       "🗣️ SAY THIS INSTEAD OF \"WRONG\"", size=13, bold=True, color=EMERALD)
    bullets(slide, Inches(7.0), Inches(5.28), Inches(5.6), Inches(1.45), PRAISE, size=11,
            sp=3)


def s54_assessment():
    slide, n = new_slide("🔒 TEACHER ONLY — End of Lesson Assessment", "TEACHER ONLY", "",
                         "Assessment", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Mark what she did today, not what you hoped for. Fill this in right after class.",
       size=12, bold=True, color=CRIMSON)
    cols = [(0.45, 4.3, "SKILL"), (4.95, 2.6, "Independent"), (7.75, 2.4, "With Hint"),
            (10.35, 2.5, "Needs More Practice")]
    header_y = Inches(1.7)
    for left, width, label in cols:
        add_round(slide, Inches(left), header_y, Inches(width), Inches(0.42), INK)
        tb(slide, Inches(left), header_y + Inches(0.07), Inches(width), Inches(0.3),
           label, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = header_y + Inches(0.48 + i * 0.38)
        band = WHITE if i % 2 == 0 else L_GREY
        add_round(slide, Inches(0.45), top, Inches(4.3), Inches(0.36), band)
        tb(slide, Inches(0.65), top + Inches(0.04), Inches(3.9), Inches(0.28), skill,
           size=12, bold=True, color=INK)
        for left, width, _l in cols[1:]:
            add_round(slide, Inches(left), top, Inches(width), Inches(0.36), band)
            tb(slide, Inches(left), top + Inches(0.02), Inches(width), Inches(0.3), "☐",
               size=13, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.45), Inches(5.72), Inches(12.4), Inches(1.15), L_AMBER)
    tb(slide, Inches(0.7), Inches(5.82), Inches(11.9), Inches(0.34),
       "NEXT LESSON FOCUS — record two or three areas to practice again", size=12,
       bold=True, color=AMBER)
    for i in range(3):
        tb(slide, Inches(0.7), Inches(6.2 + i * 0.22), Inches(11.9), Inches(0.2),
           f"{i + 1}.  ______________________________________________________________"
           "____________________________", size=10, color=SLATE)


def s55_answer_key():
    slide, n = new_slide("🔒 TEACHER ONLY — Answer Key", "TEACHER ONLY", "", "Answer key",
                         DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Hide this slide before presenting, or keep it on a second screen.", size=12,
       bold=True, color=CRIMSON)
    add_round(slide, Inches(0.45), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(0.7), Inches(1.86), Inches(5.6), Inches(0.36),
       "🧩 Vocabulary  ·  ✏️ Parts of Speech  ·  Subject/Predicate", size=13, bold=True,
       color=PURPLE)
    bullets(slide, Inches(0.7), Inches(2.28), Inches(5.6), Inches(4.4), [
        "Clue or Not: KEY, TREASURE, SECRET, MAP are clues;",
        "     RUN, BLUE, APPLE, JUMP are not",
        "GD 1 brave girl: girl/box · opened · brave, secret · carefully",
        "GD 2 young detective: detective/footprints · followed ·",
        "     young, muddy · silently",
        "GD 3 rusty key: key/shelf · fell · rusty, old · suddenly",
        "GD 4 curious students: students/puzzle · solved ·",
        "     curious, difficult · quickly",
        "GD 5 Maya unfolded: Maya/map · unfolded ·",
        "     ancient, treasure · carefully",
        "Sentence Lock: detective+hidden map · door+opened slowly ·",
        "     students+searched the room · passage+led to the treasure ·",
        "     Maya and grandmother+solved it together",
    ], size=11, sp=4)
    add_round(slide, Inches(6.75), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(7.0), Inches(1.86), Inches(5.6), Inches(0.36),
       "⏰ Tenses  ·  🛠️ Editing  ·  📖 Story  ·  🏆 Final", size=13, bold=True,
       color=EMERALD)
    bullets(slide, Inches(7.0), Inches(2.28), Inches(5.6), Inches(4.4), [
        "Name the Tense: 1 present · 2 past · 3 future ·",
        "     4 present · 5 past",
        "Transform: found / will find · opened / will open ·",
        "     searched / will search",
        "Agreement: finds · find · searches · point · shows",
        "Fix 1: The detective found an old key.",
        "Fix 2: She opened the door slowly.",
        "Fix 3: Maya and her friend searched the forest.",
        "Fix 4: The map shows a secret room.",
        "Fix 5: They will find the treasure tomorrow.",
        "Story: Maya · a wooden box with a map · a stone wall in",
        "     the forest · ancient = extremely old · key was in a",
        "     bird's nest above the door · no, she was delighted",
        "Final: adverb = quietly · \"The door was locked.\" ·",
        "     key in the nest · any correct TREASURE sentence",
    ], size=11, sp=4)


BUILDERS = [
    s01_title, s02_you_are_detective, s03_mission, s04_scene, s05_starters, s06_predict,
    s07_vocab_a, s08_vocab_b, s09_word_routine, s10_reveal_a, s11_reveal_b,
    s12_pos_intro, s13_pos_demo, s14_pos_more, s15_gd_a, s16_gd_b, s17_phrases,
    s18_clue_or_not, s19_word_game, s20_word_challenge, s21_subj_pred, s22_subj_examples,
    s23_lock_a, s24_lock_b, s25_brain_break, s26_timeline, s27_same_sentence, s28_tense_a,
    s29_tense_b, s30_transform, s31_agreement, s32_agreement_game, s33_sentence_reading,
    s34_story_intro, s35_story_1, s36_story_2, s37_story_3, s38_story_4,
    s39_evidence_method, s40_q_literal, s41_q_mixed, s42_q_evidence, s43_error_types,
    s44_fix_a, s45_fix_b, s46_writing, s47_final_board, s48_final_a, s49_final_b,
    s50_solved, s51_exit_ticket, s52_great_work, s53_support, s54_assessment,
    s55_answer_key,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

out = r"C:\Users\bhushaja\Downloads\shaip\Grade5_Mystery_Missing_Treasure_90min.pptx"
try:
    prs.save(out)
except PermissionError:
    out = out.replace(".pptx", "_new.pptx")
    prs.save(out)

story_words = sum(len(p.split()) for _t, _e, _c, _l, paras in STORY for p in paras)
with_notes = [i + 1 for i, s in enumerate(prs.slides)
              if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip()]
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
print(f"Story word count: {story_words}")
print(f"Slides carrying speaker notes: {with_notes if with_notes else 'none'}")
