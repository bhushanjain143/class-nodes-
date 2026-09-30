"""Grade 7 Reading & Grammar 60-min demo PPT — engaging, syllabus-tied games."""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Palette — bright but school-appropriate
NAVY = RGBColor(0x0F, 0x2C, 0x59)
SKY = RGBColor(0x1D, 0x9B, 0xF0)
TEAL = RGBColor(0x0D, 0x94, 0x88)
CORAL = RGBColor(0xF0, 0x73, 0x59)
GOLD = RGBColor(0xF5, 0xB3, 0x01)
PURPLE = RGBColor(0x7C, 0x3A, 0xED)
CREAM = RGBColor(0xFF, 0xFB, 0xF5)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x1F, 0x29, 0x37)
SOFT = RGBColor(0x64, 0x74, 0x8B)
LIGHT_SKY = RGBColor(0xE0, 0xF2, 0xFE)
LIGHT_TEAL = RGBColor(0xCC, 0xFB, 0xF1)
LIGHT_CORAL = RGBColor(0xFF, 0xED, 0xD5)
LIGHT_GOLD = RGBColor(0xFE, 0xF3, 0xC7)
LIGHT_PURPLE = RGBColor(0xED, 0xE9, 0xFE)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
TOTAL = 30

# Shared passage used across reading + games (syllabus cohesion)
PASSAGE_TITLE = "The Signal in the Storm"
PASSAGE = (
    "Rain hammered the windows of Room 214 as the class waited for the power to return. "
    "Ms. Reyes kept her voice calm. \"We still have ten minutes,\" she said, holding up a paper packet. "
    "\"If we finish this reading together, nobody loses progress.\"\n\n"
    "Jordan frowned at the dark projector, then noticed Maya already underlining the first paragraph. "
    "Around them, whispers rose—until someone asked a real question: \"What does the author want us to notice?\" "
    "The room shifted. Pencils moved. Ideas connected.\n\n"
    "When the lights flickered back on, the class had done more than wait. They had practiced the habit "
    "strong readers use every day: ask, find evidence, and revise your thinking. Small choices, made "
    "together, had turned a storm delay into learning."
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


def bullets(slide, left, top, width, height, items, size=16, color=DARK, spacing=8):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(spacing)
        run = p.add_run()
        run.text = "•  " + item
        set_run(run, size=size, color=color)
    return box


def numbered(slide, left, top, width, height, items, size=16, color=DARK, spacing=10):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(spacing)
        run = p.add_run()
        run.text = f"{i + 1}.  {item}"
        set_run(run, size=size, color=color)
    return box


def accent(slide, color=SKY):
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, color)


def footer(slide, page, label=""):
    add_rect(slide, Inches(0), Inches(7.15), prs.slide_width, Inches(0.35), NAVY)
    text = "Grade 7 ELA Demo  |  Reading + Grammar  |  60 min"
    if label:
        text += f"  |  {label}"
    tb(slide, Inches(0.35), Inches(7.18), Inches(11.2), Inches(0.3), text, size=11, color=WHITE)
    tb(slide, Inches(12.0), Inches(7.18), Inches(1.1), Inches(0.3),
       f"{page}/{TOTAL}", size=11, color=WHITE, align=PP_ALIGN.RIGHT)


def chip(slide, text, left=Inches(0.45), top=Inches(0.32), color=SKY, width=Inches(2.7)):
    add_round(slide, left, top, width, Inches(0.38), color)
    tb(slide, left, top + Inches(0.02), width, Inches(0.35),
       text, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def time_chip(slide, text):
    chip(slide, text, left=Inches(10.55), top=Inches(0.32), color=CORAL, width=Inches(2.35))


def game_chip(slide):
    chip(slide, "SYLLABUS GAME", left=Inches(0.45), top=Inches(0.32), color=GOLD, width=Inches(2.7))


def why_game(slide, text, top=Inches(6.55)):
    """Tiny line showing why this game connects to learning."""
    add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.45), LIGHT_GOLD)
    tb(slide, Inches(0.65), top + Inches(0.08), Inches(12.0), Inches(0.35),
       "Why this game?  " + text, size=13, bold=True, color=NAVY)


# ===================== SLIDES =====================

def s_title():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.2), GOLD)
    add_oval(slide, Inches(10.5), Inches(-1.2), Inches(4), Inches(4), RGBColor(0x1A, 0x3A, 0x6E))
    add_oval(slide, Inches(-1.5), Inches(5.0), Inches(3.5), Inches(3.5), RGBColor(0x16, 0x3A, 0x6B))
    tb(slide, Inches(0.8), Inches(1.5), Inches(11.5), Inches(0.4),
       "LIVE DEMO  •  60 MINUTES  •  GRADE 7", size=15, bold=True, color=GOLD)
    tb(slide, Inches(0.8), Inches(2.1), Inches(11.5), Inches(1.3),
       "Reading Content &\nGrammar Level-Up", size=42, bold=True, color=WHITE, font="Georgia")
    tb(slide, Inches(0.8), Inches(4.5), Inches(11.5), Inches(0.4),
       "Learn with a real passage  →  Play syllabus games  →  Walk out stronger",
       size=17, color=RGBColor(0xBF, 0xDB, 0xFE))
    add_rect(slide, Inches(0), Inches(5.9), prs.slide_width, Inches(1.6), TEAL)
    tb(slide, Inches(0.8), Inches(6.2), Inches(11.5), Inches(0.4),
       "Built for students who are already good at English — and ready to sharpen reading + grammar",
       size=16, color=WHITE)
    tb(slide, Inches(0.8), Inches(6.7), Inches(11.5), Inches(0.35),
       "CCSS-aligned focus: Main Idea  •  Inference  •  Text Evidence  •  Conventions",
       size=14, color=RGBColor(0xCC, 0xFB, 0xF1))


def s_agenda():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide)
    chip(slide, "TODAY'S FLOW")
    time_chip(slide, "60 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.5),
       "One Passage. Four Games. Real Skills.", size=28, bold=True, color=NAVY, font="Georgia")
    tb(slide, Inches(0.45), Inches(1.4), Inches(12), Inches(0.35),
       "Every game uses today’s reading or grammar — so play time = practice time.",
       size=15, color=SOFT)

    items = [
        ("04m", "Hook", "Meet the storm passage", SKY, False),
        ("07m", "Game 1", "Evidence Hunt (from passage)", GOLD, True),
        ("08m", "Skill Mini", "Main idea + inference tools", TEAL, False),
        ("07m", "Game 2", "Inference Relay (same text)", GOLD, True),
        ("10m", "Deep Read", "Annotate + answer questions", NAVY, False),
        ("07m", "Game 3", "Grammar Fix from the Story", GOLD, True),
        ("10m", "Grammar Lab", "Agreement • commas • word pairs", CORAL, False),
        ("07m", "Game 4", "Passage Remix Challenge", GOLD, True),
    ]
    for i, (mins, title, desc, color, is_game) in enumerate(items):
        col, row = i % 4, i // 4
        left = Inches(0.4 + col * 3.2)
        top = Inches(1.95 + row * 2.4)
        add_round(slide, left, top, Inches(3.05), Inches(2.15), WHITE)
        add_rect(slide, left, top, Inches(3.05), Inches(0.5), color)
        head = "GAME" if is_game else mins
        tb(slide, left, top + Inches(0.08), Inches(3.05), Inches(0.35),
           head, size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.7), Inches(2.75), Inches(0.45),
           title, size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(1.25), Inches(2.75), Inches(0.7),
           desc, size=13, color=SOFT, align=PP_ALIGN.CENTER)
    footer(slide, 2, "Overview")


def s_goals():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, TEAL)
    chip(slide, "LEARNING TARGETS", color=TEAL)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.5),
       "What You’ll Walk Away With", size=28, bold=True, color=NAVY, font="Georgia")
    goals = [
        (SKY, "Find the Big Idea", "State the main idea of a Grade 7 passage in one clear sentence"),
        (TEAL, "Prove It", "Support answers with text evidence (not guesses)"),
        (PURPLE, "Infer Smart", "Read between the lines using clues from the same story"),
        (CORAL, "Polish Grammar", "Fix agreement, commas, and confusing word pairs in context"),
    ]
    for i, (c, t, d) in enumerate(goals):
        top = Inches(1.55 + i * 1.25)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(1.1), WHITE)
        add_round(slide, Inches(0.65), top + Inches(0.25), Inches(0.6), Inches(0.6), c)
        tb(slide, Inches(0.65), top + Inches(0.35), Inches(0.6), Inches(0.4),
           str(i + 1), size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.2), Inches(11), Inches(0.35),
           t, size=20, bold=True, color=NAVY)
        tb(slide, Inches(1.5), top + Inches(0.6), Inches(11), Inches(0.35),
           d, size=15, color=SOFT)
    footer(slide, 3, "Goals")


def s_hook_passage():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, SKY)
    chip(slide, "TODAY'S PASSAGE")
    time_chip(slide, "4 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       f"Read-Aloud Hook: “{PASSAGE_TITLE}”", size=26, bold=True, color=NAVY, font="Georgia")
    add_round(slide, Inches(0.45), Inches(1.45), Inches(12.4), Inches(4.55), WHITE)
    tb(slide, Inches(0.75), Inches(1.65), Inches(11.8), Inches(4.15),
       PASSAGE, size=15, color=DARK)
    add_round(slide, Inches(0.45), Inches(6.15), Inches(12.4), Inches(0.8), LIGHT_SKY)
    tb(slide, Inches(0.7), Inches(6.3), Inches(11.9), Inches(0.5),
       "Quick feel check: What’s the mood? Who seems ready to learn? What’s one question you’d ask?",
       size=14, bold=True, color=NAVY)
    footer(slide, 4, "0–4 min")


def s_game1_intro():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), GOLD)
    tb(slide, Inches(0.8), Inches(2.0), Inches(11.7), Inches(0.4),
       "GAME 1  •  LINKED TO THE PASSAGE", size=16, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(2.6), Inches(11.7), Inches(1.0),
       "Evidence Hunt", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(3.8), Inches(11.7), Inches(0.5),
       "Find proof in “The Signal in the Storm” — not opinions.",
       size=18, color=RGBColor(0xBF, 0xDB, 0xFE), align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(4.6), Inches(11.7), Inches(0.4),
       "7 minutes  •  Syllabus skill: Text Evidence (RI.7.1)",
       size=16, color=GOLD, align=PP_ALIGN.CENTER)


def s_game1_play():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    game_chip(slide)
    time_chip(slide, "7 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Evidence Hunt — Rules", size=26, bold=True, color=NAVY, font="Georgia")

    steps = [
        ("Listen", "I’ll read a claim about the passage."),
        ("Hunt", "Find a sentence or phrase that proves it."),
        ("Quote", "Say the evidence out loud (word-for-word)."),
        ("Score", "+1 find  •  +1 exact quote  •  Max 2 each round"),
    ]
    for i, (t, d) in enumerate(steps):
        left = Inches(0.4 + i * 3.2)
        add_round(slide, left, Inches(1.5), Inches(3.05), Inches(3.5), WHITE)
        add_oval(slide, left + Inches(1.0), Inches(1.75), Inches(1.0), Inches(1.0), GOLD)
        tb(slide, left + Inches(1.0), Inches(1.95), Inches(1.0), Inches(0.6),
           str(i + 1), size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.0), Inches(2.75), Inches(0.45),
           t, size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.55), Inches(2.75), Inches(1.1),
           d, size=14, color=SOFT, align=PP_ALIGN.CENTER)
    why_game(slide, "Builds RI.7.1 — cite textual evidence to support analysis.")
    footer(slide, 6, "4–11 min")


def s_game1_rounds():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    game_chip(slide)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Evidence Hunt — Claims (from our passage)", size=24, bold=True, color=NAVY, font="Georgia")
    claims = [
        "Claim A: The class did not waste the power outage.",
        "Claim B: Ms. Reyes wanted students to keep learning.",
        "Claim C: Maya was already using a strong reading habit.",
        "Claim D: A good question changed the energy in the room.",
        "Claim E: The author believes small choices can create learning.",
    ]
    for i, c in enumerate(claims):
        top = Inches(1.45 + i * 0.9)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.8), WHITE if i % 2 == 0 else LIGHT_GOLD)
        tb(slide, Inches(0.7), top + Inches(0.22), Inches(12.0), Inches(0.45),
           c, size=16, color=DARK)
    why_game(slide, "Every claim must be proven with words from “The Signal in the Storm.”")
    footer(slide, 7, "4–11 min")


def s_game1_key():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    chip(slide, "ANSWER KEY", color=GOLD)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Sample Evidence (accept close matches)", size=24, bold=True, color=NAVY, font="Georgia")
    keys = [
        "A → “They had practiced… turned a storm delay into learning.”",
        "B → “If we finish this reading together, nobody loses progress.”",
        "C → “Maya already underlining the first paragraph.”",
        "D → “What does the author want us to notice?” / “The room shifted.”",
        "E → “Small choices… had turned a storm delay into learning.”",
    ]
    for i, k in enumerate(keys):
        top = Inches(1.5 + i * 0.95)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.85), LIGHT_TEAL)
        tb(slide, Inches(0.7), top + Inches(0.25), Inches(12.0), Inches(0.45),
           k, size=15, color=DARK)
    footer(slide, 8, "4–11 min")


def s_skills():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, TEAL)
    chip(slide, "SKILL MINI", color=TEAL)
    time_chip(slide, "8 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Reading Toolkit (Grade 7)", size=28, bold=True, color=NAVY, font="Georgia")
    cards = [
        (SKY, "Main Idea", "What is the passage mostly about?\nOne sentence. No side stories."),
        (TEAL, "Key Details", "Which moments prove the main idea?\nLook for actions + dialogue."),
        (PURPLE, "Inference", "What is suggested but not said?\nClue + your logic = inference."),
        (CORAL, "Author’s Purpose", "Why write this?\nInform / persuade / entertain?"),
    ]
    for i, (c, t, d) in enumerate(cards):
        col, row = i % 2, i // 2
        left = Inches(0.45 + col * 6.4)
        top = Inches(1.5 + row * 2.5)
        add_round(slide, left, top, Inches(6.15), Inches(2.25), WHITE)
        add_rect(slide, left, top, Inches(6.15), Inches(0.55), c)
        tb(slide, left + Inches(0.25), top + Inches(0.1), Inches(5.7), Inches(0.4),
           t, size=18, bold=True, color=WHITE)
        tb(slide, left + Inches(0.25), top + Inches(0.8), Inches(5.7), Inches(1.2),
           d, size=15, color=DARK)
    footer(slide, 9, "11–19 min")


def s_annotate():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, TEAL)
    chip(slide, "ANNOTATION", color=TEAL)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Mark the Passage Like a Pro", size=28, bold=True, color=NAVY, font="Georgia")
    codes = [
        ("★", SKY, "Star", "Main idea moment"),
        ("—", TEAL, "Underline", "Evidence you’ll quote"),
        ("?", CORAL, "Question", "Confusing line"),
        ("!", GOLD, "Aha", "Tone / shift / surprise"),
    ]
    for i, (sym, c, name, use) in enumerate(codes):
        left = Inches(0.45 + i * 3.2)
        add_round(slide, left, Inches(1.55), Inches(3.05), Inches(4.7), WHITE)
        add_oval(slide, left + Inches(0.9), Inches(1.9), Inches(1.2), Inches(1.2), c)
        tb(slide, left + Inches(0.9), Inches(2.15), Inches(1.2), Inches(0.7),
           sym, size=26, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.4), Inches(2.75), Inches(0.45),
           name, size=20, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(4.1), Inches(2.75), Inches(1.2),
           use, size=15, color=SOFT, align=PP_ALIGN.CENTER)
    footer(slide, 10, "11–19 min")


def s_game2_intro():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), GOLD)
    tb(slide, Inches(0.8), Inches(2.0), Inches(11.7), Inches(0.4),
       "GAME 2  •  SAME PASSAGE, DEEPER THINKING", size=16, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(2.6), Inches(11.7), Inches(1.0),
       "Inference Relay", size=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(3.8), Inches(11.7), Inches(0.5),
       "What can you figure out that the author never says directly?",
       size=18, color=RGBColor(0xBF, 0xDB, 0xFE), align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(4.6), Inches(11.7), Inches(0.4),
       "7 minutes  •  Syllabus skill: Make inferences (RL/RI.7.1)",
       size=16, color=GOLD, align=PP_ALIGN.CENTER)


def s_game2_play():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    game_chip(slide)
    time_chip(slide, "7 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Inference Relay — From Our Storm Story", size=24, bold=True, color=NAVY, font="Georgia")
    tb(slide, Inches(0.45), Inches(1.4), Inches(12), Inches(0.35),
       "Say: (1) your inference  (2) the clue words  (3) why the clue fits",
       size=14, color=SOFT)

    packs = [
        ("Relay 1", "Jordan frowned at the dark projector…"),
        ("Relay 2", "whispers rose—until someone asked a real question"),
        ("Relay 3", "When the lights flickered back on, the class had done more than wait."),
        ("Relay 4", "Ms. Reyes kept her voice calm… “nobody loses progress.”"),
    ]
    for i, (t, d) in enumerate(packs):
        col, row = i % 2, i // 2
        left = Inches(0.45 + col * 6.4)
        top = Inches(1.9 + row * 2.15)
        add_round(slide, left, top, Inches(6.15), Inches(1.95), WHITE)
        tb(slide, left + Inches(0.25), top + Inches(0.25), Inches(5.7), Inches(0.35),
           t, size=16, bold=True, color=GOLD)
        tb(slide, left + Inches(0.25), top + Inches(0.75), Inches(5.7), Inches(0.9),
           f'Clue: “{d}”', size=14, color=DARK)
    why_game(slide, "Inferences must come from today’s passage — keeps reading + play connected.")
    footer(slide, 12, "19–26 min")


def s_game2_key():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    chip(slide, "SAMPLE ANSWERS", color=GOLD)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Strong Inference Examples", size=26, bold=True, color=NAVY, font="Georgia")
    ans = [
        "1: Jordan feels frustrated/disappointed — clue: frowned + dark projector.",
        "2: Students were off-task until a purposeful question refocused them.",
        "3: The delay became productive; they used the time to practice reading habits.",
        "4: Ms. Reyes is steady and cares about learning continuity, not just control.",
    ]
    for i, a in enumerate(ans):
        top = Inches(1.55 + i * 1.2)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(1.05), LIGHT_PURPLE if i % 2 else LIGHT_TEAL)
        tb(slide, Inches(0.7), top + Inches(0.3), Inches(12.0), Inches(0.5), a, size=16, color=DARK)
    footer(slide, 13, "19–26 min")


def s_deep_read_task():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, SKY)
    chip(slide, "DEEP READ")
    time_chip(slide, "10 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Now You Own the Passage", size=28, bold=True, color=NAVY, font="Georgia")
    steps = [
        "Re-read “The Signal in the Storm” silently (2 min).",
        "Annotate with ★ — ? !",
        "Write a one-sentence main idea.",
        "Answer the 5 questions on the next slide.",
        "Be ready to share one starred line + why.",
    ]
    for i, s in enumerate(steps):
        top = Inches(1.5 + i * 0.95)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.85), WHITE)
        add_round(slide, Inches(0.65), top + Inches(0.18), Inches(1.35), Inches(0.5), SKY)
        tb(slide, Inches(0.65), top + Inches(0.25), Inches(1.35), Inches(0.4),
           f"Step {i + 1}", size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.25), top + Inches(0.25), Inches(10.3), Inches(0.45),
           s, size=16, color=DARK)
    footer(slide, 14, "26–36 min")


def s_questions():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, SKY)
    chip(slide, "QUESTIONS")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Comprehension — All About Today’s Passage", size=24, bold=True, color=NAVY, font="Georgia")
    qs = [
        ("Main Idea", "What is the central message of “The Signal in the Storm”?"),
        ("Detail", "What does Ms. Reyes want the class to finish during the outage?"),
        ("Inference", "Why does the room “shift” after the student’s question?"),
        ("Purpose", "Is the author mostly informing or persuading? Explain."),
        ("Evidence", "Quote one sentence that best supports your main idea."),
    ]
    for i, (tag, q) in enumerate(qs):
        top = Inches(1.45 + i * 0.95)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.85), WHITE)
        add_round(slide, Inches(0.65), top + Inches(0.2), Inches(1.9), Inches(0.45),
                  CORAL if i in (2, 3) else TEAL)
        tb(slide, Inches(0.65), top + Inches(0.25), Inches(1.9), Inches(0.35),
           tag, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.8), top + Inches(0.25), Inches(9.8), Inches(0.45),
           f"{i + 1}. {q}", size=15, color=DARK)
    footer(slide, 15, "26–36 min")


def s_q_key():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, SKY)
    chip(slide, "MODEL ANSWERS")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Discuss Variations — Reward Evidence", size=26, bold=True, color=NAVY, font="Georgia")
    answers = [
        "Main idea: During a power outage, focused questions and small choices helped the class keep learning.",
        "Detail: The reading packet — so nobody loses progress.",
        "Inference: The question redirected attention from chatter to purposeful reading.",
        "Purpose: Mostly persuade/inform — show that strong reading habits turn delays into learning.",
        "Evidence example: “Small choices… turned a storm delay into learning.”",
    ]
    bullets(slide, Inches(0.45), Inches(1.5), Inches(12.4), Inches(5.0), answers, size=16, spacing=14)
    footer(slide, 16, "26–36 min")


def s_game3_intro():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), GOLD)
    tb(slide, Inches(0.8), Inches(2.0), Inches(11.7), Inches(0.4),
       "GAME 3  •  GRAMMAR FROM THE STORY WORLD", size=16, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(2.6), Inches(11.7), Inches(1.0),
       "Grammar Fix:\nStorm Edition", size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(4.4), Inches(11.7), Inches(0.5),
       "Broken sentences about our characters — you repair them.",
       size=18, color=RGBColor(0xBF, 0xDB, 0xFE), align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(5.1), Inches(11.7), Inches(0.4),
       "7 minutes  •  Syllabus skill: Conventions (L.7.1 / L.7.2)",
       size=16, color=GOLD, align=PP_ALIGN.CENTER)


def s_game3_play():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    game_chip(slide)
    time_chip(slide, "7 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Fix These — All Tied to the Passage", size=26, bold=True, color=NAVY, font="Georgia")
    sents = [
        "The class of students are waiting for the power to return.",
        "Maya already underline the first paragraph.",
        "Its important that nobody lose progress.",
        "Jordan frowned at the projector but she kept listening.",
        "There pencils moved faster then before.",
    ]
    for i, s in enumerate(sents):
        top = Inches(1.45 + i * 0.9)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.8), WHITE)
        tb(slide, Inches(0.7), top + Inches(0.22), Inches(12.0), Inches(0.45),
           f"{i + 1}.  {s}", size=15, color=DARK)
    why_game(slide, "Same characters/setting as the reading — grammar practice feels connected, not random.")
    footer(slide, 18, "36–43 min")


def s_game3_key():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    chip(slide, "ANSWER KEY", color=GOLD)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Corrected + Rule", size=26, bold=True, color=NAVY, font="Georgia")
    fixes = [
        "The class of students is waiting…  (class = singular collective)",
        "Maya already underlined / is underlining…  (verb tense/form)",
        "It’s important that nobody loses progress.  (it’s / subject-verb)",
        "Jordan frowned at the projector, but she kept listening.  (FANBOYS comma)",
        "Their pencils moved faster than before.  (their / than)",
    ]
    for i, f in enumerate(fixes):
        top = Inches(1.45 + i * 0.95)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(0.85), LIGHT_TEAL)
        tb(slide, Inches(0.7), top + Inches(0.25), Inches(12.0), Inches(0.45),
           f"{i + 1}.  {f}", size=15, color=DARK)
    footer(slide, 19, "36–43 min")


def s_grammar_lab():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, CORAL)
    chip(slide, "GRAMMAR LAB", color=CORAL)
    time_chip(slide, "10 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "High-Impact Fixes for Grade 7", size=28, bold=True, color=NAVY, font="Georgia")
    topics = [
        ("Agreement", "Everyone / each / class → singular verbs"),
        ("Commas", "Lists • FANBOYS • intro phrases"),
        ("Word Pairs", "its/it’s • their/there/they’re • than/then"),
        ("Sentence Sense", "No fragments, no run-ons"),
    ]
    for i, (t, d) in enumerate(topics):
        left = Inches(0.45 + i * 3.2)
        add_round(slide, left, Inches(1.55), Inches(3.05), Inches(4.7), WHITE)
        add_rect(slide, left, Inches(1.55), Inches(3.05), Inches(1.1), CORAL)
        tb(slide, left + Inches(0.15), Inches(1.8), Inches(2.75), Inches(0.6),
           t, size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(3.1), Inches(2.65), Inches(2.5),
           d, size=15, color=DARK, align=PP_ALIGN.CENTER)
    footer(slide, 20, "43–53 min")


def s_grammar_rules():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, CORAL)
    chip(slide, "QUICK RULES", color=CORAL)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Keep These in Your Pocket", size=28, bold=True, color=NAVY, font="Georgia")
    rows = [
        ("Agreement", "The packet of pages is on the desk. (packet = singular)"),
        ("FANBOYS", "She wanted to quit, but she kept reading."),
        ("Intro comma", "After the lights returned, pencils kept moving."),
        ("it’s / its", "It’s (= it is) raining. The class found its focus."),
        ("than / then", "Better than chatting. Ask first, then annotate."),
        ("their / there / they’re", "They’re putting their notes over there."),
    ]
    for i, (t, d) in enumerate(rows):
        col, row = i % 2, i // 2
        left = Inches(0.45 + col * 6.4)
        top = Inches(1.5 + row * 1.65)
        add_round(slide, left, top, Inches(6.15), Inches(1.45), WHITE)
        tb(slide, left + Inches(0.25), top + Inches(0.25), Inches(5.7), Inches(0.35),
           t, size=16, bold=True, color=CORAL)
        tb(slide, left + Inches(0.25), top + Inches(0.7), Inches(5.7), Inches(0.5),
           d, size=14, color=DARK)
    footer(slide, 21, "43–53 min")


def s_grammar_practice():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, CORAL)
    chip(slide, "LIVE PRACTICE", color=CORAL)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Apply the Rules (Still in Story Context)", size=26, bold=True, color=NAVY, font="Georgia")
    items = [
        "Everyone in Room 214 have a strategy for hard paragraphs.",
        "The stack of packets are tipping near the window.",
        "Its clear that asking better questions help readers.",
        "Reading with evidence has a bigger affect then guessing.",
    ]
    for i, s in enumerate(items):
        top = Inches(1.5 + i * 1.2)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(1.05), WHITE)
        tb(slide, Inches(0.7), top + Inches(0.3), Inches(12.0), Inches(0.45),
           f"{i + 1}.  {s}", size=16, color=DARK)
    footer(slide, 22, "43–53 min")


def s_grammar_answers():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, CORAL)
    chip(slide, "ANSWER KEY", color=CORAL)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Corrected Versions", size=28, bold=True, color=NAVY, font="Georgia")
    fixes = [
        "Everyone in Room 214 has a strategy for hard paragraphs.",
        "The stack of packets is tipping near the window.",
        "It’s clear that asking better questions helps readers.",
        "Reading with evidence has a bigger effect than guessing.",
    ]
    for i, f in enumerate(fixes):
        top = Inches(1.5 + i * 1.2)
        add_round(slide, Inches(0.45), top, Inches(12.4), Inches(1.05), LIGHT_TEAL)
        tb(slide, Inches(0.7), top + Inches(0.3), Inches(12.0), Inches(0.45),
           f"{i + 1}.  {f}", size=16, color=DARK)
    footer(slide, 23, "43–53 min")


def s_game4_intro():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), GOLD)
    tb(slide, Inches(0.8), Inches(2.0), Inches(11.7), Inches(0.4),
       "GAME 4  •  FINAL BOSS — READING + GRAMMAR TOGETHER", size=15, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(2.6), Inches(11.7), Inches(1.0),
       "Passage Remix\nChallenge", size=40, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(4.4), Inches(11.7), Inches(0.5),
       "Edit a messy remix of our story, then add one strong inference.",
       size=18, color=RGBColor(0xBF, 0xDB, 0xFE), align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(5.1), Inches(11.7), Inches(0.4),
       "7 minutes  •  Integrates RI.7.1 + L.7.1/L.7.2",
       size=16, color=GOLD, align=PP_ALIGN.CENTER)


def s_game4_play():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    game_chip(slide)
    time_chip(slide, "7 minutes")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Passage Remix — Rescue This Paragraph", size=24, bold=True, color=NAVY, font="Georgia")

    add_round(slide, Inches(0.45), Inches(1.45), Inches(12.4), Inches(2.35), WHITE)
    tb(slide, Inches(0.7), Inches(1.6), Inches(12.0), Inches(0.35),
       "Broken remix (based on today’s passage):", size=14, bold=True, color=CORAL)
    tb(slide, Inches(0.7), Inches(2.05), Inches(12.0), Inches(1.5),
       "Rain hammer the windows while the class wait. Ms. Reyes stay calm and say nobody lose progress. "
       "Maya already underline the paragraph its a smart habit. Asking better questions help the room shift "
       "and turn a delay into learning.",
       size=15, color=DARK)

    add_round(slide, Inches(0.45), Inches(4.0), Inches(6.05), Inches(2.35), LIGHT_SKY)
    tb(slide, Inches(0.7), Inches(4.2), Inches(5.6), Inches(0.35),
       "Mission A (4 min)", size=16, bold=True, color=NAVY)
    bullets(slide, Inches(0.7), Inches(4.7), Inches(5.6), Inches(1.5), [
        "Rewrite the paragraph correctly",
        "Fix verbs, agreement, it’s/its",
        "Keep the original meaning",
    ], size=14)

    add_round(slide, Inches(6.75), Inches(4.0), Inches(6.1), Inches(2.35), LIGHT_GOLD)
    tb(slide, Inches(7.0), Inches(4.2), Inches(5.6), Inches(0.35),
       "Mission B (3 min)", size=16, bold=True, color=NAVY)
    bullets(slide, Inches(7.0), Inches(4.7), Inches(5.6), Inches(1.5), [
        "Add 1 inference about the class",
        "Include a clue from the passage",
        "One clear sentence is enough",
    ], size=14)
    footer(slide, 25, "53–60 min")


def s_game4_model():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    chip(slide, "MODEL RESPONSE", color=GOLD)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "What a Strong Finish Looks Like", size=26, bold=True, color=NAVY, font="Georgia")
    add_round(slide, Inches(0.45), Inches(1.5), Inches(12.4), Inches(2.5), WHITE)
    tb(slide, Inches(0.7), Inches(1.7), Inches(12.0), Inches(0.35),
       "Edited paragraph", size=15, bold=True, color=TEAL)
    tb(slide, Inches(0.7), Inches(2.15), Inches(12.0), Inches(1.5),
       "Rain hammers the windows while the class waits. Ms. Reyes stays calm and says nobody loses progress. "
       "Maya already underlines the paragraph; it’s a smart habit. Asking better questions helps the room shift "
       "and turns a delay into learning.",
       size=15, color=DARK)
    add_round(slide, Inches(0.45), Inches(4.25), Inches(12.4), Inches(2.2), LIGHT_TEAL)
    tb(slide, Inches(0.7), Inches(4.45), Inches(12.0), Inches(0.35),
       "Sample inference", size=15, bold=True, color=NAVY)
    tb(slide, Inches(0.7), Inches(4.95), Inches(12.0), Inches(1.2),
       "The class values learning even during interruptions, because they choose questions and annotation "
       "instead of only waiting for the lights.",
       size=15, color=DARK)
    footer(slide, 26, "53–60 min")


def s_wrap():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, TEAL)
    chip(slide, "WRAP-UP", color=TEAL)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Takeaways + Weekly Challenge", size=28, bold=True, color=NAVY, font="Georgia")
    left = [
        ("Evidence first", "Quote the text before you explain"),
        ("Infer with clues", "No clue = no claim"),
        ("Grammar in context", "Fix sentences about real reading"),
    ]
    for i, (t, d) in enumerate(left):
        top = Inches(1.5 + i * 1.55)
        add_round(slide, Inches(0.45), top, Inches(6.05), Inches(1.35), WHITE)
        tb(slide, Inches(0.7), top + Inches(0.25), Inches(5.6), Inches(0.35),
           t, size=18, bold=True, color=TEAL)
        tb(slide, Inches(0.7), top + Inches(0.7), Inches(5.6), Inches(0.4),
           d, size=14, color=DARK)

    add_round(slide, Inches(6.75), Inches(1.5), Inches(6.1), Inches(4.85), NAVY)
    tb(slide, Inches(7.05), Inches(1.8), Inches(5.6), Inches(0.4),
       "This Week’s Mission", size=18, bold=True, color=GOLD)
    bullets(slide, Inches(7.05), Inches(2.45), Inches(5.6), Inches(3.5), [
        "Read 15 minutes daily",
        "Annotate one short article (★ — ?)",
        "Write a 1-sentence main idea",
        "Replay Evidence Hunt on 3 claims",
        "Bring one tricky sentence next time",
    ], size=15, color=WHITE, spacing=12)
    footer(slide, 27, "Close")


def s_close():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.22), GOLD)
    tb(slide, Inches(0.8), Inches(2.1), Inches(11.7), Inches(0.45),
       "You’re already good at English.", size=20, color=GOLD, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(2.7), Inches(11.7), Inches(1.3),
       "Today you practiced becoming\nsharper — with games that teach.",
       size=32, bold=True, color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(4.4), Inches(11.7), Inches(0.45),
       "Favorite game? Best evidence quote? One skill you’ll use tomorrow?",
       size=16, color=RGBColor(0xBF, 0xDB, 0xFE), align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(5.4), Inches(11.7), Inches(0.4),
       "Great work today!", size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def s_facilitator():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide)
    chip(slide, "FACILITATOR")
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "How to Keep Energy High", size=28, bold=True, color=NAVY, font="Georgia")
    notes = [
        "Keep one passage all hour — games reuse it so learning feels cohesive.",
        "Score lightly and celebrate process (“smart quote”) more than speed.",
        "If she’s stuck, reread one sentence aloud together, then retry.",
        "Games are short on purpose: 5–7 minutes max, then back to skill.",
        "Push Grade 7 rigor: require evidence words, not “I feel like…”",
        "Materials: timer, notebook, pencil; optional printed passage page.",
    ]
    numbered(slide, Inches(0.45), Inches(1.5), Inches(12.4), Inches(5.0), notes, size=16, spacing=12)
    footer(slide, 29, "Teacher")


def s_game_map():
    """Extra clarity slide: games ↔ syllabus."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    accent(slide, GOLD)
    chip(slide, "GAME ↔ SYLLABUS", color=GOLD)
    tb(slide, Inches(0.45), Inches(0.85), Inches(12), Inches(0.45),
       "Why These Games Aren’t “Just Fun”", size=26, bold=True, color=NAVY, font="Georgia")
    rows = [
        ("Evidence Hunt", "RI.7.1 Text evidence", "Claims about the storm passage"),
        ("Inference Relay", "RL/RI.7.1 Inference", "Clues pulled from same story"),
        ("Grammar Fix: Storm Edition", "L.7.1 / L.7.2 Conventions", "Sentences about Jordan, Maya, Ms. Reyes"),
        ("Passage Remix Challenge", "Reading + Grammar together", "Edit remix + add inference"),
    ]
    add_rect(slide, Inches(0.45), Inches(1.5), Inches(12.4), Inches(0.55), NAVY)
    tb(slide, Inches(0.65), Inches(1.6), Inches(3.5), Inches(0.35), "Game", size=14, bold=True, color=WHITE)
    tb(slide, Inches(4.3), Inches(1.6), Inches(4.0), Inches(0.35), "Syllabus Target", size=14, bold=True, color=WHITE)
    tb(slide, Inches(8.5), Inches(1.6), Inches(4.1), Inches(0.35), "Tied To", size=14, bold=True, color=WHITE)
    for i, (g, s, t) in enumerate(rows):
        top = Inches(2.15 + i * 1.05)
        bg = WHITE if i % 2 == 0 else LIGHT_GOLD
        add_rect(slide, Inches(0.45), top, Inches(12.4), Inches(1.05), bg)
        tb(slide, Inches(0.65), top + Inches(0.3), Inches(3.5), Inches(0.45), g, size=14, bold=True, color=NAVY)
        tb(slide, Inches(4.3), top + Inches(0.3), Inches(4.0), Inches(0.45), s, size=14, color=DARK)
        tb(slide, Inches(8.5), top + Inches(0.3), Inches(4.1), Inches(0.45), t, size=14, color=SOFT)
    footer(slide, 30, "Map")


# Build deck order
s_title()
s_agenda()
s_goals()
s_game_map()
s_hook_passage()
s_game1_intro()
s_game1_play()
s_game1_rounds()
s_game1_key()
s_skills()
s_annotate()
s_game2_intro()
s_game2_play()
s_game2_key()
s_deep_read_task()
s_questions()
s_q_key()
s_game3_intro()
s_game3_play()
s_game3_key()
s_grammar_lab()
s_grammar_rules()
s_grammar_practice()
s_grammar_answers()
s_game4_intro()
s_game4_play()
s_game4_model()
s_wrap()
s_close()
s_facilitator()

out = r"C:\Users\bhushaja\Downloads\shaip\Grade7_Reading_Grammar_60min_Demo.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
