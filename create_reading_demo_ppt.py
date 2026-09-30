"""Generate a 60-minute Reading & Grammar demo PPT for Grade 8."""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import nsmap
from copy import deepcopy
from lxml import etree

# Colors
NAVY = RGBColor(0x1B, 0x3A, 0x5F)
TEAL = RGBColor(0x2A, 0x9D, 0x8F)
CORAL = RGBColor(0xE7, 0x6F, 0x51)
CREAM = RGBColor(0xF8, 0xF5, 0xF0)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
DARK = RGBColor(0x2D, 0x2D, 0x2D)
SOFT_GRAY = RGBColor(0x6B, 0x72, 0x80)
LIGHT_TEAL = RGBColor(0xE8, 0xF5, 0xF3)
LIGHT_CORAL = RGBColor(0xFD, 0xED, 0xE8)
LIGHT_NAVY = RGBColor(0xE8, 0xEE, 0xF5)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)


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
    # Send to back
    spTree = slide.shapes._spTree
    sp = shape._element
    spTree.remove(sp)
    spTree.insert(2, sp)
    return shape


def add_rect(slide, left, top, width, height, fill, line=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
    return shape


def add_round_rect(slide, left, top, width, height, fill):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.fill.background()
    return shape


def add_textbox(slide, left, top, width, height, text, size=18, bold=False, color=DARK,
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


def add_bullets(slide, left, top, width, height, items, size=18, color=DARK, bold_first=False,
                spacing=8):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.space_after = Pt(spacing)
        p.level = 0
        # support (text, is_bold) or plain string
        if isinstance(item, tuple):
            text, is_bold = item
        else:
            text, is_bold = item, False
        run = p.add_run()
        run.text = "•  " + text
        set_run(run, size=size, bold=is_bold or (bold_first and i == 0), color=color)
    return box


def add_numbered(slide, left, top, width, height, items, size=18, color=DARK, spacing=8):
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


def add_accent_bar(slide, color=TEAL):
    add_rect(slide, Inches(0), Inches(0), Inches(0.18), prs.slide_height, color)


def add_footer(slide, page, total=22, time_label=""):
    add_rect(slide, Inches(0), Inches(7.15), prs.slide_width, Inches(0.35), NAVY)
    label = f"Reading & Grammar Demo  |  Grade 8  |  60 Minutes"
    if time_label:
        label += f"  |  {time_label}"
    add_textbox(slide, Inches(0.4), Inches(7.18), Inches(11), Inches(0.3),
                label, size=11, color=WHITE)
    add_textbox(slide, Inches(12.2), Inches(7.18), Inches(0.9), Inches(0.3),
                f"{page}/{total}", size=11, color=WHITE, align=PP_ALIGN.RIGHT)


def add_section_chip(slide, text, left=Inches(0.5), top=Inches(0.35)):
    chip = add_round_rect(slide, left, top, Inches(2.4), Inches(0.38), TEAL)
    add_textbox(slide, left, top + Inches(0.02), Inches(2.4), Inches(0.35),
                text, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def add_time_chip(slide, text, left=Inches(10.6), top=Inches(0.35)):
    chip = add_round_rect(slide, left, top, Inches(2.3), Inches(0.38), CORAL)
    add_textbox(slide, left, top + Inches(0.02), Inches(2.3), Inches(0.35),
                text, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def title_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(5.8), prs.slide_width, Inches(1.7), TEAL)
    add_textbox(slide, Inches(0.8), Inches(1.6), Inches(11.5), Inches(0.5),
                "60-MINUTE DEMO SESSION", size=16, bold=True, color=TEAL)
    add_textbox(slide, Inches(0.8), Inches(2.2), Inches(11.5), Inches(1.2),
                "Reading Content &\nGrammar Mastery", size=44, bold=True, color=WHITE,
                font="Georgia")
    add_textbox(slide, Inches(0.8), Inches(4.6), Inches(11.5), Inches(0.5),
                "Grade 8  •  U.S. Curriculum Aligned  •  Interactive Demo",
                size=18, color=RGBColor(0xC5, 0xD4, 0xE8))
    add_textbox(slide, Inches(0.8), Inches(6.2), Inches(11.5), Inches(0.4),
                "Goal: Build stronger reading comprehension and confident grammar skills",
                size=16, color=WHITE)
    add_textbox(slide, Inches(0.8), Inches(6.7), Inches(11.5), Inches(0.35),
                "Designed for students who are already good at English and ready to level up",
                size=14, color=RGBColor(0xE0, 0xF2, 0xEF))


def agenda_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "SESSION ROADMAP")
    add_time_chip(slide, "Total: 60 min")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.6),
                "Today's Agenda", size=32, bold=True, color=NAVY, font="Georgia")

    blocks = [
        ("05 min", "Warm-Up", "Quick reading spark + grammar warm-up", TEAL),
        ("12 min", "Reading Strategies", "How strong readers think while they read", NAVY),
        ("15 min", "Guided Passage Practice", "Read, annotate, answer, discuss", TEAL),
        ("12 min", "Grammar Power-Up", "Fix the gaps that hold writing back", CORAL),
        ("10 min", "Apply & Create", "Use reading + grammar together", NAVY),
        ("06 min", "Wrap-Up", "Takeaways, tips, and next steps", TEAL),
    ]
    for i, (mins, title, desc, color) in enumerate(blocks):
        col = i % 3
        row = i // 3
        left = Inches(0.5 + col * 4.2)
        top = Inches(1.8 + row * 2.4)
        card = add_round_rect(slide, left, top, Inches(3.9), Inches(2.1), WHITE)
        add_rect(slide, left, top, Inches(0.15), Inches(2.1), color)
        add_textbox(slide, left + Inches(0.35), top + Inches(0.25), Inches(3.3), Inches(0.4),
                    mins, size=14, bold=True, color=color)
        add_textbox(slide, left + Inches(0.35), top + Inches(0.7), Inches(3.3), Inches(0.5),
                    title, size=20, bold=True, color=NAVY)
        add_textbox(slide, left + Inches(0.35), top + Inches(1.25), Inches(3.3), Inches(0.6),
                    desc, size=14, color=SOFT_GRAY)
    add_footer(slide, 2, time_label="Overview")


def goals_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "LEARNING GOALS")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.6),
                "By the End of This Hour, You Will…", size=28, bold=True, color=NAVY, font="Georgia")

    goals = [
        ("Find the Big Idea Fast", "Identify main idea, key details, and the author’s purpose"),
        ("Read Between the Lines", "Make inferences using text evidence — not guesses"),
        ("Own Your Grammar", "Fix subject-verb agreement, commas, and commonly confused words"),
        ("Write Clearer Sentences", "Combine and revise sentences so ideas sound polished"),
    ]
    for i, (title, desc) in enumerate(goals):
        top = Inches(1.7 + i * 1.15)
        add_round_rect(slide, Inches(0.5), top, Inches(12.3), Inches(1.0), WHITE)
        num = add_round_rect(slide, Inches(0.7), top + Inches(0.22), Inches(0.55), Inches(0.55), TEAL)
        add_textbox(slide, Inches(0.7), top + Inches(0.28), Inches(0.55), Inches(0.45),
                    str(i + 1), size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(1.5), top + Inches(0.18), Inches(10.8), Inches(0.4),
                    title, size=20, bold=True, color=NAVY)
        add_textbox(slide, Inches(1.5), top + Inches(0.55), Inches(10.8), Inches(0.35),
                    desc, size=15, color=SOFT_GRAY)
    add_footer(slide, 3, time_label="Goals")


def warmup_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide, CORAL)
    add_section_chip(slide, "WARM-UP")
    add_time_chip(slide, "5 minutes")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Brain Switch-On", size=28, bold=True, color=NAVY, font="Georgia")

    add_round_rect(slide, Inches(0.5), Inches(1.6), Inches(6.0), Inches(5.0), WHITE)
    add_textbox(slide, Inches(0.75), Inches(1.8), Inches(5.5), Inches(0.4),
                "Reading Spark (2 min)", size=18, bold=True, color=TEAL)
    add_textbox(slide, Inches(0.75), Inches(2.3), Inches(5.5), Inches(1.2),
                '"The library smelled like old paper and new ideas. '
                'Maya paused at the doorway, knowing that somewhere inside '
                'was the answer she needed — and maybe a few she didn\'t."',
                size=15, color=DARK)
    add_bullets(slide, Inches(0.75), Inches(3.7), Inches(5.5), Inches(2.5), [
        "What mood does this create?",
        "What can you infer about Maya?",
        "Which words create that feeling?",
    ], size=16)

    add_round_rect(slide, Inches(6.8), Inches(1.6), Inches(6.0), Inches(5.0), WHITE)
    add_textbox(slide, Inches(7.05), Inches(1.8), Inches(5.5), Inches(0.4),
                "Grammar Quick Fix (3 min)", size=18, bold=True, color=CORAL)
    add_textbox(slide, Inches(7.05), Inches(2.35), Inches(5.5), Inches(0.4),
                "Which sentence is correct?", size=15, bold=True, color=NAVY)
    add_bullets(slide, Inches(7.05), Inches(2.9), Inches(5.5), Inches(2.2), [
        "A) Neither of the players were ready.",
        "B) Neither of the players was ready.",
        "C) The team are winning the game.",
        "D) Everyone have finished the quiz.",
    ], size=15)
    add_textbox(slide, Inches(7.05), Inches(5.3), Inches(5.5), Inches(0.8),
                "Discuss: Why did you choose that answer?",
                size=14, bold=True, color=SOFT_GRAY)
    add_footer(slide, 4, time_label="0–5 min")


def warmup_answer_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide, CORAL)
    add_section_chip(slide, "WARM-UP KEY")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Warm-Up Answer + Why It Matters", size=28, bold=True, color=NAVY, font="Georgia")

    add_round_rect(slide, Inches(0.5), Inches(1.7), Inches(12.3), Inches(2.2), LIGHT_TEAL)
    add_textbox(slide, Inches(0.8), Inches(1.9), Inches(11.7), Inches(0.4),
                "Correct Answer: B — Neither of the players was ready.", size=20, bold=True, color=NAVY)
    add_textbox(slide, Inches(0.8), Inches(2.5), Inches(11.7), Inches(1.1),
                "Neither is singular. Even though “players” is plural, the subject is neither "
                "(one or the other — not both). So we use was, not were.\n"
                "Same rule for: each, everyone, somebody, either, nobody.",
                size=16, color=DARK)

    add_round_rect(slide, Inches(0.5), Inches(4.2), Inches(12.3), Inches(2.4), WHITE)
    add_textbox(slide, Inches(0.8), Inches(4.4), Inches(11.7), Inches(0.4),
                "Reading Spark Takeaways", size=18, bold=True, color=TEAL)
    add_bullets(slide, Inches(0.8), Inches(4.95), Inches(11.7), Inches(1.4), [
        "Mood = curious / hopeful / slightly mysterious",
        "Inference: Maya is searching for something important (evidence: “the answer she needed”)",
        "Strong readers notice word choice: smelled, paused, knowing — these paint a scene",
    ], size=16)
    add_footer(slide, 5, time_label="0–5 min")


def reading_overview_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "READING")
    add_time_chip(slide, "12 minutes")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "How Strong Readers Think", size=28, bold=True, color=NAVY, font="Georgia")
    add_textbox(slide, Inches(0.5), Inches(1.5), Inches(12), Inches(0.4),
                "Reading isn’t just saying the words — it’s building meaning.",
                size=16, color=SOFT_GRAY)

    pillars = [
        ("Before", "Preview", "Title, headings, first sentence — predict what you’ll learn"),
        ("During", "Question", "Ask: What’s happening? Why does this matter? What’s missing?"),
        ("During", "Annotate", "Underline claims, circle key words, star confusing parts"),
        ("After", "Summarize", "Say the main idea in one clear sentence using evidence"),
    ]
    for i, (phase, title, desc) in enumerate(pillars):
        left = Inches(0.5 + i * 3.2)
        add_round_rect(slide, left, Inches(2.2), Inches(3.0), Inches(4.3), WHITE)
        add_rect(slide, left, Inches(2.2), Inches(3.0), Inches(0.7), TEAL if i % 2 == 0 else NAVY)
        add_textbox(slide, left, Inches(2.35), Inches(3.0), Inches(0.45),
                    phase.upper(), size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, left + Inches(0.2), Inches(3.2), Inches(2.6), Inches(0.5),
                    title, size=22, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_textbox(slide, left + Inches(0.2), Inches(3.9), Inches(2.6), Inches(2.2),
                    desc, size=15, color=DARK, align=PP_ALIGN.CENTER)
    add_footer(slide, 6, time_label="5–17 min")


def reading_skills_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "READING SKILLS")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Four Skills We’ll Use Today", size=28, bold=True, color=NAVY, font="Georgia")

    skills = [
        ("Main Idea", "What is the passage mostly about?\nHint: Not every detail — the core message."),
        ("Supporting Details", "Which facts / examples prove the main idea?\nLook for because, for example, such as."),
        ("Inference", "What is suggested but not said?\nEvidence + your logic = inference."),
        ("Author’s Purpose", "Why was this written?\nTo inform, persuade, entertain, or explain?"),
    ]
    for i, (title, desc) in enumerate(skills):
        col = i % 2
        row = i // 2
        left = Inches(0.5 + col * 6.4)
        top = Inches(1.7 + row * 2.5)
        add_round_rect(slide, left, top, Inches(6.1), Inches(2.2), WHITE)
        add_rect(slide, left, top, Inches(6.1), Inches(0.55), CORAL if i % 2 else TEAL)
        add_textbox(slide, left + Inches(0.3), top + Inches(0.1), Inches(5.5), Inches(0.4),
                    title, size=18, bold=True, color=WHITE)
        add_textbox(slide, left + Inches(0.3), top + Inches(0.8), Inches(5.5), Inches(1.2),
                    desc, size=15, color=DARK)
    add_footer(slide, 7, time_label="5–17 min")


def annotation_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "STRATEGY")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Annotation Code (Use This Every Time)", size=28, bold=True, color=NAVY, font="Georgia")

    codes = [
        ("★", TEAL, "Star", "Main idea or most important claim"),
        ("—", NAVY, "Underline", "Strong supporting detail or evidence"),
        ("?", CORAL, "Question mark", "Confusing word, idea, or sentence"),
        ("→", RGBColor(0x8B, 0x5C, 0xF6), "Arrow", "Cause → effect or “this leads to…”"),
        ("\" \"", RGBColor(0xD9, 0x77, 0x06), "Quotes", "Words that show tone / feeling"),
        ("!", RGBColor(0x0E, 0xA5, 0xE9), "Exclamation", "Surprising fact or twist"),
    ]
    for i, (sym, color, name, use) in enumerate(codes):
        col = i % 3
        row = i // 3
        left = Inches(0.5 + col * 4.2)
        top = Inches(1.7 + row * 2.4)
        add_round_rect(slide, left, top, Inches(3.95), Inches(2.15), WHITE)
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, left + Inches(1.45), top + Inches(0.25), Inches(1.0), Inches(1.0)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = color
        circle.line.fill.background()
        add_textbox(slide, left + Inches(1.45), top + Inches(0.45), Inches(1.0), Inches(0.6),
                    sym, size=22, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, left + Inches(0.2), top + Inches(1.35), Inches(3.55), Inches(0.35),
                    name, size=16, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_textbox(slide, left + Inches(0.2), top + Inches(1.7), Inches(3.55), Inches(0.35),
                    use, size=13, color=SOFT_GRAY, align=PP_ALIGN.CENTER)
    add_footer(slide, 8, time_label="5–17 min")


def passage_intro_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "GUIDED PRACTICE")
    add_time_chip(slide, "15 minutes")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Guided Passage: “The Quiet Power of Curiosity”", size=26, bold=True,
                color=NAVY, font="Georgia")

    add_round_rect(slide, Inches(0.5), Inches(1.6), Inches(12.3), Inches(5.0), WHITE)
    add_textbox(slide, Inches(0.8), Inches(1.85), Inches(11.7), Inches(4.4),
                "Curiosity is often treated like a childhood habit — something people grow out of "
                "as they become “serious.” Yet some of the most useful breakthroughs begin with a "
                "simple question: Why?\n\n"
                "When students ask better questions, they don’t just collect facts; they connect ideas. "
                "A science learner who wonders why ice floats may discover density. A history student "
                "who asks why a law changed may uncover the people who fought for it. Curiosity turns "
                "passive reading into active thinking.\n\n"
                "Still, curiosity needs guidance. Without focus, questions scatter. Strong readers "
                "pause, notice patterns, and chase the most meaningful “why.” That habit — asking, "
                "checking evidence, and revising thinking — is what transforms good English into "
                "powerful understanding.",
                size=16, color=DARK)
    add_footer(slide, 9, time_label="17–32 min")


def passage_tasks_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "YOUR TASK")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Read → Annotate → Answer (10 min work + 5 min share)", size=24, bold=True,
                color=NAVY, font="Georgia")

    tasks = [
        ("Step 1", "Read once for the big picture. Don’t stop yet."),
        ("Step 2", "Read again with the annotation code (★ — ? →)."),
        ("Step 3", "Write a one-sentence main idea."),
        ("Step 4", "Answer the comprehension questions on the next slide."),
        ("Step 5", "Be ready to share one annotation and why you marked it."),
    ]
    for i, (step, text) in enumerate(tasks):
        top = Inches(1.6 + i * 0.95)
        add_round_rect(slide, Inches(0.5), top, Inches(12.3), Inches(0.85), WHITE)
        add_round_rect(slide, Inches(0.7), top + Inches(0.18), Inches(1.4), Inches(0.5), TEAL)
        add_textbox(slide, Inches(0.7), top + Inches(0.25), Inches(1.4), Inches(0.4),
                    step, size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(2.4), top + Inches(0.25), Inches(10), Inches(0.45),
                    text, size=17, color=DARK)
    add_footer(slide, 10, time_label="17–32 min")


def comprehension_q_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "COMPREHENSION")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Comprehension Questions", size=28, bold=True, color=NAVY, font="Georgia")

    qs = [
        ("Main Idea", "What is the central message of the passage?"),
        ("Detail", "According to the text, what happens when students ask better questions?"),
        ("Inference", "What does the author suggest about people who stop being curious?"),
        ("Purpose", "Is the author’s main purpose to entertain, inform, or persuade? Explain."),
        ("Evidence", "Quote one sentence that best supports your main-idea answer."),
    ]
    for i, (tag, q) in enumerate(qs):
        top = Inches(1.55 + i * 0.95)
        add_round_rect(slide, Inches(0.5), top, Inches(12.3), Inches(0.85), WHITE)
        add_round_rect(slide, Inches(0.7), top + Inches(0.2), Inches(1.8), Inches(0.45),
                       CORAL if i in (2, 3) else TEAL)
        add_textbox(slide, Inches(0.7), top + Inches(0.25), Inches(1.8), Inches(0.35),
                    tag, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, Inches(2.7), top + Inches(0.25), Inches(9.8), Inches(0.45),
                    f"{i + 1}. {q}", size=15, color=DARK)
    add_footer(slide, 11, time_label="17–32 min")


def comprehension_answers_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "ANSWER KEY")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Model Answers (Discuss Variations)", size=28, bold=True, color=NAVY, font="Georgia")

    answers = [
        "Main idea: Curiosity, when guided, turns reading into deeper understanding and stronger thinking.",
        "Detail: Students connect ideas instead of only collecting facts (examples: density; people behind laws).",
        "Inference: They may become less thoughtful / miss deeper learning — curiosity is treated as something to “grow out of.”",
        "Purpose: Mostly to inform/persuade readers that curiosity is valuable and should be practiced with focus.",
        "Evidence example: “Curiosity turns passive reading into active thinking.”",
    ]
    add_bullets(slide, Inches(0.5), Inches(1.6), Inches(12.3), Inches(5.0), answers, size=16, spacing=14)
    add_footer(slide, 12, time_label="17–32 min")


def grammar_intro_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide, CORAL)
    add_section_chip(slide, "GRAMMAR")
    add_time_chip(slide, "12 minutes")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Grammar Power-Up: The High-Impact Fixes", size=26, bold=True,
                color=NAVY, font="Georgia")
    add_textbox(slide, Inches(0.5), Inches(1.5), Inches(12), Inches(0.4),
                "You’re already good at English. These are the gaps that separate “good” from “polished.”",
                size=15, color=SOFT_GRAY)

    topics = [
        ("1", "Subject–Verb\nAgreement", "Match singular/plural subjects correctly"),
        ("2", "Comma Rules\nThat Matter", "Lists, FANBOYS, and introductory phrases"),
        ("3", "Confused\nWords", "its/it’s, affect/effect, there/their/they’re"),
        ("4", "Sentence\nPower", "Fix fragments, run-ons, and weak wording"),
    ]
    for i, (num, title, desc) in enumerate(topics):
        left = Inches(0.5 + i * 3.2)
        add_round_rect(slide, left, Inches(2.2), Inches(3.0), Inches(4.3), WHITE)
        add_rect(slide, left, Inches(2.2), Inches(3.0), Inches(1.0), CORAL)
        add_textbox(slide, left, Inches(2.4), Inches(3.0), Inches(0.6),
                    num, size=28, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_textbox(slide, left + Inches(0.15), Inches(3.5), Inches(2.7), Inches(1.2),
                    title, size=18, bold=True, color=NAVY, align=PP_ALIGN.CENTER)
        add_textbox(slide, left + Inches(0.2), Inches(5.0), Inches(2.6), Inches(1.1),
                    desc, size=14, color=SOFT_GRAY, align=PP_ALIGN.CENTER)
    add_footer(slide, 13, time_label="32–44 min")


def grammar_sva_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide, CORAL)
    add_section_chip(slide, "GRAMMAR 1")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Subject–Verb Agreement (Fast Rules)", size=28, bold=True, color=NAVY, font="Georgia")

    rules = [
        ("Indefinite pronouns are usually singular",
         "everyone, each, neither, somebody → was / has / is"),
        ("Ignore words between subject and verb",
         "The box of pencils is on the desk. (box = singular)"),
        ("Collective nouns: one unit = singular",
         "The team is ready. / The class has a project."),
        ("Compound subjects with and = plural",
         "Maya and Jordan are presenting today."),
    ]
    for i, (title, ex) in enumerate(rules):
        top = Inches(1.55 + i * 1.2)
        add_round_rect(slide, Inches(0.5), top, Inches(12.3), Inches(1.05), WHITE)
        add_textbox(slide, Inches(0.8), top + Inches(0.15), Inches(11.7), Inches(0.35),
                    title, size=17, bold=True, color=NAVY)
        add_textbox(slide, Inches(0.8), top + Inches(0.55), Inches(11.7), Inches(0.35),
                    ex, size=15, color=SOFT_GRAY)
    add_footer(slide, 14, time_label="32–44 min")


def grammar_commas_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide, CORAL)
    add_section_chip(slide, "GRAMMAR 2")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Comma Rules That Matter in Grade 8", size=28, bold=True, color=NAVY, font="Georgia")

    cards = [
        ("Lists", "We bought notebooks, pens, and folders.",
         "Use commas between items in a series."),
        ("FANBOYS", "I studied hard, but the quiz was tricky.",
         "Comma before for, and, nor, but, or, yet, so when joining two full sentences."),
        ("Intro Phrases", "After the bell rang, students settled down.",
         "Comma after an introductory word/phrase."),
        ("Nonessential Info", "My cousin, who lives in Texas, visits often.",
         "Commas around extra info you could remove."),
    ]
    for i, (title, ex, tip) in enumerate(cards):
        col = i % 2
        row = i // 2
        left = Inches(0.5 + col * 6.4)
        top = Inches(1.6 + row * 2.5)
        add_round_rect(slide, left, top, Inches(6.1), Inches(2.25), WHITE)
        add_textbox(slide, left + Inches(0.3), top + Inches(0.25), Inches(5.5), Inches(0.4),
                    title, size=18, bold=True, color=CORAL)
        add_textbox(slide, left + Inches(0.3), top + Inches(0.8), Inches(5.5), Inches(0.5),
                    ex, size=15, bold=True, color=NAVY)
        add_textbox(slide, left + Inches(0.3), top + Inches(1.4), Inches(5.5), Inches(0.6),
                    tip, size=14, color=SOFT_GRAY)
    add_footer(slide, 15, time_label="32–44 min")


def grammar_words_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide, CORAL)
    add_section_chip(slide, "GRAMMAR 3")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Commonly Confused Words", size=28, bold=True, color=NAVY, font="Georgia")

    rows = [
        ("its / it’s", "its = belonging  |  it’s = it is", "The dog wagged its tail. / It’s raining."),
        ("affect / effect", "affect = verb (to change)  |  effect = noun (result)",
         "Stress can affect sleep. / The effect was clear."),
        ("there / their / they’re", "place / belonging / they are",
         "They’re putting their bags over there."),
        ("your / you’re", "your = belonging  |  you’re = you are",
         "You’re improving your reading skills."),
        ("then / than", "then = time  |  than = comparison",
         "Read first, then revise. / Better than before."),
    ]
    # header
    add_rect(slide, Inches(0.5), Inches(1.55), Inches(12.3), Inches(0.55), NAVY)
    add_textbox(slide, Inches(0.7), Inches(1.65), Inches(3.0), Inches(0.35),
                "Pair", size=14, bold=True, color=WHITE)
    add_textbox(slide, Inches(3.8), Inches(1.65), Inches(4.5), Inches(0.35),
                "Quick Rule", size=14, bold=True, color=WHITE)
    add_textbox(slide, Inches(8.5), Inches(1.65), Inches(4.0), Inches(0.35),
                "Example", size=14, bold=True, color=WHITE)

    for i, (pair, rule, ex) in enumerate(rows):
        top = Inches(2.15 + i * 0.85)
        bg = WHITE if i % 2 == 0 else LIGHT_NAVY
        add_rect(slide, Inches(0.5), top, Inches(12.3), Inches(0.85), bg)
        add_textbox(slide, Inches(0.7), top + Inches(0.25), Inches(3.0), Inches(0.4),
                    pair, size=14, bold=True, color=NAVY)
        add_textbox(slide, Inches(3.8), top + Inches(0.25), Inches(4.5), Inches(0.4),
                    rule, size=13, color=DARK)
        add_textbox(slide, Inches(8.5), top + Inches(0.25), Inches(4.0), Inches(0.4),
                    ex, size=13, color=SOFT_GRAY)
    add_footer(slide, 16, time_label="32–44 min")


def grammar_practice_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide, CORAL)
    add_section_chip(slide, "PRACTICE")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Fix These Sentences (Live Practice)", size=28, bold=True, color=NAVY, font="Georgia")

    items = [
        "Everyone in the group have finished their annotations.",
        "The stack of books are tipping over the desk.",
        "Maya wanted to improve her writing but she forgot to revise.",
        "Its important to check your work before your done.",
        "Reading carefully has a bigger affect on scores then rushing.",
    ]
    for i, sent in enumerate(items):
        top = Inches(1.55 + i * 0.95)
        add_round_rect(slide, Inches(0.5), top, Inches(12.3), Inches(0.85), WHITE)
        add_textbox(slide, Inches(0.8), top + Inches(0.25), Inches(11.7), Inches(0.45),
                    f"{i + 1}.  {sent}", size=16, color=DARK)
    add_footer(slide, 17, time_label="32–44 min")


def grammar_answers_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide, CORAL)
    add_section_chip(slide, "ANSWER KEY")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Corrected Versions", size=28, bold=True, color=NAVY, font="Georgia")

    fixes = [
        "Everyone in the group has finished their annotations.  (everyone = singular → has)",
        "The stack of books is tipping over the desk.  (stack = singular → is)",
        "Maya wanted to improve her writing, but she forgot to revise.  (FANBOYS comma)",
        "It’s important to check your work before you’re done.  (it’s / you’re)",
        "Reading carefully has a bigger effect on scores than rushing.  (effect / than)",
    ]
    for i, text in enumerate(fixes):
        top = Inches(1.55 + i * 0.95)
        add_round_rect(slide, Inches(0.5), top, Inches(12.3), Inches(0.85), LIGHT_TEAL)
        add_textbox(slide, Inches(0.8), top + Inches(0.25), Inches(11.7), Inches(0.45),
                    f"{i + 1}.  {text}", size=15, color=DARK)
    add_footer(slide, 18, time_label="32–44 min")


def apply_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "APPLY & CREATE")
    add_time_chip(slide, "10 minutes")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Put It Together: Edit + Respond", size=28, bold=True, color=NAVY, font="Georgia")

    add_round_rect(slide, Inches(0.5), Inches(1.55), Inches(12.3), Inches(2.4), WHITE)
    add_textbox(slide, Inches(0.8), Inches(1.7), Inches(11.7), Inches(0.35),
                "Student Draft (needs help)", size=16, bold=True, color=CORAL)
    add_textbox(slide, Inches(0.8), Inches(2.15), Inches(11.7), Inches(1.5),
                "Curiosity help students learn better. When a reader ask good questions they find "
                "deeper meaning. Its more powerful then just memorizing facts. Strong readers pause "
                "check evidence and revise there thinking.",
                size=16, color=DARK)

    add_round_rect(slide, Inches(0.5), Inches(4.2), Inches(6.0), Inches(2.4), LIGHT_TEAL)
    add_textbox(slide, Inches(0.75), Inches(4.4), Inches(5.5), Inches(0.4),
                "Task A — Grammar Edit (5 min)", size=16, bold=True, color=NAVY)
    add_bullets(slide, Inches(0.75), Inches(4.95), Inches(5.5), Inches(1.4), [
        "Rewrite the paragraph correctly",
        "Fix verbs, commas, and word choice",
        "Keep the original meaning",
    ], size=14)

    add_round_rect(slide, Inches(6.8), Inches(4.2), Inches(6.0), Inches(2.4), LIGHT_CORAL)
    add_textbox(slide, Inches(7.05), Inches(4.4), Inches(5.5), Inches(0.4),
                "Task B — Reading Response (5 min)", size=16, bold=True, color=NAVY)
    add_bullets(slide, Inches(7.05), Inches(4.95), Inches(5.5), Inches(1.4), [
        "Write 3–4 sentences",
        "Explain one reading strategy you’ll use this week",
        "Include one text-based reason why it helps",
    ], size=14)
    add_footer(slide, 19, time_label="44–54 min")


def model_response_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "MODEL")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Sample Strong Response", size=28, bold=True, color=NAVY, font="Georgia")

    add_round_rect(slide, Inches(0.5), Inches(1.6), Inches(12.3), Inches(2.5), WHITE)
    add_textbox(slide, Inches(0.8), Inches(1.8), Inches(11.7), Inches(0.35),
                "Edited Paragraph", size=16, bold=True, color=TEAL)
    add_textbox(slide, Inches(0.8), Inches(2.3), Inches(11.7), Inches(1.5),
                "Curiosity helps students learn better. When a reader asks good questions, they find "
                "deeper meaning. It’s more powerful than just memorizing facts. Strong readers pause, "
                "check evidence, and revise their thinking.",
                size=16, color=DARK)

    add_round_rect(slide, Inches(0.5), Inches(4.35), Inches(12.3), Inches(2.25), LIGHT_NAVY)
    add_textbox(slide, Inches(0.8), Inches(4.55), Inches(11.7), Inches(0.35),
                "Sample Reading Response", size=16, bold=True, color=NAVY)
    add_textbox(slide, Inches(0.8), Inches(5.05), Inches(11.7), Inches(1.3),
                "This week I will annotate while I read by starring the main idea and underlining evidence. "
                "That helps me slow down and notice what the author is really saying. When I check my "
                "annotations later, I can answer questions with proof instead of guessing.",
                size=15, color=DARK)
    add_footer(slide, 20, time_label="44–54 min")


def wrapup_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "WRAP-UP")
    add_time_chip(slide, "6 minutes")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "Takeaways & Next Steps", size=28, bold=True, color=NAVY, font="Georgia")

    left_items = [
        ("Reading Habit", "Preview → Question → Annotate → Summarize"),
        ("Evidence First", "Every inference needs a text clue"),
        ("Grammar Targets", "Agreement, commas, confused words"),
    ]
    for i, (t, d) in enumerate(left_items):
        top = Inches(1.6 + i * 1.5)
        add_round_rect(slide, Inches(0.5), top, Inches(6.0), Inches(1.3), WHITE)
        add_textbox(slide, Inches(0.8), top + Inches(0.25), Inches(5.4), Inches(0.35),
                    t, size=18, bold=True, color=TEAL)
        add_textbox(slide, Inches(0.8), top + Inches(0.7), Inches(5.4), Inches(0.4),
                    d, size=14, color=DARK)

    add_round_rect(slide, Inches(6.8), Inches(1.6), Inches(6.0), Inches(4.8), NAVY)
    add_textbox(slide, Inches(7.1), Inches(1.9), Inches(5.4), Inches(0.4),
                "This Week’s Challenge", size=18, bold=True, color=TEAL)
    add_bullets(slide, Inches(7.1), Inches(2.5), Inches(5.4), Inches(3.5), [
        "Read one article / short story for 15 minutes",
        "Annotate with ★ — ?",
        "Write a 1-sentence main idea",
        "Edit 5 sentences for grammar",
        "Bring one “tricky” sentence next time",
    ], size=15, color=WHITE, spacing=12)
    add_footer(slide, 21, time_label="54–60 min")


def closing_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, NAVY)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.2), TEAL)
    add_textbox(slide, Inches(0.8), Inches(2.0), Inches(11.7), Inches(0.5),
                "You’re already good at English.", size=22, color=TEAL, align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.8), Inches(2.7), Inches(11.7), Inches(1.2),
                "Now let’s make your reading\nand grammar excellent.", size=36, bold=True,
                color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    add_textbox(slide, Inches(0.8), Inches(4.4), Inches(11.7), Inches(0.5),
                "Questions? Practice tip? Favorite sentence from today?",
                size=18, color=RGBColor(0xC5, 0xD4, 0xE8), align=PP_ALIGN.CENTER)
    add_textbox(slide, Inches(0.8), Inches(5.5), Inches(11.7), Inches(0.4),
                "Thank you — great work today!", size=20, bold=True, color=WHITE,
                align=PP_ALIGN.CENTER)


def facilitator_notes_slide():
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, CREAM)
    add_accent_bar(slide)
    add_section_chip(slide, "FACILITATOR NOTES")
    add_textbox(slide, Inches(0.5), Inches(0.9), Inches(12), Inches(0.5),
                "How to Run This Demo Smoothly", size=28, bold=True, color=NAVY, font="Georgia")

    notes = [
        "Keep energy high: praise process (“smart annotation”) not just correct answers.",
        "If she finishes early, ask her to write a tougher inference question for you.",
        "If stuck on grammar, have her read the sentence aloud — ears catch many errors.",
        "Differentiate: she’s strong overall — push for evidence quotes and precise word choice.",
        "Materials needed: printed passage (optional), pencil, notebook, timer.",
        "Tone: supportive coach, not lecture. Aim for 40% teaching / 60% doing.",
    ]
    add_numbered(slide, Inches(0.5), Inches(1.6), Inches(12.3), Inches(5.0), notes, size=16, spacing=12)
    add_footer(slide, 23, total=23, time_label="Teacher")


# Build all slides
title_slide()
agenda_slide()
goals_slide()
warmup_slide()
warmup_answer_slide()
reading_overview_slide()
reading_skills_slide()
annotation_slide()
passage_intro_slide()
passage_tasks_slide()
comprehension_q_slide()
comprehension_answers_slide()
grammar_intro_slide()
grammar_sva_slide()
grammar_commas_slide()
grammar_words_slide()
grammar_practice_slide()
grammar_answers_slide()
apply_slide()
model_response_slide()
wrapup_slide()
closing_slide()
facilitator_notes_slide()

out = r"C:\Users\bhushaja\Downloads\shaip\Grade8_Reading_Grammar_60min_Demo.pptx"
prs.save(out)
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
