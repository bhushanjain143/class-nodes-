"""Grade 5 English lesson - 90 minutes, 69 slides, no speaker notes.

"The Science Lab Mystery" - the student is a Junior Science Detective working
through the Science Lab, the Observation Desk, the Experiment Station, the
Research Table, the Mystery File and Detective HQ. Reading comprehension and
grammar carry equal weight, with subject-verb agreement as the grammar spine.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Science Hint strip, and slides 66-69 hold the support
system, the assessment table and a two-page answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x12, 0x18, 0x21)
DARK = RGBColor(0x26, 0x31, 0x3C)
SOFT = RGBColor(0x76, 0x82, 0x8E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAGE = RGBColor(0xF6, 0xF7, 0xF9)
TEAL = RGBColor(0x0E, 0x7C, 0x86)
INDIGO = RGBColor(0x2F, 0x4B, 0x9B)
VIOLET = RGBColor(0x6D, 0x4A, 0xA6)
AMBER = RGBColor(0xD1, 0x8B, 0x0E)
CRIMSON = RGBColor(0xC0, 0x3B, 0x3B)
ROSE = RGBColor(0xB0, 0x3C, 0x72)
L_TEAL = RGBColor(0xE2, 0xF1, 0xF2)
L_INDIGO = RGBColor(0xE8, 0xEC, 0xF7)
L_VIOLET = RGBColor(0xEF, 0xEA, 0xF8)
L_AMBER = RGBColor(0xFC, 0xF2, 0xDD)
L_CRIMSON = RGBColor(0xFA, 0xE9, 0xE9)
L_ROSE = RGBColor(0xF9, 0xE9, 0xF1)
L_GREY = RGBColor(0xF0, 0xF1, 0xF3)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 69
_counter = {"n": 0}

# ------------------------------------------------------------------ content

STOPS = [("🧪", "Science Lab", "Vocabulary", "0–17 min", TEAL, L_TEAL),
         ("🔍", "Observation Desk", "Subject + verb", "17–27 min", INDIGO,
          L_INDIGO),
         ("🧫", "Experiment Station", "Parts of speech", "27–40 min", VIOLET,
          L_VIOLET),
         ("📋", "Research Table", "Verb tenses", "40–50 min", AMBER, L_AMBER),
         ("📖", "Mystery File", "Reading & clues", "50–70 min", CRIMSON,
          L_CRIMSON),
         ("🏆", "Detective HQ", "Edit, report, solve", "70–90 min", ROSE, L_ROSE)]

MISSIONS = [("🔎", "Read the clues", "A real mystery passage, in five parts."),
            ("✏️", "Fix the grammar", "Subjects, verbs and tenses."),
            ("🧩", "Solve word puzzles", "Vocabulary and sentence games."),
            ("🎤", "Report your findings", "Explain the case out loud.")]

LAB_SCENE = [("🧪", "test tubes"), ("🔬", "microscope"), ("🧫", "beaker"),
             ("🔎", "magnifier"), ("📋", "clipboard"), ("🔐", "locked case")]

WARM_QS = [("👀", "What do you notice in the lab?"),
           ("❓", "What looks unusual to you?"),
           ("💭", "What do you think happened here?"),
           ("🔎", "Which detail would you check first?")]

WARM_STARTERS = ["I notice ______.", "I think ______.", "Maybe ______."]

VOCAB_CARDS = [
    ("OBSERVE", "🔍", "to look at something very carefully",
     "Scientists observe the experiment before they touch it.",
     "What do you observe on this table?", TEAL, L_TEAL),
    ("DISCOVER", "💡", "to find something new",
     "The students discover a clue near the sink.",
     "Tell me something you discovered this week.", AMBER, L_AMBER),
    ("EVIDENCE", "🔎", "information that helps us know what happened",
     "The wet mark on the floor is evidence.",
     "Is a guess evidence? Why not?", INDIGO, L_INDIGO),
    ("PREDICT", "🔮", "to say what you think will happen",
     "Can you predict the result of the experiment?",
     "Predict what happens if we leave ice in the sun.", VIOLET, L_VIOLET),
    ("RESULT", "📊", "what happens at the end of an experiment",
     "The result of the test surprised the whole class.",
     "What result did you expect?", ROSE, L_ROSE),
]

VOCAB_BANK = [("experiment", "🧪"), ("observe", "🔍"), ("discover", "💡"),
              ("evidence", "🔎"), ("predict", "🔮"), ("measure", "📏"),
              ("solution", "🧫"), ("mixture", "🥣"), ("sample", "🧴"),
              ("scientist", "🔬"), ("carefully", "🐌"), ("different", "↔️"),
              ("similar", "⚖️"), ("mystery", "🔐"), ("result", "📊"),
              ("investigate", "🕵️")]

WORD_MATCH = [("OBSERVE", ["🔍", "🥣", "📊"]), ("DISCOVER", ["📏", "💡", "🔐"]),
              ("EVIDENCE", ["🔎", "🧴", "🔮"]), ("PREDICT", ["🧫", "📊", "🔮"]),
              ("RESULT", ["📊", "🔬", "🧪"])]

WORD_OR_NOT = [("experiment", True), ("banana", False), ("evidence", True),
               ("discover", True), ("pencil", False), ("predict", True)]

SAY_IT = [("🔍", "observe"), ("🔎", "evidence"), ("📊", "result")]
SAY_STARTERS = ["I can ______.", "The scientist ______.",
                "The evidence shows ______."]

SUBJECT_VERB = [("The", False), ("scientist", "subject"), ("works", "verb"),
                ("in", False), ("the", False), ("lab.", False)]

SVA_PAIRS = [
    ("scientist", "The scientist works.", "works", "one scientist",
     "scientists", "The scientists work.", "work", "more than one scientist",
     "🔬"),
    ("experiment", "The experiment looks strange.", "looks", "one experiment",
     "experiments", "The experiments look strange.", "look",
     "more than one experiment", "🧪"),
]

SVA_RULE = [("1", "Find the subject", "Who or what is doing the action?",
             "the scientist", INDIGO),
            ("2", "One or many?", "Is the subject singular or plural?",
             "one = singular", TEAL),
            ("3", "Match the verb", "Singular subject takes the -s verb.",
             "scientist works", VIOLET),
            ("4", "Check it out loud", "Read the sentence back to yourself.",
             "does it sound right?", AMBER)]

VERB_CHOICE_A = [("The scientist ____ carefully.", ["work", "works"], "works",
                  "one scientist"),
                 ("The scientists ____ carefully.", ["works", "work"], "work",
                  "more than one"),
                 ("The experiment ____ strange.", ["look", "looks"], "looks",
                  "one experiment")]
VERB_CHOICE_B = [("The experiments ____ strange.", ["looks", "look"], "look",
                  "more than one"),
                 ("The student ____ the solution.", ["mix", "mixes"], "mixes",
                  "one student"),
                 ("The students ____ the solution.", ["mixes", "mix"], "mix",
                  "more than one")]

FIX_A = [("The scientist work in the lab.", "The scientist works in the lab.",
          "one scientist → works"),
         ("The students studies the result.", "The students study the result.",
          "more than one student → study")]
FIX_B = [("The experiments looks cold.", "The experiments look cold.",
          "more than one experiment → look"),
         ("My teacher check the samples.", "My teacher checks the samples.",
          "one teacher → checks")]

POS_LEGEND = [("NOUN", "a person, place or thing", "scientist · lab · sample",
               INDIGO, L_INDIGO),
              ("VERB", "the action word", "mixes · measures · checks", TEAL,
               L_TEAL),
              ("ADJECTIVE", "describes a noun", "careful · strange · cold",
               VIOLET, L_VIOLET),
              ("ADVERB", "describes the verb, often ends -ly",
               "carefully · quickly · slowly", AMBER, L_AMBER)]

POS_TASKS = [
    ("The student mixes the solution.",
     [("NOUN", INDIGO, L_INDIGO), ("VERB", TEAL, L_TEAL)]),
    ("The careful scientist measures the liquid.",
     [("NOUN", INDIGO, L_INDIGO), ("VERB", TEAL, L_TEAL),
      ("ADJECTIVE", VIOLET, L_VIOLET)]),
    ("The scientist carefully checks the result.",
     [("NOUN", INDIGO, L_INDIGO), ("VERB", TEAL, L_TEAL),
      ("ADVERB", AMBER, L_AMBER)]),
]

SORT_BEAKERS = [("NOUN", INDIGO, L_INDIGO), ("VERB", TEAL, L_TEAL),
                ("ADJECTIVE", VIOLET, L_VIOLET), ("ADVERB", AMBER, L_AMBER)]
SORT_WORDS = ["scientist", "quickly", "experiment", "careful", "measure",
              "bright"]

FAST_LAB = [("Find the VERB", "The scientist carefully checks the experiment.",
             ["scientist", "carefully", "checks"], TEAL, L_TEAL),
            ("Find the ADJECTIVE", "The strange experiment surprised everyone.",
             ["strange", "experiment", "surprised"], VIOLET, L_VIOLET),
            ("Find the ADVERB", "The scientist carefully measured the liquid.",
             ["scientist", "carefully", "measured"], AMBER, L_AMBER)]

SCIENTIST_SAYS = [("📓", "Scientist says touch your notebook.", "move"),
                  ("🥣", "Scientist says pretend to mix.", "move"),
                  ("✏️", "Scientist says point to the VERB.", "grammar"),
                  ("🔤", "Scientist says point to the NOUN.", "grammar"),
                  ("🧊", "Scientist says freeze!", "move"),
                  ("🔎", "Scientist says find the ADJECTIVE.", "grammar")]

FREEZE_SENTENCE = "The careful scientist quickly checked the strange sample."
FREEZE_TASKS = [("Find an ADJECTIVE", VIOLET, L_VIOLET),
                ("Find the VERB", TEAL, L_TEAL),
                ("Find the ADVERB", AMBER, L_AMBER),
                ("Find a NOUN", INDIGO, L_INDIGO)]

TENSES = [("PAST", "yesterday", "checked",
           "The scientist checked the experiment.", VIOLET, L_VIOLET),
          ("NOW", "every day", "checks",
           "The scientist checks the experiment.", TEAL, L_TEAL),
          ("FUTURE", "tomorrow", "will check",
           "The scientist will check the experiment.", AMBER, L_AMBER)]

TENSE_ID = [("The scientist measures the liquid.",
             ["Past", "Present", "Future"], "Present", "find the verb first"),
            ("The students recorded the result.",
             ["Past", "Present", "Future"], "Past", "find the verb first"),
            ("We will test the sample tomorrow.",
             ["Past", "Present", "Future"], "Future",
             "find the verb — it is two words here")]

TO_PAST = [("The scientist measures the liquid.", "measured"),
           ("The students record the result.", "recorded")]
TO_FUTURE = [("The scientist measures the liquid.", "will measure"),
             ("The students record the result.", "will record")]

TENSE_CHOICE = [("Yesterday, the scientist ____ the sample.",
                 ["checks", "checked"], "checked", "underline the time word"),
                ("Tomorrow, we ____ the result.",
                 ["recorded", "will record"], "will record",
                 "underline the time word")]

STORY_PREDICT = ["Someone took it by mistake.", "It rolled off the table.",
                 "A teacher moved it on purpose."]

STORY = [
    ("Part 1", "🧪", TEAL, L_TEAL,
     ["Mr. Alvarez's science class was getting ready for an important "
      "experiment. On the long table at the front of the room, he placed six "
      "small jars.",
      "Each jar had a neat white label. Inside the jars were different samples "
      "the class would study that morning.",
      "Mr. Alvarez reminded everyone to handle the samples carefully. Then he "
      "sent the students to the supply closet to collect goggles and "
      "notebooks."],
     "How many jars did Mr. Alvarez place on the table?"),
    ("Part 2", "🔍", INDIGO, L_INDIGO,
     ["When the students came back, something was wrong. Maya counted the jars "
      "out loud. \"One, two, three, four, five.\"",
      "She counted again, more slowly. There were only five jars on the table. "
      "The jar labeled Sample C was gone.",
      "For a moment nobody spoke. Then Mr. Alvarez smiled and said, \"You are "
      "scientists. Observe the room and gather evidence.\""],
     "Which jar was missing?"),
    ("Part 3", "🔎", VIOLET, L_VIOLET,
     ["The students began to investigate. Maya walked slowly around the table "
      "and wrote down everything she noticed.",
      "Ben found a small wet mark on the floor near the sink. Priya noticed "
      "that the door of the tall cabinet was open, even though it had been "
      "closed before.",
      "Each clue was written carefully in a notebook. The students were not "
      "guessing now. They were collecting evidence."],
     "Name one clue the students found."),
    ("Part 4", "🧊", AMBER, L_AMBER,
     ["\"A wet mark could mean someone carried something cold,\" Maya said. "
      "\"Cold things sweat when they sit in warm air.\"",
      "Ben nodded and looked at the open cabinet. Inside the cabinet was the "
      "small refrigerator the class used for storage.",
      "Priya opened the refrigerator door. There, on the middle shelf, sat the "
      "missing jar. The label read Sample C."],
     "Why does a cold jar leave a wet mark?"),
    ("Part 5", "🏆", ROSE, L_ROSE,
     ["Mr. Alvarez explained that he had moved the jar himself. Sample C had "
      "to stay cold, or the result of the experiment would be wrong.",
      "\"You did not guess,\" he said. \"You observed, you collected evidence, "
      "and you reached a conclusion. That is how real scientists solve a "
      "mystery.\"",
      "Maya wrote one last line in her notebook: Careful observation finds the "
      "answer."],
     "Why was the sample moved?"),
]

MAIN_IDEA = ["How to label jars in a science lab",
             "Students use observation and evidence to find a missing sample",
             "A class trip to the supply closet"]
SUMMARY_FRAME = ["In the story, ______ went missing.",
                 "The students ______ the room and found ______.",
                 "In the end, ______ because ______."]

CLUE_Q_A = [("🧴", "What went missing?", ["A book", "A sample", "A pencil"],
             "Part 2"),
            ("💧", "What clue did Ben find?",
             ["A wet mark near the sink", "An open window", "A broken jar"],
             "Part 3")]
CLUE_Q_B = [("❄️", "Why was the sample moved?",
             ["It had to stay cold", "It was broken", "It was not needed"],
             "Part 5"),
            ("🚪", "Where did they find it?",
             ["In the refrigerator", "In the sink", "In the closet"], "Part 4")]

SEQUENCE = [("A", "Students notice clues.", 2), ("B", "A sample is missing.", 1),
            ("C", "They investigate.", 3), ("D", "They find the sample.", 4)]

EVIDENCE_QS = [("How do you know the sample was moved on purpose?", "Part 5",
                CRIMSON, L_CRIMSON),
               ("How do you know the jar had been cold?", "Part 4", AMBER,
                L_AMBER),
               ("How do you know the students did not guess?", "Part 3", VIOLET,
                L_VIOLET)]

TRUE_FALSE = [("The students immediately knew where the sample was.", "FALSE"),
              ("Ben found a wet mark on the floor near the sink.", "TRUE"),
              ("Mr. Alvarez had moved the sample himself.", "TRUE")]

INFERENCE_OPTIONS = ["They gave up quickly.",
                     "They investigated carefully.",
                     "They ignored the clues."]

EDITOR_A = [("The students was ready.", "The students were ready.",
             "plural subject → were"),
            ("The scientist check the sample.",
             "The scientist checks the sample.", "singular subject → checks")]

EDITOR_POS = [("The students carefully looked at the clues.",
               [("NOUN", INDIGO, L_INDIGO), ("VERB", TEAL, L_TEAL),
                ("ADVERB", AMBER, L_AMBER)]),
              ("The strange sample was cold.",
               [("ADJECTIVE", VIOLET, L_VIOLET), ("NOUN", INDIGO, L_INDIGO),
                ("VERB", TEAL, L_TEAL)])]

EDITOR_CHECKS = [("🔠", "Capital letter", "Every sentence starts with one."),
                 ("⏹️", "End punctuation", "A statement ends with a period."),
                 ("🔎", "Names", "Proper nouns need a capital too.")]

REPORT_FRAMES = [("❓", "WHAT HAPPENED?", "Say it in one sentence.", CRIMSON,
                  L_CRIMSON),
                 ("🔎", "WHAT WAS THE CLUE?", "Name one piece of evidence.",
                  VIOLET, L_VIOLET),
                 ("📊", "WHAT WAS THE RESULT?", "How did it end, and why?",
                  TEAL, L_TEAL)]

REPORT_STARTERS = ["The sample was ______.", "The students noticed ______.",
                   "They discovered ______.", "This matters because ______."]

SPEAK_CHECKS = [("1️⃣", "Say what happened."), ("2️⃣", "Name one clue."),
                ("3️⃣", "Say where it was found."),
                ("4️⃣", "Say why it had been moved.")]

FINAL_GRAMMAR = [("The scientist ____ the sample.", ["check", "checks"],
                  "checks"),
                 ("The students ____ the result.", ["records", "record"],
                  "record")]
FINAL_VOCAB = [("What does EVIDENCE mean?",
                ["A clue or information", "A kind of food", "A game"],
                "A clue or information"),
               ("What does PREDICT mean?",
                ["To look closely", "To say what will happen", "To measure"],
                "To say what will happen")]
FINAL_READING = "What was missing from the table, and where was it found?"
FINAL_TENSE = ("The scientist checks the sample.", "The scientist checked the "
               "sample.")
FINAL_SPEAK = "Scientists ______ because ______."

ACHIEVEMENTS = [("🔬", "Grammar Detective"), ("📖", "Reading Detective"),
                ("🔎", "Evidence Finder"), ("🧠", "Vocabulary Explorer"),
                ("🎤", "Junior Reporter")]

CAN_DO = [("📖", "Reading fluency", "five parts, read aloud"),
          ("💡", "Main idea", "observation and evidence"),
          ("🔎", "Supporting details", "wet mark · open cabinet"),
          ("🧾", "Evidence", "pointing to the proof"),
          ("🧠", "Inference", "they investigated carefully"),
          ("🔢", "Sequencing", "B → A → C → D"),
          ("📚", "Vocabulary", "16 science words"),
          ("✏️", "Subject-verb agreement", "scientist works · scientists work"),
          ("⏳", "Verb tense", "checks · checked · will check"),
          ("🏷️", "Parts of speech", "noun · verb · adjective · adverb"),
          ("🎤", "Speaking", "you reported the case")]

EXTRA_GAMES = [("🧪", "WORD LAB", "E-V-D-I-E-N-C-E — which word is it?",
                ["EVIDENCE", "EVIDENT", "EVENING"], TEAL, L_TEAL),
               ("🔬", "GRAMMAR OR NOT?", "Correct, or needs fixing?",
                ["The samples was cold.", "The sample was cold."], INDIGO,
                L_INDIGO),
               ("🔎", "FIND THE WORD", "Find \"sample\" in Part 2.",
                ["read it aloud", "then read the sentence"], VIOLET, L_VIOLET),
               ("🧩", "SENTENCE PUZZLE", "scientist / the / carefully / works",
                ["The scientist works", "carefully."], AMBER, L_AMBER),
               ("🎯", "CLUE OR NOT?", "Does it connect to the mystery?",
                ["Wet mark near sink", "A blue notebook"], ROSE, L_ROSE)]

SUPPORT_LEVELS = [("🟢", "SCIENCE HINT", "Visual or verbal clue",
                   "Point at the part of the sentence that matters.", TEAL,
                   L_TEAL),
                  ("🟡", "DETECTIVE MISSION", "She solves it with limited help",
                   "Ask a smaller question, then go quiet.", AMBER, L_AMBER),
                  ("⭐", "EXPERT CHALLENGE", "She explains or extends the answer",
                   "Ask her to build a new sentence from the rule.", VIOLET,
                   L_VIOLET)]

SUPPORT_LADDER = [("SUPPORT 1", "Give a visual clue.", TEAL),
                  ("SUPPORT 2", "Highlight the important word.", INDIGO),
                  ("SUPPORT 3", "Read the sentence together.", VIOLET),
                  ("SUPPORT 4", "Ask a smaller question.", AMBER),
                  ("SUPPORT 5", "Model it, then she reapplies it.", ROSE)]

HINT_BANK = ["Look at the subject.", "Who is doing the action?",
             "Look at the verb.", "Read the sentence again.",
             "Look for a clue in Part 2.", "What word gives you evidence?"]

PRAISE = ["\"Good observation.\"", "\"Let's investigate that word.\"",
          "\"You found an important clue.\"", "\"Try it one more time.\"",
          "\"Excellent correction.\"", "\"Show me the evidence.\""]

RUBRIC_SKILLS = ["Reading fluency", "Main idea", "Details", "Evidence",
                 "Inference", "Vocabulary", "Subject-verb agreement",
                 "Verb tense", "Parts of speech", "Sentence editing",
                 "Speaking"]

REVIEW_NEXT = ["Words to review", "Grammar to review",
               "Reading skill to practice"]

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


def paragraph(slide, l, t, w, h, text, size=17, color=INK):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    set_run(r, size, False, color, "Calibri")
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
             RGBColor(0xE0, 0xE3, 0xE6))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), TEAL)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), INK)
    msg = "🧪 The Science Lab Mystery  |  Grade 5  |  90 min"
    if stop:
        msg += f"  |  {stop}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg,
       size=10, color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32),
       f"{n} / {TOTAL}", size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, stop, accent=TEAL, bg=PAGE):
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, bg)
    add_rect(slide, Inches(0), Inches(0), Inches(0.16), prs.slide_height, accent)
    add_round(slide, Inches(0.4), Inches(0.26), Inches(3.5), Inches(0.42), accent)
    tb(slide, Inches(0.4), Inches(0.31), Inches(3.5), Inches(0.34), tag, size=12,
       bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    if timing:
        add_round(slide, Inches(10.3), Inches(0.26), Inches(2.65), Inches(0.42),
                  INK)
        tb(slide, Inches(10.3), Inches(0.31), Inches(2.65), Inches(0.34), timing,
           size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title,
       size=28, bold=True, color=INK, font="Georgia")
    footer(slide, n, timing, stop)
    fade(slide)
    return slide, n


def one_task(slide, text, color=TEAL, top=1.34):
    """States the single job of this slide in one plain line."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text,
              size=16, bold=True, color=color)


def hint(slide, text, top=6.42, label="🔬 SCIENCE HINT", fill=L_TEAL,
         color=TEAL):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.6), Inches(0.34), label,
       size=12, bold=True, color=color)
    tb(slide, Inches(3.5), Inches(top + 0.05), Inches(9.1), Inches(0.36), text,
       size=13, bold=True, color=INK)


def model_bar(slide, top=6.42):
    steps = [("I MODEL", "I show you one", CRIMSON),
             ("WE TRY", "we do one together", AMBER),
             ("YOU GO", "your turn", TEAL)]
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


def pos_legend(slide, top=6.42):
    for i, (name, _d, _e, color, _l) in enumerate(POS_LEGEND):
        left = Inches(0.5 + i * 3.11)
        add_round(slide, left, Inches(top), Inches(2.9), Inches(0.46), color)
        tb(slide, left, Inches(top + 0.05), Inches(2.9), Inches(0.36), name,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def choice_rows(slide, items, start_index, accent, top_start=1.9, gap=1.5,
                height=1.34):
    """A sentence with a blank, then two or three answer buttons."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start, gap,
                         height)
    for (sentence, options, _answer, why), top in zip(items, tops):
        tb(slide, Inches(1.45), top + Inches(0.26), Inches(5.6), Inches(0.6),
           sentence, size=22, bold=True, color=INK)
        tb(slide, Inches(1.45), top + Inches(0.88), Inches(5.6), Inches(0.34),
           f"clue: {why}", size=11, color=SOFT)
        wide = len(options) <= 2
        width, step, size = (2.5, 2.7, 19) if wide else (1.78, 1.85, 16)
        for j, opt in enumerate(options):
            left = Inches(7.3 + j * step)
            add_round(slide, left, top + Inches(0.32), Inches(width),
                      Inches(0.7), L_GREY)
            tb(slide, left, top + Inches(0.42), Inches(width), Inches(0.5),
               f"{chr(65 + j)}.  {opt}", size=size, bold=True, color=INK,
               align=PP_ALIGN.CENTER)


def fix_rows(slide, items, start_index, accent, write_line=True):
    """A broken sentence, the reason, and a line to write the fix on."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=2.0,
                         gap=2.15, height=1.9)
    for (wrong, _right, why), top in zip(items, tops):
        add_round(slide, Inches(1.4), top + Inches(0.22), Inches(1.5),
                  Inches(0.44), L_CRIMSON)
        tb(slide, Inches(1.4), top + Inches(0.27), Inches(1.5), Inches(0.34),
           "❌ BROKEN", size=11, bold=True, color=CRIMSON,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(3.1), top + Inches(0.2), Inches(7.0), Inches(0.5),
           wrong, size=24, bold=True, color=INK)
        add_round(slide, Inches(10.3), top + Inches(0.22), Inches(2.3),
                  Inches(0.44), L_GREY)
        tb(slide, Inches(10.3), top + Inches(0.27), Inches(2.3), Inches(0.34),
           why, size=10, color=SOFT, align=PP_ALIGN.CENTER)
        if write_line:
            add_round(slide, Inches(1.4), top + Inches(0.85), Inches(11.2),
                      Inches(0.78), L_TEAL)
            tb(slide, Inches(1.65), top + Inches(0.92), Inches(2.0),
               Inches(0.3), "✅ NOW FIX IT", size=10, bold=True, color=TEAL)
            tb(slide, Inches(1.65), top + Inches(1.2), Inches(10.7),
               Inches(0.36),
               "________________________________________________________"
               "____________", size=13, color=SOFT)


def pos_task_slide(sentence, asks, title, timing, stop, accent, sub):
    """A plain sentence plus one labelled blank per part of speech."""
    slide, n = new_slide(title, "GRAMMAR", timing, stop, accent)
    one_task(slide, sub, accent)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(1.5),
              WHITE)
    tb(slide, Inches(0.5), Inches(2.25), Inches(12.35), Inches(0.8), sentence,
       size=32, bold=True, color=INK, align=PP_ALIGN.CENTER, font="Georgia")
    width = 12.35 / len(asks) - 0.2
    for i, (name, color, light) in enumerate(asks):
        left = Inches(0.5 + i * (width + 0.2))
        add_round(slide, left, Inches(3.55), Inches(width), Inches(2.1), light)
        add_round(slide, left, Inches(3.55), Inches(width), Inches(0.5), color)
        tb(slide, left, Inches(3.63), Inches(width), Inches(0.36), name,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(4.15), Inches(width), Inches(0.4),
           "write the word here", size=11, color=SOFT, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.6), Inches(width - 0.6),
                  Inches(0.8), WHITE)
        tb(slide, left + Inches(0.3), Inches(4.78), Inches(width - 0.6),
           Inches(0.5), "____________", size=20, color=SOFT,
           align=PP_ALIGN.CENTER)
    pos_legend(slide, 5.88)
    return slide, n


def question_rows(slide, items, start_index, accent=CRIMSON, top_start=1.95,
                  gap=2.16, height=1.92):
    for i, (emoji, question, options, where) in enumerate(items):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.7), Inches(0.52),
                 Inches(0.52), accent)
        tb(slide, Inches(0.75), top + Inches(0.76), Inches(0.52), Inches(0.4),
           str(start_index + i), size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.45), top + Inches(0.6), Inches(1.0), Inches(0.72),
           emoji, size=30, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.6), top + Inches(0.5), Inches(3.4), Inches(0.66),
           question, size=19, bold=True, color=INK)
        tb(slide, Inches(2.6), top + Inches(1.18), Inches(3.4), Inches(0.36),
           f"the answer is in {where}", size=11, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(6.3 + j * 2.2)
            add_round(slide, left, top + Inches(0.5), Inches(2.0),
                      Inches(0.92), L_CRIMSON)
            tb(slide, left, top + Inches(0.62), Inches(2.0), Inches(0.7), opt,
               size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)


def story_slide(part, timing):
    """One passage page: picture and quick question left, three blocks right."""
    label, emoji, color, light, blocks, quick = part
    slide, n = new_slide(f"📖 The Mystery of the Missing Sample — {label}",
                         "READING", timing, "Mystery File", color)
    add_round(slide, Inches(0.5), Inches(1.4), Inches(3.8), Inches(4.75), light)
    tb(slide, Inches(0.5), Inches(1.75), Inches(3.8), Inches(1.15), emoji,
       size=58, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(1.35), Inches(3.1), Inches(2.1), Inches(0.55),
              WHITE)
    tb(slide, Inches(1.35), Inches(3.21), Inches(2.1), Inches(0.4),
       label.upper(), size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.8), Inches(3.95), Inches(3.2), Inches(1.85),
              WHITE)
    tb(slide, Inches(1.0), Inches(4.1), Inches(2.8), Inches(0.34),
       "🔎 QUICK QUESTION", size=11, bold=True, color=color)
    tb(slide, Inches(1.0), Inches(4.5), Inches(2.8), Inches(1.2), quick,
       size=15, bold=True, color=INK)
    add_round(slide, Inches(4.6), Inches(1.4), Inches(8.25), Inches(4.75),
              WHITE)
    for i, block in enumerate(blocks):
        paragraph(slide, Inches(5.0), Inches(1.68 + i * 1.5), Inches(7.5),
                  Inches(1.4), block, size=17)
    model_bar(slide)
    return slide, n


def badge_slide(title, timing, stop, emoji, badge_label, sub, lines, next_line,
                accent=TEAL, light=L_TEAL):
    """Stop-complete slide: the badge beside what was just mastered."""
    slide, n = new_slide(title, "CASE FILE UPDATED", timing, stop, accent)
    one_task(slide, sub, accent)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.3), Inches(4.4), light)
    tb(slide, Inches(0.5), Inches(2.6), Inches(5.3), Inches(1.7), emoji,
       size=88, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.3), Inches(0.7), badge_label,
       size=25, bold=True, color=accent, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.7), Inches(0.6),
       "⭐ detective badge earned", size=15, color=SOFT, align=PP_ALIGN.CENTER)
    for i, (icon, line) in enumerate(lines):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.15), top, Inches(6.7), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.45), top + Inches(0.2), Inches(0.56),
                 Inches(0.56), WHITE)
        tb(slide, Inches(6.45), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.25), top + Inches(0.24), Inches(5.3), Inches(0.5),
           line, size=15, bold=True, color=INK)
    hint(slide, next_line, 6.45, "➡️ NEXT LOCATION", L_AMBER, AMBER)
    return slide, n


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), TEAL)
    for x, y, c in [(0.55, 5.5, VIOLET), (12.1, 5.45, AMBER)]:
        add_oval(slide, Inches(x), Inches(y), Inches(0.85), Inches(0.85), c)
    tb(slide, Inches(0.7), Inches(0.85), Inches(12), Inches(1.0), "🧪", size=52,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.82), Inches(12), Inches(0.9),
       "THE SCIENCE LAB MYSTERY", size=44, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.78), Inches(12), Inches(0.5),
       "\"Can You Solve the Mystery?\"", size=21,
       color=RGBColor(0x8F, 0xD3, 0xD8), align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, label) in enumerate(LAB_SCENE):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(3.48), Inches(1.65), Inches(1.35),
                  RGBColor(0x1D, 0x26, 0x31))
        tb(slide, left, Inches(3.66), Inches(1.65), Inches(0.6), emoji,
           size=22, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.31), Inches(1.45), Inches(0.4),
           label, size=10, color=RGBColor(0xAE, 0xBC, 0xC7),
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.3), Inches(5.12), Inches(6.7), Inches(1.15), TEAL)
    tb(slide, Inches(3.5), Inches(5.37), Inches(6.3), Inches(0.7),
       "Grade 5  •  90 Minutes  •  Reading & Grammar", size=19, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Science Lab")
    fade(slide)


def s02_detective():
    slide, n = new_slide("🔬 You Are the Junior Science Detective", "BRIEFING",
                         "0–7 min", "Science Lab", TEAL)
    one_task(slide, "Something unusual happened in the school science lab.",
             TEAL)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4),
              L_TEAL)
    tb(slide, Inches(0.5), Inches(2.4), Inches(5.4), Inches(1.7), "🔬",
       size=88, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.3), Inches(5.4), Inches(0.75),
       "JUNIOR SCIENCE DETECTIVE", size=23, bold=True, color=TEAL,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.2), Inches(4.8), Inches(0.7),
       "that is your job for the next 90 minutes", size=14, color=SOFT,
       align=PP_ALIGN.CENTER)
    for i, (icon, name, detail) in enumerate(MISSIONS):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.5), top + Inches(0.2), Inches(0.56),
                 Inches(0.56), WHITE)
        tb(slide, Inches(6.5), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.3), top + Inches(0.1), Inches(5.3), Inches(0.44),
           name, size=17, bold=True, color=INK)
        tb(slide, Inches(7.3), top + Inches(0.53), Inches(5.3), Inches(0.38),
           detail, size=12, color=SOFT)
    hint(slide, "Read the four missions out loud before you open the case.",
         6.45)


def s03_map():
    slide, n = new_slide("🗺️ The Case Map", "MAP", "0–7 min", "Science Lab",
                         INDIGO)
    one_task(slide, "Six locations. One language mission at each one.", INDIGO)
    for i, (emoji, place, job, when, color, light) in enumerate(STOPS):
        left = Inches(0.5 + i * 2.08)
        add_round(slide, left, Inches(1.95), Inches(1.9), Inches(3.5), light)
        tb(slide, left, Inches(2.18), Inches(1.9), Inches(0.8), emoji, size=30,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.05), Inches(3.05), Inches(1.8), Inches(0.9),
           place, size=13, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.95), Inches(1.7), Inches(0.6),
           job, size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(4.7), Inches(1.6),
                  Inches(0.42), WHITE)
        tb(slide, left + Inches(0.15), Inches(4.74), Inches(1.6), Inches(0.34),
           when, size=9, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.65), Inches(12.35), Inches(0.9),
              L_TEAL)
    tb(slide, Inches(0.8), Inches(5.9), Inches(11.7), Inches(0.44),
       "🧪  →  🔍  →  🧫  →  📋  →  📖  →  🏆    Clear all six and the case is "
       "solved.", size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)


def s04_warmup():
    slide, n = new_slide("👀 Warm-Up — What Do You Notice?", "SPEAKING",
                         "0–7 min", "Science Lab", VIOLET)
    one_task(slide, "No reading yet. Just look at the lab and talk.", VIOLET)
    for i, (emoji, label) in enumerate(LAB_SCENE):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.85 + row * 1.52)
        add_round(slide, left, top, Inches(3.9), Inches(1.34), L_VIOLET)
        tb(slide, left + Inches(0.2), top + Inches(0.26), Inches(1.1),
           Inches(0.8), emoji, size=30, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.5), top + Inches(0.38), Inches(2.1),
                  Inches(0.58), WHITE)
        tb(slide, left + Inches(1.5), top + Inches(0.49), Inches(2.1),
           Inches(0.42), label, size=15, bold=True, color=VIOLET,
           align=PP_ALIGN.CENTER)
    for i, (icon, question) in enumerate(WARM_QS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(5.0 + row * 0.56)
        tb(slide, left + Inches(0.1), top, Inches(0.4), Inches(0.44), icon,
           size=13, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.65), top, Inches(5.2), Inches(0.44),
           question, size=15, bold=True, color=INK)
    for i, starter in enumerate(WARM_STARTERS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(6.2), Inches(3.9), Inches(0.52), L_AMBER)
        tb(slide, left, Inches(6.29), Inches(3.9), Inches(0.38), starter,
           size=17, bold=True, color=INK, align=PP_ALIGN.CENTER)


def vocab_slide(index):
    word, emoji, meaning, sentence, ask, color, light = VOCAB_CARDS[index]
    slide, n = new_slide(f"📚 Science Word — {word}", "VOCABULARY", "7–17 min",
                         "Science Lab", color)
    one_task(slide, "Picture, then the word, then the meaning, then a sentence.",
             color)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(4.2), Inches(4.4), light)
    tb(slide, Inches(0.5), Inches(2.3), Inches(4.2), Inches(1.4), emoji,
       size=72, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.85), Inches(4.0), Inches(3.5), Inches(1.0), color)
    tb(slide, Inches(0.85), Inches(4.2), Inches(3.5), Inches(0.66), word,
       size=30, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
       font="Arial Black")
    tb(slide, Inches(0.85), Inches(5.25), Inches(3.5), Inches(0.6),
       "say it twice out loud", size=13, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(5.0), Inches(1.85), Inches(7.85), Inches(1.5),
              WHITE)
    tb(slide, Inches(5.3), Inches(2.0), Inches(7.25), Inches(0.36),
       "SIMPLE MEANING", size=11, bold=True, color=color)
    tb(slide, Inches(5.3), Inches(2.4), Inches(7.25), Inches(0.8), meaning,
       size=24, bold=True, color=INK)
    add_round(slide, Inches(5.0), Inches(3.5), Inches(7.85), Inches(1.5),
              light)
    tb(slide, Inches(5.3), Inches(3.65), Inches(7.25), Inches(0.36),
       "EXAMPLE SENTENCE", size=11, bold=True, color=color)
    tb(slide, Inches(5.3), Inches(4.05), Inches(7.25), Inches(0.8), sentence,
       size=21, bold=True, color=INK)
    add_round(slide, Inches(5.0), Inches(5.15), Inches(7.85), Inches(1.1),
              WHITE)
    tb(slide, Inches(5.3), Inches(5.28), Inches(7.25), Inches(0.36),
       "⭐ YOUR TURN", size=11, bold=True, color=VIOLET)
    tb(slide, Inches(5.3), Inches(5.65), Inches(7.25), Inches(0.5), ask,
       size=17, bold=True, color=INK)
    hint(slide, f"If she stalls, reread the example sentence and stress "
                f"\"{word.lower()}\".", 6.45)


def s05_observe():
    vocab_slide(0)


def s06_discover():
    vocab_slide(1)


def s07_evidence():
    vocab_slide(2)


def s08_predict():
    vocab_slide(3)


def s09_result():
    vocab_slide(4)


def s10_word_match():
    slide, n = new_slide("🖼️ Word and Picture Match", "PRACTICE", "7–17 min",
                         "Science Lab", TEAL)
    one_task(slide, "Read the word, then point to the picture that fits.", TEAL)
    for i, (word, options) in enumerate(WORD_MATCH):
        top = Inches(1.88 + i * 0.92)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.84), WHITE)
        add_oval(slide, Inches(0.75), top + Inches(0.17), Inches(0.5),
                 Inches(0.5), TEAL)
        tb(slide, Inches(0.75), top + Inches(0.23), Inches(0.5), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(1.45), top + Inches(0.14), Inches(2.8),
                  Inches(0.56), TEAL)
        tb(slide, Inches(1.45), top + Inches(0.22), Inches(2.8), Inches(0.42),
           word, size=19, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
           font="Arial Black")
        for j, emoji in enumerate(options):
            left = Inches(4.6 + j * 2.75)
            add_round(slide, left, top + Inches(0.1), Inches(2.5),
                      Inches(0.64), L_TEAL)
            tb(slide, left, top + Inches(0.14), Inches(2.5), Inches(0.52),
               emoji, size=22, align=PP_ALIGN.CENTER)
    hint(slide, "Cover one wrong picture with your hand if three is too many.",
         6.55)


def s11_word_or_not():
    slide, n = new_slide("🧫 Game: Science Word or Not?", "GAME", "7–17 min",
                         "Science Lab", VIOLET)
    one_task(slide, "Which of these belong in a science mystery?", VIOLET)
    for i, (word, _is_science) in enumerate(WORD_OR_NOT):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.95 + row * 2.2)
        add_round(slide, left, top, Inches(3.9), Inches(2.0), WHITE)
        tb(slide, left, top + Inches(0.3), Inches(3.9), Inches(0.7), word,
           size=30, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Georgia")
        add_round(slide, left + Inches(0.35), top + Inches(1.15), Inches(1.5),
                  Inches(0.6), L_TEAL)
        tb(slide, left + Inches(0.35), top + Inches(1.27), Inches(1.5),
           Inches(0.42), "🧪 SCIENCE", size=12, bold=True, color=TEAL,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(2.05), top + Inches(1.15), Inches(1.5),
                  Inches(0.6), L_GREY)
        tb(slide, left + Inches(2.05), top + Inches(1.27), Inches(1.5),
           Inches(0.42), "🚫 NOT", size=12, bold=True, color=SOFT,
           align=PP_ALIGN.CENTER)
    hint(slide, "Ask why. \"It is a fruit\" is a better answer than just \"no\".",
         6.45)


def s12_say_it():
    slide, n = new_slide("🎤 Say It! — Use the New Words", "SPEAKING",
                         "7–17 min", "Science Lab", ROSE)
    one_task(slide, "Pick a word. Build a whole sentence with it.", ROSE)
    for i, (emoji, word) in enumerate(SAY_IT):
        left = Inches(0.6 + i * 4.15)
        add_round(slide, left, Inches(1.95), Inches(3.9), Inches(2.4), L_ROSE)
        tb(slide, left, Inches(2.2), Inches(3.9), Inches(0.95), emoji, size=40,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.7), Inches(3.3), Inches(2.5),
                  Inches(0.7), WHITE)
        tb(slide, left + Inches(0.7), Inches(3.42), Inches(2.5), Inches(0.48),
           word, size=22, bold=True, color=ROSE, align=PP_ALIGN.CENTER)
    for i, starter in enumerate(SAY_STARTERS):
        top = Inches(4.65 + i * 0.6)
        add_round(slide, Inches(0.6), top, Inches(12.15), Inches(0.52), L_AMBER)
        tb(slide, Inches(1.0), top + Inches(0.07), Inches(11.4), Inches(0.4),
           starter, size=19, bold=True, color=INK)
    hint(slide, "One good sentence each is plenty. Say it back correctly, then "
                "she repeats it.", 6.45)


def s13_vocab_bank():
    slide, n = new_slide("📚 Your Word Bank for Today", "VOCABULARY",
                         "7–17 min", "Science Lab", INDIGO)
    one_task(slide, "Every word in this lesson comes from this bank.", INDIGO)
    for i, (word, emoji) in enumerate(VOCAB_BANK):
        col, row = i % 8, i // 8
        left = Inches(0.5 + col * 1.56)
        top = Inches(2.15 + row * 1.75)
        add_round(slide, left, top, Inches(1.42), Inches(1.55), L_INDIGO)
        tb(slide, left, top + Inches(0.14), Inches(1.42), Inches(0.6), emoji,
           size=22, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.08), top + Inches(0.84),
                  Inches(1.26), Inches(0.56), WHITE)
        tb(slide, left + Inches(0.03), top + Inches(0.94), Inches(1.36),
           Inches(0.4), word, size=11, bold=True, color=INDIGO,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.7), Inches(12.35), Inches(0.58),
              L_AMBER)
    tb(slide, Inches(0.8), Inches(5.82), Inches(11.7), Inches(0.4),
       "For any word she does not know:  picture  →  word  →  simple meaning  "
       "→  example sentence.", size=14, bold=True, color=AMBER,
       align=PP_ALIGN.CENTER)
    hint(slide, "Come back to this slide any time a word blocks her.", 6.45)


def s14_subject_verb():
    slide, n = new_slide("✏️ Subject + Verb — The Two Key Words", "EXPLAIN",
                         "17–27 min", "Observation Desk", INDIGO)
    one_task(slide, "Every sentence has someone doing something.", INDIGO)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(1.5),
              WHITE)
    tb(slide, Inches(0.5), Inches(2.25), Inches(12.35), Inches(0.8),
       "The scientist works in the lab.", size=34, bold=True, color=INK,
       align=PP_ALIGN.CENTER, font="Georgia")
    pieces = [("SUBJECT", "scientist", "Who or what is doing it?", INDIGO,
               L_INDIGO),
              ("VERB", "works", "What are they doing?", TEAL, L_TEAL)]
    for i, (name, word, detail, color, light) in enumerate(pieces):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(3.55), Inches(5.95), Inches(2.3), light)
        add_round(slide, left, Inches(3.55), Inches(5.95), Inches(0.5), color)
        tb(slide, left, Inches(3.63), Inches(5.95), Inches(0.36), name,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(4.15), Inches(5.95), Inches(0.85), word,
           size=40, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        add_round(slide, left + Inches(0.8), Inches(5.1), Inches(4.35),
                  Inches(0.58), WHITE)
        tb(slide, left + Inches(0.8), Inches(5.21), Inches(4.35), Inches(0.4),
           detail, size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "The subject usually sits right before the verb. Point at both.",
         6.45)


def sva_slide(index):
    (s_subj, s_sent, s_verb, s_note, p_subj, p_sent, p_verb, p_note,
     emoji) = SVA_PAIRS[index]
    slide, n = new_slide(f"🔍 One or Many? — {s_subj} / {p_subj}", "EXPLAIN",
                         "17–27 min", "Observation Desk", INDIGO)
    one_task(slide, "Same idea, two subjects. Watch what happens to the verb.",
             INDIGO)
    cards = [("ONE", s_subj, s_sent, s_verb, s_note, TEAL, L_TEAL),
             ("MORE THAN ONE", p_subj, p_sent, p_verb, p_note, VIOLET,
              L_VIOLET)]
    for i, (tag, subj, sent, verb, note, color, light) in enumerate(cards):
        left = Inches(0.5 + i * 6.25)
        add_round(slide, left, Inches(1.85), Inches(5.95), Inches(4.4), light)
        add_round(slide, left, Inches(1.85), Inches(5.95), Inches(0.52), color)
        tb(slide, left, Inches(1.94), Inches(5.95), Inches(0.36), tag, size=13,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(5.95), Inches(0.8), emoji, size=34,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(3.35), Inches(5.35), Inches(0.9),
           sent, size=25, bold=True, color=INK, align=PP_ALIGN.CENTER,
           font="Georgia")
        add_round(slide, left + Inches(0.5), Inches(4.35), Inches(2.3),
                  Inches(0.82), WHITE)
        tb(slide, left + Inches(0.5), Inches(4.43), Inches(2.3), Inches(0.28),
           "SUBJECT", size=10, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.5), Inches(4.7), Inches(2.3), Inches(0.42),
           subj, size=19, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(3.15), Inches(4.35), Inches(2.3),
                  Inches(0.82), WHITE)
        tb(slide, left + Inches(3.15), Inches(4.43), Inches(2.3), Inches(0.28),
           "VERB", size=10, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(3.15), Inches(4.7), Inches(2.3), Inches(0.42),
           verb, size=19, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.5), Inches(5.35), Inches(4.95), Inches(0.5),
           note, size=14, color=DARK, align=PP_ALIGN.CENTER, italic=True)
    hint(slide, "Read both sentences out loud back to back. The ear hears the "
                "difference.", 6.45)


def s15_sva_scientist():
    sva_slide(0)


def s16_sva_experiment():
    sva_slide(1)


def s17_sva_rule():
    slide, n = new_slide("🧠 The Rule in Four Steps", "RULE", "17–27 min",
                         "Observation Desk", TEAL)
    one_task(slide, "Use these four steps on every sentence today.", TEAL)
    for i, (num, label, detail, example, color) in enumerate(SVA_RULE):
        left = Inches(0.5 + i * 3.11)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(3.8), L_GREY)
        add_oval(slide, left + Inches(1.15), Inches(2.2), Inches(0.6),
                 Inches(0.6), color)
        tb(slide, left + Inches(1.15), Inches(2.3), Inches(0.6), Inches(0.42),
           num, size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.0), Inches(2.7), Inches(0.6),
           label, size=17, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(3.7), Inches(2.4), Inches(1.0),
           detail, size=13, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(4.8), Inches(2.4),
                  Inches(0.7), WHITE)
        tb(slide, left + Inches(0.25), Inches(4.95), Inches(2.4), Inches(0.45),
           example, size=13, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.9), Inches(12.35), Inches(0.44),
              L_TEAL)
    tb(slide, Inches(0.8), Inches(5.97), Inches(11.7), Inches(0.34),
       "Short version:  ONE subject takes the verb WITH -s.  MANY subjects "
       "take the verb WITHOUT -s.", size=13, bold=True, color=TEAL)


def s18_verb_choice_a():
    slide, n = new_slide("🔍 Which Verb Is Correct?", "PRACTICE", "17–27 min",
                         "Observation Desk", INDIGO)
    one_task(slide, "Find the subject first. Then choose the verb.", INDIGO)
    choice_rows(slide, VERB_CHOICE_A, 1, INDIGO)
    hint(slide, "Cover the choices. Say the sentence both ways out loud first.",
         6.42)


def s19_verb_choice_b():
    slide, n = new_slide("🔍 Which Verb Is Correct? — Rounds 4 to 6", "PRACTICE",
                         "17–27 min", "Observation Desk", INDIGO)
    one_task(slide, "Three more. Same four steps every time.", INDIGO)
    choice_rows(slide, VERB_CHOICE_B, 4, INDIGO)
    hint(slide, "Ask her to say the subject out loud before she answers.", 6.42)


def s20_fix_a():
    slide, n = new_slide("🛠️ Grammar Detective — Find the Mistake", "GAME",
                         "17–27 min", "Observation Desk", CRIMSON)
    one_task(slide, "Each sentence has one mistake. Find it and fix it.",
             CRIMSON)
    fix_rows(slide, FIX_A, 1, CRIMSON)
    hint(slide, "The mistake is always the verb. Check it against the subject.",
         6.42)


def s21_fix_b():
    slide, n = new_slide("🛠️ Grammar Detective — Rounds 3 and 4", "GAME",
                         "17–27 min", "Observation Desk", CRIMSON)
    one_task(slide, "Two more broken sentences. You know the rule now.",
             CRIMSON)
    fix_rows(slide, FIX_B, 3, CRIMSON)
    hint(slide, "Have her read her fixed version out loud. It should sound "
                "right.", 6.42)


def s22_observation_done():
    badge_slide("🔍 Observation Desk Cleared", "17–27 min", "Observation Desk",
                "🔍", "BADGE 1 EARNED",
                "You can match any verb to its subject.",
                [("✏️", "Subject and verb in every sentence."),
                 ("🔢", "Singular takes -s · plural does not."),
                 ("✅", "Six correct verb choices."),
                 ("🛠️", "Four broken sentences repaired.")],
                "Experiment Station — nouns, verbs, adjectives and adverbs.",
                INDIGO, L_INDIGO)


def s23_pos_legend():
    slide, n = new_slide("🏷️ Four Labels for Four Jobs", "EXPLAIN", "27–40 min",
                         "Experiment Station", VIOLET)
    one_task(slide, "Every word in a sentence has a job. These are the four.",
             VIOLET)
    for i, (name, detail, examples, color, light) in enumerate(POS_LEGEND):
        left = Inches(0.5 + i * 3.11)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(4.0), light)
        add_round(slide, left, Inches(1.95), Inches(2.9), Inches(0.62), color)
        tb(slide, left, Inches(2.07), Inches(2.9), Inches(0.42), name, size=16,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(2.85), Inches(2.4), Inches(1.2),
           detail, size=14, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(4.3), Inches(2.4),
                  Inches(1.4), WHITE)
        tb(slide, left + Inches(0.3), Inches(4.45), Inches(2.3), Inches(0.3),
           "EXAMPLES", size=9, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(4.8), Inches(2.3), Inches(0.8),
           examples, size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.1), Inches(12.35), Inches(0.44),
              L_AMBER)
    tb(slide, Inches(0.8), Inches(6.17), Inches(11.7), Inches(0.34),
       "Quick test:  adjectives describe a NOUN, adverbs describe a VERB — and "
       "most adverbs end in -ly.", size=13, bold=True, color=AMBER)


def s24_pos_a():
    sentence, asks = POS_TASKS[0]
    pos_task_slide(sentence, asks, "🧫 Label the Sentence — Noun and Verb",
                   "27–40 min", "Experiment Station", VIOLET,
                   "Write the noun and the verb in the jars below.")


def s25_pos_b():
    sentence, asks = POS_TASKS[1]
    pos_task_slide(sentence, asks, "🧫 Label the Sentence — Add the Adjective",
                   "27–40 min", "Experiment Station", VIOLET,
                   "Three jobs this time. The adjective describes a noun.")


def s26_pos_c():
    sentence, asks = POS_TASKS[2]
    pos_task_slide(sentence, asks, "🧫 Label the Sentence — Add the Adverb",
                   "27–40 min", "Experiment Station", VIOLET,
                   "The adverb tells you HOW the action was done.")


def s27_grammar_sort():
    slide, n = new_slide("🧪 Game: Grammar Sort", "GAME", "27–40 min",
                         "Experiment Station", TEAL)
    one_task(slide, "Drop each word into the right beaker.", TEAL)
    for i, (name, color, light) in enumerate(SORT_BEAKERS):
        left = Inches(0.5 + i * 3.11)
        add_round(slide, left, Inches(1.9), Inches(2.9), Inches(2.5), light)
        tb(slide, left, Inches(2.05), Inches(2.9), Inches(0.72), "🧪", size=28,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(2.85), Inches(2.3),
                  Inches(0.55), color)
        tb(slide, left + Inches(0.3), Inches(2.95), Inches(2.3), Inches(0.4),
           name, size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(3.55), Inches(2.3), Inches(0.7),
           "____________\n____________", size=13, color=SOFT,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.6), Inches(12.35), Inches(0.4),
       "WORDS TO SORT", size=12, bold=True, color=TEAL)
    for i, word in enumerate(SORT_WORDS):
        left = Inches(0.5 + i * 2.08)
        add_round(slide, left, Inches(5.05), Inches(1.9), Inches(0.78), WHITE)
        tb(slide, left, Inches(5.2), Inches(1.9), Inches(0.5), word, size=17,
           bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Try the word in a sentence. \"A careful ____\" proves it is an "
                "adjective.", 6.1)


def s28_fast_lab():
    slide, n = new_slide("⚡ Fast Lab Challenge", "GAME", "27–40 min",
                         "Experiment Station", AMBER)
    one_task(slide, "Three quick rounds. Read, then pick one word.", AMBER)
    for i, (task, sentence, options, color, light) in enumerate(FAST_LAB):
        top = Inches(1.9 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), light)
        add_round(slide, Inches(0.8), top + Inches(0.42), Inches(2.2),
                  Inches(0.52), color)
        tb(slide, Inches(0.8), top + Inches(0.51), Inches(2.2), Inches(0.36),
           task, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(3.2), top + Inches(0.2), Inches(9.3), Inches(0.5),
           sentence, size=19, bold=True, color=INK)
        for j, opt in enumerate(options):
            left = Inches(3.2 + j * 3.15)
            add_round(slide, left, top + Inches(0.76), Inches(2.95),
                      Inches(0.46), WHITE)
            tb(slide, left, top + Inches(0.83), Inches(2.95), Inches(0.34),
               opt, size=15, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Say the four labels out loud before round 1 as a reminder.",
         6.42)


def s29_experiment_done():
    badge_slide("🧫 Experiment Station Cleared", "27–40 min",
                "Experiment Station", "🧫", "BADGE 2 EARNED",
                "You can name the job of any word in a sentence.",
                [("🏷️", "Noun · verb · adjective · adverb."),
                 ("🧫", "Three sentences fully labelled."),
                 ("🧪", "Six words sorted into the right beaker."),
                 ("⚡", "Three fast rounds, no hints needed at the end.")],
                "Brain break, then the Research Table and verb tenses.",
                VIOLET, L_VIOLET)


def s30_break_intro():
    slide, n = new_slide("🧠 Brain Break — Scientist Says", "BREAK",
                         "35–40 min", "Experiment Station", ROSE)
    one_task(slide, "Stay by your chair. Only move if I say \"Scientist says\".",
             ROSE)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4),
              L_ROSE)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.4), Inches(1.8), "🥼",
       size=90, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.4), Inches(0.7), "5 MINUTES",
       size=30, bold=True, color=ROSE, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.8), Inches(0.6),
       "then straight into verb tenses", size=14, color=SOFT,
       align=PP_ALIGN.CENTER)
    rules = [("🪑", "Stay beside your chair the whole time."),
             ("🧊", "\"Freeze!\" means stop completely."),
             ("✏️", "Half the calls are grammar calls."),
             ("🔁", "The last two rounds, you call and I move.")]
    for i, (icon, line) in enumerate(rules):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.55), top + Inches(0.2), Inches(0.56),
                 Inches(0.56), WHITE)
        tb(slide, Inches(6.55), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.35), top + Inches(0.24), Inches(5.2), Inches(0.5),
           line, size=15, bold=True, color=INK)
    hint(slide, "Letting her give the commands is the part that resets her "
                "focus.", 6.45)


def s31_scientist_says():
    slide, n = new_slide("🥼 Scientist Says — Six Calls", "BREAK", "35–40 min",
                         "Experiment Station", ROSE)
    one_task(slide, "Listen for \"Scientist says\" — then do it.", ROSE)
    for i, (icon, line, kind) in enumerate(SCIENTIST_SAYS):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 1.5)
        color = TEAL if kind == "grammar" else ROSE
        light = L_TEAL if kind == "grammar" else L_ROSE
        add_round(slide, left, top, Inches(5.95), Inches(1.3), light)
        add_oval(slide, left + Inches(0.3), top + Inches(0.36), Inches(0.58),
                 Inches(0.58), WHITE)
        tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(0.58),
           Inches(0.44), icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.3), Inches(4.5),
           Inches(0.52), line, size=17, bold=True, color=INK)
        tb(slide, left + Inches(1.1), top + Inches(0.83), Inches(4.5),
           Inches(0.36),
           "grammar call" if kind == "grammar" else "movement call", size=11,
           color=color)
    hint(slide, "Mix the order so she has to listen rather than predict.", 6.45)


def s32_freeze_challenge():
    slide, n = new_slide("🧊 Freeze! Now Find the Word", "BREAK", "35–40 min",
                         "Experiment Station", TEAL)
    one_task(slide, "One sentence, four quick finds. Point, do not write.",
             TEAL)
    add_round(slide, Inches(0.5), Inches(1.9), Inches(12.35), Inches(1.5),
              WHITE)
    tb(slide, Inches(0.5), Inches(2.3), Inches(12.35), Inches(0.8),
       FREEZE_SENTENCE, size=28, bold=True, color=INK, align=PP_ALIGN.CENTER,
       font="Georgia")
    for i, (task, color, light) in enumerate(FREEZE_TASKS):
        left = Inches(0.5 + i * 3.11)
        add_round(slide, left, Inches(3.6), Inches(2.9), Inches(2.3), light)
        tb(slide, left, Inches(3.8), Inches(2.9), Inches(0.62), "🔎", size=22,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.25), Inches(4.5), Inches(2.4),
                  Inches(0.55), color)
        tb(slide, left + Inches(0.25), Inches(4.6), Inches(2.4), Inches(0.4),
           task, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(5.2), Inches(2.4), Inches(0.5),
           "____________", size=17, color=SOFT, align=PP_ALIGN.CENTER)
    hint(slide, "Two of these have more than one right answer. Either is fine.",
         6.15)


def s33_tenses():
    slide, n = new_slide("⏳ Past, Now, Future", "EXPLAIN", "40–50 min",
                         "Research Table", AMBER)
    one_task(slide, "Same sentence, three different times.", AMBER)
    for i, (name, when, verb, sentence, color, light) in enumerate(TENSES):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(4.0), light)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(0.62), color)
        tb(slide, left, Inches(2.02), Inches(3.9), Inches(0.42), name, size=16,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.7), Inches(3.9), Inches(0.5), when, size=15,
           color=SOFT, align=PP_ALIGN.CENTER, italic=True)
        add_round(slide, left + Inches(0.55), Inches(3.25), Inches(2.8),
                  Inches(0.9), WHITE)
        tb(slide, left + Inches(0.55), Inches(3.42), Inches(2.8), Inches(0.6),
           verb, size=26, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Arial Black")
        tb(slide, left + Inches(0.3), Inches(4.35), Inches(3.3), Inches(1.4),
           sentence, size=16, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.05), Inches(12.35), Inches(0.5),
              L_GREY)
    tb(slide, Inches(0.8), Inches(6.14), Inches(11.7), Inches(0.36),
       "PAST  ←——————  NOW  ——————→  FUTURE        -ed  ·  -s  ·  will + verb",
       size=15, bold=True, color=INK, align=PP_ALIGN.CENTER)


def s34_tense_id():
    slide, n = new_slide("⏱️ Time Machine — Which Tense Is It?", "GAME",
                         "40–50 min", "Research Table", AMBER)
    one_task(slide, "Look at the verb, then name the tense.", AMBER)
    choice_rows(slide, TENSE_ID, 1, AMBER)
    hint(slide, "No time word in the sentence? Then the verb ending is the "
                "only clue.", 6.42)


def shift_slide(items, title, sub, target_label, timing, color, light,
                hint_text):
    """Rewrite a present-tense sentence into another tense."""
    slide, n = new_slide(title, "PRACTICE", timing, "Research Table", color)
    one_task(slide, sub, color)
    for i, (sentence, _target) in enumerate(items):
        top = Inches(1.95 + i * 2.15)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.9), light)
        add_round(slide, Inches(0.8), top + Inches(0.22), Inches(1.7),
                  Inches(0.44), WHITE)
        tb(slide, Inches(0.8), top + Inches(0.27), Inches(1.7), Inches(0.34),
           "PRESENT", size=10, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.7), top + Inches(0.2), Inches(9.8), Inches(0.5),
           sentence, size=24, bold=True, color=INK)
        add_round(slide, Inches(0.8), top + Inches(0.9), Inches(1.7),
                  Inches(0.44), color)
        tb(slide, Inches(0.8), top + Inches(0.95), Inches(1.7), Inches(0.34),
           target_label, size=10, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        add_round(slide, Inches(2.7), top + Inches(0.85), Inches(9.8),
                  Inches(0.78), WHITE)
        tb(slide, Inches(2.95), top + Inches(1.05), Inches(9.3), Inches(0.42),
           "______________________________________________________________",
           size=15, color=SOFT)
    hint(slide, hint_text, 6.42)
    return slide, n


def s35_to_past():
    shift_slide(TO_PAST, "⏪ Change It to the Past",
                "Only the verb changes. Rewrite the whole sentence.", "PAST",
                "40–50 min", VIOLET, L_VIOLET,
                "Most past-tense verbs add -ed. Say it out loud to check.")


def s36_to_future():
    shift_slide(TO_FUTURE, "⏩ Change It to the Future",
                "Add one helper word in front of the verb.", "FUTURE",
                "40–50 min", AMBER, L_AMBER,
                "Future needs \"will\" plus the plain verb — not \"will "
                "measures\".")


def s37_tense_choice():
    slide, n = new_slide("⚡ Quick Challenge — Pick the Right Sentence", "GAME",
                         "40–50 min", "Research Table", TEAL)
    one_task(slide, "The time word tells you which verb to use.", TEAL)
    choice_rows(slide, TENSE_CHOICE, 1, TEAL, top_start=2.1, gap=1.9,
                height=1.6)
    add_round(slide, Inches(0.5), Inches(5.6), Inches(12.35), Inches(0.62),
              L_VIOLET)
    tb(slide, Inches(0.8), Inches(5.73), Inches(11.7), Inches(0.42),
       "⭐ EXPERT CHALLENGE — Say the same sentence in all three tenses, "
       "without looking back.", size=14, bold=True, color=VIOLET)
    hint(slide, "Underline the time word first. It decides the answer.", 6.42)


def s38_research_done():
    badge_slide("📋 Research Table Cleared", "40–50 min", "Research Table",
                "📋", "BADGE 3 EARNED",
                "You can move a sentence through time.",
                [("⏳", "Past · present · future."),
                 ("⏱️", "Three tenses identified from the verb alone."),
                 ("⏪", "Two sentences rewritten in the past."),
                 ("⏩", "Two sentences rewritten in the future.")],
                "Mystery File — the passage, and the clues inside it.",
                AMBER, L_AMBER)


def s39_story_intro():
    slide, n = new_slide("📖 Open the Mystery File", "READING", "50–60 min",
                         "Mystery File", CRIMSON)
    one_task(slide, "Before we read: what do you think happened?", CRIMSON)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4),
              L_CRIMSON)
    tb(slide, Inches(0.5), Inches(2.25), Inches(5.4), Inches(1.5), "🧴",
       size=72, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.6), Inches(3.9), Inches(5.2), Inches(1.1),
       "THE MYSTERY OF THE MISSING SAMPLE", size=24, bold=True, color=CRIMSON,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.15), Inches(4.8), Inches(0.9),
       "five parts  ·  I read first, then we read together", size=14,
       color=SOFT, align=PP_ALIGN.CENTER)
    tb(slide, Inches(6.25), Inches(1.9), Inches(6.6), Inches(0.5),
       "🤔 MAKE A PREDICTION", size=15, bold=True, color=CRIMSON)
    tb(slide, Inches(6.25), Inches(2.4), Inches(6.6), Inches(0.4),
       "There is no wrong answer here. A prediction is a smart guess.",
       size=13, color=SOFT)
    for i, guess in enumerate(STORY_PREDICT):
        top = Inches(2.95 + i * 0.95)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(0.8), WHITE)
        add_round(slide, Inches(6.5), top + Inches(0.18), Inches(0.45),
                  Inches(0.45), L_CRIMSON)
        tb(slide, Inches(6.5), top + Inches(0.22), Inches(0.45), Inches(0.36),
           chr(65 + i), size=12, bold=True, color=CRIMSON,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.2), top + Inches(0.2), Inches(5.4), Inches(0.44),
           guess, size=16, bold=True, color=INK)
    add_round(slide, Inches(6.25), Inches(5.8), Inches(6.6), Inches(0.44),
              L_AMBER)
    tb(slide, Inches(6.5), Inches(5.87), Inches(6.1), Inches(0.34),
       "Write your prediction down. We will check it at the end.", size=12,
       bold=True, color=AMBER)
    hint(slide, "Predicting first gives her a reason to read carefully.", 6.45)


def s40_story_1():
    story_slide(STORY[0], "50–60 min")


def s41_story_2():
    story_slide(STORY[1], "50–60 min")


def s42_story_3():
    story_slide(STORY[2], "50–60 min")


def s43_story_4():
    story_slide(STORY[3], "50–60 min")


def s44_story_5():
    story_slide(STORY[4], "50–60 min")


def s45_main_idea():
    slide, n = new_slide("💡 Main Idea and Summary", "COMPREHENSION",
                         "60–70 min", "Mystery File", AMBER)
    one_task(slide, "The main idea covers the WHOLE passage, not one detail.",
             AMBER)
    tb(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(0.4),
       "WHICH ONE IS THE MAIN IDEA?", size=13, bold=True, color=AMBER)
    for i, option in enumerate(MAIN_IDEA):
        top = Inches(2.25 + i * 0.8)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.7), WHITE)
        add_round(slide, Inches(0.8), top + Inches(0.13), Inches(0.5),
                  Inches(0.44), L_AMBER)
        tb(slide, Inches(0.8), top + Inches(0.18), Inches(0.5), Inches(0.36),
           chr(65 + i), size=13, bold=True, color=AMBER,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.14), Inches(10.6), Inches(0.44),
           option, size=19, bold=True, color=INK)
    tb(slide, Inches(0.5), Inches(4.7), Inches(12.35), Inches(0.4),
       "NOW SUMMARISE IT IN THREE LINES", size=13, bold=True, color=TEAL)
    for i, frame in enumerate(SUMMARY_FRAME):
        top = Inches(5.1 + i * 0.48)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.42), L_TEAL)
        tb(slide, Inches(0.85), top + Inches(0.05), Inches(11.6), Inches(0.32),
           frame, size=15, bold=True, color=INK)
    hint(slide, "Too broad or too narrow? Ask: does this cover all five parts?",
         6.6)


def s46_clue_q_a():
    slide, n = new_slide("🔎 Clue Detective — Q1 and Q2", "COMPREHENSION",
                         "60–70 min", "Mystery File", CRIMSON)
    one_task(slide, "Answer from the passage, not from memory.", CRIMSON)
    question_rows(slide, CLUE_Q_A, 1)
    hint(slide, "Turn back to the part named under the question. She finds it.",
         6.42)


def s47_clue_q_b():
    slide, n = new_slide("🔎 Clue Detective — Q3 and Q4", "COMPREHENSION",
                         "60–70 min", "Mystery File", CRIMSON)
    one_task(slide, "Two more. Point to the line that proves it.", CRIMSON)
    question_rows(slide, CLUE_Q_B, 3)
    hint(slide, "Cover one wrong option. Two choices is still real thinking.",
         6.42)


def s48_sequence():
    slide, n = new_slide("🔢 Sequence the Events", "COMPREHENSION", "60–70 min",
                         "Mystery File", INDIGO)
    one_task(slide, "These four events are out of order. Number them 1 to 4.",
             INDIGO)
    for i, (letter, event, _order) in enumerate(SEQUENCE):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.95 + row * 1.5)
        add_round(slide, left, top, Inches(5.95), Inches(1.3), L_INDIGO)
        add_oval(slide, left + Inches(0.3), top + Inches(0.36), Inches(0.58),
                 Inches(0.58), INDIGO)
        tb(slide, left + Inches(0.3), top + Inches(0.42), Inches(0.58),
           Inches(0.44), letter, size=15, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.34), Inches(3.6),
           Inches(0.6), event, size=18, bold=True, color=INK)
        add_round(slide, left + Inches(4.85), top + Inches(0.34), Inches(0.8),
                  Inches(0.62), WHITE)
        tb(slide, left + Inches(4.85), top + Inches(0.44), Inches(0.8),
           Inches(0.44), "___", size=18, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.0), Inches(12.35), Inches(1.2),
              WHITE)
    tb(slide, Inches(0.85), Inches(5.15), Inches(11.6), Inches(0.36),
       "WRITE THE ORDER HERE", size=12, bold=True, color=INDIGO)
    tb(slide, Inches(0.85), Inches(5.55), Inches(11.6), Inches(0.5),
       "____  →  ____  →  ____  →  ____", size=26, bold=True, color=SOFT)
    hint(slide, "Ask which one had to happen first. Nothing else works without "
                "it.", 6.45)


def s49_evidence():
    slide, n = new_slide("🔎 Show Me the Clue!", "COMPREHENSION", "60–70 min",
                         "Mystery File", VIOLET)
    one_task(slide, "Do not tell me the answer. Point to the sentence.", VIOLET)
    for i, (question, part, color, light) in enumerate(EVIDENCE_QS):
        top = Inches(1.9 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), light)
        add_oval(slide, Inches(0.8), top + Inches(0.42), Inches(0.52),
                 Inches(0.52), color)
        tb(slide, Inches(0.8), top + Inches(0.48), Inches(0.52), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.16), Inches(4.7), Inches(0.74),
           question, size=17, bold=True, color=INK)
        add_round(slide, Inches(1.55), top + Inches(0.88), Inches(1.8),
                  Inches(0.4), WHITE)
        tb(slide, Inches(1.55), top + Inches(0.93), Inches(1.8), Inches(0.32),
           f"look in {part}", size=11, bold=True, color=color,
           align=PP_ALIGN.CENTER)
        add_round(slide, Inches(6.4), top + Inches(0.3), Inches(6.15),
                  Inches(0.74), WHITE)
        tb(slide, Inches(6.65), top + Inches(0.4), Inches(5.7), Inches(0.34),
           "🔎  Go back, point to the line, then read it out loud.", size=13,
           bold=True, color=color)
        tb(slide, Inches(6.65), top + Inches(0.72), Inches(5.7), Inches(0.3),
           "☐ found it", size=11, color=SOFT)
    add_round(slide, Inches(0.5), Inches(6.46), Inches(12.35), Inches(0.44),
              L_TEAL)
    tb(slide, Inches(0.8), Inches(6.53), Inches(11.7), Inches(0.34),
       "Evidence means a line from the text. \"I just think so\" is not "
       "evidence — that is the whole point of this slide.", size=12, bold=True,
       color=TEAL)


def s50_true_false():
    slide, n = new_slide("✅ True or False?", "COMPREHENSION", "60–70 min",
                         "Mystery File", TEAL)
    one_task(slide, "Decide, then say which part told you.", TEAL)
    for i, (statement, _answer) in enumerate(TRUE_FALSE):
        top = Inches(1.95 + i * 1.45)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.26), WHITE)
        add_oval(slide, Inches(0.8), top + Inches(0.38), Inches(0.5),
                 Inches(0.5), TEAL)
        tb(slide, Inches(0.8), top + Inches(0.44), Inches(0.5), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.36), Inches(6.9), Inches(0.56),
           statement, size=19, bold=True, color=INK)
        for j, (label, color, light) in enumerate(
                [("TRUE", TEAL, L_TEAL), ("FALSE", CRIMSON, L_CRIMSON)]):
            left = Inches(8.7 + j * 2.0)
            add_round(slide, left, top + Inches(0.34), Inches(1.8),
                      Inches(0.6), light)
            tb(slide, left, top + Inches(0.45), Inches(1.8), Inches(0.42),
               label, size=16, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "\"Immediately\" is the word that decides number 1. Reread it.",
         6.45)


def s51_inference():
    slide, n = new_slide("🧠 Make an Inference", "COMPREHENSION", "60–70 min",
                         "Mystery File", VIOLET)
    one_task(slide, "An inference is a conclusion you build from clues.",
             VIOLET)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(0.95),
              WHITE)
    tb(slide, Inches(0.5), Inches(2.05), Inches(12.35), Inches(0.6),
       "What can we infer about these students?", size=26, bold=True,
       color=INK, align=PP_ALIGN.CENTER, font="Georgia")
    for i, option in enumerate(INFERENCE_OPTIONS):
        top = Inches(3.0 + i * 0.92)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.8),
                  L_VIOLET)
        add_round(slide, Inches(0.8), top + Inches(0.18), Inches(0.5),
                  Inches(0.44), WHITE)
        tb(slide, Inches(0.8), top + Inches(0.23), Inches(0.5), Inches(0.36),
           chr(65 + i), size=13, bold=True, color=VIOLET,
           align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.55), top + Inches(0.18), Inches(10.6),
           Inches(0.46), option, size=20, bold=True, color=INK)
    add_round(slide, Inches(0.5), Inches(5.85), Inches(12.35), Inches(0.85),
              L_AMBER)
    tb(slide, Inches(0.85), Inches(5.97), Inches(11.6), Inches(0.36),
       "⭐ EXPERT CHALLENGE", size=12, bold=True, color=AMBER)
    tb(slide, Inches(0.85), Inches(6.3), Inches(11.6), Inches(0.36),
       "The passage never says \"the students were careful\". So which clues "
       "made you choose your answer?", size=15, bold=True, color=INK)


def s52_mystery_done():
    badge_slide("📖 Mystery File Cleared", "60–70 min", "Mystery File", "📖",
                "BADGE 4 EARNED",
                "You read the case and proved your answers.",
                [("📖", "A 314-word passage, all five parts."),
                 ("💡", "Main idea, details and a three-line summary."),
                 ("🔎", "Four questions, plus the evidence for each."),
                 ("🧠", "One real inference, built from clues.")],
                "Detective HQ — edit the report, then present the case.",
                CRIMSON, L_CRIMSON)


def s53_editor_a():
    slide, n = new_slide("🧩 Grammar Lab Editor — Fix the Report", "GAME",
                         "70–78 min", "Detective HQ", CRIMSON)
    one_task(slide, "These lines came from the case report. Repair them.",
             CRIMSON)
    fix_rows(slide, EDITOR_A, 1, CRIMSON)
    hint(slide, "Same rule as before: find the subject, then check the verb.",
         6.42)


def s54_editor_pos():
    slide, n = new_slide("🏷️ Label the Report Sentences", "GRAMMAR",
                         "70–78 min", "Detective HQ", VIOLET)
    one_task(slide, "Two story sentences. Write one word in each box.", VIOLET)
    for i, (sentence, asks) in enumerate(EDITOR_POS):
        top = Inches(1.9 + i * 2.35)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(2.15),
                  L_GREY)
        tb(slide, Inches(0.8), top + Inches(0.18), Inches(11.7), Inches(0.6),
           sentence, size=24, bold=True, color=INK, font="Georgia")
        for j, (name, color, light) in enumerate(asks):
            left = Inches(0.8 + j * 3.95)
            add_round(slide, left, top + Inches(0.92), Inches(3.7),
                      Inches(1.06), light)
            add_round(slide, left, top + Inches(0.92), Inches(3.7),
                      Inches(0.42), color)
            tb(slide, left, top + Inches(0.98), Inches(3.7), Inches(0.32),
               name, size=12, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
            add_round(slide, left + Inches(0.5), top + Inches(1.44),
                      Inches(2.7), Inches(0.46), WHITE)
            tb(slide, left + Inches(0.5), top + Inches(1.52), Inches(2.7),
               Inches(0.36), "____________", size=16, color=SOFT,
               align=PP_ALIGN.CENTER)
    pos_legend(slide, 6.42)


def s55_editor_challenge():
    slide, n = new_slide("✒️ Editor's Challenge", "GAME", "70–78 min",
                         "Detective HQ", AMBER)
    one_task(slide, "Three things are wrong here. Find all three.", AMBER)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(1.2),
              L_CRIMSON)
    tb(slide, Inches(0.8), Inches(1.98), Inches(2.0), Inches(0.34),
       "❌ AS WRITTEN", size=11, bold=True, color=CRIMSON)
    tb(slide, Inches(0.8), Inches(2.32), Inches(11.7), Inches(0.6),
       "the students found sample c in the refrigerator", size=26, bold=True,
       color=INK, font="Georgia")
    for i, (icon, name, detail) in enumerate(EDITOR_CHECKS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(3.25), Inches(3.9), Inches(1.7), L_AMBER)
        tb(slide, left, Inches(3.42), Inches(3.9), Inches(0.6), icon, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(4.02), Inches(3.9), Inches(0.42), name,
           size=16, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(4.45), Inches(3.3), Inches(0.42),
           detail, size=12, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.15), Inches(12.35), Inches(1.05),
              L_TEAL)
    tb(slide, Inches(0.85), Inches(5.27), Inches(11.6), Inches(0.34),
       "✅ NOW WRITE IT CORRECTLY", size=11, bold=True, color=TEAL)
    tb(slide, Inches(0.85), Inches(5.65), Inches(11.6), Inches(0.42),
       "_______________________________________________________________"
       "_______________", size=15, color=SOFT)
    hint(slide, "Read it aloud as a finished sentence. What is missing at the "
                "start, in the middle and at the end?", 6.45)


def s56_report_frames():
    slide, n = new_slide("🎤 Junior Science Reporter", "SPEAKING", "78–84 min",
                         "Detective HQ", ROSE)
    one_task(slide, "You are reporting the case. Three questions to answer.",
             ROSE)
    for i, (icon, name, detail, color, light) in enumerate(REPORT_FRAMES):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(4.3), light)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(0.62), color)
        tb(slide, left, Inches(2.02), Inches(3.9), Inches(0.42), name,
           size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.75), Inches(3.9), Inches(0.75), icon,
           size=32, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.3), Inches(3.6), Inches(3.3), Inches(0.5),
           detail, size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.3), Inches(4.2), Inches(3.3),
                  Inches(1.75), WHITE)
        tb(slide, left + Inches(0.5), Inches(4.4), Inches(2.9), Inches(1.4),
           "______________\n\n______________\n\n______________", size=13,
           color=SOFT)
    hint(slide, "She can write notes here first, then speak from them.", 6.45)


def s57_starters():
    slide, n = new_slide("✍️ Sentence Starters and Short Write", "WRITING",
                         "78–84 min", "Detective HQ", TEAL)
    one_task(slide, "Finish each line in your own words.", TEAL)
    for i, starter in enumerate(REPORT_STARTERS):
        top = Inches(1.88 + i * 1.0)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.88), L_TEAL)
        add_oval(slide, Inches(0.8), top + Inches(0.18), Inches(0.52),
                 Inches(0.52), TEAL)
        tb(slide, Inches(0.8), top + Inches(0.24), Inches(0.52), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.6), top + Inches(0.2), Inches(4.4), Inches(0.5),
           starter, size=21, bold=True, color=INK)
        add_round(slide, Inches(6.2), top + Inches(0.17), Inches(6.35),
                  Inches(0.54), WHITE)
        tb(slide, Inches(6.45), top + Inches(0.26), Inches(5.9), Inches(0.4),
           "______________________________________", size=15, color=SOFT)
    add_round(slide, Inches(0.5), Inches(5.9), Inches(12.35), Inches(0.56),
              L_VIOLET)
    tb(slide, Inches(0.8), Inches(6.0), Inches(11.7), Inches(0.4),
       "⭐ EXPERT CHALLENGE — Join two of your lines with \"because\" or "
       "\"so\" to make one longer sentence.", size=14, bold=True, color=VIOLET)
    hint(slide, "Accept her wording. Fix only the grammar, not the ideas.",
         6.56)


def s58_speaking():
    slide, n = new_slide("🎤 Speaking Challenge — Present the Case", "SPEAKING",
                         "78–84 min", "Detective HQ", ROSE)
    one_task(slide, "Three or four sentences, out loud, no notes if you can.",
             ROSE)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4),
              L_ROSE)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.4), Inches(1.7), "🎤",
       size=84, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.6), Inches(4.4), Inches(5.2), Inches(0.8),
       "\"Here is what we found.\"", size=23, bold=True, color=ROSE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.35), Inches(4.8), Inches(0.6),
       "start with this line if you get stuck", size=13, color=SOFT,
       align=PP_ALIGN.CENTER)
    for i, (icon, line) in enumerate(SPEAK_CHECKS):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.25), top, Inches(6.6), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.5), top + Inches(0.2), Inches(0.56),
                 Inches(0.56), WHITE)
        tb(slide, Inches(6.5), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=14, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.3), top + Inches(0.24), Inches(4.6), Inches(0.5),
           line, size=16, bold=True, color=INK)
        tb(slide, Inches(12.1), top + Inches(0.22), Inches(0.5), Inches(0.5),
           "☐", size=17, color=ROSE)
    hint(slide, "Prompt with a question, never with the sentence itself.", 6.45)


def s59_final_grammar():
    slide, n = new_slide("⚡ Final Lab Game — Grammar", "CHALLENGE",
                         "84–88 min", "Detective HQ", INDIGO)
    one_task(slide, "Two sentences. Pick the verb that agrees.", INDIGO)
    for i, (sentence, options, _answer) in enumerate(FINAL_GRAMMAR):
        top = Inches(2.0 + i * 2.0)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.75), WHITE)
        add_oval(slide, Inches(0.8), top + Inches(0.6), Inches(0.52),
                 Inches(0.52), INDIGO)
        tb(slide, Inches(0.8), top + Inches(0.66), Inches(0.52), Inches(0.4),
           str(i + 1), size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.6), top + Inches(0.55), Inches(5.6), Inches(0.62),
           sentence, size=24, bold=True, color=INK)
        for j, opt in enumerate(options):
            left = Inches(7.5 + j * 2.6)
            add_round(slide, left, top + Inches(0.52), Inches(2.4),
                      Inches(0.7), L_INDIGO)
            tb(slide, left, top + Inches(0.63), Inches(2.4), Inches(0.5),
               f"{chr(65 + j)}.  {opt}", size=19, bold=True, color=INK,
               align=PP_ALIGN.CENTER)
    hint(slide, "Hints are still allowed. Finishing the case is what counts.",
         6.45)


def s60_final_reading():
    slide, n = new_slide("⚡ Final Lab Game — Reading and Vocabulary",
                         "CHALLENGE", "84–88 min", "Detective HQ", CRIMSON)
    one_task(slide, "One reading question, two word questions.", CRIMSON)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(1.1),
              L_CRIMSON)
    tb(slide, Inches(0.85), Inches(1.98), Inches(2.5), Inches(0.34),
       "📖 READING", size=11, bold=True, color=CRIMSON)
    tb(slide, Inches(0.85), Inches(2.32), Inches(11.6), Inches(0.5),
       FINAL_READING, size=21, bold=True, color=INK)
    for i, (question, options, _answer) in enumerate(FINAL_VOCAB):
        top = Inches(3.15 + i * 1.6)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.4), WHITE)
        tb(slide, Inches(0.85), top + Inches(0.18), Inches(4.4), Inches(0.5),
           question, size=19, bold=True, color=INK)
        tb(slide, Inches(0.85), top + Inches(0.74), Inches(4.4), Inches(0.36),
           "📚 vocabulary", size=11, color=SOFT)
        for j, opt in enumerate(options):
            left = Inches(5.5 + j * 2.4)
            add_round(slide, left, top + Inches(0.32), Inches(2.25),
                      Inches(0.76), L_TEAL)
            tb(slide, left, top + Inches(0.44), Inches(2.25), Inches(0.56),
               opt, size=13, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Ask her to use the word in a sentence after she picks.", 6.45)


def s61_final_tense():
    slide, n = new_slide("⚡ Final Lab Game — Tense and Speaking", "CHALLENGE",
                         "84–88 min", "Detective HQ", AMBER)
    one_task(slide, "One rewrite, then one sentence of your own.", AMBER)
    present, _past = FINAL_TENSE
    add_round(slide, Inches(0.5), Inches(1.9), Inches(12.35), Inches(1.9),
              L_AMBER)
    add_round(slide, Inches(0.8), Inches(2.12), Inches(1.7), Inches(0.44),
              WHITE)
    tb(slide, Inches(0.8), Inches(2.17), Inches(1.7), Inches(0.34), "PRESENT",
       size=10, bold=True, color=TEAL, align=PP_ALIGN.CENTER)
    tb(slide, Inches(2.7), Inches(2.1), Inches(9.8), Inches(0.5), present,
       size=24, bold=True, color=INK)
    add_round(slide, Inches(0.8), Inches(2.8), Inches(1.7), Inches(0.44),
              VIOLET)
    tb(slide, Inches(0.8), Inches(2.85), Inches(1.7), Inches(0.34), "PAST",
       size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(2.7), Inches(2.75), Inches(9.8), Inches(0.78),
              WHITE)
    tb(slide, Inches(2.95), Inches(2.95), Inches(9.3), Inches(0.42),
       "______________________________________________________________",
       size=15, color=SOFT)
    add_round(slide, Inches(0.5), Inches(4.0), Inches(12.35), Inches(2.2),
              L_ROSE)
    tb(slide, Inches(0.85), Inches(4.18), Inches(11.6), Inches(0.34),
       "🎤 SPEAKING — finish this sentence out loud", size=11, bold=True,
       color=ROSE)
    tb(slide, Inches(0.85), Inches(4.6), Inches(11.6), Inches(0.7),
       FINAL_SPEAK, size=30, bold=True, color=INK, font="Georgia")
    add_round(slide, Inches(0.85), Inches(5.45), Inches(11.6), Inches(0.6),
              WHITE)
    tb(slide, Inches(1.1), Inches(5.58), Inches(11.1), Inches(0.4),
       "example:  Scientists observe carefully because small clues matter.",
       size=14, color=SOFT)
    hint(slide, "Any sensible reason is correct. The \"because\" half is the "
                "real task.", 6.45)


def s62_champion():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), TEAL)
    for x, y in [(0.35, 3.95), (12.3, 3.95), (0.6, 5.45), (12.1, 5.45)]:
        tb(slide, Inches(x), Inches(y), Inches(0.8), Inches(0.7), "⭐", size=24,
           align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(0.85), Inches(12), Inches(1.1), "🏆",
       size=56, align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.95), Inches(12), Inches(0.9),
       "JUNIOR SCIENCE DETECTIVE", size=38, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.9), Inches(12), Inches(0.5),
       "You solved the mystery — with evidence, not guesses.", size=18,
       color=RGBColor(0x8F, 0xD3, 0xD8), align=PP_ALIGN.CENTER, italic=True)
    for i, (icon, label) in enumerate(ACHIEVEMENTS):
        left = Inches(1.35 + i * 2.15)
        add_round(slide, left, Inches(3.6), Inches(1.95), Inches(1.5),
                  RGBColor(0x1D, 0x26, 0x31))
        tb(slide, left, Inches(3.77), Inches(1.95), Inches(0.6), icon,
           size=20, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.12), Inches(4.37), Inches(1.7), Inches(0.62),
           label, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(2.6), Inches(5.3), Inches(8.1), Inches(1.15), TEAL)
    tb(slide, Inches(2.6), Inches(5.58), Inches(8.1), Inches(0.62),
       "\"I READ. I THINK. I SOLVE!\"", size=29, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(1.2), Inches(6.5), Inches(10.9), Inches(0.42),
       "I CAN READ, THINK, AND SOLVE LIKE A SCIENTIST!", size=16, bold=True,
       color=RGBColor(0x8F, 0xD3, 0xD8), align=PP_ALIGN.CENTER)
    footer(slide, n, "88–90 min", "Detective HQ")
    fade(slide)


def s63_can_do():
    slide, n = new_slide("✅ Today I Can...", "CHECKLIST", "88–90 min",
                         "Detective HQ", TEAL)
    one_task(slide, "Tick every one you did today. Read them with me.", TEAL)
    for i, (icon, skill, example) in enumerate(CAN_DO):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.85 + row * 0.8)
        add_round(slide, left, top, Inches(5.95), Inches(0.7), L_TEAL)
        tb(slide, left + Inches(0.25), top + Inches(0.11), Inches(0.5),
           Inches(0.48), icon, size=14, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.85), top + Inches(0.1), Inches(2.4),
           Inches(0.48), skill, size=14, bold=True, color=INK)
        tb(slide, left + Inches(3.35), top + Inches(0.14), Inches(2.1),
           Inches(0.42), example, size=11, color=SOFT)
        tb(slide, left + Inches(5.45), top + Inches(0.1), Inches(0.4),
           Inches(0.48), "☐", size=16, color=TEAL)
    hint(slide, "Read the list aloud together. Hearing the whole list matters.",
         6.66)


def s64_extra_a():
    slide, n = new_slide("🎲 Extra Game Bank — If There Is Time", "OPTIONAL",
                         "", "Detective HQ", VIOLET)
    one_task(slide, "Five spare games. Use any of them, in any order.", VIOLET)
    for i, (icon, name, prompt, cards, color, light) in enumerate(
            EXTRA_GAMES[:3]):
        left = Inches(0.5 + i * 4.22)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(4.3), light)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(0.5), color)
        tb(slide, left, Inches(1.98), Inches(3.9), Inches(0.36),
           f"GAME {i + 1} — {name}", size=12, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(3.9), Inches(0.62), icon, size=26,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(3.2), Inches(3.4), Inches(0.5),
           prompt, size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        for j, card in enumerate(cards):
            add_round(slide, left + Inches(0.4), Inches(3.85 + j * 0.78),
                      Inches(3.1), Inches(0.66), WHITE)
            tb(slide, left + Inches(0.4), Inches(3.97 + j * 0.78),
               Inches(3.1), Inches(0.48), card, size=14, bold=True,
               color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Every game reuses today's words and sentences. Nothing new.",
         6.45)


def s65_extra_b():
    slide, n = new_slide("🎲 Extra Game Bank — Two More", "OPTIONAL", "",
                         "Detective HQ", VIOLET)
    one_task(slide, "One sentence puzzle and one evidence game.", VIOLET)
    for i, (icon, name, prompt, cards, color, light) in enumerate(
            EXTRA_GAMES[3:]):
        left = Inches(0.5 + i * 4.22)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(4.3), light)
        add_round(slide, left, Inches(1.9), Inches(3.9), Inches(0.5), color)
        tb(slide, left, Inches(1.98), Inches(3.9), Inches(0.36),
           f"GAME {i + 4} — {name}", size=12, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tb(slide, left, Inches(2.5), Inches(3.9), Inches(0.62), icon, size=26,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(3.2), Inches(3.4), Inches(0.5),
           prompt, size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)
        for j, card in enumerate(cards):
            add_round(slide, left + Inches(0.4), Inches(3.85 + j * 0.78),
                      Inches(3.1), Inches(0.66), WHITE)
            tb(slide, left + Inches(0.4), Inches(3.97 + j * 0.78),
               Inches(3.1), Inches(0.48), card, size=14, bold=True,
               color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(8.94), Inches(1.9), Inches(3.9), Inches(4.3),
              L_GREY)
    tb(slide, Inches(9.2), Inches(2.1), Inches(3.4), Inches(0.4),
       "WHEN TO USE THESE", size=13, bold=True, color=VIOLET)
    bullets(slide, Inches(9.2), Inches(2.6), Inches(3.4), Inches(3.4),
            ["She finishes a mission early.", "Focus is dropping mid-block.",
             "A skill needs one more pass.", "You have five minutes spare.",
             "She asks to play one more."], size=12, sp=10)
    hint(slide, "Game 5 teaches that evidence must connect to the mystery.",
         6.45)


def s66_support():
    slide, n = new_slide("🧰 Support System — For the Teacher", "TEACHER", "",
                         "Detective HQ", INDIGO)
    one_task(slide, "Levels, support ladder and the words to say.", INDIGO)
    for i, (icon, name, who, action, color, light) in enumerate(
            SUPPORT_LEVELS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(2.25), light)
        tb(slide, left + Inches(0.25), Inches(2.02), Inches(0.5), Inches(0.5),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.9), Inches(2.02), Inches(2.9), Inches(0.5),
           name, size=14, bold=True, color=color)
        tb(slide, left + Inches(0.3), Inches(2.6), Inches(3.3), Inches(0.5),
           who, size=12, color=DARK)
        add_round(slide, left + Inches(0.3), Inches(3.15), Inches(3.3),
                  Inches(0.78), WHITE)
        tb(slide, left + Inches(0.45), Inches(3.26), Inches(3.0), Inches(0.58),
           action, size=11, color=INK)
    tb(slide, Inches(0.5), Inches(4.25), Inches(6.0), Inches(0.4),
       "SUPPORT LADDER — never reveal the answer first", size=13, bold=True,
       color=INDIGO)
    for i, (label, text, color) in enumerate(SUPPORT_LADDER):
        top = Inches(4.66 + i * 0.46)
        add_round(slide, Inches(0.5), top, Inches(6.0), Inches(0.42), L_GREY)
        add_round(slide, Inches(0.62), top + Inches(0.05), Inches(1.2),
                  Inches(0.32), color)
        tb(slide, Inches(0.62), top + Inches(0.07), Inches(1.2), Inches(0.28),
           label, size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.0), top + Inches(0.05), Inches(4.3), Inches(0.32),
           text, size=12, bold=True, color=INK)
    tb(slide, Inches(6.85), Inches(4.25), Inches(3.0), Inches(0.4),
       "HINT BANK", size=13, bold=True, color=INDIGO)
    bullets(slide, Inches(6.85), Inches(4.72), Inches(3.0), Inches(2.1),
            HINT_BANK, size=11)
    tb(slide, Inches(10.0), Inches(4.25), Inches(2.85), Inches(0.4),
       "WORDS TO USE", size=13, bold=True, color=TEAL)
    bullets(slide, Inches(10.0), Inches(4.72), Inches(2.85), Inches(2.1),
            PRAISE, size=11, color=TEAL)


def s67_assessment():
    slide, n = new_slide("📋 End-of-Class Assessment", "TEACHER", "",
                         "Detective HQ", AMBER)
    one_task(slide, "Tick one box per skill right after the lesson.", AMBER)
    heads = ["SKILL", "INDEPENDENT", "WITH SUPPORT", "NEEDS MORE PRACTICE"]
    widths = [5.15, 2.4, 2.4, 2.4]
    lefts = [0.5]
    for w in widths[:-1]:
        lefts.append(lefts[-1] + w)
    add_rect(slide, Inches(0.5), Inches(1.82), Inches(12.35), Inches(0.42),
             AMBER)
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
              L_AMBER)
    for i, item in enumerate(REVIEW_NEXT):
        tb(slide, Inches(0.8 + i * 4.1), Inches(6.45), Inches(3.9),
           Inches(0.34), f"{item}:  ____________", size=12, bold=True,
           color=INK)


def s68_key_a():
    slide, n = new_slide("🔑 Answer Key — Vocabulary and Grammar", "TEACHER",
                         "", "Detective HQ", TEAL)
    one_task(slide, "Slides 10 to 37.", TEAL)
    cols = [
        ("VOCABULARY (7–17)",
         ["Word match: 🔍 · 💡 · 🔎 · 🔮 · 📊",
          "Science words: experiment · evidence ·",
          " discover · predict",
          "Not science: banana · pencil"], TEAL),
        ("SUBJECT-VERB (17–27)",
         ["Verb choice: works · work · looks",
          " look · mixes · mix",
          "Fixes: works · study · look · checks",
          "Rule: one subject → -s verb"], INDIGO),
        ("PARTS OF SPEECH (27–40)",
         ["student/solution · mixes",
          "scientist/liquid · measures · careful",
          "scientist/result · checks · carefully",
          "Sort: nouns scientist, experiment · verb measure ·",
          " adjectives careful, bright · adverb quickly",
          "Fast lab: checks · strange · carefully"], VIOLET),
        ("VERB TENSES (40–50)",
         ["Tense ID: present · past · future",
          "To past: measured · recorded",
          "To future: will measure · will record",
          "Quick challenge: checked · will record",
          "Freeze sentence: careful/strange · checked ·",
          " quickly · scientist/sample"], AMBER),
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
              L_TEAL)
    tb(slide, Inches(0.8), Inches(6.76), Inches(11.7), Inches(0.32),
       "Reading and speaking answers are on the next slide.", size=11,
       bold=True, color=TEAL)


def s69_key_b():
    slide, n = new_slide("🔑 Answer Key — Reading and Editing", "TEACHER", "",
                         "Detective HQ", CRIMSON)
    one_task(slide, "Slides 45 to 61.", CRIMSON)
    cols = [
        ("COMPREHENSION (60–70)",
         ["Main idea: B", "Q1 a sample · Q2 a wet mark near the sink",
          "Q3 it had to stay cold · Q4 in the refrigerator",
          "Sequence: B → A → C → D",
          "True/False: FALSE · TRUE · TRUE", "Inference: B"], CRIMSON),
        ("EVIDENCE LINES (60–70)",
         ["\"Mr. Alvarez explained that he had moved",
          " the jar himself.\"",
          "\"Cold things sweat when they sit in warm air.\"",
          "\"The students were not guessing now.",
          " They were collecting evidence.\""], VIOLET),
        ("EDITING (70–78)",
         ["The students were ready.",
          "The scientist checks the sample.",
          "students/clues · looked · carefully",
          "strange · sample · was",
          "The students found Sample C in the",
          " refrigerator."], AMBER),
        ("FINAL GAME (84–88)",
         ["Grammar: checks · record",
          "Reading: Sample C, in the refrigerator",
          "Vocabulary: a clue or information ·",
          " to say what will happen",
          "Tense: The scientist checked the sample.",
          "Speaking: accept any sensible reason"], TEAL),
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
              L_CRIMSON)
    tb(slide, Inches(0.8), Inches(6.76), Inches(11.7), Inches(0.32),
       "Accept any sensible answer for the summary, the report and the "
       "speaking tasks — reasoning matters more than exact wording.",
       size=11, bold=True, color=CRIMSON)


BUILDERS = [
    s01_title, s02_detective, s03_map, s04_warmup, s05_observe, s06_discover,
    s07_evidence, s08_predict, s09_result, s10_word_match, s11_word_or_not,
    s12_say_it, s13_vocab_bank, s14_subject_verb, s15_sva_scientist,
    s16_sva_experiment, s17_sva_rule, s18_verb_choice_a, s19_verb_choice_b,
    s20_fix_a, s21_fix_b, s22_observation_done, s23_pos_legend, s24_pos_a,
    s25_pos_b, s26_pos_c, s27_grammar_sort, s28_fast_lab, s29_experiment_done,
    s30_break_intro, s31_scientist_says, s32_freeze_challenge, s33_tenses,
    s34_tense_id, s35_to_past, s36_to_future, s37_tense_choice,
    s38_research_done, s39_story_intro, s40_story_1, s41_story_2, s42_story_3,
    s43_story_4, s44_story_5, s45_main_idea, s46_clue_q_a, s47_clue_q_b,
    s48_sequence, s49_evidence, s50_true_false, s51_inference,
    s52_mystery_done, s53_editor_a, s54_editor_pos, s55_editor_challenge,
    s56_report_frames, s57_starters, s58_speaking, s59_final_grammar,
    s60_final_reading, s61_final_tense, s62_champion, s63_can_do, s64_extra_a,
    s65_extra_b, s66_support, s67_assessment, s68_key_a, s69_key_b,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL}, built {len(prs.slides)}"

OUT = "Grade5_Science_Lab_Mystery_90min.pptx"
prs.save(OUT)

passage_words = sum(len(block.split()) for part in STORY for block in part[4])
with_notes = sum(1 for s in prs.slides
                 if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip())
print(f"saved {OUT}")
print(f"slides: {len(prs.slides)}")
print(f"passage words: {passage_words}")
print(f"slides with notes: {with_notes}")
