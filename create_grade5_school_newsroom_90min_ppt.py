"""Grade 5 English lesson - 90 minutes, 61 slides, no speaker notes.

"The School Newsroom" - the student is a Young Reporter hired by the school
paper. Five missions carry reading, vocabulary, grammar, evidence work, editing
and speaking: Headline Hunter, Word Detective, Grammar Editor, Story Reporter
and the Final News Challenge.

Teaching support sits on visible slides rather than in speaker notes: every
activity carries a Reporter Hint strip, and slides 59-61 hold the support
system, the assessment table and the answer key.
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree

INK = RGBColor(0x11, 0x14, 0x18)
DARK = RGBColor(0x2A, 0x2F, 0x36)
SOFT = RGBColor(0x78, 0x82, 0x8E)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
PAPER = RGBColor(0xF7, 0xF6, 0xF2)
RED = RGBColor(0xC8, 0x32, 0x2B)
BLUE = RGBColor(0x1B, 0x5E, 0x9E)
GREEN = RGBColor(0x1E, 0x7A, 0x52)
AMBER = RGBColor(0xD0, 0x8A, 0x0C)
PURPLE = RGBColor(0x6A, 0x4A, 0x9E)
TEAL = RGBColor(0x15, 0x80, 0x7C)
SLATE = RGBColor(0x46, 0x58, 0x6B)
L_RED = RGBColor(0xFB, 0xE9, 0xE7)
L_BLUE = RGBColor(0xE6, 0xEF, 0xF8)
L_GREEN = RGBColor(0xE4, 0xF2, 0xEB)
L_AMBER = RGBColor(0xFC, 0xF1, 0xDA)
L_PURPLE = RGBColor(0xEE, 0xEA, 0xF7)
L_TEAL = RGBColor(0xE1, 0xF1, 0xF0)
L_SLATE = RGBColor(0xEB, 0xEE, 0xF2)
L_GREY = RGBColor(0xF2, 0xF3, 0xF4)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 61
_counter = {"n": 0}

POS_COLOR = {"noun": BLUE, "verb": GREEN, "adj": PURPLE, "adv": AMBER}

# ------------------------------------------------------------------ content

DESKS = [("🗞️", "Headline Desk", "Main idea", "0–17 min", RED, L_RED),
         ("🔎", "Word Desk", "Vocabulary", "17–27 min", PURPLE, L_PURPLE),
         ("✏️", "Editor's Desk", "Grammar", "27–49 min", GREEN, L_GREEN),
         ("📖", "Reading Desk", "The news story", "49–77 min", BLUE, L_BLUE),
         ("🎤", "Press Room", "Report & publish", "77–90 min", AMBER, L_AMBER)]

MISSIONS = [("🗞️", "Mission 1", "Headline Hunter", "Find the main idea fast.", RED),
            ("🔎", "Mission 2", "Word Detective", "Own twelve newsroom words.",
             PURPLE),
            ("✏️", "Mission 3", "Grammar Editor", "Catch every mistake.", GREEN),
            ("📖", "Mission 4", "Story Reporter", "Read and prove your answers.",
             BLUE),
            ("🏆", "Mission 5", "Final News Challenge", "Publish the edition.",
             AMBER)]

DESK_KIT = [("📰", "newspaper"), ("🎤", "microphone"), ("📷", "camera"),
            ("📋", "notebook"), ("✏️", "pencil"), ("🗞️", "headline")]

GOOD_STORY = [("A", "It has information."), ("B", "It tells what happened."),
              ("C", "It uses clear sentences."), ("D", "All of the above.")]

WARM_STARTERS = ["I think ______.", "The picture shows ______.", "Maybe ______."]

METHOD = [("1", "TEACHER MODEL", "I show you how it works.", RED),
          ("2", "TOGETHER", "We do the next one out loud.", AMBER),
          ("3", "YOU TRY", "You take one on your own.", GREEN),
          ("4", "GAME", "We turn it into a challenge.", PURPLE),
          ("5", "APPLY", "You use it in your own writing.", BLUE)]

HEADLINES = ["CLASS PLANTS A SCHOOL GARDEN", "MYSTERY PUPPY VISITS SCHOOL",
             "STUDENTS WIN SCIENCE FAIR", "NEW BOOK CLUB STARTS"]

HEADLINE_TOPICS = [("🌻", "CLASS PLANTS A SCHOOL GARDEN", "planting"),
                   ("🐶", "MYSTERY PUPPY VISITS SCHOOL", "a visitor"),
                   ("🏅", "STUDENTS WIN SCIENCE FAIR", "winning"),
                   ("📚", "NEW BOOK CLUB STARTS", "starting something")]

MATCH_A = [("🌻", ["STUDENTS WIN SCIENCE FAIR", "CLASS PLANTS A SCHOOL GARDEN",
                   "NEW BOOK CLUB STARTS"]),
           ("🐶", ["MYSTERY PUPPY VISITS SCHOOL", "PIZZA RETURNS TO THE MENU",
                   "ART SHOW FILLS THE HALL"]),
           ("🏅", ["NEW BOOK CLUB STARTS", "ART SHOW FILLS THE HALL",
                   "STUDENTS WIN SCIENCE FAIR"])]
MATCH_B = [("📚", ["NEW BOOK CLUB STARTS", "STUDENTS WIN SCIENCE FAIR",
                   "MYSTERY PUPPY VISITS SCHOOL"]),
           ("🍕", ["CLASS PLANTS A SCHOOL GARDEN", "PIZZA RETURNS TO THE MENU",
                   "NEW BOOK CLUB STARTS"]),
           ("🎨", ["ART SHOW FILLS THE HALL", "MYSTERY PUPPY VISITS SCHOOL",
                   "PIZZA RETURNS TO THE MENU"])]

MIXUP = [(["GARDEN", "SCHOOL", "GROWS", "GIANT", "PUMPKIN"],
          "SCHOOL GARDEN GROWS GIANT PUMPKIN"),
         (["WIN", "STUDENTS", "FAIR", "SCIENCE"], "STUDENTS WIN SCIENCE FAIR"),
         (["CLUB", "BOOK", "STARTS", "NEW"], "NEW BOOK CLUB STARTS")]

VOCAB_1 = [("REPORTER", "🎤", "a person who finds news and writes about it",
            "The reporter asked ten questions."),
           ("HEADLINE", "🗞️", "the big title at the top of a news story",
            "The headline used only five words."),
           ("GARDEN", "🌻", "a place where plants are grown",
            "The garden is behind the school."),
           ("PROJECT", "📋", "a piece of work with a goal",
            "Our class project took one month.")]
VOCAB_2 = [("DISCOVER", "🔎", "to find something new",
            "Students discovered a tiny bird's nest."),
           ("EXCITED", "🤩", "feeling happy and eager",
            "The excited team clapped and cheered."),
           ("CAREFULLY", "🧐", "doing something with full attention",
            "She carefully wrote down every name."),
           ("INTERVIEW", "🎙️", "to ask someone questions for a story",
            "I will interview the principal today.")]
VOCAB_3 = [("TEAM", "🤝", "a group of people working together",
            "The team met every Tuesday."),
           ("PRACTICE", "🔁", "to do something again to get better",
            "We practice reading out loud."),
           ("CELEBRATE", "🎉", "to show you are happy about something",
            "They celebrated the first red tomato."),
           ("SCHOOL", "🏫", "a place where students learn",
            "Our school has its own newspaper.")]

WORD_OR_NOT = [("REPORTER", True), ("BANANA", False), ("HEADLINE", True),
               ("GARDEN", True), ("ROCKET", False), ("INTERVIEW", True),
               ("PENGUIN", False), ("EDITOR", True)]

WHO_AM_I = [("🗞️", ["I am the big title at the top.",
                     "I give the main idea in very few words.",
                     "I am usually printed in capital letters."], "HEADLINE"),
            ("🎤", ["I ask people questions.", "I write down what they say.",
                    "I work for the newspaper."], "REPORTER"),
            ("🔎", ["I mean to find something new.",
                    "Scientists do me all the time.",
                    "The students did me when the birds arrived."], "DISCOVER")]

CONTEXT = [("Priya will ______ the principal after school.",
            ["interview", "celebrate", "practice"]),
           ("The team worked ______ so that no seeds were broken.",
            ["quickly", "carefully", "loudly"]),
           ("We will ______ when the first tomato turns red.",
            ["discover", "interview", "celebrate"])]

SUBJ_PRED = [("The students", "practice every morning."),
             ("Our school newspaper", "prints a new edition on Friday."),
             ("Priya", "interviewed the science teacher."),
             ("The excited reporters", "celebrate a good story.")]

AGREEMENT = [("The student ______ a headline.", ["write", "writes"], "writes"),
             ("The students ______ a headline.", ["write", "writes"], "write"),
             ("The reporter ______ a question.", ["ask", "asks"], "asks"),
             ("The reporters ______ a question.", ["ask", "asks"], "ask")]

AGREE_PRACTICE = [("Mr. Ellis ______ the team.", ["help", "helps"], "helps"),
                  ("Butterflies ______ to the garden.", ["come", "comes"], "come"),
                  ("My class ______ vegetables.", ["plant", "plants"], "plants"),
                  ("The birds ______ a nest.", ["build", "builds"], "build"),
                  ("Priya ______ every answer.", ["check", "checks"], "checks")]

RED_PEN_A = [("The students writes a story.", "students write"),
             ("The reporter ask a question.", "reporter asks"),
             ("My friend read the newspaper every day.", "friend reads")]
RED_PEN_B = [("The teachers helps the team.", "teachers help"),
             ("Priya and Sam plants the seeds.", "Priya and Sam plant")]

POS_KEY = [("NOUN", "a person, place or thing", "reporter, garden, story", BLUE,
            L_BLUE),
           ("VERB", "the action word", "wrote, planted, celebrate", GREEN, L_GREEN),
           ("ADJECTIVE", "describes a noun", "excited, small, busy", PURPLE,
            L_PURPLE),
           ("ADVERB", "describes a verb — often -ly", "quickly, carefully", AMBER,
            L_AMBER)]

SPIN_WHEEL = [("🔵", "NOUN", "Find the noun.", BLUE, L_BLUE),
              ("🟢", "VERB", "Find the verb.", GREEN, L_GREEN),
              ("🟣", "ADJECTIVE", "Find the adjective.", PURPLE, L_PURPLE),
              ("🟡", "ADVERB", "Find the adverb.", AMBER, L_AMBER),
              ("🔴", "FIX IT", "Fix the sentence.", RED, L_RED)]

SPIN_1 = [("The", None), ("excited", "adj"), ("reporter", "noun"),
          ("quickly", "adv"), ("wrote", "verb"), ("the", None), ("story.", "noun")]
SPIN_2 = [("The", None), ("busy", "adj"), ("students", "noun"),
          ("carefully", "adv"), ("planted", "verb"), ("the", None), ("small", "adj"),
          ("seeds.", "noun")]
SPIN_3 = [("Our", None), ("young", "adj"), ("editor", "noun"), ("proudly", "adv"),
          ("printed", "verb"), ("the", None), ("first", "adj"), ("page.", "noun")]

SPIN_TASKS_1 = [("🔵", "Find the noun", "2 of them hiding"),
                ("🟢", "Find the verb", "only 1"),
                ("🟣", "Find the adjective", "it describes the reporter"),
                ("🟡", "Find the adverb", "it ends in -ly")]

FIX_SPIN = [("the garden grow fast", "The garden grows fast."),
            ("priya write three headlines", "Priya writes three headlines."),
            ("the birds builds a nest", "The birds build a nest.")]

CONJUNCTIONS = [("and", "joins two ideas", "The students planted seeds ___ watered "
                 "them.", GREEN),
                ("but", "shows a surprise", "Priya was tired ___ she kept working.",
                 RED),
                ("so", "shows a result", "The soil was hard, ___ they added new "
                 "soil.", BLUE),
                ("because", "gives a reason", "Birds came ___ the flowers "
                 "attracted insects.", PURPLE)]

NEWSROOM_ACTIONS = [("📷", "Camera!", "Pretend to take a picture."),
                    ("🎤", "Microphone!", "Pretend to interview someone."),
                    ("📋", "Notebook!", "Pretend to write fast."),
                    ("🗞️", "Headline!", "Strike a dramatic headline pose."),
                    ("🚨", "Breaking News!", "Whisper: \"BREAKING NEWS!\""),
                    ("🔀", "Mix them up!", "Now the teacher calls them in any "
                     "order.")]

TENSE_LINE = [("PAST", "yesterday", "The reporter interviewed the teacher.", RED,
               L_RED),
              ("NOW", "today", "The reporter interviews the teacher.", GREEN,
               L_GREEN),
              ("FUTURE", "tomorrow", "The reporter will interview the teacher.",
               BLUE, L_BLUE)]

TENSE_CHART = [("interview", "interviews", "interviewed", "will interview"),
               ("plant", "plants", "planted", "will plant"),
               ("win", "wins", "won", "will win"),
               ("write", "writes", "wrote", "will write"),
               ("discover", "discovers", "discovered", "will discover")]

NAME_TENSE = [("The reporter interviews the teacher.", "PRESENT"),
              ("The team won the game.", "PAST"),
              ("We will publish the paper on Friday.", "FUTURE"),
              ("Priya wrote three headlines.", "PAST"),
              ("The students celebrate every Friday.", "PRESENT")]

TRANSFORM = [("The team wins the game.", "The team won the game.",
              "The team will win the game."),
             ("The reporter asks a question.", "The reporter asked a question.",
              "The reporter will ask a question."),
             ("We plant beans.", "We planted beans.", "We will plant beans.")]

IRREGULAR = [("grow", "grew", "will grow", "The garden ______ fast last summer."),
             ("take", "took", "will take", "She ______ a photo yesterday."),
             ("write", "wrote", "will write", "They ______ the headline last "
              "night.")]

STORY_WATCH = [("empty", "🕳️"), ("measured", "📏"), ("schedule", "🗓️"),
               ("cleared", "🧹"), ("attracted", "🦋"), ("busiest", "🏫")]

READ_STEPS = [("STEP 1", "TEACHER MODEL", "I read the first part.", RED),
              ("STEP 2", "SHARED READING", "We read the next part together.", AMBER),
              ("STEP 3", "YOU READ", "You read the sentences I point to.", GREEN),
              ("STEP 4", "CHECK", "One question before we move on.", BLUE)]

STORY = [
    ("Part 1", "The Empty Space", "🕳️", RED, L_RED,
     ["Behind Lincoln Elementary School, there was an empty space. Nobody used it. "
      "The ground was dry, and old boxes were stacked against the fence. Students "
      "walked past it every day without stopping.",
      "One morning in March, a fifth grader named Priya looked at the space and had "
      "an idea. \"We could grow something here,\" she said to her friends. Her "
      "teacher, Mr. Ellis, listened carefully and agreed to help."],
     "Who had the idea for the garden?"),
    ("Part 2", "The Plan", "🗓️", PURPLE, L_PURPLE,
     ["The students formed a team. First, they measured the space and drew a map. "
      "Next, they made a list of the plants they wanted. Some students wanted "
      "flowers. Others wanted vegetables, because they liked the idea of growing "
      "food.",
      "Mr. Ellis asked them a hard question. \"Who will water the garden in the "
      "summer?\" The team had not thought about that. They built a schedule so that "
      "every family took one week."],
     "What problem did Mr. Ellis point out?"),
    ("Part 3", "The Work", "🧹", GREEN, L_GREEN,
     ["In April, the work began. The students cleared the old boxes and carried them "
      "to the recycling bin. They dug the hard ground and mixed in fresh soil. It "
      "took three weekends. Priya said her arms hurt, but she came back every time.",
      "By the end of the month, they had planted sunflowers, tomatoes, beans, and "
      "carrots. They painted small wooden signs so visitors would know what was "
      "growing."],
     "How many weekends did the digging take?"),
    ("Part 4", "What They Discovered", "🦋", BLUE, L_BLUE,
     ["In June, something surprising happened. Butterflies arrived first, then bees, "
      "and finally a pair of small brown birds that built a nest near the fence. A "
      "science teacher explained that the flowers had attracted the insects, and the "
      "insects had attracted the birds.",
      "\"We thought we were growing vegetables,\" Priya told the school newspaper. "
      "\"We were really building a home for animals.\" The empty space is now the "
      "busiest corner of the school."],
     "What did Priya say they were really building?"),
]

MAIN_IDEA = ["A teacher taught a science lesson about insects.",
             "Students turned an unused space into a garden that brought wildlife.",
             "Priya's arms hurt after three weekends of digging."]

SEQUENCE = [("🕳️", "Priya looks at the empty space and has an idea."),
            ("🗓️", "The team measures, draws a map and builds a schedule."),
            ("🧹", "They clear the boxes, dig the ground and plant."),
            ("🦋", "Butterflies, bees and birds arrive.")]

EVIDENCE_STEPS = [("1", "ANSWER IT", "Say what you think.", RED),
                  ("2", "FIND IT", "Go back and hunt for the sentence.", AMBER),
                  ("3", "READ IT", "Read that sentence out loud to me.", GREEN)]

EVIDENCE_A = [("❓", "Why did the students start the garden?", "Part 1"),
              ("🌱", "What did they plant?", "Part 3")]
EVIDENCE_B = [("🦋", "What animals came to the garden?", "Part 4"),
              ("💡", "What did the students learn?", "Part 4")]
EVIDENCE_C = [("🏫", "How did the garden change the school area?", "Part 4")]

INFERENCE = ("\"Priya said her arms hurt, but she came back every time.\"",
             "What does this tell you about Priya?",
             ["She gives up easily.", "She does not like gardens.",
              "She keeps going even when it is hard."])

CONTEXT_STORY = [("attracted", "Part 4",
                  ["pushed away", "brought closer", "painted"]),
                 ("schedule", "Part 2",
                  ["a plan of who does what and when", "a type of plant",
                   "a kind of tool"]),
                 ("cleared", "Part 3", ["filled up", "removed things from",
                                        "measured"])]

EDITOR_CHECKS = [("🔠", "CAPITAL LETTERS", "first word, names, Mr.", RED),
                 ("❗", "PUNCTUATION", "period, comma, question mark", AMBER),
                 ("🔗", "SUBJECT-VERB", "one student writes / two students write",
                  GREEN),
                 ("⏰", "TENSE", "yesterday needs a past verb", BLUE),
                 ("🔤", "SPELLING", "read it letter by letter", PURPLE)]

FIX_NEWS_A = [("the students plants flowers in the garden",
               "The students plant flowers in the garden.",
               ["capital", "verb", "period"]),
              ("the reporter write a story yesterday",
               "The reporter wrote a story yesterday.",
               ["capital", "tense", "period"])]
FIX_NEWS_B = [("priya and her team works very hard",
               "Priya and her team work very hard.",
               ["capital", "verb", "period"]),
              ("the birds builds a nest quick",
               "The birds build a nest quickly.",
               ["capital", "verb", "adverb", "period"])]
FIX_NEWS_C = [("we planted sunflowers tomatoes and beans",
               "We planted sunflowers, tomatoes, and beans.",
               ["capital", "commas", "period"]),
              ("mr ellis ask a hard question",
               "Mr. Ellis asked a hard question.",
               ["capital", "spelling", "tense", "period"])]

REPORT_FIELDS = [("🗞️", "HEADLINE", "Five words or fewer."),
                 ("👤", "WHO", "Who is in the story?"),
                 ("⚡", "WHAT HAPPENED", "One clear action."),
                 ("📍", "WHERE", "Name the place."),
                 ("🔍", "ONE DETAIL", "Something only you noticed.")]

REPORT_STARTERS = ["The students ______.", "They are ______.",
                   "They worked ______.", "I think ______."]

FINAL = [("1", "📖", "READ", "Read this word out loud:", "INTERVIEW", RED, L_RED),
         ("2", "🟢", "GRAMMAR", "Find the verb:",
          "The excited reporter wrote the story.", GREEN, L_GREEN),
         ("3", "⏰", "TENSE", "Change this to past tense:", "The team wins.", BLUE,
          L_BLUE),
         ("4", "🔎", "READING", "From the story:",
          "What animals came to the garden?", PURPLE, L_PURPLE),
         ("5", "🎤", "SPEAK", "Make one sentence using:", "DISCOVER", AMBER,
          L_AMBER)]

GAME_BANK = [("📰", "HEADLINE MIX-UP", "Rearrange words into a headline.", RED),
             ("🔎", "WORD DETECTIVE", "Find one word inside a paragraph.", PURPLE),
             ("🎤", "WHO AM I?", "Clues about a newsroom word.", AMBER),
             ("✏️", "EDITOR'S CHOICE", "Pick the correct sentence of two.", GREEN),
             ("⏰", "TIME TRAVEL", "Present → past → future.", BLUE),
             ("🧩", "SENTENCE PUZZLE", "Arrange words into a sentence.", TEAL),
             ("🚨", "BREAKING NEWS", "Build a sentence from a picture.", SLATE)]

SUPPORT_LEVELS = [("🟢", "REPORTER HINT", "A picture or a clue is given",
                   "Point to the half of the sentence that matters.", GREEN,
                   L_GREEN),
                  ("🟡", "REPORTER MISSION", "She solves it with light support",
                   "Ask the question once, then stay quiet.", AMBER, L_AMBER),
                  ("⭐", "EDITOR CHALLENGE", "She explains it or writes her own",
                   "Ask: how do you know that?", PURPLE, L_PURPLE)]

HINT_LADDER = [("HINT 1", "Give a small clue.", RED),
               ("HINT 2", "Point to the part that matters.", AMBER),
               ("HINT 3", "Solve it together, then she answers.", GREEN)]

HINT_BANK = ["Read the sentence again.", "Look for who is doing the action.",
             "Look for the action word.", "Which word tells you more about the noun?",
             "Look back at paragraph 2.", "Find the sentence that proves it."]

PRAISE = ["\"Good try.\"", "\"Let's look at that word again.\"",
          "\"You found an important clue.\"", "\"Let's solve it together.\"",
          "\"Great evidence!\"", "\"Your reading is getting stronger.\""]

RUBRIC_SKILLS = ["Reading fluency", "Vocabulary", "Main idea", "Details", "Evidence",
                 "Parts of speech", "Subject-verb agreement", "Verb tenses",
                 "Sentence editing", "Speaking", "Sentence writing"]

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


def paragraph(slide, l, t, w, h, text, size=15, color=DARK, sp=8):
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.space_after = Pt(sp)
    r = p.add_run()
    r.text = text
    set_run(r, size, False, color)
    return box


def colored_sentence(slide, l, t, w, h, words, size=26, align=PP_ALIGN.LEFT):
    """Renders a sentence with each word tinted by its part of speech."""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    for i, (word, pos) in enumerate(words):
        r = p.add_run()
        r.text = word if i == len(words) - 1 else word + " "
        set_run(r, size, pos is not None, POS_COLOR.get(pos, INK))
    return box


def pos_legend(slide, top=1.33, left=0.5):
    labels = [("NOUN", BLUE), ("VERB", GREEN), ("ADJECTIVE", PURPLE),
              ("ADVERB", AMBER)]
    for i, (label, color) in enumerate(labels):
        l = Inches(left + i * 1.72)
        add_round(slide, l, Inches(top), Inches(1.6), Inches(0.34), color)
        tb(slide, l, Inches(top + 0.03), Inches(1.6), Inches(0.28), label, size=10,
           bold=True, color=WHITE, align=PP_ALIGN.CENTER)


def fade(slide):
    sld = slide._element
    for old in sld.findall(qn("p:transition")):
        sld.remove(old)
    tr = etree.SubElement(sld, qn("p:transition"))
    tr.set("spd", "med")
    tr.set("advClick", "1")
    etree.SubElement(tr, qn("p:fade"))


def footer(slide, n, timing="", desk=""):
    add_rect(slide, Inches(0), Inches(7.02), prs.slide_width, Inches(0.08),
             RGBColor(0xE2, 0xE2, 0xDE))
    add_rect(slide, Inches(0), Inches(7.02), int(prs.slide_width * n / TOTAL),
             Inches(0.08), RED)
    add_rect(slide, Inches(0), Inches(7.10), prs.slide_width, Inches(0.40), INK)
    msg = "📰 The School Newsroom  |  Grade 5  |  90 min"
    if desk:
        msg += f"  |  {desk}"
    if timing:
        msg += f"  |  {timing}"
    tb(slide, Inches(0.35), Inches(7.14), Inches(11.3), Inches(0.32), msg, size=10,
       color=WHITE)
    tb(slide, Inches(11.85), Inches(7.14), Inches(1.2), Inches(0.32), f"{n} / {TOTAL}",
       size=10, color=WHITE, align=PP_ALIGN.RIGHT)


def new_slide(title, tag, timing, desk, accent=RED, bg=PAPER):
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
    tb(slide, Inches(0.4), Inches(0.8), Inches(12.4), Inches(0.5), title, size=27,
       bold=True, color=INK, font="Georgia")
    footer(slide, n, timing, desk)
    fade(slide)
    return slide, n


def one_task(slide, text, color=RED, top=1.32):
    """States the single job of this slide in one plain line."""
    return tb(slide, Inches(0.45), Inches(top), Inches(11.5), Inches(0.4), text,
              size=15, bold=True, color=color)


def hint(slide, text, top=6.42, label="🟡 REPORTER HINT", fill=L_AMBER, color=AMBER):
    add_round(slide, Inches(0.5), Inches(top), Inches(12.35), Inches(0.46), fill)
    tb(slide, Inches(0.8), Inches(top + 0.06), Inches(2.5), Inches(0.34), label,
       size=12, bold=True, color=color)
    tb(slide, Inches(3.4), Inches(top + 0.05), Inches(9.2), Inches(0.36), text,
       size=13, bold=True, color=INK)


def model_bar(slide, top=6.42):
    steps = [("I MODEL", "Teacher", RED), ("WE TRY", "Together", AMBER),
             ("YOU GO", "Your turn", GREEN)]
    for i, (label, who, color) in enumerate(steps):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(top), Inches(3.9), Inches(0.46), color)
        tb(slide, left, Inches(top + 0.05), Inches(3.9), Inches(0.36),
           f"{label}  ·  {who}", size=13, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)


def numbered_rows(slide, count, start_index, accent, top_start=1.85, gap=1.5,
                  height=1.34, fill=WHITE):
    """Shared row scaffold; returns the top edge of each numbered row."""
    tops = []
    for i in range(count):
        top = Inches(top_start + i * gap)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(height), fill)
        add_oval(slide, Inches(0.75), top + Inches(height / 2 - 0.24), Inches(0.48),
                 Inches(0.48), accent)
        tb(slide, Inches(0.75), top + Inches(height / 2 - 0.18), Inches(0.48),
           Inches(0.38), str(start_index + i), size=12, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
        tops.append(top)
    return tops


def match_rows(slide, items, start_index, accent, light):
    """Headline Match: one picture, three headlines to choose from."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=1.85,
                         gap=1.5)
    for (emoji, options), top in zip(items, tops):
        tb(slide, Inches(1.45), top + Inches(0.3), Inches(1.1), Inches(0.75), emoji,
           size=32, align=PP_ALIGN.CENTER)
        for j, opt in enumerate(options):
            t = top + Inches(0.14 + j * 0.37)
            add_round(slide, Inches(2.85), t, Inches(9.7), Inches(0.33), light)
            tb(slide, Inches(3.1), t + Inches(0.02), Inches(9.3), Inches(0.28), opt,
               size=13, bold=True, color=INK)


def compact_rows(slide, count, start_index, accent, top_start=1.82, gap=0.9,
                 height=0.8):
    """Five short prompts stacked without crowding the hint strip."""
    return numbered_rows(slide, count, start_index, accent, top_start=top_start,
                         gap=gap, height=height)


def choice_rows(slide, items, start_index, accent, light, prompt_size=17):
    """A prompt on the left, three answer buttons on the right."""
    tops = compact_rows(slide, len(items), start_index, accent)
    for (prompt, options), top in zip(items, tops):
        tb(slide, Inches(1.45), top + Inches(0.18), Inches(5.0), Inches(0.5), prompt,
           size=prompt_size, bold=True, color=INK)
        for j, opt in enumerate(options):
            left = Inches(6.55 + j * 2.05)
            add_round(slide, left, top + Inches(0.15), Inches(1.9), Inches(0.5),
                      light)
            tb(slide, left, top + Inches(0.22), Inches(1.9), Inches(0.38), opt,
               size=14, bold=True, color=INK, align=PP_ALIGN.CENTER)


def vocab_cards(slide, items, accent, light):
    """Four vocabulary cards: word, picture, meaning, sentence."""
    for i, (word, emoji, meaning, sentence) in enumerate(items):
        left = Inches(0.5 + i * 3.14)
        add_round(slide, left, Inches(1.78), Inches(2.95), Inches(4.45), light)
        tb(slide, left, Inches(1.95), Inches(2.95), Inches(0.75), emoji, size=30,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(2.78), Inches(2.65),
                  Inches(0.62), accent)
        tb(slide, left + Inches(0.15), Inches(2.9), Inches(2.65), Inches(0.44), word,
           size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(3.55), Inches(2.45), Inches(1.0),
           meaning, size=12, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(4.68), Inches(2.65),
                  Inches(0.95), WHITE)
        tb(slide, left + Inches(0.3), Inches(4.8), Inches(2.35), Inches(0.75),
           sentence, size=11, color=accent, italic=True, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(5.74), Inches(2.65),
                  Inches(0.44), WHITE)
        tb(slide, left + Inches(0.25), Inches(5.82), Inches(2.45), Inches(0.32),
           "your turn:  ____________", size=10, color=SOFT, align=PP_ALIGN.CENTER)


def fix_rows(slide, items, start_index, accent, light, gap=1.5, height=1.34,
             write_line=False):
    """Editing rows: the broken sentence plus chips naming what to hunt for."""
    tops = numbered_rows(slide, len(items), start_index, accent, top_start=1.85,
                         gap=gap, height=height)
    for item, top in zip(items, tops):
        wrong, chips = item[0], item[-1]
        add_round(slide, Inches(1.4), top + Inches(0.18), Inches(8.1), Inches(0.6),
                  light)
        tb(slide, Inches(1.65), top + Inches(0.27), Inches(7.7), Inches(0.44), wrong,
           size=17, bold=True, color=accent, italic=True)
        tb(slide, Inches(1.45), top + Inches(0.86), Inches(8.0), Inches(0.34),
           "Find every mistake first, then read the fixed sentence out loud.",
           size=11, color=SOFT)
        for j, chip in enumerate(chips[:4]):
            left = Inches(9.75 + (j % 2) * 1.45)
            t = top + Inches(0.22 + (j // 2) * 0.46)
            add_round(slide, left, t, Inches(1.35), Inches(0.38), WHITE)
            tb(slide, left, t + Inches(0.03), Inches(1.35), Inches(0.3), chip,
               size=10, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
        if write_line:
            add_round(slide, Inches(1.4), top + Inches(1.24), Inches(11.05),
                      Inches(0.52), L_GREY)
            tb(slide, Inches(1.65), top + Inches(1.34), Inches(10.6), Inches(0.38),
               "Now write it correctly:  ______________________________________"
               "______________________", size=12, color=SOFT)


def evidence_rows(slide, items, start_index):
    """A comprehension question plus a space to prove it from the text."""
    tops = numbered_rows(slide, len(items), start_index, BLUE, top_start=1.85,
                         gap=2.2, height=1.95)
    for (emoji, question, where), top in zip(items, tops):
        tb(slide, Inches(1.45), top + Inches(0.55), Inches(1.0), Inches(0.8), emoji,
           size=32, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.65), top + Inches(0.3), Inches(9.8), Inches(0.62),
           question, size=21, bold=True, color=INK)
        add_round(slide, Inches(2.65), top + Inches(1.0), Inches(4.7), Inches(0.62),
                  L_GREY)
        tb(slide, Inches(2.9), top + Inches(1.13), Inches(4.3), Inches(0.42),
           "Your answer:  ____________________", size=12, color=SOFT)
        add_round(slide, Inches(7.6), top + Inches(1.0), Inches(4.85), Inches(0.62),
                  L_BLUE)
        tb(slide, Inches(7.85), top + Inches(1.11), Inches(4.4), Inches(0.44),
           f"📰 SHOW ME WHERE  ·  look in {where}", size=12, bold=True, color=BLUE)


def story_slide(part, timing):
    """One passage page: two paragraphs plus a single check question."""
    label, heading, emoji, color, light = part[0], part[1], part[2], part[3], part[4]
    paras, check = part[5], part[6]
    slide, n = new_slide(f"📖 {heading}", f"STORY · {label.upper()}", timing,
                         "Reading Desk", color)
    tb(slide, Inches(0.45), Inches(1.3), Inches(11.5), Inches(0.36),
       "STUDENTS TURN AN EMPTY SPACE INTO A GARDEN", size=14, bold=True, color=SOFT)
    add_round(slide, Inches(0.5), Inches(1.78), Inches(2.5), Inches(3.9), light)
    tb(slide, Inches(0.5), Inches(2.45), Inches(2.5), Inches(1.5), emoji, size=76,
       align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.75), Inches(4.25), Inches(2.0), Inches(0.55), WHITE)
    tb(slide, Inches(0.75), Inches(4.36), Inches(2.0), Inches(0.4), label.upper(),
       size=14, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.25), Inches(1.78), Inches(9.6), Inches(3.9), WHITE)
    paragraph(slide, Inches(3.6), Inches(2.02), Inches(8.95), Inches(1.15), paras[0],
              size=15)
    paragraph(slide, Inches(3.6), Inches(3.3), Inches(8.95), Inches(1.15), paras[1],
              size=15)
    add_round(slide, Inches(0.5), Inches(5.82), Inches(12.35), Inches(0.5), light)
    tb(slide, Inches(0.8), Inches(5.9), Inches(2.2), Inches(0.36), "✅ QUICK CHECK",
       size=12, bold=True, color=color)
    tb(slide, Inches(3.4), Inches(5.89), Inches(9.2), Inches(0.38), check, size=14,
       bold=True, color=INK)
    model_bar(slide, 6.45)
    return slide, n


# ------------------------------------------------------------------ slides


def s01_title():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), RED)
    add_round(slide, Inches(4.6), Inches(0.72), Inches(4.1), Inches(0.5), RED)
    tb(slide, Inches(4.6), Inches(0.8), Inches(4.1), Inches(0.36),
       "🚨  BREAKING NEWS  🚨", size=14, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.5), Inches(12), Inches(1.0),
       "WELCOME, YOUNG REPORTER!", size=44, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.7), Inches(2.55), Inches(12), Inches(0.5),
       "Can you create today's school newspaper?", size=21,
       color=RGBColor(0xC6, 0xCD, 0xD6), align=PP_ALIGN.CENTER, italic=True)
    for i, (emoji, label) in enumerate(DESK_KIT):
        left = Inches(1.15 + i * 1.85)
        add_round(slide, left, Inches(3.4), Inches(1.65), Inches(1.35),
                  RGBColor(0x1E, 0x23, 0x2A))
        tb(slide, left, Inches(3.58), Inches(1.65), Inches(0.6), emoji, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.23), Inches(1.45), Inches(0.4), label,
           size=10, color=RGBColor(0xA8, 0xB2, 0xBD), align=PP_ALIGN.CENTER)
    add_round(slide, Inches(3.3), Inches(5.1), Inches(6.7), Inches(1.15), RED)
    tb(slide, Inches(3.5), Inches(5.35), Inches(6.3), Inches(0.7),
       "Grade 5  •  90 Minutes  •  Reading & Grammar", size=19, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "0–7 min", "Headline Desk")
    fade(slide)


def s02_missions():
    slide, n = new_slide("🗞️ Today's Newsroom Missions", "BRIEFING", "0–7 min",
                         "Headline Desk", RED)
    one_task(slide, "Five missions. Finish all five and the edition goes to print.",
             RED)
    for i, (emoji, num, name, detail, color) in enumerate(MISSIONS):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.9), Inches(2.3), Inches(3.6), L_GREY)
        tb(slide, left, Inches(2.12), Inches(2.3), Inches(0.8), emoji, size=32,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.5), Inches(2.95), Inches(1.3),
                  Inches(0.38), color)
        tb(slide, left + Inches(0.5), Inches(2.99), Inches(1.3), Inches(0.3), num,
           size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(3.48), Inches(2.1), Inches(0.9), name,
           size=15, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.18), Inches(4.42), Inches(1.95), Inches(0.9),
           detail, size=11, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.7), Inches(12.35), Inches(0.9), L_RED)
    tb(slide, Inches(0.8), Inches(5.95), Inches(11.7), Inches(0.44),
       "🗞️  →  🔎  →  ✏️  →  📖  →  🏆    Finish all five and you are a Young "
       "Reporter.", size=16, bold=True, color=INK, align=PP_ALIGN.CENTER)


def s03_good_story():
    slide, n = new_slide("❓ What Makes a Good News Story?", "DISCUSS", "0–7 min",
                         "Headline Desk", BLUE)
    one_task(slide, "Read all four. Only one answer is complete.", BLUE)
    for i, (letter, text) in enumerate(GOOD_STORY):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.96), WHITE)
        add_oval(slide, Inches(0.85), top + Inches(0.2), Inches(0.56), Inches(0.56),
                 BLUE)
        tb(slide, Inches(0.85), top + Inches(0.27), Inches(0.56), Inches(0.42),
           letter, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.75), top + Inches(0.24), Inches(10.5), Inches(0.5), text,
           size=21, bold=True, color=INK)
    add_round(slide, Inches(0.5), Inches(6.4), Inches(12.35), Inches(0.5), L_BLUE)
    tb(slide, Inches(0.8), Inches(6.48), Inches(11.7), Inches(0.36),
       "Ask her to defend her choice out loud before you confirm anything.",
       size=13, bold=True, color=BLUE)


def s04_warmup():
    slide, n = new_slide("📷 Warm-Up — What's the News?", "WARM-UP", "0–7 min",
                         "Headline Desk", AMBER)
    one_task(slide, "No reading yet. Just look and tell me the story.", AMBER)
    add_round(slide, Inches(0.5), Inches(1.78), Inches(5.3), Inches(4.5), L_AMBER)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.3), Inches(1.8), "🎃", size=96,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.65), Inches(4.9), Inches(0.9),
       "The school garden has grown a giant pumpkin overnight.", size=15,
       bold=True, color=INK, align=PP_ALIGN.CENTER)
    questions = [("🤔", "What do you think happened?"),
                 ("🗞️", "What would your headline be?")]
    for i, (icon, q) in enumerate(questions):
        top = Inches(1.85 + i * 1.05)
        add_round(slide, Inches(6.15), top, Inches(6.7), Inches(0.9), L_RED)
        tb(slide, Inches(6.45), top + Inches(0.2), Inches(0.5), Inches(0.5), icon,
           size=16, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.15), top + Inches(0.2), Inches(5.4), Inches(0.5), q,
           size=17, bold=True, color=INK)
    for i, starter in enumerate(WARM_STARTERS):
        top = Inches(4.05 + i * 0.78)
        add_round(slide, Inches(6.15), top, Inches(6.7), Inches(0.66), WHITE)
        tb(slide, Inches(6.5), top + Inches(0.12), Inches(6.0), Inches(0.46),
           starter, size=20, bold=True, color=AMBER)
    hint(slide, "Accept any idea. This is the one task today with no wrong answer.",
         6.45)


def s05_method():
    slide, n = new_slide("🧭 How the Newsroom Works", "PROMISE", "0–7 min",
                         "Headline Desk", GREEN)
    one_task(slide, "Every skill today goes through these five steps.", GREEN)
    for i, (num, label, detail, color) in enumerate(METHOD):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.9), Inches(2.3), Inches(2.9), L_GREY)
        add_oval(slide, left + Inches(0.85), Inches(2.12), Inches(0.6), Inches(0.6),
                 color)
        tb(slide, left + Inches(0.85), Inches(2.22), Inches(0.6), Inches(0.42), num,
           size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(2.9), Inches(2.1), Inches(0.75), label,
           size=13, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.18), Inches(3.7), Inches(1.95), Inches(0.9),
           detail, size=11, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.0), Inches(12.35), Inches(1.25), L_AMBER)
    tb(slide, Inches(0.8), Inches(5.14), Inches(11.7), Inches(0.36),
       "🟡 THE REPORTER HINT PROMISE", size=13, bold=True, color=AMBER)
    tb(slide, Inches(0.8), Inches(5.56), Inches(11.7), Inches(0.6),
       "You will never be left stuck. A hint always comes before the answer — and "
       "the answer is always yours to say.", size=16, bold=True, color=INK)
    hint(slide, "Read the promise out loud. It changes how she handles the first "
                "mistake.", 6.45)


def s06_what_headline():
    slide, n = new_slide("🗞️ What Is a Headline?", "LEARN", "7–17 min",
                         "Headline Desk", RED)
    one_task(slide, "A headline gives the main idea in a few words.", RED)
    add_round(slide, Inches(0.5), Inches(1.8), Inches(12.35), Inches(1.5), INK)
    tb(slide, Inches(0.8), Inches(2.12), Inches(11.7), Inches(0.9),
       "SCHOOL GARDEN GROWS GIANT PUMPKIN!", size=36, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    facts = [("📏", "SHORT", "Usually under eight words."),
             ("💡", "MAIN IDEA", "It tells you what the story is about."),
             ("🔠", "CAPITALS", "Often printed in all capital letters."),
             ("⚡", "STRONG VERBS", "GROWS, WINS, STARTS, VISITS.")]
    for i, (icon, label, detail) in enumerate(facts):
        left = Inches(0.5 + i * 3.14)
        add_round(slide, left, Inches(3.55), Inches(2.95), Inches(2.1), L_RED)
        tb(slide, left, Inches(3.72), Inches(2.95), Inches(0.6), icon, size=22,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(4.35), Inches(2.75), Inches(0.5), label,
           size=15, bold=True, color=RED, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(4.88), Inches(2.45), Inches(0.7),
           detail, size=11, color=DARK, align=PP_ALIGN.CENTER)
    hint(slide, "Cover the headline and ask what the story is about. Then reveal it.",
         6.42)


def s07_four_headlines():
    slide, n = new_slide("📰 Four Headlines — What Is Each About?", "READ",
                         "7–17 min", "Headline Desk", RED)
    one_task(slide, "Read each one out loud. Then say the topic in one word.", RED)
    for i, (emoji, headline, topic) in enumerate(HEADLINE_TOPICS):
        top = Inches(1.85 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.96), WHITE)
        tb(slide, Inches(0.85), top + Inches(0.2), Inches(0.7), Inches(0.56), emoji,
           size=22, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.8), top + Inches(0.22), Inches(7.6), Inches(0.5),
           headline, size=20, bold=True, color=INK, font="Georgia")
        add_round(slide, Inches(9.6), top + Inches(0.24), Inches(2.9), Inches(0.5),
                  L_RED)
        tb(slide, Inches(9.6), top + Inches(0.33), Inches(2.9), Inches(0.36),
           f"topic: {topic}", size=12, bold=True, color=RED, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(6.4), Inches(12.35), Inches(0.5), L_GREY)
    tb(slide, Inches(0.8), Inches(6.48), Inches(11.7), Inches(0.36),
       "Cover the topic box first. Let her name the topic before she sees it.",
       size=13, bold=True, color=SOFT)


def s08_match_a():
    slide, n = new_slide("📰 Game: Headline Match", "GAME", "7–17 min",
                         "Headline Desk", RED)
    one_task(slide, "Which headline belongs to this picture?", RED)
    match_rows(slide, MATCH_A, 1, RED, L_RED)
    hint(slide, "Find one word in the headline that matches the picture.", 6.42)


def s09_match_b():
    slide, n = new_slide("📰 Headline Match — Rounds 4 to 6", "GAME", "7–17 min",
                         "Headline Desk", RED)
    one_task(slide, "Three more. Each time, tell me which word helped you.", RED)
    match_rows(slide, MATCH_B, 4, RED, L_RED)
    hint(slide, "Two headlines are close. Read both again before choosing.", 6.42)


def s10_mixup():
    slide, n = new_slide("🧩 Game: Headline Mix-Up", "GAME", "7–17 min",
                         "Headline Desk", PURPLE)
    one_task(slide, "The words fell off the page. Put the headline back together.",
             PURPLE)
    tops = numbered_rows(slide, len(MIXUP), 1, PURPLE, top_start=1.85, gap=1.55,
                         height=1.4)
    for (cards, _answer), top in zip(MIXUP, tops):
        for j, card in enumerate(cards):
            left = Inches(1.45 + j * 2.25)
            add_round(slide, left, top + Inches(0.2), Inches(2.1), Inches(0.58),
                      L_PURPLE)
            tb(slide, left, top + Inches(0.3), Inches(2.1), Inches(0.42), card,
               size=15, bold=True, color=PURPLE, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(1.45), top + Inches(0.86), Inches(10.95),
                  Inches(0.44), L_GREY)
        tb(slide, Inches(1.7), top + Inches(0.93), Inches(10.5), Inches(0.34),
           "Your headline:  ________________________________________________",
           size=12, color=SOFT)
    hint(slide, "Start with the verb. A headline is built around its action word.",
         6.45)


def s11_mission1_done():
    slide, n = new_slide("⭐ Mission 1 Complete — Headline Hunter", "BADGE",
                         "7–17 min", "Headline Desk", GREEN)
    one_task(slide, "You can find the main idea in seconds. That is a real skill.",
             GREEN)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.3), Inches(4.4), L_GREEN)
    tb(slide, Inches(0.5), Inches(2.6), Inches(5.3), Inches(1.7), "🗞️", size=90,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.3), Inches(0.7), "BADGE 1 EARNED",
       size=26, bold=True, color=GREEN, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.7), Inches(0.6), "Headline Hunter",
       size=15, color=SOFT, align=PP_ALIGN.CENTER)
    learned = [("💡", "A headline gives the MAIN IDEA."),
               ("📏", "Short words, strong verbs, capital letters."),
               ("🔎", "One word in the headline usually gives it away."),
               ("🧩", "You can rebuild a headline from scattered words.")]
    for i, (icon, line) in enumerate(learned):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.15), top, Inches(6.7), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.45), top + Inches(0.2), Inches(0.56), Inches(0.56),
                 WHITE)
        tb(slide, Inches(6.45), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.25), top + Inches(0.24), Inches(5.3), Inches(0.5), line,
           size=15, bold=True, color=INK)
    hint(slide, "Next up: the twelve words that run this newsroom.", 6.45,
         "➡️ COMING UP", L_GREEN, GREEN)


def s12_vocab_1():
    slide, n = new_slide("🔎 Newsroom Words — Set 1 of 3", "VOCABULARY", "17–27 min",
                         "Word Desk", PURPLE)
    one_task(slide, "Four words only. Say it, mean it, then use it.", PURPLE)
    vocab_cards(slide, VOCAB_1, PURPLE, L_PURPLE)
    hint(slide, "Say the word, she repeats it, then she reads the sentence.", 6.42)


def s13_vocab_2():
    slide, n = new_slide("🔎 Newsroom Words — Set 2 of 3", "VOCABULARY", "17–27 min",
                         "Word Desk", PURPLE)
    one_task(slide, "Four more. Two of these appear in today's news story.", PURPLE)
    vocab_cards(slide, VOCAB_2, PURPLE, L_PURPLE)
    hint(slide, "CAREFULLY ends in -ly, so it describes HOW something is done.",
         6.42)


def s14_vocab_3():
    slide, n = new_slide("🔎 Newsroom Words — Set 3 of 3", "VOCABULARY", "17–27 min",
                         "Word Desk", PURPLE)
    one_task(slide, "The last four. Now you own all twelve.", PURPLE)
    vocab_cards(slide, VOCAB_3, PURPLE, L_PURPLE)
    hint(slide, "Go back to Set 1 for twenty seconds before you move on.", 6.42)


def s15_word_or_not():
    slide, n = new_slide("🎯 Game: Word or Not?", "GAME", "17–27 min", "Word Desk",
                         TEAL)
    one_task(slide, "Which of these words belong in our newsroom?", TEAL)
    for i, (word, _belongs) in enumerate(WORD_OR_NOT):
        col, row = i % 4, i // 4
        left = Inches(0.5 + col * 3.14)
        top = Inches(1.95 + row * 1.85)
        add_round(slide, left, top, Inches(2.95), Inches(1.6), L_TEAL)
        tb(slide, left, top + Inches(0.28), Inches(2.95), Inches(0.6), word,
           size=22, bold=True, color=INK, align=PP_ALIGN.CENTER, font="Georgia")
        add_round(slide, left + Inches(0.6), top + Inches(0.98), Inches(1.75),
                  Inches(0.44), WHITE)
        tb(slide, left + Inches(0.6), top + Inches(1.05), Inches(1.75), Inches(0.34),
           "✓  YES    ✗  NO", size=12, bold=True, color=SOFT,
           align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.6), Inches(12.35), Inches(0.66), L_AMBER)
    tb(slide, Inches(0.8), Inches(5.75), Inches(11.7), Inches(0.4),
       "⭐ EDITOR CHALLENGE:  Read every YES word out loud, then use two of them in "
       "one sentence.", size=14, bold=True, color=INK)
    hint(slide, "Ask \"would a reporter write this word today?\" rather than \"is it "
                "real?\"", 6.45)


def s16_who_am_i():
    slide, n = new_slide("🎤 Game: Who Am I?", "GAME", "17–27 min", "Word Desk",
                         AMBER)
    one_task(slide, "Three clues each. Guess before the third clue if you can.",
             AMBER)
    for i, (emoji, clues, _answer) in enumerate(WHO_AM_I):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(4.4), L_AMBER)
        tb(slide, left, Inches(2.02), Inches(3.9), Inches(0.72), emoji, size=28,
           align=PP_ALIGN.CENTER)
        for j, clue in enumerate(clues):
            top = Inches(2.82 + j * 0.92)
            add_round(slide, left + Inches(0.22), top, Inches(3.45), Inches(0.8),
                      WHITE)
            tb(slide, left + Inches(0.35), top + Inches(0.22), Inches(0.35),
               Inches(0.36), f"{j + 1}.", size=11, bold=True, color=AMBER)
            tb(slide, left + Inches(0.8), top + Inches(0.18), Inches(2.75),
               Inches(0.5), clue, size=11, color=DARK)
        add_round(slide, left + Inches(0.7), Inches(5.66), Inches(2.5),
                  Inches(0.44), WHITE)
        tb(slide, left + Inches(0.7), Inches(5.73), Inches(2.5), Inches(0.34),
           "I am a ____________", size=11, color=SOFT, align=PP_ALIGN.CENTER)
    hint(slide, "Guessed after clue one? Read the other two anyway and confirm.",
         6.45)


def s17_context():
    slide, n = new_slide("🧩 Which Word Fits?", "PRACTICE", "17–27 min", "Word Desk",
                         PURPLE)
    one_task(slide, "Read the whole sentence first. The sentence tells you.", PURPLE)
    items = [(prompt, options) for prompt, options in CONTEXT]
    tops = numbered_rows(slide, len(items), 1, PURPLE, top_start=1.9, gap=1.5)
    for (prompt, options), top in zip(items, tops):
        tb(slide, Inches(1.45), top + Inches(0.24), Inches(10.9), Inches(0.5),
           prompt, size=19, bold=True, color=INK)
        for j, opt in enumerate(options):
            left = Inches(1.45 + j * 3.2)
            add_round(slide, left, top + Inches(0.8), Inches(2.95), Inches(0.44),
                      L_PURPLE)
            tb(slide, left, top + Inches(0.87), Inches(2.95), Inches(0.34), opt,
               size=14, bold=True, color=PURPLE, align=PP_ALIGN.CENTER)
    hint(slide, "Try each word in the blank out loud. The wrong ones sound wrong.",
         6.45)


def s18_subject_predicate():
    slide, n = new_slide("✏️ Subject and Predicate", "LEARN", "27–37 min",
                         "Editor's Desk", GREEN)
    one_task(slide, "Every sentence has a WHO and a WHAT THEY DO.", GREEN)
    add_round(slide, Inches(0.5), Inches(1.8), Inches(12.35), Inches(1.3), WHITE)
    add_round(slide, Inches(0.85), Inches(2.0), Inches(4.4), Inches(0.9), L_BLUE)
    tb(slide, Inches(0.85), Inches(2.2), Inches(4.4), Inches(0.55),
       "The students", size=28, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(5.6), Inches(2.0), Inches(6.9), Inches(0.9), L_GREEN)
    tb(slide, Inches(5.6), Inches(2.2), Inches(6.9), Inches(0.55),
       "practice every morning.", size=28, bold=True, color=GREEN,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.85), Inches(3.18), Inches(4.4), Inches(0.34),
       "SUBJECT — who or what", size=12, bold=True, color=BLUE,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(5.6), Inches(3.18), Inches(6.9), Inches(0.34),
       "PREDICATE — what they do", size=12, bold=True, color=GREEN,
       align=PP_ALIGN.CENTER)
    for i, (subj, pred) in enumerate(SUBJ_PRED[1:]):
        top = Inches(3.75 + i * 0.85)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.72), L_GREY)
        add_round(slide, Inches(0.85), top + Inches(0.1), Inches(4.4), Inches(0.52),
                  L_BLUE)
        tb(slide, Inches(0.85), top + Inches(0.18), Inches(4.4), Inches(0.4), subj,
           size=17, bold=True, color=BLUE, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(5.6), top + Inches(0.1), Inches(6.9), Inches(0.52),
                  L_GREEN)
        tb(slide, Inches(5.6), top + Inches(0.18), Inches(6.9), Inches(0.4), pred,
           size=17, bold=True, color=GREEN, align=PP_ALIGN.CENTER)
    hint(slide, "Ask \"who is doing it?\" first. That answer is always the subject.",
         6.45)


def s19_agreement_learn():
    slide, n = new_slide("🔗 Subject-Verb Agreement", "LEARN", "27–37 min",
                         "Editor's Desk", GREEN)
    one_task(slide, "One person adds an S to the verb. Two or more do not.", GREEN)
    pairs = [("The student", "writes", "1 person → verb gets S", BLUE, L_BLUE),
             ("The students", "write", "2+ people → no S", GREEN, L_GREEN),
             ("The reporter", "asks", "1 person → verb gets S", BLUE, L_BLUE),
             ("The reporters", "ask", "2+ people → no S", GREEN, L_GREEN)]
    for i, (subj, verb, rule, color, light) in enumerate(pairs):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(1.9 + row * 2.2)
        add_round(slide, left, top, Inches(5.95), Inches(1.95), light)
        tb(slide, left + Inches(0.35), top + Inches(0.3), Inches(5.25), Inches(0.62),
           f"{subj} ", size=26, bold=True, color=INK)
        tb(slide, left + Inches(0.35), top + Inches(0.95), Inches(2.8),
           Inches(0.55), verb.upper(), size=30, bold=True, color=color,
           font="Georgia")
        add_round(slide, left + Inches(3.3), top + Inches(1.0), Inches(2.3),
                  Inches(0.46), WHITE)
        tb(slide, left + Inches(3.3), top + Inches(1.08), Inches(2.3), Inches(0.34),
           rule, size=10, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Count the people first. The count decides the verb, nothing else.",
         6.45)


def s20_agreement_practice():
    slide, n = new_slide("🔗 Choose the Right Verb", "PRACTICE", "27–37 min",
                         "Editor's Desk", GREEN)
    one_task(slide, "Count the people. Then pick the verb.", GREEN)
    items = [(prompt, options) for prompt, options, _a in AGREE_PRACTICE]
    choice_rows(slide, items, 1, GREEN, L_GREEN)
    hint(slide, "Butterflies is more than one, even though it has no THE in front.",
         6.42)


def s21_red_pen_a():
    slide, n = new_slide("🖊️ Game: Editor's Red Pen", "GAME", "27–37 min",
                         "Editor's Desk", RED)
    one_task(slide, "One mistake per sentence. Find it before you fix it.", RED)
    items = [(wrong, ["verb"]) for wrong, _fix in RED_PEN_A]
    fix_rows(slide, items, 1, RED, L_RED)
    hint(slide, "Point at the subject and ask: is that one person, or more?", 6.42)


def s22_red_pen_b():
    slide, n = new_slide("🖊️ Editor's Red Pen — Sentences 4 and 5", "GAME",
                         "27–37 min", "Editor's Desk", RED)
    one_task(slide, "Two harder ones. Say the rule out loud as you fix it.", RED)
    items = [(wrong, ["verb"]) for wrong, _fix in RED_PEN_B]
    fix_rows(slide, items, 4, RED, L_RED)
    add_round(slide, Inches(0.5), Inches(4.9), Inches(12.35), Inches(1.3), L_PURPLE)
    tb(slide, Inches(0.8), Inches(5.05), Inches(11.7), Inches(0.36),
       "⭐ EDITOR CHALLENGE", size=13, bold=True, color=PURPLE)
    tb(slide, Inches(0.8), Inches(5.48), Inches(11.7), Inches(0.6),
       "\"Priya and Sam\" is two people. Can you write one more sentence with two "
       "names as the subject?", size=16, bold=True, color=INK)
    hint(slide, "Two names joined by AND always count as more than one.", 6.42)


def s23_pos_key():
    slide, n = new_slide("🎨 The Parts of Speech Colour Code", "LEARN", "37–44 min",
                         "Editor's Desk", BLUE)
    one_task(slide, "We colour every sentence from here on. Learn the four colours.",
             BLUE)
    for i, (name, what, examples, color, light) in enumerate(POS_KEY):
        left = Inches(0.5 + i * 3.14)
        add_round(slide, left, Inches(1.85), Inches(2.95), Inches(2.9), light)
        add_round(slide, left + Inches(0.3), Inches(2.08), Inches(2.35),
                  Inches(0.62), color)
        tb(slide, left + Inches(0.3), Inches(2.2), Inches(2.35), Inches(0.44), name,
           size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.25), Inches(2.9), Inches(2.45), Inches(0.75), what,
           size=12, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(3.75), Inches(2.55),
                  Inches(0.78), WHITE)
        tb(slide, left + Inches(0.3), Inches(3.9), Inches(2.35), Inches(0.58),
           examples, size=11, bold=True, color=color, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.0), Inches(12.35), Inches(1.2), WHITE)
    colored_sentence(slide, Inches(0.9), Inches(5.3), Inches(11.5), Inches(0.7),
                     SPIN_1, size=28, align=PP_ALIGN.CENTER)
    hint(slide, "Adverbs usually end in -ly. That one rule solves most of them.",
         6.42)


def s24_spin_wheel():
    slide, n = new_slide("🎯 Game: Grammar Spin", "GAME", "37–44 min",
                         "Editor's Desk", AMBER)
    one_task(slide, "Pick a category with your eyes closed. Then answer fast.",
             AMBER)
    for i, (icon, name, task, color, light) in enumerate(SPIN_WHEEL):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.85), Inches(2.3), Inches(3.0), light)
        tb(slide, left, Inches(2.05), Inches(2.3), Inches(0.7), icon, size=26,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.15), Inches(2.82), Inches(2.0),
                  Inches(0.55), color)
        tb(slide, left + Inches(0.15), Inches(2.92), Inches(2.0), Inches(0.4), name,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.18), Inches(3.55), Inches(1.95), Inches(0.9),
           task, size=12, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.05), Inches(12.35), Inches(1.2), L_AMBER)
    tb(slide, Inches(0.8), Inches(5.2), Inches(11.7), Inches(0.36),
       "HOW TO SPIN", size=13, bold=True, color=AMBER)
    tb(slide, Inches(0.8), Inches(5.62), Inches(11.7), Inches(0.5),
       "Close your eyes, point at the screen, open them. Whatever you landed on is "
       "your challenge. Keep each turn under twenty seconds.", size=15, color=INK)
    hint(slide, "Speed matters more than perfection here. Keep the pace up.", 6.45)


def s25_spin_1():
    slide, n = new_slide("🎯 Grammar Spin — Sentence 1", "GAME", "37–44 min",
                         "Editor's Desk", AMBER)
    pos_legend(slide)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(12.35), Inches(1.25), WHITE)
    colored_sentence(slide, Inches(0.9), Inches(2.15), Inches(11.5), Inches(0.7),
                     SPIN_1, size=32, align=PP_ALIGN.CENTER)
    for i, (icon, task, note) in enumerate(SPIN_TASKS_1):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(3.35 + row * 1.35)
        add_round(slide, left, top, Inches(5.95), Inches(1.15), L_GREY)
        tb(slide, left + Inches(0.3), top + Inches(0.28), Inches(0.6), Inches(0.56),
           icon, size=17, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.1), top + Inches(0.2), Inches(4.6), Inches(0.44),
           task, size=18, bold=True, color=INK)
        tb(slide, left + Inches(1.1), top + Inches(0.68), Inches(4.6), Inches(0.36),
           note, size=12, color=SOFT)
    hint(slide, "Colours are the answer key. Cover them with a sticky note first.",
         6.42)


def s26_spin_2():
    slide, n = new_slide("🎯 Grammar Spin — Sentences 2 and 3", "GAME", "37–44 min",
                         "Editor's Desk", AMBER)
    pos_legend(slide)
    for i, sentence in enumerate([SPIN_2, SPIN_3]):
        top = Inches(1.9 + i * 1.45)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.25), WHITE)
        colored_sentence(slide, Inches(0.9), top + Inches(0.3), Inches(11.5),
                         Inches(0.7), sentence, size=28, align=PP_ALIGN.CENTER)
    tasks = [("🔵", "Two nouns in each sentence — name them both."),
             ("🟣", "Sentence 3 has two adjectives. Find both."),
             ("🟡", "Both adverbs end in -ly."),
             ("⭐", "Swap one adjective for your own. Read it again.")]
    for i, (icon, task) in enumerate(tasks):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(4.85 + row * 0.78)
        add_round(slide, left, top, Inches(5.95), Inches(0.66), L_GREY)
        tb(slide, left + Inches(0.28), top + Inches(0.14), Inches(0.5), Inches(0.42),
           icon, size=14, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.95), top + Inches(0.13), Inches(4.85),
           Inches(0.42), task, size=14, bold=True, color=INK)
    hint(slide, "If she stalls, ask which word tells you HOW the action was done.",
         6.45)


def s27_spin_fix():
    slide, n = new_slide("🔴 Grammar Spin — Fix the Sentence", "GAME", "37–44 min",
                         "Editor's Desk", RED)
    one_task(slide, "Red means repair. Say what is wrong before you fix it.", RED)
    items = [(wrong, ["capital", "verb", "period"]) for wrong, _fix in FIX_SPIN]
    fix_rows(slide, items, 1, RED, L_RED)
    hint(slide, "Three things to check every time: capital, verb, end mark.", 6.42)


def s28_conjunctions():
    slide, n = new_slide("🔗 Joining Words — and, but, so, because", "LEARN",
                         "37–44 min", "Editor's Desk", TEAL)
    one_task(slide, "These four words glue two ideas into one sentence.", TEAL)
    for i, (word, role, example, color) in enumerate(CONJUNCTIONS):
        top = Inches(1.85 + i * 1.12)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.96), WHITE)
        add_round(slide, Inches(0.85), top + Inches(0.2), Inches(1.7), Inches(0.56),
                  color)
        tb(slide, Inches(0.85), top + Inches(0.29), Inches(1.7), Inches(0.42), word,
           size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.8), top + Inches(0.29), Inches(2.3), Inches(0.42), role,
           size=13, color=SOFT)
        tb(slide, Inches(5.3), top + Inches(0.24), Inches(7.2), Inches(0.5),
           example, size=16, bold=True, color=INK)
    hint(slide, "Read the sentence with each joining word. Only one makes sense.",
         6.45)


def s29_break_intro():
    slide, n = new_slide("🧠 Brain Break — Newsroom Actions", "BREAK", "44–49 min",
                         "Editor's Desk", TEAL)
    one_task(slide, "Stay by your chair. Act out whatever I call.", TEAL)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.4), Inches(4.4), L_TEAL)
    tb(slide, Inches(0.5), Inches(2.5), Inches(5.4), Inches(1.8), "🎬", size=90,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.4), Inches(0.7), "5 MINUTES",
       size=30, bold=True, color=TEAL, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.8), Inches(0.6),
       "then straight back to the news", size=14, color=SOFT,
       align=PP_ALIGN.CENTER)
    rules = [("🪑", "Stay beside your chair the whole time."),
             ("⚡", "React fast — that is the whole game."),
             ("🔁", "The last two rounds, you call and I act."),
             ("🚨", "\"Breaking News\" is always a whisper.")]
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


def s30_break_actions():
    slide, n = new_slide("🎬 Newsroom Actions — Six Calls", "BREAK", "44–49 min",
                         "Editor's Desk", TEAL)
    one_task(slide, "Listen for the call. Then act it out.", TEAL)
    for i, (icon, call, action) in enumerate(NEWSROOM_ACTIONS):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.9 + row * 2.3)
        add_round(slide, left, top, Inches(3.9), Inches(2.1), L_TEAL)
        tb(slide, left, top + Inches(0.18), Inches(3.9), Inches(0.7), icon, size=28,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), top + Inches(0.92), Inches(3.6),
           Inches(0.5), call, size=20, bold=True, color=TEAL,
           align=PP_ALIGN.CENTER, font="Georgia")
        tb(slide, left + Inches(0.3), top + Inches(1.46), Inches(3.3), Inches(0.5),
           action, size=12, color=DARK, align=PP_ALIGN.CENTER)
    hint(slide, "Mix in a reading call: \"Headline!\" then point at a word to read.",
         6.45)


def s31_timeline():
    slide, n = new_slide("⏰ News Through Time — Past, Now, Future", "LEARN",
                         "49–59 min", "Editor's Desk", BLUE)
    one_task(slide, "Same reporter. Same teacher. Three different times.", BLUE)
    for i, (label, when, sentence, color, light) in enumerate(TENSE_LINE):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(3.5), light)
        add_round(slide, left + Inches(1.05), Inches(2.08), Inches(1.8),
                  Inches(0.6), color)
        tb(slide, left + Inches(1.05), Inches(2.2), Inches(1.8), Inches(0.42),
           label, size=15, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(2.82), Inches(3.5), Inches(0.4), when,
           size=13, color=SOFT, align=PP_ALIGN.CENTER, italic=True)
        tb(slide, left + Inches(0.3), Inches(3.35), Inches(3.3), Inches(1.6),
           sentence, size=17, bold=True, color=INK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.55), Inches(12.35), Inches(0.72), INK)
    tb(slide, Inches(0.8), Inches(5.72), Inches(11.7), Inches(0.44),
       "PAST   ←———————   NOW   ———————→   FUTURE", size=18, bold=True,
       color=WHITE, align=PP_ALIGN.CENTER, font="Georgia")
    hint(slide, "Ask when it happened before asking which verb to use.", 6.45)


def s32_tense_chart():
    slide, n = new_slide("📋 The Reporter's Verb Chart", "LEARN", "49–59 min",
                         "Editor's Desk", BLUE)
    one_task(slide, "Five newsroom verbs in all three times.", BLUE)
    cols = [(0.5, 3.0, "VERB", SLATE), (3.7, 3.0, "NOW", GREEN),
            (6.9, 3.0, "PAST", RED), (10.1, 2.75, "FUTURE", BLUE)]
    for left, width, label, color in cols:
        add_round(slide, Inches(left), Inches(1.88), Inches(width), Inches(0.5),
                  color)
        tb(slide, Inches(left), Inches(1.98), Inches(width), Inches(0.36), label,
           size=13, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    for i, row in enumerate(TENSE_CHART):
        top = Inches(2.52 + i * 0.78)
        band = WHITE if i % 2 == 0 else L_GREY
        for (left, width, _l, color), value in zip(cols, row):
            add_round(slide, Inches(left), top, Inches(width), Inches(0.68), band)
            tb(slide, Inches(left), top + Inches(0.14), Inches(width), Inches(0.44),
               value, size=17, bold=True,
               color=INK if color == SLATE else color, align=PP_ALIGN.CENTER)
    hint(slide, "WIN → WON breaks the -ed rule. Flag it now so it is not a shock.",
         6.45)


def s33_name_tense():
    slide, n = new_slide("⏰ Game: Time Machine Reporter", "GAME", "49–59 min",
                         "Editor's Desk", BLUE)
    one_task(slide, "Read the sentence. Name the tense. Twenty seconds each.", BLUE)
    items = [(sentence, ["PAST", "PRESENT", "FUTURE"])
             for sentence, _tense in NAME_TENSE]
    choice_rows(slide, items, 1, BLUE, L_BLUE, prompt_size=16)
    hint(slide, "Hunt for a time word first: yesterday, every Friday, on Friday.",
         6.42)


def s34_transform():
    slide, n = new_slide("⏰ Time Machine — Now Change It", "GAME", "49–59 min",
                         "Editor's Desk", GREEN)
    one_task(slide, "I give you NOW. You give me PAST and FUTURE.", GREEN)
    tops = numbered_rows(slide, len(TRANSFORM), 1, GREEN, top_start=1.85, gap=1.55,
                         height=1.4)
    for (present, _past, _future), top in zip(TRANSFORM, tops):
        add_round(slide, Inches(1.45), top + Inches(0.18), Inches(10.95),
                  Inches(0.52), L_GREEN)
        tb(slide, Inches(1.7), top + Inches(0.25), Inches(10.5), Inches(0.4),
           f"NOW:   {present}", size=16, bold=True, color=GREEN)
        for j, label in enumerate(["PAST:", "FUTURE:"]):
            left = Inches(1.45 + j * 5.55)
            add_round(slide, left, top + Inches(0.8), Inches(5.4), Inches(0.46),
                      L_GREY)
            tb(slide, left + Inches(0.2), top + Inches(0.87), Inches(5.0),
               Inches(0.36), f"{label}  ___________________________", size=12,
               color=SOFT)
    hint(slide, "Say the time word out loud first: \"yesterday, the team...\"", 6.42)


def s35_irregular():
    slide, n = new_slide("⭐ Time Machine Challenge — Tricky Verbs", "BONUS",
                         "49–59 min", "Editor's Desk", PURPLE)
    one_task(slide, "These three do not take -ed. You have to know them.", PURPLE)
    for i, (now, past, future, sentence) in enumerate(IRREGULAR):
        top = Inches(1.9 + i * 1.5)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.34), L_PURPLE)
        for j, (label, value, color) in enumerate([("NOW", now, GREEN),
                                                   ("PAST", past, RED),
                                                   ("FUTURE", future, BLUE)]):
            left = Inches(0.85 + j * 2.6)
            add_round(slide, left, top + Inches(0.2), Inches(2.35), Inches(0.82),
                      WHITE)
            tb(slide, left, top + Inches(0.26), Inches(2.35), Inches(0.28), label,
               size=9, bold=True, color=SOFT, align=PP_ALIGN.CENTER)
            tb(slide, left, top + Inches(0.55), Inches(2.35), Inches(0.42), value,
               size=17, bold=True, color=color, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(8.85), top + Inches(0.2), Inches(3.65),
                  Inches(0.82), WHITE)
        tb(slide, Inches(9.05), top + Inches(0.38), Inches(3.3), Inches(0.5),
           sentence, size=12, bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Fill the blanks with the PAST form. Read the whole sentence after.",
         6.42)


def s36_mission_tense():
    slide, n = new_slide("⭐ Mission 3 Complete — Grammar Editor", "BADGE",
                         "49–59 min", "Editor's Desk", GREEN)
    one_task(slide, "You can now fix a sentence three different ways.", GREEN)
    add_round(slide, Inches(0.5), Inches(1.85), Inches(5.3), Inches(4.4), L_GREEN)
    tb(slide, Inches(0.5), Inches(2.6), Inches(5.3), Inches(1.7), "✏️", size=90,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.5), Inches(4.5), Inches(5.3), Inches(0.7), "BADGE 3 EARNED",
       size=26, bold=True, color=GREEN, align=PP_ALIGN.CENTER, font="Georgia")
    tb(slide, Inches(0.8), Inches(5.3), Inches(4.7), Inches(0.6), "Grammar Editor",
       size=15, color=SOFT, align=PP_ALIGN.CENTER)
    learned = [("🔗", "Subject and predicate in any sentence."),
               ("🔢", "One person adds an S — two or more do not."),
               ("🎨", "Noun, verb, adjective and adverb by colour."),
               ("⏰", "Past, present and future, including won and wrote.")]
    for i, (icon, line) in enumerate(learned):
        top = Inches(1.9 + i * 1.12)
        add_round(slide, Inches(6.15), top, Inches(6.7), Inches(0.96), L_GREY)
        add_oval(slide, Inches(6.45), top + Inches(0.2), Inches(0.56), Inches(0.56),
                 WHITE)
        tb(slide, Inches(6.45), top + Inches(0.26), Inches(0.56), Inches(0.44),
           icon, size=15, align=PP_ALIGN.CENTER)
        tb(slide, Inches(7.25), top + Inches(0.24), Inches(5.3), Inches(0.5), line,
           size=15, bold=True, color=INK)
    hint(slide, "Next: the real news story. Everything above is about to be used.",
         6.45, "➡️ COMING UP", L_GREEN, GREEN)


def s37_story_intro():
    slide, n = new_slide("📖 Today's Front Page Story", "STORY", "59–69 min",
                         "Reading Desk", BLUE)
    one_task(slide, "Six words to watch for. Four short parts. One question each.",
             BLUE)
    for i, (word, emoji) in enumerate(STORY_WATCH):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.85 + row * 1.35)
        add_round(slide, left, top, Inches(3.9), Inches(1.15), L_BLUE)
        tb(slide, left + Inches(0.25), top + Inches(0.26), Inches(0.8),
           Inches(0.64), emoji, size=24, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(1.35), top + Inches(0.3), Inches(2.3),
                  Inches(0.55), WHITE)
        tb(slide, left + Inches(1.35), top + Inches(0.39), Inches(2.3),
           Inches(0.4), word, size=17, bold=True, color=BLUE,
           align=PP_ALIGN.CENTER)
    for i, (step, label, detail, color) in enumerate(READ_STEPS):
        left = Inches(0.5 + i * 3.14)
        add_round(slide, left, Inches(4.7), Inches(2.95), Inches(1.5), L_GREY)
        add_round(slide, left + Inches(0.75), Inches(4.85), Inches(1.45),
                  Inches(0.34), color)
        tb(slide, left + Inches(0.75), Inches(4.88), Inches(1.45), Inches(0.28),
           step, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(5.28), Inches(2.75), Inches(0.4),
           label, size=13, bold=True, color=color, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.2), Inches(5.7), Inches(2.55), Inches(0.42),
           detail, size=10, color=DARK, align=PP_ALIGN.CENTER)
    hint(slide, "Read the six words together now. No cold surprises in the passage.",
         6.42)


def s38_story_1():
    story_slide(STORY[0], "59–69 min")


def s39_story_2():
    story_slide(STORY[1], "59–69 min")


def s40_story_3():
    story_slide(STORY[2], "59–69 min")


def s41_story_4():
    story_slide(STORY[3], "59–69 min")


def s42_main_idea():
    slide, n = new_slide("💡 Main Idea and Order of Events", "COMPREHENSION",
                         "59–69 min", "Reading Desk", PURPLE)
    one_task(slide, "One main idea. Then put the four events in order.", PURPLE)
    tb(slide, Inches(0.5), Inches(1.78), Inches(12.35), Inches(0.36),
       "WHICH SENTENCE IS THE MAIN IDEA OF THE WHOLE STORY?", size=13, bold=True,
       color=PURPLE)
    for i, option in enumerate(MAIN_IDEA):
        top = Inches(2.18 + i * 0.7)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.58), WHITE)
        add_oval(slide, Inches(0.8), top + Inches(0.11), Inches(0.36), Inches(0.36),
                 PURPLE)
        tb(slide, Inches(0.8), top + Inches(0.14), Inches(0.36), Inches(0.3),
           "ABC"[i], size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.5), top + Inches(0.1), Inches(10.7), Inches(0.4), option,
           size=16, bold=True, color=INK)
    tb(slide, Inches(0.5), Inches(4.42), Inches(12.35), Inches(0.36),
       "NOW NUMBER THESE EVENTS 1 TO 4 IN THE ORDER THEY HAPPENED.", size=13,
       bold=True, color=BLUE)
    for i, (emoji, event) in enumerate(SEQUENCE):
        col, row = i % 2, i // 2
        left = Inches(0.5 + col * 6.25)
        top = Inches(4.82 + row * 0.78)
        add_round(slide, left, top, Inches(5.95), Inches(0.66), L_BLUE)
        add_round(slide, left + Inches(0.22), top + Inches(0.13), Inches(0.42),
                  Inches(0.4), WHITE)
        tb(slide, left + Inches(0.78), top + Inches(0.12), Inches(0.5),
           Inches(0.42), emoji, size=14, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.35), top + Inches(0.14), Inches(4.45),
           Inches(0.4), event, size=12, bold=True, color=INK)
    hint(slide, "The main idea covers the WHOLE story, not one interesting detail.",
         6.42)


def s43_evidence_how():
    slide, n = new_slide("🔍 Don't Just Guess — Find the Evidence!", "LEARN",
                         "69–77 min", "Reading Desk", RED)
    one_task(slide, "From now on, every answer needs a sentence to back it up.", RED)
    for i, (num, label, detail, color) in enumerate(EVIDENCE_STEPS):
        left = Inches(0.5 + i * 4.15)
        add_round(slide, left, Inches(1.85), Inches(3.9), Inches(3.1), L_GREY)
        add_oval(slide, left + Inches(1.6), Inches(2.1), Inches(0.7), Inches(0.7),
                 color)
        tb(slide, left + Inches(1.6), Inches(2.22), Inches(0.7), Inches(0.46), num,
           size=18, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(3.0), Inches(3.6), Inches(0.5), label,
           size=19, bold=True, color=color, align=PP_ALIGN.CENTER, font="Georgia")
        tb(slide, left + Inches(0.3), Inches(3.6), Inches(3.3), Inches(0.7), detail,
           size=13, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.15), Inches(12.35), Inches(1.1), INK)
    tb(slide, Inches(0.8), Inches(5.48), Inches(11.7), Inches(0.5),
       "📰  \"SHOW ME WHERE YOU FOUND IT.\"", size=24, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    hint(slide, "Going back to the text is not cheating. It is the actual skill.",
         6.42)


def s44_evidence_a():
    slide, n = new_slide("🔍 Game: Find the Evidence — 1 and 2", "GAME", "69–77 min",
                         "Reading Desk", BLUE)
    one_task(slide, "Answer it. Then find the sentence that proves it.", BLUE)
    evidence_rows(slide, EVIDENCE_A, 1)
    hint(slide, "Part 3 lists four plants. Ask her to read all four out loud.",
         6.42)


def s45_evidence_b():
    slide, n = new_slide("🔍 Find the Evidence — 3 and 4", "GAME", "69–77 min",
                         "Reading Desk", BLUE)
    one_task(slide, "Two more. Read the proof sentence out loud each time.", BLUE)
    evidence_rows(slide, EVIDENCE_B, 3)
    hint(slide, "Question 4 is answered in Priya's own quoted words.", 6.42)


def s46_evidence_c():
    slide, n = new_slide("🔍 Find the Evidence — 5 and an Inference", "GAME",
                         "69–77 min", "Reading Desk", PURPLE)
    one_task(slide, "The last one is not written down. You work it out.", PURPLE)
    evidence_rows(slide, EVIDENCE_C, 5)
    quote, question, options = INFERENCE
    add_round(slide, Inches(0.5), Inches(4.1), Inches(12.35), Inches(2.1), L_PURPLE)
    tb(slide, Inches(0.8), Inches(4.25), Inches(11.7), Inches(0.36),
       "🧠 INFERENCE — the answer is hinted, not stated", size=12, bold=True,
       color=PURPLE)
    tb(slide, Inches(0.8), Inches(4.65), Inches(11.7), Inches(0.4), quote, size=16,
       bold=True, color=INK, italic=True)
    tb(slide, Inches(0.8), Inches(5.12), Inches(11.7), Inches(0.36), question,
       size=14, bold=True, color=PURPLE)
    for i, opt in enumerate(options):
        left = Inches(0.8 + i * 3.95)
        add_round(slide, left, Inches(5.55), Inches(3.75), Inches(0.5), WHITE)
        tb(slide, left, Inches(5.63), Inches(3.75), Inches(0.36), opt, size=12,
           bold=True, color=INK, align=PP_ALIGN.CENTER)
    hint(slide, "Ask: what kind of person comes back when their arms hurt?", 6.42)


def s47_context_story():
    slide, n = new_slide("📖 Word Meaning from the Story", "VOCABULARY", "69–77 min",
                         "Reading Desk", TEAL)
    one_task(slide, "Go back to the sentence. The sentence gives you the meaning.",
             TEAL)
    tops = numbered_rows(slide, len(CONTEXT_STORY), 1, TEAL, top_start=1.9, gap=1.5)
    for (word, where, options), top in zip(CONTEXT_STORY, tops):
        add_round(slide, Inches(1.45), top + Inches(0.22), Inches(2.5),
                  Inches(0.62), TEAL)
        tb(slide, Inches(1.45), top + Inches(0.33), Inches(2.5), Inches(0.44), word,
           size=19, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.45), top + Inches(0.92), Inches(2.5), Inches(0.34),
           f"look in {where}", size=11, color=SOFT, align=PP_ALIGN.CENTER)
        for j, opt in enumerate(options):
            left = Inches(4.25 + j * 2.78)
            add_round(slide, left, top + Inches(0.32), Inches(2.6), Inches(0.72),
                      L_TEAL)
            tb(slide, left + Inches(0.12), top + Inches(0.42), Inches(2.36),
               Inches(0.56), opt, size=12, bold=True, color=INK,
               align=PP_ALIGN.CENTER)
    hint(slide, "Read the whole sentence with each option swapped in. One fits.",
         6.42)


def s48_editor_checks():
    slide, n = new_slide("🖊️ Copy Editor — Your Five Checks", "LEARN", "77–84 min",
                         "Press Room", RED)
    one_task(slide, "Run every sentence past these five before you approve it.",
             RED)
    for i, (icon, label, detail, color) in enumerate(EDITOR_CHECKS):
        left = Inches(0.5 + i * 2.49)
        add_round(slide, left, Inches(1.9), Inches(2.3), Inches(3.3), L_GREY)
        tb(slide, left, Inches(2.1), Inches(2.3), Inches(0.7), icon, size=26,
           align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.12), Inches(2.88), Inches(2.06),
                  Inches(0.5), color)
        tb(slide, left + Inches(0.12), Inches(2.96), Inches(2.06), Inches(0.36),
           label, size=11, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.18), Inches(3.55), Inches(1.95), Inches(1.4),
           detail, size=11, color=DARK, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.5), Inches(5.4), Inches(12.35), Inches(0.85), L_AMBER)
    tb(slide, Inches(0.8), Inches(5.6), Inches(11.7), Inches(0.44),
       "⭐ Every sentence you fix completely earns one newspaper star. Five stars "
       "and the edition is approved.", size=15, bold=True, color=INK,
       align=PP_ALIGN.CENTER)
    hint(slide, "Have her name the check she is using out loud as she fixes.", 6.42)


def s49_fix_a():
    slide, n = new_slide("🖊️ Game: Fix the News! — Stars 1 and 2", "GAME",
                         "77–84 min", "Press Room", RED)
    one_task(slide, "More than one mistake in each. Name them all first.", RED)
    fix_rows(slide, FIX_NEWS_A, 1, RED, L_RED, gap=2.15, height=1.95,
             write_line=True)
    hint(slide, "Start at the very first letter. Capitals are the easiest catch.",
         6.42)


def s50_fix_b():
    slide, n = new_slide("🖊️ Fix the News! — Stars 3 and 4", "GAME", "77–84 min",
                         "Press Room", RED)
    one_task(slide, "One of these needs an adverb, not an adjective.", RED)
    fix_rows(slide, FIX_NEWS_B, 3, RED, L_RED, gap=2.15, height=1.95,
             write_line=True)
    hint(slide, "\"Quick\" describes a thing. \"Quickly\" describes how they build.",
         6.42)


def s51_fix_c():
    slide, n = new_slide("🖊️ Fix the News! — Stars 5 and 6", "GAME", "77–84 min",
                         "Press Room", RED)
    one_task(slide, "The last two. These need commas and a title.", RED)
    fix_rows(slide, FIX_NEWS_C, 5, RED, L_RED, gap=2.15, height=1.95,
             write_line=True)
    hint(slide, "A list of three things needs commas between them.", 6.42)


def s52_report_structure():
    slide, n = new_slide("🎤 Final Game: Be the Reporter", "CREATE", "84–88 min",
                         "Press Room", AMBER)
    one_task(slide, "Fill in five boxes. That is a complete news report.", AMBER)
    for i, (icon, field, detail) in enumerate(REPORT_FIELDS):
        top = Inches(1.85 + i * 0.94)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(0.8), WHITE)
        tb(slide, Inches(0.85), top + Inches(0.16), Inches(0.6), Inches(0.5), icon,
           size=17, align=PP_ALIGN.CENTER)
        add_round(slide, Inches(1.6), top + Inches(0.15), Inches(2.6),
                  Inches(0.5), L_AMBER)
        tb(slide, Inches(1.6), top + Inches(0.24), Inches(2.6), Inches(0.36), field,
           size=13, bold=True, color=AMBER, align=PP_ALIGN.CENTER)
        tb(slide, Inches(4.45), top + Inches(0.25), Inches(2.6), Inches(0.36),
           detail, size=11, color=SOFT)
        add_round(slide, Inches(7.3), top + Inches(0.15), Inches(5.2), Inches(0.5),
                  L_GREY)
        tb(slide, Inches(7.55), top + Inches(0.24), Inches(4.8), Inches(0.36),
           "____________________________________", size=12, color=SOFT)
    hint(slide, "Do the HEADLINE box last. It is easier once the facts are down.",
         6.6)


def s53_report_prompt():
    slide, n = new_slide("📷 Your Picture — Now Report It", "CREATE", "84–88 min",
                         "Press Room", AMBER)
    one_task(slide, "Two to four sentences. Say them, then write them.", AMBER)
    add_round(slide, Inches(0.5), Inches(1.8), Inches(5.3), Inches(4.45), L_AMBER)
    tb(slide, Inches(0.5), Inches(2.55), Inches(5.3), Inches(1.8), "🌻", size=90,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.8), Inches(4.6), Inches(4.7), Inches(1.0),
       "Students planting and watering at the new school garden.", size=15,
       bold=True, color=INK, align=PP_ALIGN.CENTER)
    tb(slide, Inches(6.15), Inches(1.8), Inches(6.7), Inches(0.36),
       "SENTENCE STARTERS — use them only if you need them", size=12, bold=True,
       color=SOFT)
    for i, starter in enumerate(REPORT_STARTERS):
        top = Inches(2.22 + i * 0.86)
        add_round(slide, Inches(6.15), top, Inches(6.7), Inches(0.72), WHITE)
        tb(slide, Inches(6.5), top + Inches(0.14), Inches(6.0), Inches(0.5),
           starter, size=20, bold=True, color=AMBER)
    add_round(slide, Inches(6.15), Inches(5.7), Inches(6.7), Inches(0.55), L_PURPLE)
    tb(slide, Inches(6.4), Inches(5.8), Inches(6.2), Inches(0.4),
       "⭐ EDITOR CHALLENGE: use one word from the vocabulary list.", size=12,
       bold=True, color=PURPLE)
    hint(slide, "Write down exactly what she says, then read it back to her.", 6.45)


def s54_final_a():
    slide, n = new_slide("🏆 Final Newsroom Challenge — 1 to 3", "FINAL",
                         "88–90 min", "Press Room", RED)
    one_task(slide, "No hints unless you ask. You have done all of this already.",
             RED)
    for i, (num, emoji, label, prompt, answer, color, light) in enumerate(FINAL[:3]):
        top = Inches(1.88 + i * 1.55)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.38), light)
        add_oval(slide, Inches(0.85), top + Inches(0.45), Inches(0.56),
                 Inches(0.56), color)
        tb(slide, Inches(0.85), top + Inches(0.52), Inches(0.56), Inches(0.42), num,
           size=14, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.7), top + Inches(0.42), Inches(0.9), Inches(0.62), emoji,
           size=24, align=PP_ALIGN.CENTER)
        tb(slide, Inches(2.8), top + Inches(0.3), Inches(3.6), Inches(0.44), label,
           size=14, bold=True, color=color)
        tb(slide, Inches(2.8), top + Inches(0.76), Inches(3.6), Inches(0.4), prompt,
           size=12, color=SOFT)
        add_round(slide, Inches(6.7), top + Inches(0.3), Inches(5.85), Inches(0.78),
                  WHITE)
        tb(slide, Inches(6.9), top + Inches(0.47), Inches(5.45), Inches(0.5),
           answer, size=22, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "Wait ten full seconds before offering anything. Silence helps.",
         6.45)


def s55_final_b():
    slide, n = new_slide("🏆 Final Newsroom Challenge — 4 and 5", "FINAL",
                         "88–90 min", "Press Room", AMBER)
    one_task(slide, "The last two. Then the edition goes to print.", AMBER)
    for i, (num, emoji, label, prompt, answer, color, light) in enumerate(FINAL[3:]):
        top = Inches(1.95 + i * 2.1)
        add_round(slide, Inches(0.5), top, Inches(12.35), Inches(1.85), light)
        add_oval(slide, Inches(0.85), top + Inches(0.62), Inches(0.62),
                 Inches(0.62), color)
        tb(slide, Inches(0.85), top + Inches(0.7), Inches(0.62), Inches(0.46), num,
           size=16, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.75), top + Inches(0.58), Inches(1.0), Inches(0.72),
           emoji, size=30, align=PP_ALIGN.CENTER)
        tb(slide, Inches(3.0), top + Inches(0.45), Inches(3.4), Inches(0.5), label,
           size=16, bold=True, color=color)
        tb(slide, Inches(3.0), top + Inches(0.98), Inches(3.4), Inches(0.45),
           prompt, size=13, color=SOFT)
        add_round(slide, Inches(6.7), top + Inches(0.45), Inches(5.85),
                  Inches(0.95), WHITE)
        tb(slide, Inches(6.9), top + Inches(0.62), Inches(5.45), Inches(0.62),
           answer, size=20, bold=True, color=color, align=PP_ALIGN.CENTER)
    hint(slide, "For number 5 any correct sentence counts. Praise it and move on.",
         6.42)


def s56_special_edition():
    _counter["n"] += 1
    n = _counter["n"]
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(slide, INK)
    add_rect(slide, Inches(0), Inches(0), prs.slide_width, Inches(0.26), RED)
    add_round(slide, Inches(4.35), Inches(0.6), Inches(4.65), Inches(0.5), RED)
    tb(slide, Inches(4.35), Inches(0.68), Inches(4.65), Inches(0.36),
       "🚨  SPECIAL EDITION COMPLETE  🚨", size=13, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(1.35), Inches(12), Inches(0.9),
       "YOU ARE OFFICIALLY A YOUNG REPORTER!", size=36, bold=True, color=WHITE,
       align=PP_ALIGN.CENTER, font="Georgia")
    add_oval(slide, Inches(5.67), Inches(2.4), Inches(2.0), Inches(2.0), RED)
    tb(slide, Inches(5.67), Inches(2.82), Inches(2.0), Inches(1.2), "🎤", size=44,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(0.7), Inches(4.6), Inches(12), Inches(0.6),
       "🎤  \"I can read it, fix it, and report it!\"  📰", size=24, bold=True,
       color=RGBColor(0xE8, 0xC9, 0x7A), align=PP_ALIGN.CENTER, font="Georgia")
    for i, (emoji, _num, name, _d, color) in enumerate(MISSIONS):
        left = Inches(1.05 + i * 2.35)
        add_round(slide, left, Inches(5.5), Inches(2.15), Inches(1.05), color)
        tb(slide, left, Inches(5.66), Inches(2.15), Inches(0.4), emoji, size=15,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.1), Inches(6.06), Inches(1.95), Inches(0.38),
           name, size=10, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    footer(slide, n, "88–90 min", "Press Room")
    fade(slide)


def s57_game_bank():
    slide, n = new_slide("🎲 Bonus Game Bank — If Time Remains", "EXTRA", "",
                         "Press Room", PURPLE)
    one_task(slide, "Seven spare games. Any one of them runs in three minutes.",
             PURPLE)
    for i, (emoji, name, detail, color) in enumerate(GAME_BANK[:6]):
        col, row = i % 3, i // 3
        left = Inches(0.5 + col * 4.15)
        top = Inches(1.85 + row * 2.05)
        add_round(slide, left, top, Inches(3.9), Inches(1.85), L_GREY)
        tb(slide, left + Inches(0.25), top + Inches(0.32), Inches(0.75),
           Inches(0.66), emoji, size=24, align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(1.2), top + Inches(0.28), Inches(2.5),
           Inches(0.44), name, size=14, bold=True, color=color)
        tb(slide, left + Inches(1.2), top + Inches(0.78), Inches(2.5),
           Inches(0.75), detail, size=11, color=DARK)
    emoji, name, detail, color = GAME_BANK[6]
    add_round(slide, Inches(0.5), Inches(5.95), Inches(12.35), Inches(0.72), L_GREY)
    tb(slide, Inches(0.85), Inches(6.08), Inches(0.6), Inches(0.5), emoji, size=17,
       align=PP_ALIGN.CENTER)
    tb(slide, Inches(1.6), Inches(6.1), Inches(3.0), Inches(0.44), name, size=14,
       bold=True, color=color)
    tb(slide, Inches(4.8), Inches(6.12), Inches(7.6), Inches(0.42), detail, size=12,
       color=DARK)


def s58_support():
    slide, n = new_slide("🔒 TEACHER ONLY — Support System", "TEACHER ONLY", "",
                         "Support", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.34),
       "Move between levels freely inside one activity. Never name the level out "
       "loud.", size=12, bold=True, color=RED)
    for i, (icon, name, what, how, color, light) in enumerate(SUPPORT_LEVELS):
        left = Inches(0.45 + i * 4.18)
        add_round(slide, left, Inches(1.72), Inches(3.95), Inches(2.3), light)
        tb(slide, left, Inches(1.86), Inches(3.95), Inches(0.46), icon, size=16,
           align=PP_ALIGN.CENTER)
        tb(slide, left + Inches(0.15), Inches(2.36), Inches(3.65), Inches(0.46),
           name, size=16, bold=True, color=color, align=PP_ALIGN.CENTER,
           font="Georgia")
        tb(slide, left + Inches(0.25), Inches(2.85), Inches(3.45), Inches(0.42),
           what, size=11, color=DARK, align=PP_ALIGN.CENTER)
        add_round(slide, left + Inches(0.2), Inches(3.3), Inches(3.55),
                  Inches(0.55), WHITE)
        tb(slide, left + Inches(0.3), Inches(3.41), Inches(3.35), Inches(0.38), how,
           size=10, color=SOFT, align=PP_ALIGN.CENTER, italic=True)
    add_round(slide, Inches(0.45), Inches(4.16), Inches(12.4), Inches(0.46), L_AMBER)
    tb(slide, Inches(0.7), Inches(4.23), Inches(11.9), Inches(0.34),
       "POSITIVE LABELS ONLY:   🟢 Reporter Hint   ·   🟡 Reporter Mission   ·   "
       "⭐ Editor Challenge", size=12, bold=True, color=AMBER)
    tb(slide, Inches(0.45), Inches(4.76), Inches(6.1), Inches(0.32),
       "THE HINT LADDER — never reveal the answer first", size=12, bold=True,
       color=INK)
    for i, (step, text, color) in enumerate(HINT_LADDER):
        top = Inches(5.14 + i * 0.44)
        add_round(slide, Inches(0.45), top, Inches(6.1), Inches(0.38), L_GREY)
        add_round(slide, Inches(0.6), top + Inches(0.05), Inches(1.0),
                  Inches(0.28), color)
        tb(slide, Inches(0.6), top + Inches(0.06), Inches(1.0), Inches(0.26), step,
           size=9, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
        tb(slide, Inches(1.75), top + Inches(0.04), Inches(4.6), Inches(0.3), text,
           size=11, bold=True, color=INK)
    add_round(slide, Inches(0.45), Inches(6.46), Inches(6.1), Inches(0.38), L_BLUE)
    tb(slide, Inches(0.7), Inches(6.51), Inches(5.6), Inches(0.3),
       "Then stop talking and let her answer.", size=11, bold=True, color=BLUE)
    add_round(slide, Inches(6.75), Inches(4.76), Inches(6.1), Inches(1.0), L_GREEN)
    tb(slide, Inches(7.0), Inches(4.84), Inches(5.6), Inches(0.3),
       "🟡 REPORTER HINT BANK", size=11, bold=True, color=GREEN)
    bullets(slide, Inches(7.0), Inches(5.14), Inches(2.85), Inches(0.58),
            HINT_BANK[:3], size=9, sp=1)
    bullets(slide, Inches(9.9), Inches(5.14), Inches(2.9), Inches(0.58),
            HINT_BANK[3:], size=9, sp=1)
    add_round(slide, Inches(6.75), Inches(5.86), Inches(6.1), Inches(0.98), L_RED)
    tb(slide, Inches(7.0), Inches(5.93), Inches(5.6), Inches(0.3),
       "🗣️ SAY THIS INSTEAD OF \"WRONG\"", size=11, bold=True, color=RED)
    bullets(slide, Inches(7.0), Inches(6.22), Inches(2.85), Inches(0.58),
            PRAISE[:3], size=9, sp=1)
    bullets(slide, Inches(9.9), Inches(6.22), Inches(2.9), Inches(0.58),
            PRAISE[3:], size=9, sp=1)


def s59_assessment():
    slide, n = new_slide("🔒 TEACHER ONLY — Final Assessment", "TEACHER ONLY", "",
                         "Assessment", DARK)
    tb(slide, Inches(0.45), Inches(1.28), Inches(12.4), Inches(0.32),
       "Record what she did today, not what you hoped for. Fill this in right after "
       "class.", size=12, bold=True, color=RED)
    cols = [(0.45, 4.3, "SKILL"), (4.95, 2.6, "Independent"),
            (7.75, 2.4, "With Support"), (10.35, 2.5, "Needs More Practice")]
    header_y = Inches(1.64)
    for left, width, label in cols:
        add_round(slide, Inches(left), header_y, Inches(width), Inches(0.38), INK)
        tb(slide, Inches(left), header_y + Inches(0.05), Inches(width),
           Inches(0.3), label, size=11, bold=True, color=WHITE,
           align=PP_ALIGN.CENTER)
    for i, skill in enumerate(RUBRIC_SKILLS):
        top = header_y + Inches(0.44 + i * 0.33)
        band = WHITE if i % 2 == 0 else L_GREY
        add_round(slide, Inches(0.45), top, Inches(4.3), Inches(0.31), band)
        tb(slide, Inches(0.65), top + Inches(0.02), Inches(3.9), Inches(0.27),
           skill, size=11, bold=True, color=INK)
        for left, width, _l in cols[1:]:
            add_round(slide, Inches(left), top, Inches(width), Inches(0.31), band)
            tb(slide, Inches(left), top - Inches(0.01), Inches(width), Inches(0.3),
               "☐", size=12, color=SOFT, align=PP_ALIGN.CENTER)
    add_round(slide, Inches(0.45), Inches(5.78), Inches(12.4), Inches(1.05), L_AMBER)
    tb(slide, Inches(0.7), Inches(5.87), Inches(11.9), Inches(0.34),
       "NEXT LESSON FOCUS — pick the three weakest rows above", size=12, bold=True,
       color=AMBER)
    for i in range(3):
        left = Inches(0.7 + i * 4.05)
        add_round(slide, left, Inches(6.25), Inches(3.85), Inches(0.46), WHITE)
        tb(slide, left + Inches(0.15), Inches(6.33), Inches(3.55), Inches(0.34),
           f"{i + 1}.  ____________________________", size=11, color=SOFT)


def s60_answer_key_a():
    slide, n = new_slide("🔒 TEACHER ONLY — Answer Key (1 of 2)", "TEACHER ONLY", "",
                         "Answer key", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Hide these two slides before presenting, or keep them on a second screen.",
       size=12, bold=True, color=RED)
    add_round(slide, Inches(0.45), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(0.7), Inches(1.86), Inches(5.6), Inches(0.36),
       "🗞️ Headlines  ·  🔎 Vocabulary", size=13, bold=True, color=RED)
    bullets(slide, Inches(0.7), Inches(2.28), Inches(5.6), Inches(4.3), [
        "Good news story: D — all of the above",
        "Match 1–3: garden, puppy, science fair",
        "Match 4–6: book club, pizza, art show",
        "Mix-Up 1: SCHOOL GARDEN GROWS GIANT PUMPKIN",
        "Mix-Up 2: STUDENTS WIN SCIENCE FAIR",
        "Mix-Up 3: NEW BOOK CLUB STARTS",
        "Word or Not — YES: reporter, headline, garden,",
        "     interview, editor",
        "Word or Not — NO: banana, rocket, penguin",
        "Who Am I: headline, reporter, discover",
        "Which Word Fits: interview, carefully, celebrate",
    ], size=11, sp=4)
    add_round(slide, Inches(6.75), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(7.0), Inches(1.86), Inches(5.6), Inches(0.36),
       "✏️ Grammar  ·  ⏰ Tenses", size=13, bold=True, color=GREEN)
    bullets(slide, Inches(7.0), Inches(2.28), Inches(5.6), Inches(4.3), [
        "Right verb: helps, come, plants, build, checks",
        "Red Pen 1–3: students write, reporter asks,",
        "     friend reads",
        "Red Pen 4–5: teachers help, Priya and Sam plant",
        "Spin 1 — noun: reporter, story · verb: wrote ·",
        "     adj: excited · adv: quickly",
        "Spin 2 — noun: students, seeds · verb: planted ·",
        "     adj: busy, small · adv: carefully",
        "Spin fix: The garden grows fast. / Priya writes",
        "     three headlines. / The birds build a nest.",
        "Joining words: and, but, so, because",
        "Name the tense: present, past, future, past, present",
        "Irregulars: grew, took, wrote",
    ], size=11, sp=3)


def s61_answer_key_b():
    slide, n = new_slide("🔒 TEACHER ONLY — Answer Key (2 of 2)", "TEACHER ONLY", "",
                         "Answer key", DARK)
    tb(slide, Inches(0.45), Inches(1.3), Inches(12.4), Inches(0.32),
       "Reading, evidence, editing and the final challenge.", size=12, bold=True,
       color=RED)
    add_round(slide, Inches(0.45), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(0.7), Inches(1.86), Inches(5.6), Inches(0.36),
       "📖 Reading  ·  🔍 Evidence", size=13, bold=True, color=BLUE)
    bullets(slide, Inches(0.7), Inches(2.28), Inches(5.6), Inches(4.3), [
        "Main idea: B — students turned an unused space",
        "     into a garden that brought wildlife",
        "Sequence: idea → plan → clear and plant → animals",
        "Q1 Priya saw the empty space and had an idea",
        "     (Part 1)",
        "Q2 sunflowers, tomatoes, beans, carrots (Part 3)",
        "Q3 butterflies, bees and two brown birds (Part 4)",
        "Q4 they were really building a home for animals",
        "     (Part 4)",
        "Q5 it is now the busiest corner of the school",
        "Inference: she keeps going even when it is hard",
        "Word meanings: brought closer · a plan of who does",
        "     what and when · removed things from",
    ], size=11, sp=3)
    add_round(slide, Inches(6.75), Inches(1.72), Inches(6.1), Inches(5.1), WHITE)
    tb(slide, Inches(7.0), Inches(1.86), Inches(5.6), Inches(0.36),
       "🖊️ Fix the News  ·  🏆 Final", size=13, bold=True, color=AMBER)
    bullets(slide, Inches(7.0), Inches(2.28), Inches(5.6), Inches(4.3), [
        "1. The students plant flowers in the garden.",
        "2. The reporter wrote a story yesterday.",
        "3. Priya and her team work very hard.",
        "4. The birds build a nest quickly.",
        "5. We planted sunflowers, tomatoes, and beans.",
        "6. Mr. Ellis asked a hard question.",
        "Final 1: she reads INTERVIEW out loud",
        "Final 2: the verb is WROTE",
        "Final 3: The team won.",
        "Final 4: butterflies, bees and birds",
        "Final 5: any correct sentence using DISCOVER",
    ], size=11, sp=4)


BUILDERS = [
    s01_title, s02_missions, s03_good_story, s04_warmup, s05_method,
    s06_what_headline, s07_four_headlines, s08_match_a, s09_match_b, s10_mixup,
    s11_mission1_done, s12_vocab_1, s13_vocab_2, s14_vocab_3, s15_word_or_not,
    s16_who_am_i, s17_context, s18_subject_predicate, s19_agreement_learn,
    s20_agreement_practice, s21_red_pen_a, s22_red_pen_b, s23_pos_key,
    s24_spin_wheel, s25_spin_1, s26_spin_2, s27_spin_fix, s28_conjunctions,
    s29_break_intro, s30_break_actions, s31_timeline, s32_tense_chart,
    s33_name_tense, s34_transform, s35_irregular, s36_mission_tense,
    s37_story_intro, s38_story_1, s39_story_2, s40_story_3, s41_story_4,
    s42_main_idea, s43_evidence_how, s44_evidence_a, s45_evidence_b,
    s46_evidence_c, s47_context_story, s48_editor_checks, s49_fix_a, s50_fix_b,
    s51_fix_c, s52_report_structure, s53_report_prompt, s54_final_a, s55_final_b,
    s56_special_edition, s57_game_bank, s58_support, s59_assessment,
    s60_answer_key_a, s61_answer_key_b,
]

for build in BUILDERS:
    build()

assert len(prs.slides) == TOTAL, f"expected {TOTAL} slides, built {len(prs.slides)}"

out = r"C:\Users\bhushaja\Downloads\shaip\Grade5_School_Newsroom_90min.pptx"
try:
    prs.save(out)
except PermissionError:
    out = out.replace(".pptx", "_new.pptx")
    prs.save(out)

story_words = sum(len(p.split()) for part in STORY for p in part[5])
with_notes = [i + 1 for i, s in enumerate(prs.slides)
              if s.has_notes_slide and s.notes_slide.notes_text_frame.text.strip()]
print(f"Saved: {out}")
print(f"Slides: {len(prs.slides)}")
print(f"Passage word count: {story_words}")
print(f"Slides carrying speaker notes: {with_notes if with_notes else 'none'}")
