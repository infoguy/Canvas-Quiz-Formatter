"""
docx_quiz_parser.py  v11
========================
Parses style Word documents to text2qti format.

Correct answers are detected from YELLOW HIGHLIGHT formatting in the .docx,
or from an answer-key line ("Ans: D", "Answer - A", "Correct Answer: B").

Fixes vs v10:
  - Answer-key lines are now recognized in every common WGU spelling:
    "Ans:", "Ans -", "Ans.", "Ans D", "Answer -", "Answer:", "Answer.",
    "Correct Answer:", "Correct answer is", "Answer key:", with or without
    spaces around the separator and with optional trailing punctuation.
    Previously only "Answer-" and "Correct Answer:" matched, so any other
    spelling was emitted as its own bogus essay question ("3. Ans: D" + ___)
    and the real question was left with no correct answer marked.
  - Fixed an UnboundLocalError: the fill-in-the-blank branch read
    answer_letter before it was assigned.  Answer-key detection now runs
    ahead of the FITB test, so a FITB question no longer crashes the parse
    (or silently inherits the previous question's answer letter).

Fixes vs v8:
  - Matching questions: detects two-column matching blocks where prompts are
    packed into one paragraph (soft-return separated) and answers follow as
    individually numbered paragraphs.  Outputs each pair as a text2qti
    matching item: stem line + "[match] answer" answer line.

Fixes vs v7:
  - Handles "packed" paragraphs: entire question + options in one <w:p>
    element separated by soft-return <w:br> characters.  These appear as
    embedded \n in p.text and were previously swallowed as a single line,
    causing questions to be silently dropped or mis-parsed.
  - Per-run, per-line highlight detection so the correct option is
    identified accurately even inside a packed paragraph.

Fixes vs v6:
  - Strip leading "N." / "N. " question-number prefixes from stems
  - Recognize uppercase A. / B. / C. / D. option prefixes (not just a-e)
  - Extended option letter range A-J / a-j (was capped at E/e)
  - Broader True/False section-header detection (e.g. "True/False Questions")
  - "True False" table-header row skipped correctly

Usage:
  python docx_quiz_parser.py              → opens GUI window
  python docx_quiz_parser.py <file.docx>  → command-line mode
"""

import re, sys
from pathlib import Path
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

OPTION_LETTERS = "abcdefghij"

# Matches  A)  a)  A.  a.  A<nbsp>  a<nbsp>  A<space>  a<space>  followed by text
# \xa0 = non-breaking space used in some WGU docx files instead of ". " or ") "
LETTER_PAREN_RE = re.compile(r'^\s*[A-Ja-j](?:[.)]\s*|\xa0)(.+)')

# Matches "Rationale:" lines to skip
RATIONALE_RE = re.compile(r'^Rationale:', re.I)

# Shared prefix for every answer-key style line.  Accepts, case-insensitively:
#   Answer- A     Answer - A    Answer: A    Answer:A     Answer. A
#   Ans: D        Ans - D       Ans. D       Ans D        ANS:D
#   Correct Answer: B           Correct answer is B       Answer key: B
# i.e. "ans" or "answer" (optionally plural), an optional separator drawn from
# - : . ) or an en/em dash, optional "key" / "is", then the answer value.
_ANS_PREFIX = (r'(?:correct\s+)?ans(?:wer)?s?\b\.?\s*'
               r'(?:key\b)?\s*(?:is\b)?\s*[-\u2010-\u2015:.)]?\s*')

# Trailing punctuation allowed after the value: "Answer: A." / "Ans - D)"
_ANS_SUFFIX = r'\s*[.)]?\s*$'

# Matches "Answer- True" / "Ans: False"
ANSWER_TF_RE = re.compile(r'^\s*' + _ANS_PREFIX + r'(True|False)' + _ANS_SUFFIX, re.I)

# Matches "Answer- A" / "Correct Answer: B" / "Ans: D" (MC answer key style)
ANSWER_MC_RE = re.compile(r'^\s*' + _ANS_PREFIX + r'([A-Ja-j])' + _ANS_SUFFIX, re.I)

# Combined: matches either TF or MC answer lines
ANSWER_RE = re.compile(r'^\s*' + _ANS_PREFIX + r'(True|False|[A-Ja-j])' + _ANS_SUFFIX, re.I)

# Skip rows that are just the two-column "True  False" table header
TF_TABLE_HEADER_RE = re.compile(r'^True\s+False$', re.I)

# Skip section-header lines that look like "True/False Questions" etc.
TF_SECTION_HEADER_RE = re.compile(r'^True[/\s-]*False', re.I)

# Leading question-number prefix:  "1."  "1. "  "12. "
Q_NUM_PREFIX_RE = re.compile(r'^\d+\.\s*')

# Numbered answer/prompt line: "31. Some text"  (number 1-99)
NUMBERED_LINE_RE = re.compile(r'^(\d+)\.\s+(.+)')

# ─── Two-column (tab-separated) matching support ────────────────────────────

# A standalone section header introducing a matching block: "Matching",
# "Matching Questions", "Matching - Word Parts", etc.
MATCHING_HEADER_RE = re.compile(r'^matching\b', re.I)

# Leading fill-in dashes/underscores used to draw the answer blank in the
# left column: "--------Plasty", "____otomy".  Collapsed to a single hyphen
# so Canvas shows "-Plasty" instead of a row of dashes.
LEADING_BLANK_RE = re.compile(r'^[-\u2010-\u2015_]{2,}\s*')

# Minimum consecutive two-column rows before a run is treated as a matching
# block.  Two rows is too loose — ordinary tabbed text would false-positive.
MIN_MATCHING_ROWS = 3

# Set False to keep the raw "--------Plasty" text instead of "-Plasty".
NORMALIZE_MATCHING_BLANKS = True

# Second line emitted under the matching question title.
MATCHING_INSTRUCTIONS = "Match each term to its definition."


def _clean(t):
    return (t.replace('\u2019', "'").replace('\u2018', "'")
             .replace('\u201c', '"').replace('\u201d', '"').strip())


def _strip_letter(text):
    """Remove leading 'A.' / 'a)' / 'A) ' style prefix from option text."""
    m = LETTER_PAREN_RE.match(text)
    return m.group(1).strip() if m else text


def _strip_q_num(text):
    """Remove a leading question number like '1.' or '12. ' from stem text."""
    return Q_NUM_PREFIX_RE.sub('', text).strip()


def _is_tf_header(text):
    """True for both the table-header 'True False' and section headers."""
    return bool(TF_TABLE_HEADER_RE.match(text) or TF_SECTION_HEADER_RE.match(text))


def _expand_paragraph(para):
    """
    Return a list of (text, highlighted) tuples — one per logical line.

    Most paragraphs yield exactly one tuple.  "Packed" paragraphs — where
    the author pressed Shift+Enter (soft return, <w:br>) instead of Enter
    to separate the stem from its options — contain embedded \\n characters
    and yield one tuple per soft-return-delimited line.

    Highlight is tracked per run then mapped to each line so the correct
    answer is detected even when only the correct option's run is yellow.
    """
    # Build a list of (text_fragment, is_highlighted) per run fragment,
    # splitting on the embedded newlines that soft returns produce.
    fragments = []   # list of (str, bool)
    for run in para.runs:
        hl = run.font.highlight_color == WD_COLOR_INDEX.YELLOW
        parts = run.text.split('\n')
        for part_i, part in enumerate(parts):
            if part_i > 0:
                fragments.append(('\n', False))   # line-break sentinel
            fragments.append((part, hl))

    # Collapse fragments into lines
    lines = []          # list of (text, highlighted)
    cur_text = ''
    cur_hl = False
    for frag, hl in fragments:
        if frag == '\n':
            t = _clean(cur_text)
            if t:
                lines.append((t, cur_hl))
            cur_text = ''
            cur_hl = False
        else:
            cur_text += frag
            if hl and frag.strip():
                cur_hl = True

    t = _clean(cur_text)
    if t:
        lines.append((t, cur_hl))

    return lines  # empty list if paragraph was blank


def _normalize_matching_term(text):
    """
    Tidy a left-column term: drop a leading list number ("1. -ectomy") and
    collapse a run of leading dashes/underscores to a single hyphen, so
    "--------Plasty" imports as "-Plasty" rather than a row of dashes.
    """
    t = _strip_q_num(text)
    if NORMALIZE_MATCHING_BLANKS:
        t = LEADING_BLANK_RE.sub('-', t)
    return t.strip()


def _split_two_columns(text):
    """
    Split a tab-separated two-column line into (left, right).

    WGU matching worksheets lay the term and the definition out side by side
    using literal tabs for spacing, so one paragraph looks like:

        "--------Plasty\\t\\t\\t\\t\\t\\tLiver\\xa0"

    Any number of consecutive tabs counts as one column break.  Returns None
    when the line isn't a two-column row (no tab, or one side empty).
    """
    t = _clean(text.replace('\xa0', ' '))
    if '\t' not in t:
        return None
    # "a)\tSome option" is a tabbed multiple-choice option, not a matching row.
    # Without this guard three consecutive tabbed options would be swallowed
    # as a bogus matching block.
    if LETTER_PAREN_RE.match(t):
        return None
    parts = [p.strip() for p in t.split('\t')]
    parts = [p for p in parts if p]
    if len(parts) < 2:
        return None
    left = parts[0]
    # A stray tab inside a long definition is far more likely than a real
    # third column, so everything after the first break rejoins as the right.
    right = ' '.join(parts[1:])
    if not left or not right:
        return None
    return left, right


def _extract_tab_matching_blocks(paras):
    """
    Scan raw paragraphs for tab-separated two-column matching blocks.

    Returns ({first_para_index: {'title': str, 'pairs': [(left, right), ...]}},
             skip_indices).

    A block is a run of MIN_MATCHING_ROWS or more consecutive non-blank
    paragraphs that each split cleanly into two tab-separated columns.  If the
    nearest non-blank paragraph above the run is a short "Matching…" header, it
    becomes the question title and is consumed so it doesn't fall through and
    become a stray essay question.

    IMPORTANT: pairs are taken row by row as the document lays them out.  On a
    worksheet whose right-hand column is a scrambled answer bank, those rows are
    NOT the answer key — see the mismatch warning in generate_report().
    """
    matching = {}
    skip = set()
    i = 0
    n = len(paras)

    while i < n:
        if _split_two_columns(paras[i].text) is None:
            i += 1
            continue

        rows, idxs = [], []
        j = i
        while j < n:
            pair = _split_two_columns(paras[j].text)
            if pair is None:
                break
            rows.append(pair)
            idxs.append(j)
            j += 1

        if len(rows) >= MIN_MATCHING_ROWS:
            # Look back past blank paragraphs for a "Matching" style header
            title = 'Matching'
            k = i - 1
            while k >= 0 and not _clean(paras[k].text):
                k -= 1
            if k >= 0:
                cand = _clean(paras[k].text.replace('\xa0', ' '))
                if MATCHING_HEADER_RE.match(cand) and len(cand.split()) <= 8:
                    title = cand
                    skip.add(k)

            pairs = [(_normalize_matching_term(l), r) for l, r in rows]
            matching[i] = {'title': title, 'pairs': pairs}
            skip.update(idxs[1:])   # first row's index carries the block token
            i = j
        else:
            i += 1

    return matching, skip


def _extract_matching_blocks(paras):
    """
    Scan raw paragraphs for two-column matching blocks and return a dict:
      { para_index: [(prompt_text, answer_text), ...] }

    A matching block is identified when:
      1. A paragraph's text (after clean) contains multiple \n-separated lines
         that all begin with a question number ("31. ...", "32. ...", etc.)
         forming a consecutive run of numbers — this is the "prompts" column.
      2. The N paragraphs immediately following contain single lines, each
         beginning with the same consecutive question numbers — these are the
         "answers" column.

    The para_index key is the index of the packed prompts paragraph.
    The value is a list of (prompt_text, answer_text) pairs in order.
    Answer paragraph indices are stored in the returned skip_indices set.
    """
    matching = {}   # para_idx -> list of (prompt, answer)
    skip = set()    # para indices consumed as answer columns

    for pi, p in enumerate(paras):
        if pi in skip:
            continue

        raw = _clean(p.text.replace('\xa0', ' '))
        # Split on newlines (soft returns produce \n inside p.text)
        sub_lines = [_clean(ln) for ln in raw.split('\n') if _clean(ln)]

        if len(sub_lines) < 2:
            continue

        # Check that every sub-line starts with a consecutive number
        nums = []
        texts = []
        for ln in sub_lines:
            m = NUMBERED_LINE_RE.match(ln)
            if not m:
                nums = []
                break
            nums.append(int(m.group(1)))
            texts.append(m.group(2).strip())

        if len(nums) < 2:
            continue

        # Verify numbers are consecutive
        if nums != list(range(nums[0], nums[0] + len(nums))):
            continue

        # Now look ahead for len(nums) answer paragraphs with matching numbers
        expected = nums[:]
        answers = []
        ai = pi + 1
        matched_ans_idxs = []
        while len(answers) < len(nums) and ai < len(paras):
            ans_raw = _clean(paras[ai].text.replace('\xa0', ' '))
            # Skip blank paragraphs between packed para and answer column
            if not ans_raw:
                ai += 1
                continue
            m = NUMBERED_LINE_RE.match(ans_raw)
            if m and int(m.group(1)) == expected[len(answers)]:
                answers.append(m.group(2).strip())
                matched_ans_idxs.append(ai)
                ai += 1
            else:
                break   # not a matching answer line

        if len(answers) == len(nums):
            pairs = list(zip(texts, answers))
            matching[pi] = pairs
            skip.update(matched_ans_idxs)

    return matching, skip


def _looks_like_title(p):
    """True only if this paragraph is a document title, not question content.

    The old rule was "no '?' and <= 8 words", which silently ate any short
    question that didn't end in a question mark, e.g.
    "The color of the sky is ____."  When the very first question was eaten,
    its answer lines became the next question's stem.

    A paragraph is disqualified from being a title if it shows any sign of
    being question content: a Word auto-number (the "5." lives in the
    numbering, not in p.text), a literal "N." prefix, a fill-in blank, a
    sentence-ending period, or a question mark.
    """
    t = _clean(p.text)
    if not t:
        return False

    # Explicit Heading/Title styles are always titles.
    style = (p.style.name or '') if p.style is not None else ''
    if style.lower().startswith(('title', 'heading', 'subtitle')):
        return True

    # Auto-numbered list item → it's a numbered question, not a title.
    pPr = p._p.pPr
    if pPr is not None and pPr.numPr is not None:
        return False

    # Literal "1." / "12. " question-number prefix → question.
    if Q_NUM_PREFIX_RE.match(t):
        return False

    # A fill-in blank (___) → question stem.
    if re.search(r'_{3,}', t):
        return False

    # Question mark or sentence-ending punctuation → question, not a title.
    if t.endswith(('?', '.', '!')):
        return False

    return len(t.split()) <= 8


def parse_docx(docx_path):
    doc = Document(docx_path)

    # Skip title line(s) — but only if the first paragraph really is a title.
    paras = doc.paragraphs
    first_nonblank = next((i for i, p in enumerate(paras) if _clean(p.text)), 0)
    if _looks_like_title(paras[first_nonblank]):
        start = first_nonblank + 1
    else:
        start = 0

    # ── Pre-scan for matching blocks ──────────────────────────────────────
    # Two shapes are recognized: the numbered packed-column layout, and the
    # tab-separated side-by-side layout.  Both normalize to
    # {'title': str, 'pairs': [(left, right), ...]}.
    numbered_blocks, skip_indices = _extract_matching_blocks(paras)
    matching_blocks = {k: {'title': 'Matching', 'pairs': v}
                       for k, v in numbered_blocks.items()}

    tab_blocks, tab_skip = _extract_tab_matching_blocks(paras)
    for k, v in tab_blocks.items():
        if k not in skip_indices:
            matching_blocks[k] = v
    skip_indices = set(skip_indices) | tab_skip

    # ── Expand all paragraphs into (text, highlighted) lines ──────────────
    # Packed paragraphs (soft returns inside one <w:p>) are split here so
    # the rest of the parser never has to think about them.
    all_lines = []   # list of (text, highlighted) | None for blank separators
                     # or dict {'matching': [(prompt, answer), ...]}
    for pi, p in enumerate(paras[start:], start=start):
        # Matching prompts paragraph — inject as a special token
        if pi in matching_blocks:
            all_lines.append(None)  # flush whatever came before
            all_lines.append({'matching': matching_blocks[pi]})
            all_lines.append(None)
            continue
        # Answer-column paragraphs consumed by matching — skip entirely
        if pi in skip_indices:
            continue

        expanded = _expand_paragraph(p)
        if not expanded:
            all_lines.append(None)   # blank-line separator
        else:
            all_lines.extend(expanded)
            # If a packed paragraph had >1 line, treat it as a self-contained
            # group by appending an implicit blank separator after it.
            if len(expanded) > 1:
                all_lines.append(None)

    # ── Build groups (blank-line-separated) ───────────────────────────────
    groups = []
    cur = []

    def flush():
        nonlocal cur
        if cur:
            groups.append(cur)
            cur = []

    _skip_rationale_body = [False]  # mutable flag for rationale body skip
    for item in all_lines:
        if item is None:
            _skip_rationale_body[0] = False  # blank line resets rationale skip
            # Don't split a stem from its options on an intervening blank line.
            # Only flush if we have a complete group (stem + at least 1 option)
            # or the group is empty.
            if len(cur) != 1:
                flush()
            continue

        # Matching block token — flush current, then pass through as its own group
        if isinstance(item, dict) and 'matching' in item:
            flush()
            groups.append([item])
            continue

        t, highlighted = item

        # Any True/False header → flush and create a singleton skip-group
        if _is_tf_header(t):
            flush()
            groups.append([(t, highlighted)])
            continue

        # Skip "Rationale:" lines and the rationale text that follows (soft-return split)
        if RATIONALE_RE.match(t):
            flush()
            _skip_rationale_body[0] = True
            continue
        if _skip_rationale_body[0]:
            # This line is the rationale body text — skip it
            _skip_rationale_body[0] = False
            continue

        # "Answer- True/False/X" → always closes the current group
        if ANSWER_RE.match(t) or ANSWER_TF_RE.match(t):
            cur.append((t, highlighted))
            flush()
            continue

        # A new numbered stem (e.g. "5. A patient's...") appearing while we
        # already have content in cur means two questions ran together without
        # a blank line between them — flush the current group first.
        if Q_NUM_PREFIX_RE.match(t) and cur:
            flush()

        cur.append((t, highlighted))

    flush()

    # ── Classify groups into questions ────────────────────────────────────
    questions = []
    q_num = 0
    i = 0

    while i < len(groups):
        g = groups[i]

        # ── Matching block ──
        if len(g) == 1 and isinstance(g[0], dict) and 'matching' in g[0]:
            blk = g[0]['matching']
            q_num += 1
            questions.append({
                'num': q_num,
                'stem': blk['title'],
                'instructions': MATCHING_INSTRUCTIONS,
                'type': 'matching',
                'pairs': blk['pairs'],  # list of (prompt, answer)
            })
            i += 1
            continue

        # Skip True/False header groups
        if len(g) == 1 and _is_tf_header(g[0][0]):
            i += 1
            continue

        raw_stem, _ = g[0]
        stem_text = _strip_q_num(raw_stem)

        # ── True/False: stem + "Answer- True/False" in same group ──
        if len(g) == 2 and (ANSWER_RE.match(g[1][0]) or ANSWER_TF_RE.match(g[1][0])):
            m = ANSWER_RE.match(g[1][0]) or ANSWER_TF_RE.match(g[1][0])
            correct_tf = m.group(1).capitalize()
            q_num += 1
            questions.append({
                'num': q_num,
                'stem': stem_text,
                'type': 'tf',
                'options': ['True', 'False'],
                'correct': correct_tf,
            })
            i += 1
            continue

        # ── Legacy T/F: statement alone, answer in next group ──
        if (len(g) == 1
                and i + 1 < len(groups)
                and len(groups[i + 1]) == 1
                and (ANSWER_RE.match(groups[i + 1][0][0]) or ANSWER_TF_RE.match(groups[i + 1][0][0]))):
            ans_text, _ = groups[i + 1][0]
            m = ANSWER_RE.match(ans_text) or ANSWER_TF_RE.match(ans_text)
            correct_tf = m.group(1).capitalize()
            q_num += 1
            questions.append({
                'num': q_num,
                'stem': stem_text,
                'type': 'tf',
                'options': ['True', 'False'],
                'correct': correct_tf,
            })
            i += 2
            continue

        # ── Split the group into option lines + optional answer-key line ──
        # NOTE: this must run BEFORE the fill-in-the-blank test below, which
        # reads answer_letter.  (It used to run after, so answer_letter was
        # referenced before assignment.)
        option_lines = g[1:]
        answer_letter = None
        # Strip rationale lines from tail
        while option_lines and RATIONALE_RE.match(option_lines[-1][0]):
            option_lines = option_lines[:-1]
        # Strip answer key line from tail
        if option_lines and ANSWER_MC_RE.match(option_lines[-1][0]):
            m_ans = ANSWER_MC_RE.match(option_lines[-1][0])
            answer_letter = m_ans.group(1).lower()
            option_lines = option_lines[:-1]
        elif option_lines and ANSWER_TF_RE.match(option_lines[-1][0]):
            option_lines = option_lines[:-1]

        # ── Fill-in-the-blank: stem contains ____ with NO lettered MC options ──
        BLANK_RE = re.compile(r'_{3,}')
        has_letter_options = any(LETTER_PAREN_RE.match(opt_text) for opt_text, _ in g[1:])
        has_answer_key = any(ANSWER_RE.match(opt_text) for opt_text, _ in g[1:])
        if BLANK_RE.search(stem_text) and not has_letter_options and not has_answer_key and not answer_letter:
            # Normalize blanks to exactly four underscores for the converter
            normalized_stem = BLANK_RE.sub('____', stem_text)
            # Highlighted option lines are accepted answers; fall back to all lines
            answers = [opt_text for opt_text, opt_hl in g[1:] if opt_hl]
            if not answers:
                answers = [_strip_letter(opt_text) for opt_text, _ in g[1:]]
            if answers:
                q_num += 1
                questions.append({
                    'num': q_num,
                    'stem': normalized_stem,
                    'type': 'fitb',
                    'answers': answers,
                })
                i += 1
                continue

        # ── Multiple choice: stem + 2+ option lines ──
        # Merge leading continuation-stem lines (no letter prefix AND no answer key)
        # Only do this if an answer_letter was found, confirming this IS an MC question,
        # AND the first option_line looks like a question (ends with '?'), not an answer.
        if answer_letter and option_lines:
            candidate, _ = option_lines[0]
            # A continuation stem line: doesn't match letter option, doesn't match answer key,
            # and typically ends with '?' or is clearly a question
            if (not LETTER_PAREN_RE.match(candidate)
                    and not ANSWER_RE.match(candidate)
                    and candidate.rstrip().endswith('?')):
                stem_text = stem_text + ' ' + candidate
                option_lines = option_lines[1:]

        if len(option_lines) >= 2:
            options = []
            correct_idx = None
            for j, (opt_text, opt_hl) in enumerate(option_lines):
                cleaned_opt = _strip_letter(opt_text)
                options.append(cleaned_opt)
                if opt_hl:
                    correct_idx = j
            if correct_idx is None and answer_letter:
                letter_idx = OPTION_LETTERS.find(answer_letter)
                if 0 <= letter_idx < len(options):
                    correct_idx = letter_idx
            q_num += 1
            questions.append({
                'num': q_num,
                'stem': stem_text,
                'type': 'mc',
                'options': options,
                'correct': correct_idx,
            })
            i += 1
            continue

        # ── Fallback: essay ──
        q_num += 1
        questions.append({
            'num': q_num,
            'stem': stem_text,
            'type': 'essay',
            'options': [],
            'correct': None,
        })
        i += 1

    # ── Post-pass: absorb orphaned "Answer- X" essay groups into the preceding MC ──
    # This handles cases where a blank line separates the Answer line from the options.
    cleaned = []
    for q in questions:
        if (q['type'] == 'essay'
                and not q['options']
                and ANSWER_RE.match(q['stem'])):
            # This is an orphaned answer line
            m = ANSWER_RE.match(q['stem'])
            val = m.group(1)
            if cleaned and cleaned[-1]['type'] == 'mc' and cleaned[-1]['correct'] is None:
                # Attach to the previous MC question
                letter = val.lower()
                letter_idx = OPTION_LETTERS.find(letter)
                if 0 <= letter_idx < len(cleaned[-1]['options']):
                    cleaned[-1]['correct'] = letter_idx
                    # Re-number subsequent questions is handled by the num field not mattering
                    continue  # skip adding this as its own question
            # If it couldn't be absorbed, just skip it silently
            continue
        cleaned.append(q)

    # Re-number
    for idx, q in enumerate(cleaned, 1):
        q['num'] = idx

    return cleaned


def questions_to_text2qti(questions):
    lines = []
    for q in questions:
        n = q['num']
        stem = q['stem']

        if q['type'] == 'mc':
            lines.append(f"{n}. {stem}")
            for j, opt in enumerate(q['options']):
                ltr = OPTION_LETTERS[j] if j < len(OPTION_LETTERS) else f"opt{j}"
                prefix = "*" if j == q['correct'] else ""
                lines.append(f"{prefix}{ltr}) {opt}")
            lines.append("")

        elif q['type'] == 'tf':
            lines.append(f"{n}. {stem}")
            for opt in q['options']:
                prefix = "*" if opt == q['correct'] else ""
                ltr = 'a' if opt == 'True' else 'b'
                lines.append(f"{prefix}{ltr}) {opt}")
            lines.append("")

        elif q['type'] == 'matching':
            lines.append(f"{n}. {q['stem']}")
            if q.get('instructions'):
                lines.append(q['instructions'])
            for prompt, answer in q['pairs']:
                lines.append(f"= {prompt} -> {answer}")
            lines.append("")

        elif q['type'] == 'fitb':
            lines.append(f"{n}. {q['stem']}")
            for answer in q['answers']:
                lines.append(f"* {answer}")
            lines.append("")

        elif q['type'] == 'essay':
            lines.append(f"{n}. {stem}")
            lines.append("___")
            lines.append("")

    return "\n".join(lines)


def generate_report(questions):
    from collections import Counter
    tc = Counter(q['type'] for q in questions)
    label = {'mc': 'Multiple Choice', 'tf': 'True/False', 'essay': 'Essay/Short Ans', 'matching': 'Matching', 'fitb': 'Fill in the Blank'}
    out = [
        "=" * 60,
        "QUIZ PARSE REPORT",
        "=" * 60,
        f"Total questions parsed: {len(questions)}",
        "",
        "By question type:",
    ]
    for qt, c in sorted(tc.items()):
        out.append(f"  {label.get(qt, qt):<22} {c}")

    matching_qs = [q for q in questions if q['type'] == 'matching']
    if matching_qs:
        out += [
            "",
            "!! VERIFY MATCHING PAIRS",
            "   Pairs were read row by row, exactly as the columns sit in the",
            "   Word file.  If the right-hand column is a scrambled answer",
            "   bank (typical on a student worksheet), those rows are NOT the",
            "   answer key and must be corrected before import.",
        ]
        for q in matching_qs:
            out.append(f"   Q{q['num']}  {len(q['pairs'])} pairs to check")

    out.append("=" * 60)
    return "\n".join(out)


def parse_docx_to_text2qti(docx_path):
    qs = parse_docx(docx_path)
    return questions_to_text2qti(qs), generate_report(qs)


# ─── GUI ────────────────────────────────────────────────────────────────────

def launch_gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox

    root = tk.Tk()
    root.title("DOCX Quiz Parser  v10")
    root.resizable(True, True)
    root.minsize(560, 480)

    selected_path = tk.StringVar()
    output_text   = tk.StringVar()

    pad = ttk.Frame(root, padding=16)
    pad.pack(fill="both", expand=True)

    ttk.Label(pad, text="Word document (.docx):").grid(row=0, column=0, sticky="w", pady=(0, 4))
    file_frame = ttk.Frame(pad)
    file_frame.grid(row=1, column=0, sticky="ew", pady=(0, 12))
    pad.columnconfigure(0, weight=1)
    file_frame.columnconfigure(0, weight=1)

    file_entry = ttk.Entry(file_frame, textvariable=selected_path, state="readonly")
    file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 6))

    def browse():
        p = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx"), ("All files", "*.*")])
        if p:
            selected_path.set(p)
            status_var.set("File selected. Click Parse to continue.")

    ttk.Button(file_frame, text="Browse…", command=browse).grid(row=0, column=1)

    status_var = tk.StringVar(value="Choose a .docx file to get started.")
    status_lbl = ttk.Label(pad, textvariable=status_var, foreground="gray")
    status_lbl.grid(row=2, column=0, sticky="w", pady=(0, 8))

    def run_parse():
        path = selected_path.get()
        if not path:
            messagebox.showwarning("No file", "Please select a .docx file first.")
            return
        try:
            status_var.set("Parsing…")
            root.update_idletasks()
            txt, rpt = parse_docx_to_text2qti(path)
            output_text.set(txt)
            report_box.config(state="normal")
            report_box.delete("1.0", "end")
            report_box.insert("end", rpt)
            report_box.config(state="disabled")
            save_btn.config(state="normal")
            status_var.set("Done! Review the report below, then save your output.")
            status_lbl.config(foreground="green")
        except Exception as e:
            messagebox.showerror("Parse error", str(e))
            status_var.set("Error during parsing.")
            status_lbl.config(foreground="red")

    ttk.Button(pad, text="Parse document", command=run_parse).grid(row=3, column=0, sticky="ew", pady=(0, 12))

    ttk.Label(pad, text="Parse report:").grid(row=4, column=0, sticky="w", pady=(0, 4))
    report_frame = ttk.Frame(pad)
    report_frame.grid(row=5, column=0, sticky="nsew", pady=(0, 12))
    pad.rowconfigure(5, weight=1)
    report_frame.columnconfigure(0, weight=1)
    report_frame.rowconfigure(0, weight=1)

    report_box = tk.Text(report_frame, height=12, font=("Courier", 10),
                         state="disabled", wrap="none", relief="flat",
                         background="#f4f4f4", foreground="#333")
    report_box.grid(row=0, column=0, sticky="nsew")
    sb = ttk.Scrollbar(report_frame, orient="vertical", command=report_box.yview)
    sb.grid(row=0, column=1, sticky="ns")
    report_box.config(yscrollcommand=sb.set)

    def save_output():
        src = selected_path.get()
        default_name = Path(src).stem + "_text2qti.txt"
        dest = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=default_name,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if dest:
            Path(dest).write_text(output_text.get(), encoding="utf-8")
            status_var.set(f"Saved → {Path(dest).name}")
            status_lbl.config(foreground="green")

    save_btn = ttk.Button(pad, text="Save text2qti output…", command=save_output, state="disabled")
    save_btn.grid(row=6, column=0, sticky="ew")

    root.mainloop()


# ─── Entry point ────────────────────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        launch_gui()
    else:
        path = sys.argv[1]
        txt, rpt = parse_docx_to_text2qti(path)
        out = Path(Path(path).stem + "_text2qti.txt")
        out.write_text(txt, encoding="utf-8")
        print(rpt)
        print(f"\nOutput: {out}")
