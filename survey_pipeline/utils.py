"""
Shared utilities for the 3-Step Survey Slide Automation pipeline.
"""

import re
import copy
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

PLACEHOLDER = "{Insert Finding Here}"

# ── Section divider names (order matters for detection) ──
SECTION_NAMES = [
    "Mood",
    "Favorability",
    "Ballot",
    "Positioning",
    "Pro Gill Messages",
    "Anti Gill Messages",
    "Anti Malinowski Messages",
    "Demographics",
]

# Cached style for key-finding text (from the placeholder in the template)
KEY_FINDING_STYLE: dict | None = None


# ═══════════════════════════════════════════════════════════
#  Question spec parsing
# ═══════════════════════════════════════════════════════════

def parse_question_spec(slide_text: str):
    """
    Detect question specification from slide text.
    Returns:
        ("single", int)         — e.g. Question 5:
        ("range", (int, int))   — e.g. Questions 6-16:
        None                    — no question detected
    """
    # Try range pattern first (Questions X-Y:)
    m = re.search(r"\bQuestions?\s+(\d+)\s*[-–—]\s*(\d+)\s*:", slide_text, flags=re.IGNORECASE)
    if m:
        a, b = int(m.group(1)), int(m.group(2))
        if a > b:
            a, b = b, a
        return ("range", (a, b))

    # Single question pattern (Question N:)
    m = re.search(r"\bQuestion\s+(\d+)\s*:", slide_text, flags=re.IGNORECASE)
    if m:
        return ("single", int(m.group(1)))

    return None


def get_question_ids(qspec):
    """Convert a question spec to a list of question IDs (e.g. ['Q6', 'Q7', ...])."""
    if qspec is None:
        return []
    if qspec[0] == "single":
        return [f"Q{qspec[1]}"]
    elif qspec[0] == "range":
        a, b = qspec[1]
        return [f"Q{i}" for i in range(a, b + 1)]
    return []


# ═══════════════════════════════════════════════════════════
#  Data selection
# ═══════════════════════════════════════════════════════════

def select_top_rows(ai_long: pd.DataFrame, qid: str, top_k: int = 3, exclude_net: bool = True):
    """Select top K answer options for a given question ID."""
    dfq = ai_long[ai_long["question_id"] == qid].copy()
    if dfq.empty:
        return dfq

    if exclude_net and "is_net" in dfq.columns:
        non_net = dfq[dfq["is_net"] == False]
        if not non_net.empty:
            dfq = non_net

    if "rank_pct_desc" in dfq.columns:
        dfq = dfq.sort_values(["rank_pct_desc", "pct"], ascending=[True, False])
    else:
        dfq = dfq.sort_values(["pct"], ascending=False)

    return dfq.head(top_k)


def select_top_rows_multi(ai_long: pd.DataFrame, question_ids: list, top_k: int = 3, exclude_net: bool = True):
    """
    Select top K answer options across multiple questions.
    For multi-question slides, returns data grouped by question.
    """
    all_rows = []
    for qid in question_ids:
        rows = select_top_rows(ai_long, qid, top_k=top_k, exclude_net=exclude_net)
        if not rows.empty:
            all_rows.append(rows)

    if not all_rows:
        return pd.DataFrame()

    return pd.concat(all_rows, ignore_index=True)


def format_values(rows: pd.DataFrame, pct_decimals: int = 0):
    """Format rows as 'Option – XX%; Option – XX%; Option – XX%.'"""
    parts = []
    for _, r in rows.iterrows():
        opt = str(r["answer_option"]).strip()
        pct = float(r["pct"])
        if pct_decimals == 0:
            pct_str = f"{round(pct):.0f}%"
        else:
            pct_str = f"{pct:.{pct_decimals}f}%"
        parts.append(f"{opt} – {pct_str}")
    return "; ".join(parts) + "."


def format_values_grouped(ai_long: pd.DataFrame, question_ids: list, top_k: int = 3,
                           exclude_net: bool = True, pct_decimals: int = 0):
    """
    Format top K values for multiple questions, grouped by question.
    Returns a combined string with question labels.
    """
    parts = []
    for qid in question_ids:
        rows = select_top_rows(ai_long, qid, top_k=top_k, exclude_net=exclude_net)
        if rows.empty:
            continue
        q_text = rows.iloc[0]["question_text"]
        # Shorten question text for display
        short_q = q_text[:80] + "..." if len(q_text) > 80 else q_text
        vals = format_values(rows, pct_decimals=pct_decimals)
        parts.append(f"{qid}: {vals}")

    return "\n".join(parts) if parts else ""


# ═══════════════════════════════════════════════════════════
#  Slide text extraction
# ═══════════════════════════════════════════════════════════

def get_slide_text(slide) -> str:
    """Extract all text from a slide."""
    parts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text
            if t:
                parts.append(t)
    return "\n".join(parts)


def slide_has_placeholder(slide) -> bool:
    """Check if slide contains the placeholder text."""
    return PLACEHOLDER in get_slide_text(slide)


# ═══════════════════════════════════════════════════════════
#  Text replacement with formatting preservation
# ═══════════════════════════════════════════════════════════

def _copy_run_format(src_run, dst_run):
    """Copy font formatting from source run to destination run."""
    try:
        if src_run.font.name:
            dst_run.font.name = src_run.font.name
        if src_run.font.size:
            dst_run.font.size = src_run.font.size
        if src_run.font.bold is not None:
            dst_run.font.bold = src_run.font.bold
        if src_run.font.italic is not None:
            dst_run.font.italic = src_run.font.italic
        if src_run.font.color and src_run.font.color.rgb:
            dst_run.font.color.rgb = src_run.font.color.rgb
    except Exception:
        pass  # If we can't copy some attr, continue gracefully


def ensure_key_finding_style(prs: Presentation) -> dict:
    """
    Inspect the deck once to capture the font family, size, and color
    from the {Insert Finding Here} placeholder. This is the style we
    want for key-finding sentences and related bullets.
    """
    global KEY_FINDING_STYLE
    if KEY_FINDING_STYLE is not None:
        return KEY_FINDING_STYLE

    style: dict = {}
    for slide in prs.slides:
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            tf = shape.text_frame
            if not tf.text or PLACEHOLDER not in tf.text:
                continue
            for para in tf.paragraphs:
                if para.runs:
                    r0 = para.runs[0]
                    f = r0.font
                    style = {
                        "name": f.name,
                        "size": f.size,
                        "bold": f.bold,
                        "italic": f.italic,
                        "rgb": getattr(getattr(f, "color", None), "rgb", None),
                    }
                    break
            if style:
                break
        if style:
            break

    KEY_FINDING_STYLE = style
    return KEY_FINDING_STYLE


def apply_style_to_run(run, style: dict, force_bold: bool | None = None) -> None:
    """Apply a captured style dict to a run. No-op if style is empty."""
    if not style:
        return
    f = run.font
    name = style.get("name")
    size = style.get("size")
    bold = style.get("bold")
    italic = style.get("italic")
    rgb = style.get("rgb")

    if name:
        f.name = name
    if size:
        f.size = size
    if force_bold is not None:
        f.bold = force_bold
    elif bold is not None:
        f.bold = bold
    if italic is not None:
        f.italic = italic
    if rgb:
        try:
            f.color.rgb = rgb
        except Exception:
            # If color is theme-based or invalid, ignore
            pass


def replace_placeholder_in_shape(shape, new_text: str) -> bool:
    """
    Replace {Insert Finding Here} placeholder with new_text,
    preserving the original text formatting as much as possible.
    Handles multi-line new_text by creating additional paragraphs.
    """
    if not shape.has_text_frame:
        return False

    tf = shape.text_frame
    full_text = tf.text
    if PLACEHOLDER not in full_text:
        return False

    # Find the paragraph and run containing the placeholder
    for para_idx, para in enumerate(tf.paragraphs):
        para_text = para.text
        if PLACEHOLDER not in para_text:
            continue

        # Capture formatting from the first run (or default)
        font_name = None
        font_size = None
        font_bold = None
        font_italic = None
        font_color = None
        alignment = para.alignment

        if para.runs:
            ref_run = para.runs[0]
            font_name = ref_run.font.name
            font_size = ref_run.font.size
            font_bold = ref_run.font.bold
            font_italic = ref_run.font.italic
            try:
                font_color = ref_run.font.color.rgb
            except Exception:
                pass

        # Handle multi-line replacement
        lines = new_text.split("\n")

        # Clear ALL runs in this paragraph
        for run in list(para.runs):
            r_elem = run._r
            r_elem.getparent().remove(r_elem)

        # Set first line in the existing paragraph
        run = para.add_run()
        run.text = lines[0]
        if font_name:
            run.font.name = font_name
        if font_size:
            run.font.size = font_size
        if font_bold is not None:
            run.font.bold = font_bold
        if font_italic is not None:
            run.font.italic = font_italic
        if font_color:
            run.font.color.rgb = font_color

        # Add additional paragraphs for remaining lines
        from pptx.oxml.ns import qn
        for line in lines[1:]:
            new_p = copy.deepcopy(para._p)
            # Clear runs in the copied paragraph
            for r in new_p.findall(qn('a:r')):
                new_p.remove(r)
            # Add a run with the text
            new_r = copy.deepcopy(run._r)
            new_r.text = line
            new_p.append(new_r)
            # Insert after current paragraph
            para._p.addnext(new_p)

        return True

    return False


def set_shape_text_to_single_paragraph(shape, text: str, style: dict | None = None) -> bool:
    """
    Set the shape's text frame to a single paragraph with the given text.
    Removes all other paragraphs (avoids long stats overlapping tables).
    If a style dict is provided, apply it to the run (e.g. key-finding style).
    """
    if not shape.has_text_frame:
        return False

    from pptx.oxml.ns import qn

    tf = shape.text_frame
    txBody = tf._txBody

    # Keep only the first paragraph in the text frame
    paragraphs = list(txBody.findall(qn("a:p")))
    for p in paragraphs[1:]:
        txBody.remove(p)

    first_p = txBody.find(qn("a:p"))
    if first_p is None:
        return False

    # Remove all existing runs in the first paragraph
    for r in list(first_p.findall(qn("a:r"))):
        first_p.remove(r)

    # python-pptx paragraph API for adding a run
    p0 = tf.paragraphs[0]
    run = p0.add_run()
    run.text = text

    if style:
        apply_style_to_run(run, style)

    return True


# ═══════════════════════════════════════════════════════════
#  Section detection
# ═══════════════════════════════════════════════════════════

def is_section_divider(slide) -> str | None:
    """
    Check if a slide is a section divider.
    Returns the section name if it is, None otherwise.

    Section dividers are identified as slides with a single large title
    matching one of the known section names.
    """
    texts = []
    for shape in slide.shapes:
        if shape.has_text_frame:
            t = shape.text_frame.text.strip()
            if t and not t.isdigit():  # Ignore slide numbers
                texts.append(t)

    if not texts:
        return None

    # Check if any text matches a known section name
    for text in texts:
        clean = text.strip()
        for section in SECTION_NAMES:
            if clean.lower() == section.lower():
                return section

    return None


def get_section_questions(prs, section_slide_idx: int, ai_long: pd.DataFrame) -> dict:
    """
    Get all questions belonging to a section.
    A section spans from the section divider slide to the next section divider (or end).

    Returns dict with:
        - question_ids: list of question IDs
        - question_data: DataFrame of all data for these questions
        - question_texts: dict mapping qid -> question_text
    """
    question_ids = []
    slides = list(prs.slides)

    # Scan forward from section divider until next divider or end
    for i in range(section_slide_idx + 1, len(slides)):
        slide = slides[i]

        # Stop if we hit another section divider
        if is_section_divider(slide):
            break

        slide_text = get_slide_text(slide)
        qspec = parse_question_spec(slide_text)
        if qspec:
            qids = get_question_ids(qspec)
            question_ids.extend(qids)

    # Remove duplicates while preserving order
    seen = set()
    unique_qids = []
    for qid in question_ids:
        if qid not in seen:
            seen.add(qid)
            unique_qids.append(qid)

    # Get data
    question_data = ai_long[ai_long["question_id"].isin(unique_qids)].copy()

    # Get question texts
    question_texts = {}
    for qid in unique_qids:
        rows = ai_long[ai_long["question_id"] == qid]
        if not rows.empty:
            question_texts[qid] = rows.iloc[0]["question_text"]

    return {
        "question_ids": unique_qids,
        "question_data": question_data,
        "question_texts": question_texts,
    }


# ═══════════════════════════════════════════════════════════
#  LLM helpers
# ═══════════════════════════════════════════════════════════

def call_llm(system_prompt: str, user_prompt: str, model: str = "gpt-4.1-mini", temperature: float = 0.2) -> str:
    """Call OpenAI-compatible LLM. Returns response text."""
    import os
    from openai import OpenAI

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise ValueError("OPENAI_API_KEY environment variable not set")

    client = OpenAI(api_key=api_key)
    resp = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=temperature,
    )
    return resp.choices[0].message.content.strip()


def generate_restatement(bullets: str) -> str:
    """Generate a single restatement sentence from bullet points."""
    system = (
        "You are an executive-level political survey analyst writing for senior decision-makers. "
        "Write exactly ONE expressive sentence that synthesizes the key story in the data. "
        "Use an analytical, research-oriented tone: explain what the finding implies or why it matters, "
        "not just restate the numbers. Use comparative, interpretive phrasing (e.g., 'is rated X while Y...'), "
        "stay factual and executive-neutral, and must keep under 35 words. Do not invent any numbers."
    )
    user = f"""Write EXACTLY ONE sentence (max 35 words) summarizing these survey results.

Match the analytical flow of a research finding: purpose/importance and what it implies, not just survey metadata.
Describe who is higher/lower, what stands out most, and any clear patterns—in a way that supports interpretation and decisions.
Do not add new numbers; only use the numbers already shown.

{bullets}"""
    return call_llm(system, user)


def generate_questions_asked_content(section_name: str, question_texts: dict) -> str:
    """Generate 'Questions Asked' transition slide content."""
    q_list = "\n".join([f"- {qid}: {text}" for qid, text in question_texts.items()])

    system = (
        "You are a political survey analyst writing for senior decision-makers. "
        "Explain, in analytic language, what this section of the survey is measuring and why it matters. "
        "Do NOT restate the exact question wording; summarize concepts."
    )
    user = f"""You are preparing the **"{{section_name}}: Questions Asked"** transition slide for a survey deck.

Section name: "{section_name}"

Questions in this section (IDs and abbreviated text):
{q_list}

Write 5–6 bullet points that follow THIS conceptual sequence:
1. **Section purpose** – what this section is designed to measure and why (high-level intent).
2. **Measurement approach** – how the questions collectively capture that concept (dimensions/topics being probed).
3. **Measurement refinement** – how scales, frequency, intensity, or framing sharpen interpretation (e.g., strength of feeling, tradeoffs, behavior vs attitude).
4. **Analytical importance** – why these questions matter for understanding voters/respondents in this context.
5. **Segmentation value** – how responses enable segmentation (e.g., by attitudes, behavior, intensity, demographics).
6. **Context for later findings** – how these questions set up or contextualize findings that appear later in the deck.

Guidelines:
- Focus on analytic **content structure**, not formatting. Write each bullet as a short narrative sentence that explains the concept—not a fragment or survey-question label.
- NO numeric values and NO mention of question numbers (Q1, Q2, etc.).
- Write as clean bullet points (start each line with "• ").
- Keep total length under 900 characters."""

    return call_llm(system, user)


def generate_survey_responses_content(section_name: str, question_texts: dict,
                                       question_data: pd.DataFrame) -> str:
    """Generate 'Survey Responses' transition slide content."""
    # Build data summary for each question
    data_parts = []
    for qid, text in question_texts.items():
        qdata = question_data[question_data["question_id"] == qid]
        if qdata.empty:
            continue
        top = qdata.sort_values("pct", ascending=False).head(3)
        vals = "; ".join([f"{r['answer_option'].strip()} – {r['pct']:.0f}%" for _, r in top.iterrows()])
        data_parts.append(f"{qid}: {text[:100]}\n  Top results: {vals}")

    data_summary = "\n".join(data_parts)

    system = (
        "You are a political survey analyst preparing executive briefing slides. "
        "Summarize survey RESULTS in a structured, analytic way for senior readers. "
        "Stay strictly faithful to the numbers provided."
    )
    user = f"""You are preparing the **"{{section_name}}: Survey Responses"** transition slide.

Section name: "{section_name}"

Data for this section (per question, with top answer options and percentages):
{data_summary}

Write 5–6 bullet points that follow THIS conceptual sequence:
1. **Topline pattern** – what the results broadly show in this section (e.g., overall support, divide, concern).
2. **Key comparisons** – how major options or items compare (leaders vs laggards, strongest vs weakest responses).
3. **Intensity and distribution** – where responses are concentrated (e.g., strong vs soft support, extremes vs middle).
4. **Segment or subgroup differences (if implied by options)** – call out any clear splits between types of responses (e.g., positive vs negative, incumbents vs challengers, economic vs social issues).
5. **Implications for interpretation** – what these patterns mean for how to read this section (e.g., mandate strength, vulnerability, momentum).
6. **Forward-looking context** – how these results will inform later sections, messaging, or strategy.

Guidelines:
- Use key percentages from the data where helpful (e.g., “around 6 in 10”, or explicit % when clear).
- Do NOT invent any new numbers or options beyond what appears in the data summary above.
- Write as clean bullet points (start each line with "• ").
- Keep total length under 900 characters."""

    return call_llm(system, user)
