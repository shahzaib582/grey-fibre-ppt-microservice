"""
PASS 3 — Generate Transition Slides
====================================
For each major section (Mood, Favorability, Ballot, etc.):
  - Detect section divider slides
  - Insert exactly 2 transition slides after each divider:
    Slide A: "{Section Name}: Questions Asked" — bullet summary, no numbers
    Slide B: "{Section Name}: Survey Responses" — synthesis with key percentages
  - Content is LLM-generated but data-bound

Usage:
  python -m survey_pipeline.pass3_transition_slides --data DATA.xlsx --pptx output_pass2.pptx --out FINAL.pptx
"""

import argparse
import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import PP_PLACEHOLDER_TYPE
from lxml import etree
from pptx.oxml.ns import qn

from .utils import (
    SECTION_NAMES,
    is_section_divider,
    get_section_questions,
    get_slide_text,
    parse_question_spec,
    get_question_ids,
    generate_questions_asked_content,
    generate_survey_responses_content,
    call_llm,
    ensure_key_finding_style,
    apply_style_to_run,
)
from .data_loader import load_ai_long


def get_slide_layout(prs):
    """
    Get the content-style layout for transition slides (blue/white theme).
    Prefer layouts used by content slides, NOT the section divider layout.
    """
    preferred_names = [
        "Title and Content",
        "Title Only",
        "Content with Caption",
        "Blank",
        "Custom Layout",
    ]
    for layout in prs.slide_layouts:
        if layout.name in preferred_names:
            return layout
    # Use first non–section-header layout if possible (often index 1 or 2 is content)
    for idx in (1, 2, 3, 0):
        if idx < len(prs.slide_layouts):
            return prs.slide_layouts[idx]
    return prs.slide_layouts[0]


def create_transition_slide(prs, slide_index, title_text, body_text, ref_slide=None, key_style=None):
    """
    Create a new transition slide using the deck's content layout.
    Uses the template's title and body placeholders so fonts/colors
    match the existing content slides.
    """
    layout = get_slide_layout(prs)
    slide = prs.slides.add_slide(layout)

    # Prefer using the layout's title + body placeholders so we inherit
    # font, color, and spacing from the template.
    title_shape = None
    body_shape = None

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if getattr(shape, "is_placeholder", False):
            ph_type = shape.placeholder_format.type
            if ph_type in (PP_PLACEHOLDER_TYPE.TITLE, PP_PLACEHOLDER_TYPE.CENTER_TITLE):
                title_shape = shape
            elif ph_type in (PP_PLACEHOLDER_TYPE.BODY, PP_PLACEHOLDER_TYPE.CONTENT) and body_shape is None:
                body_shape = shape

    # Fallbacks: first text shape as title, second as body
    if title_shape is None:
        for shape in slide.shapes:
            if shape.has_text_frame:
                title_shape = shape
                break
    if body_shape is None:
        for shape in slide.shapes:
            if shape.has_text_frame and shape is not title_shape:
                body_shape = shape
                break

    # If the chosen layout has no usable text shapes (e.g. Blank layout),
    # fall back to creating our own title/body boxes. Size them based on
    # the actual slide width so margins match other slides.
    slide_width = prs.slide_width
    left = Inches(0.5)
    right_margin = Inches(0.5)
    width = slide_width - left - right_margin
    title_top = Inches(0.55)
    title_height = Inches(0.5)
    body_top = Inches(1.15)
    body_height = Inches(5.5)

    if title_shape is None or not getattr(title_shape, "has_text_frame", False):
        title_shape = slide.shapes.add_textbox(left, title_top, width, title_height)
    if body_shape is None or not getattr(body_shape, "has_text_frame", False):
        body_shape = slide.shapes.add_textbox(left, body_top, width, body_height)

    # Set title text
    if title_shape is not None and getattr(title_shape, "has_text_frame", False):
        tf = title_shape.text_frame
        tf.text = ""
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title_text
        # Apply key-finding style to the title as well, but force bold
        if p.runs and key_style:
            apply_style_to_run(p.runs[0], key_style, force_bold=True)
            # Override size to 28pt as requested
            p.runs[0].font.size = Pt(28)

    # Set body content
    if body_shape is not None and getattr(body_shape, "has_text_frame", False):
        tf = body_shape.text_frame
        tf.text = ""
        tf.word_wrap = True
        _set_body_content(tf, body_text, key_style)

    return slide


def _set_body_content(text_frame, body_text, key_style=None):
    """Set body content with proper formatting. Handles bullet points."""
    lines = body_text.strip().split("\n")

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        # Determine if this is a bullet point
        is_bullet = line.startswith("•") or line.startswith("-") or line.startswith("*")
        if is_bullet:
            # Remove bullet character
            line = line.lstrip("•-* ").strip()

        if i == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()

        p.text = line
        p.space_after = Pt(6)

        if is_bullet:
            p.level = 0
            # Add bullet character back for visual
            p.text = f"• {line}"

        # Apply key-finding style (same as main bullets) if available
        if p.runs and key_style:
            apply_style_to_run(p.runs[0], key_style)


def move_slide_to_index(prs, slide, target_index):
    """Move a slide to a specific index position in the presentation."""
    sldIdLst = prs.slides._sldIdLst

    # Find the slide's rId
    slide_rel_id = None
    for rel in prs.part.rels.values():
        if rel.target_part == slide.part:
            slide_rel_id = rel.rId
            break

    if slide_rel_id is None:
        return

    # Find and remove the sldId entry
    target_entry = None
    for sldId in sldIdLst:
        if sldId.get(qn('r:id')) == slide_rel_id:
            target_entry = sldId
            break

    if target_entry is None:
        return

    sldIdLst.remove(target_entry)

    # Insert at the target position
    entries = list(sldIdLst)
    if target_index >= len(entries):
        sldIdLst.append(target_entry)
    else:
        entries[target_index].addprevious(target_entry)


def main():
    ap = argparse.ArgumentParser(description="Pass 3: Generate transition slides")
    ap.add_argument(
        "--data",
        required=True,
        help="Path to survey Excel file (either AI-ready with 'ai_long' sheet or raw 250870-style 'ExcelData')",
    )
    ap.add_argument("--pptx", required=True, help="Path to Pass 2 output")
    ap.add_argument("--out", required=True, help="Output file path")
    args = ap.parse_args()

    print("=" * 60)
    print("PASS 3 — Generate Transition Slides")
    print("=" * 60)
    print(f"Data:     {args.data}")
    print(f"Input:    {args.pptx}")
    print(f"Output:   {args.out}")
    print()

    ai_long = load_ai_long(args.data)
    prs = Presentation(args.pptx)

    key_style = ensure_key_finding_style(prs)

    original_count = len(prs.slides)
    print(f"Original slide count: {original_count}")
    print()

    # Step 1: Find all section dividers and their data
    sections_found = []
    slides_list = list(prs.slides)

    for i, slide in enumerate(slides_list):
        section_name = is_section_divider(slide)
        if section_name:
            section_data = get_section_questions(prs, i, ai_long)
            sections_found.append({
                "name": section_name,
                "slide_index": i,
                "slide": slide,
                "questions": section_data,
            })
            print(f"Found section: '{section_name}' at slide {i + 1}")
            print(f"  Questions: {section_data['question_ids']}")

    print(f"\nTotal sections found: {len(sections_found)}")
    print()

    if not sections_found:
        print("[WARN] No section dividers found. Nothing to do.")
        prs.save(args.out)
        return

    # Step 2: Generate transition slides (work backwards to preserve indices)
    slides_inserted = 0

    for section in reversed(sections_found):
        section_name = section["name"]
        q_data = section["questions"]
        divider_idx = section["slide_index"]

        if not q_data["question_ids"]:
            print(f"[SKIP] '{section_name}' — no questions found")
            continue

        print(f"Generating transition slides for '{section_name}'...")

        # Generate Slide B content first (Survey Responses)
        try:
            responses_content = generate_survey_responses_content(
                section_name, q_data["question_texts"], q_data["question_data"]
            )
            print(f"  [OK] Survey Responses content generated ({len(responses_content)} chars)")
        except Exception as e:
            print(f"  [ERROR] Failed to generate responses content: {e}")
            responses_content = "• Survey response data available for review."

        # Generate Slide A content (Questions Asked)
        try:
            questions_content = generate_questions_asked_content(
                section_name, q_data["question_texts"]
            )
            print(f"  [OK] Questions Asked content generated ({len(questions_content)} chars)")
        except Exception as e:
            print(f"  [ERROR] Failed to generate questions content: {e}")
            questions_content = "• Questions were asked in this section."

        # Create Slide B (Survey Responses) — inserted first so it ends up AFTER Slide A
        slide_b = create_transition_slide(
            prs,
            divider_idx + 1,
            f"{section_name}: Survey Responses",
            responses_content,
            ref_slide=section["slide"],
            key_style=key_style,
        )
        move_slide_to_index(prs, slide_b, divider_idx + 1)
        slides_inserted += 1

        # Create Slide A (Questions Asked)
        slide_a = create_transition_slide(
            prs,
            divider_idx + 1,
            f"{section_name}: Questions Asked",
            questions_content,
            ref_slide=section["slide"],
            key_style=key_style,
        )
        move_slide_to_index(prs, slide_a, divider_idx + 1)
        slides_inserted += 1

        print(f"  [OK] 2 transition slides inserted after '{section_name}'")

    # Step 3: Save
    prs.save(args.out)

    final_count = len(prs.slides)
    print()
    print(f"PASS 3 complete.")
    print(f"  Sections processed: {len(sections_found)}")
    print(f"  Slides inserted:    {slides_inserted}")
    print(f"  Original count:     {original_count}")
    print(f"  Final count:        {final_count}")
    print(f"  Expected count:     {original_count + slides_inserted}")
    print(f"  Output:             {args.out}")

    # Validation
    if final_count != original_count + slides_inserted:
        print(f"\n  [WARN] Slide count mismatch! Expected {original_count + slides_inserted}, got {final_count}")
    else:
        print(f"\n  [OK] Slide count validated [OK]")


if __name__ == "__main__":
    main()
