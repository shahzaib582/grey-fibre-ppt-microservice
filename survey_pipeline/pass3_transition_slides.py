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
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
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


def create_transition_slide(prs, slide_index, title_text, body_text, ref_slide=None):
    """
    Create a new transition slide using the deck's content theme (blue/white).
    Title is placed at top-left; body below so they never overlap.
    Does NOT copy the section divider's background (avoids orange theme).
    """
    layout = get_slide_layout(prs)
    slide = prs.slides.add_slide(layout)

    # Fixed positions: title at top-left, body below (matches "Research Methodology" style)
    title_top = Inches(0.55)
    title_height = Inches(0.5)
    body_top = Inches(1.15)
    body_height = Inches(5.5)
    left = Inches(0.5)
    width = Inches(9)

    # Clear any placeholder text from the layout so our boxes are the only content
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text.strip():
            try:
                shape.text_frame.clear()
            except Exception:
                pass

    # Title textbox at top-left (dark grey/black to match template)
    title_shape = slide.shapes.add_textbox(left, title_top, width, title_height)
    tf = title_shape.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0x2C, 0x2C, 0x2C)

    # Body textbox below title
    body_shape = slide.shapes.add_textbox(left, body_top, width, body_height)
    tf = body_shape.text_frame
    tf.word_wrap = True
    _set_body_content(tf, body_text)

    return slide


def _set_body_content(text_frame, body_text):
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
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
        p.space_after = Pt(6)

        if is_bullet:
            p.level = 0
            # Add bullet character back for visual
            p.text = f"• {line}"


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
