## 3‑Step Survey Slide Automation Pipeline

This project automates the creation of a **client‑ready PowerPoint deck** from a survey Excel file and a slide template. It runs a 3‑step pipeline:

1. **Pass 1 – Insert Top Numeric Findings**  
   Detects the question(s) on each slide, pulls the top answer options from the survey data, and replaces `{Insert Finding Here}` with formatted values (and updates charts where possible).

2. **Pass 2 – Add AI Restatement Sentence**  
   For each question slide, generates a single, concise, data‑bound summary sentence and replaces the numeric block with that sentence (no raw stats in the bullet copy).

3. **Pass 3 – Generate Transition Slides**  
   Detects major section dividers (Mood, Favorability, Ballot, etc.) and inserts exactly **two** transition slides after each:
   - `{Section Name}: Questions Asked`
   - `{Section Name}: Survey Responses`

The result is a clean, presentation‑ready PPTX with numbers, narrative bullets, and section transitions.

---

## Project Structure

```text
ppt client/
├── survey_pipeline/          # Main library / deployable app
│   ├── __init__.py
│   ├── api.py                # FastAPI API (POST /generate)
│   ├── run_pipeline.py       # Orchestrator for all 3 passes
│   ├── pass1_insert_numbers.py
│   ├── pass2_add_restatement.py
│   ├── pass3_transition_slides.py
│   ├── data_loader.py        # Excel → ai_long data
│   └── utils.py              # Shared helpers + LLM wrappers
├── run.py                    # Simple CLI entry point
├── requirements.txt          # Python dependencies
├── .env                      # OPENAI_API_KEY, etc. (not committed)
└── .gitignore                # Ignores env, caches, data, outputs
```

---

## How the Pipeline Works

### Data Loading (`data_loader.py`)

- Accepts either:
  - A raw **`ExcelData`** crosstab (250870‑style).
- Normalizes both into an **`ai_long` DataFrame** with:
  - `question_id` (e.g. `Q18`)
  - `question_text`
  - `answer_option`
  - `pct` (percentage)
- Handles question headers like **`Question 18:`** and **`Q18` / `Q 18` / `Q18:`**, and **BASE rows** starting with `BASE=` or `BASE:` (e.g. `BASE: DON'T KNOW / REF`).

### Pass 1 – Insert Top Numbers (`pass1_insert_numbers.py`)

- Scans slides for `{Insert Finding Here}`.
- Parses **single** (`Question 18:`) and **range** (`Questions 6–16:`) specs from slide text.
- For each question:
  - Selects top‑K answer options (excluding `NET` rows by default).
  - Formats them as:  
    `Option – XX%; Option – XX%; Option – XX%.`
  - Replaces the placeholder on the slide, and updates simple bar/column charts when present.
- If a question has no data, inserts: **“No data available for this question.”**

### Pass 2 – AI Restatement (`pass2_add_restatement.py`)

- For each question slide (after Pass 1):
  - Collects top options and percentages from `ai_long`.
  - Builds a bullet list for the LLM, e.g.  
    `- Candidate A – 51%`
  - Calls `generate_restatement(...)` in `utils.py`:
    - Exactly **one sentence**, ≤ 35 words.
    - Executive‑neutral, no hallucinated numbers.
- Output behavior:
  - If placeholder is still present, it is replaced with **just the sentence**.
  - If Pass 1 values are present, the entire text shape is replaced with **only the restatement sentence**, so bullets don’t overflow or overlap tables.

### Pass 3 – Transition Slides (`pass3_transition_slides.py`)

- Uses `utils.SECTION_NAMES` and `is_section_divider(...)` to find section headers like:
  - `Mood`, `Favorability`, `Ballot`, `Positioning`, `Pro Gill Messages`, `Anti Gill Messages`, `Anti Malinowski Messages`, `Demographics`.
- For each section:
  - Walks forward in the deck to find all question slides in the section.
  - Pulls `question_ids`, `question_texts`, and `question_data` from `ai_long`.
  - Calls LLM helpers in `utils.py` to generate:
    - **Questions Asked** content (no numbers, ≤ 1000 chars).
    - **Survey Responses** content (data‑bound, key percentages, ≤ 1000 chars).
- Inserts **two new content‑style slides** after the divider:
  - `{Section Name}: Questions Asked`
  - `{Section Name}: Survey Responses`
- Validates slide count and per‑section presence.

---

## Running as an API (FastAPI)

From project root, after installing dependencies:

```bash
pip install -r requirements.txt
set OPENAI_API_KEY=your-key-here
uvicorn survey_pipeline.api:app --host 0.0.0.0 --port 8000
```

Key endpoint:

- `POST /generate` (multipart form):
  - `data` – survey Excel (`ai_long` sheet **or** `ExcelData` sheet).
  - `template` – PPTX template with `{Insert Finding Here}` placeholders and section headers.
  - `output_name` – desired output filename (optional).

Returns the final PPTX as a file download.

---