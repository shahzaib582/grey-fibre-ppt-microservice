"""
FastAPI wrapper for the 3-Step Survey Slide Automation pipeline.
Accepts Excel + PPTX uploads, runs Pass 1–3, returns the final PPTX.

Deploy this (e.g. Render, Railway, Fly.io) and point n8n Cloud HTTP Request to it.

Usage (from project root):
  pip install -r requirements.txt
  set OPENAI_API_KEY=your-key
  uvicorn survey_pipeline.api:app --host 0.0.0.0 --port 8000
"""

import os
import re
import sys
from typing import Optional
import tempfile
import shutil
from contextlib import asynccontextmanager

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import Response
from dotenv import load_dotenv

load_dotenv()

# Bump this string whenever you deploy a new version so you can
# verify which build your n8n workflow is hitting.
API_VERSION = "Version 5.0"


def _sanitize_filename(name: str) -> str:
    """Allow only safe characters for output filename."""
    if not name or not name.strip():
        return "Brendan_Gill_Enhanced_Final.pptx"
    name = name.strip()
    if not name.lower().endswith(".pptx"):
        name = name + ".pptx"
    # Remove path components and dangerous chars
    name = os.path.basename(name)
    name = re.sub(r"[^\w\s\-\.]", "", name)
    return name or "Brendan_Gill_Enhanced_Final.pptx"


@asynccontextmanager
async def lifespan(app: FastAPI):
    yield
    # Optional: cleanup on shutdown


app = FastAPI(
    title="Survey Slide Pipeline API",
    description="Runs 3-pass survey slide automation (numbers, restatements, transition slides) and returns the final PPTX.",
    version=API_VERSION,
    lifespan=lifespan,
)


@app.get("/")
async def root():
    return {
        "service": "Survey Slide Pipeline API",
        "version": API_VERSION,
        "docs": "/docs",
        "health": "/health",
        "generate": "POST /generate (multipart: data, template, output_name, exec_summary_goals?)",
    }


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/version")
async def version():
    """Simple version endpoint to confirm deployed build."""
    return {"version": API_VERSION}


@app.post("/generate")
async def generate(
    data: UploadFile = File(..., description="Survey Excel file (AI-ready 'ai_long' or raw 'ExcelData')"),
    template: UploadFile = File(..., description="PowerPoint template with {Insert Finding Here} placeholders"),
    output_name: str = Form("Brendan_Gill_Enhanced_Final.pptx"),
    exec_summary_goals: Optional[UploadFile] = File(None, description="Optional: text file with survey goals for executive summary (required for exec summary slides)"),
):
    """
    Run the full pipeline (Pass 1, 2, 3) and return the enhanced PPTX.
    Set OPENAI_API_KEY in the environment for Pass 2 and Pass 3.
    Provide exec_summary_goals to generate executive summary slides (12-15 slides after Executive Summary header).
    """
    out_filename = _sanitize_filename(output_name)

    # Use a temp directory that we clean up manually. On Windows, pandas/openpyxl
    # can keep file handles open a bit longer, so we ignore deletion errors.
    tmpdir = tempfile.mkdtemp(prefix="pipeline_")
    data_path = os.path.join(tmpdir, "data.xlsx")
    pptx_path = os.path.join(tmpdir, "template.pptx")
    out_path = os.path.join(tmpdir, out_filename)

    try:
        try:
            content = await data.read()
            if not content:
                raise HTTPException(status_code=400, detail="Uploaded data file is empty")
            with open(data_path, "wb") as f:
                f.write(content)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Invalid data file: {e}") from e

        try:
            content = await template.read()
            if not content:
                raise HTTPException(status_code=400, detail="Uploaded template file is empty")
            with open(pptx_path, "wb") as f:
                f.write(content)
        except Exception as e:
            raise HTTPException(status_code=400, detail=f"Invalid template file: {e}") from e

        # Save exec_summary_goals to temp file if provided
        exec_goals_path = None
        if exec_summary_goals:
            try:
                content = await exec_summary_goals.read()
                if content:
                    exec_goals_path = os.path.join(tmpdir, "exec_summary_goals.txt")
                    with open(exec_goals_path, "wb") as f:
                        f.write(content)
                    print(f"[API] Executive summary goals received ({len(content)} bytes)")
            except Exception as e:
                print(f"[API] Executive summary goals read failed: {e}")
        else:
            print("[API] No executive summary goals file uploaded (exec summary slides will be skipped)")

        # Run pipeline by invoking its main() with argv
        argv_orig = sys.argv
        try:
            cmd = [
                "run_pipeline.py",
                "--data", data_path,
                "--pptx", pptx_path,
                "--out", out_path,
                "--passes", "1,2,3",
            ]
            if exec_goals_path:
                cmd.extend(["--exec-summary-goals", exec_goals_path])
            sys.argv = cmd
            from survey_pipeline.run_pipeline import main as pipeline_main
            pipeline_main()
        except SystemExit as e:
            if e.code != 0:
                raise HTTPException(
                    status_code=500,
                    detail="Pipeline failed. Check OPENAI_API_KEY and file formats.",
                )
        finally:
            sys.argv = argv_orig

        if not os.path.isfile(out_path):
            raise HTTPException(status_code=500, detail="Pipeline did not produce an output file")

        with open(out_path, "rb") as f:
            body = f.read()
    finally:
        try:
            shutil.rmtree(tmpdir, ignore_errors=True)
        except Exception:
            # Best-effort cleanup only; don't break the response on Windows file locks.
            pass

    return Response(
        content=body,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f'attachment; filename="{out_filename}"'},
    )
