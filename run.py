"""
Entry point to run the full pipeline from project root.

  python run.py --data DATA.xlsx --pptx TEMPLATE.pptx --out FINAL.pptx

Equivalent to: python -m survey_pipeline.run_pipeline ...
"""
from survey_pipeline.run_pipeline import main

if __name__ == "__main__":
    main()
