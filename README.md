# Healthians Inter-Run Precision Populator

A Streamlit app that reads daily assay PDF reports and fills a supplied Excel precision template.

## Run locally

```bash
python -m venv .venv
# Windows:
.venv\\Scripts\\activate
# macOS/Linux:
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open the local URL shown by Streamlit (usually http://localhost:8501).

## Deploy on Streamlit Community Cloud

1. Create a GitHub repository and upload `app.py`, `requirements.txt`, and `README.md`.
2. Open https://share.streamlit.io/ and sign in with GitHub.
3. Choose **New app**, select the repository and branch, and set the main file path to `app.py`.
4. Deploy. Upload the Excel template and Day 1–Day 5 PDFs in the app.

## Input naming

PDF filenames should include the day number, for example `DAY 1 INTER ALL.pdf`, `DAY 2 INTER ALL.pdf` ... `DAY 5 INTER ALL.pdf`. The app leaves template date cells unchanged.

## Important limitations

- The current parser is designed for text-based Siemens assay reports whose results appear as a standalone value plus unit line (for example `91.3 ng/mL`). Scanned/image-only PDFs need OCR support, which is not included.
- The parser uses the first five result rows for each analyte section in each daily PDF. If your report contains more than five rows, the extras are not used.
- Review the output workbook and verify analyte mapping and values against the original PDFs before using it for formal precision validation.
