# Parchi AI

Parchi AI is a local-first Streamlit application for supplier-document extraction, verification, reconciliation, and export.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Set `GEMINI_API_KEY` or add API keys in the application settings to enable Gemini extraction. Uploads are processed one document at a time; PDFs are chunked by page, and XLSX workbooks include hidden worksheets when they are readable. Local runtime databases and uploaded documents are intentionally ignored by Git.
