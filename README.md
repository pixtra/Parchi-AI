# Parchi AI

Parchi AI is a local-first Streamlit application for supplier-document extraction, verification, reconciliation, and export.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Set `GEMINI_API_KEY` or add API keys in the application settings to enable Gemini extraction. Local runtime databases and uploaded documents are intentionally ignored by Git.
