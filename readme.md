# DocFill — AI Supplier Document Extractor

Streamlit web app: upload AI supplier onboarding PDFs (IBAN letter + Emirates ID),
**Gemini Vision** extracts every field, you review/edit, then export to the
pre-formatted Arabic/English supplier Excel template (`قالب تحديث`).

No Tesseract, poppler, or pdf2image — **PyMuPDF** (pure Python) handles PDFs, so it
deploys to Streamlit Community Cloud with zero system packages.

## Project structure

```
app.py          # Streamlit premium UI (dashboard + history views)
extractor.py    # PyMuPDF PDF→images + Gemini Vision (google-genai SDK, model fallback)
validator.py    # IBAN MOD-97, SWIFT, Emirates-ID, field status, bank Arabic map
exporter.py     # Fills the supplier Excel template (append or overwrite)
history.py      # Persistent history — Google Sheets (Apps Script) or local SQLite
apps_script/
  Code.gs                # paste into Google Apps Script for the Sheets backend
requirements.txt
.streamlit/
  config.toml            # theme
  secrets.toml           # secrets (gitignored — copy from .example)
assets/supplier_template.xlsx   # bundled default template
data/                    # local SQLite history (gitignored)
docs/                    # sample PDF + Excel template (gitignored)
```

## Quick start (local)

```bash
pip install -r requirements.txt
cp .streamlit/secrets.toml.example .streamlit/secrets.toml   # then paste your key
streamlit run app.py
```

## API key

Get one at <https://aistudio.google.com> → **Get API key**.

- **Local:** `.streamlit/secrets.toml` → `GEMINI_API_KEY = "..."` (or `GOOGLE_API_KEY`).
- **Streamlit Cloud:** App → Settings → Secrets → paste the same line.

## History (Dashboard / History tabs)

Saved records are kept persistently and shown in the **History** view (search,
stats, re-export to Excel, clear). The backend is **auto-selected**:

| Condition | Backend | Notes |
|-----------|---------|-------|
| `GSHEET_WEBAPP_URL` set in secrets | **Google Sheets** | permanent in cloud *and* local |
| not set | **local SQLite** (`data/history.db`) | per-machine, survives restarts |

⚠️ On Streamlit Community Cloud the filesystem is ephemeral, so SQLite history
resets on reboot — use the Google Sheets backend there.

### Google Sheets via Apps Script (no service account, no OAuth)

1. Open a Google Sheet → **Extensions ▸ Apps Script**, paste `apps_script/Code.gs`.
2. Set `TOKEN` in the script to a long random string.
3. **Deploy ▸ New deployment ▸ Web app** — *Execute as: Me*, *Who has access: Anyone*. Copy the `/exec` URL.
4. Add to `.streamlit/secrets.toml` (or Streamlit Cloud secrets):
   ```toml
   GSHEET_WEBAPP_URL = "https://script.google.com/macros/s/XXXX/exec"
   GSHEET_TOKEN      = "the-same-TOKEN-from-Code.gs"
   ```

DocFill then reads/writes history straight to your Sheet over HTTPS.

## Model

Primary: `gemini-2.5-flash`, with automatic fallback to `gemini-2.5-flash-lite`
then `gemini-2.0-flash` if a model is overloaded/rate-limited (retry + backoff).

## Excel mapping (sheet "Supplier Template", data from row 3)

| Col | Field            | Col | Field             |
|-----|------------------|-----|-------------------|
| A   | auto-number      | K   | bank name (Arabic)|
| B   | Arabic name      | M   | account number    |
| C   | English name     | N   | IBAN              |
| G   | Emirate          | O   | account title (English name) |

Rows 1–2 are the English/Arabic header rows and are left untouched.
