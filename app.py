"""
DocFill — UAE supplier document extractor (Streamlit + Gemini Vision)
=====================================================================
Upload IBAN letters + Emirates ID PDFs → Gemini Vision extracts every field →
review / edit → export to the Arabic/English supplier Excel template.

Stack (per CLAUDE.md, final): Streamlit · PyMuPDF · Gemini Vision · openpyxl.
No Tesseract / poppler / pdf2image.

Run:   streamlit run app.py
Key:   .streamlit/secrets.toml  →  GEMINI_API_KEY = "..."
"""

import html
import os

import pandas as pd
import streamlit as st

import history
from extractor import GEMINI_MODEL, blank_record, extract_with_gemini, pdf_to_images
from exporter import existing_record_count, export_to_excel

history.init_db()

# Bundled supplier template — used by default so no upload is needed each time.
DEFAULT_TEMPLATE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "assets", "supplier_template.xlsx"
)
from validator import (
    count_ok,
    field_status,
    iban_diagnosis,
    is_export_ready,
    normalize_iban,
    validate_iban,
)

# ══════════════════════════════════════════════════════════════════════════
#  PAGE CONFIG + PREMIUM THEME CSS
# ══════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="DocFill",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

PREMIUM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&display=swap');

:root{
  --bg:#f6f7f9; --panel:#ffffff; --panel-2:#f1f3f6; --border:#e6e8ec;
  --txt:#0f172a; --txt-2:#64748b; --txt-3:#94a3b8; --brand:#2563eb;
  --ok-bg:#ecfdf5; --ok-tx:#15803d;
}
html, body, [class*="css"], .stApp{ font-family:'Inter',system-ui,sans-serif; }
.stApp{ background:var(--bg); }
.block-container{ padding-top:1.4rem; padding-bottom:3rem; max-width:1180px; }
#MainMenu, footer, header{ visibility:hidden; }

/* ---------- DARK SIDEBAR ---------- */
section[data-testid="stSidebar"]{ background:#0a0f1e; border-right:1px solid #1a2540; }
section[data-testid="stSidebar"] *{ color:#cbd5e1; }
section[data-testid="stSidebar"] .stMarkdown h1,
section[data-testid="stSidebar"] .stMarkdown h2,
section[data-testid="stSidebar"] .stMarkdown h3{ color:#e2e8f0; }
.sb-brand{ display:flex; align-items:center; gap:10px; padding:4px 0 14px;
  border-bottom:1px solid #1a2540; margin-bottom:14px; }
.sb-brand-icon{ width:34px; height:34px; background:#1d3a6e; border-radius:8px;
  display:flex; align-items:center; justify-content:center; font-size:17px; }
.sb-brand-name{ font-size:14px; font-weight:600; color:#e2e8f0; letter-spacing:-.01em; }
.sb-brand-sub{ font-size:10px; color:#475569; margin-top:1px; }
.sb-label{ font-size:9px; font-weight:600; text-transform:uppercase; letter-spacing:.11em;
  color:#475569; margin:6px 0 8px; }
.sb-ok{ display:flex; align-items:center; gap:6px; font-size:11px; color:#34d399;
  background:#022c22; padding:6px 10px; border-radius:6px; margin-top:6px; }
.sb-warn{ display:flex; align-items:center; gap:6px; font-size:11px; color:#fbbf24;
  background:#2e2305; padding:6px 10px; border-radius:6px; margin-top:6px; }
.sb-chip{ display:inline-flex; align-items:center; gap:5px; background:#0f2044;
  border:1px solid #1e3a6e; border-radius:20px; padding:4px 11px; font-size:11px;
  color:#60a5fa; }
.sb-deploy{ display:flex; align-items:center; gap:6px; font-size:10px; color:#475569;
  margin-top:10px; }
.sb-dot{ width:6px; height:6px; background:#34d399; border-radius:50%; }
section[data-testid="stSidebar"] hr{ border-color:#1a2540; }

/* ---------- TOP BAR ---------- */
.topbar{ display:flex; align-items:center; gap:10px; padding:9px 18px; background:#fff;
  border:1px solid var(--border); border-radius:12px; margin-bottom:18px; }
.crumb{ display:flex; align-items:center; gap:7px; font-size:12.5px; color:var(--txt-2); }
.crumb .dim{ color:var(--txt-3); }
.engine-pill{ margin-left:auto; display:flex; align-items:center; gap:6px; font-size:11.5px;
  background:#eff6ff; border:1px solid #bfdbfe; border-radius:20px; padding:4px 12px; color:#1d4ed8; }

/* ---------- PAGE HEAD ---------- */
.page-title{ font-size:22px; font-weight:600; color:var(--txt); letter-spacing:-.02em; }
.page-sub{ font-size:13px; color:var(--txt-2); line-height:1.55; margin-top:3px; }

/* ---------- METRIC CARDS ---------- */
.metrics{ display:grid; grid-template-columns:repeat(4,1fr); gap:12px; margin:18px 0 22px; }
.metric{ background:var(--panel); border:1px solid var(--border); border-radius:14px; padding:15px 17px; }
.metric-label{ font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:.07em;
  color:var(--txt-3); margin-bottom:8px; }
.metric-val{ font-size:25px; font-weight:600; color:var(--txt); letter-spacing:-.02em; line-height:1; }
.metric-sub{ font-size:10.5px; color:var(--txt-2); margin-top:6px; }
.badge{ display:inline-flex; align-items:center; gap:4px; font-size:10.5px; padding:3px 9px;
  border-radius:11px; margin-top:8px; font-weight:500; }
.b-green{ background:var(--ok-bg); color:var(--ok-tx); }
.b-blue{ background:#eff6ff; color:#1d4ed8; }
.b-amber{ background:#fffbeb; color:#b45309; }

/* ---------- STATUS CHECKLIST (inside card) ---------- */
.status-box{ background:var(--panel-2); border:1px solid var(--border); border-radius:12px;
  padding:14px 15px; height:100%; }
.status-title{ font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:.07em;
  color:var(--txt-3); margin-bottom:11px; }
.srow{ display:flex; align-items:center; gap:7px; padding:3px 0; font-size:12px; color:var(--txt-2); }
.srow .ic-ok{ color:#10b981; font-weight:700; }
.srow .ic-wa{ color:#f59e0b; font-weight:700; }
.gem-dot{ width:6px; height:6px; background:#f43f5e; border-radius:50%; margin-left:auto; }
.gem-note{ margin-top:11px; display:flex; align-items:center; gap:6px; font-size:10.5px; color:var(--txt-3); }

/* ---------- EDITABLE FORM ---------- */
.form-title{ font-size:10px; font-weight:600; text-transform:uppercase; letter-spacing:.07em;
  color:var(--txt-3); margin:2px 0 6px; }
.stTextInput input{ font-size:12.5px !important; border-radius:7px !important; }
.stTextInput input:focus{ border-color:var(--brand) !important; box-shadow:0 0 0 3px #2563eb18 !important; }
/* Right-to-left rendering for Arabic-valued inputs */
.stTextInput input[aria-label*="Arabic"]{ direction:rtl; text-align:right; }
div[data-testid="stTextInput"] label p{ font-size:11px !important; font-weight:500 !important; color:var(--txt-2) !important; }
.iban-note{ font-size:10.5px; color:var(--txt-3); display:inline-flex; gap:5px;
  align-items:center; margin-top:2px; }

/* ---------- EXPANDER AS CLIENT CARD ---------- */
div[data-testid="stExpander"]{ border:1px solid var(--border) !important; border-radius:14px !important;
  background:#fff !important; margin-bottom:12px; box-shadow:0 1px 2px rgba(15,23,42,.03); }
div[data-testid="stExpander"] summary{ padding:12px 16px !important; font-weight:500; }
div[data-testid="stExpander"] summary:hover{ background:var(--panel-2); border-radius:14px; }

/* ---------- EXPORT + TABLE ---------- */
.section-card{ background:var(--panel); border:1px solid var(--border); border-radius:14px;
  padding:18px 20px; margin:6px 0 14px; }
.section-title{ font-size:14px; font-weight:600; color:var(--txt); margin-bottom:4px; }
.exp-info{ display:flex; gap:8px; font-size:11.5px; color:var(--txt-2); background:var(--panel-2);
  border-radius:8px; padding:10px 13px; line-height:1.55; margin-top:10px; }
.stButton button{ border-radius:8px !important; font-size:12.5px !important; font-weight:500 !important; }
.stDownloadButton button{ border-radius:8px !important; font-size:12.5px !important; }
.empty{ text-align:center; padding:44px 20px; background:var(--panel); border:1px dashed var(--border);
  border-radius:14px; color:var(--txt-2); }
.empty .big{ font-size:30px; margin-bottom:8px; }
.steps{ display:grid; grid-template-columns:repeat(3,1fr); gap:12px; margin-top:14px; }
.step{ background:var(--panel); border:1px solid var(--border); border-radius:12px; padding:14px 16px; }
.step b{ color:var(--brand); font-size:12px; }
.step p{ font-size:12px; color:var(--txt-2); margin-top:4px; line-height:1.5; }
table.rec-tbl{ width:100%; border-collapse:collapse; font-size:11.5px; }
table.rec-tbl th{ text-align:left; padding:9px 13px; font-size:10px; text-transform:uppercase;
  letter-spacing:.06em; color:var(--txt-3); background:var(--panel-2); border-bottom:1px solid var(--border); }
table.rec-tbl td{ padding:10px 13px; border-bottom:1px solid var(--border); color:var(--txt); }
table.rec-tbl tr:last-child td{ border-bottom:none; }
.mono{ font-family:ui-monospace,monospace; font-size:10.5px; }
.pill-ok{ background:var(--ok-bg); color:var(--ok-tx); padding:2px 8px; border-radius:10px; font-weight:500; font-size:10px; }
.pill-no{ background:#fef2f2; color:#b91c1c; padding:2px 8px; border-radius:10px; font-weight:500; font-size:10px; }
</style>
"""
st.markdown(PREMIUM_CSS, unsafe_allow_html=True)

# Human labels + which fields come only from Gemini Vision (red dot in mockup).
FIELD_META = [
    ("english_name", "English name", False),
    ("arabic_name", "Arabic name", True),
    ("iban", "IBAN — verified", False),
    ("account_number", "Account number", False),
    ("swift", "SWIFT code", False),
    ("bank_name_ar", "Bank (Arabic)", False),
    ("nationality", "Nationality", True),
    ("emirates_id", "Emirates ID", True),
    ("date_of_birth", "Date of birth", True),
    ("id_expiry", "ID expiry", True),
    ("emirate", "Emirate", False),
    ("currency", "Currency", False),
]


# ══════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ══════════════════════════════════════════════════════════════════════════
if "records" not in st.session_state:
    st.session_state.records = {}          # filename -> record dict


def secret_key() -> str:
    """Gemini key from secrets — accepts GEMINI_API_KEY or GOOGLE_API_KEY."""
    try:
        return st.secrets.get("GEMINI_API_KEY") or st.secrets.get("GOOGLE_API_KEY") or ""
    except Exception:
        return ""


def get_api_key() -> str:
    """API key — backend only (secrets.toml / Streamlit Cloud secrets)."""
    return secret_key()


# ══════════════════════════════════════════════════════════════════════════
#  SIDEBAR
# ══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(
        """<div class="sb-brand">
             <div class="sb-brand-icon">🏦</div>
             <div><div class="sb-brand-name">DocFill</div>
                  <div class="sb-brand-sub">Powered by AI Vision</div></div>
           </div>""",
        unsafe_allow_html=True,
    )

    nav = st.radio("Navigate", ["📊  Dashboard", "🕘  History"], label_visibility="collapsed")
    st.markdown("---")

    st.markdown('<div class="sb-label">Engine</div>', unsafe_allow_html=True)

    # Key is read only from the backend (secrets.toml / Cloud secrets) —
    # there is no manual key entry in the UI.
    if secret_key():
        st.markdown(
            f'<div class="sb-ok">✓ AI connection Successful</div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            '<div class="sb-warn">⚠ No API key in secrets.toml — set API_KEY</div>',
            unsafe_allow_html=True,
        )

    st.markdown("---")
    st.markdown('<div class="sb-label">Excel Template</div>', unsafe_allow_html=True)
    template_file = st.file_uploader(
        "قالب_تحديث.xlsx", type=["xlsx"], label_visibility="collapsed"
    )
    if template_file:
        st.markdown(
            f'<div class="sb-ok">✓ {html.escape(template_file.name)} loaded</div>',
            unsafe_allow_html=True,
        )
        st.caption("Using your file instead of the bundled template.")
    elif os.path.exists(DEFAULT_TEMPLATE_PATH):
        st.markdown(
            '<div class="sb-ok">✓ Using bundled supplier template</div>',
            unsafe_allow_html=True,
        )
        st.caption("Upload a file above to use your own template.")
    else:
        st.markdown(
            '<div class="sb-warn">⚠ No template found — upload قالب_تحديث.xlsx</div>',
            unsafe_allow_html=True,
        )

    # st.markdown("---")
    # st.markdown(f'<div class="sb-chip">✨ {GEMINI_MODEL}</div>', unsafe_allow_html=True)
    # st.markdown(
    #     '<div style="font-size:10px;color:#334155;line-height:1.6;margin-top:8px">'
    #     'Paid — unlimited requests<br>Full Arabic + ID card OCR</div>'
    #     '<div class="sb-deploy"><div class="sb-dot"></div> Deployed on Streamlit Cloud</div>',
    #     unsafe_allow_html=True,
    # )
    # st.caption("Get a key → [aistudio.google.com](https://aistudio.google.com)")


def resolve_template() -> bytes | None:
    """Active template bytes: uploaded override, else the bundled default."""
    if template_file is not None:
        return template_file.getvalue()
    try:
        with open(DEFAULT_TEMPLATE_PATH, "rb") as f:
            return f.read()
    except OSError:
        return None


active_template = resolve_template()


# ══════════════════════════════════════════════════════════════════════════
#  HISTORY VIEW
# ══════════════════════════════════════════════════════════════════════════
def render_history():
    backend = history.active_backend()
    badge = ("☁️ Google Sheets" if backend == "gsheet" else "💾 Local SQLite")
    st.markdown(
        f"""<div class="topbar">
              <div class="crumb">🏦 <span class="dim">DocFill</span> &rsaquo; History</div>
              <div class="engine-pill">{badge}</div>
            </div>""",
        unsafe_allow_html=True,
    )
    st.markdown('<div class="page-title">Extraction history</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="page-sub">Every record you save lands here — searchable, '
        're-exportable, and persistent across sessions.</div>',
        unsafe_allow_html=True,
    )

    try:
        s = history.stats()
        rows = history.fetch_all(st.session_state.get("hist_search", ""))
    except Exception as e:
        st.error(f"Could not reach the history backend ({backend}): {e}")
        if backend == "gsheet":
            st.info("Check GSHEET_WEBAPP_URL / GSHEET_TOKEN in secrets, and that the "
                    "Apps Script Web App is deployed with access set to **Anyone**.")
        return

    st.markdown(
        f"""<div class="metrics">
          <div class="metric"><div class="metric-label">Total records</div>
            <div class="metric-val">{s['total']}</div><div class="metric-sub">all time</div></div>
          <div class="metric"><div class="metric-label">Batches</div>
            <div class="metric-val">{s['batches']}</div><div class="metric-sub">export runs</div></div>
          <div class="metric"><div class="metric-label">Valid IBANs</div>
            <div class="metric-val">{s['valid_ibans']}</div>
            <span class="badge b-green">🛡 MOD-97</span></div>
          <div class="metric"><div class="metric-label">Last activity</div>
            <div class="metric-val" style="font-size:14px;margin-top:6px">{html.escape(str(s['last']))}</div></div>
        </div>""",
        unsafe_allow_html=True,
    )

    st.text_input("🔍 Search by name, IBAN, Emirates ID or file", key="hist_search",
                  placeholder="e.g. Mohammed, AE19…, 784-…")

    if not rows:
        st.markdown(
            '<div class="empty"><div class="big">🕘</div><b>No history yet</b><br>'
            'Extract documents on the Dashboard, then click <b>Save to history</b>.</div>',
            unsafe_allow_html=True,
        )
        return

    # actions
    c1, c2, c3 = st.columns([1.4, 1, 3])
    with c1:
        if active_template is not None:
            try:
                xlsx = export_to_excel(rows, active_template, append=False)
                st.download_button("📊  Export history to Excel", data=xlsx,
                                   file_name="history_export.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            except Exception as e:
                st.error(f"Export failed: {e}")
    with c2:
        if st.button("🗑 Clear history", use_container_width=True):
            st.session_state["confirm_clear"] = True
    with c3:
        st.markdown(f'<div style="padding-top:7px;font-size:12px;color:var(--txt-2)">'
                    f'Showing {len(rows)} record(s)</div>', unsafe_allow_html=True)

    if st.session_state.get("confirm_clear"):
        st.warning("Delete ALL history permanently?")
        cc1, cc2 = st.columns([1, 5])
        if cc1.button("Yes, delete all", type="primary"):
            removed = history.clear_all()
            st.session_state["confirm_clear"] = False
            st.toast(f"Cleared {removed} record(s)")
            st.rerun()
        if cc2.button("Cancel"):
            st.session_state["confirm_clear"] = False
            st.rerun()

    # table
    show_cols = ["ts", "english_name", "arabic_name", "iban", "bank_name_ar",
                 "emirate", "nationality", "emirates_id", "source_file", "iban_valid"]
    df = pd.DataFrame(rows)
    for c in show_cols:
        if c not in df.columns:
            df[c] = ""
    df = df[show_cols].rename(columns={
        "ts": "Saved at", "english_name": "Name (EN)", "arabic_name": "Name (AR)",
        "iban": "IBAN", "bank_name_ar": "Bank (AR)", "emirate": "Emirate",
        "nationality": "Nationality", "emirates_id": "Emirates ID",
        "source_file": "File", "iban_valid": "IBAN ✓",
    })
    df["IBAN ✓"] = df["IBAN ✓"].map(lambda v: "✅" if str(v) in ("1", "True", "true") else "—")
    st.dataframe(df, use_container_width=True, hide_index=True)


if "History" in nav:
    render_history()
    st.stop()


# ══════════════════════════════════════════════════════════════════════════
#  TOP BAR + PAGE HEAD
# ══════════════════════════════════════════════════════════════════════════
n_files = len(st.session_state.records)
st.markdown(
    f"""<div class="topbar">
          <div class="crumb">🏦 <span class="dim">DocFill</span> &rsaquo; Dashboard</div>
          <div class="engine-pill">✨ AI Vision — {n_files} file{'s' if n_files != 1 else ''} processed</div>
        </div>""",
    unsafe_allow_html=True,
)
st.markdown('<div class="page-title">DocFill — Client document extraction</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="page-sub">Upload IBAN letters + Emirates ID PDFs — DocFill extracts '
    'every field automatically with AI Vision. Review, edit, and export to the supplier template.</div>',
    unsafe_allow_html=True,
)

MAX_FILES = 10

# ── File uploader ──────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    f"Upload client PDFs — up to {MAX_FILES} at once (IBAN letter + Emirates ID)",
    type=["pdf"],
    accept_multiple_files=True,
    help=f"Up to {MAX_FILES} PDFs per batch. Each PDF: page 1 = IBAN letter, page 2 = Emirates ID scan",
)

api_key = get_api_key()

# Enforce the 10-file cap (the widget itself has no count limit).
if uploaded and len(uploaded) > MAX_FILES:
    st.warning(f"⚠️ {len(uploaded)} files selected — only the first {MAX_FILES} will be processed.")
    uploaded = uploaded[:MAX_FILES]

# Files not yet extracted (dedup by filename).
pending = [p for p in (uploaded or []) if p.name not in st.session_state.records]

# ── Batch action bar ───────────────────────────────────────────────────────
extract_clicked = False
if uploaded:
    b1, b2, b3 = st.columns([1.5, 1, 3.2])
    with b1:
        extract_clicked = st.button(
            f"⚡ Extract all ({len(pending)})" if pending else "✓ All extracted",
            type="primary", use_container_width=True, disabled=not pending,
        )
    with b2:
        if st.button("🗑 Clear all", use_container_width=True,
                     disabled=not st.session_state.records):
            st.session_state.records = {}
            st.rerun()
    with b3:
        done = len(st.session_state.records)
        st.markdown(
            f'<div style="padding-top:7px;font-size:12px;color:var(--txt-2)">'
            f'{len(uploaded)} uploaded · {done} extracted · {len(pending)} pending</div>',
            unsafe_allow_html=True,
        )

# ── Process the batch on click, with a single progress bar ─────────────────
if extract_clicked and pending:
    if not api_key:
        st.error("⚠️ No API key — set GEMINI_API_KEY in .streamlit/secrets.toml (or Streamlit Cloud secrets).")
    else:
        prog = st.progress(0.0, text="Starting…")
        for i, pdf in enumerate(pending, start=1):
            prog.progress((i - 1) / len(pending), text=f"Reading {pdf.name}  ({i}/{len(pending)})…")
            try:
                images = pdf_to_images(pdf.read())
            except Exception as e:
                st.error(f"Could not open **{pdf.name}** — invalid PDF? ({e})")
                continue
            try:
                rec = extract_with_gemini(images, api_key)
            except Exception as e:
                st.warning(f"AI extraction failed for **{pdf.name}** ({e}). Fill fields manually below.")
                rec = blank_record()
            rec["source_file"] = pdf.name
            st.session_state.records[pdf.name] = rec
        prog.progress(1.0, text=f"Done — {len(pending)} file(s) processed")
        prog.empty()
        st.rerun()


# ══════════════════════════════════════════════════════════════════════════
#  METRICS
# ══════════════════════════════════════════════════════════════════════════
records = st.session_state.records
recs = list(records.values())

if recs:
    total_fields = sum(len(FIELD_META) for _ in recs)
    ok_fields = sum(count_ok(r) for r in recs)
    valid_ibans = sum(1 for r in recs if validate_iban(r.get("iban", "")))
    ready = all(is_export_ready(r) for r in recs) and active_template is not None

    st.markdown(
        f"""<div class="metrics">
          <div class="metric"><div class="metric-label">Files processed</div>
            <div class="metric-val">{len(recs)}</div><div class="metric-sub">PDFs uploaded</div></div>
          <div class="metric"><div class="metric-label">Fields extracted</div>
            <div class="metric-val">{ok_fields}</div>
            <span class="badge {'b-green' if ok_fields==total_fields else 'b-amber'}">✓ {ok_fields} / {total_fields}</span></div>
          <div class="metric"><div class="metric-label">IBANs valid</div>
            <div class="metric-val">{valid_ibans}</div>
            <span class="badge b-green">🛡 MOD-97</span></div>
          <div class="metric"><div class="metric-label">Export status</div>
            <div class="metric-val" style="font-size:16px;margin-top:5px">{'Ready' if ready else 'Pending'}</div>
            <span class="badge {'b-blue' if ready else 'b-amber'}">📄 Excel template</span></div>
        </div>""",
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        """<div class="empty"><div class="big">📥</div>
             <b>No documents yet</b><br>Upload one PDF per client to begin extraction.
             <div class="steps">
               <div class="step"><b>Step 1</b><p>Upload PDF(s) — one per client</p></div>
               <div class="step"><b>Step 2</b><p>Review extracted fields, fix if needed</p></div>
               <div class="step"><b>Step 3</b><p>Export to the supplier Excel template</p></div>
             </div></div>""",
        unsafe_allow_html=True,
    )


# ══════════════════════════════════════════════════════════════════════════
#  CLIENT CARDS (status checklist + editable form)
# ══════════════════════════════════════════════════════════════════════════
AVATAR_BG = ["#eff6ff", "#f0fdf4", "#fef2f2", "#faf5ff", "#fffbeb"]
AVATAR_FG = ["#1d4ed8", "#15803d", "#b91c1c", "#7e22ce", "#b45309"]


def initials(name: str) -> str:
    parts = [p for p in (name or "").split() if p]
    if not parts:
        return "··"
    return (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper()


for idx, (fname, rec) in enumerate(records.items()):
    name = rec.get("english_name") or fname
    bank = rec.get("bank_name_en") or rec.get("bank_name_ar") or "—"
    emirate = rec.get("emirate") or "—"
    n_ok = count_ok(rec)
    badge = f"✓ {n_ok} fields"

    with st.expander(f"  {initials(name)}   {name}   —   {fname} · {bank} · {emirate}   ·   {badge}", expanded=True):
        col_status, col_form = st.columns([1, 2.4], gap="medium")

        # ----- left: extracted-field checklist -----
        stat = field_status(rec)
        rows = ""
        for key, label, gemini_only in FIELD_META:
            ok = stat.get(key) == "ok"
            ic = '<span class="ic-ok">✓</span>' if ok else '<span class="ic-wa">!</span>'
            dot = '<span class="gem-dot"></span>' if gemini_only else ""
            rows += f'<div class="srow">{ic} {html.escape(label)} {dot}</div>'
        col_status.markdown(
            f'<div class="status-box"><div class="status-title">Extracted fields</div>{rows}'
            '<div class="gem-note"><span class="gem-dot"></span> AI Vision only</div></div>',
            unsafe_allow_html=True,
        )

        # ----- right: editable form (native widgets so edits persist) -----
        col_form.markdown('<div class="form-title">Review &amp; edit</div>', unsafe_allow_html=True)
        a, b = col_form.columns(2)
        rec["english_name"]   = a.text_input("English name", rec.get("english_name", ""), key=f"en_{idx}")
        rec["arabic_name"]    = b.text_input("Arabic name · Gemini", rec.get("arabic_name", ""), key=f"ar_{idx}")
        rec["iban"]           = a.text_input("IBAN", rec.get("iban", ""), key=f"iban_{idx}")
        rec["account_number"] = b.text_input("Account number", rec.get("account_number", ""), key=f"acct_{idx}")
        rec["swift"]          = a.text_input("SWIFT code", rec.get("swift", ""), key=f"swift_{idx}")
        rec["bank_name_ar"]   = b.text_input("Bank (Arabic)", rec.get("bank_name_ar", ""), key=f"bkar_{idx}")
        rec["nationality"]    = a.text_input("Nationality · Gemini", rec.get("nationality", ""), key=f"nat_{idx}")
        rec["emirate"]        = b.text_input("Emirate", rec.get("emirate", ""), key=f"emir_{idx}")
        rec["emirates_id"]    = a.text_input("Emirates ID · Gemini", rec.get("emirates_id", ""), key=f"eid_{idx}")
        rec["currency"]       = b.text_input("Currency", rec.get("currency", "AED"), key=f"curr_{idx}")
        rec["date_of_birth"]  = a.text_input("Date of birth · Gemini", rec.get("date_of_birth", ""), key=f"dob_{idx}")
        rec["id_expiry"]      = b.text_input("ID expiry · Gemini", rec.get("id_expiry", ""), key=f"exp_{idx}")

        # live IBAN check — tiny neutral note that says WHAT to fix
        if rec["iban"]:
            ok, hint = iban_diagnosis(rec["iban"])
            col_form.markdown(
                f'<span class="iban-note">{"✓" if ok else "ⓘ"} {html.escape(hint)}</span>',
                unsafe_allow_html=True,
            )

        if col_form.button("🗑 Remove this record", key=f"del_{idx}"):
            del st.session_state.records[fname]
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════
#  EXPORT + SUMMARY TABLE
# ══════════════════════════════════════════════════════════════════════════
if recs:
    # finalize records for export
    for r in recs:
        r["iban"] = normalize_iban(r.get("iban", ""))
        r["iban_valid"] = validate_iban(r["iban"])
    export_list = list(records.values())

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">Export records</div>', unsafe_allow_html=True)

    if active_template is not None:
        existing = existing_record_count(active_template)
        append = True
        if existing > 0:
            mode = st.radio(
                "Write mode",
                [f"➕ Append after existing {existing} record(s)", "♻️ Overwrite from row 3"],
                horizontal=True, label_visibility="collapsed",
            )
            append = mode.startswith("➕")
            note = (f"adds {len(export_list)} new record(s) after the existing {existing}"
                    if append else f"replaces existing rows with {len(export_list)} record(s)")
            st.markdown(
                f'<div class="exp-info">📋 The template already contains <b>{existing}</b> record(s) — '
                f'export {note}.</div>',
                unsafe_allow_html=True,
            )
        try:
            xlsx = export_to_excel(export_list, active_template, append=append)
            st.download_button(
                "📊  Export to Excel template",
                data=xlsx,
                file_name="supplier_template_filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Excel export failed: {e}")
    else:
        st.button("📊  Export to Excel template", disabled=True, use_container_width=True,
                  help="No template found — add assets/supplier_template.xlsx or upload one in the sidebar")

    # Save the current reviewed records to persistent history.
    if st.button(f"💾  Save {len(export_list)} record(s) to history", use_container_width=True):
        try:
            n = history.add_records(export_list)
            st.toast(f"Saved {n} record(s) to history ✓")
        except Exception as e:
            st.error(f"Could not save to history ({history.active_backend()}): {e}")

    st.markdown(
        '<div class="exp-info">ℹ️ Fills the supplier template — Arabic name (B), English name (C), '
        'Emirate (G), Bank Arabic (K), Account No. (M), IBAN (N), Account title (O). '
        'Arabic is written natively into the .xlsx, so it renders correctly in Excel. '
        f'History is saved to <b>{"Google Sheets" if history.active_backend()=="gsheet" else "local SQLite"}</b>.</div>',
        unsafe_allow_html=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

    # ----- summary table -----
    def short_iban(v: str) -> str:
        v = v or ""
        return f"{v[:6]}…{v[-6:]}" if len(v) > 14 else v

    body = ""
    for r in export_list:
        valid = validate_iban(r.get("iban", ""))
        body += (
            "<tr>"
            f'<td style="font-weight:500">{html.escape(r.get("english_name") or "—")}</td>'
            f'<td class="mono">{html.escape(short_iban(r.get("iban", "")))}</td>'
            f'<td dir="rtl">{html.escape(r.get("bank_name_ar") or "—")}</td>'
            f'<td class="mono">{html.escape(r.get("swift") or "—")}</td>'
            f'<td>{html.escape(r.get("nationality") or "—")}</td>'
            f'<td>{"<span class=pill-ok>✓ valid</span>" if valid else "<span style=\'color:var(--txt-3);font-size:10px\'>verify</span>"}</td>'
            "</tr>"
        )
    st.markdown(
        f"""<div class="section-card" style="padding:0;overflow:hidden">
          <div style="padding:13px 18px;border-bottom:1px solid var(--border);display:flex;justify-content:space-between;align-items:center">
            <span class="section-title" style="margin:0">All records — {len(export_list)} client{'s' if len(export_list)!=1 else ''}</span>
            <span class="badge b-green">Export ready</span></div>
          <table class="rec-tbl">
            <tr><th>Client name</th><th>IBAN</th><th>Bank (Arabic)</th><th>SWIFT</th><th>Nationality</th><th>Status</th></tr>
            {body}
          </table></div>""",
        unsafe_allow_html=True,
    )

    with st.expander("View full data table (all fields)"):
        st.dataframe(pd.DataFrame(export_list), use_container_width=True)
