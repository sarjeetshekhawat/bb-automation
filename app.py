"""
Blocked VILL Meter Classification — Streamlit Application
==========================================================

A professional, dark-themed web UI that lets users upload four Excel files,
select matching columns, classify every blocked VILL meter against
WC / LTCT / HTCT installation records, and download a formatted report.
"""

import streamlit as st
import pandas as pd
from engine import (
    read_excel_safe,
    get_column_names,
    normalise_key,
    build_lookup_set,
    find_internal_duplicates,
    classify_meters,
    compute_summary,
    generate_report,
)

# ──────────────────────────────────────────────────────────────────────
# Page Configuration
# ──────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Blocked VILL Meter Classifier",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ──────────────────────────────────────────────────────────────────────
# Custom CSS — Premium Dark Glassmorphism
# ──────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Root Variables ─────────────────────────────────────────────── */
:root {
    --bg-primary:    #0F172A;
    --bg-card:       rgba(30, 41, 59, 0.75);
    --bg-card-hover: rgba(51, 65, 85, 0.85);
    --accent:        #6366F1;
    --accent-light:  #818CF8;
    --success:       #10B981;
    --warning:       #F59E0B;
    --danger:        #EF4444;
    --info:          #3B82F6;
    --purple:        #A855F7;
    --text-primary:  #F1F5F9;
    --text-muted:    #94A3B8;
    --border:        rgba(100, 116, 139, 0.25);
    --glow:          0 0 20px rgba(99, 102, 241, 0.15);
}

/* ── Global ─────────────────────────────────────────────────────── */
.stApp {
    background: linear-gradient(135deg, #0F172A 0%, #1E1B4B 50%, #0F172A 100%) !important;
}

/* ── Hero Banner ────────────────────────────────────────────────── */
.hero-banner {
    background: linear-gradient(135deg, rgba(99,102,241,0.25) 0%, rgba(168,85,247,0.18) 100%);
    border: 1px solid var(--border);
    border-radius: 16px;
    padding: 2.5rem 2rem;
    margin-bottom: 2rem;
    text-align: center;
    backdrop-filter: blur(12px);
    box-shadow: var(--glow);
}
.hero-banner h1 {
    font-size: 2.2rem;
    font-weight: 800;
    background: linear-gradient(90deg, #818CF8, #C084FC, #818CF8);
    background-size: 200%;
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    animation: shimmer 3s ease-in-out infinite;
    margin-bottom: 0.5rem;
}
@keyframes shimmer {
    0%,100% { background-position: 0% 50%; }
    50% { background-position: 100% 50%; }
}
.hero-banner p {
    color: var(--text-muted);
    font-size: 1.05rem;
    margin: 0;
}

/* ── Step Cards ─────────────────────────────────────────────────── */
.step-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 1.5rem 1.8rem;
    margin-bottom: 1.5rem;
    backdrop-filter: blur(10px);
    transition: border-color 0.3s, box-shadow 0.3s;
}
.step-card:hover {
    border-color: var(--accent);
    box-shadow: var(--glow);
}
.step-title {
    font-size: 1.15rem;
    font-weight: 700;
    color: var(--accent-light);
    margin-bottom: 0.6rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.step-title .badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 28px; height: 28px;
    border-radius: 8px;
    background: var(--accent);
    color: #fff;
    font-size: 0.85rem;
    font-weight: 700;
}

/* ── Metric Cards ───────────────────────────────────────────────── */
.metric-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 1rem;
    margin: 1.5rem 0;
}
.metric-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 1.2rem 1rem;
    text-align: center;
    transition: transform 0.25s, box-shadow 0.25s;
}
.metric-card:hover {
    transform: translateY(-3px);
    box-shadow: var(--glow);
}
.metric-value {
    font-size: 2rem;
    font-weight: 800;
    line-height: 1.1;
}
.metric-label {
    font-size: 0.82rem;
    color: var(--text-muted);
    margin-top: 0.35rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}

/* Colour accents for each category */
.mc-total .metric-value  { color: var(--accent-light); }
.mc-wc    .metric-value  { color: var(--success); }
.mc-ltct  .metric-value  { color: var(--info); }
.mc-htct  .metric-value  { color: var(--warning); }
.mc-nf    .metric-value  { color: var(--danger); }
.mc-mm    .metric-value  { color: var(--purple); }
.mc-ok    .metric-value  { color: var(--success); }
.mc-pct   .metric-value  { color: #34D399; }

/* ── Duplicate Warning ──────────────────────────────────────────── */
.dup-warning {
    background: rgba(245, 158, 11, 0.12);
    border: 1px solid rgba(245, 158, 11, 0.4);
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin: 0.8rem 0;
    color: var(--warning);
    font-size: 0.92rem;
}

/* ── Table tweaks ───────────────────────────────────────────────── */
.stDataFrame { border-radius: 10px; overflow: hidden; }

/* ── Sidebar ────────────────────────────────────────────────────── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1E293B 0%, #0F172A 100%) !important;
}

/* ── Button styling ──────────────────────────────────────────────── */
.stDownloadButton > button,
button[kind="primary"] {
    background: linear-gradient(135deg, #6366F1, #8B5CF6) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    padding: 0.6rem 2rem !important;
    transition: transform 0.2s, box-shadow 0.2s !important;
}
.stDownloadButton > button:hover,
button[kind="primary"]:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 20px rgba(99, 102, 241, 0.4) !important;
}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────────────
# Hero Banner
# ──────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero-banner">
    <h1>⚡ Blocked VILL Meter Classifier</h1>
    <p>Upload your Excel / CSV files · Select matching columns · Get instant classification</p>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════
# SESSION STATE DEFAULTS
# ══════════════════════════════════════════════════════════════════════
for key in ("classified_df", "summary", "report_bytes",
            "dup_wc", "dup_ltct", "dup_htct"):
    if key not in st.session_state:
        st.session_state[key] = None


# ══════════════════════════════════════════════════════════════════════
# STEP 1 — FILE INPUT  (Upload or File Path)
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="step-card">
    <div class="step-title"><span class="badge">1</span> Provide Excel / CSV Files</div>
</div>
""", unsafe_allow_html=True)

import os

input_mode = st.radio(
    "**Choose input method:**",
    ["📂 Browse & Upload (files < 500 MB)", "📁 Paste File Path (any size, recommended for large files)"],
    horizontal=True,
    key="input_mode",
    help="Use 'Paste File Path' for large files (> 500 MB) to avoid browser upload limits.",
)

use_path_mode = "Paste File Path" in input_mode

# ── Helper: validate a local path ─────────────────────────────────
def _validate_path(path: str, label: str) -> str | None:
    """Return cleaned path if valid, else show error and return None."""
    path = path.strip().strip('"').strip("'")
    if not path:
        return None
    if not os.path.isfile(path):
        st.error(f"❌ **{label}**: File not found → `{path}`")
        return None
    ext = os.path.splitext(path)[1].lower()
    if ext not in (".xlsx", ".xls", ".csv"):
        st.error(f"❌ **{label}**: Unsupported format `{ext}`. Use .xlsx, .xls, or .csv")
        return None
    return path


if use_path_mode:
    # ── FILE PATH MODE ────────────────────────────────────────────
    st.caption(
        "Paste the **full file path** from Windows Explorer "
        "(e.g. `C:\\Users\\...\\file.csv`). This reads directly from disk — "
        "no size limit."
    )
    col1, col2 = st.columns(2)
    with col1:
        bp = st.text_input("📋 **Blocked Bill** path", key="path_blocked",
                           placeholder=r"C:\Users\...\Blocked_Bill.xlsx")
        wcp = st.text_input("🟢 **WC Consumer Installation** path", key="path_wc",
                            placeholder=r"C:\Users\...\WC_Installation.csv")
    with col2:
        lp = st.text_input("🔵 **LTCT Consumer Installation** path", key="path_ltct",
                           placeholder=r"C:\Users\...\LTCT_Installation.csv")
        hp = st.text_input("🟡 **HTCT Consumer Installation** path", key="path_htct",
                           placeholder=r"C:\Users\...\HTCT_Installation.csv")

    # Validate all paths
    blocked_path = _validate_path(bp, "Blocked Bill") if bp.strip() else None
    wc_path      = _validate_path(wcp, "WC Installation") if wcp.strip() else None
    ltct_path    = _validate_path(lp, "LTCT Installation") if lp.strip() else None
    htct_path    = _validate_path(hp, "HTCT Installation") if hp.strip() else None

    all_ready = all([blocked_path, wc_path, ltct_path, htct_path])

    if all_ready:
        # Show file sizes
        for label, p in [("Blocked Bill", blocked_path), ("WC", wc_path),
                         ("LTCT", ltct_path), ("HTCT", htct_path)]:
            size_mb = os.path.getsize(p) / (1024 * 1024)
            if size_mb >= 1024:
                st.caption(f"✅ {label}: `{os.path.basename(p)}` — {size_mb/1024:.2f} GB")
            else:
                st.caption(f"✅ {label}: `{os.path.basename(p)}` — {size_mb:.1f} MB")

    if not all_ready:
        st.info("👆 Provide all four file paths to proceed.", icon="📁")
        st.stop()

    # Create file-like references for downstream code
    blocked_file = blocked_path
    wc_file      = wc_path
    ltct_file    = ltct_path
    htct_file    = htct_path

else:
    # ── UPLOAD MODE ───────────────────────────────────────────────
    col1, col2 = st.columns(2)
    with col1:
        blocked_file = st.file_uploader(
            "📋 **Blocked Bill** (Primary File)",
            type=["xlsx", "xls", "csv"],
            key="blocked",
            help="The file containing blocked VILL meters to classify."
        )
        wc_file = st.file_uploader(
            "🟢 **WC Consumer Installation**",
            type=["xlsx", "xls", "csv"],
            key="wc",
            help="Installation records for WC meters."
        )

    with col2:
        ltct_file = st.file_uploader(
            "🔵 **LTCT Consumer Installation**",
            type=["xlsx", "xls", "csv"],
            key="ltct",
            help="Installation records for LTCT meters."
        )
        htct_file = st.file_uploader(
            "🟡 **HTCT Consumer Installation**",
            type=["xlsx", "xls", "csv"],
            key="htct",
            help="Installation records for HTCT meters."
        )

    all_ready = all([blocked_file, wc_file, ltct_file, htct_file])

    if not all_ready:
        st.info("👆 Upload all four Excel / CSV files to proceed.", icon="📂")
        st.stop()

# ══════════════════════════════════════════════════════════════════════
# Read column headers
# ══════════════════════════════════════════════════════════════════════
def _seek_reset(f):
    """Reset seek position for uploaded file objects; no-op for paths."""
    if hasattr(f, "seek"):
        f.seek(0)

try:
    blocked_cols = get_column_names(blocked_file)
    _seek_reset(blocked_file)
    wc_cols = get_column_names(wc_file)
    _seek_reset(wc_file)
    ltct_cols = get_column_names(ltct_file)
    _seek_reset(ltct_file)
    htct_cols = get_column_names(htct_file)
    _seek_reset(htct_file)
except Exception as e:
    st.error(f"❌ Error reading file headers: {e}")
    st.stop()

# ══════════════════════════════════════════════════════════════════════
# STEP 2 — COLUMN MAPPING
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="step-card">
    <div class="step-title"><span class="badge">2</span> Select Matching Column in Each File</div>
</div>
""", unsafe_allow_html=True)

st.caption(
    "Choose the column that contains the **common identifier** "
    "(e.g., Meter Number, Consumer Number, Installation Number) in each file. "
    "This identifier will be used for exact matching."
)


def _auto_detect(columns: list[str]) -> int:
    """Heuristically pick the best default column index."""
    keywords = [
        "meter", "consumer", "account", "installation",
        "inst", "acct", "con no", "meter no", "meter_no",
        "consumer no", "consumer_no", "service", "k no", "kno",
    ]
    for idx, col in enumerate(columns):
        lower = col.lower().strip()
        for kw in keywords:
            if kw in lower:
                return idx
    return 0  # fallback: first column


mc1, mc2, mc3, mc4 = st.columns(4)
with mc1:
    blocked_key = st.selectbox(
        "📋 Blocked Bill",
        blocked_cols,
        index=_auto_detect(blocked_cols),
        key="sel_blocked",
    )
with mc2:
    wc_key = st.selectbox(
        "🟢 WC Installation",
        wc_cols,
        index=_auto_detect(wc_cols),
        key="sel_wc",
    )
with mc3:
    ltct_key = st.selectbox(
        "🔵 LTCT Installation",
        ltct_cols,
        index=_auto_detect(ltct_cols),
        key="sel_ltct",
    )
with mc4:
    htct_key = st.selectbox(
        "🟡 HTCT Installation",
        htct_cols,
        index=_auto_detect(htct_cols),
        key="sel_htct",
    )


# ══════════════════════════════════════════════════════════════════════
# STEP 3 — PREVIEW  (lazy — only reads on demand, limited rows for CSVs)
# ══════════════════════════════════════════════════════════════════════
def _read_preview(f, n=5):
    """Read only the first *n* rows for preview — fast even for GB files."""
    import os
    # Determine if CSV
    if isinstance(f, str):
        ext = os.path.splitext(f)[1].lower()
    elif hasattr(f, "name"):
        ext = os.path.splitext(f.name)[1].lower()
    else:
        ext = ""

    if ext == ".csv":
        if not isinstance(f, str) and hasattr(f, "seek"):
            f.seek(0)
        return pd.read_csv(f, dtype=str, keep_default_na=False, nrows=n)
    else:
        # For Excel, read_excel_safe reads full file; just .head() it
        df = read_excel_safe(f)
        _seek_reset(f)
        return df.head(n)


show_preview = st.checkbox("🔍 Preview data (first 5 rows each)", value=False)

if show_preview:
    prev_tab1, prev_tab2, prev_tab3, prev_tab4 = st.tabs(
        ["Blocked Bill", "WC", "LTCT", "HTCT"]
    )
    with prev_tab1:
        st.dataframe(_read_preview(blocked_file), use_container_width=True)
    with prev_tab2:
        st.dataframe(_read_preview(wc_file), use_container_width=True)
    with prev_tab3:
        st.dataframe(_read_preview(ltct_file), use_container_width=True)
    with prev_tab4:
        st.dataframe(_read_preview(htct_file), use_container_width=True)


# ══════════════════════════════════════════════════════════════════════
# STEP 4 — PROCESS
# ══════════════════════════════════════════════════════════════════════
st.markdown("""
<div class="step-card">
    <div class="step-title"><span class="badge">3</span> Process &amp; Classify</div>
</div>
""", unsafe_allow_html=True)

process_btn = st.button("🚀  **Classify Meters**", type="primary",
                        use_container_width=True)

if process_btn:
    with st.spinner("Reading files… (large files may take a minute)"):
        try:
            blocked_df = read_excel_safe(blocked_file)
            _seek_reset(blocked_file)
            wc_df = read_excel_safe(wc_file)
            _seek_reset(wc_file)
            ltct_df = read_excel_safe(ltct_file)
            _seek_reset(ltct_file)
            htct_df = read_excel_safe(htct_file)
            _seek_reset(htct_file)
        except Exception as e:
            st.error(f"❌ Failed to read files: {e}")
            st.stop()

    # Validate selected columns exist
    for label, df, col in [
        ("Blocked Bill", blocked_df, blocked_key),
        ("WC Installation", wc_df, wc_key),
        ("LTCT Installation", ltct_df, ltct_key),
        ("HTCT Installation", htct_df, htct_key),
    ]:
        if col not in df.columns:
            st.error(f"❌ Column **'{col}'** not found in {label} file. "
                     f"Available columns: {list(df.columns)}")
            st.stop()

    # ── Duplicate detection inside installation files ─────────────
    with st.spinner("Checking for duplicates inside installation files…"):
        dup_wc   = find_internal_duplicates(wc_df,   wc_key)
        dup_ltct = find_internal_duplicates(ltct_df, ltct_key)
        dup_htct = find_internal_duplicates(htct_df, htct_key)
        st.session_state.dup_wc   = dup_wc   if len(dup_wc)   else None
        st.session_state.dup_ltct = dup_ltct if len(dup_ltct) else None
        st.session_state.dup_htct = dup_htct if len(dup_htct) else None

    # ── Build lookup sets ─────────────────────────────────────────
    with st.spinner("Building lookup sets…"):
        wc_set   = build_lookup_set(wc_df,   wc_key)
        ltct_set = build_lookup_set(ltct_df, ltct_key)
        htct_set = build_lookup_set(htct_df, htct_key)

    st.write(f"Unique identifiers → WC: **{len(wc_set):,}** · "
             f"LTCT: **{len(ltct_set):,}** · HTCT: **{len(htct_set):,}**")

    # ── Classify ──────────────────────────────────────────────────
    progress_bar = st.progress(0, text="Classifying meters…")
    classified_df = classify_meters(
        blocked_df, blocked_key,
        wc_set, ltct_set, htct_set,
        progress_callback=lambda p: progress_bar.progress(
            p, text=f"Classifying… {p:.0%}"
        ),
    )
    progress_bar.progress(1.0, text="✅ Classification complete!")

    # ── Summary ───────────────────────────────────────────────────
    summary = compute_summary(classified_df)

    # ── Generate report ───────────────────────────────────────────
    with st.spinner("Generating Excel report…"):
        report_bytes = generate_report(classified_df, summary)

    # Save to session state so results persist across reruns
    st.session_state.classified_df = classified_df
    st.session_state.summary       = summary
    st.session_state.report_bytes  = report_bytes


# ══════════════════════════════════════════════════════════════════════
# DISPLAY RESULTS  (from session state)
# ══════════════════════════════════════════════════════════════════════
if st.session_state.classified_df is not None:
    classified_df = st.session_state.classified_df
    summary       = st.session_state.summary
    report_bytes  = st.session_state.report_bytes

    # ── Duplicate warnings ────────────────────────────────────────
    for label, dup_df_key in [("WC", "dup_wc"), ("LTCT", "dup_ltct"),
                               ("HTCT", "dup_htct")]:
        dup_df = st.session_state.get(dup_df_key)
        if dup_df is not None:
            st.markdown(
                f'<div class="dup-warning">⚠️ <strong>{label} Installation</strong> '
                f'contains <strong>{len(dup_df)}</strong> rows with duplicate '
                f'identifiers. Review these for accuracy.</div>',
                unsafe_allow_html=True,
            )
            with st.expander(f"View {label} duplicate rows"):
                st.dataframe(dup_df, use_container_width=True)

    # ── Summary Dashboard ─────────────────────────────────────────
    st.markdown("""
    <div class="step-card">
        <div class="step-title"><span class="badge">✓</span> Classification Summary</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="metric-grid">
        <div class="metric-card mc-total">
            <div class="metric-value">{summary['Total Blocked Meters Processed']:,}</div>
            <div class="metric-label">Total Processed</div>
        </div>
        <div class="metric-card mc-wc">
            <div class="metric-value">{summary['WC Meters']:,}</div>
            <div class="metric-label">WC Meters</div>
        </div>
        <div class="metric-card mc-ltct">
            <div class="metric-value">{summary['LTCT Meters']:,}</div>
            <div class="metric-label">LTCT Meters</div>
        </div>
        <div class="metric-card mc-htct">
            <div class="metric-value">{summary['HTCT Meters']:,}</div>
            <div class="metric-label">HTCT Meters</div>
        </div>
        <div class="metric-card mc-nf">
            <div class="metric-value">{summary['Not Found']:,}</div>
            <div class="metric-label">Not Found</div>
        </div>
        <div class="metric-card mc-mm">
            <div class="metric-value">{summary['Multiple Matches']:,}</div>
            <div class="metric-label">Multiple Matches</div>
        </div>
        <div class="metric-card mc-ok">
            <div class="metric-value">{summary['Successfully Classified']:,}</div>
            <div class="metric-label">Classified</div>
        </div>
        <div class="metric-card mc-pct">
            <div class="metric-value">{summary['Classification %']}%</div>
            <div class="metric-label">Classification Rate</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Reconciliation check ──────────────────────────────────────
    total = summary["Total Blocked Meters Processed"]
    recon = (summary["WC Meters"] + summary["LTCT Meters"] +
             summary["HTCT Meters"] + summary["Not Found"] +
             summary["Multiple Matches"])
    if recon == total:
        st.success(f"✅ Counts reconcile: {recon:,} = {total:,} total", icon="✅")
    else:
        st.warning(
            f"⚠️ Reconciliation mismatch: sum of categories ({recon:,}) "
            f"≠ total ({total:,}). Review data for issues.",
            icon="⚠️",
        )

    # ── Data tabs ─────────────────────────────────────────────────
    tab_all, tab_wc, tab_ltct, tab_htct, tab_nf, tab_mm = st.tabs([
        f"All ({total:,})",
        f"WC ({summary['WC Meters']:,})",
        f"LTCT ({summary['LTCT Meters']:,})",
        f"HTCT ({summary['HTCT Meters']:,})",
        f"Not Found ({summary['Not Found']:,})",
        f"Multiple ({summary['Multiple Matches']:,})",
    ])
    with tab_all:
        st.dataframe(classified_df, use_container_width=True, height=420)
    with tab_wc:
        st.dataframe(
            classified_df[classified_df["Final_Meter_Category"] == "WC"],
            use_container_width=True, height=420,
        )
    with tab_ltct:
        st.dataframe(
            classified_df[classified_df["Final_Meter_Category"] == "LTCT"],
            use_container_width=True, height=420,
        )
    with tab_htct:
        st.dataframe(
            classified_df[classified_df["Final_Meter_Category"] == "HTCT"],
            use_container_width=True, height=420,
        )
    with tab_nf:
        st.dataframe(
            classified_df[classified_df["Final_Meter_Category"] == "NOT FOUND"],
            use_container_width=True, height=420,
        )
    with tab_mm:
        st.dataframe(
            classified_df[classified_df["Final_Meter_Category"] == "MULTIPLE MATCH"],
            use_container_width=True, height=420,
        )

    # ── Download Button ───────────────────────────────────────────
    st.markdown("---")
    st.download_button(
        label="📥  Download Excel Report",
        data=report_bytes,
        file_name="Blocked_VILL_Meter_Classification.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

# ──────────────────────────────────────────────────────────────────────
# Footer
# ──────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "Built with ❤️ using Python · pandas · openpyxl · Streamlit  —  "
    "Blocked VILL Meter Classification Automation"
)
