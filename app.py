import streamlit as st
import json
import os
import re
import textwrap
from pathlib import Path

# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Marsh AI Pitch Advisor",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# RENDER HELPERS
# ============================================================
# All raw HTML goes through render_html(), which renders it as a
# genuine DOM fragment (via st.html when available) instead of
# leaking markup into the page as visible text.

def render_html(content: str) -> None:
    """Render a block of raw HTML reliably, in a single call so
    open/close tags never get split across separate Streamlit
    markdown calls (the root cause of raw-HTML leaking into the UI)."""
    content = textwrap.dedent(content).strip()
    if hasattr(st, "html"):
        st.html(content)
    else:
        st.markdown(content, unsafe_allow_html=True)


def esc(value) -> str:
    """Minimal HTML-escaping for values interpolated into card markup."""
    text = "" if value is None else str(value)
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def slugify(name: str) -> str:
    slug = re.sub(r"[^a-zA-Z0-9]+", "_", name.strip().lower()).strip("_")
    return slug or "client"


def section_title(title: str, description: str = "") -> None:
    html = f'<div class="section-title">{esc(title)}</div>'
    if description:
        html += f'<div class="section-description">{esc(description)}</div>'
    render_html(html)


def render_metric_card(label: str, value: str, sub: str = "", value_size: str = "24px") -> str:
    return f"""
    <div class="card metric-card">
        <div class="card-title">{esc(label)}</div>
        <div class="card-value" style="font-size:{value_size};">{esc(value)}</div>
        <div class="card-small">{esc(sub)}</div>
    </div>
    """


def render_risk_card(number: int, risk: str) -> str:
    if isinstance(risk, dict):
        title = risk.get("title") or risk.get("risk") or risk.get("name") or "Risk"
        description = risk.get("description") or risk.get("short_description") or risk.get("detail") or ""
    else:
        title = str(risk)
        description = ""
    description_html = (
        f'<div class="risk-description">{esc(description)}</div>'
        if description else ""
    )
    return f"""
    <div class="risk-card">
        <span class="risk-number">{number:02d}</span>
        <div class="risk-copy">
            <div class="risk-title">{esc(title)}</div>
            {description_html}
        </div>
    </div>
    """


def render_evidence_card(benefit: str, source: str = "", page=None) -> str:
    meta = ""
    if source:
        page_part = f"&nbsp;·&nbsp; Page {esc(page)}" if page else ""
        meta = f"""
        <div class="source">
            <span class="verified-dot"></span>
            {esc(source)} {page_part}
        </div>
        """
    return f"""
    <div class="evidence">
        <div class="evidence-benefit">✓ {esc(benefit)}</div>
        {meta}
    </div>
    """


def render_audit_status(status: str, pass_rate) -> str:
    if status == "PASSED":
        return f"""
        <div class="audit-banner audit-pass">
            <div class="audit-banner-title">✓ Audit Passed</div>
            <div class="audit-banner-body">
                {esc(pass_rate)}% of generated claims were directly
                supported by policy evidence.
            </div>
        </div>
        """
    return f"""
    <div class="audit-banner audit-warning">
        <div class="audit-banner-title">⚠ Advisor Review Required</div>
        <div class="audit-banner-body">
            Audit pass rate: <b>{esc(pass_rate)}%</b><br>
            Claims that are not fully supported should be reviewed
            before client delivery.
        </div>
    </div>
    """


# ============================================================
# THEME-AWARE APP STYLES
# ============================================================

THEME_COLORS = {
    "light": {
        "bg": "#F5F7F9",
        "bg-secondary": "#EEF1F4",
        "card": "#FFFFFF",
        "card-hover": "#F8F9FB",
        "border": "#D8DEE5",
        "text-primary": "#202B35",
        "text-secondary": "#475467",
        "text-muted": "#667085",
        "success": "#16794A",
        "success-soft": "#EAF6EF",
        "warning": "#9A5B00",
        "warning-soft": "#FFF4E5",
        "error": "#B42332",
        "error-soft": "#FEF0F0",
        "accent": "#C8102E",
        "accent-strong": "#A90D27",
        "accent-soft": "#FFF0F2",
        "button-text": "#FFFFFF",
        "shadow": "rgba(16, 24, 40, 0.06)",
    },
    "dark": {
        "bg": "#111820",
        "bg-secondary": "#19232D",
        "card": "#202C37",
        "card-hover": "#2A3743",
        "border": "#3A4856",
        "text-primary": "#F4F7FA",
        "text-secondary": "#D1D9E0",
        "text-muted": "#AAB5C0",
        "success": "#62D394",
        "success-soft": "#163A2A",
        "warning": "#F4C36B",
        "warning-soft": "#44331C",
        "error": "#FF8993",
        "error-soft": "#472329",
        "accent": "#F06A7D",
        "accent-strong": "#D94A60",
        "accent-soft": "#40232B",
        "button-text": "#FFFFFF",
        "shadow": "rgba(0, 0, 0, 0.22)",
    },
}

theme_preference = str(st.session_state.get("ui_theme") or "Light").lower()
if theme_preference == "system":
    active_theme = getattr(st.context.theme, "type", None)
    active_theme = "dark" if active_theme == "dark" else "light"
else:
    active_theme = "dark" if theme_preference == "dark" else "light"
theme_variables = "\n".join(
    f"--{name}: {value};"
    for name, value in THEME_COLORS[active_theme].items()
)

render_html(
    """
    <style>

    :root {
        __THEME_VARIABLES__
    }

    /* ---------- GLOBAL ---------- */

    .stApp {
        background-color: var(--bg);
        color: var(--text-primary);
    }

    .main .block-container {
        padding-top: 1.1rem;
        padding-bottom: 2.5rem;
        max-width: 1450px;
    }

    body, .stApp, .stMarkdown, p, span, label, li, h1, h2, h3, h4, h5 {
        color: var(--text-primary);
    }

    h1, h2, h3, h4, h5 {
        font-family: Cambria, Georgia, serif;
    }

    a { color: var(--accent); }

    :focus-visible {
        outline: 2px solid var(--accent) !important;
        outline-offset: 2px;
    }

    /* ---------- SIDEBAR ---------- */

    [data-testid="stSidebar"] {
        background-color: var(--card);
        border-right: 1px solid var(--border);
    }

    [data-testid="stSidebar"] * {
        color: var(--text-primary);
    }

    [data-testid="stSidebar"] .stButton button {
        background-color: var(--bg-secondary);
        border: 1px solid var(--border);
        color: var(--text-primary);
    }

    [data-testid="stSidebar"] .stButton button:hover {
        background-color: var(--card-hover);
        border-color: var(--accent);
    }

    /* ---------- INPUTS ---------- */

    .stTextInput input,
    .stMultiSelect div[data-baseweb="select"] > div {
        background-color: var(--card) !important;
        border: 1px solid var(--border) !important;
        color: var(--text-primary) !important;
        border-radius: 7px !important;
    }

    .stMultiSelect [data-baseweb="tag"] {
        background-color: var(--accent-soft) !important;
        color: var(--accent-strong) !important;
    }

    .stMultiSelect [data-baseweb="tag"] span,
    .stMultiSelect [data-baseweb="tag"] svg {
        color: var(--accent-strong) !important;
        fill: var(--accent-strong) !important;
    }

    .stTextInput input::placeholder {
        color: var(--text-muted) !important;
    }

    label, .stTextInput label, .stMultiSelect label {
        color: var(--text-secondary) !important;
        font-size: 13px !important;
    }

    [role="listbox"], [role="option"] {
        background: var(--card) !important;
        color: var(--text-primary) !important;
    }

    [role="option"]:hover,
    [role="option"][aria-selected="true"] {
        background: var(--accent-soft) !important;
        color: var(--text-primary) !important;
    }

    /* Primary action button */

    .stButton button[kind="primary"] {
        background: var(--accent);
        color: var(--button-text);
        font-weight: 700;
        border: none;
        border-radius: 7px;
        min-height: 42px;
    }

    .stButton button[kind="primary"]:hover {
        background: var(--accent-strong);
    }

    .stButton button[kind="primary"] span,
    .stButton button[kind="primary"] svg {
        color: var(--button-text) !important;
        fill: var(--button-text) !important;
    }

    .stDownloadButton button {
        background-color: var(--card);
        border: 1px solid var(--border);
        color: var(--text-primary);
        border-radius: 7px;
        font-weight: 600;
    }

    .stDownloadButton button svg,
    .stButton button svg {
        color: currentColor;
        fill: currentColor;
    }

    .stDownloadButton button:hover {
        background-color: var(--card-hover);
        border-color: var(--accent);
    }

    /* ---------- HERO ---------- */

    .hero {
        background: var(--card);
        border: 1px solid var(--border);
        border-left: 4px solid var(--accent);
        padding: 20px 24px;
        border-radius: 8px;
        margin-bottom: 22px;
        box-shadow: 0 3px 12px var(--shadow);
    }

    .hero-title {
        font-size: 29px;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: 5px;
    }

    .hero-subtitle {
        font-size: 15px;
        color: var(--text-secondary);
        max-width: 760px;
    }

    .hero-badges {
        margin-top: 12px;
    }

    .hero-badge {
        display: inline-block;
        margin-right: 6px;
        padding: 5px 9px;
        border-radius: 5px;
        background: var(--accent-soft);
        border: 1px solid var(--border);
        color: var(--accent-strong);
        font-size: 12px;
        font-weight: 600;
    }

    /* ---------- SECTION HEADERS ---------- */

    .section-title {
        font-size: 20px;
        font-weight: 700;
        color: var(--text-primary);
        margin-top: 14px;
        margin-bottom: 3px;
    }

    .section-description {
        color: var(--text-muted);
        font-size: 14px;
        margin-bottom: 15px;
    }

    /* ---------- CARDS ---------- */

    .card {
        background: var(--card);
        border-radius: 7px;
        padding: 16px 18px;
        border: 1px solid var(--border);
        margin-bottom: 12px;
        box-shadow: 0 2px 8px var(--shadow);
    }

    .card-title {
        font-size: 12px;
        font-weight: 650;
        color: var(--text-muted);
        text-transform: uppercase;
        letter-spacing: 0.6px;
        margin-bottom: 8px;
    }

    .card-value {
        font-size: 24px;
        font-weight: 750;
        color: var(--text-primary);
        line-height: 1.2;
    }

    .card-small {
        font-size: 13px;
        color: var(--text-muted);
        margin-top: 5px;
    }

    .metric-card {
        min-height: 108px;
    }

    /* ---------- DOCUMENT CHIPS ---------- */

    .doc-chip {
        background: var(--bg-secondary);
        border: 1px solid var(--border);
        border-radius: 6px;
        padding: 9px 11px;
        font-size: 13px;
        color: var(--text-secondary);
        margin-bottom: 8px;
    }

    .doc-icon {
        color: var(--accent);
        font-weight: 700;
        margin-right: 6px;
    }

    /* ---------- RISK CARDS ---------- */

    .risk-card {
        background: var(--card);
        border-left: 3px solid var(--accent);
        border-radius: 6px;
        padding: 12px 14px;
        margin-bottom: 8px;
        border-top: 1px solid var(--border);
        border-right: 1px solid var(--border);
        border-bottom: 1px solid var(--border);
        color: var(--text-secondary);
        display: flex;
        align-items: flex-start;
        gap: 10px;
    }
    .risk-title {
        color: var(--text-primary);
        font-size: 14px;
        font-weight: 650;
        line-height: 1.35;
    }

    .risk-description {
        color: var(--text-secondary);
        font-size: 12px;
        line-height: 1.45;
        margin-top: 3px;
    }

    .profile-overview {
        color: var(--text-secondary);
        font-size: 13px;
        line-height: 1.6;
        margin-top: 10px;
    }

    .profile-note {
        background: var(--warning-soft);
        border-left: 3px solid var(--warning);
        border-radius: 4px;
        color: var(--text-secondary);
        font-size: 12px;
        line-height: 1.45;
        margin-top: 8px;
        padding: 8px 10px;
    }

    .risk-number {
        font-weight: 700;
        color: var(--accent);
        margin-right: 8px;
    }

    /* ---------- POLICY CARD ---------- */

    .policy-card {
        background: var(--card);
        border: 1px solid var(--border);
        color: var(--text-primary);
        border-left: 4px solid var(--accent);
        border-radius: 7px;
        padding: 20px 22px;
        box-shadow: 0 3px 12px var(--shadow);
        margin-bottom: 18px;
    }

    .policy-label {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: var(--text-muted);
    }

    .policy-name {
        font-size: 25px;
        font-weight: 750;
        margin-top: 5px;
        color: var(--text-primary);
    }

    .confidence {
        display: inline-block;
        margin-top: 12px;
        padding: 5px 12px;
        border-radius: 5px;
        background: var(--success-soft);
        border: 1px solid var(--border);
        color: var(--success);
        font-size: 12px;
        font-weight: 600;
    }

    /* ---------- AUDIT BANNERS ---------- */

    .audit-banner {
        border-radius: 12px;
        padding: 18px 20px;
        margin-top: 8px;
    }

    .audit-banner-title {
        font-size: 16px;
        font-weight: 700;
        margin-bottom: 6px;
    }

    .audit-banner-body {
        font-size: 14px;
        line-height: 1.5;
    }

    .audit-pass {
        background: var(--success-soft);
        border: 1px solid var(--border);
        color: var(--text-primary);
    }

    .audit-pass .audit-banner-title { color: var(--success); }

    .audit-warning {
        background: var(--warning-soft);
        border: 1px solid var(--border);
        color: var(--text-primary);
    }

    .audit-warning .audit-banner-title { color: var(--warning); }

    /* ---------- EVIDENCE ---------- */

    .evidence {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 6px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }

    .evidence-benefit {
        color: var(--text-primary);
        font-size: 14px;
        font-weight: 600;
    }

    .source {
        color: var(--text-muted);
        font-size: 12px;
        margin-top: 8px;
    }

    .verified-dot {
        display: inline-block;
        width: 7px;
        height: 7px;
        border-radius: 50%;
        background: var(--success);
        margin-right: 6px;
    }

    .claim-heading {
        align-items: center;
        display: flex;
        gap: 7px;
        margin: 8px 0 4px;
    }

    .claim-status-icon {
        align-items: center;
        border-radius: 50%;
        display: inline-flex;
        font-size: 11px;
        font-weight: 700;
        height: 20px;
        justify-content: center;
        width: 20px;
    }

    .claim-status-supported {
        background: var(--success-soft);
        color: var(--success);
    }

    .claim-status-partial {
        background: var(--warning-soft);
        color: var(--warning);
    }

    .claim-status-unsupported {
        background: var(--error-soft);
        color: var(--error);
    }

    .claim-status-review {
        background: var(--bg-secondary);
        color: var(--text-secondary);
    }

    /* ---------- SIDEBAR WORKFLOW ---------- */

    .brand-title {
        font-size: 24px;
        font-weight: 750;
        letter-spacing: 1px;
        margin-bottom: 2px;
        color: var(--text-primary);
    }

    .brand-subtitle {
        font-size: 12px;
        color: var(--text-muted);
        margin-bottom: 22px;
    }

    .workflow-label {
        font-size: 11px;
        text-transform: uppercase;
        letter-spacing: 0.8px;
        color: var(--text-muted);
        margin-bottom: 10px;
        font-weight: 650;
    }

    .workflow-heading {
        display: flex;
        align-items: center;
        justify-content: space-between;
    }

    .workflow-count {
        color: var(--text-muted);
        font-size: 11px;
        font-weight: 600;
    }

    .workflow-step {
        display: flex;
        align-items: center;
        gap: 10px;
        padding: 9px 10px;
        border-radius: 6px;
        margin-bottom: 4px;
        font-size: 13px;
        color: var(--text-secondary);
        position: relative;
    }

    .workflow-step.active {
        background: var(--accent-soft);
        color: var(--text-primary);
        font-weight: 600;
    }

    .workflow-num {
        font-size: 11px;
        font-weight: 700;
        color: var(--text-muted);
        min-width: 18px;
    }

    .workflow-step.active .workflow-num,
    .workflow-step.done .workflow-num {
        color: var(--accent-strong);
    }

    .workflow-check {
        color: var(--success);
    }

    .workflow-progress-track {
        height: 4px;
        overflow: hidden;
        background: var(--bg-secondary);
        border-radius: 4px;
        margin: 8px 0 12px;
    }

    .workflow-progress-fill {
        height: 100%;
        background: var(--accent);
        border-radius: inherit;
    }

    /* ---------- STATUS DOT ---------- */

    .status-dot {
        display: inline-block;
        width: 8px;
        height: 8px;
        border-radius: 50%;
        background: var(--success);
        margin-right: 6px;
    }

    /* ---------- MISC ---------- */

    hr { border-color: var(--border) !important; }

    [data-testid="stExpander"] {
        background: var(--card);
        border: 1px solid var(--border);
        border-radius: 7px;
    }

    [data-testid="stProgressBar"] > div > div {
        background: var(--accent) !important;
    }

    [data-testid="stProgressBar"] {
        color: var(--text-secondary);
    }

    [data-testid="stAlert"] {
        border-radius: 7px;
        background: var(--card);
        color: var(--text-primary);
        border: 1px solid var(--border);
    }

    [data-testid="stSegmentedControl"] {
        background: var(--bg-secondary);
        border-radius: 6px;
    }

    [data-testid="stSegmentedControl"] button {
        color: var(--text-secondary) !important;
    }

    [data-testid="stSegmentedControl"] button[aria-checked="true"] {
        background: var(--card) !important;
        color: var(--accent-strong) !important;
        box-shadow: 0 1px 4px var(--shadow);
    }

    [data-testid="stForm"] {
        border: 1px solid var(--border);
        border-radius: 8px;
        padding: 16px;
        background: var(--card);
        box-shadow: 0 2px 8px var(--shadow);
    }

    .footer {
        text-align: center;
        color: var(--text-muted);
        font-size: 12px;
        margin-top: 50px;
        padding-top: 20px;
        border-top: 1px solid var(--border);
    }

    @media (max-width: 760px) {
        .main .block-container {
            padding: 0.8rem 0.8rem 2rem;
        }

        .hero {
            padding: 16px 17px;
            margin-bottom: 16px;
        }

        .hero-title {
            font-size: 24px;
        }

        .section-title {
            font-size: 18px;
        }

        .card {
            padding: 13px 14px;
        }
    }

    </style>
    """.replace("__THEME_VARIABLES__", theme_variables)
)


# ============================================================
# SESSION STATE
# ============================================================

if "results" not in st.session_state:
    st.session_state.results = None

if "running" not in st.session_state:
    st.session_state.running = False


# ============================================================
# SIDEBAR
# ============================================================

WORKFLOW_STEPS = [
    "Client Profile",
    "Risk Analysis",
    "Policy Analysis",
    "AI Audit",
    "Advisor Review",
]

with st.sidebar:

    render_html(
        """
        <div class="brand-title">MARSH</div>
        <div class="brand-subtitle">AI Advisory Workspace</div>
        """
    )

    st.segmented_control(
        "Appearance",
        options=["Light", "System", "Dark"],
        default="Light",
        key="ui_theme",
    )

    has_results = st.session_state.results is not None

    current_step = len(WORKFLOW_STEPS) if has_results else 1
    progress_width = 100 if has_results else round(100 / len(WORKFLOW_STEPS))
    steps_html = (
        '<div class="workflow-heading">'
        '<span class="workflow-label">Workflow</span>'
        f'<span class="workflow-count">{current_step:02d} / {len(WORKFLOW_STEPS):02d}</span>'
        '</div>'
        '<div class="workflow-progress-track">'
        f'<div class="workflow-progress-fill" style="width:{progress_width}%"></div>'
        '</div>'
    )
    for i, step in enumerate(WORKFLOW_STEPS, 1):
        if has_results:
            css_class = "workflow-step done"
            marker = '<span class="workflow-check">✓</span>'
        elif i == 1:
            css_class = "workflow-step active"
            marker = f'<span class="workflow-num">{i:02d}</span>'
        else:
            css_class = "workflow-step"
            marker = f'<span class="workflow-num">{i:02d}</span>'
        steps_html += f'<div class="{css_class}">{marker}<span>{esc(step)}</span></div>'

    render_html(steps_html)

    st.markdown("---")

    with st.expander("System Details"):
        st.caption("Knowledge Base")
        st.success("114 policy chunks indexed")

        st.caption("LLM")
        st.info("Groq · GPT-OSS 120B")

        st.caption("Vector Store")
        st.info("ChromaDB")

        st.caption("Embeddings")
        st.info("BGE-M3")

    st.markdown("---")

    if st.button("Reset workspace", width="stretch", icon=":material/refresh:"):
        st.session_state.results = None
        st.rerun()


# ============================================================
# HERO
# ============================================================

render_html(
    """
    <div class="hero">
        <div class="hero-title">Marsh AI Pitch Advisor</div>
        <div class="hero-subtitle">
            AI-powered insurance pitch generation, policy intelligence
            and content verification.
        </div>
        <div class="hero-badges">
            <span class="hero-badge">Evidence Grounded</span>
            <span class="hero-badge">AI Assisted</span>
            <span class="hero-badge">Human Reviewed</span>
        </div>
    </div>
    """
)


# ============================================================
# INPUT SECTION
# ============================================================

section_title(
    "Client Setup",
    "Enter a client and select the policy documents to analyse."
)

with st.form("client_setup", clear_on_submit=False):
    col1, col2 = st.columns([1.2, 1])

    with col1:
        company_name = st.text_input(
            "Company name",
            placeholder="e.g. Infosys, TCS, Accenture",
            value="",
        )

    with col2:
        data_folder = Path("data")
        available_pdfs = [
            file.name for file in data_folder.glob("*.pdf")
        ] if data_folder.exists() else []
        selected_pdfs = st.multiselect(
            "Policy documents",
            available_pdfs,
            default=available_pdfs,
        )

    generate = st.form_submit_button(
        "Generate client pitch",
        type="primary",
        width="stretch",
        icon=":material/auto_awesome:",
    )


# ============================================================
# DOCUMENT STATUS
# ============================================================

if selected_pdfs:

    render_html(
        f"""
        <div class="card">
            <div class="card-title">Selected Policy Knowledge Base</div>
            <div class="card-value" style="font-size:16px;">
                {len(selected_pdfs)} policy document(s) selected
            </div>
            <div class="card-small">
                AI responses will be grounded against the selected
                policy evidence.
            </div>
        </div>
        """
    )

    cols = st.columns(min(len(selected_pdfs), 4))

    for i, pdf in enumerate(selected_pdfs):
        with cols[i % len(cols)]:
            render_html(
                f'<div class="doc-chip"><span class="doc-icon">PDF</span>{esc(Path(pdf).stem)}</div>'
            )


# ============================================================
# MAIN PIPELINE
# ============================================================

if generate:

    if not company_name.strip():
        st.error("Please enter a company name.")
        st.stop()

    if not selected_pdfs:
        st.error("Please select at least one policy document.")
        st.stop()

    stage_labels = [
        "Building client profile",
        "Identifying business risks",
        "Retrieving policy evidence",
        "Auditing AI claims",
        "Preparing recommendation",
        "Generating deliverables",
    ]

    progress = st.progress(0, text="Preparing AI and policy resources")

    try:

        from src.company_profile import generate_company_profile
        from src.risk_mapping import map_risks_to_policies
        from src.audit import audit_pitch_claims
        from src.audit_report import generate_audit_summary, save_audit_report
        from src.pitch import build_pitch
        from src.ppt_generator import create_pitch_ppt

        # ----------------------------------------------------
        # COMPANY PROFILE
        # ----------------------------------------------------

        progress.progress(10, text=f"01 · {stage_labels[0]}")

        profile = generate_company_profile(company_name)

        # ----------------------------------------------------
        # RISK MAPPING
        # ----------------------------------------------------

        progress.progress(28, text=f"02 · {stage_labels[1]}")

        mappings = map_risks_to_policies(
            profile["key_risks"],
            top_k=5
        )

        progress.progress(40, text=f"03 · {stage_labels[2]}")

        # ----------------------------------------------------
        # CANDIDATE BENEFITS
        # ----------------------------------------------------

        candidate_benefits = []

        for mapping in mappings:
            for benefit in mapping.get("relevant_benefits", []):
                candidate_benefits.append({
                    "risk": mapping["risk"],
                    "benefit": benefit["benefit"],
                    "policy": benefit["policy"],
                    "source": benefit["source"],
                    "page": benefit["page"],
                })

        # ----------------------------------------------------
        # AUDIT
        # ----------------------------------------------------

        progress.progress(58, text=f"04 · {stage_labels[3]}")

        claims = [item["benefit"] for item in candidate_benefits]

        audit_results = audit_pitch_claims(claims)

        # ----------------------------------------------------
        # VERIFIED BENEFITS
        # ----------------------------------------------------

        verified_benefits = []

        for benefit, audit in zip(candidate_benefits, audit_results):
            if audit["status"] == "SUPPORTED":
                verified_benefits.append({
                    "risk": benefit["risk"],
                    "benefit": audit["claim"],
                    "policy": benefit["policy"],
                    "source": audit["source"],
                    "page": audit["page"],
                })

        # ----------------------------------------------------
        # POLICY SCORING
        # ----------------------------------------------------

        progress.progress(74, text=f"05 · {stage_labels[4]}")

        policy_scores = {}

        for benefit in verified_benefits:
            policy = benefit["policy"]
            if policy not in policy_scores:
                policy_scores[policy] = []
            policy_scores[policy].append(benefit)

        if policy_scores:
            recommended_policy = max(
                policy_scores,
                key=lambda p: len(policy_scores[p])
            )
            supporting_benefits = policy_scores[recommended_policy]
            confidence = "High"
            reasons = [item["benefit"] for item in supporting_benefits]
        else:
            recommended_policy = "No sufficiently supported policy recommendation"
            supporting_benefits = []
            confidence = "Insufficient"
            reasons = ["No policy benefits passed the claim-level audit."]

        recommendation = {
            "recommended_policy": recommended_policy,
            "confidence": confidence,
            "reasons": reasons,
            "supporting_benefits": supporting_benefits,
            "limitations": [
                "Only directly supported policy evidence was used."
            ],
        }

        # ----------------------------------------------------
        # AUDIT SUMMARY
        # ----------------------------------------------------

        audit_report = generate_audit_summary(audit_results)

        # ----------------------------------------------------
        # PITCH
        # ----------------------------------------------------

        pitch = build_pitch(profile, recommendation, audit_report)

        # ----------------------------------------------------
        # SAVE FILES
        # ----------------------------------------------------

        progress.progress(90, text=f"06 · {stage_labels[5]}")

        os.makedirs("outputs", exist_ok=True)

        slug = slugify(company_name)

        audit_path = f"outputs/{slug}_audit_report.json"
        ppt_path = f"outputs/{slug}_final_pitch.pptx"

        save_audit_report(audit_report, audit_path)
        create_pitch_ppt(pitch, ppt_path)

        progress.progress(100, text="Analysis complete.")

        st.session_state.results = {
            "profile": profile,
            "mappings": mappings,
            "recommendation": recommendation,
            "audit": audit_report,
            "ppt_path": ppt_path,
            "audit_path": audit_path,
            "company_slug": slug,
        }

        st.success("Client pitch generated successfully.")

    except Exception as e:

        progress.empty()

        st.error(
            "Unable to generate the pitch. Please verify the company "
            "name and try again."
        )

        with st.expander("Technical details"):
            st.code(str(e), language="text")


# ============================================================
# RESULTS
# ============================================================

if st.session_state.results:

    results = st.session_state.results

    profile = results["profile"]
    recommendation = results["recommendation"]
    audit = results["audit"]

    st.markdown("---")

    # ========================================================
    # EXECUTIVE SUMMARY
    # ========================================================

    section_title("Executive Summary")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        render_html(render_metric_card(
            "Client", profile["company_name"], profile["industry"]
        ))

    with c2:
        render_html(render_metric_card(
            "Key Risks", str(len(profile["key_risks"])), "Identified by AI"
        ))

    with c3:
        render_html(render_metric_card(
            "Claims Audited", str(audit["total_claims"]), "Evidence checks"
        ))

    with c4:
        render_html(render_metric_card(
            "Audit Status", audit["overall_status"],
            f"Pass rate: {audit['pass_rate']}%", value_size="20px"
        ))

    # ========================================================
    # CLIENT PROFILE + RISKS
    # ========================================================

    section_title("Client Intelligence")

    left, right = st.columns([1, 1])

    with left:
        assumptions = profile.get("assumptions", []) or []
        assumption_html = "".join(
            f'<div class="profile-note">{esc(item)}</div>'
            for item in assumptions[:3]
        )
        overview_text = str(profile.get("business_overview", "")).strip()
        overview_html = (
            f'<div class="profile-overview">{esc(overview_text)}</div>'
            if overview_text else ""
        )
        profile_html = f"""
        <div class="card">
            <div class="card-title">Company Profile</div>
            <div style="margin-top:8px; font-size:14px; color:var(--text-secondary); line-height:1.9;">
                <b style="color:var(--text-primary);">Industry:</b> {esc(profile["industry"])}<br>
                <b style="color:var(--text-primary);">Company Size:</b> {esc(profile["company_size"])}
            </div>
            {overview_html}
            {assumption_html}
        </div>
        """
        render_html(profile_html)

    with right:
        risks_html = '<div class="card"><div class="card-title">Key Business Risks</div>'
        for i, risk in enumerate(profile["key_risks"], 1):
            risks_html += render_risk_card(i, risk)
        risks_html += "</div>"
        render_html(risks_html)

    # ========================================================
    # POLICY RECOMMENDATION
    # ========================================================

    section_title("Policy Intelligence")

    render_html(
        f"""
        <div class="policy-card">
            <div class="policy-label">Evidence-Based Policy Option</div>
            <div class="policy-name">{esc(recommendation["recommended_policy"])}</div>
            <div class="confidence">Confidence: {esc(recommendation["confidence"])}</div>
        </div>
        """
    )

    # ========================================================
    # REASONS
    # ========================================================

    st.markdown("##### Why this policy?")

    reasons_html = "".join(
        render_evidence_card(reason) for reason in recommendation["reasons"]
    )
    render_html(reasons_html)

    # ========================================================
    # SUPPORTING EVIDENCE
    # ========================================================

    st.markdown("##### Source Traceability")

    benefits = recommendation.get("supporting_benefits", [])

    if benefits:
        evidence_html = "".join(
            render_evidence_card(b["benefit"], b["source"], b["page"])
            for b in benefits
        )
        render_html(evidence_html)
    else:
        st.info("No directly supported policy evidence was found.")

    # ========================================================
    # AUDIT DASHBOARD
    # ========================================================

    st.markdown("---")
    section_title("AI Content Audit")

    a1, a2, a3, a4 = st.columns(4)

    with a1:
        render_html(render_metric_card("Total Claims", str(audit["total_claims"])))

    with a2:
        render_html(render_metric_card("Supported", str(audit["supported"])))

    with a3:
        render_html(render_metric_card("Partially supported", str(audit["partially_supported"])))

    with a4:
        render_html(render_metric_card("Unsupported", str(audit["unsupported"])))

    render_html(render_audit_status(audit["overall_status"], audit["pass_rate"]))

    # ========================================================
    # CLAIM DETAILS
    # ========================================================

    with st.expander("Detailed claim audit"):

        for i, result in enumerate(audit["claim_results"], 1):

            status = result["status"]

            if status == "SUPPORTED":
                icon = "✓"
                icon_class = "claim-status-supported"
            elif status == "PARTIALLY_SUPPORTED":
                icon = "!"
                icon_class = "claim-status-partial"
            elif status == "UNSUPPORTED":
                icon = "×"
                icon_class = "claim-status-unsupported"
            else:
                icon = "?"
                icon_class = "claim-status-review"

            render_html(
                f'<div class="claim-heading"><span class="claim-status-icon {icon_class}">'
                f'{icon}</span><strong>Claim {i} · {esc(status)}</strong></div>'
            )
            st.write(result["claim"])
            st.write(f"**Confidence:** {result['confidence']}")
            st.write(f"**Explanation:** {result['explanation']}")

            if result.get("source"):
                st.write(f"**Source:** {result['source']}")

            if result.get("page"):
                st.write(f"**Page:** {result['page']}")

            if result.get("evidence"):
                st.code(result["evidence"], language="text")

            st.divider()

    # ========================================================
    # ADVISOR WORKFLOW
    # ========================================================

    section_title("Advisor Review")

    review1, review2, review3 = st.columns(3)

    review_steps = [
        ("01 · Review", "Inspect claims flagged by the AI audit."),
        ("02 · Verify", "Confirm source document and policy clause."),
        ("03 · Approve", "Edit or remove unsupported claims before delivery."),
    ]

    for col, (title, body) in zip((review1, review2, review3), review_steps):
        with col:
            render_html(f"""
            <div class="card">
                <div class="card-title">{esc(title)}</div>
                <div style="font-size:13px; color:var(--text-secondary);">{esc(body)}</div>
            </div>
            """)

    # ========================================================
    # DOWNLOADS
    # ========================================================

    section_title("Deliverables")

    d1, d2 = st.columns(2)

    slug = results.get("company_slug", slugify(profile["company_name"]))

    with d1:
        ppt_path = results["ppt_path"]
        if os.path.exists(ppt_path):
            with open(ppt_path, "rb") as file:
                st.download_button(
                    "Download client pitch",
                    data=file,
                    file_name=f"{slug}_final_pitch.pptx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "presentationml.presentation"
                    ),
                    width="stretch",
                    icon=":material/slideshow:",
                )

    with d2:
        audit_path = results["audit_path"]
        if os.path.exists(audit_path):
            with open(audit_path, "rb") as file:
                st.download_button(
                    "Download audit report",
                    data=file,
                    file_name=f"{slug}_audit_report.json",
                    mime="application/json",
                    width="stretch",
                    icon=":material/verified_user:",
                )


# ============================================================
# FOOTER
# ============================================================

render_html(
    """
    <div class="footer">
        Marsh AI Pitch Advisor · Evidence-grounded insurance
        intelligence · Human-in-the-loop review
    </div>
    """
)
