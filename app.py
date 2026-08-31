import io
import os
import time
from dataclasses import dataclass, asdict
from typing import Optional

import pandas as pd
import requests
import streamlit as st

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------

LUSHA_ENDPOINT = "https://api.lusha.com/v3/contacts/search-and-enrich"
REQUIRED_COLUMNS = ["Full Name", "Company Name", "Role"]
REQUEST_TIMEOUT_SECONDS = 20
REQUEST_PAUSE_SECONDS = 0.15  # small pause between calls to be polite to the API
NOT_FOUND = "Not found"

# --------------------------------------------------------------------------
# "Existing contacts" tab configuration
# --------------------------------------------------------------------------

CONTACTS_XLSX_PATH = "contacts.xlsx"  # looked up relative to this script's working dir
CONTACTS_REQUIRED_COLUMNS = [
    "Nom Entreprise",
    "Prénom Contact",
    "NOM Contact",
    "TEL STANDARD",
    "TEL CONTACT Mobile",
    "Email",
    "Ville",
    "Commentaire",
    "Fonction",
]

# Color coding for the "Commentaire" column. Each card's left border + tint
# reflects which of these three statuses the contact has.
COMMENTAIRE_STYLES = {
    "Nouveau contact": {
        "border": "#2e7d32",       # green
        "bg": "#eef7ee",
        "bg_dark": "#1b2b1c",
        "text": "#1b5e20",
        "text_dark": "#8fd89a",
        "badge_bg": "#2e7d32",
    },
    "Doublon AKUITEO": {
        "border": "#e08a00",       # amber/orange
        "bg": "#fdf3e2",
        "bg_dark": "#332a15",
        "text": "#8a5a00",
        "text_dark": "#f3c26b",
        "badge_bg": "#e08a00",
    },
    "Info partielle/manquante sur Lusha": {
        "border": "#c62828",       # red
        "bg": "#fbeaea",
        "bg_dark": "#331616",
        "text": "#b71c1c",
        "text_dark": "#f19a9a",
        "badge_bg": "#c62828",
    },
}
DEFAULT_COMMENTAIRE_STYLE = {
    "border": "#8a8f98",           # neutral gray fallback for unrecognized values
    "bg": "#f0f1f3",
    "bg_dark": "#26272b",
    "text": "#4a4f58",
    "text_dark": "#c7cad1",
    "badge_bg": "#8a8f98",
}


@dataclass
class EnrichedLead:
    full_name: str
    company_name: str
    role: str
    phone_number: str
    email: str

    def to_display_dict(self) -> dict:
        return {
            "Full Name": self.full_name,
            "Company Name": self.company_name,
            "Role": self.role,
            "Phone Number": self.phone_number,
            "Email": self.email,
        }


# --------------------------------------------------------------------------
# API key resolution
# --------------------------------------------------------------------------

def get_api_key(sidebar_value: Optional[str]) -> Optional[str]:
    """Resolve the Lusha API key from (in priority order) the sidebar input,
    an environment variable, or Streamlit secrets. Never hardcode a key."""
    if sidebar_value:
        return sidebar_value.strip()

    env_key = os.environ.get("LUSHA_API_KEY")
    if env_key:
        return env_key.strip()

    try:
        secret_key = st.secrets.get("LUSHA_API_KEY")
        if secret_key:
            return str(secret_key).strip()
    except Exception:
        # st.secrets raises if no secrets.toml exists at all -- that's fine.
        pass

    return None


# --------------------------------------------------------------------------
# File parsing
# --------------------------------------------------------------------------

class FileValidationError(Exception):
    """Raised when the uploaded spreadsheet is missing or malformed."""


def _clean_str_column(series: pd.Series) -> pd.Series:
    """Coerce a column to trimmed strings, treating NaN/None as ''.

    Uses fillna("") before astype(str): with pandas' newer string dtypes,
    astype(str) alone can leave missing values as an actual float NaN
    instead of the string 'nan', so a plain .replace({"nan": ""}) misses it.
    """
    return series.fillna("").astype(str).str.strip().replace({"nan": ""})


def parse_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Read the uploaded .xlsx file and validate it has the required columns.

    Returns a DataFrame with exactly the three required columns, trimmed of
    whitespace, with fully-blank rows dropped.
    """
    if uploaded_file is None:
        raise FileValidationError("No file was uploaded.")

    filename = uploaded_file.name or ""
    if not filename.lower().endswith(".xlsx"):
        raise FileValidationError(
            f"Unsupported file type for '{filename}'. Please upload an .xlsx spreadsheet."
        )

    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
    except Exception as exc:
        raise FileValidationError(
            f"Could not read '{filename}' as an Excel file. It may be corrupted "
            f"or in an unsupported format. Details: {exc}"
        ) from exc

    if df.empty:
        raise FileValidationError("The uploaded spreadsheet has no data rows.")

    # Normalize column names (strip whitespace) before matching.
    df.columns = [str(c).strip() for c in df.columns]

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise FileValidationError(
            "The spreadsheet is missing required column(s): "
            f"{', '.join(missing)}. Expected columns: {', '.join(REQUIRED_COLUMNS)}."
        )

    df = df[REQUIRED_COLUMNS].copy()

    # Drop rows where every required field is blank.
    df = df.dropna(how="all")
    for col in REQUIRED_COLUMNS:
        df[col] = _clean_str_column(df[col])

    df = df[(df["Full Name"] != "") | (df["Company Name"] != "")]
    df = df.reset_index(drop=True)

    if df.empty:
        raise FileValidationError("No usable rows were found after removing blank rows.")

    return df


def load_contacts_file(source) -> pd.DataFrame:
    """Read the 'contacts.xlsx' sheet and validate it has the required columns.

    `source` can be a file path (str) or a Streamlit UploadedFile. Returns a
    DataFrame with exactly CONTACTS_REQUIRED_COLUMNS, values coerced to
    trimmed strings, with fully-blank rows dropped.
    """
    try:
        df = pd.read_excel(source, engine="openpyxl")
    except Exception as exc:
        raise FileValidationError(
            f"Could not read the contacts spreadsheet as an Excel file. It may be "
            f"corrupted or in an unsupported format. Details: {exc}"
        ) from exc

    if df.empty:
        raise FileValidationError("The contacts spreadsheet has no data rows.")

    df.columns = [str(c).strip() for c in df.columns]

    missing = [col for col in CONTACTS_REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise FileValidationError(
            "The contacts spreadsheet is missing required column(s): "
            f"{', '.join(missing)}. Expected columns: {', '.join(CONTACTS_REQUIRED_COLUMNS)}."
        )

    df = df[CONTACTS_REQUIRED_COLUMNS].copy()
    df = df.dropna(how="all")

    for col in CONTACTS_REQUIRED_COLUMNS:
        df[col] = _clean_str_column(df[col])

    # Drop rows with no identifying info at all (no name, no company).
    df = df[(df["Prénom Contact"] != "") | (df["NOM Contact"] != "") | (df["Nom Entreprise"] != "")]
    df = df.reset_index(drop=True)

    if df.empty:
        raise FileValidationError("No usable rows were found in the contacts spreadsheet after removing blank rows.")

    return df


def split_full_name(full_name: str) -> tuple:
    """Split 'Full Name' into (first_name, last_name) for the Lusha lookup."""
    parts = full_name.strip().split()
    if not parts:
        return "", ""
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])


# --------------------------------------------------------------------------
# Lusha API integration
# --------------------------------------------------------------------------

class LushaAPIError(Exception):
    """Raised for API-level failures (auth, network, rate limit, etc.)."""


def call_lusha_api(first_name: str, last_name: str, company_name: str, api_key: str) -> dict:
    """Query the Lusha Search & Enrich API for a single contact.

    Returns a dict: {"phone": str, "email": str}, using NOT_FOUND for any
    field Lusha could not supply. Raises LushaAPIError for hard failures
    (bad auth, network error, rate limiting) so the caller can decide how to
    surface those separately from a normal "not found" result.
    """
    payload = {
        "contacts": [
            {
                "clientReferenceId": "row-1",
                "firstName": first_name,
                "lastName": last_name,
                "companyName": company_name,
            }
        ],
        "reveal": ["emails", "phones"],
    }
    headers = {
        "api_key": api_key,
        "Content-Type": "application/json",
    }

    try:
        response = requests.post(
            LUSHA_ENDPOINT, json=payload, headers=headers, timeout=REQUEST_TIMEOUT_SECONDS
        )
    except requests.exceptions.Timeout as exc:
        raise LushaAPIError("The request to Lusha timed out.") from exc
    except requests.exceptions.RequestException as exc:
        raise LushaAPIError(f"Network error while contacting Lusha: {exc}") from exc

    if response.status_code == 401 or response.status_code == 403:
        raise LushaAPIError("Lusha rejected the API key (unauthorized). Check your API key.")
    if response.status_code == 402:
        raise LushaAPIError("Lusha account is out of credits.")
    if response.status_code == 429:
        raise LushaAPIError("Lusha rate limit exceeded. Slow down or try again later.")
    if response.status_code >= 500:
        raise LushaAPIError(f"Lusha server error (HTTP {response.status_code}).")
    if response.status_code >= 400:
        raise LushaAPIError(f"Lusha API error (HTTP {response.status_code}): {response.text[:300]}")

    try:
        data = response.json()
    except ValueError as exc:
        raise LushaAPIError("Lusha returned a response that was not valid JSON.") from exc

    return parse_lusha_response(data)


def parse_lusha_response(data: dict) -> dict:
    """Extract phone + email from a Lusha search-and-enrich response.

    Written defensively: Lusha's exact response shape has changed across API
    versions, so this looks for the fields under a few plausible paths
    rather than assuming one rigid structure.
    """
    phone = NOT_FOUND
    email = NOT_FOUND

    results = data.get("results") if isinstance(data, dict) else None
    record = None
    if isinstance(results, list) and results:
        record = results[0]
    elif isinstance(data, dict) and "contact" in data:
        # Older response shapes nest a single contact object.
        record = data.get("contact", {}).get("data", data.get("contact"))

    if not isinstance(record, dict):
        return {"phone": phone, "email": email}

    if record.get("error"):
        # e.g. {"code": "NOT_FOUND", "message": "..."}
        return {"phone": phone, "email": email}

    emails = record.get("emails") or record.get("emailAddresses")
    if isinstance(emails, list) and emails:
        first_email = emails[0]
        email = first_email.get("email") if isinstance(first_email, dict) else str(first_email)
        email = email or NOT_FOUND
    elif isinstance(record.get("email"), str) and record.get("email"):
        email = record["email"]

    phones = record.get("phones") or record.get("phoneNumbers")
    if isinstance(phones, list) and phones:
        first_phone = phones[0]
        phone = first_phone.get("number") if isinstance(first_phone, dict) else str(first_phone)
        phone = phone or NOT_FOUND
    elif isinstance(record.get("phone"), str) and record.get("phone"):
        phone = record["phone"]

    return {"phone": phone or NOT_FOUND, "email": email or NOT_FOUND}


# --------------------------------------------------------------------------
# Enrichment pipeline
# --------------------------------------------------------------------------

def enrich_leads(df: pd.DataFrame, api_key: str, progress_callback=None) -> list:
    """Run the Lusha lookup for every row in df. Never raises for a single
    row's failure -- bad rows just get 'Not found' / error markers so the
    whole batch always completes."""
    leads = []
    total = len(df)

    for i, row in df.iterrows():
        full_name = row["Full Name"]
        company_name = row["Company Name"]
        role = row["Role"]
        first_name, last_name = split_full_name(full_name)

        phone, email = NOT_FOUND, NOT_FOUND
        if first_name and company_name:
            try:
                result = call_lusha_api(first_name, last_name, company_name, api_key)
                phone, email = result["phone"], result["email"]
            except LushaAPIError as exc:
                phone, email = f"Error: {exc}", f"Error: {exc}"
            except Exception as exc:  # belt-and-braces: never crash the batch
                phone, email = f"Error: {exc}", f"Error: {exc}"
            time.sleep(REQUEST_PAUSE_SECONDS)

        leads.append(
            EnrichedLead(
                full_name=full_name or NOT_FOUND,
                company_name=company_name or NOT_FOUND,
                role=role or NOT_FOUND,
                phone_number=phone,
                email=email,
            )
        )

        if progress_callback:
            progress_callback((i + 1) / total)

    return leads


# --------------------------------------------------------------------------
# CSV export
# --------------------------------------------------------------------------

def leads_to_csv_bytes(leads: list) -> bytes:
    """Serialize enriched leads (all 5 fields) to CSV bytes for download."""
    rows = [lead.to_display_dict() for lead in leads]
    df = pd.DataFrame(rows, columns=["Full Name", "Company Name", "Role", "Phone Number", "Email"])
    buffer = io.StringIO()
    df.to_csv(buffer, index=False)
    return buffer.getvalue().encode("utf-8")


# --------------------------------------------------------------------------
# UI rendering
# --------------------------------------------------------------------------

CARD_CSS = """
<style>
.lead-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(280px, 1fr));
    gap: 1.25rem;
    margin-top: 1rem;
}
.lead-card {
    background: var(--card-bg, #ffffff);
    border: 1px solid rgba(49, 51, 63, 0.12);
    border-radius: 14px;
    padding: 1.25rem 1.4rem;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.06);
    transition: box-shadow 0.15s ease, transform 0.15s ease;
}
.lead-card:hover {
    box-shadow: 0 6px 18px rgba(0, 0, 0, 0.10);
    transform: translateY(-2px);
}
.lead-card__name {
    font-size: 1.08rem;
    font-weight: 700;
    color: #1f2430;
    margin-bottom: 0.15rem;
}
.lead-card__role {
    font-size: 0.92rem;
    color: #5b6270;
    margin-bottom: 0.85rem;
    font-weight: 500;
}
.lead-card__row {
    display: flex;
    align-items: center;
    gap: 0.55rem;
    font-size: 0.9rem;
    color: #333844;
    margin-bottom: 0.4rem;
    word-break: break-word;
}
.lead-card__icon {
    flex: 0 0 auto;
    font-size: 1rem;
}
.lead-card__row--muted {
    color: #9aa0ab;
    font-style: italic;
}
.lead-card__row--error {
    color: #c0392b;
}
@media (prefers-color-scheme: dark) {
    .lead-card {
        background: #262730;
        border-color: rgba(250, 250, 250, 0.12);
        box-shadow: 0 2px 10px rgba(0, 0, 0, 0.35);
    }
    .lead-card__name { color: #f2f2f5; }
    .lead-card__role { color: #b6bac3; }
    .lead-card__row { color: #d6d8dd; }
}
</style>
"""


def _row_class(value: str) -> str:
    if value == NOT_FOUND:
        return "lead-card__row lead-card__row--muted"
    if isinstance(value, str) and value.startswith("Error:"):
        return "lead-card__row lead-card__row--error"
    return "lead-card__row"


def _escape(text: str) -> str:
    return (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def render_lead_card_html(lead: EnrichedLead) -> str:
    """Build the HTML for a single lead card.

    IMPORTANT: this must be a single line with no embedded newlines/indentation.
    When many of these get concatenated and passed to st.markdown(...,
    unsafe_allow_html=True), Streamlit runs the text through a CommonMark
    parser first. A blank/whitespace-only line followed by an indented line
    (4+ spaces) is parsed as an *indented code block*, not HTML -- which is
    exactly what happens if this template is written as an indented
    multi-line triple-quoted string. Keeping every card as one unbroken line
    sidesteps that entirely.
    """
    return (
        '<div class="lead-card">'
        f'<div class="lead-card__name">👤 {_escape(lead.full_name)}</div>'
        f'<div class="lead-card__role">🏢 {_escape(lead.company_name)} &middot; 💼 {_escape(lead.role)}</div>'
        f'<div class="{_row_class(lead.phone_number)}"><span class="lead-card__icon">📞</span>{_escape(lead.phone_number)}</div>'
        f'<div class="{_row_class(lead.email)}"><span class="lead-card__icon">✉️</span>{_escape(lead.email)}</div>'
        '</div>'
    )


def render_leads_grid(leads: list) -> None:
    """Render every enriched lead as a card inside a responsive CSS grid."""
    st.markdown(CARD_CSS, unsafe_allow_html=True)
    cards_html = "".join(render_lead_card_html(lead) for lead in leads)
    st.markdown(f'<div class="lead-grid">{cards_html}</div>', unsafe_allow_html=True)


def render_summary(leads: list) -> None:
    total = len(leads)
    found_email = sum(1 for l in leads if l.email not in (NOT_FOUND,) and not l.email.startswith("Error:"))
    found_phone = sum(1 for l in leads if l.phone_number not in (NOT_FOUND,) and not l.phone_number.startswith("Error:"))
    col1, col2, col3 = st.columns(3)
    col1.metric("Leads processed", total)
    col2.metric("Emails found", found_email)
    col3.metric("Phone numbers found", found_phone)


# --------------------------------------------------------------------------
# "Existing contacts" tab -- card rendering
# --------------------------------------------------------------------------

CONTACT_CARD_CSS = """
<style>
.contact-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(290px, 1fr));
    gap: 1.25rem;
    margin-top: 1rem;
}
.contact-card {
    position: relative;
    border-left: 6px solid var(--cc-border);
    background: var(--cc-bg);
    border-radius: 12px;
    padding: 1.1rem 1.3rem 1.2rem 1.15rem;
    box-shadow: 0 2px 10px rgba(0, 0, 0, 0.06);
    transition: box-shadow 0.15s ease, transform 0.15s ease;
}
.contact-card:hover {
    box-shadow: 0 6px 18px rgba(0, 0, 0, 0.10);
    transform: translateY(-2px);
}
.contact-card__name {
    font-size: 1.05rem;
    font-weight: 700;
    color: #1f2430;
    margin-bottom: 0.15rem;
}
.contact-card__company {
    font-size: 0.92rem;
    color: #5b6270;
    margin-bottom: 0.6rem;
    font-weight: 500;
}
.contact-card__badge {
    display: inline-block;
    font-size: 0.72rem;
    font-weight: 700;
    color: #ffffff;
    background: var(--cc-badge-bg);
    padding: 0.18rem 0.6rem;
    border-radius: 999px;
    margin-bottom: 0.75rem;
    letter-spacing: 0.02em;
}
.contact-card__row {
    display: flex;
    align-items: flex-start;
    gap: 0.55rem;
    font-size: 0.88rem;
    color: var(--cc-text);
    margin-bottom: 0.35rem;
    word-break: break-word;
}
.contact-card__row--muted {
    color: #9aa0ab;
    font-style: italic;
}
.contact-card__icon {
    flex: 0 0 auto;
    font-size: 0.95rem;
}
@media (prefers-color-scheme: dark) {
    .contact-card { background: var(--cc-bg-dark); }
    .contact-card__name { color: #f2f2f5; }
    .contact-card__company { color: #b6bac3; }
    .contact-card__row { color: var(--cc-text-dark); }
}
</style>
"""


def _cc_field(label: str, icon: str, value: str) -> str:
    """Render one row of a contact card, muted if the value is blank."""
    value = (value or "").strip()
    if not value:
        return f'<div class="contact-card__row contact-card__row--muted"><span class="contact-card__icon">{icon}</span>{label}: —</div>'
    return f'<div class="contact-card__row"><span class="contact-card__icon">{icon}</span>{_escape(value)}</div>'


def render_contact_card_html(row: pd.Series) -> str:
    """Build the HTML for a single existing-contact card, colored by 'Commentaire'.

    Kept as one unbroken line for the same reason as render_lead_card_html:
    embedded newlines + indentation get misread as a Markdown indented code
    block by st.markdown, so cards after the first render as literal text
    instead of styled HTML.
    """
    commentaire = (row.get("Commentaire") or "").strip()
    style = COMMENTAIRE_STYLES.get(commentaire, DEFAULT_COMMENTAIRE_STYLE)

    first_name = row.get("Prénom Contact") or ""
    last_name = row.get("NOM Contact") or ""
    full_name = f"{first_name} {last_name}".strip() or "Unnamed contact"
    company = row.get("Nom Entreprise") or ""

    card_style_vars = (
        f"--cc-border:{style['border']};"
        f"--cc-bg:{style['bg']};"
        f"--cc-bg-dark:{style['bg_dark']};"
        f"--cc-text:{style['text']};"
        f"--cc-text-dark:{style['text_dark']};"
        f"--cc-badge-bg:{style['badge_bg']};"
    )

    badge_html = (
        f'<div class="contact-card__badge">{_escape(commentaire)}</div>'
        if commentaire
        else ""
    )

    rows_html = "".join(
        [
            
            _cc_field("Fonction", "💼", row.get("Fonction", "")),
            _cc_field("Tél. standard", "☎️", row.get("TEL STANDARD", "")),
            _cc_field("Mobile", "📱", row.get("TEL CONTACT Mobile", "")),
            _cc_field("Email", "✉️", row.get("Email", "")),
            _cc_field("Ville", "📍", row.get("Ville", "")),
            
        ]
    )

    return (
        f'<div class="contact-card" style="{card_style_vars}">'
        f'{badge_html}'
        f'<div class="contact-card__name">👤 {_escape(full_name)}</div>'
        f'<div class="contact-card__company">🏢 {_escape(company) if company else "—"}</div>'
        f'{rows_html}'
        '</div>'
    )


def render_contacts_grid(df: pd.DataFrame) -> None:
    """Render every row of the contacts DataFrame as a colored card grid."""
    st.markdown(CONTACT_CARD_CSS, unsafe_allow_html=True)
    cards_html = "".join(render_contact_card_html(row) for _, row in df.iterrows())
    st.markdown(f'<div class="contact-grid">{cards_html}</div>', unsafe_allow_html=True)


def render_contacts_legend() -> None:
    """Small legend explaining the card color coding."""
    swatches = "".join(
        f'<span style="display:inline-flex;align-items:center;gap:0.4rem;margin-right:1.25rem;'
        f'font-size:0.85rem;color:inherit;">'
        f'<span style="width:12px;height:12px;border-radius:3px;background:{style["border"]};'
        f'display:inline-block;"></span>{_escape(label)}</span>'
        for label, style in COMMENTAIRE_STYLES.items()
    )
    st.markdown(f'<div style="margin-bottom:0.5rem;">{swatches}</div>', unsafe_allow_html=True)


def render_contacts_summary(df: pd.DataFrame) -> None:
    total = len(df)
    counts = df["Commentaire"].value_counts()
    cols = st.columns(len(COMMENTAIRE_STYLES) + 1)
    cols[0].metric("Total contacts", total)
    for i, label in enumerate(COMMENTAIRE_STYLES.keys(), start=1):
        cols[i].metric(label, int(counts.get(label, 0)))


# --------------------------------------------------------------------------
# Main app
# --------------------------------------------------------------------------

def render_enrichment_tab(api_key: Optional[str]) -> None:
    """The original 'enrich new leads via Lusha' workflow."""


    uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"], key="leads_uploader")

    #if uploaded_file is None:
    #   st.info("Upload an .xlsx file to get started.")
    #    return

    try:
        df = parse_uploaded_file(uploaded_file)
    except FileValidationError as exc:
        st.error(str(exc))
        return

    st.success(f"Loaded {len(df)} lead(s) from '{uploaded_file.name}'.")
    with st.expander("Preview uploaded data"):
        st.dataframe(df, use_container_width=True)

    file_signature = (uploaded_file.name, uploaded_file.size, len(df))

    enrich_clicked = st.button(
        "✼ Enrich Leads",
        type="primary",
        disabled=not api_key,
        help=None if api_key else "Add a Lusha API key in the sidebar first.",
    )

    if enrich_clicked:
        progress_bar = st.progress(0.0, text="Enriching leads...")

        def _progress(fraction: float):
            progress_bar.progress(fraction, text=f"Enriching leads... {int(fraction * 100)}%")

        try:
            leads = enrich_leads(df, api_key, progress_callback=_progress)
        finally:
            progress_bar.empty()

        st.session_state["enriched_leads"] = leads
        st.session_state["enriched_signature"] = file_signature

    leads = st.session_state.get("enriched_leads")
    matches_current_file = st.session_state.get("enriched_signature") == file_signature

    if leads and matches_current_file:
        st.divider()
        render_summary(leads)
        render_leads_grid(leads)

        st.divider()
        csv_bytes = leads_to_csv_bytes(leads)
        st.download_button(
            label="⬇️ Download CSV",
            data=csv_bytes,
            file_name="enriched_leads.csv",
            mime="text/csv",
        )
    elif leads and not matches_current_file:
        st.info("A new file was uploaded. Click 'Enrich Leads' to process it.")


def render_existing_contacts_tab() -> None:
    """Displays every contact in 'contacts.xlsx' as a color-coded card,
    colored according to the 'Commentaire' column."""


    df = None

    # Try to load contacts.xlsx from the working directory first.
    if os.path.exists(CONTACTS_XLSX_PATH):
        try:
            df = load_contacts_file(CONTACTS_XLSX_PATH)
        except FileValidationError as exc:
            st.error(str(exc))
    else:
        st.info(
            f"'{CONTACTS_XLSX_PATH}' was not found next to the app. "
            "Showing historical contacts."
        )

    # Fallback / override: let the user upload the file manually.
    with st.expander("Show Historical Contacts", expanded=df is None):
        uploaded_contacts = st.file_uploader(
            "Contacts spreadsheet", type=["xlsx"], key="contacts_uploader"
        )
        if uploaded_contacts is not None:
            try:
                df = load_contacts_file(uploaded_contacts)
            except FileValidationError as exc:
                st.error(str(exc))
                df = None

    if df is None:
        return

    st.success(f"Loaded {len(df)} contacts.")

    render_contacts_summary(df)

    st.divider()

    # Filter by Commentaire status.
    status_options = list(COMMENTAIRE_STYLES.keys())
    present_statuses = [s for s in status_options if s in set(df["Commentaire"])]
    other_statuses = sorted(set(df["Commentaire"]) - set(status_options) - {""})
    all_options = present_statuses + other_statuses

    selected_statuses = st.multiselect(
        "Filter by Commentaire",
        options=all_options,
        default=all_options,
    )

    search_query = st.text_input("Search (name, company, ville, email...)", "")

    filtered_df = df[df["Commentaire"].isin(selected_statuses)] if selected_statuses else df.iloc[0:0]

    if search_query.strip():
        q = search_query.strip().lower()
        mask = filtered_df.apply(
            lambda row: q in " ".join(str(v) for v in row.values).lower(), axis=1
        )
        filtered_df = filtered_df[mask]

    render_contacts_legend()

    if filtered_df.empty:
        st.warning("No contacts match the current filters.")
        return

    st.caption(f"Showing {len(filtered_df)} of {len(df)} contact(s).")
    render_contacts_grid(filtered_df)


def main():
    st.set_page_config(page_title="CRM Contact Enrichment", page_icon="✼", layout="wide")

    st.title("⇮ CRM Contact Enrichment")

    with st.sidebar:
        st.subheader("Lusha API key")
        sidebar_key = st.text_input(
            "API key",
            type="password",
            help=(
                "You can also set this via the LUSHA_API_KEY environment variable "
                "or st.secrets['LUSHA_API_KEY'] instead of entering it here."
            ),
        )
        api_key = get_api_key(sidebar_key)
        if api_key:
            st.success("API key loaded.")
        else:
            st.warning("No API key configured yet.")

        st.divider()
        st.subheader("Spreadsheet format")
        st.write("Required columns (exact names):")
        st.code("Full Name | Company Name | Role")

    tab_enrich, tab_existing = st.tabs(["🚀 Enrich New Leads", "📇 Existing Contacts"])

    with tab_enrich:
        render_enrichment_tab(api_key)

    with tab_existing:
        render_existing_contacts_tab()


if __name__ == "__main__":
    main()