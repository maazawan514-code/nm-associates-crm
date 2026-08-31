"""
NM Associates CRM - Lead Referral & Team Sheet Dashboard (Google Sheets backend)
---------------------------------------------------------------------------------
Ye version data ko SQLite ki bajaye Google Sheets mein save karta hai, taake:
  - Data kabhi delete na ho (app restart/redeploy hone par bhi)
  - Aap khud bhi Google Sheet kholkar data dekh/edit kar sakein
  - Multiple team members ka data ek hi jagah, hamesha safe rahe

SETUP (ek dafa karna hai):
  1. Google Cloud Console (console.cloud.google.com) par jayein.
  2. Naya project banayein (ya purana use karein).
  3. "Google Sheets API" aur "Google Drive API" dono ko ENABLE karein.
  4. "APIs & Services" -> "Credentials" -> "Create Credentials" -> "Service Account".
  5. Service account bana lene ke baad, uski "Keys" tab mein jayein aur
     "Add Key" -> "Create new key" -> JSON. Ye JSON file download ho jayegi.
  6. Google Sheets par jayein, ek nayi Google Sheet banayein (naam kuch bhi rakhein,
     masalan "NM Associates Leads"). Us sheet ko OPEN karein aur "Share" button se
     us email ko Editor access dein jo JSON file mein "client_email" field mein
     likha hai (kuch is tarah: xxxx@xxxx.iam.gserviceaccount.com).
  7. Sheet ke URL se uski ID nikal lein. URL is tarah ki hoti hai:
     https://docs.google.com/spreadsheets/d/SPREADSHEET_ID_YAHAN_HAI/edit
  8. Streamlit Cloud app -> Settings -> Secrets mein ye format se daal dein:

     SPREADSHEET_ID = "yahan_apni_sheet_ki_id_dalein"
     ADMIN_PASSWORD = "apna_password_yahan_dalein"

     [gcp_service_account]
     type = "service_account"
     project_id = "..."
     private_key_id = "..."
     private_key = "-----BEGIN PRIVATE KEY-----\\n...\\n-----END PRIVATE KEY-----\\n"
     client_email = "...iam.gserviceaccount.com"
     client_id = "..."
     auth_uri = "https://accounts.google.com/o/oauth2/auth"
     token_uri = "https://oauth2.googleapis.com/token"
     auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
     client_x509_cert_url = "..."

     (Ye sab values downloaded JSON file se copy honi hain - JSON file kholkar
     har field ka value yahan paste kar dein.)

  Bas itna karne ke baad app khud hi sheet mein "Leads" naam ka worksheet
  bana lega aur data save karna shuru kar dega.
"""

import datetime as dt

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials

# ----------------------------------------------------------------------
# BASIC CONFIG
# ----------------------------------------------------------------------
ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")

TEAM_MEMBERS = [
    "Faheem",
    "Usman",
    "Ahmad",
    "Abbas",
    "Abdul Mannan Bhatti",
]

STATUS_OPTIONS = ["New", "Contacted", "In Progress", "Follow-up", "Closed - Won", "Closed - Lost"]

HEADERS = [
    "client_name", "phone_number", "requirement", "referred_to",
    "referred_date", "status", "notes", "added_by", "created_at",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

st.set_page_config(page_title="NM Associates CRM", page_icon="📋", layout="wide")


# ----------------------------------------------------------------------
# GOOGLE SHEETS CONNECTION
# ----------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_worksheet():
    if "gcp_service_account" not in st.secrets or "SPREADSHEET_ID" not in st.secrets:
        return None
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    client = gspread.authorize(creds)
    sh = client.open_by_key(st.secrets["SPREADSHEET_ID"])
    try:
        ws = sh.worksheet("Leads")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title="Leads", rows=2000, cols=len(HEADERS))
        ws.append_row(HEADERS)
    # agar sheet bilkul khali hai to header daal dein
    if not ws.get_all_values():
        ws.append_row(HEADERS)
    return ws


def get_all_leads():
    ws = get_worksheet()
    if ws is None:
        return pd.DataFrame(columns=HEADERS + ["row_number"])
    records = ws.get_all_records()
    if not records:
        return pd.DataFrame(columns=HEADERS + ["row_number"])
    df = pd.DataFrame(records)
    df["row_number"] = range(2, len(df) + 2)  # row 1 = header
    if "referred_date" in df.columns:
        df["referred_date"] = pd.to_datetime(df["referred_date"], errors="coerce").dt.date
    return df


def add_lead(client_name, phone_number, requirement, referred_to, referred_date, status, notes, added_by):
    ws = get_worksheet()
    ws.append_row([
        client_name.strip(),
        phone_number.strip(),
        requirement.strip(),
        referred_to,
        str(referred_date),
        status,
        notes.strip() if notes else "",
        added_by,
        dt.datetime.now().isoformat(timespec="seconds"),
    ])


def update_lead(row_number, status=None, notes=None):
    ws = get_worksheet()
    if status is not None:
        ws.update_cell(row_number, HEADERS.index("status") + 1, status)
    if notes is not None:
        ws.update_cell(row_number, HEADERS.index("notes") + 1, notes)


def delete_lead(row_number):
    ws = get_worksheet()
    ws.delete_rows(row_number)


# ----------------------------------------------------------------------
# SESSION STATE
# ----------------------------------------------------------------------
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

# ----------------------------------------------------------------------
# CONNECTION CHECK
# ----------------------------------------------------------------------
if "gcp_service_account" not in st.secrets or "SPREADSHEET_ID" not in st.secrets:
    st.error(
        "⚠️ Google Sheets se connection set nahi hua. Streamlit Cloud app ki "
        "**Settings → Secrets** mein `SPREADSHEET_ID`, `ADMIN_PASSWORD` aur "
        "`[gcp_service_account]` details daalein. App ke top comment mein poori "
        "instructions likhi hain."
    )
    st.stop()

# ----------------------------------------------------------------------
# SIDEBAR - ROLE SELECTION
# ----------------------------------------------------------------------
st.sidebar.title("📋 NM Associates CRM")
role = st.sidebar.radio("View karna hai?", ["Admin", "Team Member"], index=0)

if role == "Admin" and not st.session_state.is_admin:
    st.sidebar.markdown("---")
    pwd = st.sidebar.text_input("Admin Password", type="password")
    if st.sidebar.button("Login"):
        if pwd == ADMIN_PASSWORD:
            st.session_state.is_admin = True
            st.rerun()
        else:
            st.sidebar.error("Password ghalat hai.")

if st.session_state.is_admin:
    st.sidebar.success("Admin logged in ✅")
    if st.sidebar.button("Logout"):
        st.session_state.is_admin = False
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption("Team Members: " + ", ".join(TEAM_MEMBERS))
st.sidebar.caption("Data source: Google Sheets 🔗")

# ========================================================================
# ADMIN PANEL
# ========================================================================
if role == "Admin" and st.session_state.is_admin:
    st.title("Admin Panel")

    tab1, tab2 = st.tabs(["➕ Add New Referral", "📊 Team Sheet / All Leads"])

    # ---------------- TAB 1: ADD NEW REFERRAL ----------------
    with tab1:
        st.subheader("Naya Lead Refer Karein")
        with st.form("add_lead_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                client_name = st.text_input("Client ka Naam *")
                phone_number = st.text_input("Client ka Number *")
                referred_to = st.selectbox("Kisko Refer Kar Rahe Hain *", TEAM_MEMBERS)
            with col2:
                referred_date = st.date_input("Refer Karne ki Date *", value=dt.date.today())
                status = st.selectbox("Status", STATUS_OPTIONS, index=0)
                added_by = st.text_input("Aap ka Naam (Admin/Referrer)", value="Admin")

            requirement = st.text_area(
                "Client ki Requirement / Detail *",
                placeholder="Yahan har detail likhein - client ko kya chahiye, budget, timeline, koi khaas baat etc.",
                height=150,
            )
            notes = st.text_area("Extra Notes (optional)", height=80)

            submitted = st.form_submit_button("✅ Lead Add Karein", use_container_width=True)
            if submitted:
                if not client_name or not phone_number or not requirement:
                    st.error("Client Name, Phone Number aur Requirement zaroori hain.")
                else:
                    add_lead(
                        client_name, phone_number, requirement, referred_to,
                        referred_date, status, notes, added_by,
                    )
                    st.success(f"Lead '{client_name}' successfully {referred_to} ko refer ho gayi ✅")

    # ---------------- TAB 2: TEAM SHEET / ALL LEADS ----------------
    with tab2:
        st.subheader("Team Sheet")
        if st.button("🔄 Refresh Data"):
            st.rerun()

        df = get_all_leads()

        if df.empty:
            st.info("Abhi tak koi lead add nahi hui.")
        else:
            # ---- FILTERS ----
            fcol1, fcol2, fcol3, fcol4 = st.columns([1.2, 1.2, 1.2, 1.6])
            with fcol1:
                member_filter = st.multiselect("Team Member", TEAM_MEMBERS, default=TEAM_MEMBERS)
            with fcol2:
                status_filter = st.multiselect("Status", STATUS_OPTIONS, default=STATUS_OPTIONS)
            with fcol3:
                min_date = df["referred_date"].min()
                max_date = df["referred_date"].max()
                date_range = st.date_input("Date Range", value=(min_date, max_date))
            with fcol4:
                search_term = st.text_input("🔍 Search (Naam / Number)")

            filtered = df[
                df["referred_to"].isin(member_filter)
                & df["status"].isin(status_filter)
            ]

            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_d, end_d = date_range
                filtered = filtered[
                    (filtered["referred_date"] >= start_d) & (filtered["referred_date"] <= end_d)
                ]

            if search_term:
                s = search_term.lower()
                filtered = filtered[
                    filtered["client_name"].str.lower().str.contains(s)
                    | filtered["phone_number"].str.lower().str.contains(s)
                ]

            st.markdown(f"**Total Leads Found: {len(filtered)}**")

            st.dataframe(
                filtered.rename(columns={
                    "client_name": "Client Name",
                    "phone_number": "Phone Number",
                    "requirement": "Requirement",
                    "referred_to": "Referred To",
                    "referred_date": "Date Referred",
                    "status": "Status",
                    "notes": "Notes",
                    "added_by": "Added By",
                })[["Client Name", "Phone Number", "Requirement", "Referred To",
                    "Date Referred", "Status", "Notes", "Added By"]],
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                "⬇️ Download Team Sheet (CSV)",
                data=filtered.to_csv(index=False).encode("utf-8"),
                file_name=f"team_sheet_{dt.date.today()}.csv",
                mime="text/csv",
            )

            st.markdown("---")
            st.subheader("Lead Update / Status Change / Delete")
            if not filtered.empty:
                row_numbers = filtered["row_number"].tolist()
                selected_row = st.selectbox(
                    "Lead Select Karein",
                    row_numbers,
                    format_func=lambda x: filtered.loc[filtered['row_number'] == x, 'client_name'].values[0],
                )
                row = filtered[filtered["row_number"] == selected_row].iloc[0]

                ucol1, ucol2 = st.columns(2)
                with ucol1:
                    new_status = st.selectbox("Status Update Karein", STATUS_OPTIONS,
                                               index=STATUS_OPTIONS.index(row["status"]))
                with ucol2:
                    new_notes = st.text_area("Notes Update Karein", value=row["notes"])

                bcol1, bcol2 = st.columns(2)
                with bcol1:
                    if st.button("💾 Update Save Karein", use_container_width=True):
                        update_lead(selected_row, status=new_status, notes=new_notes)
                        st.success("Update ho gaya ✅")
                        st.rerun()
                with bcol2:
                    if st.button("🗑️ Lead Delete Karein", use_container_width=True, type="secondary"):
                        delete_lead(selected_row)
                        st.warning("Lead delete ho gayi.")
                        st.rerun()

elif role == "Admin" and not st.session_state.is_admin:
    st.title("Admin Panel")
    st.info("Admin panel dekhne ke liye sidebar mein password enter karein.")

# ========================================================================
# TEAM MEMBER VIEW (read-only, client bhi is se update le sakta hai)
# ========================================================================
else:
    st.title("📋 Team Member - Mere Leads")
    if st.button("🔄 Refresh"):
        st.rerun()

    selected_member = st.selectbox("Apna Naam Select Karein", TEAM_MEMBERS)

    df = get_all_leads()
    my_leads = df[df["referred_to"] == selected_member] if not df.empty else df

    st.markdown(f"### {selected_member} ke paas total **{len(my_leads)}** leads hain")

    if my_leads.empty:
        st.info("Abhi tak koi lead assign nahi hui.")
    else:
        status_filter = st.multiselect("Status ke hisaab se dekhein", STATUS_OPTIONS, default=STATUS_OPTIONS)
        my_leads = my_leads[my_leads["status"].isin(status_filter)]

        for _, row in my_leads.iterrows():
            with st.expander(f"📌 {row['client_name']} — {row['referred_date']} — {row['status']}"):
                st.write(f"**Phone Number:** {row['phone_number']}")
                st.write(f"**Requirement:** {row['requirement']}")
                st.write(f"**Notes:** {row['notes'] if row['notes'] else '—'}")
                st.write(f"**Added By:** {row['added_by']}")
