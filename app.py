"""
NM Associates CRM - Google Sheets version

Features:
- Reads existing leads from the Google Sheets "Data" worksheet.
- Existing leads remain intact; no migration/deletion is performed.
- Adds new leads to the "Data" worksheet.
- Team referral tracker uses the existing "Team Referrals" worksheet.
- Admin login via Streamlit Secrets.
- Filter/search existing leads by status, priority, country, month and source.
- Filter team referrals by team member/status/date and search client.
- Update/delete leads and referral records.
- CSV downloads.
- Automatically creates missing sheets with safe headers.
"""

import datetime as dt
from typing import Optional

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials


# ----------------------------------------------------------------------
# BASIC CONFIG
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="NM Associates CRM",
    page_icon="📋",
    layout="wide",
)

ADMIN_PASSWORD = st.secrets.get("ADMIN_PASSWORD", "admin123")

TEAM_MEMBERS = [
    "Faheem",
    "Usman",
    "Ahmad",
    "Abbas",
    "Abdul Mannan Bhatti",
]

STATUS_OPTIONS = [
    "New",
    "Contacted",
    "In Progress",
    "Follow-up",
    "Closed - Won",
    "Closed - Lost",
]

PRIORITY_OPTIONS = ["Low", "Medium", "High", "Urgent"]

DATA_HEADERS = [
    "ID",
    "Name",
    "Phone Number",
    "Country",
    "Month",
    "Lead Source",
    "Status",
    "Priority",
    "Notes",
]

REFERRAL_HEADERS = [
    "Ref #",
    "Date Referred",
    "Referred To (Team Member)",
    "Client Name",
    "Client Phone Number",
    "Client Requirement / Details",
    "Status",
    "Last Update Date",
    "Update / Notes from Member",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ----------------------------------------------------------------------
# GOOGLE SHEETS CONNECTION
# ----------------------------------------------------------------------
@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    if "gcp_service_account" not in st.secrets:
        return None
    if "SPREADSHEET_ID" not in st.secrets:
        return None

    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=SCOPES,
    )
    client = gspread.authorize(creds)
    return client.open_by_key(st.secrets["SPREADSHEET_ID"])


@st.cache_resource(show_spinner=False)
def get_data_worksheet():
    sh = get_spreadsheet()
    if sh is None:
        return None

    try:
        ws = sh.worksheet("Data")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title="Data",
            rows=2000,
            cols=len(DATA_HEADERS),
        )
        ws.append_row(DATA_HEADERS)

    # If a newly created/empty sheet exists, add headers.
    if not ws.get_all_values():
        ws.append_row(DATA_HEADERS)

    return ws


@st.cache_resource(show_spinner=False)
def get_referral_worksheet():
    sh = get_spreadsheet()
    if sh is None:
        return None

    try:
        ws = sh.worksheet("Team Referrals")
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(
            title="Team Referrals",
            rows=1000,
            cols=len(REFERRAL_HEADERS),
        )
        ws.update("A1:I1", [REFERRAL_HEADERS])

    # For an existing workbook, the real referral table starts at B14.
    # If the sheet is empty/new, use A1:I1.
    values = ws.get_all_values()
    if not values:
        ws.update("A1:I1", [REFERRAL_HEADERS])
        return ws

    # Detect the existing template header.
    header_row = None
    for idx, row in enumerate(values, start=1):
        cleaned = [str(x).strip() for x in row]
        if "Ref #" in cleaned and "Client Name" in cleaned:
            header_row = idx
            break

    if header_row is None:
        # Create a clean table at the top without destroying existing content.
        # Use the first unused row after existing content.
        new_row = len(values) + 2
        ws.update(
            f"A{new_row}:I{new_row}",
            [REFERRAL_HEADERS],
        )

    return ws


# ----------------------------------------------------------------------
# HELPERS
# ----------------------------------------------------------------------
def _safe_int(value, default=0):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def get_data_table_location(ws):
    """Return (header_row, start_col) for the Data worksheet."""
    return 1, 1


def get_referral_table_location(ws):
    """
    Return (header_row, start_col) for Team Referrals.

    Existing NM Associates template:
      headers are B14:J14.

    A newly created table:
      headers are A1:I1.
    """
    values = ws.get_all_values()
    for idx, row in enumerate(values, start=1):
        cleaned = [str(x).strip() for x in row]
        if "Ref #" in cleaned and "Client Name" in cleaned:
            ref_col = cleaned.index("Ref #") + 1
            return idx, ref_col

    # If we created A1:I1 above.
    return 1, 1


def ensure_data_headers(ws):
    current = [str(x).strip() for x in ws.row_values(1)]
    if not current:
        ws.update("A1:I1", [DATA_HEADERS])
        return

    # If the sheet is the expected Data sheet, preserve existing columns.
    # Missing expected headers are added only at the end.
    missing = [h for h in DATA_HEADERS if h not in current]
    if missing:
        start = len(current) + 1
        end = start + len(missing) - 1
        ws.update(
            f"{gspread.utils.rowcol_to_a1(1, start)}:"
            f"{gspread.utils.rowcol_to_a1(1, end)}",
            [missing],
        )


def ensure_referral_headers(ws):
    header_row, start_col = get_referral_table_location(ws)
    existing = ws.row_values(header_row)

    # Ensure the table has all required headers starting at its detected column.
    existing_slice = [
        str(x).strip()
        for x in existing[start_col - 1 : start_col - 1 + len(REFERRAL_HEADERS)]
    ]
    if existing_slice != REFERRAL_HEADERS:
        start_a1 = gspread.utils.rowcol_to_a1(header_row, start_col)
        end_a1 = gspread.utils.rowcol_to_a1(
            header_row,
            start_col + len(REFERRAL_HEADERS) - 1,
        )
        ws.update(f"{start_a1}:{end_a1}", [REFERRAL_HEADERS])

    return header_row, start_col


# ----------------------------------------------------------------------
# DATA SHEET - EXISTING LEADS
# ----------------------------------------------------------------------
def get_all_leads():
    ws = get_data_worksheet()
    if ws is None:
        return pd.DataFrame(columns=DATA_HEADERS + ["row_number"])

    ensure_data_headers(ws)
    records = ws.get_all_records(
        head=1,
        expected_headers=DATA_HEADERS,
    )

    if not records:
        return pd.DataFrame(columns=DATA_HEADERS + ["row_number"])

    df = pd.DataFrame(records)

    # Make sure all expected columns exist.
    for col in DATA_HEADERS:
        if col not in df.columns:
            df[col] = ""

    df = df[DATA_HEADERS].copy()
    df["row_number"] = range(2, len(df) + 2)

    # Normalize data types for filtering/search.
    for col in DATA_HEADERS:
        if col != "ID":
            df[col] = df[col].fillna("").astype(str)

    return df


def next_lead_id(ws):
    values = ws.get_all_values()
    if len(values) <= 1:
        return 1

    id_index = DATA_HEADERS.index("ID")
    ids = []
    for row in values[1:]:
        if len(row) > id_index:
            try:
                ids.append(int(float(row[id_index])))
            except (TypeError, ValueError):
                pass

    return max(ids, default=0) + 1


def add_lead(
    name,
    phone,
    country,
    month,
    lead_source,
    status,
    priority,
    notes,
):
    ws = get_data_worksheet()
    if ws is None:
        raise RuntimeError("Google Sheets connection is not configured.")

    ensure_data_headers(ws)
    lead_id = next_lead_id(ws)

    ws.append_row(
        [
            lead_id,
            name.strip(),
            phone.strip(),
            country.strip(),
            month.strip(),
            lead_source.strip(),
            status,
            priority,
            notes.strip() if notes else "",
        ],
        value_input_option="USER_ENTERED",
    )


def update_lead(row_number, updates):
    ws = get_data_worksheet()
    if ws is None:
        raise RuntimeError("Google Sheets connection is not configured.")

    ensure_data_headers(ws)

    for column_name, value in updates.items():
        if column_name not in DATA_HEADERS or column_name == "ID":
            continue
        col_number = DATA_HEADERS.index(column_name) + 1
        ws.update_cell(row_number, col_number, value)


def delete_lead(row_number):
    ws = get_data_worksheet()
    if ws is None:
        raise RuntimeError("Google Sheets connection is not configured.")
    ws.delete_rows(row_number)


# ----------------------------------------------------------------------
# TEAM REFERRALS SHEET
# ----------------------------------------------------------------------
def get_all_referrals():
    ws = get_referral_worksheet()
    if ws is None:
        return pd.DataFrame(columns=REFERRAL_HEADERS + ["row_number"])

    header_row, start_col = ensure_referral_headers(ws)

    end_col = start_col + len(REFERRAL_HEADERS) - 1
    values = ws.get(
        f"{gspread.utils.rowcol_to_a1(header_row, start_col)}:"
        f"{gspread.utils.rowcol_to_a1(ws.row_count, end_col)}"
    )

    if not values:
        return pd.DataFrame(columns=REFERRAL_HEADERS + ["row_number"])

    data_rows = values[1:]
    data_rows = [
        row + [""] * (len(REFERRAL_HEADERS) - len(row))
        for row in data_rows
    ]
    data_rows = [row[: len(REFERRAL_HEADERS)] for row in data_rows]

    # Remove completely empty rows.
    data_rows = [
        row
        for row in data_rows
        if any(str(cell).strip() for cell in row)
    ]

    if not data_rows:
        return pd.DataFrame(columns=REFERRAL_HEADERS + ["row_number"])

    df = pd.DataFrame(data_rows, columns=REFERRAL_HEADERS)

    # Actual Google Sheet row numbers.
    row_numbers = []
    for offset, row in enumerate(values[1:], start=1):
        if any(str(cell).strip() for cell in row):
            row_numbers.append(header_row + offset)

    df["row_number"] = row_numbers

    for col in REFERRAL_HEADERS:
        df[col] = df[col].fillna("").astype(str)

    return df


def next_referral_number(ws, header_row, start_col):
    df = get_all_referrals()
    if df.empty:
        return 1

    nums = []
    for value in df["Ref #"]:
        try:
            nums.append(int(float(value)))
        except (TypeError, ValueError):
            pass

    return max(nums, default=0) + 1


def add_referral(
    date_referred,
    referred_to,
    client_name,
    phone,
    requirement,
    status,
    notes,
):
    ws = get_referral_worksheet()
    if ws is None:
        raise RuntimeError("Google Sheets connection is not configured.")

    header_row, start_col = ensure_referral_headers(ws)
    ref_number = next_referral_number(ws, header_row, start_col)

    row = [
        ref_number,
        str(date_referred),
        referred_to,
        client_name.strip(),
        phone.strip(),
        requirement.strip(),
        status,
        str(dt.date.today()),
        notes.strip() if notes else "",
    ]

    # Find first available row after header.
    existing = ws.get_all_values()
    target_row = max(header_row + 1, len(existing) + 1)

    ws.update(
        f"{gspread.utils.rowcol_to_a1(target_row, start_col)}:"
        f"{gspread.utils.rowcol_to_a1(target_row, start_col + len(row) - 1)}",
        [row],
        value_input_option="USER_ENTERED",
    )


def update_referral(row_number, status=None, notes=None):
    ws = get_referral_worksheet()
    if ws is None:
        raise RuntimeError("Google Sheets connection is not configured.")

    header_row, start_col = ensure_referral_headers(ws)
    header_to_col = {
        header: start_col + i
        for i, header in enumerate(REFERRAL_HEADERS)
    }

    if status is not None:
        ws.update_cell(row_number, header_to_col["Status"], status)

    if notes is not None:
        ws.update_cell(
            row_number,
            header_to_col["Update / Notes from Member"],
            notes,
        )

    ws.update_cell(
        row_number,
        header_to_col["Last Update Date"],
        str(dt.date.today()),
    )


def delete_referral(row_number):
    ws = get_referral_worksheet()
    if ws is None:
        raise RuntimeError("Google Sheets connection is not configured.")
    ws.delete_rows(row_number)


# ----------------------------------------------------------------------
# CONNECTION CHECK
# ----------------------------------------------------------------------
if "gcp_service_account" not in st.secrets or "SPREADSHEET_ID" not in st.secrets:
    st.error(
        "⚠️ Google Sheets connection set nahi hai. "
        "Streamlit Cloud → Settings → Secrets mein "
        "`SPREADSHEET_ID`, `ADMIN_PASSWORD` aur `[gcp_service_account]` "
        "details add karein."
    )
    st.stop()


# ----------------------------------------------------------------------
# SIDEBAR
# ----------------------------------------------------------------------
if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

st.sidebar.title("📋 NM Associates CRM")

role = st.sidebar.radio(
    "View karna hai?",
    ["Admin", "Team Member"],
    index=0,
)

if role == "Admin" and not st.session_state.is_admin:
    st.sidebar.markdown("---")
    pwd = st.sidebar.text_input(
        "Admin Password",
        type="password",
    )

    if st.sidebar.button("Login", use_container_width=True):
        if pwd == ADMIN_PASSWORD:
            st.session_state.is_admin = True
            st.rerun()
        else:
            st.sidebar.error("Password ghalat hai.")

if st.session_state.is_admin:
    st.sidebar.success("Admin logged in ✅")
    if st.sidebar.button("Logout", use_container_width=True):
        st.session_state.is_admin = False
        st.rerun()

st.sidebar.markdown("---")
st.sidebar.caption("Team Members: " + ", ".join(TEAM_MEMBERS))
st.sidebar.caption("Data source: Google Sheets 🔗")


# ----------------------------------------------------------------------
# ADMIN PANEL
# ----------------------------------------------------------------------
if role == "Admin" and st.session_state.is_admin:
    st.title("Admin Panel")
    st.caption("Google Sheets powered CRM — existing data is preserved.")

    tab1, tab2, tab3 = st.tabs(
        [
            "➕ Add New Lead / Referral",
            "📊 All Leads",
            "👥 Team Referrals",
        ]
    )

    # ------------------------------------------------------------------
    # TAB 1: ADD
    # ------------------------------------------------------------------
    with tab1:
        st.subheader("Naya Lead Add / Refer Karein")

        mode = st.radio(
            "Action",
            ["Add to Leads", "Add Team Referral"],
            horizontal=True,
        )

        if mode == "Add to Leads":
            with st.form("add_lead_form", clear_on_submit=True):
                col1, col2 = st.columns(2)

                with col1:
                    name = st.text_input("Client ka Naam *", key="lead_client_name")
                    phone = st.text_input("Client ka Number *", key="lead_client_phone")
                    country = st.text_input("Country", value="Pakistan", key="lead_country")
                    lead_source = st.text_input(
                        "Lead Source",
                        value="Referral",
                        key="lead_source",
                    )

                with col2:
                    month = st.text_input(
                        "Month",
                        value=dt.date.today().strftime("%B %Y"),
                        key="lead_month",
                    )
                    status = st.selectbox(
                        "Status",
                        STATUS_OPTIONS,
                        key="add_referral_status",
                    )
                    priority = st.selectbox(
                        "Priority",
                        PRIORITY_OPTIONS,
                        index=1,
                        key="add_lead_priority",
                    )

                notes = st.text_area(
                    "Notes",
                    placeholder="Client details, budget, timeline etc.",
                    height=120,
                    key="lead_notes",
                )

                submitted = st.form_submit_button(
                    "✅ Lead Add Karein",
                    use_container_width=True,
                )

                if submitted:
                    if not name.strip() or not phone.strip():
                        st.error("Client Name aur Phone Number zaroori hain.")
                    else:
                        try:
                            add_lead(
                                name,
                                phone,
                                country,
                                month,
                                lead_source,
                                status,
                                priority,
                                notes,
                            )
                            st.success(
                                f"Lead '{name}' successfully add ho gayi ✅"
                            )
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Lead save nahi hui: {exc}")

        else:
            with st.form("add_referral_form", clear_on_submit=True):
                col1, col2 = st.columns(2)

                with col1:
                    client_name = st.text_input("Client ka Naam *", key="lead_client_name")
                    phone = st.text_input("Client ka Number *", key="lead_client_phone")
                    referred_to = st.selectbox(
                        "Kisko Refer Kar Rahe Hain *",
                        TEAM_MEMBERS,
                        key="add_referral_member",
                    )

                with col2:
                    referred_date = st.date_input(
                        "Refer Karne ki Date *",
                        value=dt.date.today(),
                        key="ref_date",
                    )
                    status = st.selectbox(
                        "Status",
                        STATUS_OPTIONS,
                        key="add_referral_status",
                    )

                requirement = st.text_area(
                    "Client ki Requirement / Detail *",
                    key="ref_requirement",
                    placeholder=(
                        "Client ko kya chahiye, budget, timeline, "
                        "koi khaas baat etc."
                    ),
                    height=140,
                )

                notes = st.text_area(
                    "Extra Notes / Update",
                    height=90,
                    key="ref_notes",
                )

                submitted = st.form_submit_button(
                    "🤝 Referral Save Karein",
                    use_container_width=True,
                )

                if submitted:
                    if not client_name.strip() or not phone.strip() or not requirement.strip():
                        st.error(
                            "Client Name, Phone Number aur Requirement zaroori hain."
                        )
                    else:
                        try:
                            add_referral(
                                referred_date,
                                referred_to,
                                client_name,
                                phone,
                                requirement,
                                status,
                                notes,
                            )
                            st.success(
                                f"Lead '{client_name}' successfully "
                                f"{referred_to} ko refer ho gayi ✅"
                            )
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Referral save nahi hui: {exc}")

    # ------------------------------------------------------------------
    # TAB 2: ALL LEADS
    # ------------------------------------------------------------------
    with tab2:
        st.subheader("All Leads — Data Sheet")

        if st.button("🔄 Refresh Leads", key="refresh_leads"):
            st.rerun()

        df = get_all_leads()

        if df.empty:
            st.info("Data sheet mein abhi koi lead nahi mili.")
        else:
            # Filters
            # Blank cells in the existing Data sheet are treated as "Not Set"
            # so they do not accidentally hide otherwise valid leads.
            f1, f2, f3, f4 = st.columns(4)

            filter_df = df.copy()
            for _col in ["Status", "Priority", "Country"]:
                filter_df[_col] = (
                    filter_df[_col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    .replace("", "Not Set")
                )

            with f1:
                statuses = sorted(filter_df["Status"].unique().tolist())
                selected_status = st.multiselect(
                    "Status",
                    statuses,
                    default=statuses,
                    key="all_leads_status_filter",
                )

            with f2:
                priorities = sorted(filter_df["Priority"].unique().tolist())
                selected_priority = st.multiselect(
                    "Priority",
                    priorities,
                    default=priorities,
                    key="all_leads_priority_filter",
                )

            with f3:
                countries = sorted(filter_df["Country"].unique().tolist())
                selected_country = st.multiselect(
                    "Country",
                    countries,
                    default=countries,
                    key="all_leads_country_filter",
                )

            with f4:
                search = st.text_input(
                    "🔍 Search Name / Phone / Notes",
                    key="all_leads_search",
                )

            filtered = df.copy()

            # Normalize filter values so extra spaces/case differences
            # in Google Sheets do not hide valid records.
            filtered["_status_norm"] = filtered["Status"].fillna("").astype(str).str.strip().str.casefold()
            filtered["_priority_norm"] = filtered["Priority"].fillna("").astype(str).str.strip().str.casefold()
            filtered["_country_norm"] = filtered["Country"].fillna("").astype(str).str.strip().str.casefold()

            if selected_status:
                wanted = {str(x).strip().casefold() for x in selected_status}
                filtered = filtered[filtered["_status_norm"].isin(wanted)]

            if selected_priority:
                wanted = {str(x).strip().casefold() for x in selected_priority}
                filtered = filtered[filtered["_priority_norm"].isin(wanted)]

            if selected_country:
                wanted = {str(x).strip().casefold() for x in selected_country}
                filtered = filtered[filtered["_country_norm"].isin(wanted)]

            # Internal helper columns are never displayed/downloaded.

            if search.strip():
                s = search.strip().lower()
                mask = (
                    filtered["Name"].fillna("").astype(str).str.lower().str.contains(s, na=False)
                    | filtered["Phone Number"].fillna("").astype(str).str.lower().str.contains(s, na=False)
                    | filtered["Notes"].fillna("").astype(str).str.lower().str.contains(s, na=False)
                )
                filtered = filtered[mask]

            st.markdown(
                f"**Total Leads Found: {len(filtered)} / {len(df)}**"
            )

            display = filtered[
                [
                    "ID",
                    "Name",
                    "Phone Number",
                    "Country",
                    "Month",
                    "Lead Source",
                    "Status",
                    "Priority",
                    "Notes",
                ]
            ].copy()

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                "⬇️ Download Leads CSV",
                data=filtered.drop(columns=["row_number"])
                .to_csv(index=False)
                .encode("utf-8"),
                file_name=f"nm_associates_leads_{dt.date.today()}.csv",
                mime="text/csv",
                use_container_width=True,
            )

            # Update / delete
            st.markdown("---")
            st.subheader("Lead Update / Status Change / Delete")

            if not filtered.empty:
                selected_row = st.selectbox(
                    "Lead Select Karein",
                    filtered["row_number"].tolist(),
                    format_func=lambda x: (
                        f"ID {filtered.loc[filtered['row_number'] == x, 'ID'].iloc[0]} — "
                        f"{filtered.loc[filtered['row_number'] == x, 'Name'].iloc[0]}"
                    ),
                    key="lead_update_select",
                )

                row = filtered[
                    filtered["row_number"] == selected_row
                ].iloc[0]

                u1, u2 = st.columns(2)

                with u1:
                    new_status = st.selectbox(
                        "Status Update Karein",
                        STATUS_OPTIONS,
                        index=(
                            STATUS_OPTIONS.index(row["Status"])
                            if row["Status"] in STATUS_OPTIONS
                            else 0
                        ),
                        key="lead_update_status",
                    )

                    new_priority = st.selectbox(
                        "Priority Update Karein",
                        PRIORITY_OPTIONS,
                        index=(
                            PRIORITY_OPTIONS.index(row["Priority"])
                            if row["Priority"] in PRIORITY_OPTIONS
                            else 1
                        ),
                        key="lead_update_priority",
                    )

                with u2:
                    new_notes = st.text_area(
                        "Notes Update Karein",
                        value=row["Notes"],
                        key="lead_update_notes",
                    )

                b1, b2 = st.columns(2)

                with b1:
                    if st.button(
                        "💾 Update Save Karein",
                        use_container_width=True,
                    ):
                        try:
                            update_lead(
                                selected_row,
                                {
                                    "Status": new_status,
                                    "Priority": new_priority,
                                    "Notes": new_notes,
                                },
                            )
                            st.success("Lead update ho gayi ✅")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Update failed: {exc}")

                with b2:
                    if st.button(
                        "🗑️ Lead Delete Karein",
                        use_container_width=True,
                    ):
                        try:
                            delete_lead(selected_row)
                            st.warning("Lead delete ho gayi.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Delete failed: {exc}")

    # ------------------------------------------------------------------
    # TAB 3: TEAM REFERRALS
    # ------------------------------------------------------------------
    with tab3:
        st.subheader("Team Sheet / All Referrals")

        if st.button("🔄 Refresh Referrals", key="refresh_referrals"):
            st.rerun()

        rdf = get_all_referrals()

        if rdf.empty:
            st.info(
                "Team Referrals sheet mein abhi koi referral record nahi mila."
            )
        else:
            r1, r2, r3, r4 = st.columns(4)

            with r1:
                members = sorted(
                    [
                        x
                        for x in rdf["Referred To (Team Member)"].unique()
                        if str(x).strip()
                    ]
                )
                selected_members = st.multiselect(
                    "Team Member",
                    members,
                    default=members,
                    key="referrals_member_filter",
                )

            with r2:
                referral_statuses = sorted(
                    [
                        x
                        for x in rdf["Status"].unique()
                        if str(x).strip()
                    ]
                )
                selected_ref_status = st.multiselect(
                    "Status",
                    referral_statuses,
                    default=referral_statuses,
                    key="referrals_status_filter",
                )

            with r3:
                search_ref = st.text_input(
                    "🔍 Search Client / Number",
                    key="referrals_search",
                )

            with r4:
                st.metric(
                    "Total Referrals",
                    len(rdf),
                )

            filtered_r = rdf.copy()

            if selected_members:
                filtered_r = filtered_r[
                    filtered_r["Referred To (Team Member)"].isin(
                        selected_members
                    )
                ]

            if selected_ref_status:
                filtered_r = filtered_r[
                    filtered_r["Status"].isin(selected_ref_status)
                ]

            if search_ref.strip():
                s = search_ref.strip().lower()
                mask = (
                    filtered_r["Client Name"]
                    .str.lower()
                    .str.contains(s, na=False)
                    | filtered_r["Client Phone Number"]
                    .str.lower()
                    .str.contains(s, na=False)
                )
                filtered_r = filtered_r[mask]

            st.markdown(
                f"**Total Referrals Found: {len(filtered_r)} / {len(rdf)}**"
            )

            st.dataframe(
                filtered_r[
                    REFERRAL_HEADERS
                ],
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                "⬇️ Download Team Referrals CSV",
                data=filtered_r.drop(columns=["row_number"])
                .to_csv(index=False)
                .encode("utf-8"),
                file_name=f"team_referrals_{dt.date.today()}.csv",
                mime="text/csv",
                use_container_width=True,
            )

            st.markdown("---")
            st.subheader("Referral Status / Notes / Delete")

            if not filtered_r.empty:
                selected_ref_row = st.selectbox(
                    "Referral Select Karein",
                    filtered_r["row_number"].tolist(),
                    format_func=lambda x: (
                        f"Ref #{filtered_r.loc[filtered_r['row_number'] == x, 'Ref #'].iloc[0]} — "
                        f"{filtered_r.loc[filtered_r['row_number'] == x, 'Client Name'].iloc[0]}"
                    ),
                    key="referral_update_select",
                )

                selected = filtered_r[
                    filtered_r["row_number"] == selected_ref_row
                ].iloc[0]

                c1, c2 = st.columns(2)

                with c1:
                    current_status = selected["Status"]
                    new_ref_status = st.selectbox(
                        "Status Update Karein",
                        STATUS_OPTIONS,
                        index=(
                            STATUS_OPTIONS.index(current_status)
                            if current_status in STATUS_OPTIONS
                            else 0
                        ),
                        key="referral_update_status",
                    )

                with c2:
                    new_ref_notes = st.text_area(
                        "Update / Notes",
                        value=selected[
                            "Update / Notes from Member"
                        ],
                        key="referral_update_notes",
                    )

                d1, d2 = st.columns(2)

                with d1:
                    if st.button(
                        "💾 Referral Update Save Karein",
                        use_container_width=True,
                    ):
                        try:
                            update_referral(
                                selected_ref_row,
                                status=new_ref_status,
                                notes=new_ref_notes,
                            )
                            st.success("Referral update ho gayi ✅")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Referral update failed: {exc}")

                with d2:
                    if st.button(
                        "🗑️ Referral Delete Karein",
                        use_container_width=True,
                    ):
                        try:
                            delete_referral(selected_ref_row)
                            st.warning("Referral delete ho gayi.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Referral delete failed: {exc}")


# ----------------------------------------------------------------------
# TEAM MEMBER VIEW
# ----------------------------------------------------------------------
elif role == "Admin" and not st.session_state.is_admin:
    st.title("Admin Panel")
    st.info(
        "Admin panel dekhne ke liye sidebar mein password enter karein."
    )

else:
    st.title("📋 Team Member - Mere Referrals")

    if st.button("🔄 Refresh", key="refresh_team_member"):
        st.rerun()

    selected_member = st.selectbox(
        "Apna Naam Select Karein",
        TEAM_MEMBERS,
        key="team_member_select",
    )

    rdf = get_all_referrals()

    if rdf.empty:
        st.info("Abhi tak koi referral assign nahi hui.")
    else:
        my_referrals = rdf[
            rdf["Referred To (Team Member)"] == selected_member
        ].copy()

        st.markdown(
            f"### {selected_member} ke paas total "
            f"**{len(my_referrals)}** referrals hain"
        )

        if my_referrals.empty:
            st.info(
                f"Abhi tak {selected_member} ko koi referral assign nahi hui."
            )
        else:
            status_filter = st.multiselect(
                "Status ke hisaab se dekhein",
                STATUS_OPTIONS,
                default=STATUS_OPTIONS,
                key="team_member_status_filter",
            )

            my_referrals = my_referrals[
                my_referrals["Status"].isin(status_filter)
            ]

            for _, row in my_referrals.iterrows():
                with st.expander(
                    f"📌 {row['Client Name']} — "
                    f"{row['Date Referred']} — {row['Status']}"
                ):
                    st.write(
                        f"**Phone Number:** {row['Client Phone Number']}"
                    )
                    st.write(
                        f"**Requirement:** "
                        f"{row['Client Requirement / Details']}"
                    )
                    st.write(
                        f"**Last Update:** "
                        f"{row['Last Update Date'] or '—'}"
                    )
                    st.write(
                        f"**Notes:** "
                        f"{row['Update / Notes from Member'] or '—'}"
                    )
