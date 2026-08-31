"""
NM Associates - Lead Generation CRM Dashboard (Streamlit)
Run with:  streamlit run app.py
"""

import os
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="NM Associates CRM", page_icon="📊", layout="wide")

DATA_FILE = os.path.join(os.path.dirname(__file__), "NM_Associates_CRM_Dashboard.xlsx")

STATUSES = ["New", "Contacted", "Follow-up", "Interested", "Converted", "Not Interested", "Lost"]
SOURCES = ["Facebook Ads", "Instagram Ads", "WhatsApp", "Referral", "Other"]
TEAM = ["Faheem", "Usman", "Ahmad Abbas", "Abdul Mannan Bhatti", "Me / [Your Name]"]
REF_STATUSES = ["New", "Contacted", "In Progress", "Closed - Won", "Closed - Lost"]

STATUS_COLORS = {
    "New": "#2E75B6", "Contacted": "#20B2AA", "Follow-up": "#E8A33D",
    "Interested": "#9B59B6", "Converted": "#3FA96A", "Not Interested": "#D9534F", "Lost": "#6B7280",
}


@st.cache_data
def load_leads():
    if os.path.exists(DATA_FILE):
        df = pd.read_excel(DATA_FILE, sheet_name="Data")
        df = df.dropna(subset=["Name"]).reset_index(drop=True)
        df["Status"] = df["Status"].fillna("New")
        df["Lead Source"] = df["Lead Source"].fillna("Facebook Ads")
        return df
    return pd.DataFrame(columns=["ID", "Name", "Phone Number", "Country", "Month",
                                  "Lead Source", "Status", "Priority", "Notes"])


@st.cache_data
def load_referrals():
    if os.path.exists(DATA_FILE):
        try:
            raw = pd.read_excel(DATA_FILE, sheet_name="Team Referrals", header=None)
            header_row = raw[raw[1] == "Ref #"].index[0]
            df = pd.read_excel(DATA_FILE, sheet_name="Team Referrals", header=header_row, usecols="B:J")
            df = df.dropna(how="all").reset_index(drop=True)
            return df
        except Exception:
            pass
    return pd.DataFrame(columns=["Ref #", "Date Referred", "Referred To (Team Member)", "Client Name",
                                  "Client Phone Number", "Client Requirement / Details", "Status",
                                  "Last Update Date", "Update / Notes from Member"])


if "leads" not in st.session_state:
    st.session_state.leads = load_leads()
if "referrals" not in st.session_state:
    st.session_state.referrals = load_referrals()

st.title("📊 NM Associates — Lead Generation CRM")
st.caption("Meta (Facebook/Instagram) ads lead generation • Live, interactive dashboard")

tab1, tab2 = st.tabs(["📈 Leads Dashboard", "🤝 Team Referrals"])

# ============================================================
# TAB 1 — LEADS DASHBOARD
# ============================================================
with tab1:
    df = st.session_state.leads.copy()

    st.sidebar.header("Filters")
    countries = sorted(df["Country"].dropna().unique().tolist())
    country_filter = st.sidebar.multiselect("Country", countries, default=countries)
    status_filter = st.sidebar.multiselect("Status", STATUSES, default=STATUSES)
    source_filter = st.sidebar.multiselect("Lead Source", SOURCES, default=SOURCES)

    fdf = df[df["Country"].isin(country_filter) & df["Status"].isin(status_filter) & df["Lead Source"].isin(source_filter)]

    total = len(fdf)
    converted = (fdf["Status"] == "Converted").sum()
    in_progress = fdf["Status"].isin(["New", "Contacted", "Follow-up", "Interested"]).sum()
    conv_rate = (converted / total * 100) if total else 0
    top_country = fdf["Country"].mode()[0] if total else "—"

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Leads Selected", total)
    c2.metric("Converted", int(converted))
    c3.metric("In Progress", int(in_progress))
    c4.metric("Conversion Rate", f"{conv_rate:.1f}%")
    c5.metric("Top Country", top_country)

    st.divider()

    colA, colB = st.columns(2)
    with colA:
        status_counts = fdf["Status"].value_counts().reindex(STATUSES, fill_value=0).reset_index()
        status_counts.columns = ["Status", "Count"]
        fig = px.pie(status_counts, names="Status", values="Count", hole=0.55,
                      title="Pipeline Status", color="Status", color_discrete_map=STATUS_COLORS)
        st.plotly_chart(fig, use_container_width=True)
    with colB:
        country_counts = fdf["Country"].value_counts().reset_index()
        country_counts.columns = ["Country", "Leads"]
        fig2 = px.bar(country_counts, x="Country", y="Leads", title="Leads by Country",
                       color="Country", text="Leads")
        st.plotly_chart(fig2, use_container_width=True)

    colC, colD = st.columns(2)
    with colC:
        source_counts = fdf["Lead Source"].value_counts().reset_index()
        source_counts.columns = ["Source", "Leads"]
        fig3 = px.bar(source_counts, x="Leads", y="Source", orientation="h", title="Leads by Source", text="Leads")
        st.plotly_chart(fig3, use_container_width=True)
    with colD:
        heat = pd.crosstab(fdf["Country"], fdf["Status"]).reindex(columns=STATUSES, fill_value=0)
        fig4 = px.imshow(heat, text_auto=True, aspect="auto", title="Country x Status Heatmap",
                          color_continuous_scale="Blues")
        st.plotly_chart(fig4, use_container_width=True)

    st.subheader("Leads Table (editable)")
    edited = st.data_editor(
        fdf,
        column_config={
            "Status": st.column_config.SelectboxColumn(options=STATUSES),
            "Lead Source": st.column_config.SelectboxColumn(options=SOURCES),
            "Priority": st.column_config.SelectboxColumn(options=["High", "Medium", "Low"]),
        },
        num_rows="dynamic",
        use_container_width=True,
        key="leads_editor",
    )
    if st.button("💾 Save lead changes"):
        base = st.session_state.leads.copy()
        base.update(edited)
        st.session_state.leads = base
        st.cache_data.clear()
        st.success("Saved for this session. Use 'Export' below to write it back to Excel.")

    if st.button("⬇️ Export leads to Excel"):
        with pd.ExcelWriter(DATA_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            st.session_state.leads.to_excel(writer, sheet_name="Data", index=False)
        st.success(f"Saved to {DATA_FILE}")

# ============================================================
# TAB 2 — TEAM REFERRALS
# ============================================================
with tab2:
    rdf = st.session_state.referrals.copy()

    st.subheader("Referrals by Team Member")
    if len(rdf):
        member_counts = rdf["Referred To (Team Member)"].value_counts().reindex(TEAM, fill_value=0).reset_index()
    else:
        member_counts = pd.DataFrame({"Referred To (Team Member)": TEAM, "count": [0] * len(TEAM)})
    member_counts.columns = ["Team Member", "Referrals"]
    fig5 = px.bar(member_counts, x="Team Member", y="Referrals", color="Team Member", text="Referrals")
    st.plotly_chart(fig5, use_container_width=True)

    if len(rdf):
        st.subheader("Referral Status Breakdown")
        rc1, rc2, rc3, rc4, rc5 = st.columns(5)
        counts = rdf["Status"].value_counts()
        rc1.metric("New", int(counts.get("New", 0)))
        rc2.metric("Contacted", int(counts.get("Contacted", 0)))
        rc3.metric("In Progress", int(counts.get("In Progress", 0)))
        rc4.metric("Closed - Won", int(counts.get("Closed - Won", 0)))
        rc5.metric("Closed - Lost", int(counts.get("Closed - Lost", 0)))

    st.subheader("Referral Log (editable)")
    st.caption("Add a row every time you refer a client to a team member. They update Status + notes as they follow up.")
    edited_ref = st.data_editor(
        rdf if len(rdf) else pd.DataFrame(columns=["Ref #", "Date Referred", "Referred To (Team Member)", "Client Name",
                                                     "Client Phone Number", "Client Requirement / Details", "Status",
                                                     "Last Update Date", "Update / Notes from Member"]),
        column_config={
            "Referred To (Team Member)": st.column_config.SelectboxColumn(options=TEAM),
            "Status": st.column_config.SelectboxColumn(options=REF_STATUSES),
        },
        num_rows="dynamic",
        use_container_width=True,
        key="referrals_editor",
    )
    if st.button("💾 Save referral changes"):
        st.session_state.referrals = edited_ref
        st.success("Saved for this session.")

    if st.button("⬇️ Export referrals to Excel"):
        with pd.ExcelWriter(DATA_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            edited_ref.to_excel(writer, sheet_name="Team Referrals", index=False, startcol=1)
        st.success(f"Saved to {DATA_FILE}")
