import streamlit as st


PO_OPTIONS = [
    "Select PO",
    "PO 001",
    "PO 002",
    "PO 003",
]

ROLE_OPTIONS = [
    "Select role",
    "Project Manager",
    "Business Analyst",
    "Developer",
    "QA Engineer",
]

LOCATION_OPTIONS = [
    "Select location",
    "Onsite",
    "Offshore",
    "Remote",
]


st.set_page_config(
    page_title="Auto Invoice",
    page_icon=":page_facing_up:",
    layout="centered",
)

st.title("Auto Invoice")

po = st.selectbox("PO Droplist", PO_OPTIONS)
role = st.selectbox("Role Droplist", ROLE_OPTIONS)
location = st.selectbox("Location Droplist", LOCATION_OPTIONS)

timesheets = st.file_uploader(
    "Timesheets upload area",
    type=["pdf", "png"],
    accept_multiple_files=True,
)

if timesheets:
    st.subheader("Uploaded timesheets")
    for timesheet in timesheets:
        st.write(timesheet.name)