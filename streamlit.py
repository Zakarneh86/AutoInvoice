from io import BytesIO
import pandas as pd
import streamlit as st
from modules import fill_calculation_excel


EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


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

st.divider()
st.subheader("Calculation Excel")

required_excel_inputs = {
    "ts_details": st.session_state.get("ts_details"),
    "po_hourly_rates": st.session_state.get("po_hourly_rates"),
    "po_working_hours": st.session_state.get("po_working_hours"),
}

if all(value is not None for value in required_excel_inputs.values()):
    if st.button("Generate calculation Excel"):
        excel_file = fill_calculation_excel(
            required_excel_inputs["ts_details"],
            required_excel_inputs["po_hourly_rates"],
            required_excel_inputs["po_working_hours"],
            role,
            location,
        )
        st.session_state["calculation_excel"] = excel_file.getvalue()

if "calculation_excel" in st.session_state:
    st.download_button(
        "Download calculation Excel",
        data=st.session_state["calculation_excel"],
        file_name="calculation.xlsx",
        mime=EXCEL_MIME_TYPE,
    )

    preview = pd.read_excel(BytesIO(st.session_state["calculation_excel"]), header=None)
    st.dataframe(preview, use_container_width=True)
else:
    st.info("Process the PO and timesheets first, then generate the calculation Excel.")
