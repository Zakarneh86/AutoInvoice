from io import BytesIO
import pandas as pd
import streamlit as st
import modules

apiKeys = st.secrets["API_Keys"]
openAiKey = apiKeys["openAI"]

client, error, status_text = modules.client(openAiKey)
st.write(status_text)


st.title("Auto Invoice")
st.set_page_config(
    page_title="Auto Invoice",
    page_icon=":page_facing_up:",
    layout="centered",
)

EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

po_master, po_working_hours, po_daily_rates, po_hourly_rates = modules.get_orders_data()

PO_OPTIONS = list(po_master['po_number'].unique())
po = st.selectbox("PO Droplist", PO_OPTIONS)

LOCATION_OPTIONS = list(po_hourly_rates['onshore_or_offshore'][po_hourly_rates['po_number'] == po].unique())
location = st.selectbox("Location Droplist", LOCATION_OPTIONS)

ROLE_OPTIONS = list(po_hourly_rates['role_name'][(po_hourly_rates['po_number'] == po) & (po_hourly_rates['onshore_or_offshore'] == location)])
role = st.selectbox("Role Droplist", ROLE_OPTIONS)

timesheets = st.file_uploader(
    "Timesheets upload area",
    type=["pdf", "png"],
    accept_multiple_files=True,
)

if timesheets:
    st.subheader("Uploaded timesheets")
    for timesheet in timesheets:
        st.write(timesheet.name)


ts_details = modules.get_timesheets_data(timesheets, client)

st.divider()
st.subheader("Calculation Excel")

st.session_state['ts_details']=ts_details
st.session_state['po_hourly_rates'] = po_hourly_rates
st.session_state['po_working_hours'] = po_working_hours



required_excel_inputs = {
    "ts_details": st.session_state.get("ts_details"),
    "po_hourly_rates": st.session_state.get("po_hourly_rates"),
    "po_working_hours": st.session_state.get("po_working_hours"),
}

if all(value is not None for value in required_excel_inputs.values()):
    if st.button("Generate calculation Excel"):
        excel_file = modules.fill_calculation_excel(
            required_excel_inputs["ts_details"],
            required_excel_inputs["po_hourly_rates"],
            required_excel_inputs["po_working_hours"],
            role,
            location,
            po
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
