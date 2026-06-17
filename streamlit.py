from io import BytesIO

import pandas as pd
import streamlit as st

import modules


EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


st.set_page_config(
    page_title="Auto Invoice",
    page_icon=":page_facing_up:",
    layout="wide",
)


st.markdown(
    """
    <style>
        .block-container {
            max-width: 1180px;
            padding-top: 2rem;
        }

        .hero {
            background: linear-gradient(135deg, #0f766e 0%, #155e75 55%, #334155 100%);
            border-radius: 8px;
            color: white;
            padding: 28px 32px;
            margin-bottom: 22px;
        }

        .hero h1 {
            margin: 0;
            font-size: 2.4rem;
            font-weight: 750;
        }

        .hero p {
            margin: 8px 0 0 0;
            color: #dbeafe;
            font-size: 1rem;
        }

        .panel {
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 18px;
            background: #ffffff;
        }

        .section-label {
            color: #475569;
            font-size: 0.82rem;
            font-weight: 700;
            letter-spacing: 0;
            margin-bottom: 8px;
            text-transform: uppercase;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def process_timesheets(uploaded_files, openai_client):
    ts_details = {}
    total_files = len(uploaded_files)
    progress_bar = st.progress(0)

    with st.status("Preparing timesheet extraction...", expanded=True) as status:
        for index, uploaded_file in enumerate(uploaded_files, start=1):
            st.write(f"Processing {uploaded_file.name}")
            status.update(
                label=f"Processing timesheet {index} of {total_files}: {uploaded_file.name}",
                state="running",
            )

            timesheet = modules.file_to_base64(uploaded_file)
            ts_json = modules.generate_ts_details(timesheet, openai_client)

            pro_info = {
                "project_name": ts_json["client"],
                "order_number": ts_json["client_order_number"],
            }
            eng_name = ts_json["engineer_name"]

            if "pro_info" not in ts_details:
                ts_details["pro_info"] = pro_info

            if eng_name not in ts_details:
                ts_details[eng_name] = {"ts1": ts_json["entries"]}
            else:
                timesheet_count = len(ts_details[eng_name].keys()) + 1
                ts_details[eng_name][f"ts_{timesheet_count}"] = ts_json["entries"]

            progress_bar.progress(index / total_files)

        status.update(label="Timesheets extracted successfully.", state="complete")

    return ts_details


@st.dialog("Calculation sheet is ready")
def show_ready_dialog():
    st.write("The Excel calculation sheet has been generated.")

    left, right = st.columns(2)
    with left:
        if st.button("Review", use_container_width=True):
            st.session_state["show_excel_review"] = True
            st.session_state["show_ready_dialog"] = False
            st.rerun()

    with right:
        st.download_button(
            "Download",
            data=st.session_state["calculation_excel"],
            file_name="calculation.xlsx",
            mime=EXCEL_MIME_TYPE,
            use_container_width=True,
        )


st.markdown(
    """
    <div class="hero">
        <h1>Auto Invoice</h1>
        <p>Prepare invoice calculation sheets from PO selections and uploaded timesheets.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


api_keys = st.secrets["API_Keys"]
openai_key = api_keys["openAI"]
client, error, status_text = modules.client(openai_key)

if error:
    st.error(status_text)
else:
    st.caption(status_text)

po_master, po_working_hours, po_daily_rates, po_hourly_rates = modules.get_orders_data()

po_options = list(po_master["po_number"].unique())
if not po_options:
    st.error("No purchase orders were found.")
    st.stop()

left_panel, right_panel = st.columns([1.05, 0.95], gap="large")

with left_panel:
    st.markdown('<div class="section-label">Invoice setup</div>', unsafe_allow_html=True)
    po = st.selectbox("PO", po_options)

    location_options = list(
        po_hourly_rates["onshore_or_offshore"][po_hourly_rates["po_number"] == po].unique()
    )
    location = st.selectbox("Location", location_options)

    role_options = list(
        po_hourly_rates["role_name"][
            (po_hourly_rates["po_number"] == po)
            & (po_hourly_rates["onshore_or_offshore"] == location)
        ].unique()
    )
    role = st.selectbox("Role", role_options)

with right_panel:
    st.markdown('<div class="section-label">Timesheets</div>', unsafe_allow_html=True)
    timesheets = st.file_uploader(
        "Upload timesheets",
        type=["pdf", "png"],
        accept_multiple_files=True,
    )

    if timesheets:
        st.caption(f"{len(timesheets)} file(s) ready")
        for timesheet in timesheets:
            st.write(timesheet.name)
    else:
        st.info("Upload one or more timesheets to enable generation.")


st.divider()

can_generate = bool(po and role and location and timesheets and not error)

generate_clicked = st.button(
    "Generate Calculation Sheet",
    disabled=not can_generate,
    type="primary",
    use_container_width=True,
)

if not can_generate:
    st.caption("Select a PO, role, location, and upload at least one timesheet.")

if generate_clicked:
    st.session_state["show_excel_review"] = False
    st.session_state["show_ready_dialog"] = False

    ts_details = process_timesheets(timesheets, client)
    st.session_state["ts_details"] = ts_details

    with st.status("Generating calculation sheet...", expanded=True) as status:
        st.write("Applying PO rates and working hours.")
        excel_file = modules.fill_calculation_excel(
            ts_details,
            po_hourly_rates,
            po_working_hours,
            role,
            location,
            po,
        )
        st.session_state["calculation_excel"] = excel_file.getvalue()
        status.update(label="Calculation sheet generated.", state="complete")

    st.session_state["show_ready_dialog"] = True

if st.session_state.get("show_ready_dialog") and "calculation_excel" in st.session_state:
    show_ready_dialog()

if st.session_state.get("show_excel_review") and "calculation_excel" in st.session_state:
    st.subheader("Calculation Sheet Review")
    preview = pd.read_excel(BytesIO(st.session_state["calculation_excel"]), header=None)
    st.dataframe(preview, use_container_width=True)

    st.download_button(
        "Download Calculation Sheet",
        data=st.session_state["calculation_excel"],
        file_name="calculation.xlsx",
        mime=EXCEL_MIME_TYPE,
        use_container_width=True,
    )
