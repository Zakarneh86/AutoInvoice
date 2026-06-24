from io import BytesIO

import pandas as pd
import streamlit as st

import database
import modules


EXCEL_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def show_error(message, exc):
    st.error(message)
    with st.expander("Error details"):
        st.exception(exc)


def safe_log_timesheet_extraction(
    secrets,
    source_file_name,
    po_number=None,
    engineer_name=None,
    status="success",
    error_message=None,
    extracted_json=None,
    metadata=None,
):
    try:
        return database.log_timesheet_extraction(
            secrets,
            source_file_name=source_file_name,
            po_number=po_number,
            engineer_name=engineer_name,
            status=status,
            error_message=error_message,
            extracted_json=extracted_json,
            metadata=metadata,
        )
    except Exception as exc:
        st.warning(f"Timesheet extraction was not logged: {exc}")
        return False


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


def process_timesheets(uploaded_files, openai_client, secrets=None, po_number=None):
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

            try:
                timesheet = modules.file_to_base64(uploaded_file)
                ts_json = modules.generate_ts_details(timesheet, openai_client)
                entries = modules.normalize_timesheet_records(ts_json)
            except Exception as exc:
                if secrets is not None:
                    safe_log_timesheet_extraction(
                        secrets,
                        source_file_name=uploaded_file.name,
                        po_number=po_number,
                        status="failed",
                        error_message=str(exc),
                        metadata={
                            "file_index": index,
                            "total_files": total_files,
                        },
                    )
                status.update(
                    label=f"Failed while processing {uploaded_file.name}",
                    state="error",
                )
                raise RuntimeError(
                    f"Timesheet '{uploaded_file.name}' could not be processed: {exc}"
                ) from exc

            pro_info = {
                "project_name": ts_json.get("client"),
                "order_number": ts_json.get("client_order_number"),
            }
            eng_name = ts_json.get("engineer_name") or uploaded_file.name
            log_po_number = po_number or ts_json.get("client_order_number")

            if secrets is not None:
                safe_log_timesheet_extraction(
                    secrets,
                    source_file_name=uploaded_file.name,
                    po_number=log_po_number,
                    engineer_name=eng_name,
                    status="success",
                    extracted_json=ts_json,
                    metadata={
                        "file_index": index,
                        "total_files": total_files,
                        "record_count": len(entries),
                    },
                )

            if "pro_info" not in ts_details:
                ts_details["pro_info"] = pro_info

            if eng_name not in ts_details:
                ts_details[eng_name] = {"ts1": entries}
            else:
                timesheet_count = len(ts_details[eng_name].keys()) + 1
                ts_details[eng_name][f"ts_{timesheet_count}"] = entries

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


def build_po_table_frames(po_json):
    try:
        po_info = po_json["po_info"]
        po_number = po_info["po_number"]

        po_master_df = pd.DataFrame([po_info])

        daily_rates_df = pd.DataFrame(po_json.get("daily_rate", []))
        if not daily_rates_df.empty:
            daily_rates_df.insert(0, "po_number", po_number)

        hourly_rates_df = pd.DataFrame(po_json.get("hourly_rate", []))
        if not hourly_rates_df.empty:
            hourly_rates_df.insert(0, "po_number", po_number)

        working_hours_df = pd.DataFrame(po_json.get("working_hours", []))
        if not working_hours_df.empty:
            working_hours_df.insert(0, "po_number", po_number)

        return po_master_df, daily_rates_df, hourly_rates_df, working_hours_df
    except KeyError as exc:
        raise KeyError(f"Extracted PO data is missing required field: {exc}") from exc
    except Exception as exc:
        raise RuntimeError(f"Could not prepare extracted PO tables: {exc}") from exc


st.markdown(
    """
    <div class="hero">
        <h1>Auto Invoice</h1>
        <p>Prepare invoice calculation sheets from PO selections and uploaded timesheets.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


try:
    app_secrets = st.secrets
except Exception as exc:
    app_secrets = {}
    st.warning("Secrets could not be loaded. OpenAI and Supabase features may be unavailable.")

try:
    api_keys = app_secrets["API_Keys"]
    openai_key = api_keys["openAI"]
    client, error, status_text = modules.client(openai_key)
except Exception as exc:
    client = None
    error = True
    status_text = f"OpenAI client is not configured: {exc}"

if error:
    st.error(status_text)
else:
    st.caption(status_text)

try:
    database_status = database.get_database_status(app_secrets)
    st.caption(
        f"Database Mode: {database_status['mode']} | Status: {database_status['status']}"
    )

    po_master, po_working_hours, po_daily_rates, po_hourly_rates = database.get_orders_data(
        app_secrets,
        database_status["use_supabase"],
    )
except Exception as exc:
    show_error("The order database could not be loaded.", exc)
    st.stop()

invoice_tab, add_po_tab = st.tabs(["Generate Invoice", "Add New PO"])

with invoice_tab:
    po_options = list(po_master["po_number"].unique())
    if not po_options:
        st.error("No purchase orders were found.")
        st.stop()

    left_panel, right_panel = st.columns([1.05, 0.95], gap="large")

    with left_panel:
        st.markdown('<div class="section-label">Invoice setup</div>', unsafe_allow_html=True)
        po = st.selectbox("PO", ["Select PO"] + po_options)

        invoicing_type_options = ["Select Invoicing Type", "Daily", "Hourly", "Mixed"]
        if po != "Select PO":
            po_invoicing_type = po_master["invoicing_type"][
                po_master["po_number"] == po
            ].iloc[0]
            if po_invoicing_type == "daily":
                default_invoicing_type = "Daily"
            elif po_invoicing_type == "mixed":
                default_invoicing_type = "Mixed"
            else:
                default_invoicing_type = "Hourly"
            default_invoicing_index = invoicing_type_options.index(default_invoicing_type)
        else:
            default_invoicing_index = 0

        invoicing_type_label = st.selectbox(
            "Invoicing Type",
            invoicing_type_options,
            index=default_invoicing_index,
        )

        invoicing_type = {
            "Daily": "daily",
            "Hourly": "hourly_rate",
            "Mixed": "mixed",
        }.get(invoicing_type_label)

        rate_table = po_daily_rates if invoicing_type == "daily" else po_hourly_rates

        if po != "Select PO":
            location_options = list(
                rate_table["onshore_or_offshore"][
                    rate_table["po_number"] == po
                ].unique()
            )
        else:
            location_options = []

        location = st.selectbox("Location", ["Select Location"] + location_options)

        if po != "Select PO" and location != "Select Location":
            role_options = list(
                rate_table["role_name"][
                    (rate_table["po_number"] == po)
                    & (rate_table["onshore_or_offshore"] == location)
                ].unique()
            )
        else:
            role_options = []

        role = st.selectbox("Role", ["Select Role"] + role_options)

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

    can_generate = bool(
        po != "Select PO"
        and invoicing_type_label != "Select Invoicing Type"
        and location != "Select Location"
        and role != "Select Role"
        and timesheets
        and not error
    )

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

        try:
            ts_details = process_timesheets(
                timesheets,
                client,
                secrets=app_secrets,
                po_number=po,
            )
            st.session_state["ts_details"] = ts_details

            with st.status("Generating calculation sheet...", expanded=True) as status:
                st.write("Applying PO rates and working hours.")
                excel_file = modules.fill_calculation_excel(
                    ts_details,
                    po_daily_rates,
                    po_hourly_rates,
                    po_working_hours,
                    role,
                    location,
                    po,
                    invoicing_type,
                )
                st.session_state["calculation_excel"] = excel_file.getvalue()
                status.update(label="Calculation sheet generated.", state="complete")

            st.session_state["show_ready_dialog"] = True
        except Exception as exc:
            show_error("Calculation sheet generation failed.", exc)

    if st.session_state.get("show_ready_dialog") and "calculation_excel" in st.session_state:
        show_ready_dialog()

    if st.session_state.get("show_excel_review") and "calculation_excel" in st.session_state:
        st.subheader("Calculation Sheet Review")
        try:
            preview = pd.read_excel(BytesIO(st.session_state["calculation_excel"]), header=None)
            st.dataframe(preview, use_container_width=True)
        except Exception as exc:
            show_error("The generated Excel file could not be previewed.", exc)

        st.download_button(
            "Download Calculation Sheet",
            data=st.session_state["calculation_excel"],
            file_name="calculation.xlsx",
            mime=EXCEL_MIME_TYPE,
            use_container_width=True,
        )

with add_po_tab:
    st.markdown('<div class="section-label">New PO setup</div>', unsafe_allow_html=True)

    po_has_price_list = st.checkbox("PO file includes the price list", value=True)
    po_file = st.file_uploader("Upload PO PDF", type=["pdf"], key="new_po_file")
    price_list_file = None

    if not po_has_price_list:
        price_list_file = st.file_uploader(
            "Upload price list PDF",
            type=["pdf"],
            key="new_po_price_list_file",
        )

    can_process_po = bool(
        po_file
        and not error
        and (po_has_price_list or price_list_file)
    )

    if st.button(
        "Process PO",
        disabled=not can_process_po,
        type="primary",
        use_container_width=True,
    ):
        try:
            with st.status("Extracting PO data...", expanded=True) as status:
                if po_has_price_list:
                    st.write(f"Processing {po_file.name}")
                    po_json = modules.get_po_data(po_file, None, None, True, client)
                else:
                    st.write(f"Processing {po_file.name}")
                    st.write(f"Processing {price_list_file.name}")
                    po_json = modules.get_po_data(
                        None,
                        po_file,
                        price_list_file,
                        False,
                        client,
                    )

                (
                    st.session_state["new_po_master"],
                    st.session_state["new_po_daily_rates"],
                    st.session_state["new_po_hourly_rates"],
                    st.session_state["new_po_working_hours"],
                ) = build_po_table_frames(po_json)
                status.update(label="PO data extracted.", state="complete")
        except Exception as exc:
            show_error("PO extraction failed.", exc)

    if "new_po_master" in st.session_state:
        st.subheader("Review and edit extracted PO data")

        st.markdown("PO Master")
        edited_po_master = st.data_editor(
            st.session_state["new_po_master"],
            num_rows="fixed",
            use_container_width=True,
            key="po_master_editor",
        )

        st.markdown("PO Daily Rates")
        edited_daily_rates = st.data_editor(
            st.session_state["new_po_daily_rates"],
            num_rows="dynamic",
            use_container_width=True,
            key="po_daily_rates_editor",
        )

        st.markdown("PO Hourly Rates")
        edited_hourly_rates = st.data_editor(
            st.session_state["new_po_hourly_rates"],
            num_rows="dynamic",
            use_container_width=True,
            key="po_hourly_rates_editor",
        )

        st.markdown("PO Working Hours")
        edited_working_hours = st.data_editor(
            st.session_state["new_po_working_hours"],
            num_rows="dynamic",
            use_container_width=True,
            key="po_working_hours_editor",
        )

        if st.button("Submit PO to database", type="primary", use_container_width=True):
            try:
                database.save_po_tables(
                    edited_po_master,
                    edited_daily_rates,
                    edited_hourly_rates,
                    edited_working_hours,
                    app_secrets,
                    database_status["use_supabase"],
                )
                st.success("PO database updated.")
            except Exception as exc:
                show_error("PO database update failed.", exc)
