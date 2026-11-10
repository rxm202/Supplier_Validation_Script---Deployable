from azure_blob_helper import upload_blob
from datetime import datetime
import streamlit as st
import pandas as pd
import subprocess
import os
import plotly.express as px
import time
import sys

# ======================================================
# SESSION STATE
# ======================================================

if "record_view" not in st.session_state:

    st.session_state["record_view"] = None


# ==========================================================
# CREATE WORKING FOLDERS
# ==========================================================

import os

for folder in [
    "input",
    "output",
    "temp"
]:
    os.makedirs(
        folder,
        exist_ok=True
    )

# ======================================================
# PAGE CONFIG
# ======================================================

st.set_page_config(
    page_title="Supplier Validation Framework",
    page_icon="🏢",
    layout="wide"
)

# ======================================================
# HELPER FUNCTIONS
# ======================================================

def load_summary():

    summary_file = "output/Validation_Summary.csv"

    if not os.path.exists(summary_file):
        return None

    summary_df = pd.read_csv(summary_file)

    metrics = {}

    for _, row in summary_df.iterrows():
        metrics[row["Metric"]] = row["Value"]

    return metrics


# ======================================================
# BANNER
# ======================================================

def show_banner():

    st.markdown(
        """
        <div style="
        background:linear-gradient(90deg,#0F3D77,#0057B8);
        height:10px;
        border-radius:8px;
        margin-bottom:20px;
        ">
        </div>
        """,
        unsafe_allow_html=True
    )

    col1, col2 = st.columns([1, 6])

    with col1:

        st.image(
            "images/sherwin_logo.png",
            width=80
        )

    with col2:

        st.markdown(
            """
            <h1 style="
            color:#0F3D77;
            margin-bottom:0px;
            ">
            Supplier Validation Framework
            </h1>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <p style="
            font-size:18px;
            font-weight:600;
            color:#444444;
            margin-top:0px;
            margin-bottom:5px;
            ">
            Internal On-Prem Supplier Data Validation Platform
            </p>
            """,
            unsafe_allow_html=True
        )

        st.markdown(
            """
            <p style="
            font-size:15px;
            color:#555555;
            ">
            Designed to support supplier onboarding validation,
            data quality checks, validation reporting, and Oracle
            supplier import readiness assessment within the
            Sherwin-Williams enterprise environment.
            </p>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)
# ======================================================
# KPI SECTION
# ======================================================

def show_kpis(metrics):

    col1, col2, col3, col4 = st.columns(4)

    with col1:

        st.metric(
            "📄 Total Records",
            metrics.get(
                "Total Records",
                0
            )
        )

    with col2:

        st.metric(
            "✅ Valid Records",
            metrics.get(
                "Valid Records",
                0
            )
        )

        if st.button(
            "View",
            key="view_valid"
        ):
            st.session_state[
                "record_view"
            ] = "VALID"

    with col3:

        st.metric(
            "❌ Rejected Records",
            metrics.get(
                "Rejected Records",
                0
            )
        )

        if st.button(
            "View",
            key="view_error"
        ):
            st.session_state[
                "record_view"
            ] = "ERROR"

    with col4:

        st.metric(
            "⚠ Warning Records",
            metrics.get(
                "Warning Records",
                0
            )
        )

        if st.button(
            "View",
            key="view_warning"
        ):
            st.session_state[
                "record_view"
            ] = "WARNING"
# ======================================================
# CHARTS
# ======================================================

def show_charts(metrics):

    chart_df = pd.DataFrame(
        {
            "Status": [
                "Valid",
                "Rejected",
                "Warnings"
            ],
            "Count": [
                metrics.get(
                    "Valid Records",
                    0
                ),
                metrics.get(
                    "Rejected Records",
                    0
                ),
                metrics.get(
                    "Warning Records",
                    0
                )
            ]
        }
    )

    col1, col2 = st.columns(2)

    with col1:

        fig1 = px.pie(
            chart_df,
            names="Status",
            values="Count",
            hole=0.45,
            title="Validation Distribution",
            color="Status",
            color_discrete_map={
                "Valid": "#2E7D32",
                "Warnings": "#FFC107",
                "Rejected": "#C62828"
            }
        )

        st.plotly_chart(
            fig1,
            use_container_width=True
        )

    with col2:

        fig2 = px.bar(
            chart_df,
            x="Status",
            y="Count",
            color="Status",
            title="Records by Status",
            color_discrete_map={
                "Valid": "#2E7D32",
                "Warnings": "#FFC107",
                "Rejected": "#C62828"
            }
        )

        st.plotly_chart(
            fig2,
            use_container_width=True
        )


# ======================================================
# SIDEBAR
# ======================================================

col_logo, col_title = st.sidebar.columns([1, 2])

with col_logo:

    st.image(
        "images/sherwin_logo.png",
        width=50
    )

with col_title:

    st.markdown(
        """
        **Supplier<br>
        Validation**
        """,
        unsafe_allow_html=True
    )

st.sidebar.markdown(
    """
    **Internal On-Prem**  
    Data Quality Platform

    ---

    Supplier onboarding validation,
    data quality assessment,
    validation reporting,
    and Oracle WebADI readiness.
    """
)

menu = st.sidebar.radio(
    "Navigation",
    [
        "🏠 Dashboard",
        "📂 Upload & Run",
        "⚙ Rule Management",
        "📊 Reports & Downloads",
        "ℹ About"
    ]
)


# ======================================================
# DASHBOARD
# ======================================================

if menu == "🏠 Dashboard":

    show_banner()

    st.write("")

    metrics = load_summary()

    if metrics:

        st.subheader(
            "Validation KPI Dashboard"
        )

        show_kpis(metrics)

        st.write("")

        show_charts(metrics)

        # ======================================================
        # KPI DRILL DOWN
        # ======================================================

        if st.session_state["record_view"] == "VALID":

            st.divider()

            st.subheader(
                "✅ Valid Records"
            )

            valid_file = (
                "output/Supplier_Output.xlsx"
            )

            if os.path.exists(valid_file):

                valid_df = pd.read_excel(
                    valid_file,
                    dtype=str
                )

                st.dataframe(
                    valid_df,
                    use_container_width=True,
                    height=400
                )

        elif st.session_state["record_view"] == "ERROR":

            st.divider()

            st.subheader(
                "❌ Rejected Records"
            )

            validation_file = (
                "output/Supplier_Validation_Report.xlsx"
            )

            if os.path.exists(validation_file):

                validation_df = pd.read_excel(
                    validation_file,
                    dtype=str,
                    header=2
                )

                error_df = validation_df[
                    validation_df[
                        "FAILURE_REASON"
                    ]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    != ""
                ]

                st.dataframe(
                    error_df,
                    use_container_width=True,
                    height=400
                )

        elif st.session_state["record_view"] == "WARNING":

            st.divider()

            st.subheader(
                "⚠ Warning Records"
            )

            validation_file = (
                "output/Supplier_Validation_Report.xlsx"
            )

            if os.path.exists(validation_file):

                validation_df = pd.read_excel(
                    validation_file,
                    dtype=str,
                    header=2
                )

                warning_df = validation_df[
                    validation_df[
                        "WARNING_MESSAGE"
                    ]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                    != ""
                ]

                st.dataframe(
                    warning_df,
                    use_container_width=True,
                    height=400
                )

    else:

        st.info(
            "No validation results available. Run a validation first."
        )
# ======================================================
# UPLOAD & RUN
# ======================================================

elif menu == "📂 Upload & Run":

    show_banner()

    st.write("")

    col1, col2 = st.columns([2, 1])

    # ==========================================
    # UPLOAD SECTION
    # ==========================================

    with col1:

        st.subheader("Upload Supplier File")

        uploaded_file = st.file_uploader(
            "Upload Supplier Excel File",
            type=["xlsx"]
        )

    # ==========================================
    # TEMPLATE SECTION
    # ==========================================

    with col2:

        st.subheader("Input Template")

        st.info(
            """
            Download the input template
            before preparing supplier data.

            • Do not modify column names

            • Populate supplier records from Row 2

            • Save the file in .xlsx format

            • Upload the completed file for validation
            """
        )

        template_file = (
            "template/Supplier_Input_Template.xlsx"
        )

        if os.path.exists(template_file):

            with open(
                template_file,
                "rb"
            ) as file:

                st.download_button(
                    label="📥 Download Input Template",
                    data=file.read(),
                    file_name="Supplier_Input_Template.xlsx",
                    use_container_width=True
                )

    # ==========================================
    # FILE UPLOAD PROCESSING
    # ==========================================

    if uploaded_file is not None:

        save_path = os.path.join(
            "input",
            "current_input.xlsx"
        )

        with open(save_path, "wb") as f:

            f.write(
                uploaded_file.getbuffer()
            )

        # ==========================================
        # AZURE UPLOAD TEMPORARILY DISABLED
        # ==========================================

        timestamp = datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        azure_blob_path = (
            f"data_integration/"
            f"supplier_R12/"
            f"input/"
            f"{timestamp}_{uploaded_file.name}"
        )

        st.info(
            "☁ Azure Blob upload temporarily disabled."
        )

        st.success(
            f"✅ File uploaded: {uploaded_file.name}"
        )

        if st.button("🚀 Run Validation"):

            progress_bar = st.progress(0)

            with st.spinner(
                "Validation in progress..."
            ):

                for i in range(100):

                    progress_bar.progress(
                        i + 1
                    )

                    time.sleep(0.01)

                # ==========================================
                # RUN VALIDATION ENGINE
                # ==========================================

                result = subprocess.run(
                    [
                        sys.executable,
                        "Master_Validator.py"
                    ],
                    capture_output=True,
                    text=True
                )

                # ==========================================
                # RUN WEBADI GENERATOR
                # ==========================================

                webadi_result = subprocess.run(
                    [
                        sys.executable,
                        "WebADI_Generator.py"
                    ],
                    capture_output=True,
                    text=True
                )

            progress_bar.empty()

            # ==========================================
            # VALIDATION STATUS
            # ==========================================

            if result.returncode == 0:

                st.success(
                    "✅ Validation completed successfully."
                )

                st.info(
                    "☁ Azure Blob output upload temporarily disabled."
                )

                # ==========================================
                # WEBADI STATUS
                # ==========================================

                if webadi_result.returncode == 0:

                    st.success(
                        "🚀 Oracle WebADI file generated successfully."
                    )

                else:

                    st.warning(
                        "Validation completed but WebADI generation failed."
                    )

                    st.code(
                        webadi_result.stderr
                    )

            else:

                st.error(
                    "❌ Validation Failed"
                )

                st.code(
                    result.stderr
                )

            # ==========================================
            # SUMMARY
            # ==========================================

            metrics = load_summary()

            if metrics:

                st.divider()

                st.subheader(
                    "Validation Summary"
                )

                show_kpis(metrics)
# ======================================================
# RULE MANAGEMENT
# ======================================================

elif menu == "⚙ Rule Management":

    st.markdown(
        """
        <div style="
            background: linear-gradient(90deg,#0F3D77,#0057B8);
            padding:20px;
            border-radius:12px;
            color:white;">
            <h1>⚙ Rule Management</h1>
            <p>
            Configure validation rules, severity and activation.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("")

    rules_file = "config/Validation_Rules.xlsx"

    rules_df = pd.read_excel(rules_file)

    # Only show Mandatory Rules for now
    mandatory_rules = rules_df[
        rules_df["CATEGORY"]
        .astype(str)
        .str.upper()
        == "MANDATORY"
    ]

    st.subheader("Mandatory Validation Rules")

    edited_rules = st.data_editor(
        mandatory_rules,
        use_container_width=True,
        hide_index=True
    )

    col1, col2 = st.columns(2)

    with col1:

        if st.button("💾 Save Changes"):

            updated_rules = rules_df.copy()

            for _, edited_row in edited_rules.iterrows():

                rule_id = edited_row["RULE_ID"]

                updated_rules.loc[
                    updated_rules["RULE_ID"] == rule_id,
                    "SEVERITY"
                ] = edited_row["SEVERITY"]

                updated_rules.loc[
                    updated_rules["RULE_ID"] == rule_id,
                    "ACTIVE"
                ] = edited_row["ACTIVE"]

            updated_rules.to_excel(
                rules_file,
                index=False
            )

            st.success(
                "Rule configuration saved successfully."
            )

    with col2:

        with open(
            rules_file,
            "rb"
        ) as file:

            st.download_button(
                label="⬇ Export Rule Configuration",
                data=file.read(),
                file_name="Validation_Rules.xlsx"
            )

# ======================================================
# REPORTS
# ======================================================

elif menu == "📊 Reports & Downloads":

    show_banner()

    st.write("")

    st.subheader("Download Center")

    st.markdown(
        """
        Download validation reports, summary files,
        and Oracle WebADI output generated during
        supplier validation processing.
        """
    )

    files = [
        (
            "output/Supplier_Output.xlsx",
            "✅ Supplier Output"
        ),
        (
            "output/Supplier_Validation_Report.xlsx",
            "❌ Validation Error"
        ),
        (
            "output/Validation_Summary.csv",
            "📊 Validation Summary"
        ),
        (
            "output/Supplier_WebADI_Output.xlsx",
            "🚀 Oracle WebADI File"
        )
    ]

    download_names = {

        "✅ Supplier Output":
            "Supplier_Output",

        "❌ Validation Error":
            "Supplier_Validation_Report",

        "📊 Validation Summary":
            "Validation_Summary",

        "🚀 Oracle WebADI File":
            "Supplier_WebADI_Output"
    }

    col1, col2, col3 = st.columns(3)

    columns = [col1, col2, col3]

    for idx, (file_path, label) in enumerate(files):

        if os.path.exists(file_path):

            with columns[idx % 3]:

                download_timestamp = datetime.now().strftime(
                    "%Y%m%d_%H%M%S"
                )

                file_extension = os.path.splitext(
                    file_path
                )[1]

                download_file_name = (
                    f"{download_names[label]}_"
                    f"{download_timestamp}"
                    f"{file_extension}"
                )

                with open(
                    file_path,
                    "rb"
                ) as file:

                    st.download_button(
                        label=label,
                        data=file.read(),
                        file_name=download_file_name,
                        use_container_width=True
                    )

    st.divider()

        # ==========================================
    # ERROR PREVIEW
    # ==========================================

    error_file = (
        "output/Supplier_Validation_Report.xlsx"
    )

    if os.path.exists(error_file):

        st.subheader("📋 Validation Report Preview")

        error_df = pd.read_excel(
            error_file,
            dtype=str
        )

        preview_df = error_df.drop(
            columns=[
                "ERROR_COLUMNS",
                "WARNING_COLUMNS"
            ],
            errors="ignore"
        )

        st.dataframe(
            preview_df.head(20),
            use_container_width=True,
            height=350
        )

# ======================================================
# ABOUT
# ======================================================

elif menu == "ℹ About":

    show_banner()

    st.markdown(
        """
### Supplier Validation Framework

Implemented Features

✅ Mandatory Validation

✅ Duplicate Validation

✅ Country Specific Validation

✅ Banking Validation

✅ Payment Term LOV Validation

✅ Payment Method LOV Validation

✅ Supplier Type LOV Validation

✅ Argentina Localization Validation

✅ Brazil Localization Validation

✅ Warning Management

✅ Error Management

✅ Summary Reporting

✅ Download Center

---
Supplier Data Validation Framework
"""
    )