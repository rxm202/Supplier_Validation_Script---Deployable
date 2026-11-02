import streamlit as st
import pandas as pd
import subprocess
import os
import plotly.express as px
import time

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


def show_banner():

    st.markdown(
        """
        <div style="
        background:linear-gradient(90deg,#0F3D77,#0057B8);
        padding:20px;
        border-radius:12px;
        color:white;">
            <h1>🏢 Supplier Validation Framework</h1>
            <p>
            Data Validation Framework.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )


def show_kpis(metrics):

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "📄 Total Records",
            metrics.get("Total Records", 0)
        )

    with col2:
        st.metric(
            "✅ Valid Records",
            metrics.get("Valid Records", 0)
        )

    with col3:
        st.metric(
            "❌ Rejected Records",
            metrics.get("Rejected Records", 0)
        )

    with col4:
        st.metric(
            "⚠ Warning Records",
            metrics.get("Warning Records", 0)
        )


def show_charts(metrics):

    chart_df = pd.DataFrame(
        {
            "Status": [
                "Valid",
                "Rejected",
                "Warnings"
            ],
            "Count": [
                metrics.get("Valid Records", 0),
                metrics.get("Rejected Records", 0),
                metrics.get("Warning Records", 0)
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
            title="Validation Distribution"
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
            title="Records by Status"
        )

        st.plotly_chart(
            fig2,
            use_container_width=True
        )

# ======================================================
# SIDEBAR
# ======================================================

st.sidebar.markdown(
    """
    # 🏢 Supplier Validation
    
    Data Quality Platform
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

        st.subheader("Validation KPI Dashboard")

        show_kpis(metrics)

        st.write("")

        show_charts(metrics)

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

    st.subheader("Upload Supplier File")

    uploaded_file = st.file_uploader(
        "Upload Supplier Excel File",
        type=["xlsx"]
    )

    if uploaded_file is not None:

        save_path = os.path.join(
            "input",
            "current_input.xlsx"
        )

        with open(save_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(
            f"File uploaded: {uploaded_file.name}"
        )

        if st.button("🚀 Run Validation"):

            progress_bar = st.progress(0)

            with st.spinner(
                "Validation in progress..."
            ):

                for i in range(100):

                    progress_bar.progress(i + 1)

                    time.sleep(0.01)

                # Run Validation Engine

                result = subprocess.run(
                    ["python", "Master_Validator.py"],
                    capture_output=True,
                    text=True
                )

                # Run Oracle WebADI Generator

                webadi_result = subprocess.run(
                    ["python", "WebADI_Generator.py"],
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

    files = [
        (
            "output/Supplier_Output.xlsx",
            "✅ Supplier Output"
        ),
        (
            "output/Supplier_Validation_Error.csv",
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

    for file_path, label in files:

        if os.path.exists(file_path):

            with open(
                file_path,
                "rb"
            ) as file:

                st.download_button(
                    label=label,
                    data=file.read(),
                    file_name=os.path.basename(file_path)
                )

    st.divider()

    # ERROR PREVIEW

    error_file = (
        "output/Supplier_Validation_Error.csv"
    )

    if os.path.exists(error_file):

        st.subheader("❌ Error Preview")

        error_df = pd.read_csv(error_file)

        st.dataframe(
            error_df.head(20),
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