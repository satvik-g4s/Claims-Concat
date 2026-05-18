import streamlit as st
import pandas as pd
import io
import time

st.set_page_config(layout="wide")

st.title("Excel File Concatenation Tool")

st.markdown("### Upload Files")

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

st.caption(
    "Required columns: "
    "Company Code, DepositDate, BankReference, CheckNo, AccountCustomer, "
    "DepositAmount, CompanyLoc, InstrumentPrefix, PaymentMethod, DepositSlipNo, "
    "BankLocation, BankNo, CheckDate, CustomerCode, InvoiceNo, CheckAmount, "
    "TDSAmount, DeductionAmount, OutstandingAmount, ARLocn, SOLocn, "
    "ReasonCode, TDS, Invoice_type"
)

run_button = st.button("Run")

log_container = st.container()

# Define the list of columns to read
columns_to_read = [
    'Company Code',
    'DepositDate',
    'BankReference',
    'CheckNo',
    'AccountCustomer',
    'DepositAmount',
    'CompanyLoc',
    'InstrumentPrefix',
    'PaymentMethod',
    'DepositSlipNo',
    'BankLocation',
    'BankNo',
    'CheckDate',
    'CustomerCode',
    'InvoiceNo',
    'CheckAmount',
    'TDSAmount',
    'DeductionAmount',
    'OutstandingAmount',
    'ARLocn',
    'SOLocn',
    'ReasonCode',
    'TDS',
    'Invoice_type'
]

# Create dtype mapping
dtype_mapping = {col: str for col in columns_to_read}

if run_button:

    with log_container:

        status_text = st.empty()

        if not uploaded_files:
            st.error("Please upload at least one Excel file.")
            st.stop()

        status_text.info("Starting file processing...")
        time.sleep(0.5)

        df_list = []

        progress_bar = st.progress(0)

        total_files = len(uploaded_files)

        for i, file_obj in enumerate(uploaded_files):

            status_text.info(f"Reading file {i + 1} of {total_files}...")

            try:
                # Reset file pointer
                file_obj.seek(0)

                # Read Excel file
                df_temp = pd.read_excel(
                    file_obj,
                    dtype=dtype_mapping
                )

            except Exception as e:
                st.error(f"Error reading file {file_obj.name}: {e}")
                st.stop()

            # Header finder logic (checks first 10 rows)
            header_found = False

            for row_num in range(10):

                try:
                    file_obj.seek(0)

                    temp_df = pd.read_excel(
                        file_obj,
                        header=row_num,
                        dtype=dtype_mapping
                    )

                    if all(col in temp_df.columns for col in columns_to_read):
                        df_temp = temp_df[columns_to_read]
                        header_found = True
                        break

                except Exception:
                    continue

            if not header_found:
                st.error(
                    f"Required headers not found within first 10 rows in file: {file_obj.name}"
                )
                st.stop()

            # Column validation
            missing_cols = [
                col for col in columns_to_read
                if col not in df_temp.columns
            ]

            if missing_cols:
                st.error(
                    f"Missing columns in file {file_obj.name}: {missing_cols}"
                )
                st.stop()

            try:
                df_list.append(df_temp)

            except Exception as e:
                st.error(f"Error appending dataframe: {e}")
                st.stop()

            progress = int(((i + 1) / total_files) * 100)
            progress_bar.progress(progress)

            time.sleep(0.3)

        status_text.info("Concatenating all uploaded files...")

        try:
            final_df = pd.concat(df_list, ignore_index=True)

        except Exception as e:
            st.error(f"Error concatenating files: {e}")
            st.stop()

        time.sleep(0.5)

        status_text.info("Preparing download file...")

        try:
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Combined_Data')

            output.seek(0)

        except Exception as e:
            st.error(f"Error generating output file: {e}")
            st.stop()

        status_text.success(
            f"Processing completed successfully. Total rows combined: {len(final_df)}"
        )

        st.download_button(
            label="Download Combined Excel File",
            data=output,
            file_name="combined_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("### Preview of Combined Data")
        st.dataframe(final_df.head())

        st.markdown("### Data Types")
        st.dataframe(
            pd.DataFrame({
                "Column": final_df.columns,
                "Data Type": final_df.dtypes.astype(str)
            })
        )

with st.expander("What This Tool Does"):

    st.write("""
    This tool allows users to upload multiple Excel files containing deposit,
    invoice, customer, and payment-related information.

    All uploaded files are validated, cleaned, and concatenated into one
    unified dataset for reporting and downstream processing.
    """)

with st.expander("How to Use"):

    st.write("""
    1. Upload one or more Excel files.
    2. Click the Run button.
    3. Wait for processing to complete.
    4. Download the combined output Excel file.
    """)

with st.expander("Output Details"):

    st.write("""
    The generated output contains:

    - Combined records from all uploaded files
    - Standardized column structure
    - Unified customer and invoice-level data
    - Deposit and payment tracking information

    All uploaded files are merged row-wise into a single consolidated report.
    """)

with st.expander("Financial Logic"):

    st.write("""
    The tool combines customer payment and invoice records from multiple files
    into one master dataset.

    It preserves:
    - Deposit references
    - Invoice mappings
    - TDS deductions
    - Outstanding amounts
    - Customer codes
    - Payment methods
    - Bank-related details

    This helps in consolidated reconciliation and financial tracking.
    """)
