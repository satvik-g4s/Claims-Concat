import streamlit as st
import pandas as pd
import io
import time

st.set_page_config(layout="wide")

st.title("Multi File Excel Concatenation Tool")

st.markdown("### Upload Excel Files")

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

st.caption(
    """
    Expected Structure:
    - Each uploaded Excel file must contain a sheet whose name matches the first word
      of the uploaded file name (case-insensitive).
    - Example:
        File Name: ABC_Invoice.xlsx
        Required Sheet Name: ABC

    Required Columns:
    Company Code, DepositDate, BankReference, CheckNo, AccountCustomer,
    DepositAmount, CompanyLoc, InstrumentPrefix, PaymentMethod,
    DepositSlipNo, BankLocation, BankNo, CheckDate, CustomerCode,
    InvoiceNo, CheckAmount, TDSAmount, DeductionAmount,
    OutstandingAmount, ARLocn, SOLocn, ReasonCode, TDS, Invoice_type
    """
)

run_button = st.button("Run")

log_container = st.container()

# Define required columns
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

# Define dtype mapping
dtype_mapping = {col: str for col in columns_to_read}

if run_button:

    with log_container:

        status_text = st.empty()

        if not uploaded_files:
            st.error("Please upload at least one Excel file.")
            st.stop()

        status_text.info("Starting processing...")
        time.sleep(0.5)

        df_list = []

        progress_bar = st.progress(0)

        total_files = len(uploaded_files)

        for i, file_obj in enumerate(uploaded_files):

            status_text.info(
                f"Processing file {i + 1} of {total_files}: {file_obj.name}"
            )

            # Extract first word from filename
            try:
                file_name_first_word_raw = (
                    file_obj.name
                    .replace(".xlsx", "")
                    .replace(".xls", "")
                    .replace("_", " ")
                    .strip()
                    .split()[0]
                    .lower()
                )

            except Exception as e:
                st.error(f"Error extracting filename from uploaded file: {e}")
                st.stop()

            # Read file content
            try:
                file_obj.seek(0)
                file_content = file_obj.read()

            except Exception as e:
                st.error(f"Error reading uploaded file content: {e}")
                st.stop()

            # Open Excel file
            try:
                excel_file = pd.ExcelFile(io.BytesIO(file_content))

            except Exception as e:
                st.error(f"Error opening Excel file {file_obj.name}: {e}")
                st.stop()

            # Get sheet names
            try:
                actual_sheet_names = excel_file.sheet_names

            except Exception as e:
                st.error(f"Error fetching sheet names from {file_obj.name}: {e}")
                st.stop()

            # Match sheet name
            target_sheet_name = None

            try:
                for sheet_name in actual_sheet_names:

                    if (
                        str(sheet_name)
                        .strip()
                        .lower()
                        == file_name_first_word_raw
                    ):
                        target_sheet_name = sheet_name
                        break

            except Exception as e:
                st.error(f"Error during sheet matching process: {e}")
                st.stop()

            # Validate matched sheet
            if target_sheet_name is None:
                st.error(
                    f"No matching sheet found for '{file_name_first_word_raw}' "
                    f"in file {file_obj.name}"
                )
                st.stop()

            status_text.info(
                f"Matched sheet '{target_sheet_name}' in file {file_obj.name}"
            )

            time.sleep(0.3)

            # Header detection within first 10 rows
            header_found = False

            normalized_required_cols = [
                str(col).strip().lower().replace("\n", " ")
                for col in columns_to_read
            ]

            for row_num in range(10):

                try:
                    temp_df = pd.read_excel(
                        io.BytesIO(file_content),
                        sheet_name=target_sheet_name,
                        header=row_num,
                        dtype=dtype_mapping
                    )

                    # Normalize dataframe columns
                    normalized_temp_cols = [
                        str(col).strip().lower().replace("\n", " ")
                        for col in temp_df.columns
                    ]

                    # Mapping normalized -> original
                    column_mapping = {
                        str(col).strip().lower().replace("\n", " "): col
                        for col in temp_df.columns
                    }

                    # Check all required columns exist
                    if all(
                        col in normalized_temp_cols
                        for col in normalized_required_cols
                    ):

                        selected_original_cols = [
                            column_mapping[col]
                            for col in normalized_required_cols
                        ]

                        df_temp = temp_df[selected_original_cols]

                        # Rename columns back to standard names
                        df_temp.columns = columns_to_read

                        header_found = True
                        break

                except Exception:
                    continue

            # Validate header found
            if not header_found:
                st.error(
                    f"Required headers not found within first 10 rows "
                    f"in file {file_obj.name}"
                )
                st.stop()

            # Final column validation
            missing_cols = [
                col for col in columns_to_read
                if col not in df_temp.columns
            ]

            if missing_cols:
                st.error(
                    f"Missing columns in file {file_obj.name}: {missing_cols}"
                )
                st.stop()

            # Append dataframe
            try:
                df_list.append(df_temp)

            except Exception as e:
                st.error(f"Error appending dataframe: {e}")
                st.stop()

            progress = int(((i + 1) / total_files) * 100)
            progress_bar.progress(progress)

            time.sleep(0.3)

        # Validate dataframe list
        if not df_list:
            st.error("No valid dataframes were processed.")
            st.stop()

        status_text.info("Concatenating all processed files...")

        # Concatenate dataframes
        try:
            final_df = pd.concat(df_list, ignore_index=True)

        except Exception as e:
            st.error(f"Error concatenating dataframes: {e}")
            st.stop()

        time.sleep(0.5)

        status_text.info("Generating output Excel file...")

        # Generate output file
        try:
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final_df.to_excel(
                    writer,
                    index=False,
                    sheet_name='Combined_Data'
                )

            output.seek(0)

        except Exception as e:
            st.error(f"Error generating output file: {e}")
            st.stop()

        status_text.success(
            f"Processing completed successfully. "
            f"Total combined rows: {len(final_df)}"
        )

        st.download_button(
            label="Download Combined Excel File",
            data=output,
            file_name="combined_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.markdown("### Combined Data Preview")
        st.dataframe(final_df.head())

        st.markdown("### Data Type Information")

        try:
            dtype_df = pd.DataFrame({
                "Column": final_df.columns,
                "Data Type": final_df.dtypes.astype(str)
            })

            st.dataframe(dtype_df)

        except Exception as e:
            st.error(f"Error displaying datatype information: {e}")
            st.stop()

with st.expander("What This Tool Does"):

    st.write("""
    This tool uploads multiple Excel files, validates sheet names,
    reads the required columns, and concatenates all records into
    a single consolidated output file.

    The tool automatically detects the matching sheet based on
    the uploaded file name.
    """)

with st.expander("How to Use"):

    st.write("""
    1. Upload one or more Excel files.
    2. Ensure each file contains the required sheet.
    3. Click Run.
    4. Download the final combined output file.
    """)

with st.expander("Output Details"):

    st.write("""
    The output contains:
    - Combined rows from all uploaded files
    - Standardized column structure
    - Customer transaction details
    - Invoice references
    - Payment and deduction information
    - Consolidated financial dataset

    All files are merged vertically into one master report.
    """)

with st.expander("Financial Logic"):

    st.write("""
    The process consolidates financial transaction records while preserving:

    - Deposit references
    - Customer mappings
    - Invoice numbers
    - TDS and deduction values
    - Outstanding balances
    - Bank information
    - Payment methods

    This supports reconciliation, reporting, and downstream finance operations.
    """)
