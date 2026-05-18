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
    - Each uploaded Excel file must contain a sheet whose name matches
      the first word of the uploaded file name (case-insensitive).

    Example:
    File Name: Bangalore Collection.xlsx
    Required Sheet Name: Bangalore

    The tool automatically maps source columns into a standardized output structure.
    Missing columns will be created as blank.
    """
)

run_button = st.button("Run")

log_container = st.container()

# Final required output columns
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

# Source Column -> Output Column Mapping
column_mapping_config = {
    "EFT Date": "DepositDate",
    "EFT/Cheque No to be Claimed": "CheckNo",
    "Client Code": "AccountCustomer",
    "EFT Amount": "DepositAmount",
    "Branch Remarks": "PaymentMethod",
    "Client Code": "CustomerCode",
    "Invoice Number": "InvoiceNo",
    "Deduction Amount (C)": "DeductionAmount",
    "Outstanding Amount (A)": "OutstandingAmount",
    "Branch": "SOLocn",
    "Deduction reason": "ReasonCode",
    "TDS %": "TDS"
}

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

            for row_num in range(10):

                try:
                    temp_df = pd.read_excel(
                        io.BytesIO(file_content),
                        sheet_name=target_sheet_name,
                        header=row_num,
                        dtype=str
                    )

                    # Create standardized dataframe
                    df_temp = pd.DataFrame()

                    # Loop through final required columns
                    for output_col in columns_to_read:

                        source_col = None

                        # Direct column match
                        for actual_col in temp_df.columns:

                            if (
                                str(actual_col).strip().lower()
                                == output_col.strip().lower()
                            ):
                                source_col = actual_col
                                break

                        # Mapping-based match
                        if source_col is None:

                            for source_name, mapped_name in column_mapping_config.items():

                                if mapped_name == output_col:

                                    for actual_col in temp_df.columns:

                                        if (
                                            str(actual_col).strip().lower()
                                            == source_name.strip().lower()
                                        ):
                                            source_col = actual_col
                                            break

                        # Populate column
                        if source_col is not None:
                            df_temp[output_col] = temp_df[source_col]
                        else:
                            df_temp[output_col] = ""

                    header_found = True
                    break

                except Exception:
                    continue

            # Validate processing
            if not header_found:
                st.error(
                    f"Could not process headers in file {file_obj.name}"
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
    This tool uploads multiple Excel files, automatically maps source columns,
    standardizes the structure, and combines all records into a single output file.

    Missing columns are automatically created as blank to ensure consistent output.
    """)

with st.expander("How to Use"):

    st.write("""
    1. Upload one or more Excel files.
    2. Ensure each file contains the required sheet.
    3. Click Run.
    4. Download the combined output file.
    """)

with st.expander("Output Details"):

    st.write("""
    The output contains:
    - Standardized financial transaction records
    - Customer and invoice details
    - Payment information
    - Deduction and outstanding amounts
    - Consolidated master dataset

    All uploaded files are vertically merged into one report.
    """)

with st.expander("Financial Logic"):

    st.write("""
    The tool standardizes and consolidates financial transaction data while preserving:

    - Deposit details
    - Customer mappings
    - Invoice references
    - Outstanding balances
    - Deduction information
    - TDS values
    - Branch and payment information

    Missing fields are automatically generated as blank values.
    """)
