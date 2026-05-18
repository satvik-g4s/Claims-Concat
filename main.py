import streamlit as st
import pandas as pd
import io
import time

st.set_page_config(layout="wide")#g

st.title("Multi File Excel Concatenation Tool")

st.markdown("### Upload Excel Files")

uploaded_files = st.file_uploader(
    "Upload Excel Files",
    type=["xlsx", "xls"],
    accept_multiple_files=True
)

st.caption(
    """
    Upload claims files for all locations here.

    Expected Structure:
    - Each uploaded Excel file must contain the claims sheet whose name matches
      the first word of the uploaded file name (case-insensitive).

    Example:
    File Name: Bangalore Collection.xlsx
    Required Sheet Name: Bangalore

    Required Columns: EFT/Cheque No to be Claimed, EFT Amount, EFT Date, Invoice Number, Invoice Amount, Outstanding Amount (A), Client Code, Client name, Invoice amount to be adjust, TDS (B), Deduction Amount (C), Deduction reason, TDS %, EFT Amount for Invoice (D)
    """
)

run_button = st.button("Run")

log_container = st.container()

# Final required output columns
# Final output columns
final_output_columns = [
    "EFT/Cheque No to be Claimed",
    "EFT Amount",
    "EFT Date",
    "Invoice Number",
    "Invoice Amount",
    "Outstanding Amount (A)",
    "Client Code",
    "Client name",
    "Invoice amount to be adjust",
    "TDS (B)",
    "Deduction Amount (C)",
    "Deduction reason",
    "TDS %",
    "EFT Amount for Invoice (D)",
    "Branch Remarks",
    "Branch",
    "Payment",
    "BankName",
    "BankCode",
    "BankRef",
    "EFTNO"
]
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
            
                    normalized_sheet_name = (
                        str(sheet_name)
                        .strip()
                        .lower()
                        .split()[0]
                    )
            
                    if normalized_sheet_name == file_name_first_word_raw:
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

                    # Create output dataframe
                    df_temp = pd.DataFrame()
                    
                    for output_col in final_output_columns:
                    
                        matched_col = None
                    
                        for actual_col in temp_df.columns:
                    
                            if (
                                str(actual_col).strip().lower()
                                == output_col.strip().lower()
                            ):
                                matched_col = actual_col
                                break
                    
                        if matched_col is not None:
                            df_temp[output_col] = temp_df[matched_col]
                        else:
                            df_temp[output_col] = ""
                    header_found = True
                    # Merge Unclaimed Payment sheet
                    try:
                    
                        unclaimed_df = pd.read_excel(
                            io.BytesIO(file_content),
                            sheet_name="Unclaimed Payment",
                            dtype=str
                        )
                    
                        unclaimed_df.columns = (
                            unclaimed_df.columns
                            .astype(str)
                            .str.strip()
                        )
                    
                        required_unclaimed_cols = [
                            "BankName",
                            "BankCode",
                            "EFTNO",
                            "Value",
                            "BankRef"
                        ]
                    
                        available_cols = [
                            col for col in required_unclaimed_cols
                            if col in unclaimed_df.columns
                        ]
                    
                        unclaimed_df = unclaimed_df[available_cols]
                    
                        # Standardize merge keys
                        df_temp["EFT/Cheque No to be Claimed"] = (
                            df_temp["EFT/Cheque No to be Claimed"]
                            .astype(str)
                            .str.strip()
                            .str.replace(".0", "", regex=False)
                            .str.replace(" ", "", regex=False)
                        )
                        
                        unclaimed_df["EFTNO"] = (
                            unclaimed_df["EFTNO"]
                            .astype(str)
                            .str.strip()
                            .str.replace(".0", "", regex=False)
                            .str.replace(" ", "", regex=False)
                        )
                        
                        # Remove duplicate EFTNO rows
                        unclaimed_df = unclaimed_df.drop_duplicates(
                            subset=["EFTNO"]
                        )


                        
                        # Merge
                        df_temp = df_temp.merge(
                            unclaimed_df,
                            how="left",
                            left_on="EFT/Cheque No to be Claimed",
                            right_on="EFTNO",
                            suffixes=("", "_unclaimed")
                        )
                    
                    except Exception as e:
                        st.error(f"Error processing Unclaimed Payment sheet: {e}")
                        st.stop()
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

        # Remove rows where:
        # 1. Entire row is null/blank
        # AND
        # 2. DeductionAmount is 0/blank

        try:
            final_df = pd.concat(df_list, ignore_index=True)

        except Exception as e:
            st.error(f"Error concatenating dataframes: {e}")
            st.stop()

        time.sleep(0.5)
        
        try:

            amount_cols = [
                "Deduction Amount (C)",
                "EFT Amount for Invoice (D)",
                "TDS (B)"
            ]
        
            # Standardize amount columns
            for col in amount_cols:
        
                final_df[col] = (
                    final_df[col]
                    .astype(str)
                    .str.replace(",", "", regex=False)
                    .str.strip()
                )
        
            # Check all NON-amount columns are blank
            other_cols = [
                col for col in final_df.columns
                if col not in amount_cols
            ]
        
            other_cols_blank = (
                final_df[other_cols]
                .fillna("")
                .astype(str)
                .apply(lambda col: col.str.strip())
                .replace(["nan", "None"], "")
                .eq("")
                .all(axis=1)
            )
        
            # Check all amount columns are blank/0
            amount_cols_zero_blank = (
                final_df[amount_cols]
                .replace(["", " ", "nan", "None", "<NA>"], "0")
                .apply(pd.to_numeric, errors="coerce")
                .fillna(0)
                .eq(0)
                .all(axis=1)
            )
        
            # Remove rows only where BOTH conditions are true
            final_df = final_df[
                ~(other_cols_blank & amount_cols_zero_blank)
            ]
        
        except Exception as e:
            st.error(f"Error filtering blank rows: {e}")
            st.stop()        
            
        # Remove fully blank rows
            
        try:
        
            final_df = final_df.replace(
                ["", " ", "nan", "None"],
                pd.NA
            )
        
            final_df = final_df.dropna(how="all")
        
        except Exception as e:
            st.error(f"Error removing fully blank rows: {e}")
            st.stop()
                

        # Concatenate dataframes
        

        status_text.info("Generating output Excel file...")
'''
        try:
    
            final_df = final_df.rename(columns={
        
                "EFT Date": "DepositDate",
                "EFT/Cheque No to be Claimed": "CheckNo",
                "Client Code": "AccountCustomer",
                "EFT Amount": "DepositAmount",
                "Branch Remarks": "PaymentMethod",
                "Invoice Number": "InvoiceNo",
                "Branch": "SOLocn",
                "Invoice Amount": "Inv Amt",
                "Outstanding Amount (A)": "OutstandingAmount",
                "TDS (B)": "TDSAmount",
                "Deduction Amount (C)": "DeductionAmount",
                "Payment": "CheckAmount",
                "Deduction reason": "ReasonCode",
                "TDS %": "TDS",
        
                "BankRef_unclaimed": "DepositSlipNo",
                "BankName_unclaimed": "BankLocation",
                "BankCode_unclaimed": "BankNo"
        
            })

            
        
            # Duplicate columns
            if "AccountCustomer" in final_df.columns:
                final_df["CustomerCode"] = final_df["AccountCustomer"]
        
            if "DepositSlipNo" in final_df.columns:
                final_df["BankReference"] = final_df["DepositSlipNo"]
        
            # Final required order
            final_order = [
                "Company Code",
                "DepositDate",
                "BankReference",
                "CheckNo",
                "AccountCustomer",
                "DepositAmount",
                "CompanyLoc",
                "InstrumentPrefix",
                "PaymentMethod",
                "DepositSlipNo",
                "BankLocation",
                "BankNo",
                "CheckDate",
                "CustomerCode",
                "InvoiceNo",
                "ARLocn",
                "SOLocn",
                "Inv Amt",
                "OutstandingAmount",
                "TDSAmount",
                "DeductionAmount",
                "TDS CGST",
                "TDS SGST",
                "TDS IGST",
                "Retention Amount",
                "CheckAmount",
                "ReasonCode",
                "TDS",
                "Invoice_type"
            ]
        
            # Create missing columns
            for col in final_order:
        
                if col not in final_df.columns:
                    final_df[col] = ""
        
            # Reorder columns
            final_df = final_df[final_order]
        
        except Exception as e:
            st.error(f"Error during final processing: {e}")
            st.stop()
'''
        # Generate CSV output
        try:
            output = io.StringIO()
        
            final_df.to_csv(
                output,
                index=False
            )
        
            csv_data = output.getvalue()
        
        except Exception as e:
            st.error(f"Error generating CSV file: {e}")
            st.stop()

        status_text.success(
            f"Processing completed successfully. "
            f"Total combined rows: {len(final_df)}"
        )

        st.download_button(
            label="Download Combined CSV File",
            data=csv_data,
            file_name="combined_output.csv",
            mime="text/csv"
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
