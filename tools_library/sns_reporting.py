# tools_library/sns_reporting.py
import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime

# ==========================================================
# 🏷️ MODULE METADATA & ENTRY REGISTRATION FOR STREAMLIT SaaS
# ==========================================================
TOOL_NAME = "SNS Reporting Tool"
TOOL_ICON = "📈"

def clean_headers(df):
    """Strips hidden spaces from input spreadsheet headers securely."""
    df.columns = [str(c).strip() for c in df.columns]
    return df

def run_sns_process(sales_file, stock_file, email_master_file, send_email, mail_subject, mail_body, start_date, end_date):
    """Executes clean, isolated separate sheet mappings using exact matching column sets."""
    
    # 1. Loading Input Data Streams Safely from Memory Buffers
    df_sales = pd.read_excel(sales_file) if sales_file.name.endswith('.xlsx') else pd.read_csv(sales_file)
    df_stock = pd.read_excel(stock_file) if stock_file.name.endswith('.xlsx') else pd.read_csv(stock_file)
    email_master = pd.read_excel(email_master_file) if email_master_file.name.endswith('.xlsx') else pd.read_csv(email_master_file)
    
    df_sales = clean_headers(df_sales)
    df_stock = clean_headers(df_stock)
    email_master = clean_headers(email_master)

    # Standardize 'Article Name' validation join key across all datasets
    for target_df in [df_sales, df_stock, email_master]:
        if 'Article Name' in target_df.columns:
            target_df['Article Name'] = target_df['Article Name'].astype(str).str.strip().str.upper()

    # Deduplicate Email Master by 'Article Name' to prevent cross-join array mutations
    if 'Article Name' in email_master.columns:
        email_master = email_master.drop_duplicates(subset=['Article Name'], keep='first')
        
    # Isolate Email Master columns exclusively to block cross-file column name collisions
    if 'Email' in email_master.columns and 'Article Name' in email_master.columns:
        email_master = email_master[['Article Name', 'Email']]
    else:
        raise ValueError("The uploaded Email Master sheet must contain exact 'Article Name' and 'Email' headers.")

    # 2. Strict Date Filtering Layer on Sales Sheet via "Invoice Created On"
    date_col = 'Invoice Created On'
    if date_col in df_sales.columns:
        df_sales[date_col] = pd.to_datetime(df_sales[date_col], errors='coerce')
        df_sales = df_sales[
            (df_sales[date_col].dt.date >= start_date) & 
            (df_sales[date_col].dt.date <= end_date)
        ].copy()

    # 3. Structural Renaming & Data Alignment
    columns_rename = {
        'Physical inventory': 'Physical Stock',
        'Location': 'Location Name'
    }
    df_sales = df_sales.rename(columns=columns_rename)
    df_stock = df_stock.rename(columns=columns_rename)

    # 🚨 EXACT REFINED COLUMN PICKING MATRIX - FITS YOUR EXCEL MATCH
    sales_specific_cols = [
        'Site', 'Site Name', 'Sales Office', 'Article', 'Item Description', 
        'Article Name', 'Family Name', 'Sub-Family Name', 'Category Name', 
        'Brand Name', 'RRP Price', 'Invoice Created On', 'Invoice Quantity', 'Amount@RRP'
    ]

    stock_specific_cols = [
        'Article', 'Article Name', 'Item Description', 'Site', 'Site Name', 
        'Location Name', 'Family Name', 'Sub-Family Name', 'Brand Name', 
        'Category Name', 'Physical Stock', 'Consignment Stock'
    ]

    # Safe Verification Layer: Initialize missing keys with blank markers to bypass index failures
    for col in sales_specific_cols:
        if col not in df_sales.columns:
            df_sales[col] = np.nan
            
    for col in stock_specific_cols:
        if col not in df_stock.columns:
            df_stock[col] = np.nan

    # 4. Independent Merging (Prevents 'Reindexing Objects' errors completely)
    df_sales_mapped = df_sales[sales_specific_cols].merge(email_master, how='inner', on='Article Name')
    df_stock_mapped = df_stock[stock_specific_cols].merge(email_master, how='inner', on='Article Name')

    # Intersect uniquely matched supplier emails found across the spreadsheets
    all_emails = set(df_sales_mapped['Email'].dropna().unique()).union(set(df_stock_mapped['Email'].dropna().unique()))

    processed_count = 0
    zip_buffer_dict = {}

    # 5. Segment and build workbooks individually for each supplier email group
    for email in all_emails:
        email_str = str(email).strip()
        if email_str.lower() in ['unmapped@company.com', 'na', 'nan', '', 'na;na']:
            continue

        sales_final = df_sales_mapped[df_sales_mapped['Email'] == email_str].copy()
        stock_final = df_stock_mapped[df_stock_mapped['Email'] == email_str].copy()

        # Format timestamps cleanly into standard text-based short dates for Excel display
        if not sales_final.empty and 'Invoice Created On' in sales_final.columns:
            sales_final['Invoice Created On'] = pd.to_datetime(sales_final['Invoice Created On']).dt.date

        # Drop the technical relational 'Email' column before writing out to partners
        sales_final = sales_final.drop(columns=['Email'], errors='ignore').dropna(how='all')
        stock_final = stock_final.drop(columns=['Email'], errors='ignore').dropna(how='all')

        # Skip compilation loop if both calculations return completely empty fragments
        if sales_final.empty and stock_final.empty:
            continue

        # Compile separate clean sheets inside a unified workbook directly in memory
        excel_out = io.BytesIO()
        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
            sales_final.to_excel(writer, sheet_name='Sales Report', index=False)
            stock_final.to_excel(writer, sheet_name='Stock Report', index=False)
            
        excel_bytes = excel_out.getvalue()
        
        # Remove parentheses or quotes from email keys to make clean file names
        clean_filename = f"SNS_Report_{email_str}.xlsx".replace("'", "").replace("(", "").replace(")", "")
        zip_buffer_dict[clean_filename] = excel_bytes
        processed_count += 1

    return processed_count, zip_buffer_dict

def render_ui():
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.subheader("Generate split supplier sheets mapped via Article Name keys.")
    st.markdown("---")
    
    # Structural File Intake Dashboard Panels
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sales_file = st.file_uploader("Upload Raw Sales Transactions Sheet", type=["xlsx", "csv"], key="sns_sales")
        stock_file = st.file_uploader("Upload Active Stock Balancing Sheet", type=["xlsx", "csv"], key="sns_stock")
    with col_u2:
        email_master_file = st.file_uploader("Upload Brand Email Master Mapping Sheet", type=["xlsx", "csv"], key="sns_email")
        
    st.markdown("### 📅 Filter Constraints (Invoice Created On)")
    date_col1, date_col2 = st.columns(2)
    with date_col1:
        start_date = st.date_input("Reporting From Date:", value=datetime.date(2026, 5, 1), key="sns_start")
    with date_col2:
        end_date = st.date_input("Reporting To Date:", value=datetime.date(2026, 5, 31), key="sns_end")

    st.markdown("### ✉️ Email Broadcast Template Configuration")
    send_email = st.checkbox("Enable Automated Server Email Distribution Loop?", value=False, key="sns_send")

    if send_email:
        default_body_str = (
            "Dear valued partner,\n\n"
            "Please find attached your Sales and Stock Performance Report covering the requested window.\n\n"
            "Thanks,\n"
            "Category Management Team"
        )
        mail_subject = st.text_input("Email Global Subject Line:", value="Sales & Stocks Performance Report Update", key="sns_subject_input")
        mail_body = st.text_area("Email Global Body Text Content:", value=default_body_str, height=180, key="sns_body_input")
    else:
        mail_subject = ""
        mail_body = ""

    st.markdown("---")

    if st.button("⚡ Execute Reporting Automation Loops", type="primary", use_container_width=True, key="sns_run"):
        if sales_file and stock_file and email_master_file:
            with st.spinner("Processing deep matrix calculations..."):
                try:
                    count, generated_files = run_sns_process(
                        sales_file, stock_file, email_master_file, 
                        send_email, mail_subject, mail_body, start_date, end_date
                    )
                    
                    if count > 0:
                        st.success(f"🎉 Process Complete! Successfully generated clean data splits for {count} unique suppliers.")
                        
                        if generated_files:
                            import zipfile
                            zip_out = io.BytesIO()
                            with zipfile.ZipFile(zip_out, 'w') as zip_f:
                                for fname, fbytes in generated_files.items():
                                    zip_f.writestr(fname, fbytes)
                            
                            st.download_button(
                                label="📥 Download All Generated Partner Workbooks (ZIP Archive)",
                                data=zip_out.getvalue(),
                                file_name=f"SNS_Partner_Reports_{datetime.datetime.now().strftime('%Y%m%d')}.zip",
                                mime="application/zip",
                                use_container_width=True
                            )
                    else:
                        st.warning("⚠️ No matching rows found. Ensure that your From and To dates cover the values inside your Sales spreadsheet data rows.")
                except Exception as ex:
                    st.error(f"🚨 Automation Processing Error: {str(ex)}")
        else:
            st.warning("Please upload Sales, Stock, and Email Master files simultaneously to generate workbooks.")

if __name__ == "__main__":
    render_ui()
