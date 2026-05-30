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
    """Executes column harmonization, date filtering, and multi-sheet reporting generation."""
    
    # 1. Loading Input Data Streams Safely from Memory Buffers
    df_sales = pd.read_excel(sales_file) if sales_file.name.endswith('.xlsx') else pd.read_csv(sales_file)
    df_stock = pd.read_excel(stock_file) if stock_file.name.endswith('.xlsx') else pd.read_csv(stock_file)
    email_master = pd.read_excel(email_master_file) if email_master_file.name.endswith('.xlsx') else pd.read_csv(email_master_file)
    
    df_sales = clean_headers(df_sales)
    df_stock = clean_headers(df_stock)
    email_master = clean_headers(email_master)

    # 🚨 STRUCTURAL CORRECTION: Force absolute string-case standardization to avoid key misses
    for target_df in [df_sales, df_stock, email_master]:
        if 'Article Name' in target_df.columns:
            target_df['Article Name'] = target_df['Article Name'].astype(str).str.strip().str.upper()

    # Deduplicate Email Master by 'Article Name' to prevent cross-join index expansions
    if 'Article Name' in email_master.columns:
        email_master = email_master.drop_duplicates(subset=['Article Name'], keep='first')
        
    # Isolate Email Master to avoid structural header name overlaps
    if 'Email' in email_master.columns and 'Article Name' in email_master.columns:
        email_master = email_master[['Article Name', 'Email']]

    # 2. Strict Date Filtering Layer on Sales Sheet via "Invoice Created On"
    date_col = 'Invoice Created On'
    if date_col in df_sales.columns:
        df_sales[date_col] = pd.to_datetime(df_sales[date_col], errors='coerce')
        df_sales = df_sales[
            (df_sales[date_col].dt.date >= start_date) & 
            (df_sales[date_col].dt.date <= end_date)
        ]
    
    # 3. Structural Renaming & Data Alignment
    columns_rename = {
        'Site': 'Site', 'Location': 'Location Name', 'Location Name': 'Location Name',
        'Physical Stock': 'Physical Stock', 'Physical inventory': 'Physical Stock',
        'Moving Average Price': 'Moving Average Price', 'Invoice Qty': 'Invoice Quantity',
        'Invoice Quantity': 'Invoice Quantity', 'Invoice Net Value': 'Amount@RRP',
        'Sales Value': 'Amount@RRP', 'Amount@RRP': 'Amount@RRP'
    }

    df_stock = df_stock.rename(columns=columns_rename)
    df_sales = df_sales.rename(columns=columns_rename)

    # Add logical tracking tags to segment files during the groupby split loop
    df_sales['Report Type'] = 'Sales'
    df_stock['Report Type'] = 'Stock'

    # Combine datasets safely and completely ignore/reset old index tracking positions
    combined_df = pd.concat([df_sales, df_stock], ignore_index=True)
    
    # Merge using the isolated 'Article Name' master map
    combined_df = combined_df.merge(email_master, how='left', on=['Article Name'])
    
    # Fallback string assignments for unmapped assets
    if 'Email' not in combined_df.columns:
        combined_df['Email'] = 'unmapped@company.com'
    
    combined_df['Email'] = combined_df['Email'].fillna('unmapped@company.com').astype(str).str.strip()

    # Exact column matching from your uploaded target layout sheets
    sales_specific_cols = [
        'Invoice Created On', 'Sales Office', 'Article', 'Item Description', 
        'Article Name', 'Family', 'Sub-Family', 'Category', 'Brand', 
        'Invoice Quantity', 'Amount@RRP'
    ]

    stock_specific_cols = [
        'Article', 'Article Name', 'Item Description', 'Site', 'Location Name', 
        'Family Name', 'Sub-Family Name', 'Brand Name', 'Physical Stock', 'Moving Average Price'
    ]

    grouped = combined_df.groupby('Email')
    processed_count = 0
    zip_buffer_dict = {}

    for email_key, group_item in grouped:
        # Clean string references for standard exclusion safety filters
        clean_key = str(email_key).lower().strip()
        if 'unmapped' in clean_key or clean_key in ['na', 'nan', '', 'na;na']:
            continue
            
        sales_sheet = group_item[group_item['Report Type'] == 'Sales'].copy()
        stock_sheet = group_item[group_item['Report Type'] == 'Stock'].copy()

        # Format timestamps cleanly to text-based short dates for clean Excel display
        if not sales_sheet.empty and 'Invoice Created On' in sales_sheet.columns:
            sales_sheet['Invoice Created On'] = pd.to_datetime(sales_sheet['Invoice Created On']).dt.date

        # Clear duplicate indices on the subgroup targets
        sales_sheet.reset_index(drop=True, inplace=True)
        stock_sheet.reset_index(drop=True, inplace=True)

        # Build clean column filtering sequences with completely unique objects
        sales_cols_clean = list(dict.fromkeys([c for c in sales_specific_cols if c in sales_sheet.columns]))
        stock_cols_clean = list(dict.fromkeys([c for c in stock_specific_cols if c in stock_sheet.columns]))

        # Safely slice using the unique column keys list
        sales_sheet = sales_sheet[sales_cols_clean].dropna(how='all')
        stock_sheet = stock_sheet[stock_cols_clean].dropna(how='all')

        # Skip compilation loop if both data fragments contain no items
        if sales_sheet.empty and stock_sheet.empty:
            continue

        # Compile separate sheets inside a single workbook directly in memory
        excel_out = io.BytesIO()
        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
            sales_sheet.to_excel(writer, sheet_name='Sales Report', index=False)
            stock_sheet.to_excel(writer, sheet_name='Stock Report', index=False)
            
        excel_bytes = excel_out.getvalue()
        
        # Format names file values safely
        clean_file_name = f"SNS_Report_{email_key}.xlsx".replace("'", "").replace("(", "").replace(")", "")
        zip_buffer_dict[clean_file_name] = excel_bytes
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
                    else:
                        st.warning("⚠️ No matching records found. Verify that your Reporting From/To date selection covers the values present inside your Sales spreadsheet file.")
                    
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
                except Exception as ex:
                    st.error(f"🚨 Automation Processing Error: {str(ex)}")
        else:
            st.warning("Please upload Sales, Stock, and Email Master files simultaneously to generate workbooks.")

if __name__ == "__main__":
    render_ui()
