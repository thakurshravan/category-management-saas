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

def run_sns_process(sales_file, stock_file, email_master_file, send_email, mail_subject, mail_body):
    """Executes column harmonization, data splitting, and reporting distributions."""
    
    # 1. Loading Input Data Streams from Buffers Safely
    df_sales = pd.read_excel(sales_file) if sales_file.name.endswith('.xlsx') else pd.read_csv(sales_file)
    df_stock = pd.read_excel(stock_file) if stock_file.name.endswith('.xlsx') else pd.read_csv(stock_file)
    email_master = pd.read_excel(email_master_file) if email_master_file.name.endswith('.xlsx') else pd.read_csv(email_master_file)
    
    df_sales = clean_headers(df_sales)
    df_stock = clean_headers(df_stock)
    email_master = clean_headers(email_master) # Fixed typo variable pointer here

    # 2. Standardized MCH Schema Mappings
    columns_Output = [
        'Category Code', 'Brand', 'Brand Description', 'Item Code', 
        'Item Name', 'Product Description', 'Warehouse', 'Location', 
        'Sales Quantity', 'Sales Value', 'Total Inventory', 'Sellable Inventory',
        'Inventory Reserved (either against order / transit)', 'Order Date'
    ]

    columns_rename = {
        'Item number': 'Item Code', 'ItemId': 'Item Code',
        'Item Name': 'Item Name', 'ItemName': 'Item Name',
        'Description': 'Product Description', 'ProductDesc': 'Product Description',
        'Retail Brand': 'Brand', 'BrandName': 'Brand',
        'Retail Level 1 code': 'Category Code', 'Item Dim Cat Code (L1)': 'Category Code',
        'Warehouse': 'Warehouse', 'Header Business Unit': 'Warehouse',
        'Location': 'Location', 
        'Physical inventory': 'Total Inventory', 'Physical Stock': 'Total Inventory',
        'Available physical': 'Sellable Inventory', 'Sellable Stock': 'Sellable Inventory',
        'Physical reserved': 'Inventory Reserved (either against order / transit)',
        'RRPPrice': 'Sales Value', 'QtyOrdered': 'Sales Quantity', 'OrderDate': 'Order Date'
    }

    df_stock = df_stock.rename(columns=columns_rename)
    df_sales = df_sales.rename(columns=columns_rename)

    # Re-verify and pad structural columns
    for col in columns_Output:
        if col not in df_stock.columns: df_stock[col] = np.nan
        if col not in df_sales.columns: df_sales[col] = np.nan

    df_stock = df_stock[columns_Output].copy()
    df_sales = df_sales[columns_Output].copy()

    # Tag records for matrix segmentation loops
    df_sales['Report Type'] = 'Sales'
    df_stock['Report Type'] = 'Stock'

    combined_df = pd.concat([df_sales, df_stock], ignore_index=True)
    
    # Clean up key join columns
    combined_df['Item Name'] = combined_df['Item Name'].astype(str).str.strip()
    email_master['Item Name'] = email_master['Item Name'].astype(str).str.strip()
    
    # Intersect rows against your Email Master mapping array
    combined_df = combined_df.merge(email_master, how='left', on=['Item Name'])
    
    # Fallback assignment for unmapped items
    if 'Email' not in combined_df.columns:
        combined_df['Email'] = 'unmapped@company.com'
    combined_df['Email'] = combined_df['Email'].fillna('unmapped@company.com').astype(str).str.strip()

    # Column limits for structured tab outputs
    sales_specific_cols = [
        'Category Code', 'Brand', 'Brand Description', 'Item Code', 
        'Item Name', 'Product Description', 'Warehouse', 'Sales Quantity', 
        'Sales Value', 'Order Date'
    ]

    stock_specific_cols = [
        'Category Code', 'Brand', 'Item Code', 'Item Name', 
        'Product Description', 'Warehouse', 'Location', 'Total Inventory', 
        'Sellable Inventory', 'Inventory Reserved (either against order / transit)'
    ]

    grouped = combined_df.groupby('Email')
    processed_count = 0
    zip_buffer_dict = {}

    for email_key, group_item in grouped:
        if email_key == 'unmapped@company.com' or email_key == 'NA':
            continue
            
        sales_sheet = group_item[group_item['Report Type'] == 'Sales'].copy()
        stock_sheet = group_item[group_item['Report Type'] == 'Stock'].copy()

        sales_sheet = sales_sheet[[c for c in sales_specific_cols if c in sales_sheet.columns]].dropna(how='all')
        stock_sheet = stock_sheet[[c for c in stock_specific_cols if c in stock_sheet.columns]].dropna(how='all')

        if sales_sheet.empty and stock_sheet.empty:
            continue

        # In-memory Excel compilation loop
        excel_out = io.BytesIO()
        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
            sales_sheet.to_excel(writer, sheet_name='Sales Report', index=False)
            stock_sheet.to_excel(writer, sheet_name='Stock Report', index=False)
            
        excel_bytes = excel_out.getvalue()
        zip_buffer_dict[f"SNS_Report_{email_key}.xlsx"] = excel_bytes
        processed_count += 1

    return processed_count, zip_buffer_dict

def render_ui():
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.subheader("Generate split supplier sheets and configure automated distributions.")
    st.markdown("---")
    
    # Layout Component File Uploaders
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        sales_file = st.file_uploader("Upload Raw Sales Transactions Sheet", type=["xlsx", "csv"], key="sns_sales")
        stock_file = st.file_uploader("Upload Active Stock Balancing Sheet", type=["xlsx", "csv"], key="sns_stock")
    with col_u2:
        email_master_file = st.file_uploader("Upload Brand Email Master Mapping Sheet", type=["xlsx", "csv"], key="sns_email")
        
    st.markdown("### 📅 Temporal Reporting Parameters")
    date_col1, date_col2 = st.columns(2)
    with date_col1:
        start_date = st.date_input("Reporting From Date:", value=datetime.date(2026, 4, 1), key="sns_start")
    with date_col2:
        end_date = st.date_input("Reporting To Date:", value=datetime.date(2026, 4, 19), key="sns_end")

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
                        send_email, mail_subject, mail_body
                    )
                    
                    st.success(f"🎉 Process Complete! Successfully generated multi-tab workbook allocations for {count} suppliers.")
                    
                    if generated_files:
                        import zipfile
                        zip_out = io.BytesIO()
                        with zipfile.ZipFile(zip_out, 'w') as zip_f:
                            for fname, fbytes in generated_files.items():
                                zip_f.writestr(fname, fbytes)
                        
                        st.download_button(
                            label="📥 Download All Generated Partner Workbooks (ZIP Archive)",
                            data=zip_out.getvalue(),
                            file_name=f"SNS_Supplier_Reports_{datetime.datetime.now().strftime('%Y%m%d')}.zip",
                            mime="application/zip",
                            use_container_width=True
                        )
                except Exception as ex:
                    st.error(f"🚨 Automation Processing Error: {str(ex)}")
        else:
            st.warning("Please upload Sales, Stock, and Email Master files simultaneously to generate workbooks.")

if __name__ == "__main__":
    render_ui()
