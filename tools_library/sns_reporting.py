# tools_library/sns_reporting.py
import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ==========================================================
# 🏷️ MODULE METADATA & ENTRY REGISTRATION FOR STREAMLIT SaaS
# ==========================================================
TOOL_NAME = "Enterprise Reporting & Distribution Suite"
TOOL_ICON = "🚀"

def clean_headers(df):
    """Strips hidden spaces from input spreadsheet headers securely."""
    df.columns = [str(c).strip() for c in df.columns]
    return df

def send_partner_email(smtp_server, smtp_port, sender_email, sender_password, recipient_email, subject, body_text, file_name, file_bytes):
    """SaaS Custom SMTP automation engine that scales dynamically based on current user settings."""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body_text, 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
        msg.attach(part)
        
        # Authenticate using current subscriber settings inputs dynamically
        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        return True, "Success"
    except Exception as error:
        return False, str(error)

def run_sns_process(sales_file, stock_file, email_master_file, send_email, mail_subject, mail_body, start_date, end_date, smtp_settings=None):
    """Executes isolated sheet splits matching exact partner specifications."""
    
    # Load Input Data Streams Safely from Memory Buffers
    df_sales = pd.read_excel(sales_file) if sales_file.name.endswith('.xlsx') else pd.read_csv(sales_file)
    df_stock = pd.read_excel(stock_file) if stock_file.name.endswith('.xlsx') else pd.read_csv(stock_file)
    email_master = pd.read_excel(email_master_file) if email_master_file.name.endswith('.xlsx') else pd.read_csv(email_master_file)
    
    df_sales = clean_headers(df_sales)
    df_stock = clean_headers(df_stock)
    email_master = clean_headers(email_master)

    # Standardize common join key column across all datasets
    for target_df in [df_sales, df_stock, email_master]:
        if 'Article Name' in target_df.columns:
            target_df['Article Name'] = target_df['Article Name'].astype(str).str.strip().str.upper()

    if 'Article Name' in email_master.columns:
        email_master = email_master.drop_duplicates(subset=['Article Name'], keep='first')
        
    if 'Email' in email_master.columns and 'Article Name' in email_master.columns:
        email_master = email_master[['Article Name', 'Email']]
    else:
        raise ValueError("The uploaded Email Master sheet must contain exact 'Article Name' and 'Email' headers.")

    # Strict Date Filtering Layer on Sales Sheet via "Invoice Created On"
    date_col = 'Invoice Created On'
    if date_col in df_sales.columns:
        df_sales[date_col] = pd.to_datetime(df_sales[date_col], errors='coerce')
        df_sales = df_sales[
            (df_sales[date_col].dt.date >= start_date) & 
            (df_sales[date_col].dt.date <= end_date)
        ].copy()

    # Structural Renaming & Data Alignment
    columns_rename = {
        'Physical inventory': 'Physical Stock',
        'Location': 'Location Name'
    }
    df_sales = df_sales.rename(columns=columns_rename)
    df_stock = df_stock.rename(columns=columns_rename)

    # EXACT REFINED COLUMN PICKING ARRAYS
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

    for col in sales_specific_cols:
        if col not in df_sales.columns: df_sales[col] = np.nan
            
    for col in stock_specific_cols:
        if col not in df_stock.columns: df_stock[col] = np.nan

    # Independent Merging
    df_sales_mapped = df_sales[sales_specific_cols].merge(email_master, how='inner', on='Article Name')
    df_stock_mapped = df_stock[stock_specific_cols].merge(email_master, how='inner', on='Article Name')

    all_emails = set(df_sales_mapped['Email'].dropna().unique()).union(set(df_stock_mapped['Email'].dropna().unique()))

    processed_count = 0
    zip_buffer_dict = {}
    email_delivery_report = []

    # Segment and build workbooks individually for each supplier email group
    for email in all_emails:
        email_str = str(email).strip()
        if email_str.lower() in ['unmapped@company.com', 'na', 'nan', '', 'na;na']:
            continue

        sales_final = df_sales_mapped[df_sales_mapped['Email'] == email_str].copy()
        stock_final = df_stock_mapped[df_stock_mapped['Email'] == email_str].copy()

        if not sales_final.empty and 'Invoice Created On' in sales_final.columns:
            sales_final['Invoice Created On'] = pd.to_datetime(sales_final['Invoice Created On']).dt.date

        sales_final = sales_final.drop(columns=['Email'], errors='ignore').dropna(how='all')
        stock_final = stock_final.drop(columns=['Email'], errors='ignore').dropna(how='all')

        if sales_final.empty and stock_final.empty:
            continue

        clean_filename = f"SNS_Report_{email_str}.xlsx".replace("'", "").replace("(", "").replace(")", "")
        
        excel_out = io.BytesIO()
        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
            sales_final.to_excel(writer, sheet_name='Sales Report', index=False)
            stock_final.to_excel(writer, sheet_name='Stock Report', index=False)
            
        excel_bytes = excel_out.getvalue()
        zip_buffer_dict[clean_filename] = excel_bytes
        processed_count += 1

        # DYNAMIC EMAIL DISPATCH LOOP
        if send_email and smtp_settings:
            mail_success, status_msg = send_partner_email(
                smtp_server=smtp_settings['server'],
                smtp_port=smtp_settings['port'],
                sender_email=smtp_settings['email'],
                sender_password=smtp_settings['password'],
                recipient_email=email_str,
                subject=mail_subject,
                body_text=mail_body,
                file_name=clean_filename,
                file_bytes=excel_bytes
            )
            email_delivery_report.append({
                "Partner Email": email_str,
                "Status": "Sent Successfully" if mail_success else f"Failed: {status_msg}"
            })

    return processed_count, zip_buffer_dict, email_delivery_report

def render_ui():
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.subheader("Process transactional allocations and handle partner distribution loops.")
    st.markdown("---")
    
    # 🌟 NEW: THE SUBSCRIPTION SIDEBAR WHITE-LABEL PANEL
    with st.sidebar:
        st.header("🔑 Client Gateway Authentication")
        
        # Pull system credentials dynamically fallback securely if not present
        sys_auth = st.secrets.get("email_auth", {})
        
        st.subheader("Configuration Settings")
        client_smtp_server = st.text_input("SMTP Network Gateway:", value=sys_auth.get("smtp_server", "smtp.office365.com"))
        client_smtp_port = st.number_input("SMTP Port connection:", value=int(sys_auth.get("smtp_port", 587)))
        client_sender_email = st.text_input("Corporate Username Email ID:", value=sys_auth.get("sender_email", "shravan.kumar@jumbo.ae"))
        client_sender_password = st.text_input("App-Authentication Password Token:", value=sys_auth.get("sender_password", "lhhclrzfbnphgkzp"), type="password")
        
        st.info("💡 Monthly Subscribers can adjust these sidebar credentials to run distributions directly through their own corporate mail gateways.")

    # Main UI Dashboard Intake Panels
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
    send_email = st.checkbox("Enable Automated Email Distribution Loop?", value=False, key="sns_send")

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
            with st.spinner("Processing calculations and preparing distribution sheets..."):
                try:
                    # Package the sidebar runtime credentials state variables cleanly
                    active_smtp_package = {
                        'server': client_smtp_server,
                        'port': client_smtp_port,
                        'email': client_sender_email,
                        'password': client_sender_password
                    }
                    
                    count, generated_files, email_report = run_sns_process(
                        sales_file, stock_file, email_master_file, 
                        send_email, mail_subject, mail_body, start_date, end_date,
                        smtp_settings=active_smtp_package
                    )
                    
                    if count > 0:
                        st.success(f"🎉 Process Complete! Successfully generated clean data splits for {count} unique suppliers.")
                        
                        if send_email and email_report:
                            st.markdown("### 📬 Live Distribution Tracking Grid")
                            rep_df = pd.DataFrame(email_report)
                            st.dataframe(rep_df, use_container_width=True, hide_index=True)
                        
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
