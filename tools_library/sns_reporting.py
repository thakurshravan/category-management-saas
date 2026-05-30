# tools_library/sns_reporting.py
import streamlit as st
import pandas as pd
import numpy as np
import io
import datetime
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ==========================================================
# 🏷️ MODULE METADATA & ENTRY REGISTRATION FOR STREAMLIT SaaS
# ==========================================================
TOOL_NAME = "SNS Reporting Tool"
TOOL_ICON = "📈"

def clean_headers(df):
    """Strips hidden spaces from input spreadsheet headers securely."""
    df.columns = [str(c).strip() for c in df.columns]
    return df

def send_cloud_smtp_email(recipient_email, subject, body_text, file_name, file_bytes):
    """Fallback secure SMTP automation engine when tool is running in the cloud."""
    email_creds = st.secrets.get("email_auth", {})
    if not email_creds:
        return False, "SMTP configuration keys are missing from Streamlit secrets dashboard parameters."
        
    try:
        msg = MIMEMultipart()
        msg['From'] = email_creds.get("sender_email")
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body_text, 'plain'))
        
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
        msg.attach(part)
        
        server = smtplib.SMTP(email_creds.get("smtp_server", "smtp.office365.com"), int(email_creds.get("smtp_port", 587)))
        server.starttls()
        server.login(email_creds.get("sender_email"), email_creds.get("sender_password"))
        server.sendmail(email_creds.get("sender_email"), recipient_email, msg.as_string())
        server.quit()
        return True, "Sent via Cloud SMTP Relay"
    except Exception as error:
        return False, str(error)

def run_sns_process(sales_file, stock_file, email_master_file, send_email, mail_subject, mail_body, start_date, end_date):
    """Executes isolated sheet splits matching exact partner specifications with environment auto-detection."""
    
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

    if 'Article Name' in email_master.columns:
        email_master = email_master.drop_duplicates(subset=['Article Name'], keep='first')
        
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

    # 4. Independent Merging
    df_sales_mapped = df_sales[sales_specific_cols].merge(email_master, how='inner', on='Article Name')
    df_stock_mapped = df_stock[stock_specific_cols].merge(email_master, how='inner', on='Article Name')

    all_emails = set(df_sales_mapped['Email'].dropna().unique()).union(set(df_stock_mapped['Email'].dropna().unique()))

    processed_count = 0
    zip_buffer_dict = {}
    email_delivery_report = []

    # 🚨 SYSTEM ENVIRONMENT DETECTOR 
    # Check if we are running locally on Windows OS or deployed in Linux Cloud Server container
    is_local_windows = os.name == 'nt'
    
    if send_email and is_local_windows:
        import win32com.client as win32
        temp_dir = "./temp_sns_outbound/"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

    # 5. Segment and build workbooks individually for each supplier email group
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

        # ─── ✉️ CROSS-COMPATIBLE DISPATCH ENGINE LOOP ───
        if send_email:
            if is_local_windows:
                # PATH A: Local Mode -> Route via Windows Desktop MAPI (No passwords needed)
                try:
                    local_file_path = os.path.abspath(os.path.join(temp_dir, clean_filename))
                    with open(local_file_path, "wb") as f:
                        f.write(excel_bytes)
                    
                    outlook = win32.Dispatch('outlook.application')
                    message = outlook.CreateItem(0)
                    message.Subject = mail_subject
                    message.Body = mail_body
                    message.To = email_str
                    message.Attachments.Add(local_file_path)
                    message.Send()
                    
                    email_delivery_report.append({"Partner Email": email_str, "Status": "Sent via Local Outlook App"})
                except Exception as local_err:
                    email_delivery_report.append({"Partner Email": email_str, "Status": f"Desktop Error: {str(local_err)}"})
            else:
                # PATH B: Cloud Server Mode -> Route via Secure Office 365 Network SMTP Relay using Secrets token
                success, cloud_msg = send_cloud_smtp_email(email_str, mail_subject, mail_body, clean_filename, excel_bytes)
                email_delivery_report.append({
                    "Partner Email": email_str, 
                    "Status": "Sent via Cloud Server Integration" if success else f"Cloud Server Error: {cloud_msg}"
                })

    # Clear out local file caching parameters if running in local desktop pass
    if send_email and is_local_windows and os.path.exists(temp_dir):
        for f in os.listdir(temp_dir):
            try: os.remove(os.path.join(temp_dir, f))
            except: pass

    return processed_count, zip_buffer_dict, email_delivery_report

def render_ui():
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.subheader("Smart Desktop-Cloud Hybrid Architecture Engine.")
    st.markdown("---")
    
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
            with st.spinner("Processing deep matrix calculations and routing distributions..."):
                try:
                    count, generated_files, email_report = run_sns_process(
                        sales_file, stock_file, email_master_file, 
                        send_email, mail_subject, mail_body, start_date, end_date
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
