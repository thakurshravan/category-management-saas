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

def send_cloud_smtp_email(smtp_server, smtp_port, sender_email, sender_password, recipient_email, subject, body_text, file_name, file_bytes):
    """Fallback secure SMTP automation engine when tool runs in Cloud SaaS mode."""
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
        
        server = smtplib.SMTP(smtp_server, int(smtp_port))
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_email, msg.as_string())
        server.quit()
        return True, "Sent via Cloud SMTP Gateway Relay"
    except Exception as error:
        return False, str(error)

def run_sns_process(sales_file, stock_file, email_master_file, send_email, mail_subject, mail_body, start_date, end_date, execution_mode, smtp_settings=None):
    """Executes isolated sheet splits matching your exact raw Sales & Stock column specifications."""
    
    # Load Input Data Streams Safely from Memory Buffers
    df_sales = pd.read_excel(sales_file) if sales_file.name.endswith('.xlsx') else pd.read_csv(sales_file)
    df_stock = pd.read_excel(stock_file) if stock_file.name.endswith('.xlsx') else pd.read_csv(stock_file)
    email_master = pd.read_excel(email_master_file) if email_master_file.name.endswith('.xlsx') else pd.read_csv(email_master_file)
    
    df_sales = clean_headers(df_sales)
    df_stock = clean_headers(df_stock)
    email_master = clean_headers(email_master)

    # Validate that Email Master has the columns shown in your layout image
    if 'Article Name' in email_master.columns and 'Email' in email_master.columns:
        email_master['Article Name'] = email_master['Article Name'].astype(str).str.strip().str.upper()
        email_map = email_master.drop_duplicates(subset=['Article Name']).set_index('Article Name')['Email'].to_dict()
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

    # EXACT MATCH RAW INPUT COLUMNS MATRICES
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

    # Initialize columns if missing inside source frames to prevent runtime key errors safely
    for col in sales_specific_cols:
        if col not in df_sales.columns: df_sales[col] = ""
            
    for col in stock_specific_cols:
        if col not in df_stock.columns: df_stock[col] = ""

    # Slice clean structured working views matching your exact inputs
    df_sales_clean = df_sales[sales_specific_cols].copy()
    df_stock_clean = df_stock[stock_specific_cols].copy()

    # Standardize comparison keys to perform clean case-insensitive mappings
    df_sales_clean['Article Name'] = df_sales_clean['Article Name'].astype(str).str.strip().str.upper()
    df_stock_clean['Article Name'] = df_stock_clean['Article Name'].astype(str).str.strip().str.upper()

    # Map the master recipient emails directly onto the data rows using Article Name keys
    df_sales_clean['Email'] = df_sales_clean['Article Name'].map(email_map)
    df_stock_clean['Email'] = df_stock_clean['Article Name'].map(email_map)

    # Intersect all matched unique partner email keys
    all_emails = set(df_sales_clean['Email'].dropna().unique()).union(set(df_stock_clean['Email'].dropna().unique()))

    processed_count = 0
    zip_buffer_dict = {}
    email_delivery_report = []

    # Local Windows Environment detection configuration
    use_local_win = (execution_mode == "Local Desktop App (Outlook)") and (os.name == 'nt')
    if send_email and use_local_win:
        import win32com.client as win32
        temp_dir = "./temp_sns_outbound/"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)

    # Segment and build separate data tabs for each email profile recipient group
    for email in all_emails:
        email_str = str(email).strip()
        if email_str.lower() in ['unmapped@company.com', 'na', 'nan', '', 'na;na']:
            continue

        sales_final = df_sales_clean[df_sales_clean['Email'] == email_str].copy()
        stock_final = df_stock_clean[df_stock_clean['Email'] == email_str].copy()

        # Format timestamps cleanly into standard text short dates for Excel display
        if not sales_final.empty and 'Invoice Created On' in sales_final.columns:
            sales_final['Invoice Created On'] = pd.to_datetime(sales_final['Invoice Created On']).dt.date

        # Drop runtime routing flags before output file creation
        sales_final = sales_final.drop(columns=['Email'], errors='ignore').dropna(how='all')
        stock_final = stock_final.drop(columns=['Email'], errors='ignore').dropna(how='all')

        if sales_final.empty and stock_final.empty:
            continue

        clean_filename = f"SNS_Report_{email_str}.xlsx".replace("'", "").replace("(", "").replace(")", "")
        
        # Build multi-tab clean workbook package frames inside memory
        excel_out = io.BytesIO()
        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
            sales_final.to_excel(writer, sheet_name='Sales Report', index=False)
            stock_final.to_excel(writer, sheet_name='Stock Report', index=False)
            
        excel_bytes = excel_out.getvalue()
        zip_buffer_dict[clean_filename] = excel_bytes
        processed_count += 1

        # ─── ✉️ OUTBOUND DISPATCH INTERFACE ROUTER BLOCK ───
        if send_email:
            if use_local_win:
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
                    
                    email_delivery_report.append({"Partner Email": email_str, "Status": "Sent via Desktop Outlook Application"})
                except Exception as local_err:
                    email_delivery_report.append({"Partner Email": email_str, "Status": f"Desktop Integration Error: {str(local_err)}"})
            else:
                if smtp_settings:
                    success, cloud_msg = send_cloud_smtp_email(
                        smtp_settings['server'], smtp_settings['port'], smtp_settings['email'], smtp_settings['password'],
                        email_str, mail_subject, mail_body, clean_filename, excel_bytes
                    )
                    email_delivery_report.append({
                        "Partner Email": email_str, 
                        "Status": "Sent via Cloud Server SMTP" if success else f"Cloud Security Block: {cloud_msg}"
                    })

    # Flush local disk folder parameters upon execution completion loop
    if send_email and use_local_win and os.path.exists(temp_dir):
        for f in os.listdir(temp_dir):
            try: os.remove(os.path.join(temp_dir, f))
            except: pass

    return processed_count, zip_buffer_dict, email_delivery_report

def render_ui():
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.subheader("Process transactional allocations and handle partner distribution loops.")
    st.markdown("---")
    
    with st.sidebar:
        st.header("🔑 Operational Environment Settings")
        run_mode = st.radio(
            "Select Execution Mode Routing:",
            ["Local Desktop App (Outlook)", "Cloud Server Gateway (SMTP)"],
            help="Choose 'Local Desktop' if running the app on your computer to send via your active Outlook window."
        )
        st.markdown("---")
        
        if run_mode == "Cloud Server Gateway (SMTP)":
            st.subheader("SaaS Server Configuration")
            sys_auth = st.secrets.get("email_auth", {})
            client_smtp_server = st.text_input("SMTP Network Gateway:", value=sys_auth.get("smtp_server", "smtp.office365.com"))
            client_smtp_port = st.number_input("SMTP Port Connection:", value=int(sys_auth.get("smtp_port", 587)))
            client_sender_email = st.text_input("Corporate Username Email ID:", value=sys_auth.get("sender_email", "shravan.kumar@jumbo.ae"))
            client_sender_password = st.text_input("App-Authentication Password Token:", value=sys_auth.get("sender_password", ""), type="password")
        else:
            st.success("💻 Running in Local Desktop Mode. The script will securely hook into your machine's pre-approved work Outlook session. No password entries required!")
            client_smtp_server, client_smtp_port, client_sender_email, client_sender_password = "", 587, "", ""

    # Main UI Data File Ingestors
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
            "Category Management Team\n"
            "Jumbo Electronics"
        )
        mail_subject = st.text_input("Email Global Subject Line:", value="Sales & Stocks Report Update", key="sns_subject_input")
        mail_body = st.text_area("Email Global Body Text Content:", value=default_body_str, height=180, key="sns_body_input")
    else:
        mail_subject = ""
        mail_body = ""

    st.markdown("---")

    if st.button("⚡ Execute Reporting Automation Loops", type="primary", use_container_width=True, key="sns_run"):
        if sales_file and stock_file and email_master_file:
            with st.spinner("Processing calculations and preparing distribution sheets..."):
                try:
                    active_smtp_package = {
                        'server': client_smtp_server,
                        'port': client_smtp_port,
                        'email': client_sender_email,
                        'password': client_sender_password
                    }
                    
                    count, generated_files, email_report = run_sns_process(
                        sales_file, stock_file, email_master_file, 
                        send_email, mail_subject, mail_body, start_date, end_date,
                        execution_mode=run_mode, smtp_settings=active_smtp_package
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
