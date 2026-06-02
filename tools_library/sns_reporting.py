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

TOOL_NAME = "Supplier Performance Engine"
TOOL_ICON = "📈"

# 🔒 HARDCODED LICENSE DEADLINE
EXPIRATION_DATE = datetime.date(2026, 7, 1)

def send_local_smtp_email(to_email, subject, body, attachment_bytes, attachment_filename, smtp_server, smtp_port, sender_email, sender_password):
    """Sends emails directly using secure background SMTP server configurations input by the user."""
    try:
        # Build multipart message payload envelope
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Pack the binary Excel sheet attachment stream
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{attachment_filename}"')
        msg.attach(part)

        # Establish direct secure web connection loop
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()  # Secure connection via TLS
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_email, msg.as_string())
        server.quit()
        
        return "🚀 Dispatched via Local SMTP Relay"
    except Exception as e:
        return f"❌ Transmission Error: {str(e)}"

def run_sns_process(sales_file, stock_file, email_master_file, send_email, mail_subject, mail_body, start_date, end_date, smtp_settings):
    df_sales = pd.read_excel(sales_file) if sales_file.name.endswith('.xlsx') else pd.read_csv(sales_file)
    df_stock = pd.read_excel(stock_file) if stock_file.name.endswith('.xlsx') else pd.read_csv(stock_file)
    email_master = pd.read_excel(email_master_file) if email_master_file.name.endswith('.xlsx') else pd.read_csv(email_master_file)

    df_sales.columns = [str(c).strip() for c in df_sales.columns]
    df_stock.columns = [str(c).strip() for c in df_stock.columns]
    email_master.columns = [str(c).strip() for c in email_master.columns]

    if 'Article Name' in email_master.columns and 'Email' in email_master.columns:
        email_master['Article Name'] = email_master['Article Name'].astype(str).str.strip().str.upper()
        email_map = email_master.drop_duplicates(subset=['Article Name']).set_index('Article Name')['Email'].to_dict()
    else:
        raise ValueError("Email Master must contain exact 'Article Name' and 'Email' headers.")

    if 'Invoice Created On' in df_sales.columns:
        df_sales['Invoice Created On'] = pd.to_datetime(df_sales['Invoice Created On'], errors='coerce')
        df_sales = df_sales[(df_sales['Invoice Created On'].dt.date >= start_date) & (df_sales['Invoice Created On'].dt.date <= end_date)].copy()

    sales_cols = ['Site', 'Site Name', 'Sales Office', 'Article', 'Item Description', 'Article Name', 'Family Name', 'Sub-Family Name', 'Category Name', 'Brand Name', 'RRP Price', 'Invoice Created On', 'Invoice Quantity', 'Amount@RRP']
    stock_cols = ['Article', 'Article Name', 'Item Description', 'Site', 'Site Name', 'Location Name', 'Family Name', 'Sub-Family Name', 'Brand Name', 'Category Name', 'Physical Stock', 'Consignment Stock']

    for col in sales_cols:
        if col not in df_sales.columns: df_sales[col] = np.nan
    for col in stock_cols:
        if col not in df_stock.columns: df_stock[col] = np.nan

    df_sales_clean = df_sales[sales_cols].copy()
    df_stock_clean = df_stock[stock_cols].copy()

    df_sales_clean['Article Name'] = df_sales_clean['Article Name'].astype(str).str.strip().str.upper()
    df_stock_clean['Article Name'] = df_stock_clean['Article Name'].astype(str).str.strip().str.upper()

    df_sales_clean['Email'] = df_sales_clean['Article Name'].map(email_map)
    df_stock_clean['Email'] = df_stock_clean['Article Name'].map(email_map)

    all_emails = set(df_sales_clean['Email'].dropna().unique()).union(set(df_stock_clean['Email'].dropna().unique()))
    
    processed_count = 0
    zip_buffer_dict = {}
    delivery_log = []

    for email in all_emails:
        email_str = str(email).strip()
        if email_str.lower() in ['unmapped@company.com', 'na', 'nan', '', 'na;na']: continue

        sales_final = df_sales_clean[df_sales_clean['Email'] == email_str].copy()
        stock_final = df_stock_clean[df_stock_clean['Email'] == email_str].copy()

        if not sales_final.empty and 'Invoice Created On' in sales_final.columns:
            sales_final['Invoice Created On'] = pd.to_datetime(sales_final['Invoice Created On']).dt.date

        sales_final = sales_final.drop(columns=['Email'], errors='ignore').dropna(how='all')
        stock_final = stock_final.drop(columns=['Email'], errors='ignore').dropna(how='all')

        if sales_final.empty and stock_final.empty: continue

        clean_filename = f"SNS_Report_{email_str}.xlsx".replace("'", "").replace("(", "").replace(")", "")
        
        # Build entirely in-memory to keep things lightning fast
        out_buffer = io.BytesIO()
        with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
            sales_final.to_excel(writer, sheet_name='Sales Report', index=False)
            stock_final.to_excel(writer, sheet_name='Stock Report', index=False)
        
        file_bytes = out_buffer.getvalue()
        zip_buffer_dict[clean_filename] = file_bytes
        processed_count += 1

        if send_email:
            status_message = send_local_smtp_email(
                to_email=email_str,
                subject=mail_subject,
                body=mail_body,
                attachment_bytes=file_bytes,
                attachment_filename=clean_filename,
                **smtp_settings
            )
            delivery_log.append({"Partner Email": email_str, "Status": status_message})
        else:
            delivery_log.append({"Partner Email": email_str, "Status": "💾 Workbook Generated (ZIP package)"})

    return processed_count, zip_buffer_dict, delivery_log

def render_ui():
    current_today = datetime.date.today()
    if current_today > EXPIRATION_DATE:
        st.error(f"🚨 LICENSE EXPIRED: This software expired on {EXPIRATION_DATE.strftime('%d-%b-%Y')}. Please contact administration for a token refresh key.")
        return

    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.caption(f"🔒 Corporate License Node Status: Active (Expires: {EXPIRATION_DATE.strftime('%d-%b-%Y')})")
    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    with col1:
        sales_f = st.file_uploader("📂 Drop Raw Sales Spreadsheet (Sales.xlsx)", type=["xlsx"])
    with col2:
        stock_f = st.file_uploader("📂 Drop Active Stock Balancing Sheet (stock.xlsx)", type=["xlsx"])
    with col3:
        master_f = st.file_uploader("📂 Drop Brand Email Mapping Master File", type=["xlsx"])

    st.markdown("### 📅 Execution Constraints Configuration")
    d_col1, d_col2 = st.columns(2)
    with d_col1:
        start_d = st.date_input("Reporting From Window Start Date:", value=datetime.date(2026, 5, 1))
    with d_col2:
        end_d = st.date_input("Reporting To Window End Date:", value=datetime.date(2026, 5, 31))

    # --- 🔐 NEW: LOCAL HOST SMTP CHANNEL INTERFACE ---
    st.markdown("### ⚙️ Local SMTP Server Credentials")
    with st.expander("📬 Configure Mail Server Connection", expanded=True):
        sc1, sc2 = st.columns(2)
        with sc1:
            smtp_server = st.text_input("SMTP Server Host:", value="smtp.office365.com", help="Office 365: smtp.office365.com | Gmail: smtp.gmail.com")
            sender_email = st.text_input("Your Corporate Email Address:")
        with sc2:
            smtp_port = st.number_input("SMTP Port Connection:", value=587, step=1)
            sender_password = st.text_input("Your Account App Password:", type="password", help="If using MFA, generate an App Password in your security account profile panel.")

    st.markdown("### ✉️ Distribution Messaging Template")
    send_mail_toggle = st.checkbox("Enable Background Email Delivery via SMTP?", value=True)
    
    mail_sub = st.text_input("Outbound Subject Header Line Blueprint:", value="Sales & Stocks Report | Update Window Cycle")
    mail_txt = st.text_area("Global Outbound Context Body String Layout:", value="Dear valued partner,\n\nPlease find attached your custom Sales and Stock Performance Report.\n\nThanks,\nCategory Management Team", height=120)

    st.markdown("---")

    if st.button("⚡ Execute Reporting Allocation Loops", type="primary"):
        if sales_f and stock_f and master_f:
            if send_mail_toggle and (not sender_email or not sender_password):
                st.error("Please supply both your email address and password inside the configurations panel above.")
                return

            with st.spinner("Partitioning worksheets and processing background SMTP transmission streams..."):
                try:
                    smtp_settings = {
                        "smtp_server": smtp_server,
                        "smtp_port": int(smtp_port),
                        "sender_email": sender_email,
                        "sender_password": sender_password
                    }

                    count, files_dict, logs = run_sns_process(sales_f, stock_f, master_f, send_mail_toggle, mail_sub, mail_txt, start_d, end_d, smtp_settings)
                    
                    if count > 0:
                        st.success(f"🎉 Process Complete! Successfully generated clean data splits for {count} unique vendors.")
                        
                        if logs:
                            st.markdown("### 📬 Delivery Status Tracking Log")
                            st.dataframe(pd.DataFrame(logs), use_container_width=True, hide_index=True)
                        
                        if files_dict:
                            import zipfile
                            z_io = io.BytesIO()
                            with zipfile.ZipFile(z_io, 'w') as zf:
                                for fn, fb in files_dict.items():
                                    zf.writestr(fn, fb)
                            st.download_button("📥 Download All Partner Workbooks Package (ZIP Archive File)", data=z_io.getvalue(), file_name=f"SNS_Supplier_Bundles_{datetime.datetime.now().strftime('%Y%m%d')}.zip", mime="application/zip")
                    else:
                        st.warning("⚠️ No matching row records found for the specified metrics.")
                except Exception as err:
                    st.error(f"🚨 Core Execution Failure Block: {str(err)}")
        else:
            st.warning("Please upload all three workbook project files into the uploader slots simultaneously to start parsing.")
