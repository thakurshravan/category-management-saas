# --- RUN THIS CELL IN YOUR JUPYTER NOTEBOOK ---
import ipywidgets as widgets
from IPython.display import display, HTML
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
# 🔐 HARDCODED LOCAL DEFAULTS (Change these once, never re-type)
# ==========================================================
DEFAULT_SMTP_SERVER = "smtp.office365.com"   # Office365/Outlook default
DEFAULT_SMTP_PORT = 587
DEFAULT_SENDER_EMAIL = "your-email@company.com" 
DEFAULT_SENDER_PASSWORD = "your-app-password" # Leave blank if you prefer typing it in real-time

def send_local_email(to_email, subject, body, attachment_bytes, attachment_filename):
    """Quietly dispatches emails via local network sockets without opening Outlook windows"""
    try:
        msg = MIMEMultipart()
        msg['From'] = DEFAULT_SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment_bytes)
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{attachment_filename}"')
        msg.attach(part)

        server = smtplib.SMTP(DEFAULT_SMTP_SERVER, DEFAULT_SMTP_PORT)
        server.starttls()
        server.login(DEFAULT_SENDER_EMAIL, DEFAULT_SENDER_PASSWORD)
        server.sendmail(DEFAULT_SENDER_EMAIL, to_email, msg.as_string())
        server.quit()
        return "🚀 Dispatched via Local SMTP"
    except Exception as e:
        return f"❌ Error: {str(e)}"

def process_and_split(sales_data, stock_data, master_data, start_date, end_date):
    """Processes datasets completely in-memory"""
    df_sales = pd.read_excel(io.BytesIO(sales_data))
    df_stock = pd.read_excel(io.BytesIO(stock_data))
    email_master = pd.read_excel(io.BytesIO(master_data))

    df_sales.columns = [str(c).strip() for c in df_sales.columns]
    df_stock.columns = [str(c).strip() for c in df_stock.columns]
    email_master.columns = [str(c).strip() for c in email_master.columns]

    email_master['Article Name'] = email_master['Article Name'].astype(str).str.strip().str.upper()
    email_map = email_master.drop_duplicates(subset=['Article Name']).set_index('Article Name')['Email'].to_dict()

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
    
    logs = []
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

        clean_filename = f"SNS_Report_{email_str}.xlsx"
        
        out_buffer = io.BytesIO()
        with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
            sales_final.to_excel(writer, sheet_name='Sales Report', index=False)
            stock_final.to_excel(writer, sheet_name='Stock Report', index=False)
        
        status = send_local_email(email_str, "Sales & Stock Report Update", "Dear Partner,\n\nPlease find attached your report.", out_buffer.getvalue(), clean_filename)
        logs.append({"Partner Email": email_str, "Status": status})
        
    return pd.DataFrame(logs)

# ==========================================================
# 🎨 INTERACTIVE JUPYTER UI LAYOUT WIDGETS
# ==========================================================
upload_sales = widgets.FileUpload(description="1. Sales File", accept=".xlsx", multiple=False)
upload_stock = widgets.FileUpload(description="2. Stock File", accept=".xlsx", multiple=False)
upload_master = widgets.FileUpload(description="3. Email Master", accept=".xlsx", multiple=False)

start_date_picker = widgets.DatePicker(description='From:', value=datetime.date(2026, 5, 1))
end_date_picker = widgets.DatePicker(description='To:', value=datetime.date(2026, 5, 31))
btn_run = widgets.Button(description="⚡ Run Processing Loop", button_style='primary')
output_area = widgets.Output()

def on_button_click(b):
    with output_area:
        output_area.clear_output()
        if not upload_sales.value or not upload_stock.value or not upload_master.value:
            print("❌ Please upload all three files directly via the buttons above first.")
            return
        
        print("⏳ Processing in-memory data splits... Please wait.")
        
        # Pull raw binary data out of the widgets
        sales_bytes = list(upload_sales.value.values())[0]['content']
        stock_bytes = list(upload_stock.value.values())[0]['content']
        master_bytes = list(upload_master.value.values())[0]['content']
        
        try:
            log_df = process_and_split(
                sales_bytes, stock_bytes, master_bytes, 
                start_date_picker.value, end_date_picker.value
            )
            print("🎉 Done! Review your output distribution logs below:")
            display(log_df)
        except Exception as err:
            print(f"🚨 Process Broken: {str(err)}")

btn_run.on_click(on_button_click)

# Display everything cleanly in the notebook output cell area
print("⚙️ SUPPLIER PERFORMANCE ENGINE (INTERACTIVE CELL)")
display(widgets.HBox([upload_sales, upload_stock, upload_master]))
display(widgets.HBox([start_date_picker, end_date_picker]))
display(btn_run)
display(output_area)
