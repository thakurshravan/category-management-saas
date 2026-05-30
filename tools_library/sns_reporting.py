import os
import datetime
import pandas as pd
import numpy as np
import win32com.client as win32

# ==========================================
# 🛠️ 1. SETUP PARAMETERS & CONFIGURATION
# ==========================================
# Update these paths to where your actual files are saved on your laptop
sales_file_path = "./Sales_Data.xlsx"
stock_file_path = "./Stock_Data.xlsx"
email_master_path = "./Email_Master - All.xlsx"

# Global Email Templates
mail_subject = "Sales & Stocks Performance Report Update"
mail_body = """Dear valued partner,

Please find attached your Sales and Stock Performance Report covering the requested window.

Thanks,
Category Management Team
Jumbo Electronics
"""

# Date range boundaries (Adjust these as needed)
start_date = datetime.date(2026, 5, 1)
end_date = datetime.date(2026, 5, 31)

# Set to True to send emails instantly; False will just save drafts in Outlook
send_instantly = True  

# ==========================================
# 📊 2. DATA LOADING & CLEANING
# ==========================================
print("🔄 Loading files...")
df_sales = pd.read_excel(sales_file_path) if sales_file_path.endswith('.xlsx') else pd.read_csv(sales_file_path)
df_stock = pd.read_excel(stock_file_path) if stock_file_path.endswith('.xlsx') else pd.read_csv(stock_file_path)
email_master = pd.read_excel(email_master_path) if email_master_path.endswith('.xlsx') else pd.read_csv(email_master_path)

# Strip hidden spaces from headers automatically
df_sales.columns = [str(c).strip() for c in df_sales.columns]
df_stock.columns = [str(c).strip() for c in df_stock.columns]
email_master.columns = [str(c).strip() for c in email_master.columns]

# Ensure the Email Master has the correct layout columns
if 'Article Name' not in email_master.columns or 'Email' not in email_master.columns:
    raise ValueError("Your Email Master must contain exact 'Article Name' and 'Email' columns!")

# Clean up the key column for flawless matching
email_master['Article Name'] = email_master['Article Name'].astype(str).str.strip().str.upper()
email_master = email_master.drop_duplicates(subset=['Article Name'])
email_map = email_master.set_index('Article Name')['Email'].to_dict()

# ==========================================
# 📅 3. DATE FILTERING (Sales Only)
# ==========================================
if 'Invoice Created On' in df_sales.columns:
    df_sales['Invoice Created On'] = pd.to_datetime(df_sales['Invoice Created On'], errors='coerce')
    df_sales = df_sales[
        (df_sales['Invoice Created On'].dt.date >= start_date) & 
        (df_sales[date_col].dt.date <= end_date)
    ].copy()
else:
    print("⚠️ Warning: 'Invoice Created On' column not found in Sales file. Skipping date filter.")

# ==========================================
# 🎯 4. STABILIZE EXACT SCHEMA COLUMNS
# ==========================================
# These lists match your exact raw column descriptions
sales_output_cols = [
    'Site', 'Site Name', 'Sales Office', 'Article', 'Item Description', 
    'Article Name', 'Family Name', 'Sub-Family Name', 'Category Name', 
    'Brand Name', 'RRP Price', 'Invoice Created On', 'Invoice Quantity', 'Amount@RRP'
]

stock_output_cols = [
    'Article', 'Article Name', 'Item Description', 'Site', 'Site Name', 
    'Location Name', 'Family Name', 'Sub-Family Name', 'Brand Name', 
    'Category Name', 'Physical Stock', 'Consignment Stock'
]

# Guarantee columns exist to prevent script crashes
for col in sales_output_cols:
    if col not in df_sales.columns: df_sales[col] = np.nan
for col in stock_output_cols:
    if col not in df_stock.columns: df_stock[col] = np.nan

# Slice and isolate clean working datasets
df_sales_clean = df_sales[sales_output_cols].copy()
df_stock_clean = df_stock[stock_output_cols].copy()

# Standardize match keys across dataframes
df_sales_clean['Article Name'] = df_sales_clean['Article Name'].astype(str).str.strip().str.upper()
df_stock_clean['Article Name'] = df_stock_clean['Article Name'].astype(str).str.strip().str.upper()

# Inject recipient emails via Article Name lookup map
df_sales_clean['Email'] = df_sales_clean['Article Name'].map(email_map)
df_stock_clean['Email'] = df_stock_clean['Article Name'].map(email_map)

# Isolate list of all target vendor emails
all_emails = set(df_sales_clean['Email'].dropna().unique()).union(set(df_stock_clean['Email'].dropna().unique()))

# Create a temporary local folder to hold outbound Excel attachments
temp_folder = os.path.abspath("./temp_outbound_reports")
if not os.path.exists(temp_folder):
    os.makedirs(temp_folder)

# ==========================================
# ✉️ 5. SPLIT PROCESSOR & OUTLOOK AUTOMATION LOOP
# ==========================================
print(f"🚀 Found {len(all_emails)} unique partner emails to process.")
processed_count = 0

for email in all_emails:
    email_str = str(email).strip()
    # Skip any unmapped placeholders
    if email_str.lower() in ['unmapped@company.com', 'na', 'nan', '', 'na;na']:
        continue
        
    # Extract matching row records
    sales_final = df_sales_clean[df_sales_clean['Email'] == email_str].copy()
    stock_final = df_stock_clean[df_stock_clean['Email'] == email_str].copy()
    
    # Format dates nicely for final Excel presentation sheet layout
    if not sales_final.empty and 'Invoice Created On' in sales_final.columns:
        sales_final['Invoice Created On'] = pd.to_datetime(sales_final['Invoice Created On']).dt.date
        
    # Strip distribution tracking column before writing file to vendor
    sales_final = sales_final.drop(columns=['Email'], errors='ignore').dropna(how='all')
    stock_final = stock_final.drop(columns=['Email'], errors='ignore').dropna(how='all')
    
    if sales_final.empty and stock_final.empty:
        continue
        
    # Build clean filename path string safely
    clean_filename = f"SNS_Report_{email_str}.xlsx".replace("'", "").replace("(", "").replace(")", "")
    file_output_path = os.path.join(temp_folder, clean_filename)
    
    # Write the dual-tab spreadsheet out to disk path
    with pd.ExcelWriter(file_output_path, engine='openpyxl') as writer:
        sales_final.to_excel(writer, sheet_name='Sales Report', index=False)
        stock_final.to_excel(writer, sheet_name='Stock Report', index=False)
        
    # Connect directly to your native Windows Outlook application session
    try:
        outlook = win32.Dispatch('outlook.application')
        message = outlook.CreateItem(0)
        message.Subject = mail_subject
        message.Body = mail_body
        message.To = email_str
        
        # Attach the local generated Excel document
        message.Attachments.Add(file_output_path)
        
        if send_instantly:
            message.Send()
            print(f"✅ Sent successfully to: {email_str}")
        else:
            message.Save()
            print(f"💾 Saved as draft for: {email_str}")
            
        processed_count += 1
    except Exception as mail_error:
        print(f"❌ Failed to dispatch email to {email_str}. Error: {str(mail_error)}")

# Cleanup generated folder assets on disk when loop finishes
for f in os.listdir(temp_folder):
    try: os.remove(os.path.join(temp_folder, f))
    except: pass
try: os.rmdir(temp_folder)
except: pass

print(f"\n🎉 All tasks complete! Processed and sent {processed_count} vendor reports.")
