import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# ==========================================================
# SAAS ORCHESTRATION ATTRIBUTES
# ==========================================================
TOOL_NAME = "Supplier Stock Ingest Portal"
TOOL_ICON = "🌐"

REQUIRED_COLUMNS = [
    "Date", "EAN NUMBER", "SAP CODE", "Article Description", 
    "Item Description", "Brand", "Family Name", "Type", "Cost", "Physical Stock"
]

CLOUD_STORAGE_DIR = "Data/cloud_supplier_drops"
os.makedirs(CLOUD_STORAGE_DIR, exist_ok=True)

def clean_physical_stock(val):
    if pd.isna(val): return 0
    val_str = str(val).upper().strip()
    for suffix in ['EA', 'PCS', 'UNITS', 'UNIT']:
        val_str = val_str.replace(suffix, "")
    val_str = val_str.replace(",", "").strip()
    try: return int(float(val_str))
    except ValueError: return 0

def render_ui():
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.markdown("Automate and control your inventory collection loops with batch generation and link validity rules.")
    st.markdown("---")

    url_params = st.query_params

    if "view" in url_params and url_params["view"] == "portal":
        render_public_supplier_portal(url_params)
    else:
        render_manager_dashboard()

# ==========================================================
# WORKSPACE A: PRIVATE CATEGORY MANAGER DASHBOARD
# ==========================================================
def render_manager_dashboard():
    tab_dashboard, tab_portal_generator = st.tabs(["📊 Consolidated Dashboard", "🔗 Bulk Link Generator & Validity Control"])

    with tab_portal_generator:
        st.subheader("Bulk Link Generation Control Panel")
        st.markdown("Generate secure, time-locked upload links for multiple vendors simultaneously.")
        
        # 1. Bulk Input Channel
        vendors_input = st.text_area(
            "Enter Supplier/Brand Names (Separated by commas):", 
            placeholder="e.g., SONY, SAMSUNG, APPLE, LOGITECH",
            help="Type or paste your brand list separated by commas to process them all at once."
        )
        
        # 2. Validity Window Settings
        st.markdown("### ⏳ Link Validity Settings")
        validity_option = st.selectbox(
            "How long should these links remain active?",
            ["7 Days (Standard Weekly Cycle)", "14 Days", "30 Days", "Custom Date", "Permanent (No Expiration)"]
        )
        
        # Calculate target expiration date string
        current_time = datetime.now()
        if validity_option == "7 Days (Standard Weekly Cycle)":
            expiry_date = (current_time + timedelta(days=7)).strftime("%Y-%m-%d")
        elif validity_option == "14 Days":
            expiry_date = (current_time + timedelta(days=14)).strftime("%Y-%m-%d")
        elif validity_option == "30 Days":
            expiry_date = (current_time + timedelta(days=30)).strftime("%Y-%m-%d")
        elif validity_option == "Custom Date":
            chosen_date = st.date_input("Select Expiration Date:", min_value=current_time.date())
            expiry_date = chosen_date.strftime("%Y-%m-%d")
        else:
            expiry_date = "never"

        # 3. Execution Trigger
        if st.button("Generate All Links Simultaneously", type="primary"):
            if vendors_input:
                vendor_list = [v.strip() for v in vendors_input.split(",") if v.strip()]
                st.success(f"🎉 Instantly Generated {len(vendor_list)} Security-Locked Portals!")
                
                base_url = "https://category-management-saas-fwpgyjvewktqf4repbby8j.streamlit.app"
                
                generated_records = []
                for vendor in vendor_list:
                    clean_slug = vendor.lower().replace(" ", "-")
                    public_url = f"{base_url}/?view=portal&vendor={clean_slug}&exp={expiry_date}"
                    
                    generated_records.append({
                        "Supplier Brand": vendor.upper(),
                        "Valid Until": "Indefinite" if expiry_date == "never" else expiry_date,
                        "Secure Upload Link": public_url
                    })
                
                res_df = pd.DataFrame(generated_records)
                st.dataframe(res_df, use_container_width=True, hide_index=True)
                st.info("💡 **SaaS Copy Tip:** You can highlight and copy these links directly out of the grid table above to paste into your vendor email broadcasts.")
            else:
                st.warning("Please input at least one brand name to generate links.")

    with tab_dashboard:
        st.subheader("Master Ingest Control Center")
        cloud_files = [f for f in os.listdir(CLOUD_STORAGE_DIR) if f.endswith('.csv') or f.endswith('.xlsx')]
        
        if not cloud_files:
            st.info("Awaiting uploads. No valid data records dropped off in cloud paths yet.")
        else:
            st.success(f"🔔 Online: {len(cloud_files)} supplier data sheets waiting in cloud folders.")
            
            # --- 🛠️ NEW: CENTRAL PATH RECORDS PURGE SECTION ---
            with st.expander("🚨 Purge Inward Directory for New Cycle Actions", expanded=False):
                st.warning("Warning: This operation drops all verified data sets currently loaded in the staging buffer.")
                confirm_purge = st.checkbox("Confirm: I wish to permanently delete all active data files listed below.")
                
                if st.button("Delete All Pending Files", type="primary", use_container_width=True):
                    if confirm_purge:
                        deleted_count = 0
                        for target_f in cloud_files:
                            try:
                                os.remove(os.path.join(CLOUD_STORAGE_DIR, target_f))
                                deleted_count += 1
                            except Exception as ex:
                                st.error(f"Failed to clear resource object '{target_f}': {ex}")
                        st.success(f"Wipe Completed! {deleted_count} files removed. Staging environment is ready for next cycle.")
                        st.rerun()
                    else:
                        st.error("Purge Rejected: Please verify the checkbox rule verification above to process deletion layout triggers.")
            st.markdown("---")

            combined_dfs = []
            for file_name in cloud_files:
                file_path = os.path.join(CLOUD_STORAGE_DIR, file_name)
                try:
                    df = pd.read_excel(file_path) if file_name.endswith('.xlsx') else pd.read_csv(file_path)
                    combined_dfs.append(df)
                except: pass
            
            if combined_dfs:
                master_consolidated_df = pd.concat(combined_dfs, ignore_index=True)
                csv_download_stream = master_consolidated_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Complete Consolidated List (All Suppliers Combined)",
                    data=csv_download_stream,
                    file_name=f"Master_Consolidated_Stock_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    type="primary"
                )
                st.dataframe(master_consolidated_df, use_container_width=True, hide_index=True)

# ==========================================================
# WORKSPACE B: PUBLIC SUPPLIER LANDING PORTAL (WITH TIME-GATE VALIDATION)
# ==========================================================
def render_public_supplier_portal(url_params):
    target_vendor = url_params.get("vendor", "Valued Supplier").upper().replace("-", " ")
    expiry_param = url_params.get("exp", "never")
    
    is_expired = False
    if expiry_param != "never":
        try:
            expiry_date_obj = datetime.strptime(expiry_param, "%Y-%m-%d").date()
            if datetime.now().date() > expiry_date_obj:
                is_expired = True
        except ValueError:
            pass 

    if is_expired:
        st.error("🔒 **Access Window Expired**")
        st.subheader("This upload connection link is no longer valid.")
        st.markdown(f"The inventory submission window for this cycle closed on **{expiry_param}**. Please reach out directly to your Category Manager to request an extension or a refreshed link portal.")
        st.stop() 

    st.title(f"📦 Secure Vendor Inventory Upload Portal")
    st.subheader(f"Account Portfolio Connection: {target_vendor}")
    if expiry_param != "never":
        st.caption(f"⏳ This secure submission window will automatically close on: **{expiry_param}**")
    st.markdown("---")
    
    st.warning("⚠️ **Strict Formatting Rule:** Your file columns must lock to this exact layout sequence:  \n"
               "`Date` | `EAN NUMBER` | `SAP CODE` | `Article Description` | `Item Description` | `Brand` | `Family Name` | `Type` | `Cost` | `Physical Stock`")
    
    supplier_uploaded_file = st.file_uploader("Upload your completed stock ledger:", type=["xlsx", "csv"])

    if supplier_uploaded_file:
        try:
            df = pd.read_excel(supplier_uploaded_file) if supplier_uploaded_file.name.endswith('.xlsx') else pd.read_csv(supplier_uploaded_file)
            uploaded_headers = [str(c).strip() for c in df.columns]
            missing_cols = [col for col in REQUIRED_COLUMNS if col not in uploaded_headers]
            
            if missing_cols:
                st.error("❌ **Upload Denied: Template Structure Mismatch**")
                st.markdown(f"Your file is missing the following required columns: `{missing_cols}`")
            else:
                df = df[REQUIRED_COLUMNS]
                df["Physical Stock"] = df["Physical Stock"].apply(clean_physical_stock)
                df["Supplier Origin Source"] = target_vendor
                df["Upload Timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M")
                
                clean_filename = f"{url_params.get('vendor', 'supplier')}_stock_data.csv"
                df.to_csv(os.path.join(CLOUD_STORAGE_DIR, clean_filename), index=False)
                
                st.balloons()
                st.success(f"✅ Success! Your stock ledger has been verified and synced directly with the Category Manager.")
        except Exception as e:
            st.error(f"Error handling secure validation layout checks: {e}")
