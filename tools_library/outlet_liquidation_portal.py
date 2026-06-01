# tools_library/outlet_liquidation_portal.py
import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# ==========================================================
# CORE STRUCTURAL ATTRIBUTES
# ==========================================================
TOOL_NAME = "Outlet Stock Liquidation Ingest Portal"
TOOL_ICON = "🏷️"

# LOCKED EXACTLY TO YOUR NEW OUTLET TRACKER HEADERS
REQUIRED_COLUMNS = [
    "Store", "Brand", "Item Code", "Article Description", "Type", 
    "RRP with VAT", "Category Approved Price", "Discount Applied", 
    "Display", "Open", "Repaired", "Damaged/Incomple", "Total", "SOLD"
]

STORAGE_DIR = "Data/cloud_outlet_drops"
os.makedirs(STORAGE_DIR, exist_ok=True)

def sanitize_numeric_value(val):
    """Safely cleans currency symbols, commas, and formatting to return valid floats."""
    if pd.isna(val): 
        return 0.0
    val_str = str(val).upper().strip()
    val_str = val_str.replace(",", "").replace("$", "").replace("AED", "").strip()
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def render_ui():
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.markdown("Consolidate store-wise liquidation ledgers with automated row-skip logic and format alignment checkers.")
    st.markdown("---")

    url_params = st.query_params

    if "view" in url_params and url_params["view"] == "portal":
        render_public_uploader_portal(url_params)
    else:
        render_admin_management_dashboard()

# ==========================================================
# WORKSPACE A: CATEGORY TEAM HUB
# ==========================================================
def render_admin_management_dashboard():
    tab_overview, tab_links = st.tabs(["📊 Live Data Ledger", "🔗 Supplier Portal Generators"])

    with tab_links:
        st.subheader("Generate Active Upload Links")
        raw_vendors = st.text_area(
            "Input Store Branches / Vendor Partners (Comma Separated):",
            placeholder="e.g., Dubai Mall Outlet, MOE Store, Abu Dhabi Hub"
        )

        st.markdown("### ⏳ Validity Parameter Window")
        validity_choice = st.selectbox(
            "Link Lifespan Window:",
            ["7 Days Window", "14 Days Window", "30 Days Window", "Permanent Connection"]
        )

        now = datetime.now()
        if "7 Days" in validity_choice:
            expiry_str = (now + timedelta(days=7)).strftime("%Y-%m-%d")
        elif "14 Days" in validity_choice:
            expiry_str = (now + timedelta(days=14)).strftime("%Y-%m-%d")
        elif "30 Days" in validity_choice:
            expiry_str = (now + timedelta(days=30)).strftime("%Y-%m-%d")
        else:
            expiry_str = "never"

        if st.button("Generate Secure Ingestion Matrix Links", type="primary"):
            if raw_vendors:
                vendor_nodes = [v.strip() for v in raw_vendors.split(",") if v.strip()]
                st.success(f"Successfully configured {len(vendor_nodes)} custom outlet portals!")
                
                app_base = "https://category-management-saas-fwpgyjvewktqf4repbby8j.streamlit.app"
                
                link_manifest = []
                for node in vendor_nodes:
                    node_slug = node.lower().replace(" ", "-")
                    access_url = f"{app_base}/?view=portal&vendor={node_slug}&exp={expiry_str}"
                    
                    link_manifest.append({
                        "Reporting Scope Target": node.upper(),
                        "Expiration Date": "No Expiry" if expiry_str == "never" else expiry_str,
                        "Secure Pipeline Link": access_url
                    })
                
                st.dataframe(pd.DataFrame(link_manifest), use_container_width=True, hide_index=True)

    with tab_overview:
        st.subheader("Staging Pipeline Overview")
        pending_files = [f for f in os.listdir(STORAGE_DIR) if f.endswith('.csv') or f.endswith('.xlsx')]

        if not pending_files:
            st.info("Awaiting submissions. No active store-wise sheets uploaded in this window yet.")
        else:
            st.success(f"Detected {len(pending_files)} branch reports securely stored.")
            
            with st.expander("🚨 Reset Pipeline Buffer", expanded=False):
                lock_check = st.checkbox("Confirm permanent delete of active files")
                if st.button("Execute Storage Wipe", type="primary", use_container_width=True):
                    if lock_check:
                        for target_f in pending_files:
                            os.remove(os.path.join(STORAGE_DIR, target_f))
                        st.success("Staging engine completely refreshed!")
                        st.rerun()
            
            st.markdown("---")

            consolidated_runs = []
            for document in pending_files:
                doc_path = os.path.join(STORAGE_DIR, document)
                try:
                    loaded_df = pd.read_excel(doc_path) if document.endswith('.xlsx') else pd.read_csv(doc_path)
                    consolidated_runs.append(loaded_df)
                except:
                    pass

            if consolidated_runs:
                master_outlet_df = pd.concat(consolidated_runs, ignore_index=True)
                csv_bytes = master_outlet_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Master Consolidated Liquidation Data (.csv)",
                    data=csv_bytes,
                    file_name=f"Master_Consolidated_Outlet_Stock_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    type="primary"
                )
                st.dataframe(master_outlet_df, use_container_width=True, hide_index=True)

# ==========================================================
# WORKSPACE B: DISTRIBUTED PUBLIC UPLOADER PORTAL
# ==========================================================
def render_public_uploader_portal(url_params):
    scope_target = url_params.get("vendor", "Target Branch").upper().replace("-", " ")
    expiry_deadline = url_params.get("exp", "never")

    gate_expired = False
    if expiry_deadline != "never":
        try:
            deadline_date = datetime.strptime(expiry_deadline, "%Y-%m-%d").date()
            if datetime.now().date() > deadline_date:
                gate_expired = True
        except ValueError:
            pass

    if gate_expired:
        st.error("🔒 **Submission Link Terminated**")
        st.stop()

    st.title("🏷️ Secure Stock Liquidation Portal")
    st.subheader(f"Branch Assignment: {scope_target}")
    st.markdown("---")

    uploaded_tracker = st.file_uploader("Drop your Liquidation Excel or CSV sheet here:", type=["xlsx", "csv"])

    if uploaded_tracker:
        try:
            # 1. Lookahead Scanner: bypass empty spacing or banner headers automatically
            if uploaded_tracker.name.endswith('.xlsx'):
                raw_df = pd.read_excel(uploaded_tracker, header=None)
                header_row_idx = 0
                for idx, row in raw_df.iterrows():
                    row_values = [str(val).strip() for val in row.values if pd.notna(val)]
                    if "Item Code" in row_values or "Store" in row_values:
                        header_row_idx = idx
                        break
                df = pd.read_excel(uploaded_tracker, skiprows=header_row_idx)
            else:
                df = pd.read_csv(uploaded_tracker)
            
            # 2. Drop layout artifacts
            df = df.loc[:, ~df.columns.astype(str).str.contains('^Unnamed')]

            # 3. Power wash headers: clear text wrapping and ghost spaces
            df.columns = [str(col).replace('\n', ' ').strip() for col in df.columns]
            df.columns = [" ".join(str(col).split()) for col in df.columns]

            # 4. Schema Comparison Check
            clean_required = [req.strip() for req in REQUIRED_COLUMNS]
            clean_uploaded = [c.strip() for c in df.columns]
            absent_columns = [req for req in clean_required if req not in clean_uploaded]

            if absent_columns:
                st.error("❌ **Upload Aborted: Column Alignment Failure**")
                st.write("The incoming tracker could not be matched. Missing columns:")
                st.code(f"{absent_columns}", language="json")
            else:
                # 5. Lock structure down 
                mapping_dict = {c.strip(): c for c in df.columns}
                target_columns_in_file = [mapping_dict[req.strip()] for req in REQUIRED_COLUMNS]
                
                processed_df = df[target_columns_in_file].copy()
                processed_df.columns = REQUIRED_COLUMNS  

                # 6. Clean currency parameters safely to minimize math breakdowns
                monetary_fields = ["RRP with VAT", "Category Approved Price", "Discount Applied", "Display", "Open", "Repaired", "Damaged/Incomple", "Total", "SOLD"]
                for field in monetary_fields:
                    if field in processed_df.columns:
                        processed_df[field] = processed_df[field].apply(sanitize_numeric_value)

                # Add upload source metadata
                processed_df["Ingest Source Channel"] = scope_target
                processed_df["Data Submission Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # Serialize down safely to storage matrix
                save_filename = f"outlet_drop_{url_params.get('vendor', 'unnamed')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                processed_df.to_csv(os.path.join(STORAGE_DIR, save_filename), index=False)

                st.balloons()
                st.success(f"🎉 Ledger uploaded! Your metrics for '{scope_target}' have been validated and synced perfectly.")
                
        except Exception as err:
            st.error(f"A file error occurred during validation tracking: {err}")