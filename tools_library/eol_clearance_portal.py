# tools_library/eol_clearance_portal.py
import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta

# ==========================================================
# CORE STRUCTURAL ATTRIBUTES
# ==========================================================
TOOL_NAME = "EOL Clearance Ingest Portal"
TOOL_ICON = "📉"

# Exactly matching your uploaded template schema
REQUIRED_COLUMNS = [
    "S.No", "Supplier Code", "Supplier Name", "SAP Code", "Description",
    "Category L1", "Category L3", "Brand", "Model", "Stock", "Cost",
    "Total Cost", "RRP", "Total RRP", "Clearance Price", "Total Clearance Value",
    "Action Plan", "Target Value"
]

STORAGE_DIR = "Data/cloud_eol_drops"
os.makedirs(STORAGE_DIR, exist_ok=True)

def sanitize_numeric_value(val):
    """Safely handles commas, text wrappers, and currencies to return clean floats."""
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
    st.markdown("Consolidate End-of-Life (EOL) Clearance reports with absolute format structural enforcement.")
    st.markdown("---")

    # Access current routing parameters
    url_params = st.query_params

    if "view" in url_params and url_params["view"] == "portal":
        render_public_uploader_portal(url_params)
    else:
        render_admin_management_dashboard()

# ==========================================================
# WORKSPACE A: INTERNAL MANAGEMENT DASHBOARD
# ==========================================================
def render_admin_management_dashboard():
    tab_overview, tab_links = st.tabs(["📊 Consolidated Data View", "🔗 Link Lifecycle Control"])

    with tab_links:
        st.subheader("Generate EOL Submission Connections")
        st.markdown("Produce tracking parameters to securely distribute to your field teams or vendors.")
        
        # Vendor array distribution intake
        raw_vendors = st.text_area(
            "Input Target Supplier / Team Identifiers (Comma Separated):",
            placeholder="e.g., Apple EOL, Samsung Audio, Sony Video",
            help="Separate each group tracking name with a comma."
        )

        st.markdown("### ⏳ Validity Parameter Window")
        validity_choice = st.selectbox(
            "Link Lifespan Window:",
            ["7 Days Window", "14 Days Window", "30 Days Window", "Permanent Connection"]
        )

        # Handle chronological drift settings
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
                st.success(f"Successfully configured {len(vendor_nodes)} EOL reporting portals!")
                
                # Base dynamic web context routing link
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
            else:
                st.warning("Please input at least one target scope identifier name.")

    with tab_overview:
        st.subheader("Staging Pipeline Overview")
        pending_files = [f for f in os.listdir(STORAGE_DIR) if f.endswith('.csv') or f.endswith('.xlsx')]

        if not pending_files:
            st.info("No active EOL reports have been dropped into the storage layer for this current cycle.")
        else:
            st.success(f"Detected {len(pending_files)} upstream EOL trackers submitted.")
            
            # Master Deletion Control Node
            with st.expander("🚨 Flush Database Storage Buffer", expanded=False):
                st.write("Permanently clears all active staging files to reset the collection cycle.")
                lock_check = st.checkbox("Confirm permanent delete action sequence")
                if st.button("Execute Storage Wipe", type="primary", use_container_width=True):
                    if lock_check:
                        for target_f in pending_files:
                            os.remove(os.path.join(STORAGE_DIR, target_f))
                        st.success("Staging buffer completely initialized!")
                        st.rerun()
            
            st.markdown("---")

            # Combine all incoming streams safely
            consolidated_runs = []
            for document in pending_files:
                doc_path = os.path.join(STORAGE_DIR, document)
                try:
                    loaded_df = pd.read_excel(doc_path) if document.endswith('.xlsx') else pd.read_csv(doc_path)
                    consolidated_runs.append(loaded_df)
                except:
                    pass

            if consolidated_runs:
                master_eol_df = pd.concat(consolidated_runs, ignore_index=True)
                
                # Download Stream Provisioning
                csv_bytes = master_eol_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Master Consolidated EOL Clearance Data (.csv)",
                    data=csv_bytes,
                    file_name=f"Master_Consolidated_EOL_Clearance_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True,
                    type="primary"
                )
                st.dataframe(master_eol_df, use_container_width=True, hide_index=True)

# ==========================================================
# WORKSPACE B: VENDOR/TEAM PUBLIC LANDING PORTAL
# ==========================================================
def render_public_uploader_portal(url_params):
    scope_target = url_params.get("vendor", "Target Scope").upper().replace("-", " ")
    expiry_deadline = url_params.get("exp", "never")

    # Structural Time Gate Evaluation
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
        st.subheader("The availability period allocated for this data drop collection loop has expired.")
        st.stop()

    st.title("📉 Secure EOL Clearance Tracker Portal")
    st.subheader(f"Portfolio Assignment: {scope_target}")
    if expiry_deadline != "never":
        st.caption(f"⏳ Submission timeline closing window date: **{expiry_deadline}**")
    st.markdown("---")

    # Instructions box mapping schema expectations directly
    st.info(
        "💡 **Data Standard Requirement:** Your column headers must accurately align to the following naming strings:\n\n"
        "`S.No` | `Supplier Code` | `Supplier Name` | `SAP Code` | `Description` | `Category L1` | "
        "`Category L3` | `Brand` | `Model` | `Stock` | `Cost` | `Total Cost` | `RRP` | `Total RRP` | "
        "`Clearance Price` | `Total Clearance Value` | `Action Plan` | `Target Value`"
    )

    uploaded_tracker = st.file_uploader("Drop your EOL Tracker Excel or CSV file here:", type=["xlsx", "csv"])

    if uploaded_tracker:
        try:
            # Read format safely
            raw_data_df = pd.read_excel(uploaded_tracker) if uploaded_tracker.name.endswith('.xlsx') else pd.read_csv(uploaded_tracker)
            
            # Normalize user-supplied string header spaces
            cleansed_headers = [str(col).strip() for col in raw_data_df.columns]
            raw_data_df.columns = cleansed_headers

            # Cross check missing columns
            absent_columns = [expected for expected in REQUIRED_COLUMNS if expected not in cleansed_headers]

            if absent_columns:
                st.error("❌ **Upload Aborted: Alignment Structure Invalidation**")
                st.write("The incoming tracker could not be committed because the following headers are missing:")
                st.code(f"{absent_columns}", language="json")
            else:
                # Isolate target structure frame
                processed_df = raw_data_df[REQUIRED_COLUMNS].copy()

                # Clean numeric and financial parameters safely to minimize math breakdowns
                monetary_fields = ["Stock", "Cost", "Total Cost", "RRP", "Total RRP", "Clearance Price", "Total Clearance Value", "Target Value"]
                for field in monetary_fields:
                    processed_df[field] = processed_df[field].apply(sanitize_numeric_value)

                # Add orchestration metadata attributes
                processed_df["Ingest Source Channel"] = scope_target
                processed_df["Data Submission Date"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                # Serialize down safely to storage matrix target path
                save_filename = f"eol_drop_{url_params.get('vendor', 'unnamed')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                processed_df.to_csv(os.path.join(STORAGE_DIR, save_filename), index=False)

                st.balloons()
                st.success(f"🎉 Thank you! The EOL Clearance ledger for '{scope_target}' has been processed and synced successfully.")
                
        except Exception as err:
            st.error(f"A physical validation failure stopped interpretation: {err}")