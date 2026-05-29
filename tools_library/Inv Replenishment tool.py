import pandas as pd
import numpy as np
import io
import streamlit as st

# ==========================================================
# 🏷️ MODULE METADATA & ENTRY REGISTRATION FOR STREAMLIT SaaS
# ==========================================================
TOOL_NAME = "ERP Strategic Inventory Engine"
TOOL_ICON = "📈"

def clean_headers(df):
    """Strips whitespaces and standardizes column cases for safe dictionary lookups."""
    df.columns = [str(c).strip() for c in df.columns]
    return df

def render_ui():
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.markdown("### New ERP Cross-Functional Inventory Alignment Module")
    st.markdown(
        "Upload files from your new ERP system. The pipeline dynamically maps records across "
        "all sheets using **Article** as the unique key."
    )
    st.markdown("---")

    # --- 1. CONFIGURATION PARAMETERS ---
    DAYS_IN_PERIOD = 90  
    TARGET_DAYS = 20     
    SAFETY_STOCK = 2     

    # --- 2. MULTI-STREAM FILE UPLOAD PLATFORM ---
    col1, col2 = st.columns(2)
    with col1:
        soh_upload = st.file_uploader("1. Current Stock On Hand Sheet (SOH)", type=["csv", "xlsx"])
        sales_upload = st.file_uploader("2. ERP JFM Sales Matrix File", type=["csv", "xlsx"])
    with col2:
        rtv_upload = st.file_uploader("3. Return to Vendor (RTV) Log Sheet", type=["csv", "xlsx"])
        po_upload = st.file_uploader("4. Open Purchase Orders (PO) Status Log", type=["csv", "xlsx"])

    if soh_upload and sales_upload and rtv_upload and po_upload:
        st.markdown("---")
        if st.button("Execute Cross-Functional Allocation Mapping Loop", type="primary", use_container_width=True):
            with st.spinner("Processing structural matrix intersections across your new ERP schemas..."):
                try:
                    # Load Data Streams
                    df_soh = pd.read_csv(soh_upload) if soh_upload.name.endswith('.csv') else pd.read_excel(soh_upload)
                    df_sales = pd.read_csv(sales_upload) if sales_upload.name.endswith('.csv') else pd.read_excel(sales_upload)
                    df_rtv = pd.read_csv(rtv_upload) if rtv_upload.name.endswith('.csv') else pd.read_excel(rtv_upload)
                    df_po = pd.read_csv(po_upload) if po_upload.name.endswith('.csv') else pd.read_excel(po_upload)

                    # Clean Dataframes
                    df_soh = clean_headers(df_soh)
                    df_sales = clean_headers(df_sales)
                    df_rtv = clean_headers(df_rtv)
                    df_po = clean_headers(df_po)

                    # --- 3. DYNAMIC ERP COLUMN TRANSFORMATION PIPELINE ---
                    # Map the old SAP CODE to your new ERP 'Article' identifier across datasets
                    for df_obj in [df_soh, df_sales, df_rtv, df_po]:
                        if 'Article' in df_obj.columns:
                            df_obj['Article'] = df_obj['Article'].astype(str).str.strip()
                        elif 'SAP CODE' in df_obj.columns:
                            df_obj.rename(columns={'SAP CODE': 'Article'}, inplace=True)
                            df_obj['Article'] = df_obj['Article'].astype(str).str.strip()
                        else:
                            st.error("🚨 Missing mandatory relational key column: **'Article'** or **'SAP CODE'** omitted.")
                            return

                    # Standardize Store naming constraints
                    for df_obj in [df_soh, df_sales, df_rtv, df_po]:
                        if 'Store' in df_obj.columns:
                            df_obj['Store'] = df_obj['Store'].astype(str).str.strip()
                        elif 'Store Name' in df_obj.columns:
                            df_obj.rename(columns={'Store Name': 'Store'}, inplace=True)
                            df_obj['Store'] = df_obj['Store'].astype(str).str.strip()

                    # Dynamic Target Column Recognition
                    sales_col = next((col for col in df_sales.columns if any(k in col for k in ['Total', 'Sales', 'Qty', 'QTY'])), None)
                    soh_col = next((col for col in df_soh.columns if any(k in col for k in ['SOH', 'Stock', 'On Hand', 'OnHandQuantity'])), None)
                    rtv_col = next((col for col in df_rtv.columns if any(k in col for k in ['Return', 'RTV', 'Rtn'])), None)
                    po_col = next((col for col in df_po.columns if any(k in col for k in ['PO', 'Order', 'Open', 'OpenQty'])), None)

                    # Core Fallbacks for Unmatched System Keys
                    if not sales_col: sales_col = df_sales.columns[-1]
                    if not soh_col: soh_col = 'SOH' if 'SOH' in df_soh.columns else df_soh.columns[-1]
                    if not rtv_col: rtv_col = df_rtv.columns[-1]
                    if not po_col: po_col = df_po.columns[-1]

                    # --- 4. DATA FEDERATION GRID BUILDER ---
                    df_sales_clean = df_sales.groupby(['Article', 'Store'])[sales_col].sum().reset_index()
                    df_rtv_clean = df_rtv.groupby(['Article', 'Store'])[rtv_col].sum().reset_index()
                    df_po_clean = df_po.groupby(['Article', 'Store'])[po_col].sum().reset_index()

                    # Intersect all 4 logs into a unified processing matrix
                    df = pd.merge(df_soh, df_sales_clean, on=['Article', 'Store'], how='outer')
                    df = pd.merge(df, df_rtv_clean, on=['Article', 'Store'], how='left')
                    df = pd.merge(df, df_po_clean, on=['Article', 'Store'], how='left')
                    df.fillna(0, inplace=True)

                    # Dynamic Metadata Synchronization Loop
                    meta_candidates = ['Product Name', 'ProductDesc', 'Mdse Catgry Desc', 'Brand', 'BRAND', 'Location', 'Model', 'Description']
                    meta_cols = [c for c in meta_candidates if c in df.columns]
                    
                    for col in meta_cols:
                        df[col] = df.groupby('Article')[col].transform(lambda x: x.replace(0, np.nan).ffill().bfill()).fillna("N/A")

                    # --- 5. EXECUTE ALLOCATION MATRIX CALCULATIONS ---
                    df['Daily_Sales'] = df[sales_col] / DAYS_IN_PERIOD
                    df['Avg Sales 20D'] = (df['Daily_Sales'] * TARGET_DAYS).round(2)
                    
                    # Target Demand Capacity
                    df['Capacity'] = (df['Avg Sales 20D'] - df[soh_col]).clip(lower=0).apply(np.ceil).astype(int)
                    
                    # Available Senders Surplus
                    df['Min_Keep'] = (df['Daily_Sales'] * TARGET_DAYS).apply(np.ceil).apply(lambda x: max(SAFETY_STOCK, x))
                    df['Surplus'] = (df[soh_col] - df['Min_Keep']).clip(lower=0).astype(int)

                    # --- 6. WATERFALL REBALANCING LOOPS ---
                    transfer_results = []
                    current_capacity = df.set_index(['Article', 'Store'])['Capacity'].to_dict()

                    for article in df['Article'].unique():
                        senders = df[(df['Article'] == article) & (df['Surplus'] > 0)].copy()
                        receivers = df[df['Article'] == article].sort_values(sales_col, ascending=False)
                        
                        for _, s in senders.iterrows():
                            rem_surplus = s['Surplus']
                            for _, r in receivers.iterrows():
                                if rem_surplus <= 0: break
                                if s['Store'] == r['Store']: continue
                                
                                cap = current_capacity.get((article, r['Store']), 0)
                                if cap > 0:
                                    qty = min(rem_surplus, cap)
                                    if qty > 0:
                                        transfer_results.append({
                                            'Article / SAP CODE': article,
                                            'Brand': s.get('Brand', s.get('BRAND', 'N/A')),
                                            'Product Name': s.get('Product Name', 'N/A'),
                                            'ProductDesc / Model': s.get('ProductDesc', s.get('Description', 'N/A')),
                                            'Source Store': s['Store'],
                                            'Location / Bin': s.get('Location', 'N/A'),
                                            'Mdse Catgry Desc': s.get('Mdse Catgry Desc', 'N/A'),
                                            'Source Current SOH': int(s[soh_col]),
                                            'Source Avg Sales 20D': s['Avg Sales 20D'],
                                            'Transfer Destination': r['Store'],
                                            'Transfer QTY Required': int(qty),
                                            'Receiver Avg Sales 20D': r['Avg Sales 20D']
                                        })
                                        rem_surplus -= qty
                                        current_capacity[(article, r['Store'])] -= qty

                    # --- 7. EXPORT SPREADSHEETS ---
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                        # Sheet 1: Core Master Sheet
                        df.to_excel(writer, sheet_name='Unified Inventory Grid', index=False)
                        
                        # Sheet 2: Stock Transfer Outcome (Matching User's target screenshot formatting)
                        if transfer_results:
                            pd.DataFrame(transfer_results).to_excel(writer, sheet_name='Stock Transfer Plan', index=False)
                        else:
                            pd.DataFrame([{"System Notice": "No matching inventory records found for reallocation."}]).to_excel(writer, sheet_name='Stock Transfer Plan', index=False)

                    excel_data = excel_buffer.getvalue()
                    st.balloons()
                    st.success("🎯 Multi-Stream ERP Processing Loop Finalized!")

                    # Metrics display panels
                    sc1, sc2, sc3 = st.columns(3)
                    sc1.metric("Unique Articles Processed", len(df['Article'].unique()))
                    sc2.metric("Scheduled Store Actions", len(transfer_results))
                    sc3.metric("System Operational SOH Evaluated", int(df[soh_col].sum()))

                    # File Download Endpoint
                    st.download_button(
                        label="📥 Download Unified ERP Strategic Workbook (xlsx)",
                        data=excel_data,
                        file_name="ERP_Final_Strategic_Inventory_Suite.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as ex:
                    st.error(f"🚨 Integration Engine Error: {str(ex)}")