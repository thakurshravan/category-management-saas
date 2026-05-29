# tools_library/brand_claims.py
import streamlit as st
import pandas as pd

# --- Master Connection Keys ---
TOOL_NAME = "Brand Promo Claims Tool"
TOOL_ICON = "🎯"

def render_ui():
    st.title("Brand Promotion Claims Calculator")
    st.subheader("Strict Date Window Matching ('From' & 'To') with Master Email Lookup via 'Article Name'")
    
    # 3-Column layout for clean file uploading interface
    col1, col2, col3 = st.columns(3)
    with col1:
        claim_sales_file = st.file_uploader("1. Upload Master Sales File", type=["xlsx", "csv"], key="claim_sales")
    with col2:
        claim_promo_file = st.file_uploader("2. Upload 01 File for ACC WTC Brand", type=["xlsx", "csv"], key="claim_promo")
    with col3:
        claim_email_master = st.file_uploader("3. Upload Email Master All File", type=["xlsx", "csv"], key="claim_email_master")
        
    st.markdown("---")
    
    if st.button("🚀 Calculate Total Qty Sales & Generate Claims", type="primary", key="claim_run"):
        if claim_sales_file and claim_promo_file and claim_email_master:
            with st.spinner("Processing files, applying inclusive date validation, and compiling the full report..."):
                try:
                    # 1. Dynamically read all file streams safely (.csv or .xlsx)
                    df_sales = pd.read_csv(claim_sales_file) if claim_sales_file.name.endswith('.csv') else pd.read_excel(claim_sales_file)
                    df_promo = pd.read_csv(claim_promo_file) if claim_promo_file.name.endswith('.csv') else pd.read_excel(claim_promo_file)
                    df_emails = pd.read_csv(claim_email_master) if claim_email_master.name.endswith('.csv') else pd.read_excel(claim_email_master)
                    
                    # 2. Harmonize column name headers to ensure 'Article Name' is strictly uniform
                    if "Code" in df_promo.columns and "Article Name" not in df_promo.columns:
                        df_promo = df_promo.rename(columns={"Code": "Article Name"})
                    
                    # Drop any pre-existing email references from the promo file to guarantee we ONLY use Email_Master - All
                    if "Email id" in df_promo.columns:
                        df_promo = df_promo.drop(columns=["Email id"])
                    if "Email" in df_promo.columns:
                        df_promo = df_promo.drop(columns=["Email"])
                        
                    # 3. Convert date strings to strict datetime objects for window filtering
                    df_sales["Invoice Created On"] = pd.to_datetime(df_sales["Invoice Created On"], errors="coerce")
                    df_promo["From"] = pd.to_datetime(df_promo["From"], errors="coerce")
                    df_promo["To"] = pd.to_datetime(df_promo["To"], errors="coerce")
                    
                    # 4. Clean and normalize string matching keys (strip whitespace, force uppercase)
                    df_sales["Article Name"] = df_sales["Article Name"].astype(str).str.strip().str.upper()
                    df_promo["Article Name"] = df_promo["Article Name"].astype(str).str.strip().str.upper()
                    df_emails["Article Name"] = df_emails["Article Name"].astype(str).str.strip().str.upper()
                    
                    # 5. Handle numbers safely to prevent broken aggregation calculations
                    df_sales["Invoice Quantity"] = pd.to_numeric(df_sales["Invoice Quantity"], errors="coerce").fillna(0)
                    df_promo["GV"] = pd.to_numeric(df_promo["GV"], errors="coerce").fillna(0)
                    
                    # 6. Step 1 Join: Link Sales directly to the Promotion parameters using the common 'Article Name'
                    merged = pd.merge(df_sales, df_promo, on="Article Name", how="inner")
                    
                    # CRITICAL FIX: Resolve duplicate 'Brand' columns resulting from the merge step (Brand_x vs Brand_y)
                    if "Brand_y" in merged.columns:
                        merged = merged.rename(columns={"Brand_y": "Brand"})
                    elif "Brand_x" in merged.columns:
                        merged = merged.rename(columns={"Brand_x": "Brand"})
                    
                    # 7. Inclusive day-to-day validation (includes exact same starting/ending dates)
                    valid_claims = merged[
                        (merged["Invoice Created On"].dt.date >= merged["From"].dt.date) & 
                        (merged["Invoice Created On"].dt.date <= merged["To"].dt.date)
                    ]
                    
                    if valid_claims.empty:
                        st.warning("⚠️ No records found where the Sales dates fell inside the promotion 'From' and 'To' date windows.")
                        return
                    
                    # 8. Comprehensive grouping structure to ensure Comments, Brand, and Promo Codes stay intact
                    group_cols = [
                        "Supporting Ref", "Fin No:", "Deal Sheet No", "Deal Mail Date", 
                        "CAT", "Brand", "Article Name", "From", "To", "GV",
                        "Comments", "Remarks (Promo Code)"
                    ]
                    # Verify column safety against what is present inside the filtered data
                    available_cols = [col for col in group_cols if col in valid_claims.columns]
                    
                    # dropna=False forces pandas to keep rows even if comments or promo codes are blank!
                    final_summary = valid_claims.groupby(available_cols, dropna=False).agg(
                        Total_Qty_Sold=("Invoice Quantity", "sum")
                    ).reset_index()
                    
                    # 9. Compute claims total amounts
                    final_summary["Total Claim Amount"] = final_summary["Total_Qty_Sold"] * final_summary["GV"]
                    
                    # 10. Step 2 Join: Strictly pull contact emails from Email_Master - All using common 'Article Name'
                    df_emails_clean = df_emails[["Article Name", "Email"]].drop_duplicates(subset=["Article Name"])
                    final_summary = pd.merge(final_summary, df_emails_clean, on="Article Name", how="left")
                    
                    # Rename the master email column to 'Email id' to match your exact sheet layout format
                    if "Email" in final_summary.columns:
                        final_summary = final_summary.rename(columns={"Email": "Email id"})
                        final_summary["Email id"] = final_summary["Email id"].fillna("missing@contact.com")
                    else:
                        final_summary["Email id"] = "missing@contact.com"
                    
                    # Clean visual positioning: insert 'Email id' column right after Brand column
                    if "Email id" in final_summary.columns:
                        email_col = final_summary.pop("Email id")
                        if "Brand" in final_summary.columns:
                            brand_idx = final_summary.columns.get_loc("Brand") + 1
                            final_summary.insert(brand_idx, "Email id", email_col)
                        else:
                            final_summary.insert(6, "Email id", email_col)
                        
                    # 11. Output final verified summary matrix to the dashboard screen
                    st.success(f"🎉 Success! Compiled a full report of {len(final_summary)} matching promotion claim lines.")
                    st.dataframe(final_summary, use_container_width=True)
                    
                    # 12. Provide file download action button
                    csv_bytes = final_summary.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="📥 Download Full Mapped Claims Report (.csv)",
                        data=csv_bytes,
                        file_name="Full_Brand_Promo_Claims_Report.csv",
                        mime="text/csv"
                    )
                    
                    # 13. Context-Aware Automated Email Preview Block
                    st.markdown("---")
                    st.markdown("### 📧 Generated Email Body Template")
                    
                    target_recipient = "partner@brandcontact.com"
                    if "Email id" in final_summary.columns and not final_summary["Email id"].dropna().empty:
                        target_recipient = final_summary["Email id"].dropna().iloc[0]
                        
                    st.markdown(f"**Target Destination To (Pulled strictly from Email Master):** `{target_recipient}`")
                    
                    email_msg = (
                        "Dear Value Partner,\n\n"
                        "Please find the attached claim for the promotions mentioned above. "
                        "Kindly review and verify the documents within five working days.\n\n"
                        "Dear @Ashwitha Shetty, once the claim has been verified (or after the five working days "
                        "window has passed), Kindly share the corresponding debit note.\n\n"
                        "Best regards,\nCategory Manager"
                    )
                    st.code(email_msg, language="markdown")
                    
                except Exception as ex:
                    st.error(f"An unexpected data processing error occurred: {ex}")
        else:
            st.warning("Please verify that all three required files are uploaded completely.")