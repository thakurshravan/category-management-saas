import streamlit as st
import requests
import pandas as pd
import plotly.express as px

# ==========================================================
# 🏷️ MODULE METADATA & ENTRY REGISTRATION FOR STREAMLIT
# ==========================================================
TOOL_NAME = "Sales Hub & Performance Tracker"
TOOL_ICON = "⚡"

SB_URL = st.secrets.get("SUPABASE_URL", "")
SB_KEY = st.secrets.get("SUPABASE_KEY", "")

def query_ledger(action, payload=None, target_id=None):
    headers = {
        "apiKey": SB_KEY,
        "Authorization": f"Bearer {SB_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    url = f"{SB_URL}/rest/v1/showroom_performance_ledger"
    try:
        if action == "INSERT":
            res = requests.post(url, json=payload, headers=headers)
        elif action == "SELECT":
            res = requests.get(f"{url}?order=created_at.desc", headers=headers)
        elif action == "UPDATE":
            res = requests.patch(f"{url}?id=eq.{target_id}", json=payload, headers=headers)
        return res.json() if res.status_code in [200, 201] else []
    except Exception:
        return []

def render_ui():
    BASE_COMMISSION_PERCENT = 0.02
    STRETCH_GOAL_THRESHOLD = 15000
    STRETCH_GOAL_BONUS = 500
    GOOGLE_REVIEW_BONUS = 25

    st.title(f"{TOOL_ICON} Showroom Sales & Operations Hub")
    st.markdown("---")

    hub_tabs = st.tabs(["📊 Performance KPIs", "📝 Inquiry Capture", "💰 My Earnings", "⭐ Review Stream", "👑 Admin Board"])

    raw_ledger = query_ledger("SELECT")
    df_master = pd.DataFrame(raw_ledger) if raw_ledger else pd.DataFrame()

    # ==========================================================
    # 1. FIXED PARAMETERS INTAKE FORM (Strict Numeric SAP Guard)
    # ==========================================================
    with hub_tabs[1]:
        st.subheader("Showroom Demand Intake Form")
        with st.form("showroom_intake_form", clear_on_submit=True):
            
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                f_store = st.selectbox("STORE NAME", [
                    "RALN", "RAMM", "RAWM", "RBAM", "RDCC", "RDHL", "RDLM", "RDMU", 
                    "RGMM", "RHMD", "RJH0", "RMCC", "RMOE", "RRAK", "RSCC", "RZCC"
                ])
                f_cust_name = st.text_input("Customer Name", placeholder="Enter customer full name")
                f_cust_phone = st.text_input("Customer Contact Number")
                f_cust_email = st.text_input("Customer Email Address", value="N/A")
                f_model = st.text_input("PRODUCT MODEL DETAILS", placeholder="e.g., OLED65CXPTA")
                
                # 🔢 CRITICAL: FORCE NUMERIC VALUE FOR ARTICLE CODE
                f_article = st.number_input("Article Code (SAP Code) - NUMERIC ONLY", min_value=1, step=1, value=1000001, help="Letters and symbols are restricted.")
            
            with col_f2:
                f_value = st.number_input("PRODUCT VALUE (Numeric AED)", min_value=0.0, step=100.0)
                f_brand = st.text_input("BRAND", placeholder="e.g., LG, Sony, Samsung")
                
                f_category = st.selectbox("Category", [
                    "ACC", "AV", "GAM", "IMG", "IT", "JSP", "LA", "MM", 
                    "OAD", "PG", "S&N", "SDA", "TEL", "WTC", "OTH"
                ])
                
                f_cx = st.selectbox("CX RESPONSE", [
                    "FUTURE BUYER", "JUST BROWSING", "NOT PART OF STORE ASSORTMENT", 
                    "NOT PART OF UAE STORES' ASSORTMENT", "OTH", "PRICE ISSUE", 
                    "SALES CLOSED", "STOCK ISSUE"
                ])
                
                f_remarks = st.text_area("ANY REMARKS", placeholder="Add workflow specifics...")
                
                f_emp_code = st.number_input("EMP Code (Numeric)", min_value=1, step=1, value=1001)
                f_emp_name = st.text_input("Emp Name", value=st.session_state.get("user_email", "Staff Representative"))

            if st.form_submit_button("Lock Lead to Profile Grid", type="primary", use_container_width=True):
                if f_cust_phone and f_cust_name and f_article:
                    entry_payload = {
                        "store_name": f_store,
                        "customer_name": f_cust_name,
                        "customer_phone": f_cust_phone,
                        "customer_email": f_cust_email,
                        "product_model_details": f_model,
                        "article_code": str(int(f_article)), # Formatted cleanly as a numeric string variant payload
                        "product_value": f_value,
                        "brand": f_brand,
                        "category": f_category,
                        "cx_response": f_cx,
                        "any_remarks": f_remarks,
                        "emp_code": int(f_emp_code),
                        "emp_name": f_emp_name,
                        "salesman_email": st.session_state.get("user_email", "staff@company.com")
                    }
                    if query_ledger("INSERT", payload=entry_payload):
                        st.success("🎉 Showroom entry verified and recorded successfully!")
                        st.rerun()
                    else:
                        st.error("Network synchronization timeout.")
                else:
                    st.warning("Customer Name, Contact Number, and Article Code are mandatory fields.")

    if not df_master.empty and 'salesman_email' in df_master.columns:
        df_personal = df_master[df_master['salesman_email'] == st.session_state.get("user_email", "")].copy()
    else:
        df_personal = pd.DataFrame()

    # ==========================================================
    # 2. INCENTIVE TRACKING ENGINE
    # ==========================================================
    with hub_tabs[2]:
        st.subheader("My Commission & Incentive Trackers")
        if not df_personal.empty:
            df_won = df_personal[df_personal['cx_response'] == 'SALES CLOSED']
            won_volume = df_won['product_value'].sum()
            
            base_payout = won_volume * BASE_COMMISSION_PERCENT
            five_star_reviews = df_personal[df_personal['google_rating'] == 5]
            review_commissions = len(five_star_reviews) * GOOGLE_REVIEW_BONUS
            unlocked_stretch_bonus = STRETCH_GOAL_BONUS if won_volume >= STRETCH_GOAL_THRESHOLD else 0.0
            total_accrued_payout = base_payout + review_commissions + unlocked_stretch_bonus

            p_col1, p_col2, p_col3 = st.columns(3)
            p_col1.metric("Closed Volume (AED)", f"{won_volume:,.2f}")
            p_col2.metric("Base Commission Accrued", f"{base_payout:,.2f}")
            p_col3.metric("Review Cash Rewards", f"{review_commissions:,.2f}")
            
            st.metric("Estimated Monthly Paycheck", f"{total_accrued_payout:,.2f} AED", 
                      delta=f"{STRETCH_GOAL_BONUS} AED Stretch Target Met!" if unlocked_stretch_bonus > 0 else "Target Lock Pending")
        else:
            st.info("Log your initial showroom inquiry to synchronize compensation scorecards.")

    # ==========================================================
    # 3. SALESMAN PERFORMANCE TRACKING DASHBOARD
    # ==========================================================
    with hub_tabs[0]:
        st.subheader("Personal Conversion KPIs")
        if not df_personal.empty:
            kpi_c1, kpi_c2, kpi_c3 = st.columns(3)
            total_leads_logged = len(df_personal)
            successful_closes = len(df_personal[df_personal['cx_response'] == 'SALES CLOSED'])
            conversion_rate = (successful_closes / total_leads_logged) * 100 if total_leads_logged > 0 else 0.0
            
            kpi_c1.metric("Total Inquiries Handled", total_leads_logged)
            kpi_c2.metric("Successful Deals Closed", successful_closes)
            kpi_c3.metric("Deal Conversion Rate", f"{conversion_rate:.1f}%")
            
            st.markdown("### CX Response Interactions Mix")
            mix_df = df_personal.groupby('cx_response').size().reset_index(name='Count')
            fig_pie = px.pie(mix_df, values='Count', names='cx_response', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No conversion metrics mapped to your profile yet.")

    # ==========================================================
    # 4. GOOGLE REPUTATION FEEDBACK LINK
    # ==========================================================
    with hub_tabs[3]:
        st.subheader("Showroom Reputation Pipeline")
        st.markdown(
            f'<div style="background-color:#fff; padding:20px; text-align:center; border-radius:10px; width:220px; margin:auto;">'
            f'<img src="https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=https://search.google.com/local/writereview" />'
            f'<br/><b style="color:#000;">Scan to Leave Review</b></div>', 
            unsafe_allow_url=True
        )

    # ==========================================================
    # 5. GLOBAL ADMIN BROADCAST CONTROL BOARD
    # ==========================================================
    with hub_tabs[4]:
        st.subheader("Global Operations Dashboard")
        if st.session_state.get("user_email") in ["thakurshravan@hotmail.com", "thakshravan@gmail.com"]:
            if not df_master.empty:
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    df_master.to_excel(writer, sheet_name='Master Inquiries Ledger', index=False)
                
                st.download_button(
                    label="📥 Export Master Inquiry Sheet to Excel (xlsx)",
                    data=buffer.getvalue(),
                    file_name="Master_Showroom_Inquiry_Ledger.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.markdown("---")
                st.markdown("### Active Pipeline Logs Grid")
                
                for index_pos, row_item in df_master.iterrows():
                    with st.expander(f"Lead Reference: {row_item['customer_name']} | Store: {row_item['store_name']} | Rep: {row_item['emp_name']}"):
                        col_adm1, col_adm2, col_adm3 = st.columns(3)
                        col_adm1.write(f"**Value:** {row_item['product_value']} AED | **SAP Code:** {int(row_item['article_code'])}")
                        
                        up_cx = col_adm2.selectbox("Override CX State", [
                            "FUTURE BUYER", "JUST BROWSING", "NOT PART OF STORE ASSORTMENT", 
                            "NOT PART OF UAE STORES' ASSORTMENT", "OTH", "PRICE ISSUE", 
                            "SALES CLOSED", "STOCK ISSUE"
                        ], key=f"global_cx_{row_item['id']}", index=[
                            "FUTURE BUYER", "JUST BROWSING", "NOT PART OF STORE ASSORTMENT", 
                            "NOT PART OF UAE STORES' ASSORTMENT", "OTH", "PRICE ISSUE", 
                            "SALES CLOSED", "STOCK ISSUE"
                        ].index(row_item['cx_response']))
                        
                        current_rating_idx = int(row_item['google_rating']) - 1 if row_item['google_rating'] and int(row_item['google_rating']) > 0 else 0
                        up_rating = col_adm3.selectbox("Inject Captured Google Star Rating", [1, 2, 3, 4, 5], key=f"global_rt_{row_item['id']}", index=current_rating_idx)
                        
                        txt_review = st.text_input("Inject Google Review Content Text", value=str(row_item['google_review_text'] or ''), key=f"global_tx_{row_item['id']}")
                        
                        if up_cx != row_item['cx_response'] or up_rating != row_item['google_rating'] or txt_review != row_item['google_review_text']:
                            if st.button("Commit Infrastructure Profile Update", key=f"global_btn_{row_item['id']}", type="primary"):
                                update_packet = {
                                    "cx_response": up_cx,
                                    "google_rating": int(up_rating),
                                    "google_review_text": txt_review
                                }
                                query_ledger("UPDATE", payload=update_packet, target_id=row_item['id'])
                                st.success("Database tracking parameters altered cleanly!")
                                st.rerun()
            else:
                st.info("Central database sheet contains no active logged leads.")
        else:
            st.error("Admin clearance authentication missing from this profile token.")

# Standalone execution anchor
if __name__ == "__main__":
    import io
    render_ui()
