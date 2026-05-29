import streamlit as st
import requests
import pandas as pd
import plotly.express as px

# ==========================================================
# 🏷️ MODULE METADATA & ENTRY REGISTRATION FOR STREAMLIT
# ==========================================================
TOOL_NAME = "Sales Hub & Performance Tracker"
TOOL_ICON = "⚡"

# Fetch operational keys from application environments
SB_URL = st.secrets.get("SUPABASE_URL", "")
SB_KEY = st.secrets.get("SUPABASE_KEY", "")

def query_ledger(action, payload=None, target_id=None):
    """Secure background communication tunnel with the master Supabase storage array."""
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
    active_salesman = st.session_state.get("user_email", "staff@company.com")
    
    # 💰 COMMISSION & LOGIC REWARD RULES CONFIGURATION
    BASE_COMMISSION_PERCENT = 0.02  # 2.0% Base payout on 'Won' transactions
    STRETCH_GOAL_THRESHOLD = 15000  # AED monthly tier target
    STRETCH_GOAL_BONUS = 500       # Lump sum bonus reward payout
    GOOGLE_REVIEW_BONUS = 25       # Cash prize added per 5-Star review captured

    st.title(f"{TOOL_ICON} Showroom Sales & Operations Hub")
    st.caption(f"📱 Connected Terminal License: `{active_salesman}`")
    st.markdown("---")

    # Mobile optimized system navigation selector tabs
    hub_tabs = st.tabs(["📝 Inquiry Capture", "💰 My Earnings", "📈 Performance KPIs", "⭐ Review Stream", "👑 Admin Board"])

    # 📥 Fetch records from central cloud engine once to optimize database roundtrips
    raw_ledger = query_ledger("SELECT")
    df_master = pd.DataFrame(raw_ledger) if raw_ledger else pd.DataFrame()

    # Filter localized row selections for individual salesman profiles
    if not df_master.empty:
        df_personal = df_master[df_master['salesman_email'] == active_salesman].copy()
    else:
        df_personal = pd.DataFrame()

    # ==========================================================
    # 1. CUSTOMER INQUIRY CAPTURE (Mobile Form Layout)
    # ==========================================================
    with hub_tabs[0]:
        st.subheader("Showroom Demand Intake Form")
        with st.form("showroom_intake_form", clear_on_submit=True):
            cust_name = st.text_input("Customer Full Name", value="Walk-in Lead")
            cust_phone = st.text_input("Contact Mobile Number (Required for validation)")
            prod_cat = st.selectbox("Product Categorization Segment", ["Home Electronics", "Major Domestic Appliances", "Premium Sound Ecosystems", "Mobile Devices & Wearables"])
            deal_val = st.number_input("Quoted Offer Value (AED)", min_value=0.0, step=250.0)
            initial_status = st.selectbox("Deal Status Flag", ["Pending", "Won", "Lost"])
            
            if st.form_submit_button("Lock Lead to Profile Grid", type="primary", use_container_width=True):
                if cust_phone:
                    entry_payload = {
                        "salesman_email": active_salesman,
                        "customer_name": cust_name,
                        "customer_phone": cust_phone,
                        "product_category": prod_cat,
                        "quoted_deal_value": deal_val,
                        "deal_status": initial_status
                    }
                    if query_ledger("INSERT", payload=entry_payload):
                        st.success("🎉 Showroom entry recorded and mapped under your commission ID successfully.")
                        st.rerun()
                    else:
                        st.error("Network synchronization timeout. Re-submit entry.")
                else:
                    st.warning("Customer contact number is required to run automated conversions tracking.")

    # ==========================================================
    # 2. INCENTIVE TRACKING ENGINE (Live Payout Scorecard)
    # ==========================================================
    with hub_tabs[1]:
        st.subheader("My Commission & Incentive Trackers")
        if not df_personal.empty:
            df_won = df_personal[df_personal['deal_status'] == 'Won']
            won_volume = df_won['quoted_deal_value'].sum()
            
            # Mathematical compensation logic calculations
            base_payout = won_volume * BASE_COMMISSION_PERCENT
            five_star_reviews = df_personal[df_personal['google_rating'] == 5]
            review_commissions = len(five_star_reviews) * GOOGLE_REVIEW_BONUS
            unlocked_stretch_bonus = STRETCH_GOAL_BONUS if won_volume >= STRETCH_GOAL_THRESHOLD else 0.0
            total_accrued_payout = base_payout + review_commissions + unlocked_stretch_bonus

            # Render KPI Blocks
            p_col1, p_col2, p_col3 = st.columns(3)
            p_col1.metric("Closed Volume (AED)", f"{won_volume:,.2f}")
            p_col2.metric("Base Commission Accrued", f"{base_payout:,.2f}")
            p_col3.metric("Review Cash Rewards", f"{review_commissions:,.2f}")
            
            st.metric("Estimated Monthly Paycheck", f"{total_accrued_payout:,.2f} AED", 
                      delta=f"{STRETCH_GOAL_BONUS} AED Stretch Target Met!" if unlocked_stretch_bonus > 0 else "Target Lock Pending")
            
            # Progress bar visualization tracking stretch goal completion
            stretch_progress = min(1.0, float(won_volume / STRETCH_GOAL_THRESHOLD)) if STRETCH_GOAL_THRESHOLD > 0 else 1.0
            st.markdown(f"**Monthly Volume Stretch Target: {stretch_progress * 100:.1f}% Complete ({won_volume:,.2f} / {STRETCH_GOAL_THRESHOLD:,.2f} AED)**")
            st.progress(stretch_progress)
        else:
            st.info("Log your initial closed showroom sale transaction to initialize personal tracking grids.")

    # ==========================================================
    # 3. SALESMAN PERFORMANCE TRACKING DASHBOARD (Visual Analytics)
    # ==========================================================
    with hub_tabs[2]:
        st.subheader("Personal Sales Performance KPIs")
        if not df_personal.empty:
            kpi_c1, kpi_c2, kpi_c3 = st.columns(3)
            total_leads_logged = len(df_personal)
            successful_closes = len(df_personal[df_personal['deal_status'] == 'Won'])
            conversion_rate = (successful_closes / total_leads_logged) * 100 if total_leads_logged > 0 else 0.0
            
            kpi_c1.metric("Total Inquiries Handled", total_leads_logged)
            kpi_c2.metric("Successful Deals Closed", successful_closes)
            kpi_c3.metric("Deal Conversion Rate", f"{conversion_rate:.1f}%")
            
            # Historical pipeline conversion graph
            st.markdown("### Deal Pipeline Segmentation Breakdown")
            pipeline_mix = df_personal.groupby('deal_status').size().reset_index(name='Total_Leads')
            fig_pie = px.pie(pipeline_mix, values='Total_Leads', names='deal_status', hole=0.4, color_discrete_sequence=px.colors.sequential.RdBu)
            st.plotly_chart(fig_pie, use_container_width=True)
        else:
            st.info("No conversion data trends mapped to this profile yet.")

    # ==========================================================
    # 4. GOOGLE FEEDBACK CAPTURING & AUTOMATION
    # ==========================================================
    with hub_tabs[3]:
        st.subheader("Showroom Reputation Pipeline & QR Generator")
        st.markdown(
            "When handing over an invoice, show this QR generator tool to the customer to request "
            "a 5-star rating. Captured 5-star reviews add an instant cash reward to your profile!"
        )
        
        # QR Code Mock Router Mockup for instant user experience
        st.markdown(
            f'<div style="background-color:#fff; padding:20px; text-align:center; border-radius:10px; width:220px; margin:auto;">'
            f'<img src="https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=https://search.google.com/local/writereview" />'
            f'<br/><b style="color:#000;">Scan to Leave Review</b></div>', 
            unsafe_allow_url=True
        )
        
        st.markdown("---")
        st.markdown("### My Captured Google Reviews Ledger")
        if not df_personal.empty:
            df_reviews_only = df_personal[df_personal['google_rating'].notna() & (df_personal['google_rating'] > 0)]
            if not df_reviews_only.empty:
                st.dataframe(df_reviews_only[['created_at', 'customer_name', 'google_rating', 'google_review_text']], use_container_width=True, hide_index=True)
            else:
                st.caption("No Google platform star feedback ratings linked to your deal identifier hashes yet.")

    # ==========================================================
    # 5. EXECUTIVE ADMIN CONTROL BOARD (You / Management View)
    # ==========================================================
    with hub_tabs[4]:
        st.subheader("Global Operations Dashboard")
        if active_salesman in ["thakurshravan@hotmail.com", "thakshravan@gmail.com"]:
            if not df_master.empty:
                # Global Showroom Leaderboard visualization chart
                st.markdown("### Organizational Sales Leaderboard by Representative")
                df_won_global = df_master[df_master['deal_status'] == 'Won']
                if not df_won_global.empty:
                    leaderboard_data = df_won_global.groupby('salesman_email')['quoted_deal_value'].sum().reset_index()
                    fig_leaderboard = px.bar(leaderboard_data, x='salesman_email', y='quoted_deal_value', color='quoted_deal_value', labels={'quoted_deal_value': 'Closed Sales (AED)', 'salesman_email': 'Representative ID'})
                    st.plotly_chart(fig_leaderboard, use_container_width=True)

                st.markdown("### Live Network Pipeline Database Records Log")
                for index_pos, row_item in df_master.iterrows():
                    with st.expander(f"Lead Reference: {row_item['customer_name']} | Managed by: {row_item['salesman_email']}"):
                        col_adm1, col_adm2, col_adm3 = st.columns(3)
                        col_adm1.write(f"**Value:** {row_item['quoted_deal_value']} AED")
                        
                        # Admin management fields to log reviews manually or overwrite statuses
                        up_status = col_adm2.selectbox("Override Transaction State", ["Pending", "Won", "Lost"], key=f"global_st_{row_item['id']}", index=["Pending", "Won", "Lost"].index(row_item['deal_status']))
                        
                        current_rating_idx = int(row_item['google_rating']) - 1 if row_item['google_rating'] and int(row_item['google_rating']) > 0 else 0
                        up_rating = col_adm3.selectbox("Inject Captured Google Star Rating", [1, 2, 3, 4, 5], key=f"global_rt_{row_item['id']}", index=current_rating_idx)
                        
                        txt_review = st.text_input("Inject Google Review Content Text", value=str(row_item['google_review_text'] or ''), key=f"global_tx_{row_item['id']}")
                        
                        if up_status != row_item['deal_status'] or up_rating != row_item['google_rating'] or txt_review != row_item['google_review_text']:
                            if st.button("Commit Infrastructure Profile Update", key=f"global_btn_{row_item['id']}", type="primary"):
                                update_packet = {
                                    "deal_status": up_status,
                                    "google_rating": int(up_rating),
                                    "google_review_text": txt_review
                                }
                                query_ledger("UPDATE", payload=update_packet, target_id=row_item['id'])
                                st.success("Database tracking parameters altered cleanly!")
                                st.rerun()
            else:
                st.info("Central database ledger contains no operational record sets.")
        else:
            st.error("Access Prohibited. Your active license level does not possess clearance signatures for Global Admin boards.")

if __name__ == "__main__":
    render_ui()