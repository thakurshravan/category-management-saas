# tools_library/sns_reporting.py
import streamlit as st
import datetime

# --- Master Connection Keys ---
TOOL_NAME = "SNS Reporting Tool"
TOOL_ICON = "📈"

def run_sns_process(sales, stock, emails, exclusions, send, start, end):
    """
    Your backend execution code logic goes here.
    """
    # Dummy processing return for demonstration
    return 12 

def render_ui():
    st.title("Sales & Stock Reporting Tool")
    st.subheader("Generate separated sheet reports and distribute them directly.")
    
    # Component Layout Uploaders
    sales_file = st.file_uploader("Upload Sales File (.xlsx / .csv)", type=["xlsx", "csv"], key="sns_sales")
    stock_file = st.file_uploader("Upload Stock File (.xlsx / .csv)", type=["xlsx", "csv"], key="sns_stock")
    email_master_file = st.file_uploader("Upload Email Master File (.xlsx / .csv)", type=["xlsx", "csv"], key="sns_email")
    exclusions_file = st.file_uploader("Upload Exclusions File (.xlsx / .csv)", type=["xlsx", "csv"], key="sns_exclude")
    
    st.markdown("### Date Filter Configuration")
    date_col1, date_col2 = st.columns(2)
    with date_col1:
        start_date = st.date_input("From Date:", value=datetime.date(2026, 5, 1), key="sns_start")
    with date_col2:
        end_date = st.date_input("To Date:", value=datetime.date(2026, 5, 31), key="sns_end")
        
    send_email = st.checkbox("Send live emails automatically using local Outlook?", value=True, key="sns_send")
    
    st.markdown("---")
    
    if st.button("⚡ Run Automation Engine", type="primary", key="sns_run"):
        if sales_file and stock_file and email_master_file and exclusions_file:
            with st.spinner("Running calculations and generating files..."):
                count = run_sns_process(sales_file, stock_file, email_master_file, exclusions_file, send_email, start_date, end_date)
                st.success(f"🎉 Success! Processed and compiled individual multi-sheet reports for {count} suppliers.")
        else:
            st.warning("Please upload all files before running.")