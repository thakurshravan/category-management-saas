import os
import sys
import importlib.util
import streamlit as st
import requests

# Configure the master application window layout
st.set_page_config(page_title="Category Management SaaS Suite", page_icon="🚀", layout="wide")

# Directory configurations
TOOLS_DIR = "tools_library"
if not os.path.exists(TOOLS_DIR):
    os.makedirs(TOOLS_DIR)
if TOOLS_DIR not in sys.path:
    sys.path.append(TOOLS_DIR)

# Fetch secure Supabase credentials from your Streamlit Environment Secrets
SB_URL = st.secrets.get("SUPABASE_URL", "")
SB_KEY = st.secrets.get("SUPABASE_KEY", "")

# ==========================================================
# 🔐 SUPABASE AUTHENTICATION ENGINE UTILITIES
# ==========================================================
def supabase_auth_request(endpoint, payload):
    """Handles raw HTTPS communication requests to your Supabase Auth server API."""
    headers = {
        "apiKey": SB_KEY,
        "Content-Type": "application/json"
    }
    url = f"{SB_URL}/auth/v1/{endpoint}"
    try:
        response = requests.post(url, json=payload, headers=headers)
        return response.json(), response.status_code
    except Exception as e:
        return {"error_description": str(e)}, 500

def login_user(email, password):
    payload = {"email": email.strip(), "password": password, "gotrue_meta_security": {}}
    res, status = supabase_auth_request("token?grant_type=password", payload)
    if status == 200:
        st.session_state["auth_token"] = res.get("access_token")
        st.session_state["user_email"] = res.get("user", {}).get("email")
        st.session_state["user_id"] = res.get("user", {}).get("id")
        return True, "Success"
    
    # Expose the precise API error reason
    err_msg = res.get("error_description") or res.get("error", {}).get("message") or "Invalid credentials."
    return False, f"Database Refusal: {err_msg}"

def register_user(email, password):
    payload = {"email": email.strip(), "password": password}
    res, status = supabase_auth_request("signup", payload)
    if status == 200:
        return True, "Account created! Proceed to the Login tab."
    
    err_msg = res.get("error_description") or res.get("error", {}).get("message") or "Registration blocked."
    return False, f"Database Refusal: {err_msg}"

# Initialize global authentication states in session memory
if "auth_token" not in st.session_state:
    st.session_state["auth_token"] = None

# ==========================================================
# 🗺️ SECURITY URL ROUTER (PUBLIC DEEP LINKS)
# ==========================================================
url_params = st.query_params

# If a supplier arrives via an active token drop link, bypass login completely!
if "view" in url_params and url_params["view"] == "portal":
    try:
        spec = importlib.util.spec_from_file_location("stock_aggregator", os.path.join(TOOLS_DIR, "stock_aggregator.py"))
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        mod.render_public_supplier_portal(url_params)
    except Exception as e:
        st.error(f"🚨 Supplier Gate Error: {e}")
    st.stop()

# ==========================================================
# 🛑 PRIVATE GATE: LOGIN / SIGNUP PORTAL SCREEN
# ==========================================================
if st.session_state["auth_token"] is None:
    _, col_auth, _ = st.columns([1, 1.5, 1])
    
    with col_auth:
        st.title("🚀 Category Management SaaS Suite")
        st.markdown("Commercial Portal Suite — Authenticate to access your management suite tools.")
        
        auth_mode = st.tabs(["🔒 Account Login", "✨ Create Subscriber Account"])
        
        # A. LOGIN INTERFACE TAB
        with auth_mode[0]:
            login_email = st.text_input("Business Email Address", key="login_em")
            login_password = st.text_input("Account Password", type="password", key="login_pw")
            
            if st.button("Authenticate into Workspace", type="primary", use_container_width=True):
                if login_email and login_password:
                    success, msg = login_user(login_email, login_password)
                    if success:
                        st.success("Access Granted! Launching workspace layout...")
                        st.rerun()
                    else:
                        st.error(msg)
                else:
                    st.warning("Please fill out all fields.")
                    
        # B. REGISTRATION INTERFACE TAB
        with auth_mode[1]:
            st.markdown("### Start your subscription cycle")
            reg_email = st.text_input("Enter Email Address", key="reg_em")
            reg_password = st.text_input("Create Secure Password", type="password", key="reg_pw")
            
            if st.button("Register & Initialize Tenant License", use_container_width=True):
                if reg_email and reg_password:
                    success, msg = register_user(reg_email, reg_password)
                    if success:
                        st.success(msg)
                    else:
                        st.error(msg)
                else:
                    st.warning("Please fill out all registration fields.")
    st.stop()

# ==========================================================
# 📊 INTERNAL PAGINATION LAYER (AUTHORIZED USERS ONLY)
# ==========================================================
def load_dynamic_tools():
    modules_found = {}
    for file in sorted(os.listdir(TOOLS_DIR)):
        if file.endswith(".py") and not file.startswith("__"):
            module_name = file[:-3]
            file_path = os.path.join(TOOLS_DIR, file)
            try:
                spec = importlib.util.spec_from_file_location(module_name, file_path)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                if hasattr(mod, "TOOL_NAME") and hasattr(mod, "render_ui"):
                    icon = getattr(mod, "TOOL_ICON", "⚙️")
                    modules_found[f"{icon} {mod.TOOL_NAME}"] = mod
            except Exception as e:
                st.sidebar.error(f"⚠️ Load Error '{file}': {e}")
    return modules_found

# Render Protected Sidebar Navigation Environment
st.sidebar.title("SaaS Control Center")
st.sidebar.caption(f"👤 Active License: `{st.session_state['user_email']}`")
st.sidebar.markdown("---")

available_tools = load_dynamic_tools()

if available_tools:
    selected_tool = st.sidebar.radio("Select an Automation Tool:", list(available_tools.keys()))
    st.sidebar.markdown("---")
else:
    st.sidebar.warning("No operational assets assigned to your account partition.")
    selected_tool = None

# Log out execution command
if st.sidebar.button("Log Out of Session", use_container_width=True):
    st.session_state["auth_token"] = None
    st.session_state["user_email"] = None
    st.session_state["user_id"] = None
    st.rerun()

# Run the selected tool UI directly
if selected_tool and selected_tool in available_tools:
    available_tools[selected_tool].render_ui()
else:
    st.title("Welcome to your Category Management Suite")
    st.markdown("Initialize your operational loop modules using the navigation menu dashboard selection controls.")
