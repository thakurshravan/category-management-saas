# app.py
import os
import sys
import importlib.util
import streamlit as st

# Configure the master application window layout
st.set_page_config(page_title="Category Management Automation Suite", page_icon="🚀", layout="wide")

# Directory where all backend automation tools are stored
TOOLS_DIR = "tools_library"

# Create the folder automatically if it's missing
if not os.path.exists(TOOLS_DIR):
    os.makedirs(TOOLS_DIR)

# Ensure the system path can see files inside tools_library
if TOOLS_DIR not in sys.path:
    sys.path.append(TOOLS_DIR)


def load_dynamic_tools():
    """
    Scans tools_library/ and automatically imports any Python file 
    that has TOOL_NAME and render_ui() defined.
    """
    modules_found = {}
    
    # Read files from folder and sort them alphabetically
    for file in sorted(os.listdir(TOOLS_DIR)):
        if file.endswith(".py") and not file.startswith("__"):
            module_name = file[:-3]  # Strip '.py'
            file_path = os.path.join(TOOLS_DIR, file)
            
            try:
                # Dynamic import execution setup
                spec = importlib.util.spec_from_file_location(module_name, file_path)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                
                # Check for structural requirements
                if hasattr(mod, "TOOL_NAME") and hasattr(mod, "render_ui"):
                    icon = getattr(mod, "TOOL_ICON", "⚙️")
                    sidebar_label = f"{icon} {mod.TOOL_NAME}"
                    modules_found[sidebar_label] = mod
            except Exception as e:
                st.sidebar.error(f"⚠️ Error loading file '{file}': {e}")
                
    return modules_found


# ==========================================
# SIDEBAR DYNAMIC NAVIGATION
# ==========================================
st.sidebar.title("Navigation Menu")
st.sidebar.markdown("---")

# Scan the folder for modules
available_tools = load_dynamic_tools()

if available_tools:
    selected_tool = st.sidebar.radio(
        "Select an Automation Tool:",
        list(available_tools.keys())
    )
    st.sidebar.markdown("---")
    st.sidebar.caption(f"🤖 Connected tools: {len(available_tools)}")
else:
    st.sidebar.warning("No automation scripts found inside `tools_library/` folder.")
    selected_tool = None

st.sidebar.info("💡 **Drop-and-Play:** Drop a new script into `tools_library/` and it instantly syncs up here!")

# ==========================================
# MASTER UI ROUTER RUNNER
# ==========================================
if selected_tool and selected_tool in available_tools:
    # Run the UI function of the selected module directly 
    available_tools[selected_tool].render_ui()
else:
    st.title("Welcome to your Category Management Suite")
    st.markdown("Please place your automation files into the `tools_library/` folder to activate the platform.")