import streamlit as st
import requests
import pandas as pd
import base64
from io import BytesIO
from PIL import Image

# ==========================================================
# 🏷️ MODULE METADATA & ENTRY REGISTRATION
# ==========================================================
TOOL_NAME = "Company Community Wall"
TOOL_ICON = "🥳"

# Fetch secure project parameters from secrets
SB_URL = st.secrets.get("SUPABASE_URL", "")
SB_KEY = st.secrets.get("SUPABASE_KEY", "")

def run_supabase_query(action, payload=None):
    """Handles communications with the community posts table endpoint."""
    headers = {
        "apiKey": SB_KEY,
        "Authorization": f"Bearer {SB_KEY}",
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    url = f"{SB_URL}/rest/v1/community_posts"
    
    try:
        if action == "INSERT":
            res = requests.post(url, json=payload, headers=headers)
        elif action == "SELECT":
            res = requests.get(f"{url}?order=created_at.desc", headers=headers)
        return res.json() if res.status_code in [200, 201] else []
    except Exception:
        return []

def render_ui():
    active_user = st.session_state.get("user_email", "unknown@domain.com")
    
    st.title(f"{TOOL_ICON} {TOOL_NAME}")
    st.markdown("Celebrate wins, share performance updates, and view team milestones across all showroom branches!")
    st.markdown("---")

    # Layout: Top widget handles new posts, bottom renders the feed timeline
    with st.expander("✨ Broadcast a New Team Victory (Click to Expand)", expanded=False):
        with st.form("community_post_form", clear_on_submit=True):
            post_title = st.text_input("Victory Title / Headline", placeholder="e.g., Showroom Hero of the Week! 🏆")
            post_caption = st.text_area("Caption / What happened?", placeholder="e.g., Shravan just secured a 45,000 AED appliance order! Phenomenal work! 🔥")
            
            # Mobile-friendly picture uploader
            uploaded_pic = st.file_uploader("Upload Performance Pic (JPG/PNG)", type=["jpg", "jpeg", "png"])
            
            if st.form_submit_button("Broadcast to the Company Wall", type="primary"):
                if post_title and post_caption and uploaded_pic:
                    try:
                        # Convert image stream into a web-safe Base64 data string to save directly into Supabase safely
                        img = Image.open(uploaded_pic)
                        # Resize slightly to compress and maintain lightning fast image rendering loads
                        img.thumbnail((800, 800))
                        buffer = BytesIO()
                        img_format = uploaded_pic.name.split('.')[-1].upper()
                        img_format = "JPEG" if img_format == "JPG" else img_format
                        img.save(buffer, format=img_format)
                        img_str = base64.b64encode(buffer.getvalue()).decode()
                        data_url = f"data:image/{img_format.lower()};base64,{img_str}"

                        # Construction database packet array
                        payload = {
                            "author_email": active_user,
                            "title": post_title,
                            "caption": post_caption,
                            "image_url": data_url
                        }
                        
                        response = run_supabase_query("INSERT", payload=payload)
                        if response:
                            st.success("🎯 Victory successfully pinned to the Wall of Fame feed!")
                            st.rerun()
                        else:
                            st.error("Server connection timeout. Unable to broadcast post.")
                    except Exception as ex:
                        st.error(f"Image compilation break: {ex}")
                else:
                    st.warning("All inputs (Headline, Details, and a Picture) are required to publish a victory post.")

    st.markdown("### 🏆 Live Victory Feed Timeline")
    
    # ─── TIMELINE RENDERER ENGINE ───
    posts_data = run_supabase_query("SELECT")
    if posts_data:
        for post in posts_data:
            # Render a clean card interface container layout for every logged victory
            with st.container(border=True):
                col_meta, col_del = st.columns([4, 1])
                col_meta.markdown(f"### {post['title']}")
                col_meta.caption(f"👤 **Logged by:** `{post['author_email']}` | 📅 **Published:** {post['created_at'][:10]}")
                
                # Render base64 image data payload safely back into web visuals
                try:
                    header, encoded = post['image_url'].split(",", 1)
                    image_bytes = base64.b64decode(encoded)
                    display_image = Image.open(BytesIO(image_bytes))
                    st.image(display_image, use_container_width=True)
                except Exception:
                    st.caption("⚠️ Picture data block failed to load cleanly.")
                
                st.markdown(f"📢 **Performance Insights:** {post['caption']}")
    else:
        st.info("The community wall feed timeline is currently blank. Be the first to broadcast a showroom win above!")

if __name__ == "__main__":
    render_ui()