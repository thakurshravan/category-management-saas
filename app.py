import streamlit as st
import pandas as pd
import qrcode
import io
import base64

st.set_page_config(page_title="SaaS Bulk Label Generator", layout="wide")

if "file_uploader_key" not in st.session_state:
    st.session_state["file_uploader_key"] = 0

# --- GLOBAL UTILITY LOGIC ---
def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))

def generate_qr_base64(url):
    qr = qrcode.QRCode(version=1, box_size=4, border=1)
    qr.add_data(url)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    buffered = io.BytesIO()
    img.save(buffered, format="PNG")
    return base64.b64encode(buffered.getvalue()).decode()

# --- SIDEBAR: CONFIGURATION ---
st.sidebar.header("🎨 Advanced Customization Engine")

SIZE_TEMPLATES = {
    "60x40 mm (Standard Shelf Edge)": {"w": 60, "h": 40, "qr_size": 18},
    "40x50 mm (Hang Tag)": {"w": 40, "h": 50, "qr_size": 15},
    "80x50 mm (Large Display)": {"w": 80, "h": 50, "qr_size": 22},
    "Custom Size...": None
}

selected_size = st.sidebar.selectbox("1. Select Target Ticket Size", list(SIZE_TEMPLATES.keys()))

if selected_size == "Custom Size...":
    st.sidebar.markdown("📐 **Enter Manual Dimensions (in mm):**")
    custom_w = st.sidebar.number_input("Ticket Width (mm)", min_value=10, max_value=200, value=60, step=1)
    custom_h = st.sidebar.number_input("Ticket Height (mm)", min_value=10, max_value=200, value=40, step=1)
    custom_qr = st.sidebar.number_input("QR Code Size (mm)", min_value=5, max_value=min(custom_w, custom_h)-5, value=18, step=1)
    dimensions = {"w": custom_w, "h": custom_h, "qr_size": custom_qr}
else:
    dimensions = SIZE_TEMPLATES[selected_size]

primary_color = st.sidebar.color_picker("Text & Accent Color", "#1E1E1E")
bg_style = st.sidebar.selectbox("Ticket Background Style", ["Plain White", "Light Border Box", "Solid Accent Header"])

st.sidebar.subheader("Typography")
font_choice = st.sidebar.selectbox("Select Font Family", ["Arial", "Helvetica", "Courier"])
title_size = st.sidebar.slider("Product Name Font Size", 8, 20, 11)
price_size = st.sidebar.slider("Price Font Size", 14, 32, 18)

# Control to toggle discount lines
show_was_price = st.sidebar.checkbox("Show Strikethrough 'Was' Price Row", value=True)

# --- GLOBAL FONT CONFIGURATION ---
web_font = "Courier New, monospace" if font_choice == "Courier" else f"{font_choice}, sans-serif"

# --- MAIN INTERFACE LAYOUT ---
st.title("🎟️ Custom SaaS Bulk Price Ticket Generator")
st.write("Upload a file, customize styles, and print directly onto standard A4 sticker sheets.")

# --- LIVE PREVIEW WINDOW ---
st.subheader("👀 Live Ticket Sample Preview")
preview_border = f"2px solid {primary_color}" if bg_style == "Light Border Box" or bg_style == "Solid Accent Header" else "1px dashed #ccc"
preview_header_bg = primary_color if bg_style == "Solid Accent Header" else "transparent"
preview_header_text = "#ffffff" if bg_style == "Solid Accent Header" else primary_color

was_price_html = f'<div style="text-decoration: line-through; font-size: {price_size - 4}pt; color: #888; line-height: 1;">AED 879.00</div>' if show_was_price else ''

preview_html = f"""
<div style="
    width: {dimensions['w'] * 5}px; 
    height: {dimensions['h'] * 5}px; 
    border: {preview_border}; 
    background-color: #ffffff; 
    border-radius: 4px; 
    position: relative; 
    font-family: {web_font}; 
    overflow: hidden;
    box-shadow: 0 4px 10px rgba(0,0,0,0.15);
    margin-bottom: 20px;
">
    <div style="background-color: {preview_header_bg}; padding: 8px; height: 35%; box-sizing: border-box;">
        <div style="color: {preview_header_text}; font-size: {title_size + 2}px; font-weight: bold; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">
            Sample Product Name
        </div>
    </div>
    <div style="position: absolute; top: 45%; left: 8px; color: #555; font-size: 11px;">
        SKU: J705466
    </div>
    
    <div style="position: absolute; bottom: 8px; left: 8px;">
        {was_price_html}
        <div style="color: {primary_color}; font-size: {price_size + 4}px; font-weight: bold; line-height: 1.1;">
            AED 799.00
        </div>
    </div>
    
    <div style="position: absolute; bottom: 8px; right: 8px; width: {dimensions['qr_size'] * 4.5}px; height: {dimensions['qr_size'] * 4.5}px; background-image: url('https://upload.wikimedia.org/wikipedia/commons/d/d0/QR_code_for_QR_Code_tutorial_images_section.png'); background-size: cover; border: 1px solid #eee;"></div>
</div>
"""
st.markdown(preview_html, unsafe_allow_html=True)
st.divider()

# --- DATA IMPORT ENGINE ---
def clear_file_callback():
    st.session_state["file_uploader_key"] += 1

uploaded_file = st.file_uploader(
    "Upload Product File (.xlsx or .csv)", 
    type=["xlsx", "csv"], 
    key=f"file_uploader_{st.session_state['file_uploader_key']}"
)

if uploaded_file is not None:
    if st.button("🗑️ Clear File & Reset Canvas", on_click=clear_file_callback):
        st.rerun()

    try:
        file_bytes = uploaded_file.getvalue()
        
        if uploaded_file.name.endswith('.csv'):
            data_str = file_bytes.decode('utf-8', errors='ignore')
            df = pd.read_csv(io.StringIO(data_str), header=None)
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), header=None)
        
        df = df.dropna(how='all').dropna(axis=1, how='all')
        df.columns = df.iloc[0].astype(str).str.strip()
        df = df[1:].reset_index(drop=True)
        
        mapped_cols = {}
        for col in df.columns:
            col_lower = str(col).lower()
            if "sku" in col_lower:
                mapped_cols["SKU"] = col
            elif "product" in col_lower or "name" in col_lower or "title" in col_lower:
                mapped_cols["Product Name"] = col
            elif "price" in col_lower or "rate" in col_lower or "mrp" in col_lower:
                mapped_cols["Price"] = col
            elif "url" in col_lower or "link" in col_lower or "website" in col_lower:
                mapped_cols["URL"] = col

        required_targets = ["SKU", "Product Name", "Price", "URL"]
        missing_targets = [t for t in required_targets if t not in mapped_cols]
        
        if missing_targets:
            st.error(f"Execution Halted. Missing columns: {missing_targets}")
        else:
            st.success("✨ Data payload mapped successfully!")
            st.dataframe(df.head(3), use_container_width=True)

            # --- GENERATION ENGINE ---
            card_border = f"2px solid {primary_color}" if bg_style == "Light Border Box" or bg_style == "Solid Accent Header" else "1px dashed #ccc"
            header_bg = primary_color if bg_style == "Solid Accent Header" else "transparent"
            header_text_color = "#ffffff" if bg_style == "Solid Accent Header" else primary_color

            html_cards = ""
            for idx, row in df.iterrows():
                # FIXED: Declaring safe fallbacks inside variable tracking to protect missing/corrupt values
                price_text = ""
                was_price_text = ""
                
                try:
                    price_val = float(row[mapped_cols["Price"]])
                    price_text = f"AED {price_val:.2f}"
                    was_price_text = f"AED {price_val * 1.15:.2f}"
                except:
                    price_text = f"AED {row[mapped_cols['Price']]}"
                    was_price_text = ""

                qr_b64 = generate_qr_base64(str(row[mapped_cols["URL"]]))
                
                was_row_inner = f'<div style="text-decoration: line-through; font-size: {price_size - 4}pt; color: #888; margin-bottom: 1px;">{was_price_text}</div>' if (show_was_price and was_price_text) else ''

                html_cards += f"""
                <div class="ticket-card" style="
                    width: {dimensions['w']}mm;
                    height: {dimensions['h']}mm;
                    border: {card_border};
                    box-sizing: border-box;
                    position: relative;
                    background: #fff;
                    font-family: {web_font};
                    overflow: hidden;
                    display: inline-block;
                    margin: 1mm;
                    vertical-align: top;
                    text-align: left;
                ">
                    <div style="background-color: {header_bg}; padding: 6px; height: 35%; box-sizing: border-box;">
                        <div style="color: {header_text_color}; font-size: {title_size}pt; font-weight: bold; line-height: 1.1; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden;">
                            {row[mapped_cols["Product Name"]]}
                        </div>
                    </div>
                    <div style="position: absolute; top: 42%; left: 6px; color: #555; font-size: 8pt;">
                        SKU: {row[mapped_cols["SKU"]]}
                    </div>
                    
                    <div style="position: absolute; bottom: 6px; left: 6px; line-height: 1;">
                        {was_row_inner}
                        <div style="color: {primary_color}; font-size: {price_size}pt; font-weight: bold;">
                            {price_text}
                        </div>
                    </div>
                    
                    <img src="data:image/png;base64,{qr_b64}" style="
                        position: absolute;
                        bottom: 4px;
                        right: 4px;
                        width: {dimensions['qr_size']}mm;
                        height: {dimensions['qr_size']}mm;
                    " />
                </div>
                """

            st.subheader("🖨️ Printable Document Feed")
            
            iframe_content = f"""
            <html>
            <head>
                <style>
                    body {{ margin: 0; padding: 0; font-family: sans-serif; text-align: center; background: #fafafa; }}
                    .print-btn {{
                        background-color: #25d366; color: white; border: none; padding: 12px 30px;
                        font-size: 16px; font-weight: bold; border-radius: 6px; cursor: pointer;
                        box-shadow: 0 4px 6px rgba(0,0,0,0.1); margin: 15px auto; display: block;
                    }}
                    .print-btn:hover {{ background-color: #1ebe57; }}
                    .a4-page {{
                        width: 210mm; padding: 5mm; margin: 0 auto;
                        background: white; box-sizing: border-box; text-align: left;
                    }}
                    @media print {{
                        .print-btn {{ display: none !important; }}
                        body {{ background: white; }}
                        .a4-page {{ padding: 0; margin: 0; width: 100%; }}
                    }}
                </style>
            </head>
            <body>
                <button class="print-btn" onclick="window.print()">🖨️ CLICK HERE TO PRINT BULK SHEET</button>
                <div class="a4-page">
                    {html_cards}
                </div>
            </body>
            </html>
            """
            st.components.v1.html(iframe_content, height=800, scrolling=True)

    except Exception as e:
        st.error(f"Fatal Parser Error: {e}")
