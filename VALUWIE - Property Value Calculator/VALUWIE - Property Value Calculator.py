import os
import urllib.request
import json
import difflib
import re
import io
import base64
import streamlit as st
import pandas as pd
import geopandas as gpd
import folium
from streamlit_folium import st_folium
from shapely.geometry import Point

# ReportLab PDF imports
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "Greater Luwero Base Land Valuation.xlsx")
SHAPEFILE_PATH = os.path.join(BASE_DIR, "Greater Luwero Blocks ugrf.shp")
LOGO_PATH = os.path.join(BASE_DIR, "Valuwie Logo.png")

# ==========================================
# PAGE CONFIGURATION & STYLING
# ==========================================
st.set_page_config(
    page_title="Valuwie | Property Valuation & Land Tax Portal",
    page_icon=LOGO_PATH if os.path.exists(LOGO_PATH) else "🏡",
    layout="wide"
)

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@500;700;800;900&display=swap');

    .main { background-color: #f1f5f9; }
    .block-container { padding-top: 2.8rem !important; max-width: 98%; }
    
    * { font-family: 'Plus Jakarta Sans', sans-serif !important; }

    /* Inline Side-by-Side Logo & Title Header with tight coupling */
    .brand-header-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        margin-bottom: 25px;
    }
    
    .brand-flex {
        display: flex;
        align-items: center;
        justify-content: center;
        gap: -2px;
    }

    .valuwie-title {
        font-size: 4.2rem;
        font-weight: 900;
        background: linear-gradient(45deg, #0284c7, #16a34a);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0px;
        line-height: 1;
        letter-spacing: -2px;
        text-decoration: none;
    }
    
    .valuwie-subtitle {
        font-size: 1.05rem;
        color: #0f172a;
        font-weight: 500;
        margin-top: 8px;
        letter-spacing: 0.2px;
        text-align: center;
    }

    /* Section Headings with Perfectly Aligned Larger Icons */
    .section-header-flex {
        display: flex;
        align-items: center;
        gap: 10px;
        border-bottom: 3px solid #cbd5e1;
        padding-bottom: 6px;
        margin-bottom: 10px;
    }

    .section-header-title {
        font-size: 1.25rem !important;
        font-weight: 900 !important;
        color: #0f172a !important;
        letter-spacing: -0.3px;
        margin: 0;
        line-height: 1.2;
    }

    /* High Contrast Labels */
    label, .stSelectbox label, .stNumberInput label, .stRadio label {
        font-weight: 800 !important;
        color: #0f172a !important;
        font-size: 0.88rem !important;
    }

    /* Styled Input Boxes */
    .stSelectbox div[data-baseweb="select"] > div,
    .stNumberInput input,
    .stTextInput input {
        border-radius: 12px !important;
        border: 2px solid #64748b !important;
        background-color: #ffffff !important;
        color: #0f172a !important;
        font-weight: 800 !important;
    }

    /* High Contrast Buttons */
    .stButton button, div.stDownloadButton > button {
        border-radius: 14px !important;
        font-weight: 900 !important;
        border: none !important;
        background: linear-gradient(45deg, #0284c7, #0284c7);
        color: white !important;
        box-shadow: 0 4px 12px rgba(2, 132, 199, 0.3);
        transition: all 0.2s ease-in-out;
        width: 100%;
    }
    
    .stButton button:hover, div.stDownloadButton > button:hover {
        transform: translateY(-1px);
        box-shadow: 0 6px 16px rgba(2, 132, 199, 0.4);
    }

    /* Smooth Scrolling on Dropdowns */
    div[data-baseweb="select"] > div {
        max-height: 280px;
        overflow-y: auto !important;
        border-radius: 12px !important;
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# DATA LOADING & CACHING
# ==========================================
@st.cache_data
def load_excel_data():
    try:
        df_blocks = pd.read_excel(EXCEL_PATH, sheet_name="Block Valuation Summary")
        df_agents = pd.read_excel(EXCEL_PATH, sheet_name="Agents Directory")
        df_agents = df_agents.loc[:, ~df_agents.columns.duplicated()]
        return df_blocks, df_agents
    except Exception as e:
        st.error(f"Failed to load property database: {e}")
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data
def load_gis_shapefile():
    if os.path.exists(SHAPEFILE_PATH):
        try:
            gdf = gpd.read_file(SHAPEFILE_PATH)
            if gdf.crs and gdf.crs.to_string() != "EPSG:4326":
                gdf = gdf.to_crs(epsg=4326)
            return gdf
        except Exception as e:
            st.warning(f"Map boundary file error: {e}")
            return None
    return None

df_blocks, df_agents = load_excel_data()
gdf_blocks = load_gis_shapefile()

shape_block_col = None
if gdf_blocks is not None:
    for col in ['BLOCK_NO_', 'Block_No', 'BLOCK', 'Block', 'block_no', 'BLOCK_NO']:
        if col in gdf_blocks.columns:
            shape_block_col = col
            break

# Initialize session state variables & Reset handling via URL query params
if 'reset_trigger' in st.query_params:
    for key in list(st.session_state.keys()):
        del st.session_state[key]
    st.query_params.clear()
    st.rerun()

if 'clicked_block' not in st.session_state:
    st.session_state.clicked_block = None
if 'searched_coords' not in st.session_state:
    st.session_state.searched_coords = None
if 'searched_name' not in st.session_state:
    st.session_state.searched_name = ""
if 'show_about' not in st.session_state:
    st.session_state.show_about = False
if 'prof_help_active' not in st.session_state:
    st.session_state.prof_help_active = False

LUWERO_DIRECTORY = {
    "bukalasa agricultural college": (0.7083, 32.4683),
    "wobulenzi town council": (0.7225, 32.5333),
    "bombo military barracks": (0.5703, 32.5408),
    "luwero district headquarters": (0.8456, 32.4756),
    "kasana market": (0.8500, 32.4800),
    "katikamu": (0.6850, 32.5500),
    "matugga": (0.5055, 32.5358),
    "zirobwe": (0.6333, 32.7000),
    "ndejje university": (0.6100, 32.4800),
    "kiwoko hospital": (0.8800, 32.3900),
    "wobulenzi health centre": (0.7200, 32.350)
}

# ==========================================
# HELPER FOR RENDERED SECTION HEADERS
# ==========================================
def render_header(title, icon_filename):
    icon_path = os.path.join(BASE_DIR, icon_filename)
    if os.path.exists(icon_path):
        with open(icon_path, "rb") as f:
            encoded_icon = base64.b64encode(f.read()).decode()
        st.markdown(f"""
            <div class="section-header-flex">
                <img src="data:image/png;base64,{encoded_icon}" style="width: 28px; height: 28px; object-fit: contain;" />
                <span class="section-header-title">{title}</span>
            </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="section-header-title" style="border-bottom: 3px solid #cbd5e1; padding-bottom: 6px; margin-bottom: 10px;">{title}</div>', unsafe_allow_html=True)

# ==========================================
# PDF REPORT GENERATION FUNCTION (CONTACTS COLUMN)
# ==========================================
def create_pdf_report(selected_district, selected_block, selected_estate, acres, land_value, building_val, total_property, stamp_duty, official_search, govt_total, surveyor_fee_total, title_creation_fee, agent_transfer_fee, total_acquisition, include_charges, include_survey, include_mutation, include_transfer, assigned_surveyor_info):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    elements = []
    styles = getSampleStyleSheet()
    
    centered_title_style = ParagraphStyle('DocTitleCentered', parent=styles['Heading1'], fontSize=16, textColor=colors.HexColor('#0284c7'), spaceAfter=2, fontName='Helvetica-Bold', alignment=1)
    centered_subtitle_style = ParagraphStyle('DocSubtitleCentered', parent=styles['Normal'], fontSize=8.5, textColor=colors.HexColor('#1e293b'), spaceAfter=10, fontName='Helvetica-Bold', alignment=1)
    section_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontSize=11, textColor=colors.HexColor('#0f172a'), spaceBefore=6, spaceAfter=4, fontName='Helvetica-Bold')

    if os.path.exists(LOGO_PATH):
        try:
            img = RLImage(LOGO_PATH, width=50, height=50)
            img.hAlign = 'CENTER'
            elements.append(img)
            elements.append(Spacer(1, 4))
        except Exception:
            pass

    elements.append(Paragraph("Valuwie Property Valuation & Taxation", centered_title_style))
    elements.append(Paragraph("property valuation and taxation in your hands", centered_subtitle_style))
    
    summary_data = [
        ["District", selected_district, "Land Size", f"{acres:.4f} Acres"],
        ["Block No.", f"Block {selected_block}", "Estate / Zone", selected_estate]
    ]
    t_summary = Table(summary_data, colWidths=[100, 160, 100, 180])
    t_summary.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#f1f5f9')),
        ('TEXTCOLOR', (0,0), (-1,-1), colors.HexColor('#0f172a')),
        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
        ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 8.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#94a3b8')),
    ]))
    elements.append(t_summary)
    elements.append(Spacer(1, 8))
    
    elements.append(Paragraph("Financial & Title Transfer Breakdown", section_style))
    
    fin_rows = [
        ["Cost Component", "Amount (UGX)"],
        ["Estimated Land Value", f"{land_value:,.2f}"],
        ["Building / Improvements", f"{building_val:,.2f}"],
        ["<b>Sub-total: Total Property Worth</b>", f"<b>{total_property:,.2f}</b>"],
        ["Government Stamp Duty (1.5%)", f"{stamp_duty:,.2f}"],
        ["Official Search Fee", f"{official_search:,.2f}"],
        ["<b>Sub-total: Government Statutory Dues</b>", f"<b>{govt_total:,.2f}</b>"]
    ]
    
    survey_sub_total = 0
    if include_charges:
        if include_survey and surveyor_fee_total > 0:
            fin_rows.append(["Professional Surveying Fee", f"{surveyor_fee_total:,.2f}"])
            survey_sub_total += surveyor_fee_total
        if include_mutation and title_creation_fee > 0:
            fin_rows.append(["New Title Creation Fee (Mutation)", f"{title_creation_fee:,.2f}"])
            survey_sub_total += title_creation_fee
            
        if survey_sub_total > 0:
            fin_rows.append([f"<b>Sub-total: Surveying & Titling Fees</b>", f"<b>{survey_sub_total:,.2f}</b>"])
            
        transfer_sub_total = 0
        if include_transfer and agent_transfer_fee > 0:
            fin_rows.append(["Agent / Handler Transfer Fee", f"{agent_transfer_fee:,.2f}"])
            transfer_sub_total += agent_transfer_fee
            fin_rows.append([f"<b>Sub-total: Agent & Transfer Services</b>", f"<b>{transfer_sub_total:,.2f}</b>"])
        
    fin_rows.append(["Total Acquisition Budget", f"{total_acquisition:,.2f}"])
    
    if include_charges:
        fin_rows.append([f"<b>Assigned Professional:</b> {assigned_surveyor_info.get('name', 'N/A')}", f"<b>Contact:</b> {assigned_surveyor_info.get('contacts', 'N/A')}"])

    table_paragraphs = []
    for row_idx, row in enumerate(fin_rows):
        row_cells = []
        for col_idx, cell in enumerate(row):
            if row_idx == 0:
                p_style = ParagraphStyle('THead', parent=styles['Normal'], fontName='Helvetica-Bold', textColor=colors.white, fontSize=8.5)
            elif row_idx == len(fin_rows) - (2 if include_charges else 1):
                p_style = ParagraphStyle('TFoot', parent=styles['Normal'], fontName='Helvetica-Bold', textColor=colors.HexColor('#0f172a'), fontSize=8.5)
            else:
                p_style = ParagraphStyle('TRow', parent=styles['Normal'], fontName='Helvetica', textColor=colors.HexColor('#0f172a'), fontSize=8.5)
            row_cells.append(Paragraph(str(cell), p_style))
        table_paragraphs.append(row_cells)

    t_fin = Table(table_paragraphs, colWidths=[340, 200])
    t_fin.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0284c7')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor('#e0f2fe')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#94a3b8')),
    ]))
    elements.append(t_fin)
    
    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# CENTERED HEADER WITH EMBEDDED TITLE RESET
# ==========================================
logo_img_tag = ""
if os.path.exists(LOGO_PATH):
    with open(LOGO_PATH, "rb") as f:
        encoded_logo = base64.b64encode(f.read()).decode()
        logo_img_tag = f'<img src="data:image/png;base64,{encoded_logo}" style="height: 96px; vertical-align: middle; margin-right: -4px;" />'

st.markdown(f"""
    <div class="brand-header-container">
        <a href="?reset_trigger=true" target="_self" style="text-decoration: none;">
            <div class="brand-flex" title="Click to reset Valuwie Engine">
                {logo_img_tag}
                <span class="valuwie-title">Valuwie</span>
            </div>
        </a>
        <div class="valuwie-subtitle">property valuation and taxation in your hands</div>
    </div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ==========================================
# TOP ROW: INPUTS & PROPERTY SUMMARY
# ==========================================
selected_district = "Luwero"
selected_block = "1"
selected_estate = "Unassigned"
default_rate = 0

top_col1, top_col2, top_col3_rates, top_col3_agent, top_col4 = st.columns([1.1, 1.1, 1.0, 1.0, 1.1])

with top_col1:
    render_header("Property Location", "location icon.png")
    if not df_blocks.empty:
        unique_districts = sorted([d for d in df_blocks['District'].dropna().unique() if str(d).strip().lower() not in ['mpigi', 'nan']])
        selected_district = st.selectbox("District / Area:", unique_districts)
        
        filtered_by_district = df_blocks[df_blocks['District'] == selected_district]
        
        unique_blocks = sorted(
            filtered_by_district['Block No'].dropna().astype(str).str.strip().unique(),
            key=lambda x: str(x)
        )
        
        default_block_idx = 0
        if str(st.session_state.get('clicked_block')) in unique_blocks:
            default_block_idx = unique_blocks.index(str(st.session_state.clicked_block))
            
        selected_block = st.selectbox("Block No:", unique_blocks, index=default_block_idx)
        
        filtered_by_block = filtered_by_district[filtered_by_district['Block No'].astype(str) == str(selected_block)]
        
        unique_estates = sorted(filtered_by_block['Estate'].dropna().unique()) if not filtered_by_block.empty else ["Unassigned"]
        selected_estate = st.selectbox("Village / Estate Name:", unique_estates)
        
        match_row = filtered_by_block[filtered_by_block['Estate'] == selected_estate] if not filtered_by_block.empty else pd.DataFrame()
        if not match_row.empty:
            val = match_row.iloc[0]['Baseline Value (UGX/Acre)']
            if pd.notna(val) and float(val) > 0:
                default_rate = int(val)

with top_col2:
    render_header("Land Size & Buildings", "Land icon.png")
    calculation_mode = st.radio("Size Mode:", ["Acres", "Dimensions"], horizontal=True)
    if calculation_mode == "Acres":
        acres = st.number_input("Land Size (Acres):", min_value=0.01, value=1.0, step=0.25)
    else:
        dim_unit = st.selectbox("Unit:", ["Meters", "Feet"])
        length = st.number_input("Length:", min_value=1.0, value=50.0)
        width = st.number_input("Width:", min_value=1.0, value=100.0)
        acres = (length * width) / 4046.86 if dim_unit == "Meters" else (length * width) / 43560
        st.caption(f"Computed: {acres:.3f} Acres")
    
    building_val = st.number_input("Building Val (UGX):", min_value=0, value=0, step=1000000)

with top_col3_rates:
    render_header("Valuation Rate", "valuation icon.png")
    applied_rate = st.number_input("Rate / Acre (UGX):", min_value=0, value=default_rate, step=1000000)
    st.info("Baseline market valuation rate per acre.")

with top_col3_agent:
    head_col, btn_toggle_col = st.columns([0.6, 0.4])
    with head_col:
        render_header("Professional Help", "surveyor icon.png")
    with btn_toggle_col:
        if st.session_state.prof_help_active:
            if st.button("✓ ACTIVE", key="prof_toggle_btn", type="primary"):
                st.session_state.prof_help_active = False
                st.rerun()
        else:
            if st.button("ENABLE", key="prof_toggle_btn"):
                st.session_state.prof_help_active = True
                st.rerun()

    surveyor_fee_total = 0
    title_creation_fee = 0
    agent_transfer_fee = 0
    
    base_survey_fee = 700000
    surplus_fee_per_acre = 250000
    subdivision_addon = 300000
    
    proceed_to_transfer = False
    need_mutation = False
    
    include_professional_charges = st.session_state.prof_help_active
    assigned_surveyor_info = {"name": "Ali Kisitu", "agency_name": "Datum Ventures", "reg_no": "N/A", "contacts": "N/A"}
    include_survey = False
    include_mutation = False
    include_transfer = False

    if include_professional_charges:
        help_type = st.selectbox("I need a:", ["-- Select Option --", "Surveyor", "Handler"])
        
        if help_type == "Surveyor":
            include_survey = True
            survey_scope = st.radio("Survey Scope:", ["Subdivision", "Boundary Opening"], horizontal=True)
            
            if survey_scope == "Boundary Opening":
                if acres <= 1.0:
                    surveyor_fee_total = base_survey_fee
                else:
                    surveyor_fee_total = base_survey_fee + ((acres - 1.0) * surplus_fee_per_acre)
                
                proceed_to_transfer = st.checkbox("Proceed to Transfer?", value=False)
                include_transfer = proceed_to_transfer
                
            elif survey_scope == "Subdivision":
                surveyor_fee_total = base_survey_fee + subdivision_addon
                if acres > 1.0:
                    surveyor_fee_total += (acres - 1.0) * surplus_fee_per_acre
                    
                need_mutation = st.checkbox("Create New Title (by Mutation)", value=False)
                include_mutation = need_mutation
                if need_mutation:
                    if acres < 3.0:
                        title_creation_fee = 1500000
                    elif acres < 100.0:
                        title_creation_fee = 2000000
                    else:
                        title_creation_fee = 5000000
                        
                proceed_to_transfer = st.checkbox("Transfer Title", value=False)
                include_transfer = proceed_to_transfer
                
        elif help_type == "Handler":
            handler_mode = st.radio("Handler Type:", ["JRJ Handling", "Transfer Handling"], horizontal=True)
            if handler_mode == "JRJ Handling":
                need_mutation = st.checkbox("Create New Title (by Mutation)", value=False)
                include_mutation = need_mutation
                if need_mutation:
                    if acres < 3.0:
                        title_creation_fee = 1500000
                    elif acres < 100.0:
                        title_creation_fee = 2000000
                    else:
                        title_creation_fee = 5000000
                proceed_to_transfer = st.checkbox("Transfer Title", value=False)
                include_transfer = proceed_to_transfer
            else:
                proceed_to_transfer = True
                include_transfer = True

        if help_type in ["Surveyor", "Handler"]:
            surveyor_options = []
            surveyor_mapping = {}
            if not df_agents.empty:
                for _, row in df_agents.iterrows():
                    name = str(row.get('Lead Agent', row.get('Agency Name', 'Surveyor'))).strip()
                    agency_name = str(row.get('Agency Name', 'Datum Ventures')).strip()
                    reg_no = str(row.get('Reg.No', row.get('Reg No', 'N/A'))).strip()
                    contacts = str(row.get('Contacts', row.get('Contact', row.get('Phone', 'N/A')))).strip()
                    
                    display_str = name
                    surveyor_options.append(display_str)
                    surveyor_mapping[display_str] = {
                        "name": name,
                        "agency_name": agency_name,
                        "reg_no": reg_no,
                        "contacts": contacts,
                        "display_name": name
                    }
            if not surveyor_options:
                surveyor_options = ["Ali Kisitu"]
                surveyor_mapping[surveyor_options[0]] = {"name": "Ali Kisitu", "agency_name": "Datum Ventures", "reg_no": "0000", "contacts": "N/A", "display_name": "Ali Kisitu"}

            selected_surveyor_str = st.selectbox("Assigned Professional:", surveyor_options)
            assigned_surveyor_info = surveyor_mapping.get(selected_surveyor_str, {"name": "Ali Kisitu", "agency_name": "Datum Ventures", "reg_no": "N/A", "contacts": "N/A", "display_name": "Ali Kisitu"})
            
            if proceed_to_transfer:
                if acres < 1.0:
                    agent_transfer_fee = 800000
                elif acres < 3.0:
                    agent_transfer_fee = 1500000
                elif acres <= 50.0:
                    agent_transfer_fee = 2500000
                else:
                    agent_transfer_fee = 3000000

        if surveyor_fee_total > 0:
            st.success(f"Survey Fee: UGX {surveyor_fee_total:,.0f}")
        if title_creation_fee > 0:
            st.info(f"Mutation Fee: UGX {title_creation_fee:,.0f}")
        if agent_transfer_fee > 0:
            st.info(f"Agent Transfer Fee: UGX {agent_transfer_fee:,.0f}")
    else:
        st.caption("Professional help currently inactive.")

# Pre-calculate financial metrics for report & display
land_value = applied_rate * acres
total_property = land_value + building_val
stamp_duty = total_property * 0.015
official_search = 20000
govt_total = stamp_duty + official_search

active_survey_fee = surveyor_fee_total if include_professional_charges else 0
active_mutation_fee = title_creation_fee if include_professional_charges else 0
active_transfer_fee = agent_transfer_fee if include_professional_charges else 0

total_acquisition = total_property + govt_total + active_survey_fee + active_mutation_fee + active_transfer_fee

with top_col4:
    render_header("Summary & Export", "report icon.png")
    st.markdown(f"""
    - **District:** {selected_district}
    - **Block No:** {selected_block}
    - **Estate:** {selected_estate}
    - **Land Size:** {acres:.2f} Acres
    - **Market Rate:** UGX {applied_rate:,.0f}
    """)
    
    pdf_data = create_pdf_report(selected_district, selected_block, selected_estate, acres, land_value, building_val, total_property, stamp_duty, official_search, govt_total, active_survey_fee, active_mutation_fee, active_transfer_fee, total_acquisition, include_professional_charges, include_survey, include_mutation, include_transfer, assigned_surveyor_info)
    
    st.download_button(
        label="GENERATE REPORT",
        data=pdf_data,
        file_name=f"Valuwie_Report_{selected_district}_Block_{selected_block}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

st.markdown("---")

# ==========================================
# BOTTOM ROW: MAP (LEFT) & FINANCIALS (RIGHT)
# ==========================================
map_col, fin_col = st.columns([1.3, 1.0])

with map_col:
    render_header("Interactive Block Map", "Map click.png")
    search_col, btn_col, layer_col = st.columns([1.2, 0.4, 0.9])
    with search_col:
        map_query = st.text_input("Search landmark:", placeholder="e.g. kasana market", label_visibility="collapsed")
    with btn_col:
        go_pressed = st.button("Find")
    with layer_col:
        map_style = st.selectbox("View:", ["Street", "Satellite"], label_visibility="collapsed")

    if go_pressed and map_query:
        found_coords = None
        found_name = ""
        query_clean = map_query.lower().strip()
        try:
            encoded_q = urllib.parse.quote(query_clean + ", Uganda")
            url = f"https://nominatim.openstreetmap.org/search?q={encoded_q}&format=json&limit=1"
            req = urllib.request.Request(url, headers={'User-Agent': 'ValuwieGIS/1.0'})
            with urllib.request.urlopen(req, timeout=2) as resp:
                res_data = json.loads(resp.read().decode())
                if res_data:
                    found_coords = (float(res_data[0]['lat']), float(res_data[0]['lon']))
                    found_name = res_data[0]['display_name']
        except Exception:
            pass

        if not found_coords:
            closest_matches = difflib.get_close_matches(query_clean, LUWERO_DIRECTORY.keys(), n=1, cutoff=0.25)
            if closest_matches:
                matched_key = closest_matches[0]
                found_coords = LUWERO_DIRECTORY[matched_key]
                found_name = matched_key.title()

        if found_coords and gdf_blocks is not None:
            pt = Point(found_coords[1], found_coords[0])
            matched_poly = gdf_blocks[gdf_blocks.contains(pt)] if shape_block_col else pd.DataFrame()
            if not matched_poly.empty and shape_block_col:
                matched_block_no = str(matched_poly.iloc[0][shape_block_col])
                st.session_state.clicked_block = matched_block_no
                st.session_state.searched_coords = found_coords
                st.session_state.searched_name = found_name
                st.success(f"Located Block {matched_block_no}!")
                st.rerun()

    if gdf_blocks is not None:
        if shape_block_col:
            block_gdf = pd.DataFrame()
            if 'id_2' in gdf_blocks.columns:
                block_gdf = gdf_blocks[
                    (gdf_blocks[shape_block_col].astype(str).str.strip() == str(selected_block).strip()) & 
                    (gdf_blocks['id_2'].astype(str).str.strip().str.title() == str(selected_district).strip().title())
                ]
            if block_gdf.empty:
                block_gdf = gdf_blocks[gdf_blocks[shape_block_col].astype(str).str.strip() == str(selected_block).strip()]
            target_gdf = block_gdf if not block_gdf.empty else gdf_blocks
        else:
            target_gdf = gdf_blocks

        bounds = target_gdf.total_bounds
        center_lat = (bounds[1] + bounds[3]) / 2
        center_lon = (bounds[0] + bounds[2]) / 2
        map_center = st.session_state.searched_coords if st.session_state.searched_coords else [center_lat, center_lon]
        
        tile_url = 'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}' if "Satellite" in map_style else 'OpenStreetMap'
        tile_attr = 'Tiles &copy; Esri' if "Satellite" in map_style else None

        m = folium.Map(location=map_center, zoom_start=14, tiles=tile_url, attr=tile_attr)
        if st.session_state.searched_coords:
            folium.Marker(location=st.session_state.searched_coords, popup=st.session_state.searched_name, icon=folium.Icon(color='red')).add_to(m)

        folium.GeoJson(
            target_gdf,
            style_function=lambda x: {'fillColor': '#0284c7', 'color': '#0f172a', 'weight': 2, 'fillOpacity': 0.4},
            tooltip=folium.GeoJsonTooltip(fields=[shape_block_col] if shape_block_col else [])
        ).add_to(m)
        
        if not st.session_state.searched_coords:
            m.fit_bounds([[bounds[1], bounds[0]], [bounds[3], bounds[2]]])
        
        map_output = st_folium(m, use_container_width=True, height=380, key="sketch_map")
        
        if map_output and map_output.get("last_clicked"):
            click_lat = map_output["last_clicked"]["lat"]
            click_lon = map_output["last_clicked"]["lng"]
            point = Point(click_lon, click_lat)
            
            if 'id_2' in gdf_blocks.columns:
                district_subset = gdf_blocks[gdf_blocks['id_2'].astype(str).str.strip().str.title() == str(selected_district).strip().title()]
                matched_row = district_subset[district_subset.contains(point)]
                if matched_row.empty:
                    matched_row = gdf_blocks[gdf_blocks.contains(point)]
            else:
                matched_row = gdf_blocks[gdf_blocks.contains(point)]

            if not matched_row.empty and shape_block_col:
                clicked_val = str(matched_row.iloc[0][shape_block_col])
                if st.session_state.clicked_block != clicked_val:
                    st.session_state.clicked_block = clicked_val
                    st.rerun()

with fin_col:
    render_header("Financial Breakdown", "valuation icon.png")
    
    cost_items = ["Land Value", "Building Value", "Stamp Duty (1.5%)", "Official Search Fee"]
    cost_amounts = [f"{land_value:,.0f}", f"{building_val:,.0f}", f"{stamp_duty:,.0f}", f"{official_search:,.0f}"]
    
    if active_survey_fee > 0:
        cost_items.append("Surveying Fee")
        cost_amounts.append(f"{active_survey_fee:,.0f}")
    if active_mutation_fee > 0:
        cost_items.append("New Title Creation Fee")
        cost_amounts.append(f"{active_mutation_fee:,.0f}")
    if active_transfer_fee > 0:
        cost_items.append("Agent Transfer Fee")
        cost_amounts.append(f"{active_transfer_fee:,.0f}")
        
    metrics_df = pd.DataFrame({
        "Cost Item": cost_items,
        "Amount (UGX)": cost_amounts
    })
    st.table(metrics_df)
    
    render_header("Transfer Title Cost & Total Budget", "report icon.png")
    st.info(f"**Total Estimated Acquisition Budget:** \n### UGX {total_acquisition:,.2f}")

st.markdown("---")

# ==========================================
# ABOUT CREATORS SECTION (PLAIN TOGGLE BUTTON)
# ==========================================
if st.button("ℹ️ ABOUT CREATORS", use_container_width=True):
    st.session_state.show_about = not st.session_state.show_about

if st.session_state.show_about:
    st.markdown("""
    ### Project Context
    This application was developed as a practical assignment for **Group 16** in **Computer Application and Programming (GEO 2102)** at **Ndejje University**, 2026, Semester One, Year 2.
    """)
    
    st.markdown("### Group 16 members")
    if not df_agents.empty:
        exclude_cols = [c for c in df_agents.columns if any(term in c.lower() for term in ['agency', 'title delivery', 'fee'])]
        filtered_agents_df = df_agents.drop(columns=exclude_cols).loc[:, ~df_agents.drop(columns=exclude_cols).columns.duplicated()]
        
        if 'Lead Agent' in filtered_agents_df.columns:
            filtered_agents_df = filtered_agents_df.rename(columns={'Lead Agent': 'Student Name'})
            
        filtered_agents_df.index = range(1, len(filtered_agents_df) + 1)
        st.dataframe(filtered_agents_df, use_container_width=True)
    else:
        st.info("No agents directory found in database.")