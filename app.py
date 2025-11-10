import io
from pptx import Presentation
from pptx.util import Inches
from copy import deepcopy

# --- Konstanter brugt i den primære app ---
TAG_SETTINGNAME = "{{SETTINGNAME}}"
TAG_RENDERING = "{{Rendering}}"
TAG_LINEDRAWING = "{{Linedrawing}}"
OVERVIEW_TITLE = "OVERVIEW"
PACKSHOT_TAGS = [f"{{{{ProductPackshot{i}}}}}" for i in range(1, 13)]
PROD_DESC_TAGS = [f"{{{{PRODUCT DESCRIPTION {i}}}}}" for i in range(1, 13)]
OVERVIEW_TAGS = [f"{{{{Rendering{i}}}}}" for i in range(1, 13)]

def remove_slide(prs, index: int):
    """Fjerner et slide fra præsentationen ved hjælp af dets index."""
    if index < 0 or index >= len(prs.slides._sldIdLst):
        return
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]

def create_simple_template_pptx() -> bytes:
    """Opretter en enkel, funktionel PowerPoint-skabelon med alle de nødvendige pladsholdere."""
    prs = Presentation()
    
    # Brug blank layout
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[1]
    
    # --- OVERVIEW Slide ---
    s_overview = prs.slides.add_slide(blank_layout)
    
    # Title
    s_overview.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text_frame.text = OVERVIEW_TITLE
    
    # Add 12 Rendering placeholders
    x, y = Inches(0.5), Inches(0.8)
    w, h = Inches(2.3), Inches(2.3) # Adjusted size for better fit
    for i in range(12):
        col, row = i % 4, i // 4
        
        # Opret en tekstboks som pladsholder
        tx = s_overview.shapes.add_textbox(x + col * Inches(2.5), y + row * Inches(2.5), w, h)
        tx.text_frame.text = OVERVIEW_TAGS[i]
        # Tilføj en rød ramme for synlighed i template
        try:
            tx.line.color.rgb = type('RGB', (object,), {'value': bytes([255, 0, 0])})() 
        except Exception:
            pass # Nogle miljøer understøtter ikke denne styling

    # --- SETTING Slide ---
    s_setting = prs.slides.add_slide(blank_layout)
    
    # Setting Name Title (SHOP THE LOOK - {{SETTINGNAME}})
    setting_title_box = s_setting.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    setting_title_box.text_frame.text = f"SHOP THE LOOK - {TAG_SETTINGNAME}"
    setting_title_box.text_frame.paragraphs[0].runs[0].font.bold = True
    
    # Main Rendering ({{Rendering}})
    s_setting.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(4.5), Inches(4)).text_frame.text = TAG_RENDERING
    
    # Line Drawing ({{Linedrawing}})
    s_setting.shapes.add_textbox(Inches(5.5), Inches(1.0), Inches(4.5), Inches(4)).text_frame.text = TAG_LINEDRAWING
    
    # 12 Product Slots
    x, y = Inches(0.5), Inches(5.5)
    w_pack, w_desc = Inches(0.5), Inches(1.9)
    h_slot = Inches(0.4)
    
    for i in range(12):
        slot_x = x + (i // 4) * Inches(3.2) # 4 slots per column, 3 columns total
        slot_y = y + (i % 4) * h_slot
        
        # Packshot placeholder
        pack_box = s_setting.shapes.add_textbox(slot_x, slot_y, w_pack, h_slot)
        pack_box.text_frame.text = PACKSHOT_TAGS[i]
        
        # Description placeholder
        desc_box = s_setting.shapes.add_textbox(slot_x + w_pack + Inches(0.1), slot_y, w_desc, h_slot)
        desc_box.text_frame.text = PROD_DESC_TAGS[i]

    # Clean up default blank slide if it exists
    if len(prs.slides) > 2:
        remove_slide(prs, 0)
    
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()

# Streamlit UI for downloading the template
import streamlit as st

st.title("PowerPoint Template Downloader")
st.markdown("---")
st.info("Brug knappen nedenfor til at downloade den simple, Python-kompatible skabelon. Gem den som **input-template.pptx** i samme mappe som din hovedapp.")

template_bytes = create_simple_template_pptx()

st.download_button(
    "Download input-template.pptx",
    data=template_bytes,
    file_name="input-template.pptx",
    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
)
