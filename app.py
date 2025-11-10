# app.py
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import io, os, re, requests, csv
from typing import List, Dict, Any, Tuple
from copy import deepcopy

# --------- Constants (Code only) ---------
TEMPLATE_FILE = "input-template.pptx"
MASTER_URL = "https://docs.google.com/spreadsheets/d/1blj42SbFpszWGyOrDOUwyPDJr9K1NGpTMX6eZTbt_P4/edit?gid=1152340088#gid=1152340088"
MAPPING_URL = "https://docs.google.com/spreadsheets/d/1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk/edit?gid=1056617222#gid=1056617222"

PCON_SKIPROWS = 2
IDX_SHORT, IDX_VARIANT, IDX_ARTICLE, IDX_QTY = 2, 4, 17, 30

# Template tags
TAG_SETTINGNAME = "{{SETTINGNAME}}"
TAG_PRODUCTS_LIST = "{{ProductsinSettingList}}" 
TAG_RENDERING = "{{Rendering}}"
TAG_LINEDRAWING = "{{Linedrawing}}"
OVERVIEW_TITLE = "OVERVIEW"

PACKSHOT_TAGS = [f"{{{{ProductPackshot{i}}}}}" for i in range(1, 13)]
PROD_DESC_TAGS = [f"{{{{PRODUCT DESCRIPTION {i}}}}}" for i in range(1, 13)]
OVERVIEW_TAGS = [f"{{{{Rendering{i}}}}}" for i in range(1, 13)]

# ----------------------------------------------------------------------
# --------- Utils ------------------------------------------------------
# ----------------------------------------------------------------------

def resolve_gsheet_to_csv(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url or "")
    if not m: return url
    sheet = m.group(1)
    gid_m = re.search(r"[#?&]gid=(\d+)", url)
    gid = gid_m.group(1) if gid_m else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet}/export?format=csv&gid={gid}"

def _norm_placeholder_text(s: str) -> str:
    if s is None: return ""
    s = str(s)
    s = re.sub(r"\{\{\s*", "{{", s)
    s = re.sub(r"\s*\}\}", "}}", s)
    s = re.sub(r"\{\{(\s*)([^}]*?)(\s*)\}\}", lambda m: "{{" + re.sub(r"\s+", "", m.group(2)) + "}}", s)
    return s.lower()

def _norm_tag(tag: str) -> str:
    return _norm_placeholder_text(tag)

def find_shape_by_placeholder(slide, tag: str):
    want = _norm_tag(tag)
    
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False) and shp.text_frame:
            txt = shp.text_frame.text or ""
            if want in _norm_placeholder_text(txt):
                return shp
        
        try:
            if shp.name and want in _norm_tag(shp.name):
                return shp
        except Exception:
            pass
                
    for shp in getattr(slide, "placeholders", []):
        if getattr(shp, "has_text_frame", False) and shp.text_frame:
            txt = shp.text_frame.text or ""
            if want in _norm_placeholder_text(txt):
                return shp
    return None

def set_text_preserve_style(shape, text: str):
    """
    Inserts text and preserves style from the first 'run'.
    Handles nested tags (e.g. {{SETTINGNAME}}) and multiline input.
    All text is converted to UPPERCASE.
    """
    if not shape or not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    
    final_text_content = text.upper()
    
    current_text_upper = tf.text.upper()
    if TAG_SETTINGNAME.upper() in current_text_upper:
        final_text_content = current_text_upper.replace(TAG_SETTINGNAME.upper(), text.upper())

    font_name = font_size = font_bold = None
    if tf.paragraphs and tf.paragraphs[0].runs:
        r0 = tf.paragraphs[0].runs[0]
        font_name, font_size, font_bold = r0.font.name, r0.font.size, r0.font.bold
        
    while tf.paragraphs:
        p = tf.paragraphs[0]
        for r in list(p.runs): r.text = ""
        try: tf._element.remove(p._p)
        except Exception: break
            
    lines = final_text_content.split('\n')
    p = tf.add_paragraph() 
    
    for i, line in enumerate(lines):
        if i > 0:
            p = tf.add_paragraph()
        
        run = p.add_run()
        run.text = line.strip()
        if font_name: run.font.name = font_name
        if font_size: run.font.size = font_size
        if font_bold is not None: run.font.bold = font_bold
            
    tf.word_wrap = True

def replace_image_by_tag(slide, tag: str, img_bytes: bytes):
    if not img_bytes: return
    ph = find_shape_by_placeholder(slide, tag)
    if not ph: return
    left, top, w, h = ph.left, ph.top, ph.width, ph.height
    
    img_stream = io.BytesIO(img_bytes)
    
    try:
        im = Image.open(img_stream)
        img_w, img_h = im.size
        aspect_ratio = img_w / img_h
        img_stream.seek(0)
    except Exception:
        aspect_ratio = 1.0
        img_stream.seek(0)
    
    if (w / aspect_ratio) <= h:
        new_h = w / aspect_ratio
        new_w = w
    else:
        new_w = h * aspect_ratio
        new_h = h

    new_left = left + (w - new_w) / 2
    new_top = top + (h - new_h) / 2
    
    try: ph.element.getparent().remove(ph.element)
    except Exception: pass
    
    slide.shapes.add_picture(img_stream, new_left, new_top, width=new_w, height=new_h)


def duplicate_slide(prs: Presentation, slide):
    new_slide = prs.slides.add_slide(slide.slide_layout)
    for shp in list(new_slide.shapes):
        sp = shp.element
        sp.getparent().remove(sp)
    for shp in slide.shapes:
        new_slide.shapes._spTree.append(deepcopy(shp._element))
    return new_slide

def remove_slide(prs: Presentation, index: int):
    if index < 0 or index >= len(prs.slides._sldIdLst):
        return
        
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]

def find_first_slide_with_tag(prs: Presentation, tag: str) -> Tuple[Any, int]:
    want = _norm_tag(tag)
    for i, sl in enumerate(prs.slides):
        for shp in sl.shapes:
            if getattr(shp, "has_text_frame", False) and shp.text_frame:
                if want in _norm_placeholder_text(shp.text_frame.text):
                    return sl, i
    return None, -1

def blank_layout(prs: Presentation):
    for ly in prs.slide_layouts:
        if ly.name and "blank" in ly.name.lower():
            return ly
    return prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[1]

# --- Template Creator ---
def create_simple_template_pptx() -> bytes:
    """Creates a simple, functional PowerPoint template with all required placeholders."""
    prs = Presentation()
    
    # Use blank layout
    blank_layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[1]
    
    # --- OVERVIEW Slide ---
    s_overview = prs.slides.add_slide(blank_layout)
    
    # Title
    title_box = s_overview.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    title_box.text_frame.text = OVERVIEW_TITLE
    title_box.text_frame.paragraphs[0].runs[0].font.bold = True
    
    # Add 12 Rendering placeholders
    x, y = Inches(0.5), Inches(0.8)
    w, h = Inches(2.3), Inches(2.3) # Adjusted size for better fit
    for i in range(12):
        col, row = i % 4, i // 4
        
        # Create a shape (could be Picture Placeholder or just a text box)
        tx = s_overview.shapes.add_textbox(x + col * Inches(2.5), y + row * Inches(2.5), w, h)
        tx.text_frame.text = OVERVIEW_TAGS[i]
        tx.line.color.rgb = type('RGB', (object,), {'value': bytes([255, 0, 0])})() # Red border for visibility

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


# ----------------------------------------------------------------------
# --------- Data-loaders & Lookups (Unchanged) -------------------------
# ----------------------------------------------------------------------

@st.cache_data
def load_master() -> pd.DataFrame:
    url = resolve_gsheet_to_csv(MASTER_URL)
    r = requests.get(url, timeout=20); r.raise_for_status()
    df = pd.read_csv(io.BytesIO(r.content))
    def norm(s): return re.sub(r"[\s_.-]+","",str(s
