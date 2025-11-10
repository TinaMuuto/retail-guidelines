# app.py
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import PP_PLACEHOLDER
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

# GLOBAL TIMEOUT FOR REQUESTS
REQUEST_TIMEOUT = 60 # Increased from 20 to 60 seconds to avoid ReadTimeout error

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

def find_shape_by_placeholder(slide, tag: str, find_placeholder_type=False):
    """
    Finds a shape/placeholder by its text content (tag). 
    If find_placeholder_type is True, it prioritizes returning only true placeholders.
    """
    want = _norm_tag(tag)
    
    # 1. Check true placeholders first (necessary for picture insertion)
    for ph in getattr(slide, "placeholders", []):
        try:
            if ph.has_text_frame and ph.text and want in _norm_placeholder_text(ph.text):
                return ph
        except AttributeError:
            continue
            
    if find_placeholder_type:
        return None 

    # 2. Check all other shapes (robust for non-standard text boxes and named shapes)
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
    
    img_stream = io.BytesIO(img_bytes)
    
    # 1. Try inserting into a true placeholder (stable method)
    ph = find_shape_by_placeholder(slide, tag, find_placeholder_type=True) 
    
    if ph:
        try:
            if getattr(ph.placeholder_format, 'type', None) in [PP_PLACEHOLDER.PICTURE, PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT]:
                ph.insert_picture(img_stream)
                img_stream.seek(0)
                return
        except Exception:
            pass # Fallback to aggressive replacement
    
    # 2. Aggressive delete-and-reinsert method (fallback)
    
    ph = find_shape_by_placeholder(slide, tag)
    if not ph: return
    
    left, top, w, h = ph.left, ph.top, ph.width, ph.height
    
    try:
        im = Image.open(img_stream)
        img_w, img_h = im.size
        aspect_ratio = img_w / img_h
        img_stream.seek(0)
    except Exception:
        aspect_ratio = 1.0
        img_stream.seek(0)
    
    # Calculate adjusted dimensions to PRESERVE ASPECT RATIO (fit into placeholder)
    if (w / aspect_ratio) <= h:
        new_h = w / aspect_ratio
        new_w = w
    else:
        new_w = h * aspect_ratio
        new_h = h

    new_left = left + (w - new_w) / 2
    new_top = top + (h - new_h) / 2
    
    # Remove the old shape element (aggressive cleanup)
    try: ph.element.getparent().remove(ph.element)
    except Exception: pass
    
    # Insert the new picture
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
    return prs.slide_layouts[6] if len(prs.slides) > 6 else prs.slide_layouts[1]

# --- Template Creator (For UI Download) ---
def create_simple_template_pptx() -> bytes:
    """Creates a simple, functional PowerPoint template with all required placeholders."""
    prs = Presentation()
    
    blank_layout = prs.slide_layouts[6] if len(prs.slides) > 6 else prs.slide_layouts[1]
    
    # --- OVERVIEW Slide ---
    s_overview = prs.slides.add_slide(blank_layout)
    s_overview.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5)).text_frame.text = OVERVIEW_TITLE
    
    x, y = Inches(0.5), Inches(0.8)
    w, h = Inches(2.3), Inches(2.3)
    for i in range(12):
        col, row = i % 4, i // 4
        tx = s_overview.shapes.add_textbox(x + col * Inches(2.5), y + row * Inches(2.5), w, h)
        tx.text_frame.text = OVERVIEW_TAGS[i]

    # --- SETTING Slide ---
    s_setting = prs.slides.add_slide(blank_layout)
    setting_title_box = s_setting.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.5))
    setting_title_box.text_frame.text = f"SHOP THE LOOK - {TAG_SETTINGNAME}"
    setting_title_box.text_frame.paragraphs[0].runs[0].font.bold = True
    
    s_setting.shapes.add_textbox(Inches(0.5), Inches(1.0), Inches(4.5), Inches(4)).text_frame.text = TAG_RENDERING
    s_setting.shapes.add_textbox(Inches(5.5), Inches(1.0), Inches(4.5), Inches(4)).text_frame.text = TAG_LINEDRAWING
    
    x, y = Inches(0.5), Inches(5.5)
    w_pack, w_desc = Inches(0.5), Inches(1.9)
    h_slot = Inches(0.4)
    
    for i in range(12):
        slot_x = x + (i // 4) * Inches(3.2)
        slot_y = y + (i % 4) * h_slot
        
        pack_box = s_setting.shapes.add_textbox(slot_x, slot_y, w_pack, h_slot)
        pack_box.text_frame.text = PACKSHOT_TAGS[i]
        
        desc_box = s_setting.shapes.add_textbox(slot_x + w_pack + Inches(0.1), slot_y, w_desc, h_slot)
        desc_box.text_frame.text = PROD_DESC_TAGS[i]

    if len(prs.slides) > 2:
        remove_slide(prs, 0)
    
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ----------------------------------------------------------------------
# --------- Data-loaders & Lookups -------------------------------------
# ----------------------------------------------------------------------

@st.cache_data
def load_master() -> pd.DataFrame:
    url = resolve_gsheet_to_csv(MASTER_URL)
    r = requests.get(url, timeout=REQUEST_TIMEOUT); r.raise_for_status()
    df = pd.read_csv(io.BytesIO(r.content))
    def norm(s): return re.sub(r"[\s_.-]+","",str(s).strip().lower())
    norm_map = {norm(c): c for c in df.columns}
    def mapcol(canon, alts):
        for a in [canon] + alts:
            if norm(a) in norm_map: return norm_map[norm(a)]
        return None
    col_item = mapcol("ITEM NO.", ["Item No.","ITEM","SKU","Item Number","ItemNo","ITEM_NO"])
    col_img = mapcol("IMAGE URL", ["Image URL","Image Link","Picture URL","Packshot URL","ImageURL","IMAGE DOWNLOAD LINK","Image"])
    if not col_item or not col_img:
        raise ValueError("Master is missing columns: ITEM NO. and/or IMAGE URL (IMAGE DOWNLOAD LINK is accepted).")
    out = df.rename(columns={col_item:"ITEM NO.", col_img:"IMAGE URL"})[["ITEM NO.","IMAGE URL"]]
    for c in out.columns: out[c] = out[c].astype(str).str.strip()
    return out

@st.cache_data
def load_mapping() -> pd.DataFrame:
    url = resolve_gsheet_to_csv(MAPPING_URL)
    r = requests.get(url, timeout=REQUEST_TIMEOUT); r.raise_for_status()
    df = pd.read_csv(io.BytesIO(r.content))
    def norm(s): return re.sub(r"[\s_.-]+","",str(s).strip().lower())
    norm_map = {norm(c): c for c in df.columns}
    def mapcol(canon, alts):
        for a in [canon]+alts:
            if norm(a) in norm_map: return norm_map[norm(a)]
        return None
    col_old = mapcol("OLD Item-variant", ["OLD Item variant","OLD ITEM NO.","Old Item","OLD_ITEM_VARIANT"])
    col_new = mapcol("New Item No.", ["New Item Number","NEW ITEM NO.","NEW_ITEM_NO"])
    col_desc = mapcol("Description", ["Product Description","DESC","NAME","DESCRIPTION"])
    if not col_old or not col_new or not col_desc:
        raise ValueError("Mapping is missing columns: OLD Item-variant, New Item No., Description.")
    out = df.rename(columns={
        col_old:"OLD Item-variant",
        col_new:"New Item No.",
        col_desc:"Description"
    })[["OLD Item-variant","New Item No.","Description"]]
    for c in out.columns: out[c] = out[c].astype(str).str.strip()
    return out

def _try_read_csv(fileobj, **kwargs):
    fileobj.seek(0)
    return pd.read_csv(fileobj, **kwargs)

def pcon_from_csv(uploaded_file) -> pd.DataFrame:
    attempts = [
        {"sep": ";", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "utf-8"},
        {"sep": ";", "encoding": "latin-1"},
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ",", "encoding": "utf-8"},
        {"sep": ",", "encoding": "latin-1"},
        {"sep": None, "engine": "python", "encoding": "utf-8-sig"},
        {"sep": None, "engine": "python", "encoding": "utf-8"},
        {"sep": None, "engine": "python", "encoding": "latin-1"},
    ]
    last_err, df = None, None
    need = max(IDX_SHORT, IDX_VARIANT, IDX_ARTICLE, IDX_QTY)
    for cfg in attempts:
        try:
            df = _try_read_csv(uploaded_file, header=None, skiprows=PCON_SKIPROWS,
                               on_bad_lines="skip", quoting=csv.QUOTE_MINIMAL, **cfg)
            if df.shape[1] <= need:
                last_err = ValueError(f"Too few columns with cfg={cfg}, shape={df.shape}")
                df = None; continue
            break
        except Exception as e:
            last_err = e; df = None
    if df is None:
        raise ValueError(f"Could not parse pCon CSV. Last error: {last_err}")
    sub = df.iloc[:, [IDX_SHORT, IDX_VARIANT, IDX_ARTICLE, IDX_QTY]].copy()
    sub.columns = ["SHORT_TEXT","VARIANT_TEXT","ARTICLE_NO","QUANTITY"]
    sub["ARTICLE_NO"]    = sub["ARTICLE_NO"].astype(str).str.strip()
    sub["SHORT_TEXT"]    = sub["SHORT_TEXT"].astype(str).str.strip()
    sub["VARIANT_TEXT"] = sub["VARIANT_TEXT"].astype(str).str.strip()
    sub["QUANTITY"]      = pd.to_numeric(sub["QUANTITY"], errors="coerce").fillna(1).astype(int)
    sub = sub[sub["ARTICLE_NO"].ne("")]
    if sub.empty:
        raise ValueError("pCon CSV was read, but contained no valid rows with ARTICLE_NO.")
    return sub

def fallback_key(article: str) -> str:
    base = re.sub(r"^SPECIAL-", "", str(article), flags=re.I)
    return base.split("-")[0].strip()

def packshot_lookup(master_df: pd.DataFrame, article: str) -> str:
    hit = master_df.loc[master_df["ITEM NO."] == article, "IMAGE URL"]
    if not hit.empty: return str(hit.iloc[0])
    base = fallback_key(article)
    hit = master_df.loc[master_df["ITEM NO."].apply(fallback_key) == base, "IMAGE URL"]
    return str(hit.iloc[0]) if not hit.empty else ""

def new_item_lookup(map_df: pd.DataFrame, article: str) -> str:
    hit = map_df.loc[map_df["OLD Item-variant"] == article, "New Item No."]
    if not hit.empty: return str(hit.iloc[0])
    base = fallback_key(article)
    hit = map_df.loc[map_df["OLD Item-variant"].apply(fallback_key) == base, "New Item No."]
    return str(hit.iloc[0]) if not hit.empty else ""

def mapping_description(map_df: pd.DataFrame, article: str) -> str:
    hit = map_df.loc[map_df["OLD Item-variant"] == article, "Description"]
    if not hit.empty: return str(hit.iloc[0])
    base = fallback_key(article)
    hit = map_df.loc[map_df["OLD Item-variant"].apply(fallback_key) == base, "Description"]
    return str(hit.iloc[0]) if not hit.empty else ""

# --------- Image Utilities ---------
@st.cache_data(ttl=3600)
def fetch_image(url: str) -> bytes | None:
    if not url or not url.startswith("http"): return None
    try:
        r = requests.get(url, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        
        content_type = r.headers.get("Content-Type","").lower()
        if content_type.startswith("text/") and "image" not in content_type:
             return None
             
        return r.content
    except requests.exceptions.RequestException:
        return None

def preprocess(img: bytes, max_side=1400, quality=85) -> bytes:
    try:
        im = Image.open(io.BytesIO(img))
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGB")
            
        if max(im.size) > max_side:
            ratio = min(max_side/im.width, max_side/im.height)
            im = im.resize((int(im.width*ratio), int(im.height*ratio)), Image.Resampling.LANCZOS)
            
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=85, optimize=True) 
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        return img 

# --------- PPT-byggesten CORE ---------

def add_products_table_on_blank(prs: Presentation, title: str, rows: List[List[str]]):
    """Adds a new slide with product list as an unformatted table."""
    s = prs.slides.add_slide(blank_layout(prs)) 
    
    # Title textbox
    left, top, width, height = Inches(0.6), Inches(0.3), Inches(9.2), Inches(0.6)
    tx_shape = s.shapes.add_textbox(left, top, width, height)
    set_text_preserve_style(tx_shape, title.upper()) 

    headers = ["Quantity", "Description", "Article No. / New Item No."]
    data = [headers] + rows
    left, top, width, height = Inches(0.6), Inches(1.2), Inches(9.2), Inches(5.5)
    
    try:
        tbl_shape = s.shapes.add_table(rows=len(data), cols=3, left=left, top=top, width=width, height=height)
    except Exception:
        return

    tbl = tbl_shape.table
    
    for r_i, row in enumerate(data):
        for c_i, val in enumerate(row):
            cell = tbl.cell(r_i, c_i)
            cell.text = str(val).upper() 
            for p in cell.text_frame.paragraphs:
                for run in p.runs: 
                    run.font.size = Pt(12) 
    return s


def build_presentation(master_df: pd.DataFrame,
                       mapping_df: pd.DataFrame,
                       groups: List[Dict[str, Any]],
                       overview_renderings: List[bytes]) -> bytes:

    prs = Presentation(TEMPLATE_FILE)

    setting_tpl, setting_idx = find_first_slide_with_tag(prs, TAG_SETTINGNAME)
    overview_tpl, overview_idx = find_first_slide_with_tag(prs, OVERVIEW_TITLE)
    
    # 2. OVERVIEW SLIDE GENERATION
    if overview_renderings:
        if overview_tpl is not None:
            s = duplicate_slide(prs, overview_tpl)
            for i, rb in enumerate(overview_renderings[:12]):
                replace_image_by_tag(s, OVERVIEW_TAGS[i], preprocess(rb))
        else:
            s = prs.slides.add_slide(blank_layout(prs))
            tx = s.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(9), Inches(0.6))
            set_text_preserve_style(tx, OVERVIEW_TITLE)
            cols, rows = 4, 3
            cell_w, cell_h = Inches(2.2), Inches(2.0)
            x0, y0 = Inches(0.5), Inches(1.2)
            for i, rb in enumerate(overview_renderings[:12]):
                s.shapes.add_picture(io.BytesIO(preprocess(rb)),
                                             x0 + (i % cols)*cell_w, y0 + (i // cols)*cell_h,
                                             width=cell_w, height=cell_h)

    # 3. SETTING SLIDE GENERATION
    for gi, g in enumerate(groups, 1):
        
        base = setting_tpl if setting_tpl else prs.slides[0]
        s = duplicate_slide(prs, base)
        
        # A. Populate SETTINGNAME (including "SHOP THE LOOK - {{SETTINGNAME}}")
        setting_title_shape = find_shape_by_placeholder(s, TAG_SETTINGNAME)
        if setting_title_shape:
            set_text_preserve_style(setting_title_shape, g["name"]) 

        # B. Rendering, Linedrawing, Packshots, Product Descriptions
        if g.get("rendering_bytes"):
            replace_image_by_tag(s, TAG_RENDERING, preprocess(g["rendering_bytes"]))
        if g.get("linedrawing_bytes"):
            replace_image_by_tag(s, TAG_LINEDRAWING, preprocess(g["linedrawing_bytes"]))

        for idx, it in enumerate(g["items"][:12]):
            # Packshot insertion
            if it.get("packshot_url"):
                raw = fetch_image(it["packshot_url"])
                if raw:
                    replace_image_by_tag(s, PACKSHOT_TAGS[idx], preprocess(raw))
                    
            # Product Description insertion
            desc_tag_shape = find_shape_by_placeholder(s, PROD_DESC_TAGS[idx])
            if desc_tag_shape:
                set_text_preserve_style(desc_tag_shape, it.get("desc_text",""))

        # C. Product Overview: Add a separate slide with a table
        rows = []
        for it in g["items"]:
            id_combo = it["article_no"]
            if it.get("new_item_no"): id_combo = f'{it["article_no"]} / {it["new_item_no"]}'
            rows.append([it["qty"], it["desc_text"], id_combo])
        add_products_table_on_blank(prs, f"PRODUCTS – {g['name']}", rows)

    # 4. REMOVE ORIGINAL TEMPLATE SLIDES
    indices_to_remove = []
    if overview_tpl and overview_idx is not None:
        indices_to_remove.append(overview_idx)
    if setting_tpl and setting_idx is not None:
        indices_to_remove.append(setting_idx)
        
    indices_to_remove = sorted(list(set(indices_to_remove)), reverse=True)
    
    for idx in indices_to_remove:
        remove_slide(prs, idx)


    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# --------- UI ---------
def main():
    if 'uploaded_file_names' not in st.session_state:
        st.session_state.uploaded_file_names = []

    st.set_page_config(page_title="Muuto PPT Generator", layout="wide")
    st.title("Muuto PPT Generator")
    
    st.markdown(
        """
        This tool automatically generates a **Shop the Look** PowerPoint presentation
        based on your pCon CSV exports and corresponding image files.
        It groups CSVs, Renderings, and Linedrawings by their file name prefix (e.g., 'Dining 01') 
        to create individual setting slides.
        """
    )
    st.markdown("---")
    
    # --- Template Setup Section ---
    st.subheader("Template Setup")
    if not os.path.exists(TEMPLATE_FILE):
        st.warning(f"Template file '{TEMPLATE_FILE}' not found. **Please use the download button below to generate the simple, compatible template.**")
    else:
        st.success(f"Template file '{TEMPLATE_FILE}' found. Proceed to file upload.")
        
    if st.button("Download Template File (input-template.pptx)"):
        template_bytes = create_simple_template_pptx()
        st.download_button(
            "Click to Download Template",
            data=template_bytes,
            file_name=TEMPLATE_FILE,
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    st.markdown("---")
    # --- End Template Download Section ---


    uploads = st.file_uploader(
        "Upload per setting: CSV + Rendering (JPG/PNG) + optional Linedrawing (JPG/PNG).",
        type=["csv","jpg","jpeg","png"],
        accept_multiple_files=True
    )
    
    # Update file list state for display
    if uploads:
        st.session_state.uploaded_file_names = [f.name for f in uploads]
    
    # ----------------------------------------------------------------------
    # Simplified File List Display
    # ----------------------------------------------------------------------
    
    if st.session_state.uploaded_file_names:
        st.subheader("Uploaded Files")
        st.markdown("The following files are ready for processing:")
        
        # Display list of uploaded file names
        for file_name in st.session_state.uploaded_file_names:
            st.write(f"- {file_name}")
    
    st.markdown("---")

    if st.button("Generate PPT", type="primary"):
        if not os.path.exists(TEMPLATE_FILE):
            st.error(f"Template file '{TEMPLATE_FILE}' is missing. Please use the download button above to generate it."); st.stop()
        
        if not uploads:
             st.error("Please upload files first."); st.stop()

        # USER-FRIENDLY SINGLE SPINNER
        with st.spinner("Processing files and generating presentation... This may take a moment."):
            try:
                # 1. Load data
                master_df  = load_master()
                mapping_df = load_mapping()
    
                # 2. Final Grouping (Using Automatic Guess from filenames)
                final_groups_map: Dict[str, Dict[str, Any]] = {}
                
                for f in uploads:
                    name, _ = os.path.splitext(f.name)
                    
                    base_name = name
                    lf = f.name.lower()
                    is_line_drawing = any(k in lf for k in ["line","floorplan","drawing"])
                    if is_line_drawing:
                        base_name = re.sub(r"[\s_-]*(line|floorplan|drawing)$", "", base_name, flags=re.I).strip()
                    
                    base_key = base_name.strip()
                    
                    if base_key not in final_groups_map:
                         setting_name = base_key.split(" - ", 1)[-1].title().strip()
                         if setting_name == base_key: setting_name = base_key.title()
                         final_groups_map[base_key] = {"name": setting_name, "csv": None, "rendering": None, "line": None}
                         
                    
                    if lf.endswith(".csv"):
                        final_groups_map[base_key]["csv"] = f
                    elif any(k in lf for k in ["line","floorplan","drawing"]):
                        final_groups_map[base_key]["line"] = f
                    elif lf.endswith((".jpg",".jpeg",".png")):
                        final_groups_map[base_key]["rendering"] = f
                
                settings, overview_imgs = [], []
                for key, data in final_groups_map.items():
                    if not data["csv"] or not data["rendering"]:
                        st.warning(f"⚠️ Skipping setting '{data['name']}' – requires both CSV and Rendering.")
                        continue
    
                    # 3. Read CSV and prepare items
                    df = pcon_from_csv(data["csv"])
    
                    items = []
                    for _, r in df.iterrows():
                        article = r["ARTICLE_NO"]
                        qty = int(r["QUANTITY"])
                        desc = mapping_description(mapping_df, article)
                        pack = packshot_lookup(master_df, article)
                        newno = new_item_lookup(mapping_df, article)
                        items.append({
                            "article_no": article, "qty": qty,
                            "desc_text": desc, "packshot_url": pack, "new_item_no": newno
                        })
    
                    render_bytes = data["rendering"].read()
                    overview_imgs.append(render_bytes)
                    line_bytes = data["line"].read() if data["line"] else None
    
                    settings.append({
                        "name": data["name"],
                        "items": items,
                        "rendering_bytes": render_bytes,
                        "linedrawing_bytes": line_bytes
                    })
    
                if not settings:
                    st.error("❌ No valid settings found (Requires CSV and Rendering for at least one group).")
                    st.stop()
    
                # 4. Generate PowerPoint
                ppt_bytes = build_presentation(master_df, mapping_df, settings, overview_imgs)
                
                st.success("✅ PowerPoint generated successfully!")
    
                st.download_button(
                    "Download PPT",
                    data=ppt_bytes,
                    file_name="Muuto_Settings.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                error_message = str(e)
                if "Read timed out" in error_message or "HTTPSConnectionPool" in error_message:
                     st.error("❌ Network Timeout Error: Could not connect to Google Sheets. Please ensure the links are correct and try again.")
                elif "Master is missing columns" in error_message or "Mapping is missing columns" in error_message:
                     st.error(f"❌ Data Error: {error_message}")
                elif "Template file" in error_message:
                     st.error(f"❌ Template Error: {error_message}. Please use the 'Download Template File' button to fix.")
                else:
                    st.error("❌ Generation Error: An unexpected error occurred. Please try again or check logs for details.")
                    st.exception(e)
                st.stop() 

if __name__ == "__main__":
    main()
