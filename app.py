# app.py
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from PIL import Image
import io, os, re, requests, csv
from typing import List, Dict, Any, Tuple
from copy import deepcopy

# --------- Konstanter (kun i kode) ---------
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
OVERVIEW_TITLE = "OVERVIEW" # Bruges nu kun som tjek

PACKSHOT_TAGS = [f"{{{{ProductPackshot{i}}}}}" for i in range(1, 13)]
PROD_DESC_TAGS = [f"{{{{PRODUCT DESCRIPTION {i}}}}}" for i in range(1, 13)]
OVERVIEW_TAGS = [f"{{{{Rendering{i}}}}}" for i in range(1, 13)]

# --------- Utils ---------
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
            # Brug 'in' i stedet for '==' for at fange tags med omgivende tekst (f.eks. 'SHOP THE LOOK - {{SETTINGNAME}}')
            if want in _norm_placeholder_text(txt):
                return shp
    for shp in getattr(slide, "placeholders", []):
        if getattr(shp, "has_text_frame", False) and shp.text_frame:
            txt = shp.text_frame.text or ""
            if want in _norm_placeholder_text(txt):
                return shp
    return None

def set_text_preserve_style(shape, text: str):
    """Expect a shape with text_frame. Preserves first run style if present. Handles nested tags."""
    if not shape or not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    
    # Prøv at erstatte specifikke tags, mens du beholder den omkringliggende tekst
    if TAG_SETTINGNAME in shape.text:
        # Hvis det er et felt som "SHOP THE LOOK - {{SETTINGNAME}}", erstat kun tagget
        new_text = shape.text.replace(TAG_SETTINGNAME, text)
    else:
        new_text = text

    # capture style from first run if exists
    font_name = font_size = font_bold = None
    if tf.paragraphs and tf.paragraphs[0].runs:
        r0 = tf.paragraphs[0].runs[0]
        font_name, font_size, font_bold = r0.font.name, r0.font.size, r0.font.bold
        
    # clear all paragraphs and rebuild one run
    while tf.paragraphs:
        p = tf.paragraphs[0]
        for r in list(p.runs):
            r.text = ""
        try:
            tf._element.remove(p._p)
        except Exception:
            break
            
    p = tf.add_paragraph()
    run = p.add_run()
    run.text = new_text or "" # Brug den potentielt delvist erstattede tekst
    if font_name: run.font.name = font_name
    if font_size: run.font.size = font_size
    if font_bold is not None: run.font.bold = font_bold
    tf.word_wrap = True


def replace_image_by_tag(slide, tag: str, img_bytes: bytes):
    if not img_bytes: return
    ph = find_shape_by_placeholder(slide, tag)
    if not ph: return
    left, top, w, h = ph.left, ph.top, ph.width, ph.height
    
    # Opret et midlertidigt IO stream
    img_stream = io.BytesIO(img_bytes)
    
    # Gem originale dimensioner for at beregne aspect ratio
    try:
        im = Image.open(img_stream)
        img_w, img_h = im.size
        aspect_ratio = img_w / img_h
        img_stream.seek(0) # Spol tilbage til start efter læsning
    except Exception:
        # Standard hvis billeddata ikke er gyldig (skal ikke ske efter preprocess, men sikkerhed)
        aspect_ratio = 1.0
        img_stream.seek(0)
    
    # Beregn justerede dimensioner for at BEVARE ASPECT RATIO
    # Vælg den dimension, der giver mindst skalering
    if (w / aspect_ratio) <= h:
        # Billedet er bredere end pladsen tillader (eller lig med)
        new_h = w / aspect_ratio
        new_w = w
    else:
        # Billedet er højere end pladsen tillader
        new_w = h * aspect_ratio
        new_h = h

    # Centrer billedet i placeholderen
    new_left = left + (w - new_w) / 2
    new_top = top + (h - new_h) / 2
    
    # Fjern den gamle placeholder
    try: ph.element.getparent().remove(ph.element)
    except Exception: pass
    
    # Indsæt det nye billede med bevaret aspect ratio
    slide.shapes.add_picture(img_stream, new_left, new_top, width=new_w, height=new_h)


def duplicate_slide(prs: Presentation, slide):
    new_slide = prs.slides.add_slide(slide.slide_layout)
    for shp in list(new_slide.shapes):
        sp = shp.element
        sp.getparent().remove(sp)
    for shp in slide.shapes:
        new_slide.shapes._spTree.append(deepcopy(shp._element))
    return new_slide

# --- NY FUNKTION: Fjern Slide ---
def remove_slide(prs: Presentation, index: int):
    """Fjerner et slide fra præsentationen ved hjælp af dets index."""
    rId = prs.slides._sldIdLst[index].rId
    prs.part.drop_rel(rId)
    del prs.slides._sldIdLst[index]


# --------- Data-loaders forbliver de samme ---------
# ... load_master, load_mapping, pcon_from_csv, fallback_key, packshot_lookup, new_item_lookup, mapping_description ...

# [Bemærk: Jeg udelader de uændrede data-loaders for at spare plads, men de skal være inkluderet i den endelige kode.]

# --------- Billeder forbliver de samme, med fetch_image rettelse ---------
@st.cache_data(ttl=3600)
def fetch_image(url: str) -> bytes | None:
    if not url or not url.startswith("http"): return None
    try:
        r = requests.get(url, timeout=15); r.raise_for_status()
        
        content_type = r.headers.get("Content-Type","").lower()
        if content_type.startswith("text/") and "image" not in content_type:
             return None
             
        return r.content
    except requests.RequestException:
        return None

def preprocess(img: bytes, max_side=1400, quality=85) -> bytes:
    try:
        im = Image.open(io.BytesIO(img))
        if im.mode != "RGB": im = im.convert("RGB")
        if max(im.size) > max_side:
            ratio = min(max_side/im.width, max_side/im.height)
            im = im.resize((int(im.width*ratio), int(im.height*ratio)), Image.Resampling.LANCZOS)
        buf = io.BytesIO(); im.save(buf, format="JPEG", quality=85); return buf.getvalue()
    except Exception:
        return img
# --------------------------------------------------------------------------

# --------- PPT-byggesten (Med opdateret logik) ---------

# Brug den eksisterende find_first_slide_with_tag, men den er ikke nødvendig for at finde OVERVIEW/Setting nu, 
# da vi tjekker alle slides.

def find_first_slide_with_tag(prs: Presentation, tag: str) -> Tuple[Any, int]:
    want = _norm_tag(tag)
    for i, sl in enumerate(prs.slides):
        for shp in sl.shapes:
            if getattr(shp, "has_text_frame", False) and shp.text_frame:
                if want in _norm_placeholder_text(shp.text_frame.text):
                    return sl, i
    return None, -1

# ... blank_layout, add_products_table_on_blank forbliver de samme ...

def build_presentation(master_df: pd.DataFrame,
                       mapping_df: pd.DataFrame,
                       groups: List[Dict[str, Any]],
                       overview_renderings: List[bytes],
                       status=None) -> bytes:
    prs = Presentation(TEMPLATE_FILE)

    # 1. FIND TEMPLATE SLIDES FØRST
    # Da OVERVIEW-title også kan findes i andre placeholders, leder vi efter en slide, der KUN indeholder titlen.
    setting_tpl, setting_idx = find_first_slide_with_tag(prs, TAG_SETTINGNAME)
    overview_tpl, overview_idx = find_first_slide_with_tag(prs, OVERVIEW_TITLE)
    
    # 2. OVERVIEW SLIDE GENERERING
    if status: status.write("4️⃣ Indsætter OVERVIEW…")
    if overview_renderings:
        if overview_tpl is not None:
            s = duplicate_slide(prs, overview_tpl)
            
            # Udfyld rendering tags
            for i, rb in enumerate(overview_renderings[:12]):
                replace_image_by_tag(s, OVERVIEW_TAGS[i], preprocess(rb))
        else:
            # Eksisterende fallback logik hvis ingen OVERVIEW template findes
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

    # 3. SETTING SLIDE GENERERING
    for gi, g in enumerate(groups, 1):
        if status: status.write(f"5️⃣ Opretter setting {gi}/{len(groups)}: {g['name']}")

        base = setting_tpl if setting_tpl else prs.slides[0]
        s = duplicate_slide(prs, base)
        
        # Udfyld SETTINGNAME: Bruger forbedret set_text_preserve_style
        shape_setting_name = find_shape_by_placeholder(s, TAG_SETTINGNAME)
        if shape_setting_name:
            set_text_preserve_style(shape_setting_name, g["name"])
        
        # Udfyld Rendering og Linedrawing (bruger nu billed centreringslogik)
        if g.get("rendering_bytes"):
            replace_image_by_tag(s, TAG_RENDERING, preprocess(g["rendering_bytes"]))
        if g.get("linedrawing_bytes"):
            replace_image_by_tag(s, TAG_LINEDRAWING, preprocess(g["linedrawing_bytes"]))

        # Packshots + PRODUCT DESCRIPTION
        replaced_desc = 0
        replaced_pack = 0
        for idx, it in enumerate(g["items"][:12]):
            if it.get("packshot_url"):
                raw = fetch_image(it["packshot_url"])
                if raw:
                    replace_image_by_tag(s, PACKSHOT_TAGS[idx], preprocess(raw))
                    replaced_pack += 1
            desc_tag_shape = find_shape_by_placeholder(s, PROD_DESC_TAGS[idx])
            if desc_tag_shape:
                set_text_preserve_style(desc_tag_shape, it.get("desc_text",""))
                replaced_desc += 1
        if status:
            status.write(f"    • Packshots sat: {replaced_pack} | Descriptions sat: {replaced_desc}")

        # Produkttabel på blank slide
        rows = []
        for it in g["items"]:
            id_combo = it["article_no"]
            if it.get("new_item_no"): id_combo = f'{it["article_no"]} / {it["new_item_no"]}'
            rows.append([it["qty"], it["desc_text"], id_combo])
        add_products_table_on_blank(prs, f"Products – {g['name']}", rows)

    # 4. FJERN DE ORIGINALE TEMPLATE SLIDES
    # Dette sikrer, at de tomme template-slides ikke er i det endelige output
    if setting_tpl and setting_idx is not None:
        remove_slide(prs, setting_idx)
        # Hvis overview var efter setting, er dens index faldet med 1
        if overview_tpl and overview_idx is not None and overview_idx > setting_idx:
            overview_idx -= 1
            
    if overview_tpl and overview_idx is not None:
        # Fjern overview slide. Da den duplikerede slide er indsat FØR (eller på samme plads), 
        # er det nu det duplikerede slide, der er på plads 0, og den originale (template) slide, 
        # der skal fjernes.
        remove_slide(prs, overview_idx)


    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# --------- UI forbliver den samme ---------
# ... main() funktionen forbliver den samme ...
