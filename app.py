# app.py
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
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
    for shp in getattr(slide, "placeholders", []):
        if getattr(shp, "has_text_frame", False) and shp.text_frame:
            txt = shp.text_frame.text or ""
            if want in _norm_placeholder_text(txt):
                return shp
    return None

def set_text_preserve_style(shape, text: str):
    """
    Indsætter tekst og bevarer stilen fra den første 'run'.
    Håndterer indlejrede tags (f.eks. {{SETTINGNAME}}) og multiline-input.
    Al tekst konverteres til STORE BOGSTAVER.
    """
    if not shape or not getattr(shape, "has_text_frame", False):
        return
    tf = shape.text_frame
    
    # 1. Bestem endelig tekst (håndterer tag-erstatning og uppercase)
    final_text_content = text.upper()
    
    current_text_upper = tf.text.upper()
    if TAG_SETTINGNAME.upper() in current_text_upper:
        # Hvis det er et indlejret tag, erstat kun tagget
        final_text_content = current_text_upper.replace(TAG_SETTINGNAME.upper(), text.upper())

    # 2. Fang stilen
    font_name = font_size = font_bold = None
    if tf.paragraphs and tf.paragraphs[0].runs:
        r0 = tf.paragraphs[0].runs[0]
        font_name, font_size, font_bold = r0.font.name, r0.font.size, r0.font.bold
        
    # 3. Ryd eksisterende indhold
    # Sørg for at den første paragraf fjernes sidst, hvis vi genbruger den
    while tf.paragraphs:
        p = tf.paragraphs[0]
        for r in list(p.runs): r.text = ""
        try: tf._element.remove(p._p)
        except Exception: break
            
    # 4. Genopbyg indhold linje for linje
    lines = final_text_content.split('\n')
    
    # Skab en ny paragraf til indsættelse (da vi fjernede den første i loopet)
    p = tf.add_paragraph() 
    
    # Vi bruger kun den første linie i `lines` til den paragraf, vi netop har oprettet (p)
    # og opretter nye paragraffer for de resterende linjer (hvis det er en produktliste)
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
    
    # Beregn justerede dimensioner for at BEVARE ASPECT RATIO (fit into placeholder)
    if (w / aspect_ratio) <= h:
        new_h = w / aspect_ratio
        new_w = w
    else:
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

def remove_slide(prs: Presentation, index: int):
    """Fjerner et slide fra præsentationen ved hjælp af dets index."""
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


# ----------------------------------------------------------------------
# --------- Data-loaders & Lookups (Med alle de tidligere rettelser) ----
# ----------------------------------------------------------------------

@st.cache_data
def load_master() -> pd.DataFrame:
    url = resolve_gsheet_to_csv(MASTER_URL)
    r = requests.get(url, timeout=20); r.raise_for_status()
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
        raise ValueError("Master mangler kolonner: ITEM NO. og/eller IMAGE URL (IMAGE DOWNLOAD LINK accepteres).")
    out = df.rename(columns={col_item:"ITEM NO.", col_img:"IMAGE URL"})[["ITEM NO.","IMAGE URL"]]
    for c in out.columns: out[c] = out[c].astype(str).str.strip()
    return out

@st.cache_data
def load_mapping() -> pd.DataFrame:
    url = resolve_gsheet_to_csv(MAPPING_URL)
    r = requests.get(url, timeout=20); r.raise_for_status()
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
        raise ValueError("Mapping mangler kolonner: OLD Item-variant, New Item No., Description.")
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
                last_err = ValueError(f"For få kolonner med cfg={cfg}, shape={df.shape}")
                df = None; continue
            break
        except Exception as e:
            last_err = e; df = None
    if df is None:
        raise ValueError(f"Kunne ikke parse pCon CSV. Sidste fejl: {last_err}")
    sub = df.iloc[:, [IDX_SHORT, IDX_VARIANT, IDX_ARTICLE, IDX_QTY]].copy()
    sub.columns = ["SHORT_TEXT","VARIANT_TEXT","ARTICLE_NO","QUANTITY"]
    sub["ARTICLE_NO"]    = sub["ARTICLE_NO"].astype(str).str.strip()
    sub["SHORT_TEXT"]    = sub["SHORT_TEXT"].astype(str).str.strip()
    sub["VARIANT_TEXT"] = sub["VARIANT_TEXT"].astype(str).str.strip()
    sub["QUANTITY"]      = pd.to_numeric(sub["QUANTITY"], errors="coerce").fillna(1).astype(int)
    sub = sub[sub["ARTICLE_NO"].ne("")]
    if sub.empty:
        raise ValueError("pCon CSV blev læst, men indeholdt ingen gyldige rækker med ARTICLE_NO.")
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

# --------- Billeder (Med robust fetching) ---------
@st.cache_data(ttl=3600)
def fetch_image(url: str) -> bytes | None:
    if not url or not url.startswith("http"): return None
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        
        content_type = r.headers.get("Content-Type","").lower()
        if content_type.startswith("text/") and "image" not in content_type:
             return None
             
        return r.content
    except requests.exceptions.RequestException:
        # Returnerer None ved netværksfejl (404, 500, timeout)
        return None

def preprocess(img: bytes, max_side=1400, quality=85) -> bytes:
    try:
        im = Image.open(io.BytesIO(img))
        # Konverter til RGB for at undgå problemer med pptx og PIL
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGB")
            
        if max(im.size) > max_side:
            ratio = min(max_side/im.width, max_side/im.height)
            im = im.resize((int(im.width*ratio), int(im.height*ratio)), Image.Resampling.LANCZOS)
            
        buf = io.BytesIO()
        # Brug JPEG og optimering (ligner din gamle kode)
        im.save(buf, format="JPEG", quality=quality, optimize=True) 
        buf.seek(0)
        return buf.getvalue()
    except Exception:
        # Hvis preprocess fejler, returner de originale bytes for at give pptx en chance
        return img 

# --------- PPT-byggesten ---------
def blank_layout(prs: Presentation):
    for ly in prs.slide_layouts:
        if ly.name and "blank" in ly.name.lower():
            return ly
    return prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[1]

def create_product_list_text(items: List[Dict[str, Any]]) -> str:
    """Genererer en multiline streng af produktlister i det ønskede format."""
    lines = []
    for it in items:
        article = str(it["article_no"])
        desc = str(it["desc_text"])
        qty = str(int(it["qty"])) 
        newno = str(it.get("new_item_no", ""))
        
        id_combo = article
        if newno and newno != "NAN":
            id_combo = f'{article} / {newno}'
        
        line = f'{qty} X {desc} - {id_combo}'
        lines.append(line)
        
    return "\n".join(lines)


def build_presentation(master_df: pd.DataFrame,
                       mapping_df: pd.DataFrame,
                       groups: List[Dict[str, Any]],
                       overview_renderings: List[bytes]) -> bytes:

    prs = Presentation(TEMPLATE_FILE)

    setting_tpl, setting_idx = find_first_slide_with_tag(prs, TAG_SETTINGNAME)
    overview_tpl, overview_idx = find_first_slide_with_tag(prs, OVERVIEW_TITLE)
    
    # 2. OVERVIEW SLIDE GENERERING
    if overview_renderings:
        if overview_tpl is not None:
            s = duplicate_slide(prs, overview_tpl)
            for i, rb in enumerate(overview_renderings[:12]):
                replace_image_by_tag(s, OVERVIEW_TAGS[i], preprocess(rb))
        else:
            # Fallback til blank slide
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
        
        base = setting_tpl if setting_tpl else prs.slides[0]
        s = duplicate_slide(prs, base)
        
        # A. Udfyld SETTINGNAME (inkl. "SHOP THE LOOK - {{SETTINGNAME}}")
        setting_title_shape = find_shape_by_placeholder(s, TAG_SETTINGNAME)
        if setting_title_shape:
            # set_text_preserve_style er opdateret til at håndtere korrekt erstatning og uppercase
            set_text_preserve_style(setting_title_shape, g["name"]) 

        # B. Rendering, Linedrawing, Packshots, Product Descriptions
        if g.get("rendering_bytes"):
            replace_image_by_tag(s, TAG_RENDERING, preprocess(g["rendering_bytes"]))
        if g.get("linedrawing_bytes"):
            replace_image_by_tag(s, TAG_LINEDRAWING, preprocess(g["linedrawing_bytes"]))

        for idx, it in enumerate(g["items"][:12]):
            # Indsættelse af packshot med robust fetching
            if it.get("packshot_url"):
                raw = fetch_image(it["packshot_url"])
                if raw:
                    replace_image_by_tag(s, PACKSHOT_TAGS[idx], preprocess(raw))
                    
            # Indsættelse af produktbeskrivelse
            desc_tag_shape = find_shape_by_placeholder(s, PROD_DESC_TAGS[idx])
            if desc_tag_shape:
                set_text_preserve_style(desc_tag_shape, it.get("desc_text",""))

        # C. Produktoversigt som formateret tekst i {{ProductsinSettingList}}
        product_list_text = create_product_list_text(g["items"])
        list_shape = find_shape_by_placeholder(s, TAG_PRODUCTS_LIST)
        
        if list_shape:
            # set_text_preserve_style er opdateret til at håndtere multiline tekst og uppercase
            set_text_preserve_style(list_shape, product_list_text)

    # 4. FJERN DE ORIGINALE TEMPLATE SLIDES
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
    st.set_page_config(page_title="Muuto PPT Generator", layout="wide")
    st.title("Muuto PPT Generator")

    uploads = st.file_uploader(
        "Upload pr. setting: CSV + Rendering (JPG/PNG) + valgfri Linedrawing (JPG/PNG).",
        type=["csv","jpg","jpeg","png"],
        accept_multiple_files=True
    )

    if st.button("Generér PPT", type="primary"):
        if not os.path.exists(TEMPLATE_FILE):
            st.error(f"Skabelon '{TEMPLATE_FILE}' mangler."); st.stop()

        # BRUGERVENLIG ENKEL SPINNER
        with st.spinner("Arbejder på at generere præsentationen... Dette kan tage et øjeblik"):
            try:
                # 1. Indlæsning af data
                master_df  = load_master()
                mapping_df = load_mapping()
    
                # 2. Gruppering af filer
                groups_map: Dict[str, Dict[str, Any]] = {}
                for f in uploads or []:
                    name, ext = os.path.splitext(f.name)
                    lf = f.name.lower()
                    
                    base_name = name
                    is_line_drawing = any(k in lf for k in ["line","floorplan","drawing"])
                    if is_line_drawing:
                        base_name = re.sub(r"[\s_-]*(line|floorplan|drawing)$", "", base_name, flags=re.I).strip()
                    
                    base_key = base_name.strip()
                    parts = base_key.split(" - ", 1)
                    setting_title = parts[-1].title().strip() if len(parts) > 1 else base_key.title()
                    
                    groups_map.setdefault(base_key, {"name": setting_title, "csv": None, "rendering": None, "line": None})
                    
                    if ext.lower() == ".csv":
                        groups_map[base_key]["csv"] = f
                    elif is_line_drawing:
                        groups_map[base_key]["line"] = f
                    elif lf.endswith((".jpg",".jpeg",".png")):
                        groups_map[base_key]["rendering"] = f
                
                settings, overview_imgs = [], []
                for base, data in groups_map.items():
                    if not data["csv"] or not data["rendering"]:
                        st.warning(f"⚠️ Springer over '{data['name']}' – mangler CSV eller Rendering.")
                        continue
    
                    # 3. Læsning af CSV
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
                    st.error("❌ Ingen gyldige settings fundet (kræver CSV og Rendering for mindst én gruppe).")
                    st.stop()
    
                # 4. Generering af PowerPoint
                ppt_bytes = build_presentation(master_df, mapping_df, settings, overview_imgs)
                
                st.success("✅ PowerPoint genereret succesfuldt!")
    
                st.download_button(
                    "Download PPT",
                    data=ppt_bytes,
                    file_name="Muuto_Settings.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            except Exception as e:
                st.error("❌ Fejl i generering: Prøv venligst igen.")
                st.exception(e)

if __name__ == "__main__":
    main()
