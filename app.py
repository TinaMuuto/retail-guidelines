import io
import re
import time
from typing import Dict, List, Optional, Tuple
from pathlib import Path
# from copy import deepcopy er fjernet, da den ikke blev brugt
# typing.Any er fjernet for at rydde op

import pandas as pd
import requests
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt # Pt blev ikke brugt, men bibeholdes da det kan være et artefakt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER

# --- STREAMLIT CACHING DEKORATORER ---
# Defineret her for bedre overblik og genbrug
st.cache_data.clear() # Clear cache ved app start for udvikling
st.cache_resource.clear() # Clear resource cache for prs template

# ---------------------- Constants ----------------------
TEMPLATE_PATH = Path("input-template.pptx")
# NYE URLs afledt af bruger input. Disse SKAL være offentligt delte (Anyone with the link)!
DEFAULT_MASTER_URL = "https://docs.google.com/spreadsheets/d/1blj42SbFpszWGyOrDOUwyPDJr9K1NGpTMX6eZTbt_P4/pub?output=csv&gid=1152340088"
DEFAULT_MAPPING_URL = "https://docs.google.com/spreadsheets/d/1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk/pub?output=csv&gid=1056617222"

OUTPUT_NAME = "Muuto_Settings_Generated.pptx"

MAX_OVERVIEW_IMAGES = 12
HTTP_TIMEOUT = 10
HTTP_RETRIES = 1
MAX_IMAGE_PX = 1400
JPEG_QUALITY = 85

# ---------------------- Utils - Naming & Data Lookup ----------------------

def clean_name(name: str) -> str:
    """Renser shape navne for code matching (fjerner klammer, whitespace, og lowercase)."""
    if name is None:
        return ""
    name = name.strip()
    # Bruger re.sub med en kompileret regex for at forbedre performance en smule
    name = re.sub(r"^\{\{|\}\}$", "", name).strip()
    return re.sub(r"\s+", "", name).lower()

def normalize_key(key):
    """Normaliserer et artikel-key for robust sammenligning (trim, uppercase, fjern SPECIAL- prefix)."""
    if isinstance(key, str):
        return key.strip().upper().replace('SPECIAL-', '')
    return str(key).strip().upper().replace('SPECIAL-', '')

def get_base_article_no(article_no):
    """Henter det basale artikelnummer (alt før første bindestreg, fjerner SPECIAL- prefix)."""
    article_no = normalize_key(article_no)
    if '-' in article_no:
        return article_no.split('-')[0].strip()
    return article_no.strip()

def build_shape_map(slide) -> Dict[str, list]:
    """Opretter en ordbog, der mapper rensede shape navne til shapes."""
    mapping: Dict[str, List] = {}
    for shape in slide.shapes:
        try:
            nm = clean_name(getattr(shape, "name", ""))
            if nm:
                mapping.setdefault(nm, []).append(shape)
        except Exception:
            # Ignorerer shapes uden navn
            continue
    return mapping

def safe_find_shape(shape_map: Dict[str, list], key: str, index: int = 0) -> Optional[object]:
    """Hjælpefunktion til sikkert at finde en shape ved dens rensede navn, med fleksibilitet (mindre skabelon-workarounds)."""
    clean_key = clean_name(key)
    
    if clean_key in shape_map and len(shape_map[clean_key]) > index:
        return shape_map[clean_key][index]
    
    # Template Fleksibilitet Check: Kompenserer for almindelige slåfejl (beholdes for robusthed)
    if clean_key.startswith("productpackshot"):
        alt_key = clean_key.replace("productpackshot", "packshot")
        if alt_key in shape_map and len(shape_map[alt_key]) > index:
            return shape_map[alt_key][index]
    
    if clean_key == "linedrawing":
        alt_key = "llinedrawing"
        if alt_key in shape_map and len(shape_map[alt_key]) > index:
            return shape_map[alt_key][index]

    return None

# --- Utils - Network & Data Loading (Caching er tilføjet her) ---

def http_get_bytes(url: str) -> Optional[bytes]:
    """Henter indholdet af en URL med genforsøg og timeout."""
    if not url:
        return None
    for attempt in range(HTTP_RETRIES + 1):
        try:
            resp = requests.get(url, timeout=HTTP_TIMEOUT, allow_redirects=True)
            if resp.status_code == 200 and resp.content:
                return resp.content
        except Exception:
            pass # Fejlen ignoreres, da den logges i den kaldende funktion, hvis den er kritisk
        time.sleep(0.2 * attempt)
    return None

def parse_csv_flex(buf: bytes) -> pd.DataFrame:
    """Forsøger at parse CSV data med forskellige separatorer og kodninger."""
    if buf is None:
        return pd.DataFrame()
    candidates = [
        {"sep": ";", "encoding": "utf-8-sig"},  # CSV eksport fra DK-systemer
        {"sep": ",", "encoding": "utf-8"},      # Standard CSV
        {"sep": ";", "encoding": "utf-8"},
        {"sep": "\t", "encoding": "utf-8"},
        {"sep": ",", "encoding": "utf-8-sig"},
        {"sep": ";", "encoding": "latin-1"},
    ]
    for c in candidates:
        try:
            return pd.read_csv(io.BytesIO(buf), sep=c["sep"], encoding=c["encoding"], skipinitialspace=True)
        except Exception:
            continue
    return pd.DataFrame()

@st.cache_data(show_spinner="Indlæser og normaliserer Master/Mapping Data...")
def load_remote_csv_and_normalize(url: str, normalizer_func) -> pd.DataFrame:
    """Henter og normaliserer CSV fra en ekstern URL. Bruger caching."""
    content = http_get_bytes(url)
    if content is None:
        return pd.DataFrame()
    df = parse_csv_flex(content)
    return normalizer_func(df)

# --- Utils - Data Normalization (Logikken er beholdt, da den er robust for Muuto-data) ---

def normalize_master(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliserer Master Data (Packshot URLs)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["ITEM NO.", "IMAGE"])
    
    cols = {c: c.strip() for c in df.columns}
    df = df.rename(columns=cols)
    
    col_img = next((c for c in df.columns if c.upper() == "IMAGE DOWNLOAD LINK" or c.upper() == "IMAGE URL" or ("image" in c.lower() and ("url" in c.lower() or "download" in c.lower()))), None)
    item_col = next((c for c in df.columns if c.strip().upper() == "ITEM NO." or ("item" in c.lower() and "no" in c.lower())), None)

    if item_col is None or col_img is None:
        return pd.DataFrame(columns=["ITEM NO.", "IMAGE"])
    
    out = df[[item_col, col_img]].copy()
    out.columns = ["ITEM NO.", "IMAGE"]
    out["ITEM NO."] = out["ITEM NO."].astype(str).str.strip().apply(normalize_key) 
    out["IMAGE"] = out["IMAGE"].astype(str).str.strip()      
    return out

def normalize_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliserer Mapping Data (Beskrivelser, Nye Varenumre)."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["OLD Item-variant", "Description", "New Item No."])
    
    cols = {c: c.strip() for c in df.columns}
    df = df.rename(columns=cols)
    
    col_old = next((c for c in df.columns if c.lower().strip() in ["old item-variant", "old item variant", "olditem-variant"] or ("old" in c.lower() and "variant" in c.lower())), None)
    col_new = next((c for c in df.columns if c.lower().strip() in ["new item no.", "new item no", "new item number"] or ("new" in c.lower() and ("no" in c.lower() or "number" in c.lower()))), None)
    col_desc = next((c for c in df.columns if c.lower().strip() == "description" or "desc" in c.lower()), None)
    
    if not col_old or not col_new or not col_desc:
        return pd.DataFrame(columns=["OLD Item-variant", "Description", "New Item No."])
        
    out = df[[col_old, col_desc, col_new]].copy()
    out.columns = ["OLD Item-variant", "Description", "New Item No."]
    
    out["OLD Item-variant"] = out["OLD Item-variant"].astype(str).str.strip().apply(normalize_key)
    out["New Item No."] = out["New Item No."].astype(str).str.strip()
    out["Description"] = out["Description"].astype(str).str.strip()
    return out

def normalize_pcon(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliserer pCon CSV for at udtrække Artikel No. og Antal."""
    if df is None or df.empty:
        return pd.DataFrame(columns=["ARTICLE_NO", "Quantity"])
    
    norm = {c: re.sub(r"[^a-z0-9]", "", c.lower()) for c in df.columns}
    
    article_col = next((c for c in df.columns if norm[c] in {"articleno","article","articlenumber","artno","artnr","artnumber","itemno","itemnumber","articlecode"} or ("article" in norm[c] and "no" in norm[c]) or ("item" in norm[c] and "no" in norm[c])), None)
    qty_col = next((c for c in df.columns if norm[c] in {"qty","quantity","quantities","qtytotal","qtysum"} or "qty" in norm[c]), None)

    if article_col is None:
        return pd.DataFrame(columns=["ARTICLE_NO", "Quantity"])
    
    out = pd.DataFrame()
    out["ARTICLE_NO"] = df[article_col].astype(str).fillna("").str.strip()
    
    if qty_col is not None:
        out["Quantity"] = pd.to_numeric(df[qty_col], errors="coerce").fillna(1).astype(int)
    else:
        out["Quantity"] = 1
        
    return out[["ARTICLE_NO", "Quantity"]]

# --- Utils - Data Lookup Functions (Combined) ---

def lookup_data_with_fallback(article_no: str, df: pd.DataFrame, key_col: str, return_col: str, normalize_func=normalize_key) -> str:
    """Udfører direkte match, derefter base match opslag."""
    article_no = str(article_no)
    
    # 1. Direkte Match
    try:
        match_direct = df[df[key_col].apply(normalize_func) == normalize_func(article_no)]
        if not match_direct.empty:
            val = match_direct.iloc[0][return_col]
            return "" if pd.isna(val) else str(val).strip()
    except Exception:
        pass

    # 2. Fallback Match (Base Article No)
    article_base = get_base_article_no(article_no)
    try:
        match_fallback = df[df[key_col].apply(get_base_article_no) == article_base]
        if not match_fallback.empty:
            val = match_fallback.iloc[0][return_col]
            return "" if pd.isna(val) else str(val).strip()
    except Exception:
        pass
        
    return ""

def find_packshot_url(article_no: str, mapping_df: pd.DataFrame, master_df: pd.DataFrame) -> Optional[str]:
    """Finder Packshot URL ved hjælp af 3-trins opslagslogikken (pCon -> Mapping -> Master)."""
    if master_df.empty: return None

    # 1. Find Nyt Varenummer ved hjælp af mapping data
    new_item = find_new_item(article_no, mapping_df)
    
    # 2. Opslag i Master Data (ved hjælp af Nyt Varenummer hvis tilgængeligt, ellers det originale article_no)
    lookup_key = new_item if new_item else article_no
    
    url = lookup_data_with_fallback(lookup_key, master_df, "ITEM NO.", "IMAGE", normalize_func=normalize_key)
    
    return url if url else None

def find_description(article_no: str, mapping_df: pd.DataFrame) -> str:
    """Finder produktbeskrivelse fra mapping data."""
    return lookup_data_with_fallback(article_no, mapping_df, "OLD Item-variant", "Description")

def find_new_item(article_no: str, mapping_df: pd.DataFrame) -> Optional[str]:
    """Finder det nye varenummer fra mapping data."""
    val = lookup_data_with_fallback(article_no, mapping_df, "OLD Item-variant", "New Item No.")
    return val if val else None

# --- Utils - File Grouping & PPTX Helpers ---

def group_key_from_filename(name: str) -> Tuple[str, str]:
    """Udtrækker gruppenøgle og filtype fra filnavnet (bruger alt efter ' - ')."""
    base = Path(name).stem
    lname = base.lower()
    
    if "floorplan" in lname or "line_drawing" in lname or "line drawing" in lname:
        t = "linedrawing"
    else:
        ext = Path(name).suffix.lower()
        if ext == ".csv":
            t = "csv"
        elif ext in [".jpg", ".jpeg", ".png", ".webp"]:
            t = "render"
        else:
            t = "other"
            
    if " - " in base:
        key = base.split(" - ", 1)[1]
    else:
        parts = re.split(r"[-_]", base)
        key = parts[-1] if parts else base
        
    # Renser nøglen for filtypebeskrivelser
    key = re.sub(r"\s+(floorplan|line\s*drawing|linedrawing)$", "", key, flags=re.IGNORECASE).strip()
    return key, t

def build_groups(upload_list: List[Dict]) -> Dict[str, Dict]:
    """Grupperer uploadede filer baseret på den udtrukne nøgle fra filnavnet."""
    groups: Dict[str, Dict] = {}
    for item in upload_list:
        name = item["name"]
        b = item["bytes"]
        key, t = group_key_from_filename(name)
        
        if key not in groups:
            groups[key] = {"name": key, "csv": None, "render": None, "floorplan": None}
            
        if t == "csv":
            # Hvis flere CSV'er matcher, beholdes den første/sidste. Her bibeholdes den sidste (standard)
            groups[key]["csv"] = b 
        elif t == "render":
            if groups[key]["render"] is None:
                groups[key]["render"] = b
        elif t == "linedrawing":
            if groups[key]["floorplan"] is None:
                groups[key]["floorplan"] = b
                
    valid_groups = {k: v for k, v in groups.items() if v["csv"] is not None and v["render"] is not None}
    return valid_groups

def collect_all_renderings(groups: Dict[str, Dict]) -> List[bytes]:
    """Indsamler alle rendering image bytes fra grupperne."""
    return [g["render"] for g in groups.values() if g.get("render")]

def first_run_or_none(shape):
    """Finder det første 'run' i en shape for at bevare formatering."""
    try:
        tf = shape.text_frame
        if tf and tf.paragraphs and tf.paragraphs[0].runs:
            return tf.paragraphs[0].runs[0]
    except Exception:
        return None
    return None

def set_text_preserve_format(shape, text: str):
    """Sætter tekst i en shape, mens den forsøger at bevare den originale formatering."""
    try:
        if hasattr(shape, "text_frame") and shape.text_frame:
            run0 = first_run_or_none(shape)
            if run0:
                run0.text = text
            else:
                # Nogle former (f.eks. Placeholders) har ikke runs, men har tekst
                shape.text_frame.text = text 
    except Exception:
        pass

# --- Utils - Image Processing & Insertion ---

def add_picture_contain(slide, shape, image_bytes: bytes):
    """Indsætter et billede, sikrer at det passer (contain-fit) og konverterer til kompatibel JPEG/RGB."""
    try:
        if not image_bytes:
            return
        
        im = Image.open(io.BytesIO(image_bytes))
        
        # Konverterer til RGB for at sikre kompatibilitet med JPEG-lagring
        if im.mode in ("RGBA", "LA", "P"):
              im = im.convert("RGB")

        w, h = im.size
        
        # Nedskalerer store billeder (hvis de er over MAX_IMAGE_PX i dimension)
        max_dim = min(MAX_IMAGE_PX, max(w, h))
        scale_src_cap = min(1.0, max_dim / float(max(w, h)))
        if scale_src_cap < 1.0:
            im = im.resize((int(w * scale_src_cap), int(h * scale_src_cap)), Image.Resampling.LANCZOS)
            w, h = im.size

        # Beregner contain-fit forholdet til shape
        frame_w = int(shape.width)
        frame_h = int(shape.height)
        
        s = min(frame_w / w, frame_h / h)
        s = min(s, 1.0) # Sikrer ingen opskalering over 100%
        target_w = max(1, int(w * s))
        target_h = max(1, int(h * s))

        # Gemmer billedet som JPEG i bufferen
        buf = io.BytesIO()
        im.resize((target_w, target_h), Image.Resampling.LANCZOS).save(buf, format="JPEG", quality=JPEG_QUALITY, optimize=True)
        buf.seek(0)

        # Beregner centreret position
        left = shape.left + int((shape.width - target_w) / 2)
        top = shape.top + int((shape.height - target_h) / 2)
        
        # Indsætter billedet
        slide.shapes.add_picture(buf, left, top, width=target_w, height=target_h)
        
        # Fjerner den originale shape/placeholder, hvis den ikke er en rigtig placeholder
        try:
            if not getattr(shape, "is_placeholder", False):
                shape.element.getparent().remove(shape.element)
        except Exception:
            pass
        
    except Exception as e:
        st.error(f"Image processing failed (check file format/corruption): {type(e).__name__}: {str(e)}")
        return

def add_picture_into_shape(slide, shape, image_bytes: bytes):
    """Tvinger robust contain-fit for alle billeder."""
    if not image_bytes or shape is None:
        return
    add_picture_contain(slide, shape, image_bytes)

# --- PPTX Slide Builders ---

def add_table(slide, anchor_shape, rows: int, cols: int):
    """Opretter en tabel ved hjælp af anker-shapens position og størrelse."""
    try:
        left = getattr(anchor_shape, 'left', Inches(0.5))
        top = getattr(anchor_shape, 'top', Inches(1.2))
        width = getattr(anchor_shape, 'width', Inches(9.0))
        height = getattr(anchor_shape, 'height', Inches(5.0))
        
        table = slide.shapes.add_table(rows, cols, left, top, width, height).table
        return table
    except Exception:
        return None

def build_overview_slides(prs: Presentation, overview_layout, rendering_bytes_list: List[bytes]):
    """Bygger overview slides ved hjælp af navngivne placeholders i skabelonen."""
    for batch in chunk(rendering_bytes_list, MAX_OVERVIEW_IMAGES):
        slide = prs.slides.add_slide(overview_layout)
        shape_map = build_shape_map(slide)
        
        for idx, img_bytes in enumerate(batch, start=1):
            if not img_bytes: continue
            
            pic_key = clean_name(f"Rendering{idx}")
            target_shape = safe_find_shape(shape_map, pic_key)

            if target_shape:
                add_picture_into_shape(slide, target_shape, img_bytes)

def build_setting_slide(prs: Presentation,
                        setting_layout,
                        group_name: str,
                        render_bytes: Optional[bytes],
                        floorplan_bytes: Optional[bytes],
                        products_df: pd.DataFrame,
                        mapping_df: pd.DataFrame,
                        master_df: pd.DataFrame):
    """Bygger individuelle setting slides ved hjælp af navngivne placeholders i skabelonen."""
    slide = prs.slides.add_slide(setting_layout)
    shape_map = build_shape_map(slide)
    
    # 1. Title
    title_shape = safe_find_shape(shape_map, "SETTINGNAME")
    if title_shape:
        set_text_preserve_format(title_shape, group_name.replace("-", " ").title())
        
    # 2. Images
    render_shape = safe_find_shape(shape_map, "Rendering")
    if render_shape and render_bytes:
        add_picture_into_shape(slide, render_shape, render_bytes)
    elif render_shape:
        set_text_preserve_format(render_shape, "RENDERING IMAGE MISSING")
        
    floorplan_shape = safe_find_shape(shape_map, "Linedrawing") 
    if floorplan_shape and floorplan_bytes:
        add_picture_into_shape(slide, floorplan_shape, floorplan_bytes)
    elif floorplan_shape:
        set_text_preserve_format(floorplan_shape, "LINE DRAWING MISSING")
        
    # 3. Products and Descriptions
    for i in range(1, MAX_OVERVIEW_IMAGES + 1): 
        if i > len(products_df):
            break
            
        row = products_df.iloc[i-1] 
        article_no = row["ARTICLE_NO"]
        
        # Lookups
        pack_url = find_packshot_url(article_no, mapping_df, master_df)
        desc = find_description(article_no, mapping_df)
        
        # a) Packshot image (fetch and insert)
        pic_key = clean_name(f"ProductPackshot{i}")
        pic_shape = safe_find_shape(shape_map, pic_key)
        
        if pic_shape:
            # Undgår at bruge caching her, da billed-URL'er potentielt er unikke og der er en grænse for cache størrelse
            img_bytes = http_get_bytes(pack_url) if pack_url else None
            if img_bytes:
                add_picture_into_shape(slide, pic_shape, img_bytes)
            elif pack_url:
                set_text_preserve_format(pic_shape, "IMAGE URL FOUND, DOWNLOAD FAILED")
            else:
                set_text_preserve_format(pic_shape, f"PACKSHOT LOOKUP FAILED for {article_no}")
            
        # b) Product Description
        desc_key = clean_name(f"PRODUCT DESCRIPTION {i}")
        desc_shape = safe_find_shape(shape_map, desc_key)
        
        if desc_shape:
            set_text_preserve_format(desc_shape, desc)

def build_productlist_slide(prs: Presentation,
                            layout,
                            group_name: str,
                            products_df: pd.DataFrame,
                            mapping_df: pd.DataFrame):
    """Bygger produktliste-sliden med en tabel."""
    slide = prs.slides.add_slide(layout)
    shape_map = build_shape_map(slide)

    # 1. Title
    title_shape = safe_find_shape(shape_map, "Title") or safe_find_shape(shape_map, "SETTINGNAME")
    if title_shape:
        set_text_preserve_format(title_shape, f"Products – {group_name.replace('-', ' ').title()}")
        
    # 2. Table Anchor
    anchor = safe_find_shape(shape_map, "TableAnchor")
    
    if not anchor:
        st.warning(f"ADVARSEL: 'TableAnchor' missing in '{layout.name}'. Using default position.")
    
    rows = max(1, len(products_df)) + 1
    cols = 3
    table = add_table(slide, anchor, rows, cols)
    
    if table is None:
        return
        
    # 3. Fill Table
    # Sætter header-tekst
    table.cell(0, 0).text = "Quantity"
    table.cell(0, 1).text = "Description"
    table.cell(0, 2).text = "Article No. / New Item No."
    
    r = 1
    for row in products_df.itertuples(index=False):
        article_no = row.ARTICLE_NO
        
        # Lookups
        desc = find_description(article_no, mapping_df)
        new_item = find_new_item(article_no, mapping_df)
        
        table.cell(r, 0).text = str(int(row.Quantity))
        table.cell(r, 1).text = desc
        
        article_text = f"{article_no} / {new_item}" if new_item else f"{article_no}"
        table.cell(r, 2).text = article_text
        r += 1

# --- PPTX Utility Functions (Other) ---

def chunk(lst, n):
    """Deler en liste op i chunks af størrelse n."""
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

def safe_present(prs: Presentation) -> bytes:
    """GEMMER PRÆSENTATIONEN I EN IO.BYTESIO BUFFER OG RETURNERER BYTES. RETTELSE AF FEJL."""
    binary_stream = io.BytesIO()
    try:
        prs.save(binary_stream)
        binary_stream.seek(0)
        return binary_stream.read()
    except Exception as e:
        st.error(f"FEJL: Kunne ikke gemme præsentationen til bytes. {e}")
        return b''

def find_layout_by_name(prs: Presentation, target: str):
    """Finder et layout ved navn (clean match)."""
    t = clean_name(target)
    for layout in prs.slide_layouts:
        if clean_name(layout.name) == t:
            return layout
    # Tilføjer fallback for Renderings/Overview
    if t == clean_name('Renderings'):
        return find_layout_by_name(prs, 'Overview')
    if t == clean_name('Overview'):
        return find_layout_by_name(prs, 'Renderings')
    
    return None

# Bruger st.cache_resource til at indlæse PowerPoint-skabelonen én gang
@st.cache_resource
def ensure_presentation_from_path(path: Path) -> Presentation:
    """Sikrer, at skabelonfilen eksisterer og kan indlæses ved hjælp af caching."""
    if not path.exists():
        raise FileNotFoundError(f"Template not found: {path}")
    return Presentation(str(path))

def layout_has_expected(layout, keys: List[str]) -> bool:
    """Tjekker, om layoutet indeholder de forventede placeholders."""
    try:
        names = {clean_name(getattr(sh, "name", "")) for sh in layout.shapes}
    except Exception:
        names = set()
        
    return all(clean_name(k) in names for k in keys)

def preflight_checks() -> Dict[str, str]:
    """Udfører preflight checks for template og remote CSVs."""
    results = {"template": "OK"}
    
    if not TEMPLATE_PATH.exists():
        results["template"] = f"Template not found ({TEMPLATE_PATH})."
        return results
    
    try:
        # Indlæser skabelonen via cache for hurtighed
        prs = ensure_presentation_from_path(TEMPLATE_PATH) 
        
        # 1. Renderings Check
        if not find_layout_by_name(prs, 'Renderings'):
            results["template"] = "Mangler påkrævet layout 'Renderings' (Overview-side) i Slide Master."
            return results

        # 2. Setting Check
        if not find_layout_by_name(prs, 'Setting'):
            results["template"] = "Mangler påkrævet layout 'Setting' i Slide Master."
            return results

        # 3. ProductListBlank Check
        productlist_layout = find_layout_by_name(prs, 'ProductListBlank')
        if not productlist_layout:
            results["template"] = "Mangler påkrævet layout 'ProductListBlank' i Slide Master."
            return results
        
        # 4. TableAnchor Warning
        if not layout_has_expected(productlist_layout, ["TableAnchor"]):
            results["template"] += " ADVARSEL: TableAnchor mangler i ProductListBlank."
            
    except Exception as e:
        results["template"] = f"Fejl ved indlæsning af template: {e}"
        
    return results

# ---------------------- UI ----------------------
st.set_page_config(page_title="Muuto PowerPoint Generator", layout="centered")
st.title("Muuto PowerPoint Generator")
st.write("Upload dine gruppefiler (CSV og billeder). Appen bruger den faste PowerPoint-skabelon og henter Master Data og Mapping fra faste URLs.")

# Session state initialisering (kun uploads beholdes her)
if "uploads" not in st.session_state:
    st.session_state.uploads = []

# --- Data Caching Status (uden session state) ---
master_df = load_remote_csv_and_normalize(DEFAULT_MASTER_URL, normalize_master)
mapping_df = load_remote_csv_and_normalize(DEFAULT_MAPPING_URL, normalize_mapping)


files = st.file_uploader(
    "User group files (.csv, .jpg, .png, .webp). You can add multiple files.",
    type=["csv", "jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
)

if files:
    existing = {u["name"] for u in st.session_state.uploads}
    for f in files:
        if f.name not in existing:
            st.session_state.uploads.append({"name": f.name, "bytes": f.read()})
            existing.add(f.name)
    # Tvinger en genkørsel efter upload
    st.rerun()

# Single flat file list with remove buttons
if st.session_state.uploads:
    st.subheader("Uploaded Files")
    remove_indices = []
    for idx, item in enumerate(st.session_state.uploads):
        col1, col2 = st.columns([0.9, 0.1])
        with col1:
            size_kb = len(item["bytes"]) / 1024.0
            st.write(f"**{item['name']}** — {size_kb:.1f}KB")
        with col2:
            if st.button("❌", key=f"rm_{idx}"):
                remove_indices.append(idx)
    if remove_indices:
        for i in sorted(remove_indices, reverse=True):
            del st.session_state.uploads[i]
        st.rerun()

generate = st.button("Generate Presentation")

# ---------------------- Orchestration ----------------------
if generate:
    with st.spinner("Arbejder med generering..."):
        diag = preflight_checks()
        if "Mangler påkrævet layout" in diag["template"] or "Fejl ved indlæsning" in diag["template"]:
            st.error("Template problem: " + diag["template"])
        elif not TEMPLATE_PATH.exists():
            st.error("Template filen mangler i repository'et: input-template.pptx")
        else:
            try:
                groups = build_groups(st.session_state.uploads)

                if not groups:
                    st.error("Kunne ikke danne nogen grupper. Sørg for at filer er uploadet, og at filnavne indeholder CSV og mindst ét render/floorplan billede (adskilt af ' - ').")
                    
                if groups:
                    # Henter den cachede præsentation
                    prs = ensure_presentation_from_path(TEMPLATE_PATH)

                    overview_layout = find_layout_by_name(prs, "Renderings")
                    setting_layout = find_layout_by_name(prs, "Setting")
                    productlist_layout = find_layout_by_name(prs, "ProductListBlank")
                    
                    # --- DIAGNOSTIK OG ADVARSLER (bruger de cachede DFs) ---
                    if mapping_df.empty:
                        st.warning("ADVARSEL: Mapping Data (Beskrivelser/Nye Artikelnumre) kunne IKKE indlæses.")
                    
                    if master_df.empty:
                        st.warning("ADVARSEL: Master Data (Billed-URLs) kunne IKKE indlæses.")
                        
                    # 1. Overview Slides
                    renders = collect_all_renderings(groups)
                    if renders and overview_layout:
                        build_overview_slides(prs, overview_layout, renders)

                    # 2. Per group Slides
                    for key in sorted(groups.keys()):
                        g = groups[key]
                        group_name = g["name"]
                        
                        try:
                            pcon_df = normalize_pcon(parse_csv_flex(g["csv"]))
                        except Exception as e:
                            st.warning(f"FEJL ved parsing af CSV for '{group_name}': {e}")
                            pcon_df = pd.DataFrame(columns=["ARTICLE_NO", "Quantity"])
                        
                        if pcon_df.empty:
                            st.warning(f"ADVARSEL: Kunne ikke indlæse produktdata fra CSV for gruppe '{group_name}'. Springer slides over.")
                            continue 
                            
                        render_bytes = g.get("render")
                        floorplan_bytes = g.get("floorplan")

                        # Setting Slide
                        if setting_layout:
                            build_setting_slide(prs, setting_layout, group_name, render_bytes, floorplan_bytes, pcon_df, mapping_df, master_df)
                        else:
                            st.error(f"FATAL FEJL: Layout 'Setting' blev ikke fundet. Stop.")
                            raise Exception("Missing required layout 'Setting'")

                        # Product List Slide
                        if productlist_layout:
                            build_productlist_slide(prs, productlist_layout, group_name, pcon_df, mapping_df)
                        else:
                            st.error(f"FATAL FEJL: Layout 'ProductListBlank' blev ikke fundet. Stop.")
                            raise Exception("Missing required layout 'ProductListBlank'")

                    # Gemmer den genererede præsentation til bytes
                    ppt_bytes = safe_present(prs)
                    
                    if ppt_bytes:
                        st.success("Din præsentation er klar! 🎉")
                        st.download_button(
                            "Download Muuto_Settings.pptx",
                            data=ppt_bytes,
                            file_name=OUTPUT_NAME,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        )
                    else:
                        st.error("Kunne ikke gemme præsentationen. Se venligst fejlen ovenfor.")
                        
            except Exception as e:
                st.error(f"Der opstod en fejl under generering af præsentationen: {e}")
                
# UI for data status
st.subheader("Data Forbindelses Status")
col_m, col_mp = st.columns(2)
col_m.metric("Master Data Rækker", master_df.shape[0])
col_mp.metric("Mapping Data Rækker", mapping_df.shape[0])
if master_df.empty or mapping_df.empty:
    st.warning("ADVARSEL: Nul rækker indlæst fra Master/Mapping CSV'er. Tjek URL-tilgængelighed og kolonnenavne i CSV-dataen.")
