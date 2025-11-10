# app.py
import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import io, os, re, requests, csv
from typing import List, Dict, Any, Tuple

# ---------------- Konstanter ----------------
TEMPLATE_FILE = "input-template.pptx"

# Faste Google Sheets (skjult for brugeren)
MASTER_URL  = "https://docs.google.com/spreadsheets/d/1blj42SbFpszWGyOrDOUwyPDJr9K1NGpTMX6eZTbt_P4/edit?gid=1152340088#gid=1152340088"
MAPPING_URL = "https://docs.google.com/spreadsheets/d/1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk/edit?gid=1056617222#gid=1056617222"

# pCon CSV kolonneindeks og skip
PCON_SKIPROWS = 2
IDX_SHORT, IDX_VARIANT, IDX_ARTICLE, IDX_QTY = 2, 4, 17, 30

# Template tags
TAG_SETTINGNAME   = "{{SETTINGNAME}}"
TAG_PRODUCTS_LIST = "{{ProductsinSettingList}}"
TAG_RENDERING     = "{{Rendering}}"
TAG_LINEDRAWING   = "{{Linedrawing}}"
PACKSHOT_TAGS     = [f"{{{{ProductPackshot{i}}}}}" for i in range(1, 13)]
PROD_DESC_TAGS    = [f"{{{{PRODUCT DESCRIPTION {i}}}}}" for i in range(1, 13)]
OVERVIEW_TAGS     = [f"{{{{Rendering{i}}}}}" for i in range(1, 13)]

# ---------------- Helpers ----------------
def resolve_gsheet_to_csv(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url or "")
    if not m: return url
    sheet = m.group(1)
    gid_m = re.search(r"[#?&]gid=(\d+)", url)
    gid = gid_m.group(1) if gid_m else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet}/export?format=csv&gid={gid}"

def find_shape_contains(slide, tag: str):
    tagu = tag.upper()
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False) and shp.text_frame:
            txt = shp.text_frame.text or ""
            if tagu in txt.upper():
                return shp
    return None

def set_text_preserve_style(shape, text: str):
    if not shape or not shape.has_text_frame: return
    tf = shape.text_frame
    font_name = font_size = font_bold = None
    if tf.paragraphs and tf.paragraphs[0].runs:
        r0 = tf.paragraphs[0].runs[0]
        font_name, font_size, font_bold = r0.font.name, r0.font.size, r0.font.bold
    for p in list(tf.paragraphs):
        for r in list(p.runs): r.text = ""
    p = tf.paragraphs[0] if tf.paragraphs else tf.add_paragraph()
    p.clear()
    run = p.add_run(); run.text = text or ""
    if font_name: run.font.name = font_name
    if font_size: run.font.size = font_size
    if font_bold is not None: run.font.bold = font_bold
    tf.word_wrap = True

def replace_image_by_tag(slide, tag: str, img_bytes: bytes):
    if not img_bytes: return
    ph = find_shape_contains(slide, tag)
    if not ph: return
    left, top, w, h = ph.left, ph.top, ph.width, ph.height
    try: ph.element.getparent().remove(ph.element)
    except Exception: pass
    slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, width=w, height=h)

def duplicate_slide(prs: Presentation, slide) -> any:
    new_slide = prs.slides.add_slide(slide.slide_layout)
    while len(new_slide.shapes) > 0:
        sp = new_slide.shapes[0].element
        sp.getparent().remove(sp)
    for shp in slide.shapes:
        new_slide.shapes._spTree.append(shp.element.clone())
    return new_slide

# ---------------- Data-loaders ----------------
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
    col_img  = mapcol("IMAGE URL", ["Image URL","Image Link","Picture URL","Packshot URL","ImageURL","Image","IMAGE DOWNLOAD LINK"])
    if not col_item or not col_img:
        st.error("Master mangler kolonner: ITEM NO. og/eller IMAGE URL (IMAGE DOWNLOAD LINK accepteres)."); st.stop()
    out = df.rename(columns={col_item:"ITEM NO.", col_img:"IMAGE URL"})[["ITEM NO.","IMAGE URL"]]
    out["ITEM NO."] = out["ITEM NO."].astype(str).str.strip()
    out["IMAGE URL"] = out["IMAGE URL"].astype(str).str.strip()
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
    col_old = mapcol("OLD Item-variant", ["OLD Item variant","OLD_ITEM_VARIANT","Old Item","Old SKU","OLD ITEM NO."])
    col_new = mapcol("New Item No.",   ["New Item Number","NEW_ITEM_NO","New SKU","NEW ITEM NO."])
    col_desc = mapcol("Description",   ["DESCRIPTION","Product Description","DESC","Name"])
    if not col_old or not col_new or not col_desc:
        st.error("Mapping mangler kolonner: OLD Item-variant, New Item No., Description."); st.stop()
    out = df.rename(columns={
        col_old:"OLD Item-variant",
        col_new:"New Item No.",
        col_desc:"Description"
    })[["OLD Item-variant","New Item No.","Description"]]
    for c in out.columns: out[c] = out[c].astype(str).str.strip()
    return out

# ---------------- Robust pCon CSV ----------------
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
    sub["ARTICLE_NO"]   = sub["ARTICLE_NO"].astype(str).str.strip()
    sub["SHORT_TEXT"]   = sub["SHORT_TEXT"].astype(str).str.strip()
    sub["VARIANT_TEXT"] = sub["VARIANT_TEXT"].astype(str).str.strip()
    sub["QUANTITY"]     = pd.to_numeric(sub["QUANTITY"], errors="coerce").fillna(1).astype(int)
    sub = sub[sub["ARTICLE_NO"].ne("")]
    if sub.empty:
        raise ValueError("pCon CSV blev læst, men indeholdt ingen gyldige rækker med ARTICLE_NO.")
    return sub

# ---------------- Lookups ----------------
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

# ---------------- Billeder ----------------
@st.cache_data(ttl=3600)
def fetch_image(url: str) -> bytes | None:
    if not url or not url.startswith("http"): return None
    try:
        r = requests.get(url, timeout=15); r.raise_for_status()
        if "image" not in r.headers.get("Content-Type","").lower(): return None
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
        buf = io.BytesIO(); im.save(buf, format="JPEG", quality=quality); return buf.getvalue()
    except Exception:
        return img

# ---------------- PPT byggesten ----------------
def find_first_slide_with_tag(prs: Presentation, tag: str) -> Tuple[Any, int]:
    for i, sl in enumerate(prs.slides):
        if find_shape_contains(sl, tag): return sl, i
    return None, -1

def add_products_table_on_blank(prs: Presentation, base_slide_for_layout, title: str, rows: List[List[str]]):
    if base_slide_for_layout is not None:
        s = duplicate_slide(prs, base_slide_for_layout)
        ph = find_shape_contains(s, TAG_PRODUCTS_LIST)
        if ph: set_text_preserve_style(ph, "")
    else:
        blank_layout = None
        for ly in prs.slide_layouts:
            if not ly.name or "blank" in ly.name.lower():
                blank_layout = ly; break
        s = prs.slides.add_slide(blank_layout if blank_layout else prs.slide_layouts[1])
    if s.shapes.title:
        s.shapes.title.text = title
    headers = ["Quantity", "Description", "Article No. / New Item No."]
    data = [headers] + rows
    left, top, width, height = Inches(0.6), Inches(1.8), Inches(9.2), Inches(5.2)
    tbl_shape = s.shapes.add_table(rows=len(data), cols=3, left=left, top=top, width=width, height=height)
    tbl = tbl_shape.table
    for r_i, row in enumerate(data):
        for c_i, val in enumerate(row):
            cell = tbl.cell(r_i, c_i); cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(12)
    return s

def build_presentation(master_df: pd.DataFrame,
                       mapping_df: pd.DataFrame,
                       groups: List[Dict[str, Any]],
                       overview_renderings: List[bytes]) -> bytes:
    prs = Presentation(TEMPLATE_FILE)

    # Find skabelon-slides
    setting_tpl, _  = find_first_slide_with_tag(prs, TAG_SETTINGNAME)
    list_tpl,   _   = find_first_slide_with_tag(prs, TAG_PRODUCTS_LIST)
    overview_tpl, _ = find_first_slide_with_tag(prs, "OVERVIEW")

    # OVERVIEW
    if overview_tpl and overview_renderings:
        s = duplicate_slide(prs, overview_tpl)
        for i, rb in enumerate(overview_renderings[:12]):
            replace_image_by_tag(s, OVERVIEW_TAGS[i], preprocess(rb))

    for g in groups:
        # A) Setting-slide
        base = setting_tpl if setting_tpl else prs.slides[0]
        s = duplicate_slide(prs, base)
        set_text_preserve_style(find_shape_contains(s, TAG_SETTINGNAME), g["name"])
        if g.get("rendering_bytes"):
            replace_image_by_tag(s, TAG_RENDERING, preprocess(g["rendering_bytes"]))
        if g.get("linedrawing_bytes"):
            replace_image_by_tag(s, TAG_LINEDRAWING, preprocess(g["linedrawing_bytes"]))

        # Packshots + PRODUCT DESCRIPTION tags
        for idx, it in enumerate(g["items"][:12]):
            # packshot image
            if it.get("packshot_url"):
                raw = fetch_image(it["packshot_url"])
                if raw: replace_image_by_tag(s, PACKSHOT_TAGS[idx], preprocess(raw))
            # product description text from mapping
            desc_tag = PROD_DESC_TAGS[idx]
            set_text_preserve_style(find_shape_contains(s, desc_tag), it.get("desc_text",""))

        # B) Produkttabel (Description fra mapping)
        rows = []
        for it in g["items"]:
            id_combo = it["article_no"]
            if it.get("new_item_no"):
                id_combo = f'{it["article_no"]} / {it["new_item_no"]}'
            rows.append([it["qty"], it["desc_text"], id_combo])
        add_products_table_on_blank(prs, list_tpl, f"Products – {g['name']}", rows)

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# ---------------- UI ----------------
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

        master_df  = load_master()
        mapping_df = load_mapping()

        # Gruppér filer efter prefix før første ' - ' ellers første '_'/'-'
        groups_map: Dict[str, Dict[str, Any]] = {}
        for f in uploads or []:
            name, ext = os.path.splitext(f.name)
            base = name.split(" - ", 1)[0].strip() if " - " in name else re.split(r"[_-]", name, 1)[0].strip()
            groups_map.setdefault(base, {"name": base.title(), "csv": None, "rendering": None, "line": None})
            lf = f.name.lower()
            if lf.endswith(".csv"):
                groups_map[base]["csv"] = f
            elif any(k in lf for k in ["line","floorplan","drawing"]):
                groups_map[base]["line"] = f
            elif lf.endswith((".jpg",".jpeg",".png")):
                groups_map[base]["rendering"] = f

        settings, overview_imgs = [], []
        for base, data in groups_map.items():
            if not data["csv"] or not data["rendering"]:
                st.warning(f"Ignorerer '{data['name']}' – mangler CSV eller Rendering."); continue

            df = pcon_from_csv(data["csv"])

            items = []
            for _, r in df.iterrows():
                article = r["ARTICLE_NO"]
                qty = int(r["QUANTITY"])
                # description fra mapping
                desc = mapping_description(mapping_df, article)
                pack = packshot_lookup(master_df, article)
                newno = new_item_lookup(mapping_df, article)
                items.append({
                    "article_no": article,
                    "qty": qty,
                    "desc_text": desc,
                    "packshot_url": pack,
                    "new_item_no": newno
                })

            render_bytes = data["rendering"].read(); overview_imgs.append(render_bytes)
            line_bytes = data["line"].read() if data["line"] else None

            settings.append({
                "name": data["name"],
                "items": items,
                "rendering_bytes": render_bytes,
                "linedrawing_bytes": line_bytes
            })

        if not settings:
            st.error("Ingen gyldige settings fundet."); st.stop()

        ppt_bytes = build_presentation(master_df, mapping_df, settings, overview_imgs)
        st.success("Præsentation genereret.")
        st.download_button(
            "Download PPT",
            data=ppt_bytes,
            file_name="Muuto_Settings.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

if __name__ == "__main__":
    main()
