import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import io, os, re, requests
from typing import List, Dict, Any

# ---------------- Const ----------------
TEMPLATE_FILE = "input-template.pptx"

# pCon CSV: faste kolonneindekser og skip
PCON_SKIPROWS = 2
IDX_SHORT, IDX_VARIANT, IDX_ARTICLE, IDX_QTY = 2, 4, 17, 30

# Template placeholders vi understøtter
PACKSHOT_PLACEHOLDERS = [f"{{{{ProductPackshot{i}}}}}" for i in range(1, 13)]
RENDERING_TAG = "{{Rendering}}"
LINEDRAWING_TAG = "{{Linedrawing}}"

# Master kolonnenavne efter mapping
MASTER_PACKSHOT_CANON = "IMAGE URL"

# ---------------- URL helpers ----------------
def resolve_gsheet_to_csv(url: str) -> str:
    if not isinstance(url, str) or not url.startswith(("http://", "https://")):
        return url
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if not m:
        return url
    sheet = m.group(1)
    gidm = re.search(r"[#?&]gid=(\d+)", url)
    gid = gidm.group(1) if gidm else "0"
    return f"https://docs.google.com/spreadsheets/d/{sheet}/export?format=csv&gid={gid}"

# ---------------- Loaders ----------------
@st.cache_data
def load_master(master_url: str) -> pd.DataFrame:
    url = resolve_gsheet_to_csv(master_url) if master_url.startswith("http") else master_url
    r = requests.get(url, timeout=20); r.raise_for_status()
    df = pd.read_csv(io.BytesIO(r.content))
    # kolonne-synonymer
    def norm(s): return re.sub(r"[\s_.-]+","",str(s).strip().lower())
    norm_map = {norm(c): c for c in df.columns}
    def mapcol(canon, alts):
        for a in [canon]+alts:
            if norm(a) in norm_map: return norm_map[norm(a)]
        return None
    col_item = mapcol("ITEM NO.", ["Item No.","ITEM","SKU","Item Number","ItemNo","ITEM_NO"])
    col_img  = mapcol("IMAGE URL", [     "Image URL", "Image Link", "Picture URL", "Packshot URL",     "ImageURL", "Image", "IMAGE DOWNLOAD LINK" ])
    missing = [n for n,v in {"ITEM NO.":col_item,"IMAGE URL":col_img}.items() if v is None]
    if missing:
        st.error(f"Master mangler kolonner: {', '.join(missing)}"); st.stop()
    df = df.rename(columns={col_item:"ITEM NO.", col_img:"IMAGE URL"})
    df["ITEM NO."] = df["ITEM NO."].astype(str).str.strip()
    return df[["ITEM NO.","IMAGE URL"]]

@st.cache_data
def load_mapping(mapping_url: str) -> pd.DataFrame:
    """Mapping: OLD Item-variant -> New Item No."""
    url = resolve_gsheet_to_csv(mapping_url) if mapping_url.startswith("http") else mapping_url
    r = requests.get(url, timeout=20); r.raise_for_status()
    df = pd.read_csv(io.BytesIO(r.content))
    def norm(s): return re.sub(r"[\s_.-]+","",str(s).strip().lower())
    norm_map = {norm(c): c for c in df.columns}
    def mapcol(canon, alts):
        for a in [canon]+alts:
            if norm(a) in norm_map: return norm_map[norm(a)]
        return None
    col_old = mapcol("OLD Item-variant", ["OLD Item variant","OLD_ITEM_VARIANT","Old Item","Old SKU"])
    col_new = mapcol("New Item No.", ["New Item Number","NEW_ITEM_NO","New SKU"])
    missing = [n for n,v in {"OLD Item-variant":col_old,"New Item No.":col_new}.items() if v is None]
    if missing:
        st.error(f"Mapping mangler kolonner: {', '.join(missing)}"); st.stop()
    df = df.rename(columns={col_old:"OLD Item-variant", col_new:"New Item No."})
    df["OLD Item-variant"] = df["OLD Item-variant"].astype(str).str.strip()
    df["New Item No."] = df["New Item No."].astype(str).str.strip()
    return df[["OLD Item-variant","New Item No."]]

# ---------------- pCon CSV ----------------
def pcon_from_csv(f) -> pd.DataFrame:
    df = pd.read_csv(f, skiprows=PCON_SKIPROWS, header=None)
    need = max(IDX_SHORT, IDX_VARIANT, IDX_ARTICLE, IDX_QTY)
    if df.shape[1] <= need:
        st.error("pCon CSV har for få kolonner."); st.stop()
    out = df.iloc[:, [IDX_SHORT, IDX_VARIANT, IDX_ARTICLE, IDX_QTY]].copy()
    out.columns = ["SHORT_TEXT","VARIANT_TEXT","ARTICLE_NO","QUANTITY"]
    out["ARTICLE_NO"] = out["ARTICLE_NO"].astype(str).str.strip()
    out["SHORT_TEXT"] = out["SHORT_TEXT"].astype(str).str.strip()
    out["VARIANT_TEXT"] = out["VARIANT_TEXT"].astype(str).str.strip()
    out["QUANTITY"] = pd.to_numeric(out["QUANTITY"], errors="coerce").fillna(1).astype(int)
    return out[out["ARTICLE_NO"].ne("")]

def fallback_key(article: str) -> str:
    base = re.sub(r"^SPECIAL-", "", str(article), flags=re.I)
    return base.split("-")[0].strip()

# ---------------- Lookups ----------------
def packshot_lookup(df_master: pd.DataFrame, article: str) -> str:
    hit = df_master.loc[df_master["ITEM NO."] == article, "IMAGE URL"]
    if not hit.empty: return str(hit.iloc[0])
    base = fallback_key(article)
    hit = df_master.loc[df_master["ITEM NO."].apply(fallback_key) == base, "IMAGE URL"]
    return str(hit.iloc[0]) if not hit.empty else ""

def new_item_lookup(df_map: pd.DataFrame, article: str) -> str:
    hit = df_map.loc[df_map["OLD Item-variant"] == article, "New Item No."]
    if not hit.empty: return str(hit.iloc[0])
    base = fallback_key(article)
    hit = df_map.loc[df_map["OLD Item-variant"].apply(fallback_key) == base, "New Item No."]
    return str(hit.iloc[0]) if not hit.empty else ""

# ---------------- Images ----------------
def fetch_image(url: str) -> bytes | None:
    if not url or not url.startswith("http"): return None
    try:
        r = requests.get(url, timeout=15); r.raise_for_status()
        if "image" not in r.headers.get("Content-Type","").lower(): return None
        return r.content
    except requests.RequestException:
        return None

def preprocess(img: bytes, max_side=1200, quality=85) -> bytes:
    try:
        im = Image.open(io.BytesIO(img))
        if im.mode != "RGB": im = im.convert("RGB")
        if max(im.size) > max_side:
            ratio = min(max_side/im.width, max_side/im.height)
            im = im.resize((int(im.width*ratio), int(im.height*ratio)), Image.Resampling.LANCZOS)
        buf = io.BytesIO(); im.save(buf, format="JPEG", quality=quality); return buf.getvalue()
    except Exception: return img

# ---------------- PPT helpers ----------------
def find_shape_by_text(slide, tag: str):
    tagu = tag.upper()
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False) and shp.text_frame:
            if shp.text_frame.text.strip().upper() == tagu:
                return shp
    return None

def set_text(shape, text: str):
    if not shape or not shape.has_text_frame: return
    tf = shape.text_frame
    for p in list(tf.paragraphs):
        for r in list(p.runs): r.text = ""
    p = tf.paragraphs[0] if tf.paragraphs else tf.add_paragraph()
    p.clear(); p.add_run().text = text or ""

def replace_image(slide, tag: str, img_bytes: bytes):
    ph = find_shape_by_text(slide, tag)
    if not ph or not img_bytes: return
    left, top, w, h = ph.left, ph.top, ph.width, ph.height
    try: ph.element.getparent().remove(ph.element)
    except Exception: pass
    stream = io.BytesIO(img_bytes)
    slide.shapes.add_picture(stream, left, top, width=w, height=h)

def add_table_slide(prs, title: str, rows: List[List[str]]):
    layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
    s = prs.slides.add_slide(layout)
    if s.shapes.title: s.shapes.title.text = title
    headers = ["Quantity", "Short Text", "Article No. / New Item No."]
    data = [headers] + rows
    left, top, width, height = Inches(0.5), Inches(1.8), Inches(9), Inches(5)
    tbl_shape = s.shapes.add_table(rows=len(data), cols=3, left=left, top=top, width=width, height=height)
    tbl = tbl_shape.table
    for r_i, row in enumerate(data):
        for c_i, val in enumerate(row):
            cell = tbl.cell(r_i, c_i)
            cell.text = str(val)
            for p in cell.text_frame.paragraphs:
                for run in p.runs: run.font.size = Pt(12)
    return s

# ---------------- Build PPT ----------------
def build_presentation(master_df: pd.DataFrame,
                       mapping_df: pd.DataFrame,
                       groups: List[Dict[str, Any]],
                       overview_renderings: List[bytes]) -> bytes:
    prs = Presentation(TEMPLATE_FILE)

    # OVERVIEW
    try:
        over_slide, over_idx = None, None
        for i, sl in enumerate(prs.slides):
            if find_shape_by_text(sl, "OVERVIEW"): over_slide, over_idx = sl, i; break
        layout = over_slide.slide_layout if over_slide else prs.slide_layouts[0]
        s = prs.slides.add_slide(layout)
        if find_shape_by_text(s, "{{Rendering1}}"):
            for j, rb in enumerate(overview_renderings[:12]):
                replace_image(s, f"{{{{Rendering{j+1}}}}}", preprocess(rb))
        else:
            if overview_renderings:
                replace_image(s, RENDERING_TAG, preprocess(overview_renderings[0]))
        if over_idx is not None:
            prs.slides._sldIdLst.remove(prs.slides._sldIdLst[over_idx])
    except Exception:
        pass

    # setting-layout via {{SETTINGNAME}}
    setting_layout, tmpl_idx = None, None
    for i, sl in enumerate(prs.slides):
        if find_shape_by_text(sl, "{{SETTINGNAME}}"):
            setting_layout, tmpl_idx = sl.slide_layout, i; break
    if setting_layout is None:
        setting_layout = prs.slide_layouts[0]

    for g in groups:
        # slide 1: setting
        s = prs.slides.add_slide(setting_layout)
        set_text(find_shape_by_text(s, "{{SETTINGNAME}}"), g["name"])
        if g.get("rendering_bytes"):
            replace_image(s, RENDERING_TAG, preprocess(g["rendering_bytes"]))
        if g.get("linedrawing_bytes"):
            replace_image(s, LINEDRAWING_TAG, preprocess(g["linedrawing_bytes"]))

        # udfyld packshots
        for idx, it in enumerate(g["items"][:12]):
            if it.get("packshot_url"):
                raw = fetch_image(it["packshot_url"])
                if raw: replace_image(s, PACKSHOT_PLACEHOLDERS[idx], preprocess(raw))

        # slide 2: produkttabel med New Item No.
        rows = []
        for it in g["items"]:
            combo = it["article_no"]
            if it.get("new_item_no"):
                combo = f'{it["article_no"]} / {it["new_item_no"]}'
            rows.append([it["qty"], it["short_text"], combo])
        add_table_slide(prs, f"Products – {g['name']}", rows)

    if tmpl_idx is not None:
        prs.slides._sldIdLst.remove(prs.slides._sldIdLst[tmpl_idx])

    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    return buf.getvalue()

# ---------------- UI ----------------
def main():
    st.set_page_config(page_title="Muuto PPT Generator", layout="wide")
    st.title("Muuto PPT Generator")

    master_url = st.text_input(
        "Master-data Google Sheets URL (edit eller export)",
        value="https://docs.google.com/spreadsheets/d/1blj42SbFpszWGyOrDOUwyPDJr9K1NGpTMX6eZTbt_P4/edit?gid=1152340088#gid=1152340088"
    )
    mapping_url = st.text_input(
        "Mapping (OLD Item-variant → New Item No.) Google Sheets URL (edit eller export)",
        value="https://docs.google.com/spreadsheets/d/1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk/edit?gid=1056617222#gid=1056617222"
    )

    uploads = st.file_uploader(
        "Upload alle settings: CSV + Rendering JPG/PNG + valgfri Linedrawing. Mindst CSV og Rendering pr. setting.",
        type=["csv","jpg","jpeg","png"],
        accept_multiple_files=True
    )

    if st.button("Generér PPT", type="primary"):
        if not os.path.exists(TEMPLATE_FILE):
            st.error(f"Skabelon '{TEMPLATE_FILE}' mangler."); st.stop()

        master_df = load_master(master_url)
        mapping_df = load_mapping(mapping_url)

        # gruppering pr. prefix før " - "
        groups_map: Dict[str, Dict[str, Any]] = {}
        for f in uploads or []:
            name, ext = os.path.splitext(f.name)
            base = name.split(" - ", 1)[0].strip() if " - " in name else re.split(r"[_-]", name, 1)[0].strip()
            groups_map.setdefault(base, {"name": base.title(), "csv": None, "rendering": None, "line": None})
            lf = f.name.lower()
            if lf.endswith(".csv"): groups_map[base]["csv"] = f
            elif any(k in lf for k in ["line","floorplan","drawing"]): groups_map[base]["line"] = f
            elif lf.endswith((".jpg",".jpeg",".png")): groups_map[base]["rendering"] = f

        settings, overview_imgs = [], []
        for base, data in groups_map.items():
            if not data["csv"] or not data["rendering"]:
                st.warning(f"Ignorerer '{data['name']}' – mangler CSV eller Rendering."); continue

            df = pcon_from_csv(data["csv"])
            items = []
            for _, r in df.iterrows():
                article = r["ARTICLE_NO"]
                qty = int(r["QUANTITY"])
                short = r["SHORT_TEXT"] if (not r["VARIANT_TEXT"] or r["VARIANT_TEXT"].upper()=="LIGHT OPTION: OFF") else f"{r['SHORT_TEXT']} – {r['VARIANT_TEXT']}"
                pack = packshot_lookup(master_df, article)
                newno = new_item_lookup(mapping_df, article)
                items.append({"article_no": article, "qty": qty, "short_text": short, "packshot_url": pack, "new_item_no": newno})

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
