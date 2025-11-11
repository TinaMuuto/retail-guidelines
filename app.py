import io
import re
from typing import Dict, List, Optional, Tuple
import streamlit as st
import pandas as pd
from PIL import Image
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
import requests
import time
from pathlib import Path

# -----------------------------
# Constants and defaults
# -----------------------------
TEMPLATE_PATH = Path("input-template.pptx")  # fixed template in repo root
DEFAULT_MASTER_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRdNwE1Q_aG3BntCZZPRIOgXEFJ5AHJxHmRgirMx2FJqfttgCZ8on-j1vzxM-muTTvtAHwc-ovDV1qF/pub?output=csv"
DEFAULT_MAPPING_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQPRmVmc0LYISduQdJyfz-X3LJlxiEDCNwW53LhFsWp5fFDS8V669rCd9VGoygBZSAZXeSNZ5fquPen/pub?output=csv"
OUTPUT_NAME = "Muuto_Settings.pptx"
MAX_OVERVIEW_IMAGES = 12
HTTP_TIMEOUT = 10
HTTP_RETRIES = 1
MAX_IMAGE_PX = 1400
JPEG_QUALITY = 85

# -----------------------------
# Utilities
# -----------------------------
def clean_name(name: str) -> str:
    if name is None:
        return ""
    name = name.strip()
    name = re.sub(r"^\{\{|\}\}$", "", name).strip()
    return re.sub(r"\s+", "", name).lower()

def first_run_or_none(shape):
    try:
        tf = shape.text_frame
        if tf and tf.paragraphs and tf.paragraphs[0].runs:
            return tf.paragraphs[0].runs[0]
    except Exception:
        return None
    return None

def set_text_preserve_format(shape, text: str):
    try:
        if hasattr(shape, "text_frame") and shape.text_frame:
            run0 = first_run_or_none(shape)
            if run0:
                run0.text = text
            else:
                shape.text_frame.text = text
    except Exception:
        pass

def build_shape_map(slide) -> Dict[str, object]:
    mapping = {}
    for shape in slide.shapes:
        try:
            nm = clean_name(getattr(shape, "name", ""))
            if nm:
                mapping[nm] = shape
        except Exception:
            continue
    return mapping

def add_picture_contain(slide, shape, image_bytes: bytes):
    try:
        if not image_bytes:
            return
        with Image.open(io.BytesIO(image_bytes)) as im:
            im = im.convert("RGB")
            w, h = im.size
            max_dim = min(MAX_IMAGE_PX, max(w, h))
            scale_src_cap = min(1.0, max_dim / float(max(w, h)))
            if scale_src_cap < 1.0:
                im = im.resize((int(w * scale_src_cap), int(h * scale_src_cap)), Image.LANCZOS)

            frame_w = int(shape.width)
            frame_h = int(shape.height)
            s = min(frame_w / im.width, frame_h / im.height)
            s = min(s, 1.0)  # do not upscale beyond current
            target_w = max(1, int(im.width * s))
            target_h = max(1, int(im.height * s))

            out_buf = io.BytesIO()
            im.resize((target_w, target_h), Image.LANCZOS).save(out_buf, format="JPEG", quality=JPEG_QUALITY, optimize=True)
            out_buf.seek(0)

            left = shape.left + int((shape.width - target_w) / 2)
            top = shape.top + int((shape.height - target_h) / 2)
            slide.shapes.add_picture(out_buf, left, top, width=target_w, height=target_h)
    except Exception:
        return

def add_table(slide, anchor_shape, rows: int, cols: int):
    try:
        table = slide.shapes.add_table(rows, cols, anchor_shape.left, anchor_shape.top, anchor_shape.width, anchor_shape.height).table
        return table
    except Exception:
        return None

def http_get_bytes(url: str) -> Optional[bytes]:
    if not url:
        return None
    last_err = None
    for attempt in range(HTTP_RETRIES + 1):
        try:
            resp = requests.get(url, timeout=HTTP_TIMEOUT, allow_redirects=True)
            if resp.status_code == 200 and resp.content:
                return resp.content
            last_err = f"HTTP {resp.status_code}"
        except Exception as e:
            last_err = str(e)
        time.sleep(0.2 * attempt)
    return None

def parse_csv_flex(buf: bytes) -> pd.DataFrame:
    if buf is None:
        return pd.DataFrame()
    candidates = [
        {"sep": ",", "encoding": "utf-8"},
        {"sep": ";", "encoding": "utf-8"},
        {"sep": "\t", "encoding": "utf-8"},
        {"sep": ",", "encoding": "latin-1"},
        {"sep": ";", "encoding": "latin-1"},
    ]
    for c in candidates:
        try:
            return pd.read_csv(io.BytesIO(buf), sep=c["sep"], encoding=c["encoding"])
        except Exception:
            continue
    return pd.DataFrame()

def group_key_from_filename(name: str) -> Tuple[str, str]:
    base = Path(name).stem
    m = re.split(r"[-_]", base, maxsplit=1)
    prefix = m[0].strip() if m else base.strip()
    lname = name.lower()
    if "floorplan" in lname:
        t = "floorplan"
    elif "linedrawing" in lname or "line_drawing" in lname:
        t = "linedrawing"
    else:
        ext = Path(name).suffix.lower()
        if ext == ".csv":
            t = "csv"
        elif ext in [".jpg", ".jpeg", ".png"]:
            t = "render"
        else:
            t = "other"
    return prefix, t

def base_before_dash(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if pd.notna(s) else ""
    return s.split("-")[0].strip()

def find_layout_by_name(prs: Presentation, target: str):
    t = clean_name(target)
    for layout in prs.slide_layouts:
        if clean_name(layout.name) == t:
            return layout
    for layout in prs.slide_layouts:
        if t in clean_name(layout.name):
            return layout
    return None

def ensure_presentation_from_path(path: Path) -> Presentation:
    if not path.exists():
        raise FileNotFoundError(f"Template not found: {path}")
    return Presentation(str(path))

# -----------------------------
# Domain logic
# -----------------------------
def load_remote_csv(url: str) -> pd.DataFrame:
    content = http_get_bytes(url)
    if content is None:
        return pd.DataFrame()
    df = parse_csv_flex(content)
    return df

def normalize_master(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["ITEM NO.", "IMAGE"])
    cols = {c: c.strip() for c in df.columns}
    df = df.rename(columns=cols)
    img_col = None
    for c in df.columns:
        if c.upper() in ["IMAGE URL", "IMAGE DOWNLOAD LINK"]:
            img_col = c
            break
    if img_col is None:
        for c in df.columns:
            if "image" in c.lower() and ("url" in c.lower() or "download" in c.lower()):
                img_col = c
                break
    item_col = None
    for c in df.columns:
        if c.strip().upper() == "ITEM NO.":
            item_col = c
            break
    if item_col is None:
        for c in df.columns:
            if "item" in c.lower() and "no" in c.lower():
                item_col = c
                break
    if item_col is None or img_col is None:
        return pd.DataFrame(columns=["ITEM NO.", "IMAGE"])
    out = df[[item_col, img_col]].copy()
    out.columns = ["ITEM NO.", "IMAGE"]
    out["ITEM BASE"] = out["ITEM NO."].apply(base_before_dash)
    return out

def normalize_mapping(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["OLD Item-variant", "Description", "New Item No."])
    cols = {c: c.strip() for c in df.columns}
    df = df.rename(columns=cols)
    col_old = None
    col_desc = None
    col_new = None
    for c in df.columns:
        if c.lower().strip() in ["old item-variant", "old item variant", "olditem-variant"]:
            col_old = c
        if c.lower().strip() == "description":
            col_desc = c
        if c.lower().strip() in ["new item no.", "new item no", "new item number"]:
            col_new = c
    if col_old is None:
        for c in df.columns:
            if "old" in c.lower() and "variant" in c.lower():
                col_old = c; break
    if col_new is None:
        for c in df.columns:
            if "new" in c.lower() and ("no" in c.lower() or "number" in c.lower()):
                col_new = c; break
    if col_desc is None:
        for c in df.columns:
            if "desc" in c.lower():
                col_desc = c; break
    if not col_old or not col_new:
        return pd.DataFrame(columns=["OLD Item-variant", "Description", "New Item No."])
    if col_desc is None:
        df["__desc__"] = ""
        col_desc = "__desc__"
    out = df[[col_old, col_desc, col_new]].copy()
    out.columns = ["OLD Item-variant", "Description", "New Item No."]
    out["OLD BASE"] = out["OLD Item-variant"].apply(base_before_dash)
    out["NEW BASE"] = out["New Item No."].apply(base_before_dash)
    return out

def normalize_pcon(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["ARTICLE_NO", "Quantity"])
    mapping = {}
    for c in df.columns:
        cl = c.strip().lower().replace(" ", "").replace("_", "")
        if cl in ["articleno", "article", "articlenumber", "article_no"]:
            mapping[c] = "ARTICLE_NO"
        if cl in ["qty", "quantity", "quantities"]:
            mapping[c] = "Quantity"
    if "ARTICLE_NO" not in mapping.values():
        for c in df.columns:
            if c.strip().upper() == "ARTICLE_NO":
                mapping[c] = "ARTICLE_NO"
                break
    if "Quantity" not in mapping.values():
        df["__qty__"] = 1
        mapping["__qty__"] = "Quantity"
    out = df.rename(columns=mapping)
    cols = [k for k, v in mapping.items() if v in ["ARTICLE_NO", "Quantity"]]
    out = out[cols].copy()
    out.columns = ["ARTICLE_NO", "Quantity"]
    out["ARTICLE_BASE"] = out["ARTICLE_NO"].apply(base_before_dash)
    out["Quantity"] = pd.to_numeric(out["Quantity"], errors="coerce").fillna(1).astype(int)
    return out

def find_packshot_url(article_no: str, mapping_df: pd.DataFrame, master_df: pd.DataFrame) -> Optional[str]:
    if master_df is None or master_df.empty:
        return None
    if mapping_df is not None and not mapping_df.empty:
        row = mapping_df[mapping_df["OLD Item-variant"].astype(str) == str(article_no)]
        if row.empty:
            row = mapping_df[mapping_df["OLD BASE"].astype(str) == base_before_dash(article_no)]
        if not row.empty:
            new_item = row.iloc[0]["New Item No."]
            if pd.notna(new_item):
                m = master_df[master_df["ITEM NO."].astype(str) == str(new_item)]
                if m.empty:
                    m = master_df[master_df["ITEM BASE"].astype(str) == base_before_dash(str(new_item))]
                if not m.empty:
                    return m.iloc[0]["IMAGE"]
    m = master_df[master_df["ITEM NO."].astype(str) == str(article_no)]
    if m.empty:
        m = master_df[master_df["ITEM BASE"].astype(str) == base_before_dash(str(article_no))]
    if not m.empty:
        return m.iloc[0]["IMAGE"]
    return None

def find_description(article_no: str, mapping_df: pd.DataFrame) -> str:
    if mapping_df is None or mapping_df.empty:
        return ""
    row = mapping_df[mapping_df["OLD Item-variant"].astype(str) == str(article_no)]
    if row.empty:
        row = mapping_df[mapping_df["OLD BASE"].astype(str) == base_before_dash(article_no)]
    if not row.empty:
        desc = row.iloc[0]["Description"]
        return "" if pd.isna(desc) else str(desc)
    return ""

def find_new_item(article_no: str, mapping_df: pd.DataFrame) -> Optional[str]:
    if mapping_df is None or mapping_df.empty:
        return None
    row = mapping_df[mapping_df["OLD Item-variant"].astype(str) == str(article_no)]
    if row.empty:
        row = mapping_df[mapping_df["OLD BASE"].astype(str) == base_before_dash(article_no)]
    if not row.empty:
        val = row.iloc[0]["New Item No."]
        return None if pd.isna(val) else str(val)
    return None

def chunk(lst, n):
    for i in range(0, len(lst), n):
        yield lst[i:i+n]

# -----------------------------
# Fallback slide creators
# -----------------------------
def get_blank_layout(prs: Presentation):
    for layout in prs.slide_layouts:
        if clean_name(layout.name) in ("blank", "empty"):
            return layout
    return prs.slide_layouts[0]

def create_overview_slide_fallback(prs: Presentation, images_batch):
    slide = prs.slides.add_slide(get_blank_layout(prs))
    cols, rows = 4, 3
    margin_x, margin_y = Inches(0.5), Inches(1.0)
    cell_w = Inches(2.0)
    cell_h = Inches(1.5)
    for idx, img_bytes in enumerate(images_batch, start=1):
        r = (idx-1) // cols
        c = (idx-1) % cols
        left = margin_x + c * (cell_w + Inches(0.2))
        top = margin_y + r * (cell_h + Inches(0.2))
        rect = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, cell_w, cell_h)
        rect.name = f"Rendering{idx}"
        add_picture_contain(slide, rect, img_bytes)

def create_setting_slide_fallback(prs: Presentation,
                                  group_name: str,
                                  render_bytes: Optional[bytes],
                                  floorplan_bytes: Optional[bytes],
                                  products_df: pd.DataFrame,
                                  mapping_df: pd.DataFrame,
                                  master_df: pd.DataFrame):
    slide = prs.slides.add_slide(get_blank_layout(prs))
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(8.0), Inches(0.6))
    title.name = "SETTINGNAME"
    set_text_preserve_format(title, group_name)

    render_anchor = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.5), Inches(1.2), Inches(5.5), Inches(3.5))
    render_anchor.name = "Rendering"
    if render_bytes:
        add_picture_contain(slide, render_anchor, render_bytes)

    line_anchor = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(6.2), Inches(1.2), Inches(3.0), Inches(3.5))
    line_anchor.name = "Linedrawing"
    if floorplan_bytes:
        add_picture_contain(slide, line_anchor, floorplan_bytes)

    start_top = Inches(5.0)
    cell_w = Inches(1.6)
    cell_h = Inches(1.2)
    gap = Inches(0.2)
    subset = products_df.head(12).copy() if len(products_df) > 12 else products_df.copy()
    for i, row in enumerate(subset.itertuples(index=False), start=1):
        r = (i-1) // 6
        c = (i-1) % 6
        left = Inches(0.5) + c * (cell_w + gap)
        top = start_top + r * (cell_h + Inches(0.6))
        pack_anchor = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, left, top, cell_w, cell_h)
        pack_anchor.name = f"ProductPackshot{i}"
        pack_url = find_packshot_url(row.ARTICLE_NO, mapping_df, master_df)
        img_bytes = http_get_bytes(pack_url) if pack_url else None
        if img_bytes:
            add_picture_contain(slide, pack_anchor, img_bytes)
        desc_box = slide.shapes.add_textbox(left, top + cell_h + Inches(0.05), cell_w, Inches(0.4))
        desc_box.name = f"PRODUCT DESCRIPTION {i}"
        set_text_preserve_format(desc_box, find_description(row.ARTICLE_NO, mapping_df))

def create_productlist_slide_fallback(prs: Presentation,
                                      group_name: str,
                                      products_df: pd.DataFrame,
                                      mapping_df: pd.DataFrame):
    slide = prs.slides.add_slide(get_blank_layout(prs))
    title = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.6))
    set_text_preserve_format(title, f"Products – {group_name}")
    anchor = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(0.5), Inches(1.2), Inches(9.0), Inches(5.0))
    anchor.name = "TableAnchor"
    rows = max(1, len(products_df)) + 1
    cols = 3
    table = add_table(slide, anchor, rows, cols)
    if table is None:
        return
    table.cell(0, 0).text = "Quantity"
    table.cell(0, 1).text = "Description"
    table.cell(0, 2).text = "Article No. / New Item No."
    r = 1
    for row in products_df.itertuples(index=False):
        table.cell(r, 0).text = str(int(row.Quantity))
        desc = find_description(row.ARTICLE_NO, mapping_df)
        table.cell(r, 1).text = desc
        new_item = find_new_item(row.ARTICLE_NO, mapping_df)
        table.cell(r, 2).text = f"{row.ARTICLE_NO} / {new_item}" if new_item else f"{row.ARTICLE_NO}"
        r += 1


def preflight_checks() -> Dict[str, str]:
    """Run minimal diagnostics. Returns dict of status messages."""
    results = {}
    # Template presence
    try:
        if not TEMPLATE_PATH.exists():
            results["template"] = "Template not found (input-template.pptx)."
        else:
            # Try opening template
            _ = Presentation(str(TEMPLATE_PATH))
            results["template"] = "OK"
    except Exception:
        results["template"] = "Template unreadable or not a valid .pptx."
    # Remote CSV reachability
    try:
        m = http_get_bytes(DEFAULT_MASTER_URL)
        results["master_csv"] = "OK" if m else "Unavailable"
    except Exception:
        results["master_csv"] = "Unavailable"
    try:
        mp = http_get_bytes(DEFAULT_MAPPING_URL)
        results["mapping_csv"] = "OK" if mp else "Unavailable"
    except Exception:
        results["mapping_csv"] = "Unavailable"
    return results

# -----------------------------
# Slide builders
# -----------------------------
def build_overview_slides(prs: Presentation, overview_layout, rendering_bytes_list: List[bytes]):
    for batch in chunk(rendering_bytes_list, MAX_OVERVIEW_IMAGES):
        slide = prs.slides.add_slide(overview_layout)
        shape_map = build_shape_map(slide)
        for idx, img_bytes in enumerate(batch, start=1):
            key = clean_name(f"Rendering{idx}")
            if key in shape_map and img_bytes:
                add_picture_contain(slide, shape_map[key], img_bytes)

def build_setting_slide(prs: Presentation,
                        setting_layout,
                        group_name: str,
                        render_bytes: Optional[bytes],
                        floorplan_bytes: Optional[bytes],
                        products_df: pd.DataFrame,
                        mapping_df: pd.DataFrame,
                        master_df: pd.DataFrame):
    slide = prs.slides.add_slide(setting_layout)
    shape_map = build_shape_map(slide)
    if clean_name("SETTINGNAME") in shape_map:
        set_text_preserve_format(shape_map[clean_name("SETTINGNAME")], group_name)
    if clean_name("Rendering") in shape_map and render_bytes:
        add_picture_contain(slide, shape_map[clean_name("Rendering")], render_bytes)
    if clean_name("Linedrawing") in shape_map and floorplan_bytes:
        add_picture_contain(slide, shape_map[clean_name("Linedrawing")], floorplan_bytes)
    subset = products_df.head(12).copy() if len(products_df) > 12 else products_df.copy()
    for i, row in enumerate(subset.itertuples(index=False), start=1):
        pack_url = find_packshot_url(row.ARTICLE_NO, mapping_df, master_df)
        img_bytes = http_get_bytes(pack_url) if pack_url else None
        pic_key = clean_name(f"ProductPackshot{i}")
        if pic_key in shape_map and img_bytes:
            add_picture_contain(slide, shape_map[pic_key], img_bytes)
        desc_key = clean_name(f"PRODUCT DESCRIPTION {i}")
        if desc_key in shape_map:
            desc = find_description(row.ARTICLE_NO, mapping_df)
            set_text_preserve_format(shape_map[desc_key], desc)

def build_productlist_slide(prs: Presentation,
                            layout,
                            group_name: str,
                            products_df: pd.DataFrame,
                            mapping_df: pd.DataFrame):
    slide = prs.slides.add_slide(layout)
    shape_map = build_shape_map(slide)
    title_shape = shape_map.get(clean_name("Title"), None)
    if title_shape is None:
        for s in slide.shapes:
            if hasattr(s, "text_frame") and s.text_frame:
                title_shape = s
                break
    if title_shape:
        set_text_preserve_format(title_shape, f"Products – {group_name}")
    anchor = shape_map.get(clean_name("TableAnchor"), None)
    if not anchor:
        class Dummy: pass
        anchor = Dummy()
        anchor.left = Inches(1.0)
        anchor.top = Inches(2.0)
        anchor.width = Inches(8.0)
        anchor.height = Inches(4.5)
    rows = max(1, len(products_df)) + 1
    cols = 3
    table = add_table(slide, anchor, rows, cols)
    if table is None:
        return
    table.cell(0, 0).text = "Quantity"
    table.cell(0, 1).text = "Description"
    table.cell(0, 2).text = "Article No. / New Item No."
    r = 1
    for row in products_df.itertuples(index=False):
        table.cell(r, 0).text = str(int(row.Quantity))
        desc = find_description(row.ARTICLE_NO, mapping_df)
        table.cell(r, 1).text = desc
        new_item = find_new_item(row.ARTICLE_NO, mapping_df)
        table.cell(r, 2).text = f"{row.ARTICLE_NO} / {new_item}" if new_item else f"{row.ARTICLE_NO}"
        r += 1

# -----------------------------
# App UI
# -----------------------------
st.set_page_config(page_title="Muuto PowerPoint Generator", layout="centered")
st.title("Muuto PowerPoint Generator")
st.write("Upload your group files (CSV and images). The app uses a fixed PowerPoint template from the repo and fetches Master Data and Mapping from fixed URLs.")

# Session state for uploads
if "uploads" not in st.session_state:
    st.session_state.uploads = []  # list of dict: {"name":..., "bytes":...}

# Upload widget
files = st.file_uploader(
    "User group files (.csv, .jpg, .png). You can add multiple files.",
    type=["csv", "jpg", "jpeg", "png"],
    accept_multiple_files=True,
)

if files:
    existing_names = {u["name"] for u in st.session_state.uploads}
    for f in files:
        if f.name not in existing_names:
            st.session_state.uploads.append({"name": f.name, "bytes": f.read()})
            existing_names.add(f.name)

# Single flat list without header or pagination
if st.session_state.uploads:
    to_remove = []
    for idx, f in enumerate(st.session_state.uploads):
        size_kb = f"{len(f['bytes'])/1024:.1f}KB"
        col1, col2, col3 = st.columns([6, 2, 1])
        with col1:
            st.caption(f["name"])
        with col2:
            st.caption(size_kb)
        with col3:
            if st.button("❌", key=f"rm_{idx}"):
                to_remove.append(idx)
    if to_remove:
        for i in sorted(to_remove, reverse=True):
            st.session_state.uploads.pop(i)
else:
    st.info("No user group files uploaded yet.")

# Generate button
generate = st.button("Generate presentation")

def build_groups(upload_list: List[Dict]) -> Dict[str, Dict]:
    groups: Dict[str, Dict] = {}
    for item in upload_list:
        name = item["name"]
        b = item["bytes"]
        key, t = group_key_from_filename(name)
        if key not in groups:
            groups[key] = {"name": key, "csv": None, "render": None, "floorplan": None}
        if t == "csv":
            groups[key]["csv"] = b
        elif t == "render":
            if groups[key]["render"] is None:
                groups[key]["render"] = b
        elif t in ["floorplan", "linedrawing"]:
            if groups[key]["floorplan"] is None:
                groups[key]["floorplan"] = b
    return groups

def collect_all_renderings(groups: Dict[str, Dict]) -> List[bytes]:
    lst = []
    for g in groups.values():
        if g.get("render"):
            lst.append(g["render"])
    return lst

def safe_present(prs: Presentation) -> bytes:
    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio.getvalue()


if generate:
    # Preflight diagnostics
    diag = preflight_checks()
    if diag.get("template") != "OK":
        st.error("Template issue: " + diag.get("template", "Unknown"))
    if diag.get("master_csv") != "OK":
        st.warning("Master Data source not reachable. Proceeding without packshots.")
    if diag.get("mapping_csv") != "OK":
        st.warning("Mapping source not reachable. Proceeding without descriptions and new item mapping.")

    if not TEMPLATE_PATH.exists():
        st.error("Template file is missing in the repository: input-template.pptx")
    elif diag.get("template") != "OK":
        st.error("Template unreadable. Ensure it is a valid .pptx and not a .ppt or zipped file.")
    elif not st.session_state.uploads:
        st.error("Please upload at least one group file.")
    else:
        try:
            with st.spinner("Work in progress…"):
                prs = ensure_presentation_from_path(TEMPLATE_PATH)
                overview_layout = find_layout_by_name(prs, "Overview")
                setting_layout = find_layout_by_name(prs, "Setting")
                productlist_layout = find_layout_by_name(prs, "ProductListBlank")
                groups = build_groups(st.session_state.uploads)

                master_df = load_remote_csv(DEFAULT_MASTER_URL)
                master_df = normalize_master(master_df)
                mapping_df = load_remote_csv(DEFAULT_MAPPING_URL)
                mapping_df = normalize_mapping(mapping_df)

                renders = collect_all_renderings(groups)
                if renders:
                    if overview_layout:
                        build_overview_slides(prs, overview_layout, renders)
                    else:
                        for batch in chunk(renders, MAX_OVERVIEW_IMAGES):
                            create_overview_slide_fallback(prs, batch)

                for key in sorted(groups.keys()):
                    g = groups[key]
                    group_name = g["name"]
                    pcon_df = normalize_pcon(parse_csv_flex(g["csv"]) if g["csv"] else pd.DataFrame())
                    if pcon_df.empty:
                        pcon_df = pd.DataFrame(columns=["ARTICLE_NO", "Quantity"])
                    if setting_layout:
                        build_setting_slide(prs, setting_layout, group_name, g.get("render"), g.get("floorplan"), pcon_df, mapping_df, master_df)
                    else:
                        create_setting_slide_fallback(prs, group_name, g.get("render"), g.get("floorplan"), pcon_df, mapping_df, master_df)
                    if productlist_layout:
                        build_productlist_slide(prs, productlist_layout, group_name, pcon_df, mapping_df)
                    else:
                        create_productlist_slide_fallback(prs, group_name, pcon_df, mapping_df)

                ppt_bytes = safe_present(prs)
                st.success("Your presentation is ready")
                st.download_button("Download Muuto_Settings.pptx", data=ppt_bytes, file_name=OUTPUT_NAME, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
        except Exception as e:
            # Show clearer hints without internal paths
            hint = "Template unreadable" if "PackageNotFoundError" in str(type(e)) else "Unexpected generation error"
            st.error("Generation failed. " + hint + ". Check that the template is a valid .pptx and inputs are well-formed CSV/images.")
