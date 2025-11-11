import io
import re
import time
from typing import Dict, List, Optional, Tuple
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.util import Inches

# ---------------------- Constants ----------------------
TEMPLATE_PATH = Path("input-template.pptx")
DEFAULT_MASTER_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRdNwE1Q_aG3BntCZZPRIOgXEFJ5AHJxHmRgirMx2FJqfttgCZ8on-j1vzxM-muTTvtAHwc-ovDV1qF/pub?output=csv"
DEFAULT_MAPPING_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQPRmVmc0LYISduQdJyfz-X3LJlxiEDCNwW53LhFsWp5fFDS8V669rCd9VGoygBZSAZXeSNZ5fquPen/pub?output=csv"
OUTPUT_NAME = "Muuto_Settings.pptx"

MAX_OVERVIEW_IMAGES = 12
HTTP_TIMEOUT = 10
HTTP_RETRIES = 1
MAX_IMAGE_PX = 1400
JPEG_QUALITY = 85

# ---------------------- Utils ----------------------
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

def build_shape_map(slide) -> Dict[str, list]:
    mapping: Dict[str, List] = {}
    for shape in slide.shapes:
        try:
            nm = clean_name(getattr(shape, "name", ""))
            if nm:
                mapping.setdefault(nm, []).append(shape)
        except Exception:
            continue
    return mapping

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
        {"sep": "|", "encoding": "utf-8"},
        {"sep": ",", "encoding": "utf-8-sig"},
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
    lname = base.lower()
    if "floorplan" in lname:
        t = "floorplan"
    elif "linedrawing" in lname or "line_drawing" in lname or "line drawing" in lname:
        t = "linedrawing"
    else:
        ext = Path(name).suffix.lower()
        if ext == ".csv":
            t = "csv"
        elif ext in [".jpg", ".jpeg", ".png"]:
            t = "render"
        else:
            t = "other"
    if " - " in base:
        key = base.split(" - ", 1)[1]
    else:
        parts = re.split(r"[-_]", base)
        key = parts[-1] if parts else base
    key = re.sub(r"\s+(floorplan|line\s*drawing|linedrawing)$", "", key, flags=re.IGNORECASE).strip()
    return key, t

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
        for c in df.co
