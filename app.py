import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import io
import re
from typing import List, Dict, Any, Tuple
from collections import defaultdict
import requests
import os

# -----------------------------
# KONSTANTER OG PLACEHOLDERS
# -----------------------------

# Skabelonfilen skal ligge i samme mappe som app.py
TEMPLATE_FILENAME = "input-template.pptx"

# Google Sheets skal være publiceret som CSV
LIBRARY_DEFAULT_URL = "https://docs.google.com/spreadsheets/d/1h3yaq094mBa5Yadfi9Nb_Wrnzj3gIH2DviIfU0DdwsQ/export?format=csv&gid=437866492"
MASTER_DEFAULT_URL  = "https://docs.google.com/spreadsheets/d/1blj42SbFpszWGyOrDOUwyPDJr9K1NGpTMX6eZTbt_P4/export?format=csv&gid=194572316"

# pCon-kolonneindeks (samme som i din pCon-kode)
PCON_ARTICLE_NO_COL   = 17
PCON_QUANTITY_COL     = 30
PCON_SHORT_TEXT_COL   = 2
PCON_VARIANT_TEXT_COL = 4
PCON_SKIPROWS         = 2

# Skabelonens placeholders
PRODUCT_PLACEHOLDERS     = [f"{{{{PRODUCT DESCRIPTION {i}}}}}" for i in range(1, 13)]
PACKSHOT_PLACEHOLDERS    = [f"{{{{ProductPackshot{i}}}}}" for i in range(1, 13)]
ACCESSORY_PLACEHOLDERS   = [f"{{{{accessory{i}}}}}" for i in range(1, 7)]
OVERVIEW_RENDER_TAGS     = [f"{{{{Rendering{i}}}}}" for i in range(1, 13)]
SETTING_TEXT_TAGS        = ["{{SETTINGNAME}}", "{{SETTINGSUBHEADLINE}}", "{{SettingDimensions}}", "{{SettingSize}}", "{{ProductsinSettingList}}"]
RENDERING_TAG            = "{{Rendering}}"
LINEDRAWING_TAG          = "{{Linedrawing}}"

# Masterdata-kolonne, der indeholder packshot URL
MASTER_DATA_PACKSHOT_COL = "IMAGE URL"  # vigtigt: navnet skal eksistere i master-CSV/XLSX


# -----------------------------
# DATAINDLÆSNING OG MATCH-LOGIK
# -----------------------------

@st.cache_data
def load_pcon_file(uploaded_file) -> pd.DataFrame:
    """Læs pCon CSV/XLSX og returnér DataFrame med de nødvendige felter."""
    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(uploaded_file, skiprows=PCON_SKIPROWS)
    else:
        df = pd.read_excel(uploaded_file, skiprows=PCON_SKIPROWS, engine="openpyxl")

    required_indices = [PCON_SHORT_TEXT_COL, PCON_VARIANT_TEXT_COL, PCON_ARTICLE_NO_COL, PCON_QUANTITY_COL]
    if df.shape[1] <= max(required_indices):
        raise ValueError("Utilstrækkeligt antal kolonner i pCon-filen.")

    sub = df.iloc[:, required_indices].copy()
    sub.columns = ["SHORT_TEXT", "VARIANT_TEXT", "ARTICLE_NO", "QUANTITY"]

    sub["ARTICLE_NO"] = sub["ARTICLE_NO"].astype(str).str.strip()
    sub["SHORT_TEXT"] = sub["SHORT_TEXT"].astype(str).str.strip()
    sub["VARIANT_TEXT"] = sub["VARIANT_TEXT"].astype(str).str.strip()
    return sub


@st.cache_data
def load_lookup_csv(url: str, required_cols: List[str], source_name: str) -> pd.DataFrame:
    """Læs CSV fra URL og valider kolonner."""
    if not url.startswith("http"):
        raise ValueError(f"{source_name}: ugyldig URL")
    df = pd.read_csv(url)
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"{source_name}: mangler kolonner: {', '.join(missing)}")
    for col in ["EUR ITEM NO.", "ITEM NO."]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    return df


def fallback_key(article_no: str) -> str:
    """Base-nøgle til fallback-match: fjern 'SPECIAL-' og alt efter første '-'."""
    if not article_no:
        return ""
    base = re.sub(r"^SPECIAL-", "", article_no, flags=re.IGNORECASE)
    return base.split("-")[0].strip()


def match_library(row: pd.Series, library_df: pd.DataFrame) -> Dict[str, Any]:
    """Match mod Library_data på EUR ITEM NO. med fallback."""
    article_no = row["ARTICLE_NO"]
    hit = library_df[library_df["EUR ITEM NO."] == article_no]
    if not hit.empty:
        # Ignorer 'ALL COLORS' som i din logik
        if "PRODUCT" in hit.columns and "ALL COLORS" in str(hit["PRODUCT"].iloc[0]).upper():
            pass
        else:
            return hit.iloc[0].to_dict()

    key = fallback_key(article_no)
    if key:
        alt = library_df[library_df["EUR ITEM NO."].apply(fallback_key) == key]
        if not alt.empty:
            return alt.iloc[0].to_dict()
    return {}


def match_master(row: pd.Series, master_df: pd.DataFrame) -> Dict[str, Any]:
    """Match mod Masterdata på ITEM NO. med fallback."""
    article_no = row["ARTICLE_NO"]
    hit = master_df[master_df["ITEM NO."] == article_no]
    if not hit.empty:
        return hit.iloc[0].to_dict()

    key = fallback_key(article_no)
    if key:
        alt = master_df[master_df["ITEM NO."].apply(fallback_key) == key]
        if not alt.empty:
            return alt.iloc[0].to_dict()
    return {}


def line_text_from_pcon_and_library(row: pd.Series, library_match: Dict[str, Any]) -> str:
    """Byg visningstekst for et produkt."""
    if library_match and "PRODUCT" in library_match and str(library_match["PRODUCT"]).strip():
        return str(library_match["PRODUCT"]).strip()
    short_text = str(row["SHORT_TEXT"]).strip()
    variant_text = str(row["VARIANT_TEXT"]).strip()
    if not variant_text or variant_text.upper() == "LIGHT OPTION: OFF":
        return short_text
    return f"{short_text} – {variant_text}"


def build_products(pcon_df: pd.DataFrame, library_df: pd.DataFrame, master_df: pd.DataFrame) -> Tuple[str, List[Dict[str, Any]], List[str]]:
    """
    Returnér:
      - tekstblok til {{ProductsinSettingList}} (sorteret alfabetisk, case-insensitive)
      - product_details: liste med dicts der inkluderer description, article_no, packshot_url
      - warnings: liste af str
    Linjeformat i tekstblokken skal inkludere itemnummer:  "<QTY> X <Tekst> – <ARTICLE_NO>"
    """
    warnings = []
    if MASTER_DATA_PACKSHOT_COL not in master_df.columns:
        warnings.append(f"Kolonnen '{MASTER_DATA_PACKSHOT_COL}' mangler i Master Data. Packshots bliver tomme.")

    product_lines = []
    details = []

    for _, row in pcon_df.iterrows():
        # Quantities: default 1 hvis tomt
        qty = int(row["QUANTITY"]) if pd.notna(row["QUANTITY"]) and str(row["QUANTITY"]).strip() else 1

        lib = match_library(row, library_df)
        mst = match_master(row, master_df)
        desc = line_text_from_pcon_and_library(row, lib)

        line = f"{qty} X {desc} – {row['ARTICLE_NO']}"
        product_lines.append(line)

        packshot_url = mst.get(MASTER_DATA_PACKSHOT_COL, "") if MASTER_DATA_PACKSHOT_COL in master_df.columns else ""
        if not packshot_url and mst:
            warnings.append(f"Mangler packshot URL i Master Data for {row['ARTICLE_NO']}")

        details.append({
            "description": desc,
            "article_no": str(row["ARTICLE_NO"]),
            "qty": qty,
            "packshot_url": packshot_url
        })

        if not lib:
            warnings.append(f"Ingen Library-match for artikel {row['ARTICLE_NO']}. Bruger pCon-tekst.")

    def sort_key(s: str) -> str:
        # sortér på tekstdelen efter " X "
        return s.split(" X ", 1)[-1].lower()

    product_lines_sorted = sorted(product_lines, key=sort_key)
    return "\n".join(product_lines_sorted), details, warnings


# -----------------------------
# BILLEDER
# -----------------------------

@st.cache_data(ttl=3600)
def fetch_image(url: str) -> bytes | None:
    """Hent billede og returnér bytes. Returnér None hvis ikke muligt."""
    if not url or not str(url).startswith("http"):
        return None
    try:
        r = requests.get(url, timeout=15)
        r.raise_for_status()
        if "image" not in r.headers.get("Content-Type", "").lower():
            return None
        return r.content
    except requests.RequestException:
        return None


def preprocess_image(img_bytes: bytes, max_side: int = 1200, quality: int = 85) -> bytes:
    """Konverter, skaler, komprimer til JPEG."""
    try:
        img = Image.open(io.BytesIO(img_bytes))
        if img.mode != "RGB":
            img = img.convert("RGB")
        if img.width > max_side or img.height > max_side:
            ratio = min(max_side / img.width, max_side / img.height)
            img = img.resize((int(img.width * ratio), int(img.height * ratio)), Image.Resampling.LANCZOS)
        out = io.BytesIO()
        img.save(out, format="JPEG", quality=quality)
        return out.getvalue()
    except Exception:
        return img_bytes


# -----------------------------
# POWERPOINT HJÆLPERE
# -----------------------------

def find_shape_by_tag(slide, tag: str):
    """Find placeholder ved at matche præcis tag-tekst eller shape-navn."""
    tag_u = tag.strip().upper()
    # 1) placeholders
    for shp in getattr(slide, "placeholders", []):
        if getattr(shp, "has_text_frame", False):
            if shp.text_frame and shp.text_frame.text.strip().upper() == tag_u:
                return shp
        if tag_u in shp.name.upper():
            return shp
    # 2) shapes
    for shp in slide.shapes:
        if getattr(shp, "has_text_frame", False) and shp.text_frame:
            if shp.text_frame.text.strip().upper() == tag_u:
                return shp
        if tag_u in shp.name.upper():
            return shp
    return None


def fit_replace_text(shape, value: str):
    """Erstat tekst og bevar template-font, størrelse, bold og case."""
    if shape is None or not getattr(shape, "has_text_frame", False):
        return
    v = "" if value is None else str(value)

    tf = shape.text_frame
    if not tf.paragraphs:
        p = tf.add_paragraph()
    else:
        p = tf.paragraphs[0]
    if not p.runs:
        p.add_run()

    # Hent format fra første run
    tmpl_run = p.runs[0]
    font_name = tmpl_run.font.name
    font_size = tmpl_run.font.size
    font_bold = tmpl_run.font.bold

    # Ryd eksisterende indhold uden at ændre paragraph-objektet
    for para in list(tf.paragraphs):
        for run in list(para.runs):
            run.text = ""
    # Skriv ny tekst i første paragraph
    para = tf.paragraphs[0]
    para.clear()
    run = para.add_run()
    run.text = v

    # Genskab font-egenskaber
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = font_bold
    tf.word_wrap = True


def replace_image(slide, tag: str, image_bytes: bytes, crop_to_frame: bool = False):
    """Indsæt billede i shape-ramme, skaler proportionalt, centrer, evt. crop."""
    ph = find_shape_by_tag(slide, tag)
    if ph is None:
        return
    left, top, width, height = ph.left, ph.top, ph.width, ph.height

    try:
        # Fjern placeholder-shape
        try:
            sp = ph.element
            sp.getparent().remove(sp)
        except Exception:
            pass

        # Tilføj billed-shape i samme ramme
        stream = io.BytesIO(image_bytes)
        pic = slide.shapes.add_picture(stream, left, top, width=width, height=height)

        # Forbedret skalering/centrering
        img = Image.open(io.BytesIO(image_bytes))
        img_w, img_h = img.size
        frame_w, frame_h = width.emu, height.emu

        w_ratio = frame_w / img_w
        h_ratio = frame_h / img_h
        scale = max(w_ratio, h_ratio) if crop_to_frame else min(w_ratio, h_ratio)

        new_w = int(img_w * scale)
        new_h = int(img_h * scale)

        pic.width = new_w
        pic.height = new_h
        pic.left = left + (width - pic.width) // 2
        pic.top  = top  + (height - pic.height) // 2

        if crop_to_frame:
            # Crop til præcis ramme
            off_x = (new_w - frame_w) / (2 * new_w)
            off_y = (new_h - frame_h) / (2 * new_h)
            pic.crop_left = max(0, off_x)
            pic.crop_right = max(0, off_x)
            pic.crop_top = max(0, off_y)
            pic.crop_bottom = max(0, off_y)
            pic.left, pic.top, pic.width, pic.height = left, top, width, height

    except Exception as e:
        st.warning(f"Kunne ikke indsætte billede for '{tag}': {e}")


def find_first_slide_with_tag(prs: Presentation, tag: str) -> Tuple[Any, int]:
    """Find første slide der indeholder en shape med den eksakte tag-tekst."""
    tag_u = tag.strip().upper()
    for i, slide in enumerate(prs.slides):
        for shp in slide.shapes:
            if getattr(shp, "has_text_frame", False) and shp.text_frame and shp.text_frame.text.strip().upper() == tag_u:
                return slide, i
    raise ValueError(f"Skabelonen mangler en slide med placeholder-tekst: {tag}")


def fill_overview_slides(prs: Presentation, renderings: List[bytes]) -> int:
    """Opret OVERVIEW-slides og fyld {{Rendering1..12}} i rækkefølge."""
    if not renderings:
        return 0
    overview_slide, overview_idx = find_first_slide_with_tag(prs, "OVERVIEW")
    layout = overview_slide.slide_layout

    n = len(renderings)
    groups = (n + 11) // 12
    created = 0

    for i in range(groups):
        s = prs.slides.add_slide(layout)
        created += 1
        start, end = i * 12, min((i + 1) * 12, n)
        for j, rb in enumerate(renderings[start:end]):
            replace_image(s, OVERVIEW_RENDER_TAGS[j], preprocess_image(rb), crop_to_frame=True)
        if groups > 1:
            title = find_shape_by_tag(s, "OVERVIEW")
            if title:
                fit_replace_text(title, f"OVERVIEW (SIDE {i+1} AF {groups})")

    # Fjern template-oversigtsslide
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[overview_idx])
    return created


def fill_setting_slides(
    prs: Presentation,
    settings: List[Dict[str, Any]],
    library_df: pd.DataFrame,
    master_df: pd.DataFrame
) -> int:
    """Opret slides pr. setting. Brug pCon-logik til produktliste, og EY-logik til packshots fra master."""
    if not settings:
        return 0

    template_slide, template_idx = find_first_slide_with_tag(prs, "{{SETTINGNAME}}")
    layout = template_slide.slide_layout
    total = 0

    for setting in settings:
        setting_name = setting["name"]
        pcon_df = load_pcon_file(setting["pcon_file"])
        product_text, product_details, warnings = build_products(pcon_df, library_df, master_df)

        for w in warnings:
            st.warning(f"[{setting_name}] {w}")

        # Produkt-paginering i grupper af 12
        n = len(product_details)
        pages = (n + 11) // 12 or 1

        for page in range(pages):
            s = prs.slides.add_slide(layout)
            total += 1

            # Tekstfelter
            fit_replace_text(find_shape_by_tag(s, "{{SETTINGNAME}}"),
                             f"{setting_name}" if pages == 1 else f"{setting_name} (Produkter: Side {page+1} af {pages})")

            if page == 0:
                fit_replace_text(find_shape_by_tag(s, "{{SETTINGSUBHEADLINE}}"), setting.get("subheadline", ""))
                fit_replace_text(find_shape_by_tag(s, "{{SettingDimensions}}"), setting.get("dimensions", ""))
                fit_replace_text(find_shape_by_tag(s, "{{SettingSize}}"), setting.get("size", ""))
                fit_replace_text(find_shape_by_tag(s, "{{ProductsinSettingList}}"), product_text)

                # Billeder: Rendering og Linedrawing
                if setting.get("rendering_bytes"):
                    replace_image(s, RENDERING_TAG, preprocess_image(setting["rendering_bytes"]))
                if setting.get("linedrawing_bytes"):
                    replace_image(s, LINEDRAWING_TAG, preprocess_image(setting["linedrawing_bytes"]))
                # Nulstil accessory-felter eksplicit
                for tag in ACCESSORY_PLACEHOLDERS:
                    fit_replace_text(find_shape_by_tag(s, f"{{{{{tag}}}}}"), "")
            else:
                # Kun navn og produkter på fortsættelsessider
                fit_replace_text(find_shape_by_tag(s, "{{SETTINGSUBHEADLINE}}"), "")
                fit_replace_text(find_shape_by_tag(s, "{{SettingDimensions}}"), "")
                fit_replace_text(find_shape_by_tag(s, "{{SettingSize}}"), "")
                fit_replace_text(find_shape_by_tag(s, "{{ProductsinSettingList}}"), "")

            # Udfyld op til 12 produkter og packshots pr. side
            start, end = page * 12, min((page + 1) * 12, n)
            chunk = product_details[start:end]
            for idx, prod in enumerate(chunk):
                # Produktbeskrivelse i tekstplaceholder
                prod_tag = PRODUCT_PLACEHOLDERS[idx]
                fit_replace_text(find_shape_by_tag(s, prod_tag), prod["description"])

                # Packshot fra masterdata
                ps_tag = PACKSHOT_PLACEHOLDERS[idx]
                if prod.get("packshot_url"):
                    raw = fetch_image(prod["packshot_url"])
                    if raw:
                        replace_image(s, f"{{{{{ps_tag}}}}}", preprocess_image(raw))

    # Fjern skabelonens setting-slide
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[template_idx])
    return total


# -----------------------------
# GRUPPERING AF UPLOADS I SETTINGS
# -----------------------------

def group_uploaded_files(uploaded_files: List[io.BytesIO]) -> Dict[str, Dict[str, Any]]:
    """
    Gruppér filer pr. setting baseret på fælles prefix før første '_' eller '-'.
    For hver setting kræves: CSV/XLSX + Rendering (jpg/png).
    Linedrawing er valgfri (filnavn med 'floorplan' i navnet).
    """
    groups = defaultdict(lambda: {"csv": None, "rendering": None, "floorplan": None, "name": None})

    for f in uploaded_files:
        fname = f.name
        base = re.split(r"[_-]", fname, 1)[0].strip()
        if not base:
            continue
        std_name = base.replace("_", " ").strip().title()
        groups[base]["name"] = std_name

        lf = fname.lower()
        if lf.endswith(".csv") or lf.endswith(".xlsx"):
            groups[base]["csv"] = f
        elif "floorplan" in lf:
            groups[base]["floorplan"] = f
        elif lf.endswith(".jpg") or lf.endswith(".jpeg") or lf.endswith(".png"):
            groups[base]["rendering"] = f

    # filtrér kun gyldige settings
    out = {}
    for base, data in groups.items():
        if data["csv"] and data["rendering"]:
            out[base] = data
        else:
            st.warning(f"Ignorerer '{data.get('name', base)}': mangler CSV/XLSX eller rendering.")
    return out


# -----------------------------
# STREAMLIT UI
# -----------------------------

def main():
    st.set_page_config(page_title="Muuto PowerPoint Generator", layout="wide")
    st.title("Muuto PowerPoint Generator")
    st.caption("Kombinerer EY-præsentationslogik (packshots fra Master Data) og pCon-logik (produktliste fra CSV) i én samlet PPT.")

    st.subheader("1) Upload alle setting-filer på én gang")
    st.write("CSV/XLSX fra pCon, rendering JPG/PNG, valgfri floorplan/linedrawing JPG/PNG. Filer med samme prefix grupperes.")
    uploads = st.file_uploader("Filer", type=["csv", "xlsx", "jpg", "jpeg", "png"], accept_multiple_files=True)

    st.subheader("2) Manuelle felter (gælder for alle settings)")
    c1, c2, c3 = st.columns(3)
    with c1:
        subhead = st.text_input("SETTINGSUBHEADLINE", value="")
    with c2:
        dims = st.text_input("SettingDimensions", value="")
    with c3:
        size = st.text_input("SettingSize", value="")

    st.subheader("3) Generér PowerPoint")
    if st.button("Generér"):
        errors = []
        if not os.path.exists(TEMPLATE_FILENAME):
            errors.append(f"Skabelonen '{TEMPLATE_FILENAME}' blev ikke fundet i mappen.")
        if not uploads:
            errors.append("Ingen filer uploadet.")
        if errors:
            for e in errors:
                st.error(e)
            st.stop()

        grouped = group_uploaded_files(uploads)
        if not grouped:
            st.error("Ingen gyldige settings. Tjek filnavne og upload mindst CSV + rendering pr. setting.")
            st.stop()

        settings = []
        all_renderings = []
        for _, data in grouped.items():
            rendering_bytes = data["rendering"].read()
            all_renderings.append(rendering_bytes)
            linedrawing_bytes = data["floorplan"].read() if data["floorplan"] else None
            settings.append({
                "name": data["name"],
                "subheadline": subhead,
                "dimensions": dims,
                "size": size,
                "pcon_file": data["csv"],
                "rendering_bytes": rendering_bytes,
                "linedrawing_bytes": linedrawing_bytes
            })

        with st.spinner("Indlæser opslagsdata og skabelon…"):
            library_df = load_lookup_csv(LIBRARY_DEFAULT_URL, ["PRODUCT", "EUR ITEM NO."], "Library_data")
            master_df  = load_lookup_csv(MASTER_DEFAULT_URL,  ["ITEM NO.", MASTER_DATA_PACKSHOT_COL], "Master_data")
            prs = Presentation(TEMPLATE_FILENAME)

        with st.spinner("Bygger OVERVIEW og settings, henter packshots…"):
            fill_overview_slides(prs, all_renderings)
            fill_setting_slides(prs, settings, library_df, master_df)

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        st.success("PowerPoint klar.")
        st.download_button(
            "Download præsentation",
            data=buf,
            file_name="Muuto_Setting_Presentation_Auto.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )


if __name__ == "__main__":
    main()
