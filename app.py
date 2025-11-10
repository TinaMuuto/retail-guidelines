# app.py

import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import io
import re
from typing import List, Dict, Any, Tuple
from collections import defaultdict
import numpy as np
import requests # Nødvendigt for at hente billeder via URL

# --- KONSTANTER ---
PCON_ARTICLE_NO_COL = 17
PCON_QUANTITY_COL = 30
PCON_SHORT_TEXT_COL = 2
PCON_VARIANT_TEXT_COL = 4
PCON_SKIPROWS = 2
PRODUCT_PLACEHOLDERS = [f"PRODUCT DESCRIPTION {i}" for i in range(1, 13)]
PACKSHOT_PLACEHOLDERS = [f"ProductPackshot{i}" for i in range(1, 13)]
ACCESSORY_PLACEHOLDERS = [f"accessory{i}" for i in range(1, 7)]
OVERVIEW_RENDER_PLACEHOLDERS = [f"Rendering{i}" for i in range(1, 13)]
# Antages at denne kolonne eksisterer i Master Data Sheets
MASTER_DATA_PACKSHOT_COL = 'IMAGE URL' 

# --- DATA OG FILHÅNDTERINGSFUNKTIONER ---

@st.cache_data
def load_pcon_file(uploaded_file) -> pd.DataFrame:
    """Indlæser pCon Excel/CSV og returnerer en DataFrame med de nødvendige kolonner."""
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file, skiprows=PCON_SKIPROWS)
    else:
        df = pd.read_excel(uploaded_file, skiprows=PCON_SKIPROWS, engine='openpyxl')
    
    required_indices = [PCON_SHORT_TEXT_COL, PCON_VARIANT_TEXT_COL, PCON_ARTICLE_NO_COL, PCON_QUANTITY_COL]
    max_index = max(required_indices)
    
    if df.shape[1] <= max_index:
        raise ValueError("Utilstrækkeligt antal kolonner i pCon-filen.")

    df_subset = df.iloc[:, required_indices].copy()
    df_subset.columns = ['SHORT_TEXT', 'VARIANT_TEXT', 'ARTICLE_NO', 'QUANTITY']
    
    df_subset['ARTICLE_NO'] = df_subset['ARTICLE_NO'].astype(str).str.strip().fillna('')
    df_subset['SHORT_TEXT'] = df_subset['SHORT_TEXT'].astype(str).str.strip().fillna('')
    df_subset['VARIANT_TEXT'] = df_subset['VARIANT_TEXT'].astype(str).str.strip().fillna('')
    
    return df_subset

@st.cache_data
def load_library_data(input_url: str, expected_cols: List[str], source_name: str) -> pd.DataFrame:
    """Indlæser opslagsdata fra Google Sheets CSV-eksport URL og validerer nødvendige kolonner."""
    
    if not input_url.startswith('http'):
        st.error(f"Fejl: '{source_name}' mangler en gyldig URL.")
        raise ValueError("URL mangler.")

    try:
        # Læser direkte fra CSV-eksport URL'en
        df = pd.read_csv(input_url) 
        
        missing_cols = [col for col in expected_cols if col not in df.columns]
        if missing_cols:
            st.error(f"Fejl: Dataen fra '{source_name}' (via URL) mangler de forventede kolonner: {', '.join(missing_cols)}")
            raise ValueError("Manglende kolonner i opslagsdata.")
            
        for col in ['EUR ITEM NO.', 'ITEM NO.']:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip() 
        return df
        
    except Exception as e:
        st.error(f"Fejl under indlæsning af '{source_name}' fra URL: {e}. Tjek URL'en og om dokumentet er offentliggjort som CSV-eksport.")
        raise

@st.cache_data(ttl=3600) # Cache i 1 time for at undgå gentagne netværkskald
def fetch_image_bytes_from_url(url: str, article_no: str) -> bytes | None:
    """Henter billed-bytes fra en given URL."""
    if not url or not url.startswith('http'):
        return None
    
    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status() # Hæv exception for dårlige statuskoder (4xx eller 5xx)
        
        content_type = response.headers.get('Content-Type', '').lower()
        if 'image' not in content_type:
            return None

        return response.content
    except requests.exceptions.RequestException as e:
        st.warning(f"Advarsel: Kunne ikke hente packshot for {article_no} fra URL. Fejl: {e}")
        return None
        
def fallback_key(article_no: str) -> str:
    """Genererer en fallback-nøgle fra ARTICLE_NO."""
    if pd.isna(article_no) or not article_no:
        return ""
    
    base_key = re.sub(r'^SPECIAL-', '', article_no, flags=re.IGNORECASE)
    base_key = base_key.split('-')[0].strip()
    
    return base_key

def match_library(row: pd.Series, library_df: pd.DataFrame) -> Dict[str, Any]:
    """Matcher en pCon-række mod Library_data."""
    article_no = row['ARTICLE_NO']
    primary_match = library_df[library_df['EUR ITEM NO.'] == article_no]
    
    if not primary_match.empty:
        if 'PRODUCT' in primary_match.columns and 'ALL COLORS' in primary_match['PRODUCT'].iloc[0]:
            pass 
        else:
            return primary_match.iloc[0].to_dict()

    key = fallback_key(article_no)
    if key:
        fallback_match = library_df[library_df['EUR ITEM NO.'].apply(fallback_key) == key]
        if not fallback_match.empty:
            return fallback_match.iloc[0].to_dict()
            
    return {}

def match_master(row: pd.Series, master_df: pd.DataFrame) -> Dict[str, Any]:
    """Matcher en pCon-række mod Masterdata."""
    article_no = row['ARTICLE_NO']
    primary_match = master_df[master_df['ITEM NO.'] == article_no]
    
    if not primary_match.empty:
        return primary_match.iloc[0].to_dict()

    key = fallback_key(article_no)
    if key:
        fallback_match = master_df[master_df['ITEM NO.'].apply(fallback_key) == key]
        if not fallback_match.empty:
            return fallback_match.iloc[0].to_dict()
            
    return {}

def get_product_description(row: pd.Series, library_match: Dict[str, Any]) -> str:
    """Genererer produktbeskrivelsen baseret på matches og pCon-data."""
    if library_match and 'PRODUCT' in library_match:
        return str(library_match['PRODUCT'])
    else:
        short_text = str(row['SHORT_TEXT']).strip()
        variant_text = str(row['VARIANT_TEXT']).strip()
        
        if not variant_text or variant_text.upper() == 'LIGHT OPTION: OFF':
            return short_text
        else:
            return f"{short_text} – {variant_text}"

def build_products_list(pcon_df: pd.DataFrame, library_df: pd.DataFrame, master_df: pd.DataFrame) -> Tuple[str, List[Dict[str, Any]], List[str]]:
    """Genererer den sorterede produktliste og forbereder produktdata, inkl. packshot URL'er."""
    
    product_lines = []
    product_details = []
    warnings = []
    
    if MASTER_DATA_PACKSHOT_COL not in master_df.columns:
        warnings.append(f"ADVARSEL: Kolonnen '{MASTER_DATA_PACKSHOT_COL}' blev ikke fundet i Master Data. Packshots vil blive tomme.")
        
    for _, row in pcon_df.iterrows():
        qty = int(row['QUANTITY']) if pd.notna(row['QUANTITY']) and row['QUANTITY'] else 1
        library_match = match_library(row, library_df)
        master_match = match_master(row, master_df)
        
        product_desc = get_product_description(row, library_match)
        
        list_line = f"{qty} X {product_desc}"
        product_lines.append(list_line)
        
        packshot_url = master_match.get(MASTER_DATA_PACKSHOT_COL, '') if MASTER_DATA_PACKSHOT_COL in master_df.columns else ''
        
        if not packshot_url and master_match:
             warnings.append(f"Advarsel: Packshot URL'en for {row['ARTICLE_NO']} er tom i Master Data.")

        detail = {
            'description': product_desc,
            'article_no': row['ARTICLE_NO'],
            'library_match': library_match,
            'packshot_url': packshot_url, 
            'pcon_row': row
        }
        product_details.append(detail)
        
        if not library_match:
            warnings.append(f"Advarsel: Ingen Library-match for artikel: {row['ARTICLE_NO']}. Bruger pCon-tekst.")

    def sort_key(line):
        return line.split(' X ', 1)[-1].lower() if ' X ' in line else line.lower()
        
    sorted_lines = sorted(product_lines, key=sort_key)
    
    return '\n'.join(sorted_lines), product_details, warnings

def preprocess_image(img_bytes: bytes) -> bytes:
    """Behandler et billede: konverterer til RGB, sætter max størrelse, komprimerer som JPEG."""
    try:
        img = Image.open(io.BytesIO(img_bytes))
        if img.mode != 'RGB':
            img = img.convert('RGB')
            
        max_size = 1200
        if img.width > max_size or img.height > max_size:
            ratio = min(max_size / img.width, max_size / img.height)
            new_size = (int(img.width * ratio), int(img.height * ratio))
            img = img.resize(new_size, Image.Resampling.LANCZOS)
            
        output = io.BytesIO()
        img.save(output, format='JPEG', quality=85) 
        return output.getvalue()
        
    except Exception as e:
        st.error(f"Fejl i billedbehandling: {e}")
        return img_bytes

def group_uploaded_files(uploaded_files: List[io.BytesIO]) -> Dict[str, Dict[str, Any]]:
    """Grupperer filer i settings baseret på det fælles filnavn prefix."""
    settings_grouped = defaultdict(lambda: {'csv': None, 'rendering': None, 'floorplan': None, 'name': None})
    
    for file in uploaded_files:
        filename = file.name
        
        base_name = re.split(r'[_-]', filename, 1)[0].strip() 
        
        if not base_name:
            continue
            
        setting_name_standardized = base_name.replace('_', ' ').strip().title()
        
        settings_grouped[base_name]['name'] = setting_name_standardized

        filename_lower = filename.lower()
        
        if filename_lower.endswith('.csv') or filename_lower.endswith('.xlsx'):
            settings_grouped[base_name]['csv'] = file
        elif 'floorplan' in filename_lower:
            settings_grouped[base_name]['floorplan'] = file
        elif filename_lower.endswith('.jpg') or filename_lower.endswith('.jpeg') or filename_lower.endswith('.png'):
            settings_grouped[base_name]['rendering'] = file
            
    final_settings = {}
    for base_name, data in settings_grouped.items():
        if data['csv'] and data['rendering']:
            final_settings[base_name] = data
        else:
            st.warning(f"⚠️ Setting '{data['name']}' blev ignoreret: Mangler CSV ({'OK' if data['csv'] else 'Mangler'}) eller Rendering ({'OK' if data['rendering'] else 'Mangler'}).")

    return final_settings


# --- POWERPOINT GENERERING FUNKTIONER ---

def get_placeholder_by_tag(slide, tag: str):
    """Finder den første placeholder på en slide, der matcher en tag."""
    for shape in slide.placeholders:
        if shape.has_text_frame and shape.text_frame.text.strip().upper() == tag.upper():
            return shape
        if tag.upper() in shape.name.upper():
            return shape
    for shape in slide.shapes:
        if shape.has_text_frame and shape.text_frame.text.strip().upper() == tag.upper():
            return shape
        if tag.upper() in shape.name.upper():
            return shape
    return None

def fit_replace_text(shape, value: str):
    """Erstatter tekst i en shape og bevarer det originale formatering."""
    
    value_str = str(value).strip() if value is not None and pd.notna(value) else ""
    
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame
    
    if not text_frame.paragraphs:
        p = text_frame.add_paragraph()
    else:
        p = text_frame.paragraphs[0]
        
    if not p.runs:
        p.add_run()
        
    template_run = p.runs[0]
    
    font_name = template_run.font.name
    font_size = template_run.font.size
    font_bold = template_run.font.bold
    
    while len(text_frame.paragraphs) > 0:
        p_to_remove = text_frame.paragraphs[0]
        for run in p_to_remove.runs:
            run.text = ""
        
    p = text_frame.paragraphs[0]
    p.clear() 
    run = p.add_run()
    run.text = value_str
    
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = font_bold
    
    text_frame.word_wrap = True

def replace_image(slide, placeholder_tag: str, image_bytes: bytes, crop_to_frame: bool = False):
    """Erstatter et billede i en placeholder. Skalerer proportionelt og centrerer i rammen."""
    
    placeholder = get_placeholder_by_tag(slide, placeholder_tag)
    if placeholder is None:
        return
        
    left, top, width, height = placeholder.left, placeholder.top, placeholder.width, placeholder.height
    
    try:
        image_stream = io.BytesIO(image_bytes)
        pic = slide.shapes.add_picture(image_stream, left, top, width, height)
        
        sp = placeholder.element
        sp.getparent().remove(sp)
        
        image_stream.seek(0)
        pic = slide.shapes.add_picture(image_stream, left, top, width, height)

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
        pic.top = top + (height - pic.height) // 2

        if crop_to_frame:
            offset_x = (new_w - frame_w) / (2 * new_w) 
            offset_y = (new_h - frame_h) / (2 * new_h) 

            pic.crop_left = offset_x
            pic.crop_right = offset_x
            pic.crop_top = offset_y
            pic.crop_bottom = offset_y
            
            pic.left = left
            pic.top = top
            pic.width = width
            pic.height = height

    except Exception as e:
        st.warning(f"Advarsel: Kunne ikke indsætte billede for placeholder '{placeholder_tag}'. Fejl: {e}")

# Rettelser på type hints: Erstatter Presentation.slide med Any
def find_first_slide_with_tag(prs: Presentation, tag: str) -> Tuple[Any, int]:
    """Finder det første slide, der indeholder en bestemt placeholder-tag (til skabelonsøgning)."""
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip().upper() == tag.upper():
                return slide, i
    st.error(f"Fejl: Skabelonen ('input-template.pptx') mangler en slide med placeholder-tekst: {tag}")
    raise ValueError("Placeholder ikke fundet i skabelonen.")

def get_slide_index_by_tag(prs: Presentation, tag: str) -> int:
    """Returnerer indexet for det første slide, der indeholder en bestemt placeholder-tag."""
    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip().upper() == tag.upper():
                return i
    return -1

def fill_overview_slides(prs: Presentation, all_renderings: List[bytes]):
    """Opretter OVERVIEW-slides og indsætter renderinger."""
    if not all_renderings:
        return 0
        
    overview_slide, overview_index = find_first_slide_with_tag(prs, 'OVERVIEW')
    overview_layout = overview_slide.slide_layout
    
    num_renderings = len(all_renderings)
    num_overview_slides = (num_renderings + 11) // 12
    slides_created = 0

    for i in range(num_overview_slides):
        current_slide = prs.slides.add_slide(overview_layout)
        slides_created += 1
        
        start_index = i * 12
        end_index = min((i + 1) * 12, num_renderings)
        
        for j, render_bytes in enumerate(all_renderings[start_index:end_index]):
            placeholder_tag = OVERVIEW_RENDER_PLACEHOLDERS[j]
            replace_image(current_slide, placeholder_tag, render_bytes, crop_to_frame=True)
            
        if num_overview_slides > 1:
            title_shape = get_placeholder_by_tag(current_slide, 'OVERVIEW')
            if title_shape:
                 fit_replace_text(title_shape, f"OVERVIEW (SIDE {i+1} AF {num_overview_slides})")


    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[overview_index])
    
    return slides_created

def fill_setting_slides(prs: Presentation, setting_data: List[Dict[str, Any]], library_df: pd.DataFrame, master_df: pd.DataFrame) -> int:
    """Opretter setting-slides, fylder tekst og billeder (inkl. dynamiske packshots)."""
    
    if not setting_data:
        return 0
        
    setting_template_slide, template_index = find_first_slide_with_tag(prs, '{{SETTINGNAME}}')
    setting_layout = setting_template_slide.slide_layout
    
    total_slides_created = 0
    
    for setting in setting_data:
        setting_name = setting['name']
        
        pcon_df = load_pcon_file(setting['pcon_file'])
        product_list_text, product_details, warnings = build_products_list(pcon_df, library_df, master_df) 
        
        for warning in warnings:
            st.warning(f"[{setting_name}] {warning}")
        
        num_products = len(product_details)
        num_product_slides = (num_products + 11) // 12
        
        for i in range(num_product_slides):
            is_first_slide = (i == 0)
            
            current_slide = prs.slides.add_slide(setting_layout)
            total_slides_created += 1
            
            fit_replace_text(get_placeholder_by_tag(current_slide, '{{SETTINGNAME}}'), setting_name)
            
            if is_first_slide:
                fit_replace_text(get_placeholder_by_tag(current_slide, '{{SETTINGSUBHEADLINE}}'), setting['subheadline'])
                fit_replace_text(get_placeholder_by_tag(current_slide, '{{SettingDimensions}}'), setting['dimensions'])
                fit_replace_text(get_placeholder_by_tag(current_slide, '{{SettingSize}}'), setting['size'])
                fit_replace_text(get_placeholder_by_tag(current_slide, '{{ProductsinSettingList}}'), product_list_text)

                replace_image(current_slide, '{{Rendering}}', preprocess_image(setting['rendering_bytes']))
                
                if setting['linedrawing_bytes']:
                    replace_image(current_slide, '{{Linedrawing}}', preprocess_image(setting['linedrawing_bytes']))
                
                # Sikrer at Accessory placeholders er tomme
                for k in range(6):
                    placeholder_tag = ACCESSORY_PLACEHOLDERS[k]
                    fit_replace_text(get_placeholder_by_tag(current_slide, placeholder_tag), "")
            else:
                 # Ryd felter på efterfølgende slides
                fit_replace_text(get_placeholder_by_tag(current_slide, '{{SETTINGSUBHEADLINE}}'), "")
                fit_replace_text(get_placeholder_by_tag(current_slide, '{{ProductsinSettingList}}'), "")
            
            # --- Udfyld produkt- og packshot-placeholders ---
            start_prod_index = i * 12
            end_prod_index = min((i + 1) * 12, num_products)
            
            for j, product_detail in enumerate(product_details[start_prod_index:end_prod_index]):
                prod_index = j
                
                # 1. Produktbeskrivelse
                prod_desc_tag = PRODUCT_PLACEHOLDERS[prod_index]
                fit_replace_text(get_placeholder_by_tag(current_slide, prod_desc_tag), product_detail['description'])
                
                # 2. Packshot (Hent dynamisk)
                packshot_tag = PACKSHOT_PLACEHOLDERS[prod_index]
                packshot_url = product_detail.get('packshot_url')

                if packshot_url:
                    packshot_bytes = fetch_image_bytes_from_url(packshot_url, product_detail['article_no'])
                    
                    if packshot_bytes:
                         replace_image(current_slide, packshot_tag, preprocess_image(packshot_bytes))
                
            # Opdater sidens titel for pagination
            if num_product_slides > 1:
                fit_replace_text(get_placeholder_by_tag(current_slide, '{{SETTINGNAME}}'), f"{setting_name} (Produkter: Side {i+1} af {num_product_slides})")


    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[template_index])
    
    return total_slides_created

def export_pptx(prs: Presentation) -> bytes:
    """Gemmer præsentationen til en BytesIO-buffer."""
    with io.BytesIO() as buffer:
        prs.save(buffer)
        return buffer.getvalue()


# --- STREAMLIT UI OG HOVEDPROGRAM ---

def main():
    st.set_page_config(page_title="PowerPoint Generator", layout="wide")
    st.title("📄 Muuto PowerPoint Generator")
    st.markdown("---")

    # --- Sektion: 1. Opslagsfiler og Skabelon ---
    st.header("1. Opslagsfiler (Google Sheets) og Skabelon")
    
    # Standard URL'er (Skal være publiceret som CSV)
    LIBRARY_DEFAULT_URL = "https://docs.google.com/spreadsheets/d/1h3yaq094mBa5Yadfi9Nb_Wrnzj3gIH2DviIfU0DdwsQ/export?format=csv&gid=437866492"
    MASTER_DEFAULT_URL = "https://docs.google.com/spreadsheets/d/1blj42SbFpszWGyOrDOUwyPDJr9K1NGpTMX6eZTbt_P4/export?format=csv&gid=194572316"

    col_lib, col_master, col_template = st.columns(3)

    with col_lib:
        library_url = st.text_input(
            "URL til **Library_data**", 
            key="library_url", 
            value=LIBRARY_DEFAULT_URL
        )
    with col_master:
        master_url = st.text_input(
            "URL til **Muuto Master Data**", 
            key="master_url", 
            value=MASTER_DEFAULT_URL
        )
    with col_template:
        st.write("Upload **input-template.pptx** (Kræves)")
        template_file = st.file_uploader(
            "input-template.pptx", 
            type=['pptx'], 
            key="template_upload",
            label_visibility="collapsed" # Skjul standardlabel
        )
        if template_file:
            st.success("Skabelon uploadet.")
        
    st.markdown("---")
    
    # --- Sektion: 2. Upload Alle Setting Filer ---
    st.header("2. Upload Alle Setting Filer")
    st.warning("Upload alle filer (CSV/XLSX, Rendering.jpg, Floorplan.jpg) på én gang. Filerne skal have et **fælles prefix** for at blive grupperet.")

    st.write("Multi-upload: CSV/XLSX, Rendering JPG/PNG, Floorplan JPG/PNG")
    all_setting_files = st.file_uploader(
        "Setting Filer", 
        type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], 
        accept_multiple_files=True,
        key="all_setting_files_upload",
        label_visibility="collapsed" # Skjul standardlabel
    )
    
    if all_setting_files:
        st.success(f"{len(all_setting_files)} filer er uploadet og klar til gruppering.")

    # Valgfri manuelle inputs (gælder for *Alle* settings)
    st.subheader("Manuelle Inputs (Gælder for *Alle* settings)")
    
    col_sub, col_dim, col_size = st.columns(3)
    with col_sub:
        manual_subheadline = st.text_input("SETTINGSUBHEADLINE", key="manual_subheadline", value="")
    with col_dim:
        manual_dimensions = st.text_input("SettingDimensions", key="manual_dimensions", value="")
    with col_size:
        manual_size = st.text_input("SettingSize", key="manual_size", value="")

    st.markdown("---")
    
    # --- Sektion: 3. Generér PowerPoint ---
    st.header("3. Generér PowerPoint")
    
    if st.button("🚀 Generér PowerPoint", type="primary"):
        
        # --- 1. Validering og Gruppering ---
        errors = []
        if not template_file:
            errors.append("❌ Skabelonen (input-template.pptx) mangler.")
        if not library_url or not master_url:
            errors.append("❌ URL til Library eller Master Data mangler.")
            
        if errors:
            for error in errors:
                st.error(error)
            st.stop()
            
        grouped_settings = group_uploaded_files(all_setting_files)
        
        if not grouped_settings and all_setting_files:
             st.error("❌ Ingen gyldige settings kunne dannes fra de uploadede filer. Tjek filnavnekonventionen.")
             st.stop()

        valid_setting_data_for_processing = []
        all_renderings_bytes = []
        
        for base_name, data in grouped_settings.items():
            
            rendering_bytes = data['rendering'].read()
            all_renderings_bytes.append(rendering_bytes)
            
            linedrawing_bytes = data['floorplan'].read() if data['floorplan'] else None
            
            setting = {
                'name': data['name'], 
                'subheadline': manual_subheadline,
                'dimensions': manual_dimensions,
                'size': manual_size,
                'pcon_file': data['csv'],
                'rendering_bytes': rendering_bytes,
                'linedrawing_bytes': linedrawing_bytes
            }
            valid_setting_data_for_processing.append(setting)

        # --- 2. Databehandling ---
        with st.spinner("Læser opslagsdata og skabelon..."):
            try:
                # Bemærk: 'IMAGE URL' kolonnen forventes her
                library_df = load_library_data(library_url, ['PRODUCT', 'EUR ITEM NO.'], 'Library_data')
                master_df = load_library_data(master_url, ['ITEM NO.', MASTER_DATA_PACKSHOT_COL], 'Muuto Master Data') 
                
                template_bytes = template_file.read()
                prs = Presentation(io.BytesIO(template_bytes))
            except Exception as e:
                st.error(f"Kritisk fejl under indlæsning af skabelon eller opslagsdata: {e}")
                st.stop()

        # --- 3. Udfyld PowerPoint ---
        with st.spinner("Genererer PowerPoint-slides og henter packshots..."):
            
            try:
                fill_overview_slides(prs, all_renderings_bytes)
                fill_setting_slides(prs, valid_setting_data_for_processing, library_df, master_df)
                
            except Exception as e:
                st.error(f"Kritisk fejl under generering af slides: {e}")
                st.exception(e)
                st.stop()


        # --- 4. Eksport ---
        final_pptx_bytes = export_pptx(prs)
        
        st.success("✅ PowerPoint genereret succesfuldt!")
        
        st.download_button(
            label="⬇️ Download Færdig PowerPoint",
            data=final_pptx_bytes,
            file_name="Muuto_Setting_Oversigt_Automatisk.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        st.balloons()
        
if __name__ == "__main__":
    main()
