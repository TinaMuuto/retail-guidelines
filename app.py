# app.py

import streamlit as st
import pandas as pd
from pptx import Presentation
# Importer nødvendige pptx/PIL/io moduler (forkortet for plads, men de skal være inkluderet)
from pptx.util import Inches
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import io
import re
from typing import List, Dict, Any, Tuple
from collections import defaultdict
import numpy as np

# --- Konstanter ---
PCON_ARTICLE_NO_COL = 17
PCON_QUANTITY_COL = 30
PCON_SHORT_TEXT_COL = 2
PCON_VARIANT_TEXT_COL = 4
PCON_SKIPROWS = 2
PRODUCT_PLACEHOLDERS = [f"PRODUCT DESCRIPTION {i}" for i in range(1, 13)]
PACKSHOT_PLACEHOLDERS = [f"ProductPackshot{i}" for i in range(1, 13)]
ACCESSORY_PLACEHOLDERS = [f"accessory{i}" for i in range(1, 7)]
OVERVIEW_RENDER_PLACEHOLDERS = [f"Rendering{i}" for i in range(1, 13)]

# --- HJÆLPEFUNKTIONER (Beholdes uændret) ---

@st.cache_data
def load_pcon_file(uploaded_file) -> pd.DataFrame:
    # Funktion til at indlæse pCon-filen... (Beholdes uændret)
    # ... (Se tidligere kode for detaljer)
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
    # Funktion til at indlæse data fra Google Sheets URL... (Beholdes uændret)
    # ... (Se tidligere kode for detaljer)
    if not input_url.startswith('http'):
        st.error(f"Fejl: '{source_name}' mangler en gyldig URL.")
        raise ValueError("URL mangler.")

    try:
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
        
# Beholder: fallback_key, match_library, match_master, get_product_description, build_products_list
# Beholder: preprocess_image
# Beholder: PowerPoint-funktioner (get_placeholder_by_tag, fit_replace_text, replace_image, etc.)

# --- NY FILGRUPPERINGSFUNKTION ---

def group_uploaded_files(uploaded_files: List[io.BytesIO]) -> Dict[str, Dict[str, Any]]:
    """
    Grupperer filer i settings baseret på det fælles filnavn.

    Filnavne forventes at være i format: "UniktNavn_Detalje.ext".
    Det unikke navn (første segment) bruges til at gruppere.
    """
    settings_grouped = defaultdict(lambda: {'csv': None, 'rendering': None, 'floorplan': None, 'name': None})
    
    for file in uploaded_files:
        filename = file.name
        
        # Forsøg at finde det fælles navn (første segment, fx "OsloLivingRoom")
        # Vi splitter ved den første understregning eller bindestreg som en robust separator
        base_name = re.split(r'[_-]', filename, 1)[0].strip() 
        
        if not base_name:
            continue # Ignorer filer uden et genkendeligt navn
            
        # Standardiser navnet (første bogstav stort, resten småt for placeholders)
        setting_name_standardized = base_name.replace('_', ' ').strip().title()
        
        # Tildel filen til den korrekte type
        settings_grouped[base_name]['name'] = setting_name_standardized

        filename_lower = filename.lower()
        
        if filename_lower.endswith('.csv') or filename_lower.endswith('.xlsx'):
            settings_grouped[base_name]['csv'] = file
        elif 'floorplan' in filename_lower:
            settings_grouped[base_name]['floorplan'] = file
        elif filename_lower.endswith('.jpg') or filename_lower.endswith('.jpeg') or filename_lower.endswith('.png'):
            # Antager at en resterende JPG/PNG er renderingen
            settings_grouped[base_name]['rendering'] = file
            
    # Konverter defaultdict til en almindelig dict for at bevare de indlæste navne
    final_settings = {}
    for base_name, data in settings_grouped.items():
        if data['csv'] and data['rendering']:
            final_settings[base_name] = data
        else:
            # Fejlfinding: En setting mangler den obligatoriske CSV eller Rendering
            st.warning(f"⚠️ Setting '{data['name']}' blev ignoreret: Mangler CSV ({'OK' if data['csv'] else 'Mangler'}) eller Rendering ({'OK' if data['rendering'] else 'Mangler'}).")

    return final_settings

# --- HOVEDLOGIK (Opdateret UI) ---

def main():
    st.set_page_config(page_title="PowerPoint Generator", layout="wide")
    st.title("📄 Muuto PowerPoint Generator (Automatisk)")
    st.markdown("---")

    # --- Sektion: 1. Opslagsfiler og Skabelon ---
    st.header("1. Opslagsfiler (Google Sheets) og Skabelon")
    
    # Standard URL'er (Brugere kan redigere dem, men de er forudfyldt)
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
        template_file = st.file_uploader(
            "Upload **input-template.pptx** (Kræves)", 
            type=['pptx'], 
            key="template_upload"
        )
        
    st.markdown("---")
    
    # --- Sektion: 2. Upload Alle Setting Filer (NYT) ---
    st.header("2. Upload Alle Setting Filer")
    st.warning("Upload alle filer (CSV/XLSX, Rendering.jpg, Floorplan.jpg) på én gang. Filerne skal have et **fælles prefix** for at blive grupperet (fx 'OsloRoom_PCon.csv', 'OsloRoom_Render.jpg', 'OsloRoom_Floorplan.jpg').")

    all_setting_files = st.file_uploader(
        "Multi-upload: CSV/XLSX, Rendering JPG/PNG, Floorplan JPG/PNG", 
        type=['csv', 'xlsx', 'jpg', 'jpeg', 'png'], 
        accept_multiple_files=True,
        key="all_setting_files_upload"
    )
    
    # Valgfri manuelle inputs (da de ikke er tilgængelige i filnavnet)
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
            
        if not all_setting_files:
            errors.append("❌ Der er ikke uploadet nogen filer til settings.")
        
        if errors:
            for error in errors:
                st.error(error)
            st.stop()
            
        # Gruppér filer automatisk (Returns Dict[base_name, Dict[file_type, file_object]])
        grouped_settings = group_uploaded_files(all_setting_files)
        
        if not grouped_settings:
            st.error("❌ Ingen gyldige settings at behandle. Tjek, at du har mindst én CSV/XLSX og én Rendering (.jpg/.png) med et fælles navn.")
            st.stop()

        # Konverter til den gamle list-struktur for at genbruge fill-funktionerne
        valid_setting_data_for_processing = []
        all_renderings_bytes = []
        
        for base_name, data in grouped_settings.items():
            
            # Læs alle bytes i hukommelsen
            rendering_bytes = data['rendering'].read()
            all_renderings_bytes.append(rendering_bytes)
            
            linedrawing_bytes = data['floorplan'].read() if data['floorplan'] else None
            
            setting = {
                'name': data['name'], # Navnet hentet fra filnavn
                'subheadline': manual_subheadline,
                'dimensions': manual_dimensions,
                'size': manual_size,
                'pcon_file': data['csv'],
                'rendering_bytes': rendering_bytes,
                'linedrawing_bytes': linedrawing_bytes,
                'packshot_bytes_list': [], # Tom liste (ingen uploadet)
                'accessories': [] # Tom liste (ingen uploadet)
            }
            valid_setting_data_for_processing.append(setting)

        # --- 2. Databehandling ---
        with st.spinner("Læser opslagsdata fra Google Sheets og skabelon..."):
            try:
                library_df = load_library_data(library_url, ['PRODUCT', 'EUR ITEM NO.'], 'Library_data')
                master_df = load_library_data(master_url, ['ITEM NO.'], 'Muuto Master Data') 
                
                template_bytes = template_file.read()
                prs = Presentation(io.BytesIO(template_bytes))
            except Exception as e:
                st.error(f"Kritisk fejl under indlæsning af skabelon eller opslagsdata: {e}")
                st.stop()

        # --- 3. Udfyld PowerPoint ---
        with st.spinner("Genererer PowerPoint-slides..."):
            
            try:
                # Udfyld OVERVIEW før settings
                fill_overview_slides(prs, all_renderings_bytes)
                
                # Udfyld settings
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
