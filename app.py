import io
import re
import time
from typing import Dict, List, Optional, Tuple, Any
from pathlib import Path
from copy import deepcopy

import pandas as pd
import requests
import streamlit as st
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE, PP_PLACEHOLDER

# --- CONSTANTS ---
OUTPUT_NAME = "Muuto_Settings_Generated.pptx"
TEMPLATE_PATH = Path("input-template.pptx")

# NYE URLs afledt af bruger input. Disse SKAL være offentligt delte (Anyone with the link)!
# Master Data URL (tidligere 1blj42SbFpszWGyOrDOUwyPDJr9K1NGpTMX6eZTbt_P4)
DEFAULT_MASTER_URL = "https://docs.google.com/spreadsheets/d/1blj42SbFpszWGyOrDOUwyPDJr9K1NGpTMX6eZTbt_P4/pub?output=csv&gid=1152340088"
# Mapping Data URL (tidligere 1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk)
DEFAULT_MAPPING_URL = "https://docs.google.com/spreadsheets/d/1S50it_q1BahpZCPW8dbuN7DyOMnyDgFIg76xIDSoXEk/pub?output=csv&gid=1056617222"


# Caching for dataframes and images
@st.cache_data(ttl=3600)
def http_get_bytes(url: str, timeout: int = 15) -> Optional[io.BytesIO]:
    """Henter indhold fra URL som bytes (bruges til CSV og billeder)."""
    try:
        response = requests.get(url, timeout=timeout)
        response.raise_for_status()
        return io.BytesIO(response.content)
    except requests.exceptions.RequestException as e:
        st.error(f"HTTP GET fejlede for {url}: {e}")
        return None

def parse_csv_flex(csv_bytes: io.BytesIO) -> pd.DataFrame:
    """Læser CSV bytes fleksibelt og håndterer forskellige separators."""
    # Prøver med standard separatorer: ';' (europeisk), ',' (US/UK) og '\t' (tab)
    separators = [';', ',', '\t']
    for sep in separators:
        try:
            csv_bytes.seek(0)
            df = pd.read_csv(csv_bytes, sep=sep, encoding='utf-8', on_bad_lines='skip', skipinitialspace=True)
            if not df.empty and df.shape[1] > 1: # Tjekker om DF ikke er tom og har mere end én kolonne
                return df
        except Exception:
            pass
    
    # Sidste forsøg med Latin-1, hvis UTF-8 fejlede
    for sep in separators:
        try:
            csv_bytes.seek(0)
            df = pd.read_csv(csv_bytes, sep=sep, encoding='latin-1', on_bad_lines='skip', skipinitialspace=True)
            if not df.empty and df.shape[1] > 1:
                return df
        except Exception:
            continue

    return pd.DataFrame()

def normalize_column_names(df: pd.DataFrame, target_columns: Dict[str, List[str]]) -> pd.DataFrame:
    """Normaliserer kolonnenavne baseret på en liste af potentielle navne."""
    df.columns = [str(col).strip() for col in df.columns]
    mapping = {}
    
    for standard_name, potential_names in target_columns.items():
        found = False
        for potential in potential_names:
            # Case-insensitive match, fjerner mellemrum og tegn.
            clean_potential = re.sub(r'[^a-z0-9]', '', potential.lower())
            for current_col in df.columns:
                clean_current = re.sub(r'[^a-z0-9]', '', current_col.lower())
                if clean_current == clean_potential:
                    mapping[current_col] = standard_name
                    found = True
                    break
            if found:
                break
    
    df.rename(columns=mapping, inplace=True)
    return df

def normalize_master(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliserer Master Data kolonnenavne (til Packshot URL)."""
    target_columns = {
        "ARTICLE_NO": ["Article No.", "Item No.", "Item", "Artikelnummer"],
        "PACKSHOT_URL": ["Packshot_URL", "URL_Packshot", "Image_URL"]
    }
    df = normalize_column_names(df, target_columns)
    # Beholder kun de nødvendige kolonner
    df = df.loc[:, df.columns.isin(target_columns.keys())]
    if "ARTICLE_NO" in df.columns:
        df["ARTICLE_NO"] = df["ARTICLE_NO"].astype(str).str.strip()
    return df

def normalize_mapping(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliserer Mapping Data kolonnenavne (til Beskrivelse og Nyt Nummert)."""
    target_columns = {
        "OLD_ARTICLE_NO": ["OLD Item-variant", "Old Item", "Gammelt Varenummer"],
        "NEW_ARTICLE_NO": ["New Item No.", "New Item", "Nyt Varenummer"],
        "DESCRIPTION": ["Description", "Beskrivelse", "Product Description"]
    }
    df = normalize_column_names(df, target_columns)
    # Beholder kun de nødvendige kolonner
    df = df.loc[:, df.columns.isin(target_columns.keys())]
    if "OLD_ARTICLE_NO" in df.columns:
        df["OLD_ARTICLE_NO"] = df["OLD_ARTICLE_NO"].astype(str).str.strip()
    return df

def normalize_pcon(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliserer Pcon CSV kolonnenavne (fra upload)."""
    target_columns = {
        "ARTICLE_NO": ["Article No.", "External Item Number", "Artikelnummer"],
        "Quantity": ["Quantity", "Antal"]
    }
    df = normalize_column_names(df, target_columns)
    # Beholder kun de nødvendige kolonner
    df = df.loc[:, df.columns.isin(target_columns.keys())]
    if "ARTICLE_NO" in df.columns:
        df["ARTICLE_NO"] = df["ARTICLE_NO"].astype(str).str.strip()
    # Sikrer Quantity er tilgængelig
    if "Quantity" not in df.columns:
        df["Quantity"] = 1 
    # Fjerner rækker uden et varenummer
    df.dropna(subset=['ARTICLE_NO'], inplace=True)
    return df

# --- Utils - PowerPoint ---

def ensure_presentation_from_path(path: Path) -> Presentation:
    """Opretter eller sikrer en Presentation fra en Path-objekt."""
    return Presentation(path)

def find_layout_by_name(prs: Presentation, name: str) -> Optional[Any]:
    """Finder et Slide Layout efter navn (case-insensitive)."""
    for layout in prs.slide_masters[0].slide_layouts:
        if layout.name.lower() == name.lower():
            return layout
    return None

def layout_has_expected(layout: Any, expected_placeholders: List[str]) -> bool:
    """Tjekker om et layout har de forventede placeholders (Shape names)."""
    found_names = {shape.name for shape in layout.shapes}
    return all(name in found_names for name in expected_placeholders)

def replace_text_in_placeholder(slide, name: str, new_text: str):
    """Finder en shape med 'name' og erstatter dens tekst."""
    for shape in slide.shapes:
        if shape.has_text_frame and shape.name == name:
            text_frame = shape.text_frame
            text_frame.clear()
            p = text_frame.paragraphs[0]
            p.text = new_text
            p.font.size = Pt(24)
            return

def replace_image_in_placeholder(slide, name: str, image_bytes: io.BytesIO):
    """Finder et placeholder-billede og erstatter kilden."""
    for shape in slide.shapes:
        if shape.name == name:
            # Gemmer position og størrelse fra den eksisterende shape
            left, top, width, height = shape.left, shape.top, shape.width, shape.height
            
            # Sletter den gamle shape
            sp = shape.element
            sp.getparent().remove(sp)
            
            # Indsætter det nye billede og tilpasser dets størrelse og position
            try:
                slide.shapes.add_picture(image_bytes, left, top, width, height)
            except Exception as e:
                st.warning(f"Kunne ikke indsætte billede for '{name}': {e}")
            return

def add_table_to_slide(slide, table_data: List[List[str]], table_anchor_name: str):
    """Tilføjer en tabel til et slide og bruger TableAnchor for positionering."""
    anchor = None
    for shape in slide.shapes:
        if shape.name == table_anchor_name:
            anchor = shape
            break
    
    if anchor is None:
        st.warning(f"WARNING: '{table_anchor_name}' mangler i '{slide.slide_layout.name}'. Bruger standard position.")
        left = Inches(1)
        top = Inches(2)
        width = Inches(8)
        height = Inches(4)
    else:
        left, top, width, height = anchor.left, anchor.top, anchor.width, anchor.height

    if not table_data:
        return
        
    num_rows = len(table_data)
    num_cols = len(table_data[0])

    shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
    table = shape.table

    # Overskrift (Headers)
    for col_idx, text in enumerate(table_data[0]):
        cell = table.cell(0, col_idx)
        cell.text = text
        text_frame = cell.text_frame
        p = text_frame.paragraphs[0]
        p.font.bold = True
        p.font.size = Pt(12)

    # Indhold (Body)
    for row_idx, row_data in enumerate(table_data[1:]):
        for col_idx, text in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = text
            text_frame = cell.text_frame
            p = text_frame.paragraphs[0]
            p.font.size = Pt(10)

def safe_present(prs: Presentation) -> io.BytesIO:
    """Gemmer præsentationen i en BytesIO-buffer."""
    with io.BytesIO() as buffer:
        prs.save(buffer)
        buffer.seek(0)
        return buffer

# --- Utils - Grouping and Mapping Logic ---

def find_packshot_url(article_no: str, master_df: pd.DataFrame) -> Optional[str]:
    """Slår Packshot URL op i Master Data DF. Prioriterer eksakt match."""
    match = master_df[master_df["ARTICLE_NO"] == article_no]
    if not match.empty and "PACKSHOT_URL" in match.columns:
        return match["PACKSHOT_URL"].iloc[0]
    return None

def find_mapping_data(article_no: str, mapping_df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    """Slår New Article No. og Beskrivelse op i Mapping Data DF."""
    # Prioriterer eksakt match på det fulde varenummer
    match = mapping_df[mapping_df["OLD_ARTICLE_NO"] == article_no]
    
    if match.empty:
        # Fallback: Prøver at matche på basis-varenummeret (første del)
        base_article = article_no.split('-')[0].split('_')[0]
        match = mapping_df[mapping_df["OLD_ARTICLE_NO"] == base_article]
    
    if not match.empty:
        desc = match["DESCRIPTION"].iloc[0] if "DESCRIPTION" in match.columns else None
        new_no = match["NEW_ARTICLE_NO"].iloc[0] if "NEW_ARTICLE_NO" in match.columns else None
        return desc, new_no
    
    return None, None # Hvis intet match

def get_product_details(article_no: str, mapping_df: pd.DataFrame, master_df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Samler alle detaljer for et givent varenummer."""
    details = {"DESCRIPTION": None, "NEW_ARTICLE_NO": None, "PACKSHOT_URL": None}
    
    if mapping_df.empty:
        st.info("Mapping data er tom. Kan ikke finde beskrivelse/nyt varenummer.")
    else:
        desc, new_no = find_mapping_data(article_no, mapping_df)
        details["DESCRIPTION"] = desc
        details["NEW_ARTICLE_NO"] = new_no
        
    # Packshot URL skal findes via det NYE varenummer, hvis det findes, ellers det GAMLE.
    lookup_no = details["NEW_ARTICLE_NO"] if details["NEW_ARTICLE_NO"] else article_no
    
    if master_df.empty:
        st.info("Master data er tom. Kan ikke finde Packshot URL.")
    else:
        packshot_url = find_packshot_url(lookup_no, master_df)
        if packshot_url is None:
            # Prøv fallback til det gamle nummer, hvis det nye nummer ikke virkede
            if details["NEW_ARTICLE_NO"] and details["NEW_ARTICLE_NO"] != article_no:
                packshot_url = find_packshot_url(article_no, master_df)

        details["PACKSHOT_URL"] = packshot_url
        
    return details

# --- Utils - Preflight and Grouping ---

def preflight_checks() -> Dict[str, str]:
    """Kører tjek på template filen."""
    diag = {"template": "OK"}
    if not TEMPLATE_PATH.exists():
        diag["template"] = f"Template filen '{TEMPLATE_PATH}' mangler."
        return diag
    
    try:
        prs = Presentation(TEMPLATE_PATH)
        required_layouts = ["Overview", "Renderings", "Setting", "ProductListBlank"]
        for name in required_layouts:
            if not find_layout_by_name(prs, name):
                diag["template"] = f"Mangler påkrævet layout '{name}' i Slide Master."
                return diag

        # Specifik tjek for TableAnchor i ProductListBlank
        productlist_layout = find_layout_by_name(prs, "ProductListBlank")
        if productlist_layout and not layout_has_expected(productlist_layout, ["TableAnchor"]):
            st.warning("WARNING: 'TableAnchor' mangler i 'ProductListBlank'. Bruger standard position.")
            
    except Exception as e:
        diag["template"] = f"Fejl ved indlæsning af template: {e}"
        
    return diag

def build_groups(uploads: List[Any]) -> Dict[str, Dict[str, Any]]:
    """Grupperer uploads baseret på deres filnavn prefix."""
    groups: Dict[str, Dict[str, Any]] = {}
    
    for uploaded_file in uploads:
        name_parts = uploaded_file.name.rsplit('.', 1)
        if len(name_parts) < 2:
            continue
        
        base_name = name_parts[0]
        ext = name_parts[1].lower()
        
        # Ekstraherer gruppens navn fra filnavnet (navnet efter sidste bindestreg)
        group_match = re.search(r'-\s*([^-]+)$', base_name)
        group_key = group_match.group(1).strip().lower().replace(" ", "") if group_match else base_name.lower().replace(" ", "")

        if group_key not in groups:
            groups[group_key] = {"name": group_key, "csv": None, "render": None, "floorplan": None}

        if ext == 'csv':
            groups[group_key]['csv'] = uploaded_file.getvalue()
        elif ext in ['jpg', 'jpeg', 'png', 'webp']:
            # Tjekker om det er en floorplan ved at søge efter "floorplan" i navnet.
            if "floorplan" in uploaded_file.name.lower():
                groups[group_key]['floorplan'] = io.BytesIO(uploaded_file.getvalue())
            else:
                groups[group_key]['render'] = io.BytesIO(uploaded_file.getvalue())

    # Filtrer grupper der mangler CSV
    valid_groups = {k: v for k, v in groups.items() if v["csv"] is not None}
    return valid_groups

def collect_all_renderings(groups: Dict[str, Dict[str, Any]]) -> List[io.BytesIO]:
    """Samler alle renderings for overview-slides."""
    renders = []
    for g in groups.values():
        if g["render"]:
            renders.append(g["render"])
    return renders

# --- Slide Builders ---

def build_overview_slides(prs: Presentation, layout: Any, renders: List[io.BytesIO]):
    """Bygger Overview-slides med op til 12 renderings pr. slide."""
    
    # Henter placeholder navne, der ligner 'Rendering1', 'Rendering2', osv.
    render_placeholders = [shape.name for shape in layout.shapes if re.match(r"Rendering\d+", shape.name)]
    render_placeholders.sort(key=lambda x: int(re.search(r'\d+', x).group(0)))
    
    batch_size = len(render_placeholders)

    for i in range(0, len(renders), batch_size):
        slide = prs.slides.add_slide(layout)
        batch = renders[i:i + batch_size]
        
        for j, render_bytes in enumerate(batch):
            placeholder_name = render_placeholders[j]
            replace_image_in_placeholder(slide, placeholder_name, render_bytes)

def build_setting_slide(prs: Presentation, layout: Any, group_name: str, render_bytes: Optional[io.BytesIO], floorplan_bytes: Optional[io.BytesIO], pcon_df: pd.DataFrame, mapping_df: pd.DataFrame, master_df: pd.DataFrame):
    """Bygger et Setting-slide med billeder og produktdetaljer."""
    slide = prs.slides.add_slide(layout)
    
    replace_text_in_placeholder(slide, "SETTINGNAME", group_name.replace("-", " ").title())
    
    if render_bytes:
        replace_image_in_placeholder(slide, "Rendering", render_bytes)
    if floorplan_bytes:
        # Tjekker for det korrekte navn (Linedrawing)
        replace_image_in_placeholder(slide, "Linedrawing", floorplan_bytes)

    # Udfyld produktinformation (Packshot og Beskrivelse)
    for i, (_, row) in enumerate(pcon_df.head(12).iterrows()): # Max 12 produkter
        article_no = row["ARTICLE_NO"]
        
        # Henter detaljer (Packshot URL, Beskrivelse, Nyt Varenummer)
        details = get_product_details(article_no, mapping_df, master_df)
        
        desc = details["DESCRIPTION"]
        packshot_url = details["PACKSHOT_URL"]
        
        # Udfyld Beskrivelse
        desc_placeholder_name = f"PRODUCT DESCRIPTION {i + 1}"
        description_text = desc if desc else f"*** BESKRIVELSE MANGLER ({article_no}) ***"
        replace_text_in_placeholder(slide, desc_placeholder_name, description_text)

        # Udfyld Packshot
        if packshot_url:
            packshot_bytes = http_get_bytes(packshot_url)
            if packshot_bytes:
                packshot_placeholder_name = f"ProductPackshot{i + 1}"
                replace_image_in_placeholder(slide, packshot_placeholder_name, packshot_bytes)
            else:
                st.warning(f"Kunne ikke hente Packshot fra URL: {packshot_url}")
        else:
            st.warning(f"Packshot URL mangler for: {article_no}")

def build_productlist_slide(prs: Presentation, layout: Any, group_name: str, pcon_df: pd.DataFrame, mapping_df: pd.DataFrame):
    """Bygger et slide med en komplet produktliste i tabelformat."""
    slide = prs.slides.add_slide(layout)
    replace_text_in_placeholder(slide, "SETTINGNAME", f"Produkter – {group_name.replace('-', ' ').title()}")

    table_data = [["Quantity", "Description", "Article No. / New Item No."]]

    for _, row in pcon_df.iterrows():
        article_no = row["ARTICLE_NO"]
        quantity = str(int(row["Quantity"]))
        
        # Henter beskrivelse og nyt varenummer
        desc, new_no = find_mapping_data(article_no, mapping_df)
        
        description_text = desc if desc else "*** BESKRIVELSE MANGLER ***"
        
        article_text = article_no
        if new_no:
            article_text = f"{article_no} / {new_no}"
        
        table_data.append([quantity, description_text, article_text])

    add_table_to_slide(slide, table_data, "TableAnchor")


# ---------------------- UI ----------------------
st.set_page_config(page_title="Muuto PowerPoint Generator", layout="centered")
st.title("Muuto PowerPoint Generator")
st.write("Upload dine grupperede filer (CSV og billeder). App'en bruger den faste PowerPoint-template og henter Master Data og Mapping fra faste URLs.")

# Initialize session state for uploads and generated dataframes
if 'uploads' not in st.session_state:
    st.session_state.uploads = []
if 'last_master_df' not in st.session_state:
    st.session_state.last_master_df = None
if 'last_mapping_df' not in st.session_state:
    st.session_state.last_mapping_df = None

uploaded_files = st.file_uploader(
    "Upload dine indstillingsfiler (CSV, Rendering, Floorplan).",
    type=['csv', 'jpg', 'jpeg', 'png', 'webp', 'pptx'],
    accept_multiple_files=True
)

if uploaded_files:
    st.session_state.uploads = uploaded_files
    st.info(f"{len(st.session_state.uploads)} filer er klar til behandling.")

generate = st.button("Generér PowerPoint")

# ---------------------- Orchestration ----------------------
if generate:
    with st.spinner("Arbejder..."):
        diag = preflight_checks()
        if diag.get("template") != "OK":
            st.error("Template problem: " + diag.get("template", "Ukendt"))
        elif not TEMPLATE_PATH.exists():
            st.error("Template filen mangler i repository: input-template.pptx")
        else:
            try:
                groups = build_groups(st.session_state.uploads)

                if not groups:
                    st.error("Kunne ikke danne nogen grupper. Sørg for at filer er uploadet og filnavne indeholder CSV og mindst ét billede/floorplan.")
                    
                if groups:
                    prs = ensure_presentation_from_path(TEMPLATE_PATH)

                    overview_layout = find_layout_by_name(prs, "Overview") or find_layout_by_name(prs, "Renderings")
                    setting_layout = find_layout_by_name(prs, "Setting")
                    productlist_layout = find_layout_by_name(prs, "ProductListBlank")

                    # --- LIVE DATA LOAD MED DIAGNOSTICS ---
                    master_raw = load_remote_csv(DEFAULT_MASTER_URL)
                    mapping_raw = load_remote_csv(DEFAULT_MAPPING_URL)

                    master_df = normalize_master(master_raw)
                    mapping_df = normalize_mapping(mapping_raw)
                    
                    st.session_state.last_master_df = master_df
                    st.session_state.last_mapping_df = mapping_df
                    
                    # --- DIAGNOSTIC OUTPUT (HVIS MAPPING FEJLEDE) ---
                    if mapping_df.empty:
                        st.warning("ADVARSEL: Mapping Data (Beskrivelser/Nye Artikelnumre) kunne ikke indlæses. Beskrivelser og nye varenumre vil mangle.")
                        if not mapping_raw.empty:
                            st.warning("Mapping CSV blev hentet, men data normalisering fejlede. Tjek kolonnerne for korrekt navngivning (f.eks. 'OLD Item-variant', 'New Item No.', 'Description'):")
                            st.dataframe(mapping_raw.head(3).T) # Viser de første 3 rækker og kolonner for fejlfinding
                    
                    if master_df.empty:
                         st.warning("ADVARSEL: Master Data (Billed-URLs) kunne ikke indlæses. Packshots vil mangle.")
                    # ---------------------------------------------

                    # 1. Overview Slides
                    renders = collect_all_renderings(groups)
                    if renders:
                        if overview_layout and layout_has_expected(overview_layout, ["Rendering1"]):
                            build_overview_slides(prs, overview_layout, renders)
                    

                    # 2. Per group Slides
                    for key in sorted(groups.keys()):
                        g = groups[key]
                        group_name = g["name"]
                        
                        try:
                            pcon_df = normalize_pcon(parse_csv_flex(g["csv"]))
                        except Exception as e:
                            pcon_df = pd.DataFrame(columns=["ARTICLE_NO", "Quantity"])
                            st.error(f"Fejl ved parsing af CSV for gruppe '{group_name}': {e}")
                            
                        
                        if pcon_df.empty:
                            st.warning(f"ADVARSEL: Kunne ikke indlæse produktdata fra CSV for gruppe '{group_name}'. Springer slides over.")
                            continue 
                        
                        render_bytes = g.get("render")
                        floorplan_bytes = g.get("floorplan")

                        # Setting Slide
                        if setting_layout and layout_has_expected(setting_layout, ["SETTINGNAME", "Rendering", "Linedrawing", "ProductPackshot1"]):
                            build_setting_slide(prs, setting_layout, group_name, render_bytes, floorplan_bytes, pcon_df, mapping_df, master_df)
                        else:
                            st.info(f"INFO: Setting slide for '{group_name}' kunne ikke bruge template layout (Tjek placeholder navne).")

                        # Product List Slide
                        if productlist_layout:
                            build_productlist_slide(prs, productlist_layout, group_name, pcon_df, mapping_df)
                        else:
                            st.info(f"INFO: Product List slide for '{group_name}' kunne ikke bruge template layout.")

                    ppt_bytes = safe_present(prs)
                    st.success("Din præsentation er klar!")
                    st.download_button(
                        "Download Muuto_Settings.pptx",
                        data=ppt_bytes,
                        file_name=OUTPUT_NAME,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
            except Exception as e:
                st.error(f"Der opstod en fejl under generering af præsentationen: {e}")
                
# UI for data status
if st.session_state.last_master_df is not None and st.session_state.last_mapping_df is not None:
    st.subheader("Data Forbindelses Status")
    col_m, col_mp = st.columns(2)
    col_m.metric("Master Data Rækker", st.session_state.last_master_df.shape[0])
    col_mp.metric("Mapping Data Rækker", st.session_state.last_mapping_df.shape[0])
    if st.session_state.last_master_df.empty or st.session_state.last_mapping_df.empty:
        st.warning("ADVARSEL: Nul rækker indlæst fra Master/Mapping CSV'er. Tjek URL-tilgængelighed og kolonnenavne i CSV-dataen.")
