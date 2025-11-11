import io
import re
import math
import csv
from collections import defaultdict

# Imports for Google Sheets (Pandas simulation)
# NOTE: We assume this function is robustly implemented to handle flexible column naming
# and returns a list of dictionaries (rows) for each sheet.
def load_google_sheet_csv(filename):
    """
    Simulates loading a Google Sheet CSV export.
    In a real environment, this would load the CSV/Excel file.
    For this implementation, we return mock data structure based on the user's rules.
    """
    if 'Master' in filename:
        return [
            # --------------------------- FULDE MASTER DATA (SIMULERING) ---------------------------
            # Til Cover ArmChair
            {'ITEM NO.': '15172-OAK_COGN', 'IMAGE DOWNLOAD LINK': 'http://master/cover-armchair.jpg', 'IMAGE URL': 'http://master/cover-armchair-old.jpg'},
            {'ITEM NO.': 'COVARMU01RA101', 'IMAGE DOWNLOAD LINK': 'http://master/cover-armchair-new.jpg', 'IMAGE URL': ''}, 
            # Til Strand Pendant
            {'ITEM NO.': '22453', 'IMAGE DOWNLOAD LINK': 'http://master/strand-pendant.jpg', 'IMAGE URL': ''},
            {'ITEM NO.': 'SDPENC6001', 'IMAGE DOWNLOAD LINK': 'http://master/strand-pendant-new.jpg', 'IMAGE URL': ''},
            # Til Raise Glasses (Small)
            {'ITEM NO.': 'RAIGLS20S202', 'IMAGE DOWNLOAD LINK': 'http://master/raise-glasses-small.jpg', 'IMAGE URL': ''},
            # Til Raise Glasses (Large)
            {'ITEM NO.': 'RAIGLS30S202', 'IMAGE DOWNLOAD LINK': 'http://master/raise-glasses-large.jpg', 'IMAGE URL': ''},
            # Til Cosy Lamp
            {'ITEM NO.': '01033', 'IMAGE DOWNLOAD LINK': 'http://master/cosy-lamp.jpg', 'IMAGE URL': ''},
            {'ITEM NO.': 'CSYTBL02', 'IMAGE DOWNLOAD LINK': 'http://master/cosy-lamp-new.jpg', 'IMAGE URL': ''},
            # Til Raise Carafe
            {'ITEM NO.': 'RAICAR02', 'IMAGE DOWNLOAD LINK': 'http://master/raise-carafe.jpg', 'IMAGE URL': ''},
            # Outline Sofa (ingen ny varenummer, så vi matcher det gamle direkte)
            {'ITEM NO.': '09170', 'IMAGE DOWNLOAD LINK': 'http://master/outline-sofa.jpg', 'IMAGE URL': ''},
            # Ukendt (1234567) - Mangler stadig match
            # ------------------------------------------------------------------------------------
        ]
    elif 'Mapping' in filename:
        return [
            # Mapping data forbliver uændret for at teste matches:
            {'OLD Item-variant': '15172-OAK_COGN', 'New Item No.': 'COVARMU01RA101', 'Description': 'Cover ArmChair - OAK/Refine Cognac'},
            {'OLD Item-variant': '22453', 'New Item No.': 'SDPENC6001', 'Description': 'Strand Pendant – Ø60 - Closed'},
            {'OLD Item-variant': '09170', 'New Item No.': '', 'Description': 'Outline Sofa 2-seater - Steelcut 190'}, # Ingen nyt nummer
            {'OLD Item-variant': '01033', 'New Item No.': 'CSYTBL02', 'Description': 'Cosy in Grey'},
            {'OLD Item-variant': '12429', 'New Item No. ' + '\ufeff': 'RAICAR02', 'Description': 'Raise - Carafe - Clear'},
            {'OLD Item-variant': '12431', 'New Item No.': 'RAIGLS30S202', 'Description': 'Raise - Glasses - Large - Clear'},
            {'OLD Item-variant': '12430', 'New Item No.': 'RAIGLS20S202', 'Description': 'Raise - Glasses - Small - Clear'},
            # Urelaterede varer som tester fallback
            {'OLD Item-variant': '15172', 'New Item No.': 'COVARMBASE', 'Description': 'Cover ArmChair Base Description'},
        ]
    return []

# --- Resten af koden er uændret ---

def normalize_key(key):
# ... (Funktioner er ikke vist for brevity, men er uændrede)
def get_base_article_no(article_no):
# ...
def find_column_index(header_row, *possible_names):
# ...
def safe_int(value, default=1):
# ...
def lookup_data(article_no, data_source, key_field, value_fields):
# ...
def parse_pcon_csv(csv_content):
# ...
class MuutoPPTXGenerator:
# ... (Alle metoder er uændrede)
    def is_rendering(self, filename):
# ...
    def is_linedrawing(self, filename):
# ...
    def get_setting_name_from_filename(self, filename):
# ...
    def group_files_by_setting(self):
# ...
    def process_all_settings(self):
# ...
    def generate_output_pptx(self):
# ...

# --- Udfør simuleringen og generer den endelige fil ---

generator = MuutoPPTXGenerator(
    uploaded_files=[
        # Dining 01 Group
        {"fileName": "Shop-the-look_2025_Q2 - dining 01.jpg", "fileMimeType": "image/jpeg", "contentFetchId": "id-dining-r"},
        {"fileName": "Shop-the-look_2025_Q2 - dining 01 floorplan.jpg", "fileMimeType": "image/jpeg", "contentFetchId": "id-dining-l"},
        {"fileName": "Shop-the-look_2025_Q2 - dining 01.csv", "fileMimeType": "text/csv", "contentFetchId": "id-dining-c",
         "snippetFromFront": "Position;Image;Short Text;Long Text;Variant Text;...;Article No.;...;Quantity;...\r\n",
         "snippetFromBack": "1;;Cover ArmChair - OAK/Refine Cognac;...;15172-OAK_COGN;...;6\r\n2;;Strand Pendant – Ø60 - Closed;...;22453;...;1\r\n3;;Raise - Glasses - Small - Clear;...;12430;...;3\r\n4;;Outline Sofa 2-seater;...;09170;...;1\r\n5;;Ukendt Produkt;...;1234567;...;1\r\n"}
        ,
        # Reading Corner 01 Group
        {"fileName": "Shop-the-look_2025_Q2 - reading corner 01.jpg", "fileMimeType": "image/jpeg", "contentFetchId": "id-reading-r"},
        {"fileName": "Shop-the-look_2025_Q2 - reading corner 01 floorplan.jpg", "fileMimeType": "image/jpeg", "contentFetchId": "id-reading-l"},
        {"fileName": "Shop-the-look_2025_Q2 - reading corner 01.csv", "fileMimeType": "text/csv", "contentFetchId": "id-reading-c",
         "snippetFromFront": "Position;Image;Short Text;Long Text;Variant Text;...;Article No.;...;Quantity;...\r\n",
         "snippetFromBack": "1;;Cosy in Grey;...;01033;...;1\r\n2;;Raise Carafe;...;12429;...;1\r\n"}
        ,
        # Ugyldig gruppe (mangler CSV)
        {"fileName": "Shop-the-look_2025_Q2 - invalid setting 01.jpg", "fileMimeType": "image/jpeg", "contentFetchId": "id-invalid-r"},
    ]
)

# Generer den simulerede output-fil
pptx_output = generator.generate_output_pptx()
