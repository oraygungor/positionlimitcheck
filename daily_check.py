import requests
import pandas as pd
import os
import shutil
import json
import hashlib
from datetime import datetime
from collections import Counter

# --- AYARLAR ---
URL = "https://www.esma.europa.eu/sites/default/files/position_limits_publication.xlsx"
LATEST_FILE = "positionlimit-latest.xlsx"
PREVIOUS_FILE = "positionlimit-previous.xlsx"

# Eşleştirme için kullanılacak "Kimlik" kolonları (Dosya yapısına göre)
# Bu kolonlar aynıysa, satır aynıdır; sadece içindeki diğer değerler değişmiştir.
# Unnamed: 1 -> MIC Code (XPAR vs)
# Unnamed: 3 -> Instrument Name (Euronext Milling Wheat vs)
KEY_COLUMNS = ["Unnamed: 1", "Unnamed: 3"]

def get_file_hash(filepath):
    """Dosyanın SHA256 parmak izini hesaplar."""
    sha256_hash = hashlib.sha256()
    with open(filepath, "rb") as f:
        for byte_block in iter(lambda: f.read(4096), b""):
            sha256_hash.update(byte_block)
    return sha256_hash.hexdigest()

def download_file():
    print(f"Downloading new file from {URL}...")
    try:
        response = requests.get(URL, timeout=60)
        response.raise_for_status()
        content_type = response.headers.get('Content-Type', '')
        if 'text/html' in content_type:
            raise ValueError("Indirilen dosya Excel degil, HTML sayfasi.")
        with open(LATEST_FILE, 'wb') as f:
            f.write(response.content)
        print("Download complete.")
    except Exception as e:
        print(f"CRITICAL ERROR downloading file: {e}")
        exit(1)

def load_all_sheets(filepath):
    if not os.path.exists(filepath):
        return {}
    try:
        xls_dict = pd.read_excel(filepath, sheet_name=None, dtype=str, engine='openpyxl')
        clean_dict = {}
        for sheet_name, df in xls_dict.items():
            df = df.fillna("")
            df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
            clean_dict[sheet_name] = df
        return clean_dict
    except Exception as e:
        print(f"Error reading Excel file {filepath}: {e}")
        exit(1)

def compare_dataframes(df_latest, df_previous, sheet_name):
    # Kolon Hizalama
    if df_latest.empty and not df_previous.empty:
        all_cols = df_previous.columns
    elif df_previous.empty and not df_latest.empty:
        all_cols = df_latest.columns
    else:
        all_cols = sorted(set(df_latest.columns) | set(df_previous.columns))
    
    df_latest = df_latest.reindex(columns=all_cols, fill_value="")
    df_previous = df_previous.reindex(columns=all_cols, fill_value="")

    # Counter Mantığı (Satır Bazlı Fark)
    latest_rows = [tuple(r) for r in df_latest.to_numpy()]
    prev_rows = [tuple(r) for r in df_previous.to_numpy()]

    c_latest = Counter(latest_rows)
    c_prev = Counter(prev_rows)

    additions_rows = list((c_latest - c_prev).elements())
    deletions_rows = list((c_prev - c_latest).elements())

    add_df = pd.DataFrame(additions_rows, columns=all_cols)
    del_df = pd.DataFrame(deletions_rows, columns=all_cols)

    if not add_df.empty:
        add_df["_SHEET_SOURCE"] = sheet_name
    if not del_df.empty:
        del_df["_SHEET_SOURCE"] = sheet_name

    return add_df, del_df

def detect_updates(additions_list, deletions_list):
    """
    Eklenenler ve Silinenler listesini karşılaştırır.
    Eğer 'Anahtar Kolonlar' (Borsa Kodu + İsim) aynıysa, bunu UPDATE olarak işaretler.
    """
    updates = []
    final_additions = []
    final_deletions = deletions_list.copy() # Üzerinde oynama yapacağız

    for item_add in additions_list:
        match_found = False
        
        # Bu eklenen satırın "Anahtar" (Key) değerini oluştur
        # Örn: "XPAR_Euronext Milling Wheat"
        key_val_add = "_".join([str(item_add.get(k, "")) for k in KEY_COLUMNS])

        # Silinenler listesinde bu anahtara sahip biri var mı?
        for item_del in final_deletions:
            key_val_del = "_".join([str(item_del.get(k, "")) for k in KEY_COLUMNS])
            
            # Eğer Sheet aynıysa ve Anahtarlar (ID) tutuyorsa -> GÜNCELLEME
            if item_add.get("_SHEET_SOURCE") == item_del.get("_SHEET_SOURCE") and key_val_add == key_val_del:
                
                # Değişen kolonları bul
                diffs = {}
                # Kimlik için 1-2 bilgi ekle
                diffs["_Identity_MIC"] = item_add.get("Unnamed: 1")
                diffs["_Identity_Instrument"] = item_add.get("Unnamed: 3")
                diffs["_Sheet"] = item_add.get("_SHEET_SOURCE")
                
                changes_found = False
                for key in item_add.keys():
                    if key == "_SHEET_SOURCE": continue
                    
                    val_new = item_add.get(key)
                    val_old = item_del.get(key)
                    
                    if val_new != val_old:
                        # Değişikliği kaydet: "50000 -> 120000"
                        diffs[key] = f"{val_old} -> {val_new}"
                        changes_found = True
                
                if changes_found:
                    updates.append(diffs)
                    # Eşleşen silinen satırı listeden çıkar (artık deleted değil updated oldu)
                    final_deletions.remove(item_del)
                    match_found = True
                    break # Diğerlerine bakmaya gerek yok
        
        if not match_found:
            final_additions.append(item_add)

    return final_additions, final_deletions, updates

def generate_diff():
    download_file()

    if not os.path.exists(PREVIOUS_FILE):
        print("First run detected. Setting current download as previous.")
        shutil.copy(LATEST_FILE, PREVIOUS_FILE)
        return

    print("Performing SHA-256 check...")
    if get_file_hash(LATEST_FILE) == get_file_hash(PREVIOUS_FILE):
        print("SHA-256 MATCH: Files are binary identical. Stopping.")
        if os.path.exists(LATEST_FILE): os.remove(LATEST_FILE)
        return
    
    print("SHA-256 MISMATCH. Processing sheets...")
    dict_latest = load_all_sheets(LATEST_FILE)
    dict_previous = load_all_sheets(PREVIOUS_FILE)

    all_sheet_names = set(dict_latest.keys()) | set(dict_previous.keys())
    
    raw_additions = []
    raw_deletions = []

    for sheet in all_sheet_names:
        df_lat = dict_latest.get(sheet, pd.DataFrame())
        df_prev = dict_previous.get(sheet, pd.DataFrame())

        if df_lat.empty and df_prev.empty: continue
            
        add_df, del_df = compare_dataframes(df_lat, df_prev, sheet)
        
        if not add_df.empty: raw_additions.extend(add_df.to_dict(orient='records'))
        if not del_df.empty: raw_deletions.extend(del_df.to_dict(orient='records'))

    # --- AKILLI KARŞILAŞTIRMA (UPDATES) ---
    print("Detecting specific updates...")
    final_additions, final_deletions, final_updates = detect_updates(raw_additions, raw_deletions)

    if not final_additions and not final_deletions and not final_updates:
        print("Result: Content identical (Metadata change only).")
        shutil.move(LATEST_FILE, PREVIOUS_FILE)
    else:
        print(f"CHANGES: Additions: {len(final_additions)}, Deletions: {len(final_deletions)}, Updates: {len(final_updates)}")
        
        date_str = datetime.now().strftime("%d.%m.%Y")
        json_filename = f"previousVSlatest-{date_str}.json"
        
        output_data = {
            "date": date_str,
            "summary": {
                "status": "changes_found",
                "additions_count": len(final_additions),
                "deletions_count": len(final_deletions),
                "updates_count": len(final_updates)
            },
            "updates": final_updates,     # Sadece değişen hücreleri gösterir
            "additions": final_additions, # Yepyeni satırlar
            "deletions": final_deletions  # Tamamen silinen satırlar
        }

        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=4, ensure_ascii=False)
            print(f"Diff file created: {json_filename}")

        shutil.move(LATEST_FILE, PREVIOUS_FILE)
        print(f"Updated {PREVIOUS_FILE}.")

if __name__ == "__main__":
    generate_diff()
