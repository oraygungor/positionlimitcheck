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
    """
    Excel dosyasındaki TÜM sheet'leri okur.
    Geriye { 'Sheet1': DataFrame, 'Sheet2': DataFrame } şeklinde bir sözlük döndürür.
    """
    if not os.path.exists(filepath):
        return {}
    
    try:
        # sheet_name=None -> Tüm sayfaları okur
        xls_dict = pd.read_excel(filepath, sheet_name=None, dtype=str, engine='openpyxl')
        
        clean_dict = {}
        
        for sheet_name, df in xls_dict.items():
            # Temizlik İşlemleri (NaN temizle, Trim yap)
            df = df.fillna("")
            df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
            clean_dict[sheet_name] = df
            
        return clean_dict

    except Exception as e:
        print(f"Error reading Excel file {filepath}: {e}")
        exit(1)

def compare_dataframes(df_latest, df_previous, sheet_name):
    """
    İki DataFrame'i karşılaştırır ve eklenen/silinen satırları döndürür.
    """
    # Kolon Hizalama (Eğer sheet boşsa kolonları diğerinden al)
    if df_latest.empty and not df_previous.empty:
        all_cols = df_previous.columns
    elif df_previous.empty and not df_latest.empty:
        all_cols = df_latest.columns
    else:
        all_cols = sorted(set(df_latest.columns) | set(df_previous.columns))
    
    df_latest = df_latest.reindex(columns=all_cols, fill_value="")
    df_previous = df_previous.reindex(columns=all_cols, fill_value="")

    # Counter Mantığı
    latest_rows = [tuple(r) for r in df_latest.to_numpy()]
    prev_rows = [tuple(r) for r in df_previous.to_numpy()]

    c_latest = Counter(latest_rows)
    c_prev = Counter(prev_rows)

    additions_rows = list((c_latest - c_prev).elements())
    deletions_rows = list((c_prev - c_latest).elements())

    # Sonuçları DataFrame yap ve Sheet ismini ekle
    add_df = pd.DataFrame(additions_rows, columns=all_cols)
    del_df = pd.DataFrame(deletions_rows, columns=all_cols)

    # Hangi sheet'te olduğunu bilmek için kolon ekle (JSON çıktısı için)
    if not add_df.empty:
        add_df["_SHEET_SOURCE"] = sheet_name
    if not del_df.empty:
        del_df["_SHEET_SOURCE"] = sheet_name

    return add_df, del_df

def generate_diff():
    # 1. İndir
    download_file()

    # 2. Previous kontrolü
    if not os.path.exists(PREVIOUS_FILE):
        print("First run detected. Setting current download as previous.")
        shutil.copy(LATEST_FILE, PREVIOUS_FILE)
        return

    # 3. SHA Kontrolü
    print("Performing SHA-256 check...")
    if get_file_hash(LATEST_FILE) == get_file_hash(PREVIOUS_FILE):
        print("SHA-256 MATCH: Files are binary identical. Stopping.")
        if os.path.exists(LATEST_FILE):
            os.remove(LATEST_FILE)
        return
    else:
        print("SHA-256 MISMATCH. Starting Deep Content Comparison (All Sheets)...")

    # 4. Verileri Yükle (Dictionary olarak: {'Sheet1': DF, 'Sheet2': DF})
    print("Loading all sheets...")
    dict_latest = load_all_sheets(LATEST_FILE)
    dict_previous = load_all_sheets(PREVIOUS_FILE)

    # Tüm sheet isimlerini birleştir (Yeni gelen veya silinen sheet olabilir)
    all_sheet_names = set(dict_latest.keys()) | set(dict_previous.keys())
    
    all_additions = []
    all_deletions = []

    # 5. Her Sheet İçin Döngü
    for sheet in all_sheet_names:
        # Eğer sheet dosyada yoksa boş DataFrame oluştur
        df_lat = dict_latest.get(sheet, pd.DataFrame())
        df_prev = dict_previous.get(sheet, pd.DataFrame())

        if df_lat.empty and df_prev.empty:
            continue
            
        print(f"Comparing sheet: {sheet}...")
        add_df, del_df = compare_dataframes(df_lat, df_prev, sheet)
        
        if not add_df.empty:
            all_additions.extend(add_df.to_dict(orient='records'))
        if not del_df.empty:
            all_deletions.extend(del_df.to_dict(orient='records'))

    # 6. Sonuç
    if not all_additions and not all_deletions:
        print("Result: Content is identical across all sheets.")
        shutil.move(LATEST_FILE, PREVIOUS_FILE)
    else:
        print(f"CHANGES DETECTED! Total Additions: {len(all_additions)}, Total Deletions: {len(all_deletions)}")
        
        date_str = datetime.now().strftime("%d.%m.%Y")
        json_filename = f"previousVSlatest-{date_str}.json"
        
        output_data = {
            "date": date_str,
            "summary": {
                "status": "changes_found",
                "additions_count": len(all_additions),
                "deletions_count": len(all_deletions)
            },
            "additions": all_additions,
            "deletions": all_deletions
        }

        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=4, ensure_ascii=False)
            print(f"Diff file created: {json_filename}")

        shutil.move(LATEST_FILE, PREVIOUS_FILE)
        print(f"Updated {PREVIOUS_FILE} for the next run.")

if __name__ == "__main__":
    generate_diff()
