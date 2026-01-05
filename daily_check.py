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
    """Dosyanın SHA256 parmak izini (hash) hesaplar."""
    sha256_hash = hashlib.sha256()
    with open(filepath, "rb") as f:
        # Dosyayı parça parça oku (Büyük dosyalar için RAM dostu)
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

def load_data(filepath):
    if not os.path.exists(filepath):
        return None
    try:
        df = pd.read_excel(filepath, dtype=str, engine='openpyxl')
        df = df.fillna("")
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        return df
    except Exception as e:
        print(f"Error reading Excel file {filepath}: {e}")
        exit(1)

def generate_diff():
    # 1. Yeni dosyayı indir
    download_file()

    # 2. Previous dosya yoksa, yeniyi previous yap ve bitir.
    if not os.path.exists(PREVIOUS_FILE):
        print("First run detected. Setting current download as previous.")
        shutil.copy(LATEST_FILE, PREVIOUS_FILE)
        return

    # --- ADIM 3: SHA-256 KONTROLÜ (Hızlı Ön Eleme) ---
    print("Performing SHA-256 binary check...")
    latest_hash = get_file_hash(LATEST_FILE)
    previous_hash = get_file_hash(PREVIOUS_FILE)

    if latest_hash == previous_hash:
        print("SHA-256 MATCH: Files are binary identical.")
        print("No need to parse Excel. Stopping.")
        # Dosyalar birebir aynı, işleme gerek yok.
        if os.path.exists(LATEST_FILE):
            os.remove(LATEST_FILE)
        return
    else:
        print("SHA-256 MISMATCH: Binary content differs.")
        print("Starting deep content comparison (Pandas)...")
        # Hash farklı ama içerik aynı olabilir (Metadata değişmiştir).
        # O yüzden devam ediyoruz...

    # --- ADIM 4: PANDAS & COUNTER ANALİZİ ---
    print("Loading datasets...")
    df_latest = load_data(LATEST_FILE)
    df_previous = load_data(PREVIOUS_FILE)

    # Kolon Hizalama
    all_cols = sorted(set(df_latest.columns) | set(df_previous.columns))
    df_latest = df_latest.reindex(columns=all_cols, fill_value="")
    df_previous = df_previous.reindex(columns=all_cols, fill_value="")

    # Counter ile Satır Karşılaştırma
    latest_rows = [tuple(r) for r in df_latest.to_numpy()]
    prev_rows = [tuple(r) for r in df_previous.to_numpy()]

    c_latest = Counter(latest_rows)
    c_prev = Counter(prev_rows)

    additions_rows = list((c_latest - c_prev).elements())
    deletions_rows = list((c_prev - c_latest).elements())

    additions_df = pd.DataFrame(additions_rows, columns=all_cols)
    deletions_df = pd.DataFrame(deletions_rows, columns=all_cols)

    # 5. Sonuç
    if deletions_df.empty and additions_df.empty:
        print("Result: Content is identical (Only metadata/timestamps changed).")
        # İçerik aynı çıktı, ama dosya hash'i farklıydı.
        # Yine de en son inen dosyayı 'previous' yapalım ki hash'ler güncellensin.
        shutil.move(LATEST_FILE, PREVIOUS_FILE)
        print("Updated previous file to match latest metadata.")
    else:
        print(f"CHANGES DETECTED! Additions: {len(additions_df)}, Deletions: {len(deletions_df)}")
        
        date_str = datetime.now().strftime("%d.%m.%Y")
        json_filename = f"previousVSlatest-{date_str}.json"
        
        output_data = {
            "date": date_str,
            "summary": {
                "status": "changes_found",
                "additions_count": len(additions_df),
                "deletions_count": len(deletions_df)
            },
            "additions": additions_df.to_dict(orient='records'),
            "deletions": deletions_df.to_dict(orient='records')
        }

        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=4, ensure_ascii=False)
            print(f"Diff file created: {json_filename}")

        shutil.move(LATEST_FILE, PREVIOUS_FILE)
        print(f"Updated {PREVIOUS_FILE} for the next run.")

if __name__ == "__main__":
    generate_diff()
