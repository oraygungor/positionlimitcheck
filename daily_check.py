import requests
import pandas as pd
import os
import shutil
import json
from datetime import datetime
from collections import Counter

# --- AYARLAR ---
URL = "https://www.esma.europa.eu/sites/default/files/position_limits_publication.xlsx"
LATEST_FILE = "positionlimit-latest.xlsx"
PREVIOUS_FILE = "positionlimit-previous.xlsx"

def download_file():
    print(f"Downloading new file from {URL}...")
    try:
        response = requests.get(URL, timeout=60)
        response.raise_for_status()
        
        # Basit İçerik Kontrolü: İnen şey bir Excel dosyası mı?
        # Excel dosyaları genellikle 'application/vnd.openxmlformats' veya binary başlar.
        # En basit kontrol: HTML inip inmediğine bakmak.
        content_type = response.headers.get('Content-Type', '')
        if 'text/html' in content_type:
            raise ValueError("Indirilen dosya Excel degil, HTML sayfasi (Muhtemelen hata sayfasi).")
            
        with open(LATEST_FILE, 'wb') as f:
            f.write(response.content)
        print("Download complete.")
        
    except Exception as e:
        print(f"CRITICAL ERROR downloading file: {e}")
        exit(1)

def load_data(filepath):
    """
    Excel dosyasını yükler, temizler ve dataframe döndürür.
    """
    if not os.path.exists(filepath):
        return None
    
    try:
        # Tüm verileri 'str' oku, sheet belirtilmezse ilk sheet okunur.
        df = pd.read_excel(filepath, dtype=str, engine='openpyxl')
        
        # Boşlukları ve NaN'ları temizle
        df = df.fillna("")
        # Baştaki/sondaki görünmez boşlukları (trim) temizle
        df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
        
        return df
    except Exception as e:
        print(f"Error reading Excel file {filepath}: {e}")
        # Eğer dosya bozuksa scriptin durması daha iyidir
        exit(1)

def generate_diff():
    # 1. İndir
    download_file()

    # 2. Previous kontrolü
    if not os.path.exists(PREVIOUS_FILE):
        print("ALERT: No previous file found. Setting current download as previous for the NEXT run.")
        shutil.copy(LATEST_FILE, PREVIOUS_FILE)
        return

    print("Loading datasets...")
    df_latest = load_data(LATEST_FILE)
    df_previous = load_data(PREVIOUS_FILE)

    # --- KRİTİK DÜZELTME 1: Kolon Hizalama ---
    # İki dosyada kolonlar yer değiştirmiş veya yeni kolon gelmiş olabilir.
    # İkisini de tüm kolonların birleşimine (union) göre genişletiyoruz.
    all_cols = sorted(set(df_latest.columns) | set(df_previous.columns))
    
    df_latest = df_latest.reindex(columns=all_cols, fill_value="")
    df_previous = df_previous.reindex(columns=all_cols, fill_value="")

    # --- KRİTİK DÜZELTME 2: Counter ile Multiset Farkı ---
    # Pandas merge yerine, satırları tuple'a çevirip sayıyoruz.
    # Bu yöntem duplicate satırları ve kolon kaymalarını %100 doğru yönetir.
    
    print("Comparing rows using Counter logic...")
    # DataFrame -> Numpy -> List of Tuples
    latest_rows = [tuple(r) for r in df_latest.to_numpy()]
    prev_rows = [tuple(r) for r in df_previous.to_numpy()]

    # Satırları say (Hangi satırdan kaç tane var?)
    c_latest = Counter(latest_rows)
    c_prev = Counter(prev_rows)

    # Farkları bul (Multiset subtraction)
    # c_latest - c_prev -> Latest'ta olup Previous'ta olmayanlar (EKLENENLER)
    # c_prev - c_latest -> Previous'ta olup Latest'ta olmayanlar (SİLİNENLER)
    additions_rows = list((c_latest - c_prev).elements())
    deletions_rows = list((c_prev - c_latest).elements())

    # Tekrar DataFrame'e çevir
    additions_df = pd.DataFrame(additions_rows, columns=all_cols)
    deletions_df = pd.DataFrame(deletions_rows, columns=all_cols)

    # 3. Sonuç Kontrolü
    if deletions_df.empty and additions_df.empty:
        print("No content changes detected.")
        if os.path.exists(LATEST_FILE):
            os.remove(LATEST_FILE)
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

        # Update previous file
        shutil.move(LATEST_FILE, PREVIOUS_FILE)
        print(f"Updated {PREVIOUS_FILE} for the next run.")

if __name__ == "__main__":
    generate_diff()
