import requests
import pandas as pd
import os
import shutil
import json
from datetime import datetime

# --- AYARLAR ---
URL = "https://www.esma.europa.eu/sites/default/files/position_limits_publication.xlsx"
LATEST_FILE = "positionlimit-latest.xlsx"
PREVIOUS_FILE = "positionlimit-previous.xlsx"

def download_file():
    print(f"Downloading new file from {URL}...")
    try:
        response = requests.get(URL, timeout=30)
        response.raise_for_status() # Hata varsa (404 vs) programı durdur
        with open(LATEST_FILE, 'wb') as f:
            f.write(response.content)
        print("Download complete.")
    except Exception as e:
        print(f"Error downloading file: {e}")
        exit(1)

def load_data(filepath):
    """
    Excel dosyasını yükler, her şeyi string'e çevirir ve 
    baştaki/sondaki görünmez boşlukları (whitespace) temizler.
    """
    if not os.path.exists(filepath):
        return None
    
    # Tüm verileri 'str' (yazı) olarak oku ki 100 ile "100" farkı olmasın
    df = pd.read_excel(filepath, dtype=str)
    
    # Boş (NaN) değerleri boş string yap
    df = df.fillna("")
    
    # Tüm hücrelerdeki gereksiz boşlukları (trim) temizle
    # Örneğin: "  DE000... " -> "DE000..." olur
    df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
    
    return df

def generate_diff():
    # 1. Yeni dosyayı indir (DİKKAT: GitHub'daki latest dosyasının üzerine yazar)
    download_file()

    # 2. Previous (Eski) dosya var mı kontrol et
    if not os.path.exists(PREVIOUS_FILE):
        print("ALERT: No previous file found. Setting current download as previous for the NEXT run.")
        shutil.copy(LATEST_FILE, PREVIOUS_FILE)
        return

    # 3. Verileri Yükle
    print("Loading datasets...")
    df_latest = load_data(LATEST_FILE)
    df_previous = load_data(PREVIOUS_FILE)

    # 4. Karşılaştırma Yap
    # indicator=True bize satırın hangi dosyada olduğunu söyler
    # left_only = Sadece eskide var (SİLİNMİŞ)
    # right_only = Sadece yenide var (EKLENMİŞ)
    merged = df_previous.merge(df_latest, how='outer', indicator=True)

    deletions_df = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])
    additions_df = merged[merged['_merge'] == 'right_only'].drop(columns=['_merge'])

    # 5. Sonuçları Değerlendir
    if deletions_df.empty and additions_df.empty:
        print("No changes detected. Files are content-wise identical.")
        # Değişiklik yoksa, yeni inen latest dosyasını silebiliriz, kirlilik yapmasın.
        if os.path.exists(LATEST_FILE):
            os.remove(LATEST_FILE)
    else:
        print(f"CHANGES DETECTED! Additions: {len(additions_df)}, Deletions: {len(deletions_df)}")
        
        # JSON Dosyası Hazırla
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

        # JSON Kaydet
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=4, ensure_ascii=False)
            print(f"Diff file created: {json_filename}")

        # ÖNEMLİ: Yeni dosyayı, bir sonraki günün 'eski' dosyası olması için güncelle
        shutil.move(LATEST_FILE, PREVIOUS_FILE)
        print(f"Updated {PREVIOUS_FILE} with the new content for tomorrow's comparison.")

if __name__ == "__main__":
    generate_diff()
