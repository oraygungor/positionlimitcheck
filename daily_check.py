import pandas as pd
import requests
import os
import json
import shutil
from datetime import datetime

# --- AYARLAR ---
FILES = {
    "UK": {
        "url": "https://www.fca.org.uk/publication/data/position-limits-contract-names-vpc.xlsx",
        "local_file": "positionlimit-previous-UK.xlsx"
    },
    "EU": {
        # ESMA linki (Örnek)
        "url": "https://www.esma.europa.eu/sites/default/files/position_limits_publication.xlsx", 
        "local_file": "positionlimit-previous-EU.xlsx"
    }
}

REPORT_FILE = "report_list.json"

def download_file(url, filename):
    try:
        response = requests.get(url, verify=False) # Verify false SSL hatalarını bypass eder
        if response.status_code == 200:
            with open(filename, 'wb') as f:
                f.write(response.content)
            return True
    except Exception as e:
        print(f"Hata - İndirme başarısız ({url}): {e}")
    return False

def load_history():
    if os.path.exists(REPORT_FILE):
        with open(REPORT_FILE, 'r', encoding='utf-8') as f:
            try:
                return json.load(f)
            except:
                return []
    return []

def save_history(history_data):
    with open(REPORT_FILE, 'w', encoding='utf-8') as f:
        json.dump(history_data, f, ensure_ascii=False, indent=4)

def check_updates():
    history = load_history()
    changes_detected = False
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for region, config in FILES.items():
        print(f"--- {region} Kontrol Ediliyor ---")
        temp_file = f"temp_{region}.xlsx"
        target_file = config["local_file"]
        
        # 1. Yeni dosyayı indir
        if not download_file(config["url"], temp_file):
            print(f"{region} indirilemedi.")
            continue

        # 2. Karşılaştırma Mantığı
        try:
            df_new = pd.read_excel(temp_file)
        except Exception as e:
            print(f"{region} dosyası bozuk veya okunamadı: {e}")
            continue

        if os.path.exists(target_file):
            df_old = pd.read_excel(target_file)
            
            # Değişiklik var mı? (Basitçe tüm dataframeleri kıyaslar)
            if not df_new.equals(df_old):
                print(f"⚠️ {region} için YENİ veri bulundu!")
                
                # Tarihçeye EKLE (Append)
                history.append({
                    "date": current_date,
                    "region": region,
                    "type": "UPDATE",
                    "message": f"{region} limitlerinde değişiklik tespit edildi. (Satır: {len(df_old)} -> {len(df_new)})"
                })
                changes_detected = True
                shutil.move(temp_file, target_file) # Dosyayı güncelle
            else:
                print(f"{region} güncel, değişiklik yok.")
                os.remove(temp_file) # Temp sil
        else:
            print(f"{region} ilk kez indirildi.")
            history.append({
                "date": current_date,
                "region": region,
                "type": "INIT",
                "message": f"{region} takibi başlatıldı."
            })
            changes_detected = True
            shutil.move(temp_file, target_file)

    # Değişiklik varsa JSON dosyasını güncelle
    if changes_detected:
        save_history(history)
        print(">> report_list.json güncellendi.")
    else:
        print(">> Herhangi bir güncelleme yapılmadı.")

if __name__ == "__main__":
    check_updates()
