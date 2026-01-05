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
        response = requests.get(url, verify=False)
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

def detailed_compare(df_old, df_new, sheet_name):
    """
    İki tabloyu hücre bazında kıyaslar ve detaylı rapor döner.
    """
    changes = []
    
    # 1. Boyut Kontrolü
    if df_old.shape != df_new.shape:
        changes.append(f"[{sheet_name}] Tablo boyutu değişti! (Eski: {df_old.shape}, Yeni: {df_new.shape})")
        # Boyutlar farklıysa hücre hücre kıyaslama zordur, sadece satır farkını raporlarız.
        diff_rows = len(df_new) - len(df_old)
        if diff_rows > 0:
            changes.append(f"[{sheet_name}] {diff_rows} adet yeni satır eklendi.")
        elif diff_rows < 0:
            changes.append(f"[{sheet_name}] {abs(diff_rows)} adet satır silindi.")
        return changes

    # 2. Hücre Hücre Kıyaslama (Boyutlar aynıysa)
    # Pandas 'compare' fonksiyonu farkları yan yana koyar
    try:
        diff = df_old.compare(df_new)
        if not diff.empty:
            # Çok fazla detay olmaması için ilk 10 değişikliği alalım (veya hepsini)
            count = 0
            for index, row in diff.iterrows():
                for col in diff.columns.levels[0]: # Sütun isimleri
                    old_val = row[(col, 'self')]
                    new_val = row[(col, 'other')]
                    
                    # NaN kontrolü (Pandas bazen boşlukları NaN görür)
                    if pd.isna(old_val) and pd.isna(new_val):
                        continue
                        
                    changes.append(f"[{sheet_name}] Satır {index+2}, '{col}': {old_val} -> {new_val}")
                    count += 1
                    if count >= 20: # Güvenlik: Her sayfa için en fazla 20 detay yaz (Log şişmesin)
                        changes.append(f"[{sheet_name}] ... ve diğer değişiklikler.")
                        return changes
    except Exception as e:
        changes.append(f"[{sheet_name}] Karşılaştırma hatası: {str(e)}")

    return changes

def check_updates():
    history = load_history()
    changes_detected = False
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for region, config in FILES.items():
        print(f"--- {region} Kontrol Ediliyor ---")
        temp_file = f"temp_{region}.xlsx"
        target_file = config["local_file"]
        
        if not download_file(config["url"], temp_file):
            print(f"{region} indirilemedi.")
            continue

        # TÜM SEKMELERİ OKU
        try:
            dfs_new = pd.read_excel(temp_file, sheet_name=None)
        except Exception as e:
            print(f"Excel okuma hatası: {e}")
            continue

        if os.path.exists(target_file):
            dfs_old = pd.read_excel(target_file, sheet_name=None)
            
            all_changes_log = [] # Bu bölge için tüm değişiklikleri burada toplayacağız
            
            # A. Sekme İsimlerini Kontrol Et
            old_sheets = set(dfs_old.keys())
            new_sheets = set(dfs_new.keys())
            
            added_sheets = new_sheets - old_sheets
            removed_sheets = old_sheets - new_sheets
            common_sheets = old_sheets & new_sheets

            if added_sheets:
                all_changes_log.append(f"Yeni sekmeler eklendi: {', '.join(added_sheets)}")
            if removed_sheets:
                all_changes_log.append(f"Sekmeler silindi: {', '.join(removed_sheets)}")

            # B. Ortak Sekmeleri Tek Tek Tara
            for sheet in common_sheets:
                # BREAK YOK! Hepsini tara.
                sheet_changes = detailed_compare(dfs_old[sheet], dfs_new[sheet], sheet)
                all_changes_log.extend(sheet_changes)

            # Değişiklik varsa kaydet
            if all_changes_log:
                print(f"⚠️ {region} için değişiklikler tespit edildi!")
                
                # Liste halindeki mesajları HTML'de alt alta göstermek için <br> ile birleştir
                formatted_message = "<br>".join(all_changes_log)
                
                history.append({
                    "date": current_date,
                    "region": region,
                    "type": "UPDATE",
                    "message": formatted_message
                })
                changes_detected = True
                shutil.move(temp_file, target_file)
            else:
                print(f"{region} güncel.")
                os.remove(temp_file)
        else:
            print(f"{region} ilk kurulum.")
            history.append({
                "date": current_date,
                "region": region,
                "type": "INIT",
                "message": "İlk dosya indirildi ve takibe başlandı."
            })
            changes_detected = True
            shutil.move(temp_file, target_file)

    if changes_detected:
        save_history(history)
        print(">> report_list.json güncellendi.")

if __name__ == "__main__":
    check_updates()
