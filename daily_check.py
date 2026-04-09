import pandas as pd
import requests
import os
import json
import shutil
from datetime import datetime
from openai import OpenAI

FILES = {
    "UK": {
        "url": "https://www.fca.org.uk/publication/data/position-limits-contract-names-vpc.xlsx",
        "local_file": "positionlimit-previous-UK.xlsx"
    },
    "EU": {
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
    except:
        pass
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

def excel_to_text(file_path):
    text_output = []
    try:
        excel_data = pd.read_excel(file_path, sheet_name=None)
        for sheet_name, df in excel_data.items():
            text_output.append(f"--- SHEET: {sheet_name} ---")
            text_output.append(df.to_string(index=False))
            text_output.append("\n")
    except:
        pass
    return "\n".join(text_output)

def analyze_full_difference_with_ai(old_text, new_text, region):
    if not os.getenv("OPENAI_API_KEY"):
        return "Error: OPENAI_API_KEY environment variable not found."

    client = OpenAI()
    
    prompt = f"""
    You are a financial data analyst.
    I am providing you with the full text content of two versions (OLD and NEW) of the {region} Position Limits Excel files.
    
    Please compare them thoroughly and summarize all changes. Include:
    - Any new rows/contracts added.
    - Any rows/contracts deleted.
    - Any specific values that have changed.
    
    Provide your response as a clear, professional summary in English.
    
    === OLD VERSION ===
    {old_text}
    
    === NEW VERSION ===
    {new_text}
    """
    
    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You are a precise data analysis assistant."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.1
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error during AI analysis: {str(e)}"

def check_updates():
    history = load_history()
    changes_detected = False
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for region, config in FILES.items():
        temp_file = f"temp_{region}.xlsx"
        target_file = config["local_file"]
        
        if not download_file(config["url"], temp_file):
            continue

        if os.path.exists(target_file):
            old_text = excel_to_text(target_file)
            new_text = excel_to_text(temp_file)
            
            if old_text != new_text:
                ai_summary = analyze_full_difference_with_ai(old_text, new_text, region)
                
                history.append({
                    "date": current_date,
                    "region": region,
                    "type": "UPDATE",
                    "ai_summary": ai_summary
                })
                changes_detected = True
                shutil.move(temp_file, target_file)
            else:
                os.remove(temp_file)
        else:
            history.append({
                "date": current_date,
                "region": region,
                "type": "INIT",
                "message": "Initial file downloaded."
            })
            changes_detected = True
            shutil.move(temp_file, target_file)

    if changes_detected:
        save_history(history)

if __name__ == "__main__":
    check_updates()
