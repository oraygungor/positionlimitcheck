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
You are a senior compliance officer reviewing exchange position limit files for an energy trading company.

Task:
Compare the OLD and NEW versions of the {region} Position Limits file using ONLY the text provided below. Do not rely on outside knowledge. Do not infer missing values. Do not guess.

Your objectives:
1. Identify every contract/row that was added in the NEW version.
2. Identify every contract/row that was removed compared with the OLD version.
3. Identify every field/value that changed for contracts/rows that exist in both versions.
4. Assess whether any of the identified changes could affect a pure energy trading company, and explain why in practical compliance/trading terms.

Instructions:
- Use only the supplied OLD and NEW text.
- If a point cannot be verified from the text, say: "Not verifiable from the provided text."
- Do not invent row names, contract names, values, or business impacts.
- If the files appear identical in substance, state that clearly.
- If formatting/layout changed but substance did not, separate that from substantive changes.
- Treat the output as a controlled compliance comparison, so be precise and conservative.
- Keep factual comparison separate from interpretation.
- Where possible, quote the exact old value and new value.
- If the text quality is incomplete, messy, or ambiguous, flag that explicitly.
- Do not summarize vaguely. Be exhaustive.

Output format:

A. Executive Summary
- 3-6 bullet points summarizing the key changes.
- State clearly whether the changes are likely material, potentially material, or not material for a pure energy trading company.

B. Added Rows / Contracts
- List each newly added row/contract.
- For each, include the contract name and the relevant values visible in the NEW version.

C. Deleted Rows / Contracts
- List each deleted row/contract.
- For each, include the contract name and the relevant values visible in the OLD version.

D. Changed Values
For each row/contract present in both versions but changed:
- Contract / row name
- Field name
- OLD value
- NEW value
- Short description of the change

E. Compliance / Trading Impact Assessment
- Explain how the changes could affect an energy trading company.
- Focus on position monitoring, pre-trade controls, reporting, exchange limit usage, and operational risk.
- If the changes do not materially affect a pure energy trading company, state exactly:
  "These are the identified changes, but they are not expected to materially affect a pure energy trading company."
- If impact depends on trading activity in specific contracts/venues, say so explicitly.

F. Data Quality / Limitations
- Note any comparison limitations caused by missing text, OCR issues, broken tables, inconsistent naming, or formatting problems.

Important:
Do not provide generic commentary. Only discuss changes that can be tied back to the provided text.

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
