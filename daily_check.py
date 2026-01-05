import requests
import pandas as pd
import os
import shutil
import json
from datetime import datetime

# Configuration
URL = "https://www.esma.europa.eu/sites/default/files/position_limits_publication.xlsx"
LATEST_FILE = "positionlimit-latest.xlsx"
PREVIOUS_FILE = "positionlimit-previous.xlsx"

def download_file():
    print(f"Downloading from {URL}...")
    response = requests.get(URL)
    with open(LATEST_FILE, 'wb') as f:
        f.write(response.content)
    print("Download complete.")

def load_data(filepath):
    """Loads excel and converts all data to string to ensure clean comparison"""
    if not os.path.exists(filepath):
        return None
    # Read Excel, treating all columns as strings to avoid type mismatch issues
    df = pd.read_excel(filepath, dtype=str)
    # Fill NaNs with empty string
    df = df.fillna("")
    return df

def generate_diff():
    # 1. Download the new file
    download_file()

    # 2. Check if a previous file exists
    if not os.path.exists(PREVIOUS_FILE):
        print("No previous file found. Setting current download as previous for next run.")
        shutil.copy(LATEST_FILE, PREVIOUS_FILE)
        return

    # 3. Load both files into Pandas
    df_latest = load_data(LATEST_FILE)
    df_previous = load_data(PREVIOUS_FILE)

    # 4. Compare DataFrames
    # We merge with indicator=True to find which rows are only in left (previous) or right (latest)
    merged = df_previous.merge(df_latest, how='outer', indicator=True)

    # 'left_only' implies it was in Previous but not Latest -> DELETED
    deletions_df = merged[merged['_merge'] == 'left_only'].drop(columns=['_merge'])
    
    # 'right_only' implies it was in Latest but not Previous -> ADDED
    additions_df = merged[merged['_merge'] == 'right_only'].drop(columns=['_merge'])

    # 5. Check if there are differences
    if deletions_df.empty and additions_df.empty:
        print("Files are identical. No action taken.")
        # Cleanup: remove latest since we don't need it if it's the same
        os.remove(LATEST_FILE)
    else:
        print(f"Differences found! {len(additions_df)} additions, {len(deletions_df)} deletions.")
        
        # Prepare JSON output
        date_str = datetime.now().strftime("%d.%m.%Y")
        json_filename = f"previousVSlatest-{date_str}.json"
        
        output_data = {
            "date": date_str,
            "summary": {
                "additions_count": len(additions_df),
                "deletions_count": len(deletions_df)
            },
            "additions": additions_df.to_dict(orient='records'),
            "deletions": deletions_df.to_dict(orient='records')
        }

        # Write the JSON file
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(output_data, f, indent=4, ensure_ascii=False)
            print(f"Created diff file: {json_filename}")

        # Update the 'previous' file to match the new 'latest' for the next run
        shutil.move(LATEST_FILE, PREVIOUS_FILE)
        print(f"Updated {PREVIOUS_FILE} with the new content.")

if __name__ == "__main__":
    generate_diff()
