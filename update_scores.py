import openpyxl
import requests
from bs4 import BeautifulSoup
import json
import re
import time

def load_config():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def get_scores_from_web():
    """Fetch all match data from livescore.bz"""
    print("Fetching data from livescore.bz...")
    
    urls = [
        "https://www.livescore.bz/",
        "https://www.livescore.bz/en/football/league/131/results/",
        "https://www.livescore.bz/en/football/league/398/results/"
    ]
    
    all_matches = []
    headers = {"User-Agent": "Mozilla/5.0"}
    
    for url in urls:
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, 'html.parser')
                for a in soup.find_all('a', href=True):
                    if 'event' in a['href'] or 'match' in a['href']:
                        all_matches.append(a.get_text(" ", strip=True))
            time.sleep(0.5)
        except Exception as e:
            print(f"Error fetching {url}: {e}")
    
    print(f"Found {len(all_matches)} match entries")
    return all_matches

def parse_score(text):
    """Extract score from text like '1 - 2' or '1-2'"""
    match = re.search(r'(\d+)\s*-\s*(\d+)', text)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None, None

def check_keywords(keywords, text):
    """Check if any keyword is in text"""
    text_lower = text.lower()
    if isinstance(keywords, list):
        return any(k.lower() in text_lower for k in keywords)
    return keywords.lower() in text_lower

def find_match_score(match_config, web_data):
    """Find score for a match using its keywords"""
    home_kw = match_config.get('home_keywords', [match_config['home'][:5]])
    away_kw = match_config.get('away_keywords', [match_config['away'][:5]])
    
    for line in web_data:
        if check_keywords(home_kw, line) and check_keywords(away_kw, line):
            return parse_score(line)
    return None, None

def update_scores():
    config = load_config()
    matches = config['matches']
    
    file_path = 'predictions.xlsx'
    
    try:
        open(file_path, 'a+').close()
    except PermissionError:
        print("[ERROR] Excel file is open! Close it and try again.")
        input("Press Enter to exit...")
        return
    
    web_data = get_scores_from_web()
    
    try:
        wb = openpyxl.load_workbook(file_path)
        sheet = wb.active
        updates = 0
        
        print("\nUpdating scores...\n")
        
        for row in sheet.iter_rows(min_row=1, max_col=6, values_only=False):
            if row[3].value == "Real Score:":
                title = row[0].value
                if title and " vs " in title:
                    home, away = title.split(" vs ")
                    
                    # Find matching config
                    match_cfg = None
                    for m in matches:
                        if m['home'] == home and m['away'] == away:
                            match_cfg = m
                            break
                    
                    if not match_cfg:
                        match_cfg = {'home': home, 'away': away}
                    
                    score_h, score_a = find_match_score(match_cfg, web_data)
                    
                    if score_h is not None:
                        print(f"[+] {title}: {score_h}-{score_a}")
                        row[4].value = score_h
                        row[5].value = score_a
                        updates += 1
                    else:
                        if row[4].value is not None:
                            print(f"[=] {title}: keeping ({row[4].value}-{row[5].value})")
                        else:
                            print(f"[-] {title}: not found")
        
        wb.save(file_path)
        print(f"\nDone! {updates} updates made.")
        
    except Exception as e:
        print(f"Error: {e}")
    
    input("\nPress Enter to close...")

if __name__ == "__main__":
    update_scores()
