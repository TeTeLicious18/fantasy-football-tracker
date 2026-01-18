import requests
from bs4 import BeautifulSoup
import json
import re
import time

def load_config():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

def fetch_fixtures_from_web(league_id, matchday_start, matchday_end):
    """
    Fetch fixtures from livescore.bz for a specific league.
    Returns dict: {matchday: [{"home": ..., "away": ..., "date": ...}, ...]}
    """
    print(f"Fetching fixtures for {league_id}...")
    
    # Map common leagues to livescore.bz IDs
    LEAGUE_IDS = {
        "romania/superliga": 131,
        "romania/liga-1": 131,
        "england/premier-league": 2,
        "spain/la-liga": 8,
        "germany/bundesliga": 9,
        "italy/serie-a": 10,
        "france/ligue-1": 11,
    }
    
    lid = LEAGUE_IDS.get(league_id.lower(), 131)
    
    urls = [
        f"https://www.livescore.bz/en/football/league/{lid}/fixture/",
        f"https://www.livescore.bz/en/football/league/{lid}/results/"
    ]
    
    all_text = []
    headers = {"User-Agent": "Mozilla/5.0"}
    
    for url in urls:
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, 'html.parser')
                for elem in soup.find_all(['a', 'div', 'span']):
                    text = elem.get_text(" ", strip=True)
                    if text:
                        all_text.append(text)
            time.sleep(1)
        except Exception as e:
            print(f"Error fetching {url}: {e}")
    
    print(f"  Found {len(all_text)} text elements")
    
    # Parse into matchdays (simplified - returns placeholder structure)
    # In production, you would parse the actual HTML structure
    matchdays = {}
    
    for md in range(matchday_start, matchday_end + 1):
        matchdays[md] = []
        print(f"  Note: Matchday {md} - add fixtures manually or enhance parser")
    
    return matchdays

def search_match_score(home, away, league_id=None):
    """Search for a specific match score"""
    print(f"Searching score for: {home} vs {away}")
    
    urls = [
        "https://www.livescore.bz/",
    ]
    
    if league_id:
        LEAGUE_IDS = {
            "romania/superliga": 131,
            "england/premier-league": 2,
        }
        lid = LEAGUE_IDS.get(league_id.lower(), 131)
        urls.append(f"https://www.livescore.bz/en/football/league/{lid}/results/")
    
    headers = {"User-Agent": "Mozilla/5.0"}
    
    for url in urls:
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            if resp.status_code == 200:
                soup = BeautifulSoup(resp.text, 'html.parser')
                for a in soup.find_all('a', href=True):
                    text = a.get_text(" ", strip=True).lower()
                    if home.lower()[:4] in text and away.lower()[:4] in text:
                        # Try to extract score
                        match = re.search(r'(\d+)\s*-\s*(\d+)', a.get_text())
                        if match:
                            return int(match.group(1)), int(match.group(2))
            time.sleep(0.5)
        except Exception as e:
            print(f"Error: {e}")
    
    return None, None

def save_fixtures(matchdays, filename='fixtures.json'):
    """Save fetched fixtures to JSON file"""
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(matchdays, f, indent=2, ensure_ascii=False)
    print(f"Fixtures saved to {filename}")

def load_fixtures(filename='fixtures.json'):
    """Load fixtures from JSON file"""
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

if __name__ == "__main__":
    config = load_config()
    league = config['league']['flashscore_id']
    start = config['matchdays']['start']
    end = config['matchdays']['end']
    
    print("=" * 50)
    print("FIXTURE FETCHER")
    print("=" * 50)
    print(f"League: {config['league']['name']}")
    print(f"Matchdays: {start} to {end}")
    print()
    
    fixtures = fetch_fixtures_from_web(league, start, end)
    save_fixtures(fixtures)
    
    print("\nTo add specific fixtures, edit fixtures.json manually or")
    print("use the web interface to input match details.")
