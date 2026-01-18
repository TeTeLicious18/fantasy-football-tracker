# Fantasy Football Tracker

A tool to track football match predictions among friends. Supports automatic fixture fetching and score updates.

## Features

- Multi-matchday Excel sheets with automatic point calculation
- Configurable league and matchday range
- Auto-fetch scores from livescore.bz
- Leaderboard with cumulative totals

## Scoring System

| Result | Points |
|--------|--------|
| Exact score (e.g., predicted 2-1, actual 2-1) | 3 |
| Correct outcome (winner or draw) | 1 |
| Wrong | 0 |

## Quick Start

### 1. Configure your league

Edit `config.json`:

```json
{
  "league": {
    "name": "Superliga Romania",
    "country": "Romania",
    "flashscore_id": "romania/superliga"
  },
  "matchdays": {
    "start": 23,
    "end": 28
  },
  "players": ["Player1", "Player2", "Player3", "Player4"]
}
```

**Supported leagues:**
- `romania/superliga`
- `england/premier-league`
- `spain/la-liga`
- `germany/bundesliga`
- `italy/serie-a`
- `france/ligue-1`

### 2. (Optional) Fetch fixtures

```bash
python fetch_fixtures.py
```

This creates `fixtures.json` with match data. Edit this file to add specific matches with dates.

### 3. Generate Excel

```bash
pip install openpyxl xlsxwriter requests beautifulsoup4
python generate_excel.py
```

Creates `predictions.xlsx` with:
- One sheet per matchday
- Date/time field for each match
- Prediction columns for each player
- Automatic point formulas
- Leaderboard sheet

### 4. Fill predictions

Open `predictions.xlsx` and fill in your predictions in columns B (Pred Home) and C (Pred Away).

### 5. Update Scores

Run anytime to fetch latest scores:

```bash
python update_scores.py
```

Or double-click `run_update.bat` (Windows)

## Files

| File | Description |
|------|-------------|
| `config.json` | League and player configuration |
| `fixtures.json` | Match fixtures with dates (auto-generated or manual) |
| `fetch_fixtures.py` | Fetches fixtures from web |
| `generate_excel.py` | Creates the Excel file |
| `update_scores.py` | Fetches scores from web |
| `run_update.bat` | Windows shortcut |
| `predictions.xlsx` | Generated spreadsheet |

## Adding Custom Fixtures

Edit `fixtures.json`:

```json
{
  "23": [
    {"home": "Team A", "away": "Team B", "date": "25.01.2026 20:00"},
    {"home": "Team C", "away": "Team D", "date": "25.01.2026 17:00"}
  ],
  "24": [
    {"home": "Team E", "away": "Team F", "date": "01.02.2026 15:00"}
  ]
}
```

Then run `python generate_excel.py` to regenerate the Excel.

## Requirements

- Python 3.8+
- Libraries: `openpyxl`, `requests`, `beautifulsoup4`

## License

MIT
