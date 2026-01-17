# Fantasy Football Tracker

A simple tool to track fantasy football match predictions among friends. Automatically fetches live scores and calculates points.

## Features

- Excel spreadsheet with automatic point calculation
- Auto-fetch scores from livescore.bz
- Total points leaderboard
- Configurable matches and players via JSON

## Scoring System

| Result | Points |
|--------|--------|
| Exact score (e.g., predicted 2-1, actual 2-1) | 3 |
| Correct outcome (winner or draw) | 1 |
| Wrong | 0 |

## Quick Start

### 1. Configure your matches

Edit `config.json`:

```json
{
  "players": ["Alice", "Bob", "Charlie"],
  "matches": [
    {
      "home": "Team A",
      "away": "Team B",
      "predictions": {
        "Alice": [2, 1],
        "Bob": [1, 1],
        "Charlie": [0, 2]
      }
    }
  ]
}
```

### 2. Generate Excel

```bash
pip install xlsxwriter openpyxl requests beautifulsoup4
python generate_excel.py
```

### 3. Update Scores

Run anytime to fetch latest scores:

```bash
python update_scores.py
```

Or double-click `run_update.bat` (Windows)

## Files

| File | Description |
|------|-------------|
| `config.json` | Match list and predictions |
| `generate_excel.py` | Creates the Excel file |
| `update_scores.py` | Fetches scores from web |
| `run_update.bat` | Windows shortcut |
| `predictions.xlsx` | Generated spreadsheet |

## Requirements

- Python 3.8+
- Libraries: `xlsxwriter`, `openpyxl`, `requests`, `beautifulsoup4`

## License

MIT - Do whatever you want with it!
