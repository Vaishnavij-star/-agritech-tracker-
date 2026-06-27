# AgriTech Startup Discovery Tracker

An automated solution that periodically scans online sources to identify newly emerging AgriTech startups 

---

## What It Does

- **Discovers** real AgriTech startups from live news using AI-powered web search
- **Extracts** structured information: startup name, category, country, funding stage, description, why it was in the news, website, and source URL
- **Deduplicates** entries using MD5 hashing so no startup is saved twice
- **Stores** everything in a formatted Excel file (`agritech_startups.xlsx`)
- **Runs automatically** every Monday at 08:00 UTC via GitHub Actions

---

## Project Structure

```
agritech-tracker/
├── tracker.py                        # Main script
├── agritech_startups.xlsx            # Output Excel database
├── seen_hashes.json                  # Deduplication store
├── requirements.txt                  # Python dependencies
├── env                               # API keys (not committed to GitHub)
├── .gitignore
└── .github/
    └── workflows/
        └── weekly_scan.yml           # GitHub Actions weekly automation
```

---

## How It Works

### 1. Discovery
The script sends 10 targeted search queries to Claude AI (with OpenAI as fallback), which uses its built-in web search tool to scan:
- AgFunder News
- TechCrunch AgTech
- Crunchbase
- Business Wire
- PR Newswire
- StartUs Insights
- And more

Example queries:
```
"agritech startup raised funding 2026"
"farm robotics startup funding 2026"
"precision agriculture startup new product launch 2026"
```

### 2. AI Extraction
Claude reads the search results and extracts structured data for each startup it finds:

```json
{
  "startup_name": "AgZen",
  "category": "Precision Farming",
  "country": "United States",
  "funding_stage": "Series B",
  "description": "MIT spinout that provides real-time droplet-level control over crop spraying.",
  "news_summary": "Raised $10M Series B in March 2026 after acreage grew 15x in a single season.",
  "startup_website": "https://www.agzen.com",
  "source_url": "https://agfundernews.com/..."
}
```

### 3. Deduplication
Every startup name is hashed (MD5) and stored in `seen_hashes.json`. On every subsequent run, already-tracked startups are skipped automatically.

### 4. Excel Output
Results are saved to a professionally formatted Excel file with:
- Green header row with frozen panes
- Alternating row colours
- Auto-sized columns
- A Summary sheet with live `COUNTA` formula tracking total startups

---

## Sample Data (Current Excel)

| Startup | Category | Country | Stage | Source |
|---|---|---|---|---|
| AgZen | Precision Farming | USA | Series B | AgFunderNews |
| Halter | Livestock Tech | New Zealand | Series C+ | BusinessWire |
| Tropic Biosciences | AgBiotech | UK | Series C | Crunchbase |
| 4AG Robotics | Farm Robotics | Canada | Series B | Crunchbase |
| Orchard Robotics | Precision Farming | USA | Series A | AgritechDigest |
| TRIC Robotics | Farm Robotics | USA | Seed | The Robot Report |
| Biographica | AgBiotech | UK | Seed | NewMarketPitch |
| Verdant Impact | Livestock Tech | India | Seed | NewMarketPitch |
| ArkeaBio | Livestock Tech | USA | Series A | NewMarketPitch |
| SwarmFarm Robotics | Farm Robotics | Australia | Series A | Grain Central |
| Miraterra | Precision Farming | Canada | Seed | NewMarketPitch |
| Agriodor | Crop Tech | France | Seed | NewMarketPitch |
| AgriPass | Farm Robotics | USA | Seed | Global Agriculture |
| Singrow | Crop Tech | Singapore | Seed | AgFunderNews |
| Nbryo | Livestock Tech | New Zealand | Seed | NewMarketPitch |

> All startups are real and verified from live sources (June 2026). Every row includes a source URL.

---

## Weekly Automation

The included GitHub Actions workflow runs the tracker every Monday automatically:

```yaml
on:
  schedule:
    - cron: "0 8 * * 1"   # Every Monday 08:00 UTC
  workflow_dispatch:        # Manual trigger available
```

After each scan it commits the updated Excel and `seen_hashes.json` back to the repo.

To enable: add `ANTHROPIC_API_KEY` as a GitHub repository secret.

---

## Setup & Running Locally

```bash
# 1. Clone the repo
git clone https://github.com/yourusername/agritech-tracker
cd agritech-tracker

# 2. Install dependencies
pip install -r requirements.txt

# 3. Add your API key to env file
echo "ANTHROPIC_API_KEY=sk-ant-xxxx" > env
# Optional fallback:
echo "OPENAI_API_KEY=sk-xxxx" >> env

# 4. Run
python tracker.py
```

---

## Requirements

```
requests
openpyxl
python-dotenv
```

---

## Tech Stack

| Component | Technology |
|---|---|
| AI Discovery | Claude claude-sonnet-4-6 with web_search tool |
| Fallback | OpenAI GPT-4o-mini |
| Data Storage | Excel (.xlsx) via openpyxl |
| Deduplication | MD5 hashing + JSON store |
| Automation | GitHub Actions (weekly cron) |
| Language | Python 3.11+ |

---

## Scalability

- Add more search queries in `SEARCH_QUERIES` list to cover more sources
- Add more API keys in `env` file for higher rate limits
- Schedule more frequent scans by editing the cron expression
- Swap Excel for a database (PostgreSQL, Supabase) for larger scale

---

*Built by Vaishnavi Jaiswal 