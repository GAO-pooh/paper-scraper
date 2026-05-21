<div align="right">

[English](README.md) | [简体中文](README_zh.md)

</div>

# Academic Paper Scraper

Automated scrapers for **ScienceDirect** and **INFORMS PubsOnLine** — search papers by keyword, author, or journal, and batch-download PDFs using your institutional access.

---

## Supported Platforms

| Script | Platform | PDF Download |
|--------|----------|--------------|
| `sd_scraper.py` / `sd_scraper_en.py` | [ScienceDirect](https://www.sciencedirect.com) (Elsevier) | Via Chrome DevTools (bypasses Cloudflare) |
| `informs_scraper.py` / `informs_scraper_en.py` | [INFORMS PubsOnLine](https://pubsonline.informs.org) | Direct HTTP with session cookie |

Files ending in `_en.py` are full English-interface versions; the others are Chinese-interface versions. Logic is identical.

---

## Features

- **Multiple search modes**: keyword, journal browse, journal+keyword, author, ISSN, advanced (combine any criteria)
- **Bulk PDF download** with institutional access (university SSO / CARSI)
- **Anti-bot measures**: Chrome TLS fingerprint spoofing via `curl_cffi`, stealth JS injection, automatic rate-limit handling
- **ScienceDirect**: Chrome DevTools Protocol (CDP) captures PDF bytes directly — no Playwright required
- **INFORMS**: simple direct HTTP download; Atypon does not enforce JS challenges on PDF endpoints
- Output formats: **CSV / JSON / XLSX**
- Interactive wizard mode (run with no arguments)

---

## Installation

**Step 1 — Install dependencies**

```bash
pip install curl_cffi websocket-client browser-cookie3 openpyxl
```

For INFORMS scraper, also install:

```bash
pip install beautifulsoup4 lxml
```

**Step 2 — Run**

```bash
# ScienceDirect (interactive wizard)
python3 sd_scraper_en.py

# INFORMS (interactive wizard)
python3 informs_scraper_en.py
```

> **macOS only** (Chrome path is hardcoded to `/Applications/Google Chrome.app`).  
> On Linux/Windows, set `CHROME_BIN` in the scraper class to your Chrome executable path.

Full dependency list: [`requirements.txt`](requirements.txt)

---

## Quick Start

### ScienceDirect

```bash
# Interactive wizard (recommended for first use)
python sd_scraper_en.py

# Keyword search — save metadata as XLSX
python sd_scraper_en.py -m keyword -q "machine learning" -n 100 --browser-cookies

# Keyword search + download PDFs
python sd_scraper_en.py -m keyword -q "deep learning" -n 50 --browser-cookies --download-pdfs

# Browse a journal (most recent first)
python sd_scraper_en.py -m journal -j "Energy" -n 200 --browser-cookies --sort date

# Keyword search within a specific journal
python sd_scraper_en.py -m journal_keyword -j "Renewable Energy" -q "solar cell" -n 50 --browser-cookies

# Search by author
python sd_scraper_en.py -m author -a "Zhang Wei" -n 30 --browser-cookies

# Advanced search (combine criteria)
python sd_scraper_en.py -m advanced -q "deep learning" --date 2021-2024 --type REV -n 50 --browser-cookies --download-pdfs
```

### INFORMS PubsOnLine

```bash
# Interactive wizard
python informs_scraper_en.py

# Keyword search
python informs_scraper_en.py -m keyword -q "supply chain" -n 100 --browser-cookies

# Browse a journal
python informs_scraper_en.py -m journal -j mnsc -n 200 --browser-cookies

# Specific volume/issue TOC
python informs_scraper_en.py -m toc -j mnsc -v 71 -i 3 --browser-cookies

# Keyword search + download PDFs
python informs_scraper_en.py -m keyword -q "inventory" -n 50 --browser-cookies --download-pdf

# Best login method: pop up Chrome, log in manually
python informs_scraper_en.py -m keyword -q "machine learning" -n 50 --chrome-login --download-pdf
```

#### INFORMS Journal Codes

| Code | Journal |
|------|---------|
| `mnsc` | Management Science |
| `opre` | Operations Research |
| `ijoc` | INFORMS Journal on Computing |
| `mksc` | Marketing Science |
| `msom` | Manufacturing & Service Operations Management |
| `trsc` | Transportation Science |
| `isre` | Information Systems Research |
| `orsc` | Organization Science |

---

## How PDF Download Works

### ScienceDirect

Cloudflare blocks direct HTTP requests to PDF endpoints. The scraper works around this using Chrome DevTools Protocol (CDP):

1. Script auto-launches Chrome in debug mode (copies your existing profile — no re-login needed)
2. Navigates to the article page, establishing cookie context
3. Intercepts the PDF response bytes via `Network` / `Fetch` DevTools events
4. Writes the PDF directly to disk — no Save dialog, no Playwright required

**First-time setup**: open Chrome, log in to ScienceDirect via your institution (SSO/CARSI), and open at least one article PDF to confirm access. The script handles everything else automatically.

```bash
# If Chrome isn't already logged in, use this first:
python sd_scraper_en.py --open-browser-login
# Then run your actual scrape:
python sd_scraper_en.py -m keyword -q "turbine" -n 50 --browser-cookies --download-pdfs
```

### INFORMS

INFORMS (Atypon platform) does not enforce JS challenges on PDF endpoints — a valid session cookie is sufficient.

```bash
# Option 1: read cookies from Chrome (must be logged in already)
python informs_scraper_en.py -m keyword -q "inventory" -n 30 --browser-cookies --download-pdf

# Option 2: pop up Chrome, log in, then auto-extract cookies
python informs_scraper_en.py -m keyword -q "inventory" -n 30 --chrome-login --download-pdf

# Option 3: member credentials (direct login)
python informs_scraper_en.py -m keyword -q "inventory" -n 30 --member 123456 --password MyPwd --download-pdf
```

---

## Output

All output goes to `./results/` (ScienceDirect) or `./informs_result/` (INFORMS) by default.

```
results/
└── keyword_machine_learning_20250101_120000/
    ├── keyword_machine_learning_20250101_120000.xlsx   ← metadata
    └── pdfs/
        ├── 001_Zhang_2024_Deep learning for...pdf
        ├── 002_Li_2023_Transfer learning in...pdf
        └── ...
```

---

## All CLI Options

### ScienceDirect (`sd_scraper_en.py`)

| Option | Description |
|--------|-------------|
| `-m`, `--mode` | `keyword` / `journal` / `journal_keyword` / `author` / `issn` / `advanced` |
| `-q`, `--query` | Search keywords (supports `AND` / `OR` / `NOT`) |
| `-j`, `--journal` | Journal name |
| `-a`, `--author` | Author name |
| `--issn` | Journal ISSN |
| `-n`, `--count` | Max papers to fetch (default: 50) |
| `--date` | Year range, e.g. `2020-2024` |
| `--sort` | `relevance` (default) or `date` |
| `--type` | Article type: `FLA` / `REV` / `SCO` |
| `--browser-cookies` | Auto-read cookies from local Chrome |
| `--cookies` | Path to a cookie JSON file |
| `--format` | Output format: `xlsx` (default) / `csv` / `json` / `all` |
| `--download-pdfs` | Download PDFs after saving metadata |
| `--output` | Custom output directory |
| `--open-browser-login` | Open Chrome for manual institutional login |
| `--interactive` | Launch interactive wizard |

### INFORMS (`informs_scraper_en.py`)

| Option | Description |
|--------|-------------|
| `-m`, `--mode` | `keyword` / `journal` / `toc` / `advanced` |
| `-q`, `--query` | Search keywords |
| `-j`, `--journal` | Journal code (e.g. `mnsc`) |
| `-v`, `--volume` | Volume number (toc mode) |
| `-i`, `--issue` | Issue number (toc mode) |
| `--author` | Author name (advanced mode) |
| `--date` | Year range, e.g. `2020-2024` |
| `-n`, `--count` | Max papers (default: 100) |
| `--chrome-login` | ★ Pop up Chrome for manual login (most reliable) |
| `--browser-cookies` | Read cookies from local Chrome |
| `--cookies-file` | Load cookies from a JSON file |
| `--member` | INFORMS member ID |
| `--password` | Account password |
| `--format` | `csv` (default) / `json` / `xlsx` |
| `--download-pdf` | Download PDFs after scraping |
| `-o`, `--output-dir` | Output directory |

---

## Notes

- **Institutional access required** for full-text PDF download. Open-access papers can be downloaded without login.
- **Rate limiting**: the scraper adds random delays between requests (default 2–5 s). Do not reduce these aggressively.
- **macOS only** in current form. The Chrome binary path (`/Applications/Google Chrome.app/...`) is hardcoded. Linux/Windows users: update `CHROME_BIN` in the class definition.
- **Cookie expiry**: session cookies expire (typically days to weeks). Re-login if you encounter 403 errors.
- Use responsibly and in accordance with your institution's and the publishers' terms of service.

---

## License

MIT
