# Software QA & Testing — Assignment 6 (AITU)

![Python](https://img.shields.io/badge/python-3.11-blue)
![Selenium](https://img.shields.io/badge/tool-Selenium-43B02A)
![BrowserStack](https://img.shields.io/badge/cloud-BrowserStack-E86C1A)
![Course](https://img.shields.io/badge/course-SQAT-lightgrey)

## Overview

Assignment 6 for the **Software Quality Assurance and Testing** course at Astana IT University.
Two tasks in one script:
1. **Data-driven testing** — all test inputs are read from an Excel file via `openpyxl`; no values are hardcoded in the test logic.
2. **Cross-browser cloud execution** — the Selenium test runs on **BrowserStack Remote WebDriver** against Chrome and Firefox on Windows 11.

Target site: [booking.com](https://www.booking.com) — hotel search scenario (destination → dates → guests → results).

## What it covers

- **Data-driven testing** with `openpyxl` — URL, destination, date offsets, and guest counts all come from `testdata.xlsx`
- **Selenium Remote WebDriver** via BrowserStack W3C capabilities (Selenium 4 style)
- **Cross-browser execution** — Chrome latest and Firefox latest, Windows 11
- **Structured step logging** — JSON run log persisted per browser per run
- **Screenshot capture** at each step (home, destination, dates, results, error if failed)
- **PDF report generation** with `reportlab` and `matplotlib` (data proof table + screenshots embedded)
- Best-effort overlay and cookie-consent handling for real-world sites

## Project structure

```
assignment6_sqat/
├── run_assignment6.py        # Main script: Excel read, BrowserStack run, PDF report
├── testdata.xlsx             # Test data (destination, dates, guests, URL)
├── requirements.txt
└── artifacts/                # Generated per run (timestamped folder)
    └── YYYYMMDD_HHMMSS/
        ├── excel_data_proof.png
        ├── Assignment6_Report.pdf
        ├── Windows_11_Chrome/
        │   ├── 01_home.png
        │   ├── 02_destination_filled.png
        │   ├── run_log.json
        │   └── ...
        └── Windows_11_Firefox/
            └── ...
```

## Getting started

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Set BrowserStack credentials:

```bash
export BROWSERSTACK_USERNAME=<your_username>
export BROWSERSTACK_ACCESS_KEY=<your_access_key>
```

Run:

```bash
python run_assignment6.py
```

Artifacts (screenshots, JSON logs, PDF report) are written to `artifacts/<timestamp>/`.

---

Adil Ormanov — [GitHub](https://github.com/Adilforest)
