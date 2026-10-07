# Book Catalog Scraper

A Python portfolio project that collects book data from the Books to Scrape practice website.

## Features

- Visits up to three catalog pages
- Extracts book titles, prices, ratings, availability, and product links
- Removes duplicate records by product URL
- Saves the results as CSV and Excel files
- Creates a report with collection statistics

## Requirements

- Python 3.12 or newer
- Packages listed in `requirements.txt`

## Setup

Open PowerShell in the project folder and install the packages:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Run

```powershell
.\.venv\Scripts\python.exe scraper.py
```

The program saves its results in the `output` folder:

- `books.csv` — collected records in CSV format
- `books.xlsx` — collected records in Excel format
- `report.txt` — collection summary

## Data fields

- Catalog page number
- Book title
- Price in GBP
- Rating from 1 to 5
- Availability
- Product page URL

## Notes

This project uses Books to Scrape, a practice website made for web scraping exercises. Its prices and ratings are randomly assigned and do not represent real products. The scraper is limited to three pages for this portfolio demonstration.