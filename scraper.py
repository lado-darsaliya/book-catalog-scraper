import csv
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font


START_URL = "https://books.toscrape.com/"
MAX_PAGES = 3
OUTPUT_DIR = Path("output")

RATINGS = {
    "One": 1,
    "Two": 2,
    "Three": 3,
    "Four": 4,
    "Five": 5,
}

COLUMNS = [
    "page",
    "title",
    "price_gbp",
    "rating",
    "availability",
    "product_url",
]


def get_page(url):
    """Download one page and turn its HTML into a searchable document."""
    try:
        response = requests.get(
            url,
            headers={"User-Agent": "PortfolioCatalogScraper/1.0"},
            timeout=20,
        )
        response.raise_for_status()
    except requests.RequestException as error:
        raise SystemExit(f"Could not download the page: {error}") from error

    return BeautifulSoup(response.text, "html.parser")


def extract_books(soup, page_number, page_url):
    """Read book details from one catalog page."""
    books = []

    for card in soup.select("article.product_pod"):
        title_link = card.select_one("h3 a")
        price_element = card.select_one(".price_color")
        availability_element = card.select_one(".availability")
        rating_element = card.select_one("p.star-rating")

        if title_link is None or price_element is None:
            continue

        title = title_link.get("title", title_link.get_text(strip=True))
        price_text = price_element.get_text(strip=True).replace("£", "").strip()

        try:
            price = float(price_text)
        except ValueError:
            price = None

        rating = None
        if rating_element is not None:
            for css_class in rating_element.get("class", []):
                if css_class in RATINGS:
                    rating = RATINGS[css_class]
                    break

        availability = ""
        if availability_element is not None:
            availability = " ".join(
                availability_element.get_text(" ", strip=True).split()
            )

        product_url = urljoin(page_url, title_link.get("href", ""))

        books.append(
            {
                "page": page_number,
                "title": title,
                "price_gbp": price,
                "rating": rating,
                "availability": availability,
                "product_url": product_url,
            }
        )

    return books


def save_csv(books, path):
    with path.open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.DictWriter(file, fieldnames=COLUMNS)
        writer.writeheader()
        writer.writerows(books)


def save_excel(books, path):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Books"

    sheet.append(COLUMNS)

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for book in books:
        sheet.append([book[column] for column in COLUMNS])

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    column_widths = {
        "A": 10,
        "B": 45,
        "C": 15,
        "D": 12,
        "E": 25,
        "F": 60,
    }
    for column, width in column_widths.items():
        sheet.column_dimensions[column].width = width

    workbook.save(path)


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)

    all_books = []
    current_url = START_URL
    pages_visited = 0

    for page_number in range(1, MAX_PAGES + 1):
        print(f"Reading page {page_number}: {current_url}")

        soup = get_page(current_url)
        pages_visited += 1

        books_on_page = extract_books(soup, page_number, current_url)
        all_books.extend(books_on_page)

        next_link = soup.select_one("li.next a")
        if next_link is None or page_number == MAX_PAGES:
            break

        current_url = urljoin(current_url, next_link.get("href", ""))
        time.sleep(0.5)

    unique_books = []
    seen_urls = set()

    for book in all_books:
        if book["product_url"] not in seen_urls:
            unique_books.append(book)
            seen_urls.add(book["product_url"])

    csv_path = OUTPUT_DIR / "books.csv"
    excel_path = OUTPUT_DIR / "books.xlsx"
    report_path = OUTPUT_DIR / "report.txt"

    save_csv(unique_books, csv_path)
    save_excel(unique_books, excel_path)

    missing_titles = sum(not book["title"] for book in unique_books)
    missing_prices = sum(book["price_gbp"] is None for book in unique_books)

    report = (
        "Book Catalog Scraper Report\n"
        f"Source: {START_URL}\n"
        f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"Pages visited: {pages_visited}\n"
        f"Books collected: {len(unique_books)}\n"
        f"Duplicate records removed: {len(all_books) - len(unique_books)}\n"
        f"Books with missing titles: {missing_titles}\n"
        f"Books with missing prices: {missing_prices}\n"
    )
    report_path.write_text(report, encoding="utf-8")

    print("\nDone!")
    print(f"Books collected: {len(unique_books)}")
    print(f"CSV file: {csv_path}")
    print(f"Excel file: {excel_path}")
    print(f"Report: {report_path}")


if __name__ == "__main__":
    main()