# 🕷️ Competitor Price Tracker

A Python web-scraping demo that parses a local HTML product catalog, extracts structured product information, normalizes prices, and exports the results to Excel.

`Python 3.x` · `BeautifulSoup4` · `OpenPyXL` · `MIT License`

## 🚀 Features

- **Product extraction:** Reads product names, SKUs, categories, prices, stock information, and ratings from HTML.
- **Data normalization:** Converts price strings into numeric values suitable for analysis.
- **Excel export:** Writes the extracted dataset into a formatted `.xlsx` workbook.
- **Local test environment:** Includes a mock catalog so the scraper can be tested without depending on a third-party website.

## 🧩 Project Structure

- `mock_marketplace.py` — creates the local catalog data.
- `competitor_catalog.html` — sample HTML catalog used for parsing.
- `intelligence_bot.py` — scraping, transformation, and Excel export logic.
- `Competitor_Market_Intelligence.xlsx` — sample generated output.

## 📋 Setup

```bash
git clone https://github.com/afaqkhan-io/automated-competitor-price-tracker.git
cd automated-competitor-price-tracker
python -m venv .venv
```

### Windows

```powershell
.venv\Scripts\activate
```

### macOS / Linux

```bash
source .venv/bin/activate
```

Install dependencies:

```bash
pip install beautifulsoup4 openpyxl
```

## ▶️ Run

```bash
python intelligence_bot.py
```

Open `Competitor_Market_Intelligence.xlsx` to inspect the exported dataset.

> **Note:** The included catalog is a local simulation for development and demonstration. When adapting the scraper to a real website, respect that site's terms, robots rules, and applicable laws.

## 📄 License

Distributed under the **MIT License**. See `LICENSE` for details.
