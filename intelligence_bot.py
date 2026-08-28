import os
import pandas as pd
from bs4 import BeautifulSoup
from mock_marketplace import generate_competitor_site


def run_intelligence_bot():
    print("[BOT ENGINE] Triggering automated scraping pipeline...")

    # 1. Ensure target mock web interface exists locally
    if not os.path.exists("competitor_catalog.html"):
        generate_competitor_site()

    # 2. Ingest and parse the HTML structure
    with open("competitor_catalog.html", "r", encoding="utf-8") as f:
        html_content = f.read()

    soup = BeautifulSoup(html_content, "html.parser")
    product_cards = soup.find_all("div", class_="product-card")

    scraped_data = []

    # 3. Extract discrete data attributes using strict CSS selectors
    for card in product_cards:
        sku = card.get("data-sku", "N/A")
        title = card.find("h2", class_="product-title").text.strip()
        category = card.find("span", class_="product-category").text.strip()

        # Parse currency string to float for downstream accounting analytics
        raw_price = card.find("span", class_="competitor-price").text.strip()
        numeric_price = float(raw_price.replace("$", "").replace(",", ""))

        stock = card.find("span", class_="stock-status").text.strip()
        rating_text = card.find("div", class_="rating").text.strip()

        scraped_data.append(
            {
                "SKU": sku,
                "Product_Name": title,
                "Category": category,
                "Competitor_Price": numeric_price,
                "Inventory_Status": stock,
                "Market_Feedback": rating_text,
            }
        )

    # 4. Convert structural records into a professional Pandas DataFrame
    df = pd.DataFrame(scraped_data)

    # 5. Output raw dataset and an executive Excel spreadsheet
    output_excel = "Competitor_Market_Intelligence.xlsx"

    # Standard format conversion with automated column width fitting
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Live Price Index", index=False)

        # Access openpyxl back-end worksheet for clean formatting
        workbook = writer.book
        worksheet = writer.sheets["Live Price Index"]

        # Correct way to enable gridlines in openpyxl sheets
        worksheet.views.sheetView[0].showGridLines = True

        # Apply standard currency format to the price column
        for cell in worksheet["D"]:
            if cell.row != 1:  # Skip header row
                cell.number_format = "$#,##0.00"

        # Fit columns dynamically
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)

    print(
        f"[BOT SUCCESS] Successfully extracted {len(df)} structural competitor products."
    )
    print(f"[BOT SUCCESS] Standardized ledger exported to '{output_excel}'")


if __name__ == "__main__":
    run_intelligence_bot()
