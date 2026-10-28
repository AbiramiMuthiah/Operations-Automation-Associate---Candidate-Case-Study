# Luxury inventory snapshot: LuxLexicon vs Carousell Luxury

Code behind a one-day inventory snapshot (28 Oct 2026) comparing every live listing on
[luxlexicon.com](https://luxlexicon.com) and [carousellluxury.com](https://carousellluxury.com):
2,074 and 831 products respectively, each with its Malaysia (RM) storefront price.

## How it works

| Step | File | What it does |
|---|---|---|
| 1. Collect | `collect_luxury.js` | Both stores run on Shopify. The script reads the public product feed (`/products.json`, 250 per page) for the full catalogue in S$, then reads `/en-my/products/<handle>.js` for each product to get the price a Malaysian shopper sees in RM. Requests are paced (3 in parallel, with pauses and back-off on rate limits). Output: one JSON file per store. |
| 2. Clean | `clean.py` | Standardises brand names, maps 167 product types to 8 categories, parses condition and target audience from descriptions, flags sale items, exotics and price bands, and logs every data-quality issue it finds. Output: `raw.csv`, `clean.csv`, `dq_log.csv`. |
| 3. Report | `build_xlsx.py` | Builds the workbook (Dashboard, Clean_Data, Raw_Data, Mappings, Methodology, Code). Every dashboard figure is a live formula over the data, and brand, category and price-band rules are lookup tables, so editing a mapping updates every product. |

## Run it

1. Open https://luxlexicon.com/en-my in Chrome, open the developer console (F12 → Console), paste collect_luxury.js and press Enter. (Chrome may ask you to type "allow pasting" first.) It downloads luxlexicon_data.json. Repeat on https://carousellluxury.com/en-my to get carousell_luxury_data.json.
2. Create a folder called data/ next to the scripts and move both JSON files into it.
3. pip install -r requirements.txt
4. python clean.py
5. python build_xlsx.py (writes out/Luxury_Inventory_Snapshot.xlsx; open it in Google Sheets via File > Import > Upload)

## Notes

- Only public storefront data is read; no logins and no HTML scraping.
- Prices are asking prices, not sold prices. Data reflects the stores on the snapshot date.
- Scheduling step 1 daily would turn this into a running record of new listings, price changes and items that sell.

