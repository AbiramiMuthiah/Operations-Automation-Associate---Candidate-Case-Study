# Luxury inventory snapshot: LuxLexicon vs Carousell Luxury

Code behind a one-day inventory snapshot (28 Oct 2026) comparing every live listing on
[luxlexicon.com](https://luxlexicon.com) and [carousellluxury.com](https://carousellluxury.com):
2,074 and 831 products respectively, each with its Malaysia (RM) storefront price.

## How it works

| Step | File | What it does |
|---|---|---|
| 1. Collect | `collect_luxury.js` | Both stores run on Shopify. The script reads the public product feed (`/products.json`, 250 per page) for the full catalogue in S$, then reads `/en-my/products/<handle>.js` for each product to get the price a Malaysian shopper sees in RM. Requests are paced (3 in parallel, with pauses and back-off on rate limits). Output: one JSON file per store. |
| 2. Clean | `clean.py` | Standardises brand names, maps 167 product types to 9 categories, parses condition and target audience from descriptions, flags sale items, exotics and price bands, and logs every data-quality issue it finds. Output: `raw.csv`, `clean.csv`, `dq_log.csv`. |
| 3. Report | `build_xlsx.py` | Builds the workbook (Raw_Data, Clean_Data, Mappings, Dashboard, Methodology). Cleaning rules and every dashboard figure are live formulas over the data, so editing a mapping re-cleans the whole dataset. |

## Run it

1. Open `https://luxlexicon.com/en-my` in a browser, paste `collect_luxury.js` into the developer console and run it. Repeat on `https://carousellluxury.com/en-my`. Put both JSON files in `data/`.
2. `pip install -r requirements.txt`
3. `python clean.py`
4. `python build_xlsx.py` (writes `out/Luxury_Inventory_Snapshot.xlsx`, which opens in Google Sheets via File > Import)

## Notes

- Only public storefront data is read; no logins and no HTML scraping.
- Prices are asking prices, not sold prices. Data reflects the stores on the snapshot date.
- Scheduling step 1 daily would turn this into a running record of new listings, price changes and items that sell.
