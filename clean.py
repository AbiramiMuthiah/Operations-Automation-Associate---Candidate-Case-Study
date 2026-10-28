"""Clean and standardise the LuxLexicon + Carousell Luxury snapshot.

Inputs : data/luxlexicon_data.json, data/carousell_luxury_data.json
         (full Shopify /products.json feed + exact RM price per product from /en-my/products/<handle>.js)
Outputs: out/raw.csv, out/clean.csv, out/dq_log.csv
"""
import json
import re
from collections import Counter
from datetime import datetime, timezone

import pandas as pd

FILES = ["data/luxlexicon_data.json", "data/carousell_luxury_data.json"]

# ---------- brand normalisation ----------
BRAND_FIX = {
    "See By Chole": "See By Chloé",
    "ALAïA": "Alaïa",
    "Tiffany & Co": "Tiffany & Co.",
    "St Louis": "Saint-Louis",
    "Faure Le page": "Faure Le Page",
    "Gucci X Balenciaga": "Gucci x Balenciaga",
}
COLLAB_PARENT = {  # collaborations roll up to the parent house for brand analysis
    "Fendi x Sarah Coleman": "Fendi",
    "Gucci x Balenciaga": "Gucci",
    "Louis Vuitton x Yayoi Kusama": "Louis Vuitton",
    "Louis Vuitton x Stephen Sprouse": "Louis Vuitton",
}
BIG_HOUSES = ["Hermès", "Chanel", "Louis Vuitton", "Dior"]

# ---------- category mapping (product_type -> category) ----------
CAT = {
    "Bags": ["Shoulder Bag", "Top Handle Bag", "Crossbody Bag", "Tote Bag", "Bucket Bag", "Backpack", "Clutch",
             "Wallet On Chain", "Vanity Bag", "Vanity Case", "Belt Bag", "Messenger", "Briefcase", "Duffle Bag",
             "Hobo Bag", "Box Bag", "Phone Bag", "Handbag", "Pochette on Strap Shoulder Bag", "Clutch On Chain",
             "Clutch With Chain Camera Bag", "Vanity Top Handle"],
    "Small Leather Goods": ["Wallet", "Long Wallet", "Compact Wallet", "Card Holder", "Card Wallet", "Pouch",
                            "Key Ring", "Keychain", "Toiletries Case", "Toiletries", "Case", "Book cover",
                            "Jewellery Pouch"],
    "Jewellery": ["Bracelet", "Ring", "Necklace", "Earrings", "Brooch", "Long Necklace", "Choker Necklace",
                  "Jewellery", "Jewellery set"],
    "Watches": ["Watch"],
    "Accessories": ["Bag Accessory", "Charm", "Scarf", "Twilly", "Twilly (Pair)", "Twilly Ring", "Scarf Ring",
                    "Shawl", "Belt", "Belt Kit", "Strap", "Bag Strap", "Tie", "Hat", "Bucket Hat", "Glove Clip",
                    "Gloves"],
    "Shoes": ["Sandals", "Sneakers", "Loafers", "Heels", "Flats", "Ballet Flats", "Shoes"],
    "Apparel": ["Ready-To-Wear"],
    "Home & Tableware": ["Tableware", "Porcelaine", "Homeware", "Change Tray", "Tray", "Tea Cup And Saucer"],
}
TYPE_TO_CAT = {t.lower(): c for c, ts in CAT.items() for t in ts}
TYPE_FIX = {"wallet on chain": "Wallet On Chain", "ready-to-wear": "Ready-To-Wear"}
LL_GROUP_TO_CAT = {"Bags": "Bags", "Small Leather Goods": "Small Leather Goods", "Jewellery": "Jewellery",
                   "Watches": "Watches", "Accessories": "Accessories", "Apparel": "Apparel",
                   "Tableware": "Home & Tableware"}

COND_ORDER = ["Brand New", "Like New", "Very Good", "Good", "Fair", "Unknown"]
PRICE_BANDS_SGD = [(0, 1000, "Under S$1k"), (1000, 2500, "S$1k–2.5k"), (2500, 5000, "S$2.5k–5k"),
                   (5000, 10000, "S$5k–10k"), (10000, 20000, "S$10k–20k"), (20000, 50000, "S$20k–50k"),
                   (50000, 1e12, "S$50k+")]


def text(html):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html or "")).replace("&amp;", "&").strip()


def condition(site, desc, tags):
    if site == "LuxLexicon":
        fields = [f.strip() for f in desc.split("|")]
        for f in fields:
            fl = f.lower()
            if fl in ("brand new", "new"):
                return "Brand New"
            for c in ["Like New", "Very Good", "Good", "Fair"]:
                if fl.startswith(c.lower()):
                    return c
    m = re.search(r"\b(Brand New|Like New|Very Good|Good|Fair)\s+Condition", desc, re.I)
    if m:
        return m.group(1).title()
    for t in tags:
        m = re.match(r"(Brand New|Like New|Very Good|Good|Fair) Condition", t, re.I)
        if m:
            return m.group(1).title()
    m = re.search(r"\b(Brand New|Like New|Pristine)\b", desc, re.I)
    if m:
        return "Like New" if m.group(1).lower() == "pristine" else m.group(1).title()
    return "Unknown"


def material_from(site, desc, title):
    if site == "LuxLexicon":
        f = [x.strip() for x in desc.split("|")]
        if len(f) >= 4 and f[3] not in ("", "Others"):
            return f[3]
    m = re.search(r"\bin ([A-Za-z' ]+?)(?:,|$| Leather\b)", title)
    if m:
        mat = m.group(1).strip()
        return mat if mat else None
    return None


KEYWORD_CAT = [(r"sandal|sneaker|loafer|heel|flat|boot|mule|espadrille", "Shoes"),
               (r"pillow|plate|cup|tray|display|sculpture|blanket|cushion|vase", "Home & Tableware"),
               (r"bag|kelly|birkin|lindy|constance|evelyne|picotin|bolide|flap|tote|clutch", "Bags"),
               (r"wallet|card holder|pouch", "Small Leather Goods"),
               (r"ring|bracelet|necklace|earring|pendant", "Jewellery")]


def keyword_category(name):
    for pat, c in KEYWORD_CAT:
        if re.search(pat, name, re.I):
            return c
    return None


def audience(desc):
    tail = desc[-60:].lower()
    if "women" in tail or "lady" in tail:
        return "Women" if "unisex" not in tail else "Unisex"
    if "unisex" in tail:
        return "Unisex"
    if re.search(r"\bmen\b", tail):
        return "Men"
    return "Not stated"


def band(p):
    for lo, hi, name in PRICE_BANDS_SGD:
        if lo <= p < hi:
            return name
    return None


def main():
    raw_rows, rows, dq = [], [], Counter()
    for f in FILES:
        d = json.load(open(f, encoding="utf-8"))
        site, base, rm = d["site"], d["base"], d["rm"]
        snap = datetime.fromisoformat(d["snapshot_utc"].replace("Z", "+00:00"))
        seen = set()
        for p in d["products"]:
            v = p["variants"][0]
            desc = text(p.get("body_html"))
            price = float(v["price"])
            cmp_ = float(v["compare_at_price"]) if v.get("compare_at_price") else None
            raw_rows.append({
                "site": site, "snapshot_utc": d["snapshot_utc"][:16].replace("T", " "),
                "product_id": str(p["id"]), "product_name": p["title"], "brand_raw": p["vendor"],
                "product_type_raw": p["product_type"], "tags": " | ".join(p["tags"]), "sku": v.get("sku"),
                "price_sgd_raw": price, "compare_at_price_sgd_raw": cmp_, "price_myr_raw": rm.get(p["handle"]),
                "available": v.get("available"), "created_at": p["created_at"], "published_at": p["published_at"],
                "image_count": len(p["images"]), "url": f"{base}/products/{p['handle']}",
                "url_my": f"{base}/en-my/products/{p['handle']}", "description": desc[:500],
            })

            # ---- cleaning ----
            if p["id"] in seen:
                dq["duplicate product_id removed"] += 1
                continue
            seen.add(p["id"])
            name = re.sub(r"\s+", " ", p["title"]).strip()
            brand_raw = (p["vendor"] or "").strip()
            brand_full = BRAND_FIX.get(brand_raw, brand_raw)
            brand = COLLAB_PARENT.get(brand_full, brand_full)
            if brand != brand_raw:
                dq["brand name standardised"] += 1
            ptype = (p["product_type"] or "").strip()
            ptype = TYPE_FIX.get(ptype.lower(), ptype)
            cat = TYPE_TO_CAT.get(ptype.lower())
            if not cat and site == "LuxLexicon":
                grp = (desc.split("|") + ["", ""])[1].strip()
                cat = LL_GROUP_TO_CAT.get(grp)
                if cat:
                    dq["category filled from description"] += 1
            if not cat:
                cat = keyword_category(name)
                if cat:
                    dq["category inferred from product name"] += 1
            if not ptype:
                dq["blank product type"] += 1
                ptype = "Unspecified"
            if not cat:
                cat = "Other"
                dq["category = Other (unmapped)"] += 1
            cond = condition(site, desc, p["tags"])
            if cond == "Unknown":
                dq["condition not stated"] += 1
            rm_price = rm.get(p["handle"])
            if rm_price is None:
                dq["missing RM price"] += 1
            if price <= 0:
                dq["zero/negative price"] += 1
            if site == "Carousell Luxury" and "price less than 500 sgd" in p["tags"] and price >= 500:
                dq["tag 'price less than 500 sgd' contradicts price"] += 1
            on_sale = bool(cmp_ and cmp_ > price)
            if cmp_ and cmp_ <= price:
                dq["compare-at price not above price (ignored)"] += 1
            created = datetime.fromisoformat(p["created_at"])
            rows.append({
                "site": site, "product_id": str(p["id"]), "product_name": name, "brand": brand,
                "brand_detail": brand_full, "is_collab": brand_full != brand,
                "brand_group": brand if brand in BIG_HOUSES else "Other brands",
                "category": cat, "product_type": ptype, "condition": cond,
                "condition_rank": COND_ORDER.index(cond) + 1,
                "is_new": cond == "Brand New",
                "price_sgd": price, "price_myr": rm_price,
                "my_price_multiple": round(rm_price / price, 4) if rm_price and price else None,
                "price_band_sgd": band(price),
                "compare_at_sgd": cmp_ if on_sale else None, "on_sale": on_sale,
                "discount_pct": round(1 - price / cmp_, 4) if on_sale else None,
                "available": bool(v.get("available")),
                "exotic": "exotic" in [t.lower() for t in p["tags"]] or bool(
                    re.search(r"crocodile|alligator|ostrich|lizard|python|niloticus|porosus", name + " " + desc, re.I)),
                "audience": audience(desc),
                "listed_date": created.date().isoformat(),
                "listed_month": created.strftime("%Y-%m"),
                "days_listed": (snap - created.astimezone(timezone.utc)).days,
                "image_count": len(p["images"]),
                "url": f"{base}/products/{p['handle']}",
            })
    raw, clean = pd.DataFrame(raw_rows), pd.DataFrame(rows)
    import os
    os.makedirs("out", exist_ok=True)
    raw.to_csv("out/raw.csv", index=False, encoding="utf-8-sig")
    clean.to_csv("out/clean.csv", index=False, encoding="utf-8-sig")
    pd.DataFrame(sorted(dq.items(), key=lambda x: -x[1]), columns=["issue", "count"]).to_csv(
        "out/dq_log.csv", index=False, encoding="utf-8-sig")
    print(len(raw), len(clean)); print(pd.DataFrame(dq.items()))


if __name__ == "__main__":
    main()
