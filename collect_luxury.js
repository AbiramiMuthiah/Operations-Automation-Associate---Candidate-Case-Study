// Collect every live product from a Shopify store, plus the Malaysia (RM) price of each one.
// Run in the browser console on https://luxlexicon.com/en-my or https://carousellluxury.com/en-my.
// Output: a JSON file with the full product feed, the RM price per product, and store metadata.
(async () => {
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const products = [];

  // 1. Full product feed in the store's base currency (S$), 250 products per page
  for (let page = 1; ; page++) {
    const d = await fetch(`/products.json?limit=250&page=${page}`).then(r => r.json());
    if (!d.products || !d.products.length) break;
    products.push(...d.products);
    await sleep(800);
  }

  // 2. Malaysia price for each product: /en-my/products/<handle>.js returns RM prices in cents
  const rm = {};
  let i = 0;
  async function worker() {
    while (i < products.length) {
      const h = products[i++].handle;
      for (let attempt = 0; attempt < 3; attempt++) {
        try {
          const r = await fetch(`/en-my/products/${h}.js`);
          if (r.status === 429) { await sleep(5000); continue; }   // rate limited: back off
          const d = await r.json();
          rm[h] = Math.min(...d.variants.map(v => v.price)) / 100;
          break;
        } catch (e) { await sleep(1500); }
      }
      await sleep(250);
    }
  }
  await Promise.all([worker(), worker(), worker()]);   // 3 requests in parallel

  // 3. Store metadata (name, country, base currency, product count) and save
  const meta = await fetch('/meta.json').then(r => r.json());
  const out = { site: meta.name, base: location.origin, snapshot_utc: new Date().toISOString(),
                meta, shopify_currency: window.Shopify && Shopify.currency, rm, products };
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([JSON.stringify(out)], { type: 'application/json' }));
  a.download = `${meta.name.toLowerCase().replace(/ /g, '_')}_data.json`;
  a.click();
  console.log(`Saved ${products.length} products, ${Object.keys(rm).length} RM prices`);
})();
