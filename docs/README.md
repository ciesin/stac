# CIESIN SEDAC STAC Browser

Preview of the CIESIN STAC with data incorporated from GRID3, Harvard Dataverse, and NASA EarthData. 

## Contents

- `index.html` - static Vue browser with category filters, dashboard highlights, details, and asset tables
- `app.js` - search, filters, STAC item loading, and asset table logic
- `styles.css` - browser styling
- `stac/catalog.json` - STAC root catalog
- `stac/collections/*/collection.json` - one STAC collection per SEDAC category
- `stac/collections/*/items/*.json` - generated STAC items
- `stac/items-index.json` - compact index used by the browser
- `stac/summary.json` - generated build summary
- `stac/asset-rows-not-attached.json` - asset rows that could not be attached to an item

The browser is a static Vue page that loads Vue from a CDN. It does not require a Vue/Vite build step.

## Regenerate

From `WORKIN DIRECTORY`:

```powershell
python .\scripts\build_category_stac_catalog.py --items-csv .\outputs\main_item_list\ciesin_main_stac_item_list.csv --assets-csv .\outputs\main_asset_list\ciesin_main_stac_asset_list.csv --output-dir .\outputs\category_stac --clean
python .\scripts\build_category_stac_browser.py --stac-source-dir .\outputs\category_stac\stac --docs-dir .\docs --clean
```

## GitHub Pages

Commit this `docs` folder, then configure GitHub Pages to deploy from the `docs` folder on the selected branch.
