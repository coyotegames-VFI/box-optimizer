# Build Status

## Implemented

- Python package structure for `box_optimizer`
- Dimension normalization to centimeters, including flat two-dimension items
- Weight normalization for kg, g, lb, and oz
- Dimensional, packed actual, and chargeable weight helpers
- Tiered per-item padding plus final exterior carton padding
- Core dataclasses for SKU, order, packed item, carton, and unmatched SKU records
- Rotation and geometry helpers for 3D packing
- Deterministic heuristic 3D packer with rotations, boundaries, overlap checks, and placements
- Optimized carton dimension search capped at `74 x 37 x 44 cm`
- Multi-box splitter for orders that cannot fit in one capped carton
- CSV/XLSX intake for SKU masters and orders
- Combined dimension parsing from columns like `Dimensions`, `Size`, `Dims`, and `LxWxH`
- Wide-format order parsing where product/SKU names are column headers and cells are quantities
- Improved SKU matching using SKU, product name, aliases, and punctuation-insensitive normalized keys
- Stable exact SKU combination keys
- Campaign box standardization with upward-only rounding and `+2 cm` default tolerance
- XLSX report writer with required tabs and formatting
- Public `optimize_workbook(...)` workflow function
- FastAPI app with `/health`, `/optimize`, and `/openapi.json`
- API-key auth for `/optimize` via `Authorization: Bearer <key>` or `X-API-Key: <key>`
- Railway-ready Dockerfile using `PORT`
- `.env.example`

## Tests Passing

Current full test suite:

```text
97 passed, 1 warning
```

The warning is a local pytest cache permission warning in this Windows/Codex workspace and does not indicate a failing test.

## Tests Failing

No tests are currently failing.

## Known Parser Gaps

- Some real workbook SKU names still remain unmatched when headers are variants of SKU master names, for example free-gift prefixes, language variants, or renamed token pack columns.
- Region/country/state metadata may need campaign-specific inference when the order workbook only encodes geography in sheet names.
- Current XLSX reader is lightweight and supports the workbook shapes tested so far, but it is not a full Excel engine.
- Config fields like custom `max_carton_cm`, `dimensional_divisor`, and `packing_weight_uplift` are accepted but not fully applied throughout every algorithm yet.
- Wide-format order IDs are generated from sheet name and row number when no explicit order/backer ID column exists.

## Railway Deployment URL

Not set in this repo yet.

Once deployed, record the Railway public URL here, for example:

```text
https://your-service-name.up.railway.app
```

## API Endpoints

- `GET /health`
  - Public health check
  - Returns `{ "status": "ok" }`

- `GET /openapi.json`
  - Public OpenAPI schema

- `POST /optimize`
  - Requires API key if `BOX_OPTIMIZER_API_KEY` is set
  - Accepted auth headers:
    - `Authorization: Bearer <key>`
    - `X-API-Key: <key>`
  - Accepts multipart form fields:
    - `sku_master_file`
    - `orders_file`
    - optional `config_json`
  - Returns generated `.xlsx` workbook

## Next Recommended Task

Improve real-campaign SKU alias matching for unmatched wide-format headers, especially free gift prefixes, language variants, and product names that differ slightly from the SKU master.
