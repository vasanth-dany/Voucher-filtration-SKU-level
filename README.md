# Voucher Filtration Automation

Automated voucher eligibility filtering for Lazada campaigns.
Processes Price/Stock, Content, zeCOM Tracking, BAU Voucher, and AM Exclusion files
to produce a filtered Excel output with campaign-tagged SKUs.

---

## Folder Structure

```
voucher-filtration/
├── inputs/                  ← put your Excel files here (not pushed to GitHub)
├── outputs/                 ← filtered output goes here (not pushed to GitHub)
├── config.py                ← !! ONLY FILE YOU CHANGE EACH CAMPAIGN !!
├── voucher_filtration.py    ← main script (do not edit)
├── requirements.txt         ← Python dependencies
├── .gitignore
└── README.md
```

---

## First-Time Setup

```bash
# 1. Clone the repo
git clone https://github.com/YOUR_USERNAME/voucher-filtration.git
cd voucher-filtration

# 2. Install dependencies
pip install -r requirements.txt
```

---

## Every Campaign Run

### Step 1 — Drop files into `inputs/`
Copy your new Excel files into the `inputs/` folder:
- `SellerPriceTemplate_....xlsx`
- `Content_file_....xlsx`
- `zeCOM_Tracking_File_....xlsx`
- `BAU_....xlsx` *(optional)*
- `AM_Exclusion.xlsx` *(optional)*

### Step 2 — Update `config.py`
Open `config.py` and update:
1. **File paths** (Section 1) — point to your new files
2. **Campaign rules** (Section 4) — update `match_values` if zeCOM Exclusion wording changed

```python
# Example: adding a new campaign
{
    'name'        : '40% NMS (Max 40% VC)',
    'output_col'  : '40%_NMS (Max 40% VC)',
    'sheet_name'  : '40% NMS',
    'tab_color'   : '#B45309',
    'cell_color'  : '#FEF3C7',
    'font_color'  : '#92400E',
    'match_type'  : 'contains',
    'match_values': ['MAX 40% VC', '40% VC'],
    'excludes'    : [],
},
```

### Step 3 — Run
```bash
python voucher_filtration.py
```
Output is saved to `outputs/`.

### Step 4 — Save rule changes to GitHub
```bash
git add config.py
git commit -m "Jun 2026 BAU — added 40% NMS, updated 50% NMS wording"
git push
```

---

## Campaign Rule Reference

| match_type | Behaviour |
|---|---|
| `contains` | Exclusion value contains any of `match_values` (case-insensitive) |
| `exact`    | Exclusion value exactly equals one of `match_values` (case-insensitive) |

Use `excludes` to block variants even if `match_values` matched:
```python
'match_values': ['20% VC'],
'excludes'    : ['No platform VC'],   # blocks "20% VC - No platform VC"
```

---

## Eligibility Rules

A SKU is **ELIGIBLE** only if ALL of these pass:

| Rule | Config key |
|---|---|
| Status Lazada (zeCOM col V) = YES | `status_lazada_must_be_yes` |
| Status BAU = YES *(if BAU file provided)* | `status_bau_must_be_yes` |
| Live Date ≤ today *(if BAU file provided)* | `future_live_date` |
| Price AND Special Price ≥ threshold (default 39) | `min_price_threshold` |

---

## zeCOM Column Reference

| Column | Excel Col | Index | Field |
|---|---|---|---|
| Style# | C | 2 | ALU_NO |
| Launch Date | T | 19 | Launch_Date |
| Lazada Status | V | 21 | Status_Lazada |
| Price | AU | 46 | Price |
| Special Price | AV | 47 | Special_Price |
| Disc % | AW | 48 | Disc_Pct |
| Exclusion | AX | 49 | Exclusion |

> If zeCOM file structure changes, update `ZECOM_COLS` in `config.py`.

---

## Output Sheets

| Sheet | Contents |
|---|---|
| All SKUs | Full dataset with all columns |
| 20% NMS | ELIGIBLE SKUs matching 20% VC rule |
| 30% NMS | ELIGIBLE SKUs matching 30% VC rule |
| 50% NMS | ELIGIBLE SKUs matching Open for All / 10days rule |
| Price Mismatch | ELIGIBLE SKUs where Price ≠ Special Price |
| Summary | Count per campaign + criteria |
