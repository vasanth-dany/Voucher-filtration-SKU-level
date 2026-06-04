# ============================================================
# VOUCHER FILTRATION — CAMPAIGN CONFIG
# !! THIS IS THE ONLY FILE YOU NEED TO UPDATE EACH CAMPAIGN !!
# ============================================================

# ------------------------------------------------------------
# 1. FILE PATHS — update to your new files each run
# ------------------------------------------------------------
PRICE_FILE   = 'inputs/SellerPriceTemplate_2026-05-18T034436_0800.xlsx'
CONTENT_FILE = 'inputs/Content_file_12_05_2026_.xlsx'
ZECOM_FILE   = 'inputs/zeCOM_Tracking_File_15_05.xlsx'
BAU_FILE     = None    # e.g. 'inputs/BAU_01_Jun_voucher_filteration.xlsx'
AM_EXCL_FILE = None    # e.g. 'inputs/AM_Exclusion.xlsx'

OUTPUT_FILE  = 'outputs/Voucher_Filtration_May2026.xlsx'

# ------------------------------------------------------------
# 2. ZECOM COLUMN POSITIONS — update if tracking file changes
#    (0-based index; check zeCOM MY sheet header row 4)
# ------------------------------------------------------------
ZECOM_COLS = {
    'style':         2,   # C  — Style# = ALU_NO
    'launch_date':  19,   # T  — Launch Date
    'status_laz':   21,   # V  — Lazada status YES/NO
    'price':        46,   # AU — Price (MY RRP)
    'special_price':47,   # AV — Special Price (MY EC SRP)
    'disc_pct':     48,   # AW — Discount %
    'exclusion':    49,   # AX — Exclusion (main campaign column)
}

# ------------------------------------------------------------
# 3. ELIGIBILITY RULES — rarely changes
# ------------------------------------------------------------
MIN_PRICE_THRESHOLD = 39   # flag if Price OR Special Price is below this

PRE_FILTER_RULES = {
    'status_lazada_must_be_yes': True,   # exclude if Status_Lazada != YES
    'status_bau_must_be_yes':    True,   # exclude if BAU status != YES (only if BAU file provided)
    'future_live_date':          True,   # exclude if Live Date > today (only if BAU file provided)
    'min_price_threshold':       True,   # exclude if Price OR Special Price < MIN_PRICE_THRESHOLD
}

# ------------------------------------------------------------
# 4. CAMPAIGN RULES — update match_values when zeCOM wording changes
# ------------------------------------------------------------
#
# match_type options:
#   "contains" — exclusion value contains any of match_values (case-insensitive)
#   "exact"    — exclusion value exactly equals one of match_values (case-insensitive)
#
# excludes: list of substrings that DISQUALIFY even if match_values matched
#   e.g. excludes: ["No platform VC"] will block "20% VC - No platform VC"
#
# To ADD a new campaign: copy one block below and fill in the details
# To REMOVE a campaign:  comment out the block with #
# ------------------------------------------------------------

CAMPAIGNS = [

    {
        'name'        : '20% NMS (All 20% VC Remark)',
        'output_col'  : '20%_NMS (All 20% VC Remark)',
        'sheet_name'  : '20% NMS',
        'tab_color'   : '#7C3AED',
        'cell_color'  : '#EDE9FE',
        'font_color'  : '#5B21B6',
        'match_type'  : 'contains',
        'match_values': ['20% VC'],          # ← update if wording changes
        'excludes'    : [],                  # e.g. ['No platform VC'] to block variants
    },

    {
        'name'        : '30% NMS (Max 30% VC Remark)',
        'output_col'  : '30%_NMS (Max 30% VC Remark)',
        'sheet_name'  : '30% NMS',
        'tab_color'   : '#1D4ED8',
        'cell_color'  : '#DBEAFE',
        'font_color'  : '#1E40AF',
        'match_type'  : 'contains',
        'match_values': ['MAX 30% VC', '30% VC'],   # ← update if wording changes
        'excludes'    : [],
    },

    {
        'name'        : '50% NMS (Open for All + Open for All 10 days)',
        'output_col'  : '50%_NMS (Open for All + 10days)',
        'sheet_name'  : '50% NMS',
        'tab_color'   : '#065F46',
        'cell_color'  : '#D1FAE5',
        'font_color'  : '#065F46',
        'match_type'  : 'exact',
        'match_values': [                    # ← update if wording changes
            'Open for all',
            'OPEN for 10days up to 50%',
        ],
        'excludes'    : [],
    },

    # ── TEMPLATE: copy-paste to add a new campaign ──────────
    # {
    #     'name'        : 'NEW CAMPAIGN NAME',
    #     'output_col'  : 'XX%_NMS (Label)',
    #     'sheet_name'  : 'XX% NMS',
    #     'tab_color'   : '#HEX',
    #     'cell_color'  : '#HEX',
    #     'font_color'  : '#HEX',
    #     'match_type'  : 'contains',   # or 'exact'
    #     'match_values': ['EXCLUSION VALUE 1', 'EXCLUSION VALUE 2'],
    #     'excludes'    : [],
    # },

]

# ------------------------------------------------------------
# 5. OUTPUT STYLING — rarely changes
# ------------------------------------------------------------
STYLING = {
    'header_bg'     : '#1E293B',
    'eligible_bg'   : '#DCFCE7',
    'eligible_font' : '#166534',
    'flagged_bg'    : '#FEE2E2',
    'flagged_font'  : '#991B1B',
    'am_excl_bg'    : '#FEF3C7',
    'am_excl_font'  : '#92400E',
    'mismatch_bg'   : '#FEE2E2',
    'mismatch_font' : '#991B1B',
    'match_bg'      : '#DCFCE7',
    'match_font'    : '#166534',
}
