"""
Voucher Filtration Automation — Streamlit Web App
==================================================
Run: streamlit run app.py
"""

import streamlit as st
import pandas as pd
import openpyxl
import xlsxwriter
import io
import warnings
from datetime import date, datetime

warnings.filterwarnings('ignore')

# ── Page config ───────────────────────────────────────────────
st.set_page_config(
    page_title="Voucher Filtration",
    page_icon="🏷️",
    layout="wide"
)

# ── Styles ────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #0f1117; }
    .block-container { padding-top: 2rem; }
    .stButton > button {
        background-color: #6c63ff;
        color: white;
        font-weight: 700;
        border: none;
        padding: 0.6rem 2rem;
        border-radius: 6px;
        font-size: 16px;
        width: 100%;
    }
    .stButton > button:hover { background-color: #8076ff; }
    .metric-card {
        background: #1e293b;
        border-radius: 10px;
        padding: 1rem 1.2rem;
        margin-bottom: 0.5rem;
        border-left: 4px solid;
    }
    .upload-section {
        background: #1a1a2e;
        border-radius: 10px;
        padding: 1.2rem;
        margin-bottom: 1rem;
        border: 1px solid #2d2d44;
    }
</style>
""", unsafe_allow_html=True)

TODAY = date.today()

# ── Header ────────────────────────────────────────────────────
st.title("🏷️ Voucher Filtration Automation")
st.caption(f"Lazada Campaign SKU Filtering  ·  Run date: {TODAY}")
st.divider()

# ============================================================
# SIDEBAR — Campaign Rules (editable every run)
# ============================================================
with st.sidebar:
    st.header("⚙️ Campaign Rules")
    st.caption("Update these each campaign run")

    st.subheader("Eligibility Threshold")
    min_price = st.number_input(
        "Min Price / Special Price (flag if below)",
        value=39, min_value=0, step=1
    )

    st.subheader("Active Campaigns")
    st.caption("Tick to include, edit match values if wording changes")

    # 20% NMS
    run_20 = st.checkbox("20% NMS", value=True)
    match_20 = st.text_area(
        "20% NMS — match values (one per line, contains)",
        value="20% VC",
        height=70,
        disabled=not run_20
    )
    exclude_20 = st.text_input(
        "20% NMS — exclude if contains",
        value="",
        placeholder="e.g. No platform VC",
        disabled=not run_20
    )

    st.divider()

    # 30% NMS
    run_30 = st.checkbox("30% NMS", value=True)
    match_30 = st.text_area(
        "30% NMS — match values (one per line, contains)",
        value="MAX 30% VC\n30% VC",
        height=70,
        disabled=not run_30
    )
    exclude_30 = st.text_input(
        "30% NMS — exclude if contains",
        value="",
        disabled=not run_30
    )

    st.divider()

    # 50% NMS
    run_50 = st.checkbox("50% NMS", value=True)
    match_50 = st.text_area(
        "50% NMS — match values (one per line, exact)",
        value="Open for all\nOPEN for 10days up to 50%",
        height=85,
        disabled=not run_50
    )
    exclude_50 = st.text_input(
        "50% NMS — exclude if contains",
        value="",
        disabled=not run_50
    )

    st.divider()
    st.subheader("zeCOM Column Positions")
    st.caption("Only change if file structure changed")
    col_style  = st.number_input("Style# (ALU_NO) index", value=2)
    col_laz    = st.number_input("Status Lazada index",    value=21)
    col_price  = st.number_input("Price (AU) index",       value=46)
    col_sp     = st.number_input("Special Price (AV) index", value=47)
    col_disc   = st.number_input("Disc% (AW) index",       value=48)
    col_excl   = st.number_input("Exclusion (AX) index",   value=49)

# ── Build campaign config from sidebar ───────────────────────
CAMPAIGNS = []
if run_20:
    CAMPAIGNS.append({
        'name': '20% NMS (All 20% VC Remark)', 'col': '20%_NMS',
        'sheet': '20% NMS', 'tab': '#7C3AED', 'bg': '#EDE9FE', 'fg': '#5B21B6',
        'type': 'contains',
        'match': [v.strip() for v in match_20.splitlines() if v.strip()],
        'excl':  [v.strip() for v in exclude_20.split(',') if v.strip()],
    })
if run_30:
    CAMPAIGNS.append({
        'name': '30% NMS (Max 30% VC Remark)', 'col': '30%_NMS',
        'sheet': '30% NMS', 'tab': '#1D4ED8', 'bg': '#DBEAFE', 'fg': '#1E40AF',
        'type': 'contains',
        'match': [v.strip() for v in match_30.splitlines() if v.strip()],
        'excl':  [v.strip() for v in exclude_30.split(',') if v.strip()],
    })
if run_50:
    CAMPAIGNS.append({
        'name': '50% NMS (Open for All + 10days)', 'col': '50%_NMS',
        'sheet': '50% NMS', 'tab': '#065F46', 'bg': '#D1FAE5', 'fg': '#065F46',
        'type': 'exact',
        'match': [v.strip() for v in match_50.splitlines() if v.strip()],
        'excl':  [v.strip() for v in exclude_50.split(',') if v.strip()],
    })

ZECOM_COLS = {
    'style': int(col_style), 'status_laz': int(col_laz),
    'price': int(col_price), 'special_price': int(col_sp),
    'disc_pct': int(col_disc), 'exclusion': int(col_excl),
}

# ============================================================
# FILE UPLOADS
# ============================================================
st.subheader("📂 Upload Input Files")

col1, col2, col3 = st.columns(3)
with col1:
    st.markdown("**1. Price / Stock File** *(required)*")
    price_file = st.file_uploader("SellerPriceTemplate", type=['xlsx','xls'], key='price')
with col2:
    st.markdown("**2. Content File** *(required)*")
    content_file = st.file_uploader("Content_file", type=['xlsx','xls'], key='content')
with col3:
    st.markdown("**3. zeCOM Tracking File** *(required)*")
    zecom_file = st.file_uploader("zeCOM_Tracking_File", type=['xlsx','xls'], key='zecom')

col4, col5 = st.columns(2)
with col4:
    st.markdown("**4. BAU Voucher File** *(optional)*")
    bau_file = st.file_uploader("BAU_...xlsx", type=['xlsx','xls'], key='bau')
with col5:
    st.markdown("**5. AM Exclusion File** *(optional)*")
    am_file = st.file_uploader("AM_Exclusion.xlsx", type=['xlsx','xls'], key='am')

st.divider()

# ============================================================
# PROCESSING FUNCTIONS
# ============================================================
def normalise_ean(v):
    return str(v).strip().split('.')[0].strip()

def matches_campaign(excl_val, camp):
    e  = str(excl_val).strip() if pd.notna(excl_val) else ''
    el = e.lower()
    mv = [v.lower() for v in camp['match']]
    ex = [v.lower() for v in camp['excl']]
    matched = any(m in el for m in mv) if camp['type'] == 'contains' else el in mv
    if not matched: return False
    if any(x in el for x in ex if x): return False
    return True

def check_eligibility(row, min_p):
    reasons = []
    sl = row.get('Status_Lazada')
    if pd.notna(sl) and str(sl).strip().upper() != 'YES':
        reasons.append('Status Lazada ≠ YES')
    sb = row.get('Status_BAU')
    if pd.notna(sb) and str(sb).strip().upper() != 'YES':
        reasons.append('Status BAU ≠ YES')
    ld = row.get('Live_Date')
    if pd.notna(ld):
        ld_d = ld.date() if isinstance(ld, datetime) else ld
        if isinstance(ld_d, date) and ld_d > TODAY:
            reasons.append('Future Live Date')
    try:
        pv  = float(row['Price'])         if pd.notna(row.get('Price'))         else None
        spv = float(row['Special_Price'])  if pd.notna(row.get('Special_Price')) else None
        if (pv is not None and pv < min_p) or (spv is not None and spv < min_p):
            reasons.append(f'Price or Special Price < {min_p}')
    except: pass
    return 'ELIGIBLE' if not reasons else ' | '.join(reasons)

def run_filtration(price_f, content_f, zecom_f, bau_f, am_f, campaigns, zecom_cols, min_p):
    logs = []
    def log(msg): logs.append(msg)

    # STEP 1: Price file
    log("▶ Loading Price/Stock file...")
    price_df = pd.read_excel(price_f, sheet_name=0, header=0)
    price_df = price_df.rename(columns={'SellerSku':'SellerSKU','Name':'Product_Name'})
    price_df['SellerSKU']     = price_df['SellerSKU'].astype(str).str.strip()
    price_df['Product_ID']    = price_df.get('ProductId', '')
    price_df['Status_Active'] = 'active'
    price_df['Quantity']      = price_df.get('Quantity', 0)
    price_df.drop(columns=[c for c in ['Price','SalePrice','Special_Price'] if c in price_df.columns], inplace=True)
    log(f"   ✓ {len(price_df):,} SKUs loaded")

    # STEP 2: Content file
    log("▶ Mapping ALU_NO via Content file...")
    content_df = pd.read_excel(content_f, sheet_name='content')[['Color_No','EAN']].copy()
    content_df['EAN_norm'] = content_df['EAN'].astype(str).str.strip().str.split('.').str[0]
    content_df = content_df.drop_duplicates(subset='EAN_norm')
    price_df['EAN_norm'] = price_df['SellerSKU'].apply(normalise_ean)
    price_df = price_df.merge(content_df[['EAN_norm','Color_No']], on='EAN_norm', how='left')
    price_df.rename(columns={'Color_No':'ALU_NO'}, inplace=True)
    log(f"   ✓ {price_df['ALU_NO'].notna().sum():,} / {len(price_df):,} SKUs matched to ALU_NO")

    # STEP 3: zeCOM
    log("▶ Loading zeCOM Tracking file...")
    wb_z = openpyxl.load_workbook(zecom_f, read_only=True, data_only=True)
    ws_z = wb_z['MY']
    records = []
    for row in ws_z.iter_rows(min_row=5, values_only=True):
        style = row[zecom_cols['style']] if len(row) > zecom_cols['style'] else None
        if not style: continue
        def g(k): i=zecom_cols.get(k); return row[i] if i and len(row)>i else None
        records.append({'ALU_NO':str(style).strip(),'Status_Lazada':g('status_laz'),
            'Price':g('price'),'Special_Price':g('special_price'),
            'Disc_Pct':g('disc_pct'),'Exclusion':g('exclusion')})
    wb_z.close()
    zecom_df = pd.DataFrame(records).drop_duplicates(subset='ALU_NO')
    log(f"   ✓ {len(zecom_df):,} styles loaded")

    excl_vals = zecom_df['Exclusion'].value_counts().to_dict()
    log(f"   Exclusion values: {excl_vals}")

    main_df = price_df.merge(zecom_df, on='ALU_NO', how='left')

    # STEP 4: BAU
    main_df['Live_Date'] = None; main_df['Status_BAU'] = None
    if bau_f:
        log("▶ Loading BAU Voucher file...")
        bau_raw = pd.read_excel(bau_f, sheet_name='template', header=0)
        bau_df  = bau_raw.iloc[3:].copy(); bau_df.columns = bau_raw.columns
        bau_df  = bau_df.drop_duplicates(subset='SellerSKU')
        bau_df['SellerSKU'] = bau_df['SellerSKU'].astype(str).str.strip()
        live_col   = next((c for c in bau_df.columns if 'live' in str(c).lower()), None)
        status_col = next((c for c in bau_df.columns if str(c).strip().lower()=='status'), None)
        cols = ['SellerSKU']+([live_col] if live_col else [])+([status_col] if status_col else [])
        bau_sub = bau_df[cols].copy()
        if live_col:   bau_sub.rename(columns={live_col:'Live_Date'}, inplace=True)
        if status_col: bau_sub.rename(columns={status_col:'Status_BAU'}, inplace=True)
        main_df = main_df.merge(bau_sub, on='SellerSKU', how='left', suffixes=('','_bau'))
        for col in ['Live_Date','Status_BAU']:
            if f'{col}_bau' in main_df.columns:
                main_df[col] = main_df[f'{col}_bau'].combine_first(main_df[col])
                main_df.drop(columns=[f'{col}_bau'], inplace=True)
        log(f"   ✓ BAU merged")

    # STEP 5: AM Exclusion
    am_set = set()
    if am_f:
        log("▶ Loading AM Exclusion file...")
        am_df = pd.read_excel(am_f)
        am_set = set(am_df.iloc[:,0].dropna().astype(str).str.strip())
        log(f"   ✓ {len(am_set):,} excluded ALU_NOs")
    main_df['AM_Exclusion'] = main_df['ALU_NO'].apply(
        lambda x: 'EXCLUDE' if pd.notna(x) and str(x).strip() in am_set else '')

    # STEP 6: Eligibility
    log("▶ Calculating eligibility...")
    main_df['Eligibility'] = main_df.apply(lambda r: check_eligibility(r, min_p), axis=1)
    elig_count = (main_df['Eligibility']=='ELIGIBLE').sum()
    log(f"   ✓ ELIGIBLE: {elig_count:,} | Flagged: {len(main_df)-elig_count:,}")

    # STEP 7: Price check
    def price_check(row):
        try:
            p=row.get('Price'); sp=row.get('Special_Price')
            if pd.isna(p) or pd.isna(sp): return 'NO DATA'
            return 'MATCH' if float(p)==float(sp) else 'MISMATCH'
        except: return 'NO DATA'
    main_df['Price_vs_SpecialPrice'] = main_df.apply(price_check, axis=1)

    # STEP 8: Campaigns
    log("▶ Applying campaign rules...")
    for camp in campaigns:
        main_df[camp['col']] = main_df.apply(
            lambda r: 'YES' if (r['Eligibility']=='ELIGIBLE'
                and r['AM_Exclusion']!='EXCLUDE'
                and matches_campaign(r.get('Exclusion'), camp)) else '', axis=1)
        count = (main_df[camp['col']]=='YES').sum()
        log(f"   {camp['name']}: {count:,} SKUs")

    # Output columns
    camp_cols = [c['col'] for c in campaigns]
    OUT_COLS = ['Product_ID','Product_Name','SellerSKU','ALU_NO',
        'Status_Active','Quantity','Price','Special_Price','Disc_Pct',
        'Live_Date','Status_BAU','Status_Lazada',
        'Price_vs_SpecialPrice','Exclusion','AM_Exclusion','Eligibility'] + camp_cols
    for c in OUT_COLS:
        if c not in main_df.columns: main_df[c] = ''
    out_df = main_df[OUT_COLS].copy()

    return out_df, logs, campaigns, elig_count

def build_excel(out_df, campaigns, elig_count):
    output = io.BytesIO()
    wb = xlsxwriter.Workbook(output, {'strings_to_numbers': False, 'in_memory': True})

    base = {'font_name':'Arial','font_size':9,'valign':'vcenter','border':1,'border_color':'#D1D5DB'}
    def mf(props): return wb.add_format({**base, **props})

    FMT = {
        'header'  : mf({'bold':True,'bg_color':'#1E293B','font_color':'#FFFFFF','text_wrap':True,'align':'center'}),
        'data'    : mf({'bg_color':'#FFFFFF'}),
        'alt'     : mf({'bg_color':'#F8FAFC'}),
        'eligible': mf({'bg_color':'#DCFCE7','font_color':'#166534','bold':True}),
        'flagged' : mf({'bg_color':'#FEE2E2','font_color':'#991B1B','bold':True}),
        'amber'   : mf({'bg_color':'#FEF3C7','font_color':'#92400E','bold':True}),
        'mismatch': mf({'bg_color':'#FEE2E2','font_color':'#991B1B'}),
        'match'   : mf({'bg_color':'#DCFCE7','font_color':'#166534'}),
    }
    CFMT = {c['col']: mf({'bg_color':c['bg'],'font_color':c['fg'],'bold':True,'align':'center'}) for c in campaigns}

    cols = list(out_df.columns)
    widths = {'Product_ID':14,'Product_Name':32,'SellerSKU':18,'ALU_NO':14,
        'Status_Active':11,'Quantity':10,'Price':12,'Special_Price':14,'Disc_Pct':10,
        'Live_Date':14,'Status_BAU':12,'Status_Lazada':13,
        'Price_vs_SpecialPrice':18,'Exclusion':30,'AM_Exclusion':14,'Eligibility':30}

    def write_ws(name, df, tab):
        ws = wb.add_worksheet(name[:31])
        ws.set_tab_color(tab); ws.freeze_panes(1,0)
        ws.autofilter(0,0,len(df),len(cols)-1); ws.set_row(0,32)
        for ci,col in enumerate(cols):
            ws.write(0,ci,col,FMT['header'])
            ws.set_column(ci,ci,widths.get(col,16))
        for ri,(_,row) in enumerate(df.iterrows(),1):
            bf = FMT['alt'] if ri%2==0 else FMT['data']
            for ci,col in enumerate(cols):
                val = row[col]
                if val is None or (isinstance(val,float) and pd.isna(val)): val=''
                elif isinstance(val,(datetime,date)): val=str(val)[:10]
                if col=='Eligibility': fmt=FMT['eligible'] if val=='ELIGIBLE' else FMT['flagged']
                elif col=='AM_Exclusion' and val=='EXCLUDE': fmt=FMT['amber']
                elif col=='Price_vs_SpecialPrice': fmt=FMT['match'] if val=='MATCH' else (FMT['mismatch'] if val=='MISMATCH' else bf)
                elif col in CFMT and val=='YES': fmt=CFMT[col]
                else: fmt=bf
                ws.write(ri,ci,val,fmt)

    write_ws('All SKUs', out_df, '#1E293B')
    for camp in campaigns:
        write_ws(camp['sheet'], out_df[out_df[camp['col']]=='YES'], camp['tab'])

    mismatch = out_df[(out_df['Eligibility']=='ELIGIBLE')&(out_df['Price_vs_SpecialPrice']=='MISMATCH')]
    write_ws('Price Mismatch', mismatch, '#B91C1C')

    # Summary
    ws_s = wb.add_worksheet('Summary')
    ws_s.set_tab_color('#0F172A')
    ws_s.set_column(0,0,40); ws_s.set_column(1,1,14); ws_s.set_column(2,2,52)
    hf = wb.add_format({'font_name':'Arial','font_size':11,'bold':True,'bg_color':'#0F172A','font_color':'#FFFFFF','border':1,'align':'center'})
    ws_s.write(0,0,'Campaign',hf); ws_s.write(0,1,'SKU Count',hf); ws_s.write(0,2,'Criteria',hf)
    rows_s = [('All SKUs',len(out_df),'All rows'),
              ('ELIGIBLE',elig_count,'No pre-filter flags'),
              ('Flagged',len(out_df)-elig_count,'Status/Price/Date flags'),
              ('AM Excluded',(out_df['AM_Exclusion']=='EXCLUDE').sum(),'In AM Exclusion list')] + \
             [(c['name'],(out_df[c['col']]=='YES').sum(),f"Exclusion {c['type']} {', '.join(c['match'])}") for c in campaigns] + \
             [('Price Mismatch',len(mismatch),'Price ≠ Special Price')]
    for ri,(label,count,crit) in enumerate(rows_s,1):
        rf=wb.add_format({'font_name':'Arial','font_size':10,'border':1,'bg_color':'#F8FAFC'})
        nf=wb.add_format({'font_name':'Arial','font_size':10,'border':1,'bold':True,'align':'center','bg_color':'#F8FAFC'})
        ws_s.write(ri,0,label,rf); ws_s.write(ri,1,count,nf); ws_s.write(ri,2,crit,rf)

    wb.close()
    output.seek(0)
    return output

# ============================================================
# RUN BUTTON
# ============================================================
ready = price_file and content_file and zecom_file

if not ready:
    st.info("👆 Upload the 3 required files above to get started")

if st.button("▶  Run Voucher Filtration", disabled=not ready):
    with st.spinner("Processing..."):
        try:
            progress = st.progress(0, text="Loading files...")
            out_df, logs, camps, elig_count = run_filtration(
                price_file, content_file, zecom_file, bau_file, am_file,
                CAMPAIGNS, ZECOM_COLS, min_price
            )
            progress.progress(80, text="Building Excel output...")
            excel_bytes = build_excel(out_df, camps, elig_count)
            progress.progress(100, text="Done!")

            # ── Results ──────────────────────────────────────
            st.success("✅ Filtration complete!")
            st.divider()
            st.subheader("📊 Results Summary")

            camp_cols = [c['col'] for c in CAMPAIGNS]
            metrics_cols = st.columns(3 + len(CAMPAIGNS))

            with metrics_cols[0]:
                st.metric("Total SKUs", f"{len(out_df):,}")
            with metrics_cols[1]:
                st.metric("✅ Eligible", f"{elig_count:,}")
            with metrics_cols[2]:
                st.metric("❌ Flagged", f"{len(out_df)-elig_count:,}")
            for i, camp in enumerate(CAMPAIGNS):
                with metrics_cols[3+i]:
                    count = (out_df[camp['col']]=='YES').sum()
                    st.metric(camp['sheet'], f"{count:,}")

            # Exclusion breakdown
            st.divider()
            st.subheader("📋 Exclusion Values Found")
            excl_counts = out_df['Exclusion'].value_counts().reset_index()
            excl_counts.columns = ['Exclusion Value', 'SKU Count']
            st.dataframe(excl_counts, use_container_width=True, hide_index=True)

            # Log
            with st.expander("📝 Processing Log"):
                for line in logs:
                    st.text(line)

            # Download
            st.divider()
            st.download_button(
                label="⬇️  Download Filtered Excel",
                data=excel_bytes,
                file_name=f"Voucher_Filtration_{TODAY}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ Error: {e}")
            st.exception(e)
