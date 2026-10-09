import pandas as pd
import argparse
import sys

# Logika percabangan khusus (PDC -> AX -> DC)
decision_tree = {
    'P11AA': {'11CX': '11031', 'default': '11032'},
    'D11AP': {'11PDX': '11742', 'default': '11741'},
    'D11AD': {'11PCX': '11641', 'default': '11642'},
    'P11AR': {'11PEX': '11181', '11PFX': '11181', 'default': '11182'}
}

# Mapping 1:1 untuk PDC yang tidak memiliki percabangan (default dari PDF)
flat_mapping = {
    'P11AG': '11141',
    'P11AH': '11071',
    'P11AB': '11041',
    'P11AC': '11021',
    'P11AK': '11091',
    'P11AE': '11131',
    'P11AF': '11051',
    'P11AJ': '11151',
    'P11AP': '11161',
    'P11AN': '11101',
    'P11AL': '11081',
    'P11AM': '11121',
    'P11AQ': '11171',
    'D11AB': '11621',
    'D11AN': '11731',
    'D11AM': '11721',
    'D11AL': '11711',
    'D11AA': '11611',
    'D11AC': '11631',
    'D11AF': '11661',
    'D11AU': '11791',
    'D11AK': '11701',
    'D11AH': '11681',
    'D11AS': '11771',
    'D11AV': '11801',
    'D11AJ': '11692',
    'D11AQ': '11751',
    'D11AR': '11761',
    'D11AE': '11651',
    'D11AG': '11671',
    'D11AT': '11781'
}

master_file = r'C:\Users\PUSBIKES-KEMKES\Downloads\mdc\DRG versi September 2025-Final Publish & DRG Master.xlsx'
try:
    df_master = pd.read_excel(master_file, sheet_name='Rencana Usulan per 29092026', dtype=str)
except Exception as e:
    print(f"Error reading DRG Master: {e}")
    sys.exit(1)

def get_drg_info(dc_code):
    """Mencari semua DRG turunan dari sebuah DC di DRG Master"""
    # Cari di kolom dc update
    matches = df_master[df_master['dc update'].str.replace(r'\.0$', '', regex=True).str.strip() == dc_code]
    if matches.empty:
        # Cari di kolom dc lama jika tidak ada di update
        matches = df_master[df_master['dc'].str.replace(r'\.0$', '', regex=True).str.strip() == dc_code]
    
    if matches.empty:
        # Fallback jika kode belum tercatat di Master
        return [{"drg": "Tidak Terdaftar di Master", "desc": "Tidak Terdaftar di Master"}]
        
    results = []
    for _, row in matches.iterrows():
        drg = row.get('drg update')
        desc = row.get('desc drg update')
        if pd.isna(drg):
            drg = row.get('drg')
            desc = row.get('desc drg')
        results.append({'drg': str(drg), 'desc': str(desc)})
    
    return [dict(t) for t in {tuple(d.items()) for d in results}] # remove duplicates

def run_test(pdc, ax=None):
    pdc = pdc.strip().upper()
    ax = ax.strip().upper() if ax else None
    
    print(f"\n============================================")
    print(f"TESTING ALGORITHM - MDC 11 DRAW.IO LOGIC")
    print(f"============================================")
    print(f"Input PDC (Cluster) : {pdc}")
    print(f"Input AX (Kondisi)  : {ax if ax else 'TIDAK ADA'}")
    
    dc_result = None
    
    if pdc in decision_tree:
        if ax in decision_tree[pdc]:
            dc_result = decision_tree[pdc][ax]
            print(f"[*] Logic Match: PDC {pdc} + AX {ax} -> Menghasilkan DC {dc_result}")
        else:
            dc_result = decision_tree[pdc]['default']
            print(f"[*] Logic Match: PDC {pdc} TANPA AX SPESIFIK -> Menghasilkan DC {dc_result}")
    elif pdc in flat_mapping:
        dc_result = flat_mapping[pdc]
        print(f"[*] Logic Match: PDC {pdc} adalah Single-Branch -> Menghasilkan DC {dc_result}")
    else:
        print(f"[!] ERROR: PDC '{pdc}' tidak dikenali di dalam Draw.io MDC 11.")
        return
        
    print(f"\n>> DC DITEMUKAN: {dc_result}")
    print(f">> MENCARI DRG KE MASTER DRG...")
    
    drgs = get_drg_info(dc_result)
    for i, item in enumerate(drgs, 1):
        print(f"   - DRG {i}: {item['drg']} ({item['desc']})")
    print("============================================\n")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Test MDC 11 Draw.io Logic")
    parser.add_argument("--pdc", type=str, help="Kode PDC (contoh: P11AA)", required=False)
    parser.add_argument("--ax", type=str, help="Kode AX opsional (contoh: 11CX)", required=False)
    args = parser.parse_args()
    
    if args.pdc:
        run_test(args.pdc, args.ax)
    else:
        print("Mode Interaktif - Ketik 'exit' untuk keluar.")
        while True:
            pdc_input = input("Masukkan Kode PDC (contoh: D11AP): ")
            if pdc_input.lower() == 'exit': break
            ax_input = input("Masukkan Kode AX (kosongkan dan tekan Enter jika tidak ada): ")
            if ax_input.lower() == 'exit': break
            run_test(pdc_input, ax_input)
