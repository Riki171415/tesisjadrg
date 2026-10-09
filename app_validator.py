import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="MDC 11 Logic Validator", layout="wide", page_icon="🔍")

# ==========================================
# 1. CORE LOGIC & MAPPING
# ==========================================
# Draw.io Logic Translation
decision_tree = {
    'P11AA': {'11CX': '11031', 'default': '11032'},
    'D11AP': {'11PDX': '11742', 'default': '11741'},
    'D11AD': {'11PCX': '11641', 'default': '11642'},
    'P11AR': {'11PEX': '11181', '11PFX': '11181', 'default': '11182'}
}

flat_mapping = {
    'P11AG': '11141', 'P11AH': '11071', 'P11AB': '11041', 'P11AC': '11021',
    'P11AK': '11091', 'P11AE': '11131', 'P11AF': '11051', 'P11AJ': '11151',
    'P11AP': '11161', 'P11AN': '11101', 'P11AL': '11081', 'P11AM': '11121',
    'P11AQ': '11171', 'D11AB': '11621', 'D11AN': '11731', 'D11AM': '11721',
    'D11AL': '11711', 'D11AA': '11611', 'D11AC': '11631', 'D11AF': '11661',
    'D11AU': '11791', 'D11AK': '11701', 'D11AH': '11681', 'D11AS': '11771',
    'D11AV': '11801', 'D11AJ': '11692', 'D11AQ': '11751', 'D11AR': '11761',
    'D11AE': '11651', 'D11AG': '11671', 'D11AT': '11781'
}

def get_expected_dc(pdc):
    """Mengembalikan daftar kemungkinan DC untuk PDC tertentu berdasarkan Draw.io"""
    if pd.isna(pdc) or not str(pdc).strip():
        return []
    pdc = str(pdc).strip().upper()
    if pdc in flat_mapping:
        return [flat_mapping[pdc]]
    if pdc in decision_tree:
        # Mengembalikan semua kemungkinan DC dari cabang AX
        return list(set(decision_tree[pdc].values()))
    return []

# ==========================================
# 2. UI LAYOUT
# ==========================================
st.title("🏥 Sistem Validasi Logika MDC 11")
st.markdown("""
Aplikasi ini secara otomatis memvalidasi file Excel usulan terhadap standar logika percabangan di **Draw.io**.
Silakan unggah file Excel yang sudah di-generate (contoh: `MDC_11_Terstruktur_DC_V4.xlsx`) atau file *Usulan* yang memiliki kolom `PDC_Baru` dan `DC_Baru`.
""")

st.sidebar.header("📁 Upload File")
uploaded_excel = st.sidebar.file_uploader("Upload File Excel (.xlsx)", type=["xlsx"])
uploaded_pdf = st.sidebar.file_uploader("Upload File Draw.io (.pdf) [Referensi/Opsional]", type=["pdf"])

if uploaded_pdf:
    st.sidebar.success("PDF berhasil dimuat sebagai referensi!")

if uploaded_excel:
    with st.spinner('Menganalisis logika data...'):
        try:
            xls = pd.ExcelFile(uploaded_excel)
            sheets = xls.sheet_names
            
            selected_sheets = st.multiselect("Pilih Sheet untuk dianalisis:", sheets, default=[s for s in sheets if 'ICD' in s.upper() and ('10' in s or '9' in s)])
            
            if not selected_sheets:
                st.warning("Silakan pilih minimal satu sheet.")
            else:
                results = []
                error_count = 0
                valid_count = 0
                has_dc_col = False
                
                for sheet in selected_sheets:
                    df = pd.read_excel(xls, sheet_name=sheet)
                    
                    # Coba deteksi kolom yang relevan (Fleksibel)
                    pdc_col = next((c for c in df.columns if str(c).strip().lower() in ['pdc_baru', 'new cluster code', 'pdc', 'cluster code']), None)
                    dc_col = next((c for c in df.columns if str(c).strip().lower() in ['dc_baru', 'dc', 'dc output', 'dc_awal']), None)
                    icd_col = next((c for c in df.columns if str(c).strip().lower() in ['icd_code', 'icd 10 code', 'icd-9-cm code', 'icd-10 code', 'icd-9 code', 'icd 10', 'icd 9']), None)
                    desc_col = next((c for c in df.columns if str(c).strip().lower() in ['deskripsi icd', 'icd-10 description', 'icd-9-cm description', 'deskripsi icd 10', 'desc_baru_excel', 'new cluster description']), None)
                    
                    if dc_col: has_dc_col = True
                    
                    if not pdc_col:
                        st.error(f"Tidak dapat menemukan kolom PDC (Cluster Code) di sheet '{sheet}'.")
                        continue
                        
                    for idx, row in df.iterrows():
                        pdc_val = str(row[pdc_col]).strip().upper() if pd.notna(row[pdc_col]) else ""
                        icd_val = str(row[icd_col]).strip() if icd_col and pd.notna(row[icd_col]) else "-"
                        desc_val = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else "-"
                        
                        if not pdc_val or pdc_val == 'NAN':
                            continue
                            
                        expected_dcs = get_expected_dc(pdc_val)
                        actual_dc = str(row[dc_col]).strip() if dc_col and pd.notna(row[dc_col]) else ""
                        
                        status = "✅ Valid"
                        notes = ""
                        
                        if not expected_dcs:
                            status = "⚠️ Warning"
                            notes = "PDC ini tidak ada di pakem Draw.io."
                            error_count += 1
                        elif dc_col and actual_dc and actual_dc != 'nan':
                            if any(actual_dc == e_dc for e_dc in expected_dcs) or actual_dc in ",".join(expected_dcs):
                                valid_count += 1
                            else:
                                status = "❌ Mismatch"
                                notes = f"Draw.io mengharuskan DC: {', '.join(expected_dcs)}."
                                error_count += 1
                        else:
                            notes = f"Seharusnya mengarah ke DC: {', '.join(expected_dcs)} (Tergantung AX jika ada)."
                    
                        row_data = {
                            "Sheet": sheet,
                            "Baris Excel": idx + 2,
                            "Kode ICD": icd_val,
                            "Deskripsi": desc_val,
                            "PDC Input": pdc_val,
                        }
                        if dc_col:
                            row_data["DC (Di Excel)"] = actual_dc
                        
                        row_data["DC Seharusnya (Draw.io)"] = ", ".join(expected_dcs) if expected_dcs else "Tidak Diketahui"
                        row_data["Status"] = status
                        row_data["Catatan Logika"] = notes
                        
                        results.append(row_data)

                # ==============================
                # 3. MENAMPILKAN HASIL
                # ==============================
                st.markdown("### 📊 Ringkasan Validasi Gabungan")
                col1, col2, col3 = st.columns(3)
                col1.metric("Total Data Diperiksa", len(results))
                if has_dc_col:
                    col2.metric("✅ Logika Sesuai", valid_count)
                    col3.metric("❌ Logika Meleset", error_count)
                
                res_df = pd.DataFrame(results)
                
                st.markdown("### 📋 Detail Ketidaksesuaian (Mismatch)")
                if not res_df.empty:
                    # Filter hanya yang bermasalah jika ada
                    errors_df = res_df[res_df['Status'].isin(["❌ Mismatch", "⚠️ Warning"])]
                    if not errors_df.empty:
                        st.dataframe(errors_df.style.applymap(lambda x: "background-color: #ffcccc; color: #900" if x == "❌ Mismatch" else ("background-color: #fff3cd; color: #856404" if x == "⚠️ Warning" else ""), subset=['Status']), use_container_width=True)
                    else:
                        st.info("Luar biasa! Semua data di sheet yang dipilih sesuai dengan logika Draw.io.")
                
                st.markdown("### 📜 Semua Data Kombinasi")
                st.dataframe(res_df, use_container_width=True)
                
        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses file: {e}")
else:
    st.info("⬅️ Silakan unggah file Excel di panel sebelah kiri untuk memulai.")
