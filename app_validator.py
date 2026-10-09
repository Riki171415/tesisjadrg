import streamlit as st
import pandas as pd
import re
from drawio_parser import extract_xml_from_pdf, parse_drawio_xml

st.set_page_config(page_title="MDC Logic Validator & Simulator", layout="wide", page_icon="🔍")

# ==========================================
# 1. CORE LOGIC & MAPPING
# ==========================================
# Default fallback (MDC 11) jika tidak ada PDF yang diunggah
default_decision_tree = {
    'P11AA': {'11CX': '11031', 'default': '11032'},
    'D11AP': {'11PDX': '11742', 'default': '11741'},
    'D11AD': {'11PCX': '11641', 'default': '11642'},
    'P11AR': {'11PEX': '11181', '11PFX': '11181', 'default': '11182'}
}
default_flat_mapping = {
    'P11AG': '11141', 'P11AH': '11071', 'P11AB': '11041', 'P11AC': '11021',
    'P11AK': '11091', 'P11AE': '11131', 'P11AF': '11051', 'P11AJ': '11151',
    'P11AP': '11161', 'P11AN': '11101', 'P11AL': '11081', 'P11AM': '11121',
    'P11AQ': '11171', 'D11AB': '11621', 'D11AN': '11731', 'D11AM': '11721',
    'D11AL': '11711', 'D11AA': '11611', 'D11AC': '11631', 'D11AF': '11661',
    'D11AU': '11791', 'D11AK': '11701', 'D11AH': '11681', 'D11AS': '11771',
    'D11AV': '11801', 'D11AJ': '11692', 'D11AQ': '11751', 'D11AR': '11761',
    'D11AE': '11651', 'D11AG': '11671', 'D11AT': '11781'
}

def get_expected_dc(pdc, flat_map, dec_tree):
    if pd.isna(pdc) or not str(pdc).strip():
        return []
    pdc = str(pdc).strip().upper()
    if pdc in flat_map:
        return [flat_map[pdc]]
    if pdc in dec_tree:
        return list(set(dec_tree[pdc].values()))
    return []

def evaluate_logic(pdc, ax_list, flat_map, dec_tree):
    if pdc in flat_map:
        return flat_map[pdc], "Jalur langsung (tanpa AX)"
    
    if pdc in dec_tree:
        logic_branch = dec_tree[pdc]
        for ax in ax_list:
            if ax in logic_branch:
                return logic_branch[ax], f"Masuk cabang AX: {ax}"
        return logic_branch.get('default'), "Masuk cabang default (tanpa AX terkait)"
    
    return None, "PDC tidak ditemukan di Draw.io"

# ==========================================
# 2. UI & DATA LOADING
# ==========================================
st.title("🏥 Sistem Validasi & Simulator Logika MDC")
st.markdown("Aplikasi ini menggunakan **Draw.io Logic Engine** untuk memvalidasi master data dan menyimulasikan *DRG Grouping* pasien berdasarkan input **Kode ICD**.")

st.sidebar.header("📁 Upload File Kamus (Excel)")
uploaded_excel = st.sidebar.file_uploader("Upload File Usulan (.xlsx)", type=["xlsx"])

st.sidebar.header("📁 Upload File Logika (Draw.io PDF)")
uploaded_pdf = st.sidebar.file_uploader("Upload PDF Draw.io (.pdf)", type=["pdf"])

st.sidebar.header("📁 Upload Rekap DC (Opsional)")
uploaded_rekap_dc = st.sidebar.file_uploader("Upload Rekap DC Final (.xlsx)", type=["xlsx"])

decision_tree = default_decision_tree
flat_mapping = default_flat_mapping
dc_desc_dict = {}

if uploaded_rekap_dc:
    with st.spinner("Membaca kamus Deskripsi DC..."):
        try:
            df_rekap = pd.read_excel(uploaded_rekap_dc)
            if 'dc update' in df_rekap.columns and 'desc dc update' in df_rekap.columns:
                for _, r in df_rekap.iterrows():
                    if pd.notna(r['dc update']):
                        dc_desc_dict[str(r['dc update']).strip()] = str(r['desc dc update']).strip()
                st.sidebar.success(f"Berhasil memuat {len(dc_desc_dict)} deskripsi DC.")
            else:
                st.sidebar.error("Kolom 'dc update' atau 'desc dc update' tidak ditemukan di file Rekap DC.")
        except Exception as e:
            st.sidebar.error(f"Gagal membaca Rekap DC: {e}")

if uploaded_pdf:
    with st.spinner("Mengekstrak logika dari PDF Draw.io..."):
        try:
            xml_data = extract_xml_from_pdf(uploaded_pdf)
            if xml_data:
                decision_tree, flat_mapping = parse_drawio_xml(xml_data)
                st.sidebar.success(f"Logika berhasil diekstrak! ({len(flat_mapping)} PDC langsung, {len(decision_tree)} PDC bercabang)")
            else:
                st.sidebar.error("Tidak ditemukan metadata Draw.io di dalam PDF ini.")
        except Exception as e:
            st.sidebar.error(f"Gagal mem-parsing PDF: {e}")
else:
    st.sidebar.info("Menggunakan logika MDC 11 (Bawaan). Unggah PDF untuk MDC lain.")

if uploaded_excel:
    with st.spinner('Membaca file Excel...'):
        try:
            xls = pd.ExcelFile(uploaded_excel)
            sheets = xls.sheet_names
            
            # Buat TABS
            tab1, tab2 = st.tabs(["📊 Validasi Master Data", "🏥 DRG Simulator"])
            
            with tab1:
                st.subheader("Validasi Sinkronisasi Master Data")
                selected_sheets = st.multiselect("Pilih Sheet untuk divalidasi:", sheets, default=[s for s in sheets if 'ICD' in s.upper() and ('10' in s or '9' in s)])
                
                if not selected_sheets:
                    st.warning("Silakan pilih minimal satu sheet.")
                else:
                    results = []
                    error_count, valid_count = 0, 0
                    has_dc_col = False
                    
                    for sheet in selected_sheets:
                        df = pd.read_excel(xls, sheet_name=sheet)
                        
                        pdc_col = next((c for c in df.columns if str(c).strip().lower() in ['pdc_baru', 'new cluster code', 'pdc', 'cluster code']), None)
                        dc_col = next((c for c in df.columns if str(c).strip().lower() in ['dc_baru', 'dc', 'dc output', 'dc_awal']), None)
                        icd_col = next((c for c in df.columns if str(c).strip().lower() in ['icd_code', 'icd 10 code', 'icd-9-cm code', 'icd-10 code', 'icd-9 code', 'icd 10', 'icd 9']), None)
                        desc_col = next((c for c in df.columns if str(c).strip().lower() in ['deskripsi icd', 'icd-10 description', 'icd-9-cm description', 'deskripsi icd 10', 'desc_baru_excel', 'new cluster description']), None)
                        
                        if dc_col: has_dc_col = True
                        if not pdc_col: continue
                            
                        for idx, row in df.iterrows():
                            pdc_val = str(row[pdc_col]).strip().upper() if pd.notna(row[pdc_col]) else ""
                            icd_val = str(row[icd_col]).strip() if icd_col and pd.notna(row[icd_col]) else "-"
                            desc_val = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else "-"
                            
                            if not pdc_val or pdc_val == 'NAN': continue
                                
                            expected_dcs = get_expected_dc(pdc_val, flat_mapping, decision_tree)
                            actual_dc = str(row[dc_col]).strip() if dc_col and pd.notna(row[dc_col]) else ""
                            status, notes = "✅ Valid", ""
                            
                            if not expected_dcs:
                                status, notes, error_count = "⚠️ Warning", "PDC ini tidak ada di Draw.io.", error_count + 1
                            elif dc_col and actual_dc and actual_dc != 'nan':
                                if any(actual_dc == e_dc for e_dc in expected_dcs) or actual_dc in ",".join(expected_dcs):
                                    valid_count += 1
                                else:
                                    status, notes, error_count = "❌ Mismatch", f"Draw.io mengharuskan DC: {', '.join(expected_dcs)}.", error_count + 1
                            else:
                                notes = f"Seharusnya mengarah ke DC: {', '.join(expected_dcs)} (Tergantung AX jika ada)."
                        
                            row_data = {"Sheet": sheet, "Baris Excel": idx + 2, "Kode ICD": icd_val, "Deskripsi": desc_val, "PDC Input": pdc_val}
                            if dc_col: row_data["DC (Di Excel)"] = actual_dc
                            row_data["DC Seharusnya (Draw.io)"] = ", ".join(expected_dcs) if expected_dcs else "Tidak Diketahui"
                            
                            if dc_desc_dict and expected_dcs:
                                expected_descs = [dc_desc_dict.get(dc, "Deskripsi tidak ditemukan") for dc in expected_dcs]
                                row_data["Deskripsi DC Seharusnya"] = " | ".join(expected_descs)
                                
                            row_data["Status"] = status
                            row_data["Catatan Logika"] = notes
                            results.append(row_data)

                    res_df = pd.DataFrame(results)
                    st.markdown("### 📋 Detail Ketidaksesuaian (Mismatch)")
                    if not res_df.empty:
                        errors_df = res_df[res_df['Status'].isin(["❌ Mismatch", "⚠️ Warning"])]
                        if not errors_df.empty:
                            # Gunakan map alih-alih applymap untuk versi pandas terbaru
                            st.dataframe(errors_df.style.map(lambda x: "background-color: #ffcccc; color: #900" if x == "❌ Mismatch" else ("background-color: #fff3cd; color: #856404" if x == "⚠️ Warning" else ""), subset=['Status']), use_container_width=True)
                        else:
                            st.info("Semua data di sheet yang dipilih sesuai dengan logika Draw.io.")
                    st.markdown("### 📜 Semua Data Kombinasi")
                    st.dataframe(res_df, use_container_width=True)

            with tab2:
                st.subheader("Simulasi DRG Berdasarkan Input ICD")
                
                # 3. BUILD ICD DICTIONARY FROM EXCEL
                icd_dict = {}
                
                for sheet in sheets:
                    if 'ICD' in sheet.upper():
                        df = pd.read_excel(xls, sheet_name=sheet)
                        icd_col = next((c for c in df.columns if str(c).strip().lower() in ['icd_code', 'icd 10 code', 'icd-9-cm code', 'icd-10 code', 'icd-9 code', 'icd 10', 'icd 9']), None)
                        cluster_col = next((c for c in df.columns if str(c).strip().lower() in ['new cluster code', 'pdc_baru', 'cluster code', 'pdc']), None)
                        desc_col = next((c for c in df.columns if str(c).strip().lower() in ['deskripsi icd', 'icd-10 description', 'icd-9-cm description', 'deskripsi icd 10']), None)
                        
                        if icd_col and cluster_col:
                            for idx, row in df.iterrows():
                                icd = str(row[icd_col]).strip()
                                cluster = str(row[cluster_col]).strip().upper()
                                desc = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else ""
                                
                                if icd != 'nan' and cluster != 'NAN' and icd and cluster:
                                    icd_dict[icd] = {"cluster": cluster, "desc": desc}
                
                icd_list = list(icd_dict.keys())
                
                if not icd_list:
                    st.error("Gagal mengekstrak kamus ICD dari file Excel.")
                else:
                    st.markdown("Silakan pilih kode Diagnosa (ICD 10) dan Tindakan (ICD 9 CM). Sistem akan otomatis mencari PDC/AX-nya dan mengevaluasi DC final.")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("**Diagnosa (ICD 10)**")
                        primary_diag = st.selectbox("Diagnosa Utama:", [""] + icd_list, key="p_diag")
                        secondary_diags = st.multiselect("Diagnosa Sekunder (Bisa lebih dari 1):", icd_list, key="s_diag")
                    
                    with col2:
                        st.markdown("**Prosedur / Tindakan (ICD 9 CM)**")
                        procedures = st.multiselect("Tindakan Utama/Sekunder:", icd_list, key="proc")
                    
                    if st.button("🚀 Jalankan Grouper Simulator"):
                        st.markdown("---")
                        st.subheader("💡 Hasil Analisis Grouper")
                        
                        # A. Ekstrak Cluster dari Input
                        pdc_candidates = []
                        ax_candidates = []
                        
                        all_inputs = []
                        if primary_diag: all_inputs.append(("Diagnosa Utama", primary_diag))
                        for d in secondary_diags: all_inputs.append(("Diagnosa Sekunder", d))
                        for p in procedures: all_inputs.append(("Prosedur", p))
                        
                        st.markdown("**1. Pemetaan ICD ke Cluster (Kamus):**")
                        for label, code in all_inputs:
                            c_info = icd_dict.get(code, {})
                            cluster = c_info.get('cluster', 'Tidak Ditemukan')
                            desc = c_info.get('desc', '')
                            
                            st.write(f"- {label} **{code}** ({desc}) ➡️ Masuk ke Cluster: **{cluster}**")
                            
                            # Identifikasi apakah dia PDC (awalannya P/D) atau AX (contoh 11CX, dll)
                            if cluster.startswith('P') or cluster.startswith('D'):
                                pdc_candidates.append(cluster)
                            elif cluster != 'Tidak Ditemukan':
                                ax_candidates.append(cluster)
                        
                        # B. Tentukan Base PDC
                        st.markdown("**2. Penentuan Base PDC:**")
                        final_pdc = None
                        if pdc_candidates:
                            # Prioritaskan Prosedur (Surgical Partition - P) daripada Diagnosa (Medical Partition - D)
                            surgical = [c for c in pdc_candidates if c.startswith('P')]
                            if surgical:
                                final_pdc = surgical[0]
                                st.write(f"✅ Sistem mendeteksi adanya Tindakan Operasi (Surgical). Base PDC yang digunakan adalah **{final_pdc}**.")
                            else:
                                final_pdc = pdc_candidates[0]
                                st.write(f"✅ Tidak ada tindakan operasi. Base PDC yang digunakan dari Diagnosa adalah **{final_pdc}**.")
                        else:
                            st.error("❌ Tidak ada Base PDC (D/P) yang terdeteksi dari input ICD Anda!")
                        
                        # C. Evaluasi Engine Draw.io
                        if final_pdc:
                            st.markdown("**3. Eksekusi Engine Draw.io:**")
                            if ax_candidates:
                                st.write(f"Menemukan AX tambahan dari diagnosa sekunder/prosedur: **{', '.join(ax_candidates)}**")
                            else:
                                st.write("Tidak ada AX tambahan yang terdeteksi.")
                            
                            dc_result, logic_note = evaluate_logic(final_pdc, ax_candidates, flat_mapping, decision_tree)
                            
                            if dc_result:
                                desc = dc_desc_dict.get(dc_result, "")
                                st.success(f"🎉 **HASIL FINAL DC: {dc_result}**" + (f" - {desc}" if desc else ""))
                                st.info(f"Keterangan Logika: {logic_note}")
                            else:
                                st.error(f"❌ {logic_note}")

        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
else:
    st.info("⬅️ Silakan unggah file Excel Master/Usulan di panel sebelah kiri untuk memulai.")
