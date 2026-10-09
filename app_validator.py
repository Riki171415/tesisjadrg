import streamlit as st
import pandas as pd
import re
from drawio_parser import extract_xml_from_pdf, parse_drawio_xml, extract_page_names

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

decision_tree = default_decision_tree
flat_mapping = default_flat_mapping
dc_desc_dict = {}

# Muat otomatis Kamus Deskripsi DC (Tertanam)
import json
import os
if os.path.exists('dc_desc.json'):
    with open('dc_desc.json', 'r') as f:
        dc_desc_dict = json.load(f)

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
            tab1, tab2, tab3 = st.tabs(["📊 Validasi Master Data", "🏥 DRG Batch Simulator", "✅ Validasi SOP Format Data"])
            
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
                        df = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                        
                        cols_lower = {str(c).strip().lower(): c for c in df.columns}
                        
                        pdc_col = None
                        for target in ['new cluster code', 'pdc_baru', 'pdc', 'cluster code']:
                            if target in cols_lower: pdc_col = cols_lower[target]; break
                            
                        dc_col = None
                        for target in ['dc_baru', 'dc', 'dc output', 'dc_awal']:
                            if target in cols_lower: dc_col = cols_lower[target]; break
                            
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
                            try:
                                styled_df = errors_df.style.map(lambda x: "background-color: #ffcccc; color: #900" if x == "❌ Mismatch" else ("background-color: #fff3cd; color: #856404" if x == "⚠️ Warning" else ""), subset=['Status'])
                            except AttributeError:
                                styled_df = errors_df.style.applymap(lambda x: "background-color: #ffcccc; color: #900" if x == "❌ Mismatch" else ("background-color: #fff3cd; color: #856404" if x == "⚠️ Warning" else ""), subset=['Status'])
                            st.dataframe(styled_df, use_container_width=True)
                        else:
                            st.info("Semua data di sheet yang dipilih sesuai dengan logika Draw.io.")
                    
                    if st.checkbox("Tampilkan seluruh data kombinasi (Berpotensi berat untuk data besar)"):
                        st.markdown("### 📜 Semua Data Kombinasi")
                        st.dataframe(res_df, use_container_width=True)

            with tab2:
                st.subheader("Simulasi DRG Batch Berdasarkan Kombinasi Kasus")
                
                st.markdown("Pilih sheet Kamus (MDC Master) yang akan digunakan. **Penting:** Pisahkan antara sheet Ranap dan Rajal agar kodenya tidak bertabrakan (karena Draw.io biasanya terpisah untuk Ranap).")
                selected_dict_sheets = st.multiselect("Pilih Sheet Kamus untuk Simulasi:", sheets, default=[s for s in sheets if 'ranap' in s.lower() and 'icd' in s.lower()])
                
                # Build ICD Dictionary from Kamus
                icd_dict = {}
                diags_master = set()
                tinds_master = set()
                
                for sheet in selected_dict_sheets:
                    if 'ICD' in sheet.upper():
                        df_icd = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                        icd_col = next((c for c in df_icd.columns if str(c).strip().lower() in ['icd_code', 'icd 10 code', 'icd-9-cm code', 'icd-10 code', 'icd-9 code', 'icd 10', 'icd 9']), None)
                        
                        cols_lower = {str(c).strip().lower(): c for c in df_icd.columns}
                        cluster_col = None
                        for target in ['new cluster code', 'pdc_baru', 'cluster code', 'pdc']:
                            if target in cols_lower:
                                cluster_col = cols_lower[target]
                                break
                        
                        if icd_col and cluster_col:
                            # Tentukan tipe berdasarkan nama sheet atau kolom
                            is_icd10 = '10' in sheet.lower() or '10' in icd_col.lower()
                            is_icd9 = '9' in sheet.lower() or '9' in icd_col.lower()
                            
                            for _, row in df_icd.iterrows():
                                icd = str(row[icd_col]).strip()
                                cluster = str(row[cluster_col]).strip().upper()
                                if icd != 'nan' and cluster != 'NAN' and icd and cluster:
                                    if icd not in icd_dict:
                                        icd_dict[icd] = set()
                                    icd_dict[icd].add(cluster)
                                    if is_icd10:
                                        diags_master.add(icd)
                                    elif is_icd9:
                                        tinds_master.add(icd)
                                    else:
                                        diags_master.add(icd) # Default fallback
                
                cluster_to_icd = {}
                for icd, clusters in icd_dict.items():
                    for cluster in clusters:
                        if cluster not in cluster_to_icd:
                            cluster_to_icd[cluster] = []
                        cluster_to_icd[cluster].append(icd)
                        
                st.markdown(f"Aplikasi telah memetakan total **{len(cluster_to_icd)} Cluster Code** dari file Kamus Master yang dipilih.")
                st.markdown("Klik tombol di bawah ini untuk melahirkan Skenario Test Cases (Pasien Dummy) yang 100% mewakili seluruh jalur logika Draw.io.")
                
                if st.button("🚀 Generate Test Cases dari Draw.io (Reverse Engineering)"):
                    with st.spinner("Membaca rute Draw.io dan menyusun skenario pasien..."):
                        test_cases = []
                        
                        # 1. Jalur Flat Mapping (Tanpa AX)
                        for pdc, dc in flat_mapping.items():
                            pdc_icds = cluster_to_icd.get(pdc, ["KODE_ICD_TIDAK_ADA_DI_KAMUS"])
                            desc = dc_desc_dict.get(dc, "")
                            test_cases.append({
                                "Skenario Logika": f"Jalur Lurus (PDC -> DC)",
                                "Input Kode ICD": pdc_icds[0],
                                "Cluster Code Dibutuhkan": pdc,
                                "Expected DC": dc,
                                "Deskripsi DC": desc
                            })
                            
                        # 2. Jalur Decision Tree (Dengan AX)
                        for pdc, branches in decision_tree.items():
                            pdc_icds = cluster_to_icd.get(pdc, ["KODE_ICD_TIDAK_ADA_DI_KAMUS"])
                            
                            # Jalur Default (PDC Saja)
                            default_dc = branches.get('default')
                            if default_dc:
                                desc = dc_desc_dict.get(default_dc, "")
                                test_cases.append({
                                    "Skenario Logika": f"Jalur Default (Tanpa AX)",
                                    "Input Kode ICD": pdc_icds[0],
                                    "Cluster Code Dibutuhkan": pdc,
                                    "Expected DC": default_dc,
                                    "Deskripsi DC": desc
                                })
                                
                            # Jalur AX (PDC + 1 atau Lebih AX)
                            for ax_cond, dc in branches.items():
                                if ax_cond == 'default': continue
                                
                                # ax_cond bisa berisi multiple AX (misal: "11PEX & 11PFX")
                                required_axs = ax_cond.split('&')
                                icd_inputs = [pdc_icds[0]]
                                cluster_inputs = [pdc]
                                
                                for req_ax in required_axs:
                                    req_ax = req_ax.strip()
                                    ax_icds = cluster_to_icd.get(req_ax, ["KODE_ICD_TIDAK_ADA_DI_KAMUS"])
                                    icd_inputs.append(ax_icds[0])
                                    cluster_inputs.append(req_ax)
                                    
                                desc = dc_desc_dict.get(dc, "")
                                test_cases.append({
                                    "Skenario Logika": f"Jalur Bersyarat ({ax_cond})",
                                    "Input Kode ICD": " + ".join(icd_inputs),
                                    "Cluster Code Dibutuhkan": " + ".join(cluster_inputs),
                                    "Expected DC": dc,
                                    "Deskripsi DC": desc
                                })
                                
                        if test_cases:
                            st.success(f"Berhasil me-reverse-engineer **{len(test_cases)} Jalur Logika** dari Draw.io menjadi Skenario Uji Coba.")
                            df_tests = pd.DataFrame(test_cases)
                            
                            # Cek jika ada kamus yang bolong (KODE_ICD_TIDAK_ADA_DI_KAMUS)
                            missing_df = df_tests[df_tests["Input Kode ICD"].str.contains("KODE_ICD_TIDAK_ADA_DI_KAMUS")]
                            if not missing_df.empty:
                                st.warning(f"⚠️ Perhatian: Ditemukan {len(missing_df)} jalur Draw.io yang terputus karena Cluster Code-nya tidak ada di file Excel Kamus yang Anda pilih.")
                                
                            st.dataframe(df_tests, use_container_width=True)
                        else:
                            st.error("Tidak ada jalur logika yang bisa diekstrak. Pastikan PDF Draw.io sudah diunggah dan terbaca.")

                st.markdown("---")
                st.subheader("🕵️ Analisis Celah Logika (Error Hunter)")
                st.markdown("Fitur ini menganalisis secara matematis apakah ada kode di Kamus yang **berujung jalan buntu (Dead End / GAGAL)** di Draw.io, atau sebaliknya.")
                
                if st.button("🔍 Mulai Perburuan Error (Dead End)"):
                    with st.spinner("Menelusuri celah antara Kamus Excel dan Draw.io..."):
                        error_reports = []
                        
                        # Ambil semua Cluster yang ada di Excel
                        excel_clusters = set(cluster_to_icd.keys())
                        
                        # Ambil semua PDC dan AX yang ada di Draw.io
                        drawio_pdcs = set(flat_mapping.keys()).union(set(decision_tree.keys()))
                        drawio_axs = set()
                        for branches in decision_tree.values():
                            for ax_cond in branches.keys():
                                if ax_cond != 'default':
                                    for req_ax in ax_cond.split('&'):
                                        drawio_axs.add(req_ax.strip())
                                        
                        drawio_all_clusters = drawio_pdcs.union(drawio_axs)
                        
                        # 1. PDCs di Kamus yang BUNTU (Tidak ada di Draw.io)
                        for cluster in excel_clusters:
                            if cluster.startswith('P') or cluster.startswith('D'):
                                if cluster not in drawio_pdcs:
                                    icds = cluster_to_icd[cluster]
                                    error_reports.append({
                                        "Tipe Error": "PDC Buntu (Dead End)",
                                        "Cluster Code": cluster,
                                        "Keterangan": f"PDC ini ada di Excel (untuk ICD: {', '.join(icds[:3])}...), TAPI TIDAK DITEMUKAN di Draw.io! Semua pasien dengan diagnosa ini pasti GAGAL."
                                    })
                                    
                            # 2. AXs di Kamus yang BUNTU (Tidak pernah diwajibkan di Draw.io)
                            else:
                                if cluster not in drawio_axs:
                                    icds = cluster_to_icd[cluster]
                                    error_reports.append({
                                        "Tipe Error": "AX Tidak Terpakai (Orphan)",
                                        "Cluster Code": cluster,
                                        "Keterangan": f"AX ini ada di Excel (untuk ICD: {', '.join(icds[:3])}...), TAPI tidak pernah dipersyaratkan oleh jalur manapun di Draw.io."
                                    })
                                    
                        # 3. Cluster di Draw.io yang KOSONG (Tidak ada ICD-nya di Kamus)
                        for draw_cluster in drawio_all_clusters:
                            if draw_cluster not in excel_clusters:
                                error_reports.append({
                                    "Tipe Error": "Missing ICD (Kamus Bolong)",
                                    "Cluster Code": draw_cluster,
                                    "Keterangan": f"Draw.io mensyaratkan Cluster ini, TAPI tidak ada satupun Kode ICD di file Excel yang di-mapping ke Cluster ini."
                                })
                                
                        # 4. PDC Draw.io tanpa Default Branch (Risiko GAGAL jika AX tidak terpenuhi)
                        for pdc, branches in decision_tree.items():
                            if 'default' not in branches:
                                error_reports.append({
                                    "Tipe Error": "Missing Default Branch (Risiko Dead End)",
                                    "Cluster Code": pdc,
                                    "Keterangan": f"Draw.io memiliki rute untuk PDC ini, tapi TIDAK ADA jalur 'No/Default'. Jika pasien memiliki PDC ini tapi tidak memiliki AX yang tepat, statusnya pasti GAGAL."
                                })
                                
                        if error_reports:
                            st.error(f"🚨 Ditemukan **{len(error_reports)} Titik Error / Celah Logika** antara Kamus dan Draw.io!")
                            st.dataframe(pd.DataFrame(error_reports), use_container_width=True)
                        else:
                            st.success("✅ Luar biasa! Tidak ada celah logika, *Dead End*, atau kamus yang terputus antara Excel dan Draw.io.")
            with tab3:
                st.subheader("Pengecekan Kriteria & Validasi SOP Format Data")
                
                if st.button("🚀 Jalankan Pengecekan Kriteria (Audit Format)"):
                    with st.spinner("Mengaudit format Excel dan Draw.io..."):
                        audit_results = []
                        
                        # 1. Pengecekan Penamaan Sheet
                        req_sheets = ['ranap cluster (icd 10)', 'ranap cluster (icd 9 cm)', 'rajal cluster (icd 10)', 'rajal cluster (icd 9 cm)']
                        sheet_lower = [s.strip().lower() for s in sheets]
                        for req in req_sheets:
                            if req not in sheet_lower:
                                audit_results.append(f"❌ **Penamaan Sheet:** Sheet '{req}' tidak ditemukan atau penulisannya salah.")
                        
                        if not any("Penamaan Sheet" in r for r in audit_results):
                            audit_results.append("✅ **Penamaan Sheet:** Semua sheet Ranap & Rajal terdeteksi benar.")
                            
                        # Iterasi ke semua sheet ICD untuk cek aturan baris
                        icd_sheets = [s for s in sheets if 'ICD' in s.upper()]
                        for sheet in icd_sheets:
                            df_audit = pd.read_excel(xls, sheet_name=sheet, dtype=str)
                            
                            # 2. Kolom A wajib terisi
                            if df_audit.iloc[:, 0].isna().any():
                                audit_results.append(f"❌ **Kolom A Kosong:** Terdapat baris kosong di Kolom A pada sheet '{sheet}'.")
                                
                            # 3. Klasifikasi Cluster (PDC/AX) harus ada
                            cols_lower = {str(c).strip().lower(): c for c in df_audit.columns}
                            cluster_col = None
                            for target in ['new cluster code', 'pdc_baru', 'cluster code', 'pdc']:
                                if target in cols_lower: cluster_col = cols_lower[target]; break
                            if cluster_col:
                                if df_audit[cluster_col].isna().any():
                                    audit_results.append(f"❌ **Klasifikasi Cluster Kosong:** Ada kode yang tidak memiliki cluster (PDC/AX) di sheet '{sheet}'.")
                            
                            # 4. Konsistensi Procedures (ICD-9-CM)
                            if '9' in sheet:
                                desc_col = next((c for c in df_audit.columns if str(c).strip().lower() in ['new cluster description', 'desc_baru_excel']), None)
                                if desc_col:
                                    proc_issues = df_audit[df_audit[desc_col].astype(str).str.contains('Proc.', regex=False, na=False)]
                                    if not proc_issues.empty:
                                        audit_results.append(f"❌ **Konsistensi Narasi:** Ditemukan kata 'Proc.' (seharusnya 'Procedures') di sheet '{sheet}' sebanyak {len(proc_issues)} baris.")
                                        
                            # 5. Pengecekan Duplikat
                            icd_col = next((c for c in df_audit.columns if str(c).strip().lower() in ['icd_code', 'icd 10 code', 'icd-9-cm code', 'icd-10 code', 'icd-9 code', 'icd 10', 'icd 9']), None)
                            if icd_col:
                                # Buang baris kosong agar NaN tidak dihitung sebagai duplikat
                                valid_df = df_audit.dropna(subset=[icd_col])
                                dups = valid_df[valid_df.duplicated(subset=[icd_col], keep=False)]
                                if not dups.empty:
                                    audit_results.append(f"❌ **Duplikasi Data:** Ditemukan {len(dups)} baris duplikat Kode ICD di sheet '{sheet}'.")
                                
                        # 6. Konsistensi Draw.io (Cek Tab 'Current')
                        if uploaded_pdf:
                            pdf_xml = extract_xml_from_pdf(uploaded_pdf)
                            if pdf_xml:
                                pages = extract_page_names(pdf_xml)
                                if 'Current' not in pages:
                                    audit_results.append("⚠️ **Draw.io Tab:** Tidak ditemukan sheet/tab bernama 'Current' di dalam file Draw.io yang membuktikan pembaruan.")
                                else:
                                    audit_results.append("✅ **Draw.io Tab:** Tab 'Current' terdeteksi.")
                                    
                                # Pengecekan silang PDC Code / Description bisa ditambah disini (tapi Tab 1 sudah memfasilitasi pengecekan kode).
                                audit_results.append("✅ **Konsistensi Draw.io vs Excel:** (Lihat hasil di Tab 1 'Validasi Master Data' untuk kecocokan logika PDC/AX).")
                        else:
                            audit_results.append("⚠️ **Draw.io:** File Draw.io tidak diunggah, pengecekan tab 'Current' dilewati.")
                        
                        # Tampilkan Hasil Audit
                        for res in audit_results:
                            if res.startswith("✅"):
                                st.success(res)
                            elif res.startswith("❌"):
                                st.error(res)
                            else:
                                st.warning(res)
                                
                        st.info("💡 **Catatan Validasi Akhir:** Pastikan Anda juga mengecek silang di Web INAGROUPER setelah seluruh perbaikan di atas dilakukan.")
        except Exception as e:
            st.error(f"Terjadi kesalahan: {e}")
else:
    st.info("⬅️ Silakan unggah file Excel Master/Usulan di panel sebelah kiri untuk memulai.")
