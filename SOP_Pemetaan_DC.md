# SOP (Standard Operating Procedure) - Pemetaan Distribution Center (DC)

**Dokumen ini merupakan standar panduan (SOP) untuk penulisan script Python terkait pemetaan/mapping DC dan Deskripsi DC yang bersumber dari file `DRG Master.xlsx`.**

---

## 1. Latar Belakang
Pada file sumber data (`DRG versi September 2025-Final Publish & DRG Master.xlsx`, sheet `Rencana Usulan per 29092026`), terdapat dua sumber informasi DC:
- Kolom **Original**: `dc` dan `dc desc`
- Kolom **Update (Usulan/Revisi)**: `dc update` dan `desc dc update`

Berdasarkan pengecekan, terdapat lebih dari 200 inkonsistensi/perubahan antara kolom original dan update. Jika script hanya membaca salah satu (misalnya hanya kolom original), maka akan ada DC baru atau deskripsi usulan yang gagal terpetakan (menjadi kosong / "Tidak terdaftar").

## 2. Aturan Standar Ekstraksi (SOP)
Untuk menjamin seluruh kode DC dan deskripsinya terpetakan dengan benar dan konsisten di seluruh script (baik di `generate_flow.py`, `map_by_desc.py`, `verify_desc.py`, dsb.), **setiap script wajib menggabungkan (concat) informasi dari kedua pasangan kolom tersebut.**

### Snippet Kode Standar yang Wajib Digunakan:
Gunakan snippet pandas di bawah ini setiap kali Anda perlu memetakan DC (membuat dictionary atau DataFrame untuk mapping):

```python
import pandas as pd

# 1. Baca master file
master_file = r'C:\Users\PUSBIKES-KEMKES\Downloads\mdc\DRG versi September 2025-Final Publish & DRG Master.xlsx'
xl_m = pd.ExcelFile(master_file)
master_df = xl_m.parse('Rencana Usulan per 29092026', dtype=str)

# 2. Ambil dari kolom Original
dc_map_1 = master_df[['dc', 'dc desc']].dropna().rename(columns={'dc':'dc_code', 'dc desc':'dc_desc'})

# 3. Ambil dari kolom Update
dc_map_2 = master_df[['dc update', 'desc dc update']].dropna().rename(columns={'dc update':'dc_code', 'desc dc update':'dc_desc'})

# 4. Gabungkan (Concat) dan hapus duplikat
dc_map = pd.concat([dc_map_1, dc_map_2]).drop_duplicates()

# 5. Standarisasi format string kode DC (misal menghapus akhiran .0)
dc_map['dc_str'] = dc_map['dc_code'].str.replace(r'\.0$', '', regex=True).str.strip()

# OPTIONAL: Jika Anda membutuhkan dictionary mapping dari DC Code ke Deskripsi
dc_dict = dict(zip(dc_map['dc_str'], dc_map['dc_desc']))

# OPTIONAL: Jika script lama butuh nama kolom secara harfiah 'dc' dan 'dc desc'
dc_map['dc'] = dc_map['dc_code']
dc_map['dc desc'] = dc_map['dc_desc']
```

## 3. Ketentuan Validasi
1. Jangan menggunakan `dropna()` sebelum melakukan pemilihan kolom (kolom usulan dan original dipisah).
2. Hindari *hardcode* nama kolom di luar `['dc', 'dc desc']` dan `['dc update', 'desc dc update']` untuk Master DC ini, kecuali jika format master sheet dari Kemenkes mengalami perubahan struktur.
3. Semua pembuatan script `mapping_*.py` atau `generate_flow_*.py` ke depannya **wajib** mengikuti pedoman ini.

---
*Diperbarui berdasarkan hasil penyesuaian inkonsistensi DC pada 6 Oktober 2026.*
