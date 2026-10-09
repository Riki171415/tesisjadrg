import pandas as pd
import numpy as np
import re

master_file = r'C:\Users\PUSBIKES-KEMKES\Downloads\mdc\DRG versi September 2025-Final Publish & DRG Master.xlsx'
xl_m = pd.ExcelFile(master_file)
master_df = xl_m.parse('Rencana Usulan per 29092026', dtype=str)
dc_map_1 = master_df[['dc', 'dc desc']].dropna().rename(columns={'dc':'dc_code', 'dc desc':'dc_desc'})
dc_map_2 = master_df[['dc update', 'desc dc update']].dropna().rename(columns={'dc update':'dc_code', 'desc dc update':'dc_desc'})
dc_map = pd.concat([dc_map_1, dc_map_2]).drop_duplicates()
dc_map['dc_str'] = dc_map['dc_code'].str.replace(r'\.0$', '', regex=True).str.strip()
dc_dict = dict(zip(dc_map['dc_str'], dc_map['dc_desc']))

pdf_mapping = [
    ('P11AA', '11032, 11031', 'Endovascular Procedures for Neurologic Diseases'),
    ('P11AG', '11141', 'Complex Intracranial Procedures'),
    ('P11AH', '11071', 'Simple Intracranial Procedures'),
    ('P11AB', '11041', 'Complex Spinal Procedures'),
    ('P11AC', '11021', 'Complex Craniotomy Procedures'),
    ('P11AK', '11091', 'Simple Spinal Procedures'),
    ('P11AE', '11131', 'Moderate Craniotomy Procedures'),
    ('P11AF', '11051', 'Simple Craniotomy Procedures'),
    ('P11AJ', '11151', 'Moderate Spinal Procedures'),
    ('P11AP', '11161', 'Neuromodulation Procedures'),
    ('P11AN', '11101', 'Extracranial Vascular Procedures'),
    ('P11AL', '11081', 'Complex Peripheral Nerve Procedures'),
    ('P11AD', '11061', 'Ventricular System Procedures'),
    ('P11AM', '11121', 'Simple Peripheral Nerve Procedures'),
    ('P11AQ', '11171', 'Neurologic Angiography'),
    ('P11AR', '11182, 11181', 'Percutaneous Transluminal Angiography (PTA) Cerebral'),
    ('D11AB', '11621', 'Nervous System Neoplasm'),
    ('D11AP', '11741, 11742', 'Ischemic Stroke'),
    ('D11AN', '11731', 'Subarachnoid Hemorrhage'),
    ('D11AM', '11721', 'Hemorrhagic Stroke'),
    ('D11AD', '11642, 11641', 'Multiple Sclerosis & Other Nervous Diseases'),
    ('D11AL', '11711', 'Congenital Nervous System Disorders'),
    ('D11AA', '11611', 'Nervous System Infections'),
    ('D11AC', '11631', 'Systemic Atrophy'),
    ('D11AF', '11661', 'Neurobehaviour & Other Neurodegenerative Disease'),
    ('D11AU', '11791', 'Delirium & Coma'),
    ('D11AK', '11701', 'Cerebral Palsy & Others Paralytic Syndrome'),
    ('D11AH', '11681', 'Nerves, Radix, & Plexus Disorders'),
    ('D11AS', '11771', 'Medula Spinalis, Spine & Peripheral Nervous Trauma'),
    ('D11AV', '11801', 'Others Nervous System Disorders'),
    ('D11AJ', '11692', 'Myoneural Junction & Muscle Disease'),
    ('D11AQ', '11751', 'Transient Ischemic Attact & Others Neurovascular Disorders'),
    ('D11AR', '11761', 'Head Trauma'),
    ('D11AE', '11651', 'Extrapyramidal & Movement Disorders'),
    ('D11AG', '11671', 'Epilepsy & Convulsion Attack'),
    ('D11AT', '11781', 'Headache'),
    ('11CX', '11031', 'Endovascular Proc. for Complex Neurologic Diseases'),
    ('11PDX', '11742', 'Ischemic Stroke w Thrombolytic'),
    ('11PCX', '11641', 'Multiple Sclerosis & Other Nervous Diseases w Plasmapheresis Proc.'),
    ('11PEX', '11181', 'Percutaneous Transluminal Angiography (PTA) Cerebral w Stents Proc.'),
    ('11PFX', '11181', 'Percutaneous Transluminal Angiography (PTA) Cerebral w Stents Proc.')
]

for pdc, dc_num, desc in reversed(pdf_mapping):
    if dc_num:
        for num in dc_num.split(','):
            num = num.strip()
            if num and num not in dc_dict:
                dc_dict[num] = desc

df_pdf = pd.DataFrame(pdf_mapping, columns=['PDC', 'DC_Drawio', 'Desc_Drawio'])

def get_drg_desc(dcs):
    if pd.isna(dcs) or dcs == '': return ''
    dc_list = [d.strip() for d in str(dcs).split(',')]
    descs = []
    for d in dc_list:
        if d in dc_dict:
            descs.append(dc_dict[d])
    return ', '.join(descs)

df_pdf['Desc_DRG_Master'] = df_pdf['DC_Drawio'].apply(get_drg_desc)

file_path = r'C:\Users\PUSBIKES-KEMKES\Downloads\mdc\MDC 11 Neuro - Cluster - Follow Up Usulan.xlsx'
xl = pd.ExcelFile(file_path)

def normalize_cols(df, code_col_variations):
    icd_col = None
    for col in code_col_variations:
        if col in df.columns:
            icd_col = col
            df = df.rename(columns={col: 'ICD_Code'})
            break
            
    df = df.dropna(subset=['ICD_Code']).copy()
            
    if 'New Cluster Code' in df.columns:
        df['Final_Cluster_Code'] = np.where(df['New Cluster Code'].notna(), df['New Cluster Code'], df['Cluster Code'])
    else:
        df['Final_Cluster_Code'] = df['Cluster Code']
    
    df['ICD_Code'] = df['ICD_Code'].astype(str).str.strip().str.upper()
    df['Final_Cluster_Code'] = df['Final_Cluster_Code'].astype(str).str.strip().str.upper()
    
    df['ICD_Code'] = df['ICD_Code'].replace('NAN', np.nan)
    df['Final_Cluster_Code'] = df['Final_Cluster_Code'].replace('NAN', np.nan)
    
    df = df.dropna(subset=['ICD_Code']).copy()
    return df

def compare(sheet_web, sheet_usulan, icd_cols):
    df_web = xl.parse(sheet_web, dtype=str)
    df_usulan = xl.parse(sheet_usulan, dtype=str)
    
    df_web = normalize_cols(df_web, icd_cols)
    df_usulan = normalize_cols(df_usulan, icd_cols)
    
    web_sub = df_web[['ICD_Code', 'Final_Cluster_Code', 'Cluster Description']].copy()
    web_sub.columns = ['ICD_Code', 'Cluster_Web', 'Desc_Web_Excel']
    
    usulan_sub = df_usulan[['ICD_Code', 'Final_Cluster_Code', 'Cluster Description']].copy()
    if 'Keterangan' in df_usulan.columns:
        usulan_sub['Keterangan_Original'] = df_usulan['Keterangan']
    else:
        usulan_sub['Keterangan_Original'] = np.nan
    usulan_sub.columns = ['ICD_Code', 'Cluster_Usulan', 'Desc_Usulan_Excel', 'Keterangan_Original']
    
    merged = pd.merge(web_sub, usulan_sub, on='ICD_Code', how='outer')
    
    conditions = [
        merged['Cluster_Web'].isna(),
        merged['Cluster_Usulan'].isna(),
        (merged['Cluster_Web'].notna()) & (merged['Cluster_Usulan'].notna()) & (merged['Cluster_Web'] != merged['Cluster_Usulan'])
    ]
    choices = ['Baru Ditambahkan', 'Dihapus (Drop)', 'Pindah Cluster']
    merged['Status'] = np.select(conditions, choices, default='Tetap')
    
    return merged

df_r10 = compare('ranap cluster icd10_web20260922', 'ranap clusters (ICD 10)', ['ICD-10 Code', 'ICD 10 Code'])
df_r9 = compare('ranap cluster icd9cm_web2026092', 'ranap clusters (ICD 9)', ['ICD-9-CM Code'])
df_j9 = compare('rajal cluster icd9cm_web2026092', 'rajal clusters (icd 9 cm)', ['ICD-9-CM Code'])

all_dfs = {'Ranap ICD 10': df_r10, 'Ranap ICD 9': df_r9, 'Rajal ICD 9': df_j9}
res = {}

def normalize_text(t):
    if pd.isna(t): return ''
    t = str(t).lower().strip()
    t = re.sub(r'[^a-z0-9]', '', t)
    return t

def generate_diff_warning(row, phase):
    excel_desc = row[f'Desc_{phase}_Excel']
    drawio_desc = row[f'Desc_{phase}_Drawio']
    if pd.isna(excel_desc) or pd.isna(drawio_desc) or excel_desc == '' or drawio_desc == '':
        return ''
    
    norm_ex = normalize_text(excel_desc)
    norm_dr = normalize_text(drawio_desc)
    
    if norm_ex != norm_dr:
        return f'{phase.upper()}: Excel (\"{excel_desc}\") != Draw.io (\"{drawio_desc}\")'
    return ''

for sheet, df in all_dfs.items():
    df = df.merge(df_pdf, how='left', left_on='Cluster_Web', right_on='PDC')
    df.rename(columns={'DC_Drawio': 'DC_Awal_Drawio', 'Desc_Drawio': 'Desc_Web_Drawio', 'Desc_DRG_Master': 'Desc_Awal_DRG_Master'}, inplace=True)
    df.drop('PDC', axis=1, inplace=True, errors='ignore')
    
    df = df.merge(df_pdf, how='left', left_on='Cluster_Usulan', right_on='PDC')
    df.rename(columns={'DC_Drawio': 'DC_Baru_Drawio', 'Desc_Drawio': 'Desc_Usulan_Drawio', 'Desc_DRG_Master': 'Desc_Baru_DRG_Master'}, inplace=True)
    df.drop('PDC', axis=1, inplace=True, errors='ignore')
    
    df['Warning_Web'] = df.apply(lambda r: generate_diff_warning(r, 'Web'), axis=1) 
    df['Warning_Usulan'] = df.apply(lambda r: generate_diff_warning(r, 'Usulan'), axis=1)
    
    df['Keterangan_Beda_Deskripsi'] = df[['Warning_Web', 'Warning_Usulan']].apply(lambda x: ' | '.join([i for i in x if i]), axis=1)
    
    cols = ['ICD_Code', 'Status', 
            'Cluster_Web', 'Desc_Web_Excel', 'DC_Awal_Drawio', 'Desc_Web_Drawio', 'Desc_Awal_DRG_Master',
            'Cluster_Usulan', 'Desc_Usulan_Excel', 'DC_Baru_Drawio', 'Desc_Usulan_Drawio', 'Desc_Baru_DRG_Master',
            'Keterangan_Beda_Deskripsi']
    if 'Keterangan_Original' in df.columns:
        cols.append('Keterangan_Original')
    
    df = df[cols]
    df = df[df['Status'] != 'Tetap']
    
    df = df.replace({'nan': np.nan, 'NAN': np.nan, 'None': np.nan})
    df = df.fillna('')
    
    res[sheet] = df

out_path = r'C:\Users\PUSBIKES-KEMKES\Downloads\mdc\MDC_11_Final_Drawio_DRG_Clean_Verified.xlsx'
with pd.ExcelWriter(out_path) as writer:
    for sheet, df in res.items():
        df.to_excel(writer, sheet_name=sheet, index=False)

print('Verified Excel Generated at:', out_path)
