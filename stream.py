import pandas as pd
import glob
import numpy as np
import time
import datetime as dt
import streamlit as st
from io import BytesIO
import pytz
import requests
import os
import zipfile
from xlsxwriter import Workbook
import tempfile
import re
import io

def load_excel(file_path):
    with open(file_path, 'rb') as file:
        model = pd.read_excel(file, engine='openpyxl')
    return model
 
def to_excel(df, sheet_name='Sheet1'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        # Mengakses workbook dan worksheet untuk format header
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]
        
        # Menambahkan format khusus untuk header
        header_format = workbook.add_format({'border': 0, 'bold': False, 'font_size': 12})
        
        # Menulis header manual dengan format khusus
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
    processed_data = output.getvalue()
    return processed_data
  
def get_current_time_gmt7():
    tz = pytz.timezone('Asia/Jakarta')
    return dt.datetime.now(tz).strftime('%Y%m%d_%H%M%S')
    
st.title('SCM-Cleaning')

selected_option = st.selectbox("Pilih salah satu:", ['REKAP MENTAH','REKAP PENYESUAIAN INPUTAN IA','PROMIX','REKAP DATA 42.02','WEBSMART (DINE IN/TAKEAWAY)','PENYESUAIAN IA'])
if selected_option == 'REKAP MENTAH':
    st.write('Upload file format *zip')
if selected_option == 'REKAP PENYESUAIAN INPUTAN IA':
    st.write('Upload file format *zip')
if selected_option == 'PROMIX':
    st.write('Upload file format *xlsx')
if selected_option == 'REKAP DATA 42.02':
    st.write('Upload file format *zip')
if selected_option == 'WEBSMART (DINE IN/TAKEAWAY)':
    st.write('Upload file format *zip')
if selected_option == 'PENYESUAIAN IA':
    st.write('Upload file format *zip')
 
def download_file_from_github(url, save_path):
    response = requests.get(url)
    if response.status_code == 200:
        with open(save_path, 'wb') as file:
            file.write(response.content)
        print(f"File downloaded successfully and saved to {save_path}")
    else:
        print(f"Failed to download file. Status code: {response.status_code}")
     
url = 'https://raw.githubusercontent.com/Analyst-FPnA/Rekap-SCM/main/DATABASE_IA.xlsx'

# Path untuk menyimpan file yang diunduh
save_path = 'DATABASE_IA.xlsx'

# Unduh file dari GitHub
download_file_from_github(url, save_path)

# Muat model dari file yang diunduh
if os.path.exists(save_path):
    db_ia = load_excel(save_path)
else:
    print("file does not exist") 
uploaded_file = st.file_uploader("Pilih file", type=["zip",'xlsx'])

if uploaded_file is not None:
  if st.button('Process'):
      with st.spinner('Data sedang diproses...'):
        if selected_option == 'REKAP MENTAH':
            with tempfile.TemporaryDirectory() as tmpdirname:
                # Ekstrak file ZIP ke direktori sementara
                with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                    zip_ref.extractall(tmpdirname)
                  
                dfs=[]
                for file in os.listdir(tmpdirname):
                    if file.endswith('.xlsx'):
                            df = pd.read_excel(tmpdirname+'/'+file, sheet_name='REKAP MENTAH')
                            if 'NAMA RESTO' not in df.columns:
                                df = df.loc[:,[x for x in df.columns if 'Unnamed' not in str(x)][:-1]].fillna('')
                                df['NAMA RESTO'] = file.split('-')[0]
                            dfs.append(df)
                      
                dfs = pd.concat(dfs, ignore_index=True)
                excel_data = to_excel(dfs, sheet_name="REKAP MENTAH")
                st.download_button(
                    label="Download Excel",
                    data=excel_data,
                    file_name=f'LAPORAN SO HARIAN RESTO_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )   

        if selected_option == 'REKAP PENYESUAIAN INPUTAN IA':
            nama_file = uploaded_file.name.replace('.zip','')
            db_ia = load_excel(save_path)
            with tempfile.TemporaryDirectory() as tmpdirname:
                # Ekstrak file ZIP ke direktori sementara
                with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                    zip_ref.extractall(tmpdirname)
                non_com = ['SUPPLIES [OTHERS]','00.COST','20.ASSET.ASSET','21.COST.ASSET']
                concatenated_df = []
             
                for file in os.listdir(tmpdirname):
                    if file.startswith('4217'):
                        df_4217     =   pd.read_excel(tmpdirname+'/'+file, header=4).fillna('')
                        df_4217 = df_4217.drop(columns=[x for x in df_4217.reset_index().T[(df_4217.reset_index().T[1]=='')].index if 'Unnamed' in x])
                        df_4217.columns = df_4217.T.reset_index()['index'].apply(lambda x: np.nan if 'Unnamed' in x else x).ffill().values
                        df_4217 = df_4217.iloc[1:,:-3]

                        df_melted =pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang'],
                                            value_vars=df_4217.columns[6:].values,
                                            var_name='Nama Cabang', value_name='Total Stok').reset_index(drop=True)

                        df_melted2 = pd.melt(pd.melt(df_4217, id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Satuan #1','Satuan #2','Satuan #3'],
                                            value_vars=df_4217.columns[6:].values,
                                            var_name='Nama Cabang', value_name='Total Stok').drop_duplicates(),
                                            id_vars=['Kode Barang', 'Nama Barang','Kategori Barang','Nama Cabang','Total Stok'],
                                            var_name='Variabel', value_name='Satuan')

                        df_melted2 = df_melted2[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Satuan','Variabel']].drop_duplicates().reset_index(drop=True)

                        df_melted = df_melted.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)
                        df_melted2 = df_melted2.sort_values(['Kode Barang','Nama Cabang']).reset_index(drop=True)

                        df_4217_final = pd.concat([df_melted2, df_melted[['Total Stok']]], axis=1)
                        df_4217_final = df_4217_final[['Kode Barang','Nama Barang','Kategori Barang','Nama Cabang','Variabel','Satuan','Total Stok']]
                        df_4217_final['Kode Barang'] = df_4217_final['Kode Barang'].astype('int')
                        df_4217_final['Total Stok'] = df_4217_final['Total Stok'].astype('float')

                        df_4217_final=df_4217_final[df_4217_final['Variabel'] == "Satuan #1"].rename(columns={"Total Stok":"Saldo Akhir"})

                        #df_4217_final.insert(0, 'No. Urut', range(1, len(df_4217_final) + 1))

                        def format_nama_cabang(cabang):
                            match1 = re.match(r"\((\d+),\s*([A-Z]+)\)", cabang)
                            if match1:
                                return f"{match1.group(1)}.{match1.group(2)}"
                            else:
                                match2 = re.match(r"^(\d+)\..*?\((.*?)\)$", cabang)
                                if match2:
                                    return f"{match2.group(1)}.{match2.group(2)}"
                                else:
                                    return cabang

                        df_4217_final['Cabang'] = df_4217_final['Nama Cabang'].apply(format_nama_cabang)

                        #df_4217_final=df_4217_final.loc[:,["No. Urut", "Kategori Barang","Kode Barang","Nama Barang","Satuan","Saldo Akhir", "Cabang"]]
                        concatenated_df.append(df_4217_final)
                    else:
                        df_so = pd.read_excel(tmpdirname+'/'+file)
                        df_so['CABANG'] = df_so['CABANG'].str.upper().str[:6]

                df_4217 = pd.concat(concatenated_df)
                df_4217['CABANG'] = df_4217['Cabang'].str[-6:]
                df_4217 = df_4217.merge(df_so, left_on=['CABANG','Nama Barang'], right_on=['CABANG','NAMA BARANG'], how='left').drop(columns=['CABANG','NAMA BARANG'])
                df_4217['#Hasil Stock Opname'] = df_4217['#Hasil Stock Opname'].fillna(0)
                df_4217['DEVIASI(Rumus)'] = df_4217['Saldo Akhir'] - df_4217['#Hasil Stock Opname']
                df_4217 = df_4217[df_4217['DEVIASI(Rumus)']!=0].reset_index()
                df_4217['Tipe Penyesuaian'] = ''
                df_4217.loc[df_4217[df_4217['DEVIASI(Rumus)']>0].index, 'Tipe Penyesuaian'] = 'Pengurangan'
                df_4217.loc[df_4217[df_4217['DEVIASI(Rumus)']<0].index, 'Tipe Penyesuaian'] = 'Penambahan'
                df_4217['DEVIASI(Rumus)'] = df_4217['DEVIASI(Rumus)'].abs()

                for cab in df_4217['Cabang'].unique():
                    folder = f'{tmpdirname}/{nama_file}/{df_4217[df_4217['Cabang']==cab]['Nama Cabang'].iloc[0,]}'
                    if not os.path.exists(folder):
                        os.makedirs(folder)
                    for kat in db_ia['KATEGORI'].unique():
                        if kat in ['Raw Material', 'Packaging']:
                            df_ia = df_4217[(df_4217['Kategori Barang'].isin(db_ia[db_ia['KATEGORI']==kat]['FILTER'])) 
                                            & ~(df_4217['Nama Barang'].isin(db_ia[db_ia['KATEGORI']=='Consume']['FILTER']))
                                            & (df_4217['Cabang']==cab)] 
                            df_ia = df_ia.rename(columns={'Kode Barang':'Kode','Satuan':'UNIT','DEVIASI(Rumus)':'Kuantitas','Nama Cabang':'Gudang'}).loc[:,['Nama Barang','Kode','UNIT','Kuantitas','Gudang','Tipe Penyesuaian']]
                            if not df_ia.empty:
                                df_ia.to_excel(f'{folder}/{kat}_{cab}_{nama_file}.xlsx', index=False)
                        if kat in ['Consume']:
                            df_ia = df_4217[(df_4217['Nama Barang'].isin(db_ia[db_ia['KATEGORI']==kat]['FILTER']))
                                            & (df_4217['Cabang']==cab)] 
                            df_ia = df_ia.rename(columns={'Kode Barang':'Kode','Satuan':'UNIT','DEVIASI(Rumus)':'Kuantitas','Nama Cabang':'Gudang'}).loc[:,['Nama Barang','Kode','UNIT','Kuantitas','Gudang','Tipe Penyesuaian']]
                            if not df_ia.empty:
                                df_ia.to_excel(f'{folder}/{kat}_{cab}_{nama_file}.xlsx', index=False)
                        else:         
                            df_ia = df_4217[(df_4217['Nama Barang'].isin(db_ia[db_ia['KATEGORI']==kat]['FILTER']))
                                            & (df_4217['Cabang']==cab) & (df_4217['Kategori Barang'].isin(non_com))] 
                            df_ia = df_ia.rename(columns={'Kode Barang':'Kode','Satuan':'UNIT','DEVIASI(Rumus)':'Kuantitas','Nama Cabang':'Gudang'}).loc[:,['Nama Barang','Kode','UNIT','Kuantitas','Gudang','Tipe Penyesuaian']]
                            if not df_ia.empty:
                                df_ia.to_excel(f'{folder}/{kat}_{cab}_{nama_file}.xlsx', index=False)

                folder_path = f'{tmpdirname}/{nama_file}'
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    for root, dirs, files in os.walk(folder_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, start=folder_path)
                            zip_file.write(file_path, arcname)

                # Pindahkan ke awal buffer agar bisa dibaca
                zip_buffer.seek(0)
                st.download_button(
                    label="Download Zip",
                    data=zip_buffer,
                    file_name=f"REKAP PENYESUAIAN STOK (IA)_{nama_file}_{get_current_time_gmt7()}.zip",
                    mime="application/zip"
                )
             
        if selected_option == 'PROMIX':
                df_promix = pd.read_excel(uploaded_file,header=1)
                df_cab = pd.read_excel(uploaded_file,header=2).dropna(subset=df_promix.iloc[0,0]).iloc[:,:5].drop_duplicates()
                df_promix = df_promix.T
                df_promix[0] = df_promix[0].ffill()
                df_promix = df_promix.reset_index()
                df_promix['index'] = df_promix['index'].apply(lambda x: np.nan if 'Unnamed' in str(x) else x).ffill()
                df_promix.columns = df_promix.loc[0,:].fillna('')
                df_promix = df_promix.iloc[5:,:].groupby(df_promix.columns[:3].to_list())[df_promix.columns[3:]].sum().reset_index()
                df_promix = df_promix.melt(id_vars=df_promix.columns[:3], value_vars=df_promix.columns[3:])
                df_promix.columns = ['TANGGAL','NAMA BAHAN','SUMBER','CABANG','QTY']
                df_promix = df_promix.merge(df_cab,
                                how='left', left_on='CABANG', right_on=df_cab.columns[0]).drop(columns='CABANG').iloc[:,[0,4,5,6,7,8,1,2,3]]
                st.download_button(
                        label="Download Excel",
                        data=to_excel(df_promix),
                        file_name=f'promix_{get_current_time_gmt7()}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                    )
         
        if selected_option == 'REKAP DATA 42.02':
            with tempfile.TemporaryDirectory() as tmpdirname:
               # Ekstrak file ZIP ke folder sementara
               with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                   zip_ref.extractall(tmpdirname)
               
               all_dfs = []
               for filename in os.listdir(tmpdirname):
                   if filename.endswith('.xlsx'):
                       file_path = os.path.join(tmpdirname, filename)
       
                       # Ambil nama file dan ekstrak kode cabang
                       match = re.search(r'_(\d{4}\.[A-Z]+)', filename)
                       cabang = match.group(1) if match else ''
       
                       # Baca Excel: header ke-5 (index ke-4)
                       df = pd.read_excel(file_path, header=4).fillna('')
                       df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
                       df['Cabang'] = cabang
       
                       all_dfs.append(df)
    
               if all_dfs:
                   df_combined = pd.concat(all_dfs, ignore_index=True)
       
                   # Tombol download hasil
                   st.download_button(
                       label="Download Gabungan Excel",
                       data=to_excel(df_combined),
                       file_name=f'42.02 Combine_{get_current_time_gmt7()}.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                   )
               else:
                   st.warning("Tidak ada file .xlsx ditemukan dalam ZIP.")

        if selected_option == 'WEBSMART (DINE IN/TAKEAWAY)':
            with tempfile.TemporaryDirectory() as tmpdirname:
               # Ekstrak file ZIP ke folder sementara
               with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                   zip_ref.extractall(tmpdirname)
               
               all_dfs = [] 
               for filename in os.listdir(tmpdirname):
                   if filename.endswith('.xls'):
                        file_path = os.path.join(tmpdirname, filename)
                        df_qty = pd.read_html(f'{file_path}')[0]
                        df_qty = df_qty[df_qty.iloc[1:,].columns[df_qty.iloc[1:,].apply(lambda col: col.astype(str).str.contains('QTY', case=False, na=False).any())]]
                        df_qty.iloc[0,:] = df_qty.iloc[0,:] + '_QTY'
                        df_qty.columns = df_qty.iloc[0,:]
                        try:
                            df_qty = df_qty.iloc[2:,:-2].iloc[:-1].drop(columns='OFFLINE_QTY')
                        except:
                            df_qty = df_qty.iloc[2:,:-2].iloc[:-1]
                        df_qty.iloc[:,1:] = df_qty.iloc[:,1:].astype(float)
                        df_rp = pd.read_html(f'{file_path}')[0]
                        df_rp = df_rp[df_rp.iloc[1:,].columns[df_rp.iloc[1:,].apply(lambda col: col.astype(str).str.contains('Rp', case=False, na=False).any())]]
                        df_rp.iloc[0,1:] = df_rp.iloc[0,1:] + '_RP'
                        df_rp.columns = df_rp.iloc[0,:]
                        try:
                            df_rp = df_rp.iloc[2:,:-2].iloc[:-1].drop(columns='OFFLINE_QTY')
                        except:
                            df_rp = df_rp.iloc[2:,:-2].iloc[:-1]
                        df_rp.iloc[:,1:] = df_rp.iloc[:,1:].astype(float)
                        df = df_rp.reset_index().merge(df_qty.reset_index(), on='index')
                        df = df.melt(id_vars=['RESTO'], value_vars=df.iloc[:,1:].columns,value_name='RP',var_name='CATEGORY')
                        df[['CATEGORY','Variable']] = df['CATEGORY'].str.split('_', n=1, expand=True)
                        df = df[df['Variable']=='QTY'].reset_index(drop=True).reset_index().merge(df[df['Variable']=='RP'].reset_index(drop=True).reset_index(), on=['index','RESTO','CATEGORY']).drop(columns=['index','Variable_x','Variable_y']).rename(columns={'RP_x':'QTY','RP_y':'VALUE'})
                        df['TYPE'] = df['CATEGORY'].apply(lambda x: x if x=='DINE IN' else 'TAKE AWAY')
                        df['MONTH'] = re.findall(r'_(\w+)', filename)[-1]
                        df = df.sort_values(['RESTO','TYPE']).reset_index(drop=True)
                        all_dfs.append(df)
    
               if all_dfs:
                   df_combined = pd.concat(all_dfs, ignore_index=True)
       
                   # Tombol download hasil
                   st.download_button(
                       label="Download Gabungan Excel",
                       data=to_excel(df_combined),
                       file_name=f'WEBSMART (DINE IN/TAKEAWAY) Combine_{get_current_time_gmt7()}.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                   )

        if selected_option == 'PENYESUAIAN IA':
            with tempfile.TemporaryDirectory() as tmpdirname:
                with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                    zip_ref.extractall(tmpdirname)
        
                    all_dfs = []
                    for file_info in zip_ref.infolist():
                        if file_info.filename.endswith(('.xlsx', '.xls', '.csv')):
                            with zip_ref.open(file_info.filename) as file:
                                try:
                                    if file_info.filename.endswith('.csv'):
                                        df = pd.read_csv(file, skiprows=9)
                                    else:
                                        df = pd.read_excel(file, skiprows=9)
        
                                    df.drop(df.columns[0], axis=1, inplace=True)
                                    if 'Kode' in df.columns:
                                        df = df[~df['Kode'].isin(['Penyesuaian Persediaan', 'Keterangan', 'Kode'])]
                                        df = df[df['Kode'].notna()]
                                    if 'Tipe' in df.columns:
                                        mask_penambahan = df['Tipe'].str.lower() == 'penambahan'
                                        for col in ['Kts.', 'Total Biaya']:
                                            if col in df.columns:
                                                df.loc[mask_penambahan, col] = df.loc[mask_penambahan, col] * -1
                                    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
                                    if 'Gudang' in df.columns:
                                        def transform_gudang(val):
                                            try:
                                                prefix = re.search(r'^(\d+)', str(val))
                                                kode = re.search(r'\((.*?)\)', str(val))
                                                if prefix and kode:
                                                    return f"{prefix.group(1)}.{kode.group(1)}"
                                                else:
                                                    return val
                                            except:
                                                return val
                                        df['Gudang'] = df['Gudang'].apply(transform_gudang)
                                    all_dfs.append(df)
                                except Exception as e:
                                    st.error(f"Error processing {file_info.filename}: {e}")
        
            if all_dfs:
                final_df = pd.concat(all_dfs, ignore_index=True)
        
                # Load Harga
                try:
                    harga = pd.read_excel(f"{tmpdirname}/Penyesuaian IA/bahan/Harga.xlsx", skiprows=4).fillna("").drop(
                        columns={'Kategori Barang', 'Kode Barang', 'Nama Satuan', 'Saldo Awal', 'Masuk', 'Keluar'})
                except Exception as e:
                    st.error(f"Gagal membaca Harga.xlsx: {e}")
                    st.stop()
        
                harga = harga[~harga['Nama Barang'].isin(['Total Nama Barang', ''])].rename(
                    columns={'Saldo Akhir': 'Kuantitas', 'Unnamed: 14': 'Nilai'})
                harga = harga[harga['Nama Barang'].notna()]
                harga = harga.loc[:, ~harga.columns.str.startswith('Unnamed')]
        
                def safe_divide(row):
                    try:
                        return abs(row['Nilai'] / row['Kuantitas'])
                    except:
                        return 0
        
                harga['Harga'] = harga.apply(safe_divide, axis=1)
        
                # Cari folder REKAP PENYESUAIAN STOK (IA) secara dinamis
                rekap_path = ""
                for root, dirs, files in os.walk(tmpdirname):
                    for dir_name in dirs:
                        if dir_name.strip().lower() == "rekap penyesuaian stok (ia)":
                            rekap_path = os.path.join(root, dir_name)
                            break
                    if rekap_path:
                        break
        
                if not rekap_path:
                    st.error("Folder 'REKAP PENYESUAIAN STOK (IA)' tidak ditemukan di dalam ZIP.")
                    st.stop()
        
                # Gabungkan semua file dari folder tersebut
                all_files = []
                for root, dirs, files in os.walk(rekap_path):
                    for file in files:
                        if file.endswith('.xlsx') or file.endswith('.xls'):
                            all_files.append(os.path.join(root, file))
        
                combined_df = pd.DataFrame()
                for file in all_files:
                    try:
                        df = pd.read_excel(file)
                        combined_df = pd.concat([combined_df, df], ignore_index=True)
                    except Exception as e:
                        print(f"Gagal membaca file: {file} karena {e}")
        
                def format_gudang(g):
                    match = re.match(r"(\d+)(?:\.\d+)?-.*\((\w+)\)", str(g))
                    if match:
                        return f"{match.group(1)}.{match.group(2)}"
                    return g
        
                combined_df['Gudang'] = combined_df['Gudang'].apply(format_gudang)
        
                final_df['Kode'] = final_df['Kode'].astype(str)
                combined_df['Kode'] = combined_df['Kode'].astype(str)
        
                final_df = final_df.merge(
                    combined_df[['Kode', 'Gudang', 'Kuantitas']],
                    on=['Kode', 'Gudang'],
                    how='left'
                )
                final_df.rename(columns={'Kuantitas': 'REKAP'}, inplace=True)
                final_df['Kts.'] = final_df['Kts.'].astype(int)
                final_df['REKAP'] = final_df['REKAP'].astype(int)
                final_df['Total Biaya'] = final_df['Total Biaya'].astype(int)
        
                def selisih(row):
                    try:
                        return row['REKAP'] - row['Kts.']
                    except:
                        return 0
        
                final_df['SELISIH'] = final_df.apply(selisih, axis=1)
        
                final_df['Nama Barang'] = final_df['Nama Barang'].astype(str)
                harga['Nama Barang'] = harga['Nama Barang'].astype(str)
        
                final_df = final_df.merge(
                    harga[['Nama Barang', 'Harga']],
                    on='Nama Barang',
                    how='left'
                )
        
                final_df['Harga'] = final_df['Harga'].astype(int)
        
                st.download_button(
                    label="Download Gabungan Excel",
                    data=to_excel(final_df),
                    file_name=f'Penyesuaian IA Combine_{get_current_time_gmt7()}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
