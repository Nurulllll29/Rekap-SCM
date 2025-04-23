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

selected_option = st.selectbox("Pilih salah satu:", ['LAPORAN SO HARIAN','PROMIX','REKAP SO'])
if selected_option == 'LAPORAN SO HARIAN':
    st.write('Upload file format *zip')
if selected_option == 'PROMIX':
    st.write('Upload file format *xlsx')
if selected_option == 'REKAP SO':
    st.write('Upload file format *zip')
    
uploaded_file = st.file_uploader("Pilih file", type=["zip",'xlsx'])
if uploaded_file is not None:
  if st.button('Process'):
      with st.spinner('Data sedang diproses...'):
        if selected_option == 'LAPORAN SO HARIAN':
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
         
        if selected_option == 'REKAP SO':
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
