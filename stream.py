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

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

        # Mengakses workbook dan worksheet untuk format header
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
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
    
st.title('Rekap SCM')
uploaded_file = st.file_uploader("Pilih file ZIP", type="zip")

if uploaded_file is not None:
  if st.button('Process'):
    with tempfile.TemporaryDirectory() as tmpdirname:
        # Ekstrak file ZIP ke direktori sementara
        with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
            zip_ref.extractall(tmpdirname)
          
        dfs=[]
        for file in os.listdir(tmpdirname):
            if file.endswith('.xlsx'):
                    df = pd.read_excel(tmpdirname+'/'+file, sheet_name='REKAP MENTAH')
                    df = df.loc[:,[x for x in df.columns if 'Unnamed' not in str(x)][:-1]].fillna('')
                    df['NAMA RESTO'] = file.split('-')[0]
                    dfs.append(df)
              
        dfs = pd.concat(dfs, ignore_index=True)
        excel_data = to_excel(dfs)
        st.download_button(
            label="Download Excel",
            data=excel_data,
            file_name=f'LAPORAN SO HARIAN RESTO_{get_current_time_gmt7()}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )   
