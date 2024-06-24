import streamlit as st
import os
from pathlib import Path
import pandas as pd
import xlsxwriter
import glob
import shutil
import zipfile




st.set_page_config(layout="wide")





user_input_excel = st.sidebar.file_uploader("Upload database excel", type=['csv','xlsx'], accept_multiple_files=False, key='file_uploader')
user_input_folder = st.sidebar.file_uploader("Upload pdf folder", type=['zip'], accept_multiple_files=False, key='file_uploader_2')


if user_input_excel is not None:
    if user_input_excel.name.endswith('.xlsx'):
        df=pd.read_excel(user_input_excel, dtype='string')
        st.sidebar.success('Database File Uploaded Successfully!')
        # st.sidebar.write(df)
    else:
        st.sidebar.warning('You need to upload an excel file')

    if user_input_folder is not None:
        if user_input_folder.name.endswith('.zip'):
            target_path = os.path.join(os.getcwd(), os.path.splitext(user_input_folder.name)[0])
            if os.path.exists(target_path) == False:
                os.mkdir(target_path)
            with zipfile.ZipFile(user_input_folder, 'r') as z:
                z.extractall(target_path)
            st.sidebar.success('Folder Uploaded Successfully!')
        else:
            st.sidebar.warning('You need to upload zip type file')
        
        # tab1, tab2 = st.tabs(["Setting", "Run Apps"])
        # with tab1 :


        col_1, col_2, col_3, col_4, col_5 = st.columns(5)
        with st.container():
            with col_1:
                user_input_npwp = 'IDENTITAS_PENERIMA_PENGHASILAN'
        with st.container():
            with col_2:
                user_input_perusahaan = 'NAMA_PENERIMA_PENGHASILAN'
        with st.container():
            with col_3:
                user_input_masa_pajak = 'MASA_PAJAK'
        with st.container():
            with col_4:
                user_input_tahun_pajak = 'TAHUN_PAJAK'
        with st.container():
            with col_5:
                user_input_ID = 'ID_SISTEM'

        submit_button_clicked = st.button("Submit", type="primary", use_container_width=True)

        if submit_button_clicked :


            a = os.listdir(os.path.join(os.getcwd(),os.path.splitext(user_input_folder.name)[0]))
            lst = []
            for x in a :
                lst.append(os.path.splitext(x)[0][-36 :])

            if len(lst) != len(df) :
                st.error('Data length not matched!')


            for i in range(len(lst)):
                matching_index = df.index[df[user_input_ID] == lst[i]]
                nama_perusahaan = []
                npwp_perusahaan = []
                nama_perusahaan = df.loc[matching_index, user_input_perusahaan]
                npwp_perusahaan = df.loc[matching_index, user_input_npwp]
                st.write(str(nama_perusahaan))
                nama_npwp_perusahaan = str(nama_perusahaan.item()) + ' (' + str(npwp_perusahaan.item()) + ')'
                tahun_pajak = df.loc[matching_index, user_input_tahun_pajak]
                masa_pajak = df.loc[matching_index, user_input_masa_pajak]
                tahun_pajak = []
                masa_pajak = []
                tahun_masa_pajak = str(tahun_pajak.item()) + '-' + str(masa_pajak.item())

                result_path = os.path.join(os.getcwd(),'Result')
                if os.path.exists(result_path) == False:
                    os.mkdir(result_path)
                if os.path.exists(os.path.join(result_path, nama_npwp_perusahaan)) == False:
                    os.mkdir(os.path.join(result_path, nama_npwp_perusahaan))
                if os.path.exists(os.path.join(result_path, nama_npwp_perusahaan, tahun_masa_pajak)) == False:
                    os.mkdir(os.path.join(result_path, nama_npwp_perusahaan, tahun_masa_pajak))
                path_to_save = os.path.join(result_path, nama_npwp_perusahaan, tahun_masa_pajak)
                
                shutil.copy(glob.glob(os.path.join(os.getcwd(),os.path.splitext(user_input_folder.name)[0],'*pdf'))[i], path_to_save)
            
            # st.write(os.listdir(os.getcwd()))

            shutil.make_archive('Result', 'zip', result_path)
            
            result_path_zipped = os.path.join(os.getcwd(),'Result.zip')

            with open(result_path_zipped, "rb") as fp :
                button_clicked = st.download_button(label=':cloud: Download Result', type="secondary", data=fp, file_name='Result.zip', mime="application/zip")




    else :
        st.error("You have to upload pdf folder in the sidebar")

else :
    st.error("You have to upload a csv or an excel file in the sidebar")




