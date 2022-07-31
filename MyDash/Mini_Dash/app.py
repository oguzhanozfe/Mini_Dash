from cgitb import enable
import pandas as pd
import streamlit as st
from st_aggrid import AgGrid
import xlsxwriter
from io import BytesIO


output = BytesIO()

st.set_page_config(page_title="MyCooms Data",
                    page_icon=":bar_chart:",
                    layout="wide")

st.header("MyCooms Ariza Data")


excel_file1 = 'data.xlsx'
excel_file2 = 'Mycooms_eleman_listes.xlsx'
sheet_name1 ='DETAY'
sheet_name2 ='Sheet1'
df = pd.read_excel(excel_file1, sheet_name= sheet_name1)
df_eleman = pd.read_excel(excel_file2, sheet_name= sheet_name2)


st.sidebar.header("Filter")
musteri_list = st.sidebar.multiselect("Müşteri Seçiniz",
                                    options = df['MUSTERI'].unique(),
                                    default = df['MUSTERI'].unique())


il = df[df['MUSTERI'].isin(musteri_list)]


il_list = st.sidebar.multiselect("Il Seçiniz",
                                    options = il['IL'].unique(),
                                    default = il['IL'].unique())

ilce = il[il['IL'].isin(il_list)]

ilce_list = st.sidebar.multiselect("Ilce Seçiniz",
                                    options = ilce['ILCE'].unique(),
                                    default = ilce['ILCE'].unique())
col_list = list(df)
col_list = st.sidebar.multiselect("Sutun Seçiniz",
                                    options = col_list,
                                    default = ['MUSTERI','IL','ILCE','TERMINAL_ADI','IS','TERM_ID','serIno'])

new_df = df.query(
    "MUSTERI in @musteri_list and IL in @il_list and ILCE in @ilce_list"
)

new_df = new_df[col_list]

response = AgGrid(
    new_df,
    fit_columns_on_grid_load=False,
    enable_exporting=True,
    theme="dark",
    enable_enterprise_modules=True,
    enable_range_selection=True,
    editable=True
)


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    format1 = workbook.add_format({'num_format': '0.00'}) 
    worksheet.set_column('A:A', None, format1)  
    writer.save()
    processed_data = output.getvalue()
    return processed_data
df_xlsx = to_excel(response['data'])
st.sidebar.download_button(label='📥 Download Current Result',
                                data=df_xlsx ,
                                file_name= 'df_test.xlsx')

grid2 = AgGrid(
    df_eleman,
    editable=True,
    fit_columns_on_grid_load=True,

    theme="dark"
)









