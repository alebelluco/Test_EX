import pandas as pd
import streamlit as st

st.set_page_config(page_title="Test", layout='wide')

percorso = st.sidebar.file_uploader(f"Caricare l' elenco interventi estratto da Byron")
while percorso is None:
    st.stop()

df_raw=pd.read_excel(percorso)

salvataggio = st.text_input('incollare percorso file')


df_raw[0:10].to_excel('{}\output.xlsx'.format(salvataggio))


