import streamlit as st
from streamlit_folium import folium_static
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from datetime import time, timedelta, datetime, date
import math
import xlrd
import folium

st.set_page_config(page_title="Planner interventi", layout='wide')

col1, col2 = st.columns([4,1])
with col1:
    st.title('Visualizzazione programma di lavoro')

with col2:
    st.image('https://github.com/alebelluco/Test_EX/blob/main/exera_logo.png?raw=True')

percorso = st.sidebar.file_uploader("Caricare il programma Ferrara")
if not percorso:
    st.stop()
visualizzazione=pd.read_csv(percorso)
#st.write(df)


tab1, tab2, tab3, tab4 = st.tabs(['FERRARA | Visualizzazione per operatore','FERRARA | Visualizzazione per data - tutti gli op', 'FERRARA | Ricerca per cliente', 'ALTRI SITI'])

with tab1:
    operatore_unico = visualizzazione.operatore.unique()
    data_unica = visualizzazione.data.unique()
    scelta_op = st.selectbox('Selezionare operatore', operatore_unico)
    scelta_data = st.selectbox('Selezionare data', data_unica)
    st.dataframe(visualizzazione[(visualizzazione.operatore == scelta_op) & (visualizzazione.data == scelta_data)])
    
with tab2:

    pivot=pd.pivot_table(visualizzazione[['data','Cliente','Sito','Servizio','operatore','Durata_stimata']],
                          index=['data','Cliente','Sito','Servizio'],
                          values = 'Durata_stimata',
                          columns='operatore',                       
                          aggfunc='sum',
                          #fill_value = '-',
                          
                          ).reset_index()
    
    scelta_data2 = st.selectbox('Selezionare data', data_unica, key='data2')
    st.dataframe(pivot[pivot['data']==scelta_data2].dropna(axis='columns',how='all').fillna('-'), width=2500)

with tab3:
    cliente_unico = visualizzazione.Cliente.unique()
    scelta_cliente = st.selectbox('Selezionare Cliente', cliente_unico, key='op')
    st.dataframe(visualizzazione[visualizzazione.Cliente == scelta_cliente])

with tab4:
    vincolo_nop = visualizzazione[['ID','N_op','Op_vincolo']]
    
    percorso_altri = st.sidebar.file_uploader("Caricare altri siti")
    if not percorso_altri:
        st.stop()
    altri_siti=pd.read_csv(percorso_altri)

    #altri_siti  = altri_siti.merge(vincolo_nop, how='left',left_on='ID',right_on='ID')

    altri_siti['Durata_stimata'] = altri_siti['Durata_stimata'].str.replace(',','.')
    altri_siti['lat'] = altri_siti['lat'].str.replace(',','.')
    altri_siti['lng'] = altri_siti['lng'].str.replace(',','.')
    altri_siti['Durata_stimata'] = altri_siti['Durata_stimata'].astype(float)
    siti_unici = altri_siti['SitoTerritoriale'].unique()
    #scelta_sito = st.selectbox('Selezionare Sito', siti_unici)
    scelta_sito = st.multiselect('Selezionare Sito', siti_unici)

    altri_siti = altri_siti[[any(sito in word for sito in scelta_sito) for word in altri_siti.SitoTerritoriale.astype(str)]]
    altri_siti = altri_siti.rename(columns={'N_op_x':'N_op','Op_vincolo_x':'Op_vincolo'})
    st.dataframe(altri_siti[altri_siti.columns[1:18]].drop('Note', axis=1), width=2500)
    st.subheader('{:0.2f} ore totali di intervento sul sito territoriale'.format(altri_siti.Durata_stimata.sum()/60)) 
    altri_siti  = altri_siti.merge(vincolo_nop, how='left',left_on='ID',right_on='ID')
    siti_unici = altri_siti['SitoTerritoriale'].unique()

    mappa=folium.Map(location=(44.813057,12.021541),zoom_start=11)
    for i in range(len(altri_siti)):
        try:
            folium.Marker(location=(altri_siti.lat.iloc[i],altri_siti.lng.iloc[i]),popup = altri_siti.Cliente.iloc[i]).add_to(mappa)
        except:
            st.write('Cliente {} non visibile sulla mappa per mancanza di coordinate su Byron'.format(altri_siti.Cliente.iloc[i]))
            pass
    folium_static(mappa,width=1800,height=800)





