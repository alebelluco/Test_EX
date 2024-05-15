# To-do
# Corretto bug per interventi 2 OP, inserito reseti index df prima di calcolare la lunghezza x inserimento riga sdoppiata
# da correggere: se il secondo op non ha abbastanza tempo gli sdoppia l'intervento e non va bene --- verificare se succede o se era problema di cache

import streamlit as st
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
from datetime import time, timedelta, datetime, date
import math
import xlrd

st.set_page_config(page_title="Planner interventi", layout='wide')

tab1, tab4 = st.tabs(['Scaduto', 'Input programmazione'])

with tab1:
    mesi_scaduto = {1:'Gennaio', 2:'Febbraio',3:'Marzo',4:'Aprile',5:'Maggio',6:'Giugno',7:'Luglio',8:'Agosto',9:'Settembre',10:'Ottobre',11:'Novembre', 12:'Dicembre'}
    key_list_scaduto = list(mesi_scaduto.keys())
    val_list_scaduto = list(mesi_scaduto.values())

    oggi = datetime.now().date()
    def highlight_tipo(s):
        return ['background-color: #818105']*len(s) if s['Cntr']=='TIPO' else ['background-color: black']*len(s)

    percorso_scaduto = st.sidebar.file_uploader(f"Caricare l'elenco interventi estratto da Byron da 1 gennaio a oggi, tutti gli interventi")
    if not percorso_scaduto:
        st.stop()
    df_scaduto=pd.read_excel(percorso_scaduto, engine='xlrd')
    df_scaduto = df_scaduto[(df_scaduto.S != 'C') & ([data.month < oggi.month for data in df_scaduto['Data Inizio']])]
    df_scaduto['Cntr'] = np.where(['TIPO' in word for word in df_scaduto['Descrizione Contratto'].astype(str)],'TIPO',None)
    df_scaduto = df_scaduto[['S','Data Inizio','Cliente','IstruzioniOperative','Descrizione Contratto','Referente/Amm. Condominio','Sito','Citta','Indirizzo Sito','SitoTerritoriale','Servizio','Cntr']]
    df_scaduto_standby = df_scaduto[['(STANDBY)' in istruzione for istruzione in df_scaduto['IstruzioniOperative'].astype(str)]]
    df_scaduto_no_standby = df_scaduto[['(STANDBY)' not in istruzione for istruzione in df_scaduto['IstruzioniOperative'].astype(str)]]


    lista_unica_scaduto = list(set([ data.month for data in df_scaduto['Data Inizio']]))
    selezione_scaduto= [ mesi_scaduto[key] for key in lista_unica_scaduto ]
    mese_scelto_scaduto = st.multiselect('Selezionare mesi', selezione_scaduto)

    if not mese_scelto_scaduto:
        mese_scelto_scaduto = val_list_scaduto

    valore_scaduto = [key_list_scaduto[val_list_scaduto.index(n)] for n in mese_scelto_scaduto ]
    #mese = df_raw[[ data.month in valore for data in df_raw['Data Inizio']]]
    #mese_n = min(valore)


    st.title('Interventi non svolti')


    st.subheader('Interventi in Standby a partire dai mesi selezionati', divider='orange')
                                
    st.dataframe(df_scaduto_standby[[ data.month in valore_scaduto for data in df_scaduto_standby['Data Inizio']]].style.apply(highlight_tipo, axis=1),width=2500)
    st.markdown('{:0.0f} interventi in Standby'.format(len(df_scaduto_standby[[ data.month in valore_scaduto for data in df_scaduto_standby['Data Inizio']]])))
    
    st.subheader('Interventi scaduti da inizio anno non in Standby a partire dai mesi selezionati', divider='orange')
    st.dataframe(df_scaduto_no_standby[[ data.month in valore_scaduto for data in df_scaduto_no_standby['Data Inizio']]].style.apply(highlight_tipo, axis=1),width=2500)
    st.markdown('{:0.0f} interventi non in Standby e non eseguiti'.format(len(df_scaduto_no_standby[[ data.month in valore_scaduto for data in df_scaduto_no_standby['Data Inizio']]])))

with tab4:
    
    path_clara = 'https://github.com/alebelluco/Test_EX/blob/main/Allowed.csv?raw=True'
    #path_clara = 'https://github.com/alebelluco/Test_EX/blob/main/Clara_allowed.csv?raw=True'
    
    #ordine_siti = pd.read_excel('/Users/Alessandro/Documents/AB/Clienti/AB/Exera/Programmazione/Siti_territoriali.xlsx')
    #ordine_siti = pd.read_csv('https://github.com/alebelluco/Test_EX/blob/main/Siti_territoriali.csv', delimiter=';')#, on_bad_lines='skip')
    ordine_siti = pd.read_csv('https://github.com/alebelluco/Test_EX/blob/main/Siti_territoriali2.csv?raw=True', delimiter=';')

    ordine_siti['ViaggioAR'] = [stringa.replace(',','.') for stringa in ordine_siti.ViaggioAR]
    ordine_siti['ViaggioAR'] = ordine_siti['ViaggioAR'].astype(float)

    #siti=pd.read_excel('/Users/Alessandro/Documents/AB/Clienti/AB/Exera/cantieri.xls')
    #siti=pd.read_excel('https://github.com/alebelluco/Exera/blob/main/cantieri.xls')
    #siti = pd.read_csv('/Users/Alessandro/Documents/AB/Clienti/AB/Exera/Programmazione/Flatfiles/cantieri.csv',delimiter=';')
    siti = pd.read_csv('https://github.com/alebelluco/Test_EX/blob/main/cantieri.csv?raw=True',delimiter=';')
    siti = siti[['id sito','lat','lng']]

    anno_corrente = 2024
    # oggi = datetime.now().date()

    # modificare bimensile +- 5 gg
    deltas = {30:3,15:2,10:2,7:1,7.5:1,6:1,1:0,2:0, 4:1}
    viaggi = dict(zip(ordine_siti.SitoTerritoriale, ordine_siti.ViaggioAR))

    def calendario_no_domeniche (mese, anno, operatore): #in realtà non ci sono neanche i sabati
        calendario = []
        inizio = date(anno, mese, 1) 
        calendario = [inizio]
        if mese in [11,4,6,9]:
            for i in range (29):#29
                calendario.append(inizio+timedelta(days=i+1))
        elif mese == 2:
            for i in range (30):#27 metto 30 
                calendario.append(inizio+timedelta(days=i+1))
        else:
            for i in range (30):#30
                calendario.append(inizio+timedelta(days=i+1))
                
        cal=pd.DataFrame(calendario).rename(columns={0:'Data'})
        cal['day']=[ data.weekday() for data in cal.Data]
        cal['day_n']=[ data.day for data in cal.Data]

        for i in range(len(cal)):
            if cal.Data.iloc[i].month != mese_n:
                cal['day_n'].iloc[i]=cal['day_n'].iloc[i-1]
                cal['Data'].iloc[i]=cal['Data'].iloc[i-1]


        cal = cal[(cal.day != 6)]# & (cal.day != 5)]
        cal['operatore']=operatore
        cal = cal.reset_index(drop=True)
                
        return cal

    col_a, col_b = st.columns([4,1])
    with col_a:
        st.title('Schedulatore interventi')
        st.write ('Il cliente FERRARA TUA è escluso dalla schedulazione')
        st.write('Versione aggionata del 04/04/2024 - Modifiche introdotte:')
        st.write('- Correzione Siti Territoriali  dopo modifica Byron')
            
    with col_b:
        #st.image('/Users/Alessandro/Documents/AB/Clienti/AB/Exera/exera_logo.png')
        st.image('https://github.com/alebelluco/Test_EX/blob/main/exera_logo.png?raw=True')

    st.subheader('',divider='grey')

    # Caricamento file estratto da Byron
    percorso = st.sidebar.file_uploader(f"Caricare l' elenco interventi estratto da Byron")
    while percorso is None:
        st.stop()

    df_raw=pd.read_excel(percorso, engine='xlrd')

    # pulizia dati nan
    df_raw.Sito = df_raw.Sito.astype(str).replace({'nan':'non disponibile'})
    df_raw['SitoTerritoriale'] = [word[:-3] for word in df_raw.SitoTerritoriale.astype(str)]

    df_raw['Indirizzo Sito'] = df_raw['Indirizzo Sito'].astype(str).replace({'nan':'non disponibile'})
    df_raw = df_raw.merge(siti,how='left',left_on='IdSito',right_on='id sito')

    #rimozione voci fatturazione
    #fattura = ['FATTURAZIONE','fatturazione','Fatturazione']
    #df_raw = df_raw[[all(voce not in check for voce in fattura) for check in df_raw['IstruzioniOperative'].astype(str)]]
    #rimozione ferrara tua
    df_raw = df_raw[df_raw.Cliente != 'FERRARA TUA SPA']
    df_raw = df_raw[df_raw['Operatore'] != ' (FAT) (FAT)']

    #creazione delle chiavi
    df_raw['key'] = df_raw.Cliente + df_raw.Sito + " | " + df_raw['Indirizzo Sito'] + df_raw.Servizio
    #df_raw['key_sito'] =  df_raw.Cliente + df_raw.Sito + " | " + df_raw['Indirizzo Sito']
    df_raw['key_sito'] =  df_raw.Cliente.astype(str) + df_raw.Sito.astype(str) + df_raw['Indirizzo Sito'].astype(str)
    df_raw['key_univoca'] = df_raw.Cliente + df_raw.Sito + " | " + df_raw['Indirizzo Sito'] + df_raw.Servizio + df_raw.Periodicita



    df_raw = df_raw[df_raw['key'] != 'VETRORESINA SPA sede op, QUARTIERE |  VIA PRAFITTA BERTOLINA 4/A , QUARTIERE 44015 , FESM DEMUSCAZIONE']



    # Suddivisione  dei dataframe in: 
    # Chiuso
    # Pianificato
    # Pianificabile

    filtro_pianificati = [' ANTILARVALE 1',' SQUADRA 2',' SQUADRA 1',' SQUADRA 4',' SQUADRA 3',' SQUADRA 6',' 2 OPERATORI', ' VINCOLO ,' ,' SQUADRA 5' ]

    # creo un dataframe di un mese (mese di pianificazione)

    mesi = {1:'Gennaio', 2:'Febbraio',3:'Marzo',4:'Aprile',5:'Maggio',6:'Giugno',7:'Luglio',8:'Agosto',9:'Settembre',10:'Ottobre',11:'Novembre', 12:'Dicembre'}
    key_list = list(mesi.keys())
    val_list = list(mesi.values())
    lista_unica = list(set([ data.month for data in df_raw['Data Inizio']]))
    selezione = [ mesi[key] for key in lista_unica ]
    mese_scelto = st.multiselect('Selezionare mese da pianificare', selezione)

    while mese_scelto == []:
        st.stop()

    valore = [key_list[val_list.index(n)] for n in mese_scelto ]
    mese = df_raw[[ data.month in valore for data in df_raw['Data Inizio']]]
    mese_n = min(valore)
    

    # Calcolo della prima data pianificabile

    calendario_pianificazione = calendario_no_domeniche(mese_n,anno_corrente,'standard')
    #st.write(calendario_pianificazione.Data.iloc[-1])
    oggi_day = oggi.day#-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
    index_oggi = calendario_pianificazione[calendario_pianificazione.day_n==oggi_day].day_n.index[0]

#next working
    
    if oggi.month != valore[0]:
        next_working=calendario_pianificazione.Data.iloc[0]
    else:
        next_working = calendario_pianificazione.Data.loc[index_oggi+1]#----------------------------------------dopodomani
    data_first = next_working
    data_last = calendario_pianificazione.day_n.iloc[-1]

#correzione no tempo
    
    media_intervento = mese[mese['Min']>0]['Min'].mean()
    mese['Durata_stimata'] = np.where(mese['Min']>0, mese['Min'], media_intervento)

# interventi in standby
    
    standby = df_raw[['(STANDBY)' in nota for nota in df_raw['IstruzioniOperative'].astype(str)]]
    standby = standby[standby.S != 'C']
    mese = mese[['(STANDBY)' not in nota for nota in mese['IstruzioniOperative'].astype(str)]]
   

    chiusi = df_raw[df_raw.S == "C"]
    chiusi_save = chiusi
    chiusi['data_AB'] = pd.to_datetime(chiusi['Data Esecuzione'].str[0:10], format="%d/%m/%Y")
    #chiusi['data_AB'] = pd.to_datetime([data[0:10] for data in chiusi['Data Esecuzione'].astype(str)], format="%d/%m/%Y")

    chiusi = chiusi[['data_AB','key']]
    chiusi = chiusi.reset_index(drop=True)
    len_c = len(chiusi)
    da_c = chiusi.data_AB.min().date()
    a_c = chiusi.data_AB.max().date()

    col1, col2, col3,col10,col20,col30 = st.columns([1,1,1,1,1,1])
    with col1:
        st.title(f':orange[{len_c}]')
        st.write('Interventi chiusi dal {} al {}'.format(da_c,a_c))

# 23/02/2023 inserito intervento N
    pianificabile = mese[(mese.S == "F")  | (mese.S == "F*") | (mese.S == "N")] 
    pianificati = pianificabile[[(all(check not in word for check in filtro_pianificati)) and (word != 'nan') for word in pianificabile.Operatore.astype(str)]]
    pianificati = pianificati[['Data Inizio','key']]
    pianificati = pianificati.rename(columns={'Data Inizio':'data_AB'})
    pianificati = pianificati.reset_index(drop=True)

    len_p = len(pianificati)
    da_p = pianificati.data_AB.min().date()
    a_p = pianificati.data_AB.max().date()

    with col3:
        st.title(f':white[{len_p}]')
        st.write('Interventi pianificati in attesa di esecuzione dal {} al {}'.format(da_p,a_p))

    pianificabile = pianificabile[[(any(check in word for check in filtro_pianificati)) or (word == 'nan') for word in pianificabile.Operatore.astype(str)]]
    pianificabile = pianificabile.reset_index(drop=True)

    pianificabile = pianificabile[['ID','S','Data Inizio','Cliente','Sito','Indirizzo Sito','Dispositivi Installati','SitoTerritoriale','Servizio','ID Contratto','Codice Contratto','Citta',
                                'Operatore','Min','Periodicita','Operatore2','key','key_sito','key_univoca','Durata_stimata','IstruzioniOperative', 'Note','lat','lng','Descrizione Contratto']]

    len_todo = len(pianificabile)
    da_pian = pianificabile['Data Inizio'].min().date()
    a_pian = pianificabile['Data Inizio'].max().date()

    with col20:
        st.title(f':red[{len_todo}]')
        st.write('Interventi da pianificare dal {} al {}'.format(da_pian,a_pian))


    st.subheader('Interventi in standby', divider='orange')
    standby = standby[['S','Cliente','IstruzioniOperative','Referente/Amm. Condominio','Sito','Indirizzo Sito','SitoTerritoriale','Servizio','ID Contratto','Descrizione Contratto','Citta']]
    st.dataframe(standby, width=2500)

    # inizio elaborazione------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
    # i calcoli sulla frequenza degli interventi vengono fatti sul dataframe del mese da pianificare.    
    count_row = mese[['key','Periodicita']].groupby(by='key').count().sort_values(by='Periodicita', ascending = False)
    count_row_distinct = mese[['key','Periodicita']].groupby(by='key').nunique().sort_values(by='Periodicita', ascending = False)
    rank = mese[['key','Periodicita']]
    rank['rank'] = rank.groupby('key').rank(method='dense')
    sito = mese[['key_sito','ID']].groupby(by='key_sito').count()
    chiusi = pd.concat([chiusi,pianificati]) #-------------------------------------considero chiusi + pianificati per calcolare l'ultima data

    chiusi = chiusi.groupby(by='key').max()
    pianificati = pianificati.groupby(by='key').max()

    df_out = rank.merge(count_row, how='left', left_on='key', right_on='key')
    df_out = df_out.rename(columns={'Periodicita_x':'Periodicita', 'Periodicita_y':'Conteggio'})
    df_out = df_out.merge(count_row_distinct, how='left', left_on='key', right_on='key')
    df_out = df_out.rename(columns={'Periodicita_x':'Periodicita', 'Periodicita_y':'Conteggio_distinti'})
    df_out=df_out.merge(chiusi, how='left', left_on='key', right_on='key')
    df_out = df_out.rename(columns={'data_AB':'ultimo_intervento'})
    df_out=df_out.merge(pianificati, how='left', left_on='key', right_on='key')
    df_out = df_out.rename(columns={'data_AB':'ultimo_pianificato'})
    df_out['ultimo_intervento_check'] = df_out.ultimo_intervento.astype(str).replace({'NaT':'Pianificazione libera'})

    df_out['Target'] = np.where(df_out.ultimo_intervento_check != 'Pianificazione libera', df_out.ultimo_intervento + timedelta(days = 30), np.datetime64('NaT'))

    df_out['mese']=[data.month for data in df_out['Target']]
    data_corretta = calendario_pianificazione.Data.iloc[-1]
    for i in range(len(df_out)):
        if df_out.mese.iloc[i]!= mese_n:
            df_out.Target.iloc[i]=pd.to_datetime(data_corretta)

    df_out = df_out.drop_duplicates()
    df_out['key_univoca']=df_out.key + df_out.Periodicita

    output = pianificabile[['ID','S','Data Inizio','Cliente','Sito','Indirizzo Sito','Servizio','key','key_sito','key_univoca','Durata_stimata','Descrizione Contratto','SitoTerritoriale','lat','lng','IstruzioniOperative','Citta']]#------------------
    output = output.merge(df_out,how='left', left_on='key_univoca', right_on='key_univoca')
    output = output.merge(sito, how='left', left_on='key_sito', right_on='key_sito')
    output = output.sort_values(by=['Cliente','Sito','Periodicita','ID_y','rank'])
    output = output.reset_index(drop=True)
    output = output.rename(columns={'ID_y' : 'Servizi sul sito'})

# se il contratto è SPOT, metto di default 2 interventi sul mese, così diventa 15 gg la distanza
    
    output['Conteggio_distinti'] = np.where((['SPOT' in stringa for stringa in output['Descrizione Contratto'].astype(str)]) & (output['Conteggio_distinti'].astype(int) < 2 ) , 2, output.Conteggio_distinti)

    colonne = output.columns
   
    # Dataframe PLANNED--------------------------------------------------------------------------------------------------------------------------
    planned = pd.DataFrame(columns=colonne)
    planned['TGT_new']=[]
    ultimo = {}

    planned['last']=np.datetime64('NaT')
    planned['periodo'] = 1
    planned['range'] = None

    for i in range(len(output)):
        frequenza = output['Conteggio_distinti'].iloc[i]
        periodo = math.floor(30/frequenza)
        rank = output['rank'].iloc[i]
        sito = output['key_sito'].iloc[i]  
        
        if frequenza == 1: # servizi previsti una sola volta nel mese        
            planned.loc[i]=output.iloc[i]
            if output['ultimo_intervento'].astype(str).iloc[i]=='NaT': #---------sono in stato di pianificazione libera e un singolo intervento
            #if output['ultimo_intervento_check'].astype(str).iloc[i]=='Pianificazione libera': #---------sono in stato di pianificazione libera e un singolo intervento    
                planned['TGT_new'].iloc[i]=date(anno_corrente,mese_n,15)
                
            else: #------------distanzio di un mese dall'ultimo
                planned['TGT_new'].loc[i] = output['Target'].iloc[i] #---------distanziato 30gg 
        else:
            chiave = output['key_x'].iloc[i]
            periodo = math.floor(30/frequenza)

#*** Inserisco qui la condizione 15 giorni per i contratti spot??------------------------------- if spot --> periodo = 15

            if output['rank'].iloc[i] == 1: # è il primo degli interventi multpli è il 1 di n
                planned.loc[i]=output.iloc[i]
                if output['ultimo_intervento'].astype(str).iloc[i]=='NaT':
                    planned['TGT_new'].loc[i] = data_first
                    ultimo[chiave] = planned['TGT_new'].loc[i]
                else:
                    last = output['ultimo_pianificato'].iloc[i]
                    if str(last)=='NaT':
                        last = output['ultimo_intervento'].iloc[i]
                    planned['TGT_new'].loc[i] = last + timedelta(days = periodo)
                    planned['last'].loc[i]=str(last)
                    #planned['TGT_new'].loc[i] = data_first # oppure metto il target?
                    ultimo[chiave] = planned['TGT_new'].loc[i]
            else:
                chiave = output['key_x'].loc[i]
                periodo = math.floor(30/frequenza)
                try:
                    last = ultimo[chiave]                 
                except:
                    last = output['ultimo_pianificato'].iloc[i]
                    if str(last)=='NaT':
                        last = output['ultimo_intervento'].iloc[i]

                #if output.key_univoca.iloc[i] == output.key_univoca.iloc[i-1]: # per raggruppare sullo stesso giorno gli interventi dell'università
                #    tgt_new = planned.TGT_new.loc[i-1]
                #else:
                tgt_new = last + timedelta(days = periodo)

                planned.loc[i]=output.iloc[i]
                if tgt_new.month == mese_n:
                    planned['TGT_new'].loc[i] = tgt_new
                else:
                    planned['TGT_new'].loc[i] = calendario_pianificazione.Data.iloc[-1]

                planned['last'].loc[i]=str(last)     
                ultimo[chiave] = tgt_new

        planned['periodo'].loc[i]=periodo
        
    planned = planned.rename(columns={'ID_x':'ID'})
    planned = planned[['ID','S','Data Inizio','Cliente','Sito','Citta','Indirizzo Sito','Servizio','Periodicita','rank','Conteggio','Conteggio_distinti','ultimo_intervento','ultimo_pianificato',
                    'ultimo_intervento_check','Target','Servizi sul sito','TGT_new','last','periodo','key_sito','Durata_stimata','range','Descrizione Contratto','SitoTerritoriale','lat','lng','IstruzioniOperative']]

    st.subheader('Interventi scaduti', divider='orange')

    with st.expander('Visualizza interventi in ritardo'):
    
    #------------------------------------------------------------------------------------------------------da qui inizio il raggruppamento di disponibilità vicine 

        planned['day']=[data.day for data in planned.TGT_new]
        planned['Ritardo']=None

        count_scaduto = 0
        for i in range(len(planned)):   
            if planned['S'].iloc[i]=='F*': # metto il vincolo sulla data degli interventi improrogabili
                disp_intervento = [planned['Data Inizio'].iloc[i].day]
                planned['range'].iloc[i] = disp_intervento

            elif planned['S'].iloc[i]=="N": # vincolo interveenti N
                giorno_tgt = planned['Data Inizio'].iloc[i].day
                disp_intervento = np.arange(giorno_tgt, giorno_tgt + 5, step=1)
                #st.write(disp_intervento)
                disp_adj = [elemento for elemento in disp_intervento if (elemento >= (data_first).day) & (elemento <= data_last)]
                #st.write(disp_adj)
                if disp_adj != []:
                    planned['range'].iloc[i] = disp_adj
                else:                    
                    #if any([elemento <= data_first.day for elemento in disp_intervento]): #se non sono tutti nel futuro
                    planned['range'].iloc[i] = np.arange(data_first.day,data_first.day+3,step=1) # riproposto in un range a partire dal primo giorno disponibile + 3 giorni 
                    if planned.Target.astype(str).iloc[i] != 'nan':
                        if planned.Target.iloc[i].month == mese_n and planned.Target.iloc[i].month is not None :
                            planned['range'].iloc[i] = np.arange(data_first.day,data_first.day+3,step=1) # riproposto in un range a partire dal primo giorno disponibile + 3 giorni 
                            #st.write(f'Attenzione: intervento ID {planned.ID.iloc[i]} | :orange[Cliente: {planned.Cliente.iloc[i]} ]| fuori target - ripianificato ')
                            planned['Ritardo'].iloc[i]='x'
                            st.write(f'Attenzione: intervento ID {planned.ID.iloc[i]}  |  :orange[Cliente: {planned.Cliente.iloc[i]} ]  |  ultimo intervento: :orange[{str(planned.ultimo_intervento.iloc[i])[:10]}]')
                            count_scaduto += 1
                    #else:
                      #  planned['range'].iloc[i] = [elemento for elemento in disp_intervento if (elemento <= data_last)]


            else:   # tutti gli altri interventi
                periodo_intervento = planned.periodo.iloc[i]
                delta = deltas[periodo_intervento]
                giorno_tgt = planned.day.iloc[i]
                disp_intervento = np.arange(giorno_tgt-delta, giorno_tgt + delta, step=1)
                disp_adj = [elemento for elemento in disp_intervento if (elemento >= (data_first).day) & (elemento <= data_last) ] # per mettere prima data pianificata da oggi in poi
                
                if disp_adj != []:
                    planned['range'].iloc[i] = disp_adj
                else:
                    if planned['ultimo_intervento_check'].iloc[i]!='Pianificazione libera':
                        if any([elemento <= data_first.day for elemento in disp_intervento]): #se non  sono tutti nel futuro
                            planned['range'].iloc[i] = np.arange(data_first.day,data_first.day+3,step=1) # riproposto in un range a partire dal primo giorno disponibile + 3 giorni 
                            if planned.Target.astype(str).iloc[i] != 'nan':
                                if planned.Target.iloc[i].month == mese_n and planned.Target.iloc[i].month is not None :
                                    planned['range'].iloc[i] = np.arange(data_first.day,data_first.day+3,step=1) # riproposto in un range a partire dal primo giorno disponibile + 3 giorni 
                                    st.write(f'Attenzione: intervento ID {planned.ID.iloc[i]}  |  :orange[Cliente: {planned.Cliente.iloc[i]} ]  |  ultimo intervento: :orange[{str(planned.ultimo_intervento.iloc[i])[:10]}]')
                                    planned['Ritardo'].iloc[i]='x'
                                    count_scaduto += 1
                        else:
                            planned['range'].iloc[i] = [elemento for elemento in disp_intervento if (elemento <= data_last)]
            
            if (planned['ultimo_intervento_check'].iloc[i]=='Pianificazione libera') and (planned['periodo'].iloc[i]==30): # se non c'è vincolo metto disponibili tutti i giorni
                disp_intervento = np.arange(1,31, step=1)
                
                disp_adj = [elemento for elemento in disp_intervento if elemento >= (data_first).day]
                
                planned['range'].iloc[i] = disp_adj

            if planned['S'].iloc[i]=='F*': # metto il vincolo sulla data degli interventi improrogabili
                disp_intervento = [planned['Data Inizio'].iloc[i].day]
                planned['range'].iloc[i] = disp_intervento

    
#--------    altri_siti = planned.copy()
#--------    altri_siti['TGT_new']=altri_siti['TGT_new'].astype(str)

    st.write(':red[Attenzione {} interventi in ritardo]'.format(count_scaduto))
    planned_print=planned[['ID','Cliente','Indirizzo Sito','Servizio','Periodicita','ultimo_intervento','periodo','range']]
    #planned_print['TGT_new']=planned_print['TGT_new'].astype(str)
    st.divider()
    st.write('**Date disponibili per interventi**')
    st.dataframe(planned_print)


    #–-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    

    #date_target = planned[['ID','available','gruppo']]

    #-----------------------------------------------------------------------------------------------------------------------------------------assegnazione agli operatori

    colonne=['ID','Data Inizio','Cliente', 'Sito', 'Indirizzo Sito', 'SitoTerritoriale','Servizio','Periodicita','ID Contratto', 'Citta','Operatore','Operatore2','Durata_stimata','IstruzioniOperative','Note', 'lat','lng','key_sito','Descrizione Contratto']

    #------------------------------------------------------------------------------------------------------------ Disponibilità siti Clara
    #clara_allowed = pd.read_excel(path_clara).fillna(0)
    clara_allowed = pd.read_csv(path_clara, delimiter=';').fillna(0)
    clara_allowed = clara_allowed[0:963]
    #clara_allowed['key']=clara_allowed.Cliente.astype(str) +' '+ clara_allowed.Sito.astype(str) +' | '+ clara_allowed.Indirizzo_sito.astype(str)

    clara_allowed['key']=clara_allowed.Cliente.astype(str) +' ' +clara_allowed.Sito.astype(str) +' '+ clara_allowed.Indirizzo_sito.astype(str)
 

    clara_allowed['DALLE_M'] = clara_allowed['DALLE_M'].astype(str).str.replace(',','.')
    clara_allowed['ALLE_M'] = clara_allowed['ALLE_M'].astype(str).replace(',','.')
    clara_allowed['DALLE_P'] = clara_allowed['DALLE_P'].astype(str).replace(',','.')
    clara_allowed['ALLE_P'] = clara_allowed['ALLE_P'].astype(str).replace(',','.')

    clara_allowed['draft_orari1'] = [ f'Mattino | dalle {dalle_m} alle 'for dalle_m in clara_allowed.DALLE_M.astype(str) ]
    clara_allowed['draft_orari2'] = [ f'{alle_m}'for alle_m in clara_allowed.ALLE_M.astype(str)]
    clara_allowed['draft_orari3'] = [ f' - Pomeriggio | dalle {dalle_p} alle 'for dalle_p in clara_allowed.DALLE_P.astype(str) ]
    clara_allowed['draft_orari4'] = [ f'{alle_p}'for alle_p in clara_allowed.ALLE_P.astype(str)]

    
    clara_allowed['orari']=np.where((clara_allowed.DALLE_M.astype(str) != '0') & (clara_allowed.DALLE_P.astype(str) != '0.0'), clara_allowed['draft_orari1']+
                                   clara_allowed['draft_orari2']+clara_allowed['draft_orari3']+clara_allowed['draft_orari4'], None )
    
    clara_allowed['orari']=np.where((clara_allowed.DALLE_M.astype(str) != '0') & (clara_allowed.DALLE_P.astype(str) == '0.0') & (clara_allowed['orari'].astype(str) == 'None'),
                                   clara_allowed['draft_orari1'] + clara_allowed['draft_orari2'],clara_allowed['orari'])

    #clara_allowed = clara_allowed[clara_allowed.columns[0:8]]
    clara_allowed['allowed']=None

    for j in range(len(clara_allowed)):
        lista = []
        for i in [0,1,2,3,4,5]:
            #colonna = f'{i}'
            if clara_allowed[str(i)].iloc[j]=='x':
                lista.append(i)
        clara_allowed['allowed'].iloc[j]=lista
        lista=[]


    clara_allowed = clara_allowed[['allowed','key','orari']]

    #-----------------------------------------------------------------------------------------------

    #altri_siti = df[['ID','Cliente','Sito','Indirizzo Sito','IstruzioniOperative','orari','Note','Servizio','Periodicita','SitoTerritoriale','Citta','Durata_stimata','Target_range','lat','lng','Descrizione Contratto']]
    altri_siti = planned.copy()
    #--------------------------------------------------------------------------------------------------------------------------FILTRO SOLO FERRARA 


    vincolo_nop = df_raw[['ID','Operatore','Operatore2']]
    altri_siti  = altri_siti.merge(vincolo_nop, how='left',left_on='ID',right_on='ID')


    altri_siti['Mensile1'] = [ctr[1:2] for ctr in altri_siti['Descrizione Contratto'].astype(str) ]
    altri_siti['Mensile'] = np.where(altri_siti['Mensile1'].astype(str)=='*','si','no')
    altri_siti = altri_siti.drop(columns=['Mensile1'])
    altri_siti = altri_siti.rename(columns={'Operatore':'N_op', 'Operatore2':'Op_vincolo','range':'Target_range'})
    altri_siti['TGT_new']=altri_siti['TGT_new'].astype(str)
    altri_siti['Durata_stimata']=altri_siti['Durata_stimata'].astype(float)
    altri_siti['Note'] = None
    altri_siti['orari']=None

    st.write(altri_siti)
    #altri_siti = altri_siti.merge(clara_allowed[['allowed','key']], how='left',left_on='key_sito',right_on='key')

    def convert_df(df):
        return df.to_csv(index=False,decimal=',').encode() 

    csv_altri = convert_df(altri_siti)
    st.download_button(
        label="Download interventi altri siti",
        data=csv_altri,
        file_name='Altri_siti.csv',
        mime='text/csv',
        )  



# altri siti è una copia di planned


#riga 380
#---------------
# planned = planned[['ID','S','Data Inizio','Cliente','Sito','Indirizzo Sito','Servizio','Periodicita','rank','Conteggio','Conteggio_distinti','ultimo_intervento','ultimo_pianificato',
#                    'ultimo_intervento_check','Target','Servizi sul sito','TGT_new','last','periodo','key_sito','Durata_stimata','range','Descrizione Contratto','SitoTerritoriale','lat','lng','IstruzioniOperative','Citta']]


# colonne = output.columns planned ha queste colonne

# r 298
# output = pianificabile[['ID','S','Data Inizio','Cliente','Sito','Indirizzo Sito','Servizio','key','key_sito','key_univoca','Durata_stimata','Descrizione Contratto','SitoTerritoriale','lat','lng','IstruzioniOperative','Citta']]
    
# r 249
# pianificabile = pianificabile[['ID','S','Data Inizio','Cliente','Sito','Indirizzo Sito','Dispositivi Installati','SitoTerritoriale','Servizio','ID Contratto','Codice Contratto','Citta',
#                                'Operatore','Min','Periodicita','Operatore2','key','key_sito','key_univoca','Durata_stimata','IstruzioniOperative', 'Note','lat','lng','Descrizione Contratto']]


# pianificabile parte da mese (tutte colonne)


