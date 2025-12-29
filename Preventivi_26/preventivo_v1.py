import streamlit as st
st.set_page_config(layout='wide')
import pandas as pd
#import random
#from utils import word
from io import BytesIO
#from docx import Document
#from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
#from docx.enum.style import WD_STYLE_TYPE
st.title('Configuratore preventivi')


# caricamento informazioni
path = st.file_uploader('Caricare file')
if not path:
    st.stop()

df_consumabili= pd.read_excel(path, sheet_name='Consumabili')
df_consumabili_for_dic = df_consumabili[['Materiale','UM','Costo_unitario']]

dic_unit_cons = dict(zip(df_consumabili_for_dic.Materiale, df_consumabili_for_dic.Costo_unitario))
dic_unit_cons['-']= 0

dic_um_cons = dict(zip(df_consumabili_for_dic.Materiale, df_consumabili_for_dic.UM))
dic_um_cons['-']= '-'

df_automezzi = pd.read_excel(path, sheet_name='Automezzi')
df_attrezzature = pd.read_excel(path, sheet_name='Attrezzature')

df_altro = pd.read_excel(path, sheet_name='Altro')

dic_mezzi_h = dict(zip(df_automezzi.Mezzo, df_automezzi['Costo orario']))
dic_mezzi_consumo = dict(zip(df_automezzi.Mezzo, df_automezzi['Consumo [km/l]']))

dic_altro = dict(zip(df_altro.Voce, df_altro.Importo))
mdo = dic_altro['Manodopera']

dic_attr = dict(zip(df_attrezzature.Risorsa, df_attrezzature.Costo_orario))
dic_attr['-']=0

df_ple = df_automezzi[['Ple' in check for check in df_automezzi.Mezzo]]
dic_ple = dict(zip(df_ple.Mezzo, df_ple['Nolo giornaliero']))

# CONFIGURAZIONE SEZIONI
# PRODOTTI E MATERIALI
# Prodotti chimici
# definisco un numero di voci da inserire in base al tipo di intervento
dic_voci = {
        'Disinfestazione insetti striscianti':5,
        'Zanzare':5,
        'Disinfestazione insetti volanti':5,
        'Demuscazione':5,
        'Disinfestazione - Saturazione':5,
        'Derattizzazione':5,
        'Piccioni':5
    }

# Materiali
dic_voci_mat = {
        'Disinfestazione insetti striscianti':3,
        'Zanzare':3,
        'Disinfestazione insetti volanti':3,
        'Demuscazione':3,
        'Disinfestazione - Saturazione':3,
        'Derattizzazione':3,
        'Piccioni':10
    }

# DPI
dic_voci_dpi = {
        'Disinfestazione insetti striscianti':3,
        'Zanzare':3,
        'Disinfestazione insetti volanti':3,
        'Demuscazione':3,
        'Disinfestazione - Saturazione':3,
        'Derattizzazione':3,
        'Piccioni':5
    }

dic_voci_a = {
        'Disinfestazione insetti striscianti':3,
        'Zanzare':4,
        'Disinfestazione insetti volanti':3,
        'Demuscazione':3,
        'Disinfestazione - Saturazione':3,
        'Derattizzazione':3,
        'Piccioni':3
    }



# selezione

tipologie = df_consumabili.Interventi.unique()
tipo = st.selectbox('Selezionare tipo intervento', options=tipologie)
intervento = st.text_input('Titolo preventivo')

t_materiali, t_viaggio, t_att, t_riep, t_word = st.tabs(['Materiali','Viaggio','Attrezzature','Totale', 'Esporta Word'])

with t_materiali:
    m10,m20,m30,m40,m50,m60 = st.columns([2,1,3,1,3,1])
        
    with m10: # chimici selezione
        voci = []
        st.subheader('Prodotti chimici')
        for i in range(dic_voci[tipo]):
            opzioni = ['-'] + list(df_consumabili[(df_consumabili.Categoria == 'Chimici') & (df_consumabili.Interventi == tipo)].Materiale.unique())
            globals()[f'qty{i}']=st.selectbox('mat',options=opzioni, label_visibility='hidden', key=f'{i}')
            voci.append(globals()[f'qty{i}'])
        #voci

    with m20: #chimici quantità
        qty = []
        st.subheader('')
        for i in range(dic_voci[tipo]):
            
            globals()[f'qty{i}']=st.number_input('qty',step=0.1, label_visibility='hidden', key=f'{i+2000}')
            qty.append(globals()[f'qty{i}'])
        #qty

        tabella_c = pd.DataFrame({'Categoria':'Chimici','Voce':voci, 'Quantità':qty})
        tabella_c['UM'] = [dic_um_cons[prodotto] for prodotto in tabella_c.Voce]
        tabella_c['Valore unitario'] = [dic_unit_cons[prodotto] for prodotto in tabella_c.Voce]
        tabella_c['Importo'] = tabella_c['Quantità']*tabella_c['Valore unitario']
        tabella_c = tabella_c[tabella_c.Voce != '-']


    with m30: # materiali selezione
        voci_m = []
        st.subheader('Materiali')
        for i in range(dic_voci_mat[tipo]):
            opzioni = ['-'] + list(df_consumabili[(df_consumabili.Categoria == 'Materiali') & (df_consumabili.Interventi == tipo)].Materiale.unique())
            globals()[f'voce_m{i}']=st.selectbox('mat',options=opzioni, label_visibility='hidden', key=f'{i+3000}')
            voci_m.append(globals()[f'voce_m{i}'])

    with m40:  # materiale quantità
        qty_m = []
        st.subheader('')
        for i in range(dic_voci_mat[tipo]):
            
            globals()[f'qty_m{i}']=st.number_input('qty',step=0.1, label_visibility='hidden', key=f'{i+4000}')
            qty_m.append(globals()[f'qty_m{i}'])

        tabella_m = pd.DataFrame({'Categoria':'Materiali','Voce':voci_m, 'Quantità':qty_m})
        tabella_m['UM'] = [dic_um_cons[prodotto] for prodotto in tabella_m.Voce]
        tabella_m['Valore unitario'] = [dic_unit_cons[prodotto] for prodotto in tabella_m.Voce]
        tabella_m['Importo'] = tabella_m['Quantità']*tabella_m['Valore unitario']
        tabella_m= tabella_m[tabella_m.Voce != '-']


    with m50: # selezione dpi
        voci_d = []
        st.subheader('DPI')
        for i in range(dic_voci_dpi[tipo]):
            opzioni = ['-'] + list(df_consumabili[(df_consumabili.Categoria == 'DPI') & (df_consumabili.Interventi == tipo)].Materiale.unique())
            globals()[f'voce_d{i}']=st.selectbox('dpi',options=opzioni, label_visibility='hidden', key=f'{i+5000}')
            voci_d.append(globals()[f'voce_d{i}'])


    with m60: # dpi quantità
        qty_d = []
        st.subheader('')
        for i in range(dic_voci_dpi[tipo]):
            
            globals()[f'qty_d{i}']=st.number_input('qty',step=1, label_visibility='hidden', key=f'{i+6000}')
            qty_d.append(globals()[f'qty_d{i}'])

        tabella_d = pd.DataFrame({'Categoria':'DPI','Voce':voci_d, 'Quantità':qty_d})
        tabella_d['UM'] = [dic_um_cons[prodotto] for prodotto in tabella_d.Voce]
        tabella_d['Valore unitario'] = [dic_unit_cons[prodotto] for prodotto in tabella_d.Voce]
        tabella_d['Importo'] = tabella_d['Quantità']*tabella_d['Valore unitario']
        tabella_d= tabella_d[tabella_d.Voce != '-']

    st.subheader('Riepilogo')
    st.divider()
    materials = pd.concat([tabella_c, tabella_m, tabella_d])
    materials
    
with t_viaggio:
    cat_v = ['Viaggio','Viaggio','Viaggio']
    voci_viaggio = []
    qty_v = []
    um_v = ['H','LT','H']
    unit_v = []
    imp_v = []

    st.subheader('Calcolo spese di viaggio e nolo mezzi')
    

    st.divider()

    v10,v15,v20 = st.columns([1,1,2])
    
    with v10:

        st.subheader('Inserimento dati')

        if st.checkbox('2 Operatori'):
            op=2

        elif st.checkbox('Più di 2 Operatori'):
            op = st.number_input('Inserire il numero di operatori', step=1)

        else:
            op=1

        durata_intervento = st.number_input('Durata intervento (escluso il viaggio)', step=5,key='d1')
        durata_viaggio = st.number_input('Inserire durata viaggio in minuti', step=5)
        mezzo = st.selectbox('Selezionare mezzo', options = df_automezzi[['Ple' not in check for check in df_automezzi.Mezzo]].Mezzo.unique()) #escludo le ple
        distanza = st.number_input('Inserire km A/R', step=5)
        gasolio = dic_unit_cons['Gasolio']
        litri = distanza / dic_mezzi_consumo[mezzo]

        durata_mdo = durata_intervento*op

        # ammortamento mezzo
        voci_viaggio.append(f'Ammortamento mezzo {mezzo}')
        qty_v.append((durata_viaggio+durata_intervento)/60)
        #um_v.append('H')
        unit_v.append(dic_mezzi_h[mezzo])
        imp_v.append(dic_mezzi_h[mezzo]*((durata_viaggio+durata_intervento)/60))

        # gasolio
        voci_viaggio.append(f'Gasolio (viaggio A/R:{distanza}km | consumo:{dic_mezzi_consumo[mezzo]} km/l | tot: {litri} litri)')
        qty_v.append(litri)
        unit_v.append(gasolio)
        imp_v.append(gasolio * litri)

        # manodopera viaggio
        voci_viaggio.append(f'Manodopera ({durata_viaggio} minuti di viaggio, {op} operatori | tot: {op*durata_viaggio} minuti totali)')
        qty_v.append(durata_viaggio*op/60)
        unit_v.append(mdo)
        imp_v.append(mdo*durata_viaggio*op/60)


        if st.toggle('Autostrada'):
            autostrada = st.number_input('Inserire importo pedaggio A/R', step=0.1, key='aut1')
            voci_viaggio.append(f'Pedaggio autostrada A/R')
            cat_v.append('Viaggio')
            qty_v.append(1)
            um_v.append('ND')
            unit_v.append(autostrada)
            imp_v.append(autostrada)

        if st.toggle('Nolo PLE'):
            mezzo_ple = st.selectbox('Selezionare modello PLE', options = dic_ple.keys())
            giorni_nolo = st.number_input('Giorni di nolo PLE', value=1, step=1)
            voci_viaggio.append(f'Nolo giornaliero {mezzo_ple}')
            cat_v.append('Nolo')
            qty_v.append(giorni_nolo)
            um_v.append('ND')
            unit_v.append(dic_ple[mezzo_ple])
            imp_v.append(dic_ple[mezzo_ple]*giorni_nolo)
            voci_viaggio.append('Quota fissa gasolio PLE (50€)')
            cat_v.append('Nolo')
            qty_v.append(giorni_nolo)
            um_v.append('ND')
            unit_v.append(50)
            imp_v.append(50*giorni_nolo)


        if st.toggle('Buono pasto'):
            giorni_intervento = st.number_input('Selezionare giorni di intervento', value=1, step=1)
            pasto = dic_altro['Buono pasto']
            voci_viaggio.append(f'Buono pasto per {op} operatori per {giorni_intervento} giorni')
            cat_v.append('Pasti')
            qty_v.append(op*giorni_intervento)
            um_v.append('ND')
            unit_v.append(pasto)
            imp_v.append(pasto*op*giorni_intervento)

        #cat_v
        #voci_viaggio
        #qty_v
        #um_v
        #unit_v
        #imp_v

        if st.toggle('Interventi diversi dal primo'):

            st.subheader('Interventi diversi dal primo')

            #cat_v = ['Viaggio','Viaggio','Viaggio','Viaggio','Viaggio','Viaggio']
            cat_v = cat_v + ['Viaggio'] *3
            um_v = um_v + ['H','LT','H']
         




            if st.checkbox('2 Operatori', key='multi'):
                op=2

            elif st.checkbox('Più di 2 Operatori', key='multi1'):
                op = st.number_input('Inserire il numero di operatori', step=1, key='multi2')

            else:
                op=1

            durata_intervento2 = st.number_input('Durata intervento (escluso il viaggio)', step=5,key='multi3')

            numero_interventi = st.number_input('Numero di interventi totali (escluso il primo)', step=1, value=2, key='multi11')
            
            durata_altri = (durata_intervento2) * (numero_interventi)
            
            durata_mdo += (durata_altri*op)

            mezzo = st.selectbox('Selezionare mezzo', options = df_automezzi[['Ple' not in check for check in df_automezzi.Mezzo]].Mezzo.unique(), key='multi4') #escludo le ple
            
            gasolio = dic_unit_cons['Gasolio']

            litri = distanza * numero_interventi / dic_mezzi_consumo[mezzo]

            # ammortamento mezzo
            voci_viaggio.append(f'Ammortamento mezzo {mezzo} per interventi diversi dal primo')
            qty_v.append((durata_viaggio+durata_intervento2) * numero_interventi/60)
            #um_v.append('H')
            unit_v.append(dic_mezzi_h[mezzo])
            imp_v.append(dic_mezzi_h[mezzo]*((durata_viaggio+durata_intervento2) * numero_interventi/60))

            # gasolio
            voci_viaggio.append(f'Gasolio (viaggio A/R:{distanza}km | {numero_interventi} interventi | consumo:{dic_mezzi_consumo[mezzo]} km/l | tot: {litri} litri)')
            qty_v.append(litri)
            unit_v.append(gasolio)
            imp_v.append(gasolio * litri)

            # manodopera viaggio
            voci_viaggio.append(f'Manodopera ({durata_viaggio} minuti di viaggio, {op} operatori | {numero_interventi} interventi | tot: {op*durata_viaggio*numero_interventi} minuti totali)')
            qty_v.append(durata_viaggio*op*numero_interventi/60)
            unit_v.append(mdo)
            imp_v.append(mdo*durata_viaggio*op*numero_interventi/60)


            if st.toggle('Autostrada', key='multi5'):
                autostrada = st.number_input('Inserire importo pedaggio A/R', step=0.1, key='aut5')
                voci_viaggio.append(f'Pedaggio autostrada A/R')
                cat_v.append('Viaggio')
                qty_v.append(numero_interventi)
                um_v.append('ND')
                unit_v.append(autostrada)
                imp_v.append(autostrada*numero_interventi)

            if st.toggle('Nolo PLE', key='multi6'):
                mezzo_ple = st.selectbox('Selezionare modello PLE', options = dic_ple.keys(), key='multi7')
                giorni_nolo = st.number_input('Giorni di nolo PLE', value=1, step=1, key='multi8')
                voci_viaggio.append(f'Nolo giornaliero {mezzo_ple}')
                cat_v.append('Nolo')
                qty_v.append(giorni_nolo)
                um_v.append('ND')
                unit_v.append(dic_ple[mezzo_ple])
                imp_v.append(dic_ple[mezzo_ple]*giorni_nolo)
                voci_viaggio.append('Quota fissa gasolio PLE (50€)')
                cat_v.append('Nolo')
                qty_v.append(giorni_nolo)
                um_v.append('ND')
                unit_v.append(50)
                imp_v.append(50*giorni_nolo)


            if st.toggle('Buono pasto', key='multi9'):
                giorni_intervento = st.number_input('Selezionare giorni di intervento', value=1, step=1, key='multi10')
                pasto = dic_altro['Buono pasto']
                voci_viaggio.append(f'Buono pasto per {op} operatori per {giorni_intervento} giorni per {numero_interventi} interventi')
                cat_v.append('Pasti')
                qty_v.append(op*numero_interventi*giorni_intervento)
                um_v.append('ND')
                unit_v.append(pasto)
                imp_v.append(pasto*op*numero_interventi *giorni_intervento)

           
        tabella_viaggio = pd.DataFrame({'Categoria':cat_v,'Voce':voci_viaggio,'Quantità':qty_v, 'UM':um_v, 'Valore unitario':unit_v, 'Importo':imp_v})
        
    with v20:
        st.subheader('Tabella di riepilogo')
        st.divider()
        tabella_viaggio
        st.metric('Importo totale viaggio',tabella_viaggio.Importo.sum(), border=True)

    pass

with t_att:
    a10, a20, a30 = st.columns([2,1,1])
    with a10: # attrezz
        voci_a = []
        st.subheader('Attrezzature e manodopera intervento')
        for i in range(dic_voci_a[tipo]):
            opzioni = ['-'] + list(df_attrezzature[(df_attrezzature.Interventi == tipo)].Risorsa.unique())
            globals()[f'voce_a{i}']=st.selectbox('att',options=opzioni, label_visibility='hidden', key=f'{i+8000}')
            voci_a.append(globals()[f'voce_a{i}'])

        tabella_a = pd.DataFrame({'Categoria':'Attrezzature','Voce':voci_a, 'Quantità': durata_intervento/60})
        tabella_a['UM'] = 'H'
        tabella_a['Valore unitario'] = [dic_attr[attr] for attr in tabella_a.Voce]
        tabella_a['Importo'] = tabella_a['Quantità']*tabella_a['Valore unitario']
        tabella_a= tabella_a[tabella_a.Voce != '-']
        tabella_a.loc[len(tabella_a)]=['Manodopera',
                                       f'Manodopera totale ({durata_mdo/60} ore uomo totali )',
                                       durata_mdo/60,
                                       'H',
                                       mdo,
                                       mdo*durata_mdo/60]
        
        
        st.subheader('Importo a corpo')

        testo_corpo = st.text_input('Descrizione voce a corpo', value='Altri importi a corpo')
        importo_corpo = st.number_input('Importo a corpo', step=1)
        
        if importo_corpo != 0:
            tabella_a.loc[len(tabella_a)]=['Altro',
                                        f'{testo_corpo}',
                                        None,
                                        None,
                                        None,
                                        importo_corpo]
        
        
        tabella_a

    with a30:
        st.subheader('Riepilogo attrezzature')
        # devo aggiungere una riga per la manodopera
      
        st.metric('Importo totale attrezzature',tabella_a.Importo.sum(), border=True)

with t_riep:

    st.subheader(intervento)

    tabella_totale = pd.concat([materials, tabella_viaggio, tabella_a])

    
    costo_totale = tabella_totale.Importo.sum()
    tabella_totale.loc[len(tabella_totale)] = [None,'Costo diretto',None,None,None,costo_totale]
    #tabella_totale

   

    r10, r20 = st.columns([1,2])
    with r10:
        st.subheader('Maggiorazioni, spese generali e profitto')
        st.metric('Costo diretto', value='€ {:0.2f}'.format(costo_totale), border=True)
        st.subheader(':orange[Calcolo maggiorazioni]', divider='orange')
        mag_not=0
        mag_urg=0
        mag_fes=0
        mag_not=0
        if st.checkbox('Maggiorazione Notturna'):
            mag_not = dic_altro['Maggiorazione notturna']
            testo_magnot = f'Maggiorazione notturna {mag_not*100}%'
            importo_magnot = '€ {:0.2f}'.format(mag_not*costo_totale)
            st.metric(testo_magnot, value=importo_magnot)
            tabella_totale.loc[len(tabella_totale)] = [None,testo_magnot,None,None,None,mag_not*costo_totale]

        if st.checkbox('Maggiorazione Festivo'):
            mag_fes = dic_altro['Maggiorazione festivo']
            testo_magfes = f'Maggiorazione festivo {mag_fes*100}%'
            importo_magfes = '€ {:0.2f}'.format(mag_fes*costo_totale)
            st.metric(testo_magfes, value=importo_magfes)
            tabella_totale.loc[len(tabella_totale)] = [None,testo_magfes,None,None,None,mag_fes*costo_totale]

        if st.checkbox('Maggiorazione Urgenza'):
            mag_urg = dic_altro['Urgenza']*(durata_intervento+durata_viaggio)/60
            testo_magurg= f'Maggiorazione urgenza'
            importo_magurg = '€ {:0.2f}'.format(mag_urg)
            st.metric(testo_magurg, value=importo_magurg)
            tabella_totale.loc[len(tabella_totale)] = [None,testo_magurg,None,None,None,mag_urg]
        
        costo_diretto = costo_totale+mag_not*costo_totale+mag_fes*costo_totale+mag_urg
        tabella_totale.loc[len(tabella_totale)] = [None,'Costo incluse maggiorazioni',None,None,None,costo_diretto]
        st.metric('Costo diretto totale', value=('€ {:0.2f}'.format(costo_diretto)), border=True)
        st.subheader(':orange[Calcolo costi generali + Profitto]', divider='orange')
        oh = st.number_input('Inserire % markup costi generali', step=0.05, value=0.2)
        st.metric('Costi generali ({:0.2f}%)'.format(oh*100), value='€ {:0.2f}'.format(costo_diretto*oh),border=True)

        tabella_totale.loc[len(tabella_totale)] = [None,'Costi generali ({:0.2f}%)'.format(oh*100),None,None,None,costo_diretto*oh]

        costo_totale = costo_diretto*(1+oh)
        tabella_totale.loc[len(tabella_totale)] = [None,'Costo totale',None,None,None,costo_totale]
        profit = st.number_input('Inserire % markup profitto', step=0.05, value=0.2)
        st.metric('Margine aziendale {:0.2f}%'.format(profit*100), value='€ {:0.2f}'.format(costo_totale*(profit)), border=True)
        tabella_totale.loc[len(tabella_totale)] = [None,'Margine aziendale {:0.2f}%'.format(profit*100),None,None,None,profit*costo_totale]

        prezzo = costo_totale * (1+profit)
        st.subheader(':orange[Prezzo finale]', divider='orange')
        st.metric('Prezzo finale', value='€ {:0.2f}'.format(prezzo))
        tabella_totale.loc[len(tabella_totale)] = [None,'Prezzo finale',None,None,None,prezzo]

    with r20:
        st.subheader('Tabella riepilogativa')
        st.dataframe(tabella_totale, height=800)

        def scarica_excel(df, filename):
            output = BytesIO()
            writer = pd.ExcelWriter(output, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1',index=False)
            writer.close()

            st.download_button(
                label="Scarica Excel",
                icon=':material/download:',
                data=output.getvalue(),
                file_name=filename,
                mime="application/vnd.ms-excel"
            )
        
        scarica_excel(tabella_totale, f'Preventivo - {intervento}.xlsx')

    with t_word:

        st.subheader('Esporta preventivo in Word')
        
        # Input aggiuntivi per i nuovi placeholder nel modello
        c1, c2 = st.columns(2)
        with c1:
            intestatario = st.text_input('Intestatario', value='Sig. Nicola Ferrari')
            indirizzo = st.text_input('Indirizzo', value='Via Roma, 1')
            cap = st.text_input('CAP', value='12345')
            città = st.text_input('Città', value='Roma')
            mail = st.text_input('Email', value='cliente@email.it')
        with c2:
            cellulare = st.text_input('Cellulare', value='347/0000000')
            descrizione = st.text_area('Descrizione intervento', value='Offerta per gestire l’attuale criticità.\nSono previste n. 3 fasi operative.')
            note_tecniche = st.text_area('Note tecniche', value='Mirato al controllo dei topi e ratti;\nEseguito sulle aree di proprietà.')

        # Template path
        #template_path = 'modellonew.docx' 
        template_path = 'https://github.com/alebelluco/Test_EX/blob/main/Preventivi_26/modellonew.docx'
        
        if st.button('Genera Documento Word'):
            try:
                # Carica il documento
                from docx import Document
                import datetime
                
                doc = Document(template_path)
                
                # Sostituzioni semplici nei paragrafi
                today = datetime.date.today().strftime('%d/%m/%Y')
                
                # Mappa dei placeholder
                placeholders = {
                    '[DATA ODIERNA]': today,
                    '[INTESTATARIO]': intestatario,
                    '[INDIRIZZO]': indirizzo,
                    '[CAP]': cap,
                    '[CITY]': città,
                    '[MAIL]': mail,
                    '[CELLULARE]': cellulare,
                    '[TITOLO PREVENTIVO]': intervento if intervento else "Offerta tecnico economica",
                    '[DESCRIZIONE INTERVENTO]': descrizione,
                    '[IMPORTO INTERVENTO]': '€ {:0.2f} + Iva di Legge'.format(prezzo),
                    '[NOTE TECNICHE]': note_tecniche
                }
                
                for para in doc.paragraphs:
                    for key, value in placeholders.items():
                        if key in para.text:
                            # Gestione sostituzione nei runs per mantenere il font
                            for run in para.runs:
                                if key in run.text:
                                    # Se il valore ha dei ritorni a capo, dobbiamo gestirli diversamente
                                    if "\n" in str(value):
                                        parts = str(value).split("\n")
                                        run.text = run.text.replace(key, parts[0])
                                        for part in parts[1:]:
                                            run.add_break()
                                            run.add_text(part)
                                    else:
                                        run.text = run.text.replace(key, str(value))
                    
                    # Pulizia vecchi placeholder
                    for run in para.runs:
                        for old_key in ['[[TITOLO]]', '[[PREZZO]]', '[[TABELLA_1]]', '[[TABELLA_2]]']:
                            if old_key in run.text:
                                run.text = run.text.replace(old_key, '')

                # Salvataggio in buffer
                target_stream = BytesIO()
                doc.save(target_stream)
                
                st.success('Documento Word generato con successo!')
                
                st.download_button(
                    label="Scarica Word",
                    data=target_stream.getvalue(),
                    file_name=f"Preventivo_{intervento if intervento else 'nuovo'}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Errore durante la generazione del Word: {e}")
                st.info("Assicurati che 'modellonew.docx' sia presente nella cartella dello script.")

