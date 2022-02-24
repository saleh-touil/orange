#Les Modules utilisées :

import streamlit as st  # pip3 install streamlit
import pandas as pd  # pip3 install pandas
import base64  # Standard Python Module
import plotly.express as px  # pip3 install plotly-express
from io import StringIO, BytesIO  # Standard Python Module
import matplotlib.pyplot as plt #pip3 install matplotlib
import emojis #pip3 install emojis

#Les fonctions :

def generate_excel_download_link(df):
    # Credit Excel: https://discuss.streamlit.io/t/how-to-add-a-download-excel-csv-function-to-a-button/4474/5
    towrite = BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)  # write to BytesIO buffer
    towrite.seek(0)  # reset pointer
    b64 = base64.b64encode(towrite.read()).decode()
    href = f"<a href='data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}' download='data_download.xlsx'>Télécharger le fichier Excel </a>"
    return st.markdown(href, unsafe_allow_html=True)

def emojiss(ch):
    if ((ch == "nan")or(ch == "NaT") or (ch == "")):
        ch = emojis.decode(":no_entry_sign:")
    return ch    

#Menu Principal:

st.set_page_config(
    page_title='Rapport',layout="centered",page_icon=":bar_chart:",
    menu_items={'Get Help': None,'Report a bug': None,'About': None},initial_sidebar_state="auto")
title = st.title('Excel Report 📈')
st.subheader('À vaillant coeur rien d’impossible. -Jacques Cœur')
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            </style>
            """            
st.markdown(hide_streamlit_style, unsafe_allow_html=True) 
#Upload fichier
uploaded_file = st.file_uploader('Choisir Un Fichier XLSX ', type='xlsx')
if uploaded_file: #Si le fichier est mis :
    st.markdown('---')
    df = pd.read_excel(uploaded_file, engine='openpyxl',skiprows=4, usecols=lambda x: 'Unnamed' not in x)
    x=df.astype(str)
    #Tableau du basé sur le fichier Excel donné
    st.dataframe(x)
    #Liste à choix multiple      
    groupby_column = st.selectbox(
        'Quel type de Rapport?',
        ('Rapport Complet', "Rapport d'un Client"),
    )
    if groupby_column=="Rapport d'un Client":    
     if "Demande client" in x.columns: 
      st.title('Mise en Service :')
      text_input=st.text_input("Filtrer par Demande Client: ")
      text_input = str(text_input)
      k=0
      if text_input in x.values:
        st.markdown('---')
        c = st.container() 
        #Affichage De chaque case concernant le code donné 
        for j in range(0,len(df.index)):
             v = x.iloc[j]['Demande client']
             if v == text_input:
              for z in x.columns: 
               varr = emojiss(x.iloc[j][f"{z}"])
               c.write(f'{z} : { varr }')  
               k=1 
      #Message d'erreur Demande client         
      elif(k==0)and(text_input):
       st.warning("Demande Client N'existe pas")
     if "T0 INI" in x.columns:
         st.title('Changement :')
         text_input=st.text_input("Filtrer par Demade Changement Client : ")
         text_input = str(text_input)
         k=0
         if text_input in x.values:
          st.markdown('---')
          c = st.container()
          #Affichage De chaque case concernant le demande changement donné 
          for j in range(0,len(df.index)):
             v = x.iloc[j]['Dde chgt client']
             if v == text_input:
              for z in x.columns: 
               varr = emojiss(x.iloc[j][f"{z}"])
               c.write(f'{z} : { varr }')  
               k=1 
         #Message d'erreur Changement client      
         elif(k==0)and(text_input):
           st.warning("Demade Changement Client N'existe pas")
    elif (groupby_column=="Rapport Complet")and("Demande client" in x.columns):
#Liste à choix multiple              
        options = st.selectbox(
         "Type d'accès Demande 1 :",
         ['','Tout','FO', 'FH', 'LTE TDD', 'VSAT1.2'],
         )
        e=["Date installation FO","Date installation & MES FH","Date installation&MES LTE TDD réelle","Date étape installation & MES VSAT 1.2"]
        en=["FO","FH","LTE TDD","VSAT1.2"]
#Type d'access 1        
        ops = {
            '':(),
            'Tout':e,
            'FO':'Date installation FO',
            'FH':'Date installation & MES FH',
            'LTE TDD':"Date installation&MES LTE TDD réelle",
            'VSAT1.2':"Date étape installation & MES VSAT 1.2"
        }
#Type d'access 2       
        ops2 = {
            '':(),
            'Tout':e,
            'FO':'Date installation FO',
            'FH':'Date installation & MES FH',
            'LTE TDD':"Date installation&MES LTE TDD réelle",
            'VSAT1.2':"Date étape installation & MES VSAT 1.2"
        }
#case valeur : resultat        
        tutti = {
            1:'FO',
            2:'FH',
            3:"LTE TDD",
            4:"VSAT1.2"
        }
#Liste à choix multiple        
        options2 = st.selectbox(
         "Type d'accès Demande 2 :",
         ['','Tout','FO', 'FH', 'LTE TDD', 'VSAT1.2'],
         )
#Boutton Radio        
        choice = st.radio (
            "Filtrer par : ",
            ('Demande Complete','Demande en Cours')
        )
#Initiation :
        k=0;m=0;l=0;z=0

        if (choice == "Demande Complete")and(options!='')or(options2!=''):
             st.markdown("---")
             ch = emojis.decode(":ok:")
             st.title(f"Demande Complete {ch}")
             c=st.container()
             for j in range(0,len(df.index)):
               v = x.iloc[j]["Dde 1 : Type d'accès"]
               w = x.iloc[j]["Dde 2 : Type d'accès"]
               if options!='': 
                cat = x.iloc[j][ops[f"{options}"]]
                on = ops[f"{options}"]
               if (options2!='')and(options2!='Tout'): 
                cat2 = x.iloc[j][ops2[f"{options2}"]]
                on2 = ops2[f"{options2}"]
                #Addition des résultats
               if v==options:
                 k=k+1
               if(v==options)and(cat!="NaT"):
                   z=z+1  
               if w==options2:
                   l=l+1  
               if (w==options2)and(cat2!="NaT"):
                   m=m+1
             if (options!='')and(options!='Tout'): 
              st.subheader(f"Demande 1 : {options}")  
              st.write(f"Total de {k} de type d'accès {options}.")
              st.write(f"{z} Demandes Completes par rapport au {on}.")
              labels = "Demande complete","Demande en cours"
              h = k-z
              sizes = [(z*100)/k,((h)*100)/k]
              fig1, ax1 = plt.subplots()
              ax1.pie(sizes, labels=labels, autopct='%1.1f%%',
                   shadow=True, startangle=90)
              ax1.axis('equal')  
              st.pyplot(fig1)                 
             elif(options=='Tout'):
                 st.subheader(f"Demande 1 : {options2}") 
                 total = 0
                 for i in range(1,5):
                     s=0
                     options=tutti[i]
                     for j in range(0,len(df.index)): 
                      v = x.iloc[j]["Dde 1 : Type d'accès"]
                      cat = x.iloc[j][ops[f"{options}"]]
                      #Addition des résultats
                      if (v==options):
                          s=s+1                          
                      if i == 1:
                        fo = s 
                      if i == 2:
                         fh = s 
                      if i == 3:
                         lte_tdd = s 
                      if i == 4:
                        vsat = s 
                     total=total+s
                     on = ops[f"{options}"]
                     st.write(f"Total de {s} de type d'accès {options}.") 
                 labels = "FO","FH","LTE TDD","VSAT1.2"
                 sizes = [(fo*100)/total,(fh*100)/total,(lte_tdd*100)/total,(vsat*100)/total]
                 fig1, ax1 = plt.subplots()
                 ax1.pie(sizes, labels=labels, autopct='%1.1f%%',
                   shadow=True, startangle=90)
                 ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

                 st.pyplot(fig1)  
             if (options2!="")and(options2!='Tout'):
                 st.subheader(f"Demande 2 : {options2}")   
                 st.write(f"Total de {l} de type d'accès {options2}.") 
                 st.write(f"{m} Demandes Completes par rapport au {on2}.")
                 labels = "Demande complete","Demande en cours"
                 r = l-m
                 sizes = [(m*100)/l,((r)*100)/l]
                 fig1, ax1 = plt.subplots()
                 ax1.pie(sizes, labels=labels, autopct='%1.1f%%',
                   shadow=True, startangle=90)
                 ax1.axis('equal')  
                 st.pyplot(fig1)            
             elif(options2=='Tout'):
                 st.subheader(f"Demande 2 : {options2}") 
                 total = 0
                 for i in range(1,5):
                     s=0
                     options=tutti[i]
                     for j in range(0,len(df.index)): 
                      v = x.iloc[j]["Dde 2 : Type d'accès"]
                      cat = x.iloc[j][ops[f"{options2}"]]
                      #Addition des résultats
                      if (v==options):
                          s=s+1                          
                      if i == 1:
                        fo = s 
                      if i == 2:
                         fh = s 
                      if i == 3:
                         lte_tdd = s 
                      if i == 4:
                        vsat = s 
                     total=total+s
                     on = ops[f"{options}"]
                     st.write(f"Total de {s} de type d'accès {options}.") 
                 labels = "FO","FH","LTE TDD","VSAT1.2"
                 sizes = [(fo*100)/total,(fh*100)/total,(lte_tdd*100)/total,(vsat*100)/total]
                 fig1, ax1 = plt.subplots()
                 ax1.pie(sizes, labels=labels, autopct='%1.1f%%',
                   shadow=True, startangle=90)
                 ax1.axis('equal')  
                 st.pyplot(fig1)
        elif (choice == "Demande en Cours"):
             st.markdown("---")
             ch = emojis.decode(":hourglass:")
             st.title(f"Demande en cours {ch}")
             
    if ("T0 INI" in x.columns) and(groupby_column=="Rapport Complet"):           
     options = st.selectbox(
         "Type de changement :",
         ["","Tout","Upgrade de débit", "Downgrade de débit", "Changement nb appels simultanés OU SDA", "Changement pack", "Portabilité VoIP fixe"],
         )
     tt=["Upgrade de débit","Downgrade de débit","Changement nb appels simultanés OU SDA","Changement pack","Portabilité VoIP fixe"]
     tt2=["Date valid changement","Date valid changement", "Date valid changement", "Date validation changement de pack", " Date validation MES"]
     ops = {
            "Upgrade de débit":"Date valid changement",
            "Downgrade de débit":"Date valid changement",
            "Changement nb appels simultanés OU SDA" : "Date valid changement",
            "Changement pack" : "Date validation changement de pack",
            "Portabilité VoIP fixe" : "Date validation MES"
        }


     z=0
     k=0
     if(options!="Tout")and(options!=""):   
      st.markdown("---")
      ch = emojis.decode(":ok:")
      st.title(f"{options} {ch}")
      c=st.container()
      for j in range(0,len(df.index)):
               v = x.iloc[j]["Type de changement"]
               cat = x.iloc[j][ops[f"{options}"]]
               on = ops[f"{options}"]
               #Addition des résultats
               if v==options:
                   k=k+1
               if(v==options)and(cat!="NaT"):
                   z=z+1  
     if (options!='')and(options!='Tout'):
              h=k-z 
              st.subheader(f"Type de changement : {options}")  
              st.write(f"Total de {k} de type d'accès {options}.")
              st.write(f"{z} Demandes Completes par rapport au {on}.")
              #st.write(f"{k-z} Demandes Non Completes") 
              labels = "Demande complete","Demande en cours"
              h = k-z
              sizes = [(z*100)/k,((h)*100)/k]
              fig1, ax1 = plt.subplots()
              ax1.pie(sizes, labels=labels, autopct='%1.1f%%',
                   shadow=True, startangle=90)
              ax1.axis('equal')  
              st.pyplot(fig1)                 
     if(options=='Tout'):
                 st.markdown('___')
                 ch = emojis.decode(":ok:")
                 st.title(f"{options} {ch}")
                 total=0;up=0;down=0;nb_appel=0;chan_pack=0;voip=0;
                 for i in range(0,5):
                     s=0
                     options=tt[i]
                     for j in range(0,len(df.index)): 
                      v = x.iloc[j]["Type de changement"]
                      #Addition des résultats
                      if (v==options):
                          s=s+1                          
                          if i == 0:
                            up = s 
                          if i == 1:
                            down = s 
                          if i == 2:
                            nb_appel = s 
                          if i == 3:
                            chan_pack = s 
                          if i == 4:
                            voip = s
                              
                     total=total+s
                 st.subheader(f"Total de {total} type de changement .") 
                 
                 labels = f"Upgrade [{up}]",f"Downgrade [{down}]",f"Changement nb d'appels [{nb_appel}]",f"Changement pack [{chan_pack}]",f"Portabilité VoIP fixe [{voip}]"
                 sizes = [(up*100)/total,(down*100)/total,(nb_appel*100)/total,(chan_pack*100)/total,(voip*100)/total]
                 fig1, ax1 = plt.subplots()
                 ax1.pie(sizes, labels=labels, autopct='%1.1f%%',
                     shadow=True, startangle=90)
                 ax1.axis('equal')  
                 st.pyplot(fig1)  
#Téléchargements :                             
    st.subheader('Téléchargements:')
    generate_excel_download_link(x)                  