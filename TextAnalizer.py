from logging import NullHandler
import pandas as pd
import spacy as sp
import os
import PySimpleGUI as sg
import webview as wv
from spacy.matcher import DependencyMatcher as match
from spacy.lang.la.stop_words import STOP_WORDS
from spacy import displacy
from string import punctuation

stopwords = list(STOP_WORDS)


def list_proper_nouns(text):
    try:
        nombres_list=[]
        for token in text:
            if (token.text in stopwords or token.text in punctuation):
                continue
            if token.pos_ in ['PROPN']:
                nombre_propio={"token":token,"nombre":token.text}
                nombres_list.append(nombre_propio)       
        return nombres_list
    except Exception as e:
            sg.popup(f"Ocurrió el siguiente error mientras se detectaban los nombres propios: {e}, seguramente tengas que descargar el motor de idioma de Spacy en las opciones.")
def list_all_lemmas(text):
    try:
        lemmas_list=[]
        for token in text:
            if (token.text in stopwords or token.text in punctuation):
                continue
            if token.pos_ in ['NOUN', 'VERB']:
                lemma = {"token":token,"lemma":token.lemma_.lower()}
                lemmas_list.append(lemma)
        return lemmas_list
    except Exception as e:
            sg.popup(f"Ocurrió el siguiente error mientras se detectaban los nombres propios: {e}, seguramente tengas que descargar el motor de idioma de Spacy en las opciones.")
def list_morphology(text):
    try:
        morph_list=[]
        for token in text:
            if (token.text in stopwords or token.text in punctuation):
                continue
            morfologia_dict = {"token":token}
            morfologia_dict.update(token.morph.to_dict())
            morph_list.append(morfologia_dict)
        return morph_list
    except Exception as e:
            sg.popup(f"Ocurrió el siguiente error mientras se detectaban los nombres propios: {e}, seguramente tengas que descargar el motor de idioma de Spacy en las opciones.")
#Despues crear una lista con un diccionario por cada token para poder consultarlos.


def detect_all_nouns(text):
    sustantivos = []

    return sustantivos      



sg.theme('Reddit')  
# El diseño de la ventana.
#primero definimos el menu superior:
menu_def = ['Opciones', ['Descargar Modelo Spacy de latin grande','Descargar Modelo Transformer para Latin Spacy','Cerrar']],['&Ayuda',['&Ayuda', 'About...']],

#luego creamos el layout
layout = [  [sg.Menu(menu_def)],
            [sg.Text('Bienvenido a la interfaz para Latincy')],
            [sg.Text('Nombre del archivo de salida'), sg.InputText(key="-OUTFILE-"),sg.Radio('.xlsx',"FILETYPE", key="-XLSX-", default=True), sg.Radio('.csv',"FILETYPE", key="-CSV-", default=True), sg.Radio('.json',"FILETYPE", key="-JSON-", default=True)],
            [sg.Text('¿Dónde querés guardar el archivo?'), sg.InputText(key="-OUTFOLDER-"), sg.FolderBrowse(target="-OUTFOLDER-")],
            [sg.Text('Elegí archivo a analizar'),sg.Input(), sg.FileBrowse(key="-IN-")],
            [sg.Text('¿Cuál es el nombre de la pestaña?'), sg.InputText(key="-SHEET-")],
            [sg.Text('¿Cuál es el nombre de la columna?'), sg.InputText(key="-COL-")],
            [sg.Text('Orden de palabras en el archivo generado'),sg.Radio('Mantener palabras en filas originales',"WORDORDER", key="-KEEPROW-", default=True), sg.Radio('Una palabra por fila',"WORDORDER", key="-ONEWORDROW-", default=False), sg.Radio('Todas las palabras en la misma columna, una fila por palabra',"WORDORDER", key="-KEEPROWONECOL-", default=False), sg.Radio('Mantener filas, todo en una columna',"WORDORDER", key="-ONECOL-", default=False)],
            [sg.Text('Orden de columnas en el archivo generado'),sg.Radio('Mantener palabras en filas originales, una columna por opción',"WORDORDER", key="-KEEPROW-", default=True), sg.Radio('Una palabra por fila y una columna para opción',"WORDORDER", key="-ONEWORDROW-", default=False), sg.Radio('Todas las palabras en la misma columna, una fila por palabra',"WORDORDER", key="-KEEPROWONECOL-", default=False), sg.Radio('Mantener filas, todo en una columna',"WORDORDER", key="-ONECOL-", default=False)],
            [sg.Text('Elegir un motor de NLP'),sg.Radio('Spacy',"ENGINE", key="-SPAC-", default=True)],
            [sg.Text('Tipo de Análisis'), sg.Checkbox('Entidades Naturales', key="-NER-", default=False),sg.Checkbox('Morfología', key="-MORPH-", default=False), sg.Checkbox('Separar el label en las entidades naturales', key="-NERL-", default=False)],
            [sg.Text('PoS Tags'), sg.Checkbox('Nombres Propios', key="-PROPN-", default=True), sg.Checkbox('Sustantivos y Verbos', key="-NOUNV-", default=True),sg.Checkbox('Morfología', key="-MORPH-", default=False), sg.Checkbox('Separar el label en las entidades naturales', key="-NERL-", default=False)],
            [sg.Text('Visualización'), sg.Checkbox('Visualizar entidades de cada fila.', key="-DISPLAY-", default=False), sg.Checkbox('Mostrar la morfología de cada fila.', key="-DISPLAYMORPH-", default=False)],            
            [sg.Button('OK'), sg.Button('Cerrar')]]

#Y la ventana:
window = sg.Window('GN Analizer', layout)

# Hacemos un loop para procesar eventos y tomar los inputs de la ventana como values
while True:
    event, values = window.read()
    #primero esperamos eventos del dropdown menu:
    if event == 'Descargar Modelo Spacy de latin grande': 
        os.system('pip install https://huggingface.co/latincy/la_core_web_lg/resolve/main/la_core_web_lg-any-py3-none-any.whl')
        #os.system('python -m spacy download la_core_web_lg')
    if event == 'Descargar Modelo Transformer para Latin Spacy':
        os.system('pip install https://huggingface.co/latincy/la_core_web_trf/resolve/main/la_core_web_trf-any-py3-none-any.whl')
        #os.system('python -m spacy download la_core_web_trf')
    if event == 'About...':     
        sg.popup('Este programa fue creado para probar spacy en latin', 'Version 0.1', 'PySimpleGUI rocks... Spacy is da bomb...')   
    #despues, si el usuario clickea en Ok procesamos todo lo que sigue
    if event == 'OK':
        #Lo primero que hacemos despues del OK es chequear que estén todas las opciones completadas y sino prompteo un error
        if values["-OUTFILE-"] == "":
            sg.popup(f"No determinaste el nombre del archivo que vamos a generar.")
        elif values["-OUTFOLDER-"] == "":
            sg.popup(f"No determinaste la carpeta en la que se va a guardar el archivo generado.")    
        elif values["-IN-"] == "" :
            sg.popup(f"No seleccionaste un archivo para analizar.")
        elif values["-SHEET-"] == "" :
            sg.popup(f"No determinaste en qué hoja del archivo está la columna de texto.")
        elif values["-COL-"] == "":
            sg.popup(f"No determinaste cuál es la columna de texto.")
        elif values["-NOUNV-"] == False and values["-PROPN-"] == False and values["-NER-"] == False and values["-DISPLAY-"] == False and values["-MORPH-"] == False and values["-SENTIMENT-"] == False: 
            sg.popup(f"No clickeaste ninguna opción para analizar el texto")
        #Si se llenaron todas las opciones, ninguna está vacía, procedemos a ejecutar el programa.
        else:
            original_file = values["-IN-"]
            sheet_name = values["-SHEET-"]
            column = values["-COL-"]
            #A parte de tomar los valores inputeados, usamos os para normalizar la ruta de la carpeta (porque va a cambiar segun el OS) y para joinearla con el nombre del archivo
            new_file = os.path.join(os.path.normcase(values["-OUTFOLDER-"]), values["-OUTFILE-"])
            #Ahora, segun que opcion clickeo el usuario, vamos a definir el motor y modelo con las variables nlp y nlpner
            # Defino las variables:
            if  values["-SPAC-"] == True:
                try:
                    nlpner = sp.load('la_core_web_lg')
                    nlp = sp.load('la_core_web_lg')

                #Con spacy usamos dos modelos diferentes, uno sirve para NER y el otro deptrees y lematizacion
                except Exception as s:
                    sg.popup(f"Ups! Ocurrió el siguiente error: {s}, seguramente tengas que descargar el motor de idioma de Spacy en las opciones.")
            elif  values["-SPACTRF-"] == True:
                try:
                    nlpner = sp.load('la_core_web_trf')
                    nlp = sp.load('la_core_web_trf')

                #Con spacy usamos dos modelos diferentes, uno sirve para NER y el otro deptrees y lematizacion
                except Exception as s:
                    sg.popup(f"Ups! Ocurrió el siguiente error: {s}, seguramente tengas que descargar el transformer del motor de idioma de Spacy en las opciones.")

            try:
                my_file = pd.read_excel(original_file, sheet_name= sheet_name)
                my_file_index = my_file.index
                number_of_rows = len(my_file_index)
                sg.popup(f"La columna que vamos a analizar tiene {number_of_rows} filas.")
# Ejecuto acá los fors:
                i = 0
                stopwords= list(STOP_WORDS)
                for row_info in my_file[column]:
                    text = nlp(str(row_info))
                    morph_list= list_morphology(text)
                    lemmas_list = list_all_lemmas(text)
                    nombres_list = list_proper_nouns(text)
                    sg.OneLineProgressMeter('Avance de la operación.',i, number_of_rows, 'OK')
                    i=i+1



                if values["-KEEPROWONECOL-"] == True or values["-ONECOL-"] == True:
                    new_df = pd.DataFrame({"Todas las palabras" : columna})
                if values["-NOUNV-"] == True:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})                            
                else:
                    pass
                if values["-PROPN-"] == True:
                    pass
                elif values["-NOUNV-"] == True and values["-PROPN-"] == True and values["-NER-"] == True:
                    if values["-NERL-"] == True:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})
                        new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df4 = pd.DataFrame({"Labels" : labels})
                        new_df5 = pd.DataFrame({"Morfologia" : morfologia})                        
                        new_df = new_df1.join(new_df2.join(new_df3.join(new_df4.join(new_df5))))
                    else:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})
                        new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df = new_df1.join(new_df2.join(new_df3))
                elif values["-NOUNV-"] == True and values["-PROPN-"] == True and values["-NER-"] == False:
                    new_df1 = pd.DataFrame({"Lemas" : palabras})
                    new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                    new_df = new_df1.join(new_df2)
                elif values["-NOUNV-"] == True and values["-PROPN-"] == False and values["-NER-"] == True : 
                    if values["-NERL-"] == True:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df4 = pd.DataFrame({"Labels" : labels})                   
                        new_df = new_df1.join(new_df3.join(new_df4))
                    else:
                        new_df1 = pd.DataFrame({"Lemas" : palabras})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df = new_df1.join(new_df3)
                elif values["-NOUNV-"] == False and values["-PROPN-"] == True and values["-NER-"] == True :
                    if values["-NERL-"] == True:
                        new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df4 = pd.DataFrame({"Labels" : labels}) 
                        new_df = new_df2.join(new_df3.join(new_df4))                      
                    else:
                        new_df2 = pd.DataFrame({"Nombres propios" : nombres})
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df = new_df2.join(new_df3)
                elif values["-NOUNV-"] == False and values["-PROPN-"] == True and values["-NER-"] == False : 
                    new_df = pd.DataFrame({"Nombres propios" : nombres})
                elif values["-NOUNV-"] == True and values["-PROPN-"] == False and values["-NER-"] == False:
                    new_df = pd.DataFrame({"Lemas" : palabras})
                elif values["-NOUNV-"] == False and values["-PROPN-"] == False and values["-NER-"] == True:
                    if values["-NERL-"] == True:
                        new_df3 = pd.DataFrame({"Entidades Naturales" : entidades})
                        new_df4 = pd.DataFrame({"Labels" : labels}) 
                        new_df = new_df3.join(new_df4)                     
                    else:
                        new_df = pd.DataFrame({"Entidades Naturales" : entidades})
#Ahora según la opción de archivo a generar, vamos a generar el archivo final                
                if values["-XLSX-"] == True:
                    new_file = new_file+".xlsx"
                    new_df.to_excel(new_file,sheet_name='NLP', index = False)
                elif values["-CSV-"] == True:
                    new_file = new_file+".csv"
                    new_df.to_csv(new_file, index = False)
                elif values["-JSON-"] == True:
                    new_file = new_file+".json"
                    new_df.to_json(path_or_buf=new_file)
                sg.popup(f'Archivo {new_file} creado correctamente.')
                if values["-DISPLAY-"] == True :
                    options = {"fine_grained": True, "compact": False, "add_lema": True, "color": "blue"}                        
                    wv.create_window('Current Spacy Row', html= displacy.render(sentencener, style="ent", options=options))
                    wv.start()
                    wv.create_window('Current Spacy Row', html= displacy.render(sentencener, style="dep", options=options))
                    wv.start()
            except Exception as e:
                sg.popup(f"Ups! Ocurrió el siguiente error: {e}")
    if event == "Cerrar" or event == sg.WIN_CLOSED:
        force_exit = True
        break
window.close()



