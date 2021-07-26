from bs4 import *
import requests
import pandas as pd

ListFind = {'2020-02 Cumulative Update for Windows 10 Version 1909 for x64-based Systems (KB4532693)':'KB4532693',
            'MS15-081: Security Update for Microsoft Office 2016 (KB3085538) 64-Bit Edition':'KB3085538'
            }

diccionario = {}
ValorBuscado = []
ValorResultado = []


texto1 = " "
titulo1 = " "
urlArq = " "

publico = ""
arQuitect = 'n/d'

''' 
--- Función que recorre los titulos de ListFind vector por vector e identifica los Id ---
'''

def arreglarid(idd):
    title = 0
    positionnext = ""
    for i in range(len(idd)):
        for ii in list(idd[i]):
            if title == 0:
                if ii == '_':
                    title = 1
                else:
                    positionnext = positionnext + ii
                    #print(positionnext) # --- Verificación del recorrido del Id --- '''
            else:
                continue
    return positionnext

''' 
--- Ciclo que permite incluir en la URL cada uno de los KB de la lista que se van a validar (ListFind) --- 
'''

for actual, kb in ListFind.items():

    url = 'https://www.catalog.update.microsoft.com/Search.aspx?q={}'.format(kb)
    pagina = requests.get(url) # --- Con esta linea se valida si la petición responde 200 ---

    soup = BeautifulSoup(pagina.content, 'html.parser')

    #print(url) # --- Con esta l{inea se valida lo que sucede con la asignación ---
    #print(pagina) # --- Esta linea imprime la petición ---
    #print(soup) # --- Esta linea imprime el recorrido de cada uno de los elementos representados en etiquetas del HTML del Catalogo ---


    listkb = soup.find_all('td', class_='resultsbottomBorder resultspadding')   # Recorrido de las etiquetas de la página web Windows Update Catalog

    if listkb:

        # --- Este segmento de Código recorre cada titulo ---

        for i in listkb:
            if 'KB' in (i.text).strip():
                cadena = arreglarid(i['id'])
                diccionario[(i.text).strip()] = cadena # --- Alimenta un nuevo diccionario ---

                #print('resultado: ', diccionario) # --- Recupera el Id junto con el título y la arquitectura de cada uno de los titulos que contiene el KB a validar ---
                #print('cadena: ', cadena) # --- Recupera lista de Id de los titulos que contiene el KB a validar ---

        '''
        --- Consulta por arquitectura ---
        '''

        for titleArq, idArq in diccionario.items():
            for ii, kbArq in ListFind.items():
                if titleArq == ii:
                    urlArq = 'https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid={}'.format(idArq)
                    arQuitect = titleArq

                    #print(titleArq) # --- Mantener el título del KB que se esta consultado en la sección Arquitectura ---
                    #print(idArq) # --- Mantener el Id del KB que se esta consultado en la sección Arquitectura ---
                    #print(kbArq)  # --- Mantener el KB que se esta consultado en la sección Arquitectura ---

                elif kbArq in titleArq:
                    urlArq = 'https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid={}'.format(idArq)
                    arQuitect = titleArq
                    #print(arQuitect)  # --- Mantener el KB que se esta consultado en la sección Arquitectura ---
                else:
                    continue

        """
        --- Consulta del último KB actualizado y lanzado por Microsoft, para ello se validan todos los KB anteriores.
        """
        ventana = requests.get(urlArq) # --- Con esta linea se valida si la petición responde 200 ---
        #print(ventana) # --- Esta linea imprime la petición ---

        soup2 = BeautifulSoup(ventana.content, 'html.parser')
        #print(soup2) # --- Esta linea imprime el recorrido de cada uno de los elementos representados en etiquetas del HTML del Catalogo. ---

        listkb2 = soup2.find_all("div", id="supersededbyInfo")

        #print(listkb2) # Imprcmdime la captura de la lista completa de los KB que se han lanzado para la version y arquitectura del Sistema Operativo. ---

        texto = " "

        for i in listkb2: # --- Recorrido de la lista para confirmar la actualización ---

            if (i.text).strip() != 'n/a':
                urlIdRecord = i.a.get('href')
                texto = i.a.text
                conf = 0

                while texto != 'n/a':
                    if conf == 0:
                        urlIdRecord = i.a.get('href')
                        texto = i.a.text
                        conf = 1

                    else:
                        texto = texto1
                        urlIdRecord = titulo1

                        #print(texto1)  # --- Imprime el final del recorrido de los titulos generales
                        #print(titulo1)  # --- Imprime el final del recorrido de los titulos generales

                    url3 = 'https://www.catalog.update.microsoft.com/{}'.format(urlIdRecord)
                    pagina3 = requests.get(url3)
                    soup3 = BeautifulSoup(pagina3.content, 'html.parser')
                    listkb3 = soup3.find_all("div", id="supersededbyInfo")
                    final3 = soup3.find_all("span", id="ScopedViewHandler_titleText")

                    for ii in listkb3:
                        try:
                            titulo1 = ii.a.get('href')
                            texto1 = ii.a.text

                        except:
                            texto = 'n/a'

                for i in final3: # --- Ciclo que agrupa el elemento que se quiere validar jutno con el resultado, validandp si este tiene una actualización. ---
                    ValorBuscado.append(actual)
                    ValorResultado.append(i.text)
                publico = texto
                break
            else:
                publico = (i.text).strip()
                ValorBuscado.append(actual)
                ValorResultado.append(actual)
                break
    else:
        texto1 = "El parche {} no esta disponible".format(kb) # --- En caso de no estar la actualización, envia mensaje de no esta disponible. ---
        ValorBuscado.append(actual)
        ValorResultado.append(texto1)

fd = pd.DataFrame({'ValorBuscado': ValorBuscado, 'ValorResultado': ValorResultado}, index=range(len(ValorBuscado))) # --- Construcción de columnas y datos en formato JSON. ---

fd.to_csv(r'C:\Users\pccar\Desktop\ScriptKBpy\resultado.csv', index=False) # --- Ruta donde se quiere descargar el archivo resultante .CSV. ---

#print('Archivo: ', fd) # Línea para imprimir resultado final.