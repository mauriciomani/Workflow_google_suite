# -*- coding: utf-8 -*-
"""
Created on Wed Sep 18 21:06:09 2019

@author: mauricioman12
"""

from workflow_google import workflow, correo_electronico, hoja_calculo, documents
import json

def sheets_docs(aut, service_sheets, sheet_id, name_hoja, registros,
                 service_doc, doc_id, titulo, parrafo, color, size, size_p, font, font_p):
    calc = hoja_calculo(sheet_id, service_sheets, titulo, parrafo, color)
    #trae las columnas 
    range_name = calc.hoja_extraer(name_hoja ,registros['sheets']['columna'],
                                   registros['sheets']['ultima_celda'] + 1,
                                   registros['sheets']['columna'], True)
    valores = calc.extraer_valores(range_name)
    #Almacena los valores del ultimo registro, la columna, y el ultimo valor
    registro = aut.registro_json(registros['sheets']['columna'], registros['sheets']['ultima_celda'] + len(valores))
    aut.save_json('registro', path2, registro)
    doc = documents(service_doc, doc_id, titulo, parrafo, color)
    result= doc.insertar_texto(size, size_p, font, font_p)
    return(valores, result)
    
def sheets_gmail(me, service_email, subject, texto, to):
    email = correo_electronico(me, service_email)
    texto = email.mensaje_a_enviar(to, subject, texto)
    mensaje = email.enviar_email(texto)
    lista = email.lista_de_mensajes()




SCOPES = ['https://mail.google.com/', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/documents']
#Replace for paths
path = 'C:/Users/path/first'
path2 = 'C:/Users/path/second'
aut = workflow(SCOPES, path)
cred = aut.autenticacion()
service_sheets = aut.creacion_de_servicios('sheets', cred)
service_doc = aut.creacion_de_servicios('docs', cred)
service_email = aut.creacion_de_servicios('gmail', cred)

registros = json.loads(path2 + 'registro.json')
with open('C:/Users/save/json/', 'r') as myfile:
    data=myfile.read()
json.loads(data)


