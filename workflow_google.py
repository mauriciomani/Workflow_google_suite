from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from email.mime.text import MIMEText
from googleapiclient import errors
import json
import base64

class workflow():
    def __init__(self, SCOPES, CREDENTIALS_PATH):
        self.SCOPES = SCOPES
        #Se debe de ingresar la ruta completa para el archivo de credenciales
        #Por ejemplo: 'C:/Name/Desktop/credentials.json'
        self.CREDENTIALS_PATH = CREDENTIALS_PATH
    
    # Versiones bajo las que fue construido
    def app_version(app):
        app = app.lower()
        if app == 'drive':
            return('drive', 'v3')
        if app == 'sheets':
            return('sheets', 'v4')
        if app == 'docs':
            return('docs', 'v1')
        if app== 'gmail':
            return('gmail', 'v1')
    
    #Almacena el archivo json con el nombre (name) deseado, el path y la variable que lo contiene
    def save_json(self, name, path, doc):
        with open(path + name, 'w') as json_file:
            json.dump(doc, json_file)
        
    # APP es la version con la que se conectaran: sheets, drive, documents
    # Version es la version del api    
    def autenticacion(self):
        """El archivo token.pickle almacena el acceso de usuario y actualiza. 
        Es creado automáticamente cuando el flow de autorización se completa por 
        primera vez"""
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        #Si no existe el archivo token.pickle. El usuario deberá hacer un LOGIN
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                self.CREDENTIALS_PATH, self.SCOPES)
                creds = flow.run_local_server(port=0)
                # Crea y almacena el archivo tocken.pickle
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        return(creds)

    def creacion_de_servicios(self, app, creds):
        #Para crear los servicios de los diferentes scopes
        app, version = workflow.app_version(app)
        #Crea el servicio
        service = build(app, version, credentials=creds)
        return(service)
    
    def registro_json(self, columna, ult_celda):
        registro = {'sheets':{'columna':columna,
                    'ultima_celda': ult_celda,}, 'doc':{}}
        return(registro)

class hoja_calculo():
    #Enable la API de shetes
    def __init__(self, SHEET_ID, service):
        #El id del sheet que se puede obtener del link del mismo
        self.SHEET_ID = SHEET_ID
        self.service = service
        
    def hoja_extraer(self, name, column_start, number_start, column_end, number_end = '', all_column = True):
        #Tiene como objetivo retornar el nombre de la hoja sobre la que se quiere trabajar en el archivo
        if all_column == True:
            #Retorno para Regresar toda la columna
            return(name + '!' + column_start + str(number_start) + ':' + column_end)
        else:
            return(name + '!' + column_start + str(number_start) + ':' + column_end + str(number_end))
    
    def extraer_valores(self, range_name): 
        #Se conecta con la autenticación
        sheet = self.service.spreadsheets()
        #Extrae los valores de la hoja_extraer
        result = sheet.values().get(spreadsheetId=self.SHEET_ID,
                                range=range_name).execute()
        values = result.get('values', [])
        return(values)

class documents():
    #Enable la API de docs
    def __init__(self, service, doc_id, titulo = 'titulo', parrafo = 'parrafo', color = [1.0, 0.0, 0.0]):
        self.service = service
        #Titulo del insert
        self.titulo = titulo
        #Parrafo del insert
        self.parrafo = parrafo
        self.color = color
        self.doc_id = doc_id
    
    #La terminación p son los atributos del parrafo
    def insertar_texto(self, size, size_p, font = "Times New Roman", bold = True, italic = False, align = 'CENTER',
                        italic_p = False, font_p = "Times New Roman", align_p = 'START'):
        #Concatena el titulo y el parrafo, titulo un espacio y despues 2 espacios al final del parrafo
        texto = self.titulo + '\n\n' + self.parrafo + '\n\n\n'
        #Tamaño de la cadena de texto titulo más 2 caracteres para funcionar
        n_titulo = len(self.titulo) + 2
        n_parrafo = len(self.parrafo) + 1
        requests = [{"insertText":{"text": texto, "location":{"index":1}}},
                     {'updateTextStyle': {
                'range': {'startIndex': 1,'endIndex': n_titulo},
                'textStyle': {'bold': bold,'italic': italic,
                        'weightedFontFamily': {'fontFamily': font}, 
                                'fontSize': {'magnitude': size,'unit': 'PT'},
                                'foregroundColor': {'color': {'rgbColor': {
                                'blue': self.color[0],'green': self.color[1],'red': self.color[2]}}}},
                'fields': 'bold,italic,weightedFontFamily,fontSize,foregroundColor'}}, 
                {'updateParagraphStyle': 
                    {'range': {'startIndex': 1,'endIndex':  n_titulo},
                'paragraphStyle':{'alignment': align},'fields': 'alignment'}},
                {'updateTextStyle': {
                'range': {'startIndex': n_titulo + 1,'endIndex': n_titulo + n_parrafo},
                'textStyle': {'italic': italic_p,
                        'weightedFontFamily': {'fontFamily': font_p}, 
                                'fontSize': {'magnitude': size_p,'unit': 'PT'}},
                'fields': 'italic,weightedFontFamily,fontSize,foregroundColor'}}, 
                {'updateParagraphStyle': 
                    {'range': {'startIndex': n_titulo + 1,'endIndex':  n_titulo + n_parrafo},
                'paragraphStyle':{'alignment': align_p},'fields': 'alignment'}}]
        result = self.service.documents().batchUpdate(documentId=self.doc_id, body={'requests': requests}).execute()
        return(result)

    #Sirve para conocer la plantilla en la cula será insertada el texto
    #Esto es para insertar texto varias veces en un mismo artículo
    #Si se buca crear un nuevo archivo cada vez que se inserta texto usar Merge Text
    def extraer_plantilla(self, path_to_info, name):
        result = self.service.documents().get(documentId=self.doc_id).execute()
        workflow.save_json(name, path_to_info, result)
    
class correo_electronico():
    #Enable la API de Gmail
    def __init__(self, email_from, service):
        self.email_from = email_from
        self.service = service 
        
    def mensaje_a_enviar(self, email_to, text, subject):
        message = MIMEText(text)
        message['to'] = email_to
        message['from'] = self.email_from
        message['subject'] = subject
        message_decode = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
        return({'raw': message_decode})
    
    def enviar_email(self, message):
        #Envia el mensaje proveniente de la funcion message_to_send
        try:
            message = (self.service.users().messages().send(userId=self.email_from, body=message).execute())
            print('Message Id: %s' % message['id'])
            return (message)
        except errors.HttpError as error:
            print ('An error occurred: %s' % error)
    
    def lista_de_mensajes(self):
        results = self.service.users().messages().list(userId='me', labelIds = ['INBOX']).execute()
        info = {'sender':[], 'subject':[], 'ids':[]}
        for i in range(11):
            ids = results['messages'][i]['id']
            result = self.service.users().messages().get(userId='me', id=ids).execute()
            info['ids'].append(ids)
            for e in result['payload']['headers']:
                if e['name']=='from':
                    msg_from = e['value']
                    info['sender'].append(msg_from)
                elif e['name']=='subject':
                    msg_subject = e['value']
                    info['subject'].append(msg_subject)
                else:
                    pass
        return(info)
    
    def info_thread(self, t_id):
        t_data = self.service.users().threads().get(userId='me', id=t_id).execute()
        largo_thread = len(t_data['messages'])
        mensajes = []
        if largo_thread>1:
            for i in largo_thread:
                mensajes.append(t_data['messages'][i]['snippet'])
        else:
            print('SIN MENSAJE RECIBIDO')
        return(mensajes)
        