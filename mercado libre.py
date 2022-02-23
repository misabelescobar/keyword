# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""
import poplib #Librería para hacer conexión a servidor de gmail
from email import parser #Para ingresar a los objetos del correo
import pandas as pd #Para manejo de BD (excel)
import os 

ubicacion_actual = os.getcwd()

#Se crean listas para ser llenadas
fecha_correo = []
remitente_correo = []
asunto_correo = []

#Input para solicitar correo y contraseña
print("Ingresa correo electrónico: ")
username = input()

print("Ingresa contraseña: ")
password = input()

#Declarando el servidor al que me conecto
pop3_server = 'pop.gmail.com' 
#username = 'retomariameli@gmail.com'
#password = 'M12345678*'

try:
    email_obj = poplib.POP3_SSL(pop3_server) #conectarme al servidor de gmail
    print(email_obj.getwelcome()) #para confirmar la conexión exitosa
    email_obj.user(username) #al objeto le asigno usuario y contraseña ingresados
    email_obj.pass_(password)
    conexion_exitosa = 1
    
except:
    print("Problemas en la conexión, verifique usuario y contraseña.")
    print("Vuelva a ejecutar el programa")
    
    conexion_exitosa = 0

if conexion_exitosa == 1:
     
    #Obtengo los mensajes del servidor
    messages = [email_obj.retr(i) for i in range(1, len(email_obj.list()[1]) + 1)]
    
    
    # Concatena los mensajes
    messages = ['\n'.join(map(bytes.decode, mssg[1])) for mssg in messages]
    
    #Analiza gramaticalmente los mensajes
    messages = [parser.Parser().parsestr(mssg) for mssg in messages]
    
    try:
        df_origen = pd.read_excel(ubicacion_actual+"/listado_correos.xlsx")
    except FileNotFoundError:
        df_origen = pd.DataFrame(columns=['Fecha correo', 'Remitente', 'Asunto'])
        
    
    #Se inicia recorrer cada uno de los mensajes
    for message in messages:
        print ('***************')
        print (message['subject'])
        remitente = message['from']
        remitente=remitente[remitente.find("<")+1:remitente.find(">") ]
        print (remitente)
        fecha = (message['Date'])
        print (fecha)
        
        existe = df_origen.isin([fecha,remitente,message['subject']]).any().any()
        
        if existe == True:
            continue
        else:
        #Busca el contenido del cuerpo del mensaje
            for parte_cuerpo in message.walk():
                if parte_cuerpo.get_content_type(): #Valida si el contenido del cuerpo existe
                    body = parte_cuerpo.get_payload(decode=True) #Coge el HTML del cuerpo
                    encoding = 'utf-8'
                    try: #Haga lo que está dentro del try si es un string
                        cuerpo = str(body, encoding) #Codifica de HTML a utf-8 (string)
                        if cuerpo.find("Risk") > 0: #Busco la palabra Risk
                            print('Contiene Risk')
                            fecha_correo.append(fecha) #Rellena la lista con los datos
                            remitente_correo.append(remitente)
                            asunto_correo.append(message['subject'])
                            break #pasar al siguiente correo porque encontró Risk
                    except: #En caso que no sea string continue a la siguiente parte del cuerpo
                        continue #Pasar a la siguiente parte del cuerpo
          
    df = pd.DataFrame() #Para crear BD a Excel
    df['Fecha correo'] = fecha_correo
    df['Remitente'] = remitente_correo
    df['Asunto'] = asunto_correo
    
    df = pd.concat([df_origen, df])
    df.to_excel(ubicacion_actual+"/listado_correos.xlsx", index=False)
    
