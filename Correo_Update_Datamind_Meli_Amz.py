import subprocess
import win32com.client as win32
import time

# Ruta completa del archivo a ejecutar primero
ruta_archivo_datamind = r"C:\Users\SSN0609\OneDrive - Stanley Black & Decker\Dashboards LAG\Data Flow\Datamind\VS Code Datamind\Data\Code\Proceso_Update_Datamind.py"

ruta_archivo_meli_amz = r"C:\Users\SSN0609\OneDrive - Stanley Black & Decker\Dashboards LAG\Data Flow\Datamind\VS Code Mercado Libre Amazon\Code\ProcesoETL_Meli_Amz_Update.py"

ruta_archivo_inventario= r'C:\Users\SSN0609\OneDrive - Stanley Black & Decker\Dashboards LAG\Data Flow\Datamind\Inventario 28 dias\code\ProcesoETL_Inventario.py'


# Intentar ejecutar el archivo primero
proceso_datamind = subprocess.Popen(["python", ruta_archivo_datamind])

# Esperar a que el proceso termine y obtener el código de salida
codigo_salida_datamind = proceso_datamind.wait()
print("Esperando para el siguiente proceso de Mercado Libre y Amazon")
time.sleep(30)


# Intentar ejecutar el archivo primero
proceso_meli_amz = subprocess.Popen(["python", ruta_archivo_meli_amz])
# Esperar a que el proceso termine y obtener el código de salida
codigo_salida_meli_amz = proceso_meli_amz.wait()
print("Esperando para el siguiente proceso de Inventario")
time.sleep(30)

# Intentar ejecutar el archivo primero
proceso_inventario = subprocess.Popen(["python", ruta_archivo_inventario])
# Esperar a que el proceso termine y obtener el código de salida
codigo_salida_inventario = proceso_inventario.wait()
print("Esperando para el siguiente proceso envio de correo")
time.sleep(30)

#Proceso_Update_Datamind
#ProcesoETL_Meli_Amz_Update
#ProcesoETL_Inventario
# Configuración del correo electrónico
correo_exitoso = {
    "subject": "Actualizacion Datamind Mercado libre-Amazon e Inventario exitosa",
    "body": "Los codigos de actualizacion:\n Proceso_Update_Datamind.py\n ProcesoETL_Meli_Amz_Update.py\n  ProcesoETL_Inventario\n Se han ejecutado de manera existosa"
}

correo_fallido = {
    "subject": "Actualizacion Datamind Mercado libre-Amazon e Inventario fallida",
    "body": "Los codigos de actualizacion:\n Proceso_Update_Datamind.py\n ProcesoETL_Meli_Amz_Update.py\n  ProcesoETL_Inventario\n NO se han ejecutado de manera existosa"
}


# Configuración del correo electrónico
correo_datamind = {
    "subject": "Actualizacion datamind exitosa pero  mercado libre-amazon e inventario fallida ",
    "body": "El codigos de actualizacion:\n Proceso_Update_Datamind.py se ejecuto correctamente pero los codigos\n ProcesoETL_Meli_Amz_Update.py\n  ProcesoETL_Inventario\n NO se han ejecutado de manera existosa"
}

# Configuración del correo electrónico
correo_meli_amz = {
    "subject": "Actualizacion mercado libre-amazon exitosa pero  datamind e inventario fallida ",
    "body": "El codigos de actualizacion:\n ProcesoETL_Meli_Amz_Update.py  se ejecuto correctamente pero los codigos\n Proceso_Update_Datamind.py\n  ProcesoETL_Inventario\n NO se han ejecutado de manera existosa"
}

correo_inventario = {
    "subject": "Actualizacion inventario exitosa pero  datamind e mercado libre-amazon fallida ",
    "body": "El codigos de actualizacion: ProcesoETL_Inventario \n   se ejecuto correctamente pero los codigos\n Proceso_Update_Datamind.py\n ProcesoETL_Meli_Amz_Update.py\n NO se han ejecutado de manera existosa"
}

correo_Datamind_inventario = {
    "subject": "Actualizacion Datamind e inventario exitosa pero mercado libre-amazon fallida ",
    "body": "El codigos de actualizacion: ProcesoETL_Inventario \n Proceso_Update_Datamind.py\n  se ejecuto correctamente pero los codigos\n ProcesoETL_Meli_Amz_Update.py\n NO se han ejecutado de manera existosa"
}
correo_Datamind_Meli_Amz = {
    "subject": "Actualizacion Datamind y Mercado libre-Amazon  exitosa pero inventario fallida ",
    "body": "El codigos de actualizacion:ProcesoETL_Meli_Amz_Update.py  \n Proceso_Update_Datamind.py\n  se ejecuto correctamente pero los codigos\n  ProcesoETL_Inventario\n NO se han ejecutado de manera existosa"
}

correo_Inventario_Meli_Amz = {
    "subject": "Actualizacion Inventario y Mercado libre-Amazon  exitosa pero Datamind fallida ",
    "body": "El codigos de actualizacion:ProcesoETL_Meli_Amz_Update.py  \n ProcesoETL_Inventario  \n  se ejecuto correctamente pero los codigos\n  Proceso_Update_Datamind.py\n NO se han ejecutado de manera existosa"
}




# Función para enviar correo electrónico a una lista de destinatarios
def enviar_correo(destinatarios, subject, body):
    outlook = win32.Dispatch("Outlook.Application")
    for destinatario in destinatarios:
        mail = outlook.CreateItem(0)
        mail.To = destinatario
        mail.Subject = subject
        mail.Body = body
        mail.Send()

# Lista de destinatarios de correo electrónico
#lista_correos = ['sebastian.nunez@sbdinc.com', 'adrian.orozco@sbdinc.com']
lista_correos = ['sebastian.nunez@sbdinc.com']

# Verificar el código de salida y enviar correo correspondiente a todos los destinatarios
if (codigo_salida_datamind == 0 and codigo_salida_meli_amz ==0 and codigo_salida_inventario==0):
    print("Los codigos de actualizacion:\n Proceso_Update_Datamind.py\n ProcesoETL_Meli_Amz_Update.py\n  ProcesoETL_Inventario\n Se han ejecutado de manera existosa\n Enviando correo...")
    enviar_correo(lista_correos, correo_exitoso["subject"], correo_exitoso["body"])

elif(codigo_salida_datamind != 0 and codigo_salida_meli_amz ==0 and codigo_salida_inventario ==0):
    print("El codigos de actualizacion:ProcesoETL_Meli_Amz_Update.py  \n ProcesoETL_Inventario  \n  se ejecuto correctamente pero los codigos\n  Proceso_Update_Datamind.py\n NO se han ejecutado de manera existosa\n Enviando correo...")
    enviar_correo(lista_correos, correo_datamind["subject"], correo_meli_amz["body"])

elif(codigo_salida_datamind == 0 and codigo_salida_meli_amz !=0 and codigo_salida_inventario ==0):
    print("El codigos de actualizacion: ProcesoETL_Inventario \n Proceso_Update_Datamind.py\n  se ejecuto correctamente pero los codigos\n ProcesoETL_Meli_Amz_Update.py\n NO se han ejecutado de manera existosa\n Enviando correo...")
    enviar_correo(lista_correos, correo_meli_amz["subject"], correo_meli_amz["body"])

elif(codigo_salida_datamind == 0 and codigo_salida_meli_amz ==0 and codigo_salida_inventario !=0):
    print("El codigos de actualizacion:ProcesoETL_Meli_Amz_Update.py  \n Proceso_Update_Datamind.py\n  se ejecuto correctamente pero los codigos\n  ProcesoETL_Inventario\n NO se han ejecutado de manera existosa\n Enviando correo...")
    enviar_correo(lista_correos, correo_inventario["subject"], correo_meli_amz["body"])

elif(codigo_salida_datamind != 0 and codigo_salida_meli_amz !=0 and codigo_salida_inventario ==0):
    print("El codigos de actualizacion: ProcesoETL_Inventario \n   se ejecuto correctamente pero los codigos\n Proceso_Update_Datamind.py\n ProcesoETL_Meli_Amz_Update.py\n NO se han ejecutado de manera existosa\n Enviando correo...")
    enviar_correo(lista_correos, correo_Datamind_Meli_Amz["subject"], correo_meli_amz["body"])


elif(codigo_salida_datamind == 0 and codigo_salida_meli_amz !=0 and codigo_salida_inventario !=0):
    print("El codigos de actualizacion:\n Proceso_Update_Datamind.py se ejecuto correctamente pero los codigos\n ProcesoETL_Meli_Amz_Update.py\n  ProcesoETL_Inventario\n NO se han ejecutado de manera existosa\n Enviando correo...")
    enviar_correo(lista_correos, correo_Inventario_Meli_Amz["subject"], correo_datamind["body"])



elif(codigo_salida_datamind != 0 and codigo_salida_meli_amz ==0 and codigo_salida_inventario !=0):
    print("El codigos de actualizacion:\n ProcesoETL_Meli_Amz_Update.py  se ejecuto correctamente pero los codigos\n Proceso_Update_Datamind.py\n  ProcesoETL_Inventario\n NO se han ejecutado de manera existosa\n Enviando correo...")
    enviar_correo(lista_correos, correo_Datamind_inventario["subject"], correo_meli_amz["body"])




else:
    print("Los codigos de actualizacion:\n Proceso_Update_Datamind.py\n ProcesoETL_Meli_Amz_Update.py\n  ProcesoETL_Inventario\n NO Se han ejecutado de manera existosa\n Enviando correo...")
    enviar_correo(lista_correos, correo_fallido["subject"], correo_fallido["body"])
time.sleep(30)
