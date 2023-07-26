import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tabulate import tabulate
from tkinter import *
from tkinter import messagebox
import os

def verificar_archivo():
    ruta_archivo = 'C:/Proyectos/infDef_python/informes_defender/informe_modificado.xlsx'

    if not os.path.exists(ruta_archivo):
        messagebox.showerror("Error", f"El archivo {ruta_archivo} no existe.")
        return False

    try:
        pd.read_excel(ruta_archivo)
    except PermissionError:
        messagebox.showerror("Error", f"El archivo {ruta_archivo} está abierto. Cierre el archivo y vuelva a intentarlo.")
        return False

    return True

def mostrar_resumen():
    if not verificar_archivo():
        return

    df = pd.read_excel(ruta_archivo)
    grupo_correo = df.groupby('correo')
    resumen = {}

    for correo, grupo in grupo_correo:
        usuario = grupo.iloc[0]['usuario']
        cantidad_items = len(grupo)
        resumen[correo] = {
            'Usuario': usuario,
            'Cantidad de items': cantidad_items
        }

    resumen_texto = ""
    for correo, info in resumen.items():
        resumen_texto += f"- Correo: {correo}\n"
        resumen_texto += f"  Usuario: {info['Usuario']}\n"
        resumen_texto += f"  Cantidad de items: {info['Cantidad de items']}\n\n"

    messagebox.showinfo("Resumen de Correos", f"Se enviarán {len(resumen)} correos:\n\n{resumen_texto}")

def enviar_correos():
    if not verificar_archivo():
        return

    confirmacion = messagebox.askquestion("Confirmación", "¿Deseas enviar los correos?")

    if confirmacion == 'yes':
        df = pd.read_excel(ruta_archivo)
        grupo_correo = df.groupby('correo')
        resumen = {}

        for correo, grupo in grupo_correo:
            destinatario = correo
            usuario = grupo.iloc[0]['usuario']
            cantidad_items = len(grupo)

            datos_tabla = grupo[['VulnerabilitySeverityLevel', 'SoftwareName', 'RecommendedSecurityUpdate']]
            tabla = tabulate(datos_tabla, headers=['Urgencia', 'Aplicacion', 'Recomendacion'], tablefmt='orgtbl')

            cuerpo_correo = f"Estimado {usuario}, detectamos que su equipo tiene los siguientes problemas de updates:\n\n{tabla}"

            servidor_smtp = 'smtp.gmail.com'
            puerto_smtp = 587
            remitente = 'neorisp@gmail.com'
            password = 'hrwzhgivekxmibir'

            mensaje = MIMEMultipart()
            mensaje['From'] = remitente
            mensaje['To'] = destinatario
            mensaje['Subject'] = 'Problemas de updates en tu equipo'
            mensaje.attach(MIMEText(cuerpo_correo, 'plain'))

            with smtplib.SMTP(servidor_smtp, puerto_smtp) as servidor:
                servidor.starttls()
                servidor.login(remitente, password)
                servidor.send_message(mensaje)

        messagebox.showinfo("Envío de Correos", "Los correos han sido enviados correctamente.")
    else:
        messagebox.showinfo("Envío de Correos", "No se enviaron los correos.")

# Crear ventana
ventana = Tk()
ventana.title("Envío de Correos")
ventana.geometry("300x200")

# Botón para mostrar el resumen
btn_resumen = Button(ventana, text="Mostrar Resumen", command=mostrar_resumen)
btn_resumen.pack(pady=20)

# Botón para enviar los correos
btn_enviar = Button(ventana, text="Enviar Correos", command=enviar_correos)
btn_enviar.pack()

# Ejecutar la ventana
ventana.mainloop()
