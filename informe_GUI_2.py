import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tabulate import tabulate
import os.path
import tkinter as tk
from tkinter import messagebox

ruta_archivo = 'C:/Proyectos/infDef_python/informes_defender/informe_modificado.xlsx'

def verificar_archivo():
    if not os.path.isfile(ruta_archivo):
        messagebox.showerror("Error", "El archivo no existe.")
        return False
    elif os.path.exists(ruta_archivo):
        messagebox.showerror("Error", "El archivo está abierto por otro programa.")
        return True
    else:
        return True

def mostrar_resumen(ruta_archivo):
    if not verificar_archivo():
        return

    df = pd.read_excel(ruta_archivo)
    grupo_correo = df.groupby('correo')
    resumen = {}

    for correo, grupo in grupo_correo:
        cantidad_items = len(grupo)
        resumen[correo] = cantidad_items

    resumen_total = f"Total de correos: {len(resumen)}\n"

    opcion = messagebox.askyesno("Resumen", resumen_total + "¿Deseas ver el resumen detallado?")
    if opcion:
        resumen_detallado = "\n".join([f"{correo}: {cantidad_items} items" for correo, cantidad_items in resumen.items()])
        messagebox.showinfo("Resumen Detallado", resumen_detallado)

    opcion_enviar = messagebox.askyesno("Confirmar Envío", "¿Deseas enviar los correos?")
    if opcion_enviar:
        enviar_correos(df)

def enviar_correos(df):
    servidor_smtp = 'smtp.gmail.com'  # Cambia esto según tu proveedor de correo
    puerto_smtp = 587  # Cambia esto según tu proveedor de correo
    remitente = 'neorisp@gmail.com'  # Cambia esto por tu dirección de correo electrónico
    password = 'hrwzhgivekxmibir'  # Cambia esto por tu contraseña de correo electrónico

    grupo_correo = df.groupby('correo')
    for correo, grupo in grupo_correo:
        usuario = grupo.iloc[0]['usuario']
        datos_tabla = grupo[['VulnerabilitySeverityLevel', 'SoftwareName', 'RecommendedSecurityUpdate']]
        tabla = tabulate(datos_tabla, headers=['Urgencia', 'Aplicacion', 'Recomendacion'], tablefmt='orgtbl')

        # Construir el cuerpo del correo
        cuerpo_correo = f"Estimado {usuario}, detectamos que su equipo tiene los siguientes problemas de updates:\n\n{tabla}"

        # Crear el objeto MIMEMultipart para el correo
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = correo
        mensaje['Subject'] = 'Problemas de updates en tu equipo'

        # Agregar el cuerpo del correo al objeto MIMEText
        mensaje.attach(MIMEText(cuerpo_correo, 'plain'))

        # Enviar el correo utilizando el servidor SMTP de Office 365
        with smtplib.SMTP(servidor_smtp, puerto_smtp) as servidor:
            servidor.starttls()
            servidor.login(remitente, password)
            servidor.send_message(mensaje)

        print(f"Correo enviado a: {correo}")

# Mostrar el resumen al iniciar la aplicación
mostrar_resumen(ruta_archivo)

# Iniciar la interfaz gráfica
root = tk.Tk()
root.withdraw()
root.mainloop()
