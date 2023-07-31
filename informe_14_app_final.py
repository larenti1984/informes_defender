import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
import os.path
import tkinter as tk
from tkinter import ttk
import win32com.client as win32
import tkinter.messagebox as messagebox
from tabulate import tabulate

# Obtener la ruta del directorio actual
directorio_actual = os.getcwd()

# Construir la ruta relativa al archivo
ruta_archivo = os.path.join(directorio_actual, 'Reporte_IT_Pruebas.xlsx')

# Construir la ruta completa a la imagen neorisIT.jpg en la carpeta raíz del script
ruta_imagen = os.path.join(directorio_actual, 'neorisIT.jpg')

def verificar_archivo():
    if not os.path.isfile(ruta_archivo):
        messagebox.showerror("Error", "El archivo no existe.")
        return False
    elif os.path.exists(ruta_archivo):
        try:
            pd.read_excel(ruta_archivo).to_excel(ruta_archivo) # Intenta abrir y cerrar el archivo
            return True
        except:
            messagebox.showerror("Error", "El archivo está abierto por otro programa.")
            return False
    else:
        return False

def enviar_correos():
    if not verificar_archivo():
        return

    df = pd.read_excel(ruta_archivo)
    servidor_smtp = 'smtp.office365.com'  # Cambia esto según tu proveedor de correo
    puerto_smtp = 587  # Cambia esto según tu proveedor de correo
    remitente = 'matias.larenti@neoris.com'  # Cambia esto por tu dirección de correo electrónico
    password = 'Micha.6567!'  # Cambia esto por tu contraseña de correo electrónico

    grupo_correo = df.groupby('mail')
    for correo, grupo in grupo_correo:
        usuario = correo.split('@')[0] + '@neoris.com'
        datos_tabla = grupo[['VulnerabilitySeverityLevel', 'SoftwareName', 'RecommendedSecurityUpdate']]
        tabla = tabulate(datos_tabla, headers=['Urgencia', 'Aplicación', 'Recomendación'], tablefmt='pretty')

        # Leer el contenido del archivo correo2.html
        with open('correo2.html', 'r', encoding='utf-8') as archivo_html:
            contenido_html = archivo_html.read()

        # Incorporar el contenido del informe detallado en el HTML
        contenido_html = contenido_html.replace('Detalle de la notificación', tabla)

        # Crear el objeto MIMEMultipart para el correo
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = correo
        mensaje['Subject'] = 'Problemas de updates en tu equipo'

        # Agregar el cuerpo del correo en formato HTML al objeto MIMEText
        mensaje.attach(MIMEText(contenido_html, 'html'))

        # Agregar la imagen adjunta
        with open(ruta_imagen, 'rb') as img_file:
            imagen_adjunta = MIMEImage(img_file.read(), name=os.path.basename(ruta_imagen))
        imagen_adjunta.add_header('Content-ID', '<neorisIT.jpg>')
        mensaje.attach(imagen_adjunta)

        # Enviar el correo utilizando el cliente de Outlook
        outlook = win32.Dispatch('outlook.application')
        correo_salida = outlook.CreateItem(0)
        correo_salida.To = correo
        correo_salida.Subject = 'Problemas de updates en tu equipo'
        correo_salida.HTMLBody = contenido_html
        correo_salida.Send()

        print(f"Correo enviado a: {correo}")
# Crear la interfaz gráfica
root = tk.Tk()
root.title("Informe y Envío de Correos")
root.geometry("400x200")
root.resizable(False, False)

def verificar_archivo_click():
    verificar_archivo()

def ver_informe_click():
    ver_informe()

def enviar_correos_click():
    enviar_correos()

# Crear el botón para verificar archivo
btn_verificar_archivo = ttk.Button(root, text="Verificar Archivo", command=verificar_archivo_click)
btn_verificar_archivo.pack(pady=20)

# Crear el botón para ver el informe
btn_ver_informe = ttk.Button(root, text="Ver informe", command=ver_informe_click)
btn_ver_informe.pack(pady=5)

# Crear el botón para enviar correos
btn_enviar_correos = ttk.Button(root, text="Enviar correos", command=enviar_correos_click)
btn_enviar_correos.pack()

# Iniciar la interfaz gráfica
root.mainloop()
