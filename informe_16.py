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
# ruta absoluta
# ruta_imagen = 'C:/Users/matias.larenti/OneDrive - neoris.com/Desktop/Git_Codigos/infDef_python/neorisIT.jpg'

def verificar_archivo():
    if not os.path.isfile(ruta_archivo):
        messagebox.showerror("Error", "El archivo no existe.")
        return False
    elif os.path.exists(ruta_archivo):
        try:
            # Verificar que el archivo se pueda leer correctamente
            pd.read_excel(ruta_archivo)
            messagebox.showinfo("Información", "Archivo verificado correctamente.")
            return True
        except:
            messagebox.showerror("Error", "El archivo está abierto por otro programa.")
            return False
    else:
        return False
    
def verificar_archivo_imagen():
    if not os.path.isfile(ruta_imagen):
        messagebox.showerror("Error", "La imagen 'neorisIT.jpg' no existe en la ubicación especificada.")
        return False
    else:
        return True

def ver_informe():
    if not verificar_archivo():
        return

    df = pd.read_excel(ruta_archivo)
    grupo_correo = df.groupby('mail')
    resumen = {}

    for correo, grupo in grupo_correo:
        cantidad_items = len(grupo)
        resumen[correo] = cantidad_items

    usuarios_correo = [f"{correo} - {resumen[correo]} items" for correo in resumen.keys()]
    resumen_detallado = "\n".join(usuarios_correo)

    # Mostrar el informe detallado en un cuadro de diálogo de mensaje
    messagebox.showinfo("Informe Detallado", resumen_detallado)

def enviar_correos():
    if not verificar_archivo() or not verificar_archivo_imagen():
        return

    df = pd.read_excel(ruta_archivo)
    servidor_smtp = 'smtp.office365.com'  # Cambia esto según tu proveedor de correo
    puerto_smtp = 587  # Cambia esto según tu proveedor de correo
    remitente = 'matias.larenti@neoris.com'  # Cambia esto por tu dirección de correo electrónico
    password = 'Micha.6567!'  # Cambia esto por tu contraseña de correo electrónico

    # Crear el objeto outlook una sola vez fuera del bucle for
    outlook = win32.Dispatch('outlook.application')

    grupo_correo = df.groupby('mail')
    for correo, grupo in grupo_correo:
        datos_tabla = grupo[['VulnerabilitySeverityLevel', 'SoftwareName', 'RecommendedSecurityUpdate']]

        # Formatear la tabla como HTML utilizando tabulate
        tabla_html = datos_tabla.to_html(index=False, justify='left')

        # Leer el contenido del archivo correo2.html
        with open('correo2.html', 'r', encoding='utf-8') as archivo_html:
            contenido_html = archivo_html.read()

        # Incorporar el contenido del informe detallado en el HTML
        contenido_html = contenido_html.replace('Detalle de la notificación', tabla_html)

        # Crear el objeto MIMEMultipart para el correo
        mensaje = MIMEMultipart('related')
        mensaje['From'] = remitente
        mensaje['To'] = correo
        mensaje['Subject'] = 'Problemas de updates en tu equipo'

        # Agregar el cuerpo del correo en formato HTML al objeto MIMEText
        mensaje.attach(MIMEText(contenido_html, 'html'))

        # Agregar la imagen adjunta
        with open(ruta_imagen, 'rb') as img_file:
            imagen_adjunta = MIMEImage(img_file.read(), _subtype='jpeg', name=os.path.basename(ruta_imagen))
        imagen_adjunta.add_header('Content-ID', '<neorisIT.jpg>')
        mensaje.attach(imagen_adjunta)

        # Enviar el correo utilizando el cliente de Outlook
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
btn_ver_informe = ttk.Button(root, text="Ver Informe", command=ver_informe_click)
btn_ver_informe.pack()

# Crear el botón para enviar correos
btn_enviar_correos = ttk.Button(root, text="Enviar Correos", command=enviar_correos_click)
btn_enviar_correos.pack()

# Iniciar la interfaz gráfica
root.mainloop()
