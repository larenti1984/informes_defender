import pandas as pd
import os.path
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
import os
import win32com.client as win32
from email.mime.multipart import MIMEMultipart  # Cambiar la importación a email.mime.multipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from tabulate import tabulate

# Resto del código (el mismo que antes)



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

def ver_informe():
    if not verificar_archivo():
        return

    df = pd.read_excel(ruta_archivo)
    grupo_correo = df.groupby('mail')
    resumen = {}

    for correo, grupo in grupo_correo:
        cantidad_items = len(grupo)
        resumen[correo] = cantidad_items

    usuarios_correo = [f"{correo.split('@')[0]} - {resumen[correo]} items" for correo in resumen.keys()]
    resumen_detallado = "\n".join(usuarios_correo)

    # Crear una nueva ventana para mostrar el informe detallado
    ventana_informe = tk.Toplevel()
    ventana_informe.title("Informe Detallado")
    ventana_informe.geometry("400x400")

    # Crear un widget de desplazamiento vertical
    scroll_y = tk.Scrollbar(ventana_informe, orient="vertical")

    # Crear un widget de texto con desplazamiento
    txt_informe = scrolledtext.ScrolledText(ventana_informe, wrap=tk.WORD, yscrollcommand=scroll_y.set)
    txt_informe.insert(tk.END, resumen_detallado)
    txt_informe.pack(fill=tk.BOTH, expand=True)

    # Configurar la barra de desplazamiento
    scroll_y.config(command=txt_informe.yview)
    scroll_y.pack(side=tk.RIGHT, fill=tk.Y)

def enviar_correos():
    if not verificar_archivo():
        return

    df = pd.read_excel(ruta_archivo)
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    servidor_smtp = 'smtp.gmail.com'  # Cambia esto según tu proveedor de correo
    puerto_smtp = 587  # Cambia esto según tu proveedor de correo
    remitente = 'neorisp@gmail.com'  # Cambia esto por tu dirección de correo electrónico
    password = 'hrwzhgivekxmibir'  # Cambia esto por tu contraseña de correo electrónico

    grupo_correo = df.groupby('mail')
    for correo, grupo in grupo_correo:
        usuario = correo.split('@')[0] + '@neoris.com'
        datos_tabla = grupo[['VulnerabilitySeverityLevel', 'SoftwareName', 'CveId', 'RecommendedSecurityUpdate']]
        tabla_html = tabulate(datos_tabla.values.tolist(), headers='keys', tablefmt='html')  # Convertir DataFrame a lista de listas

        # Leer el contenido del archivo correo2.html
        with open('correo2.html', 'r', encoding='utf-8') as archivo_html:
            contenido_html = archivo_html.read()

        # Incorporar el contenido del informe detallado en el HTML
        contenido_html = contenido_html.replace('Detalle de la notificación', tabla_html)

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
        mail = outlook.CreateItem(0)
        mail.To = f'"{correo}"'
        mail.Subject = 'Problemas de updates en tu equipo'
        mail.HTMLBody = contenido_html
        mail.Send()

        print(f"Correo enviado a: {correo}")

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Informe y Envío de Correos")
root.geometry("400x200")
root.resizable(False, False)

# Crear el botón para ver el informe
btn_ver_informe = ttk.Button(root, text="Ver informe", command=ver_informe)
btn_ver_informe.pack(pady=20)

# Crear el botón para enviar correos
btn_enviar_correos = ttk.Button(root, text="Enviar correos", command=enviar_correos)
btn_enviar_correos.pack()

# Iniciar la interfaz gráfica
root.mainloop()
