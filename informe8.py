import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import os.path
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
from email.mime.image import MIMEImage
from tabulate import tabulate

# Obtener la ruta del directorio actual
directorio_actual = os.getcwd()

# Construir la ruta relativa al archivo
ruta_archivo = os.path.join(directorio_actual, 'informe_modificado.xlsx')

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
    grupo_correo = df.groupby('correo')
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
    servidor_smtp = 'smtp.gmail.com'  # Cambia esto según tu proveedor de correo
    puerto_smtp = 587  # Cambia esto según tu proveedor de correo
    remitente = 'neorisp@gmail.com'  # Cambia esto por tu dirección de correo electrónico
    password = 'hrwzhgivekxmibir'  # Cambia esto por tu contraseña de correo electrónico

    grupo_correo = df.groupby('correo')
    for correo, grupo in grupo_correo:
        usuario = correo.split('@')[0] + '@neoris.com'
        datos_tabla = grupo[['VulnerabilitySeverityLevel', 'SoftwareName', 'RecommendedSecurityUpdate']]
        tabla_html = tabulate(datos_tabla, headers='keys', tablefmt='html')

        # Leer el contenido del archivo correo2.html
        with open('correo2.html', 'r', encoding='utf-8') as archivo_html:
            contenido_html = archivo_html.read()

        # Incorporar el contenido del informe detallado en el HTML
        contenido_html = contenido_html.replace('Detalle de la notificación', tabla_html)
        contenido_html = contenido_html.replace('<img width=625 height=278 style=\'width:6.5104in;height:2.8958in\' id="Content_x0020_Placeholder_x0020_4" src="cid:image001.jpg@01D9BFE3.7A905120"></span>', '<img width=625 height=278 style=\'width:6.5104in;height:2.8958in\' id="Content_x0020_Placeholder_x0020_4" src="cid:neorisIT.jpg">')

        # Crear el objeto MIMEMultipart para el correo
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = correo
        mensaje['Subject'] = 'Notificacion Defender'

        # Agregar el cuerpo del correo en formato HTML al objeto MIMEText
        mensaje.attach(MIMEText(contenido_html, 'html'))

        # Agregar la imagen adjunta
        with open(ruta_imagen, 'rb') as img_file:
            imagen_adjunta = MIMEImage(img_file.read(), name=os.path.basename(ruta_imagen))
        imagen_adjunta.add_header('Content-ID', '<neorisIT.jpg>')
        mensaje.attach(imagen_adjunta)

        # Enviar el correo utilizando el servidor SMTP de Office 365
        with smtplib.SMTP(servidor_smtp, puerto_smtp) as servidor:
            servidor.starttls()
            servidor.login(remitente, password)
            servidor.send_message(mensaje)

        print(f"Correo enviado a: {correo}")

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Informe y Envío de Correos")
root.geometry("400x200")
root.resizable(False, False)

def ver_informe_click():
    ver_informe()

def enviar_correos_click():
    enviar_correos()

# Crear el botón para ver el informe
btn_ver_informe = ttk.Button(root, text="Ver informe", command=ver_informe_click)
btn_ver_informe.pack(pady=20)

# Crear el botón para enviar correos
btn_enviar_correos = ttk.Button(root, text="Enviar correos", command=enviar_correos_click)
btn_enviar_correos.pack()

# Iniciar la interfaz gráfica
root.mainloop()
