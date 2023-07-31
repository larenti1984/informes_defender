import pandas as pd
import os.path
import tkinter as tk
from tkinter import ttk
from tkinter import scrolledtext
import win32com.client as win32
from tabulate import tabulate  # Agregar esta línea para importar la función tabulate

# El resto del código se mantiene igual
# ...


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
    remitente = 'tu_cuenta@tu_empresa.com'  # Cambiar por la dirección de correo electrónico del remitente

    grupo_correo = df.groupby('mail')
    for correo, grupo in grupo_correo:
        usuario = correo.split('@')[0] + '@neoris.com'
        datos_tabla = grupo[['VulnerabilitySeverityLevel', 'SoftwareName', 'RecommendedSecurityUpdate']]
        tabla_html = tabulate(datos_tabla, headers='keys', tablefmt='html')

        outlook = win32.Dispatch('outlook.application')
        mensaje = outlook.CreateItem(0)  # 0 significa un correo nuevo

        mensaje.To = correo
        mensaje.Subject = 'Problemas de updates en tu equipo'

        # Cuerpo del correo en formato HTML
        contenido_html = f'<html><body><p>Detalle de la notificación:</p>{tabla_html}</body></html>'
        mensaje.HTMLBody = contenido_html

        # Agregar la imagen adjunta
        mensaje.Attachments.Add(ruta_imagen)

        # Enviar el correo
        mensaje.Send()

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
