import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tabulate import tabulate
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

# Obtener la ruta del directorio actual
directorio_actual = os.getcwd()

# Construir la ruta relativa al archivo
ruta_archivo = os.path.join(directorio_actual, 'informe_original_no_modificar.xlsx')

def verificar_archivo():
    if not os.path.isfile(ruta_archivo):
        messagebox.showerror("Error", "El archivo no existe.")
        return False
    elif os.path.exists(ruta_archivo + ".lock"):
        messagebox.showerror("Error", "El archivo está abierto por otro programa.")
        return False
    else:
        return True

def mostrar_resumen():
    if not verificar_archivo():
        return

    df = pd.read_excel(ruta_archivo)
    total_registros = len(df)

    # Crear la ventana de progreso
    ventana_progreso = tk.Toplevel(root)
    ventana_progreso.title("Progreso")
    ventana_progreso.geometry("200x100")

    # Etiqueta de progreso
    lbl_progreso = tk.Label(ventana_progreso, text="Cargando datos...")
    lbl_progreso.pack(pady=10)

    # Barra de progreso
    barra_progreso = ttk.Progressbar(ventana_progreso, length=200, mode='determinate')
    barra_progreso.pack(pady=10)
    barra_progreso["maximum"] = total_registros

    # Función para actualizar la barra de progreso
    def actualizar_progreso(valor):
        barra_progreso["value"] = valor
        ventana_progreso.update_idletasks()

    # Función para cerrar la ventana de progreso
    def cerrar_ventana_progreso():
        ventana_progreso.destroy()

    # Agrupar los datos por correo y generar el cuerpo del correo
    grupo_correo = df.groupby('LoggedOnUsers')
    resumen = {}

    for i, (correo, grupo) in enumerate(grupo_correo):
        cantidad_items = len(grupo)
        resumen[correo] = cantidad_items

        # Actualizar la barra de progreso
        actualizar_progreso(i+1)

    resumen_total = f"Total de correos: {len(resumen)}\n"

    opcion = messagebox.askyesno("Resumen", resumen_total + "¿Deseas ver el resumen detallado?")
    if opcion:
        resumen_detallado = "\n".join([f"{correo}: {cantidad_items} items" for correo, cantidad_items in resumen.items()])
        messagebox.showinfo("Resumen Detallado", resumen_detallado)

    opcion_enviar = messagebox.askyesno("Confirmar Envío", "¿Deseas enviar los correos?")
    if opcion_enviar:
        enviar_correos(df)

    # Cerrar la ventana de progreso
    cerrar_ventana_progreso()

def enviar_correos(df):
    servidor_smtp = 'smtp.gmail.com'  # Cambia esto según tu proveedor de correo
    puerto_smtp = 587  # Cambia esto según tu proveedor de correo
    remitente = 'neorisp@gmail.com'  # Cambia esto por tu dirección de correo electrónico
    password = 'hrwzhgivekxmibir'  # Cambia esto por tu contraseña de correo electrónico

    grupo_correo = df.groupby('LoggedOnUsers')
    for correo, grupo in grupo_correo:
        usuario = grupo.iloc[0]['LoggedOnUsers']
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

# Crear la interfaz gráfica
root = tk.Tk()
root.title("Informe de Problemas de Updates")
root.geometry("300x150")

# Botón "Ver Informe"
btn_ver_informe = tk.Button(root, text="Ver Informe", command=mostrar_resumen)
btn_ver_informe.pack(pady=20)

# Botón "Enviar Correos"
btn_enviar_correos = tk.Button(root, text="Enviar Correos", command=lambda: enviar_correos(df))
btn_enviar_correos.pack()

# Iniciar la interfaz gráfica
root.mainloop()
