import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tabulate import tabulate

# Leer la planilla de Excel
df = pd.read_excel('C:/Proyectos/infDef_python/informes_defender/informe_modificado.xlsx')

# Agrupar los datos por correo y generar el cuerpo del correo
grupo_correo = df.groupby('correo')
for correo, grupo in grupo_correo:
    usuario = grupo.iloc[0]['usuario']
    datos_tabla = grupo[['VulnerabilitySeverityLevel', 'RegistryDeviceTag', 'RecommendedSecurityUpdate']]
    tabla = tabulate(datos_tabla, headers=['Urgencia', 'Aplicacion', 'Recomendacion'], tablefmt='orgtbl')

    # Construir el cuerpo del correo
    cuerpo_correo = f"Estimado {usuario}, detectamos que su equipo tiene los siguientes problemas de updates:\n\n{tabla}"

    # Configurar los detalles del correo saliente
    remitente = 'tu_correo@tudominio.com'  # Cambia esto por tu dirección de correo electrónico
    password = 'tu_contraseña'  # Cambia esto por tu contraseña de correo electrónico
    servidor_smtp = 'smtp.office365.com'  # Servidor SMTP de Office 365
    puerto_smtp = 587  # Puerto SMTP de Office 365

    destinatario = correo

    # Crear el objeto MIMEMultipart para el correo
    mensaje = MIMEMultipart()
    mensaje['From'] = remitente
    mensaje['To'] = destinatario
    mensaje['Subject'] = 'Problemas de updates en tu equipo'

    # Agregar el cuerpo del correo al objeto MIMEText
    mensaje.attach(MIMEText(cuerpo_correo, 'plain'))

    # Enviar el correo utilizando el servidor SMTP de Office 365
    with smtplib.SMTP(servidor_smtp, puerto_smtp) as servidor:
        servidor.starttls()
        servidor.login(remitente, password)
        servidor.send_message(mensaje)

    print(f"Correo enviado a: {destinatario}")
