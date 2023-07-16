import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from tabulate import tabulate
import os

# Ruta del archivo Excel
ruta_archivo = 'C:/Proyectos/infDef_python/informes_defender/informe_modificado.xlsx'

# Verificar si el archivo existe y está cerrado
if not os.path.exists(ruta_archivo):
    print(f"El archivo {ruta_archivo} no existe.")
    # Puedes agregar aquí el código para manejar esta situación de error
    # Por ejemplo, puedes lanzar una excepción o enviar una notificación.

try:
    df = pd.read_excel(ruta_archivo)
except PermissionError:
    print(f"El archivo {ruta_archivo} está abierto. Cierre el archivo y vuelva a intentarlo.")
    # Puedes agregar aquí el código para manejar esta situación de error
    # Por ejemplo, puedes lanzar una excepción o enviar una notificación.

# Agrupar los datos por correo y generar el cuerpo del correo
grupo_correo = df.groupby('correo')
resumen = {}

for correo, grupo in grupo_correo:
    usuario = grupo.iloc[0]['usuario']
    cantidad_items = len(grupo)

    # Actualizar el resumen
    resumen[correo] = {
        'Usuario': usuario,
        'Cantidad de items': cantidad_items
    }

# Mostrar resumen de correos
print(f"Se enviarán {len(resumen)} correos:")
for correo, info in resumen.items():
    print(f"- Correo: {correo}")
    print(f"  Usuario: {info['Usuario']}")
    print(f"  Cantidad de items: {info['Cantidad de items']}")

# Preguntar si se desea ver el resumen detallado
ver_resumen_detallado = input("¿Deseas ver el resumen detallado? (y/n): ")

if ver_resumen_detallado.lower() == 'y':
    for correo, info in resumen.items():
        usuario = info['Usuario']
        cantidad_items = info['Cantidad de items']

        datos_tabla = grupo[grupo['correo'] == correo][['VulnerabilitySeverityLevel', 'SoftwareName', 'RecommendedSecurityUpdate']]
        tabla = tabulate(datos_tabla, headers=['Urgencia', 'Aplicacion', 'Recomendacion'], tablefmt='orgtbl')

        print(f"\nDetalles para el correo: {correo}")
        print(f"Usuario: {usuario}")
        print(f"Cantidad de items: {cantidad_items}")
        print(f"Tabla de problemas:\n{tabla}\n")

# Confirmar el envío de los correos
confirmacion = input("¿Deseas enviar los correos? (y/n): ")

if confirmacion.lower() == 'y':
    # Enviar los correos
    servidor_smtp = 'smtp.gmail.com'  # Cambia esto según tu proveedor de correo
    puerto_smtp = 587  # Cambia esto según tu proveedor de correo
    remitente = 'neorisp@gmail.com'  # Cambia esto por tu dirección de correo electrónico
    password = 'hrwzhgivekxmibir'  # Cambia esto por tu contraseña de correo electrónico

    for correo, info in resumen.items():
        destinatario = correo

        # Construir el cuerpo del correo
        usuario = info['Usuario']
        cuerpo_correo = f"Estimado {usuario}, detectamos que su equipo tiene los siguientes problemas de updates:\n\n{tabla}"

        # Crear el objeto MIMEMultipart para el correo
        mensaje = MIMEMultipart()
        mensaje['From'] = remitente
        mensaje['To'] = destinatario
        mensaje['Subject'] = 'Problemas de updates en tu equipo'

        # Agregar el cuerpo del correo al objeto MIMEText
        mensaje.attach(MIMEText(cuerpo_correo, 'plain'))

        # Enviar el correo utilizando el servidor SMTP
        with smtplib.SMTP(servidor_smtp, puerto_smtp) as servidor:
            servidor.starttls()
            servidor.login(remitente, password)
            servidor.send_message(mensaje)

        print(f"Correo enviado a: {destinatario}")

else:
    print("No se enviaron los correos.")

