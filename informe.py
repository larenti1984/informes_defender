import pandas as pd

# Leer la planilla de Excel
df = pd.read_excel('C:/Proyectos/infDef_python/informes_defender/informe.xlsx')

# Obtener los valores de la columna "usuario" como cadenas de caracteres
usuarios = df['usuario'].astype(str)

# Construir los correos
correos = [usuario + '@neoris.com' for usuario in usuarios]

# Agregar la columna "correo" al inicio del DataFrame
df.insert(0, 'correo', correos)

# Guardar el DataFrame modificado en un nuevo archivo Excel
df.to_excel('C:/Proyectos/infDef_python/informes_defender/informe_modificado.xlsx', index=False)


