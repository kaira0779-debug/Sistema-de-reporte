import sys
import os

# Ruta donde estará tu proyecto en PythonAnywhere (ajústala)
path = '/home/tu_usuario/misistema'
if path not in sys.path:
    sys.path.append(path)

# Importar la aplicación Flask
from app import app as application

# Si usas variables de entorno, PythonAnywhere las cargará desde el panel,
# pero también puedes cargarlas aquí manualmente:
# os.environ['MYSQL_HOST'] = 'tu_host'
# os.environ['MYSQL_USER'] = 'tu_usuario'
# os.environ['MYSQL_PASSWORD'] = 'tu_contraseña'
# os.environ['MYSQL_DB'] = 'tu_base'