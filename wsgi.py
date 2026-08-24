import sys
import os

# Ajusta esta ruta a donde esté tu carpeta de proyecto
path = 'C:\Users\kaira\Desktop\sistema_reportes'   # o /home/tu_usuario/mysite
if path not in sys.path:
    sys.path.append(path)

# Importa tu aplicación Flask
from app import app as application