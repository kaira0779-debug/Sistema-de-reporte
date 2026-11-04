import sys
import os

# Detecta automáticamente la carpeta donde está este archivo (wsgi.py)
# Sin importar si está en C:\, D:\ o en el escritorio de otro usuario.
current_dir = os.path.dirname(os.path.abspath(__file__))

if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Importa tu aplicación Flask
from app import app as application