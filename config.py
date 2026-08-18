import os
from dotenv import load_dotenv

# Cargar variables de entorno desde el archivo .env
basedir = os.path.abspath(os.path.dirname(__file__))
load_dotenv(os.path.join(basedir, '.env'))

class Config:
    # Clave secreta de sesión
    SECRET_KEY = os.environ.get('SECRET_KEY', 'una_clave_secreta_muy_segura_2024!')
    
    # Configuración de base de datos MySQL / MariaDB / XAMPP
    MYSQL_HOST = os.environ.get('MYSQL_HOST', 'localhost')
    MYSQL_PORT = int(os.environ.get('MYSQL_PORT', 3306))
    MYSQL_USER = os.environ.get('MYSQL_USER', 'root')
    MYSQL_PASSWORD = os.environ.get('MYSQL_PASSWORD', 'Admin123*')
    MYSQL_DB = os.environ.get('MYSQL_DB', 'reportes_fallas')
    MYSQL_CHARSET = 'utf8mb4'
    MYSQL_CONNECT_TIMEOUT = 10
    MYSQL_CURSORCLASS = 'DictCursor'
    
    # Ruta al ejecutable mysqldump para respaldos automáticos
    MYSQLDUMP_PATH = os.environ.get('MYSQLDUMP_PATH', '')
    
    # Carpeta de almacenamiento de respaldos
    BACKUP_FOLDER = os.path.join(basedir, 'backups')