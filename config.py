import os

class Config:
    SECRET_KEY = os.environ.get('SECRET_KEY', 'una_clave_secreta_muy_segura_2024!')
    MYSQL_HOST = os.environ.get('MYSQL_HOST', 'localhost')
    MYSQL_USER = os.environ.get('MYSQL_USER', 'root')
    MYSQL_PASSWORD = os.environ.get('MYSQL_PASSWORD', 'Admin123*')
    MYSQL_DB = os.environ.get('MYSQL_DB', 'reportes_fallas')
    MYSQL_CURSORCLASS = 'DictCursor'

    # Ruta al ejecutable mysqldump (se detecta automáticamente si se deja vacío)
    # Si quieres forzar una ruta específica, escríbela aquí.
    MYSQLDUMP_PATH = os.environ.get('MYSQLDUMP_PATH', '')

    BACKUP_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'backups')