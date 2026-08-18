import os
import re
import csv
import random
import subprocess
import datetime
import atexit
import shutil
from datetime import datetime
from functools import lru_cache
from time import time
from flask import Flask, render_template, request, redirect, url_for, flash, session, g, jsonify, send_from_directory
from flask_login import LoginManager, login_user, logout_user, login_required, current_user, UserMixin
import bcrypt
import openpyxl
import pymysql
from config import Config
from apscheduler.schedulers.background import BackgroundScheduler
from reportlab.lib.pagesizes import letter, landscape
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.lib import colors
from reportlab.lib.units import inch
from io import BytesIO
import difflib
import tempfile
import chardet
import unicodedata

app = Flask(__name__)
app.config.from_object(Config)

pymysql.install_as_MySQLdb()

# ============================================================
#   CACHÉ SIMPLE
# ============================================================
CACHE_DURATION = 300
cache = {
    'nodos': {'data': None, 'timestamp': 0},
    'motivos': {'data': None, 'timestamp': 0},
    'tipos_falla': {'data': None, 'timestamp': 0},
    'tecnicos': {'data': None, 'timestamp': 0},
}

def get_cache(key, fetch_func):
    now = time()
    if cache[key]['data'] is not None and (now - cache[key]['timestamp']) < CACHE_DURATION:
        return cache[key]['data']
    data = fetch_func()
    cache[key]['data'] = data
    cache[key]['timestamp'] = now
    return data

def invalidar_cache():
    for key in cache:
        cache[key]['data'] = None
        cache[key]['timestamp'] = 0

# ============================================================
#   CONEXIÓN A LA BASE DE DATOS
# ============================================================
def get_db():
    if 'db' not in g:
        g.db = pymysql.connect(
            host=app.config['MYSQL_HOST'],
            port=app.config['MYSQL_PORT'],
            user=app.config['MYSQL_USER'],
            password=app.config['MYSQL_PASSWORD'],
            database=app.config['MYSQL_DB'],
            cursorclass=pymysql.cursors.DictCursor
        )
    return g.db

@app.teardown_appcontext
def close_db(error):
    db = g.pop('db', None)
    if db is not None:
        db.close()

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = 'login'
login_manager.login_message = 'Por favor inicia sesión para acceder.'

class User(UserMixin):
    def __init__(self, id, username, nombre_completo, rol, online):
        self.id = id
        self.username = username
        self.nombre_completo = nombre_completo
        self.rol = rol
        self.online = online

@login_manager.user_loader
def load_user(user_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id, username, nombre_completo, rol, online FROM trabajadores WHERE id = %s", (user_id,))
    user = cur.fetchone()
    cur.close()
    if user:
        return User(user['id'], user['username'], user['nombre_completo'], user['rol'], user['online'])
    return None

@app.before_request
def update_online():
    if current_user.is_authenticated:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("UPDATE trabajadores SET online = 1, ultima_actividad = NOW() WHERE id = %s", (current_user.id,))
        conn.commit()
        cur.close()

@app.context_processor
def inject_user():
    return dict(current_user=current_user, now=datetime.now)

# ============================================================
#   CREAR ADMIN Y COLUMNAS SI NO EXISTEN
# ============================================================
def crear_admin_si_no_existe():
    conn = get_db()
    cur = conn.cursor()

    columnas_requeridas_trab = {
        'telefono': "ALTER TABLE trabajadores ADD COLUMN telefono VARCHAR(20) NULL AFTER cedula",
        'ultima_actividad': "ALTER TABLE trabajadores ADD COLUMN ultima_actividad DATETIME NULL AFTER online"
    }
    for col, ddl in columnas_requeridas_trab.items():
        cur.execute(f"SHOW COLUMNS FROM trabajadores LIKE '{col}'")
        if not cur.fetchone():
            cur.execute(ddl)
            conn.commit()

    columnas_requeridas_clientes = {
        'nodo': "ALTER TABLE clientes ADD COLUMN nodo VARCHAR(100) NULL AFTER direccion_servicio",
        'mbps_contratados': "ALTER TABLE clientes ADD COLUMN mbps_contratados VARCHAR(50) NULL AFTER nodo"
    }
    for col, ddl in columnas_requeridas_clientes.items():
        cur.execute(f"SHOW COLUMNS FROM clientes LIKE '{col}'")
        if not cur.fetchone():
            try:
                cur.execute(ddl)
                conn.commit()
            except Exception as e:
                print(f"No se pudo agregar columna {col} a clientes: {e}")

    cur.execute("SHOW COLUMNS FROM tickets LIKE 'tecnico_id'")
    if not cur.fetchone():
        try:
            cur.execute("ALTER TABLE tickets ADD COLUMN tecnico_id INT NULL AFTER responsable")
            conn.commit()
        except Exception as e:
            print(f"No se pudo agregar columna tecnico_id a tickets: {e}")

    cur.execute("SELECT id FROM trabajadores WHERE username = 'admin'")
    if not cur.fetchone():
        hashed = bcrypt.hashpw(b'Admin123*', bcrypt.gensalt()).decode('utf-8')
        cur.execute(
            "INSERT INTO trabajadores (username, password_hash, nombre_completo, rol, online) VALUES (%s,%s,%s,%s,%s)",
            ('admin', hashed, 'Administrador del Sistema', 'admin', 0)
        )
        conn.commit()
        print("✅ Usuario administrador creado (admin / Admin123*)")
    cur.close()

with app.app_context():
    crear_admin_si_no_existe()

# ============================================================
#   GENERAR NÚMERO DE TICKET AUTOMÁTICO
# ============================================================
def generar_numero_ticket():
    conn = get_db()
    cur = conn.cursor()
    year = datetime.now().strftime('%Y')
    cur.execute("""
        CREATE TABLE IF NOT EXISTS ticket_sequences (
            year INT PRIMARY KEY,
            next_num INT NOT NULL DEFAULT 5000
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """)
    conn.commit()
    cur.execute("""
        INSERT INTO ticket_sequences (year, next_num) VALUES (%s, 5000)
        ON DUPLICATE KEY UPDATE next_num = LAST_INSERT_ID(next_num + 1)
    """, (int(year),))
    conn.commit()
    cur.execute("SELECT LAST_INSERT_ID() as next_val")
    row = cur.fetchone()
    next_num = row['next_val'] if row else 5000
    cur.close()
    return f"{year}-{next_num}"

# ============================================================
#   GENERAR NÚMERO CORRELATIVO
# ============================================================
def generar_numero_correlativo(tipo):
    conn = get_db()
    cur = conn.cursor()
    anio = datetime.now().year

    cur.execute("SHOW COLUMNS FROM correlativos LIKE 'año'")
    tiene_columna_tilde = cur.fetchone() is not None

    if tiene_columna_tilde:
        cur.execute("DROP TABLE IF EXISTS correlativos")
        conn.commit()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS correlativos (
            id INT AUTO_INCREMENT PRIMARY KEY,
            tipo VARCHAR(20) NOT NULL,
            anio INT NOT NULL DEFAULT 0,
            ultimo_numero INT NOT NULL DEFAULT 0,
            UNIQUE KEY unique_tipo_anio (tipo, anio)
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4
    """)
    conn.commit()

    cur.execute("""
        INSERT INTO correlativos (tipo, anio, ultimo_numero) VALUES (%s, %s, 0)
        ON DUPLICATE KEY UPDATE ultimo_numero = LAST_INSERT_ID(ultimo_numero + 1)
    """, (tipo, anio))
    conn.commit()
    cur.execute("SELECT LAST_INSERT_ID() as next_val")
    row = cur.fetchone()
    nuevo = row['next_val'] if row else 1
    cur.close()
    return nuevo

# ============================================================
#   FUNCIONES DE CLIENTES
# ============================================================
def guardar_cliente(datos):
    conn = get_db()
    cur = conn.cursor()
    n_contrato = datos.get('n_contrato', '').strip()
    n_documento = datos.get('n_documento', '').strip()
    nombre = datos.get('nombre_apellido', '').strip()

    if not n_contrato:
        if n_documento and nombre:
            cur.execute("SELECT id FROM clientes WHERE n_documento = %s AND nombre_apellido = %s LIMIT 1", (n_documento, nombre))
            cliente = cur.fetchone()
        else:
            cliente = None

        if cliente:
            cur.execute("""
                UPDATE clientes SET n_documento=%s, nombre_apellido=%s, telefono=%s, direccion_servicio=%s,
                nodo=%s, mbps_contratados=%s, estatus_usuario=%s
                WHERE id=%s
            """, (n_documento, nombre,
                  datos.get('telefono','').strip(), datos.get('direccion_servicio','').strip(),
                  datos.get('nodo','').strip(), datos.get('mbps_contratados','').strip(),
                  datos.get('estatus_usuario','').strip(), cliente['id']))
            conn.commit()
            cur.close()
            invalidar_cache()
            return 'updated'
        else:
            cur.execute("""
                INSERT INTO clientes (n_contrato, n_documento, nombre_apellido, telefono, direccion_servicio, nodo, mbps_contratados, estatus_usuario)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, ('', n_documento, nombre,
                  datos.get('telefono','').strip(), datos.get('direccion_servicio','').strip(),
                  datos.get('nodo','').strip(), datos.get('mbps_contratados','').strip(),
                  datos.get('estatus_usuario','').strip()))
            conn.commit()
            cur.close()
            invalidar_cache()
            return 'inserted'
    else:
        cur.execute("SELECT id FROM clientes WHERE n_contrato = %s", (n_contrato,))
        cliente = cur.fetchone()
        if cliente:
            cur.execute("""
                UPDATE clientes SET n_documento=%s, nombre_apellido=%s, telefono=%s, direccion_servicio=%s,
                nodo=%s, mbps_contratados=%s, estatus_usuario=%s
                WHERE id=%s
            """, (n_documento, nombre,
                  datos.get('telefono','').strip(), datos.get('direccion_servicio','').strip(),
                  datos.get('nodo','').strip(), datos.get('mbps_contratados','').strip(),
                  datos.get('estatus_usuario','').strip(), cliente['id']))
            conn.commit()
            cur.close()
            invalidar_cache()
            return 'updated'
        else:
            cur.execute("""
                INSERT INTO clientes (n_contrato, n_documento, nombre_apellido, telefono, direccion_servicio, nodo, mbps_contratados, estatus_usuario)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (n_contrato, n_documento, nombre,
                  datos.get('telefono','').strip(), datos.get('direccion_servicio','').strip(),
                  datos.get('nodo','').strip(), datos.get('mbps_contratados','').strip(),
                  datos.get('estatus_usuario','').strip()))
            conn.commit()
            cur.close()
            invalidar_cache()
            return 'inserted'

# ============================================================
#   LECTURA Y MAPEO DE ARCHIVOS
# ============================================================
CAMPOS_CLIENTES = [
    'n_contrato', 'n_documento', 'nombre_apellido', 'telefono',
    'direccion_servicio', 'nodo', 'mbps_contratados', 'estatus_usuario'
]

CAMPOS_TICKETS = [
    'fecha', 'num_ticket', 'nodo', 'cliente', 'motivo_ticket',
    'visita_tecnica', 'tipo_falla', 'departamento', 'status', 'responsable', 'observaciones'
]

SINONIMOS = {
    'n_contrato': ['contrato', 'n° contrato', 'nro contrato', 'número contrato', 'contrato cliente', 'id contrato', 'n_contrato', 'contrato', 'nro. contrato', 'nº contrato'],
    'n_documento': ['documento', 'cédula', 'rif', 'cedula', 'n° documento', 'nro documento', 'identificación', 'n_documento', 'documento', 'cedula', 'rif'],
    'nombre_apellido': ['nombre', 'cliente', 'nombre cliente', 'nombre y apellido', 'razón social', 'empresa', 'persona', 'nombre_apellido', 'nombre', 'cliente'],
    'telefono': ['teléfono', 'celular', 'movil', 'contacto', 'tel', 'phone', 'telefono', 'tel'],
    'direccion_servicio': ['dirección', 'direccion', 'servicio', 'direccion de servicio', 'ubicación', 'domicilio', 'direccion_prestacion_servicio', 'dirección de servicio'],
    'nodo': ['zona', 'sector', 'nodo', 'torre', 'zona de cobertura'],
    'mbps_contratados': ['mbps', 'velocidad', 'plan', 'mb', 'ancho de banda', 'contratado', 'mbps', 'velocidad contratada'],
    'estatus_usuario': ['estatus', 'estado', 'condición', 'situación', 'status', 'estatus'],
    'fecha': ['fecha', 'dia', 'date', 'fecha de ticket'],
    'num_ticket': ['ticket', 'nº ticket', 'numero ticket', 'nro ticket', 'id ticket', 'número de ticket', 'num_ticket', 'nş de ticket', 'n° ticket', 'nro. ticket', 'n de ticket'],
    'cliente': ['cliente', 'nombre cliente', 'usuario', 'afiliado', 'contratante', 'nombre'],
    'motivo_ticket': ['motivo', 'razón', 'descripción', 'falla reportada', 'problema', 'motivo de ticket'],
    'visita_tecnica': ['visita', 'técnico', 'visita tecnica', 'requiere visita', 'desplazamiento'],
    'tipo_falla': ['falla', 'tipo falla', 'problema', 'categoría', 'clasificación'],
    'departamento': ['área', 'departamento', 'unidad', 'gerencia'],
    'status': ['estado', 'estatus', 'situación', 'condición'],
    'responsable': ['técnico', 'responsable', 'asignado a', 'encargado', 'personal'],
    'observaciones': ['observación', 'notas', 'comentarios', 'detalles', 'adicional']
}

def normalizar_texto(texto):
    texto = texto.lower()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    texto = re.sub(r'[^a-z0-9 ]', ' ', texto)
    texto = re.sub(r'\s+', ' ', texto).strip()
    return texto

def detectar_delimitador(archivo):
    with open(archivo, 'rb') as f:
        raw = f.read(2048)
        encoding = chardet.detect(raw)['encoding'] or 'utf-8-sig'
    with open(archivo, 'r', encoding=encoding) as f:
        sample = f.read(2048)
    delimitadores = [',', ';', '\t', '|']
    first_line = sample.split('\n')[0] if '\n' in sample else sample
    counts = {}
    for delim in delimitadores:
        counts[delim] = first_line.count(delim)
    if ';' in first_line and counts.get(';', 0) > counts.get(',', 0):
        return ';'
    if counts.get(',', 0) > 0:
        return ','
    if '\t' in first_line:
        return '\t'
    total_counts = {}
    for delim in delimitadores:
        total_counts[delim] = sample.count(delim)
    if total_counts:
        return max(total_counts, key=total_counts.get)
    return ','

def leer_archivo_datos(filepath, extension):
    if extension in ['.xlsx', '.xls']:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        try:
            hoja = wb.active
            filas = list(hoja.iter_rows(values_only=True))
            datos = [[str(cell).strip() if cell is not None else '' for cell in row] for row in filas if any(cell for cell in row)]
            return datos
        finally:
            wb.close()
    else:
        with open(filepath, 'rb') as f:
            raw = f.read(10000)
            encoding = chardet.detect(raw)['encoding'] or 'utf-8-sig'
        with open(filepath, 'r', encoding=encoding) as f:
            sample = f.read(2048)
            f.seek(0)
            delim = detectar_delimitador(filepath)
            reader = csv.reader(f, delimiter=delim)
            filas = list(reader)
            if len(filas) < 2:
                f.seek(0)
                for d in [',', ';', '\t', '|']:
                    if d != delim:
                        f.seek(0)
                        reader = csv.reader(f, delimiter=d)
                        filas = list(reader)
                        if len(filas) >= 2:
                            break
            return [[cell.strip() for cell in row] for row in filas if any(cell for cell in row)]

def mapear_columnas(filas, campos_esperados, sinonimos):
    if not filas:
        return None, None
    header = [str(c).strip() for c in filas[0]]
    header_norm = [normalizar_texto(h) for h in header]
    mapeo = {}
    usado = set()
    for campo in campos_esperados:
        sinonimos_campo = sinonimos.get(campo, [])
        candidatos = [campo] + sinonimos_campo
        candidatos_norm = [normalizar_texto(c) for c in candidatos]
        mejor_idx = None
        mejor_score = 0
        for idx, col_norm in enumerate(header_norm):
            if idx in usado:
                continue
            for cand in candidatos_norm:
                if col_norm == cand:
                    score = 1.0
                else:
                    score = difflib.SequenceMatcher(None, col_norm, cand).ratio()
                if cand in col_norm:
                    score += 0.2
                if score > mejor_score:
                    mejor_score = score
                    mejor_idx = idx
        if mejor_idx is not None and mejor_score >= 0.6:
            mapeo[campo] = mejor_idx
            usado.add(mejor_idx)
    campos_clientes_mapped = sum(1 for c in CAMPOS_CLIENTES if c in mapeo)
    campos_tickets_mapped = sum(1 for c in CAMPOS_TICKETS if c in mapeo)
    if campos_clientes_mapped >= campos_tickets_mapped:
        tipo = 'clientes'
    else:
        tipo = 'tickets'
    return mapeo, tipo

def convertir_fecha(valor):
    if not valor:
        return None
    if isinstance(valor, datetime):
        return valor.strftime('%Y-%m-%d')
    if isinstance(valor, str):
        valor = valor.strip()
        if ' ' in valor:
            valor = valor.split(' ')[0]
        formatos = [
            '%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%m/%d/%Y',
            '%d.%m.%Y', '%Y/%m/%d', '%b %d %Y', '%d %b %Y'
        ]
        for fmt in formatos:
            try:
                dt = datetime.strptime(valor, fmt)
                return dt.strftime('%Y-%m-%d')
            except ValueError:
                continue
        return None
    return None

def procesar_fila_datos(fila, mapeo, tipo):
    datos = {}
    for campo, idx in mapeo.items():
        if idx < len(fila):
            val = fila[idx].strip() if isinstance(fila[idx], str) else fila[idx]
        else:
            val = ''
        if campo == 'fecha':
            val = convertir_fecha(val)
        elif campo in ['n_contrato', 'n_documento', 'telefono', 'mbps_contratados']:
            if isinstance(val, str):
                val = re.sub(r'[^\w\-\.\+]', '', val)
        datos[campo] = val
    return datos

# ============================================================
#                           RUTAS
# ============================================================
@app.route('/')
def index():
    if current_user.is_authenticated:
        if current_user.rol == 'admin':
            return redirect(url_for('dashboard'))
        else:
            return redirect(url_for('panel_tecnico'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if current_user.is_authenticated:
        if current_user.rol == 'admin':
            return redirect(url_for('dashboard'))
        else:
            return redirect(url_for('panel_tecnico'))
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password'].encode('utf-8')
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT id, username, password_hash, nombre_completo, rol FROM trabajadores WHERE username = %s", (username,))
        user = cur.fetchone()
        cur.close()
        if user and bcrypt.checkpw(password, user['password_hash'].encode('utf-8')):
            user_obj = User(user['id'], user['username'], user['nombre_completo'], user['rol'], True)
            login_user(user_obj)
            if user['rol'] == 'admin':
                return redirect(url_for('dashboard'))
            else:
                return redirect(url_for('panel_tecnico'))
        else:
            flash('Credenciales inválidas', 'danger')
    return render_template('login.html')

@app.route('/logout')
@login_required
def logout():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE trabajadores SET online = 0, ultima_actividad = NULL WHERE id = %s", (current_user.id,))
    conn.commit()
    cur.close()
    logout_user()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    if current_user.rol != 'admin':
        return redirect(url_for('panel_tecnico'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) as total FROM trabajadores WHERE rol = 'tecnico'")
    total_tecnicos = cur.fetchone()['total']
    cur.execute("SELECT COUNT(*) as total FROM clientes")
    total_clientes = cur.fetchone()['total']
    cur.execute("SELECT COUNT(*) as total FROM tickets")
    total_tickets = cur.fetchone()['total']
    cur.close()
    return render_template('dashboard.html', total_tecnicos=total_tecnicos, total_clientes=total_clientes, total_tickets=total_tickets)

# ============================================================
#   PANEL DEL TÉCNICO
# ============================================================
@app.route('/panel_tecnico')
@login_required
def panel_tecnico():
    if current_user.rol != 'tecnico':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT t.*, tec.nombre_completo as tecnico_nombre
        FROM tickets t
        LEFT JOIN trabajadores tec ON t.tecnico_id = tec.id
        WHERE t.tecnico_id = %s
        ORDER BY t.fecha DESC, t.num_ticket DESC
    """, (current_user.id,))
    tickets = cur.fetchall()
    cur.close()
    return render_template('panel_tecnico.html', tickets=tickets)

# ---------- CLIENTES ----------
@app.route('/clientes')
@login_required
def clientes():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM clientes ORDER BY creado_en DESC LIMIT 500")
    clientes = cur.fetchall()
    cur.execute("SELECT DISTINCT nodo FROM clientes WHERE nodo IS NOT NULL AND nodo != '' ORDER BY nodo")
    nodos = [n['nodo'] for n in cur.fetchall()]
    cur.close()
    return render_template('clientes.html', clientes=clientes, nodos=nodos)

@app.route('/preview_import', methods=['POST'])
@login_required
def preview_import():
    if current_user.rol != 'admin':
        return jsonify({'error': 'Acceso denegado'}), 403
    archivo = request.files['archivo']
    if not archivo:
        return jsonify({'error': 'No se seleccionó archivo'}), 400
    ext = os.path.splitext(archivo.filename)[1].lower()
    if ext not in ['.xlsx', '.xls', '.csv']:
        return jsonify({'error': 'Formato no soportado'}), 400
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    archivo.save(temp.name)
    temp.close()
    try:
        filas = leer_archivo_datos(temp.name, ext)
        try:
            os.unlink(temp.name)
        except Exception:
            pass
        if not filas:
            return jsonify({'error': 'El archivo está vacío'}), 400
        mapeo, tipo = mapear_columnas(filas, CAMPOS_CLIENTES + CAMPOS_TICKETS, SINONIMOS)
        if not mapeo:
            return jsonify({'error': 'No se pudieron reconocer las columnas'}), 400
        preview = []
        for fila in filas[1:6]:
            datos = procesar_fila_datos(fila, mapeo, tipo)
            preview.append(datos)
        return jsonify({
            'tipo': tipo,
            'mapeo': mapeo,
            'preview': preview,
            'total_filas': len(filas) - 1
        })
    except Exception as e:
        try:
            os.unlink(temp.name)
        except Exception:
            pass
        return jsonify({'error': str(e)}), 500

@app.route('/cargar_clientes', methods=['GET', 'POST'])
@login_required
def cargar_clientes():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    if request.method == 'POST':
        archivo = request.files['archivo']
        if not archivo:
            flash('No se seleccionó archivo', 'danger')
            return redirect(url_for('clientes'))
        ext = os.path.splitext(archivo.filename)[1].lower()
        if ext not in ['.xlsx', '.xls', '.csv']:
            flash('Solo se permiten archivos .xlsx, .xls y .csv', 'danger')
            return redirect(url_for('clientes'))
        temp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        archivo.save(temp.name)
        temp.close()
        try:
            filas = leer_archivo_datos(temp.name, ext)
            try:
                os.unlink(temp.name)
            except Exception as e:
                print(f"Advertencia: no se pudo eliminar {temp.name}: {e}")
            if not filas:
                flash('El archivo está vacío', 'danger')
                return redirect(url_for('clientes'))
            mapeo, tipo = mapear_columnas(filas, CAMPOS_CLIENTES, SINONIMOS)
            if tipo != 'clientes' or not mapeo:
                flash('No se pudieron identificar las columnas como clientes. Verifica el archivo.', 'danger')
                return redirect(url_for('clientes'))
            conn = get_db()
            cur = conn.cursor()
            insertados = 0
            actualizados = 0
            errores = 0
            for fila in filas[1:]:
                if not any(fila):
                    continue
                datos = procesar_fila_datos(fila, mapeo, 'clientes')
                if not datos.get('nombre_apellido'):
                    errores += 1
                    continue
                resultado = guardar_cliente(datos)
                if resultado == 'inserted':
                    insertados += 1
                elif resultado == 'updated':
                    actualizados += 1
                else:
                    errores += 1
            cur.close()
            flash(f'Clientes importados: {insertados} nuevos, {actualizados} actualizados, {errores} errores.', 'success')
        except Exception as e:
            try:
                os.unlink(temp.name)
            except Exception:
                pass
            flash(f'Error al procesar el archivo: {str(e)}', 'danger')
        return redirect(url_for('clientes'))
    return redirect(url_for('clientes'))

@app.route('/agregar_cliente', methods=['POST'])
@login_required
def agregar_cliente():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    datos = {
        'n_contrato': request.form.get('n_contrato', '').strip(),
        'n_documento': request.form.get('n_documento', '').strip(),
        'nombre_apellido': request.form.get('nombre_apellido', '').strip(),
        'telefono': request.form.get('telefono', '').strip(),
        'direccion_servicio': request.form.get('direccion_servicio', '').strip(),
        'nodo': request.form.get('nodo', '').strip(),
        'mbps_contratados': request.form.get('mbps_contratados', '').strip(),
        'estatus_usuario': request.form.get('estatus_usuario', '').strip()
    }
    if not datos['nombre_apellido']:
        flash('El nombre del cliente es obligatorio', 'danger')
        return redirect(url_for('clientes'))
    resultado = guardar_cliente(datos)
    if resultado == 'inserted': flash('Cliente agregado correctamente', 'success')
    elif resultado == 'updated': flash('Cliente actualizado (existía con ese número de contrato)', 'info')
    else: flash('Error al guardar el cliente', 'danger')
    return redirect(url_for('clientes'))

@app.route('/cliente/<int:cliente_id>/editar', methods=['GET', 'POST'])
@login_required
def editar_cliente(cliente_id):
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM clientes WHERE id = %s", (cliente_id,))
    cliente = cur.fetchone()
    if not cliente:
        flash('Cliente no encontrado', 'danger')
        return redirect(url_for('clientes'))
    if request.method == 'POST':
        n_contrato = request.form.get('n_contrato', '').strip()
        n_documento = request.form.get('n_documento', '').strip()
        nombre_apellido = request.form.get('nombre_apellido', '').strip()
        telefono = request.form.get('telefono', '').strip()
        direccion_servicio = request.form.get('direccion_servicio', '').strip()
        nodo = request.form.get('nodo', '').strip()
        mbps_contratados = request.form.get('mbps_contratados', '').strip()
        estatus_usuario = request.form.get('estatus_usuario', '').strip()
        if not nombre_apellido:
            flash('El nombre del cliente es obligatorio', 'danger')
            return render_template('editar_cliente.html', cliente=cliente)
        cur.execute("""
            UPDATE clientes SET n_contrato=%s, n_documento=%s, nombre_apellido=%s, telefono=%s,
            direccion_servicio=%s, nodo=%s, mbps_contratados=%s, estatus_usuario=%s
            WHERE id=%s
        """, (n_contrato, n_documento, nombre_apellido, telefono, direccion_servicio, nodo, mbps_contratados, estatus_usuario, cliente_id))
        conn.commit()
        cur.close()
        invalidar_cache()
        flash('Cliente actualizado correctamente', 'success')
        return redirect(url_for('detalle_cliente', cliente_id=cliente_id))
    cur.execute("SELECT DISTINCT nodo FROM clientes WHERE nodo IS NOT NULL AND nodo != '' ORDER BY nodo")
    nodos = [row['nodo'] for row in cur.fetchall()]
    cur.close()
    return render_template('editar_cliente.html', cliente=cliente, nodos=nodos)

@app.route('/cliente/<int:cliente_id>/eliminar', methods=['POST'])
@login_required
def eliminar_cliente(cliente_id):
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM clientes WHERE id = %s", (cliente_id,))
    conn.commit()
    cur.close()
    invalidar_cache()
    flash('Cliente eliminado correctamente', 'success')
    return redirect(url_for('clientes'))

@app.route('/cliente/<int:cliente_id>')
@login_required
def detalle_cliente(cliente_id):
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM clientes WHERE id = %s", (cliente_id,))
    cliente = cur.fetchone()
    cur.close()
    if not cliente:
        flash('Cliente no encontrado', 'danger')
        return redirect(url_for('clientes'))
    return render_template('detalle_cliente.html', cliente=cliente)

# ---------- TICKETS ----------
@app.route('/tickets')
@login_required
def tickets():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT t.*, tec.nombre_completo as tecnico_nombre
        FROM tickets t
        LEFT JOIN trabajadores tec ON t.tecnico_id = tec.id
        ORDER BY t.fecha DESC, t.num_ticket DESC
        LIMIT 500
    """)
    tickets = cur.fetchall()

    cur.execute("SELECT DISTINCT nodo FROM clientes WHERE nodo IS NOT NULL AND nodo != '' UNION SELECT DISTINCT nodo FROM tickets WHERE nodo IS NOT NULL AND nodo != '' ORDER BY nodo")
    nodos = [row['nodo'] for row in cur.fetchall()]

    cur.execute("SELECT DISTINCT motivo_ticket FROM tickets WHERE motivo_ticket IS NOT NULL AND motivo_ticket != '' ORDER BY motivo_ticket")
    motivos = [row['motivo_ticket'] for row in cur.fetchall()]

    cur.execute("SELECT DISTINCT tipo_falla FROM tickets WHERE tipo_falla IS NOT NULL AND tipo_falla != '' ORDER BY tipo_falla")
    tipos_falla = [row['tipo_falla'] for row in cur.fetchall()]

    cur.execute("SELECT id, nombre_completo, cedula, telefono FROM trabajadores WHERE rol = 'tecnico' ORDER BY nombre_completo")
    tecnicos = cur.fetchall()

    cur.close()
    return render_template('tickets.html', tickets=tickets, nodos=nodos, motivos=motivos, tipos_falla=tipos_falla, tecnicos=tecnicos)

@app.route('/api/buscar_tickets_json')
@login_required
def api_buscar_tickets():
    q = request.args.get('q', '').strip()
    conn = get_db()
    cur = conn.cursor()
    if q:
        like = f'%{q}%'
        cur.execute("""
            SELECT * FROM tickets 
            WHERE num_ticket LIKE %s OR cliente LIKE %s OR nodo LIKE %s 
               OR motivo_ticket LIKE %s OR tipo_falla LIKE %s OR status LIKE %s 
               OR responsable LIKE %s OR observaciones LIKE %s
            ORDER BY fecha DESC, num_ticket DESC 
            LIMIT 50
        """, (like, like, like, like, like, like, like, like))
    else:
        cur.execute("SELECT * FROM tickets ORDER BY fecha DESC, num_ticket DESC LIMIT 50")
    tickets = cur.fetchall()
    cur.close()
    for t in tickets:
        if t.get('fecha') and hasattr(t['fecha'], 'strftime'):
            t['fecha'] = t['fecha'].strftime('%Y-%m-%d')
    return jsonify(tickets)

@app.route('/api/buscar_cliente_json')
@login_required
def api_buscar_cliente():
    q = request.args.get('q', '').strip()
    if not q:
        return jsonify([])
    conn = get_db()
    cur = conn.cursor()
    q_upper = q.upper().strip()
    like = f'%{q_upper}%'
    cur.execute("""
        SELECT id, n_contrato, nombre_apellido, telefono, nodo, n_documento, direccion_servicio, mbps_contratados, estatus_usuario
        FROM clientes
        WHERE UPPER(n_contrato) LIKE %s 
           OR UPPER(nombre_apellido) LIKE %s 
           OR UPPER(n_documento) LIKE %s 
           OR UPPER(telefono) LIKE %s 
           OR UPPER(nodo) LIKE %s
           OR UPPER(direccion_servicio) LIKE %s
        ORDER BY 
            CASE 
                WHEN UPPER(nombre_apellido) = %s THEN 1
                WHEN UPPER(nombre_apellido) LIKE %s THEN 2
                WHEN UPPER(n_contrato) = %s THEN 3
                WHEN UPPER(n_documento) = %s THEN 4
                ELSE 5
            END
        LIMIT 20
    """, (like, like, like, like, like, like,
         q_upper, f'%{q_upper}%', q_upper, q_upper))
    clientes = []
    for c in cur.fetchall():
        clientes.append({
            'id': c['id'],
            'nombre': c['nombre_apellido'],
            'contrato': c['n_contrato'] or '',
            'documento': c['n_documento'] or '',
            'telefono': c['telefono'] or '',
            'direccion': c['direccion_servicio'] or '',
            'nodo': c['nodo'] or '',
            'mbps': c['mbps_contratados'] or '',
            'estatus': c['estatus_usuario'] or ''
        })
    cur.close()
    return jsonify(clientes)

@app.route('/cargar_tickets', methods=['POST'])
@login_required
def cargar_tickets():
    archivo = request.files['archivo']
    if not archivo:
        flash('No se seleccionó archivo', 'danger')
        return redirect(url_for('tickets'))
    ext = os.path.splitext(archivo.filename)[1].lower()
    if ext not in ['.xlsx', '.xls', '.csv']:
        flash('Solo se permiten archivos .xlsx, .xls y .csv', 'danger')
        return redirect(url_for('tickets'))
    temp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    archivo.save(temp.name)
    temp.close()
    try:
        filas = leer_archivo_datos(temp.name, ext)
        try:
            os.unlink(temp.name)
        except Exception as e:
            print(f"Advertencia: no se pudo eliminar {temp.name}: {e}")
        if not filas:
            flash('El archivo está vacío', 'danger')
            return redirect(url_for('tickets'))
        mapeo, tipo = mapear_columnas(filas, CAMPOS_TICKETS, SINONIMOS)
        if tipo != 'tickets' or not mapeo:
            flash('No se pudieron identificar las columnas como tickets. Verifica el archivo.', 'danger')
            return redirect(url_for('tickets'))
        conn = get_db()
        cur = conn.cursor()
        insertados = 0
        errores = 0
        for idx, fila in enumerate(filas[1:], start=2):
            if not any(fila):
                continue
            datos = procesar_fila_datos(fila, mapeo, 'tickets')
            if not datos.get('cliente') or not datos.get('fecha'):
                errores += 1
                print(f"Fila {idx}: faltan cliente o fecha")
                continue
            num_ticket = str(datos.get('num_ticket', '')).strip()
            if not num_ticket:
                errores += 1
                print(f"Fila {idx}: num_ticket vacío")
                continue
            cur.execute("SELECT id FROM tickets WHERE num_ticket = %s AND fecha = %s", (num_ticket, datos['fecha']))
            if cur.fetchone():
                errores += 1
                print(f"Fila {idx}: duplicado de {num_ticket} en {datos['fecha']}")
                continue
            if datos.get('tipo_falla') and 'ap caida' in datos['tipo_falla'].lower():
                datos['tipo_falla'] = 'AP caída'
            cur.execute("""
                INSERT INTO tickets (fecha, num_ticket, nodo, cliente, motivo_ticket, visita_tecnica, 
                                    tipo_falla, departamento, status, responsable, observaciones, tecnico_id)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,NULL)
            """, (datos.get('fecha'), num_ticket, datos.get('nodo', ''), datos.get('cliente'),
                  datos.get('motivo_ticket', ''), datos.get('visita_tecnica', ''),
                  datos.get('tipo_falla', ''), datos.get('departamento', ''),
                  datos.get('status', 'Pendiente'), datos.get('responsable', ''),
                  datos.get('observaciones', '')))
            insertados += 1
        conn.commit()
        cur.close()
        flash(f'{insertados} tickets importados correctamente, {errores} errores.', 'success')
    except Exception as e:
        try:
            os.unlink(temp.name)
        except Exception:
            pass
        flash(f'Error al procesar el archivo: {str(e)}', 'danger')
    return redirect(url_for('tickets'))

@app.route('/agregar_ticket', methods=['POST'])
@login_required
def agregar_ticket():
    nodo = request.form.get('nodo', '').strip()
    cliente_id = request.form.get('cliente_id', '').strip()
    motivo_ticket = request.form.get('motivo_ticket', '').strip()
    visita_tecnica = request.form.get('visita_tecnica', '').strip()
    tipo_falla = request.form.get('tipo_falla', '').strip()
    status = request.form.get('status', 'Pendiente').strip()
    responsable = request.form.get('responsable', '').strip()
    tecnico_id = request.form.get('tecnico_id', '').strip()
    observaciones = request.form.get('observaciones', '').strip()
    cliente_nombre_busqueda = request.form.get('cliente_nombre_busqueda', '').strip()

    conn = get_db()
    cur = conn.cursor()

    if not cliente_id and cliente_nombre_busqueda:
        cur.execute("SELECT id, nombre_apellido, nodo FROM clientes WHERE nombre_apellido = %s LIMIT 1", (cliente_nombre_busqueda,))
        cliente = cur.fetchone()
        if cliente:
            cliente_id = cliente['id']
            nodo = nodo or cliente['nodo']

    if not cliente_id:
        cur.close()
        flash('Debe seleccionar un cliente de la lista', 'danger')
        return redirect(url_for('tickets'))

    cur.execute("SELECT * FROM clientes WHERE id = %s", (cliente_id,))
    cliente = cur.fetchone()
    if not cliente:
        cur.close()
        flash('Cliente no encontrado', 'danger')
        return redirect(url_for('tickets'))

    if not nodo:
        nodo = cliente.get('nodo', '')

    cliente_nombre = cliente['nombre_apellido']
    num_ticket = generar_numero_ticket()
    fecha = datetime.now().strftime('%Y-%m-%d')
    if re.search(r'ap\s*ca[ií]da', tipo_falla, re.IGNORECASE):
        tipo_falla = 'AP caída'

    tecnico_id = int(tecnico_id) if tecnico_id else None

    cur.execute("""
        INSERT INTO tickets (fecha, num_ticket, nodo, cliente, motivo_ticket, visita_tecnica, 
                            tipo_falla, departamento, status, responsable, observaciones, tecnico_id)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    """, (fecha, num_ticket, nodo, cliente_nombre, motivo_ticket, visita_tecnica, 
          tipo_falla, '', status, responsable, observaciones, tecnico_id))
    conn.commit()
    ticket_id = cur.lastrowid
    cur.close()
    flash(f'Ticket {num_ticket} agregado correctamente', 'success')
    return redirect(url_for('planilla_soporte', ticket_id=ticket_id))

@app.route('/ticket/<int:ticket_id>/eliminar', methods=['POST'])
@login_required
def eliminar_ticket(ticket_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM tickets WHERE id = %s", (ticket_id,))
    conn.commit()
    cur.close()
    flash('Ticket eliminado correctamente', 'success')
    return redirect(url_for('tickets'))

@app.route('/ticket/<int:ticket_id>')
@login_required
def detalle_ticket(ticket_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT t.*, tec.nombre_completo as tecnico_nombre, tec.cedula as tecnico_cedula, tec.telefono as tecnico_telefono
        FROM tickets t
        LEFT JOIN trabajadores tec ON t.tecnico_id = tec.id
        WHERE t.id = %s
    """, (ticket_id,))
    ticket = cur.fetchone()
    cur.close()
    if not ticket:
        flash('Ticket no encontrado', 'danger')
        return redirect(url_for('tickets'))
    return render_template('detalle_ticket.html', ticket=ticket)

@app.route('/ticket/<int:ticket_id>/editar', methods=['GET', 'POST'])
@login_required
def editar_ticket(ticket_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM tickets WHERE id = %s", (ticket_id,))
    ticket = cur.fetchone()
    if not ticket:
        flash('Ticket no encontrado', 'danger')
        cur.close()
        return redirect(url_for('tickets'))

    if request.method == 'POST':
        nodo = request.form.get('nodo', '').strip()
        motivo_ticket = request.form.get('motivo_ticket', '').strip()
        visita_tecnica = request.form.get('visita_tecnica', '').strip()
        tipo_falla = request.form.get('tipo_falla', '').strip()
        departamento = request.form.get('departamento', '').strip()
        status = request.form.get('status', '').strip()
        responsable = request.form.get('responsable', '').strip()
        tecnico_id = request.form.get('tecnico_id', '').strip()
        observaciones = request.form.get('observaciones', '').strip()

        if re.search(r'ap\s*ca[ií]da', tipo_falla, re.IGNORECASE):
            tipo_falla = 'AP caída'
        tecnico_id = int(tecnico_id) if tecnico_id else None

        cur.execute("""
            UPDATE tickets SET nodo=%s, motivo_ticket=%s, visita_tecnica=%s, tipo_falla=%s,
            departamento=%s, status=%s, responsable=%s, tecnico_id=%s, observaciones=%s
            WHERE id=%s
        """, (nodo, motivo_ticket, visita_tecnica, tipo_falla, departamento, status,
              responsable, tecnico_id, observaciones, ticket_id))
        conn.commit()
        cur.close()
        flash('Ticket actualizado correctamente', 'success')
        return redirect(url_for('detalle_ticket', ticket_id=ticket_id))

    cur.execute("SELECT DISTINCT nodo FROM clientes WHERE nodo IS NOT NULL AND nodo != '' UNION SELECT DISTINCT nodo FROM tickets WHERE nodo IS NOT NULL AND nodo != '' ORDER BY nodo")
    nodos = [row['nodo'] for row in cur.fetchall()]
    cur.execute("SELECT DISTINCT motivo_ticket FROM tickets WHERE motivo_ticket IS NOT NULL AND motivo_ticket != '' ORDER BY motivo_ticket")
    motivos = [row['motivo_ticket'] for row in cur.fetchall()]
    cur.execute("SELECT DISTINCT tipo_falla FROM tickets WHERE tipo_falla IS NOT NULL AND tipo_falla != '' ORDER BY tipo_falla")
    tipos_falla = [row['tipo_falla'] for row in cur.fetchall()]
    cur.execute("SELECT DISTINCT status FROM tickets WHERE status IS NOT NULL AND status != '' ORDER BY status")
    estados = [row['status'] for row in cur.fetchall()]
    cur.execute("SELECT id, nombre_completo FROM trabajadores WHERE rol = 'tecnico' ORDER BY nombre_completo")
    tecnicos = cur.fetchall()
    cur.close()
    return render_template('editar_ticket.html', ticket=ticket, nodos=nodos, motivos=motivos,
                           tipos_falla=tipos_falla, estados=estados, tecnicos=tecnicos)

# ---------- API TICKET DETALLE ----------
@app.route('/api/ticket/<int:ticket_id>')
@login_required
def api_ticket_detalle(ticket_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM tickets WHERE id = %s", (ticket_id,))
    ticket = cur.fetchone()
    if not ticket:
        cur.close()
        return jsonify({'error': 'Ticket no encontrado'}), 404
    cliente = None
    tecnico = None
    if ticket.get('cliente'):
        cur.execute("SELECT * FROM clientes WHERE nombre_apellido = %s LIMIT 1", (ticket['cliente'],))
        cliente = cur.fetchone()
    if ticket.get('tecnico_id'):
        cur.execute("SELECT id, username, nombre_completo, cedula, telefono FROM trabajadores WHERE id = %s AND rol = 'tecnico'", (ticket['tecnico_id'],))
        tecnico = cur.fetchone()
    cur.close()
    if ticket.get('fecha') and hasattr(ticket['fecha'], 'strftime'):
        ticket['fecha'] = ticket['fecha'].strftime('%Y-%m-%d')
    if cliente and cliente.get('creado_en') and hasattr(cliente['creado_en'], 'strftime'):
        cliente['creado_en'] = cliente['creado_en'].strftime('%Y-%m-%d %H:%M:%S')
    return jsonify({'ticket': ticket, 'cliente': cliente, 'tecnico': tecnico})

# ---------- ADMINISTRACIÓN ----------
@app.route('/admin')
@login_required
def admin_panel():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) as total FROM trabajadores WHERE rol = 'tecnico'")
    total_tecnicos = cur.fetchone()['total']
    cur.execute("SELECT COUNT(*) as total FROM clientes")
    total_clientes = cur.fetchone()['total']
    cur.execute("SELECT COUNT(*) as total FROM tickets")
    total_tickets = cur.fetchone()['total']
    cur.close()
    return render_template('admin_panel.html', total_tecnicos=total_tecnicos, total_clientes=total_clientes, total_tickets=total_tickets)

@app.route('/admin/tecnicos', methods=['GET', 'POST'])
@login_required
def admin_tecnicos():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        nombre_completo = request.form.get('nombre_completo', '').strip()
        cedula = request.form.get('cedula', '').strip()
        telefono = request.form.get('telefono', '').strip()
        if not username or not nombre_completo:
            flash('El usuario y nombre son obligatorios', 'danger')
            return redirect(url_for('admin_tecnicos'))
        try:
            cur.execute("SELECT id FROM trabajadores WHERE username = %s", (username,))
            if cur.fetchone():
                flash('El nombre de usuario ya existe', 'danger')
                return redirect(url_for('admin_tecnicos'))
            password_default = 'Tecnico123*'
            hashed = bcrypt.hashpw(password_default.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            cur.execute("""
                INSERT INTO trabajadores (username, password_hash, nombre_completo, cedula, telefono, rol, online)
                VALUES (%s,%s,%s,%s,%s,'tecnico',0)
            """, (username, hashed, nombre_completo, cedula, telefono))
            conn.commit()
            nuevo_id = cur.lastrowid
            flash(f'Técnico registrado exitosamente con ID {nuevo_id}. Contraseña temporal: {password_default}', 'success')
        except Exception as e:
            flash(f'Error: {str(e)}', 'danger')
        finally:
            cur.close()
        invalidar_cache()
        return redirect(url_for('admin_tecnicos'))
    cur.execute("SELECT id, username, nombre_completo, cedula, telefono, rol, online, ultima_actividad FROM trabajadores WHERE rol = 'tecnico' ORDER BY creado_en DESC")
    tecnicos = cur.fetchall()
    cur.close()
    return render_template('admin_tecnicos.html', tecnicos=tecnicos)

@app.route('/admin/tecnico/<int:id>/eliminar', methods=['POST'])
@login_required
def eliminar_tecnico(id):
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    if id == current_user.id:
        flash('No puedes eliminar tu propia cuenta', 'danger')
        return redirect(url_for('admin_tecnicos'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM trabajadores WHERE id = %s AND rol = 'tecnico'", (id,))
    conn.commit()
    cur.close()
    invalidar_cache()
    flash('Técnico eliminado', 'success')
    return redirect(url_for('admin_tecnicos'))

# ============================================================
#   DUPLICADOS
# ============================================================
@app.route('/admin/duplicados')
@login_required
def admin_duplicados():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT fecha, num_ticket, nodo, cliente, motivo_ticket, visita_tecnica,
               tipo_falla, departamento, status, responsable, observaciones,
               GROUP_CONCAT(id ORDER BY id DESC) as ids, COUNT(*) as cnt
        FROM tickets
        GROUP BY fecha, num_ticket, nodo, cliente, motivo_ticket, visita_tecnica,
                 tipo_falla, departamento, status, responsable, observaciones
        HAVING cnt > 1
        ORDER BY fecha DESC
    """)
    grupos = cur.fetchall()
    duplicados = []
    for g in grupos:
        ids = [int(x) for x in g['ids'].split(',')]
        cur.execute("SELECT * FROM tickets WHERE id IN (%s)" % ','.join(['%s']*len(ids)), ids)
        tickets = cur.fetchall()
        tickets.sort(key=lambda x: x['id'], reverse=True)
        duplicados.append({
            'fecha': g['fecha'],
            'cliente': g['cliente'],
            'motivo_ticket': g['motivo_ticket'],
            'tipo_falla': g['tipo_falla'],
            'tickets': tickets,
            'conservar': tickets[0]['id']
        })
    cur.close()
    return render_template('admin_duplicados.html', duplicados=duplicados)

@app.route('/admin/eliminar_duplicados_auto', methods=['POST'])
@login_required
def eliminar_duplicados_auto():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT fecha, num_ticket, nodo, cliente, motivo_ticket, visita_tecnica,
               tipo_falla, departamento, status, responsable, observaciones,
               GROUP_CONCAT(id ORDER BY id DESC) as ids
        FROM tickets
        GROUP BY fecha, num_ticket, nodo, cliente, motivo_ticket, visita_tecnica,
                 tipo_falla, departamento, status, responsable, observaciones
        HAVING COUNT(*) > 1
    """)
    grupos = cur.fetchall()
    eliminados = 0
    for g in grupos:
        ids = [int(x) for x in g['ids'].split(',')]
        for id_ticket in ids[1:]:
            cur.execute("DELETE FROM tickets WHERE id = %s", (id_ticket,))
            eliminados += 1
    conn.commit()
    cur.close()
    flash(f'Se eliminaron automáticamente {eliminados} tickets duplicados exactos.', 'success')
    return redirect(url_for('admin_duplicados'))

# ============================================================
#   PLANILLA DE SOPORTE
# ============================================================
GENERADOS_FOLDER = os.path.join(app.static_folder, 'plantillas', 'generados')
PDF_FOLDER = os.path.join(app.static_folder, 'plantillas', 'pdfs')
for folder in [GENERADOS_FOLDER, PDF_FOLDER]:
    if not os.path.exists(folder):
        os.makedirs(folder)

MOTIVOS_TICKET = [
    "Sin servicio", "Lentitud", "Intermitencia", "Instalación nueva",
    "Cambio de equipo", "Reubicación", "Mantenimiento", "Otro"
]
TIPOS_FALLA = [
    "Fibra cortada", "Equipo dañado", "Configuración incorrecta",
    "Sin electricidad", "Actualización de firmware", "Interferencia",
    "No aplica", "Otro"
]

def generar_pdf_planilla(cliente, tecnico, motivo, tipo_falla, observaciones, num_ticket, hora_inicio, hora_fin,
                         tipo_planilla='soporte', numero_correlativo=None):
    buffer = BytesIO()
    c = pdfcanvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    margin = 40
    x_left = margin
    x_right = width - margin
    y = height - margin

    # --- Encabezado ---
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(width/2, y, "DAITNVOZ C.A")
    c.setFont("Helvetica", 8)
    c.drawCentredString(width/2, y - 13, "Avenida Principal de los Chorros, Qta Datinvoz, Caracas 1071, Distrito Capital, Venezuela")
    c.drawCentredString(width/2, y - 24, "Teléfono: 0424-1637944")
    c.setFont("Helvetica-Bold", 10)
    if tipo_planilla.lower() == 'soporte':
        c.drawCentredString(width/2, y - 38, "ORDEN DE SERVICIO TECNICO - SOPORTE DE INTERNET")
    else:
        c.drawCentredString(width/2, y - 38, f"ORDEN DE SERVICIO TECNICO - {tipo_planilla.upper()}")

    # Número correlativo en esquina superior derecha
    c.setFont("Helvetica", 8)
    c.drawRightString(x_right, y - 12, f"N° Correlativo: {numero_correlativo if numero_correlativo else '________'}")

    # Línea separadora
    c.setStrokeColor(colors.black)
    c.line(x_left, y - 48, x_right, y - 48)

    # Bloque de datos del ticket/fecha/hora
    y = y - 62
    c.setFont("Helvetica", 8)
    if tipo_planilla.lower() == 'soporte':
        c.drawString(x_left, y, f"N° Ticket: {num_ticket if num_ticket else '______________'}")
        c.drawString(x_left + 200, y, f"Fecha: {datetime.now().strftime('%d/%m/%Y')}")
        c.drawString(x_left + 350, y, f"Hora Inicio: {hora_inicio if hora_inicio else '________'}")
        c.drawString(x_left + 500, y, f"Hora Fin: {hora_fin if hora_fin else '________'}")
    else:
        c.drawString(x_left, y, f"Fecha: {datetime.now().strftime('%d/%m/%Y')}")
        c.drawString(x_left + 200, y, f"Hora Inicio: {hora_inicio if hora_inicio else '________'}")
        c.drawString(x_left + 400, y, f"Hora Fin: {hora_fin if hora_fin else '________'}")

    y -= 22
    # --- Datos del cliente ---
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_left, y, "DATOS DEL CLIENTE")
    y -= 14
    c.setFont("Helvetica", 8)
    nombre_cliente = cliente.get('nombre_apellido', '') if cliente.get('nombre_apellido') else '__________________________'
    direccion = cliente.get('direccion_servicio', '') if cliente.get('direccion_servicio') else '__________________________'
    telefono = cliente.get('telefono', '') if cliente.get('telefono') else '____________________'
    c.drawString(x_left, y, f"Nombre/Empresa: {nombre_cliente}")
    y -= 12
    c.drawString(x_left, y, f"Dirección: {direccion}")
    y -= 12
    c.drawString(x_left, y, f"Teléfono: {telefono}")

    # --- Descripción del servicio ---
    y -= 24
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_left, y, "DESCRIPCION DEL SERVICIO REALIZADO:")
    y -= 12
    c.setFont("Helvetica", 8)
    for i in range(4):
        c.line(x_left, y - i*13, x_right, y - i*13)
    if observaciones:
        lines = observaciones.split('\n')
        for i, line in enumerate(lines[:4]):
            c.drawString(x_left + 4, y - i*13 - 3, line[:75])

    # --- Materiales ---
    y -= 62
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_left, y, "MATERIALES ENTREGADOS / INSTALADOS")
    y -= 12
    c.setFont("Helvetica", 7)
    c.drawString(x_left, y, "Cant.")
    c.drawString(x_left + 50, y, "Descripción")
    c.drawString(x_left + 200, y, "Serial/N° Lote")
    c.drawString(x_left + 350, y, "Observaciones")
    y -= 6
    c.line(x_left, y, x_right, y)
    y -= 10
    for i in range(4):
        c.drawString(x_left, y, "____")
        c.drawString(x_left + 50, y, "________________")
        c.drawString(x_left + 200, y, "_______________")
        c.drawString(x_left + 350, y, "________________")
        y -= 13

    # --- Datos del técnico ---
    y -= 18
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_left, y, "DATOS DEL TÉCNICO")
    y -= 14
    c.setFont("Helvetica", 8)
    if tecnico and tecnico.get('nombre_completo'):
        c.drawString(x_left, y, f"Nombre: {tecnico['nombre_completo']}")
    else:
        c.drawString(x_left, y, "Nombre: __________________________")
    if tecnico and tecnico.get('cedula'):
        c.drawString(x_left + 250, y, f"Cédula: {tecnico['cedula']}")
    else:
        c.drawString(x_left + 250, y, "Cédula: ________________")
    if tecnico and tecnico.get('telefono'):
        c.drawString(x_left + 430, y, f"Celular: {tecnico['telefono']}")
    else:
        c.drawString(x_left + 430, y, "Celular: ________________")
    y -= 12
    c.drawString(x_left, y, "Firma: __________________________")

    # --- Verificaciones y visita adicional ---
    y -= 22
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_left, y, "VERIFICACIONES")
    y -= 12
    c.setFont("Helvetica", 8)
    c.drawString(x_left, y, "[ ] Servicio funcional correctamente")
    c.drawString(x_left + 200, y, "[ ] Velocidad verificada")
    c.drawString(x_left + 400, y, "[ ] Cliente satisfecho")

    y -= 14
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_left, y, "REQUIERE VISITA ADICIONAL/INSPECCION")
    y -= 12
    c.setFont("Helvetica", 8)
    c.drawString(x_left, y, "[ ] Cableado dañado")
    c.drawString(x_left + 160, y, "[ ] Equipo defectuoso")
    c.drawString(x_left + 320, y, "[ ] Servicio no resuelto")
    c.drawString(x_left + 480, y, "[ ] Solo inspección")

    # --- Validación final ---
    y -= 24
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_left, y, "VALIDACION FINAL")
    y -= 14
    c.setFont("Helvetica", 8)
    c.drawString(x_left, y, "Cliente: ________________")
    c.drawString(x_left + 180, y, "Técnico: ________________")
    c.drawString(x_left + 360, y, "Revisado por: ________________")
    y -= 12
    c.drawString(x_left, y, "Firma: ________________")
    c.drawString(x_left + 180, y, "Firma: ________________")
    c.drawString(x_left + 360, y, "Firma: ________________")

    c.setFont("Helvetica-Oblique", 6)
    c.drawCentredString(width/2, 20, "Documento generado automáticamente - Sistema de Reportes de Fallas v1.0")

    c.save()
    buffer.seek(0)
    return buffer

# ============================================================
#   RUTAS DE PLANILLA
# ============================================================
@app.route('/planilla_soporte')
@login_required
def planilla_soporte():
    ticket_id = request.args.get('ticket_id', type=int)
    ticket = None
    cliente = None
    if ticket_id:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM tickets WHERE id = %s", (ticket_id,))
        ticket = cur.fetchone()
        if ticket:
            cur.execute("SELECT * FROM clientes WHERE nombre_apellido = %s LIMIT 1", (ticket['cliente'],))
            cliente = cur.fetchone()
        cur.close()
    generados = []
    if os.path.exists(GENERADOS_FOLDER):
        for f in os.listdir(GENERADOS_FOLDER):
            if f.endswith('.pdf'):
                partes = f.replace('.pdf', '').split('_')
                if len(partes) >= 3:
                    cliente_id = partes[1]
                    timestamp_str = partes[2]
                    fecha_mostrar = timestamp_str
                    try:
                        fecha_dt = datetime.strptime(timestamp_str, '%Y%m%d_%H%M%S')
                        fecha_mostrar = fecha_dt.strftime('%d/%m/%Y %H:%M')
                    except:
                        pass
                    conn = get_db()
                    cur = conn.cursor()
                    cur.execute("SELECT nombre_apellido FROM clientes WHERE id = %s", (cliente_id,))
                    row = cur.fetchone()
                    nombre_cliente = row['nombre_apellido'] if row else "Desconocido"
                    cur.close()
                    generados.append({
                        'archivo': f,
                        'cliente': nombre_cliente,
                        'fecha': fecha_mostrar,
                        'cliente_id': cliente_id
                    })
                else:
                    generados.append({'archivo': f, 'cliente': '?', 'fecha': '', 'cliente_id': ''})
    return render_template('planilla_soporte.html', 
                           generados=generados, 
                           motivos=MOTIVOS_TICKET, 
                           tipos_falla=TIPOS_FALLA,
                           ticket=ticket,
                           cliente=cliente)

@app.route('/planilla_soporte/generar_planilla', methods=['POST'])
@login_required
def generar_planilla_cliente():
    ticket_id = request.form.get('ticket_id', '').strip()
    cliente_id = request.form.get('cliente_id', '').strip()
    tipo_planilla = request.form.get('tipo_planilla', 'soporte').strip().lower()

    conn = get_db()
    cur = conn.cursor()

    ticket = None
    cliente = None
    tecnico = None

    if ticket_id:
        cur.execute("SELECT * FROM tickets WHERE id = %s", (ticket_id,))
        ticket = cur.fetchone()
        if not ticket:
            cur.close()
            flash('Ticket no encontrado', 'danger')
            return redirect(url_for('planilla_soporte'))

        # Obtener cliente asociado
        if cliente_id:
            cur.execute("SELECT * FROM clientes WHERE id = %s", (cliente_id,))
            cliente = cur.fetchone()
        elif ticket.get('cliente'):
            cur.execute("SELECT * FROM clientes WHERE nombre_apellido = %s LIMIT 1", (ticket['cliente'],))
            cliente = cur.fetchone()

        # Si no se encuentra cliente, creamos uno provisional con el nombre del ticket
        if not cliente:
            cliente = {
                'id': None,
                'nombre_apellido': ticket.get('cliente', ''),
                'direccion_servicio': '',
                'telefono': '',
                'nodo': '',
                'mbps_contratados': '',
                'estatus_usuario': '',
                'asignador': ''
            }

        # Obtener técnico por tecnico_id si existe
        if ticket.get('tecnico_id'):
            cur.execute("SELECT id, username, nombre_completo, cedula, telefono FROM trabajadores WHERE id = %s AND rol = 'tecnico'", (ticket['tecnico_id'],))
            tecnico = cur.fetchone()
    else:
        if not cliente_id:
            cur.close()
            flash('Debe seleccionar un ticket o cliente', 'danger')
            return redirect(url_for('planilla_soporte'))
        cur.execute("SELECT * FROM clientes WHERE id = %s", (cliente_id,))
        cliente = cur.fetchone()
        if not cliente:
            cur.close()
            flash('Cliente no encontrado', 'danger')
            return redirect(url_for('planilla_soporte'))
        num_ticket = generar_numero_ticket()
        ticket = {
            'num_ticket': num_ticket,
            'motivo_ticket': '',
            'tipo_falla': '',
            'observaciones': '',
            'cliente': cliente['nombre_apellido']
        }

    # Si no es soporte, no usar número de ticket
    if tipo_planilla != 'soporte':
        num_ticket = ''
    else:
        num_ticket = ticket.get('num_ticket', '') if ticket else ''

    motivo = ticket.get('motivo_ticket', '') if ticket else ''
    tipo_falla = ticket.get('tipo_falla', '') if ticket else ''
    observaciones = ticket.get('observaciones', '') if ticket else ''

    numero_correlativo = generar_numero_correlativo(tipo_planilla)

    try:
        pdf_buffer = generar_pdf_planilla(
            cliente, tecnico, motivo, tipo_falla, observaciones,
            num_ticket, '', '', tipo_planilla, numero_correlativo
        )
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        # Usar el id del cliente o 0 si es provisional
        cliente_id_final = cliente['id'] if cliente.get('id') else 0
        output_filename = f"planilla_{cliente_id_final}_{timestamp}.pdf"
        output_path = os.path.join(GENERADOS_FOLDER, output_filename)
        with open(output_path, 'wb') as f:
            f.write(pdf_buffer.read())
        flash(f'Planilla PDF generada correctamente', 'success')
        return redirect(url_for('preview_planilla', filename=output_filename))
    except Exception as e:
        flash(f'Error al generar planilla: {str(e)}', 'danger')
        return redirect(url_for('planilla_soporte'))

@app.route('/planilla_soporte/preview/<filename>')
@login_required
def preview_planilla(filename):
    filepath = os.path.join(GENERADOS_FOLDER, filename)
    if not os.path.exists(filepath):
        flash('El archivo no existe o fue eliminado.', 'danger')
        return redirect(url_for('planilla_soporte'))
    pdf_url = url_for('descargar_generado', filename=filename)
    return render_template('preview_planilla.html', filename=filename, pdf_url=pdf_url)

@app.route('/planilla_soporte/descargar_generado/<filename>')
@login_required
def descargar_generado(filename):
    return send_from_directory(GENERADOS_FOLDER, filename, as_attachment=False)

@app.route('/planilla_soporte/eliminar_generado/<filename>', methods=['POST'])
@login_required
def eliminar_generado(filename):
    filename = os.path.basename(filename)
    filepath = os.path.join(GENERADOS_FOLDER, filename)
    if os.path.exists(filepath):
        os.remove(filepath)
        flash('Planilla eliminada correctamente', 'success')
    else:
        flash('El archivo no existe', 'warning')
    return redirect(url_for('planilla_soporte'))

@app.route('/planilla_soporte/eliminar_seleccionadas', methods=['POST'])
@login_required
def eliminar_seleccionadas():
    archivos = request.form.getlist('archivos')
    if not archivos:
        flash('No se seleccionó ninguna planilla.', 'warning')
        return redirect(url_for('planilla_soporte'))
    eliminados = 0
    for filename in archivos:
        filename = os.path.basename(filename)
        filepath = os.path.join(GENERADOS_FOLDER, filename)
        if os.path.exists(filepath):
            os.remove(filepath)
            eliminados += 1
    flash(f'Se eliminaron {eliminados} planillas correctamente.', 'success')
    return redirect(url_for('planilla_soporte'))

# ============================================================
#   RESPALDOS
# ============================================================
@app.route('/admin/respaldos')
@login_required
def listar_respaldos():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    backup_folder = app.config['BACKUP_FOLDER']
    backups = []
    if os.path.exists(backup_folder):
        for f in os.listdir(backup_folder):
            if f.endswith('.sql'):
                ruta = os.path.join(backup_folder, f)
                tamaño = os.path.getsize(ruta)
                fecha_mod = datetime.fromtimestamp(os.path.getmtime(ruta))
                backups.append({
                    'nombre': f,
                    'tamaño': tamaño,
                    'fecha': fecha_mod.strftime('%d/%m/%Y %H:%M:%S'),
                    'tamaño_legible': f"{tamaño / 1024:.1f} KB" if tamaño < 1024*1024 else f"{tamaño / (1024*1024):.2f} MB"
                })
    backups.sort(key=lambda x: x['fecha'], reverse=True)
    return render_template('respaldos.html', backups=backups)

@app.route('/admin/respaldo/descargar/<filename>')
@login_required
def descargar_respaldo(filename):
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    return send_from_directory(app.config['BACKUP_FOLDER'], filename, as_attachment=True)

@app.route('/admin/respaldo/eliminar/<filename>', methods=['POST'])
@login_required
def eliminar_respaldo(filename):
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    ruta = os.path.join(app.config['BACKUP_FOLDER'], filename)
    if os.path.exists(ruta):
        os.remove(ruta)
        flash(f'Respaldo {filename} eliminado', 'success')
    else:
        flash('El archivo no existe', 'warning')
    return redirect(url_for('listar_respaldos'))

@app.route('/backup_manual', methods=['POST'])
@login_required
def backup_manual():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    filepath, error = crear_respaldo('manual')
    if error:
        flash(f'Error al generar respaldo: {error}', 'danger')
    else:
        flash(f'Respaldo generado correctamente: {os.path.basename(filepath)}', 'success')
    return redirect(url_for('listar_respaldos'))

def crear_respaldo(tipo='manual'):
    mysqldump_path = app.config.get('MYSQLDUMP_PATH')
    if not mysqldump_path:
        mysqldump_path = shutil.which('mysqldump')
    if not mysqldump_path:
        common_paths = [
            r'C:\Program Files\MariaDB 10.11\bin\mysqldump.exe',
            r'C:\Program Files\MariaDB 10.6\bin\mysqldump.exe',
            r'C:\Program Files\MariaDB 10.5\bin\mysqldump.exe',
            r'C:\Program Files (x86)\MariaDB 10.11\bin\mysqldump.exe',
            r'C:\Program Files\MySQL\MySQL Server 8.0\bin\mysqldump.exe',
            r'C:\xampp\mysql\bin\mysqldump.exe',
        ]
        for path in common_paths:
            if os.path.exists(path):
                mysqldump_path = path
                break
    filepath = None
    error = None
    if mysqldump_path and os.path.exists(mysqldump_path):
        try:
            os.makedirs(app.config['BACKUP_FOLDER'], exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            filename = f"backup_{timestamp}.sql"
            filepath = os.path.join(app.config['BACKUP_FOLDER'], filename)
            cmd = [
                mysqldump_path,
                f"--host={app.config['MYSQL_HOST']}",
                f"--user={app.config['MYSQL_USER']}",
                f"--password={app.config['MYSQL_PASSWORD']}",
                app.config['MYSQL_DB'],
                "--result-file=" + filepath,
                "--skip-extended-insert"
            ]
            subprocess.run(cmd, check=True, capture_output=True)
        except Exception as e:
            print(f"mysqldump falló: {e}, usando alternativa nativa...")
            filepath, error = generar_respaldo_sql()
    else:
        filepath, error = generar_respaldo_sql()
    return filepath, error

def generar_respaldo_sql():
    try:
        os.makedirs(app.config['BACKUP_FOLDER'], exist_ok=True)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"backup_{timestamp}.sql"
        filepath = os.path.join(app.config['BACKUP_FOLDER'], filename)
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(f"-- Respaldo generado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"-- Base de datos: {app.config['MYSQL_DB']}\n\n")
            f.write("SET FOREIGN_KEY_CHECKS=0;\n\n")
            conn = get_db()
            cur = conn.cursor()
            cur.execute("SHOW TABLES")
            tables = [list(row.values())[0] for row in cur.fetchall()]
            for table in tables:
                cur.execute(f"SHOW CREATE TABLE `{table}`")
                result = cur.fetchone()
                if result and 'Create Table' in result:
                    f.write(f"DROP TABLE IF EXISTS `{table}`;\n")
                    f.write(f"{result['Create Table']};\n\n")
                cur.execute(f"SELECT * FROM `{table}`")
                rows = cur.fetchall()
                if rows:
                    columns = [desc[0] for desc in cur.description]
                    f.write(f"-- Datos de la tabla: {table}\n")
                    for row in rows:
                        values = []
                        for val in row.values():
                            if val is None:
                                values.append('NULL')
                            elif isinstance(val, (int, float)):
                                values.append(str(val))
                            elif isinstance(val, datetime):
                                values.append(f"'{val.strftime('%Y-%m-%d %H:%M:%S')}'")
                            else:
                                escaped = str(val).replace("'", "''")
                                values.append(f"'{escaped}'")
                        f.write(f"INSERT INTO `{table}` (`{'`, `'.join(columns)}`) VALUES ({', '.join(values)});\n")
                    f.write("\n")
            f.write("SET FOREIGN_KEY_CHECKS=1;\n")
            cur.close()
        return filepath, None
    except Exception as e:
        return None, str(e)

# ============================================================
#   API CLIENTE DETALLE
# ============================================================
@app.route('/api/cliente/<int:cliente_id>')
@login_required
def api_cliente_detalle(cliente_id):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM clientes WHERE id = %s", (cliente_id,))
    cliente = cur.fetchone()
    cur.close()
    if not cliente:
        return jsonify({'error': 'Cliente no encontrado'}), 404
    return jsonify({
        'id': cliente['id'],
        'n_contrato': cliente['n_contrato'] or '',
        'n_documento': cliente['n_documento'] or '',
        'nombre_apellido': cliente['nombre_apellido'] or '',
        'nombre': cliente['nombre_apellido'] or '',
        'telefono': cliente['telefono'] or '',
        'direccion_servicio': cliente['direccion_servicio'] or '',
        'nodo': cliente['nodo'] or '',
        'mbps_contratados': cliente['mbps_contratados'] or '',
        'estatus_usuario': cliente['estatus_usuario'] or ''
    })

# ---------- API USUARIOS EN LÍNEA ----------
@app.route('/api/usuarios_online')
@login_required
def api_usuarios_online():
    if current_user.rol != 'admin':
        return jsonify({'error': 'Acceso denegado'}), 403
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, username, nombre_completo, rol, ultima_actividad 
        FROM trabajadores 
        WHERE online = 1 AND ultima_actividad > DATE_SUB(NOW(), INTERVAL 5 MINUTE)
    """)
    usuarios = cur.fetchall()
    cur.close()
    return jsonify(usuarios)

# ============================================================
#        GRÁFICAS
# ============================================================
@app.route('/graficas')
@login_required
def graficas():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    return render_template('graficas.html')

@app.route('/api/graficas/datos')
@login_required
def api_graficas_datos():
    if current_user.rol != 'admin':
        return jsonify({'error': 'Acceso denegado'}), 403
    filtro = request.args.get('filtro', 'tickets')
    conn = get_db()
    cur = conn.cursor()
    try:
        if filtro == 'tickets':
            cur.execute("""
                SELECT YEARWEEK(fecha, 1) as semana, COUNT(*) as total 
                FROM tickets 
                WHERE fecha >= DATE_SUB(CURDATE(), INTERVAL 8 WEEK)
                GROUP BY semana 
                ORDER BY semana ASC
            """)
            semanas_data = cur.fetchall()
            semanas = [f"Sem {s['semana'] % 100}" for s in semanas_data]
            totales_semana = [s['total'] for s in semanas_data]

            cur.execute("""
                SELECT DATE_FORMAT(fecha, '%Y-%m') as mes, COUNT(*) as total 
                FROM tickets 
                WHERE fecha >= DATE_SUB(CURDATE(), INTERVAL 12 MONTH)
                GROUP BY mes 
                ORDER BY mes ASC
            """)
            meses_data = cur.fetchall()
            meses = [datetime.strptime(m['mes'], '%Y-%m').strftime('%b %Y') for m in meses_data]
            totales_mes = [m['total'] for m in meses_data]

            cur.execute("""
                SELECT DATE_FORMAT(fecha, '%Y-%m-%d') as dia, COUNT(*) as total 
                FROM tickets 
                WHERE fecha >= DATE_SUB(CURDATE(), INTERVAL 30 DAY)
                GROUP BY dia 
                ORDER BY dia ASC
            """)
            dias_data = cur.fetchall()
            dias = [d['dia'] for d in dias_data]
            totales_dia = [d['total'] for d in dias_data]

            cur.execute("""
                SELECT tipo_falla, COUNT(*) as total 
                FROM tickets 
                WHERE tipo_falla IS NOT NULL AND tipo_falla != '' 
                GROUP BY tipo_falla 
                ORDER BY total DESC 
                LIMIT 10
            """)
            fallas = cur.fetchall()
            fallas_labels = [f['tipo_falla'] for f in fallas]
            fallas_data = [f['total'] for f in fallas]

            return jsonify({
                'semanas': semanas,
                'totales_semana': totales_semana,
                'meses': meses,
                'totales_mes': totales_mes,
                'dias': dias,
                'totales_dia': totales_dia,
                'fallas_labels': fallas_labels,
                'fallas_data': fallas_data
            })

        elif filtro == 'clientes_zona':
            cur.execute("""
                SELECT nodo, COUNT(*) as total
                FROM clientes
                WHERE nodo IS NOT NULL AND nodo != ''
                GROUP BY nodo
                ORDER BY total DESC
            """)
            nodos = cur.fetchall()
            dona_labels = [n['nodo'] for n in nodos]
            dona_data = [n['total'] for n in nodos]
            colores = []
            for _ in nodos:
                r, g, b = random.randint(50, 200), random.randint(50, 200), random.randint(50, 200)
                colores.append(f'rgba({r},{g},{b},0.8)')
            return jsonify({
                'dona_labels': dona_labels,
                'dona_data': dona_data,
                'colores_dona': colores
            })

        return jsonify({'error': 'Filtro no válido'}), 400

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500
    finally:
        cur.close()

# ---------- UTILIDADES ----------
@app.context_processor
def utility_processor():
    import time
    return dict(now=datetime.now, version=int(time.time()))

# ============================================================
#   EJECUCIÓN
# ============================================================
if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    debug_mode = os.environ.get('FLASK_DEBUG', '0') == '1'
    app.run(debug=debug_mode, host='0.0.0.0', port=port)