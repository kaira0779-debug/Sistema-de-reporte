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
    'asignadores': {'data': None, 'timestamp': 0},
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
        cur.execute("UPDATE trabajadores SET online = 1 WHERE id = %s", (current_user.id,))
        conn.commit()
        cur.close()

@app.context_processor
def inject_user():
    return dict(current_user=current_user)

# ============================================================
#   CREAR ADMIN SI NO EXISTE
# ============================================================
def crear_admin_si_no_existe():
    conn = get_db()
    cur = conn.cursor()
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
    cur.execute("SELECT num_ticket FROM tickets WHERE num_ticket LIKE %s ORDER BY num_ticket DESC LIMIT 1", (f'{year}-%',))
    row = cur.fetchone()
    cur.close()
    if row:
        try:
            num_str = row['num_ticket'].split('-')[1]
            next_num = int(num_str) + 1
        except (ValueError, IndexError):
            next_num = 5000
    else:
        next_num = 5000
    return f"{year}-{next_num}"

# ============================================================
#   FUNCIONES DE CLIENTES
# ============================================================
def guardar_cliente(datos):
    conn = get_db()
    cur = conn.cursor()
    n_contrato = datos.get('n_contrato', '').strip()
    if not n_contrato:
        cur.execute("""
            INSERT INTO clientes (n_contrato, n_documento, nombre_apellido, telefono, direccion_servicio, nodo, mbps_contratados, estatus_usuario, asignador)
            VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
        """, ('', datos.get('n_documento','').strip(), datos.get('nombre_apellido','').strip(),
              datos.get('telefono','').strip(), datos.get('direccion_servicio','').strip(),
              datos.get('nodo','').strip(), datos.get('mbps_contratados','').strip(),
              datos.get('estatus_usuario','').strip(), datos.get('asignador','').strip()))
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
                nodo=%s, mbps_contratados=%s, estatus_usuario=%s, asignador=%s
                WHERE id=%s
            """, (datos.get('n_documento','').strip(), datos.get('nombre_apellido','').strip(),
                  datos.get('telefono','').strip(), datos.get('direccion_servicio','').strip(),
                  datos.get('nodo','').strip(), datos.get('mbps_contratados','').strip(),
                  datos.get('estatus_usuario','').strip(), datos.get('asignador','').strip(), cliente['id']))
            conn.commit()
            cur.close()
            invalidar_cache()
            return 'updated'
        else:
            cur.execute("""
                INSERT INTO clientes (n_contrato, n_documento, nombre_apellido, telefono, direccion_servicio, nodo, mbps_contratados, estatus_usuario, asignador)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
            """, (n_contrato, datos.get('n_documento','').strip(), datos.get('nombre_apellido','').strip(),
                  datos.get('telefono','').strip(), datos.get('direccion_servicio','').strip(),
                  datos.get('nodo','').strip(), datos.get('mbps_contratados','').strip(),
                  datos.get('estatus_usuario','').strip(), datos.get('asignador','').strip()))
            conn.commit()
            cur.close()
            invalidar_cache()
            return 'inserted'

# ============================================================
#   LECTURA Y MAPEO DE ARCHIVOS (funciones auxiliares)
# ============================================================
CAMPOS_CLIENTES = [
    'n_contrato', 'n_documento', 'nombre_apellido', 'telefono',
    'direccion_servicio', 'nodo', 'mbps_contratados', 'estatus_usuario', 'asignador'
]

CAMPOS_TICKETS = [
    'fecha', 'num_ticket', 'nodo', 'cliente', 'motivo_ticket',
    'visita_tecnica', 'tipo_falla', 'departamento', 'status', 'responsable', 'observaciones'
]

SINONIMOS = {
    'n_contrato': ['contrato', 'n° contrato', 'nro contrato', 'número contrato', 'contrato cliente', 'id contrato'],
    'n_documento': ['documento', 'cédula', 'rif', 'cedula', 'n° documento', 'nro documento', 'identificación'],
    'nombre_apellido': ['nombre', 'cliente', 'nombre cliente', 'nombre y apellido', 'razón social', 'empresa', 'persona'],
    'telefono': ['teléfono', 'celular', 'movil', 'contacto', 'tel', 'phone'],
    'direccion_servicio': ['dirección', 'direccion', 'servicio', 'direccion de servicio', 'ubicación', 'domicilio'],
    'nodo': ['zona', 'sector', 'nodo', 'torre', 'zona de cobertura'],
    'mbps_contratados': ['mbps', 'velocidad', 'plan', 'mb', 'ancho de banda', 'contratado'],
    'estatus_usuario': ['estatus', 'estado', 'condición', 'situación', 'status'],
    'asignador': ['asignado por', 'vendedor', 'tecnico asignador', 'responsable instalación'],
    'fecha': ['fecha', 'dia', 'date', 'fecha de ticket'],
    'num_ticket': ['ticket', 'nº ticket', 'numero ticket', 'nro ticket', 'id ticket', 'número de ticket'],
    'cliente': ['cliente', 'nombre cliente', 'usuario', 'afiliado', 'contratante'],
    'motivo_ticket': ['motivo', 'razón', 'descripción', 'falla reportada', 'problema', 'motivo de ticket'],
    'visita_tecnica': ['visita', 'técnico', 'visita tecnica', 'requiere visita', 'desplazamiento'],
    'tipo_falla': ['falla', 'tipo falla', 'problema', 'categoría', 'clasificación'],
    'departamento': ['área', 'departamento', 'unidad', 'gerencia'],
    'status': ['estado', 'estatus', 'situación', 'condición'],
    'responsable': ['técnico', 'responsable', 'asignado a', 'encargado', 'personal'],
    'observaciones': ['observación', 'notas', 'comentarios', 'detalles', 'adicional']
}

def detectar_delimitador(archivo):
    with open(archivo, 'r', encoding='utf-8-sig') as f:
        sample = f.read(1024)
        f.seek(0)
    delimitadores = [',', ';', '\t', '|']
    for delim in delimitadores:
        if delim in sample:
            return delim
    return ','

def leer_archivo_datos(filepath, extension):
    if extension in ['.xlsx', '.xls']:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        hoja = wb.active
        filas = list(hoja.iter_rows(values_only=True))
        return [[str(cell).strip() if cell is not None else '' for cell in row] for row in filas if any(cell for cell in row)]
    else:
        with open(filepath, 'rb') as f:
            raw = f.read(10000)
            encoding = chardet.detect(raw)['encoding'] or 'utf-8-sig'
        with open(filepath, 'r', encoding=encoding) as f:
            sample = f.read(1024)
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
    header = [str(c).strip().lower() for c in filas[0]]
    header = [re.sub(r'[^\w\s]', '', h) for h in header]
    mapeo = {}
    usado = set()
    for campo in campos_esperados:
        sinonimos_campo = sinonimos.get(campo, [])
        candidatos = [campo] + sinonimos_campo
        mejor_idx = None
        mejor_score = 0
        for idx, col in enumerate(header):
            if idx in usado:
                continue
            for cand in candidatos:
                score = difflib.SequenceMatcher(None, col, cand).ratio()
                if cand in col or col in cand:
                    score += 0.3
                if score > mejor_score:
                    mejor_score = score
                    mejor_idx = idx
        if mejor_idx is not None and mejor_score >= 0.5:
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
                val = re.sub(r'[^\w\-\.]', '', val)
        datos[campo] = val
    return datos

def buscar_cliente_flexible(nombre_busqueda, conn):
    if not nombre_busqueda:
        return None
    nombre_normalizado = ' '.join(nombre_busqueda.split()).upper()
    cur = conn.cursor()
    cur.execute("SELECT * FROM clientes WHERE UPPER(REPLACE(nombre_apellido, '  ', ' ')) = %s", (nombre_normalizado,))
    cliente = cur.fetchone()
    if cliente:
        cur.close()
        return cliente
    cur.execute("SELECT * FROM clientes WHERE UPPER(nombre_apellido) LIKE %s", (f'%{nombre_normalizado}%',))
    cliente = cur.fetchone()
    if cliente:
        cur.close()
        return cliente
    palabras = nombre_normalizado.split()
    if len(palabras) > 1:
        condiciones = ' AND '.join(['UPPER(nombre_apellido) LIKE %s'] * len(palabras))
        params = [f'%{p}%' for p in palabras]
        cur.execute(f"SELECT * FROM clientes WHERE {condiciones}", params)
        cliente = cur.fetchone()
        if cliente:
            cur.close()
            return cliente
    cur.close()
    return None

# ============================================================
#                           RUTAS PRINCIPALES
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
    cur.execute("UPDATE trabajadores SET online = 0 WHERE id = %s", (current_user.id,))
    conn.commit()
    cur.close()
    logout_user()
    return redirect(url_for('login'))

@app.route('/dashboard')
@login_required
def dashboard():
    if current_user.rol != 'admin':
        return redirect(url_for('panel_tecnico'))
    return render_template('dashboard.html')

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
        SELECT a.*, t.num_ticket, t.cliente 
        FROM asignaciones a 
        LEFT JOIN tickets t ON a.ticket_id = t.id 
        WHERE a.tecnico_id = %s 
        ORDER BY a.fecha_asignacion DESC
    """, (current_user.id,))
    asignaciones = cur.fetchall()
    cur.close()
    return render_template('panel_tecnico.html', asignaciones=asignaciones)

@app.route('/api/tecnicos/notificaciones')
@login_required
def api_notificaciones_tecnico():
    if current_user.rol != 'tecnico':
        return jsonify({'error': 'Acceso denegado'}), 403
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) as nuevas FROM asignaciones WHERE tecnico_id = %s AND estado = 'Pendiente'", (current_user.id,))
    row = cur.fetchone()
    cur.close()
    return jsonify({'nuevas': row['nuevas']})

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
        os.unlink(temp.name)
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
            os.unlink(temp.name)
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
        'estatus_usuario': request.form.get('estatus_usuario', '').strip(),
        'asignador': request.form.get('asignador', '').strip()
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
        asignador = request.form.get('asignador', '').strip()
        if not nombre_apellido:
            flash('El nombre del cliente es obligatorio', 'danger')
            return render_template('editar_cliente.html', cliente=cliente)
        cur.execute("""
            UPDATE clientes SET n_contrato=%s, n_documento=%s, nombre_apellido=%s, telefono=%s,
            direccion_servicio=%s, nodo=%s, mbps_contratados=%s, estatus_usuario=%s, asignador=%s
            WHERE id=%s
        """, (n_contrato, n_documento, nombre_apellido, telefono, direccion_servicio, nodo, mbps_contratados, estatus_usuario, asignador, cliente_id))
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
    cur.execute("SELECT * FROM tickets ORDER BY fecha DESC, num_ticket DESC LIMIT 500")
    tickets = cur.fetchall()
    cur.execute("SELECT DISTINCT nodo FROM clientes WHERE nodo IS NOT NULL AND nodo != '' UNION SELECT DISTINCT nodo FROM tickets WHERE nodo IS NOT NULL AND nodo != '' ORDER BY nodo")
    nodos = [row['nodo'] for row in cur.fetchall()]
    cur.execute("SELECT DISTINCT motivo_ticket FROM tickets WHERE motivo_ticket IS NOT NULL AND motivo_ticket != '' ORDER BY motivo_ticket")
    motivos = [row['motivo_ticket'] for row in cur.fetchall()]
    cur.execute("SELECT DISTINCT tipo_falla FROM tickets WHERE tipo_falla IS NOT NULL AND tipo_falla != '' ORDER BY tipo_falla")
    tipos_falla = [row['tipo_falla'] for row in cur.fetchall()]
    cur.execute("SELECT DISTINCT nombre_completo FROM trabajadores WHERE rol = 'tecnico' ORDER BY nombre_completo")
    tecnicos = [row['nombre_completo'] for row in cur.fetchall()]
    cur.close()
    return render_template('tickets.html', tickets=tickets, nodos=nodos, motivos=motivos, tipos_falla=tipos_falla, tecnicos=tecnicos)

@app.route('/api/buscar_tickets_json')
@login_required
def api_buscar_tickets():
    q = request.args.get('q', '').strip()
    conn = get_db()
    cur = conn.cursor()
    if q:
        cur.execute("""SELECT * FROM tickets WHERE num_ticket LIKE %s OR cliente LIKE %s OR nodo LIKE %s OR motivo_ticket LIKE %s OR tipo_falla LIKE %s OR status LIKE %s OR responsable LIKE %s OR observaciones LIKE %s ORDER BY fecha DESC, num_ticket DESC LIMIT 50""", (f'%{q}%',)*8)
    else:
        cur.execute("SELECT * FROM tickets ORDER BY fecha DESC, num_ticket DESC LIMIT 50")
    tickets = cur.fetchall()
    cur.close()
    for t in tickets:
        if t['fecha']:
            t['fecha'] = t['fecha'].strftime('%Y-%m-%d')
    return jsonify(tickets)

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
        os.unlink(temp.name)
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
        for fila in filas[1:]:
            if not any(fila):
                continue
            datos = procesar_fila_datos(fila, mapeo, 'tickets')
            if not datos.get('cliente') or not datos.get('fecha'):
                errores += 1
                continue
            num_ticket = str(datos.get('num_ticket', ''))
            if num_ticket and datos['fecha']:
                cur.execute("SELECT id FROM tickets WHERE num_ticket = %s AND fecha = %s", (num_ticket, datos['fecha']))
                if cur.fetchone():
                    errores += 1
                    continue
            if 'ap caida' in datos.get('tipo_falla', '').lower():
                datos['tipo_falla'] = 'AP caída'
            cur.execute("""INSERT INTO tickets (fecha, num_ticket, nodo, cliente, motivo_ticket, visita_tecnica, tipo_falla, departamento, status, responsable, observaciones) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""",
                (datos.get('fecha'), num_ticket, datos.get('nodo', ''), datos.get('cliente'), datos.get('motivo_ticket', ''),
                 datos.get('visita_tecnica', ''), datos.get('tipo_falla', ''), datos.get('departamento', ''),
                 datos.get('status', 'Pendiente'), datos.get('responsable', ''), datos.get('observaciones', '')))
            insertados += 1
        conn.commit()
        cur.close()
        flash(f'{insertados} tickets importados correctamente, {errores} errores.', 'success')
    except Exception as e:
        flash(f'Error al procesar el archivo: {str(e)}', 'danger')
    return redirect(url_for('tickets'))

@app.route('/agregar_ticket', methods=['POST'])
@login_required
def agregar_ticket():
    nodo = request.form.get('nodo')
    cliente_id = request.form.get('cliente_id')
    motivo_ticket = request.form.get('motivo_ticket')
    visita_tecnica = request.form.get('visita_tecnica')
    tipo_falla = request.form.get('tipo_falla')
    status = request.form.get('status')
    responsable = request.form.get('responsable')
    observaciones = request.form.get('observaciones')
    if not cliente_id:
        flash('El cliente es obligatorio', 'danger')
        return redirect(url_for('tickets'))
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT nombre_apellido FROM clientes WHERE id = %s", (cliente_id,))
    cliente = cur.fetchone()
    if not cliente:
        flash('Cliente no encontrado', 'danger')
        cur.close()
        return redirect(url_for('tickets'))
    cliente_nombre = cliente['nombre_apellido']
    num_ticket = generar_numero_ticket()
    fecha = datetime.now().strftime('%Y-%m-%d')
    if re.search(r'ap\s*ca[ií]da', tipo_falla, re.IGNORECASE):
        tipo_falla = 'AP caída'
    cur.execute("""INSERT INTO tickets (fecha, num_ticket, nodo, cliente, motivo_ticket, visita_tecnica, tipo_falla, departamento, status, responsable, observaciones) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)""", (fecha, num_ticket, nodo, cliente_nombre, motivo_ticket, visita_tecnica, tipo_falla, '', status, responsable, observaciones))
    conn.commit()
    cur.close()
    flash(f'Ticket {num_ticket} agregado correctamente', 'success')
    return redirect(url_for('tickets'))

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
    cur.execute("SELECT * FROM tickets WHERE id = %s", (ticket_id,))
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
        observaciones = request.form.get('observaciones', '').strip()
        if re.search(r'ap\s*ca[ií]da', tipo_falla, re.IGNORECASE):
            tipo_falla = 'AP caída'
        cur.execute("""UPDATE tickets SET nodo=%s, motivo_ticket=%s, visita_tecnica=%s, tipo_falla=%s, departamento=%s, status=%s, responsable=%s, observaciones=%s WHERE id=%s""", (nodo, motivo_ticket, visita_tecnica, tipo_falla, departamento, status, responsable, observaciones, ticket_id))
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
    cur.execute("SELECT DISTINCT nombre_completo FROM trabajadores WHERE rol = 'tecnico' ORDER BY nombre_completo")
    tecnicos = [row['nombre_completo'] for row in cur.fetchall()]
    cur.close()
    return render_template('editar_ticket.html', ticket=ticket, nodos=nodos, motivos=motivos, tipos_falla=tipos_falla, estados=estados, tecnicos=tecnicos)

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
        password = request.form.get('password', '')
        if not username or not password or not nombre_completo:
            flash('Todos los campos son obligatorios', 'danger')
            return redirect(url_for('admin_tecnicos'))
        try:
            cur.execute("SELECT id FROM trabajadores WHERE username = %s", (username,))
            if cur.fetchone():
                flash('El nombre de usuario ya existe', 'danger')
                return redirect(url_for('admin_tecnicos'))
            hashed = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            cur.execute("INSERT INTO trabajadores (username, password_hash, nombre_completo, cedula, rol, online) VALUES (%s,%s,%s,%s,'tecnico',0)", (username, hashed, nombre_completo, cedula))
            conn.commit()
            nuevo_id = cur.lastrowid
            flash(f'Técnico registrado exitosamente con ID {nuevo_id}', 'success')
        except Exception as e:
            flash(f'Error: {str(e)}', 'danger')
        finally:
            cur.close()
        invalidar_cache()
        return redirect(url_for('admin_tecnicos'))
    cur.execute("SELECT id, username, nombre_completo, cedula, rol, online FROM trabajadores WHERE rol = 'tecnico' ORDER BY creado_en DESC")
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
#   DUPLICADOS (EXACTOS)
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
#   ASIGNACIÓN DE TRABAJOS
# ============================================================
@app.route('/admin/asignaciones', methods=['GET', 'POST'])
@login_required
def admin_asignaciones():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    conn = get_db()
    cur = conn.cursor()
    if request.method == 'POST':
        tecnico_id = request.form.get('tecnico_id')
        descripcion = request.form.get('descripcion')
        ticket_id = request.form.get('ticket_id') or None
        tipo_trabajo = request.form.get('tipo_trabajo', '').strip() or None
        prioridad = request.form.get('prioridad', 'Media')
        fecha_limite = request.form.get('fecha_limite') or None
        zona = request.form.get('zona', '').strip() or None
        if not tecnico_id or not descripcion:
            flash('Todos los campos obligatorios deben completarse', 'danger')
        else:
            cur.execute("""
                INSERT INTO asignaciones (tecnico_id, admin_id, descripcion, ticket_id, tipo_trabajo, prioridad, fecha_limite, zona)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s)
            """, (tecnico_id, current_user.id, descripcion, ticket_id, tipo_trabajo, prioridad, fecha_limite, zona))
            conn.commit()
            flash('Asignación creada correctamente', 'success')
        return redirect(url_for('admin_asignaciones'))
    cur.execute("SELECT id, nombre_completo FROM trabajadores WHERE rol = 'tecnico' ORDER BY nombre_completo")
    tecnicos = cur.fetchall()
    cur.execute("SELECT DISTINCT nodo FROM clientes WHERE nodo IS NOT NULL AND nodo != '' ORDER BY nodo")
    nodos = [row['nodo'] for row in cur.fetchall()]
    cur.execute("""
        SELECT a.*, t.nombre_completo as tecnico_nombre, tk.num_ticket, tk.cliente
        FROM asignaciones a
        JOIN trabajadores t ON a.tecnico_id = t.id
        LEFT JOIN tickets tk ON a.ticket_id = tk.id
        ORDER BY a.fecha_asignacion DESC
    """)
    asignaciones = cur.fetchall()
    cur.close()
    return render_template('admin_asignaciones.html', tecnicos=tecnicos, asignaciones=asignaciones, nodos=nodos)

@app.route('/admin/asignacion/<int:id>/estado', methods=['POST'])
@login_required
def cambiar_estado_asignacion(id):
    if current_user.rol not in ('admin', 'tecnico'):
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    nuevo_estado = request.form.get('estado')
    conn = get_db()
    cur = conn.cursor()
    cur.execute("UPDATE asignaciones SET estado=%s WHERE id=%s", (nuevo_estado, id))
    conn.commit()
    cur.close()
    flash('Estado actualizado', 'success')
    return redirect(request.referrer or url_for('panel_tecnico'))

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

def generar_pdf_planilla(cliente, tecnico, motivo, tipo_falla, observaciones, num_ticket, hora_inicio, hora_fin):
    buffer = BytesIO()
    c = pdfcanvas.Canvas(buffer, pagesize=letter)
    width, height = letter
    margin = 50
    x_left = margin
    x_right = width - margin
    y_start = height - margin

    c.setFont("Helvetica-Bold", 16)
    c.drawString(x_left, y_start, "DAITNVOZ C.A")
    c.setFont("Helvetica", 9)
    c.drawString(x_left, y_start - 18, "Avenida Principal de los Chorros, Qta Datinvoz, Caracas 1071, Distrito Capital, Venezuela")
    c.drawString(x_left, y_start - 32, "Teléfono: 0424-1637944")
    c.setFont("Helvetica-Bold", 13)
    c.drawCentredString(width/2, y_start - 52, "ORDEN DE SERVICIO TECNICO - SOPORTE DE INTERNET")

    c.setFont("Helvetica", 9)
    y_line = y_start - 72
    c.drawString(x_left, y_line, f"N° Ticket: {num_ticket if num_ticket else '______________'}")
    c.drawString(x_left + 180, y_line, f"Fecha: {datetime.now().strftime('%d/%m/%Y')}")
    c.drawString(x_left + 340, y_line, f"Hora Inicio: {hora_inicio if hora_inicio else '________'}")
    c.drawString(x_left + 500, y_line, f"Hora Fin: {hora_fin if hora_fin else '________'}")

    y = y_line - 25
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_left, y, "DATOS DEL CLIENTE")
    y -= 18
    c.setFont("Helvetica", 10)
    nombre_cliente = cliente.get('nombre_apellido', '') if cliente.get('nombre_apellido') else '__________________________'
    direccion = cliente.get('direccion_servicio', '') if cliente.get('direccion_servicio') else '__________________________'
    telefono = cliente.get('telefono', '') if cliente.get('telefono') else '____________________'
    c.drawString(x_left, y, f"Nombre/Empresa: {nombre_cliente}")
    y -= 16
    c.drawString(x_left, y, f"Dirección: {direccion}")
    y -= 16
    c.drawString(x_left, y, f"Teléfono: {telefono}")
    c.drawString(x_left + 230, y, f"Correo: ____________________")

    y -= 25
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_left, y, "DESCRIPCION DEL SERVICIO REALIZADO:")
    y -= 16
    c.setFont("Helvetica", 10)
    for i in range(4):
        c.drawString(x_left, y - i*14, "_______________________________________________________________________________")
    if observaciones:
        lines = observaciones.split('\n')
        for i, line in enumerate(lines[:4]):
            c.setFillColor(colors.white)
            c.rect(x_left+2, y - i*14 - 10, 500, 12, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.drawString(x_left+4, y - i*14, line[:70])
    y -= 60

    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_left, y, "MATERIALES ENTREGADOS / INSTALADOS")
    y -= 16
    c.setFont("Helvetica", 9)
    c.drawString(x_left, y, "Cant.")
    c.drawString(x_left + 50, y, "Descripción")
    c.drawString(x_left + 220, y, "Serial/N° Lote")
    c.drawString(x_left + 380, y, "Observaciones")
    y -= 10
    c.line(x_left, y, x_right, y)
    y -= 8
    for i in range(4):
        c.drawString(x_left, y, "____")
        c.drawString(x_left + 50, y, "________________")
        c.drawString(x_left + 220, y, "_______________")
        c.drawString(x_left + 380, y, "________________")
        y -= 16

    y -= 10
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_left, y, "DATOS DEL TÉCNICO")
    y -= 18
    c.setFont("Helvetica", 10)
    if tecnico and tecnico.get('nombre_completo'):
        c.drawString(x_left, y, f"Nombre: {tecnico['nombre_completo']}")
    else:
        c.drawString(x_left, y, "Nombre: __________________________")
    c.drawString(x_left + 250, y, "ID/C.I.: ________________")
    c.drawString(x_left + 430, y, "Celular: ________________")
    y -= 16
    c.drawString(x_left, y, "Firma: __________________________")

    y -= 20
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_left, y, "VERIFICACIONES")
    y -= 16
    c.setFont("Helvetica", 10)
    c.drawString(x_left, y, "[ ] Servicio funcional correctamente")
    c.drawString(x_left + 220, y, "[ ] Velocidad verificada")
    c.drawString(x_left + 430, y, "[ ] Cliente satisfecho")

    y -= 20
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_left, y, "REQUIERE VISITA ADICIONAL/INSPECCION")
    y -= 16
    c.setFont("Helvetica", 10)
    c.drawString(x_left, y, "[ ] Cableado dañado")
    c.drawString(x_left + 160, y, "[ ] Equipo defectuoso")
    c.drawString(x_left + 320, y, "[ ] Servicio no resuelto")
    c.drawString(x_left + 480, y, "[ ] Solo inspección")

    y -= 25
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x_left, y, "VALIDACION FINAL")
    y -= 18
    c.setFont("Helvetica", 10)
    c.drawString(x_left, y, "Cliente: ________________")
    c.drawString(x_left + 180, y, "Técnico: ________________")
    c.drawString(x_left + 360, y, "Revisado por: ________________")
    y -= 14
    c.drawString(x_left, y, "Firma: ________________")
    c.drawString(x_left + 180, y, "Firma: ________________")
    c.drawString(x_left + 360, y, "Firma: ________________")

    c.setFont("Helvetica-Oblique", 8)
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
    cliente_nombre = request.args.get('cliente_nombre', '')
    motivo = request.args.get('motivo', '')
    tipo_falla = request.args.get('tipo_falla', '')
    observaciones = request.args.get('observaciones', '')
    ticket_id = request.args.get('ticket_id', '')
    return render_template('planilla_soporte.html', generados=generados, motivos=MOTIVOS_TICKET, tipos_falla=TIPOS_FALLA,
                           cliente_nombre=cliente_nombre, motivo=motivo, tipo_falla=tipo_falla,
                           observaciones=observaciones, ticket_id=ticket_id)

@app.route('/planilla_soporte/generar_planilla', methods=['POST'])
@login_required
def generar_planilla_cliente():
    cliente_id = request.form.get('cliente_id')
    tecnico_id = request.form.get('tecnico_id')
    motivo = request.form.get('motivo_ticket')
    tipo_falla = request.form.get('tipo_falla')
    observaciones_adicionales = request.form.get('observaciones_adicionales', '').strip()
    num_ticket = request.form.get('num_ticket', '').strip()
    hora_inicio = request.form.get('hora_inicio', '').strip()
    hora_fin = request.form.get('hora_fin', '').strip()

    if not cliente_id:
        flash('Falta seleccionar un cliente', 'danger')
        return redirect(url_for('planilla_soporte'))

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM clientes WHERE id = %s", (cliente_id,))
    cliente = cur.fetchone()
    if not cliente:
        flash('Cliente no encontrado', 'danger')
        cur.close()
        return redirect(url_for('planilla_soporte'))
    cur.close()

    tecnico = None
    if tecnico_id:
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT id, username, nombre_completo FROM trabajadores WHERE id = %s AND rol = 'tecnico'", (tecnico_id,))
        tecnico = cur.fetchone()
        cur.close()

    if not num_ticket:
        num_ticket = generar_numero_ticket()

    try:
        pdf_buffer = generar_pdf_planilla(cliente, tecnico, motivo, tipo_falla, observaciones_adicionales, num_ticket, hora_inicio, hora_fin)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"planilla_{cliente_id}_{timestamp}.pdf"
        output_path = os.path.join(GENERADOS_FOLDER, output_filename)
        with open(output_path, 'wb') as f:
            f.write(pdf_buffer.read())
        flash(f'Planilla PDF generada correctamente con ticket {num_ticket}', 'success')
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
        filepath = os.path.join(GENERADOS_FOLDER, filename)
        if os.path.exists(filepath):
            os.remove(filepath)
            eliminados += 1
    flash(f'Se eliminaron {eliminados} planillas correctamente.', 'success')
    return redirect(url_for('planilla_soporte'))

# ============================================================
#   RESPALDOS (COMPLETO)
# ============================================================
def registrar_respaldo(filename, filepath, tipo):
    try:
        size = os.path.getsize(filepath)
        conn = get_db()
        cur = conn.cursor()
        cur.execute("INSERT INTO backup_log (filename, filepath, size, tipo) VALUES (%s,%s,%s,%s)", (filename, filepath, size, tipo))
        conn.commit()
        cur.close()
    except Exception as e:
        print("Error registrando respaldo:", e)

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
            tables = get_table_names()
            conn = get_db()
            cur = conn.cursor()
            for table in tables:
                create_sql = get_table_schema(table)
                if create_sql:
                    f.write(f"DROP TABLE IF EXISTS `{table}`;\n")
                    f.write(f"{create_sql};\n\n")
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
    if filepath and not error:
        registrar_respaldo(os.path.basename(filepath), filepath, tipo)
    return filepath, error

def get_table_names():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SHOW TABLES")
    tables = [list(row.values())[0] for row in cur.fetchall()]
    cur.close()
    return tables

def get_table_schema(table):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(f"SHOW CREATE TABLE `{table}`")
    result = cur.fetchone()
    cur.close()
    return result['Create Table'] if result else None

def respaldo_semanal():
    with app.app_context():
        filepath, error = crear_respaldo('semanal')
        if error:
            print(f"[ERROR] Respaldo semanal falló: {error}")
        else:
            print(f"[OK] Respaldo semanal generado: {filepath}")

def respaldo_mensual():
    with app.app_context():
        filepath, error = crear_respaldo('mensual')
        if error:
            print(f"[ERROR] Respaldo mensual falló: {error}")
        else:
            print(f"[OK] Respaldo mensual generado: {filepath}")

scheduler = BackgroundScheduler()
scheduler.add_job(func=respaldo_semanal, trigger='cron', day_of_week='sun', hour=3, minute=0)
scheduler.add_job(func=respaldo_mensual, trigger='cron', day=1, hour=4, minute=0)
scheduler.start()
atexit.register(lambda: scheduler.shutdown())

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
    return redirect(url_for('admin_panel'))

@app.route('/admin/respaldos')
@login_required
def listar_respaldos():
    if current_user.rol != 'admin':
        flash('Acceso denegado', 'danger')
        return redirect(url_for('dashboard'))
    backups = []
    backup_dir = app.config['BACKUP_FOLDER']
    if os.path.exists(backup_dir):
        for f in os.listdir(backup_dir):
            if f.endswith('.sql'):
                ruta = os.path.join(backup_dir, f)
                tamaño = os.path.getsize(ruta)
                fecha_mod = datetime.fromtimestamp(os.path.getmtime(ruta))
                backups.append({
                    'nombre': f,
                    'tamaño': tamaño,
                    'fecha': fecha_mod.strftime('%d/%m/%Y %H:%M:%S'),
                    'tamaño_legible': f"{tamaño / 1024:.1f} KB" if tamaño < 1024*1024 else f"{tamaño / (1024*1024):.2f} MB"
                })
    backups.sort(key=lambda x: x['fecha'], reverse=True)
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM backup_log ORDER BY fecha DESC")
    logs = cur.fetchall()
    cur.close()
    return render_template('respaldos.html', backups=backups, logs=logs)

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

# ---------- API TÉCNICOS ----------
@app.route('/api/tecnicos_json')
@login_required
def api_tecnicos_json():
    def fetch_tecnicos():
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT id, username, nombre_completo, online FROM trabajadores WHERE rol = 'tecnico' ORDER BY nombre_completo")
        data = cur.fetchall()
        cur.close()
        return data
    return jsonify(get_cache('tecnicos', fetch_tecnicos))

# ---------- API BÚSQUEDA CLIENTES ----------
@app.route('/api/buscar_cliente_json')
@login_required
def api_buscar_cliente():
    q = request.args.get('q', '').strip()
    if not q:
        return jsonify([])
    conn = get_db()
    cur = conn.cursor()
    q_upper = q.upper().strip()
    cur.execute("""
        SELECT id, n_contrato, nombre_apellido, telefono, nodo, n_documento, direccion_servicio, mbps_contratados, estatus_usuario, asignador
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
                ELSE 3
            END
        LIMIT 20
    """, (f'%{q_upper}%', f'%{q_upper}%', f'%{q_upper}%', f'%{q_upper}%', f'%{q_upper}%', f'%{q_upper}%', q_upper, f'%{q_upper}%'))
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
            'estatus': c['estatus_usuario'] or '',
            'asignador': c['asignador'] or ''
        })
    cur.close()
    return jsonify(clientes)

@app.route('/api/cliente/<int:cliente_id>')
@login_required
def api_cliente_detalle(cliente_id):
    if current_user.rol != 'admin':
        return jsonify({'error': 'Acceso denegado'}), 403
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
        'telefono': cliente['telefono'] or '',
        'direccion_servicio': cliente['direccion_servicio'] or '',
        'nodo': cliente['nodo'] or '',
        'mbps_contratados': cliente['mbps_contratados'] or '',
        'estatus_usuario': cliente['estatus_usuario'] or '',
        'asignador': cliente.get('asignador', '') or ''
    })

# ---------- API ASIGNADORES ----------
@app.route('/api/asignadores')
@login_required
def api_asignadores():
    def fetch_asignadores():
        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT DISTINCT asignador FROM clientes WHERE asignador IS NOT NULL AND asignador != '' ORDER BY asignador")
        data = [row['asignador'] for row in cur.fetchall()]
        cur.close()
        return data
    return jsonify(get_cache('asignadores', fetch_asignadores))

# ---------- API USUARIOS EN LÍNEA ----------
@app.route('/api/usuarios_online')
@login_required
def api_usuarios_online():
    if current_user.rol != 'admin':
        return jsonify({'error': 'Acceso denegado'}), 403
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT id, username, nombre_completo, rol FROM trabajadores WHERE online = 1")
    usuarios = cur.fetchall()
    cur.close()
    return jsonify(usuarios)

# ============================================================
#        GRÁFICAS (SOLO ADMIN) - CORREGIDO PARA DIAS
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
            # Semana
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

            # Mes
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

            # Día (últimos 30 días) - USAR DATE_FORMAT para asegurar formato
            cur.execute("""
                SELECT DATE_FORMAT(fecha, '%Y-%m-%d') as dia, COUNT(*) as total 
                FROM tickets 
                WHERE fecha >= DATE_SUB(CURDATE(), INTERVAL 30 DAY)
                GROUP BY dia 
                ORDER BY dia ASC
            """)
            dias_data = cur.fetchall()
            dias = [d['dia'] for d in dias_data]  # ya vienen como strings
            totales_dia = [d['total'] for d in dias_data]

            # Fallas
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
#   EJECUCIÓN (Waitress)
# ============================================================
if __name__ == '__main__':
    from waitress import serve
    port = int(os.environ.get('PORT', 5000))
    serve(app, host='0.0.0.0', port=port, threads=8, channel_timeout=120)