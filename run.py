import os
import sys
import socket
import time

def obtener_ips_locales():
    """Detecta las direcciones IP locales para compartirlas en la red."""
    ips = []
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(0.5)
        s.connect(('8.8.8.8', 80))
        ip_principal = s.getsockname()[0]
        ips.append(ip_principal)
        s.close()
    except Exception:
        pass
        
    try:
        hostname = socket.gethostname()
        for ip in socket.gethostbyname_ex(hostname)[2]:
            if not ip.startswith('127.') and ip not in ips:
                ips.append(ip)
    except Exception:
        pass
        
    return ips

def verificar_base_datos():
    """Comprueba si MySQL está activo antes de arrancar el servidor."""
    from config import Config
    import pymysql
    
    intentos = 5
    for i in range(1, intentos + 1):
        try:
            conn = pymysql.connect(
                host=Config.MYSQL_HOST,
                port=Config.MYSQL_PORT,
                user=Config.MYSQL_USER,
                password=Config.MYSQL_PASSWORD,
                database=Config.MYSQL_DB,
                connect_timeout=3
            )
            conn.close()
            print("✅ Conexión a la base de datos MySQL establecida correctamente.")
            return True
        except Exception:
            print(f"⏳ Intento {i}/{intentos}: Esperando a MySQL ({Config.MYSQL_HOST}:{Config.MYSQL_PORT})...")
            time.sleep(2)
            
    print("\n⚠️ ADVERTENCIA: No se pudo conectar a MySQL. Asegúrate de que el servicio MySQL/XAMPP esté iniciado.")
    return False

if __name__ == '__main__':
    print("=" * 60)
    print("   SISTEMA DE REPORTES DE FALLAS Y TICKETS (DATINVOZ)")
    print("=" * 60)
    
    verificar_base_datos()
    
    try:
        from app import app
    except ImportError as e:
        print(f"❌ Error importando módulos: {e}")
        print("   Ejecuta: pip install -r requirements.txt")
        sys.exit(1)
        
    port = int(os.environ.get('PORT', 5000))
    ips = obtener_ips_locales()
    
    print("\n✅ SERVIDOR LISTO Y ESCUCHANDO:")
    print(f"   • Acceso local (en esta PC):       http://localhost:{port}")
    print(f"   • Acceso local IP:                 http://127.0.0.1:{port}")
    
    if ips:
        print("\n📡 ACCESO DESDE OTROS DISPOSITIVOS (Misma Red WiFi / LAN):")
        for ip in ips:
            print(f"   👉 http://{ip}:{port}")
    else:
        print(f"\n📡 Para acceder desde otra PC en tu misma red: http://TU_IP_LOCAL:{port}")
        
    print("\n🔄 Modo DEBUG activado: los cambios se recargan automáticamente al guardar.")
    print("🛑 Presiona CTRL+C para detener el servidor.")
    print("=" * 60 + "\n")
    
    # Modo debug: recarga automática al guardar cambios
    app.run(debug=True, host='0.0.0.0', port=port)