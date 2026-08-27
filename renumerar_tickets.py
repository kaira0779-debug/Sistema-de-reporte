import pymysql
from config import Config

# Conexión a la base de datos usando la configuración existente
conn = pymysql.connect(
    host=Config.MYSQL_HOST,
    port=Config.MYSQL_PORT,
    user=Config.MYSQL_USER,
    password=Config.MYSQL_PASSWORD,
    database=Config.MYSQL_DB,
    cursorclass=pymysql.cursors.DictCursor
)

try:
    with conn.cursor() as cur:
        # Obtener todos los tickets del año 2026 ordenados por fecha y luego por id
        cur.execute("""
            SELECT id, num_ticket, fecha
            FROM tickets
            WHERE num_ticket LIKE '2026-5%' AND CAST(SUBSTRING_INDEX(num_ticket,'-',-1) AS UNSIGNED) >= 5006
            ORDER BY fecha ASC, id ASC
        """)
        tickets = cur.fetchall()

        if not tickets:
            print("No hay tickets del año 2026 para renumerar.")
            exit()

        # Mostrar cantidad encontrada
        print(f"Se encontraron {len(tickets)} tickets del año 2026.")

        # Preguntar confirmación
        confirmacion = input("¿Desea renumerar todos estos tickets comenzando en 2026-5000? (s/n): ")
        if confirmacion.lower() != 's':
            print("Operación cancelada.")
            exit()

        # Asignar nuevos números secuencialmente
        nuevo_numero = 5000
        for ticket in tickets:
            nuevo_num_ticket = f"2026-{nuevo_numero}"
            cur.execute("""
                UPDATE tickets SET num_ticket = %s WHERE id = %s
            """, (nuevo_num_ticket, ticket['id']))
            print(f"  {ticket['num_ticket']} -> {nuevo_num_ticket}")
            nuevo_numero += 1

        # Actualizar la secuencia para futuros tickets
        ultimo_numero = nuevo_numero - 1  # El último asignado
        cur.execute("""
            INSERT INTO ticket_sequences (year, next_num) VALUES (2026, %s)
            ON DUPLICATE KEY UPDATE next_num = %s
        """, (ultimo_numero, ultimo_numero))
        print(f"\nSecuencia actualizada: el próximo ticket será 2026-{nuevo_numero}")

        conn.commit()
        print("✅ Renumeración completada exitosamente.")

except Exception as e:
    conn.rollback()
    print(f"❌ Error: {e}")
finally:
    conn.close()