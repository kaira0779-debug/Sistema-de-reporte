# limpiar_duplicados_tickets.py
"""
Detecta y elimina tickets duplicados (misma fecha + cliente + motivo).

USO:
  python limpiar_duplicados_tickets.py              -> solo muestra (no modifica)
  python limpiar_duplicados_tickets.py --ejecutar   -> elimina los duplicados
"""
import sys
import pymysql
import re
from config import Config

def conectar():
    return pymysql.connect(
        host=Config.MYSQL_HOST,
        port=Config.MYSQL_PORT,
        user=Config.MYSQL_USER,
        password=Config.MYSQL_PASSWORD,
        database=Config.MYSQL_DB,
        cursorclass=pymysql.cursors.DictCursor
    )

def hallar_grupos(cur):
    """Agrupa tickets con misma fecha + cliente + motivo_ticket."""
    cur.execute("""
        SELECT 
            DATE(fecha) AS fecha_dia,
            cliente,
            motivo_ticket,
            COUNT(*) AS cnt,
            GROUP_CONCAT(id ORDER BY id ASC) AS ids
        FROM tickets
        WHERE cliente IS NOT NULL AND cliente != ''
          AND motivo_ticket IS NOT NULL AND motivo_ticket != ''
        GROUP BY DATE(fecha), cliente, motivo_ticket
        HAVING cnt > 1
        ORDER BY fecha_dia DESC, cliente ASC
    """)
    return cur.fetchall()

def elegir_a_conservar(detalles):
    """
    Prefiere conservar el ticket con menor número (el original).
    Descarta los auto-numerados que tienen números más altos.
    """
    def score(t):
        nt = str(t.get('num_ticket') or '')
        m = re.match(r'^(\d{4})-(\d+)$', nt)
        if m:
            return int(m.group(2))
        return 99999999
    return sorted(detalles, key=score)[0]['id']

def main():
    ejecutar = '--ejecutar' in sys.argv
    print("=" * 70)
    print("LIMPIEZA DE TICKETS DUPLICADOS")
    print(f"Modo: {'EJECUTAR' if ejecutar else 'SOLO ANÁLISIS'}")
    print("=" * 70)

    conn = conectar()
    try:
        with conn.cursor() as cur:
            cur.execute("SELECT COUNT(*) AS total FROM tickets")
            total = cur.fetchone()['total']
            print(f"Tickets actuales: {total}\n")

            grupos = hallar_grupos(cur)
            print(f"Grupos con posibles duplicados: {len(grupos)}\n")

            if not grupos:
                print("✅ No hay duplicados detectados.")
                return

            a_eliminar = sum(g['cnt'] - 1 for g in grupos)
            print(f"Tickets a eliminar: {a_eliminar}")
            print(f"Tickets quedarán:   {total - a_eliminar}\n")
            print("-" * 70)
            print("VISTA PREVIA (primeros 20 grupos)")
            print("-" * 70)

            for i, g in enumerate(grupos[:20], 1):
                fecha_str = g['fecha_dia'].strftime('%d/%m/%Y') if g['fecha_dia'] else 'Sin fecha'
                ids = [int(x) for x in g['ids'].split(',')]
                ph = ','.join(['%s'] * len(ids))
                cur.execute(f"""
                    SELECT id, num_ticket, tipo_falla, status
                    FROM tickets WHERE id IN ({ph}) ORDER BY id ASC
                """, ids)
                detalles = cur.fetchall()
                conservar = elegir_a_conservar(detalles)

                print(f"\n[{i}] {fecha_str} | {g['cliente']}")
                print(f"    Motivo: {g['motivo_ticket']}  ({g['cnt']} registros)")
                for d in detalles:
                    tag = "CONSERVAR" if d['id'] == conservar else "ELIMINAR"
                    print(f"      [{tag}] ID {d['id']}: {d['num_ticket']}")

            if len(grupos) > 20:
                print(f"\n... y {len(grupos) - 20} grupos más")

            if not ejecutar:
                print("\n" + "=" * 70)
                print("MODO ANÁLISIS: no se modificó nada.")
                print("Para eliminar, ejecuta:")
                print("   python limpiar_duplicados_tickets.py --ejecutar")
                print("=" * 70)
                return

            print("\n⚠️  Asegúrate de tener un BACKUP de la base de datos.")
            resp = input("Escribe 'SI' para confirmar eliminación: ")
            if resp.strip().upper() != 'SI':
                print("Cancelado.")
                return

            eliminados = 0
            for g in grupos:
                ids = [int(x) for x in g['ids'].split(',')]
                ph = ','.join(['%s'] * len(ids))
                cur.execute(f"""
                    SELECT id, num_ticket FROM tickets WHERE id IN ({ph})
                """, ids)
                detalles = cur.fetchall()
                conservar = elegir_a_conservar(detalles)

                for d in detalles:
                    if d['id'] != conservar:
                        cur.execute("DELETE FROM ticket_tecnicos WHERE ticket_id = %s", (d['id'],))
                        cur.execute("DELETE FROM tickets WHERE id = %s", (d['id'],))
                        eliminados += 1

            conn.commit()
            print(f"\n✅ Eliminados: {eliminados} tickets duplicados.")
            cur.execute("SELECT COUNT(*) AS total FROM tickets")
            print(f"Tickets restantes: {cur.fetchone()['total']}")

    except Exception as e:
        conn.rollback()
        print(f"\n❌ Error: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()

if __name__ == '__main__':
    main()