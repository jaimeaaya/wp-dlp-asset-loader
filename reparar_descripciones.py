"""
reparar_descripciones.py
Actualiza post_excerpt de todos los dlp_document con el texto literal
de la columna DESCRIPCION (Col D) del Excel, sumando la sección Archivos.

Formato resultante en post_excerpt:
  {descripcion}<br><br>Archivos:<br>{formato}

Uso:
  python reparar_descripciones.py             # preview de los primeros 5 cambios
  python reparar_descripciones.py --aplicar   # aplica todos los cambios
  python reparar_descripciones.py --solo hero # solo una marca
"""

import re, csv, os, sys
import openpyxl, pymysql
from sshtunnel import SSHTunnelForwarder
from dotenv import load_dotenv

load_dotenv()

EXCEL_PATH = 'Listado de Assets - PD.xlsx'
SKIP_HOJAS = {'CRONOGRAMA'}

SSH = dict(
    host     = os.environ['SSH_HOST'],
    port     = int(os.environ['SSH_PORT']),
    username = os.environ['SSH_USERNAME'],
    password = os.environ['SSH_PASSWORD'],
)
DB = dict(
    user     = os.environ['DB_USER'],
    password = os.environ['DB_PASSWORD'],
    db       = os.environ['DB_NAME'],
    charset  = 'utf8mb4',
)

COL = dict(nombre=0, categoria=1, subcategoria=2, descripcion=3, formato=4,
           img_principal=5, preview=6, editables=7, piezas=8, consolidado=9, hoja=10)

LANG_TTI = {'Español': 3, 'English': 6}


def limpiar(val):
    return '' if val is None else str(val).strip()

def extraer_url_amplify(texto):
    if not texto:
        return ''
    m = re.search(r'https://www\.amplify\.churchdwight\.com/transfer/[a-f0-9]+', texto)
    return m.group(0) if m else ''

def build_excerpt(descripcion, formato):
    desc = (descripcion or '').strip()
    fmt  = (formato or '').strip()
    if not desc and not fmt:
        return ''
    if not fmt:
        return desc
    return f"{desc}<br><br>Archivos:<br>{fmt}"


def leer_consolidado(path):
    wb  = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws  = wb['Consolidado']
    rows = list(ws.iter_rows(values_only=True))

    docs = {}   # (amplify_url, idioma) -> {nombre, descripcion, formato, hoja}
    orden = []

    for row in rows[1:]:
        if not any(row):
            continue
        nombre      = limpiar(row[COL['nombre']])
        hoja        = limpiar(row[COL['hoja']])
        consolidado = limpiar(row[COL['consolidado']])
        editables   = limpiar(row[COL['editables']])
        if hoja in SKIP_HOJAS or not nombre:
            continue

        url    = extraer_url_amplify(consolidado) or extraer_url_amplify(editables)
        idioma = 'English' if 'English' in hoja else 'Español'
        clave  = (url, idioma)
        if not url or clave in docs:
            continue

        orden.append(clave)
        docs[clave] = {
            'nombre':      nombre,
            'hoja':        hoja,
            'idioma':      idioma,
            'amplify_url': url,
            'descripcion': limpiar(row[COL['descripcion']]),
            'formato':     limpiar(row[COL['formato']]),
        }

    wb.close()
    return [docs[k] for k in orden]


def get_connection():
    tunnel = SSHTunnelForwarder(
        (SSH['host'], SSH['port']),
        ssh_username=SSH['username'], ssh_password=SSH['password'],
        remote_bind_address=('127.0.0.1', 3306),
    )
    tunnel.start()
    conn = pymysql.connect(host='127.0.0.1', port=tunnel.local_bind_port, **DB)
    return tunnel, conn


def main():
    aplicar    = '--aplicar' in sys.argv
    solo_marca = None
    if '--solo' in sys.argv:
        idx = sys.argv.index('--solo')
        if idx + 1 < len(sys.argv):
            solo_marca = sys.argv[idx + 1].lower()

    if not aplicar:
        print("=== MODO PREVIEW — mostrando primeras diferencias, no se escribe nada ===")
        print("    Usar --aplicar para actualizar la DB.\n")

    print("Leyendo Excel...")
    docs = leer_consolidado(EXCEL_PATH)
    if solo_marca:
        docs = [d for d in docs if solo_marca in d['hoja'].lower()]
    print(f"  Documentos en Excel: {len(docs)}")

    print("Conectando via SSH tunnel...")
    tunnel, conn = get_connection()

    cur = conn.cursor()

    actualizados = 0
    sin_cambio   = 0
    no_encontrado = 0
    preview_count = 0

    for doc in docs:
        url    = doc['amplify_url']
        tti    = LANG_TTI[doc['idioma']]
        nuevo_excerpt = build_excerpt(doc['descripcion'], doc['formato'])

        # Buscar el post en DB por URL + idioma
        cur.execute("""
            SELECT pm.post_id, p.post_excerpt
            FROM wp_postmeta pm
            JOIN wp_posts p ON p.ID = pm.post_id
            JOIN wp_term_relationships tr ON tr.object_id = pm.post_id
            WHERE pm.meta_key = '_dlp_direct_link_url' AND pm.meta_value = %s
              AND tr.term_taxonomy_id = %s
              AND p.post_type = 'dlp_document' AND p.post_status = 'publish'
            LIMIT 1
        """, (url, tti))
        row = cur.fetchone()

        if not row:
            no_encontrado += 1
            continue

        post_id, excerpt_actual = row

        if excerpt_actual == nuevo_excerpt:
            sin_cambio += 1
            continue

        if not aplicar:
            if preview_count < 5:
                print(f"\npost_id={post_id} | {doc['idioma']} | {doc['nombre']}")
                print(f"  ACTUAL:  {repr(excerpt_actual[:120])}")
                print(f"  NUEVO:   {repr(nuevo_excerpt[:120])}")
                preview_count += 1
            actualizados += 1
            continue

        cur.execute("UPDATE wp_posts SET post_excerpt=%s WHERE ID=%s",
                    (nuevo_excerpt, post_id))
        conn.commit()
        actualizados += 1

        if actualizados % 50 == 0:
            print(f"  Actualizados: {actualizados}...")

    conn.close()
    tunnel.stop()

    print(f"\nResultado:")
    print(f"  Actualizados:    {actualizados}")
    print(f"  Sin cambio:      {sin_cambio}")
    print(f"  No encontrados:  {no_encontrado}")

    if not aplicar and actualizados > 0:
        print(f"\n  -> Correr con --aplicar para escribir los {actualizados} cambios en la DB.")


if __name__ == '__main__':
    main()
