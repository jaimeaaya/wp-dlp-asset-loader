"""
diagnostico_assets.py
Lee la hoja Consolidado del Excel, agrupa por URL de CONSOLIDADO,
cruza contra la DB y produce un CSV con el estado de cada documento.

Ejecutar en el servidor:
  python3 diagnostico_assets.py
"""

import re
import csv
import json
import os
import openpyxl
import pymysql
from sshtunnel import SSHTunnelForwarder
from dotenv import load_dotenv

load_dotenv()

# ── Configuración ─────────────────────────────────────────────────────────────
EXCEL_PATH = 'Listado de Assets - PD.xlsx'
OUTPUT_CSV = 'diagnostico_assets.csv'

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

WP_PATH = os.environ['WP_PATH']

# Índices de columnas en la hoja Consolidado (0-based)
COL = dict(
    nombre       = 0,
    categoria    = 1,
    subcategoria = 2,
    descripcion  = 3,
    formato      = 4,
    img_principal= 5,
    preview      = 6,
    editables    = 7,
    piezas       = 8,
    consolidado  = 9,
    hoja         = 10,
)

# Hojas que no son documentos
SKIP_HOJAS = {'CRONOGRAMA'}

# ── Helpers ───────────────────────────────────────────────────────────────────

def limpiar(val):
    if val is None:
        return ''
    return str(val).strip()

def extraer_url_amplify(texto):
    """Extrae la primera URL de amplify.churchdwight.com del texto."""
    if not texto:
        return ''
    match = re.search(r'https://www\.amplify\.churchdwight\.com/transfer/[a-f0-9]+', texto)
    return match.group(0) if match else ''

def extraer_titulo_pieza(texto):
    """
    De 'Post 1: https://...' extrae 'Post 1'.
    Si no hay prefijo, devuelve cadena vacía.
    """
    if not texto:
        return ''
    match = re.match(r'^([^:]+):\s*https://', texto.strip())
    return match.group(1).strip() if match else ''

def normalizar_marca(hoja):
    """'Hero - English' → 'Hero', 'A&H' → 'arm-hammer', etc."""
    nombre = hoja.replace(' - English', '').strip()
    slugs = {
        'A&H':          'arm-hammer',
        'Stérimar':     'sterimar',
        'Stérimar': 'sterimar',
        'Brand Camp':   'brand-camp',
        'Ecommerce':    'ecommerce',
        'Elementos globales': 'elementos-globales',
    }
    return slugs.get(nombre, nombre.lower().replace(' ', '-'))


# ── Lectura del Excel ─────────────────────────────────────────────────────────

def leer_consolidado(path):
    """
    Lee la hoja Consolidado y devuelve una lista de grupos.
    Cada grupo = un dlp_document con sus filas agrupadas.
    """
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Consolidado']
    rows = list(ws.iter_rows(values_only=True))

    # Saltar fila de cabecera (row 0)
    grupos  = {}   # (consolidado_url, idioma) → dict del documento
    orden   = []   # mantener orden de aparición

    for row in rows[1:]:
        if not any(row):
            continue

        nombre      = limpiar(row[COL['nombre']])
        hoja        = limpiar(row[COL['hoja']])
        preview     = limpiar(row[COL['preview']])
        editables   = limpiar(row[COL['editables']])
        consolidado = limpiar(row[COL['consolidado']])

        if hoja in SKIP_HOJAS or not nombre:
            continue

        url_consolidado = extraer_url_amplify(consolidado)
        url_editables   = extraer_url_amplify(editables)

        # Clave de agrupación: (URL, idioma) — misma URL puede existir como ES y EN
        idioma = 'English' if 'English' in hoja else 'Español'
        clave  = (url_consolidado or url_editables, idioma)
        if not clave[0]:
            continue

        if clave not in grupos:
            orden.append(clave)
            grupos[clave] = {
                'nombre':         nombre,
                'hoja':           hoja,
                'marca':          normalizar_marca(hoja),
                'idioma':         idioma,
                'categoria':      limpiar(row[COL['categoria']]),
                'subcategoria':   limpiar(row[COL['subcategoria']]),
                'descripcion':    limpiar(row[COL['descripcion']]),
                'formato':        limpiar(row[COL['formato']]),
                'img_principal':  limpiar(row[COL['img_principal']]),
                'amplify_url':    clave[0],
                'previews':       [],
                'single_assets':  [],
            }

        doc = grupos[clave]

        # Acumular preview_slider
        if preview:
            doc['previews'].append(preview)

        # Acumular single_assets (cuando editables ≠ consolidado)
        if url_editables and url_editables != url_consolidado:
            titulo_pieza = extraer_titulo_pieza(editables)
            doc['single_assets'].append({
                'titulo': titulo_pieza,
                'url':    url_editables,
            })

    wb.close()
    return [grupos[k] for k in orden]


# ── Conexión vía túnel SSH ────────────────────────────────────────────────────

def get_connection():
    tunnel = SSHTunnelForwarder(
        (SSH['host'], SSH['port']),
        ssh_username        = SSH['username'],
        ssh_password        = SSH['password'],
        remote_bind_address = ('127.0.0.1', 3306),
    )
    tunnel.start()
    conn = pymysql.connect(
        host    = '127.0.0.1',
        port    = tunnel.local_bind_port,
        **DB,
    )
    return tunnel, conn


# ── Cruce con la DB ───────────────────────────────────────────────────────────

LANG_TTI = {'Español': 3, 'English': 6}

def cargar_db_urls(conn):
    """Devuelve {(amplify_url, lang_tti): post_id} para todos los dlp_document publicados.
    La clave incluye idioma para que el mismo URL pueda existir como ES y como EN.
    """
    cur = conn.cursor()
    cur.execute("""
        SELECT p.ID, pm.meta_value, tr.term_taxonomy_id
        FROM wp_posts p
        JOIN wp_postmeta pm ON pm.post_id = p.ID
        JOIN wp_term_relationships tr ON tr.object_id = p.ID
        WHERE p.post_type   = 'dlp_document'
          AND p.post_status = 'publish'
          AND pm.meta_key   = '_dlp_direct_link_url'
          AND pm.meta_value != ''
          AND tr.term_taxonomy_id IN (3, 6)
    """)
    return {(row[1].strip(), row[2]): row[0] for row in cur.fetchall()}


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("Leyendo Excel…")
    documentos = leer_consolidado(EXCEL_PATH)
    print(f"  Documentos identificados: {len(documentos)}")

    print("Conectando vía SSH tunnel…")
    tunnel, conn = get_connection()
    db_urls      = cargar_db_urls(conn)
    conn.close()
    tunnel.stop()
    print(f"  Registros en DB: {len(db_urls)}")

    # Diagnóstico
    nuevo   = 0
    existe  = 0
    sin_url = 0

    filas = []
    for doc in documentos:
        url     = doc['amplify_url']
        tti     = LANG_TTI.get(doc['idioma'], 3)
        post_id = db_urls.get((url, tti)) if url else None
        if not url:
            estado  = 'SIN_URL'
            sin_url += 1
        elif post_id:
            estado  = 'EXISTE'
            existe  += 1
        else:
            estado  = 'NUEVO'
            nuevo   += 1

        filas.append({
            'estado':           estado,
            'post_id':          post_id or '',
            'hoja':             doc['hoja'],
            'marca':            doc['marca'],
            'idioma':           doc['idioma'],
            'nombre':           doc['nombre'],
            'categoria':        doc['categoria'],
            'subcategoria':     doc['subcategoria'],
            'formato':          doc['formato'],
            'amplify_url':      url,
            'descripcion':      doc['descripcion'],
            'img_principal':    doc['img_principal'],
            'previews_count':   len(doc['previews']),
            'previews':         ' | '.join(doc['previews']),
            'single_assets_count': len(doc['single_assets']),
            'single_assets':    json.dumps(doc['single_assets'], ensure_ascii=False),
        })

    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=filas[0].keys())
        writer.writeheader()
        writer.writerows(filas)

    print(f"\nResultado -> {OUTPUT_CSV}")
    print(f"  NUEVO:   {nuevo}")
    print(f"  EXISTE:  {existe}")
    print(f"  SIN_URL: {sin_url}")
    print(f"  TOTAL:   {len(filas)}")


if __name__ == '__main__':
    main()
