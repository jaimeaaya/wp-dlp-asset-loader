"""
validacion_cargue.py
Cruza la hoja Consolidado del Excel contra la DB de WordPress.
Para cada dlp_document verifica:
  - post existe y está publicado
  - _thumbnail_id seteado (imagen principal)
  - preview_slider_N: cantidad coincide con Excel
  - asset_N: cantidad coincide con Excel
  - _dlp_document_visibility meta presente (requerido por DLP)
  - _dlp_document_file_size   meta presente (requerido por DLP)
  - file_type term tti=14 asignado (requerido por DLP para docs URL)
  - par Polylang (post_translations) si existe la versión opuesta

Genera:
  validacion_cargue.csv  — fila por documento con estado y problemas
  vincular_traducciones.sql — parches SQL para pares Polylang faltantes (si hay)

Uso:
  python validacion_cargue.py
  python validacion_cargue.py --reparar   # además ejecuta el SQL de vinculacion
"""

import re, csv, json, os, sys, time, hashlib
import openpyxl, pymysql
from sshtunnel import SSHTunnelForwarder
from dotenv import load_dotenv

load_dotenv()

EXCEL_PATH  = 'Listado de Assets - PD.xlsx'
OUTPUT_CSV  = 'validacion_cargue.csv'
OUTPUT_SQL  = 'vincular_traducciones.sql'
SKIP_HOJAS  = {'CRONOGRAMA'}

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

COL = dict(
    nombre=0, categoria=1, subcategoria=2, descripcion=3, formato=4,
    img_principal=5, preview=6, editables=7, piezas=8, consolidado=9, hoja=10,
)
LANG_TTI  = {'Español': 3, 'English': 6}
FILE_TYPE_URL_TTI = 14


# ── Helpers ───────────────────────────────────────────────────────────────────

def limpiar(val):
    return '' if val is None else str(val).strip()

def extraer_url_amplify(texto):
    if not texto:
        return ''
    m = re.search(r'https://www\.amplify\.churchdwight\.com/transfer/[a-f0-9]+', texto)
    return m.group(0) if m else ''

def extraer_titulo_pieza(texto):
    if not texto:
        return ''
    m = re.match(r'^([^:]+):\s*https://', texto.strip())
    return m.group(1).strip() if m else ''


# ── Lectura Excel (misma lógica que diagnostico) ──────────────────────────────

def leer_consolidado(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb['Consolidado']
    rows = list(ws.iter_rows(values_only=True))

    grupos, orden = {}, []
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

        url_c = extraer_url_amplify(consolidado)
        url_e = extraer_url_amplify(editables)
        idioma = 'English' if 'English' in hoja else 'Español'
        clave  = (url_c or url_e, idioma)
        if not clave[0]:
            continue

        if clave not in grupos:
            orden.append(clave)
            grupos[clave] = {
                'nombre': nombre, 'hoja': hoja, 'idioma': idioma,
                'amplify_url': clave[0],
                'img_principal': limpiar(row[COL['img_principal']]),
                'previews': [], 'single_assets': [],
            }
        doc = grupos[clave]
        if preview:
            doc['previews'].append(preview)
        if url_e and url_e != url_c:
            doc['single_assets'].append({'titulo': extraer_titulo_pieza(editables), 'url': url_e})

    wb.close()
    return [grupos[k] for k in orden]


# ── Conexión ──────────────────────────────────────────────────────────────────

def get_connection():
    tunnel = SSHTunnelForwarder(
        (SSH['host'], SSH['port']),
        ssh_username=SSH['username'], ssh_password=SSH['password'],
        remote_bind_address=('127.0.0.1', 3306),
    )
    tunnel.start()
    conn = pymysql.connect(host='127.0.0.1', port=tunnel.local_bind_port, **DB)
    return tunnel, conn


# ── Carga masiva de datos DB ──────────────────────────────────────────────────

def cargar_posts_por_url(conn, urls_tti):
    """
    urls_tti: lista de (amplify_url, lang_tti)
    Devuelve {(url, lang_tti): post_id}
    """
    if not urls_tti:
        return {}
    cur = conn.cursor()
    url_list = list({u for u, _ in urls_tti})
    placeholders = ','.join(['%s'] * len(url_list))
    cur.execute(f"""
        SELECT pm.post_id, pm.meta_value, tr.term_taxonomy_id
        FROM wp_posts p
        JOIN wp_postmeta pm ON pm.post_id = p.ID AND pm.meta_key = '_dlp_direct_link_url'
        JOIN wp_term_relationships tr ON tr.object_id = p.ID AND tr.term_taxonomy_id IN (3, 6)
        WHERE p.post_type = 'dlp_document' AND p.post_status = 'publish'
          AND pm.meta_value IN ({placeholders})
    """, url_list)
    result = {}
    for post_id, meta_value, tti in cur.fetchall():
        result[(meta_value.strip(), tti)] = post_id
    return result


def cargar_meta_posts(conn, post_ids):
    """Devuelve {post_id: {meta_key: [meta_value, ...]}}"""
    if not post_ids:
        return {}
    ids_ph = ','.join(['%s'] * len(post_ids))
    cur = conn.cursor()
    cur.execute(f"""
        SELECT post_id, meta_key, meta_value FROM wp_postmeta
        WHERE post_id IN ({ids_ph})
    """, list(post_ids))
    result = {}
    for post_id, key, val in cur.fetchall():
        result.setdefault(post_id, {}).setdefault(key, []).append(val)
    return result


def cargar_term_relationships(conn, post_ids):
    """Devuelve {post_id: set(term_taxonomy_ids)}"""
    if not post_ids:
        return {}
    ids_ph = ','.join(['%s'] * len(post_ids))
    cur = conn.cursor()
    cur.execute(f"""
        SELECT object_id, term_taxonomy_id FROM wp_term_relationships
        WHERE object_id IN ({ids_ph})
    """, list(post_ids))
    result = {}
    for oid, tti in cur.fetchall():
        result.setdefault(oid, set()).add(tti)
    return result


def cargar_polylang_pairs(conn, post_ids):
    """
    Devuelve {post_id: set(post_ids del mismo grupo post_translations)}
    """
    if not post_ids:
        return {}
    ids_ph = ','.join(['%s'] * len(post_ids))
    cur = conn.cursor()
    cur.execute(f"""
        SELECT tr1.object_id, tr2.object_id
        FROM wp_term_relationships tr1
        JOIN wp_term_relationships tr2
            ON tr1.term_taxonomy_id = tr2.term_taxonomy_id AND tr1.object_id != tr2.object_id
        JOIN wp_term_taxonomy tt ON tt.term_taxonomy_id = tr1.term_taxonomy_id
            AND tt.taxonomy = 'post_translations'
        WHERE tr1.object_id IN ({ids_ph})
    """, list(post_ids))
    result = {}
    for pid, partner in cur.fetchall():
        result.setdefault(pid, set()).add(partner)
    return result


# ── Reparar pares Polylang ────────────────────────────────────────────────────

def reparar_polylang(conn, pares_faltantes):
    """
    pares_faltantes: lista de (es_id, en_id)
    Crea los términos post_translations y los vincula.
    """
    cur = conn.cursor()
    reparados = 0
    for es_id, en_id in pares_faltantes:
        slug = 'pll_' + hashlib.md5(f"{es_id}{en_id}{time.time()}".encode()).hexdigest()[:13]
        cur.execute("INSERT INTO wp_terms (name, slug, term_group) VALUES (%s,%s,0)", (slug, slug))
        term_id = cur.lastrowid
        cur.execute(
            "INSERT INTO wp_term_taxonomy (term_id, taxonomy, description, parent, count) VALUES (%s,'post_translations','',0,0)",
            (term_id,))
        tti = cur.lastrowid
        for pid in (es_id, en_id):
            cur.execute(
                "INSERT IGNORE INTO wp_term_relationships (object_id, term_taxonomy_id, term_order) VALUES (%s,%s,0)",
                (pid, tti))
        cur.execute("UPDATE wp_term_taxonomy SET count=2 WHERE term_taxonomy_id=%s", (tti,))
        conn.commit()
        reparados += 1
        print(f"  Polylang vinculado: ES={es_id} <-> EN={en_id}")
    return reparados


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    reparar = '--reparar' in sys.argv

    print("Leyendo Excel...")
    docs = leer_consolidado(EXCEL_PATH)
    print(f"  Documentos en Excel: {len(docs)}")

    print("Conectando via SSH tunnel...")
    tunnel, conn = get_connection()

    # Cargar post_ids por (url, lang_tti)
    urls_tti = [(d['amplify_url'], LANG_TTI[d['idioma']]) for d in docs]
    url_to_post = cargar_posts_por_url(conn, urls_tti)

    # Obtener todos los post_ids encontrados
    all_post_ids = set(url_to_post.values())
    print(f"  Posts encontrados en DB: {len(all_post_ids)}")

    # Cargar meta y terms de todos los posts de una vez
    meta_db    = cargar_meta_posts(conn, all_post_ids)
    terms_db   = cargar_term_relationships(conn, all_post_ids)
    pll_db     = cargar_polylang_pairs(conn, all_post_ids)

    # Cruce por idioma inverso (para buscar par Polylang faltante)
    # url -> {es: post_id, en: post_id}
    url_lang_map = {}
    for (url, tti), pid in url_to_post.items():
        lang = 'es' if tti == 3 else 'en'
        url_lang_map.setdefault(url, {})[lang] = pid

    # ── Validación ────────────────────────────────────────────────────────────
    filas = []
    pares_faltantes = []

    ok_count = warn_count = error_count = missing_count = 0

    for doc in docs:
        url    = doc['amplify_url']
        idioma = doc['idioma']
        tti    = LANG_TTI[idioma]
        lang   = 'en' if idioma == 'English' else 'es'
        post_id = url_to_post.get((url, tti))

        issues = []
        estado = 'OK'

        if not post_id:
            filas.append({
                'estado': 'FALTANTE',
                'post_id': '',
                'idioma': idioma,
                'hoja': doc['hoja'],
                'nombre': doc['nombre'],
                'amplify_url': url,
                'thumbnail_id': '',
                'previews_excel': len(doc['previews']),
                'previews_db': 0,
                'assets_excel': len(doc['single_assets']),
                'assets_db': 0,
                'has_visibility_meta': False,
                'has_filesize_meta': False,
                'has_filetype_term': False,
                'has_polylang_pair': False,
                'issues': 'Post no encontrado en DB',
            })
            missing_count += 1
            continue

        meta  = meta_db.get(post_id, {})
        terms = terms_db.get(post_id, set())
        pll   = pll_db.get(post_id, set())

        # 1. Thumbnail
        thumb_vals = meta.get('_thumbnail_id', [])
        thumb_id   = thumb_vals[0] if thumb_vals else ''
        has_thumb  = bool(thumb_id and thumb_id != '0')
        if doc['img_principal'] and not has_thumb:
            issues.append('SIN_THUMBNAIL')

        # 2. Previews
        preview_keys   = [k for k in meta if k.startswith('preview_slider_')]
        previews_db    = len(preview_keys)
        previews_excel = len(doc['previews'])
        if previews_excel > 0 and previews_db == 0:
            issues.append(f'SIN_PREVIEWS (Excel:{previews_excel} DB:0)')
        elif previews_db != previews_excel:
            issues.append(f'PREVIEWS_MISMATCH (Excel:{previews_excel} DB:{previews_db})')

        # 3. Single assets
        asset_keys   = [k for k in meta if re.match(r'^asset_\d+$', k)]
        assets_db    = len(asset_keys)
        assets_excel = len(doc['single_assets'])
        if assets_excel > 0 and assets_db == 0:
            issues.append(f'SIN_ASSETS (Excel:{assets_excel} DB:0)')
        elif assets_db != assets_excel:
            issues.append(f'ASSETS_MISMATCH (Excel:{assets_excel} DB:{assets_db})')

        # 4. Metas requeridos por DLP
        has_vis      = '_dlp_document_visibility' in meta
        has_filesize = '_dlp_document_file_size' in meta
        if not has_vis:
            issues.append('FALTA_VISIBILITY_META')
        if not has_filesize:
            issues.append('FALTA_FILESIZE_META')

        # 5. file_type term (tti=14)
        has_filetype = FILE_TYPE_URL_TTI in terms
        if not has_filetype:
            issues.append('FALTA_FILETYPE_TERM(14)')

        # 6. Par Polylang
        pair_lang = 'en' if lang == 'es' else 'es'
        partner_id = url_lang_map.get(url, {}).get(pair_lang)
        has_pll = bool(partner_id and partner_id in pll)
        if partner_id and not has_pll:
            issues.append(f'SIN_POLYLANG_PAIR (partner={partner_id})')
            # Solo agregar una vez por par
            pair_tuple = (min(post_id, partner_id), max(post_id, partner_id))
            if pair_tuple not in [(min(a, b), max(a, b)) for a, b in pares_faltantes]:
                if lang == 'es':
                    pares_faltantes.append((post_id, partner_id))
                else:
                    pares_faltantes.append((partner_id, post_id))

        # Determinar estado final
        critical = {'SIN_THUMBNAIL', 'FALTA_VISIBILITY_META', 'FALTA_FILESIZE_META', 'FALTA_FILETYPE_TERM(14)'}
        has_critical  = any(any(c in iss for c in critical) for iss in issues)
        has_warning   = bool(issues)

        if has_critical:
            estado = 'ERROR'
            error_count += 1
        elif has_warning:
            estado = 'WARN'
            warn_count += 1
        else:
            estado = 'OK'
            ok_count += 1

        filas.append({
            'estado': estado,
            'post_id': post_id,
            'idioma': idioma,
            'hoja': doc['hoja'],
            'nombre': doc['nombre'],
            'amplify_url': url,
            'thumbnail_id': thumb_id,
            'previews_excel': previews_excel,
            'previews_db': previews_db,
            'assets_excel': assets_excel,
            'assets_db': assets_db,
            'has_visibility_meta': has_vis,
            'has_filesize_meta': has_filesize,
            'has_filetype_term': has_filetype,
            'has_polylang_pair': has_pll,
            'issues': ' | '.join(issues) if issues else '',
        })

    conn.close()
    tunnel.stop()

    # ── Output CSV ────────────────────────────────────────────────────────────
    with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=filas[0].keys())
        writer.writeheader()
        writer.writerows(filas)

    print(f"\nResultado -> {OUTPUT_CSV}")
    print(f"  OK:       {ok_count}")
    print(f"  WARN:     {warn_count}")
    print(f"  ERROR:    {error_count}")
    print(f"  FALTANTE: {missing_count}")
    print(f"  TOTAL:    {len(filas)}")
    print(f"\n  Pares Polylang faltantes: {len(pares_faltantes)}")

    # ── Reparar Polylang ──────────────────────────────────────────────────────
    if pares_faltantes:
        if reparar:
            print("\nReparando pares Polylang...")
            tunnel2, conn2 = get_connection()
            reparados = reparar_polylang(conn2, pares_faltantes)
            conn2.close()
            tunnel2.stop()
            print(f"  Reparados: {reparados}")
        else:
            print(f"  -> Correr con --reparar para vincularlos automaticamente")
            print(f"  -> O revisar {OUTPUT_SQL} para aplicar manualmente")
            # Generar SQL de referencia
            with open(OUTPUT_SQL, 'w', encoding='utf-8') as f:
                f.write("-- Pares Polylang faltantes. Usar --reparar para aplicar automaticamente.\n")
                for es_id, en_id in pares_faltantes:
                    f.write(f"-- ES={es_id} <-> EN={en_id}\n")
            print(f"  IDs en {OUTPUT_SQL}")


if __name__ == '__main__':
    main()
