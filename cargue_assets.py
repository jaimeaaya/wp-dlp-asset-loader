"""
cargue_assets.py
Carga documentos NUEVOS desde diagnostico_assets.csv a WordPress.
Orden: ES primero, luego EN, finalmente vincula pares Polylang.

Uso:
  python cargue_assets.py             # carga todos los NUEVO
  python cargue_assets.py --dry-run   # solo muestra lo que haría
  python cargue_assets.py --solo Hero # solo una marca
"""

import re
import csv
import json
import os
import sys
import time
import logging
import unicodedata
import pymysql
import paramiko
from sshtunnel import SSHTunnelForwarder
from dotenv import load_dotenv

load_dotenv()

# ── Config ────────────────────────────────────────────────────────────────────
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

WP_PATH        = os.environ['WP_PATH']
WP_URL         = os.environ['WP_URL']
DEFAULT_AUTHOR = int(os.getenv('DEFAULT_AUTHOR', '3'))
INPUT_CSV      = 'diagnostico_assets.csv'
LOG_FILE       = 'cargue_assets.log'
ASSET_FILE_SIZE = '5 MB'

# ── Mappings taxonomía ────────────────────────────────────────────────────────
LANG_TTI        = {'es': 3, 'en': 6}     # term_taxonomy_id de language
FILE_TYPE_URL_TTI = 14                   # file_type taxonomy: "www" (term para documentos URL)

# brand slug → term_taxonomy_id (doc_categories bajo 'brands')
BRAND_ES = {
    'hero': 30, 'arm-hammer': 31, 'batiste': 34, 'oxiclean': 33,
    'sterimar': 35, 'therabreath': 36, 'trojan': 37, 'waterpik': 38, 'nair': 75,
}
# brand slug → term_taxonomy_id EN (bajo 'brands-en'). waterpik-en ya existe.
BRAND_EN = {'waterpik': 129}
BRANDS_PARENT_ES_ID = 32   # term_id del término 'brands'
BRANDS_PARENT_EN_ID = 127  # term_id del término 'brands-en'

# CATEGORIA Excel → term_taxonomy_id ES
CAT_ES = {
    'Productos':       29,
    'Producto':        29,
    'Assets':          40,
    'Documento':       42,
    'Piezas Digitales':43,
    'Piezas Impresas': 44,
    'Brand Camp':      59,
    'Ecommerce':       72,
}
# CATEGORIA Excel → term_taxonomy_id EN (productos-en=125, resto se crea)
CAT_EN = {'Productos': 125, 'Producto': 125}

# HOJA → brand slug
HOJA_BRAND = {
    'Hero': 'hero', 'Hero - English': 'hero',
    'TheraBreath': 'therabreath', 'TheraBreath - English': 'therabreath',
    ' TheraBreath - English': 'therabreath',
    'Batiste': 'batiste', 'Batiste - English': 'batiste',
    ' Batiste - English': 'batiste',
    'Waterpik': 'waterpik', 'Waterpik - English': 'waterpik',
    'OxiClean': 'oxiclean', 'OxiClean - English': 'oxiclean',
    'Trojan': 'trojan', 'Trojan - English': 'trojan',
    'A&H': 'arm-hammer', 'A&H - English': 'arm-hammer',
    'Stérimar': 'sterimar', 'Stérimar - English': 'sterimar',
    'Nair': 'nair', 'Nair - English': 'nair',
    'Brand Camp': 'brand-camp', 'Ecommerce': 'ecommerce',
}
# HOJA → brand display name (para meta 'brand')
HOJA_BRAND_NAME = {
    'Hero': 'Hero', 'Hero - English': 'Hero',
    'TheraBreath': 'TheraBreath', 'TheraBreath - English': 'TheraBreath',
    ' TheraBreath - English': 'TheraBreath',
    'Batiste': 'Batiste', 'Batiste - English': 'Batiste',
    ' Batiste - English': 'Batiste',
    'Waterpik': 'Waterpik', 'Waterpik - English': 'Waterpik',
    'OxiClean': 'OxiClean', 'OxiClean - English': 'OxiClean',
    'Trojan': 'Trojan', 'Trojan - English': 'Trojan',
    'A&H': 'Arm & Hammer', 'A&H - English': 'Arm & Hammer',
    'Stérimar': 'Stérimar', 'Stérimar - English': 'Stérimar',
    'Nair': 'Nair', 'Nair - English': 'Nair',
    'Brand Camp': 'Brand Camp', 'Ecommerce': 'Ecommerce',
}

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s',
    handlers=[logging.FileHandler(LOG_FILE, encoding='utf-8'),
              logging.StreamHandler(sys.stdout)]
)
log = logging.getLogger()

# ── Helpers generales ─────────────────────────────────────────────────────────
def slugify(text):
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')
    text = re.sub(r'[^a-z0-9\s-]', '', text.lower())
    return re.sub(r'[\s-]+', '-', text).strip('-')

def dropbox_direct(url):
    if not url or 'dropbox.com' not in url:
        return url
    return re.sub(r'dl=\d', 'dl=1', url) if 'dl=' in url else url + '&dl=1'

def build_excerpt(descripcion, formato):
    desc = (descripcion or '').strip()
    fmt  = (formato or '').strip()
    if not desc and not fmt:
        return ''
    if not fmt:
        return desc
    return f"{desc}\r\n<br><br>\r\n<b>Archivos:</b><br>\r\n{fmt}"

# ── SSH / WP-CLI ──────────────────────────────────────────────────────────────
def open_ssh():
    ssh = paramiko.SSHClient()
    ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    ssh.connect(SSH['host'], port=SSH['port'],
                username=SSH['username'], password=SSH['password'])
    return ssh

def wp_cli(ssh, cmd, dry_run=False):
    full = f"wp {cmd} --path={WP_PATH} --url={WP_URL} 2>/dev/null"
    if dry_run:
        log.info(f"[DRY] {full}")
        return '0'
    _, stdout, stderr = ssh.exec_command(full)
    out = stdout.read().decode().strip()
    err = stderr.read().decode().strip()
    if err:
        log.warning(f"WP-CLI stderr: {err}")
    return out

def import_media(ssh, url, post_id, dry_run=False):
    direct = dropbox_direct(url)
    if not direct:
        return None
    result = wp_cli(ssh, f'media import "{direct}" --post_id={post_id} --porcelain', dry_run)
    if dry_run:
        return 9999
    try:
        return int(result) if result and result.isdigit() else None
    except Exception:
        return None

def db_ping(conn):
    """Ping/reconnect MySQL after potentially long WP-CLI calls."""
    try:
        conn.ping(reconnect=True)
    except Exception:
        pass

# ── DB helpers ────────────────────────────────────────────────────────────────
def unique_slug(cur, base_slug, post_type):
    slug = base_slug
    n = 1
    while True:
        cur.execute("SELECT ID FROM wp_posts WHERE post_name=%s AND post_type=%s", (slug, post_type))
        if not cur.fetchone():
            return slug
        slug = f"{base_slug}-{n}"
        n += 1

def create_post(cur, title, excerpt, post_type, dry_run=False):
    slug = unique_slug(cur, slugify(title)[:190] or 'doc', post_type)
    if dry_run:
        log.info(f"[DRY] INSERT post: {post_type} | {title}")
        return 0
    cur.execute("""
        INSERT INTO wp_posts
            (post_author, post_date, post_date_gmt, post_content, post_excerpt,
             post_status, post_type, post_title, post_name,
             post_modified, post_modified_gmt, to_ping, pinged, post_content_filtered)
        VALUES (%s, NOW(), UTC_TIMESTAMP(), '', %s,
                'publish', %s, %s, %s,
                NOW(), UTC_TIMESTAMP(), '', '', '')
    """, (DEFAULT_AUTHOR, excerpt, post_type, title, slug))
    return cur.lastrowid

def set_meta(cur, post_id, key, value, dry_run=False):
    if dry_run:
        return
    cur.execute("DELETE FROM wp_postmeta WHERE post_id=%s AND meta_key=%s", (post_id, key))
    cur.execute("INSERT INTO wp_postmeta (post_id, meta_key, meta_value) VALUES (%s,%s,%s)",
                (post_id, key, value))

def assign_term(cur, post_id, term_taxonomy_id, dry_run=False):
    if dry_run:
        return
    cur.execute(
        "INSERT IGNORE INTO wp_term_relationships (object_id, term_taxonomy_id, term_order) VALUES (%s,%s,0)",
        (post_id, term_taxonomy_id))
    cur.execute(
        "UPDATE wp_term_taxonomy SET count=count+1 WHERE term_taxonomy_id=%s",
        (term_taxonomy_id,))

def get_or_create_term(cur, name, slug, taxonomy, parent_term_id=0, dry_run=False):
    cur.execute(
        "SELECT tt.term_taxonomy_id FROM wp_terms t JOIN wp_term_taxonomy tt ON tt.term_id=t.term_id WHERE t.slug=%s AND tt.taxonomy=%s",
        (slug, taxonomy))
    row = cur.fetchone()
    if row:
        return row[0]
    if dry_run:
        log.info(f"[DRY] CREATE TERM: {taxonomy} / {slug}")
        return 0
    cur.execute("INSERT INTO wp_terms (name, slug, term_group) VALUES (%s,%s,0)", (name, slug))
    term_id = cur.lastrowid
    cur.execute(
        "INSERT INTO wp_term_taxonomy (term_id, taxonomy, description, parent, count) VALUES (%s,%s,'', %s,0)",
        (term_id, taxonomy, parent_term_id))
    return cur.lastrowid

def link_polylang(cur, es_id, en_id, dry_run=False):
    """Crea un par de traduccion Polylang entre es_id y en_id."""
    import hashlib, time
    slug = 'pll_' + hashlib.md5(f"{es_id}{en_id}{time.time()}".encode()).hexdigest()[:13]
    tti = get_or_create_term(cur, slug, slug, 'post_translations', dry_run=dry_run)
    assign_term(cur, es_id, tti, dry_run)
    assign_term(cur, en_id, tti, dry_run)

# ── Lógica principal de un documento ─────────────────────────────────────────
def process_document(conn, ssh, row, dry_run=False):
    hoja      = row['hoja']
    idioma    = row['idioma']
    nombre    = row['nombre']
    categoria = row['categoria']
    subcat    = row['subcategoria']
    formato   = row['formato']
    amplify   = row['amplify_url']
    img_p     = row['img_principal']
    previews  = [p.strip() for p in row['previews'].split(' | ') if p.strip()]
    sa_list   = json.loads(row['single_assets']) if row['single_assets'] else []
    lang      = 'en' if idioma == 'English' else 'es'
    brand_slug = HOJA_BRAND.get(hoja, slugify(hoja.replace(' - English', '')))
    brand_name = HOJA_BRAND_NAME.get(hoja, hoja.replace(' - English', '').strip())

    log.info(f"  Procesando [{lang.upper()}]: {nombre}")

    # Guard: skip if this Amplify URL already exists for THIS language
    # Same URL can exist as both ES and EN (independent documents)
    if not dry_run:
        cur = conn.cursor()
        cur.execute("""
            SELECT pm.post_id FROM wp_postmeta pm
            JOIN wp_term_relationships tr ON tr.object_id = pm.post_id
            WHERE pm.meta_key = '_dlp_direct_link_url' AND pm.meta_value = %s
              AND tr.term_taxonomy_id = %s
            LIMIT 1
        """, (amplify, LANG_TTI[lang]))
        existing = cur.fetchone()
        if existing:
            log.warning(f"    SKIP (ya existe post {existing[0]} con esta URL en {lang.upper()})")
            return existing[0]

    excerpt = build_excerpt(row['descripcion'], formato)
    cur = conn.cursor()

    # PHASE 1: Post + terms + basic meta + single assets (all fast DB ops)
    # 1. Crear post dlp_document
    post_id = create_post(cur, nombre, excerpt, 'dlp_document', dry_run)
    if not dry_run:
        log.info(f"    Post creado: ID={post_id}")

    # 2. Idioma (Polylang language taxonomy)
    assign_term(cur, post_id, LANG_TTI[lang], dry_run)

    # 2b. File type (DLP requiere este término; "www" = documento de tipo URL)
    assign_term(cur, post_id, FILE_TYPE_URL_TTI, dry_run)

    # 3. Brand category
    if lang == 'es':
        brand_tti = BRAND_ES.get(brand_slug)
        if not brand_tti:
            brand_tti = get_or_create_term(
                cur, brand_name, brand_slug, 'doc_categories', BRANDS_PARENT_ES_ID, dry_run)
        assign_term(cur, post_id, brand_tti, dry_run)
    else:
        brand_tti = BRAND_EN.get(brand_slug)
        if not brand_tti:
            en_slug = f"{brand_slug}-en"
            brand_tti = get_or_create_term(
                cur, brand_name, en_slug, 'doc_categories', BRANDS_PARENT_EN_ID, dry_run)
            if not dry_run:
                BRAND_EN[brand_slug] = brand_tti
        assign_term(cur, post_id, brand_tti, dry_run)

    # 4. Content category
    if lang == 'es':
        cat_tti = CAT_ES.get(categoria)
        if not cat_tti and categoria:
            cat_tti = get_or_create_term(cur, categoria, slugify(categoria), 'doc_categories', dry_run=dry_run)
    else:
        cat_tti = CAT_EN.get(categoria)
        if not cat_tti and categoria:
            en_cat_slug = slugify(categoria) + '-en'
            cat_tti = get_or_create_term(cur, categoria, en_cat_slug, 'doc_categories', dry_run=dry_run)
            if not dry_run and cat_tti:
                CAT_EN[categoria] = cat_tti
    if cat_tti:
        assign_term(cur, post_id, cat_tti, dry_run)

    # Subcategoria (si existe como término)
    if subcat:
        sub_slug = slugify(subcat)
        cur.execute(
            "SELECT tt.term_taxonomy_id FROM wp_terms t JOIN wp_term_taxonomy tt ON tt.term_id=t.term_id WHERE t.slug=%s AND tt.taxonomy='doc_categories'",
            (sub_slug,))
        sub_row = cur.fetchone()
        if sub_row:
            assign_term(cur, post_id, sub_row[0], dry_run)

    # 5. Meta básico
    set_meta(cur, post_id, '_dlp_direct_link_url',      amplify,  dry_run)
    set_meta(cur, post_id, '_dlp_document_link_type',   'url',    dry_run)
    set_meta(cur, post_id, '_dlp_document_visibility',  '',       dry_run)
    set_meta(cur, post_id, '_dlp_document_file_size',   '',       dry_run)
    set_meta(cur, post_id, 'dlp_tipo_archivo',          formato,  dry_run)
    set_meta(cur, post_id, 'brand',                     brand_name, dry_run)
    set_meta(cur, post_id, 'order',                     '0',      dry_run)
    set_meta(cur, post_id, '_dlp_download_count',       '0',      dry_run)

    # 6. Single assets (no WP-CLI needed — all DB)
    ext = formato.split('-')[0].strip() if '-' in formato else formato.strip()
    for i, sa in enumerate(sa_list, start=1):
        sa_title = f"{nombre} — {sa['titulo']}" if sa.get('titulo') else f"{nombre} — Pieza {i:02d}"
        sa_id = create_post(cur, sa_title, '', 'single_asset', dry_run)
        set_meta(cur, sa_id, 'asset_extension', ext,             dry_run)
        set_meta(cur, sa_id, 'asset_file_size', ASSET_FILE_SIZE, dry_run)
        set_meta(cur, sa_id, 'url_file',        sa['url'],       dry_run)
        set_meta(cur, post_id, f'asset_{i}', str(sa_id),         dry_run)
        log.info(f"    single_asset_{i}: ID={sa_id} | {sa_title}")

    # Commit everything before media imports (WP-CLI can be slow and drop the tunnel)
    if not dry_run:
        conn.commit()
        log.info(f"    Post base committed: ID={post_id}")

    # PHASE 2: Media imports via WP-CLI (slow — Dropbox downloads)
    # Reconnect after each import in case tunnel dropped

    # 7. Imagen principal -> featured image
    if img_p:
        att_id = import_media(ssh, img_p, post_id, dry_run)
        if att_id:
            db_ping(conn)
            cur = conn.cursor()
            set_meta(cur, post_id, '_thumbnail_id', str(att_id), dry_run)
            if not dry_run:
                conn.commit()
            log.info(f"    Featured image: attachment={att_id}")

    # 8. Preview slider images
    for i, prev_url in enumerate(previews, start=1):
        att_id = import_media(ssh, prev_url, post_id, dry_run)
        if att_id:
            db_ping(conn)
            cur = conn.cursor()
            set_meta(cur, post_id, f'preview_slider_{i}', str(att_id), dry_run)
            if not dry_run:
                conn.commit()
            log.info(f"    preview_slider_{i}: attachment={att_id}")

    if not dry_run:
        log.info(f"    OK post_id={post_id}")

    return post_id

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    dry_run   = '--dry-run' in sys.argv
    solo_marca = None
    if '--solo' in sys.argv:
        idx = sys.argv.index('--solo')
        if idx + 1 < len(sys.argv):
            solo_marca = sys.argv[idx + 1].lower()

    if dry_run:
        log.info("=== MODO DRY-RUN — no se escribe nada ===")

    # Leer CSV
    with open(INPUT_CSV, encoding='utf-8-sig') as f:
        rows = [r for r in csv.DictReader(f) if r['estado'] == 'NUEVO']

    if solo_marca:
        rows = [r for r in rows if solo_marca in r['marca'].lower()]

    log.info(f"Documentos NUEVO a procesar: {len(rows)}")

    # Separar ES y EN
    es_rows = [r for r in rows if r['idioma'] != 'English']
    en_rows = [r for r in rows if r['idioma'] == 'English']
    log.info(f"  ES: {len(es_rows)}  EN: {len(en_rows)}")

    # Conexiones en dict mutable para poder reconectar desde run_batch
    ctx = {'tunnel': None, 'conn': None, 'ssh': None}

    def open_connections():
        for key in ('tunnel', 'conn', 'ssh'):
            try:
                if ctx[key]:
                    ctx[key].stop() if key == 'tunnel' else ctx[key].close()
            except Exception:
                pass
        t = SSHTunnelForwarder(
            (SSH['host'], SSH['port']),
            ssh_username=SSH['username'], ssh_password=SSH['password'],
            remote_bind_address=('127.0.0.1', 3306))
        t.start()
        if hasattr(t, '_transport') and t._transport:
            t._transport.set_keepalive(10)
        ctx['tunnel'] = t
        ctx['conn'] = pymysql.connect(host='127.0.0.1', port=t.local_bind_port, **DB)
        s = open_ssh()
        s.get_transport().set_keepalive(30)
        ctx['ssh'] = s
        log.info("Connections established.")

    open_connections()

    # Registros para vincular traducciones al final
    # {consolidado_url: {'es': post_id, 'en': post_id}}
    translation_map = {}

    RECONNECT_ERRORS = (pymysql.err.InterfaceError, pymysql.err.OperationalError)

    def run_batch(batch, lang_key):
        for row in batch:
            url = row['amplify_url']
            for attempt in range(2):
                try:
                    pid = process_document(ctx['conn'], ctx['ssh'], row, dry_run)
                    if url not in translation_map:
                        translation_map[url] = {}
                    translation_map[url][lang_key] = pid
                    break
                except RECONNECT_ERRORS as e:
                    if attempt == 0:
                        log.warning(f"Connection lost on '{row['nombre']}' — reconnecting...")
                        open_connections()
                    else:
                        log.error(f"ERROR (retry failed) en '{row['nombre']}': {e}")
                except Exception as e:
                    try:
                        db_ping(ctx['conn'])
                        ctx['conn'].rollback()
                    except Exception:
                        pass
                    log.error(f"ERROR en '{row['nombre']}': {e}", exc_info=True)
                    break

    log.info("--- Cargando ES ---")
    run_batch(es_rows, 'es')

    log.info("--- Cargando EN ---")
    run_batch(en_rows, 'en')

    # Vincular traducciones Polylang
    log.info("--- Vinculando traducciones Polylang ---")
    pll_cur = ctx['conn'].cursor()
    pares = 0
    for url, ids in translation_map.items():
        if 'es' in ids and 'en' in ids:
            link_polylang(pll_cur, ids['es'], ids['en'], dry_run)
            pares += 1
            if not dry_run:
                ctx['conn'].commit()
    log.info(f"  Pares vinculados: {pares}")

    ctx['ssh'].close()
    ctx['conn'].close()
    ctx['tunnel'].stop()
    log.info("=== Cargue finalizado ===")


if __name__ == '__main__':
    main()
