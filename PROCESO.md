# Proceso de cargue CYD Assets — Detalle técnico

## Macro VBA — ConsolidarHojas

El Excel tiene una hoja por marca (Hero, TheraBreath, Batiste, etc.). La macro genera la hoja **Consolidado** que usan todos los scripts Python.

**Archivo**: `macro/ConsolidarHojas.bas` en el repo `wp-dlp-asset-loader`

**Qué hace**:
1. Limpia la hoja Consolidado (mantiene cabecera)
2. Recorre todas las hojas salvo `Consolidado` y `CRONOGRAMA`
3. Copia columnas B–K de cada hoja fuente
4. Agrega el nombre de la hoja en columna K (`Hoja`) — así `diagnostico_assets.py` sabe el idioma y la marca

**Estructura columnas por hoja fuente**:
- A: UBICACION EN DB (ignorada) | B: NOMBRE | C: CATEGORIA | D: SUBCATEGORIA | E: DESCRIPCION | F: FORMATO | G: IMAGEN PRINCIPAL (Dropbox) | H: PREVIEW CONTENIDOS (Dropbox) | I: EDITABLES | J: PIEZAS | K: CONSOLIDADO (Amplify URL)

**Columna K "Hoja" en Consolidado** es clave: `diagnostico_assets.py` determina idioma por si el nombre de hoja contiene "English" o no. Si el nombre de hoja fuente cambia, revisar `SKIP_HOJAS` y el mapeo en `leer_consolidado()`.

**Ejecutar**: `Alt+F8 > ConsolidarHojas > Ejecutar`. El archivo debe estar guardado como `.xlsm`.

---

## Archivos del proyecto

| Archivo | Rol |
|---|---|
| `Listado de Assets - PD.xlsx` | Fuente de datos; debe tener hoja **Consolidado** |
| `diagnostico_assets.py` | Lee Excel, cruza contra DB, genera CSV |
| `diagnostico_assets.csv` | Resultado del diagnóstico (NUEVO / EXISTE / SIN_URL) |
| `cargue_assets.py` | Cargue masivo via SSH tunnel + WP-CLI |
| `cargue_assets.log` | Log completo de cada corrida |
| `validacion_cargue.py` | Cruce de calidad Excel↔DB; verifica thumbnail, previews, assets, Polylang |
| `validacion_cargue.csv` | Resultado de la validación (OK / WARN / ERROR / FALTANTE) |
| `.env` | Credenciales SSH y DB (nunca subir al repo) |

## Conexión SSH / DB

```python
SSH = dict(host='35.236.219.140', port=25628, username='cdportal', password='cJobniBeuPb8jE3')
DB  = dict(user='cdportal', password='gfoj4eGoeBDVIwF', db='cdportal', charset='utf8mb4')
WP_PATH = '/www/cdportal_742/public'
```

Requiere `paramiko<4.0` (3.5.1) y `sshtunnel`. Con paramiko 5+ falla con `AttributeError: module 'paramiko' has no attribute 'DSSKey'`.

## Estructura del cargue por documento

**Phase 1 — DB rápida (se commitea antes de media):**
- INSERT `wp_posts` (post_type=dlp_document, status=publish)
- Términos: idioma (Polylang), **file_type=www** (tti=14, requerido por DLP para docs URL), marca (doc_categories), categoría, subcategoría
- Meta: `_dlp_direct_link_url`, `_dlp_document_link_type`, `_dlp_document_visibility` (vacío, requerido por DLP `Visibility.php`), `_dlp_document_file_size` (vacío, requerido por DLP para no causar critical error), `dlp_tipo_archivo`, `brand`, `order`, `_dlp_download_count`
- Sub-posts `single_asset` con `asset_extension`, `asset_file_size`, `url_file`
- Meta `asset_N` en el documento padre

> **IMPORTANTE**: `_dlp_document_visibility` y `_dlp_document_file_size` deben crearse con valor vacío en Phase 1. Si faltan, DLP genera un "critical error" al renderizar el single del documento porque `Visibility.php` y el renderer de file size los esperan siempre presentes.

**Phase 2 — Media via WP-CLI (puede caer el tunnel):**
- `wp media import <dropbox_url> --post_id=N --porcelain` → attachment ID
- `db_ping(conn)` antes de cada write post-import para reconectar si es necesario
- Commit después de cada media

## Idempotencia

Al inicio de cada documento el script consulta:
```sql
SELECT pm.post_id FROM wp_postmeta pm
JOIN wp_term_relationships tr ON tr.object_id = pm.post_id
WHERE pm.meta_key = '_dlp_direct_link_url' AND pm.meta_value = <amplify_url>
  AND tr.term_taxonomy_id = <LANG_TTI>
```
La clave incluye el idioma (lang_tti). Así el mismo amplify_url puede existir como doc ES y doc EN sin colisión. Si ya existe, emite WARNING SKIP y continúa.

## Polylang y cargues parciales (IMPORTANTE)

El script vincula pares ES↔EN al final del cargue (`link_polylang`) **solo si ambos idiomas se procesaron en el mismo lote**. Si se carga solo EN (porque los ES ya estaban), los pares no se crean.

**Después de cualquier cargue parcial (solo ES o solo EN):**
```
python validacion_cargue.py --reparar
```
Esto detecta y crea todos los pares `post_translations` faltantes.

Señal de alerta: en el log del cargue, si `Pares vinculados: 0` pero había docs EN → correr --reparar.

## Reconexión automática

Cuando el tunnel SSH muere (InterfaceError/OperationalError), `run_batch` detecta el error, llama `open_connections()` para recrear tunnel+conn+ssh y reintenta el documento fallido una vez.

## Agrupación de documentos (diagnostico)

La clave de agrupación es la **URL de Amplify del campo CONSOLIDADO**. Dentro de cada grupo:
- Filas con distinta URL en EDITABLES → `single_assets`
- Filas con URL en PREVIEW CONTENIDOS (DROPBOX) → `previews` (slider)
- Primera fila del grupo → datos principales del documento

## Taxonomías WordPress

| Meta / término | Valor |
|---|---|
| Idioma ES | term_taxonomy_id=3 |
| Idioma EN | term_taxonomy_id=6 |
| Brands ES (padre) | term_id=32 |
| Brands EN (padre) | term_id=127 |
| Hero ES | term_taxonomy_id=30 |
| Arm & Hammer ES | term_taxonomy_id=31 |
| Nair ES | term_taxonomy_id=75 |
| Waterpik EN | term_taxonomy_id=129 |

Los términos de marca EN que no existen se crean on-demand con slug `{marca}-en`.

## Polylang

Los pares de traducción ES/EN se crean cuando dos documentos comparten la misma URL Amplify. Se crea un término en `post_translations` taxonomy y ambos posts se asignan a ese término via `wp_term_relationships`.

## Validación de calidad post-cargue

Correr después de cada cargue masivo:
```
python validacion_cargue.py
```

### Estados del CSV de validación

| Estado | Descripción |
|---|---|
| OK | Todos los campos correctos |
| WARN | Algún mismatch no crítico (previews/assets) |
| ERROR | Falta thumbnail o meta requerido por DLP |
| FALTANTE | El post no existe en DB aunque está en Excel |

### Issues comunes en validación

| Issue | Causa | Acción |
|---|---|---|
| `SIN_POLYLANG_PAIR` | Cargue parcial (solo ES o solo EN) | `python validacion_cargue.py --reparar` |
| `PREVIEWS_MISMATCH DB>Excel` | Cargue original acumuló ES+EN antes del fix de idioma; no es error real | Ignorar |
| `PREVIEWS_MISMATCH Excel>DB` | Dropbox URL expirada o import fallido | Re-importar manualmente vía WP-CLI |
| `SIN_THUMBNAIL` | Dropbox URL principal expirada o import fallido | Re-importar manualmente vía WP-CLI |
| `ASSETS_MISMATCH DB>Excel` | Cargue original acumuló assets de ambos idiomas | Ignorar (datos extras, no faltan) |

### Re-importar media fallida manualmente

```bash
# SSH al servidor
wp media import "<dropbox_url_dl=1>" --post_id=<post_id> --porcelain
# Luego setear el meta correspondiente
wp post meta set <post_id> _thumbnail_id <attachment_id>
wp post meta set <post_id> preview_slider_N <attachment_id>
```

## Errores conocidos

| Error | Causa | Fix |
|---|---|---|
| `AttributeError: module 'paramiko' has no attribute 'DSSKey'` | paramiko 5.x incompatible con sshtunnel | `pip install "paramiko<4.0"` |
| `InterfaceError: (0, '')` / SSH session not active | Tunnel SSH cayó durante descargas largas | Volver a correr (reconexión automática + idempotencia) |
| KeyError: 'descripcion' | diagnostico_assets.csv desactualizado | Regenerar con `python diagnostico_assets.py` |
| `UnicodeEncodeError` con caracteres `→` | Windows terminal cp1252 | Usar `->` en lugar de `→` en prints |
| Critical error en single del documento | Faltan términos/meta requeridos por DLP: `file_type=www` (tti=14), `_dlp_document_visibility=''`, `_dlp_document_file_size=''` | Parchar con INSERT IGNORE en `wp_term_relationships` y `wp_postmeta` |
| `PHP Fatal: ltrim(): Argument must be string, WP_Error given` en `functions.php:707` | `get_term_link()` devuelve `WP_Error` porque **Kinsta Redis object cache** tiene cacheado un resultado negativo para un término creado vía SQL directo (no `wp_insert_term`). El cache guarda `false` cuando alguien visita la página antes de que el término existiera. WP-CLI no reproduce el error porque no pasa por Redis. | (1) `wp cache flush` en el servidor para limpiar el cache negativo. (2) Guard defensivo `is_wp_error($term_link)` ya agregado en `functions.php` línea 704-711 para que nunca explote aunque el cache vuelva a ser stale. |

## Notas sobre Redis y términos creados vía SQL directo

Kinsta usa Redis como WordPress object cache. Cuando el script crea términos de categoría vía `INSERT INTO wp_terms`/`wp_term_taxonomy` (fuera del flujo normal de `wp_insert_term`), el cache de Redis **no se invalida**. Si alguien visita un documento antes de que el término exista, Redis cachea `get_term_by('slug', ...) = false`. Después de crear el término, Redis sigue devolviendo `false` hasta que se expira o se flushea.

**Fix permanente ya aplicado**: `is_wp_error()` guard en `functions.php:704-711` (`asylum_get_main_content_html`). Si `get_term_link()` falla, renderiza el nombre como `<span>` en lugar de explotar.

**Si el error vuelve a aparecer después de un cargue**: correr `wp cache flush` vía SSH:
```bash
cd /www/cdportal_742/public && wp cache flush
```
