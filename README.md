# wp-dlp-asset-loader

Herramienta para carga masiva de documentos (`dlp_document`) en WordPress con Document Library Pro, via SSH tunnel + WP-CLI + pymysql.

## Scripts

| Script | Uso |
|---|---|
| `diagnostico_assets.py` | Lee el Excel y cruza contra la DB. Genera `diagnostico_assets.csv` con NUEVO / EXISTE / SIN_URL |
| `cargue_assets.py` | Carga los documentos NUEVO en WordPress (posts, metas, términos, media) |
| `validacion_cargue.py` | Cruce de calidad post-cargue: thumbnail, previews, assets, Polylang |

## Requisitos

```
pip install -r requirements.txt
```

> **Importante**: `paramiko<4.0` requerido. La versión 5+ rompe `sshtunnel`.

## Configuración

Copiar `.env.example` a `.env` y completar las credenciales:

```
cp .env.example .env
```

## Flujo

```bash
# 1. Ver qué falta
python diagnostico_assets.py

# 2. Validar sin escribir
python cargue_assets.py --dry-run

# 3. Cargar
python cargue_assets.py

# 4. Flush de cache Redis (Kinsta)
# SSH: cd /www/site/public && wp cache flush

# 5. Cruce de calidad
python validacion_cargue.py

# 6. Reparar pares Polylang si hubo cargue parcial
python validacion_cargue.py --reparar
```

## Notas importantes

- El Excel debe tener hoja **Consolidado** (generada con macro VBA).
- El cargue es **idempotente**: re-ejecutable sin duplicados (guard por URL + idioma).
- Los pares Polylang se vinculan automáticamente solo si ES y EN se cargan en el mismo lote. Si se carga solo EN, ejecutar `--reparar`.
- Después de cada cargue masivo, hacer `wp cache flush` en el servidor para evitar errores de Redis cache.
- Ver `PROCESO.md` para documentación técnica completa.
