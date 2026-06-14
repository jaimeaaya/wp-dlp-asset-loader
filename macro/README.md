# Macro VBA — ConsolidarHojas

Consolida todas las hojas de marca del Excel en la hoja **Consolidado**, que es la que leen los scripts Python.

## Cómo instalar la macro

### Opción A — Importar el módulo .bas (recomendado)

1. Abrir el Excel con las hojas de marca
2. Guardar como `.xlsm` (Excel con macros habilitadas)
3. Ir a `Desarrollador > Visual Basic` (o `Alt+F11`)
4. Clic derecho en el proyecto > `Importar archivo`
5. Seleccionar `ConsolidarHojas.bas`
6. Cerrar el editor VBA

### Opción B — Pegar el código manualmente

1. `Alt+F11` para abrir el editor VBA
2. `Insertar > Módulo`
3. Pegar el contenido de `ConsolidarHojas.bas`
4. Cerrar y guardar como `.xlsm`

## Cómo ejecutar

- `Alt+F8` > seleccionar `ConsolidarHojas` > `Ejecutar`
- O desde la cinta: `Desarrollador > Macros`

## Qué hace

1. Limpia la hoja **Consolidado** (mantiene la cabecera)
2. Recorre todas las hojas excepto `Consolidado` y `CRONOGRAMA`
3. Por cada hoja, copia las columnas B–K al Consolidado
4. Agrega el nombre de la hoja fuente en la columna K (`Hoja`)

## Estructura esperada de cada hoja fuente

| Col | Campo |
|-----|-------|
| A | UBICACION EN DB (provisional) — ignorada por la macro |
| B | NOMBRE |
| C | CATEGORIA |
| D | SUBCATEGORIA (Solo si aplica) |
| E | DESCRIPCION |
| F | FORMATO |
| G | IMAGEN PRINCIPAL (DROPBOX) |
| H | PREVIEW CONTENIDOS (DROPBOX) |
| I | EDITABLES |
| J | PIEZAS |
| K | CONSOLIDADO |

## Notas

- El archivo DEBE guardarse como `.xlsm` para mantener la macro.
- Si el archivo se distribuye como `.xlsx`, la macro se pierde — volver a importar `ConsolidarHojas.bas`.
- Ejecutar la macro ANTES de correr `diagnostico_assets.py`.
