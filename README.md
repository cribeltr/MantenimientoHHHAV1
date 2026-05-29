# Programación de Mantención Preventiva 2026 · HHHA

Visor HTML interactivo de la **programación y registro de mantenciones preventivas**
de los equipos médicos críticos del Hospital Hernán Henríquez Aravena, generado a
partir del libro Excel `Programacion_MP_2026.xlsm`.

## Archivos

| Archivo | Descripción |
|---|---|
| `programacion_mp_2026.html` | **Visor final** (un solo archivo, sin dependencias). Ábrelo en cualquier navegador. |
| `build_html.py` | Script que extrae los datos del Excel y regenera el HTML. |
| `Programacion_MP_2026.xlsm` | Libro Excel de origen. |

## Vistas

El archivo HTML incluye **tres vistas** seleccionables por pestañas:

1. **Registro MP 2026** — Hoja `Registro_MP-2026`, columnas **B a AQ**, filas 8 a 973.
   Cada mes tiene dos subcolumnas: **P** (programado) y **R** (realizado).
2. **PMP 2026 · Carta Gantt** — Hoja `PMP_2026`, con la familia de equipo (`Fam`),
   los 12 meses y el resumen `PMP` / `MP-R` / `MP-RA`.
3. **📋 Pendientes** — Vista agregada con *todos* los pendientes de todos los equipos
   (a partir del historial), con los mismos filtros, ordenamiento y exportación.

Las vistas 1 y 2 comparten **idénticas funcionalidades** de tabla.

## Funcionalidades (en ambas vistas)

- 🔍 **Buscador global** sobre todas las columnas.
- ▼ **Filtros por columna tipo Excel**: lista desplegable con casillas, conteo de
  registros por valor, buscador de valores y "Seleccionar todo". Los filtros de
  varias columnas se combinan (Y lógico).
- ↕ **Ordenamiento** ascendente / descendente (texto y numérico) desde cada columna.
- ✕ **Limpiar filtros** y **contador** de equipos mostrados.
- 🎨 **Coloreado de estados** con leyenda (X, R, PM, RA, Si, Si-RA, C1–C8, FS, Baja, No, NU).
- 📌 **Encabezado y primeras columnas fijas** (ID / N° Carpeta / Fam) al desplazar.
- ⬇ **Exportar CSV** del resultado filtrado.

## Historial por equipo (eventos · pendientes · tareas)

En cada fila, el botón **📋** (en la columna ID) abre un **panel lateral** con el
historial del equipo. La jerarquía es:

```
Equipo
└─ Evento            (fecha · tipo · qué ocurrió)
   └─ Pendiente      (título · estado · responsable · fecha límite)
      └─ Tarea/Gestión (tipo · descripción · estado · nota)
```

- Crear / editar / eliminar en los tres niveles; marcar tareas como hechas.
- El botón 📋 muestra un **contador** de pendientes abiertos y se pone **rojo si hay vencidos**.
- Filtro de barra **"Con pendientes abiertos / vencidos"** y pestaña **Pendientes** agregada.
- **Persistencia local**: se guarda en el navegador (`localStorage`). Con **⬇ Historial**
  descargas un respaldo `.json` y con **⬆ Importar** lo restauras o lo llevas a otro equipo.

> Como el historial se guarda en el navegador, al **regenerar el HTML desde el Excel**
> los datos se conservan (es una capa independiente de la tabla). Para compartirlo entre
> personas/computadores, usa Exportar/Importar el respaldo JSON.

## Regenerar el HTML

```bash
pip install openpyxl
python3 build_html.py            # usa Programacion_MP_2026.xlsm del repo
# o bien:
python3 build_html.py /ruta/a/otro_libro.xlsm
```
