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

El archivo HTML incluye **dos vistas** seleccionables por pestañas, ambas con
**idénticas funcionalidades**:

1. **Registro MP 2026** — Hoja `Registro_MP-2026`, columnas **B a AQ**, filas 8 a 973.
   Cada mes tiene dos subcolumnas: **P** (programado) y **R** (realizado).
2. **PMP 2026 · Carta Gantt** — Hoja `PMP_2026`, con la familia de equipo (`Fam`),
   los 12 meses y el resumen `PMP` / `MP-R` / `MP-RA`.

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

## Regenerar el HTML

```bash
pip install openpyxl
python3 build_html.py            # usa Programacion_MP_2026.xlsm del repo
# o bien:
python3 build_html.py /ruta/a/otro_libro.xlsm
```
