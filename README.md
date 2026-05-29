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

## Motor de eventos por equipo (estado · ciclos · pendientes)

En cada fila, el botón **📋** (columna ID) abre un **panel lateral** con la bitácora del
equipo. Cada equipo tiene un **estado** (`operativo · no operativo · en servicio técnico ·
baja`) y los eventos lo van cambiando. El estado y los **ciclos correctivos** se **derivan**
del registro de eventos: al **anular** un evento, todo se recalcula.

### Los 7 tipos de evento

| Tipo | Efecto principal |
|---|---|
| **Solicitud de trabajo** | Abre un ciclo correctivo (Folio SIGEM, manual o automático) → **no operativo**. |
| **Visita técnica** | Diagnóstica/Correctiva; estado resultante elegible. Correctiva + operativo → **cierra** el ciclo. |
| **Orden de Compra** | Gestión dentro del ciclo (Trato directo / Compra ágil). No cambia el estado. |
| **Envío a servicio técnico** | N° de envío correlativo. Estado **fijo: en servicio técnico**. |
| **Recepción** | Retorno; si queda operativo **cierra** el ciclo (avisa si no hay ciclo abierto). |
| **Reparación** | Operativo cierra el ciclo; "en servicio técnico" lo deja abierto. |
| **Mantención preventiva** | Estado **automático** según el código de Resultado (Si, C1–C8, FS, Baja, NU, No). |

- **Folio SIGEM**: hilo del ciclo. Lo abre la Solicitud; los demás eventos lo **heredan
  preseleccionado**. Si se deja vacío en la Solicitud, se genera `SIGEM-2026-####`.
- **MP con C1–C8** → crea un **pendiente de reprogramación** y marca una **"R"** (reversible,
  borde naranja) en el **mes siguiente** de las tablas Registro/PMP. **NU** → pendiente
  "Localizar equipo". **Baja** → el equipo pasa a baja.
- **Anulación en vez de borrado**: los eventos no se eliminan; se **anulan** (quedan con fecha
  y motivo, tachados) y el estado se recalcula.
- **Pendientes** (automáticos o manuales) con **tareas/gestiones**; el botón 📋 muestra el
  contador de pendientes abiertos (rojo si vencidos) y un punto con el estado del equipo.
- Barra: filtro **"Pendientes abiertos/vencidos"** y **"Estado del equipo"**; pestaña
  **Pendientes** agregada de todos los equipos.

### Persistencia

Se guarda en el navegador (`localStorage`). Con **⬇ Historial** descargas un respaldo `.json`
y con **⬆ Importar** lo restauras o lo llevas a otro computador. Como es una capa
independiente, **el historial se conserva al regenerar el HTML desde el Excel**.

> **Pendiente de definir (no implementado):** el origen automático **"Auto-maestro"** (crear
> eventos al subir el archivo maestro) y la **"Conciliación"** de conflictos.

## Regenerar el HTML

```bash
pip install openpyxl
python3 build_html.py            # usa Programacion_MP_2026.xlsm del repo
# o bien:
python3 build_html.py /ruta/a/otro_libro.xlsm
```
