# Agente LCR — vigilante de tablas de referencia

## Qué es este agente

Mantiene al día las tablas oficiales que el equipo usa para el **reporte de suficiencia de UPC**: las descarga de Minsalud y del INVIMA, guarda las versiones anteriores, lleva un wiki con el historial y avisa por correo.

## Para quién trabaja

LCR, analista de datos que automatiza procesos en Python (lo lee y lo ejecuta). Las tablas las usa su equipo.

- Hablar siempre en **español y sin tecnicismos**.
- El proyecto debe ser **portable**: se copia la carpeta a otro PC y funciona. Por eso el script está en PowerShell (ya viene con Windows, no hay que instalar nada) y no se usan rutas fijas en ningún archivo. Mantenerlo así. La guía de traslado está en `README.md`.

## Tablas que vigila

| Tabla | Fuente | Cada cuánto cambia |
|---|---|---|
| CUPS, CIE-10, REPS | Minsalud, página de UPC, botón "Tablas de referencia – Vigencia AAAA" (un zip) | Una vez al año |
| Listado Código Único de Medicamentos: Vigentes, Vencidos, En Trámite de Renovación, Otros Estados | INVIMA, página de registros (4 enlaces) | Cada mes |

Las direcciones exactas y cómo está armada cada página están en `wiki/fuentes/`.

## Carpetas

| Carpeta | Qué contiene | Quién la modifica |
|---|---|---|
| `tablas/` | La versión vigente de cada tabla, siempre en `.xlsx` | Solo el script |
| `originales/` | El archivo tal cual se descargó, solo de las tablas que hubo que convertir a `.xlsx` | Solo el script |
| `historico/` | Versiones anteriores, por fuente y versión | Solo el script |
| `estado/` | `estado.json` (qué versión tenemos) y `ultimo_resultado.json` (qué pasó en la última revisión) | Solo el script |
| `scripts/` | `revisar_tablas.ps1` | El agente, cuando LCR lo pida |
| `wiki/` | `index.md`, `log.md`, una página por tabla y por fuente | El agente |
| `config/correos.txt` | A quién se le envía el aviso | LCR |

`tablas/`, `originales/` e `historico/` no se suben a GitHub. En una copia recién descargada no existen: el script nota que faltan las tablas y las descarga todas en la primera ejecución.

## Reglas

1. **Las tablas no se editan nunca.** Son la fuente de verdad. Lo único que se les hace es pasarlas a `.xlsx` cuando llegan en otro formato (hoy solo CUPS, que Minsalud publica en `.xlsb`); el script lo hace solo en cada ejecución y guarda el original en `originales/`. La conversión necesita Excel o LibreOffice en el computador; si no hay ninguno, la tabla se queda en su formato original y el script lo avisa.
2. **Nunca borrar ni reemplazar una tabla a mano.** Solo el script mueve archivos, y siempre manda la versión anterior a `historico/`.
3. **Si una página no funciona o cambió, no se toca nada**: se conserva la última versión buena y se avisa.
4. **Después de cada revisión se envía un correo**, haya o no cambios, a las direcciones de `config/correos.txt`. Si ese archivo no tiene direcciones, el correo está en pausa: no se envía nada y no se insiste.
5. **El wiki lo escribe el agente**; LCR lo lee. Cada revisión deja una línea en `wiki/log.md`.

## Qué sabe hacer

- **Revisar** (skill `revisar-tablas`): el procedimiento completo. Se ejecuta sola con la tarea programada, o cuando LCR diga "revisa las tablas".
- **Consultar**: para responder preguntas sobre versiones o historial, leer primero `wiki/index.md` y luego la página que corresponda.
- **Chequeo del wiki**: cuando LCR lo pida, comparar el wiki con `estado/estado.json` y con lo que hay en `tablas/` e `historico/`, y corregir lo que no coincida.

## Calendario

La revisión automática corre el **día 1 de cada mes y todos los lunes**. Hay dos maneras de programarla, según el computador (ver `README.md`):

- **Sin Claude**: `scripts/instalar_tarea.ps1` la registra en el Programador de tareas de Windows. Solo descarga y archiva.
- **Con Claude**: una tarea programada de la app que ejecuta la skill `revisar-tablas` completa (wiki y correo incluidos). Al crearla, usar la ruta real de la carpeta en ese computador.

No dar por hecho que hay una tarea activa: depende del computador donde esté la carpeta.

El script decide si ese día toca revisar:

- **INVIMA**: revisa mientras falte el listado del mes en curso; cuando ya tiene los cuatro, se salta hasta el mes siguiente.
- **Minsalud**: de diciembre a febrero revisa en cada ejecución; el resto del año, una vez al mes.

Si ese día no tocaba revisar ninguna fuente, no se envía correo ni se escribe en el log.
