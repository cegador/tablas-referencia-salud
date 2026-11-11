# Agente LCR — tablas de referencia de salud siempre al día

Descarga y mantiene actualizadas las tablas oficiales que se usan en Colombia para el **reporte de suficiencia de UPC**, consultando directamente las páginas del Ministerio de Salud y del INVIMA.

Se puede usar de dos maneras:

- **En local**: un script de PowerShell que revisa, descarga y archiva. No necesita instalar nada.
- **Con un agente** (Claude Code): además lleva un wiki con el historial de cada tabla, avisa por correo e investiga cuando una página cambia.

> Proyecto no oficial. Las tablas pertenecen a Minsalud y al INVIMA; este repositorio no las incluye, solo las descarga de sus páginas públicas.

## Qué tablas mantiene

| Tabla | Fuente | Cada cuánto cambia |
|---|---|---|
| CUPS (con cobertura) | [Minsalud – UPC](https://www.minsalud.gov.co/salud/POS/Paginas/unidad-de-pago-por-capitacion-upc.aspx), «Tablas de referencia» | Una vez al año |
| CIE-10 | Minsalud – UPC | Una vez al año |
| REPS | Minsalud – UPC | Una vez al año |
| Código Único de Medicamentos: Vigentes | [INVIMA – registros](https://www.invima.gov.co/productos-vigilados/medicamentos-y-productos-biologicos/autorizacion-de-comercializacion-registros) | Cada mes |
| Código Único de Medicamentos: Vencidos | INVIMA – registros | Cada mes |
| Código Único de Medicamentos: En Trámite de Renovación | INVIMA – registros | Cada mes |
| Código Único de Medicamentos: Otros Estados | INVIMA – registros | Cada mes |

## Qué hace

1. Abre las dos páginas y busca si hay una versión más reciente que la que tienes.
2. Si la hay, la descarga, mueve la anterior a `historico/` y deja la nueva en `tablas/`.
3. Comprueba que todas las tablas estén en `.xlsx`. Si alguna llega en otro formato (CUPS viene en `.xlsb`), la convierte y guarda el original en `originales/`.
4. Si una página no funciona o cambió, no toca nada: conserva la última versión buena y lo avisa.

## Requisitos

- **Windows 10 u 11.** PowerShell ya viene incluido.
- **Internet.**
- **Excel o LibreOffice** (opcional): solo para convertir a `.xlsx`. Sin ellos todo funciona igual, pero CUPS se queda en `.xlsb`.
- **Claude Code** (opcional): solo para el uso con agente.

## Descarga

Con Git:

```
git clone https://github.com/<usuario>/<repositorio>.git
cd <repositorio>
```

Sin Git: en la página del repositorio, botón **Code → Download ZIP**, y descomprime la carpeta donde quieras. Puede estar en cualquier ubicación; no usa rutas fijas.

## Uso en local

Abre PowerShell dentro de la carpeta del proyecto.

### 1. Primera ejecución

```
powershell -ExecutionPolicy Bypass -File scripts\revisar_tablas.ps1
```

La primera vez descarga las siete tablas (unos 60 MB; tarda uno o dos minutos) y las deja en `tablas/`. Verás una línea por fuente y, debajo, cada tabla descargada.

### 2. Revisar cuando quieras

El mismo comando. Para obligarlo a revisar aunque ese día no toque según el calendario, añade `-Forzar`:

```
powershell -ExecutionPolicy Bypass -File scripts\revisar_tablas.ps1 -Forzar
```

### 3. Dejarlo automático

```
powershell -ExecutionPolicy Bypass -File scripts\instalar_tarea.ps1
```

Crea dos tareas en el Programador de tareas de Windows: todos los lunes y el día 1 de cada mes, a las 8:00. No pide permisos de administrador. El computador tiene que estar encendido a esa hora.

- Otra hora: añade `-Hora 07:30`.
- Quitarlas: añade `-Quitar`.

### Dónde ver qué pasó

| Archivo | Qué contiene |
|---|---|
| `estado/historial.log` | Un resumen de cada revisión, en texto |
| `estado/ultimo_resultado.json` | El detalle de la última revisión |
| `estado/estado.json` | Qué versión de cada tabla tienes ahora |

### Calendario de revisión

Aunque la tarea corra cada lunes, el script decide si ese día hace falta revisar:

- **INVIMA**: revisa mientras falte el listado del mes en curso. Cuando ya tiene los cuatro, se salta hasta el mes siguiente.
- **Minsalud**: de diciembre a febrero revisa en cada ejecución (es cuando sale la vigencia nueva); el resto del año, una vez al mes.

## Uso con agente

El repositorio trae lo que un agente necesita para trabajar con disciplina: `CLAUDE.md` (quién es y sus reglas), la skill `revisar-tablas` (el procedimiento paso a paso) y un wiki que el propio agente mantiene.

### 1. Abrir el proyecto

Abre la carpeta en [Claude Code](https://claude.com/claude-code) (app de escritorio o terminal). El agente lee `CLAUDE.md` al empezar.

### 2. Pedirle la revisión

Escribe:

```
revisa las tablas
```

El agente ejecuta el script, lee el resultado, actualiza el wiki y te cuenta qué cambió. Si una página falló, la abre para ver qué pasó y lo deja anotado.

### 3. Activar el correo (opcional)

Escribe las direcciones en `config/correos.txt`, una por línea. Después de cada revisión el agente envía un solo correo a esas direcciones, haya o no cambios. Necesita el conector de Gmail activo en Claude. Mientras el archivo esté vacío, no se envía nada.

### 4. Dejarlo automático (opcional)

Pídele:

```
programa la revisión de tablas para los lunes y el día 1 de cada mes
```

Crea una tarea programada en la app de Claude que ejecuta la skill completa, con wiki y correo. Corre mientras la app esté abierta.

### Otras cosas que puedes pedirle

- «¿Qué versión de CUPS tenemos y desde cuándo?»
- «¿Qué cambió en las tablas este año?»
- «Haz un chequeo del wiki» (compara el wiki con lo que hay realmente en las carpetas y corrige lo que no coincida).

### El wiki

Está en `wiki/` y se puede abrir con [Obsidian](https://obsidian.md) o con cualquier editor de texto.

- `wiki/index.md`: catálogo de las tablas con su versión vigente.
- `wiki/log.md`: registro de cada revisión, en orden.
- `wiki/tablas/`: una página por tabla, con su historial de versiones.
- `wiki/fuentes/`: cómo está armada cada página oficial y qué se ha observado.

El wiki del repositorio refleja el estado en que se publicó. Después de tu primera descarga, pide un chequeo del wiki para que quede al día con tus archivos.

Si usas otro agente, `CLAUDE.md` y `.claude/skills/revisar-tablas/SKILL.md` son texto plano y sirven como instrucciones para cualquiera.

## Estructura del proyecto

```
├── README.md
├── CLAUDE.md                   Instrucciones del agente
├── .claude/skills/revisar-tablas/SKILL.md
├── scripts/
│   ├── revisar_tablas.ps1      Revisa, descarga, archiva y convierte
│   └── instalar_tarea.ps1      Programa la revisión en Windows
├── config/correos.txt          Destinatarios del aviso
├── estado/estado.json          Qué versión se tiene de cada tabla
├── wiki/                       Historial y notas, mantenido por el agente
│
│   Se crean al usarlo y no se suben al repositorio:
├── tablas/                     Versión vigente de cada tabla (.xlsx)
├── originales/                 Archivos tal cual se descargaron, si hubo que convertirlos
└── historico/                  Versiones anteriores
```

## Problemas frecuentes

| Qué ves | Qué hacer |
|---|---|
| «La ejecución de scripts está deshabilitada» | Usa el comando completo, con `-ExecutionPolicy Bypass` |
| «AVISO de formato: … no se pudo convertir» | El computador no tiene Excel ni LibreOffice. La tabla está al día, pero en su formato original |
| «La página … no funciona» | La página oficial está caída. No se tocó nada; se reintenta en la próxima revisión |
| «No se encontró en la página el enlace…» | La página oficial cambió de estructura y hay que ajustar el script. Con el agente, pídele que lo revise |
| Borraste `estado/estado.json` | Recupéralo del repositorio; sin él, el script no sabe qué versión tiene |

## Limitaciones

- Solo funciona en Windows.
- La conversión real con Excel no se ha probado todavía; el resto del recorrido sí.
- Depende de cómo estén armadas hoy las páginas de Minsalud e INVIMA. Si cambian, el script avisa en vez de descargar algo equivocado.
