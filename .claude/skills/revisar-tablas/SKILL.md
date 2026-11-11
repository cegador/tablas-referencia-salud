---
name: revisar-tablas
description: Revisa las páginas de Minsalud e INVIMA, descarga las tablas de referencia nuevas (CUPS, CIE-10, REPS y listados de Código Único de Medicamentos), actualiza el wiki y envía el correo de aviso. Usar cuando digan «revisa las tablas», «hay tablas nuevas», «actualiza las tablas», o cuando lo dispare la tarea programada.
---

# Revisar tablas

Sigue los pasos en orden. Háblale a la persona en español y sin tecnicismos.

## 1. Ejecutar la revisión

Desde la carpeta del proyecto:

```
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\revisar_tablas.ps1
```

Añade `-Forzar` solo si la persona pidió expresamente revisar ahora (por ejemplo «revisa ya», «fuerza la revisión»). La tarea programada nunca usa `-Forzar`.

El script puede tardar varios minutos si hay descargas (los listados del INVIMA pesan unos 20 MB cada uno).

## 2. Leer el resultado

Lee `estado/ultimo_resultado.json`. Cada fuente trae un `estado`:

| Estado | Qué significa |
|---|---|
| `omitido` | Hoy no tocaba revisar esa fuente |
| `sin_cambios` | Se revisó y no hay versión nueva |
| `actualizado` | Se descargó al menos una tabla nueva (ver `novedades`) |
| `error` | La página no funciona o cambió (ver `errores`) |
| `actualizado_con_errores` | Unas tablas se actualizaron y otras fallaron |

El resultado trae además dos listas sobre el formato de los archivos:

- `conversiones`: tablas que llegaron en otro formato y el script pasó a `.xlsx` (el original queda en `originales/`).
- `avisos`: tablas que no están en `.xlsx` y no se pudieron convertir, normalmente porque el computador no tiene Excel ni LibreOffice.

**Si las dos fuentes están en `omitido` y las dos listas están vacías, termina aquí**: no escribas en el wiki ni envíes correo.

Si hay conversiones o avisos, cuéntalos en el log, en la página de la tabla (el nombre del archivo vigente cambia a `.xlsx`) y en el correo. Un aviso de formato no es una falla de la página: la tabla está al día, solo que en su formato original.

## 3. Si hubo un error

- No muevas, borres ni reemplaces ninguna tabla a mano.
- Abre la página de la fuente en el navegador y mira qué pasó: si no carga, si el enlace cambió de nombre o de lugar. Anota lo que veas en `wiki/fuentes/<fuente>.md` y cuéntalo en el correo.
- Si el problema es que la página cambió de estructura, di en el correo que hay que ajustar el script, pero no lo modifiques sin que LCR lo pida.

## 4. Actualizar el wiki

- **`wiki/log.md`**: añade al final una entrada con este formato exacto de título:
  `## [AAAA-MM-DD] revisión | <resumen corto>`
  y debajo una línea por fuente revisada (no menciones las omitidas).
- **Por cada novedad**: en `wiki/tablas/<tabla>.md` cambia la versión vigente, el nombre del archivo y la fecha de descarga, y añade la versión anterior a la tabla «Historial» con la carpeta de `historico/` donde quedó.
- **`wiki/index.md`**: actualiza la columna de versión de las tablas que cambiaron y la fecha de «Última revisión».
- Si no hubo novedades, solo se tocan `log.md` y la fecha de «Última revisión» del índice.

## 5. Enviar el correo

Lee `config/correos.txt` (una dirección por línea; las líneas que empiezan con `#` no cuentan).

- **Si no hay direcciones**: el correo está en pausa. No envíes nada y menciónalo en una sola línea en el resumen final.
- **Si hay direcciones**: envía **un solo correo** a todas, con el conector de Gmail. Solo a las direcciones de ese archivo, a nadie más.

Asunto, según el caso:

- `Tablas de referencia: sin cambios (AAAA-MM-DD)`
- `Tablas de referencia: hay versiones nuevas (AAAA-MM-DD)`
- `Tablas de referencia: una página no funciona (AAAA-MM-DD)`

Cuerpo: corto y en lenguaje sencillo. Para cada fuente revisada, di qué pasó. Si hay tablas nuevas, lista el nombre de la tabla, el archivo nuevo y la versión que reemplaza. Si hubo un error, di qué página falló, que se conservó la última versión buena y que se volverá a intentar en la próxima revisión. Cierra indicando que las tablas vigentes están en la carpeta `tablas` y las anteriores en `historico`.

## 6. Cerrar

Resume en dos o tres líneas lo que pasó: qué se revisó, qué cambió y a quién se le envió el correo.
