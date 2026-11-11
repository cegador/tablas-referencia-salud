# INVIMA — listados de Código Único de Medicamentos

- **Página:** https://www.invima.gov.co/productos-vigilados/medicamentos-y-productos-biologicos/autorizacion-de-comercializacion-registros
- **Qué se busca:** cuatro enlaces que empiezan por «Listado Código Único de Medicamentos…».
- **Tablas que tomamos:** [[cum-vigentes]], [[cum-vencidos]], [[cum-renovacion]] y [[cum-otros-estados]].
- No pide usuario.

## Cómo se descarga

Cada enlace lleva a una ficha intermedia (por ejemplo `/biblioteca/listadocodigounicovigentesoctubre2026xlsx`) y esa ficha tiene el botón de descarga real (`/biblioteca/download/<número>`). El archivo llega con el nombre `ListadoCodigoUnico<Tipo><Mes><Año>.xlsx`.

## Cómo se reconoce una versión nueva

El mes y el año van en el propio enlace. Hay versión nueva cuando el mes del enlace es posterior al del archivo que tenemos. Cada listado se controla por separado, por si el INVIMA no publica los cuatro el mismo día.

## Observaciones

- El listado de un mes aparece durante ese mismo mes, pero no el día 1: el 11 de noviembre de 2026 todavía estaban los de octubre.
- Vigentes, Vencidos y Otros Estados pesan entre 15 y 21 MB; Renovación, unos 200 KB.
