# Minsalud — tablas de referencia de UPC

- **Página:** https://www.minsalud.gov.co/salud/POS/Paginas/unidad-de-pago-por-capitacion-upc.aspx
- **Qué se busca:** el botón «Tablas de referencia», enlace «Vigencia AAAA (descargable)».
- **Qué se descarga:** un zip llamado `tablas-vigencia-AAAA.zip` (unos 4,6 MB). No pide usuario.
- **Tablas que tomamos:** [[cups]], [[cie-10]] y [[reps]].

## Cómo se reconoce una versión nueva

1. Aparece un enlace con un año mayor (por ejemplo `tablas-vigencia-2027.zip`).
2. O el zip del mismo año fue modificado. En ese caso se comparan las tres tablas una por una y solo se reemplaza la que de verdad cambió.

## Observaciones

- CUPS viene en `.xlsb`; CIE-10 y REPS, en `.xlsx`. El script pasa CUPS a `.xlsx` y guarda el original en `originales/`. Para saber si Minsalud cambió la tabla se compara siempre contra el original, no contra el convertido.

- El zip de 2026 fue modificado por Minsalud el 22 de mayo de 2026, a mitad de año, sin cambiar de nombre. Ese cambio fue en otra tabla (productos de soporte nutricional); CUPS, CIE-10 y REPS siguen siendo las de febrero de 2026. Por eso no basta con mirar el año: también se vigila la fecha de modificación.
- El zip trae otras tablas que hoy no usamos: unidades de concentración, unidades de dispensación, forma farmacéutica, insumos y dispositivos, productos de soporte nutricional.
- Dentro del zip la carpeta se llama `TABLAS DE REFENCIA 2026` (así, con el error de escritura) y hay una carpeta `__MACOSX` que se ignora.
