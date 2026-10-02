# Verificación de la actualización

- La vista previa confirma una interfaz en español con carga de tres plantillas, rango escrito y relación visible de variables.
- La interfaz muestra que solo se edita el expediente seleccionado.
- Se documentan las variables individuales de padre, madre, padrino, madrina y cuatro abuelos.
- Los campos compuestos `mis_padres` y `mis_padrinos` se generan en dos líneas.
- Cada etiqueta individual tiene un destino único; `(nombre)` se rechaza por ser ambiguo.
- `(padrino)` y `(madrina)` se guardan por separado. Los grupos `(padrinos)` y `(mis padrinos)` requieren exactamente dos nombres.
- Repetir el mismo campo con el mismo valor es idempotente; un valor distinto genera advertencia y conserva el primer valor.
- Los campos compuestos (`mis_padres`, `mis_padrinos`, `abuelos_paternos`, `abuelos_maternos`, `el_sr`, `la_sra`) se calculan al generar el DOCX y no son entradas independientes.
- Los marcadores se detectan desde el contenido XML de la plantilla cargada para mostrarlos en la vista de correspondencias.
- `pnpm check` y `pnpm build` completaron correctamente.
