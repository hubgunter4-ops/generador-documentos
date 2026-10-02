# Notas de verificación

- La aplicación usa una sola plantilla `.docx` y un Excel con una fila por documento.
- Se admiten plantillas Word de varias páginas; cada fila recibe una copia completa.
- Las copias se combinan en el orden de las filas y se separan con un salto de página en un único DOCX.
- Los únicos campos canónicos son `persona_nombre`, `mis_padres`, `mis_padrinos`, `abuelos_paternos`, `abuelos_maternos`, `el_sr` y `la_sra`.
- Los campos compuestos vienen directamente de Excel; no se vuelven a calcular ni se mezclan.
- Se rechazan columnas Excel repetidas después de normalizar llaves y se muestran las columnas desconocidas como no utilizadas.
- Los marcadores de Word no reconocidos y los campos presentes en Word pero ausentes en Excel se muestran antes de generar.
- `pnpm check`, `pnpm build` y la validación del DOCX completaron correctamente antes de esta migración.
