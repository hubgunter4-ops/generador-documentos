# Prueba de confirmación: Excel → Word

## Archivos

- `plantilla_confirmacion_excel_proporcionado.docx`: plantilla basada en el texto de la parroquia.
- `datos_sinteticos_familia.xlsx`: Excel proporcionado previamente.
- `resultado_confirmaciones_50_registros.docx`: documento combinado final.

## Verificaciones realizadas

- Registros leídos del Excel: **50**.
- Campos detectados en la plantilla:
  - `YO`
  - `MIS_PADRES`
  - `MIS_PADRINOS`
  - `ABUELOS_PATERNOS`
  - `EL_SR_Y_LA_SRA`
- Campos faltantes: **ninguno**.
- Documentos combinados: **50**.
- Encabezados de confirmación encontrados: **50**.
- Marcadores `{{...}}` sin reemplazar: **0**.
- Páginas verificadas mediante LibreOffice/PDF: **50**.
- Integridad ZIP del DOCX: **correcta**.
- Pruebas unitarias del proyecto: **4 aprobadas**.

La combinación inserta un salto de página entre registros para que cada fila del Excel ocupe una página independiente.

La prueba se ejecutó con la suite del proyecto (`pytest -q`), la combinación real Excel → Word y una conversión posterior a PDF para contar las páginas.
