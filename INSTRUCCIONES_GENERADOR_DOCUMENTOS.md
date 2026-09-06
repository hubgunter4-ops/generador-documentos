# Instrucciones de prueba

1. Abra la aplicación.
2. En la sección **Preparar plantillas**, pulse **Crear lote** sin cargar archivos para usar las tres plantillas de prueba.
3. Cambie el número de expedientes si desea un lote mayor; el rango permitido es de 50 a 500.
4. Seleccione un expediente.
5. Escriba los datos en el cuadro compartido utilizando etiquetas como `(yo)`, `(papá)`, `(mamá)`, `(padrino)` y `(madrina)`.
6. Espere el estado **Cambios guardados**.
7. Descargue un documento individual, el ZIP del expediente o el ZIP completo.
8. Use **Exportar JSON** para crear un respaldo editable antes de importar otro respaldo.

El cambio en un expediente no modifica los demás. Los padres y padrinos se guardan como personas individuales y los campos compuestos `mis_padres` y `mis_padrinos` se construyen al generar cada documento.

## Plantillas definitivas

Al crear un nuevo lote, cargue exactamente tres archivos `.docx`. Prepare los marcadores dentro de Word con las claves documentadas en el README. Las plantillas de un lote no se modifican al generar descargas.
