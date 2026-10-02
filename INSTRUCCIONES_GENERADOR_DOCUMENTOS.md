# Instrucciones de uso

1. Prepare una plantilla Word `.docx` de una o varias páginas.
2. En Word, coloque marcadores como `{yo}`, `{mis_padres}`, `{mis_padrinos}`, `{abuelos_paternos}`, `{abuelos_maternos}` y `{el_sr_y_la_sra}`. La plantilla también puede usar categorías visibles sin llaves: `mis_padres`, `mis_padrinos`, `abuelos_paternos`, `el_sr_y_la_sra` y `yo`, seguidas de líneas de captura.
3. Prepare un Excel con una fila de encabezados y una fila por documento. Use exactamente los mismos nombres, sin llaves: `yo`, `mis_padres`, `mis_padrinos`, `abuelos_paternos` y `el_sr_y_la_sra`.
4. Abra la aplicación y seleccione la plantilla Word.
5. Seleccione el archivo Excel.
6. Revise la correspondencia: los marcadores verdes tienen una columna; los amarillos requieren atención.
7. Escriba el nombre del archivo final y pulse **Generar un Word único**.

Cada fila de Excel produce una copia completa de la plantilla. Las copias se unen en el orden de las filas y se separan con un salto de página. Una plantilla de cuatro páginas con tres filas produce doce páginas, salvo que el contenido provoque saltos adicionales.

## Datos y campos

Los campos compuestos (`mis_padres`, `mis_padrinos`, `abuelos_paternos`, `el_sr_y_la_sra` y `yo`) vienen ya preparados en Excel. La aplicación los pasa a Word tal como están escritos y no solicita datos adicionales ni los combina con otros nombres.

La columna `persona_nombre` corresponde al marcador `{persona_nombre}`. No uses etiquetas genéricas o columnas duplicadas; el programa valida las columnas repetidas y avisa de nombres no reconocidos.

## Privacidad

El Excel y el DOCX se procesan localmente en el navegador. No se envían a un servidor ni deben guardarse dentro del repositorio.
