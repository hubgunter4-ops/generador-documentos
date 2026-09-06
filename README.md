# Generador de documentos por expediente

Aplicación estática local-first para cargar tres plantillas `.docx`, definir un rango de 50 a 500 expedientes, introducir variables familiares en un solo cuadro y descargar documentos Word o ZIP.

## Rango de expedientes

Se puede escribir un número (`50`) o un rango (`1-50`, `101-150`). La aplicación crea expedientes independientes y conserva el número indicado.

## Variables y correlación

La entrada compartida reconoce, entre otras, estas etiquetas:

- `(yo)` → `{persona_nombre}` en los documentos 1, 2 y 3.
- `(papá)` o `(padre)` → `{padre_nombre}`.
- `(mamá)` o `(madre)` → `{madre_nombre}`.
- `(abuelo paterno)` → `{abuelo_paterno}`.
- `(abuela paterna)` → `{abuela_paterna}`.
- `(abuelo materno)` → `{abuelo_materno}`.
- `(abuela materna)` → `{abuela_materna}`.
- `(padrino)` → `{padrino_nombre}`.
- `(madrina)` → `{madrina_nombre}`.

Los grupos compuestos `(mis padres)` y `(mis padrinos)` aceptan dos líneas en el orden definido. También se acepta el separador `E` cuando el usuario escribe una sola línea compuesta. Al generar documentos se construyen `{mis_padres}` y `{mis_padrinos}` con saltos de línea.

Los abuelos se almacenan como cuatro variables independientes y se incorporan únicamente en las plantillas que contengan esos marcadores. Si un documento no tiene los marcadores de abuelos, no se le agregan datos de abuelos.

## Vista previa

Después de seleccionar un expediente y un documento, la sección **Vista previa de la modificación** muestra cada marcador detectado en la plantilla y el valor que recibirá. Las palabras normales de Word no se sustituyen automáticamente.

## Privacidad

Los DOCX se procesan dentro del navegador. Las plantillas y expedientes se guardan en IndexedDB del perfil local; los respaldos se exportan manualmente como JSON. No existe servidor, autenticación ni sincronización remota.

GitHub Pages puede publicar el código, pero no es un repositorio privado de expedientes. No coloque documentos personales ni respaldos JSON dentro del repositorio.
