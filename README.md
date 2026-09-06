# Generador de documentos por expediente

Aplicación estática local-first para cargar tres plantillas `.docx`, definir un rango de 50 a 500 expedientes, introducir variables en un solo cuadro y descargar documentos Word o ZIP.

## Rango de expedientes

Se puede escribir un número (`50`) o un rango (`1-50`, `101-150`). La aplicación crea expedientes independientes y conserva el número indicado.

## Variables y correlación

La entrada compartida reconoce:

- `(yo)` → `{persona_nombre}` en los documentos 1, 2 y 3.
- `(papá)` o `(padre)` → `{padre_nombre}`.
- `(mamá)` o `(madre)` → `{madre_nombre}`.
- `(abuelo paterno)` → `{abuelo_paterno}`.
- `(abuela paterna)` → `{abuela_paterna}`.
- `(abuelo materno)` → `{abuelo_materno}`.
- `(abuela materna)` → `{abuela_materna}`.
- `(padrinos)` o `(mis padrinos)` → **dos nombres de padrinos**, capturados en dos líneas o unidos por `E`.

Los padrinos se ingresan una sola vez. No se capturan por separado como `(padrino)` y `(madrina)`.

La plantilla define cómo se presentan:

| Marcador Word | Resultado |
|---|---|
| `{el_sr}` | Primer nombre de padrino, con el tratamiento `El Sr.` colocado en la plantilla |
| `{la_sra}` | Segundo nombre de padrino, con el tratamiento `La Sra.` colocado en la plantilla |
| `{mis_padrinos}` | Los dos nombres, en dos líneas y en el orden capturado |

Así, un documento puede usar `{el_sr}` y `{la_sra}`, mientras otro utiliza `{mis_padrinos}`, sin volver a solicitar los nombres.

Los grupos `(mis padres)`, `(abuelos paternos)` y `(abuelos maternos)` aceptan dos líneas en el orden definido. Los abuelos se incorporan únicamente en las plantillas que contengan sus marcadores; si un documento no los solicita, no se le agregan datos.

## Privacidad

Los DOCX se procesan dentro del navegador. Las plantillas y expedientes se guardan en IndexedDB del perfil local; los respaldos se exportan manualmente como JSON. No existe servidor, autenticación ni sincronización remota.

GitHub Pages puede publicar el código, pero no es un repositorio privado de expedientes. No coloque documentos personales ni respaldos JSON dentro del repositorio.
