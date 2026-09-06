# Generador de documentos por expediente

Aplicación estática local-first para cargar tres plantillas `.docx`, crear entre 50 y 500 expedientes, introducir datos familiares mediante etiquetas y descargar documentos Word o ZIP.

## Inicio

```bash
pnpm install
pnpm dev
```

En la interfaz puede crear un lote con las tres plantillas de prueba integradas o cargar sus propias plantillas `.docx`.

## Marcadores de Word

Use marcadores completos, sin espacios ni acentos en las claves:

- `{persona_nombre}`
- `{padre_nombre}`
- `{madre_nombre}`
- `{abuelos_maternos}`
- `{abuelos_paternos}`
- `{padrino_nombre}`
- `{madrina_nombre}`
- `{mis_padres}`
- `{mis_padrinos}`

## Entrada compartida

Ejemplo:

```text
(yo): Juan Pérez López
(papá): Carlos Pérez Hernández
(mamá): Ana López Martínez
(abuelos maternos):
Roberto López
Elena Martínez
(padrino): Miguel Torres Ramírez
(madrina): Laura Gómez Castillo
```

Las equivalencias aceptadas incluyen `yo`, `nombre`, `papá`, `papa`, `padre`, `mamá`, `mama`, `madre`, `padrino` y `madrina`. También se admiten bloques `(mis padres)` y `(mis padrinos)` con exactamente dos líneas no vacías.

## Privacidad

Los documentos se procesan dentro del navegador. Las plantillas y los expedientes se guardan en IndexedDB del perfil local y los respaldos se exportan manualmente como JSON. La aplicación no incluye servidor, autenticación ni sincronización remota.

GitHub Pages puede publicar el código, pero no es un repositorio privado de expedientes. No coloque documentos personales ni respaldos JSON dentro del repositorio.

## Estado de esta versión

Incluye creación configurable de lotes, parser contextual, modelo `sharedData`, actualización independiente de expedientes, generación DOCX, ZIP individual y masivo, exportación/importación JSON y plantillas sintéticas de prueba.

La validación más rigurosa de plantillas reales debe ejecutarse con sus documentos definitivos, especialmente si contienen marcadores divididos entre fragmentos XML, encabezados, pies de página o tablas complejas.
