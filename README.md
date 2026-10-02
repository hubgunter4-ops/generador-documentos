# Generador de documentos desde Excel

Aplicación local-first para cargar **una plantilla Word `.docx` de varias páginas** y un archivo Excel `.xlsx`/`.xls`. Cada fila del Excel genera una copia completa de la plantilla; todas las copias se unen en un único documento Word descargable.

## Flujo

```text
Plantilla Word de varias páginas + Excel con una fila por registro
                              ↓
              Un documento Word final con todas las copias
```

Los archivos se procesan dentro del navegador. No se suben a un servidor ni se guardan en GitHub.

## Columnas Excel y marcadores Word

Usa exactamente estos nombres como encabezados de Excel y marcadores de Word:

| Campo | Encabezado Excel | Marcador Word |
|---|---|---|
| Yo | `yo` | `{yo}` |
| Mis padres | `mis_padres` | `{mis_padres}` |
| Mis padrinos | `mis_padrinos` | `{mis_padrinos}` |
| Abuelos paternos | `abuelos_paternos` | `{abuelos_paternos}` |
| Abuelos maternos | `abuelos_maternos` | `{abuelos_maternos}` |
| El Sr. y la Sra. | `el_sr_y_la_sra` | `{el_sr_y_la_sra}` |

Los campos compuestos se toman directamente de Excel. La aplicación no intenta reconstruirlos, separarlos ni mezclarlos con otros campos. También acepta `{persona_nombre}` como alias de `{yo}` cuando se usa una plantilla anterior.

> Cada fila representa un documento. Si la plantilla tiene cuatro páginas y Excel tiene tres filas, el resultado tendrá las tres copias completas, separadas por saltos de página.

## Reglas para evitar confusiones

- Cada campo debe tener una sola columna en Excel.
- No repitas columnas equivalentes con nombres distintos; usa los siete nombres canónicos anteriores.
- Los marcadores deben escribirse con una sola llave a cada lado, por ejemplo `{mis_padres}`.
- Los marcadores desconocidos se muestran como advertencia antes de generar.
- Las columnas Excel no reconocidas se muestran como no utilizadas y no se copian a Word.
- Si falta una columna correspondiente a un marcador, ese marcador quedará vacío y se mostrará una advertencia.

## Uso

1. Abre la aplicación.
2. Selecciona una plantilla `.docx` de varias páginas.
3. Selecciona el archivo Excel con una fila por registro.
4. Revisa los marcadores detectados, las columnas reconocidas y la primera fila de muestra.
5. Define el nombre del documento final.
6. Pulsa **Generar un Word único**.

## Desarrollo

```bash
pnpm install --frozen-lockfile
pnpm check
pnpm build
```

El build genera `dist/public` para GitHub Pages y `dist/index.js` para el servidor de preview.

## Privacidad

El procesamiento del Excel y del DOCX ocurre completamente en el navegador mediante SheetJS, PizZip y Docxtemplater. No existe servidor de datos ni autenticación. No subas datos personales al repositorio.


## Aplicación de escritorio Tkinter

La aplicación de escritorio está en [`desktop/combina-word`](desktop/combina-word). Carga primero el DOCX, detecta sus marcadores o categorías visibles, después carga el Excel, detecta el número de registros y genera una copia de la plantilla por cada fila seleccionada.

El workflow [`Build desktop installers`](.github/workflows/build-desktop.yml) ejecuta las pruebas y construye automáticamente:

- `combina-word-deb`: paquete `.deb` para Debian/Ubuntu.
- `combina-word-exe`: ejecutable `.exe` para Windows.

Los archivos se descargan desde la sección **Artifacts** de la ejecución de GitHub Actions. La compilación se activa cuando cambian los archivos de `desktop/combina-word` o manualmente desde **Actions → Build desktop installers → Run workflow**.
