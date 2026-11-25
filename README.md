# Consulta Masiva RUC

Aplicación web estática para consultar listas de RUC peruanos con datos derivados del Padrón Reducido del RUC de SUNAT. Está preparada para publicarse en GitHub Pages y no necesita un backend ni una instalación local.

## Funciones

- Tres formas de entrada: archivo TXT/CSV, libro Excel o lista pegada/escrita directamente en la página.
- En TXT/CSV y en la lista pegada se admite un RUC por línea, o varios separados por espacios, comas, punto y coma o barras verticales.
- En Excel se utiliza la primera hoja. Si existe una columna cuyo encabezado es `RUC`, se usa; en caso contrario, se utiliza la primera columna.
- Conserva el orden de entrada y las apariciones duplicadas.
- Mantiene las entradas con formato inválido y los números que no aparecen en la copia consultada.
- Distingue los registros `Encontrado`, `BAJA DEFINITIVA` e `Identificador interno` (cuando el estado del registro figura como `NUM. INTERNO IDENTIF`).
- Conserva el domicilio que figure en el registro, incluso para personas naturales. El dato se identifica como información procedente del padrón y no se verifica de forma independiente.
- Descarga Excel `.xlsx` con hojas `Resultados` y `Resumen`, colores por resultado, filtros y primera fila inmovilizada.
- Descarga CSV UTF-8 con separador punto y coma. El CSV contiene las filas de resultados; no permite estilos ni hojas adicionales.
- Procesa las listas localmente en el navegador. El archivo de entrada no se envía a un servidor de esta aplicación.

## Fuente de datos y funcionamiento

La fuente original es el [Padrón Reducido del RUC de SUNAT](https://www.sunat.gob.pe/descargaPRR/mrc137_padron_reducido.html).

Para no descargar el padrón completo en cada visita, la aplicación consulta el espejo público [`alb3rt0ru1z/tribio-padron-ruc`](https://github.com/alb3rt0ru1z/tribio-padron-ruc), que distribuye fragmentos JSON mediante jsDelivr y GitHub Raw como respaldo. La aplicación obtiene el manifiesto y descarga solo los fragmentos correspondientes a los primeros cinco dígitos de los RUC incluidos en la lista.

**Dependencia externa:** la aplicación no administra ese espejo. Su disponibilidad, estructura y fecha de actualización dependen de su mantenedor y de los servicios de distribución. La página muestra la fecha `updated_at` que informa el manifiesto y advierte cuando la copia tiene más de tres días. Esa fecha es la fecha declarada de generación de la copia, no una confirmación independiente de la fecha oficial de corte de SUNAT.

La aplicación no llama a una API oficial de consulta individual ni verifica en tiempo real el estado actual del contribuyente. Los datos son una copia del padrón y pueden no reflejar cambios recientes. Un resultado «No encontrado en el padrón» no demuestra por sí solo que el RUC sea inexistente.

### Respaldo ante errores del CDN

Cada fragmento se solicita primero a jsDelivr. Si la descarga responde con HTTP 403/429, otro error HTTP, un fallo de red o una respuesta JSON inválida, la aplicación intenta obtener el mismo archivo desde GitHub Raw. Se consultan como máximo dos fragmentos simultáneamente. Si ambas fuentes fallan, la consulta se detiene y no genera archivos incompletos.

## Colores del Excel

| Resultado | Color |
|---|---|
| Encontrado | Verde |
| BAJA DEFINITIVA | Rojo claro |
| Identificador interno | Azul |
| No encontrado en el padrón | Ámbar |
| Formato inválido | Rojo |

La primera fila de `Resultados` queda inmovilizada al desplazarse. La hoja cuenta con filtros en sus encabezados.

## Privacidad

Los archivos y textos de entrada se leen y procesan en el navegador; no se cargan en un servidor de esta aplicación. Para consultar, el navegador solicita al CDN los fragmentos correspondientes a los prefijos de cinco dígitos necesarios. El proveedor del CDN puede registrar esas solicitudes. El sitio no almacena las listas ni los archivos generados.

## Publicar en GitHub Pages

1. En GitHub, crea un repositorio público, por ejemplo `consulta-masiva-ruc`.
2. Extrae el ZIP del proyecto y sube su contenido (no el ZIP como único archivo). `index.html` debe quedar en la raíz del repositorio.
3. Abre **Settings → Pages**.
4. En **Build and deployment**, elige **Deploy from a branch**.
5. Selecciona la rama `main` y la carpeta `/(root)`. Pulsa **Save**.
6. GitHub mostrará la URL pública cuando termine la publicación.

La dirección tendrá esta estructura: `https://TU-USUARIO.github.io/consulta-masiva-ruc/`.

Para actualizar una página ya publicada, sube los archivos nuevos en la misma raíz y confirma los cambios. GitHub Pages volverá a publicar el sitio.

## Bibliotecas

- [SheetJS Community Edition 0.20.3](https://docs.sheetjs.com/docs/getting-started/installation/standalone/) para leer archivos Excel.
- [ExcelJS 4.4.0](https://github.com/exceljs/exceljs) para generar el libro Excel con estilos y paneles inmovilizados. Se carga al solicitar la descarga Excel, con distribuidores alternativos.
- JavaScript nativo para lectura de texto, agrupación de RUC, solicitudes `fetch` y generación CSV.

## Estructura

```text
ConsultaMasivaRUC_Web/
├── index.html
├── styles.css
├── app.js
├── favicon.svg
├── .nojekyll
├── README.md
├── LICENSE
├── CITAS_Y_FUENTES.md
└── SUBIR_A_GITHUB.md
```

## Aviso

Proyecto independiente, no afiliado, patrocinado ni respaldado por SUNAT. La información es referencial y no sustituye la consulta oficial ni la Ficha RUC.
