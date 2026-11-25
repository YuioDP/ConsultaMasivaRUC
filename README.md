# Consulta Masiva RUC

Aplicación web estática para consultar listas de RUC peruanos con datos derivados del Padrón Reducido del RUC de SUNAT. Está preparada para publicarse en GitHub Pages y no necesita un backend ni una instalación local.

## Funciones

- Entrada TXT, CSV y libros Excel (`.xls`, `.xlsx`, `.xlsm`, `.xlsb`).
- En TXT se detectan los RUC de 11 dígitos. Se admite un RUC por línea y también varias apariciones en una línea.
- En Excel se usa la primera hoja. Si existe una columna cuyo encabezado es `RUC`, se utiliza; en caso contrario, se utiliza la primera columna.
- Conserva el orden de entrada y las apariciones duplicadas.
- Mantiene las entradas con formato inválido y los RUC que no aparecen en la copia consultada.
- Exporta siempre un libro `.xlsx` con dos hojas: `Resultados` y `Resumen`.
- Procesa los archivos localmente en el navegador. El archivo de entrada no se envía a un servidor de la aplicación.

## Fuente de datos y funcionamiento

La fuente original es el [Padrón Reducido del RUC de SUNAT](https://www.sunat.gob.pe/descargaPRR/mrc137_padron_reducido.html).

Para no descargar el padrón completo en cada visita, la aplicación consulta el espejo público [`alb3rt0ru1z/tribio-padron-ruc`](https://github.com/alb3rt0ru1z/tribio-padron-ruc), que distribuye fragmentos JSON mediante jsDelivr. La aplicación obtiene el manifiesto y descarga solo los fragmentos correspondientes a los primeros cinco dígitos de los RUC incluidos en la lista.

**Dependencia externa:** la aplicación no administra ese espejo. Su disponibilidad, estructura y fecha de actualización dependen de su mantenedor y de jsDelivr. La página muestra la fecha `updated_at` que informa el manifiesto y advierte cuando la copia tiene más de tres días. Esa fecha es la fecha declarada de generación de la copia, no una confirmación independiente de la fecha oficial de corte de SUNAT.

**Verificación al preparar esta versión (24/11/2026):** el manifiesto público consultado informaba `updated_at: 2026-11-11T18:40:55.496Z`, 18.522.303 registros y 1.972 prefijos. Por tanto, en este momento la aplicación debe mostrar una advertencia de antigüedad. Aunque el README del espejo describe una tarea diaria, no se debe asumir que se ejecutó correctamente hasta comprobar una fecha reciente en el manifiesto. La aplicación no oculta esta situación ni etiqueta esa fecha como fecha oficial de SUNAT.

La aplicación no llama a una API oficial de consulta individual ni verifica en tiempo real el estado actual del contribuyente. Los datos son una copia del padrón y pueden no reflejar cambios recientes. Un resultado «No encontrado en el padrón» no demuestra por sí solo que el RUC sea inexistente.

### Respaldo ante errores del CDN

Cada fragmento se solicita primero a jsDelivr. Si la descarga responde con HTTP 403/429, otro error HTTP, un fallo de red o una respuesta JSON inválida, la aplicación intenta obtener el mismo archivo desde GitHub Raw. Se consultan como máximo dos fragmentos simultáneamente para moderar el volumen de solicitudes. Si ambas fuentes fallan, la consulta se detiene y no genera un Excel incompleto. Esto mejora la tolerancia a fallos, pero no elimina la dependencia de servicios externos.

## Privacidad

El archivo de entrada se lee y procesa en el navegador; no se carga en un servidor de esta aplicación. Para consultar, el navegador solicita al CDN los fragmentos correspondientes a los prefijos de cinco dígitos necesarios. Por tanto, el proveedor del CDN puede registrar esas solicitudes. El sitio no almacena las listas ni los archivos Excel generados.

## Publicar en GitHub Pages

1. Descarga y extrae el ZIP del proyecto.
2. En GitHub, crea un repositorio nuevo llamado, por ejemplo, `consulta-masiva-ruc`.
3. Puedes elegir `Public` y no necesitas inicializarlo con README, `.gitignore` o licencia.
4. Dentro del repositorio, elige **Add file → Upload files**.
5. Arrastra el contenido de esta carpeta (no subas el ZIP como único archivo). `index.html` debe quedar en la raíz del repositorio.
6. Escribe un mensaje como `Publicar primera versión` y pulsa **Commit changes**.
7. Abre **Settings → Pages**.
8. En **Build and deployment**, selecciona **Deploy from a branch**.
9. En **Branch**, selecciona `main` y la carpeta `/ (root)`. Pulsa **Save**.
10. Cuando finalice la publicación, GitHub mostrará la dirección del sitio. Normalmente tendrá este formato: `https://TU-USUARIO.github.io/consulta-masiva-ruc/`.

También se incluye una guía independiente: [`SUBIR_A_GITHUB.md`](SUBIR_A_GITHUB.md).

Si el repositorio se llama exactamente `TU-USUARIO.github.io`, la dirección será `https://TU-USUARIO.github.io/`.

## Probar localmente

Se puede abrir `index.html` directamente para revisar el diseño y la lectura de archivos. Para probar la consulta real, es preferible servir la carpeta mediante un servidor estático local (por ejemplo, la extensión Live Server de Visual Studio Code), porque así se reproducen mejor las condiciones de una página publicada.

## Bibliotecas

- [SheetJS Community Edition 0.20.3](https://docs.sheetjs.com/docs/getting-started/installation/standalone/) para leer y generar Excel. Se carga desde el CDN oficial de SheetJS.
- JavaScript nativo para lectura de TXT, agrupación de RUC y llamadas `fetch` a los fragmentos JSON.

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
└── CITAS_Y_FUENTES.md
```

## Nota sobre licencias de fuentes externas

La licencia MIT incluida cubre el código propio de esta aplicación. No concede derechos sobre el contenido de SUNAT ni sobre los archivos transformados y distribuidos por repositorios externos. El repositorio espejo se consulta como fuente externa y no se copia dentro de este proyecto; antes de usarlo con fines comerciales o redistribuir sus datos, revisa las condiciones aplicables y confirma con sus responsables si corresponde.

## Aviso

Proyecto independiente, no afiliado, patrocinado ni respaldado por SUNAT. La información se ofrece como herramienta de consulta referencial y no sustituye la consulta oficial ni la Ficha RUC.
