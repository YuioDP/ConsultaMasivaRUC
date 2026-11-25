# Fuentes y notas técnicas

## Fuente oficial de datos

- SUNAT, Padrón Reducido del RUC: https://www.sunat.gob.pe/descargaPRR/mrc137_padron_reducido.html
- SUNAT Orientación, campos publicados: https://orientacion.sunat.gob.pe/7-padron-reducido-del-ruc

La descripción oficial del padrón indica que contiene RUC, nombre o razón social, estado del contribuyente, condición de domicilio, ubigeo y domicilio fiscal.

## Distribución de fragmentos JSON

- Repositorio de datos: https://github.com/alb3rt0ru1z/tribio-padron-ruc
- Manifiesto actual: https://cdn.jsdelivr.net/gh/alb3rt0ru1z/tribio-padron-ruc@main/manifest.json
- CDN: https://www.jsdelivr.com/

El script de generación del repositorio citado particiona el padrón por los primeros cinco dígitos del RUC y produce archivos `chunks/XXXXX.json`. La aplicación se adapta a las columnas declaradas en cada fragmento.

## Biblioteca Excel

- SheetJS Community Edition, documentación oficial de scripts para navegador: https://docs.sheetjs.com/docs/getting-started/installation/standalone/
- Versión fijada: 0.20.3

## Respaldo de descarga

La aplicación utiliza como respaldo los archivos públicos de GitHub Raw:

- https://raw.githubusercontent.com/alb3rt0ru1z/tribio-padron-ruc/main/manifest.json
- https://raw.githubusercontent.com/alb3rt0ru1z/tribio-padron-ruc/main/chunks/20506.json (ejemplo de fragmento)

El respaldo se usa únicamente cuando la descarga principal mediante jsDelivr falla.
