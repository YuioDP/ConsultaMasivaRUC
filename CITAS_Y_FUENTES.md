# Fuentes y notas técnicas

## Fuente oficial de datos

- SUNAT, Padrón Reducido del RUC: https://www.sunat.gob.pe/descargaPRR/mrc137_padron_reducido.html
- SUNAT Orientación, campos publicados: https://orientacion.sunat.gob.pe/7-padron-reducido-del-ruc

El padrón oficial incluye RUC, nombre o razón social, estado del contribuyente, condición de domicilio, ubigeo y domicilio fiscal. La aplicación muestra los campos utilizados en su interfaz y exportaciones, conservando el domicilio cuando consta en el registro del fragmento consultado.

## Distribución de fragmentos JSON

- Repositorio de datos de terceros: https://github.com/alb3rt0ru1z/tribio-padron-ruc
- Manifiesto distribuido por jsDelivr: https://cdn.jsdelivr.net/gh/alb3rt0ru1z/tribio-padron-ruc@main/manifest.json
- Distribuidor CDN: https://www.jsdelivr.com/
- Acceso de respaldo mediante GitHub Raw: https://raw.githubusercontent.com/alb3rt0ru1z/tribio-padron-ruc/main/manifest.json

El script de generación del repositorio citado particiona el padrón por los primeros cinco dígitos del RUC y produce archivos `chunks/XXXXX.json`. La aplicación obtiene la fecha declarada de generación desde `manifest.json`.

**Limitación:** este repositorio es mantenido por un tercero. Aunque su propósito sea transformar el padrón oficial, la disponibilidad y actualización efectiva de los fragmentos dependen de su proceso automatizado y de los distribuidores. La fecha del manifiesto no debe confundirse con una verificación independiente de la última publicación oficial de SUNAT.

## Bibliotecas web

- SheetJS Community Edition, documentación de scripts para navegador: https://docs.sheetjs.com/docs/getting-started/installation/standalone/
- SheetJS versión fijada: 0.20.3
- ExcelJS, repositorio y documentación: https://github.com/exceljs/exceljs
- ExcelJS versión fijada: 4.4.0. Se carga bajo demanda desde cdnjs, jsDelivr o unpkg, según disponibilidad.

## Notas de clasificación

- `Encontrado`: el número ingresado coincide con un registro del fragmento consultado y su estado no es `BAJA DEFINITIVA` ni `NUM. INTERNO IDENTIF`.
- `BAJA DEFINITIVA`: el registro coincide y el campo estado contiene ese valor.
- `Identificador interno`: el registro coincide y su estado figura como `NUM. INTERNO IDENTIF`. Se conservan los campos del registro, incluido el domicilio que allí conste.
- `No encontrado en el padrón`: no se halló coincidencia en los fragmentos consultados. No acredita por sí solo que el número sea inexistente.
- `Formato inválido`: el elemento de entrada no pudo interpretarse como una secuencia de once dígitos.

El domicilio se reproduce desde el registro consultado. La aplicación no realiza una verificación independiente del dato ni determina por sí misma si es un domicilio fiscal vigente.
