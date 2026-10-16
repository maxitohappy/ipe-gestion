# Gestión IPE

Aplicación web para la gestión mensual de **ingresos percibidos en exceso (IPE)** a partir del Excel de control.

- **Memos (APLICAR):** agrupa por empresa los IPE con estado `APLICAR` en *MIS APUNTES* y genera un memo en Word (.docx) por empresa con el formato institucional (Arial 11, firma de Eliana Rojas Silva). El Nº de memo queda en blanco para completarlo en Word.
  - Un IPE puede aplicarse a **varios documentos**: el saldo se descuenta en cada aplicación (columna *Monto* de la tabla) y, si queda diferencia, el memo agrega *"Nota: Diferencia de … mantener en la cuenta 'acreedores por ingresos percibidos en exceso' …"*.
- **Historial de memos:** cada memo descargado queda registrado con el saldo de cada IPE. Al cargar el Excel de otro mes, un IPE con memo previo se oculta solo si su saldo **no cambió** (Contabilidad aún no lo procesa); si cambió (aplicación parcial o nuevo saldo del cliente) vuelve a *Memos* indicando el memo anterior. Desde el historial se puede volver a descargar un memo, devolverlo a pendientes, exportar a Excel o respaldar/importar (el historial se guarda en el navegador).
- **Devoluciones:** consolida los IPE con estado `DEVOLVER` por empresa y descarga una planilla Excel (hojas *Resumen* y *Detalle*) con columnas para los datos bancarios.
- **Resumen y alertas:** saldo por estado, antigüedad y alertas (APLICAR sin documento a aplicar, sin justificación, estados no estándar, correos sin respuesta por más de 30 días, posibles duplicados, filas sin gestión).

**Privacidad:** el Excel se procesa solo en el navegador. Ningún dato se envía a internet ni se guarda en este repositorio.

## Uso mensual

1. Abre el Excel IPE del mes y verifica que tenga, al final de la hoja, las columnas:
   - `DOCTO A APLICAR` (ej. `Pagar FEE 466934`)
   - `MONTO A APLICAR` (ej. `701483`)

   Complétalas en las filas con estado `APLICAR` usando la deuda del cliente del sistema contable. Para varios documentos sepáralos con `;` en el mismo orden (ej. `Pagar FEE 263576; Pagar FEE 262729` y `2005; 2327`), o agrégalos en la app con **+ Agregar documento**.
2. Entra a la página publicada y arrastra el Excel.
3. Revisa la pestaña **Resumen y alertas**.
4. En **Memos**: revisa fecha e iniciales, marca o desmarca documentos, ajusta los documentos a aplicar y descarga cada memo o todos en un .zip. Los memos descargados pasan al **Historial**.
5. En **Devoluciones**: descarga la planilla para Tesorería.
6. En **Historial**: descarga un respaldo periódicamente.

## Uso sin internet (respaldo)

Si GitHub o internet no están disponibles, la app funciona igual desde el computador:

1. Descarga la versión local desde el enlace **"Descargar versión sin internet"** de la página (o https://github.com/maxitohappy/ipe-gestion/archive/refs/heads/main.zip). Hazlo con anticipación y vuelve a descargarla cuando haya actualizaciones.
2. Descomprime el .zip en una carpeta del computador.
3. Abre **index.html** con doble clic (Chrome o Edge). Arriba aparecerá *"Modo sin internet"*.

El historial de memos se guarda por separado en cada modo (en línea y sin internet). Para pasarlo de uno a otro usa **Historial → Descargar respaldo** y luego **Importar respaldo**.

## Excel de un mes nuevo (sin MIS APUNTES)

El Excel de cada mes puede traer varias hojas (por ejemplo, el mayor contable y la hoja de trabajo). La app revisa todas y usa la hoja de trabajo IPE; si hay más de una, se puede cambiar con el selector **Hoja del Excel**.

Si la hoja todavía no tiene la columna *MIS APUNTES*:

1. Presiona **Copiar MIS APUNTES desde el Excel del mes anterior** y elige el Excel del mes pasado. Los registros que ya existían toman su estado (coincidencia por RUT + documento + número + comprobante, o por documento si cambió el comprobante); los nuevos quedan *SIN ESTADO*.
2. Opcional: **Descargar hoja de trabajo con MIS APUNTES** para completar los estados nuevos en Excel y volver a cargarla.

El estado se lee de la columna *MIS APUNTES* (también acepta *APUNTES* o *ESTADO*; no importan mayúsculas ni espacios). Las columnas se reconocen por su nombre, así que el orden y el nombre de la hoja pueden cambiar de un mes a otro. La columna ACREEDOR/DEUDOR se reconoce aunque venga sin el encabezado *XX*.

## Desarrollo

Sitio estático (HTML + JavaScript), sin compilación. Librerías incluidas en `vendor/`: SheetJS 0.20.3, PizZip 3.1.7, docxtemplater 3.60.0, FileSaver 2.0.5.

```bash
# probar: basta abrir index.html; o con servidor local
npx http-server .
# si cambia templates/memo_template.docx, regenerar la plantilla incrustada
python scripts/embed_template.py
# pruebas
node --test tests/*.test.js
```

La plantilla `templates/memo_template.docx` es el MEMO institucional con marcadores. Para cambiar el formato, edítala en Word conservando los marcadores.

Marcadores disponibles: `{fecha}`, `{razon_social}`, `{rut}`, `{iniciales}`, párrafos `{#motivos}{#primero}MOTIVO :{/primero} {doc} {nro}: {justificacion}{/motivos}` y fila de tabla `{#items}{doc} {nro} {monto} {docto_aplicar} {monto_aplicar}{/items}` y párrafos `{#notas}… {doc} {nro} por {remanente}…{/notas}`.
