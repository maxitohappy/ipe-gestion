# Gestión IPE

Aplicación web para la gestión mensual de **ingresos percibidos en exceso (IPE)** a partir del Excel de control.

**Privacidad:** el Excel se procesa solo en el navegador y no se envía a internet. Solo las **imágenes de correos que el usuario sube** salen del navegador: pasan por el backend en Google Apps Script hacia Claude (Anthropic) para leerlas, y el backend no las guarda. Nada de esto queda en este repositorio.

## Estructura del Excel

Hoja de trabajo con, entre otras, las columnas: `TIPO DE IPE` (B), `MI GESTION` (C), `C.F.`, `COD. AUX.`, `DV`, `RAZÓN SOCIAL`, `DOC.`, `NRO.`, `COMPROB.`, `FECHA`, `SALDO`, `XX` (ACREEDOR/DEUDOR), `Justificación Ingreso anticipado` y `Gestión realizada`.

- Las columnas se reconocen por su nombre: el orden y el nombre de la hoja pueden cambiar.
- `MI GESTION` indica el estado. También se acepta el nombre anterior `MIS APUNTES`. Si contiene *aplicar* o *aplicación* el registro va a **Memos**; si contiene *devolver* o *devolución* va a **Devoluciones** (si nombra ambas, queda pendiente). Si dice *finalizado*, no se considera en los totales ni se muestra.
- `TIPO DE IPE` debe coincidir con un tipo del documento GLOSAS: *Pago en exceso, Pago duplicado, Depósito sin identificar, Nota de Crédito, Error en pago, Regularización, Anulación RLE*. "Error en pago" tiene tres redacciones: `Error en pago` (o `Error en pago 1`), `Error en pago 2` y `Error en pago 3`. Si el tipo está vacío, el MOTIVO usa la columna *Justificación*.
- Si el libro trae varias hojas (por ejemplo, el mayor contable), la app usa la hoja de trabajo IPE; se puede cambiar con el selector **Hoja del Excel**.
- La columna ACREEDOR/DEUDOR se reconoce aunque venga sin el encabezado `XX`.
- Cuando un mismo IPE (empresa + documento + número) tiene saldo deudor y acreedor, se muestra **un solo registro** con el saldo acreedor menos el deudor.

## Pestañas

- **Resumen y alertas:** empresas en total (por razón social), en Memos, en Devoluciones y pendientes (las que no están en ninguna de las dos), y el listado de los ingresos anticipados de las **empresas pendientes**, numerado, con su MI GESTION y **categoría** (bandera): 🟥 años anteriores, 🟨 primer semestre del año en curso, 🟩 segundo semestre.
- **Memos (APLICAR):** un **anexo de aplicación** en Word por empresa (con buscador por RUT o empresa). El formato se elige solo según la fecha del IPE: **año en curso** o **años anteriores** (si una empresa tiene IPE de ambos, se generan dos anexos). El **MOTIVO** se redacta con la glosa del TIPO DE IPE de la planilla; los datos a completar aparecen como **[N° factura]**, **[fecha factura]**, etc., con fondo amarillo. Cada IPE puede aplicarse a varios documentos; si se aplica una parte (por ejemplo $20 de $100), lo aplicado pasa al historial y **el resto ($80) sigue pendiente**, con la Nota de diferencia. En el anexo siguiente del mismo IPE, el MOTIVO agrega *"Asimismo, con fecha …, se aplicaron $20, según MEMO N°[N° memo], de fecha …"*.
- **Devoluciones:** un **anexo de devolución** en Word por empresa, con el formato de **año en curso** o **años anteriores** según la fecha del IPE (con buscador por RUT o empresa). **OBSERVACIONES** usa la glosa del TIPO DE IPE; en la app se completan el tipo de cuenta, el N° de cuenta, el **banco** (lista de bancos de Chile u "Otro") y los correos. Lo que falte queda resaltado en el Word. Se puede devolver el total o una parte.
- **Historial de memos** e **Historial de devoluciones:** cada documento descargado queda registrado; se puede volver a descargar, devolver a pendientes, exportar a Excel y respaldar/importar.

### Lectura de imágenes de correos

Usa la misma herramienta de Lector de Boletas: Claude a través de un backend en Google Apps Script ([backend/LEEME.md](backend/LEEME.md)). Lo leído siempre queda editable, con la imagen a la vista y advertencias cuando algo no cuadra.

- **Memos:** con **Leer correo (imagen)** (arriba o en la tarjeta), la app identifica al cliente por RUT o razón social, recorre sus IPE incluidos y asigna los documentos en el orden del correo: *Pagar* si el saldo cubre el documento y *Abonar* si cubre una parte. Advierte si el saldo o el RUT del correo difieren del Excel.
- **Devoluciones:** con **Leer datos bancarios (imagen)** se completan tipo de cuenta, número, banco y correos (excluye los de la DGAC). Advierte si el beneficiario tiene otro RUT.
- **Datos bancarios guardados:** al descargar el Word quedan guardados por RUT en este navegador (se incluyen en el respaldo). En una devolución futura aparece *"Este cliente registra los siguientes datos bancarios… ¿Confirma?"*: **Sí** los pasa a la tarjeta y al Word; **No** pide subir la imagen.

### Saldo pendiente entre meses

Cada movimiento del historial guarda el saldo que tenía el IPE en el Excel. Al cargar un Excel nuevo, los movimientos aún no reflejados (mismo saldo) se descuentan; si el saldo del Excel ya cambió, se toma el saldo del Excel. Un IPE desaparece de Memos/Devoluciones cuando su saldo pendiente llega a cero.

## Uso mensual

1. Entra a la página publicada y arrastra el Excel del mes.
2. Si la hoja no tiene `MI GESTION`, usa **Copiar desde el Excel del mes anterior** (copia MI GESTION y TIPO DE IPE de los registros que ya existían). Opcional: **Descargar hoja de trabajo** para completar los nuevos en Excel y volver a cargarla.
3. Revisa **Resumen y alertas**.
4. En **Memos** y **Devoluciones**: revisa fecha e iniciales, elige el TIPO DE IPE, completa los datos y descarga los Word.
5. Completa en Word los datos resaltados en amarillo.
6. En los **historiales**: descarga un respaldo periódicamente (el historial se guarda en el navegador).

## Uso sin internet (respaldo)

1. Descarga la versión local desde **"Descargar versión sin internet"** (o https://github.com/maxitohappy/ipe-gestion/archive/refs/heads/main.zip).
2. Descomprime el .zip y abre **index.html** con doble clic (Chrome o Edge). Arriba aparecerá *"Modo sin internet"*. La lectura de imágenes sí necesita internet; sin conexión, los datos se completan a mano.

El historial y los datos bancarios se guardan por separado en cada modo; para pasarlo de uno a otro usa **Descargar respaldo** e **Importar respaldo**.

## Desarrollo

Sitio estático (HTML + JavaScript), sin compilación. Librerías incluidas en `vendor/`: SheetJS 0.20.3, PizZip 3.1.7, docxtemplater 3.60.0, FileSaver 2.0.5.

```bash
# probar: basta abrir index.html; o con servidor local
npx http-server .
# si cambia alguna plantilla de templates/, regenerar las plantillas incrustadas
python scripts/embed_template.py
# pruebas
node --test tests/*.test.js
```

Plantillas (`templates/`), editables en Word conservando los marcadores:

Los textos con datos a completar se pasan como listas de partes `{#campo}{#n}{t}{/n}{#m}{t}{/m}{/campo}` (`m` = resaltado).

- `aplicacion_actual.docx` / `aplicacion_anterior.docx` (anexo de aplicación, año en curso / años anteriores): `{razon_social}`, `{rut}`, párrafos `{#motivos}{#primero}MOTIVO:{/primero} {doc} {nro}: {justificacion…}{/motivos}`, fila `{#items}{doc} {nro} {monto} {docto_aplicar} {monto_aplicar}{/items}` y párrafos `{#notas}… {doc} {nro} por {remanente}…{/notas}`.
- `devolucion_actual.docx` / `devolucion_anterior.docx` (anexo de devolución): `{rut}`, `{razon_social}`, fila `{#items}{cf} {tipo_docto} {monto} {observaciones…}{/items}`, `{tipo_cuenta…}`, `{numero_cuenta…}`, `{banco…}`, `{monto_total}` y correos (`{#correos}{correo}{/correos}` en año en curso; `{correos_linea}` en años anteriores).

Las glosas por tipo de IPE están en `js/glosas.js`.
