# Gestión IPE

Aplicación web para la gestión mensual de **ingresos percibidos en exceso (IPE)** a partir del Excel de control.

**Privacidad:** el Excel se procesa solo en el navegador. Ningún dato se envía a internet ni se guarda en este repositorio.

## Estructura del Excel

Hoja de trabajo con, entre otras, las columnas: `TIPO DE IPE` (B), `MI GESTION` (C), `C.F.`, `COD. AUX.`, `DV`, `RAZÓN SOCIAL`, `DOC.`, `NRO.`, `COMPROB.`, `FECHA`, `SALDO`, `XX` (ACREEDOR/DEUDOR), `Justificación Ingreso anticipado` y `Gestión realizada`.

- Las columnas se reconocen por su nombre: el orden y el nombre de la hoja pueden cambiar.
- `MI GESTION` indica el estado (`APLICAR`, `DEVOLVER`, `correo enviado …`). También se acepta el nombre anterior `MIS APUNTES`.
- `TIPO DE IPE` debe coincidir con un tipo del documento GLOSAS: *Pago en exceso, Pago duplicado, Depósito sin identificar, Nota de Crédito, Error en pago, Regularización, Anulación RLE*. "Error en pago" tiene tres redacciones: `Error en pago` (o `Error en pago 1`), `Error en pago 2` y `Error en pago 3`. Si el tipo está vacío, el MOTIVO usa la columna *Justificación*.
- Si el libro trae varias hojas (por ejemplo, el mayor contable), la app usa la hoja de trabajo IPE; se puede cambiar con el selector **Hoja del Excel**.
- La columna ACREEDOR/DEUDOR se reconoce aunque venga sin el encabezado `XX`.
- Cuando un mismo IPE (empresa + documento + número) tiene saldo deudor y acreedor, se muestra **un solo registro** con el saldo acreedor menos el deudor.

## Pestañas

- **Resumen y alertas:** empresas en total (por razón social), en Memos, en Devoluciones y pendientes (las que no están en ninguna de las dos), y el listado de los ingresos anticipados de las **empresas pendientes**, numerado, con su MI GESTION y **categoría** (bandera): 🟥 años anteriores, 🟨 primer semestre del año en curso, 🟩 segundo semestre.
- **Memos (APLICAR):** un memo Word por empresa (con buscador por RUT o empresa). El **MOTIVO** se redacta en el Word con la glosa del TIPO DE IPE de la planilla; los datos a completar aparecen como **[N° factura]**, **[fecha factura]**, etc., con fondo amarillo. Cada IPE puede aplicarse a varios documentos; si se aplica una parte (por ejemplo $20 de $100), lo aplicado pasa al historial y **el resto ($80) sigue pendiente**. En el memo siguiente del mismo IPE, el MOTIVO agrega *"Asimismo, con fecha …, se aplicaron $20, según MEMO N°[N° memo], de fecha …"* (el N° se completa en el Word). La fecha del Word va como mes.año (`SANTIAGO, OCT.2026`) y el N° de memo queda en blanco.
- **Devoluciones:** un oficio Word por empresa con el formato de devolución (con buscador por RUT o empresa). **OBSERVACIONES** usa la glosa del TIPO DE IPE; en la app se completan la REF., el tipo de cuenta, el N° de cuenta, el **banco** (lista de bancos de Chile u "Otro") y los correos. Lo que falte queda resaltado en el Word. Se puede devolver el total o una parte.
- **Historial de memos** e **Historial de devoluciones:** cada documento descargado queda registrado; se puede volver a descargar, devolver a pendientes, exportar a Excel y respaldar/importar.

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
2. Descomprime el .zip y abre **index.html** con doble clic (Chrome o Edge). Arriba aparecerá *"Modo sin internet"*.

El historial se guarda por separado en cada modo; para pasarlo de uno a otro usa **Descargar respaldo** e **Importar respaldo**.

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

- `memo_template.docx`: `{fecha}`, `{razon_social}`, `{rut}`, `{iniciales}`, párrafos `{#motivos}{#primero}MOTIVO :{/primero} {doc} {nro}: {justificacion…}{/motivos}`, fila `{#items}{doc} {nro} {monto} {docto_aplicar} {monto_aplicar}{/items}` y párrafos `{#notas}… {doc} {nro} por {remanente}…{/notas}`.
- `devolucion_template.docx`: `{referencia…}`, `{fecha}`, `{rut}`, `{razon_social}`, fila `{#items}{cf} {tipo_docto} {monto} {observaciones…}{/items}`, `{tipo_cuenta…}`, `{numero_cuenta…}`, `{banco…}`, `{monto_total}`, párrafos `{#correos}{correo}{/correos}` e `{iniciales}`.

Las glosas por tipo de IPE están en `js/glosas.js`.
