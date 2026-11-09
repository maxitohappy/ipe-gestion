# Backend de lectura de imágenes (Google Apps Script)

Misma herramienta de Lector de Boletas: la imagen llega al script, este la envía a Claude y devuelve los datos. No guarda imágenes ni datos.

## Primera configuración (una vez)

1. Abre el proyecto en <https://script.google.com> (proyecto **Gestion IPE - Lector de correos**).
2. ⚙️ **Configuración del proyecto → Propiedades del script → Agregar propiedad**:
   - `ANTHROPIC_API_KEY` = tu clave de <https://console.anthropic.com/settings/keys>
   - Opcionales: `LIMITE_DIARIO` (por defecto 100), `MODELO` (por defecto `claude-haiku-5-5`).
3. En el editor, elige la función **`inicializar`** y pulsa **Ejecutar**. Autoriza los permisos (Google avisa que la app no está verificada porque es tuya: *Configuración avanzada → Ir a Gestion IPE*).

Para probar: abre `URL_DEL_BACKEND?accion=ping`. Debe responder `{"ok":true,...}`.

## Actualizar el código sin cambiar la URL

Requiere `clasp login` y el `.clasp.json` local (no se publica):

```bash
cd backend
clasp push -f
clasp update-deployment <ID de la implementación> -d "descripción"
```

## Acciones

- GET `ping`: estado y uso del día.
- POST `leer_documentos` (`tipo`, `datos` en base64): correo con saldo a favor y documentos por pagar.
- POST `leer_banco` (`tipo`, `datos` en base64): correo con datos bancarios.

Cada lectura descuenta 1 del tope diario, salvo las que la API rechaza.

## Notas

- La URL `/exec` es pública: quien la tenga puede enviar imágenes hasta el tope diario, lo que consume saldo de la cuenta de Anthropic. No da acceso a datos, porque el script no guarda nada.
- Las imágenes pasan por Google y Anthropic: pueden contener datos bancarios y de clientes.
