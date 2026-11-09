// Lectura de imágenes de correos con Claude, a través del backend en Apps Script (backend/Code.gs).
(function () {
  // URL pública del backend. No es un secreto: la clave de Anthropic vive en las Propiedades del script.
  const BACKEND_URL = 'https://script.google.com/macros/s/AKfycbzL6d-iHn-tQPL1BPup5nrdl2A9CHpZH0_d3cQNWl3JyjMZ0iSNt0D3vIlRsfrpUVJ6/exec';
  const MAX_LADO = 2000; // px; se reduce solo si es muy grande, para que el texto pequeño de las tablas siga legible

  async function prepararImagen(archivo) {
    if (!archivo.type.startsWith('image/')) throw new Error('Formato no permitido. Usa una imagen JPG o PNG.');
    const bitmap = await createImageBitmap(archivo);
    const escala = Math.min(1, MAX_LADO / Math.max(bitmap.width, bitmap.height));
    const lienzo = document.createElement('canvas');
    lienzo.width = Math.round(bitmap.width * escala);
    lienzo.height = Math.round(bitmap.height * escala);
    const ctx = lienzo.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, lienzo.width, lienzo.height);
    ctx.drawImage(bitmap, 0, 0, lienzo.width, lienzo.height);
    return lienzo.toDataURL('image/jpeg', 0.92);
  }

  /** accion: 'leer_documentos' o 'leer_banco'. Devuelve { datos, imagen } (imagen: data URL para mostrarla junto a lo leído). */
  async function leer(accion, archivo) {
    const imagen = await prepararImagen(archivo);
    let j;
    try {
      const r = await fetch(BACKEND_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({ accion, tipo: 'image/jpeg', datos: imagen.split(',')[1] }),
      });
      j = await r.json();
    } catch (err) {
      throw new Error('No se pudo conectar con el servicio de lectura. Revisa tu conexión a internet o que el servicio esté habilitado (backend/LEEME.md). Puedes completar los datos a mano.');
    }
    if (!j.ok) throw new Error(j.error || 'Error desconocido');
    return { datos: j.datos, imagen };
  }

  (window.IPE = window.IPE || {}).lectura = { leer };
})();
