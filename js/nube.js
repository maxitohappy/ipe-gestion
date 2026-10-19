// Conexión con Firebase: inicio de sesión (Authentication) e historiales compartidos (Firestore).
//
// - Solo entran usuarios creados por el administrador en la consola de Firebase (el registro público está
//   desactivado). Al primer ingreso se exige cambiar la contraseña temporal.
// - Los memos y devoluciones generados quedan en las colecciones "memos" y "devoluciones", visibles para todos
//   los usuarios; las reglas de Firestore (firestore.rules) solo permiten borrar un documento a quien lo creó.
(function (root) {
  const COLECCIONES = { memos: 'memos', devoluciones: 'devoluciones' };
  let auth = null;
  let db = null;

  /** Disponible solo si hay configuración y la página se abrió desde internet (no en la versión sin internet). */
  function disponible() {
    return Boolean(root.IPE_FIREBASE_CONFIG && root.firebase && location.protocol !== 'file:');
  }

  function iniciar() {
    if (auth) return;
    root.firebase.initializeApp(root.IPE_FIREBASE_CONFIG);
    auth = root.firebase.auth();
    auth.languageCode = 'es';
    db = root.firebase.firestore();
  }

  const usuario = () => auth && auth.currentUser;

  function alCambiarSesion(cb) {
    return auth.onAuthStateChanged(cb);
  }

  function ingresar(correo, clave) {
    return auth.signInWithEmailAndPassword(correo.trim(), clave);
  }

  function salir() {
    return auth.signOut();
  }

  /** Envía el correo de Firebase para crear una nueva contraseña. */
  function recuperarClave(correo) {
    return auth.sendPasswordResetEmail(correo.trim(), { url: location.origin + location.pathname });
  }

  /** ¿El usuario aún tiene la contraseña temporal que le dio el administrador? */
  async function debeCambiarClave() {
    const doc = await db.collection('usuarios').doc(usuario().uid).get();
    return !doc.exists || !doc.data().claveDefinida;
  }

  /** Cambia la contraseña (pide la actual para confirmar identidad) y la marca como definitiva. */
  async function cambiarClave(actual, nueva) {
    const u = usuario();
    const cred = root.firebase.auth.EmailAuthProvider.credential(u.email, actual);
    await u.reauthenticateWithCredential(cred);
    await u.updatePassword(nueva);
    await db.collection('usuarios').doc(u.uid).set({
      correo: u.email,
      claveDefinida: true,
      actualizado: root.firebase.firestore.FieldValue.serverTimestamp(),
    }, { merge: true });
  }

  /** Escucha un historial en tiempo real; `cb` recibe la lista de registros (cada uno con su autor). */
  function escuchar(tipo, cb, onError) {
    return db.collection(COLECCIONES[tipo]).onSnapshot(
      (snap) => cb(snap.docs.map((d) => ({ ...d.data(), id: d.id }))),
      onError,
    );
  }

  /** Guarda registros del historial a nombre del usuario actual. */
  async function guardar(tipo, entradas) {
    const u = usuario();
    const lote = db.batch();
    for (const e of entradas) {
      // Firestore no admite valores undefined: se limpia el registro antes de guardarlo.
      const limpio = JSON.parse(JSON.stringify(e));
      lote.set(db.collection(COLECCIONES[tipo]).doc(e.id), { ...limpio, creadoPor: u.uid, creadoPorCorreo: u.email });
    }
    await lote.commit();
  }

  /** Elimina un registro (las reglas de Firestore solo lo permiten a quien lo creó). */
  function eliminar(tipo, id) {
    return db.collection(COLECCIONES[tipo]).doc(id).delete();
  }

  /** Mensajes de error de Firebase en español. */
  function mensaje(err) {
    const c = (err && err.code) || '';
    const m = {
      'auth/invalid-credential': 'Correo o contraseña incorrectos.',
      'auth/wrong-password': 'La contraseña actual no es correcta.',
      'auth/user-not-found': 'Correo o contraseña incorrectos.',
      'auth/invalid-email': 'El correo no es válido.',
      'auth/missing-password': 'Ingresa la contraseña.',
      'auth/too-many-requests': 'Demasiados intentos. Espera unos minutos o recupera tu contraseña.',
      'auth/user-disabled': 'Este usuario está deshabilitado. Contacta al administrador.',
      'auth/weak-password': 'La nueva contraseña es muy débil (mínimo 8 caracteres).',
      'auth/network-request-failed': 'Sin conexión a internet.',
      'permission-denied': 'No tienes permiso para esta acción (solo quien generó el documento puede volverlo a pendientes).',
    };
    return m[c] || (err && err.message) || 'Ocurrió un error.';
  }

  (root.IPE = root.IPE || {}).nube = {
    disponible, iniciar, usuario, alCambiarSesion, ingresar, salir, recuperarClave,
    debeCambiarClave, cambiarClave, escuchar, guardar, eliminar, mensaje,
  };
})(typeof self !== 'undefined' ? self : this);
