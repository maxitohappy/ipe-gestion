// Utilidades de formato: RUT, montos en pesos, fechas.
(function (root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else (root.IPE = root.IPE || {}).format = api;
})(typeof self !== 'undefined' ? self : this, function () {
  const MESES = ['ENE', 'FEB', 'MAR', 'ABR', 'MAY', 'JUN', 'JUL', 'AGO', 'SEP', 'OCT', 'NOV', 'DIC'];

  /** Normaliza un encabezado: sin tildes, mayúsculas y solo letras/números ("COD. AUX." -> "CODAUX"). */
  function normalizeHeader(s) {
    return String(s ?? '')
      .normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toUpperCase().replace(/[^A-Z0-9]/g, '');
  }

  function miles(n) {
    const abs = Math.round(Math.abs(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.');
    return (n < 0 ? '-' : '') + abs;
  }

  /** 96755500, "2" -> "96.755.500-2" */
  function formatRut(num, dv) {
    const digits = String(num ?? '').replace(/\D/g, '');
    if (!digits) return '';
    return `${miles(Number(digits))}-${String(dv ?? '').trim().toUpperCase()}`;
  }

  /** 22955 -> "$ 22.955" */
  function formatCLP(n) {
    return `$ ${miles(Number(n) || 0)}`;
  }

  /** Convierte "$ 7.926", " $ - ", 7926, "-24.570" a número. */
  function toNumber(v) {
    if (typeof v === 'number') return v;
    const s = String(v ?? '').trim();
    if (!s || /^\$?\s*-\s*$/.test(s)) return 0;
    const neg = /^-|\(.*\)|\$\s*-\s*\d/.test(s);
    const n = Number(s.replace(/[^\d,]/g, '').replace(',', '.'));
    return Number.isFinite(n) ? (neg ? -n : n) : 0;
  }

  /** Monto a aplicar: si es número lo deja con separador de miles ("701.483"); si es texto, lo respeta. */
  function formatMontoLibre(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') return miles(v);
    const s = String(v).trim();
    return /^[\d.\s$]+$/.test(s) && /\d/.test(s) ? miles(toNumber(s)) : s;
  }

  /** Acepta Date, número de serie Excel o texto dd-mm-yyyy / dd.mm.yyyy / dd/mm/yyyy. */
  function parseFecha(v) {
    if (v instanceof Date) return isNaN(v) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate());
    if (typeof v === 'number' && v > 20000 && v < 80000) {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    const m = String(v ?? '').match(/(\d{1,2})[-./](\d{1,2})[-./](\d{4})/);
    if (!m) return null;
    const d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    return isNaN(d) ? null : d;
  }

  /** Date -> "14.OCT.2026" (formato del memo) */
  function fechaMemo(d) {
    return `${String(d.getDate()).padStart(2, '0')}.${MESES[d.getMonth()]}.${d.getFullYear()}`;
  }

  /** Date -> "OCT.2026" (fecha del memo y de la devolución: solo mes y año; el día lo completa el usuario) */
  function fechaMesAnio(d) {
    return `${MESES[d.getMonth()]}.${d.getFullYear()}`;
  }

  /**
   * Bandera según la fecha del ingreso anticipado, respecto del año en curso:
   * roja = años anteriores, amarilla = primer semestre, verde = segundo semestre.
   */
  function bandera(fecha, hoy = new Date()) {
    if (!fecha) return null;
    if (fecha.getFullYear() < hoy.getFullYear()) return 'roja';
    if (fecha.getFullYear() > hoy.getFullYear()) return 'verde';
    return fecha.getMonth() < 6 ? 'amarilla' : 'verde';
  }

  /** Date -> "14-10-2026" */
  function fechaCorta(d) {
    if (!d) return '';
    return `${String(d.getDate()).padStart(2, '0')}-${String(d.getMonth() + 1).padStart(2, '0')}-${d.getFullYear()}`;
  }

  function diasEntre(desde, hasta) {
    return Math.floor((hasta - desde) / 86400000);
  }

  return { normalizeHeader, formatRut, formatCLP, toNumber, formatMontoLibre, parseFecha, fechaMemo, fechaMesAnio, bandera, fechaCorta, diasEntre, miles };
});
