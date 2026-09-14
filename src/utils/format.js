/**
 * Utilidades puras de formato y parseo. Sin dependencias de Office.js ni del DOM,
 * para que se puedan probar con Node normal (ver /test).
 */

export function esc(s) {
  return String(s == null ? "" : s).replace(/[&<>"]/g, function (c) {
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;" }[c];
  });
}

export function num(v) {
  return (typeof v === "number" && isFinite(v)) ? v : null;
}

export function fmt(v, d) {
  var n = num(v);
  if (n === null) return "—";
  return n.toLocaleString("es-CO", { minimumFractionDigits: d || 0, maximumFractionDigits: d || 0 });
}

export function fmtMM(v) {
  var n = num(v);
  if (n === null) return "—";
  return fmt(n / 1e6, 0);
}

export function parseNum(s) {
  if (s === null || s === undefined) return null;
  s = String(s).trim();
  if (!s) return null;
  s = s.replace(/\./g, "").replace(/,/g, ".").replace(/[^0-9.\-]/g, "");
  var n = parseFloat(s);
  return isFinite(n) ? n : null;
}

/**
 * true si alguna celda de la matriz de fórmulas (Range.formulas) contiene una fórmula.
 * Es la base de la protección "el panel nunca escribe sobre una celda con fórmula".
 */
export function hasFormula(formulas) {
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      var f = formulas[i][j];
      if (typeof f === "string" && f.charAt(0) === "=") return true;
    }
  }
  return false;
}
