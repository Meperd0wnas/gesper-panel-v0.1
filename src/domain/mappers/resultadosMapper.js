/**
 * Mapea los datos crudos del libro a los contratos de dominio
 * ResultadosHidraulicos y ResultadosFinancieros (ver docs/ARQUITECTURA_MVP.md
 * §4). Es la misma lógica que hoy vive en app/render.js:pintaResultados() y
 * app/render.js:sparkIPUF(), separada de la generación de HTML/SVG.
 */
import { num } from "../../utils/format.js";

/**
 * Replica a propósito una particularidad del cálculo que ya existía en
 * app/render.js:pintaResultados() antes de la Fase 2D: si hay AL MENOS un
 * año válido en la serie pero son menos de `desde` (p. ej. 2 años válidos
 * con desde=3), el promedio da 0 -- no null -- porque el código original
 * dividía por `Math.max(1, tramo.length)` para evitar una división por
 * cero. Solo se devuelve null cuando la serie completa no tiene ningún año
 * válido. No es el comportamiento "más correcto" (sería razonable esperar
 * null también en ese caso), pero esta fase exige paridad exacta con la
 * pantalla anterior, no una corrección -- ver informe de la Fase 2D.
 */
function promedioDesde(serie, desde) {
  if (!serie.length) return null;
  var tramo = serie.slice(desde);
  return tramo.reduce(function (a, b) { return a + b; }, 0) / Math.max(1, tramo.length);
}

/**
 * @param {Object} raw
 * @returns {Object} ResultadosHidraulicos
 */
export function mapResultadosHidraulicos(raw) {
  var serie = raw.ipuf2 ? raw.ipuf2.v[0].map(num) : [];
  var validos = serie.filter(function (x) { return x !== null; });
  // El promedio del horizonte descarta los primeros 3 años, igual que hoy en
  // app/render.js:pintaResultados() (Alternativas!M62:V62 en la nota de UI).
  var promedio = promedioDesde(validos, 3);
  var meta = raw.meta ? num(raw.meta.v[0][0]) : null;

  var ianc = raw.ianc2 ? raw.ianc2.v[0].map(num).filter(function (x) { return x !== null; }) : [];
  var anios = raw.anios ? raw.anios.t[0] : [];

  var serieIPUF = serie.map(function (valor, i) {
    return { anio: anios[i] !== undefined ? String(anios[i]) : "", valor: valor };
  });

  return {
    ipufPromedio: promedio,
    metaIPUF: meta,
    cumpleMeta: (promedio !== null && meta !== null) ? promedio < meta : null,
    ianc: { inicio: ianc.length ? ianc[0] : null, fin: ianc.length ? ianc[ianc.length - 1] : null },
    serieIPUF: serieIPUF
  };
}

/**
 * @param {Object} raw
 * @returns {Object} ResultadosFinancieros
 */
export function mapResultadosFinancieros(raw) {
  var f = raw.fcaja ? raw.fcaja.t : null; // C65 payback, C66 TIR, C67 VNA
  var tir = f ? String(f[1][0]) : null;
  var tirmCruda = raw.tirm ? raw.tirm.v[0][0] : null;
  var tirm = num(tirmCruda);

  var alc = raw.par ? num(raw.par.v[13][0]) : null;  // C17
  var alc2 = raw.par ? num(raw.par.v[14][0]) : null; // C18

  return {
    // true si se encontró la hoja FCAJA PROYECTO en el libro -- antes de la
    // Fase 2D, render.js mostraba un mensaje distinto ("No encuentro la
    // hoja...") en vez de las tarjetas de VNA/TIR/payback cuando esta hoja
    // faltaba (p. ej. una plantilla de otra versión). Sin este campo, la UI
    // no podría distinguir "hoja ausente" de "hoja presente pero vacía".
    flujoCajaEncontrado: f !== null,
    vna: f ? String(f[2][0]) : null,
    tir: tir,
    tirDefinida: tir !== null && tir.indexOf("N.A") < 0,
    tirModificada: tirm,
    // Texto crudo de FCJ!D66 para cuando no es un número (p. ej. vacía o con
    // una nota) -- paridad con el `esc(tirm)` que usaba render.js cuando
    // `num(tirm)` no daba un número.
    tirModificadaTexto: tirm === null && tirmCruda !== null ? String(tirmCruda) : null,
    payback: f ? String(f[0][0]) : null,
    parametros: {
      waccCorriente: raw.par ? String(raw.par.t[4][0]) : "",
      cargoFijoAcueducto: raw.par ? String(raw.par.t[11][0]) : "",
      cargoVariableAcueducto: raw.par ? String(raw.par.t[12][0]) : "",
      gradiente: raw.grad ? String(raw.grad.t[0][0]) : null
    },
    alcantarilladoEnCero: (alc === 0 || alc === null) && (alc2 === 0 || alc2 === null)
  };
}
