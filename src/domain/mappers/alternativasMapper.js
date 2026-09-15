/**
 * Mapea los datos crudos del libro al contrato de dominio Alternativa[] (ver
 * docs/ARQUITECTURA_MVP.md §4). Es la misma lógica que hoy vive en
 * app/render.js:pintaAlt(), separada de la generación de HTML.
 */
import { num } from "../../utils/format.js";

function bloqueBC(rango) {
  var beneficio = rango ? num(rango.v[0][0]) : null;
  var costo = rango ? num(rango.v[1][0]) : null;
  var bc = (beneficio !== null && costo !== null && costo !== 0) ? beneficio / costo : null;
  return { beneficio: beneficio, costo: costo, bc: bc };
}

function promedioIPUF(rango) {
  if (!rango) return null;
  var valores = rango.v[0].map(num).filter(function (x) { return x !== null; });
  if (valores.length < 4) return null;
  var tramo = valores.slice(3);
  return tramo.reduce(function (a, b) { return a + b; }, 0) / tramo.length;
}

function cumpleMeta(rango) {
  if (!rango) return null;
  var texto = String(rango.t[0][0] || "");
  if (texto.indexOf("NO CUMPLE") >= 0) return false;
  if (texto.indexOf("CUMPLE") >= 0) return true;
  return null;
}

function comoTexto(rango) {
  return rango ? String(rango.t[0][0] || "") : "";
}

/**
 * @param {Object} raw
 * @returns {Array<Object>} Alternativa[]
 */
export function mapAlternativas(raw) {
  var b1 = bloqueBC(raw.a1bc);
  var b2 = bloqueBC(raw.a2bc);
  var b3 = bloqueBC(raw.a3bc);

  return [
    {
      id: 1,
      beneficio: b1.beneficio,
      costo: b1.costo,
      bc: b1.bc,
      ipufPromedio: promedioIPUF(raw.ipuf1),
      cumpleMeta: cumpleMeta(raw.ver1),
      composicionTexto: comoTexto(raw.selA1)
    },
    {
      id: 2,
      beneficio: b2.beneficio,
      costo: b2.costo,
      bc: b2.bc,
      ipufPromedio: promedioIPUF(raw.ipuf2),
      cumpleMeta: cumpleMeta(raw.ver2),
      composicionTexto: comoTexto(raw.a2txt)
    },
    {
      id: 3,
      beneficio: b3.beneficio,
      costo: b3.costo,
      bc: b3.bc,
      // cellMap.js (LECTURAS) no trae una serie de IPUF para la Alternativa 3
      // -- no se inventa una fuente; se deja en null, igual que hoy en
      // app/render.js:pintaAlt() (columna "—" fija para Alt. 3).
      ipufPromedio: null,
      cumpleMeta: cumpleMeta(raw.ver3),
      composicionTexto: comoTexto(raw.a3txt)
    }
  ];
}
