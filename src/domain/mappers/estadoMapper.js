/**
 * Mapea los datos crudos del libro (raw = lo que devuelve
 * excelService.leerLibro(LECTURAS).datos) al contrato de dominio
 * EstadoModelo (ver docs/ARQUITECTURA_MVP.md §4). Es la misma lógica que hoy
 * vive en app/render.js:pintaEstado(), separada de la generación de HTML
 * para poder probarla sin DOM ni Excel.
 */
import { num } from "../../utils/format.js";
import { mapProgramas } from "./programasMapper.js";

function texto(v) {
  return String(v || "").trim();
}

/**
 * @param {Object} raw
 * @returns {Object} EstadoModelo
 */
export function mapEstado(raw) {
  var ct = raw.ct005 ? raw.ct005.t.map(function (r) { return String(r[0] || ""); }) : [];
  var ctOK = ct.length === 3 && ct.every(function (x) { return x.indexOf("OK") >= 0; });
  var motorSin = ct.length > 2 && String(ct[2]).indexOf("SIN EJECUTAR") >= 0;
  var integridad = ctOK ? "ok" : (motorSin ? "motor_sin_ejecutar" : "revisar");

  var waccReal = raw.parE ? num(raw.parE.v[0][0]) : null;
  var waccCorriente = raw.par ? num(raw.par.v[4][0]) : null; // C8, 5.ª celda de C4:C18
  var ipcValor = raw.par ? num(raw.par.v[5][0]) : null;      // C9, 6.ª celda de C4:C18

  var desfase = false;
  if (raw.flagsV && raw.flags2) {
    var a = raw.flagsV.t[0].map(function (x) { return texto(x).toUpperCase(); });
    var b = raw.flags2.t[0].map(function (x) { return texto(x).toUpperCase(); });
    for (var i = 0; i < Math.max(a.length, b.length); i++) {
      if ((a[i] || "") !== (b[i] || "")) { desfase = true; break; }
    }
  }

  var anios = raw.anios ? raw.anios.t[0].filter(function (x) { return texto(x); }) : [];
  var horizonte = anios.length
    ? { inicio: texto(anios[0]), fin: texto(anios[anios.length - 1]), cantidadAnios: anios.length }
    : null;

  var programasSinDatos = mapProgramas(raw).filter(function (p) { return p.estado === "sin_datos"; });

  return {
    empresa: raw.empresa ? texto(raw.empresa.t[0][0]) : "",
    horizonte: horizonte,
    metodoProyeccionUsuarios: raw.metodo ? texto(raw.metodo.t[0][0]) : "",
    metaIPUF: raw.meta ? num(raw.meta.v[0][0]) : null,
    integridad: integridad,
    controlesIntegridad: ct,
    wacc: { real: waccReal, corriente: waccCorriente },
    alternativaDesfasada: desfase,
    programasSinDatos: programasSinDatos,
    ipc: { valor: ipcValor, desactualizado: ipcValor !== null && ipcValor < 0.045 },
    configuracionMotor: {
      presupuesto: raw.monto ? texto(raw.monto.t[0][0]) : "",
      bcMinimo: raw.minbc ? texto(raw.minbc.t[0][0]) : "",
      combinacionA1: raw.selA1 ? texto(raw.selA1.t[0][0]) : ""
    }
  };
}
