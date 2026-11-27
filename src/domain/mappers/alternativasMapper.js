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

/**
 * Promedio del IPUF desde el 4.º año (columna M), el mismo tramo que usa la
 * hoja para decidir si una alternativa cumple la meta (T22: AVERAGE(M20:V20)).
 */
function promedioDesdeCuarto(fila) {
  var valores = fila.map(num).filter(function (x) { return x !== null; });
  if (valores.length < 4) return null;
  var tramo = valores.slice(3);
  return tramo.reduce(function (a, b) { return a + b; }, 0) / tramo.length;
}

function promedioIPUF(rango) {
  return rango ? promedioDesdeCuarto(rango.v[0]) : null;
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

/* ---------- Detalle por alternativa y plan de gestión (informe de resultados) ---------- */

/** Rótulo comparable: sin tildes, en mayúsculas, sin la unidad entre corchetes ("IANC [%]" -> "IANC"). */
export function normalizarRotulo(s) {
  return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "")
    .toUpperCase().replace(/\[.*?\]/g, "").replace(/\s+/g, " ").trim();
}

/**
 * Filas de los bloques H:V de cada alternativa (alt1bal, alt2bal, alt3bal),
 * ubicadas por el rótulo de la columna H. "PRODUCCIÓN" y "PRODUCCIÓN AL
 * APLICAR EL PROGRAMA" se distinguen porque la comparación es exacta.
 */
export var FILAS_BALANCE = [
  ["produccion", "PRODUCCION"],
  ["facturacion", "FACTURACION"],
  ["usuarios", "PROYECCION DE USUARIOS"],
  ["vrecProduccion", "V RECUPERADO DE PRODUCCION"],
  ["vrecFacturacion", "V RECUPERADO DE FACTURACION"],
  ["produccionConProgramas", "PRODUCCION AL APLICAR EL PROGRAMA"],
  ["facturacionConProgramas", "FACTURACION AL APLICAR EL PROGRAMA"],
  ["ianc", "IANC"],
  ["isuf", "ISUF"],
  ["icuf", "ICUF"],
  ["ipuf", "IPUF"]
];

function seriesDeBloque(rango) {
  var series = {};
  FILAS_BALANCE.forEach(function (f) { series[f[0]] = null; });
  if (!rango) return series;
  for (var i = 0; i < rango.v.length; i++) {
    var rotulo = normalizarRotulo(rango.t ? rango.t[i][0] : rango.v[i][0]);
    for (var k = 0; k < FILAS_BALANCE.length; k++) {
      if (FILAS_BALANCE[k][1] === rotulo && series[FILAS_BALANCE[k][0]] === null) {
        // columnas: H rótulo, I unidad, J:V un valor por año
        series[FILAS_BALANCE[k][0]] = rango.v[i].slice(2).map(num);
      }
    }
  }
  return series;
}

/**
 * Programas con "SI" en la fila de banderas de una alternativa. El número de
 * cada columna sale de la fila 6 (I6:AA6); si no se pudo leer, se asume el
 * orden de la hoja (columna I = P.1 ... AA = P.19).
 */
function programasMarcados(banderas, numeros) {
  if (!banderas) return [];
  var nums = numeros ? numeros.v[0] : [];
  var out = [];
  banderas.t[0].forEach(function (b, i) {
    if (String(b || "").trim().toUpperCase() !== "SI") return;
    var n = num(nums[i]);
    out.push(n !== null ? n : i + 1);
  });
  return out;
}

/** "1, 2, 3, 9, 10, " -> [1, 2, 3, 9, 10] */
export function listaDeProgramas(texto) {
  return String(texto || "").split(/[^0-9]+/).filter(Boolean).map(function (x) { return parseInt(x, 10); });
}

/**
 * @typedef {Object} DetalleAlternativa
 * @property {1|2|3} id
 * @property {number[]} programas     números de los programas marcados con "SI"
 * @property {string} veredicto       texto de T22/T63/T105 ("... CUMPLE con las metas ...")
 * @property {number|null} ipufPromedio  promedio del 4.º al último año, el tramo de T22/T63/T105
 * @property {string} nota            nota de la hoja sobre la alternativa (J45 para la 2)
 * @property {Object<string,(number|null)[]|null>} series  una por clave de FILAS_BALANCE, un valor por año
 */

/**
 * @param {Object} raw
 * @returns {{anios: string[], inicioPromedio: number, alternativas: DetalleAlternativa[]}}
 */
export function mapDetalleAlternativas(raw) {
  var anios = raw.anios ? raw.anios.t[0].map(function (x) { return String(x == null ? "" : x).trim(); }) : [];
  var fuentes = [
    { id: 1, bloque: raw.alt1bal, banderas: raw.flagsV, ver: raw.ver1, nota: null },
    { id: 2, bloque: raw.alt2bal, banderas: raw.flags2, ver: raw.ver2, nota: raw.alt2nota },
    { id: 3, bloque: raw.alt3bal, banderas: raw.flags3, ver: raw.ver3, nota: null }
  ];
  return {
    anios: anios,
    inicioPromedio: 3, // índice del primer año del promedio (columna M)
    alternativas: fuentes.map(function (f) {
      var series = seriesDeBloque(f.bloque);
      return {
        id: f.id,
        programas: programasMarcados(f.banderas, raw.altNums),
        veredicto: comoTexto(f.ver).trim(),
        ipufPromedio: series.ipuf ? promedioDesdeCuarto(series.ipuf) : null,
        nota: comoTexto(f.nota).replace(/^\*+/, "").trim(),
        series: series
      };
    })
  };
}

function celdaTexto(rango, i, j) {
  if (!rango || !rango.t[i] || rango.t[i][j] == null) return "";
  return String(rango.t[i][j]).trim();
}

function celdaNumero(rango, i, j) {
  return rango && rango.v[i] ? num(rango.v[i][j]) : null;
}

/**
 * Plan de gestión de pérdidas (Alternativas!A26:E35), combinaciones óptimas
 * (A37:E51) y clasificación de programas por criterio (A55:D65). Las filas
 * del resumen del plan se ubican por su rótulo de la columna B.
 * @param {Object} raw
 */
export function mapPlanGestion(raw) {
  var p = raw.plan;
  var plan = {
    descripcion: celdaTexto(p, 0, 0),
    requisitoEsencial: [],
    programasTexto: "",
    programas: [],
    costo: null,
    beneficio: null,
    bc: null,
    volumenRecuperado: null
  };
  if (p) {
    for (var i = 0; i < p.v.length; i++) {
      var a = normalizarRotulo(celdaTexto(p, i, 0));
      var b = normalizarRotulo(celdaTexto(p, i, 1));
      if (a.indexOf("REQUISITO") >= 0) {
        for (var j = 1; j < p.v[i].length; j++) { var r = celdaTexto(p, i, j); if (r) plan.requisitoEsencial.push(r); }
      } else if (a.indexOf("TOTAL DE PROGRAMAS") === 0) {
        plan.programasTexto = celdaTexto(p, i, 1).replace(/[,\s]+$/, "");
        plan.programas = listaDeProgramas(plan.programasTexto);
      } else if (b === "COSTO") plan.costo = celdaNumero(p, i, 2);
      else if (b === "BENEFICIO") plan.beneficio = celdaNumero(p, i, 2);
      else if (b.indexOf("RELACION") === 0) plan.bc = celdaNumero(p, i, 2);
      else if (b.indexOf("VOL") === 0) plan.volumenRecuperado = celdaNumero(p, i, 2);
    }
  }

  var combinaciones = [];
  var c = raw.combos;
  if (c) {
    for (var k = 0; k < c.v.length; k++) {
      var progs = celdaTexto(c, k, 0).replace(/[,\s]+$/, "");
      if (!progs) continue;
      combinaciones.push({
        programasTexto: progs,
        programas: listaDeProgramas(progs),
        bc: celdaNumero(c, k, 1),
        volumenRecuperado: celdaNumero(c, k, 2),
        costo: celdaNumero(c, k, 3),
        beneficio: celdaNumero(c, k, 4)
      });
    }
  }

  var clasificacion = { criterios: [], grupos: [] };
  var g = raw.categ;
  if (g) {
    clasificacion.criterios = [celdaTexto(g, 0, 1), celdaTexto(g, 0, 2), celdaTexto(g, 0, 3)];
    var actual = null;
    for (var m = 1; m < g.v.length; m++) {
      var nombre = celdaTexto(g, m, 0);
      if (nombre) { actual = { nombre: nombre, columnas: [[], [], []] }; clasificacion.grupos.push(actual); }
      if (!actual) continue;
      for (var n = 0; n < 3; n++) { var prog = celdaTexto(g, m, n + 1); if (prog) actual.columnas[n].push(prog); }
    }
  }

  return { plan: plan, combinaciones: combinaciones, clasificacion: clasificacion };
}
