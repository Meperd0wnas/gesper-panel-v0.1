/**
 * Pruebas de alternativasMapper.js: transformar los rangos crudos de la hoja
 * Alternativas en Alternativa[]. Son funciones puras -- no requieren Excel ni
 * DOM.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { mapAlternativas } from "../src/domain/mappers/alternativasMapper.js";

function bc(beneficio, costo) { return { v: [[beneficio], [costo]] }; }
function serie(valores) { return { v: [valores] }; }
function texto(valor) { return { t: [[valor]] }; }

test("mapAlternativas arma las 3 alternativas con beneficio, costo y B/C calculado", () => {
  var raw = {
    a1bc: bc(1200, 800),
    a2bc: bc(1000, 1000),
    a3bc: bc(0, 0),
    selA1: texto("Combinación 3"),
    a2txt: texto("Compuesta a mano"),
    a3txt: texto("")
  };
  var out = mapAlternativas(raw);
  assert.equal(out.length, 3);
  assert.deepEqual(out.map((a) => a.id), [1, 2, 3]);
  assert.equal(out[0].bc, 1.5);
  assert.equal(out[1].bc, 1);
  assert.equal(out[2].bc, null); // costo 0 -> no se puede dividir
  assert.equal(out[0].composicionTexto, "Combinación 3");
  assert.equal(out[1].composicionTexto, "Compuesta a mano");
});

test("mapAlternativas calcula el IPUF promedio de la Alternativa 1 y 2 desde el 4.º año", () => {
  var raw = { ipuf1: serie([9, 9, 9, 5, 3]), ipuf2: serie([9, 9, 9, 4, 4]) };
  var out = mapAlternativas(raw);
  assert.equal(out[0].ipufPromedio, 4); // (5+3)/2
  assert.equal(out[1].ipufPromedio, 4); // (4+4)/2
});

test("mapAlternativas deja ipufPromedio en null para la Alternativa 3 (sin lectura en cellMap)", () => {
  var out = mapAlternativas({ ipuf1: serie([9, 9, 9, 5, 3]) });
  assert.equal(out[2].ipufPromedio, null);
});

test("mapAlternativas interpreta 'CUMPLE'/'NO CUMPLE' como booleano, y null si no hay dato reconocible", () => {
  var raw = { ver1: texto("CUMPLE"), ver2: texto("NO CUMPLE"), ver3: texto("") };
  var out = mapAlternativas(raw);
  assert.equal(out[0].cumpleMeta, true);
  assert.equal(out[1].cumpleMeta, false);
  assert.equal(out[2].cumpleMeta, null);
});

test("mapAlternativas devuelve null en beneficio/costo/bc/ipufPromedio/cumpleMeta cuando falta el rango de origen", () => {
  var out = mapAlternativas({});
  assert.deepEqual(out[0], {
    id: 1, beneficio: null, costo: null, bc: null, ipufPromedio: null, cumpleMeta: null, composicionTexto: ""
  });
});

test("mapAlternativas ignora un rango de IPUF con menos de 4 años válidos (dato insuficiente)", () => {
  var raw = { ipuf1: serie([9, 9]) };
  assert.equal(mapAlternativas(raw)[0].ipufPromedio, null);
});

/* ---------- Detalle por alternativa y plan de gestión (informe de resultados) ---------- */
import { mapDetalleAlternativas, mapPlanGestion, normalizarRotulo, listaDeProgramas } from "../src/domain/mappers/alternativasMapper.js";

/** Bloque H:V: rótulo, unidad y un valor por año. */
function bloque(filas) {
  var v = filas.map(function (f) { return [f[0], f[1]].concat(f[2]); });
  return { v: v, t: v.map(function (r) { return r.map(String); }) };
}
function banderas(marcas) { return { v: [marcas], t: [marcas] }; }

test("normalizarRotulo quita tildes, unidad entre corchetes y espacios de sobra", () => {
  assert.equal(normalizarRotulo(" FACTURACIÓN AL APLICAR EL PROGRAMA"), "FACTURACION AL APLICAR EL PROGRAMA");
  assert.equal(normalizarRotulo("IPUF [m³/usuarios/mes]"), "IPUF");
  assert.equal(normalizarRotulo(null), "");
});

test("listaDeProgramas convierte el texto de la hoja en números, con coma final o sin ella", () => {
  assert.deepEqual(listaDeProgramas("1, 2, 3, 9, 10, "), [1, 2, 3, 9, 10]);
  assert.deepEqual(listaDeProgramas(""), []);
});

test("mapDetalleAlternativas ubica cada fila por su rótulo, no por su posición", () => {
  var raw = {
    anios: { v: [["2023", "2024", "2025", "2026", "2027"]], t: [["2023", "2024", "2025", "2026", "2027"]] },
    // Orden cambiado a propósito: IPUF antes que PRODUCCIÓN, y PRODUCCIÓN AL APLICAR antes que PRODUCCIÓN.
    alt2bal: bloque([
      ["IPUF [m³/usuarios/mes]", "m³/usuarios/mes", [5, 5, 5, 4, 2]],
      ["PRODUCCIÓN AL APLICAR EL PROGRAMA", "m³/año", [90, 90, 90, 80, 70]],
      ["PRODUCCIÓN", "m³/año", [100, 100, 100, 100, 100]],
      ["IANC [%]", "%", [30, 30, 30, 25, 20]]
    ])
  };
  var a2 = mapDetalleAlternativas(raw).alternativas[1];
  assert.deepEqual(a2.series.ipuf, [5, 5, 5, 4, 2]);
  assert.deepEqual(a2.series.produccion, [100, 100, 100, 100, 100]);
  assert.deepEqual(a2.series.produccionConProgramas, [90, 90, 90, 80, 70]);
  assert.deepEqual(a2.series.ianc, [30, 30, 30, 25, 20]);
  assert.equal(a2.series.icuf, null, "una fila que no está en el bloque queda en null");
  assert.equal(a2.ipufPromedio, 3, "promedio desde el 4.º año: (4 + 2) / 2");
});

test("mapDetalleAlternativas toma los programas marcados con SI según el número de la fila 6", () => {
  var raw = {
    altNums: { v: [[1, 2, 3, 4]], t: [["1", "2", "3", "4"]] },
    flagsV: banderas(["SI", "", "si ", "NO"]),
    flags2: banderas(["", "SI", "", "SI"]),
    ver1: { t: [["La ALTERNATIVA No. 1 ... NO CUMPLE con las metas"]] },
    alt2nota: { t: [["**EVALUADA EN EVOLUCIÓN DEL BALANCE HÍDRICO Y FINANCIERAMENTE"]] }
  };
  var out = mapDetalleAlternativas(raw).alternativas;
  assert.deepEqual(out[0].programas, [1, 3]);
  assert.deepEqual(out[1].programas, [2, 4]);
  assert.deepEqual(out[2].programas, [], "sin fila de banderas, sin programas");
  assert.equal(out[0].veredicto, "La ALTERNATIVA No. 1 ... NO CUMPLE con las metas");
  assert.equal(out[1].nota, "EVALUADA EN EVOLUCIÓN DEL BALANCE HÍDRICO Y FINANCIERAMENTE");
});

test("mapDetalleAlternativas asume el orden de la hoja si no pudo leer los números de programa", () => {
  var out = mapDetalleAlternativas({ flags3: banderas(["", "SI", "SI"]) }).alternativas;
  assert.deepEqual(out[2].programas, [2, 3]);
});

test("mapDetalleAlternativas no lanza sin datos y deja todo vacío", () => {
  var out = mapDetalleAlternativas({});
  assert.deepEqual(out.anios, []);
  assert.equal(out.inicioPromedio, 3);
  assert.equal(out.alternativas.length, 3);
  assert.equal(out.alternativas[0].ipufPromedio, null);
  assert.equal(out.alternativas[0].series.ipuf, null);
});

test("mapPlanGestion lee el resumen del plan por rótulo, las combinaciones y la clasificación", () => {
  function rango(filas) { return { v: filas, t: filas.map(function (r) { return r.map(function (x) { return x === "" ? "" : String(x); }); }) }; }
  var raw = {
    plan: rango([
      ["El plan de gestión de pérdidas incluye los programas requeridos", "", "", "", ""],
      ["", "", "", "", ""],
      ["Programas requisito esencial", "1. Macromedición", "2. Sectorización", "", ""],
      ["Total de Programas Plan de Gestión", "1, 2, 9, ", "", "", ""],
      ["", " Costo", 300, "", ""],
      ["", " Beneficio", 450, "", ""],
      ["", "Relación BC", 1.5, "", ""],
      ["", "Vol. Recuperado", 1000, "", ""]
    ]),
    combos: rango([
      ["1, 2, 9", 1.5, 1000, 300, 450],
      ["1, 2", 1.2, 800, 200, 240],
      ["", "", "", "", ""]
    ]),
    categ: rango([
      ["", "Obligatorios", "Mayor Volumen recuperado", "Mayor Relación B/C"],
      ["Técnicos", "P.1 Macromedición", "P.7 Reposición", "P.5 Daños"],
      ["", "P.2 Sectorización", "", ""],
      ["Comerciales", "", "P.9 Medidores", "P.16 Comunidades"]
    ])
  };
  var out = mapPlanGestion(raw);
  assert.deepEqual(out.plan.requisitoEsencial, ["1. Macromedición", "2. Sectorización"]);
  assert.equal(out.plan.programasTexto, "1, 2, 9");
  assert.deepEqual(out.plan.programas, [1, 2, 9]);
  assert.equal(out.plan.costo, 300);
  assert.equal(out.plan.beneficio, 450);
  assert.equal(out.plan.bc, 1.5);
  assert.equal(out.plan.volumenRecuperado, 1000);
  assert.equal(out.combinaciones.length, 2, "las filas vacías no cuentan");
  assert.deepEqual(out.combinaciones[1], { programasTexto: "1, 2", programas: [1, 2], bc: 1.2, volumenRecuperado: 800, costo: 200, beneficio: 240 });
  assert.deepEqual(out.clasificacion.criterios, ["Obligatorios", "Mayor Volumen recuperado", "Mayor Relación B/C"]);
  assert.deepEqual(out.clasificacion.grupos[0], { nombre: "Técnicos", columnas: [["P.1 Macromedición", "P.2 Sectorización"], ["P.7 Reposición"], ["P.5 Daños"]] });
  assert.deepEqual(out.clasificacion.grupos[1].columnas, [[], ["P.9 Medidores"], ["P.16 Comunidades"]]);
});

test("mapPlanGestion no lanza sin datos", () => {
  var out = mapPlanGestion({});
  assert.equal(out.plan.costo, null);
  assert.deepEqual(out.combinaciones, []);
  assert.deepEqual(out.clasificacion, { criterios: [], grupos: [] });
});
