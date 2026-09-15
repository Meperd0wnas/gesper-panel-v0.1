/**
 * Pruebas de programasMapper.js: transformar la tabla cruda de Alternativas!A6:E24
 * en ProgramaResumen[]. Son funciones puras -- no requieren Excel ni DOM.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { mapProgramas } from "../src/domain/mappers/programasMapper.js";

function tabla(filas) {
  return { prog: { v: filas } };
}

test("mapProgramas transforma una fila normal en un ProgramaResumen completo", () => {
  var raw = tabla([["P.9 Renovación de medidores", 9, 1200, 800, 1.5]]);
  var out = mapProgramas(raw);
  assert.equal(out.length, 1);
  assert.deepEqual(out[0], {
    numero: 9,
    nombre: "P.9 Renovación de medidores",
    tipo: "comercial",
    beneficioVPN: 1200,
    costoVPN: 800,
    bc: 1.5,
    estado: "viable",
    nota: ""
  });
});

test("mapProgramas clasifica tipo técnico (1-8) y comercial (9-19)", () => {
  var raw = tabla([
    ["P.1 Macromedición", 1, 500, 400, 1.25],
    ["P.19 Comercial", 19, 300, 300, 1]
  ]);
  var out = mapProgramas(raw);
  assert.equal(out[0].tipo, "tecnico");
  assert.equal(out[1].tipo, "comercial");
});

test("mapProgramas marca 'no_viable' cuando el B/C está entre 0 y 1 sin ser cero", () => {
  var raw = tabla([["P.10 Fraudes", 10, 100, 500, 0.2]]);
  assert.equal(mapProgramas(raw)[0].estado, "no_viable");
});

test("mapProgramas marca 'viable' cuando el B/C es exactamente 1 o más", () => {
  var raw = tabla([["P.13 Anomalías", 13, 100, 100, 1]]);
  assert.equal(mapProgramas(raw)[0].estado, "viable");
});

test("mapProgramas marca 'sin_datos' cuando beneficio y costo son 0 (ambigüedad documentada en §8)", () => {
  var raw = tabla([["P.16 Software", 16, 0, 0, 0]]);
  assert.equal(mapProgramas(raw)[0].estado, "sin_datos");
});

test("mapProgramas marca 'sin_datos' cuando el B/C no es un número (celda vacía)", () => {
  var raw = tabla([["P.X sin diligenciar", 5, 100, 200, ""]]);
  var out = mapProgramas(raw)[0];
  assert.equal(out.estado, "sin_datos");
  assert.equal(out.bc, null);
});

test("mapProgramas marca 'no_aplica' y conserva la nota cuando la columna de beneficio trae texto (caso P.8)", () => {
  var raw = tabla([["P.8 Técnico", 8, "Se evalúa junto con P.19", null, null]]);
  var out = mapProgramas(raw)[0];
  assert.equal(out.estado, "no_aplica");
  assert.equal(out.nota, "Se evalúa junto con P.19");
  assert.equal(out.beneficioVPN, null);
});

test("mapProgramas ignora filas vacías o sin número de programa válido", () => {
  var raw = tabla([
    ["", "", "", "", ""],
    ["Encabezado de sección", "no es un número", "", "", ""],
    ["P.13 Anomalías", 13, 50, 40, 1.25]
  ]);
  var out = mapProgramas(raw);
  assert.equal(out.length, 1);
  assert.equal(out[0].numero, 13);
});

test("mapProgramas devuelve una lista vacía cuando falta la clave 'prog' en raw", () => {
  assert.deepEqual(mapProgramas({}), []);
});

test("mapProgramas procesa una lista con los 19 programas reales sin lanzar", () => {
  var filas = [];
  for (var n = 1; n <= 19; n++) filas.push(["P." + n, n, n * 10, n * 5, n % 3 === 0 ? 0 : 1.1]);
  var out = mapProgramas(tabla(filas));
  assert.equal(out.length, 19);
});
