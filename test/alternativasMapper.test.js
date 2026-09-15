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
