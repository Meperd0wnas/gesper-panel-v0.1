/**
 * Pruebas de resultadosMapper.js: transformar los rangos crudos de
 * Alternativas/FCAJA/PARÁMETROS en ResultadosHidraulicos y
 * ResultadosFinancieros. Son funciones puras -- no requieren Excel ni DOM.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { mapResultadosHidraulicos, mapResultadosFinancieros } from "../src/domain/mappers/resultadosMapper.js";

function renglon(valores) { return { v: [valores], t: [valores.map(String)] }; }
function celda(valor) { return { v: [[valor]], t: [[valor]] }; }

function parFixture(vPorIndice, tPorIndice) {
  var v = [], t = [];
  for (var i = 0; i < 15; i++) {
    v.push([vPorIndice && vPorIndice[i] !== undefined ? vPorIndice[i] : 0]);
    t.push([tPorIndice && tPorIndice[i] !== undefined ? tPorIndice[i] : ""]);
  }
  return { v: v, t: t };
}

/* ---------- mapResultadosHidraulicos ---------- */

test("mapResultadosHidraulicos calcula el IPUF promedio desde el 4.º año y evalúa la meta", () => {
  var raw = {
    ipuf2: renglon([9, 9, 9, 5, 3]),
    meta: celda(4.5),
    ianc2: renglon([40, 35, 30, 25, 20]),
    anios: renglon(["2021", "2022", "2023", "2024", "2025"])
  };
  var out = mapResultadosHidraulicos(raw);
  assert.equal(out.ipufPromedio, 4); // (5+3)/2
  assert.equal(out.cumpleMeta, true); // 4 < 4.5
  assert.deepEqual(out.ianc, { inicio: 40, fin: 20 });
  assert.equal(out.serieIPUF.length, 5);
  assert.deepEqual(out.serieIPUF[4], { anio: "2025", valor: 3 });
});

test("mapResultadosHidraulicos marca cumpleMeta=false cuando el promedio supera la meta", () => {
  var raw = { ipuf2: renglon([9, 9, 9, 6, 7]), meta: celda(4.5) };
  assert.equal(mapResultadosHidraulicos(raw).cumpleMeta, false);
});

test("mapResultadosHidraulicos devuelve valores vacíos/null cuando no hay datos", () => {
  var out = mapResultadosHidraulicos({});
  assert.equal(out.ipufPromedio, null);
  assert.equal(out.cumpleMeta, null);
  assert.deepEqual(out.ianc, { inicio: null, fin: null });
  assert.deepEqual(out.serieIPUF, []);
});

test("mapResultadosHidraulicos deja como null las celdas con texto en la serie IPUF (valor inesperado)", () => {
  var raw = { ipuf2: renglon([9, "n/d", 9, 4, 3]) };
  var out = mapResultadosHidraulicos(raw);
  assert.equal(out.serieIPUF[1].valor, null);
});

/* ---------- mapResultadosFinancieros ---------- */

test("mapResultadosFinancieros conserva vna/tir/payback como texto tal como vienen del libro", () => {
  var raw = {
    fcaja: { t: [["4 años"], ["18,3 %"], ["$ 1.234 millones"]] },
    tirm: celda(0.201),
    par: parFixture({}, { 4: "14,5 %", 11: "$ 3.200", 12: "$ 1.100" }),
    grad: celda("2,0 %")
  };
  var out = mapResultadosFinancieros(raw);
  assert.equal(out.payback, "4 años");
  assert.equal(out.tir, "18,3 %");
  assert.equal(out.tirDefinida, true);
  assert.equal(out.tirModificada, 0.201);
  assert.equal(out.vna, "$ 1.234 millones");
  assert.deepEqual(out.parametros, {
    waccCorriente: "14,5 %", cargoFijoAcueducto: "$ 3.200", cargoVariableAcueducto: "$ 1.100", gradiente: "2,0 %"
  });
});

test("mapResultadosFinancieros marca tirDefinida=false cuando la TIR no está definida ('N.A.')", () => {
  var raw = { fcaja: { t: [["—"], ["N.A. (flujo sin cambio de signo)"], ["—"]] } };
  assert.equal(mapResultadosFinancieros(raw).tirDefinida, false);
});

test("mapResultadosFinancieros marca alcantarilladoEnCero cuando ambos cargos son 0 (booleano derivado)", () => {
  var raw = { par: parFixture({ 13: 0, 14: 0 }) };
  assert.equal(mapResultadosFinancieros(raw).alcantarilladoEnCero, true);
});

test("mapResultadosFinancieros no marca alcantarilladoEnCero cuando algún cargo tiene valor", () => {
  var raw = { par: parFixture({ 13: 1500, 14: 900 }) };
  assert.equal(mapResultadosFinancieros(raw).alcantarilladoEnCero, false);
});

test("mapResultadosFinancieros no lanza y devuelve null cuando faltan todas las hojas de origen", () => {
  assert.doesNotThrow(() => mapResultadosFinancieros({}));
  var out = mapResultadosFinancieros({});
  assert.equal(out.vna, null);
  assert.equal(out.tir, null);
  assert.equal(out.tirDefinida, false);
  assert.equal(out.payback, null);
});
