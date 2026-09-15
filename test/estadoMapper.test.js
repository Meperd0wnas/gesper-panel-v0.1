/**
 * Pruebas de estadoMapper.js: transformar los rangos crudos de PARÁMETROS y
 * Alternativas en un EstadoModelo. Son funciones puras -- no requieren Excel
 * ni DOM.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { mapEstado } from "../src/domain/mappers/estadoMapper.js";

function columna(valores) { return { v: valores.map((v) => [v]), t: valores.map((v) => [v]) }; }
function renglon(valores) { return { v: [valores], t: [valores.map(String)] }; }
function celda(valor) { return { v: [[valor]], t: [[valor]] }; }

function parFixture(vPorIndice) {
  var v = [];
  for (var i = 0; i < 15; i++) v.push([vPorIndice && vPorIndice[i] !== undefined ? vPorIndice[i] : 0]);
  return { v: v, t: v.map(() => [""]) };
}

test("mapEstado arma un EstadoModelo saludable a partir de datos completos", () => {
  var raw = {
    ct005: columna(["Control X: OK", "Control Y: OK", "Control Z: OK"]),
    parE: columna([0.12, 0.09]),
    par: parFixture({ 4: 0.145, 5: 0.052 }),
    flagsV: renglon(["SI", "NO"]),
    flags2: renglon(["SI", "NO"]),
    empresa: celda("Empresa X"),
    anios: renglon(["2021", "2022", "2023"]),
    meta: celda(4.5),
    metodo: celda("Geométrico"),
    monto: celda("$ 5.000.000.000"),
    minbc: celda("1,0"),
    selA1: celda("Combinación 3"),
    prog: { v: [["P.9", 9, 100, 50, 1.5]] }
  };
  var out = mapEstado(raw);
  assert.equal(out.empresa, "Empresa X");
  assert.deepEqual(out.horizonte, { inicio: "2021", fin: "2023", cantidadAnios: 3 });
  assert.equal(out.metodoProyeccionUsuarios, "Geométrico");
  assert.equal(out.metaIPUF, 4.5);
  assert.equal(out.integridad, "ok");
  assert.deepEqual(out.wacc, { real: 0.12, corriente: 0.145 });
  assert.equal(out.alternativaDesfasada, false);
  assert.equal(out.ipc.valor, 0.052);
  assert.equal(out.ipc.desactualizado, false);
  assert.deepEqual(out.programasSinDatos, []);
  assert.deepEqual(out.configuracionMotor, {
    presupuesto: "$ 5.000.000.000", bcMinimo: "1,0", combinacionA1: "Combinación 3"
  });
});

test("mapEstado detecta 'motor_sin_ejecutar' en los controles de integridad", () => {
  var raw = { ct005: columna(["Control 1: OK", "Control 2: OK", "Control 3: SIN EJECUTAR"]) };
  assert.equal(mapEstado(raw).integridad, "motor_sin_ejecutar");
});

test("mapEstado marca 'revisar' cuando algún control no está en OK ni dice SIN EJECUTAR", () => {
  var raw = { ct005: columna(["Control 1: OK", "Control 2: ERROR", "Control 3: OK"]) };
  assert.equal(mapEstado(raw).integridad, "revisar");
});

test("mapEstado devuelve wacc.real null cuando PARÁMETROS!E8 está vacía", () => {
  var raw = { parE: columna(["", 0.09]) };
  assert.equal(mapEstado(raw).wacc.real, null);
});

test("mapEstado marca ipc.desactualizado cuando el IPC es menor a 4.5% (porcentaje en forma decimal)", () => {
  var raw = { par: parFixture({ 5: 0.042 }) };
  var out = mapEstado(raw);
  assert.equal(out.ipc.valor, 0.042);
  assert.equal(out.ipc.desactualizado, true);
});

test("mapEstado detecta desfase entre las banderas de vista y las del consolidado", () => {
  var raw = { flagsV: renglon(["SI", "NO"]), flags2: renglon(["SI", "SI"]) };
  assert.equal(mapEstado(raw).alternativaDesfasada, true);
});

test("mapEstado arma programasSinDatos a partir de mapProgramas (0 beneficio, 0 costo)", () => {
  var raw = {
    prog: {
      v: [
        ["P.9", 9, 100, 50, 1.5],
        ["P.16", 16, 0, 0, 0]
      ]
    }
  };
  var out = mapEstado(raw);
  assert.equal(out.programasSinDatos.length, 1);
  assert.equal(out.programasSinDatos[0].numero, 16);
});

test("mapEstado devuelve horizonte null cuando no hay años", () => {
  assert.equal(mapEstado({}).horizonte, null);
});

test("mapEstado no lanza y devuelve valores por defecto cuando raw viene vacío", () => {
  assert.doesNotThrow(() => mapEstado({}));
  var out = mapEstado({});
  assert.equal(out.empresa, "");
  assert.equal(out.integridad, "revisar");
  assert.deepEqual(out.wacc, { real: null, corriente: null });
  assert.deepEqual(out.programasSinDatos, []);
});
