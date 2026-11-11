import { test } from "node:test";
import assert from "node:assert/strict";
import { H, LECTURAS, BLOQUES, ANCLAS, ANEXOS } from "../src/config/cellMap.js";

test("toda lectura es una tupla [clave, hoja, direccion] con nombre de hoja conocido en H", () => {
  const hojasValidas = new Set(Object.values(H));
  for (const lectura of LECTURAS) {
    assert.equal(lectura.length, 3, `lectura mal formada: ${JSON.stringify(lectura)}`);
    const [clave, hoja, addr] = lectura;
    assert.equal(typeof clave, "string");
    assert.ok(hojasValidas.has(hoja), `hoja "${hoja}" de la lectura "${clave}" no está en H`);
    assert.match(addr, /^[A-Z]+\d+(:[A-Z]+\d+)?$/, `dirección "${addr}" de "${clave}" no parece un rango de Excel`);
  }
});

test("las claves de LECTURAS son únicas", () => {
  const claves = LECTURAS.map((l) => l[0]);
  assert.equal(new Set(claves).size, claves.length);
});

test("cada bloque de navegación apunta a una hoja+celda", () => {
  for (const b of BLOQUES) {
    assert.equal(typeof b.n, "string");
    assert.equal(b.ir.length, 2);
  }
});

test("cada anexo tiene id único, hoja destino conocida y bloques de horizonte que apuntan a una fila de bandera", () => {
  const hojasValidas = new Set(Object.values(H));
  const ids = ANEXOS.map((a) => a.id);
  assert.equal(new Set(ids).size, ids.length);
  for (const a of ANEXOS) {
    assert.match(a.id, /^A\d{2}$/);
    assert.ok(hojasValidas.has(a.hojaDestino), `hoja destino "${a.hojaDestino}" de ${a.id} no está en H`);
    assert.ok(Number.isInteger(a.filaAnios) && a.filaAnios > 0);
    assert.ok(["concepto", "grupo"].includes(a.alcanceControles), `alcanceControles de ${a.id}`);
    for (const c of a.celdasProtegidas || []) assert.match(c, /^[A-Z]+\d+$/, `celda protegida "${c}" de ${a.id}`);
    for (const b of a.banderasHorizonte) {
      assert.ok(b.desde <= b.hasta && b.hasta < b.fila, `bloque ${b.desde}-${b.hasta} de ${a.id}: la bandera va debajo`);
      assert.ok(a.filasBandera.includes(b.fila), `la bandera ${b.fila} de ${a.id} debe estar en filasBandera`);
    }
  }
});

test("los 19 programas tienen ancla: técnicos (1-8) en EV PROG TÉCNICOS, comerciales (9-19) en EV PROG CCIALES", () => {
  for (let n = 1; n <= 19; n++) {
    assert.ok(ANCLAS[n], `el programa ${n} no tiene ancla: el panel caería a buscar su rótulo`);
    assert.equal(ANCLAS[n][0], n <= 8 ? H.EVT : H.EVC, `hoja del ancla del programa ${n}`);
    assert.match(ANCLAS[n][1], /^[A-Z]+\d+(:[A-Z]+\d+)?$/, `dirección del ancla del programa ${n}`);
  }
});

test("cada ancla de programa tiene hoja, dirección y una etiqueta legible", () => {
  for (const [n, ancla] of Object.entries(ANCLAS)) {
    assert.equal(ancla.length, 3, `ancla del programa ${n} mal formada`);
    const [, , etiqueta] = ancla;
    assert.ok(etiqueta.length > 0);
  }
});
