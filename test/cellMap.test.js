import { test } from "node:test";
import assert from "node:assert/strict";
import { H, LECTURAS, BLOQUES, ANCLAS } from "../src/config/cellMap.js";

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

test("cada ancla de programa tiene hoja, dirección y una etiqueta legible", () => {
  for (const [n, ancla] of Object.entries(ANCLAS)) {
    assert.equal(ancla.length, 3, `ancla del programa ${n} mal formada`);
    const [, , etiqueta] = ancla;
    assert.ok(etiqueta.length > 0);
  }
});
