import { test } from "node:test";
import assert from "node:assert/strict";
import { esc, num, fmt, fmtMM, parseNum, hasFormula } from "../src/utils/format.js";

test("esc escapa HTML pero deja pasar texto normal", () => {
  assert.equal(esc('<b>"a" & b</b>'), "&lt;b&gt;&quot;a&quot; &amp; b&lt;/b&gt;");
  assert.equal(esc(null), "");
  assert.equal(esc(undefined), "");
  assert.equal(esc(42), "42");
});

test("num solo acepta números finitos", () => {
  assert.equal(num(3.5), 3.5);
  assert.equal(num("3.5"), null);
  assert.equal(num(NaN), null);
  assert.equal(num(Infinity), null);
  assert.equal(num(null), null);
});

test("fmt formatea en es-CO con los decimales pedidos, o un guion si no hay dato", () => {
  assert.equal(fmt(1234.5, 1), "1.234,5");
  assert.equal(fmt(null), "—");
  assert.equal(fmt("texto"), "—");
});

test("fmtMM divide entre un millón y redondea a entero", () => {
  assert.equal(fmtMM(2500000), "3");
  assert.equal(fmtMM(null), "—");
});

test("parseNum interpreta el formato numérico latinoamericano (punto miles, coma decimal)", () => {
  assert.equal(parseNum("1.234,5"), 1234.5);
  assert.equal(parseNum("  7  "), 7);
  assert.equal(parseNum(""), null);
  assert.equal(parseNum(null), null);
  assert.equal(parseNum("no es un número"), null);
});

test("hasFormula detecta una fórmula en cualquier celda de la matriz", () => {
  assert.equal(hasFormula([["=SUMA(A1:A2)"]]), true);
  assert.equal(hasFormula([[10, 20], ["texto", "=A1"]]), true);
  assert.equal(hasFormula([[10, 20], ["texto", 30]]), false);
  assert.equal(hasFormula([[]]), false);
});
