/**
 * Pruebas de la capa de integración con Excel, usando un doble de prueba
 * (fake) mínimo de Office.js en lugar de un Excel real. No cubren Office.js
 * en sí -- eso solo se verifica cargando el panel dentro de Excel -- pero sí
 * verifican que excelService.js llama a la API como se espera y que la
 * protección "no escribir sobre una fórmula" funciona.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { escribirRango, escribirCampos, leerLibro } from "../src/services/excelService.js";

function fakeRange({ formulas, values, text, address } = {}) {
  return {
    formulas: formulas || [[""]],
    values: values || [[""]],
    text: text || [[""]],
    address: address || "Hoja1!A1",
    load() {},
    select() {}
  };
}

function fakeWorksheet(name, rangesByAddr = {}) {
  return {
    name,
    visibility: null,
    load() {},
    activate() {},
    getRange(addr) { return rangesByAddr[addr] || fakeRange({ address: name + "!" + addr }); }
  };
}

function fakeWorkbook(sheets, stats) {
  const items = Object.values(sheets);
  return {
    worksheets: {
      items,
      load() {},
      getItem(nombre) {
        if (!sheets[nombre]) throw new Error("Hoja no encontrada: " + nombre);
        return sheets[nombre];
      },
      getItemOrNullObject(nombre) {
        return sheets[nombre] || { isNullObject: true, load() {} };
      },
      add(nombre) {
        const s = fakeWorksheet(nombre);
        sheets[nombre] = s;
        items.push(s);
        return s;
      }
    },
    application: { calculate() { stats.calculateCalls++; } }
  };
}

/**
 * Instala el doble de Office.js. Devuelve `stats` (runCalls, calculateCalls)
 * para que los tests de escribirCampos puedan verificar cuántas veces se
 * llamó Excel.run y application.calculate, sin cambiar el comportamiento de
 * los tests que ya existían y no usan ese valor de retorno.
 */
function installFakeOffice(sheets) {
  const stats = { runCalls: 0, calculateCalls: 0 };
  global.Excel = {
    CalculationType: { full: "full" },
    SheetVisibility: { visible: "visible", hidden: "hidden" },
    run: async (callback) => {
      stats.runCalls++;
      const ctx = { workbook: fakeWorkbook(sheets, stats), sync: async () => {} };
      return callback(ctx);
    }
  };
  return stats;
}

test("escribirRango rechaza escribir sobre una celda con fórmula", async () => {
  installFakeOffice({
    PAR: fakeWorksheet("PAR", { C8: fakeRange({ formulas: [["=SUMA(1,2)"]], address: "PAR!C8" }) })
  });
  await assert.rejects(
    () => escribirRango("PAR", "C8", [[99]]),
    /contiene una fórmula/
  );
});

test("escribirRango escribe cuando la celda de destino no tiene fórmula", async () => {
  const celda = fakeRange({ formulas: [[""]], address: "EVC!C181" });
  installFakeOffice({ EVC: fakeWorksheet("EVC", { C181: celda }) });
  await escribirRango("EVC", "C181", [[12.5]]);
  assert.deepEqual(celda.values, [[12.5]]);
});

test("leerLibro solo lee las direcciones cuya hoja existe en el libro", async () => {
  const rango = fakeRange({ values: [["Empresa X"]], text: [["Empresa X"]] });
  installFakeOffice({ EVC: fakeWorksheet("EVC", { F2: rango }) });
  const { hojasLibro, datos } = await leerLibro([
    ["empresa", "EVC", "F2"],
    ["noExiste", "HOJA_QUE_NO_EXISTE", "A1"]
  ]);
  assert.deepEqual(hojasLibro, ["EVC"]);
  assert.equal(datos.empresa.t[0][0], "Empresa X");
  assert.equal(datos.noExiste, undefined);
});

/* ---------- escribirCampos (escritura por lotes) ---------- */

test("escribirCampos escribe varios campos, en distintas hojas, en una sola ejecución de Excel.run", async () => {
  const a1 = fakeRange({ formulas: [[""]], address: "PAR!A1" });
  const b2 = fakeRange({ formulas: [[""]], address: "PAR!B2" });
  const c3 = fakeRange({ formulas: [[""]], address: "EVC!C3" });
  const stats = installFakeOffice({
    PAR: fakeWorksheet("PAR", { A1: a1, B2: b2 }),
    EVC: fakeWorksheet("EVC", { C3: c3 })
  });

  await escribirCampos([
    { hoja: "PAR", direccion: "A1", valores: [[1]] },
    { hoja: "PAR", direccion: "B2", valores: [[2]] },
    { hoja: "EVC", direccion: "C3", valores: [[3]] }
  ]);

  assert.deepEqual(a1.values, [[1]]);
  assert.deepEqual(b2.values, [[2]]);
  assert.deepEqual(c3.values, [[3]]);
  assert.equal(stats.runCalls, 1, "las 3 escrituras deben ir en una sola ejecución de Excel.run");
});

test("escribirCampos recalcula una sola vez, sin importar cuántos campos se escriban", async () => {
  const celdas = ["A1", "A2", "A3", "A4"].map((addr) => fakeRange({ formulas: [[""]], address: "PAR!" + addr }));
  const rangesByAddr = Object.fromEntries(["A1", "A2", "A3", "A4"].map((addr, i) => [addr, celdas[i]]));
  const stats = installFakeOffice({ PAR: fakeWorksheet("PAR", rangesByAddr) });

  await escribirCampos([
    { hoja: "PAR", direccion: "A1", valores: [[1]] },
    { hoja: "PAR", direccion: "A2", valores: [[2]] },
    { hoja: "PAR", direccion: "A3", valores: [[3]] },
    { hoja: "PAR", direccion: "A4", valores: [[4]] }
  ]);

  assert.equal(stats.calculateCalls, 1, "calculate(full) debe llamarse exactamente una vez por lote");
});

test("escribirCampos rechaza el lote si alguna entrada apunta a una celda con fórmula", async () => {
  const c8 = fakeRange({ formulas: [["=SUMA(1,2)"]], address: "PAR!C8" });
  installFakeOffice({ PAR: fakeWorksheet("PAR", { C8: c8 }) });

  await assert.rejects(
    () => escribirCampos([{ hoja: "PAR", direccion: "C8", valores: [[99]] }]),
    /contiene una fórmula/
  );
});

test("escribirCampos no escribe ninguna entrada del lote si una sola tiene fórmula (todo o nada)", async () => {
  const a1 = fakeRange({ formulas: [[""]], values: [["sin tocar"]], address: "PAR!A1" });
  const b2 = fakeRange({ formulas: [[""]], values: [["sin tocar"]], address: "PAR!B2" });
  const c3 = fakeRange({ formulas: [["=A1+B2"]], values: [["formula"]], address: "PAR!C3" });
  installFakeOffice({ PAR: fakeWorksheet("PAR", { A1: a1, B2: b2, C3: c3 }) });

  await assert.rejects(
    () => escribirCampos([
      { hoja: "PAR", direccion: "A1", valores: [[111]] },
      { hoja: "PAR", direccion: "B2", valores: [[222]] },
      { hoja: "PAR", direccion: "C3", valores: [[333]] } // esta tiene fórmula
    ]),
    /contiene una fórmula/
  );

  // Ninguna de las tres debió escribirse, ni siquiera las que no tenían fórmula.
  assert.deepEqual(a1.values, [["sin tocar"]]);
  assert.deepEqual(b2.values, [["sin tocar"]]);
  assert.deepEqual(c3.values, [["formula"]]);
});

test("escribirCampos rechaza una lista vacía", async () => {
  await assert.rejects(() => escribirCampos([]), /al menos una entrada/);
});

test("escribirCampos rechaza una entrada sin 'valores' en forma de matriz", async () => {
  installFakeOffice({ PAR: fakeWorksheet("PAR", {}) });
  await assert.rejects(
    () => escribirCampos([{ hoja: "PAR", direccion: "A1", valores: 5 }]),
    /matriz/
  );
});

test("escribirCampos rechaza una entrada sin hoja", async () => {
  await assert.rejects(
    () => escribirCampos([{ direccion: "A1", valores: [[1]] }]),
    /hoja válida/
  );
});

test("escribirRango sigue funcionando igual que antes (regresión: ahora delega en escribirCampos)", async () => {
  const c8 = fakeRange({ formulas: [["=SUMA(1,2)"]], address: "PAR!C8" });
  installFakeOffice({ PAR: fakeWorksheet("PAR", { C8: c8 }) });
  await assert.rejects(
    () => escribirRango("PAR", "C8", [[99]]),
    /contiene una fórmula/
  );

  const celda = fakeRange({ formulas: [[""]], address: "EVC!C181" });
  installFakeOffice({ EVC: fakeWorksheet("EVC", { C181: celda }) });
  await escribirRango("EVC", "C181", [[12.5]]);
  assert.deepEqual(celda.values, [[12.5]]);
});
