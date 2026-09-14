/**
 * Pruebas de la capa de integración con Excel, usando un doble de prueba
 * (fake) mínimo de Office.js en lugar de un Excel real. No cubren Office.js
 * en sí -- eso solo se verifica cargando el panel dentro de Excel -- pero sí
 * verifican que excelService.js llama a la API como se espera y que la
 * protección "no escribir sobre una fórmula" funciona.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { escribirRango, leerLibro } from "../src/services/excelService.js";

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

function fakeWorkbook(sheets) {
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
    application: { calculate() {} }
  };
}

function installFakeOffice(sheets) {
  global.Excel = {
    CalculationType: { full: "full" },
    SheetVisibility: { visible: "visible", hidden: "hidden" },
    run: async (callback) => {
      const ctx = { workbook: fakeWorkbook(sheets), sync: async () => {} };
      return callback(ctx);
    }
  };
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
