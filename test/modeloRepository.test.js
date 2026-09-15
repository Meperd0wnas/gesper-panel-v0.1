/**
 * Pruebas de modeloRepository.js: verifica que cargarModelo() lee el libro en
 * una sola ejecución de Excel.run y compone correctamente los 4 mappers.
 * Usa el mismo tipo de doble de prueba de Office.js que excelService.test.js
 * -- no depende de Excel real.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { cargarModelo } from "../src/domain/repositories/modeloRepository.js";

/**
 * Rango vacío por defecto lo bastante grande (20x20) para que cualquier
 * lectura LECTURAS -- sea vertical, horizontal o de una sola celda -- se
 * resuelva a "" en vez de a un índice indefinido.
 */
function vacio() {
  return Array.from({ length: 20 }, function () { return Array.from({ length: 20 }, function () { return ""; }); });
}

function fakeRange({ values, text, address } = {}) {
  return {
    values: values || vacio(),
    text: text || vacio(),
    address: address || "Hoja1!A1",
    load() {},
    select() {}
  };
}

function fakeWorksheet(name, rangesByAddr = {}) {
  return {
    name,
    load() {},
    getRange(addr) { return rangesByAddr[addr] || fakeRange({ address: name + "!" + addr }); }
  };
}

function installFakeOffice(sheets) {
  var stats = { runCalls: 0 };
  var items = Object.values(sheets);
  global.Excel = {
    run: async (callback) => {
      stats.runCalls++;
      var ctx = {
        workbook: {
          worksheets: {
            items: items,
            load() {},
            getItem(nombre) { return sheets[nombre]; }
          }
        },
        sync: async () => {}
      };
      return callback(ctx);
    }
  };
  return stats;
}

test("cargarModelo lee el libro en una sola ejecución de Excel.run y compone los 4 mappers", async () => {
  var stats = installFakeOffice({
    "Alternativas": fakeWorksheet("Alternativas", {
      "A6:E24": fakeRange({ values: [["P.9 Renovación de medidores", 9, 100, 50, 2]] })
    }),
    "PARÁMETROS": fakeWorksheet("PARÁMETROS", {})
  });

  var modelo = await cargarModelo();

  assert.equal(stats.runCalls, 1, "debe leer el libro en una sola ejecución de Excel.run");
  assert.equal(modelo.programas.length, 1);
  assert.equal(modelo.programas[0].numero, 9);
  assert.equal(modelo.estado.empresa, "");
  assert.deepEqual(modelo.alternativas.map((a) => a.id), [1, 2, 3]);
  assert.ok(modelo.leidoEn instanceof Date);
  assert.deepEqual(modelo.hojasLibro.slice().sort(), ["Alternativas", "PARÁMETROS"]);
});

test("cargarModelo expone 'raw' como puente temporal para pantallas aún no migradas", async () => {
  installFakeOffice({
    "Alternativas": fakeWorksheet("Alternativas", {
      "A6:E24": fakeRange({ values: [["P.9", 9, 100, 50, 2]] })
    })
  });
  var modelo = await cargarModelo();
  assert.ok(modelo.raw.prog);
  assert.deepEqual(modelo.raw.prog.v, [["P.9", 9, 100, 50, 2]]);
});

test("cargarModelo no lanza y devuelve valores por defecto cuando el libro no tiene ninguna hoja esperada", async () => {
  installFakeOffice({});
  var modelo = await cargarModelo();
  assert.deepEqual(modelo.programas, []);
  assert.equal(modelo.estado.integridad, "revisar");
  assert.deepEqual(modelo.hojasLibro, []);
});
