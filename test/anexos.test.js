/**
 * Pruebas de la carga de anexos: contrato SALIDA_GESPER
 * (src/domain/anexos/salidaGesper.js) y su composición con excelService
 * (src/domain/repositories/anexosRepository.js). El libro del anexo se arma
 * directamente con la forma que devuelve utils/xlsx.js (el lector tiene sus
 * propias pruebas en xlsx.test.js) y GESPER se simula con un doble de Office.js.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { ANEXOS } from "../src/config/cellMap.js";
import {
  COLUMNAS, interpretarSalida, evaluarAnexo, direccionesAInspeccionar, planificarCarga, dependeDecision, normalizarTexto
} from "../src/domain/anexos/salidaGesper.js";
import { ejecutarCarga, anexoPorId } from "../src/domain/repositories/anexosRepository.js";
import { columnaALetra } from "../src/utils/xlsx.js";

const A01 = ANEXOS[0];
const EVT = "EV PROG TÉCNICOS";
const ETQ_INST = 'Suministro e instalación de Medidor Ultrasónico de Caudal Clamp-On 14"-60"';
const ETQ_REEMP = 'Reemplazo macromedidor entre 14" hasta 60"  incluye suministro e instalación';

/** Fila de SALIDA_GESPER a partir de las columnas que importan en cada prueba. */
function fila(o) {
  return Object.assign({
    PROGRAMA: "P1", SITIO: "Captación", CONCEPTO: "concepto de " + o.ID, RANGO_CLAVE: null, ANIO_IDX: null,
    UNIDAD: "un", VALOR_CONTROL: null, ORIGEN_HOJA: "x", ORIGEN_REF: "x", DESTINO_HOJA: EVT,
    TRANSFORMACION: "Ninguna", DEPENDE_DECISION: "NO", NOTA: null
  }, o);
}

const FILAS = [
  fila({ ID: "A01.CAP.PUNTOS", TIPO_FILA: "DATO", VALOR: 7, ESTADO: "OK", DESTINO_CELDA: "D108", DESTINO_ETIQUETA: "Número de puntos de captación" }),
  fila({ ID: "A01.CAP.INST.R14_60.Y02", TIPO_FILA: "DATO", ANIO_IDX: 2, VALOR: 0, ESTADO: "OK", DESTINO_CELDA: "M115", DESTINO_ETIQUETA: ETQ_INST }),
  fila({ ID: "A01.SAL.REEMP.R14_60.Y02", TIPO_FILA: "DATO", SITIO: "Salida PTAP", ANIO_IDX: 2, VALOR: 0, VALOR_CONTROL: 1, ESTADO: "ADVERTENCIA",
    DESTINO_CELDA: "M347", DESTINO_ETIQUETA: ETQ_REEMP, DEPENDE_DECISION: "SI – filtro de la tabla dinámica" }),
  fila({ ID: "A01.CAP.REEMP.R6.Y01", TIPO_FILA: "DATO", ANIO_IDX: 1, VALOR: 4, ESTADO: "REVISAR", DESTINO_CELDA: "L178", DESTINO_ETIQUETA: "Reemplazo <=6" }),
  fila({ ID: "A01.CAP.CTRL_INST", TIPO_FILA: "CONTROL", VALOR: 2, VALOR_CONTROL: 2, ESTADO: "OK", DESTINO_HOJA: null }),
  fila({ ID: "A01.GEN.CONTEMPLADO", TIPO_FILA: "NO_DISPONIBLE", SITIO: "General", VALOR: null, ESTADO: "N/A", DESTINO_HOJA: null })
];

/** Libro con la forma de leerXlsx: metadatos en A3:H4 y tabla desde A6, con o sin nombres definidos. */
function libroAnexo({ meta = {}, filas = FILAS, encabezados = COLUMNAS, conNombres = true } = {}) {
  const m = Object.assign({ SPEC_VERSION: "1.0", ANEXO_ID: "A01", PROGRAMAS: "P1", ESTADO_GLOBAL: "OK CON ADVERTENCIAS" }, meta);
  const celdas = { A1: "SALIDA_GESPER · título" };
  Object.keys(m).forEach((k, i) => { celdas[columnaALetra(i + 1) + 3] = k; celdas[columnaALetra(i + 1) + 4] = m[k]; });
  encabezados.forEach((h, i) => { celdas[columnaALetra(i + 1) + 6] = h; });
  filas.forEach((f, r) => COLUMNAS.forEach((col, i) => {
    if (f[col] !== null && f[col] !== undefined) celdas[columnaALetra(i + 1) + (7 + r)] = f[col];
  }));
  return {
    hojas: ["ResumenGeneral", "SALIDA_GESPER"],
    nombres: conNombres ? {
      SALIDA_GESPER_META: { hoja: "SALIDA_GESPER", ref: "A3:H4" },
      SALIDA_GESPER_DATOS: { hoja: "SALIDA_GESPER", ref: "A6:S" + (6 + filas.length) }
    } : {},
    celdas: { SALIDA_GESPER: celdas }
  };
}

/** GESPER tal como lo devuelve excelService.leerCeldas, con los destinos de FILAS válidos. */
function gesper(cambios = {}) {
  const c = {
    B108: "Número de puntos de captación", D108: 5,
    B115: ETQ_INST, M2: 2, M115: 3, M116: "NO",
    B347: ETQ_REEMP.replace("  ", " "), M347: 0, M349: "NO",
    L2: 1, B178: "Reemplazo <=6", L178: 0, L182: "NO"
  };
  Object.assign(c, cambios);
  const celdas = {};
  Object.keys(c).forEach((k) => {
    const v = c[k];
    const esF = typeof v === "string" && v.charAt(0) === "=";
    celdas[k] = { valor: esF ? 0 : v, texto: String(esF ? 0 : v), formula: v };
  });
  return { hojaExiste: true, celdas };
}

function evaluar(opts) { return evaluarAnexo(interpretarSalida(libroAnexo(opts)), A01); }
function motivoDe(plan, id) { const o = plan.omitidos.filter((x) => x.fila.ID === id)[0]; return o ? o.motivo : null; }

/* ---------- estructura ---------- */

test("interpretarSalida lee metadatos y filas por nombre de columna", () => {
  const s = interpretarSalida(libroAnexo());
  assert.equal(s.meta.ANEXO_ID, "A01");
  assert.deepEqual(s.version, { mayor: 1, menor: 0 });
  assert.equal(s.filas.length, FILAS.length);
  assert.equal(s.filas[1].DESTINO_CELDA, "M115");
  assert.equal(s.filas[1].VALOR, 0);
  assert.equal(s.filas[1].filaAnexo, 8);
});

test("interpretarSalida encuentra la tabla por la fila «ID» si faltan los nombres definidos", () => {
  const s = interpretarSalida(libroAnexo({ conNombres: false }));
  assert.equal(s.filas.length, FILAS.length);
  assert.equal(s.meta.ESTADO_GLOBAL, "OK CON ADVERTENCIAS");
});

test("interpretarSalida acepta una versión menor nueva y rechaza una versión mayor distinta", () => {
  assert.doesNotThrow(() => interpretarSalida(libroAnexo({ meta: { SPEC_VERSION: "1.3" } })));
  assert.throws(() => interpretarSalida(libroAnexo({ meta: { SPEC_VERSION: "2.0" } })), /especificación 2\.0/);
  assert.throws(() => interpretarSalida(libroAnexo({ meta: { SPEC_VERSION: "" } })), /SPEC_VERSION/);
});

test("interpretarSalida rechaza un archivo sin la hoja SALIDA_GESPER", () => {
  const lib = libroAnexo();
  lib.hojas = ["ResumenGeneral"];
  assert.throws(() => interpretarSalida(lib), /no tiene la hoja SALIDA_GESPER/);
});

test("interpretarSalida rechaza columnas renombradas o reordenadas", () => {
  const enc = COLUMNAS.slice();
  enc[7] = "VALOR_FINAL";
  assert.throws(() => interpretarSalida(libroAnexo({ encabezados: enc })), /columna H se esperaba «VALOR»/);
});

/* ---------- clasificación ---------- */

test("evaluarAnexo: solo DATO en OK/ADVERTENCIA son candidatos; REVISAR se omite; CONTROL y NO_DISPONIBLE no se cargan", () => {
  const ev = evaluar();
  assert.deepEqual(ev.candidatos.map((f) => f.ID), ["A01.CAP.PUNTOS", "A01.CAP.INST.R14_60.Y02", "A01.SAL.REEMP.R14_60.Y02"]);
  assert.equal(ev.datos.length, 4);
  assert.match(ev.omitidos[0].motivo, /REVISAR/);
  assert.equal(ev.controles.length, 1);
  assert.equal(ev.noDisponibles.length, 1);
  assert.equal(ev.bloqueo, null);
});

test("evaluarAnexo rechaza el archivo de otro anexo", () => {
  assert.throws(() => evaluar({ meta: { ANEXO_ID: "A02" } }), /anexo A02, no al A01/);
});

test("evaluarAnexo bloquea la carga si ESTADO_GLOBAL es REVISAR", () => {
  assert.match(evaluar({ meta: { ESTADO_GLOBAL: "REVISAR" } }).bloqueo, /REVISAR/);
});

test("un CONTROL en REVISAR invalida solo los datos de su sitio y concepto", () => {
  const filas = FILAS.map((f) => (f.ID === "A01.CAP.CTRL_INST" ? Object.assign({}, f, { ESTADO: "REVISAR" }) : f));
  const ev = evaluar({ filas });
  assert.deepEqual(ev.candidatos.map((f) => f.ID), ["A01.CAP.PUNTOS", "A01.SAL.REEMP.R14_60.Y02"]);
  assert.ok(ev.omitidos.some((o) => o.fila.ID === "A01.CAP.INST.R14_60.Y02" && /CTRL_INST/.test(o.motivo)));
  assert.equal(ev.controlesRevisar.length, 1);
});

/* ---------- validación de destinos ---------- */

test("direccionesAInspeccionar pide etiqueta, destino, año y bandera de horizonte", () => {
  const dirs = direccionesAInspeccionar(evaluar(), A01).sort();
  assert.deepEqual(dirs, ["B108", "B115", "B347", "D108", "M115", "M116", "M2", "M347", "M349"].sort());
});

test("planificarCarga escribe los candidatos válidos, incluidos los ceros, y pide confirmar la advertencia", () => {
  const plan = planificarCarga(evaluar(), gesper(), A01);
  assert.deepEqual(plan.escrituras.map((e) => [e.celda, e.valor, e.cambia]), [["D108", 7, true], ["M115", 0, true], ["M347", 0, false]]);
  assert.equal(plan.escrituras[1].actual, 3, "un 0 del anexo reemplaza el 3 que tenía GESPER");
  assert.equal(plan.avisos.length, 1);
  assert.equal(plan.avisos[0].tipo, "ADVERTENCIA");
  assert.equal(plan.requiereConfirmacion, true);
  assert.equal(plan.omitidos.length, 1, "solo el REVISAR");
});

test("sin advertencias y con ESTADO_GLOBAL OK no hace falta confirmar", () => {
  const filas = FILAS.filter((f) => f.ESTADO !== "ADVERTENCIA");
  const plan = planificarCarga(evaluar({ filas, meta: { ESTADO_GLOBAL: "OK" } }), gesper(), A01);
  assert.equal(plan.requiereConfirmacion, false);
});

test("se omite y reporta un destino cuya etiqueta en la columna B no coincide", () => {
  const plan = planificarCarga(evaluar(), gesper({ B115: "Otra fila que se movió aquí" }), A01);
  assert.match(motivoDe(plan, "A01.CAP.INST.R14_60.Y02"), /etiqueta de GESPER en B115/);
  assert.ok(!plan.escrituras.some((e) => e.celda === "M115"));
});

test("se omite un destino cuya columna no es el año ANIO_IDX en la fila 2", () => {
  const plan = planificarCarga(evaluar(), gesper({ M2: 3 }), A01);
  assert.match(motivoDe(plan, "A01.CAP.INST.R14_60.Y02"), /columna M es el año 3/);
});

test("se omite un destino que tiene fórmula", () => {
  const plan = planificarCarga(evaluar(), gesper({ D108: "=SUM(D1:D2)" }), A01);
  assert.match(motivoDe(plan, "A01.CAP.PUNTOS"), /fórmula/);
});

test("nunca se escribe en una fila de banderas «¿contemplado?», aunque la etiqueta coincida", () => {
  const filas = [fila({ ID: "A01.CAP.X.Y02", TIPO_FILA: "DATO", ANIO_IDX: 2, VALOR: 1, ESTADO: "OK", DESTINO_CELDA: "M116", DESTINO_ETIQUETA: "NO" })];
  const plan = planificarCarga(evaluar({ filas }), gesper({ B116: "NO" }), A01);
  assert.equal(plan.escrituras.length, 0);
  assert.match(motivoDe(plan, "A01.CAP.X.Y02"), /contemplado/);
});

test("se omite un destino en otra hoja y un VALOR que no es número", () => {
  const filas = [
    fila({ ID: "A01.CAP.OTRA", TIPO_FILA: "DATO", VALOR: 1, ESTADO: "OK", DESTINO_HOJA: "EV PROG CCIALES", DESTINO_CELDA: "D108", DESTINO_ETIQUETA: "x" }),
    fila({ ID: "A01.CAP.TXT", TIPO_FILA: "DATO", VALOR: "siete", ESTADO: "OK", DESTINO_CELDA: "D108", DESTINO_ETIQUETA: "Número de puntos de captación" })
  ];
  const plan = planificarCarga(evaluar({ filas }), gesper(), A01);
  assert.match(motivoDe(plan, "A01.CAP.OTRA"), /no EV PROG TÉCNICOS/);
  assert.match(motivoDe(plan, "A01.CAP.TXT"), /no es un número/);
});

test("dos filas al mismo destino: se escribe la primera y se omite la segunda", () => {
  const filas = [FILAS[0], Object.assign({}, FILAS[0], { ID: "A01.CAP.PUNTOS_BIS" })];
  const plan = planificarCarga(evaluar({ filas }), gesper(), A01);
  assert.equal(plan.escrituras.length, 1);
  assert.match(motivoDe(plan, "A01.CAP.PUNTOS_BIS"), /ya escribe en D108/);
});

test("un valor distinto de 0 en un año N/A de la bandera de horizonte genera un aviso", () => {
  const filas = [Object.assign({}, FILAS[1], { VALOR: 2 })];
  const plan = planificarCarga(evaluar({ filas, meta: { ESTADO_GLOBAL: "OK" } }), gesper({ M116: "N/A" }), A01);
  assert.equal(plan.escrituras.length, 1, "se carga igual: el aviso no corrige nada");
  assert.equal(plan.avisos[0].tipo, "HORIZONTE");
  assert.equal(plan.requiereConfirmacion, true);
  const conCero = planificarCarga(evaluar({ meta: { ESTADO_GLOBAL: "OK" }, filas: [FILAS[1]] }), gesper({ M116: "N/A" }), A01);
  assert.equal(conCero.avisos.length, 0, "un 0 en un año N/A no cambia nada en GESPER");
});

test("si el libro abierto no tiene la hoja destino, la carga se bloquea", () => {
  const plan = planificarCarga(evaluar(), { hojaExiste: false, celdas: {} }, A01);
  assert.match(plan.bloqueo, /no tiene la hoja EV PROG TÉCNICOS/);
  assert.equal(plan.escrituras.length, 0);
});

test("dependeDecision y normalizarTexto", () => {
  assert.equal(dependeDecision({ DEPENDE_DECISION: "SI – años escritos a mano" }), true);
  assert.equal(dependeDecision({ DEPENDE_DECISION: "SÍ - filtro" }), true);
  assert.equal(dependeDecision({ DEPENDE_DECISION: "NO" }), false);
  assert.equal(normalizarTexto('  Reemplazo  14" hasta 60" '), 'Reemplazo 14" hasta 60"');
  assert.equal(normalizarTexto("TÉCNICOS"), "TÉCNICOS");
});

/* ---------- repositorio: escritura por lotes ---------- */

function instalarGesper(valores) {
  const stats = { runCalls: 0, calculateCalls: 0 };
  const rangos = {};
  const rango = (addr) => {
    if (!rangos[addr]) {
      const v = valores[addr] === undefined ? "" : valores[addr];
      rangos[addr] = { values: [[v]], text: [[String(v)]], formulas: [[v]], address: EVT + "!" + addr, load() {} };
    }
    return rangos[addr];
  };
  const hoja = { name: EVT, load() {}, getRange: rango };
  global.Excel = {
    CalculationType: { full: "full" },
    run: async (cb) => {
      stats.runCalls++;
      return cb({
        workbook: {
          worksheets: {
            getItemOrNullObject: (n) => (n === EVT ? hoja : { isNullObject: true, load() {} }),
            getItem: (n) => { if (n !== EVT) throw new Error("no existe " + n); return hoja; }
          },
          application: { calculate() { stats.calculateCalls++; } }
        },
        sync: async () => {}
      });
    }
  };
  return { stats, rangos };
}

test("ejecutarCarga revalida contra GESPER y escribe todo en un lote con un solo recálculo", async () => {
  const valores = {};
  Object.entries(gesper().celdas).forEach(([k, c]) => { valores[k] = c.valor; });
  const { stats, rangos } = instalarGesper(valores);

  const r = await ejecutarCarga(evaluar(), anexoPorId("A01"));
  assert.equal(r.escritos, 3);
  assert.equal(r.cambiados, 2);
  assert.deepEqual(rangos.D108.values, [[7]]);
  assert.deepEqual(rangos.M115.values, [[0]]);
  assert.equal(stats.runCalls, 2, "una lectura de validación + un lote de escritura");
  assert.equal(stats.calculateCalls, 1, "un solo recálculo para todo el anexo");
});

test("ejecutarCarga no escribe nada si la carga está bloqueada", async () => {
  const { stats } = instalarGesper({});
  await assert.rejects(() => ejecutarCarga(evaluar({ meta: { ESTADO_GLOBAL: "REVISAR" } }), A01), /REVISAR/);
  assert.equal(stats.calculateCalls, 0);
});
