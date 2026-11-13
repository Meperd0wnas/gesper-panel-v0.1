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
  COLUMNAS, interpretarSalida, evaluarAnexo, direccionesAInspeccionar, planificarCarga, dependeDecision, normalizarTexto,
  seleccionInicial, grupoDe, coinciden
} from "../src/domain/anexos/salidaGesper.js";
import { ejecutarCarga, prepararCarga, anexoPorId } from "../src/domain/repositories/anexosRepository.js";
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
      SALIDA_GESPER_META: { hoja: "SALIDA_GESPER", ref: "A3:" + columnaALetra(Math.max(8, Object.keys(m).length)) + "4" },
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
  return inspeccion(Object.assign(c, cambios));
}

/**
 * Celda simulada de GESPER: un valor ("texto", 3), una fórmula sin resultado
 * que importe ("=SUM(...)", vale 0) o una fórmula con su resultado ({ f, v }).
 */
function celdaSimulada(v) {
  if (v && typeof v === "object") return { valor: v.v, texto: String(v.v), formula: v.f };
  const esF = typeof v === "string" && v.charAt(0) === "=";
  return { valor: esF ? 0 : v, texto: String(esF ? 0 : v), formula: v };
}

/** { "B115": "texto", "M115": 3, "L479": "=SUM(...)" } -> forma de excelService.leerCeldas. */
function inspeccion(c) {
  const celdas = {};
  Object.keys(c).forEach((k) => { celdas[k] = celdaSimulada(c[k]); });
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

test("direccionesAInspeccionar pide etiqueta, destino, año, año calendario y bandera de horizonte", () => {
  const dirs = direccionesAInspeccionar(evaluar(), A01).sort();
  assert.deepEqual(dirs, ["B108", "B115", "B347", "D108", "M115", "M116", "M2", "M3", "M347", "M349"].sort());
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

test("sin advertencias, con ESTADO_GLOBAL OK y sin valores que cambien, no hace falta confirmar", () => {
  const filas = FILAS.filter((f) => f.ESTADO !== "ADVERTENCIA");
  const plan = planificarCarga(evaluar({ filas, meta: { ESTADO_GLOBAL: "OK" } }), gesper({ D108: 7, M115: 0 }), A01);
  assert.equal(plan.cambios, 0);
  assert.equal(plan.requiereConfirmacion, false);
});

test("un valor nuevo distinto del actual de GESPER exige confirmación, aunque no haya advertencias", () => {
  const filas = FILAS.filter((f) => f.ESTADO !== "ADVERTENCIA");
  const plan = planificarCarga(evaluar({ filas, meta: { ESTADO_GLOBAL: "OK" } }), gesper({ D108: 7 }), A01);
  assert.equal(plan.avisos.length, 0);
  assert.equal(plan.cambios, 1, "M115: 3 -> 0");
  assert.equal(plan.requiereConfirmacion, true);
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
      const c = celdaSimulada(valores[addr] === undefined ? "" : valores[addr]);
      rangos[addr] = { values: [[c.valor]], text: [[c.texto]], formulas: [[c.formula]], address: EVT + "!" + addr, load() {} };
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

  const ev = evaluar();
  const vista = await prepararCarga(ev, anexoPorId("A01")); // lo que el usuario ve y confirma
  const r = await ejecutarCarga(ev, anexoPorId("A01"), vista);
  assert.equal(r.escritos, 3);
  assert.equal(r.cambiados, 2);
  assert.deepEqual(rangos.D108.values, [[7]]);
  assert.deepEqual(rangos.M115.values, [[0]]);
  assert.equal(stats.runCalls, 3, "vista previa + revalidación + un lote de escritura");
  assert.equal(stats.calculateCalls, 1, "un solo recálculo para todo el anexo");
});

test("ejecutarCarga no escribe nada si la carga está bloqueada", async () => {
  const { stats } = instalarGesper({});
  await assert.rejects(() => ejecutarCarga(evaluar({ meta: { ESTADO_GLOBAL: "REVISAR" } }), A01), /REVISAR/);
  assert.equal(stats.calculateCalls, 0);
});

/* ---------- Anexo 2 — Sectorización (P2) ---------- */
/*
 * Fixture con la estructura de Especificacion_SALIDA_GESPER_A02_v1 (§4, §7,
 * §11): 70 DATO, 13 CONTROL (6 de referencia con destino en L479:Q479) y 9
 * NO_DISPONIBLE. GESPER simulado: mismos valores que el anexo salvo L465 = 10
 * (hallazgo H-03), y fórmulas en L479:Q479.
 */
const A02 = anexoPorId("A02");
const COL_ANIO = "KLMNOPQRSTU"; // columnas de los años 0..10 en GESPER
const RANGOS_A02 = ["R6", "R8_12", "R14_60"];
const SITIO_A02 = { SEC: "2.1.1 Sectores", TAN: "2.1.2 Tanques", TEL: "2.1.3 Telemetría", REE: "2.1.7 Reemplazo", VAL: "2.2 Válvulas", AIS: "2.3 Aislamiento", GEN: "General" };
const NO_CERO_A02 = {
  "SEC.MACRO_EXIST": 31, "SEC.MACRO_FALT": 40,
  "SEC.INST.R6.Y01": 5, "SEC.INST.R6.Y02": 4, "SEC.INST.R6.Y03": 4, "SEC.INST.R6.Y05": 1,
  "SEC.INST.R8_12.Y01": 1, "SEC.INST.R8_12.Y02": 4, "SEC.INST.R8_12.Y03": 5, "SEC.INST.R8_12.Y04": 9, "SEC.INST.R8_12.Y05": 6,
  "SEC.INST.R14_60.Y04": 1,
  "TAN.PUNTOS": 25, "TAN.MACRO_EXIST": 12, "TAN.INST.R6.Y02": 1,
  "TAN.INST.R8_12.Y01": 1, "TAN.INST.R8_12.Y02": 1, "TAN.INST.R8_12.Y03": 1, "TAN.INST.R8_12.Y04": 2,
  "TAN.INST.R14_60.Y01": 3, "TAN.INST.R14_60.Y02": 2, "TAN.INST.R14_60.Y03": 1, "TAN.INST.R14_60.Y04": 1,
  "TEL.REQ.Y01": 7, "TEL.REQ.Y02": 11, "TEL.REQ.Y03": 10, "TEL.REQ.Y04": 15, "TEL.REQ.Y05": 10,
  "REE.REEMP_TOT": 12, "REE.REEMP.R8_12.Y01": 2, "REE.REEMP.R8_12.Y06": 3,
  "REE.REEMP.R14_60.Y01": 2, "REE.REEMP.R14_60.Y02": 3, "REE.REEMP.R14_60.Y03": 1, "REE.REEMP.R14_60.Y06": 1,
  "VAL.FALT": 262, "VAL.INST.Y02": 59, "VAL.INST.Y03": 54, "VAL.INST.Y04": 137, "VAL.INST.Y05": 12,
  "AIS.SECTORES": 37
};
const ETQ_CONTEMPLADO = "¿Dentro de los proyectos de la empresa ya se encuentra contemplado este subprograma?";
function etiquetaA02(celda) { return "Rótulo de la fila " + celda.replace(/^[A-Z]+/, ""); }

function filasA02() {
  const filas = [];
  const dato = (clave, celda, anio) => filas.push(fila({
    ID: "A02." + clave, TIPO_FILA: "DATO", PROGRAMA: "P2", SITIO: SITIO_A02[clave.split(".")[0]], CONCEPTO: "Concepto " + clave,
    ANIO_IDX: anio === undefined ? null : anio, VALOR: NO_CERO_A02[clave] || 0, ESTADO: "OK",
    DESTINO_CELDA: celda, DESTINO_ETIQUETA: etiquetaA02(celda)
  }));
  const anual = (base, filaG, desde, hasta) => { for (let y = desde; y <= hasta; y++) dato(base + ".Y0" + y, COL_ANIO[y] + filaG, y); };
  dato("SEC.MACRO_EXIST", "D363"); dato("SEC.MACRO_FALT", "D364");
  RANGOS_A02.forEach((r, i) => anual("SEC.INST." + r, 366 + i, 0, 5));
  dato("TAN.PUNTOS", "D380"); dato("TAN.MACRO_EXIST", "D381");
  RANGOS_A02.forEach((r, i) => anual("TAN.INST." + r, 384 + i, 0, 4));
  anual("TEL.REQ", 399, 0, 5);
  dato("REE.REEMP_TOT", "K448");
  RANGOS_A02.forEach((r, i) => anual("REE.REEMP." + r, 449 + i, 1, 6));
  dato("VAL.FALT", "D463");
  anual("VAL.INST", 465, 0, 5);
  dato("AIS.SECTORES", "D477");
  const control = (o) => filas.push(fila(Object.assign({ TIPO_FILA: "CONTROL", PROGRAMA: "P2", VALOR: 1, VALOR_CONTROL: 1, ESTADO: "OK", DESTINO_HOJA: null }, o)));
  ["SEC.CTRL_PUNTOS", "SEC.CTRL_INST", "TAN.CTRL_INST", "TEL.CTRL_REQ", "REE.CTRL_REEMP", "VAL.CTRL_INST", "AIS.CTRL_VERIF"]
    .forEach((k) => control({ ID: "A02." + k, SITIO: SITIO_A02[k.split(".")[0]] }));
  [5, 8, 10, 6, 5, 3].forEach((v, i) => control({
    ID: "A02.AIS.VERIF.Y0" + (i + 1), SITIO: SITIO_A02.AIS, ANIO_IDX: i + 1, VALOR: v, VALOR_CONTROL: v,
    DESTINO_HOJA: EVT, DESTINO_CELDA: COL_ANIO[i + 1] + "479", DESTINO_ETIQUETA: etiquetaA02("L479")
  }));
  ["CONTEMPLADO", "HORIZONTE", "TEL_NUEVOS", "TEL_REPARTO", "ALMACEN_LOCAL", "TEL_EXIST", "AIS_REPARTO", "VAL_TAPAS", "DERIVADOS"]
    .forEach((k) => filas.push(fila({ ID: "A02.GEN." + k, TIPO_FILA: "NO_DISPONIBLE", PROGRAMA: "P2", SITIO: "General", VALOR: null, ESTADO: "N/A", DESTINO_HOJA: null })));
  return filas;
}

/** Celdas de EV PROG TÉCNICOS para el A02: fila 2 (0..10), fila 3 (2025..2035), banderas en NO, destinos = anexo salvo L465. */
function gesperA02(cambios = {}) {
  const c = {};
  for (let y = 0; y <= 10; y++) { c[COL_ANIO[y] + "2"] = y; c[COL_ANIO[y] + "3"] = 2025 + y; }
  A02.filasBandera.forEach((r) => { c["B" + r] = ETQ_CONTEMPLADO; for (let y = 0; y <= 10; y++) c[COL_ANIO[y] + r] = "NO"; });
  filasA02().forEach((f) => { if (f.DESTINO_CELDA) { c["B" + f.DESTINO_CELDA.replace(/^[A-Z]+/, "")] = f.DESTINO_ETIQUETA; c[f.DESTINO_CELDA] = f.VALOR; } });
  for (let y = 1; y <= 6; y++) c[COL_ANIO[y] + "479"] = "=SUM(" + COL_ANIO[y] + "480:" + COL_ANIO[y] + "482)";
  c.L465 = 10; // H-03: GESPER tiene 10 válvulas en el año 1; el anexo trae 0
  c.B447 = "Tiempo estimado para Reemplazo"; c.D447 = 6; c.D478 = 6;
  return Object.assign(c, cambios);
}

function libroA02({ meta = {}, filas = filasA02() } = {}) {
  return libroAnexo({ filas, meta: Object.assign({ ANEXO_ID: "A02", PROGRAMAS: "P2", ESTADO_GLOBAL: "OK" }, meta) });
}
function evaluarA02(opts) { return evaluarAnexo(interpretarSalida(libroA02(opts)), A02); }
function planA02(ev, cambios) { return planificarCarga(ev, inspeccion(gesperA02(cambios)), A02); }

test("A02: reconoce el anexo A02 y el programa P2, y rechaza el archivo en la opción equivocada", () => {
  const ev = evaluarA02();
  assert.equal(ev.meta.ANEXO_ID, "A02");
  assert.equal(ev.estadoGlobal, "OK");
  assert.throws(() => evaluarAnexo(interpretarSalida(libroA02()), A01), /anexo A02, no al A01/);
  assert.throws(() => evaluarA02({ meta: { PROGRAMAS: "P3" } }), /programas «P3»; para el A02 se esperaba P2/);
});

test("A02: acepta el contrato 1.x y rechaza una versión mayor distinta", () => {
  assert.doesNotThrow(() => evaluarA02({ meta: { SPEC_VERSION: "1.1" } }));
  assert.throws(() => evaluarA02({ meta: { SPEC_VERSION: "2.0" } }), /especificación 2\.0/);
});

test("A02: lee los 70 datos, nunca escribe las 13 filas CONTROL y deja las 9 NO_DISPONIBLE como pendientes", () => {
  const ev = evaluarA02();
  assert.equal(ev.datos.length, 70);
  assert.equal(ev.candidatos.length, 70);
  assert.equal(ev.controles.length, 13);
  assert.equal(ev.noDisponibles.length, 9);
  const plan = planA02(ev);
  assert.equal(plan.escrituras.length, 70);
  assert.equal(plan.omitidos.length, 0);
  assert.ok(plan.escrituras.every((e) => e.fila.TIPO_FILA === "DATO"));
  assert.ok(!plan.escrituras.some((e) => /479$/.test(e.celda)), "las referencias de aislamiento (L479:Q479) no se escriben");
  assert.ok(!direccionesAInspeccionar(ev, A02).includes("B479"), "no se validan como destino (sin etiqueta ni año)");
  assert.deepEqual(plan.referencias.map((r) => [r.celda, r.formula]),
    ["L479", "M479", "N479", "O479", "P479", "Q479"].map((c) => [c, true]), "solo se leen, como referencia");
});

test("A02: transfiere los ceros y detecta la única diferencia con GESPER (L465: 10 → 0), que exige confirmación", () => {
  const plan = planA02(evaluarA02());
  assert.equal(plan.escrituras.filter((e) => e.valor === 0).length, 29);
  assert.equal(plan.cambios, 1);
  const cambio = plan.escrituras.filter((e) => e.cambia)[0];
  assert.deepEqual([cambio.fila.ID, cambio.celda, cambio.actual, cambio.valor, cambio.anio, cambio.anioCalendario],
    ["A02.VAL.INST.Y01", "L465", 10, 0, 1, 2026]);
  assert.equal(plan.avisos.length, 0);
  assert.equal(plan.requiereConfirmacion, true, "sin advertencias, la confirmación la exige el cambio de valor");
});

test("A02: omite los destinos cuya etiqueta en B o cuyo año en la fila 2 no coinciden", () => {
  const ev = evaluarA02();
  const etq = planA02(ev, { B465: "Fila movida" });
  assert.equal(etq.omitidos.length, 6);
  assert.ok(etq.omitidos.every((o) => /^A02\.VAL\.INST\./.test(o.fila.ID) && /etiqueta de GESPER en B465/.test(o.motivo)));

  const anio = planA02(ev, { M2: 5 });
  // Año 2 (columna M): 3 de sectores, 3 de tanques, 1 de telemetría, 3 de reemplazo, 1 de válvulas.
  assert.equal(anio.omitidos.length, 11);
  assert.ok(anio.omitidos.every((o) => /columna M es el año 5/.test(o.motivo)));
});

test("A02: no escribe sobre fórmulas, filas «¿contemplado?» ni celdas de horizonte (D447, D478)", () => {
  const formula = planA02(evaluarA02(), { D463: "=SUM(K465:P465)" });
  assert.match(motivoDe(formula, "A02.VAL.FALT"), /fórmula/);

  const filas = filasA02().concat([
    fila({ ID: "A02.VAL.BANDERA.Y02", TIPO_FILA: "DATO", PROGRAMA: "P2", ANIO_IDX: 2, VALOR: 1, ESTADO: "OK", DESTINO_CELDA: "M467", DESTINO_ETIQUETA: ETQ_CONTEMPLADO }),
    fila({ ID: "A02.REE.TIEMPO", TIPO_FILA: "DATO", PROGRAMA: "P2", VALOR: 6, ESTADO: "OK", DESTINO_CELDA: "D447", DESTINO_ETIQUETA: "Tiempo estimado para Reemplazo" })
  ]);
  const plan = planA02(evaluarA02({ filas }));
  assert.match(motivoDe(plan, "A02.VAL.BANDERA.Y02"), /contemplado/);
  assert.match(motivoDe(plan, "A02.REE.TIEMPO"), /D447 es una celda de horizonte/);
  assert.equal(plan.escrituras.length, 70);
});

test("A02: un CONTROL en REVISAR invalida todos los datos de su grupo (spec A02 §6)", () => {
  const conRevisar = (id) => filasA02().map((f) => (f.ID === id ? Object.assign({}, f, { ESTADO: "REVISAR" }) : f));
  const val = evaluarA02({ filas: conRevisar("A02.VAL.CTRL_INST") });
  assert.deepEqual(val.omitidos.map((o) => o.fila.ID).sort(),
    ["A02.VAL.FALT", "A02.VAL.INST.Y00", "A02.VAL.INST.Y01", "A02.VAL.INST.Y02", "A02.VAL.INST.Y03", "A02.VAL.INST.Y04", "A02.VAL.INST.Y05"]);
  assert.equal(val.candidatos.length, 63);
  const ais = evaluarA02({ filas: conRevisar("A02.AIS.VERIF.Y03") });
  assert.deepEqual(ais.omitidos.map((o) => o.fila.ID), ["A02.AIS.SECTORES"]);
});

test("A02: no carga filas DATO con ID de otro anexo o de otro programa", () => {
  const filas = filasA02().concat([
    fila({ ID: "A01.CAP.PUNTOS", TIPO_FILA: "DATO", PROGRAMA: "P2", VALOR: 7, ESTADO: "OK", DESTINO_CELDA: "D108", DESTINO_ETIQUETA: "x" }),
    fila({ ID: "A02.SEC.OTRO", TIPO_FILA: "DATO", PROGRAMA: "P1", VALOR: 7, ESTADO: "OK", DESTINO_CELDA: "D363", DESTINO_ETIQUETA: "x" })
  ]);
  const ev = evaluarA02({ filas });
  assert.match(ev.omitidos.filter((o) => o.fila.ID === "A01.CAP.PUNTOS")[0].motivo, /no empieza por A02\./);
  assert.match(ev.omitidos.filter((o) => o.fila.ID === "A02.SEC.OTRO")[0].motivo, /programa «P1»/);
  assert.equal(ev.candidatos.length, 70);
});

test("A02: un valor distinto de 0 en un año N/A de la bandera de su bloque genera aviso (telemetría → fila 404)", () => {
  const plan = planA02(evaluarA02(), { P404: "N/A" });
  assert.deepEqual(plan.avisos.map((a) => a.tipo + ":" + a.fila.ID), ["HORIZONTE:A02.TEL.REQ.Y05"]);
});

test("A02: ESTADO_GLOBAL REVISAR bloquea la carga y no escribe nada", async () => {
  const { stats } = instalarGesper(gesperA02());
  const ev = evaluarA02({ meta: { ESTADO_GLOBAL: "REVISAR" } });
  assert.match(planA02(ev).bloqueo, /REVISAR/);
  await assert.rejects(() => ejecutarCarga(ev, A02, null), /REVISAR/);
  assert.equal(stats.calculateCalls, 0);
});

test("A02: sin confirmación (el usuario cancela o no marca la casilla) no se escribe nada", async () => {
  const { stats, rangos } = instalarGesper(gesperA02());
  await assert.rejects(() => ejecutarCarga(evaluarA02(), A02, null), /necesita tu confirmación/);
  assert.deepEqual(rangos.L465.values, [[10]], "L465 conserva el 10");
  assert.equal(stats.calculateCalls, 0);
});

test("A02: con la confirmación del plan visto escribe los 70 valores, incluido el 0 en L465, en un lote y un recálculo", async () => {
  const { stats, rangos } = instalarGesper(gesperA02());
  const ev = evaluarA02();
  const vista = await prepararCarga(ev, A02);
  const r = await ejecutarCarga(ev, A02, vista);
  assert.equal(r.escritos, 70);
  assert.equal(r.cambiados, 1);
  assert.deepEqual(rangos.L465.values, [[0]]);
  assert.deepEqual(rangos.D463.values, [[262]]);
  assert.deepEqual(rangos.L479.formulas, [["=SUM(L480:L482)"]], "la fila de fórmulas de aislamiento no se escribe");
  assert.deepEqual(rangos.L479.values, [[0]]);
  assert.ok(!r.escrituras.some((e) => /479$/.test(e.celda)));
  assert.equal(r.referencias.length, 6, "se releen después de recalcular, para compararlas");
  assert.equal(stats.calculateCalls, 1);
});

test("A02: si GESPER cambió entre la vista previa y la carga, la confirmación ya no vale y no se escribe nada", async () => {
  const { stats, rangos } = instalarGesper(gesperA02());
  const ev = evaluarA02();
  const vista = await prepararCarga(ev, A02);
  rangos.L465.values = [[12]]; rangos.L465.formulas = [[12]]; // alguien editó L465 después de ver la vista previa
  const err = await ejecutarCarga(ev, A02, vista).then(() => null, (e) => e);
  assert.match(err.message, /cambió desde la vista previa/);
  assert.equal(err.planActual.escrituras.filter((e) => e.cambia)[0].actual, 12, "devuelve la vista previa actualizada");
  assert.deepEqual(rangos.L465.values, [[12]]);
  assert.equal(stats.calculateCalls, 0);
});

/* ---------- Anexo 3 — Presiones, daños y fugas (P3, P5, P6) ---------- */
/*
 * Fixture con los casos de Especificacion_SALIDA_GESPER_A03 v1.1 (§2, §4, §6,
 * §11): contrato 1.1 con ESTADO_GLOBAL REVISAR y GRUPOS_REVISAR en la columna
 * I, varios grupos independientes, un CONTROL de una serie en REVISAR
 * (DAN.CTRL_TAD), una referencia con destino en una fórmula de GESPER (D548),
 * un rótulo de GESPER que es fórmula (B697), ruido de punto flotante
 * (L509), un ajuste de decimales (L623) y diferencias reales (D501, D673).
 */
const A03 = anexoPorId("A03");
const SITIO_A03 = {
  ERP: "3.1 Instalación y mantenimiento de ERP's", FIJ: "3.3.1 Equipos de presión fijos", ROT: "3.3.2 Equipos de presión rotativos",
  MOD: "3.4 Modelo hidráulico", DAN: "5.1 Atención de daños", DET: "6.1 Detección de fugas", RMEN: "6.2.1 Reparación en redes menores"
};
const PROG_A03 = { ERP: "P3", FIJ: "P3", ROT: "P3", MOD: "P3", DAN: "P5", DET: "P6", RMEN: "P6" };
const ETQ_RMEN_VOL = "Volumen de agua recuperado por Reparación de fugas en redes menores";

function filasA03() {
  const g = (id) => id.split(".")[1];
  const dato = (id, celda, valor, o = {}) => fila(Object.assign({
    ID: id, TIPO_FILA: "DATO", PROGRAMA: PROG_A03[g(id)], SITIO: SITIO_A03[g(id)], CONCEPTO: "Concepto " + id,
    ANIO_IDX: /\.Y(\d\d)$/.test(id) ? Number(id.slice(-2)) : null, VALOR: valor, ESTADO: "OK",
    DESTINO_CELDA: celda, DESTINO_ETIQUETA: celda === "M697" ? ETQ_RMEN_VOL : "Rótulo de la fila " + celda.replace(/^[A-Z]+/, "")
  }, o));
  const control = (id, o = {}) => fila(Object.assign({
    ID: id, TIPO_FILA: "CONTROL", PROGRAMA: PROG_A03[g(id)], SITIO: SITIO_A03[g(id)], VALOR: 1, VALOR_CONTROL: 1, ESTADO: "OK", DESTINO_HOJA: null
  }, o));
  return [
    dato("A03.ERP.FALT", "D497", 21),
    dato("A03.ERP.MANT_TOT", "D501", 14, { NOTA: "H-05" }),
    dato("A03.ERP.VOL_MANT.Y01", "L509", 92653.59999999999, { ESTADO: "ADVERTENCIA", VALOR_CONTROL: 92653.59999999999, NOTA: "Total vacío en la hoja fuente (H-12)." }),
    control("A03.ERP.CTRL_MANT"),
    control("A03.ERP.COBERT_VOL_MANT", { ESTADO: "ADVERTENCIA", VALOR: 2647562.6, VALOR_CONTROL: 2844796 }),
    dato("A03.FIJ.EXIST", "D530", 120),
    dato("A03.FIJ.CAJAS.Y01", "L533", 6),
    control("A03.FIJ.CTRL_CAJAS"),
    dato("A03.ROT.CAJAS_DISP", "D547", 53),
    control("A03.ROT.CAJAS_FALT", { VALOR: 16, VALOR_CONTROL: 16, DESTINO_HOJA: EVT, DESTINO_CELDA: "D548", DESTINO_ETIQUETA: "No. de cajas faltantes" }),
    dato("A03.MOD.CAL.Y01", "L623", 30582.84153118256),
    dato("A03.MOD.ACT.Y05", "P624", 33635.03, { ESTADO: "REVISAR", VALOR_CONTROL: 40215.62 }),
    control("A03.MOD.CTRL_ACT"),
    dato("A03.DAN.TAD.Y01", "L666", 132091.57),
    dato("A03.DAN.CAPAC.Y01", "L660", 1),
    control("A03.DAN.CTRL_TAD", { ESTADO: "REVISAR", VALOR: 132091.57, VALOR_CONTROL: 252358.16 }),
    dato("A03.DET.LONG_RED", "D673", 1532829.0556638443),
    dato("A03.RMEN.VOL.Y02", "M697", 181314.49),
    fila({ ID: "A03.GEN.CONTEMPLADO", TIPO_FILA: "NO_DISPONIBLE", PROGRAMA: "P3, P5, P6", SITIO: "General", VALOR: null, ESTADO: "N/A", DESTINO_HOJA: null }),
    fila({ ID: "A03.GEN.HORIZONTE", TIPO_FILA: "NO_DISPONIBLE", PROGRAMA: "P3, P5, P6", SITIO: "General", VALOR: null, ESTADO: "N/A", DESTINO_HOJA: null })
  ];
}

/** EV PROG TÉCNICOS para el A03: años en las filas 2 y 3, banderas en NO, destinos = anexo salvo las diferencias de la spec §6. */
function gesperA03(cambios = {}) {
  const c = {};
  for (let y = 0; y <= 10; y++) { c[COL_ANIO[y] + "2"] = y; c[COL_ANIO[y] + "3"] = 2025 + y; }
  A03.filasBandera.forEach((r) => { c["B" + r] = ETQ_CONTEMPLADO; for (let y = 0; y <= 10; y++) c[COL_ANIO[y] + r] = "NO"; });
  filasA03().forEach((f) => {
    if (f.TIPO_FILA === "DATO") { c["B" + f.DESTINO_CELDA.replace(/^[A-Z]+/, "")] = f.DESTINO_ETIQUETA; c[f.DESTINO_CELDA] = f.VALOR; }
  });
  c.B697 = { f: '=CONCATENATE("Volumen de agua recuperado por ",B688)', v: ETQ_RMEN_VOL }; // rótulo calculado (spec A03 §3, columna P)
  c.D548 = { f: "=D546-D547", v: 16 };
  c.D501 = 15;                       // H-05: GESPER 15, el anexo 14
  c.L509 = 92653.6;                  // mismo valor salvo el redondeo binario
  c.L623 = 30582.84;                 // GESPER guarda 2 decimales
  c.D673 = 1532829.061276;           // H-07: 0,0056 m de diferencia
  c.B545 = "Tiempo estimado para inversión total - Años de priorización"; c.D545 = 5;
  return Object.assign(c, cambios);
}

function libroA03({ meta = {}, filas = filasA03() } = {}) {
  return libroAnexo({
    filas,
    meta: Object.assign({
      SPEC_VERSION: "1.1", ANEXO_ID: "A03", PROGRAMAS: "P3, P5, P6", ESTADO_GLOBAL: "REVISAR",
      N_DATO: 12, N_DATO_OK: 11, N_ADVERTENCIA: 2, N_REVISAR: 2, GRUPOS_REVISAR: "MOD DAN" // GRUPOS_REVISAR en la columna I, como en el anexo
    }, meta)
  });
}
function evaluarA03(opts) { return evaluarAnexo(interpretarSalida(libroA03(opts)), A03); }
function planA03(ev, cambios) { return planificarCarga(ev, inspeccion(gesperA03(cambios)), A03); }
function celdasDe(plan) { return plan.escrituras.map((e) => e.celda); }

test("A03: reconoce el anexo, los programas P3, P5 y P6 y lee GRUPOS_REVISAR de la columna I", () => {
  const s = interpretarSalida(libroA03());
  assert.equal(s.meta.GRUPOS_REVISAR, "MOD DAN");
  assert.deepEqual(s.version, { mayor: 1, menor: 1 });
  const ev = evaluarA03();
  assert.equal(ev.datos.length, 12);
  assert.equal(ev.controles.length, 6);
  assert.equal(ev.noDisponibles.length, 2);
  assert.throws(() => evaluarA03({ meta: { PROGRAMAS: "P3, P5" } }), /para el A03 se esperaba P3, P5, P6/);
  assert.throws(() => evaluarAnexo(interpretarSalida(libroA03()), A02), /anexo A03, no al A02/);
  assert.equal(grupoDe("A03.ERP.INST.R6.Y02"), "ERP");
});

test("A03: con la spec 1.1, ESTADO_GLOBAL REVISAR bloquea solo las filas en REVISAR y deja cargar el resto", () => {
  const ev = evaluarA03();
  assert.equal(ev.bloqueo, null);
  assert.equal(ev.cargaParcial, true);
  assert.deepEqual(ev.gruposRevisar, ["MOD", "DAN"]);
  assert.match(ev.omitidos.filter((o) => o.fila.ID === "A03.MOD.ACT.Y05")[0].motivo, /REVISAR/);
  const plan = planA03(ev);
  assert.equal(plan.bloqueo, null);
  assert.equal(plan.requiereConfirmacion, true, "una carga parcial siempre se confirma");
  assert.ok(celdasDe(plan).includes("L623"), "la calibración del modelo (MOD) sí se carga");
  assert.ok(!celdasDe(plan).includes("P624"), "la actualización en REVISAR no");
});

test("A03: el mismo REVISAR con la spec 1.0, o sin grupos identificables, bloquea toda la carga", async () => {
  assert.match(evaluarA03({ meta: { SPEC_VERSION: "1.0" } }).bloqueo, /REVISAR/);
  const sinGrupos = filasA03().map((f) => (f.ESTADO === "REVISAR" ? Object.assign({}, f, { ESTADO: "OK" }) : f));
  const ev = evaluarA03({ filas: sinGrupos, meta: { GRUPOS_REVISAR: "" } });
  assert.match(ev.bloqueo, /REVISAR/);
  const { stats } = instalarGesper(gesperA03());
  await assert.rejects(() => ejecutarCarga(ev, A03, null, ["FIJ"]), /REVISAR/);
  assert.equal(stats.calculateCalls, 0);
});

test("A03: un CONTROL de una serie en REVISAR invalida solo esa serie (DAN.TAD), no el resto del grupo (DAN.CAPAC)", () => {
  const ev = evaluarA03();
  assert.ok(ev.omitidos.some((o) => o.fila.ID === "A03.DAN.TAD.Y01" && /CTRL_TAD/.test(o.motivo)));
  assert.ok(ev.candidatos.some((f) => f.ID === "A03.DAN.CAPAC.Y01"));
  // Un control de cobertura (COBERT_) en REVISAR invalida su concepto igual que un CTRL_.
  const filas = filasA03().map((f) => (f.ID === "A03.ERP.COBERT_VOL_MANT" ? Object.assign({}, f, { ESTADO: "REVISAR" }) : f));
  const ev2 = evaluarA03({ filas });
  assert.ok(ev2.omitidos.some((o) => o.fila.ID === "A03.ERP.VOL_MANT.Y01" && /COBERT_VOL_MANT/.test(o.motivo)));
  assert.ok(ev2.candidatos.some((f) => f.ID === "A03.ERP.MANT_TOT"), "MANT_TOT no es de la serie VOL_MANT");
  assert.ok(ev2.gruposRevisar.includes("ERP"));
});

test("A03: los CONTROL con destino (D548) se leen como referencia y nunca se escriben", async () => {
  const ev = evaluarA03();
  assert.ok(direccionesAInspeccionar(ev, A03).includes("D548"));
  const plan = planA03(ev);
  assert.ok(!celdasDe(plan).includes("D548"));
  assert.deepEqual(plan.referencias.map((r) => [r.fila.ID, r.celda, r.gesper, r.formula, r.coincide]),
    [["A03.ROT.CAJAS_FALT", "D548", 16, true, true]]);
  assert.equal(plan.grupos.filter((g) => g.id === "ROT")[0].referencias.length, 1);

  const { rangos } = instalarGesper(gesperA03());
  const vista = await prepararCarga(ev, A03);
  const r = await ejecutarCarga(ev, A03, vista, ["ROT"]);
  assert.deepEqual(rangos.D548.formulas, [["=D546-D547"]]);
  assert.equal(r.referencias[0].despues, 16, "se relee después del recálculo");
});

test("A03: un rótulo de GESPER que es fórmula se compara por su valor calculado", () => {
  const plan = planA03(evaluarA03());
  assert.ok(celdasDe(plan).includes("M697"));
  const movida = planA03(evaluarA03(), { B697: { f: "=CONCATENATE(B1)", v: "Otra fila" } });
  assert.match(motivoDe(movida, "A03.RMEN.VOL.Y02"), /etiqueta de GESPER en B697 \(«Otra fila»\)/);
});

test("A03: el redondeo binario no es un cambio, un ajuste de decimales no pide confirmación y 0,0056 m sí es un cambio", () => {
  const plan = planA03(evaluarA03());
  const e = (celda) => plan.escrituras.filter((x) => x.celda === celda)[0];
  assert.deepEqual([e("L509").cambia, e("L509").ajuste], [false, false], "92653.6 contra 92653.59999999999");
  assert.deepEqual([e("L623").cambia, e("L623").ajuste], [false, true], "30582.84 contra 30582.8415: se escribe, sin confirmar");
  assert.equal(e("D673").cambia, true, "longitud de redes, H-07");
  assert.equal(e("D501").cambia, true, "15 -> 14, H-05");
  assert.equal(plan.cambios, 2);
  assert.equal(plan.ajustes, 1);
  assert.equal(coinciden(16, 16.004), true);
  assert.equal(coinciden(16, "16"), false);
});

test("A03: resume por grupo qué pide revisión, y la selección inicial deja marcados solo los grupos limpios", () => {
  const plan = planA03(evaluarA03());
  const g = Object.fromEntries(plan.grupos.map((x) => [x.id, x]));
  assert.deepEqual(plan.grupos.map((x) => x.id), ["ERP", "FIJ", "ROT", "MOD", "DAN", "DET", "RMEN"]);
  assert.equal(g.ERP.sitio, SITIO_A03.ERP);
  assert.equal(g.ERP.programa, "P3");
  assert.deepEqual([g.ERP.cambios, g.ERP.avisos.length, g.ERP.controles.length], [1, 1, 1], "MANT_TOT, la advertencia y el COBERT");
  assert.equal(g.MOD.revisar, true);
  assert.equal(g.MOD.omitidos.length, 1, "incompleto: ACT en REVISAR");
  assert.equal(g.DAN.omitidos.length, 1);
  assert.deepEqual(seleccionInicial(plan), { ERP: false, FIJ: true, ROT: true, MOD: false, DAN: false, DET: false, RMEN: true });
});

test("A03: ejecutarCarga escribe solo los grupos marcados, en un lote y con un solo recálculo", async () => {
  const { stats, rangos } = instalarGesper(gesperA03());
  const ev = evaluarA03();
  const vista = await prepararCarga(ev, A03);
  const r = await ejecutarCarga(ev, A03, vista, ["FIJ", "DET"]);
  assert.deepEqual(r.escrituras.map((e) => e.celda).sort(), ["D530", "D673", "L533"]);
  assert.equal(r.escritos, 3);
  assert.equal(r.cambiados, 1, "D673");
  assert.deepEqual(r.grupos, ["FIJ", "DET"]);
  assert.deepEqual(rangos.D673.values, [[1532829.0556638443]]);
  assert.deepEqual(rangos.D501.values, [[15]], "ERP no se marcó: GESPER conserva el 15");
  assert.equal(stats.calculateCalls, 1);
});

test("A03: sin ningún grupo marcado, o si GESPER cambió desde la vista previa, no se escribe nada", async () => {
  const { stats, rangos } = instalarGesper(gesperA03());
  const ev = evaluarA03();
  const vista = await prepararCarga(ev, A03);
  await assert.rejects(() => ejecutarCarga(ev, A03, vista, []), /No marcaste ningún grupo/);
  await assert.rejects(() => ejecutarCarga(ev, A03, null, ["FIJ"]), /necesita tu confirmación/);
  rangos.D530.values = [[119]]; // alguien editó GESPER después de ver la vista previa
  const err = await ejecutarCarga(ev, A03, vista, ["FIJ"]).then(() => null, (e) => e);
  assert.match(err.message, /cambió desde la vista previa/);
  assert.ok(err.planActual.grupos.filter((g) => g.id === "FIJ")[0].requiereConfirmacion, "ahora FIJ trae un cambio");
  assert.deepEqual(rangos.L533.values, [[6]]);
  assert.equal(stats.calculateCalls, 0);
});

test("A03: nunca escribe banderas ni celdas de horizonte, y avisa un año N/A aunque la bandera esté encima de la fila", () => {
  const filas = filasA03().concat([
    fila({ ID: "A03.ROT.HORIZ", TIPO_FILA: "DATO", PROGRAMA: "P3", VALOR: 5, ESTADO: "OK", DESTINO_CELDA: "D545",
      DESTINO_ETIQUETA: "Tiempo estimado para inversión total - Años de priorización" }),
    fila({ ID: "A03.ROT.BANDERA.Y01", TIPO_FILA: "DATO", PROGRAMA: "P3", ANIO_IDX: 1, VALOR: 1, ESTADO: "OK", DESTINO_CELDA: "L555", DESTINO_ETIQUETA: ETQ_CONTEMPLADO })
  ]);
  const plan = planA03(evaluarA03({ filas }), { L503: "N/A" });
  assert.match(motivoDe(plan, "A03.ROT.HORIZ"), /D545 es una celda de horizonte/);
  assert.match(motivoDe(plan, "A03.ROT.BANDERA.Y01"), /contemplado/);
  // L509 (volumen de mantenimiento de ERP's) va debajo de su bandera, la fila 503.
  assert.deepEqual(plan.avisos.filter((a) => a.tipo === "HORIZONTE").map((a) => a.fila.ID), ["A03.ERP.VOL_MANT.Y01"]);
});

/* ---------- Anexo 4: P4 y P6.3, contrato 1.1, carga por grupo (valores inventados) ---------- */

const A04 = anexoPorId("A04");
const SITIO_A04 = {
  IMP: "4.1 Pruebas de estanqueidad e impermeabilización de tanques", REB: "6.3.1 Instrumentación y control de nivel en tanques",
  SEN: "6.3.5 Reemplazo de sensor de nivel", GEN: "General"
};
const PROG_A04 = { IMP: "P4", REB: "P6", SEN: "P6", GEN: "P4, P6" };
const ETQ_IMP = "Lavado e impermeabilizado de tanques de almacenamiento de agua potable"; // filas 641 y 642
const ETQ_IMP_VOL = "Volumen de agua recuperado Pruebas de estanqueidad e impermeabilización de tanques";

function filasA04() {
  const g = (id) => id.split(".")[1];
  const etiqueta = (celda) => ({ 641: ETQ_IMP, 648: ETQ_IMP_VOL })[celda.replace(/^[A-Z]+/, "")] || "Rótulo de la fila " + celda.replace(/^[A-Z]+/, "");
  const dato = (id, celda, valor, o = {}) => fila(Object.assign({
    ID: id, TIPO_FILA: "DATO", PROGRAMA: PROG_A04[g(id)], SITIO: SITIO_A04[g(id)], CONCEPTO: "Concepto " + id,
    ANIO_IDX: /\.Y(\d\d)$/.test(id) ? Number(id.slice(-2)) : null, VALOR: valor, ESTADO: "OK", DESTINO_CELDA: celda, DESTINO_ETIQUETA: etiqueta(celda)
  }, o));
  const control = (id, o = {}) => fila(Object.assign({
    ID: id, TIPO_FILA: "CONTROL", PROGRAMA: PROG_A04[g(id)], SITIO: SITIO_A04[g(id)], VALOR: 1, VALOR_CONTROL: 1, ESTADO: "OK", DESTINO_HOJA: null
  }, o));
  return [
    dato("A04.IMP.AREA_TOT", "D639", 300.5, { ESTADO: "ADVERTENCIA", VALOR_CONTROL: 300.5, NOTA: "Incluye tanques marcados NA (H-02)." }),
    dato("A04.IMP.AREA.Y02", "M641", 120.25),
    dato("A04.IMP.VOL.Y02", "M648", 410.5),
    dato("A04.IMP.VOL.Y07", "R648", 0),
    control("A04.IMP.CTRL_VOL"),
    control("A04.IMP.COBERT_PRIOR", { ESTADO: "ADVERTENCIA", VALOR: 300.5, VALOR_CONTROL: 520 }),
    dato("A04.REB.PTOS_FALT", "D731", 6),
    dato("A04.REB.INSTAL.Y02", "M733", 2),
    dato("A04.REB.VOL.Y02", "M739", 150),
    control("A04.REB.CTRL_INSTAL"),
    control("A04.REB.TANQ_CON_CONTROL", { VALOR: 12, VALOR_CONTROL: 12, DESTINO_HOJA: EVT, DESTINO_CELDA: "D730",
      DESTINO_ETIQUETA: "No. de tanques con control de nivel", NOTA: "Fila de referencia: en GESPER es una fórmula (H-03)." }),
    dato("A04.SEN.EQ_EDAD", "D779", 7),
    dato("A04.SEN.REEMPL.Y02", "M780", 3),
    control("A04.SEN.CTRL_REEMPL"),
    dato("A04.GEN.TANQUES_TOT", "D638", 18, { DEPENDE_DECISION: "SI – qué registros cuentan como tanques" }),
    fila({ ID: "A04.GEN.CONTEMPLADO", TIPO_FILA: "NO_DISPONIBLE", PROGRAMA: "P4, P6", SITIO: "General", VALOR: null, ESTADO: "N/A", DESTINO_HOJA: null }),
    fila({ ID: "A04.GEN.HORIZONTE", TIPO_FILA: "NO_DISPONIBLE", PROGRAMA: "P4, P6", SITIO: "General", VALOR: null, ESTADO: "N/A", DESTINO_HOJA: null })
  ];
}

/** EV PROG TÉCNICOS para el A04: destinos = anexo salvo el volumen del año 7 (H-01) y el total de tanques (H-03). */
function gesperA04(cambios = {}) {
  const c = {};
  for (let y = 0; y <= 10; y++) { c[COL_ANIO[y] + "2"] = y; c[COL_ANIO[y] + "3"] = 2025 + y; }
  A04.filasBandera.forEach((r) => { c["B" + r] = ETQ_CONTEMPLADO; for (let y = 0; y <= 10; y++) c[COL_ANIO[y] + r] = "NO"; });
  filasA04().forEach((f) => {
    if (f.TIPO_FILA === "DATO") { c["B" + f.DESTINO_CELDA.replace(/^[A-Z]+/, "")] = f.DESTINO_ETIQUETA; c[f.DESTINO_CELDA] = f.VALOR; }
  });
  c.B642 = ETQ_IMP;                                                   // mismo rótulo que la 641 (spec A04 §3)
  c.M642 = "=IFERROR(HLOOKUP(M636-$D$637,$L$636:$U$642,6,FALSE),0)";  // repite la impermeabilización cada D637 años
  c.B648 = { f: '=CONCATENATE("Volumen de agua recuperado ",B635)', v: ETQ_IMP_VOL };
  c.D730 = { f: "=D729-D731", v: 13 };                                // D729 = TQTOTAL = D638
  c.R648 = 500;                                                       // GESPER tiene volumen en un año sin dato del anexo
  c.D638 = 19;                                                        // GESPER cuenta un tanque más
  c.B637 = "Frecuencia para impermeabilización de tanques"; c.D637 = 10;
  c.B732 = "Tiempo estimado para inversión total"; c.D732 = 5;
  c.B778 = "Tiempo estimado para Reemplazo"; c.D778 = 4;
  return Object.assign(c, cambios);
}

function libroA04({ meta = {}, filas = filasA04() } = {}) {
  return libroAnexo({
    filas,
    meta: Object.assign({
      SPEC_VERSION: "1.1", ANEXO_ID: "A04", PROGRAMAS: "P4, P6", ESTADO_GLOBAL: "OK CON ADVERTENCIAS",
      N_DATO: 10, N_DATO_OK: 9, N_ADVERTENCIA: 2, N_REVISAR: 0, GRUPOS_REVISAR: ""
    }, meta)
  });
}
function evaluarA04(opts) { return evaluarAnexo(interpretarSalida(libroA04(opts)), A04); }
function planA04(ev, cambios) { return planificarCarga(ev, inspeccion(gesperA04(cambios)), A04); }

test("A04: reconoce el anexo y los programas P4 y P6, y carga el total de tanques aunque su programa sea «P4, P6»", () => {
  const ev = evaluarA04();
  assert.equal(ev.datos.length, 10);
  assert.equal(ev.candidatos.length, 10);
  assert.equal(ev.controles.length, 5);
  assert.equal(ev.noDisponibles.length, 2);
  assert.deepEqual([ev.bloqueo, ev.cargaParcial], [null, false]);
  assert.ok(ev.candidatos.some((f) => f.ID === "A04.GEN.TANQUES_TOT"));
  assert.throws(() => evaluarAnexo(interpretarSalida(libroA04()), A03), /anexo A04, no al A03/);
  assert.throws(() => evaluarA04({ meta: { PROGRAMAS: "P4" } }), /para el A04 se esperaba P4, P6/);
  // Un programa ajeno al anexo se omite aunque venga junto a uno propio.
  const filas = filasA04().map((f) => (f.ID === "A04.GEN.TANQUES_TOT" ? Object.assign({}, f, { PROGRAMA: "P4, P5" }) : f));
  assert.match(evaluarA04({ filas }).omitidos.filter((o) => o.fila.ID === "A04.GEN.TANQUES_TOT")[0].motivo, /«P4, P5» no es de este anexo/);
});

test("A04: detecta el 0 que reemplaza un volumen de GESPER y el total de tanques, y deja marcados solo REB y SEN", () => {
  const plan = planA04(evaluarA04());
  assert.equal(plan.escrituras.length, 10);
  assert.deepEqual(plan.escrituras.filter((e) => e.cambia).map((e) => [e.celda, e.actual, e.valor]), [["R648", 500, 0], ["D638", 19, 18]]);
  assert.deepEqual(plan.avisos.map((a) => [a.tipo, a.fila.ID]), [["ADVERTENCIA", "A04.IMP.AREA_TOT"]]);
  assert.equal(plan.requiereConfirmacion, true);
  const g = Object.fromEntries(plan.grupos.map((x) => [x.id, x]));
  assert.deepEqual(plan.grupos.map((x) => x.id), ["IMP", "REB", "SEN", "GEN"]);
  assert.deepEqual([g.IMP.cambios, g.IMP.avisos.length, g.IMP.controles.length], [1, 1, 1], "R648, AREA_TOT y COBERT_PRIOR");
  assert.equal(g.GEN.programa, "P4, P6");
  assert.deepEqual(seleccionInicial(plan), { IMP: false, REB: true, SEN: true, GEN: false });
});

test("A04: la referencia D730 se lee y nunca se escribe; que no coincida no obliga a confirmar su grupo", async () => {
  const ev = evaluarA04();
  assert.ok(direccionesAInspeccionar(ev, A04).includes("D730"));
  const plan = planA04(ev);
  assert.deepEqual(plan.referencias.map((r) => [r.fila.ID, r.celda, r.gesper, r.formula, r.coincide]),
    [["A04.REB.TANQ_CON_CONTROL", "D730", 13, true, false]]);
  const reb = plan.grupos.filter((x) => x.id === "REB")[0];
  assert.deepEqual([reb.referencias.length, reb.requiereConfirmacion], [1, false], "REB no cambia nada en GESPER");

  const { stats, rangos } = instalarGesper(gesperA04());
  const vista = await prepararCarga(ev, A04);
  const r = await ejecutarCarga(ev, A04, vista, ["REB", "GEN"]);
  assert.deepEqual(rangos.D730.formulas, [["=D729-D731"]]);
  assert.deepEqual(rangos.D638.values, [[18]]);
  assert.deepEqual(rangos.R648.values, [[500]], "IMP no se marcó: GESPER conserva su volumen");
  assert.deepEqual(r.grupos, ["REB", "GEN"]);
  assert.equal(r.referencias[0].despues, 13, "se relee después del recálculo");
  assert.equal(stats.calculateCalls, 1);
});

test("A04: las filas 641 y 642 tienen el mismo rótulo; la 642 es una fórmula y nunca se escribe", () => {
  const filas = filasA04().concat([
    fila({ ID: "A04.IMP.REPITE.Y02", TIPO_FILA: "DATO", PROGRAMA: "P4", ANIO_IDX: 2, VALOR: 1, ESTADO: "OK", DESTINO_CELDA: "M642", DESTINO_ETIQUETA: ETQ_IMP })
  ]);
  const plan = planA04(evaluarA04({ filas }));
  assert.match(motivoDe(plan, "A04.IMP.REPITE.Y02"), /M642 tiene una fórmula/);
  assert.ok(celdasDe(plan).includes("M641"));
  assert.ok(celdasDe(plan).includes("M648"), "el rótulo de la fila 648 es una fórmula y se compara por su valor");
});

test("A04: nunca escribe la frecuencia ni los horizontes (D637, D732, D778), y avisa un año N/A del reemplazo de sensor", () => {
  const extra = [["A04.IMP.FREC", "D637"], ["A04.REB.HORIZ", "D732"], ["A04.SEN.HORIZ", "D778"]].map(([id, celda]) =>
    fila({ ID: id, TIPO_FILA: "DATO", PROGRAMA: PROG_A04[id.split(".")[1]], VALOR: 3, ESTADO: "OK", DESTINO_CELDA: celda,
      DESTINO_ETIQUETA: gesperA04()["B" + celda.slice(1)] }));
  const plan = planA04(evaluarA04({ filas: filasA04().concat(extra) }), { M781: "N/A", R643: "N/A" });
  for (const [id, celda] of [["A04.IMP.FREC", "D637"], ["A04.REB.HORIZ", "D732"], ["A04.SEN.HORIZ", "D778"]]) {
    assert.match(motivoDe(plan, id), new RegExp(celda + " es una celda de horizonte o frecuencia"));
  }
  // M780 va encima de su bandera (fila 781); el 0 de R648 en un año N/A no avisa.
  assert.deepEqual(plan.avisos.filter((a) => a.tipo === "HORIZONTE").map((a) => a.fila.ID), ["A04.SEN.REEMPL.Y02"]);
});

/* ---------- Anexo 7: P7, contrato 1.1, carga por grupo (valores inventados) ---------- */

const A07 = anexoPorId("A07");
const SITIO_A07 = { RED: "Reposición de redes", ACO: "Reposición de acometidas", GEN: "General" };
const ETQ_RED_VOL6 = 'Volumen recuperado por redes <6" '; // GESPER rotula 807 a 809 con un espacio al final (spec A07 §3)
const ETQ_HORIZ_A07 = "Tiempo estimado para inversión total - Años de priorización";

function filasA07() {
  const g = (id) => id.split(".")[1];
  const etiqueta = (celda) => ({ 807: ETQ_RED_VOL6 })[celda.replace(/^[A-Z]+/, "")] || "Rótulo de la fila " + celda.replace(/^[A-Z]+/, "");
  const dato = (id, celda, valor, o = {}) => fila(Object.assign({
    ID: id, TIPO_FILA: "DATO", PROGRAMA: "P7", SITIO: SITIO_A07[g(id)], CONCEPTO: "Concepto " + id,
    ANIO_IDX: /\.Y(\d\d)$/.test(id) ? Number(id.slice(-2)) : null, VALOR: valor, ESTADO: "OK", DESTINO_CELDA: celda, DESTINO_ETIQUETA: etiqueta(celda),
    DEPENDE_DECISION: "SI – año de priorización escrito a mano"
  }, o));
  const control = (id, o = {}) => fila(Object.assign({
    ID: id, TIPO_FILA: "CONTROL", PROGRAMA: "P7", SITIO: SITIO_A07[g(id)], VALOR: 1, VALOR_CONTROL: 1, ESTADO: "OK", DESTINO_HOJA: null
  }, o));
  return [
    dato("A07.RED.D6.Y02", "M798", 120.5),
    dato("A07.RED.D6.Y03", "N798", 80),
    dato("A07.RED.D14.Y03", "N800", 45.25),
    dato("A07.RED.V6.Y02", "M807", 9000),
    dato("A07.RED.V6.Y03", "N807", 4000),
    dato("A07.RED.V14.Y03", "N809", 700),
    control("A07.RED.LONG_TOT", { VALOR: 205.75, VALOR_CONTROL: 205.75, DESTINO_HOJA: EVT, DESTINO_CELDA: "D794",
      DESTINO_ETIQUETA: "Longitud total de redes a reponer", NOTA: "Fila de referencia: en GESPER es una fórmula redondeada (H-06)." }),
    control("A07.RED.CTRL_D6"),
    control("A07.RED.CTRL_V6"),
    control("A07.RED.CTRL_BLOQUE_LONG"),
    control("A07.RED.COBERT_LONG"),
    dato("A07.ACO.CANT_TOT", "D795", 9),
    dato("A07.ACO.CANT.Y02", "M816", 4),
    dato("A07.ACO.VOL.Y02", "M822", 1500),
    control("A07.ACO.CTRL_CANT"),
    control("A07.GEN.LONG_RED_SIST", { VALOR: 5000, VALOR_CONTROL: 5000, DESTINO_HOJA: EVT, DESTINO_CELDA: "D792",
      DESTINO_ETIQUETA: "Longitud total de redes en el sistema" }),
    fila({ ID: "A07.GEN.CONTEMPLADO", TIPO_FILA: "NO_DISPONIBLE", PROGRAMA: "P7", SITIO: "General", VALOR: null, ESTADO: "N/A", DESTINO_HOJA: null }),
    fila({ ID: "A07.GEN.HORIZONTE", TIPO_FILA: "NO_DISPONIBLE", PROGRAMA: "P7", SITIO: "General", VALOR: null, ESTADO: "N/A", DESTINO_HOJA: null })
  ];
}

/** EV PROG TÉCNICOS para el A07: destinos = anexo salvo N800, que GESPER guarda en km dentro de una fila en m (H-01). */
function gesperA07(cambios = {}) {
  const c = {};
  for (let y = 0; y <= 10; y++) { c[COL_ANIO[y] + "2"] = y; c[COL_ANIO[y] + "3"] = 2025 + y; }
  A07.filasBandera.forEach((r) => { c["B" + r] = ETQ_CONTEMPLADO; for (let y = 0; y <= 10; y++) c[COL_ANIO[y] + r] = y <= 6 ? "NO" : "N/A"; });
  filasA07().forEach((f) => {
    if (f.TIPO_FILA === "DATO") { c["B" + f.DESTINO_CELDA.replace(/^[A-Z]+/, "")] = f.DESTINO_ETIQUETA; c[f.DESTINO_CELDA] = f.VALOR; }
  });
  c.N800 = 0.04525;
  c.B794 = "Longitud total de redes a reponer"; c.D794 = { f: "=0.2*1000", v: 200 };
  c.B792 = "Longitud total de redes en el sistema"; c.D792 = { f: "=$D$673", v: 5000 };
  for (const r of [796, 815]) { c["B" + r] = ETQ_HORIZ_A07; c["D" + r] = { f: "=COUNT(L" + r + ":U" + r + ")", v: 6 }; c["K" + r] = 0; }
  return Object.assign(c, cambios);
}

function libroA07({ meta = {}, filas = filasA07() } = {}) {
  return libroAnexo({
    filas,
    meta: Object.assign({
      SPEC_VERSION: "1.1", ANEXO_ID: "A07", PROGRAMAS: "P7", ESTADO_GLOBAL: "OK",
      N_DATO: 9, N_DATO_OK: 9, N_ADVERTENCIA: 0, N_REVISAR: 0, GRUPOS_REVISAR: ""
    }, meta)
  });
}
function evaluarA07(opts) { return evaluarAnexo(interpretarSalida(libroA07(opts)), A07); }
function planA07(ev, cambios) { return planificarCarga(ev, inspeccion(gesperA07(cambios)), A07); }

test("A07: reconoce el anexo y P7; agrupa en RED y ACO, y la referencia general D792 no crea un grupo", () => {
  assert.throws(() => evaluarAnexo(interpretarSalida(libroA07()), A04), /anexo A07, no al A04/);
  const ev = evaluarA07();
  assert.equal(ev.candidatos.length, 9);
  const plan = planA07(ev);
  assert.equal(plan.escrituras.length, 9, "incluido M807, cuyo rótulo termina en espacio");
  assert.deepEqual(plan.escrituras.filter((e) => e.cambia).map((e) => [e.celda, e.actual, e.valor]), [["N800", 0.04525, 45.25]]);
  assert.deepEqual(plan.grupos.map((x) => x.id), ["RED", "ACO"]);
  assert.deepEqual(plan.referencias.map((r) => [r.fila.ID, r.grupo, r.celda, r.gesper, r.formula, r.coincide]), [
    ["A07.RED.LONG_TOT", "RED", "D794", 200, true, false],
    ["A07.GEN.LONG_RED_SIST", "GEN", "D792", 5000, true, true]
  ]);
  assert.equal(plan.grupos.filter((x) => x.id === "RED")[0].referencias.length, 1, "D792 no es de ninguna tarjeta: la pantalla la lista aparte");
  assert.deepEqual(seleccionInicial(plan), { RED: false, ACO: true });
});

test("A07: cargar solo ACO escribe sus datos, deja la fila 800 como estaba y nunca escribe D794 ni D792", async () => {
  const ev = evaluarA07();
  const { stats, rangos } = instalarGesper(gesperA07());
  const vista = await prepararCarga(ev, A07);
  const r = await ejecutarCarga(ev, A07, vista, ["ACO"]);
  assert.equal(r.escritos, 3);
  assert.deepEqual([rangos.D795.values, rangos.M816.values, rangos.M822.values], [[[9]], [[4]], [[1500]]]);
  assert.deepEqual(rangos.N800.values, [[0.04525]], "RED no se marcó: GESPER conserva su valor");
  assert.deepEqual([rangos.D794.formulas, rangos.D792.formulas], [[["=0.2*1000"]], [["=$D$673"]]]);
  assert.deepEqual(r.referencias.map((x) => x.celda), ["D794", "D792"], "las dos referencias se releen tras la carga");
  assert.equal(stats.calculateCalls, 1);
});

test("A07: un control de serie en REVISAR (RED.CTRL_V6) bloquea solo esa serie; el resto de RED y ACO se puede cargar", () => {
  // El anexo pasa a REVISAR la serie V6 del año 2; el panel también omite el
  // año 3, que llegó en OK, porque su control la invalida (alcance "concepto").
  const filas = filasA07().map((f) => (f.ID === "A07.RED.CTRL_V6" || f.ID === "A07.RED.V6.Y02" ? Object.assign({}, f, { ESTADO: "REVISAR" }) : f));
  const ev = evaluarA07({ filas, meta: { ESTADO_GLOBAL: "REVISAR", N_DATO_OK: 8, N_REVISAR: 2, GRUPOS_REVISAR: "RED" } });
  assert.equal(ev.cargaParcial, true);
  assert.equal(ev.bloqueo, null);
  assert.deepEqual(ev.omitidos.map((o) => o.fila.ID), ["A07.RED.V6.Y02", "A07.RED.V6.Y03"]);
  assert.match(ev.omitidos[1].motivo, /A07\.RED\.CTRL_V6 está en REVISAR/);
  const plan = planA07(ev);
  assert.deepEqual(plan.escrituras.map((e) => e.celda), ["M798", "N798", "N800", "N809", "D795", "M816", "M822"]);
  const red = plan.grupos.filter((x) => x.id === "RED")[0];
  assert.deepEqual([red.revisar, red.omitidos.length, red.requiereConfirmacion], [true, 2, true]);
  assert.deepEqual(seleccionInicial(plan), { RED: false, ACO: true });
});

test("A07: nunca escribe banderas (802, 817) ni horizontes (D796, K796, D815, K815); las acometidas avisan con la bandera 817", () => {
  const horiz = [["A07.RED.HORIZ", "D796"], ["A07.RED.HORIZ.Y00", "K796"], ["A07.ACO.HORIZ", "D815"], ["A07.ACO.HORIZ.Y00", "K815"]];
  const extra = horiz.map(([id, celda]) =>
    fila({ ID: id, TIPO_FILA: "DATO", PROGRAMA: "P7", ANIO_IDX: /Y00$/.test(id) ? 0 : null, VALOR: 5, ESTADO: "OK", DESTINO_CELDA: celda, DESTINO_ETIQUETA: ETQ_HORIZ_A07 }))
    .concat([fila({ ID: "A07.ACO.BANDERA.Y02", TIPO_FILA: "DATO", PROGRAMA: "P7", ANIO_IDX: 2, VALOR: 1, ESTADO: "OK", DESTINO_CELDA: "M817", DESTINO_ETIQUETA: ETQ_CONTEMPLADO })]);
  const plan = planA07(evaluarA07({ filas: filasA07().concat(extra) }), { M817: "N/A" });
  for (const [id, celda] of horiz) assert.match(motivoDe(plan, id), new RegExp(celda + " es una celda de horizonte"));
  assert.match(motivoDe(plan, "A07.ACO.BANDERA.Y02"), /contemplado/);
  // M816 y M822 dependen de la bandera 817; las filas de redes del mismo año (802 en NO) y D795 (no anual) no avisan.
  assert.deepEqual(plan.avisos.filter((a) => a.tipo === "HORIZONTE").map((a) => a.fila.ID), ["A07.ACO.CANT.Y02", "A07.ACO.VOL.Y02"]);
});
