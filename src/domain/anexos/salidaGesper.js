/**
 * Contrato de la hoja SALIDA_GESPER de un anexo técnico (Especificacion_SALIDA_GESPER_v1).
 *
 * Funciones puras, en el orden en que las usa la carga:
 *   1. interpretarSalida(libro)      -> { meta, filas }   estructura y versión
 *   2. evaluarAnexo(salida, anexo)   -> qué filas son candidatas a cargar (spec §4, §6)
 *   3. direccionesAInspeccionar(...) -> qué celdas de GESPER hay que leer para validarlas
 *   4. planificarCarga(..., inspeccion) -> qué se escribe, qué se omite y por qué,
 *      y qué avisos requieren confirmación (spec §6 pasos 4-8)
 *
 * No conoce Office.js ni el DOM: el libro del anexo llega ya leído por
 * utils/xlsx.js y la inspección de GESPER por excelService.leerCeldas. Así es
 * reutilizable para los anexos A02, A03... (mismas columnas, spec §10) con
 * solo agregar su entrada en ANEXOS de config/cellMap.js.
 *
 * Nada de aquí corrige un dato: si un destino no cumple, se omite y se reporta.
 */
import { parsearCelda, rangoComoMatriz } from "../../utils/xlsx.js";
import { hasFormula } from "../../utils/format.js";

export const HOJA_SALIDA = "SALIDA_GESPER";
export const NOMBRE_DATOS = "SALIDA_GESPER_DATOS";
export const NOMBRE_META = "SALIDA_GESPER_META";
export const SPEC_MAYOR_SOPORTADA = 1;

/** Columnas A..S, en orden (spec §3). Una columna nueva al final es compatible. */
export const COLUMNAS = [
  "ID", "TIPO_FILA", "PROGRAMA", "SITIO", "CONCEPTO", "RANGO_CLAVE", "ANIO_IDX", "VALOR", "UNIDAD",
  "VALOR_CONTROL", "ESTADO", "ORIGEN_HOJA", "ORIGEN_REF", "DESTINO_HOJA", "DESTINO_CELDA",
  "DESTINO_ETIQUETA", "TRANSFORMACION", "DEPENDE_DECISION", "NOTA"
];

const ESTADOS_CARGABLES = ["OK", "ADVERTENCIA"];
const ESTADOS_GLOBALES = ["OK", "OK CON ADVERTENCIAS", "REVISAR"];

/** Compara textos de Excel sin que cuenten espacios dobles, no separables o la forma Unicode. */
export function normalizarTexto(s) {
  return String(s == null ? "" : s).normalize("NFC").replace(/[\s ]+/g, " ").trim();
}

function vacio(v) { return v === null || v === undefined || (typeof v === "string" && v.trim() === ""); }

/** "1.0" -> { mayor: 1, menor: 0 }; null si no tiene la forma mayor.menor. */
export function parsearVersion(v) {
  var m = /^(\d+)(?:\.(\d+))?$/.exec(String(v == null ? "" : v).trim());
  return m ? { mayor: parseInt(m[1], 10), menor: m[2] ? parseInt(m[2], 10) : 0 } : null;
}

/* ---------- 1. estructura ---------- */

/** Busca la fila cuya columna A dice `texto` (respaldo cuando faltan los nombres definidos, spec §2). */
function filaConTextoEnA(celdas, texto) {
  var filas = Object.keys(celdas)
    .filter(function (k) { return /^A\d+$/.test(k) && normalizarTexto(celdas[k]) === texto; })
    .map(function (k) { return parseInt(k.slice(1), 10); });
  return filas.length ? Math.min.apply(null, filas) : 0;
}

function ultimaFila(celdas) {
  var max = 0;
  Object.keys(celdas).forEach(function (k) { var f = parseInt(k.replace(/^[A-Z]+/, ""), 10); if (f > max) max = f; });
  return max;
}

function refDeNombre(libro, nombre) {
  var n = libro.nombres[nombre];
  return n && n.hoja === HOJA_SALIDA ? n.ref : null;
}

/**
 * Lee metadatos y tabla de SALIDA_GESPER desde el libro ya leído por leerXlsx.
 * Lanza un Error con un mensaje para el usuario si la estructura no sirve
 * (falta la hoja, faltan columnas, versión mayor incompatible).
 *
 * @returns {{meta: Object, filas: Object[]}} cada fila trae las 19 columnas
 *   por nombre y `filaAnexo` (número de fila en la hoja del anexo).
 */
export function interpretarSalida(libro) {
  if (libro.hojas.indexOf(HOJA_SALIDA) < 0 || !libro.celdas[HOJA_SALIDA]) {
    throw new Error("El archivo no tiene la hoja " + HOJA_SALIDA + ". Verifica que sea la versión del anexo preparada para GESPER.");
  }
  var celdas = libro.celdas[HOJA_SALIDA];

  // Metadatos: nombre definido; si no está, la fila donde A dice SPEC_VERSION y la siguiente.
  var refMeta = refDeNombre(libro, NOMBRE_META);
  if (!refMeta) {
    var fm = filaConTextoEnA(celdas, "SPEC_VERSION");
    if (!fm) throw new Error("No encuentro los metadatos de " + HOJA_SALIDA + " (SPEC_VERSION, ANEXO_ID…).");
    refMeta = "A" + fm + ":Z" + (fm + 1);
  }
  var mm = rangoComoMatriz(celdas, refMeta);
  var meta = {};
  (mm[0] || []).forEach(function (clave, i) {
    if (!vacio(clave)) meta[normalizarTexto(clave)] = mm[1] ? mm[1][i] : null;
  });

  var version = parsearVersion(meta.SPEC_VERSION);
  if (!version) throw new Error("El anexo no declara SPEC_VERSION válida (se leyó «" + (meta.SPEC_VERSION == null ? "" : meta.SPEC_VERSION) + "»).");
  if (version.mayor !== SPEC_MAYOR_SOPORTADA) {
    throw new Error("El anexo usa la especificación " + meta.SPEC_VERSION + "; el panel soporta la versión " +
      SPEC_MAYOR_SOPORTADA + ".x. No se carga para no interpretar mal sus columnas.");
  }
  ["ANEXO_ID", "ESTADO_GLOBAL"].forEach(function (k) {
    if (vacio(meta[k])) throw new Error("Falta el metadato " + k + " en " + HOJA_SALIDA + ".");
  });

  // Tabla: nombre definido; si no está, desde la fila donde A dice ID hasta la última con datos.
  var refDatos = refDeNombre(libro, NOMBRE_DATOS);
  if (!refDatos) {
    var fi = filaConTextoEnA(celdas, "ID");
    if (!fi) throw new Error("No encuentro la tabla de " + HOJA_SALIDA + " (fila de encabezados con «ID» en la columna A).");
    refDatos = "A" + fi + ":S" + Math.max(fi, ultimaFila(celdas));
  }
  var m = rangoComoMatriz(celdas, refDatos);
  var enc = m[0] || [];
  COLUMNAS.forEach(function (col, i) {
    if (normalizarTexto(enc[i]) !== col) {
      throw new Error("La tabla de " + HOJA_SALIDA + " no tiene las columnas esperadas: en la columna " +
        String.fromCharCode(65 + i) + " se esperaba «" + col + "» y dice «" + normalizarTexto(enc[i]) + "».");
    }
  });

  var filaInicial = parseInt(refDatos.split(":")[0].replace(/^[A-Z]+/, ""), 10);
  var filas = [];
  for (var r = 1; r < m.length; r++) {
    if (m[r].every(vacio)) continue;
    var o = { filaAnexo: filaInicial + r };
    COLUMNAS.forEach(function (col, i) { o[col] = m[r][i]; });
    o.ID = normalizarTexto(o.ID);
    o.TIPO_FILA = normalizarTexto(o.TIPO_FILA).toUpperCase();
    o.ESTADO = normalizarTexto(o.ESTADO).toUpperCase();
    filas.push(o);
  }
  return { meta: meta, version: version, filas: filas };
}

/* ---------- 2. clasificación ---------- */

/**
 * Qué datos invalida un CONTROL en REVISAR, según `anexo.alcanceControles`:
 *  - "concepto" (A01, spec v1 §6): A01.CAP.CTRL_INST invalida A01.CAP.INST.*;
 *    A01.CAP.CTRL_REEMP invalida A01.CAP.REEMP.* y A01.CAP.REEMP_TOT.
 *  - "grupo" (A02, spec A02 §6): cualquier control de A02.VAL.* invalida
 *    todos los datos A02.VAL.*.
 */
function controlCubre(idControl, idDato, alcance) {
  var c = idControl.split("."), d = idDato.split(".");
  if (c.length < 3 || d.length < 3 || c[1] !== d[1]) return false;
  if (alcance === "grupo") return true;
  if (c[2].indexOf("CTRL_") !== 0) return false;
  var concepto = c[2].slice(5);
  return d[2] === concepto || d[2].indexOf(concepto + "_") === 0;
}

/** "P3, P5 y P6" -> ["P3","P5","P6"]. */
function listaProgramas(v) {
  return normalizarTexto(v).toUpperCase().split(/[^A-Z0-9]+/).filter(function (p) { return /^P\d+$/.test(p); }).sort();
}

/**
 * Separa las filas según TIPO_FILA y ESTADO (spec §4, §6 pasos 2-3).
 * Lanza Error si el archivo es de otro anexo.
 */
export function evaluarAnexo(salida, anexo) {
  var meta = salida.meta;
  var anexoId = normalizarTexto(meta.ANEXO_ID);
  if (anexoId !== anexo.id) {
    throw new Error("El archivo corresponde al anexo " + anexoId + ", no al " + anexo.id + " (" + anexo.nombre + ").");
  }
  var esperados = anexo.programas.slice().sort();
  if (listaProgramas(meta.PROGRAMAS).join(",") !== esperados.join(",")) {
    throw new Error("El anexo declara los programas «" + normalizarTexto(meta.PROGRAMAS) + "»; para el " + anexo.id +
      " se esperaba " + esperados.join(", ") + ".");
  }

  var ev = {
    meta: meta, version: salida.version, totalFilas: salida.filas.length,
    datos: [], candidatos: [], omitidos: [], controles: [], controlesRevisar: [], noDisponibles: [], otras: [],
    bloqueo: null
  };

  salida.filas.forEach(function (f) {
    if (f.TIPO_FILA === "DATO") ev.datos.push(f);
    else if (f.TIPO_FILA === "CONTROL") { ev.controles.push(f); if (f.ESTADO === "REVISAR") ev.controlesRevisar.push(f); }
    else if (f.TIPO_FILA === "NO_DISPONIBLE") ev.noDisponibles.push(f);
    else ev.otras.push(f);
  });

  ev.datos.forEach(function (f) {
    if (f.ID.indexOf(anexo.id + ".") !== 0) { ev.omitidos.push({ fila: f, motivo: "El ID no empieza por " + anexo.id + ".: no es un dato de este anexo." }); return; }
    if (anexo.programas.indexOf(normalizarTexto(f.PROGRAMA).toUpperCase()) < 0) {
      ev.omitidos.push({ fila: f, motivo: "El programa «" + normalizarTexto(f.PROGRAMA) + "» no es de este anexo (" + anexo.programas.join(", ") + ")." });
      return;
    }
    if (f.ESTADO === "REVISAR") { ev.omitidos.push({ fila: f, motivo: "El anexo marca este dato en REVISAR." }); return; }
    if (ESTADOS_CARGABLES.indexOf(f.ESTADO) < 0) { ev.omitidos.push({ fila: f, motivo: "Estado desconocido «" + f.ESTADO + "»." }); return; }
    var ctrl = ev.controlesRevisar.filter(function (c) { return controlCubre(c.ID, f.ID, anexo.alcanceControles); })[0];
    if (ctrl) { ev.omitidos.push({ fila: f, motivo: "El control " + ctrl.ID + " está en REVISAR e invalida este dato." }); return; }
    ev.candidatos.push(f);
  });

  var global = normalizarTexto(meta.ESTADO_GLOBAL).toUpperCase();
  if (global === "REVISAR") {
    ev.bloqueo = "El anexo está en REVISAR: hay datos o controles que no cuadran. Corrígelos en el anexo, actualiza y guarda antes de cargar.";
  } else if (ESTADOS_GLOBALES.indexOf(global) < 0) {
    ev.bloqueo = "ESTADO_GLOBAL desconocido («" + meta.ESTADO_GLOBAL + "»). No se carga.";
  }
  ev.estadoGlobal = global;
  return ev;
}

/* ---------- 3. qué leer de GESPER ---------- */

function banderaDe(anexo, fila) {
  var b = (anexo.banderasHorizonte || []).filter(function (x) { return fila >= x.desde && fila <= x.hasta; })[0];
  return b ? b.fila : 0;
}

/**
 * Celdas de la hoja destino que planificarCarga necesita: etiqueta (B), año
 * (fila de años), año calendario (solo para mostrarlo), destino y bandera de horizonte.
 */
export function direccionesAInspeccionar(evaluacion, anexo) {
  var set = {};
  evaluacion.candidatos.forEach(function (f) {
    var c = parsearCelda(f.DESTINO_CELDA);
    if (!c) return;
    set["B" + c.fila] = 1;
    set[c.columna + c.fila] = 1;
    if (!vacio(f.ANIO_IDX)) {
      set[c.columna + anexo.filaAnios] = 1;
      if (anexo.filaAniosCalendario) set[c.columna + anexo.filaAniosCalendario] = 1;
      var fb = banderaDe(anexo, c.fila);
      if (fb) set[c.columna + fb] = 1;
    }
  });
  return Object.keys(set);
}

/* ---------- 4. plan de carga ---------- */

function motivoDestino(f, c, insp, anexo, ocupadas) {
  if (normalizarTexto(f.DESTINO_HOJA) !== normalizarTexto(anexo.hojaDestino)) {
    return "La hoja destino es «" + normalizarTexto(f.DESTINO_HOJA) + "», no " + anexo.hojaDestino + ".";
  }
  if (!c) return "DESTINO_CELDA «" + normalizarTexto(f.DESTINO_CELDA) + "» no es una celda válida.";
  if (typeof f.VALOR !== "number" || !isFinite(f.VALOR)) return "VALOR no es un número (vacío, texto o error en el anexo).";

  if ((anexo.celdasProtegidas || []).indexOf(c.columna + c.fila) >= 0) {
    return "La celda " + c.columna + c.fila + " es una celda de horizonte que no se carga desde el anexo.";
  }
  var etiqueta = (insp.celdas["B" + c.fila] || {}).texto;
  if ((anexo.filasBandera || []).indexOf(c.fila) >= 0 || normalizarTexto(etiqueta).toLowerCase().indexOf("contemplado") >= 0) {
    return "La fila " + c.fila + " es una bandera «¿contemplado?»: no se carga desde el anexo.";
  }
  if (normalizarTexto(etiqueta) !== normalizarTexto(f.DESTINO_ETIQUETA)) {
    return "La etiqueta de GESPER en B" + c.fila + " («" + normalizarTexto(etiqueta) + "») no coincide con la del anexo («" +
      normalizarTexto(f.DESTINO_ETIQUETA) + "»).";
  }
  if (!vacio(f.ANIO_IDX)) {
    var anio = Number(f.ANIO_IDX);
    if (!Number.isInteger(anio) || anio < 0 || anio > 10) return "ANIO_IDX «" + f.ANIO_IDX + "» no es un año entre 0 y 10.";
    var hdr = (insp.celdas[c.columna + anexo.filaAnios] || {}).valor;
    if (vacio(hdr) || Number(hdr) !== anio) {
      return "La columna " + c.columna + " es el año " + (vacio(hdr) ? "(vacío)" : hdr) + " en GESPER (fila " + anexo.filaAnios +
        "), no el año " + anio + " del anexo.";
    }
  }
  var destino = insp.celdas[c.columna + c.fila] || {};
  if (hasFormula([[destino.formula]])) return "La celda " + c.columna + c.fila + " tiene una fórmula en GESPER.";
  if (ocupadas[c.columna + c.fila]) return "Otra fila del anexo (" + ocupadas[c.columna + c.fila] + ") ya escribe en " + c.columna + c.fila + ".";
  return null;
}

/**
 * Cruza los candidatos con lo que hoy tiene GESPER (inspeccion =
 * resultado de excelService.leerCeldas sobre direccionesAInspeccionar).
 *
 * La confirmación del usuario se exige si hay avisos, si el anexo está en
 * OK CON ADVERTENCIAS o si algún valor nuevo difiere del que hoy tiene
 * GESPER (p. ej. un 0 del anexo que reemplaza un 10): regla general, no una
 * excepción por celda.
 *
 * @returns {{
 *   escrituras: Array<{fila: Object, hoja: string, celda: string, valor: number, actual: any, cambia: boolean,
 *                      anio: number|null, anioCalendario: any}>,
 *   omitidos: Array<{fila: Object, motivo: string}>,
 *   avisos: Array<{fila: Object, tipo: string, texto: string}>,
 *   cambios: number,
 *   requiereConfirmacion: boolean,
 *   bloqueo: string|null
 * }}
 */
export function planificarCarga(evaluacion, inspeccion, anexo) {
  var plan = {
    escrituras: [], omitidos: evaluacion.omitidos.slice(), avisos: [], cambios: 0,
    requiereConfirmacion: false, bloqueo: evaluacion.bloqueo
  };
  if (!inspeccion.hojaExiste) {
    plan.bloqueo = "El libro abierto no tiene la hoja " + anexo.hojaDestino + ". Abre el libro GESPER antes de cargar el anexo.";
    return plan;
  }

  var ocupadas = {};
  evaluacion.candidatos.forEach(function (f) {
    var c = parsearCelda(f.DESTINO_CELDA);
    var motivo = motivoDestino(f, c, inspeccion, anexo, ocupadas);
    if (motivo) { plan.omitidos.push({ fila: f, motivo: motivo }); return; }

    var celda = c.columna + c.fila;
    ocupadas[celda] = f.ID;
    var actual = (inspeccion.celdas[celda] || {}).valor;
    var anual = !vacio(f.ANIO_IDX);
    plan.escrituras.push({
      fila: f, hoja: anexo.hojaDestino, celda: celda, valor: f.VALOR, actual: actual,
      cambia: !(typeof actual === "number" && actual === f.VALOR),
      anio: anual ? Number(f.ANIO_IDX) : null,
      anioCalendario: anual && anexo.filaAniosCalendario ? (inspeccion.celdas[c.columna + anexo.filaAniosCalendario] || {}).valor : null
    });

    if (f.ESTADO === "ADVERTENCIA") {
      plan.avisos.push({
        fila: f, tipo: "ADVERTENCIA",
        texto: "Valor " + f.VALOR + (vacio(f.VALOR_CONTROL) ? "" : " · control " + f.VALOR_CONTROL) +
          (vacio(f.NOTA) ? "" : " · " + normalizarTexto(f.NOTA))
      });
    }
    var fb = vacio(f.ANIO_IDX) ? 0 : banderaDe(anexo, c.fila);
    var bandera = fb ? (inspeccion.celdas[c.columna + fb] || {}).valor : null;
    if (f.VALOR !== 0 && fb && normalizarTexto(bandera).toUpperCase() === "N/A") {
      plan.avisos.push({
        fila: f, tipo: "HORIZONTE",
        texto: "El año " + f.ANIO_IDX + " está en N/A en GESPER (" + c.columna + fb + "): las fórmulas descartan este valor en silencio."
      });
    }
  });

  plan.cambios = plan.escrituras.filter(function (e) { return e.cambia; }).length;
  plan.requiereConfirmacion = plan.avisos.length > 0 || plan.cambios > 0 || evaluacion.estadoGlobal === "OK CON ADVERTENCIAS";
  return plan;
}

/**
 * Huella de lo que el usuario confirma: cada escritura (celda, valor actual,
 * valor nuevo) y cada aviso. Si GESPER cambia entre la vista previa y la
 * carga, la huella cambia y la confirmación deja de valer.
 */
export function firmaPlan(plan) {
  return plan.escrituras.map(function (e) { return e.celda + ":" + JSON.stringify(e.actual) + ">" + e.valor; }).join("|") +
    "#" + plan.avisos.map(function (a) { return a.tipo + ":" + a.fila.ID; }).join("|");
}

/** "SI – motivo" en DEPENDE_DECISION (spec §3, columna R). */
export function dependeDecision(fila) {
  return /^S[IÍ](?![A-Za-z])/i.test(normalizarTexto(fila.DEPENDE_DECISION));
}
