/**
 * Lector mínimo de archivos .xlsx/.xlsm, sin dependencias.
 *
 * Existe para leer un archivo que el usuario elige desde el panel (p. ej. un
 * anexo técnico) sin abrirlo en Excel y sin confundirlo con el libro GESPER
 * abierto: Office.js solo da acceso al libro activo. Ver docs/DECISIONES.md §11.
 *
 * Alcance deliberadamente corto: lee la lista de hojas, los nombres definidos
 * y los valores guardados (el último resultado calculado de cada fórmula) de
 * las hojas que se pidan. No evalúa fórmulas, no lee formatos ni tablas
 * dinámicas. Un .xlsx es un ZIP de XML; la descompresión usa
 * DecompressionStream("deflate-raw"), nativo en WebView2/Safari 16.4+ y en
 * Node 18+ (por eso se puede probar con node:test).
 *
 * Sin Office.js ni DOM: se puede probar con Node normal (ver /test).
 */

/* ---------- direcciones de celda ---------- */

/** "A" -> 1, "Z" -> 26, "AA" -> 27. */
export function letraAColumna(letras) {
  var n = 0;
  for (var i = 0; i < letras.length; i++) n = n * 26 + (letras.charCodeAt(i) - 64);
  return n;
}

/** 1 -> "A", 27 -> "AA". */
export function columnaALetra(n) {
  var s = "";
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/**
 * "M115" o "$M$115" -> { columna: "M", col: 13, fila: 115 }; null si no es
 * una sola celda válida.
 */
export function parsearCelda(addr) {
  var m = /^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/.exec(String(addr == null ? "" : addr).trim().toUpperCase());
  if (!m) return null;
  return { columna: m[1], col: letraAColumna(m[1]), fila: parseInt(m[2], 10) };
}

/**
 * "A6:S127" o "$A$3:$H$4" (o una sola celda) -> { c1, f1, c2, f2 }; null si
 * no es un rango válido.
 */
export function parsearRango(ref) {
  var partes = String(ref == null ? "" : ref).split(":");
  if (partes.length > 2) return null;
  var a = parsearCelda(partes[0]);
  var b = partes.length === 2 ? parsearCelda(partes[1]) : a;
  if (!a || !b) return null;
  return {
    c1: Math.min(a.col, b.col), f1: Math.min(a.fila, b.fila),
    c2: Math.max(a.col, b.col), f2: Math.max(a.fila, b.fila)
  };
}

/**
 * Matriz 2D (filas x columnas) de los valores de `celdas` dentro de `ref`.
 * Las celdas vacías quedan en null.
 */
export function rangoComoMatriz(celdas, ref) {
  var r = parsearRango(ref);
  if (!r) throw new Error("Rango inválido: " + ref);
  var out = [];
  for (var f = r.f1; f <= r.f2; f++) {
    var fila = [];
    for (var c = r.c1; c <= r.c2; c++) {
      var k = columnaALetra(c) + f;
      fila.push(Object.prototype.hasOwnProperty.call(celdas, k) ? celdas[k] : null);
    }
    out.push(fila);
  }
  return out;
}

/* ---------- ZIP ---------- */

function u16(b, i) { return b[i] | (b[i + 1] << 8); }
function u32(b, i) { return (b[i] | (b[i + 1] << 8) | (b[i + 2] << 16) | (b[i + 3] << 24)) >>> 0; }

/**
 * Índice del ZIP: nombre de entrada -> { metodo, tamComprimido, offsetLocal }.
 * Lee el directorio central (no recorre los encabezados locales en orden,
 * que pueden traer tamaños en 0 cuando el ZIP se escribió en streaming).
 */
function indiceZip(bytes) {
  var fin = -1;
  for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 22 - 65535); i--) {
    if (u32(bytes, i) === 0x06054b50) { fin = i; break; }
  }
  if (fin < 0) {
    throw new Error("El archivo no es un .xlsx válido (no es un ZIP). Si es un .xls antiguo o está protegido con contraseña, ábrelo en Excel y guárdalo como .xlsx sin contraseña.");
  }
  var total = u16(bytes, fin + 10);
  var p = u32(bytes, fin + 16);
  var dec = new TextDecoder("utf-8");
  var entradas = {};
  for (var n = 0; n < total; n++) {
    if (u32(bytes, p) !== 0x02014b50) throw new Error("El directorio del archivo .xlsx está dañado.");
    var metodo = u16(bytes, p + 10);
    var tamComp = u32(bytes, p + 20);
    var largoNombre = u16(bytes, p + 28), largoExtra = u16(bytes, p + 30), largoComent = u16(bytes, p + 32);
    var offsetLocal = u32(bytes, p + 42);
    var nombre = dec.decode(bytes.subarray(p + 46, p + 46 + largoNombre));
    if (tamComp === 0xffffffff || offsetLocal === 0xffffffff) {
      throw new Error("El archivo es demasiado grande para este lector (ZIP64 no soportado).");
    }
    entradas[nombre] = { metodo: metodo, tamComprimido: tamComp, offsetLocal: offsetLocal };
    p += 46 + largoNombre + largoExtra + largoComent;
  }
  return entradas;
}

async function inflarRaw(datos) {
  if (typeof DecompressionStream === "undefined") {
    throw new Error("Esta versión de Excel no permite leer archivos desde el panel (falta DecompressionStream). Usa Excel de Microsoft 365 actualizado.");
  }
  var flujo = new Blob([datos]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(flujo).arrayBuffer());
}

async function leerEntrada(bytes, entradas, nombre) {
  var e = entradas[nombre];
  if (!e) return null;
  var p = e.offsetLocal;
  if (u32(bytes, p) !== 0x04034b50) throw new Error("La entrada " + nombre + " del archivo .xlsx está dañada.");
  var inicio = p + 30 + u16(bytes, p + 26) + u16(bytes, p + 28);
  var datos = bytes.subarray(inicio, inicio + e.tamComprimido);
  var crudo;
  if (e.metodo === 0) crudo = datos;
  else if (e.metodo === 8) crudo = await inflarRaw(datos);
  else throw new Error("Método de compresión no soportado (" + e.metodo + ") en " + nombre + ".");
  return new TextDecoder("utf-8").decode(crudo);
}

/* ---------- XML (solo el subconjunto que escribe Excel para estas partes) ---------- */

function decodificarXml(s) {
  return s
    .replace(/&(?:#(\d+)|#x([0-9a-fA-F]+)|(amp|lt|gt|quot|apos));/g, function (_, dec, hex, ent) {
      if (dec) return String.fromCodePoint(parseInt(dec, 10));
      if (hex) return String.fromCodePoint(parseInt(hex, 16));
      return { amp: "&", lt: "<", gt: ">", quot: "\"", apos: "'" }[ent];
    })
    // Escape propio de OOXML para caracteres de control: _x000D_ etc.
    .replace(/_x([0-9A-Fa-f]{4})_/g, function (_, h) { return String.fromCharCode(parseInt(h, 16)); });
}

function atributos(s) {
  var out = {};
  var re = /([\w:]+)\s*=\s*"([^"]*)"/g, m;
  while ((m = re.exec(s))) out[m[1]] = decodificarXml(m[2]);
  return out;
}

/** Texto de un <si> o un <is>: concatena todos los <t>, ignorando la fonética (<rPh>). */
function textoRico(xml) {
  var limpio = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/g, "");
  var re = /<t\b[^>]*?(?:\/>|>([\s\S]*?)<\/t>)/g, m, s = "";
  while ((m = re.exec(limpio))) s += m[1] ? decodificarXml(m[1]) : "";
  return s;
}

function leerCadenasCompartidas(xml) {
  if (!xml) return [];
  var out = [], re = /<si\b[^>]*?(?:\/>|>([\s\S]*?)<\/si>)/g, m;
  while ((m = re.exec(xml))) out.push(m[1] ? textoRico(m[1]) : "");
  return out;
}

function valorCelda(attrs, contenido, compartidas) {
  var t = attrs.t || "n";
  if (t === "inlineStr") {
    var is = /<is\b[^>]*>([\s\S]*?)<\/is>/.exec(contenido);
    return is ? textoRico(is[1]) : "";
  }
  var v = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(contenido);
  if (!v) return null; // fórmula sin resultado guardado, o celda solo con formato
  var crudo = decodificarXml(v[1]);
  if (t === "s") { var i = parseInt(crudo, 10); return i >= 0 && i < compartidas.length ? compartidas[i] : null; }
  if (t === "str") return crudo;
  if (t === "b") return crudo === "1";
  if (t === "e") return { error: crudo };
  var n = Number(crudo);
  return crudo.trim() !== "" && isFinite(n) ? n : null;
}

/** Valores guardados de una hoja, como { "A1": valor, ... } (solo celdas con valor). */
function leerCeldasHoja(xml, compartidas) {
  var celdas = {};
  var reFila = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g, mf, filaPrevia = 0;
  while ((mf = reFila.exec(xml))) {
    var af = atributos(mf[1]);
    var fila = af.r ? parseInt(af.r, 10) : filaPrevia + 1;
    filaPrevia = fila;
    if (!mf[2]) continue;
    var reCel = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, mc, colPrevia = 0;
    while ((mc = reCel.exec(mf[2]))) {
      var ac = atributos(mc[1]);
      var pos = ac.r ? parsearCelda(ac.r) : null;
      var col = pos ? pos.col : colPrevia + 1;
      colPrevia = col;
      if (!mc[2]) continue;
      var v = valorCelda(ac, mc[2], compartidas);
      if (v !== null && v !== "") celdas[columnaALetra(col) + (pos ? pos.fila : fila)] = v;
    }
  }
  return celdas;
}

/** "SALIDA_GESPER!$A$6:$S$127" o "'Mi hoja'!$A$1" -> { hoja, ref }; null si no es una referencia simple. */
function parsearReferenciaNombre(texto) {
  var m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/.exec(String(texto).trim());
  if (!m) return null;
  return { hoja: m[1] ? m[1].replace(/''/g, "'") : m[2], ref: m[3].replace(/\$/g, "") };
}

function resolverRuta(destino) {
  if (destino.charAt(0) === "/") return destino.slice(1);
  var partes = ("xl/" + destino).split("/"), out = [];
  partes.forEach(function (p) { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); });
  return out.join("/");
}

/**
 * Abre un .xlsx/.xlsm desde sus bytes y lee los valores guardados de las
 * hojas pedidas.
 *
 * @param {ArrayBuffer|Uint8Array} buffer
 * @param {{hojas?: string[]}} [opciones] hojas cuyas celdas leer (por defecto, ninguna)
 * @returns {Promise<{hojas: string[], nombres: Object<string,{hoja:string,ref:string}>, celdas: Object<string,Object>}>}
 *   `nombres`: solo los nombres definidos de alcance de libro que apuntan a
 *   un rango simple. `celdas[hoja]`: { "A1": valor } de cada hoja pedida que exista.
 */
export async function leerXlsx(buffer, opciones) {
  var bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  var entradas = indiceZip(bytes);

  var wbXml = await leerEntrada(bytes, entradas, "xl/workbook.xml");
  if (!wbXml) throw new Error("El archivo no es un libro de Excel (.xlsx): no tiene xl/workbook.xml.");
  var relsXml = (await leerEntrada(bytes, entradas, "xl/_rels/workbook.xml.rels")) || "";

  var rutas = {};
  var reRel = /<Relationship\b([^>]*?)\/?>/g, mr;
  while ((mr = reRel.exec(relsXml))) { var ar = atributos(mr[1]); if (ar.Id && ar.Target) rutas[ar.Id] = resolverRuta(ar.Target); }

  var hojas = [], archivoHoja = {};
  var reHoja = /<sheet\b([^>]*?)\/?>/g, mh;
  while ((mh = reHoja.exec(wbXml))) {
    var ah = atributos(mh[1]);
    hojas.push(ah.name);
    archivoHoja[ah.name] = rutas[ah["r:id"]];
  }

  var nombres = {};
  var reNom = /<definedName\b([^>]*?)>([\s\S]*?)<\/definedName>/g, mn;
  while ((mn = reNom.exec(wbXml))) {
    var an = atributos(mn[1]);
    if (an.localSheetId !== undefined) continue;
    var refNom = parsearReferenciaNombre(decodificarXml(mn[2]));
    if (refNom) nombres[an.name] = refNom;
  }

  var celdas = {};
  var pedidas = (opciones && opciones.hojas) || [];
  var compartidas = null;
  for (var i = 0; i < pedidas.length; i++) {
    var ruta = archivoHoja[pedidas[i]];
    if (!ruta) continue;
    if (compartidas === null) compartidas = leerCadenasCompartidas(await leerEntrada(bytes, entradas, "xl/sharedStrings.xml"));
    var xml = await leerEntrada(bytes, entradas, ruta);
    if (xml) celdas[pedidas[i]] = leerCeldasHoja(xml, compartidas);
  }

  return { hojas: hojas, nombres: nombres, celdas: celdas };
}
