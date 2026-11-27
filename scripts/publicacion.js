/**
 * Funciones puras de la publicación del panel (las usa scripts/empaquetar.js):
 * el manifiesto de producción, las plantillas del instalador y el ZIP que
 * descarga cada usuario. No leen ni escriben archivos, así que se prueban con
 * node:test (test/publicacion.test.js). Ver docs/DECISIONES.md §13.
 *
 * Viven en scripts/ y no en src/utils/ porque no son parte del panel: lo que
 * hay en src/ se publica tal cual en GitHub Pages.
 */
import zlib from "node:zlib";

/** Origen con el que el manifiesto de desarrollo (manifest.xml) apunta a node server.js. */
export const ORIGEN_DESARROLLO = "https://localhost:3000";

/**
 * Sufijo de los rótulos de manifest.xml («Panel GESPER (desarrollo)») para
 * distinguir los dos botones cuando ambos paneles están instalados. El
 * manifiesto de producción lo quita.
 */
export const SUFIJO_DESARROLLO = " (desarrollo)";

const RE_GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

/**
 * "https://x.github.io/repo" -> "https://x.github.io/repo/". Excel solo carga
 * complementos por HTTPS; sin consulta ni fragmento, porque la URL se
 * concatena con las rutas del manifiesto.
 */
export function normalizarUrlBase(url) {
  var u;
  try { u = new URL(String(url || "")); } catch (e) { throw new Error("URL de publicación inválida: «" + url + "»."); }
  if (u.protocol !== "https:") throw new Error("Excel solo carga complementos por HTTPS: «" + url + "».");
  if (u.search || u.hash) throw new Error("La URL de publicación no puede llevar «?» ni «#»: «" + url + "».");
  return u.href.endsWith("/") ? u.href : u.href + "/";
}

/** Reemplaza exactamente una aparición de `re`; si no hay ninguna o hay varias, falla. */
function reemplazarUna(texto, re, nuevo, que) {
  var global = new RegExp(re.source, re.flags.replace("g", "") + "g");
  var n = (texto.match(global) || []).length;
  if (n !== 1) throw new Error("manifest.xml: esperaba un solo " + que + " y hay " + n + ".");
  return texto.replace(re, nuevo);
}

/**
 * Manifiesto de producción a partir del de desarrollo: cambia el Id, la
 * versión y todas las URL de https://localhost:3000 por `urlBase`, y quita el
 * sufijo « (desarrollo)» de los rótulos. El Id es otro para que el panel
 * publicado y el de desarrollo puedan estar instalados a la vez en el mismo
 * Excel sin pisarse.
 *
 * @param {string} xml contenido de manifest.xml
 * @param {{urlBase: string, id: string, version: string}} opciones
 * @returns {string}
 */
export function manifiestoProduccion(xml, opciones) {
  var base = normalizarUrlBase(opciones.urlBase);
  var id = String(opciones.id || "");
  var version = String(opciones.version || "");
  if (!RE_GUID.test(id)) throw new Error("publicacion.idManifiesto (package.json) no es un GUID: «" + id + "».");
  if (!/^\d+(\.\d+){0,3}$/.test(version)) throw new Error("La versión «" + version + "» no tiene el formato de Office (hasta cuatro números: 1.2.3.4).");

  var idDesarrollo = (/<Id>([^<]*)<\/Id>/.exec(xml) || [])[1];
  if (idDesarrollo && idDesarrollo.toLowerCase() === id.toLowerCase()) {
    throw new Error("El Id de producción es el mismo que el de manifest.xml; deben ser distintos para que no se pisen.");
  }

  var out = reemplazarUna(xml, /<Id>[^<]*<\/Id>/, "<Id>" + id.toUpperCase() + "</Id>", "<Id>");
  out = reemplazarUna(out, /<Version>[^<]*<\/Version>/, "<Version>" + version + "</Version>", "<Version>");
  // <AppDomain> lleva solo el origen, sin ruta ni barra final.
  out = reemplazarUna(out, new RegExp("<AppDomain>" + ORIGEN_DESARROLLO + "/?</AppDomain>"),
    "<AppDomain>" + new URL(base).origin + "</AppDomain>", "<AppDomain> de " + ORIGEN_DESARROLLO);
  out = out.split(ORIGEN_DESARROLLO + "/").join(base);
  out = out.split(SUFIJO_DESARROLLO).join("");

  if (/localhost/i.test(out)) {
    throw new Error("El manifiesto de producción todavía menciona localhost. Revisa manifest.xml: cada URL debe empezar por " + ORIGEN_DESARROLLO + "/.");
  }
  if (/\(\s*desarrollo\s*\)/i.test(out)) {
    throw new Error("El manifiesto de producción todavía dice «(desarrollo)». En manifest.xml el sufijo debe escribirse exactamente «" + SUFIJO_DESARROLLO + "».");
  }
  return out;
}

/** Sustituye {{CLAVE}} por `valores.CLAVE`. Una clave sin valor es un error, no un texto vacío. */
export function rellenarPlantilla(texto, valores) {
  return texto.replace(/\{\{(\w+)\}\}/g, function (_, k) {
    if (!Object.prototype.hasOwnProperty.call(valores, k)) throw new Error("La plantilla usa {{" + k + "}} y no tiene valor.");
    return String(valores[k]);
  });
}

/** Saltos de línea de Windows: cmd.exe lee mal algunos .cmd con LF solo. */
export function aCrlf(texto) {
  return texto.replace(/\r?\n/g, "\r\n");
}

/* ---------- ZIP ---------- */

var TABLA_CRC = (function () {
  var t = new Uint32Array(256);
  for (var n = 0; n < 256; n++) {
    var c = n;
    for (var k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

/** CRC-32 (el de ZIP), sin depender de zlib.crc32, que no existe en todo Node 20. */
export function crc32(bytes) {
  var c = 0xffffffff;
  for (var i = 0; i < bytes.length; i++) c = TABLA_CRC[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

function fechaDos(d) {
  var anio = Math.max(1980, d.getFullYear());
  return {
    hora: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
    dia: ((anio - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()
  };
}

/**
 * ZIP con los archivos en la raíz (sin carpetas). Comprime con deflate
 * cuando reduce el tamaño. Es lo que Windows abre con «Extraer todo».
 *
 * @param {{nombre: string, datos: string|Uint8Array}[]} archivos
 * @param {Date} [fecha] fecha de modificación de las entradas
 * @returns {Buffer}
 */
export function crearZip(archivos, fecha) {
  var dos = fechaDos(fecha || new Date());
  var vistos = {};
  var locales = [], centrales = [], offset = 0;
  archivos.forEach(function (a) {
    if (!a.nombre || /[\\/]/.test(a.nombre)) throw new Error("Nombre de entrada no válido para el ZIP: «" + a.nombre + "».");
    if (vistos[a.nombre.toLowerCase()]) throw new Error("Entrada repetida en el ZIP: «" + a.nombre + "».");
    vistos[a.nombre.toLowerCase()] = true;

    var nombre = Buffer.from(a.nombre, "utf8");
    var crudo = typeof a.datos === "string" ? Buffer.from(a.datos, "utf8") : Buffer.from(a.datos);
    var comprimido = zlib.deflateRawSync(crudo);
    var metodo = comprimido.length < crudo.length ? 8 : 0;
    var datos = metodo === 8 ? comprimido : crudo;
    var crc = crc32(crudo);

    var lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0);
    lh.writeUInt16LE(20, 4);         // versión necesaria para extraer: 2.0
    lh.writeUInt16LE(0x0800, 6);     // bit 11: nombre en UTF-8
    lh.writeUInt16LE(metodo, 8);
    lh.writeUInt16LE(dos.hora, 10);
    lh.writeUInt16LE(dos.dia, 12);
    lh.writeUInt32LE(crc, 14);
    lh.writeUInt32LE(datos.length, 18);
    lh.writeUInt32LE(crudo.length, 22);
    lh.writeUInt16LE(nombre.length, 26);

    var ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0);
    ch.writeUInt16LE(20, 4);         // hecho por: 2.0
    ch.writeUInt16LE(20, 6);
    ch.writeUInt16LE(0x0800, 8);
    ch.writeUInt16LE(metodo, 10);
    ch.writeUInt16LE(dos.hora, 12);
    ch.writeUInt16LE(dos.dia, 14);
    ch.writeUInt32LE(crc, 16);
    ch.writeUInt32LE(datos.length, 20);
    ch.writeUInt32LE(crudo.length, 24);
    ch.writeUInt16LE(nombre.length, 28);
    ch.writeUInt32LE(offset, 42);

    locales.push(lh, nombre, datos);
    centrales.push(ch, nombre);
    offset += lh.length + nombre.length + datos.length;
  });

  var directorio = Buffer.concat(centrales);
  var fin = Buffer.alloc(22);
  fin.writeUInt32LE(0x06054b50, 0);
  fin.writeUInt16LE(archivos.length, 8);
  fin.writeUInt16LE(archivos.length, 10);
  fin.writeUInt32LE(directorio.length, 12);
  fin.writeUInt32LE(offset, 16);
  return Buffer.concat(locales.concat([directorio, fin]));
}
