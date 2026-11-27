/**
 * Pruebas de la publicación del panel (scripts/publicacion.js): el manifiesto
 * de producción que se genera a partir de manifest.xml, las plantillas del
 * instalador y el ZIP que descarga cada usuario.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import fs from "node:fs";
import zlib from "node:zlib";
import {
  manifiestoProduccion, normalizarUrlBase, rellenarPlantilla, aCrlf, crc32, crearZip
} from "../scripts/publicacion.js";

const XML_DEV = fs.readFileSync(new URL("../manifest.xml", import.meta.url), "utf8");
const PKG = JSON.parse(fs.readFileSync(new URL("../package.json", import.meta.url), "utf8"));
const URL_BASE = "https://ejemplo.github.io/gesper-panel-v0.1/";
const ID = "86EBD15E-F5F7-4619-A605-65F3CD4028E6";

function prod(cambios) {
  return manifiestoProduccion(XML_DEV, Object.assign({ urlBase: URL_BASE, id: ID, version: "0.1.0" }, cambios));
}

/* ---------- URL base ---------- */

test("normalizarUrlBase agrega la barra final y no la duplica", () => {
  assert.equal(normalizarUrlBase("https://x.github.io/repo"), "https://x.github.io/repo/");
  assert.equal(normalizarUrlBase("https://x.github.io/repo/"), "https://x.github.io/repo/");
  assert.equal(normalizarUrlBase("https://X.GitHub.io/repo"), "https://x.github.io/repo/");
});

test("normalizarUrlBase rechaza HTTP, consultas y URL vacías", () => {
  assert.throws(() => normalizarUrlBase("http://x.github.io/repo/"), /HTTPS/);
  assert.throws(() => normalizarUrlBase("https://x.github.io/repo/?v=1"), /«\?»/);
  assert.throws(() => normalizarUrlBase(""), /inválida/);
  assert.throws(() => normalizarUrlBase(undefined), /inválida/);
});

/* ---------- Manifiesto de producción ---------- */

test("el manifiesto de producción no menciona localhost y apunta al sitio publicado", () => {
  const xml = prod();
  assert.doesNotMatch(xml, /localhost/i);
  assert.match(xml, /<SourceLocation DefaultValue="https:\/\/ejemplo\.github\.io\/gesper-panel-v0\.1\/src\/taskpane\.html"\/>/);
  assert.match(xml, /<bt:Url id="Taskpane\.Url" DefaultValue="https:\/\/ejemplo\.github\.io\/gesper-panel-v0\.1\/src\/taskpane\.html"\/>/);
  assert.match(xml, /<IconUrl DefaultValue="https:\/\/ejemplo\.github\.io\/gesper-panel-v0\.1\/assets\/icon-64\.png"\/>/);
});

test("AppDomain lleva solo el origen, sin la ruta del repositorio", () => {
  assert.match(prod(), /<AppDomain>https:\/\/ejemplo\.github\.io<\/AppDomain>/);
});

test("el manifiesto de producción cambia el Id y la versión, y conserva el resto", () => {
  const xml = prod({ version: "0.2.0" });
  assert.match(xml, new RegExp("<Id>" + ID + "</Id>"));
  assert.match(xml, /<Version>0\.2\.0<\/Version>/);
  // Lo que no es URL, Id, versión ni rótulo queda igual: misma cantidad de líneas y mismos textos.
  assert.equal(xml.split("\n").length, XML_DEV.split("\n").length);
  assert.match(xml, /<Permissions>ReadWriteDocument<\/Permissions>/);
});

test("el panel de desarrollo se rotula «(desarrollo)» y el publicado no", () => {
  assert.match(XML_DEV, /<DisplayName DefaultValue="Panel GESPER \(desarrollo\)"\/>/);
  assert.match(XML_DEV, /id="Button\.Label" DefaultValue="Panel GESPER \(desarrollo\)"/);
  const xml = prod();
  assert.doesNotMatch(xml, /desarrollo/i);
  assert.match(xml, /<DisplayName DefaultValue="Panel GESPER"\/>/);
  assert.match(xml, /id="Group\.Label" DefaultValue="GESPER"/);
  assert.match(xml, /id="Button\.Label" DefaultValue="Panel GESPER"/);
  assert.match(xml, /id="GetStarted\.Title" DefaultValue="Panel GESPER instalado"/);
  assert.match(xml, /Busca el boton Panel GESPER en la pestana Inicio/);
});

test("falla si queda un «(desarrollo)» escrito de otra forma", () => {
  const raro = XML_DEV.replace(' DefaultValue="GESPER (desarrollo)"', ' DefaultValue="GESPER(Desarrollo)"');
  assert.notEqual(raro, XML_DEV);
  assert.throws(() => manifiestoProduccion(raro, { urlBase: URL_BASE, id: ID, version: "0.1.0" }), /desarrollo/);
});

test("el Id de producción tiene que ser distinto del de desarrollo", () => {
  const idDev = /<Id>([^<]*)<\/Id>/.exec(XML_DEV)[1];
  assert.throws(() => prod({ id: idDev }), /mismo que el de manifest\.xml/);
  assert.throws(() => prod({ id: idDev.toLowerCase() }), /mismo que el de manifest\.xml/);
});

test("rechaza un Id que no es GUID y una versión que Office no acepta", () => {
  assert.throws(() => prod({ id: "gesper" }), /no es un GUID/);
  assert.throws(() => prod({ id: undefined }), /no es un GUID/);
  assert.throws(() => prod({ version: "0.1.0-beta" }), /formato de Office/);
  assert.throws(() => prod({ version: "1.2.3.4.5" }), /formato de Office/);
});

test("falla si manifest.xml trae una URL de localhost que no se puede reemplazar", () => {
  const raro = XML_DEV.replace("<Permissions>", "<!-- http://localhost:3000/x -->\n  <Permissions>");
  assert.notEqual(raro, XML_DEV);
  assert.throws(() => manifiestoProduccion(raro, { urlBase: URL_BASE, id: ID, version: "0.1.0" }), /localhost/);
});

test("falla si manifest.xml no tiene exactamente un <AppDomain> de localhost", () => {
  const sin = XML_DEV.replace(/<AppDomain>[^<]*<\/AppDomain>/, "");
  assert.throws(() => manifiestoProduccion(sin, { urlBase: URL_BASE, id: ID, version: "0.1.0" }), /AppDomain/);
});

test("package.json trae una configuración de publicación válida", () => {
  const xml = manifiestoProduccion(XML_DEV, { urlBase: PKG.publicacion.url, id: PKG.publicacion.idManifiesto, version: PKG.version });
  assert.doesNotMatch(xml, /localhost/i);
});

/* ---------- Plantillas ---------- */

test("rellenarPlantilla sustituye cada clave y falla si falta alguna", () => {
  assert.equal(rellenarPlantilla("set \"ID={{ID}}\" v{{VERSION}} {{ID}}", { ID: "A", VERSION: "1" }), "set \"ID=A\" v1 A");
  assert.throws(() => rellenarPlantilla("{{URL}}", { ID: "A" }), /\{\{URL\}\}/);
});

test("aCrlf deja todos los saltos de línea como CRLF, sin duplicar los que ya lo eran", () => {
  assert.equal(aCrlf("a\nb\r\nc\n"), "a\r\nb\r\nc\r\n");
});

test("las plantillas del instalador solo usan claves conocidas y los .cmd son ASCII", () => {
  const valores = { ID: ID, VERSION: "0.1.0", URL: URL_BASE, ZIP: "x.zip" };
  for (const nombre of ["Instalar Panel GESPER.cmd", "Desinstalar Panel GESPER.cmd", "LEEME.txt", "index.html"]) {
    const texto = fs.readFileSync(new URL("../instalador/" + nombre, import.meta.url), "utf8");
    const lleno = rellenarPlantilla(texto, valores);
    assert.doesNotMatch(lleno, /\{\{/, nombre);
    if (nombre.endsWith(".cmd")) {
      assert.doesNotMatch(texto, /[^\x00-\x7f]/, nombre + " debe ser ASCII");
      assert.match(lleno, new RegExp('set "ID=' + ID + '"'), nombre);
    }
  }
});

/* ---------- ZIP ---------- */

test("crc32 da el valor de referencia de «123456789»", () => {
  assert.equal(crc32(Buffer.from("123456789")), 0xcbf43926);
  assert.equal(crc32(Buffer.alloc(0)), 0);
});

/** Lee un ZIP desde su directorio central y verifica el CRC de cada entrada. */
function leerZip(buf) {
  const fin = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  assert.ok(fin >= 0, "sin fin de directorio central");
  const total = buf.readUInt16LE(fin + 10);
  let p = buf.readUInt32LE(fin + 16);
  const out = {};
  for (let i = 0; i < total; i++) {
    assert.equal(buf.readUInt32LE(p), 0x02014b50);
    const metodo = buf.readUInt16LE(p + 10), crc = buf.readUInt32LE(p + 16);
    const tamComp = buf.readUInt32LE(p + 20), tam = buf.readUInt32LE(p + 24);
    const largo = buf.readUInt16LE(p + 28), offset = buf.readUInt32LE(p + 42);
    const nombre = buf.subarray(p + 46, p + 46 + largo).toString("utf8");
    assert.equal(buf.readUInt32LE(offset), 0x04034b50, nombre);
    assert.equal(buf.readUInt32LE(offset + 14), crc, nombre + ": CRC distinto en el encabezado local");
    const inicio = offset + 30 + buf.readUInt16LE(offset + 26) + buf.readUInt16LE(offset + 28);
    const datos = buf.subarray(inicio, inicio + tamComp);
    const crudo = metodo === 8 ? zlib.inflateRawSync(datos) : Buffer.from(datos);
    assert.equal(crudo.length, tam, nombre);
    assert.equal(crc32(crudo), crc, nombre + ": CRC no coincide con los datos");
    out[nombre] = crudo;
    p += 46 + largo + buf.readUInt16LE(p + 30) + buf.readUInt16LE(p + 32);
  }
  return out;
}

test("crearZip produce un ZIP legible, con los datos y CRC correctos", () => {
  const largo = "Panel GESPER\r\n".repeat(200); // se comprime
  const zip = crearZip([
    { nombre: "Instalar Panel GESPER.cmd", datos: largo },
    { nombre: "LEEME.txt", datos: "﻿Instalación" },
    { nombre: "x.bin", datos: new Uint8Array([1, 2, 3]) } // más chico sin comprimir
  ], new Date(2026, 10, 27, 12, 30, 10));
  const e = leerZip(zip);
  assert.deepEqual(Object.keys(e), ["Instalar Panel GESPER.cmd", "LEEME.txt", "x.bin"]);
  assert.equal(e["Instalar Panel GESPER.cmd"].toString("utf8"), largo);
  assert.equal(e["LEEME.txt"].toString("utf8"), "﻿Instalación");
  assert.deepEqual([...e["x.bin"]], [1, 2, 3]);
});

test("crearZip rechaza nombres con carpetas y entradas repetidas", () => {
  assert.throws(() => crearZip([{ nombre: "a/b.txt", datos: "" }]), /no válido/);
  assert.throws(() => crearZip([{ nombre: "a.txt", datos: "" }, { nombre: "A.TXT", datos: "" }]), /repetida/);
});
