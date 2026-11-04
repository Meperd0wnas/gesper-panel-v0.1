/**
 * Pruebas del lector mínimo de .xlsx (src/utils/xlsx.js). Arma en memoria un
 * .xlsx diminuto (ZIP + XML, igual que lo escribe Excel) en vez de depender
 * de un archivo real del cliente.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import zlib from "node:zlib";
import {
  leerXlsx, rangoComoMatriz, parsearCelda, parsearRango, letraAColumna, columnaALetra
} from "../src/utils/xlsx.js";

/** ZIP mínimo: { "ruta": "contenido" } -> Buffer. `comprimir` usa deflate (método 8), si no "stored" (método 0). */
function crearZip(archivos, comprimir) {
  const locales = [], centrales = [];
  let offset = 0;
  for (const [nombre, contenido] of Object.entries(archivos)) {
    const n = Buffer.from(nombre, "utf8");
    const crudo = Buffer.from(contenido, "utf8");
    const datos = comprimir ? zlib.deflateRawSync(crudo) : crudo;
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(comprimir ? 8 : 0, 8);
    lh.writeUInt32LE(datos.length, 18); lh.writeUInt32LE(crudo.length, 22); lh.writeUInt16LE(n.length, 26);
    const ch = Buffer.alloc(46);
    ch.writeUInt32LE(0x02014b50, 0); ch.writeUInt16LE(20, 4); ch.writeUInt16LE(20, 6); ch.writeUInt16LE(comprimir ? 8 : 0, 10);
    ch.writeUInt32LE(datos.length, 20); ch.writeUInt32LE(crudo.length, 24); ch.writeUInt16LE(n.length, 28);
    ch.writeUInt32LE(offset, 42);
    locales.push(lh, n, datos);
    centrales.push(ch, n);
    offset += lh.length + n.length + datos.length;
  }
  const cd = Buffer.concat(centrales);
  const fin = Buffer.alloc(22);
  fin.writeUInt32LE(0x06054b50, 0);
  fin.writeUInt16LE(Object.keys(archivos).length, 8); fin.writeUInt16LE(Object.keys(archivos).length, 10);
  fin.writeUInt32LE(cd.length, 12); fin.writeUInt32LE(offset, 16);
  return Buffer.concat([...locales, cd, fin]);
}

const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

const LIBRO = {
  "xl/workbook.xml":
    `<?xml version="1.0" encoding="UTF-8"?><workbook ${NS}><sheets>` +
    `<sheet name="Datos &amp; más" sheetId="1" r:id="rId1"/><sheet name="SALIDA_GESPER" sheetId="2" state="hidden" r:id="rId2"/>` +
    `</sheets><definedNames>` +
    `<definedName name="_xlnm.Print_Area" localSheetId="0">'Datos &amp; más'!$A$1:$B$2</definedName>` +
    `<definedName name="SALIDA_GESPER_DATOS">SALIDA_GESPER!$A$1:$C$3</definedName>` +
    `<definedName name="ROTO">#REF!</definedName>` +
    `</definedNames></workbook>`,
  "xl/_rels/workbook.xml.rels":
    `<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="x/worksheet" Target="worksheets/sheet1.xml"/>` +
    `<Relationship Id="rId2" Type="x/worksheet" Target="/xl/worksheets/sheet2.xml"/>` +
    `</Relationships>`,
  "xl/sharedStrings.xml":
    `<sst ${NS} count="2" uniqueCount="2"><si><t>&lt;=6</t></si>` +
    `<si><r><t>14"</t></r><r><rPr><b/></rPr><t xml:space="preserve">-60"</t></r><rPh sb="0" eb="1"><t>fonética</t></rPh></si></sst>`,
  "xl/worksheets/sheet1.xml": `<worksheet ${NS}><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`,
  "xl/worksheets/sheet2.xml":
    `<worksheet ${NS}><cols><col min="1" max="3" width="9"/></cols><sheetData>` +
    `<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="inlineStr"><is><t>TÉCNICOS</t></is></c></row>` +
    `<row r="2"><c r="A2"><f>1+1</f><v>2</v></c><c r="B2" t="str"><f>IF(1,"OK","REVISAR")</f><v>OK</v></c><c r="C2" s="3"/></row>` +
    `<row r="3"><c r="A3" t="b"><v>1</v></c><c r="B3" t="e"><v>#N/A</v></c><c r="C3"><f>A1</f></c></row>` +
    `<row r="5"><c r="A5"><v>0</v></c><c r="B5"><v>-1.5E-2</v></c></row>` +
    `</sheetData></worksheet>`
};

test("leerXlsx lee hojas, nombres definidos de libro y los valores guardados de cada tipo de celda", async () => {
  const lib = await leerXlsx(crearZip(LIBRO, true), { hojas: ["SALIDA_GESPER"] });
  assert.deepEqual(lib.hojas, ["Datos & más", "SALIDA_GESPER"]);
  assert.deepEqual(lib.nombres, { SALIDA_GESPER_DATOS: { hoja: "SALIDA_GESPER", ref: "A1:C3" } });
  const c = lib.celdas.SALIDA_GESPER;
  assert.equal(c.A1, "<=6");                 // cadena compartida con entidad
  assert.equal(c.B1, '14"-60"');             // texto enriquecido, sin la fonética
  assert.equal(c.C1, "TÉCNICOS");            // cadena en línea
  assert.equal(c.A2, 2);                     // fórmula numérica: resultado guardado
  assert.equal(c.B2, "OK");                  // fórmula de texto
  assert.equal(c.A3, true);
  assert.deepEqual(c.B3, { error: "#N/A" });
  assert.equal(c.A5, 0);                     // los ceros se conservan
  assert.equal(c.B5, -0.015);
  assert.ok(!("C2" in c), "una celda solo con formato no tiene valor");
  assert.ok(!("C3" in c), "una fórmula sin resultado guardado no tiene valor");
  assert.equal(lib.celdas["Datos & más"], undefined, "solo se leen las hojas pedidas");
});

test("leerXlsx también lee entradas sin comprimir (método stored)", async () => {
  const lib = await leerXlsx(crearZip(LIBRO, false), { hojas: ["Datos & más"] });
  assert.equal(lib.celdas["Datos & más"].A1, 1);
});

test("leerXlsx rechaza un archivo que no es un ZIP (.xls antiguo, protegido o cualquier otra cosa)", async () => {
  await assert.rejects(() => leerXlsx(Buffer.from("no soy un zip, soy un .xls")), /no es un ZIP/);
});

test("leerXlsx rechaza un ZIP que no es un libro de Excel", async () => {
  await assert.rejects(() => leerXlsx(crearZip({ "hola.txt": "x" }, true)), /no tiene xl\/workbook.xml/);
});

test("direcciones: columnas, celdas y rangos", () => {
  assert.equal(letraAColumna("A"), 1);
  assert.equal(letraAColumna("S"), 19);
  assert.equal(letraAColumna("AA"), 27);
  assert.equal(columnaALetra(21), "U");
  assert.equal(columnaALetra(28), "AB");
  assert.deepEqual(parsearCelda("$M$115"), { columna: "M", col: 13, fila: 115 });
  assert.equal(parsearCelda("M115:N116"), null);
  assert.equal(parsearCelda("115"), null);
  assert.equal(parsearCelda(null), null);
  assert.deepEqual(parsearRango("$A$6:$S$127"), { c1: 1, f1: 6, c2: 19, f2: 127 });
});

test("rangoComoMatriz arma la matriz del rango con null en las celdas vacías", () => {
  assert.deepEqual(rangoComoMatriz({ A1: "x", B2: 0 }, "A1:B2"), [["x", null], [null, 0]]);
});
