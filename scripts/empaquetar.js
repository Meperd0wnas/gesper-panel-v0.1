/**
 * Arma la versión del panel para los usuarios finales (docs/DECISIONES.md §13).
 * No compila nada: copia src/ y assets/ tal cual y genera el manifiesto de
 * producción y el instalador.
 *
 *   dist/sitio/                        lo que publica GitHub Pages
 *     index.html                       página de descarga (instalador/index.html)
 *     src/, assets/                    el panel, sin cambios
 *     manifest.xml                     manifiesto de producción (para TI / Microsoft 365)
 *     Panel-GESPER-instalador.zip      lo que descarga cada usuario
 *
 * Uso:   npm run empaquetar
 *        node scripts/empaquetar.js --verificar-url <url>
 *          (el workflow de Pages lo usa: falla si la URL donde se va a publicar
 *           no es la de package.json, porque el manifiesto apuntaría a otro sitio)
 */
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { manifiestoProduccion, normalizarUrlBase, rellenarPlantilla, aCrlf, crearZip } from "./publicacion.js";

const RAIZ = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const DIST = path.join(RAIZ, "dist");
const SITIO = path.join(DIST, "sitio");
const NOMBRE_ZIP = "Panel-GESPER-instalador.zip";
const INSTALADOR = "Instalar Panel GESPER.cmd";
const DESINSTALADOR = "Desinstalar Panel GESPER.cmd";

function leer(rel) { return fs.readFileSync(path.join(RAIZ, rel), "utf8"); }

function fallar(msg) {
  console.error("\n  No se pudo empaquetar: " + msg + "\n");
  process.exit(1);
}

/** cmd.exe lee el .cmd con la página de códigos de la consola: fuera de ASCII, los acentos salen rotos. */
function cmd(rel, valores) {
  var texto = rellenarPlantilla(leer(rel), valores);
  var i = texto.search(/[^\x00-\x7f]/);
  if (i >= 0) fallar(rel + " tiene un carácter fuera de ASCII («" + texto[i] + "»). Escríbelo sin tilde.");
  return aCrlf(texto);
}

try {
  var pkg = JSON.parse(leer("package.json"));
  var pub = pkg.publicacion || {};
  var urlBase = normalizarUrlBase(pub.url);

  var args = process.argv.slice(2);
  var iv = args.indexOf("--verificar-url");
  if (iv >= 0) {
    var destino = normalizarUrlBase(args[iv + 1]);
    if (destino !== urlBase) {
      fallar("GitHub Pages va a publicar en " + destino + " pero package.json (publicacion.url) dice " + urlBase +
        ". Corrige publicacion.url para que el manifiesto apunte al sitio real.");
    }
  }

  var manifiesto = manifiestoProduccion(leer("manifest.xml"), { urlBase: urlBase, id: pub.idManifiesto, version: pkg.version });
  var valores = { ID: pub.idManifiesto.toUpperCase(), VERSION: pkg.version, URL: urlBase, ZIP: NOMBRE_ZIP };

  var zip = crearZip([
    { nombre: INSTALADOR, datos: cmd("instalador/" + INSTALADOR, valores) },
    { nombre: DESINSTALADOR, datos: cmd("instalador/" + DESINSTALADOR, valores) },
    { nombre: "manifest.xml", datos: manifiesto },
    // BOM: el Bloc de notas de Windows 10 anterior a 1903 no reconoce UTF-8 sin él.
    { nombre: "LEEME.txt", datos: "﻿" + aCrlf(rellenarPlantilla(leer("instalador/LEEME.txt"), valores)) }
  ]);
  var pagina = rellenarPlantilla(leer("instalador/index.html"), valores);

  fs.rmSync(DIST, { recursive: true, force: true });
  fs.mkdirSync(SITIO, { recursive: true });
  fs.cpSync(path.join(RAIZ, "src"), path.join(SITIO, "src"), { recursive: true });
  fs.cpSync(path.join(RAIZ, "assets"), path.join(SITIO, "assets"), { recursive: true });
  fs.writeFileSync(path.join(SITIO, "manifest.xml"), manifiesto);
  fs.writeFileSync(path.join(SITIO, "index.html"), pagina);
  fs.writeFileSync(path.join(SITIO, NOMBRE_ZIP), zip);
} catch (e) {
  fallar(e && e.message ? e.message : String(e));
}

console.log("");
console.log("  Panel GESPER " + pkg.version + " empaquetado");
console.log("  ------------------------------------------------");
console.log("  Carpeta a publicar: " + SITIO);
console.log("  Instalador:         " + path.join(SITIO, NOMBRE_ZIP));
console.log("  Se publicará en:    " + urlBase);
console.log("  Panel:              " + urlBase + "src/taskpane.html");
console.log("");
