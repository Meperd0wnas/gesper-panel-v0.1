/**
 * Servidor local HTTPS para el Panel GESPER.
 *
 * Excel exige que el complemento se sirva por HTTPS con un certificado en el
 * que el equipo confie. Este script usa los certificados de desarrollo de
 * Office (office-addin-dev-certs). Si no estan instalados, lo dice y explica
 * como instalarlos.
 *
 * Uso:   node server.js
 */

import https from "https";
import fs from "fs";
import path from "path";
import os from "os";
import crypto from "crypto";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const PORT = 3000;
const ROOT = __dirname;

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".xml": "text/xml; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon"
};

function fallbackCerts() {
  // Ubicacion por defecto de los certificados de office-addin-dev-certs
  const dir = path.join(os.homedir(), ".office-addin-dev-certs");
  const key = path.join(dir, "localhost.key");
  const crt = path.join(dir, "localhost.crt");
  if (fs.existsSync(key) && fs.existsSync(crt)) {
    return { key: fs.readFileSync(key), cert: fs.readFileSync(crt) };
  }
  return null;
}

async function getCerts() {
  try {
    const devCerts = await import("office-addin-dev-certs");
    const o = await devCerts.getHttpsServerOptions();
    return { key: o.key, cert: o.cert, ca: o.ca };
  } catch (e) {
    const fb = fallbackCerts();
    if (fb) return fb;
    console.error("\n  No encuentro los certificados de desarrollo de Office.\n");
    console.error("  Ejecuta una sola vez, en esta misma carpeta:\n");
    console.error("      npx office-addin-dev-certs install\n");
    console.error("  Windows te va a pedir confirmar que confias en el certificado. Acepta.");
    console.error("  Despues vuelve a ejecutar:  node server.js\n");
    process.exit(1);
  }
}

/*
 * Informes de resultados que el panel pide abrir en el navegador (boton
 * "Abrir en el navegador" de la pantalla Alternativas). Viven solo en memoria:
 * nunca se escriben a disco y se pierden al cerrar el servidor. Se guardan los
 * ultimos MAX_INFORMES, cada uno bajo un id aleatorio.
 */
const INFORMES = new Map();
const MAX_INFORMES = 10;
const MAX_BYTES_INFORME = 20 * 1024 * 1024;

function recibirInforme(req, res) {
  // Exigir text/html obliga al navegador a pedir permiso (CORS) antes de que
  // otra pagina pueda enviar algo aqui; este servidor no lo concede.
  if (!String(req.headers["content-type"] || "").startsWith("text/html")) {
    res.writeHead(415, { "Content-Type": "text/plain; charset=utf-8" }); res.end("Se espera text/html"); return;
  }
  const partes = [];
  let total = 0;
  req.on("data", (trozo) => {
    total += trozo.length;
    if (total > MAX_BYTES_INFORME) {
      res.writeHead(413, { "Content-Type": "text/plain; charset=utf-8" }); res.end("Informe demasiado grande");
      req.destroy();
      return;
    }
    partes.push(trozo);
  });
  req.on("end", () => {
    if (total > MAX_BYTES_INFORME) return;
    const id = crypto.randomUUID();
    INFORMES.set(id, Buffer.concat(partes));
    while (INFORMES.size > MAX_INFORMES) INFORMES.delete(INFORMES.keys().next().value);
    res.writeHead(200, { "Content-Type": "application/json; charset=utf-8", "Cache-Control": "no-store" });
    res.end(JSON.stringify({ ruta: "/informe/" + id }));
  });
}

function servirInforme(id, res) {
  const html = INFORMES.get(id);
  if (!html) {
    res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    res.end("Este informe ya no esta disponible. Vuelve a generarlo desde el panel.");
    return;
  }
  res.writeHead(200, { "Content-Type": "text/html; charset=utf-8", "Cache-Control": "no-store" });
  res.end(html);
}

function serve(req, res) {
  let urlPath;
  try {
    urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  } catch (e) {
    res.writeHead(400); res.end("URL invalida"); return;
  }
  if (urlPath === "/informe" && req.method === "POST") { recibirInforme(req, res); return; }
  if (urlPath.startsWith("/informe/")) { servirInforme(urlPath.slice("/informe/".length), res); return; }
  if (urlPath === "/") urlPath = "/src/taskpane.html";

  const filePath = path.normalize(path.join(ROOT, urlPath));
  if (!filePath.startsWith(ROOT)) { res.writeHead(403); res.end("Prohibido"); return; }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      res.end("No encontrado: " + urlPath);
      return;
    }
    res.writeHead(200, {
      "Content-Type": MIME[path.extname(filePath).toLowerCase()] || "application/octet-stream",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*"
    });
    res.end(data);
  });
}

(async () => {
  const opts = await getCerts();
  https.createServer(opts, serve).listen(PORT, () => {
    console.log("");
    console.log("  Panel GESPER — servidor local activo");
    console.log("  ------------------------------------------------");
    console.log("  Panel:      https://localhost:" + PORT + "/src/taskpane.html");
    console.log("  Manifiesto: " + path.join(ROOT, "manifest.xml"));
    console.log("");
    console.log("  Deja esta ventana abierta mientras uses el panel.");
    console.log("  Para detenerlo: Ctrl + C");
    console.log("");
  });
})();
