/**
 * Pruebas de app/informe.js: el informe de resultados en HTML. Es una función
 * pura (ModeloGesper -> texto), así que se prueba sin Excel ni DOM. Los datos
 * son inventados; no se usa el caso de ningún cliente.
 */
import { test } from "node:test";
import assert from "node:assert/strict";
import { generarInformeHtml, nombreArchivoInforme, escalaEje } from "../src/app/informe.js";

var ANIOS = ["2023", "2024", "2025", "2026", "2027", "2028"];

function series(ipuf) {
  return {
    produccion: [100, 100, 100, 100, 100, 100], facturacion: [70, 70, 70, 70, 70, 70], usuarios: [1000, 1000, 1000, 1000, 1000, 1000],
    vrecProduccion: [0, 0, 0, 5, 6, 7], vrecFacturacion: [0, 0, 0, 1, 1, 1],
    produccionConProgramas: [100, 100, 100, 95, 94, 93], facturacionConProgramas: [70, 70, 70, 71, 71, 71],
    ianc: [30, 30, 30, 25, 24, 23], isuf: [8, 8, 8, 7.9, 7.8, 7.7], icuf: [5.8, 5.8, 5.8, 5.9, 5.9, 5.9], ipuf: ipuf
  };
}

function modeloEjemplo() {
  return {
    estado: {
      empresa: "Acueducto Ejemplo S.A. E.S.P.",
      horizonte: { inicio: "2023", fin: "2028", cantidadAnios: 6 },
      metaIPUF: 2,
      integridad: "ok",
      controles: [{ nombre: "Tabla PROGRAMAS (clave)", estado: "OK" }],
      controlesIntegridad: ["OK"],
      wacc: { real: 0.12, corriente: 0.15 },
      alternativaDesfasada: true,
      programasSinDatos: [],
      ipc: { valor: 0.05, desactualizado: false },
      configuracionMotor: { presupuesto: "$1.000.000", bcMinimo: "1,0", combinacionA1: "1, 9" }
    },
    programas: [
      { numero: 1, nombre: "P.1 Macromedición", tipo: "tecnico", beneficioVPN: 5e8, costoVPN: 1e9, bc: 0.5, estado: "no_viable", nota: "" },
      { numero: 8, nombre: "P.8 Redes provisionales", tipo: "tecnico", beneficioVPN: null, costoVPN: null, bc: null, estado: "no_aplica", nota: "PROGRAMA INCLUIDO DENTRO DEL P19" },
      { numero: 9, nombre: "P.9 Renovación de medidores", tipo: "comercial", beneficioVPN: 3e9, costoVPN: 1e9, bc: 3, estado: "viable", nota: "" }
    ],
    resultadosHidraulicos: null,
    resultadosFinancieros: {
      flujoCajaEncontrado: true, vna: "1.234,5", tir: "N.A.", tirDefinida: false, tirModificada: 0.25, tirModificadaTexto: null, payback: "3",
      parametros: { waccCorriente: "15,00%", cargoFijoAcueducto: "$10", cargoVariableAcueducto: "$2", gradiente: null }, alcantarilladoEnCero: false
    },
    alternativas: [
      { id: 1, beneficio: 3.5e9, costo: 2e9, bc: 1.75, ipufPromedio: 2.5, cumpleMeta: false, composicionTexto: "1, 9" },
      // B/C deliberadamente distinto de beneficio ÷ costo: el informe debe mostrar el del modelo, no recalcularlo.
      { id: 2, beneficio: 3e9, costo: 1e9, bc: 1.234, ipufPromedio: 1.5, cumpleMeta: true, composicionTexto: "9" },
      { id: 3, beneficio: 0, costo: 0, bc: null, ipufPromedio: null, cumpleMeta: false, composicionTexto: "" }
    ],
    detalleAlternativas: {
      anios: ANIOS, inicioPromedio: 3,
      alternativas: [
        { id: 1, programas: [1, 9], veredicto: "La ALTERNATIVA No. 1 ... NO CUMPLE", ipufPromedio: 2.5, nota: "", series: series([2.2, 2.2, 2.2, 2.6, 2.5, 2.4]) },
        { id: 2, programas: [9], veredicto: "La ALTERNATIVA No. 2 ... CUMPLE", ipufPromedio: 1.5, nota: "", series: series([2.2, 2.2, 2.2, 1.6, 1.5, 1.4]) },
        { id: 3, programas: [], veredicto: "", ipufPromedio: 2.9, nota: "", series: series([2.2, 2.2, 2.2, 2.8, 2.9, 3.0]) }
      ]
    },
    planGestion: {
      plan: { descripcion: "Plan de prueba", requisitoEsencial: ["1. Macromedición"], programasTexto: "1, 9", programas: [1, 9], costo: 2e9, beneficio: 3.5e9, bc: 1.75, volumenRecuperado: 2e6 },
      combinaciones: [{ programasTexto: "1, 9", programas: [1, 9], bc: 1.75, volumenRecuperado: 2e6, costo: 2e9, beneficio: 3.5e9 }],
      clasificacion: { criterios: ["Obligatorios", "Mayor Volumen recuperado", "Mayor Relación B/C"], grupos: [{ nombre: "Técnicos", columnas: [["P.1 Macromedición"], [], []] }] }
    }
  };
}

var FECHA = new Date(2026, 10, 24, 10, 30);

test("generarInformeHtml produce un documento completo con todas las secciones", () => {
  var html = generarInformeHtml(modeloEjemplo(), { fecha: FECHA });
  assert.match(html, /^<!DOCTYPE html>/);
  assert.match(html, /<title>Informe de resultados GESPER · Acueducto Ejemplo S\.A\. E\.S\.P\.<\/title>/);
  ["resumen", "estado", "comparacion", "indicadores", "programas", "plan", "financiero", "notas", "anexo"].forEach(function (id) {
    assert.ok(html.indexOf('id="' + id + '"') >= 0, "falta la sección " + id);
  });
  for (var f = 1; f <= 6; f++) assert.ok(html.indexOf('id="fig' + f + '"') >= 0, "falta la figura " + f);
});

test("cada gráfica de líneas tiene sus datos para el tooltip, y el JSON se puede leer", () => {
  var html = generarInformeHtml(modeloEjemplo(), { fecha: FECHA });
  var graficas = html.match(/<svg class="graf"/g) || [];
  assert.equal(graficas.length, 7, "IPUF, IANC, 2 de volumen y 3 de ISUF/ICUF");
  var json = html.match(/<script type="application\/json" id="datos-graficas">([\s\S]*?)<\/script>/)[1];
  var datos = JSON.parse(json);
  assert.equal(datos.length, 7);
  assert.deepEqual(datos[0].anios, ANIOS);
  assert.deepEqual(datos[0].series.map(function (s) { return s.nombre; }), ["Alternativa 1", "Alternativa 2", "Alternativa 3"]);
  assert.deepEqual(datos[0].series[1].valores, [2.2, 2.2, 2.2, 1.6, 1.5, 1.4]);
});

test("el texto que viene del libro se escapa, también dentro del JSON de las gráficas", () => {
  var m = modeloEjemplo();
  m.estado.empresa = '<img src=x onerror="alert(1)">';
  m.programas[0].nombre = "P.1 <script>alert(1)</script>";
  m.detalleAlternativas.anios = ["2023</script><script>alert(1)", "2024", "2025", "2026", "2027", "2028"];
  var html = generarInformeHtml(m, { fecha: FECHA });
  assert.ok(html.indexOf("<img src=x") < 0);
  assert.ok(html.indexOf("&lt;img src=x onerror=&quot;alert(1)&quot;&gt;") >= 0);
  assert.ok(html.indexOf("<script>alert(1)") < 0);
  assert.equal((html.match(/<\/script>/g) || []).length, 2, "solo cierran los dos <script> propios del documento");
});

test("muestra las cifras del modelo tal cual, sin recalcularlas", () => {
  var html = generarInformeHtml(modeloEjemplo(), { fecha: FECHA });
  // Celda B/C de la Alternativa 2 en la tabla comparativa (columna c2).
  assert.ok(html.indexOf('<td class="n c2"><b>1,23</b></td>') >= 0, "B/C de la Alternativa 2 tal como viene (1,234), no 3e9 / 1e9");
  assert.ok(html.indexOf('<td class="n c2"><b>3,00</b></td>') < 0);
});

test("explica el veredicto de cada alternativa y la que no tiene programas", () => {
  var html = generarInformeHtml(modeloEjemplo(), { fecha: FECHA });
  assert.match(html, /Alternativa 2<\/span> <span class="nota-in">\(compuesta a mano · alimenta el consolidado financiero\)<\/span>: 1 programa/);
  assert.ok(html.indexOf("Cumple la meta de la empresa (2,0)") >= 0);
  assert.ok(html.indexOf("No cumple la meta de la empresa (2,0)") >= 0);
  assert.ok(html.indexOf("no tiene programas marcados en la hoja") >= 0);
  assert.ok(html.indexOf("Programa incluido dentro del P19") >= 0, "las notas en mayúsculas sostenidas se pasan a oración");
});

test("avisa cuando los controles de integridad no están en OK", () => {
  var m = modeloEjemplo();
  m.estado.integridad = "motor_sin_ejecutar";
  var html = generarInformeHtml(m, { fecha: FECHA });
  assert.ok(html.indexOf("El motor de optimización no se ha ejecutado") >= 0);
  m.estado.integridad = "revisar";
  assert.ok(generarInformeHtml(m, { fecha: FECHA }).indexOf("Revisar los controles de integridad del modelo") >= 0);
});

test("con logo usa la imagen; sin logo escribe el nombre de la marca", () => {
  var con = generarInformeHtml(modeloEjemplo(), { fecha: FECHA, logo: "data:image/png;base64,AAAA" });
  assert.ok(con.indexOf('<img src="data:image/png;base64,AAAA" alt="Dual Ingeniería Hidráulica">') >= 0);
  var sin = generarInformeHtml(modeloEjemplo(), { fecha: FECHA });
  assert.ok(sin.indexOf('<span class="marca-txt">Dual Ingeniería Hidráulica</span>') >= 0);
});

test("no lanza con un modelo vacío (libro sin las hojas esperadas)", () => {
  var html = generarInformeHtml({}, { fecha: FECHA });
  assert.match(html, /<\/html>\n$/);
  assert.ok(html.indexOf("No se pudieron leer los años del horizonte") >= 0);
  assert.ok(html.indexOf("No se pudo leer la tabla de programas") >= 0);
});

test("nombreArchivoInforme arma un nombre sin tildes, puntos ni espacios", () => {
  assert.equal(nombreArchivoInforme("amb S.A. E.S.P.", FECHA), "Informe_GESPER_amb_SA_ESP_2026-11-24.html");
  assert.equal(nombreArchivoInforme("Acueducto de Bogotá", FECHA), "Informe_GESPER_Acueducto_de_Bogota_2026-11-24.html");
  assert.equal(nombreArchivoInforme("", new Date(2026, 0, 5)), "Informe_GESPER_2026-01-05.html");
});

test("escalaEje elige pasos redondos y los decimales que necesitan", () => {
  var a = escalaEje(3.54, 6, 5);
  assert.deepEqual([a.min, a.max, a.decimales], [3.5, 6, 1]);
  assert.deepEqual(a.marcas, [3.5, 4, 4.5, 5, 5.5, 6]);
  var b = escalaEje(0, 8.9, 4);
  assert.deepEqual([b.min, b.max, b.decimales], [0, 10, 1]);
  assert.deepEqual(b.marcas, [0, 2.5, 5, 7.5, 10]);
  var c = escalaEje(4, 4, 5);
  assert.ok(c.max > c.min, "un rango sin variación no divide por cero");
  var d = escalaEje(Infinity, -Infinity, 5);
  assert.deepEqual([d.min, d.max], [0, 1], "sin datos, una escala por defecto");
});
