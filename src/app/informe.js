/**
 * Informe de resultados en HTML para descargar: la comparación de las tres
 * alternativas de la hoja Alternativas, la evolución de sus indicadores, los
 * 19 programas, el plan de gestión de pérdidas y los resultados financieros
 * del consolidado, con cada gráfica explicada.
 *
 * Función pura: recibe el ModeloGesper ya mapeado (domain/repositories/
 * modeloRepository.js) y devuelve el documento como texto. No toca el DOM,
 * ni `state`, ni Excel, así que se prueba con Node (test/informe.test.js).
 *
 * El documento es autocontenido (CSS, gráficas en SVG y un script corto de
 * tooltips, todo en línea; sin librerías ni recursos externos): se abre igual
 * en el navegador, sin conexión, y se imprime o guarda como PDF.
 *
 * No recalcula el modelo. Las cifras se leen tal cual del libro; lo único que
 * hace este archivo es presentarlas (pasar pesos a millones, m³ a millones de
 * m³) y describirlas (primer y último año, máximos, quién cumple la meta). El
 * IPUF promedio viene del mapper, con el mismo tramo de años que usa la hoja
 * para su veredicto (T22/T63/T105).
 */
import { LECTURAS } from "../config/cellMap.js";
import { esc, num, fmt, parseNum } from "../utils/format.js";

/**
 * Un color por alternativa, el mismo en todo el informe. Son los tres primeros
 * colores categóricos de la paleta de referencia (skill dataviz), validados
 * sobre fondo blanco para todos los pares, incluido daltonismo. El azul de la
 * marca (#063F5C) no sirve como color de serie: es demasiado oscuro y poco
 * saturado; queda para títulos y para los programas en la figura 6.
 */
var COLOR_ALT = { 1: "#2a78d6", 2: "#eb6834", 3: "#1baf7a" };
var COLOR_PROGRAMA = "#063F5C";
var COLOR_GRIS = "#9AA7A3";
var MAXIMO_NORMATIVO_IPUF = 6; // m³/usuario/mes, Res. 330/2017 art. 9 (el mismo que muestra la pantalla Estado)

var ROL_ALT = {
  1: "Propuesta por el motor de optimización",
  2: "Compuesta a mano · alimenta el consolidado financiero",
  3: "Compuesta a mano"
};

/* ---------- formato ---------- */

function millones(v) { return num(v) === null ? "—" : fmt(v / 1e6, 0); }
function millonesM3(v, d) { return num(v) === null ? "—" : fmt(v / 1e6, d === undefined ? 2 : d); }
function plural(n, uno, varios) { return n + " " + (n === 1 ? uno : varios); }
function listaP(nums) { return nums.length ? nums.map(function (n) { return "P." + n; }).join(", ") : "ninguno"; }
function nombrePrograma(p) { return String(p.nombre || "").replace(/^P\.\s*\d+\s*/, "").trim() || "Programa " + p.numero; }
/** Texto de la hoja escrito todo en mayúsculas -> oración ("PROGRAMA INCLUIDO DENTRO DEL P19" -> "Programa incluido dentro del P19"). */
function oracion(t) {
  t = String(t || "").trim();
  if (!t || t !== t.toUpperCase() || !/[A-ZÁÉÍÓÚÑ]{3}/.test(t)) return t;
  t = t.toLowerCase().replace(/\bp\.?\s?(\d+)/g, function (m, n) { return "P" + n; });
  return t.charAt(0).toUpperCase() + t.slice(1);
}
function fuente(clave) {
  for (var i = 0; i < LECTURAS.length; i++) if (LECTURAS[i][0] === clave) return LECTURAS[i][1] + "!" + LECTURAS[i][2];
  return "";
}
/** JSON que se puede incrustar en un <script> sin cerrarlo antes de tiempo. */
function jsonSeguro(o) { return JSON.stringify(o).replace(/</g, "\\u003c"); }
function fechaLarga(d) { return d.toLocaleDateString("es-CO", { day: "numeric", month: "long", year: "numeric" }); }
function hora(d) { return d.toLocaleTimeString("es-CO", { hour: "2-digit", minute: "2-digit" }); }
function dos(n) { return (n < 10 ? "0" : "") + n; }

/**
 * Nombre del archivo descargado: Informe_GESPER_<empresa>_<AAAA-MM-DD>.html,
 * sin tildes ni caracteres que Windows no acepte.
 */
export function nombreArchivoInforme(empresa, fecha) {
  var d = fecha || new Date();
  var slug = String(empresa || "").normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\./g, "")
    .replace(/[^A-Za-z0-9]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 40);
  return "Informe_GESPER_" + (slug ? slug + "_" : "") + d.getFullYear() + "-" + dos(d.getMonth() + 1) + "-" + dos(d.getDate()) + ".html";
}

/* ---------- series ---------- */

function serie(a, clave) { return (a.series && a.series[clave]) || []; }
function primero(s) { for (var i = 0; i < s.length; i++) if (num(s[i]) !== null) return { i: i, v: s[i] }; return null; }
function ultimo(s) { for (var i = s.length - 1; i >= 0; i--) if (num(s[i]) !== null) return { i: i, v: s[i] }; return null; }
function maximo(s) {
  var m = null;
  for (var i = 0; i < s.length; i++) if (num(s[i]) !== null && (m === null || s[i] > m.v)) m = { i: i, v: s[i] };
  return m;
}
function primerPositivo(s) { for (var i = 0; i < s.length; i++) if (num(s[i]) !== null && s[i] > 0) return i; return -1; }
function escalar(s, factor) { return s.map(function (v) { return num(v) === null ? null : v / factor; }); }

/* ---------- datos del informe ---------- */

function armarAlternativas(modelo) {
  var resumen = modelo.alternativas || [];
  var detalle = (modelo.detalleAlternativas && modelo.detalleAlternativas.alternativas) || [];
  return [1, 2, 3].map(function (id) {
    var r = resumen.filter(function (a) { return a.id === id; })[0] || {};
    var d = detalle.filter(function (a) { return a.id === id; })[0] || {};
    var programas = d.programas || [];
    return {
      id: id,
      nombre: "Alternativa " + id,
      corto: "Alt. " + id,
      color: COLOR_ALT[id],
      rol: ROL_ALT[id],
      beneficio: num(r.beneficio),
      costo: num(r.costo),
      bc: num(r.bc),
      cumpleMeta: r.cumpleMeta === true || r.cumpleMeta === false ? r.cumpleMeta : null,
      composicionTexto: r.composicionTexto || "",
      programas: programas,
      veredicto: d.veredicto || "",
      ipufPromedio: num(d.ipufPromedio),
      nota: d.nota || "",
      series: d.series || {},
      vacia: programas.length === 0
    };
  });
}

function contexto(modelo, opciones) {
  var det = modelo.detalleAlternativas || {};
  var anios = det.anios && det.anios.length ? det.anios : [];
  var ini = det.inicioPromedio === undefined ? 3 : det.inicioPromedio;
  var estado = modelo.estado || {};
  return {
    fecha: opciones.fecha || new Date(),
    logo: opciones.logo || null,
    version: opciones.version || "v0.1",
    nombreArchivo: nombreArchivoInforme(estado.empresa, opciones.fecha),
    estado: estado,
    alts: armarAlternativas(modelo),
    programas: modelo.programas || [],
    fin: modelo.resultadosFinancieros || null,
    plan: modelo.planGestion || { plan: {}, combinaciones: [], clasificacion: { criterios: [], grupos: [] } },
    anios: anios,
    inicioPromedio: ini,
    periodo: anios.length > ini ? anios[ini] + "–" + anios[anios.length - 1] : "del 4.º año al final del horizonte",
    meta: num(estado.metaIPUF),
    graficas: [],      // datos de cada gráfica de líneas, para el script de tooltips
    figura: 0
  };
}

/* ---------- piezas de HTML ---------- */

function punto(color) { return '<span class="pt" style="background:' + color + '"></span>'; }
function chipAlt(a) { return '<span class="chip-alt">' + punto(a.color) + esc(a.nombre) + '</span>'; }

function veredictoChip(cumple, meta) {
  if (cumple === null) return '<span class="estado nd">Sin veredicto</span>';
  return '<span class="estado ' + (cumple ? "ok" : "no") + '"><span aria-hidden="true">' + (cumple ? "✓" : "✕") + '</span> ' +
    (cumple ? "Cumple" : "No cumple") + (meta !== null && meta !== undefined ? " la meta (" + fmt(meta, 1) + ")" : "") + '</span>';
}

function aviso(tipo, titulo, texto) {
  var icono = { ok: "✓", info: "i", wn: "!", cr: "✕" }[tipo];
  return '<div class="aviso ' + tipo + '"><span class="ic" aria-hidden="true">' + icono + '</span><div><div class="at">' + titulo +
    '</div>' + (texto ? '<div class="ad">' + texto + '</div>' : '') + '</div></div>';
}

function kpi(etiqueta, valor, unidad, sub) {
  return '<div class="kpi"><div class="l">' + etiqueta + '</div><div class="v">' + valor +
    (unidad ? '<span class="u">' + unidad + '</span>' : '') + '</div>' + (sub ? '<div class="s">' + sub + '</div>' : '') + '</div>';
}

function figura(ctx, titulo, queMuestra, cuerpo, comoLeer, lectura, fuenteTexto) {
  ctx.figura++;
  return '<figure id="fig' + ctx.figura + '"><figcaption><div class="fn">Figura ' + ctx.figura + '</div><div class="ft">' + titulo +
    '</div><div class="fq">' + queMuestra + '</div></figcaption>' + cuerpo +
    '<div class="explica"><div><h4>Cómo leerla</h4>' + comoLeer + '</div><div><h4>Lectura de este caso</h4>' +
    (lectura.length ? '<ul>' + lectura.map(function (l) { return '<li>' + l + '</li>'; }).join("") + '</ul>' : '<p>Sin datos suficientes.</p>') +
    '</div></div>' + (fuenteTexto ? '<div class="fuente">Fuente: ' + fuenteTexto + '</div>' : '') + '</figure>';
}

function leyenda(items) {
  return '<div class="leyenda">' + items.map(function (it) {
    var key;
    if (it.tipo === "banda") key = '<span class="kb"></span>';
    else if (it.tipo === "area") key = '<span class="ka" style="background:' + it.color + '"></span>';
    else if (it.tipo === "barra") key = '<span class="kr" style="background:' + it.color + '"></span>';
    else key = '<svg width="20" height="8" aria-hidden="true"><line x1="1" y1="4" x2="19" y2="4" stroke="' + it.color +
      '" stroke-width="' + (it.guion ? 1.5 : 2.5) + '"' + (it.guion ? ' stroke-dasharray="4 3"' : '') + ' stroke-linecap="round"/></svg>';
    return '<span class="li">' + key + it.texto + '</span>';
  }).join("") + '</div>';
}

/* ---------- gráfica de líneas (SVG) ---------- */

/** Escala "redonda" para el eje Y: paso de 1, 2, 2,5 o 5 por potencia de 10. */
export function escalaEje(lo, hi, marcas) {
  if (!(isFinite(lo) && isFinite(hi))) { lo = 0; hi = 1; }
  if (lo === hi) { lo = lo - 1; hi = hi + 1; }
  var bruto = (hi - lo) / (marcas || 5);
  var e = Math.pow(10, Math.floor(Math.log10(bruto)));
  var r = bruto / e;
  var paso = (r <= 1 ? 1 : r <= 2 ? 2 : r <= 2.5 ? 2.5 : r <= 5 ? 5 : 10) * e;
  var min = Math.floor(lo / paso + 1e-9) * paso;
  var max = Math.ceil(hi / paso - 1e-9) * paso;
  var lista = [];
  for (var v = min; v <= max + paso / 2; v += paso) lista.push(Number(v.toFixed(10)));
  var dec = 0;
  while (dec < 4 && Math.abs(Math.round(paso * Math.pow(10, dec)) - paso * Math.pow(10, dec)) > 1e-6) dec++;
  return { min: min, max: max, marcas: lista, decimales: dec };
}

function rangoDe(series, extras) {
  var lo = Infinity, hi = -Infinity;
  series.forEach(function (s) { s.forEach(function (v) { if (num(v) !== null) { lo = Math.min(lo, v); hi = Math.max(hi, v); } }); });
  (extras || []).forEach(function (v) { if (num(v) !== null) { lo = Math.min(lo, v); hi = Math.max(hi, v); } });
  return { lo: lo, hi: hi };
}

/**
 * @param {Object} ctx
 * @param {Object} o
 *  anios, series [{nombre, etiqueta, color, valores}], unidad, decimales,
 *  refs [{valor, etiqueta, color}], banda {desde, etiqueta}, area {sup, inf, color},
 *  escala (compartida entre paneles), desdeCero, ancho, alto, margenDer, fuente (tamaño), aria
 */
function graficaLineas(ctx, o) {
  var W = o.ancho || 880, H = o.alto || 330, fs = o.fuente || 12;
  var anios = o.anios, n = anios.length;
  if (n < 2) return '<p class="nota">Sin años suficientes para graficar.</p>';
  var x0 = o.margenIzq || 48, x1 = W - (o.margenDer || 112), y0 = 14, y1 = H - 26;
  var es = o.escala;
  if (!es) {
    var r = rangoDe(o.series.map(function (s) { return s.valores; }), (o.refs || []).map(function (rf) { return rf.valor; }));
    if (o.desdeCero) r.lo = Math.min(0, r.lo);
    es = escalaEje(r.lo, r.hi, o.marcas || 5);
  }
  var dx = (x1 - x0) / (n - 1);
  function X(i) { return x0 + dx * i; }
  function Y(v) { return y1 - (y1 - y0) * (v - es.min) / (es.max - es.min); }
  function r1(v) { return Math.round(v * 10) / 10; }

  var id = ctx.graficas.length;
  var s = '<svg class="graf" viewBox="0 0 ' + W + ' ' + H + '" role="img" tabindex="0" data-graf="' + id + '" aria-label="' + esc(o.aria || "") + '"' +
    ' font-size="' + fs + '">';

  if (o.banda && o.banda.desde > 0 && o.banda.desde < n) {
    var bx = X(o.banda.desde) - dx / 2;
    s += '<rect x="' + r1(bx) + '" y="' + y0 + '" width="' + r1(x1 - bx) + '" height="' + (y1 - y0) + '" fill="#F0F4F3"/>';
    if (o.banda.etiqueta) s += '<text x="' + r1(bx + 6) + '" y="' + (y0 + fs + 2) + '" fill="#6A7975">' + esc(o.banda.etiqueta) + '</text>';
  }

  es.marcas.forEach(function (m, k) {
    var y = r1(Y(m));
    s += '<line x1="' + x0 + '" x2="' + x1 + '" y1="' + y + '" y2="' + y + '" stroke="' + (k === 0 ? "#C9D3D0" : "#E8EEEC") + '" stroke-width="1"/>';
    s += '<text x="' + (x0 - 8) + '" y="' + r1(y + fs / 3) + '" text-anchor="end" fill="#6A7975" class="tn">' + fmt(m, es.decimales) + '</text>';
  });

  var cada = Math.max(1, Math.ceil(n * (fs * 2.9) / (x1 - x0)));
  for (var i = 0; i < n; i++) {
    if ((n - 1 - i) % cada !== 0) continue;
    s += '<text x="' + r1(X(i)) + '" y="' + (H - 6) + '" text-anchor="middle" fill="#6A7975" class="tn">' + esc(anios[i]) + '</text>';
  }

  if (o.area) {
    var arriba = [], abajo = [];
    for (var j = 0; j < n; j++) {
      if (num(o.area.sup[j]) === null || num(o.area.inf[j]) === null) continue;
      arriba.push(r1(X(j)) + "," + r1(Y(o.area.sup[j])));
      abajo.unshift(r1(X(j)) + "," + r1(Y(o.area.inf[j])));
    }
    if (arriba.length > 1) s += '<polygon points="' + arriba.concat(abajo).join(" ") + '" fill="' + o.area.color + '" fill-opacity="0.13"/>';
  }

  (o.refs || []).forEach(function (rf) {
    if (num(rf.valor) === null) return;
    var y = r1(Y(rf.valor));
    s += '<line x1="' + x0 + '" x2="' + x1 + '" y1="' + y + '" y2="' + y + '" stroke="' + rf.color + '" stroke-width="1.5" stroke-dasharray="6 4"/>';
    s += '<text x="' + (x0 + 6) + '" y="' + (y - 5) + '" fill="#3D4A46" class="halo">' + esc(rf.etiqueta) + '</text>';
  });

  o.series.forEach(function (se) {
    var d = "", abierto = false;
    se.valores.forEach(function (v, k) {
      if (num(v) === null) { abierto = false; return; }
      d += (abierto ? "L" : "M") + r1(X(k)) + " " + r1(Y(v));
      abierto = true;
    });
    if (d) s += '<path d="' + d + '" fill="none" stroke="' + se.color + '" stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>';
  });

  // Etiquetas al final de cada línea. Si dos chocan se separan y se unen a su
  // línea con un conector, en vez de apilarse sueltas.
  var etq = [];
  o.series.forEach(function (se) {
    var p = ultimo(se.valores);
    if (!p) return;
    etq.push({ x: X(p.i), y: Y(p.v), yObj: Y(p.v), color: se.color, nombre: se.etiqueta || se.nombre, valor: fmt(p.v, o.decimales) });
  });
  etq.sort(function (a, b) { return a.yObj - b.yObj; });
  var gap = fs + 4;
  for (var a = 1; a < etq.length; a++) if (etq[a].y - etq[a - 1].y < gap) etq[a].y = etq[a - 1].y + gap;
  if (etq.length && etq[etq.length - 1].y > y1) {
    etq[etq.length - 1].y = y1;
    for (var b = etq.length - 2; b >= 0; b--) if (etq[b + 1].y - etq[b].y < gap) etq[b].y = etq[b + 1].y - gap;
  }
  etq.forEach(function (e) {
    s += '<circle cx="' + r1(e.x) + '" cy="' + r1(e.yObj) + '" r="4" fill="' + e.color + '" stroke="#fff" stroke-width="2"/>';
    if (Math.abs(e.y - e.yObj) > 1.5) {
      s += '<polyline points="' + r1(e.x + 6) + ',' + r1(e.yObj) + ' ' + r1(x1 + 10) + ',' + r1(e.y) + '" fill="none" stroke="#96A29E" stroke-width="1"/>';
    }
    s += '<text x="' + (x1 + 13) + '" y="' + r1(e.y + fs / 3) + '"><tspan fill="#3D4A46" font-weight="600">' + esc(e.nombre) +
      '</tspan><tspan fill="#1A2320" dx="5" class="tn">' + e.valor + '</tspan></text>';
  });

  s += '<line class="xh" x1="0" x2="0" y1="' + y0 + '" y2="' + y1 + '" stroke="#3D4A46" stroke-width="1" visibility="hidden" pointer-events="none"/>';
  s += '<rect class="hit" x="' + x0 + '" y="' + y0 + '" width="' + (x1 - x0) + '" height="' + (y1 - y0) + '" fill="transparent"/>';
  s += '</svg>';

  ctx.graficas.push({
    id: id, anios: anios, x0: x0, dx: dx, y0: y0, y1: y1, min: es.min, max: es.max,
    dec: o.decimales, unidad: o.unidad || "",
    series: o.series.map(function (se) { return { nombre: se.nombre, color: se.color, valores: se.valores.map(function (v) { return num(v); }) }; })
  });
  return s;
}

/* ---------- secciones ---------- */

function portada(ctx) {
  var e = ctx.estado, cfg = e.configuracionMotor || {};
  var h = '<header class="portada"><div class="top">' +
    (ctx.logo ? '<img src="' + esc(ctx.logo) + '" alt="Dual Ingeniería Hidráulica">' : '<span class="marca-txt">Dual Ingeniería Hidráulica</span>') +
    '<div class="eyebrow">Informe de resultados · GESPER</div></div>' +
    '<h1>Alternativas de reducción de pérdidas de agua</h1>' +
    '<div class="empresa">' + esc(e.empresa || "Caso sin identificar") + '</div>' +
    '<div class="generado">Generado el ' + esc(fechaLarga(ctx.fecha)) + ' a las ' + esc(hora(ctx.fecha)) + ' con el Panel GESPER ' + esc(ctx.version) + '</div>' +
    '<dl class="ficha">' +
    '<div><dt>Horizonte</dt><dd>' + (e.horizonte ? esc(e.horizonte.inicio) + "–" + esc(e.horizonte.fin) + ' · ' + e.horizonte.cantidadAnios + ' años' : "—") + '</dd></div>' +
    '<div><dt>Meta de IPUF de la empresa</dt><dd>' + (ctx.meta !== null ? fmt(ctx.meta, 1) + ' m³/usuario/mes' : "—") + '</dd></div>' +
    '<div><dt>Presupuesto disponible</dt><dd>' + (esc(cfg.presupuesto) || "—") + '</dd></div>' +
    '<div><dt>B/C mínimo exigido al motor</dt><dd>' + (esc(cfg.bcMinimo) || "—") + '</dd></div>' +
    '</dl></header>';
  return h;
}

function indice() {
  var items = [
    ["resumen", "Resumen ejecutivo"], ["estado", "Estado del modelo"], ["comparacion", "Comparación de las alternativas"],
    ["indicadores", "Evolución de los indicadores"], ["programas", "Programas"], ["plan", "Plan de gestión de pérdidas"],
    ["financiero", "Resultados financieros del consolidado"], ["notas", "Notas metodológicas y fuentes"], ["anexo", "Anexo A · Balance hídrico por alternativa"]
  ];
  return '<nav class="indice" aria-label="Contenido"><div class="eyebrow">Contenido</div><ol>' +
    items.map(function (it, i) {
      return '<li><span class="in">' + (it[0] === "anexo" ? "A" : i + 1) + '</span><a href="#' + it[0] + '">' + it[1].replace(/^Anexo A · /, "Anexo · ") + '</a></li>';
    }).join("") + '</ol></nav>';
}

function cumpleFrase(a, ctx) {
  if (a.cumpleMeta === null) return "La hoja no da un veredicto sobre la meta";
  return (a.cumpleMeta ? "Cumple" : "No cumple") + " la meta de la empresa" + (ctx.meta !== null ? " (" + fmt(ctx.meta, 1) + ")" : "");
}

function fraseAlternativa(a, ctx) {
  if (a.vacia) return chipAlt(a) + ' no tiene programas marcados en la hoja.';
  return chipAlt(a) + ' <span class="nota-in">(' + esc(a.rol.toLowerCase()) + ')</span>: ' + plural(a.programas.length, "programa", "programas") +
    ', relación B/C de <b>' + fmt(a.bc, 2) + '</b> e IPUF promedio de <b>' + fmt(a.ipufPromedio, 2) + '</b> m³/usuario/mes en ' + ctx.periodo + '. ' +
    cumpleFrase(a, ctx) + '.';
}

function seccionResumen(ctx) {
  var a2 = ctx.alts[1], fin = ctx.fin;
  var ianc = serie(a2, "ianc"), iancIni = primero(ianc), iancFin = ultimo(ianc);
  var h = '<section id="resumen"><h2><span class="num">1</span>Resumen ejecutivo</h2>';
  h += '<p class="lead">Este informe resume los resultados de la hoja <b>Alternativas</b> del modelo GESPER' +
    (ctx.estado.empresa ? ' para <b>' + esc(ctx.estado.empresa) + '</b>' : '') + ': las tres alternativas de inversión en programas de reducción de pérdidas, ' +
    'la evolución de sus indicadores' + (ctx.anios.length ? ' entre ' + esc(ctx.anios[0]) + ' y ' + esc(ctx.anios[ctx.anios.length - 1]) : '') +
    ' y la evaluación de los ' + ctx.programas.length + ' programas.</p>';

  h += '<div class="kpis">';
  h += kpi("Relación B/C · Alternativa 2", fmt(a2.bc, 2), "",
    "Beneficio " + millones(a2.beneficio) + " · costo " + millones(a2.costo) + " millones $");
  h += kpi("IPUF promedio " + ctx.periodo + " · Alternativa 2", fmt(a2.ipufPromedio, 2), "m³/usuario/mes", veredictoChip(a2.cumpleMeta, ctx.meta));
  h += kpi("IANC " + (iancFin ? esc(ctx.anios[iancFin.i] || "") : "") + " · Alternativa 2", iancFin ? fmt(iancFin.v, 1) : "—", "%",
    iancIni ? "En " + esc(ctx.anios[iancIni.i] || "") + ": " + fmt(iancIni.v, 1) + " %" : "");
  if (fin && fin.flujoCajaEncontrado) {
    h += kpi("VNA del proyecto", esc(fin.vna), "millones $", "TIR " + esc(fin.tir) +
      (fin.payback ? " · recuperación: " + esc(fin.payback) + (String(fin.payback).trim() === "1" ? " año" : " años") : ""));
  } else {
    h += kpi("VNA del proyecto", "—", "", "No se encontró la hoja de flujo de caja");
  }
  h += '</div><p class="nota">Las cifras destacadas son de la Alternativa 2: es la que alimenta el consolidado financiero (fila 49 de la hoja Alternativas).</p>';

  h += '<ul class="frases">' + ctx.alts.map(function (a) { return '<li>' + fraseAlternativa(a, ctx) + '</li>'; }).join("") + '</ul>';
  return h + '</section>';
}

function seccionEstado(ctx) {
  var e = ctx.estado;
  var h = '<section id="estado"><h2><span class="num">2</span>Estado del modelo al generar el informe</h2>';
  var controles = (e.controles || []).map(function (c) { return esc(c.nombre) + ': <b>' + esc(c.estado || "—") + '</b>'; }).join(" · ");
  if (e.integridad === "ok") h += aviso("ok", "Controles de integridad del modelo en OK", controles);
  else if (e.integridad === "motor_sin_ejecutar") h += aviso("wn", "El motor de optimización no se ha ejecutado", controles + ". La Alternativa 1 puede no corresponder a la combinación óptima.");
  else h += aviso("cr", "Revisar los controles de integridad del modelo", (controles || "No se pudieron leer los controles.") + " Los resultados de este informe pueden no ser fiables.");

  h += aviso("info", "Los resultados financieros corresponden a la Alternativa 2",
    "El consolidado (las 19 evaluaciones, el bloque tarifario y el flujo de caja) se calcula sobre la fila 49 de la hoja Alternativas. " +
    (e.alternativaDesfasada ? "La Alternativa 1, la que propuso el motor, incluye programas distintos y <b>no afecta</b> esos resultados. " : "Hoy coincide con la selección de la Alternativa 1. ") +
    "Si la selección cambió y no se volvió a ejecutar el motor, el consolidado puede estar desactualizado: Excel no lo advierte.");

  if (e.wacc && (e.wacc.real === null || e.wacc.real === 0)) {
    h += aviso("cr", "Falta el WACC de la empresa", "Con <code>PARÁMETROS!E8</code> vacía, el modelo descuenta a la tasa de inflación y no al WACC real.");
  }
  if (e.ipc && e.ipc.desactualizado) {
    h += aviso("wn", "Series de referencia desactualizadas", "El modelo proyecta con un IPC de " + fmt(e.ipc.valor * 100, 2) +
      " %. El DANE cerró 2025 en 5,10 %: el horizonte se indexa por debajo de la inflación real.");
  }
  var sin = e.programasSinDatos || [];
  if (sin.length) {
    h += aviso("wn", plural(sin.length, "programa sin datos", "programas sin datos"),
      esc(listaP(sin.map(function (p) { return p.numero; }))) + " tienen beneficio y costo en 0. No se distingue si no aplican o si falta diligenciarlos.");
  }
  return h + '</section>';
}

function barrasBeneficioCosto(ctx) {
  var max = 0;
  ctx.alts.forEach(function (a) { max = Math.max(max, a.beneficio || 0, a.costo || 0); });
  var h = leyenda([
    { tipo: "barra", color: "linear-gradient(90deg," + COLOR_ALT[1] + " 33%," + COLOR_ALT[2] + " 33% 66%," + COLOR_ALT[3] + " 66%)", texto: "Beneficio (color de cada alternativa)" },
    { tipo: "barra", color: COLOR_GRIS, texto: "Costo" }
  ]);
  h += '<div class="barras">';
  ctx.alts.forEach(function (a) {
    if (a.vacia && !a.beneficio && !a.costo) {
      h += '<div class="bfila"><div class="bnom">' + chipAlt(a) + '</div><div class="bpar"><span class="bv nd">Sin programas marcados: beneficio y costo en 0</span></div><div class="bbc">B/C —</div></div>';
      return;
    }
    h += '<div class="bfila"><div class="bnom">' + chipAlt(a) + '</div><div class="bpar">';
    [["Beneficio", a.beneficio, a.color], ["Costo", a.costo, COLOR_GRIS]].forEach(function (b) {
      var ancho = max > 0 && b[1] ? Math.max(0.5, b[1] / max * 80) : 0;
      h += '<div class="bl">' + (ancho ? '<span class="barra" tabindex="0" style="width:' + ancho.toFixed(2) + '%;background:' + b[2] +
        '" data-tip="' + esc(a.nombre + " · " + b[0].toLowerCase()) + '" data-tip-valor="' + esc(millones(b[1]) + " millones $") + '"></span>' : '') +
        '<span class="bv">' + (b[1] ? millones(b[1]) : (b[1] === 0 ? "0" : "—")) + '</span></div>';
    });
    h += '</div><div class="bbc">' + (a.bc !== null ? 'B/C <b>' + fmt(a.bc, 2) + '</b>' : '<span class="nota">' + (a.vacia ? "sin programas" : "B/C —") + '</span>') + '</div></div>';
  });
  return h + '</div><p class="nota">Millones de pesos, valor presente (VNA).</p>';
}

function seccionComparacion(ctx) {
  var A = ctx.alts;
  var h = '<section id="comparacion"><h2><span class="num">3</span>Comparación de las alternativas</h2>';
  h += '<p>La hoja Alternativas arma tres portafolios de programas. La <b>Alternativa 1</b> la propone el motor de optimización (VBA) con el presupuesto y el B/C mínimo de la portada; ' +
    'las <b>alternativas 2 y 3</b> se componen a mano marcando «SI» en los programas que se quieren incluir. Solo la Alternativa 2 se evalúa en el balance hídrico y financieramente.</p>';

  var ultimoAnio = ctx.anios.length ? ctx.anios[ctx.anios.length - 1] : "";
  function fila(etq, f, cls) {
    return '<tr><th scope="row">' + etq + '</th>' + A.map(function (a) { return '<td class="' + (cls || "n") + (a.id === 2 ? " c2" : "") + '">' + f(a) + '</td>'; }).join("") + '</tr>';
  }
  h += '<div class="desborde"><table class="tabla comp"><colgroup><col class="k"><col><col><col></colgroup><thead><tr><th></th>' + A.map(function (a) {
    return '<th class="n' + (a.id === 2 ? " c2" : "") + '">' + chipAlt(a) + (a.id === 2 ? '<div class="badge">Consolidado</div>' : '') + '</th>';
  }).join("") + '</tr></thead><tbody>';
  h += fila("Origen", function (a) { return esc(a.rol.split(" · ")[0]); }, "t");
  h += fila("Programas incluidos", function (a) { return esc(listaP(a.programas)); }, "t");
  h += fila("Número de programas", function (a) { return a.programas.length; });
  h += fila("Beneficio (millones $)", function (a) { return millones(a.beneficio); });
  h += fila("Costo (millones $)", function (a) { return millones(a.costo); });
  h += fila("Relación B/C", function (a) { return '<b>' + fmt(a.bc, 2) + '</b>'; });
  h += fila("IPUF promedio " + ctx.periodo, function (a) { return fmt(a.ipufPromedio, 2); });
  h += fila("IANC " + esc(ultimoAnio) + " (%)", function (a) { var u = ultimo(serie(a, "ianc")); return u ? fmt(u.v, 1) : "—"; });
  h += fila("Meta de IPUF", function (a) { return veredictoChip(a.cumpleMeta, null); }, "n");
  h += '</tbody></table></div><p class="nota">Beneficio y costo en valor presente (VNA). IPUF en m³/usuario/mes. El veredicto es el de la hoja (celdas ' +
    esc([fuente("ver1"), fuente("ver2"), fuente("ver3")].map(function (f) { return f.split("!")[1]; }).join(", ")) + ').</p>';

  // Figura 1
  var lect = [];
  A.forEach(function (a) {
    if (a.vacia) { lect.push(chipAlt(a) + ': sin programas, beneficio y costo en 0.'); return; }
    if (a.bc === null) return;
    lect.push(chipAlt(a) + ': B/C ' + fmt(a.bc, 2) + ' — ' + (a.bc >= 1 ? "el beneficio supera al costo." : "el costo supera al beneficio."));
  });
  var bcMin = parseNum(ctx.estado.configuracionMotor ? ctx.estado.configuracionMotor.bcMinimo : null);
  if (bcMin !== null) lect.push("El B/C mínimo configurado para el motor es " + fmt(bcMin, 1) + "; las alternativas compuestas a mano no pasan por ese filtro.");
  h += figura(ctx, "Beneficio y costo de cada alternativa",
    "Valor presente de los beneficios y de los costos de los programas que incluye cada alternativa.",
    barrasBeneficioCosto(ctx),
    '<p>Cada alternativa tiene dos barras: la de color es el beneficio (VNA de ingresos de sus programas) y la gris es el costo (VNA de egresos e inversión). ' +
    'Si la barra de color es más larga que la gris, la relación B/C es mayor que 1: los programas devuelven más de lo que cuestan. A la derecha, la B/C de cada alternativa (beneficio ÷ costo).</p>',
    lect, esc(fuente("a1bc") + ", " + fuente("a2bc") + ", " + fuente("a3bc")));

  // Matriz de programas
  h += '<h3>Programas que incluye cada alternativa</h3>';
  h += '<div class="desborde"><table class="tabla matriz"><thead><tr><th>Programa</th><th>Tipo</th><th class="n">B/C</th>' +
    A.map(function (a) { return '<th class="c"><span class="chip-alt">' + punto(a.color) + esc(a.corto) + '</span></th>'; }).join("") + '</tr></thead><tbody>';
  ctx.programas.forEach(function (p) {
    h += '<tr><td><span class="pn">P.' + p.numero + '</span> ' + esc(nombrePrograma(p)) + '</td><td>' + (p.tipo === "tecnico" ? "Técnico" : "Comercial") + '</td>' +
      (p.nota ? '<td class="nw"><span class="nota">' + esc(oracion(p.nota)) + '</span></td>' : '<td class="n">' + fmt(p.bc, 2) + '</td>') +
      A.map(function (a) {
        var si = a.programas.indexOf(p.numero) >= 0;
        return '<td class="c">' + (si ? '<span class="incl" style="background:' + a.color + '"><span class="sr">Incluido</span></span>' : '<span class="sr">No incluido</span>') + '</td>';
      }).join("") + '</tr>';
  });
  h += '</tbody><tfoot><tr><th colspan="3">Programas incluidos</th>' + A.map(function (a) { return '<td class="c"><b>' + a.programas.length + '</b></td>'; }).join("") + '</tr></tfoot></table></div>';
  h += '<p class="nota">Un punto de color marca que el programa está incluido. Fuente: filas de banderas ' +
    esc([fuente("flagsV"), fuente("flags2"), fuente("flags3")].map(function (f) { return f.split("!")[1]; }).join(", ")) + ' de la hoja Alternativas.</p>';
  return h + '</section>';
}

function lineasPorAlternativa(ctx, clave, factor) {
  return ctx.alts.map(function (a) {
    var v = serie(a, clave);
    return { nombre: a.nombre, etiqueta: a.corto, color: a.color, valores: factor ? escalar(v, factor) : v };
  });
}

function seccionIndicadores(ctx) {
  var A = ctx.alts, anios = ctx.anios, ini = ctx.inicioPromedio;
  var h = '<section id="indicadores"><h2><span class="num">4</span>Evolución de los indicadores</h2>';
  h += '<p>Cada alternativa tiene en la hoja un bloque de balance hídrico: producción, facturación y usuarios proyectados, el volumen que recuperan sus programas y, con eso, cuatro indicadores por año. ' +
    'Las gráficas de la hoja («INDICADORES ALTERNATIVA No. 1, 2 y 3») ponen los cuatro en un mismo eje aunque el IANC está en % y los demás en m³/usuario/mes; aquí se separan por unidad y se comparan las alternativas entre sí.</p>';
  if (anios.length < 2) return h + '<p class="nota">No se pudieron leer los años del horizonte (' + esc(fuente("anios")) + ').</p></section>';

  var banda = { desde: ini, etiqueta: "Periodo del promedio " + ctx.periodo };
  var leyAlts = A.map(function (a) { return { color: a.color, texto: esc(a.nombre) }; });

  // Iguales al principio: los años en que las tres alternativas coinciden porque no recuperan volumen.
  var iniciales = 0;
  for (var i = 0; i < anios.length; i++) {
    var todosCero = A.every(function (a) {
      var p = serie(a, "vrecProduccion")[i], f = serie(a, "vrecFacturacion")[i];
      return (p === 0 || p === null || p === undefined) && (f === 0 || f === null || f === undefined);
    });
    if (!todosCero) break;
    iniciales++;
  }
  var notaIniciales = iniciales > 0 && iniciales < anios.length
    ? ' Entre ' + esc(anios[0]) + ' y ' + esc(anios[iniciales - 1]) + ' las tres alternativas coinciden: ninguna recupera volumen en esos años.' : '';

  // Figura 2 · IPUF
  var refs = [];
  if (ctx.meta !== null) refs.push({ valor: ctx.meta, etiqueta: "Meta de la empresa · " + fmt(ctx.meta, 1), color: "#063F5C" });
  refs.push({ valor: MAXIMO_NORMATIVO_IPUF, etiqueta: "Máximo normativo · " + fmt(MAXIMO_NORMATIVO_IPUF, 1), color: "#A93A2B" });
  var lect2 = [];
  A.forEach(function (a) {
    var s = serie(a, "ipuf"), p = primero(s), u = ultimo(s);
    if (!p) return;
    lect2.push(chipAlt(a) + ': de ' + fmt(p.v, 2) + ' (' + esc(anios[p.i]) + ') a ' + fmt(u.v, 2) + ' (' + esc(anios[u.i]) + '); promedio ' + ctx.periodo + ' de <b>' +
      fmt(a.ipufPromedio, 2) + '</b>. ' + cumpleFrase(a, ctx) + '.');
  });
  var sobre = A.filter(function (a) { return serie(a, "ipuf").some(function (v) { return num(v) !== null && v > MAXIMO_NORMATIVO_IPUF; }); });
  lect2.push(sobre.length ? "Superan el máximo normativo de " + fmt(MAXIMO_NORMATIVO_IPUF, 1) + " en algún año: " + sobre.map(chipAlt).join(", ") + "."
    : "Ninguna alternativa supera el máximo normativo de " + fmt(MAXIMO_NORMATIVO_IPUF, 1) + " m³/usuario/mes en el horizonte.");
  h += figura(ctx, "IPUF de cada alternativa frente a la meta",
    "Índice de pérdidas por usuario facturado (m³/usuario/mes), año a año.",
    leyenda(leyAlts.concat([{ color: "#063F5C", guion: true, texto: "Meta de la empresa" }, { color: "#A93A2B", guion: true, texto: "Máximo normativo (Res. 330/2017)" },
      { tipo: "banda", texto: "Periodo del promedio" }])) +
    graficaLineas(ctx, { anios: anios, series: lineasPorAlternativa(ctx, "ipuf"), unidad: "m³/usuario/mes", decimales: 2, refs: refs, banda: banda,
      aria: "IPUF por año de las tres alternativas, con la meta de la empresa y el máximo normativo" }),
    '<p>El IPUF es el agua que se pierde por cada usuario facturado al mes: cuanto más baja la línea, mejor. La franja gris (' + ctx.periodo + ') es el tramo que la hoja promedia para dar su veredicto: ' +
    'una alternativa <b>cumple</b> si ese promedio queda por debajo de la línea de meta.' + notaIniciales + '</p>',
    lect2, esc(fuente("alt1bal") + ", " + fuente("alt2bal") + ", " + fuente("alt3bal") + " (fila IPUF)"));

  // Figura 3 · IANC
  var lect3 = [], mejor = null;
  A.forEach(function (a) {
    var s = serie(a, "ianc"), p = primero(s), u = ultimo(s);
    if (!p) return;
    lect3.push(chipAlt(a) + ': de ' + fmt(p.v, 1) + ' % (' + esc(anios[p.i]) + ') a ' + fmt(u.v, 1) + ' % (' + esc(anios[u.i]) + ').');
    if (!mejor || u.v < mejor.v) mejor = { a: a, v: u.v };
  });
  if (mejor && A.filter(function (a) { return ultimo(serie(a, "ianc")); }).length > 1) {
    lect3.push("El IANC más bajo al final del horizonte es el de la " + chipAlt(mejor.a) + " (" + fmt(mejor.v, 1) + " %).");
  }
  h += figura(ctx, "IANC de cada alternativa",
    "Índice de agua no contabilizada: porcentaje del agua producida que no se factura.",
    leyenda(leyAlts.concat([{ tipo: "banda", texto: "Periodo del promedio del IPUF" }])) +
    graficaLineas(ctx, { anios: anios, series: lineasPorAlternativa(ctx, "ianc"), unidad: "%", decimales: 1, banda: { desde: ini, etiqueta: "" },
      aria: "IANC por año de las tres alternativas" }),
    '<p>El IANC compara lo que se produce con lo que se factura: es la diferencia entre ambos como porcentaje de la producción. Una línea más baja significa que una parte mayor del agua producida llega a facturarse. ' +
    'A diferencia del IPUF, no depende del número de usuarios.</p>',
    lect3, esc(fuente("alt1bal") + ", " + fuente("alt2bal") + ", " + fuente("alt3bal") + " (fila IANC)"));

  // Figura 4 · volumen recuperado (dos paneles, misma escala)
  var vp = lineasPorAlternativa(ctx, "vrecProduccion", 1e6), vf = lineasPorAlternativa(ctx, "vrecFacturacion", 1e6);
  var rv = rangoDe(vp.concat(vf).map(function (s) { return s.valores; }), [0]);
  var escV = escalaEje(Math.min(0, rv.lo), rv.hi, 4);
  var panel = function (titulo, series, aria) {
    return '<div><h5>' + titulo + '</h5>' + graficaLineas(ctx, { anios: anios, series: series, unidad: "millones de m³/año", decimales: 2, escala: escV,
      ancho: 440, alto: 270, margenIzq: 42, margenDer: 104, fuente: 13, aria: aria }) + '</div>';
  };
  var lect4 = [];
  A.forEach(function (a) {
    var p = serie(a, "vrecProduccion"), f = serie(a, "vrecFacturacion");
    var desde = Math.min.apply(null, [primerPositivo(p), primerPositivo(f)].filter(function (k) { return k >= 0; }).concat([Infinity]));
    if (!isFinite(desde)) { lect4.push(chipAlt(a) + ': no recupera volumen' + (a.vacia ? ' (no tiene programas marcados).' : '.')); return; }
    var mp = maximo(p), mf = maximo(f);
    lect4.push(chipAlt(a) + ': recupera desde ' + esc(anios[desde]) + '. Máximo anual de ' + (mp ? millonesM3(mp.v) + ' millones de m³ en producción (' + esc(anios[mp.i]) + ')' : '—') +
      ' y de ' + (mf ? millonesM3(mf.v) + ' en facturación (' + esc(anios[mf.i]) + ')' : '—') + '.');
  });
  h += figura(ctx, "Volumen recuperado por año",
    "Agua que recuperan cada año los programas de cada alternativa, en millones de m³.",
    leyenda(leyAlts) + '<div class="multiples dos">' +
    panel("Recuperado de producción", vp, "Volumen recuperado de producción por año, tres alternativas") +
    panel("Recuperado de facturación", vf, "Volumen recuperado de facturación por año, tres alternativas") + '</div>',
    '<p>El volumen recuperado <b>de producción</b> es agua que deja de producirse porque se pierde menos en la red; el recuperado <b>de facturación</b> es agua que ya se consumía y pasa a facturarse. ' +
    'El primero baja la producción y el segundo sube la facturación; ambos se reflejan en el IPUF y el IANC de las figuras 2 y 3. Los dos paneles usan la misma escala.</p>',
    lect4, esc("filas «V RECUPERADO DE PRODUCCIÓN» y «V RECUPERADO DE FACTURACIÓN» de " + fuente("alt1bal") + ", " + fuente("alt2bal") + ", " + fuente("alt3bal")));

  // Figura 5 · ISUF e ICUF por alternativa (el IPUF es la franja entre ambas)
  var todas = [];
  A.forEach(function (a) { todas.push(serie(a, "isuf"), serie(a, "icuf")); });
  var ri = rangoDe(todas, []);
  var escI = escalaEje(ri.lo, ri.hi, 4);
  var lect5 = [];
  var paneles = A.map(function (a) {
    var isuf = serie(a, "isuf"), icuf = serie(a, "icuf");
    var pi = primero(isuf), ui = ultimo(isuf), pc = primero(icuf), uc = ultimo(icuf);
    if (pi && pc) {
      lect5.push(chipAlt(a) + ': ISUF de ' + fmt(pi.v, 2) + ' a ' + fmt(ui.v, 2) + ' e ICUF de ' + fmt(pc.v, 2) + ' a ' + fmt(uc.v, 2) +
        ' (' + esc(anios[pi.i]) + ' → ' + esc(anios[ui.i]) + ').');
    }
    return '<div><h5>' + chipAlt(a) + '</h5>' + graficaLineas(ctx, {
      anios: anios, unidad: "m³/usuario/mes", decimales: 2, escala: escI, ancho: 300, alto: 250, margenIzq: 30, margenDer: 86, fuente: 12.5,
      series: [{ nombre: "ISUF", color: a.color, valores: isuf }, { nombre: "ICUF", color: "#52514e", valores: icuf }],
      area: { sup: isuf, inf: icuf, color: a.color },
      aria: "ISUF e ICUF por año de la " + a.nombre
    }) + '</div>';
  });
  h += figura(ctx, "Cómo se forma el IPUF en cada alternativa",
    "Suministro por usuario (ISUF) y consumo facturado por usuario (ICUF), en m³/usuario/mes. La franja entre ambas líneas es el IPUF.",
    leyenda([{ color: "#6A7975", texto: "ISUF: línea del color de la alternativa" }, { color: "#52514e", texto: "ICUF: línea gris oscuro" },
      { tipo: "area", color: "rgba(106,121,117,.25)", texto: "IPUF = ISUF − ICUF (franja)" }]) +
    '<div class="multiples">' + paneles.join("") + '</div>',
    '<p>El IPUF es la resta ISUF − ICUF: lo que se suministra por usuario y no se le factura. Una alternativa lo reduce acercando las dos líneas, ya sea bajando el suministro (menos pérdidas en la red) ' +
    'o subiendo el consumo facturado. Los tres paneles usan la misma escala, así que el ancho de la franja se puede comparar entre alternativas.</p>',
    lect5, esc("filas ISUF e ICUF de " + fuente("alt1bal") + ", " + fuente("alt2bal") + ", " + fuente("alt3bal")));
  return h + '</section>';
}

function barrasProgramas(ctx) {
  var P = ctx.programas.slice();
  var conBC = P.filter(function (p) { return !p.nota && p.bc !== null && p.bc !== 0; }).sort(function (a, b) { return b.bc - a.bc; });
  var resto = P.filter(function (p) { return conBC.indexOf(p) < 0; });
  var max = Math.max.apply(null, [1].concat(conBC.map(function (p) { return p.bc; })));
  var escala = escalaEje(0, max, 4).max;
  var pos = function (v) { return (v / escala * 82).toFixed(2); };
  var h = leyenda([{ tipo: "barra", color: COLOR_PROGRAMA, texto: "B/C ≥ 1" }, { tipo: "barra", color: COLOR_GRIS, texto: "B/C < 1" },
    { color: "#1A2320", guion: true, texto: "B/C = 1 (beneficio igual al costo)" }]);
  h += '<div class="barras prog"><div class="bfila eje"><div class="bnom"></div><div class="bpista"><span class="umbral-et" style="left:' + pos(1) + '%">B/C = 1</span></div></div>';
  conBC.forEach(function (p) {
    var viable = p.estado === "viable";
    h += '<div class="bfila"><div class="bnom" title="' + esc(nombrePrograma(p)) + '"><span class="pn">P.' + p.numero + '</span> ' + esc(nombrePrograma(p)) + '</div>' +
      '<div class="bpista"><span class="umbral" style="left:' + pos(1) + '%"></span><span class="barra" tabindex="0" style="width:' + pos(p.bc) + '%;background:' + (viable ? COLOR_PROGRAMA : COLOR_GRIS) + '"' +
      ' data-tip="' + esc("P." + p.numero + " " + nombrePrograma(p)) + '" data-tip-valor="' + esc("B/C " + fmt(p.bc, 2)) + '"></span><span class="bv">' + fmt(p.bc, 2) + '</span></div></div>';
  });
  if (resto.length) h += '<div class="bresto">';
  resto.forEach(function (p) {
    h += '<div class="bfila"><div class="bnom" title="' + esc(nombrePrograma(p)) + '"><span class="pn">P.' + p.numero + '</span> ' + esc(nombrePrograma(p)) + '</div>' +
      '<div class="bpista"><span class="bv nd">' + (p.nota ? esc(oracion(p.nota)) : "Sin datos: beneficio y costo en 0") + '</span></div></div>';
  });
  if (resto.length) h += '</div>';
  return h + '</div>';
}

function seccionProgramas(ctx) {
  var P = ctx.programas;
  var h = '<section id="programas"><h2><span class="num">5</span>Programas</h2>';
  h += '<p>La tabla de programas de la hoja Alternativas trae, para cada uno de los ' + P.length + ' programas, el valor presente de sus ingresos y de sus egresos más la inversión, y su relación B/C. ' +
    'Los programas técnicos (P.1 a P.8) y los comerciales (P.9 a P.19) calculan su B/C con dos fórmulas distintas (Especificación Técnica, §4.8).</p>';
  if (!P.length) return h + '<p class="nota">No se pudo leer la tabla de programas (' + esc(fuente("prog")) + ').</p></section>';

  var viables = P.filter(function (p) { return p.estado === "viable"; }).sort(function (a, b) { return b.bc - a.bc; });
  var noViables = P.filter(function (p) { return p.estado === "no_viable"; });
  var sinDatos = P.filter(function (p) { return p.estado === "sin_datos"; });
  var conNota = P.filter(function (p) { return p.estado === "no_aplica"; });
  var lect = [];
  lect.push(viables.length + " de " + P.length + " programas tienen B/C de 1 o más; " + noViables.length + " la tienen por debajo de 1.");
  if (viables.length) lect.push("Mayor B/C: " + viables.slice(0, 3).map(function (p) { return "P." + p.numero + " (" + fmt(p.bc, 2) + ")"; }).join(", ") + ".");
  if (sinDatos.length) lect.push("Sin datos (beneficio y costo en 0): " + esc(listaP(sinDatos.map(function (p) { return p.numero; }))) + ". No se distingue si no aplican o si falta diligenciarlos.");
  conNota.forEach(function (p) { lect.push("P." + p.numero + " no tiene B/C propia: «" + esc(oracion(p.nota)) + "»."); });

  h += figura(ctx, "Relación beneficio/costo de cada programa",
    "B/C de los programas, de mayor a menor. Al final, los que no tienen una B/C que graficar.",
    barrasProgramas(ctx),
    '<p>La línea vertical marca B/C = 1. Las barras que la pasan (azul oscuro) son programas que devuelven más de lo que cuestan; las grises, menos. ' +
    'El motor de optimización combina programas para la Alternativa 1 respetando el presupuesto y el B/C mínimo; un programa con B/C menor que 1 puede entrar en una combinación si el conjunto cumple.</p>',
    lect, esc(fuente("prog")));

  h += '<div class="desborde"><table class="tabla"><thead><tr><th>Programa</th><th>Tipo</th><th class="n">Beneficio (VNA, millones $)</th><th class="n">Costo (VNA, millones $)</th><th class="n">B/C</th><th>Estado</th></tr></thead><tbody>';
  P.forEach(function (p) {
    var est = { viable: '<span class="estado ok"><span aria-hidden="true">✓</span> B/C ≥ 1</span>', no_viable: '<span class="estado no"><span aria-hidden="true">✕</span> B/C &lt; 1</span>',
      sin_datos: '<span class="estado nd">Sin datos</span>', no_aplica: '<span class="estado nd">Sin B/C propia</span>' }[p.estado] || "";
    h += '<tr><td><span class="pn">P.' + p.numero + '</span> ' + esc(nombrePrograma(p)) + '</td><td>' + (p.tipo === "tecnico" ? "Técnico" : "Comercial") + '</td>' +
      '<td class="n">' + (p.nota ? '<span class="nota">' + esc(oracion(p.nota)) + '</span>' : millones(p.beneficioVPN)) + '</td><td class="n">' + (p.nota ? "" : millones(p.costoVPN)) + '</td>' +
      '<td class="n"><b>' + (p.nota ? "—" : fmt(p.bc, 2)) + '</b></td><td>' + est + '</td></tr>';
  });
  return h + '</tbody></table></div></section>';
}

function mismosProgramas(a, b) {
  if (a.length !== b.length) return false;
  var x = a.slice().sort(function (p, q) { return p - q; }), y = b.slice().sort(function (p, q) { return p - q; });
  return x.every(function (v, i) { return v === y[i]; });
}

function seccionPlan(ctx) {
  var pg = ctx.plan, p = pg.plan || {}, combos = pg.combinaciones || [], cl = pg.clasificacion || { criterios: [], grupos: [] };
  var alt1 = ctx.alts[0];
  var h = '<section id="plan"><h2><span class="num">6</span>Plan de gestión de pérdidas</h2>';
  if (p.descripcion) h += '<p>' + esc(p.descripcion) + '.</p>';
  if (p.requisitoEsencial && p.requisitoEsencial.length) {
    h += '<p><b>Programas requisito esencial:</b> ' + p.requisitoEsencial.map(function (r) { return '<span class="tag">' + esc(r) + '</span>'; }).join(" ") + '</p>';
  }
  if (p.programas && p.programas.length) {
    h += '<p><b>Programas del plan (' + p.programas.length + '):</b> ' + esc(listaP(p.programas)) + '. Según la hoja, el plan es la combinación de mayor volumen recuperado entre las que evaluó el motor.</p>';
  }
  h += '<div class="kpis">' +
    kpi("Costo del plan", millones(p.costo), "millones $", "valor presente") +
    kpi("Beneficio del plan", millones(p.beneficio), "millones $", "valor presente") +
    kpi("Relación B/C", fmt(p.bc, 2), "", "") +
    kpi("Volumen recuperado", millonesM3(p.volumenRecuperado), "millones de m³", "en todo el horizonte") + '</div>';

  if (combos.length) {
    h += '<h3>Combinaciones evaluadas por el motor</h3><p class="nota">Las ' + combos.length + ' combinaciones óptimas que registró el motor de optimización, en el orden de la hoja. ' +
      'Se marcan la que la hoja toma como plan y la que corresponde a la Alternativa 1.</p>';
    h += '<div class="desborde"><table class="tabla"><thead><tr><th>#</th><th>Programas</th><th class="n">B/C</th><th class="n">Vol. recuperado (millones de m³)</th><th class="n">Costo (millones $)</th><th class="n">Beneficio (millones $)</th></tr></thead><tbody>';
    combos.forEach(function (c, i) {
      var marcas = [];
      if (p.programas && p.programas.length && mismosProgramas(c.programas, p.programas)) marcas.push('<span class="badge">Plan</span>');
      if (!alt1.vacia && mismosProgramas(c.programas, alt1.programas)) marcas.push('<span class="badge a1">' + punto(alt1.color) + 'Alt. 1</span>');
      h += '<tr' + (marcas.length ? ' class="marcada"' : '') + '><td class="n">' + (i + 1) + '</td><td>' + esc(c.programasTexto) + ' ' + marcas.join(" ") + '</td><td class="n">' + fmt(c.bc, 3) + '</td>' +
        '<td class="n">' + millonesM3(c.volumenRecuperado) + '</td><td class="n">' + millones(c.costo) + '</td><td class="n">' + millones(c.beneficio) + '</td></tr>';
    });
    h += '</tbody></table></div>';
  }

  if (cl.grupos && cl.grupos.length) {
    h += '<h3>Programas destacados por criterio</h3><p class="nota">Clasificación de la hoja Alternativas: los programas obligatorios y los que más aportan en volumen recuperado y en relación B/C, por tipo.</p>';
    h += '<div class="desborde"><table class="tabla"><thead><tr><th></th>' + cl.criterios.map(function (c) { return '<th>' + esc(c) + '</th>'; }).join("") + '</tr></thead><tbody>';
    cl.grupos.forEach(function (g) {
      h += '<tr><th scope="row">' + esc(g.nombre) + '</th>' + g.columnas.map(function (col) {
        return '<td>' + (col.length ? col.map(esc).join("<br>") : '<span class="nota">—</span>') + '</td>';
      }).join("") + '</tr>';
    });
    h += '</tbody></table></div>';
  }
  h += '<p class="fuente">Fuente: ' + esc(fuente("plan") + ", " + fuente("combos") + ", " + fuente("categ")) + '.</p>';
  return h + '</section>';
}

function seccionFinanciero(ctx) {
  var fin = ctx.fin;
  var h = '<section id="financiero"><h2><span class="num">7</span>Resultados financieros del consolidado</h2>';
  h += '<p>Flujo de caja del proyecto con la Alternativa 2, la que alimenta el consolidado. Ver la advertencia de la sección 2 sobre su vigencia.</p>';
  if (!fin) return h + '<p class="nota">Sin datos financieros.</p></section>';
  if (fin.flujoCajaEncontrado) {
    h += '<div class="kpis">' +
      kpi("VNA del proyecto", esc(fin.vna), "millones $", "") +
      kpi("TIR", esc(fin.tir), "", fin.tirDefinida ? "" : "No es un error: el flujo de caja no cambia de signo en el horizonte, así que la TIR no está definida.") +
      kpi("TIR modificada", fin.tirModificada !== null ? fmt(fin.tirModificada * 100, 1) : esc(fin.tirModificadaTexto || "—"), fin.tirModificada !== null ? "%" : "", "") +
      kpi("Periodo de recuperación", esc(fin.payback), String(fin.payback).trim() === "1" ? "año" : "años", "") + '</div>';
  } else {
    h += aviso("wn", "No se encontró la hoja de flujo de caja", "El libro no tiene la hoja <code>" + esc(fuente("fcaja").split("!")[0]) + "</code>.");
  }
  var pr = fin.parametros || {};
  h += '<h3>Parámetros que gobiernan estos números</h3><table class="tabla parametros"><tbody>' +
    '<tr><th scope="row">WACC en pesos corrientes</th><td class="n">' + (esc(pr.waccCorriente) || "—") + '</td></tr>' +
    '<tr><th scope="row">Cargo fijo acueducto</th><td class="n">' + (esc(pr.cargoFijoAcueducto) || "—") + '</td></tr>' +
    '<tr><th scope="row">Cargo variable acueducto</th><td class="n">' + (esc(pr.cargoVariableAcueducto) || "—") + '</td></tr>' +
    (pr.gradiente !== null && pr.gradiente !== undefined ? '<tr><th scope="row">Gradiente de perpetuidad</th><td class="n">' + esc(pr.gradiente) + '</td></tr>' : '') +
    '</tbody></table>';
  if (fin.alcantarilladoEnCero) h += '<p class="nota">Los cargos de alcantarillado están en cero: confirmar si la empresa no presta ese servicio o si falta el dato.</p>';
  h += '<p class="fuente">Fuente: ' + esc(fuente("fcaja") + ", " + fuente("tirm") + ", " + fuente("par") + ", " + fuente("grad")) + '.</p>';
  return h + '</section>';
}

function seccionNotas(ctx) {
  var h = '<section id="notas"><h2><span class="num">8</span>Notas metodológicas y fuentes</h2>';
  h += '<h3>Cómo se calculan los indicadores en la hoja</h3><dl class="defs">' +
    '<div><dt>Producción y facturación con programas</dt><dd>Producción − volumen recuperado de producción; facturación + volumen recuperado de facturación.</dd></div>' +
    '<div><dt>IANC (%)</dt><dd>(Producción con programas − facturación con programas) ÷ producción con programas × 100.</dd></div>' +
    '<div><dt>ISUF (m³/usuario/mes)</dt><dd>Producción con programas ÷ (usuarios × 12).</dd></div>' +
    '<div><dt>ICUF (m³/usuario/mes)</dt><dd>Facturación con programas ÷ (usuarios × 12).</dd></div>' +
    '<div><dt>IPUF (m³/usuario/mes)</dt><dd>ISUF − ICUF. El máximo normativo es ' + fmt(MAXIMO_NORMATIVO_IPUF, 1) + ' (Res. 330/2017, art. 9).</dd></div>' +
    '<div><dt>Veredicto de la meta</dt><dd>La hoja dice que una alternativa cumple si el promedio de su IPUF en ' + ctx.periodo + ' es menor que la meta de la empresa' +
    (ctx.meta !== null ? ' (' + fmt(ctx.meta, 1) + ')' : '') + '.</dd></div>' +
    '<div><dt>Relación B/C de una alternativa</dt><dd>Suma de los VNA de ingresos de sus programas ÷ suma de los VNA de egresos e inversión.</dd></div>' +
    '<div><dt>Programa 8</dt><dd>No tiene evaluación B/C propia: se evalúa junto con el programa 19.</dd></div>' +
    '</dl>';
  h += '<h3>De dónde sale cada dato</h3><p>Todas las cifras se leyeron del libro de Excel al generar el informe; el panel no recalcula el modelo. ' +
    'Lo único que hace el informe es presentarlas: pasa pesos a millones de pesos y m³ a millones de m³, y promedia el IPUF en el mismo tramo de años que usa la hoja para su veredicto.</p>';
  var filas = [
    ["Tabla de programas", ["prog"]], ["Composición y B/C de cada alternativa", ["selA1", "a1bc", "a2txt", "a2bc", "a3txt", "a3bc"]],
    ["Programas marcados", ["altNums", "flagsV", "flags2", "flags3"]], ["Balance hídrico e indicadores", ["anios", "alt1bal", "alt2bal", "alt3bal"]],
    ["Veredicto de la meta", ["meta", "ver1", "ver2", "ver3"]], ["Plan de gestión", ["plan", "combos", "categ"]],
    ["Controles de integridad", ["ctNom", "ct005"]], ["Resultados financieros", ["fcaja", "tirm"]], ["Parámetros", ["par", "parE", "grad"]], ["Empresa", ["empresa"]]
  ];
  h += '<table class="tabla fuentes"><tbody>' + filas.map(function (f) {
    return '<tr><th scope="row">' + f[0] + '</th><td>' + f[1].map(function (k) { return '<code>' + esc(fuente(k)) + '</code>'; }).join(" ") + '</td></tr>';
  }).join("") + '</tbody></table>';
  return h + '</section>';
}

var FILAS_ANEXO = [
  ["produccion", "Producción", "Mm³/año", 1e6, 2],
  ["facturacion", "Facturación", "Mm³/año", 1e6, 2],
  ["usuarios", "Usuarios", "usuarios", 1, 0],
  ["vrecProduccion", "Recuperado de producción", "Mm³/año", 1e6, 3],
  ["vrecFacturacion", "Recuperado de facturación", "Mm³/año", 1e6, 3],
  ["produccionConProgramas", "Producción con programas", "Mm³/año", 1e6, 2],
  ["facturacionConProgramas", "Facturación con programas", "Mm³/año", 1e6, 2],
  ["ianc", "IANC", "%", 1, 2],
  ["isuf", "ISUF", "m³/usuario/mes", 1, 2],
  ["icuf", "ICUF", "m³/usuario/mes", 1, 2],
  ["ipuf", "IPUF", "m³/usuario/mes", 1, 2]
];

function seccionAnexo(ctx) {
  var anios = ctx.anios, ini = ctx.inicioPromedio;
  var h = '<section id="anexo" class="anexo"><h2><span class="num">A</span>Anexo · Balance hídrico por alternativa</h2>';
  h += '<p class="nota">Valores de la hoja por año. Mm³ = millones de m³. Las columnas sombreadas son el periodo del promedio del IPUF (' + ctx.periodo + ').</p>';
  ctx.alts.forEach(function (a) {
    h += '<h3>' + chipAlt(a) + ' <span class="nota-in">· ' + esc(a.rol) + '</span></h3>';
    if (!Object.keys(a.series).some(function (k) { return a.series[k]; })) { h += '<p class="nota">Sin datos.</p>'; return; }
    h += '<div class="desborde"><table class="tabla anexo-t"><thead><tr><th>Concepto</th><th>Unidad</th>' +
      anios.map(function (y, i) { return '<th class="n' + (i >= ini ? " prom" : "") + '">' + esc(y) + '</th>'; }).join("") + '</tr></thead><tbody>';
    FILAS_ANEXO.forEach(function (f) {
      var s = serie(a, f[0]);
      h += '<tr' + (f[0] === "ipuf" ? ' class="destacada"' : '') + '><th scope="row">' + f[1] + '</th><td class="u">' + f[2] + '</td>' +
        anios.map(function (y, i) { var v = num(s[i]); return '<td class="n' + (i >= ini ? " prom" : "") + '">' + (v === null ? "—" : fmt(v / f[3], f[4])) + '</td>'; }).join("") + '</tr>';
    });
    h += '</tbody></table></div>';
    if (a.veredicto && !a.vacia) h += '<p class="veredicto"><b>Veredicto de la hoja:</b> «' + esc(a.veredicto) + '»</p>';
  });
  return h + '</section>';
}

/* ---------- documento ---------- */

var CSS = [
  ":root{--ink:#1A2320;--ink2:#3D4A46;--muted:#6A7975;--subtle:#96A29E;--border:#E8EEEC;--border-strong:#D2DCD9;--canvas:#F5F5F7;--surface:#fff;--sunk:#F0F4F3;",
  "--primary:#063F5C;--primary-wash:#E6ECEF;--brand:#FD8916;--ok:#1E7A4D;--ok-w:#DEF2E7;--wn:#96650D;--wn-w:#FCF0DA;--cr:#A93A2B;--cr-w:#FAE3E0}",
  "*{box-sizing:border-box}html{color-scheme:light}",
  "body{margin:0;background:var(--canvas);color:var(--ink);font:14px/1.55 'Segoe UI',system-ui,-apple-system,sans-serif;-webkit-print-color-adjust:exact;print-color-adjust:exact}",
  ".acciones{max-width:1040px;margin:16px auto 0;padding:0 16px;display:flex;justify-content:flex-end;gap:8px}",
  ".acciones button{font:inherit;font-size:13px;font-weight:600;border-radius:6px;padding:7px 14px;cursor:pointer;border:1px solid var(--border-strong);background:var(--surface);color:var(--primary)}",
  ".acciones button.pri{background:var(--primary);border-color:var(--primary);color:#fff}",
  ".acciones button:hover{filter:brightness(.96)}",
  ".hoja{max-width:1040px;margin:12px auto 32px;background:var(--surface);border:1px solid var(--border);border-radius:10px;overflow:hidden}",
  ".marca{height:5px;background:linear-gradient(90deg,var(--primary),var(--brand))}",
  ".portada{padding:28px 44px 26px;border-bottom:1px solid var(--border)}",
  ".portada .top{display:flex;align-items:center;justify-content:space-between;gap:16px;flex-wrap:wrap}",
  ".portada img{height:58px;width:auto;display:block}",
  ".marca-txt{font-weight:700;color:var(--primary)}",
  ".eyebrow{font-size:11px;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);font-weight:600}",
  "h1{font-size:30px;line-height:1.2;margin:26px 0 6px;color:var(--primary);font-weight:600;letter-spacing:-.015em}",
  ".empresa{font-size:17px;color:var(--ink2);font-weight:600}",
  ".generado{font-size:12.5px;color:var(--muted);margin-top:2px}",
  ".ficha{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));margin:20px 0 0;border:1px solid var(--border);border-radius:8px}",
  ".ficha>div{padding:10px 14px;border-left:1px solid var(--border)}.ficha>div:first-child{border-left:0}",
  ".ficha dt{font-size:11.5px;color:var(--muted)}.ficha dd{margin:2px 0 0;font-weight:600;font-size:14px}",
  "main{padding:6px 44px 36px}",
  ".indice{margin-top:22px;padding:14px 18px;background:var(--sunk);border-radius:8px}",
  ".indice ol{margin:6px 0 0;padding:0;list-style:none;columns:2;column-gap:32px;font-size:13.5px}",
  ".indice li{margin:3px 0}.indice .in{display:inline-block;min-width:20px;color:var(--brand);font-weight:700;font-variant-numeric:tabular-nums}",
  ".indice a{color:var(--primary);text-decoration:none}.indice a:hover{text-decoration:underline}",
  "section{padding-top:30px}",
  "h2{font-size:20px;color:var(--primary);margin:0 0 6px;font-weight:600;display:flex;gap:12px;align-items:baseline;border-bottom:2px solid var(--primary-wash);padding-bottom:8px}",
  "h2 .num{font-size:13px;color:var(--brand);font-weight:700;min-width:14px}",
  "h3{font-size:15.5px;margin:24px 0 8px;color:var(--ink);font-weight:600;display:flex;align-items:center;gap:6px;flex-wrap:wrap}",
  "h5{margin:0 0 4px;font-size:13px;font-weight:600;color:var(--ink2)}",
  "p{margin:10px 0}.lead{font-size:15px;color:var(--ink2)}",
  ".nota,.nota-in{font-size:12.5px;color:var(--muted)}.nota-in{font-weight:400}",
  ".fuente{font-size:11.5px;color:var(--subtle);margin-top:10px}",
  "code{font-family:Consolas,'Courier New',monospace;font-size:12px;background:var(--sunk);padding:1px 5px;border-radius:3px;color:var(--ink2)}",
  ".pt{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px;vertical-align:baseline;flex:none}",
  ".chip-alt{display:inline-flex;align-items:center;font-weight:600;white-space:nowrap}",
  ".kpis{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;margin:16px 0 8px}",
  ".kpi{border:1px solid var(--border);border-radius:8px;padding:12px 14px}",
  ".kpi .l{font-size:12px;color:var(--muted)}.kpi .v{font-size:26px;font-weight:600;letter-spacing:-.02em;margin-top:2px;line-height:1.2}",
  ".kpi .u{font-size:12.5px;color:var(--muted);font-weight:400;margin-left:5px;letter-spacing:0}.kpi .s{font-size:12.5px;color:var(--ink2);margin-top:6px}",
  ".estado{display:inline-flex;align-items:center;gap:5px;font-size:12px;font-weight:600;padding:2px 8px;border-radius:999px;white-space:nowrap}",
  ".estado.ok{background:var(--ok-w);color:var(--ok)}.estado.no{background:var(--cr-w);color:var(--cr)}.estado.nd{background:var(--sunk);color:var(--muted)}",
  ".frases{padding-left:18px;margin:14px 0 0}.frases li{margin:6px 0}",
  ".aviso{display:flex;gap:12px;align-items:flex-start;border:1px solid var(--border);border-left:3px solid var(--x);border-radius:8px;padding:11px 14px;margin:10px 0}",
  ".aviso .ic{flex:none;width:20px;height:20px;border-radius:50%;background:var(--x);color:#fff;font-size:12px;font-weight:700;display:grid;place-items:center;margin-top:1px}",
  ".aviso .at{font-weight:600}.aviso .ad{font-size:13px;color:var(--ink2);margin-top:2px}",
  ".aviso.ok{--x:var(--ok)}.aviso.wn{--x:var(--wn)}.aviso.cr{--x:var(--cr)}.aviso.info{--x:var(--primary)}",
  ".tabla{width:100%;border-collapse:collapse;font-size:13px;font-variant-numeric:tabular-nums}",
  ".tabla th{font-size:11.5px;text-transform:uppercase;letter-spacing:.04em;color:var(--muted);font-weight:600;text-align:left;padding:8px;border-bottom:1px solid var(--border-strong);vertical-align:bottom}",
  ".tabla tbody th,.tabla tfoot th{text-transform:none;letter-spacing:0;font-size:13px;color:var(--ink);font-weight:600;vertical-align:top;border-bottom:1px solid var(--border)}",
  ".tabla td{padding:8px;border-bottom:1px solid var(--border);vertical-align:top}",
  ".tabla .n{text-align:right;white-space:nowrap}.tabla .c{text-align:center}.tabla td.t{min-width:150px}.tabla td.nw{text-align:right;max-width:160px}",
  ".matriz th.c,.matriz td.c{width:64px}",
  ".tabla .c2{background:#FBFCFD}.tabla thead .c2{background:var(--primary-wash)}",
  ".tabla tfoot td,.tabla tfoot th{border-bottom:0;border-top:1px solid var(--border-strong)}",
  ".tabla tr.marcada td{background:#FFF8F0}",
  ".comp{table-layout:fixed}.comp col.k{width:22%}.comp td.t{font-size:12.5px}",
  ".badge{display:inline-flex;align-items:center;font-size:10.5px;font-weight:700;letter-spacing:.04em;text-transform:uppercase;color:var(--primary);background:var(--primary-wash);padding:2px 7px;border-radius:999px;margin-top:4px}",
  ".badge.a1{color:var(--ink2);background:var(--sunk)}",
  ".tag{display:inline-block;background:var(--sunk);border-radius:999px;padding:2px 10px;font-size:12.5px;margin:2px 2px}",
  ".pn{font-variant-numeric:tabular-nums;color:var(--muted);font-weight:600;margin-right:2px}",
  ".incl{display:inline-block;width:11px;height:11px;border-radius:50%}",
  ".sr{position:absolute;width:1px;height:1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap}",
  ".desborde{overflow-x:auto}",
  "figure{margin:22px 0 8px;border:1px solid var(--border);border-radius:10px;padding:18px 20px}",
  "figcaption .fn{font-size:11px;font-weight:700;color:var(--brand);letter-spacing:.1em;text-transform:uppercase}",
  "figcaption .ft{font-size:16px;font-weight:600;margin:2px 0;color:var(--ink)}figcaption .fq{font-size:13px;color:var(--muted)}",
  ".leyenda{display:flex;flex-wrap:wrap;gap:6px 18px;margin:12px 0 8px;font-size:12.5px;color:var(--ink2)}",
  ".leyenda .li{display:inline-flex;align-items:center;gap:6px}",
  ".leyenda .kb{width:16px;height:10px;background:#F0F4F3;border:1px solid var(--border-strong);border-radius:2px}",
  ".leyenda .ka,.leyenda .kr{width:14px;height:10px;border-radius:2px}",
  "svg.graf{width:100%;height:auto;display:block;overflow:visible}svg.graf:focus{outline:2px solid var(--primary);outline-offset:4px;border-radius:4px}",
  "svg.graf .tn{font-variant-numeric:tabular-nums}",
  "svg.graf .halo{paint-order:stroke;stroke:#fff;stroke-width:4px;stroke-linejoin:round}",
  ".multiples{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;margin-top:6px}.multiples.dos{grid-template-columns:repeat(2,minmax(0,1fr))}",
  ".explica{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:16px}",
  ".explica>div{background:var(--sunk);border-radius:8px;padding:12px 14px;font-size:13px}",
  ".explica h4{margin:0 0 4px;font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:var(--muted)}",
  ".explica p{margin:0}.explica ul{margin:0;padding-left:18px}.explica li{margin:4px 0}",
  ".barras{margin-top:6px}",
  ".bfila{display:grid;grid-template-columns:150px minmax(0,1fr) 92px;gap:12px;align-items:center;padding:7px 0;border-bottom:1px solid var(--border)}",
  ".bfila:last-child{border-bottom:0}",
  ".bpar{display:flex;flex-direction:column;gap:2px}.bl{display:flex;align-items:center;gap:8px;height:16px}",
  ".barra{display:block;height:14px;border-radius:0 4px 4px 0;flex:none;outline:none}.barra:hover,.barra:focus{filter:brightness(1.12)}",
  ".bv{position:relative;z-index:1;background:var(--surface);padding:0 3px;font-size:12.5px;font-variant-numeric:tabular-nums;color:var(--ink);white-space:nowrap}.bv.nd{color:var(--muted);font-style:italic;padding:0}",
  ".bresto{margin-top:8px;padding-top:6px;border-top:1px solid var(--border)}",
  ".bbc{font-size:13px;text-align:right}",
  ".prog .bfila{grid-template-columns:minmax(0,300px) minmax(0,1fr);padding:4px 0;border-bottom:0}",
  ".prog .bnom{font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}",
  ".bpista{position:relative;display:flex;align-items:center;gap:8px;min-height:18px}",
  ".umbral{position:absolute;top:-4px;bottom:-4px;border-left:1.5px dashed var(--ink)}",
  ".prog .eje{padding:0}.prog .eje .bpista{min-height:16px}",
  ".umbral-et{position:absolute;transform:translateX(-50%);font-size:11.5px;color:var(--ink2);white-space:nowrap}",
  ".defs{display:grid;grid-template-columns:1fr 1fr;gap:10px 24px;margin:0}",
  ".defs dt{font-weight:600;font-size:13px}.defs dd{margin:2px 0 0;font-size:13px;color:var(--ink2)}",
  ".fuentes th{width:36%}.fuentes code{display:inline-block;margin:1px 4px 1px 0}",
  ".parametros{max-width:560px}",
  ".anexo-t{font-size:12px}.anexo-t th,.anexo-t td{padding:6px 5px}.anexo-t tbody th{white-space:nowrap;font-size:12px}.anexo-t td.u{color:var(--muted);white-space:nowrap;font-size:11.5px}",
  ".anexo-t .prom{background:#F7F9F8}.anexo-t thead .prom{background:#EEF2F1}",
  ".anexo-t tr.destacada th,.anexo-t tr.destacada td{font-weight:700}",
  ".veredicto{font-size:12.5px;color:var(--ink2);font-style:italic}",
  "footer.pie{border-top:1px solid var(--border);padding:16px 44px;font-size:12px;color:var(--muted);display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap}",
  "#tip{position:fixed;z-index:10;display:none;pointer-events:none;background:#fff;border:1px solid var(--border-strong);border-radius:8px;box-shadow:0 6px 20px rgba(20,30,28,.14);padding:8px 10px;font-size:12.5px;min-width:150px}",
  "#tip .tt{font-weight:600;color:var(--muted);font-size:11.5px;margin-bottom:4px}#tip .tr{display:flex;align-items:center;gap:7px;margin:2px 0}",
  "#tip .k{width:12px;height:3px;border-radius:2px;flex:none}#tip b{font-variant-numeric:tabular-nums;color:var(--ink)}#tip .n{color:var(--muted)}#tip .u{color:var(--subtle);font-size:11px;margin-top:4px}",
  "@media screen and (max-width:820px){.portada,main{padding-left:16px;padding-right:16px}footer.pie{padding:16px}.hoja{margin:8px 0 0;border-radius:0;border-left:0;border-right:0}",
  ".ficha,.kpis{grid-template-columns:1fr 1fr}.ficha>div:nth-child(3){border-left:0}.ficha>div:nth-child(n+3){border-top:1px solid var(--border)}",
  ".explica,.multiples,.multiples.dos,.defs{grid-template-columns:1fr}.indice ol{columns:1}h1{font-size:24px}",
  ".bfila{grid-template-columns:96px minmax(0,1fr) 72px;gap:8px}.prog .bfila{grid-template-columns:1fr}.prog .eje .bnom{display:none}figure{padding:14px}}",
  "@page{size:A4;margin:14mm 12mm}@page anexo{size:A4 landscape;margin:12mm}",
  "@media print{body{background:#fff;font-size:10pt}.acciones,#tip{display:none!important}.hoja{margin:0;border:0;border-radius:0;max-width:none}",
  ".portada,main{padding-left:0;padding-right:0}footer.pie{padding:12px 0}.indice{break-inside:avoid}",
  "figure,.kpis,.aviso,.kpi,.explica,tr,.bfila,.defs>div{break-inside:avoid}h2,h3,figcaption{break-after:avoid}",
  "section#comparacion,section#indicadores,section#programas,section#plan{break-before:page}",
  ".anexo{page:anexo;break-before:page}.desborde{overflow:visible}a{color:inherit;text-decoration:none}}"
].join("\n");

/**
 * Script del documento (ES5, sin dependencias): tooltip con cruz en las
 * gráficas de líneas (se ajusta al año más cercano; también con las flechas
 * del teclado), tooltip en las barras, e imprimir / guardar. Los textos se
 * insertan con textContent, nunca como HTML.
 */
var SCRIPT = [
  "(function(){",
  "var tip=document.getElementById('tip');if(!tip)return;var D={};",
  "try{JSON.parse(document.getElementById('datos-graficas').textContent).forEach(function(g){D[g.id]=g;});}catch(e){}",
  "function f(v,d){return v===null||v===undefined?'—':v.toLocaleString('es-CO',{minimumFractionDigits:d,maximumFractionDigits:d});}",
  "function el(t,c,x){var e=document.createElement(t);if(c)e.className=c;if(x!==undefined)e.textContent=x;return e;}",
  "function vaciar(){while(tip.firstChild)tip.removeChild(tip.firstChild);}",
  "function ocultar(){tip.style.display='none';}",
  "function poner(cx,cy){tip.style.display='block';var w=tip.offsetWidth,h=tip.offsetHeight,x=cx+16,y=cy+16;",
  "if(x+w>window.innerWidth-8)x=cx-w-16;if(y+h>window.innerHeight-8)y=cy-h-16;tip.style.left=Math.max(8,x)+'px';tip.style.top=Math.max(8,y)+'px';}",
  "var NS='http://www.w3.org/2000/svg';",
  "Array.prototype.forEach.call(document.querySelectorAll('svg[data-graf]'),function(svg){",
  "var g=D[svg.getAttribute('data-graf')];if(!g)return;var xh=svg.querySelector('.xh'),hit=svg.querySelector('.hit'),act=-1;",
  "var pts=g.series.map(function(s){var c=document.createElementNS(NS,'circle');c.setAttribute('class','xh-pt');c.setAttribute('r','4');c.setAttribute('fill',s.color);",
  "c.setAttribute('stroke','#fff');c.setAttribute('stroke-width','2');c.setAttribute('visibility','hidden');c.setAttribute('pointer-events','none');svg.appendChild(c);return c;});",
  "function Y(v){return g.y1-(g.y1-g.y0)*(v-g.min)/(g.max-g.min);}",
  "function ver(i,cx,cy){act=i;var x=g.x0+g.dx*i;xh.setAttribute('x1',x);xh.setAttribute('x2',x);xh.setAttribute('visibility','visible');",
  "g.series.forEach(function(s,k){var v=s.valores[i];if(v===null){pts[k].setAttribute('visibility','hidden');return;}pts[k].setAttribute('cx',x);pts[k].setAttribute('cy',Y(v));pts[k].setAttribute('visibility','visible');});",
  "vaciar();tip.appendChild(el('div','tt',g.anios[i]));",
  "g.series.forEach(function(s){var r=el('div','tr'),k=el('span','k');k.style.background=s.color;r.appendChild(k);r.appendChild(el('b','',f(s.valores[i],g.dec)));r.appendChild(el('span','n',s.nombre));tip.appendChild(r);});",
  "if(g.unidad)tip.appendChild(el('div','u',g.unidad));poner(cx,cy);}",
  "function fuera(){act=-1;xh.setAttribute('visibility','hidden');pts.forEach(function(p){p.setAttribute('visibility','hidden');});ocultar();}",
  "function idx(cx){var r=svg.getBoundingClientRect(),vb=svg.viewBox.baseVal,x=(cx-r.left)*vb.width/r.width;return Math.max(0,Math.min(g.anios.length-1,Math.round((x-g.x0)/g.dx)));}",
  "hit.addEventListener('pointermove',function(ev){ver(idx(ev.clientX),ev.clientX,ev.clientY);});hit.addEventListener('pointerleave',fuera);",
  "svg.addEventListener('keydown',function(ev){if(ev.key!=='ArrowLeft'&&ev.key!=='ArrowRight')return;ev.preventDefault();",
  "var i=act<0?g.anios.length-1:act+(ev.key==='ArrowRight'?1:-1);i=Math.max(0,Math.min(g.anios.length-1,i));",
  "var r=svg.getBoundingClientRect(),vb=svg.viewBox.baseVal;ver(i,r.left+(g.x0+g.dx*i)*r.width/vb.width,r.top+r.height/3);});",
  "svg.addEventListener('blur',fuera);});",
  "Array.prototype.forEach.call(document.querySelectorAll('[data-tip]'),function(b){",
  "function ver(cx,cy){vaciar();tip.appendChild(el('div','tt',b.getAttribute('data-tip')));tip.appendChild(el('b','',b.getAttribute('data-tip-valor')));poner(cx,cy);}",
  "b.addEventListener('pointermove',function(ev){ver(ev.clientX,ev.clientY);});b.addEventListener('pointerleave',ocultar);",
  "b.addEventListener('focus',function(){var r=b.getBoundingClientRect();ver(r.right,r.top);});b.addEventListener('blur',ocultar);});",
  "var bi=document.getElementById('btnImprimir');if(bi)bi.addEventListener('click',function(){window.print();});",
  "var bg=document.getElementById('btnGuardarHtml');if(bg)bg.addEventListener('click',function(){ocultar();",
  "var c=document.documentElement.cloneNode(true);Array.prototype.forEach.call(c.querySelectorAll('.xh-pt'),function(n){n.parentNode.removeChild(n);});",
  "var blob=new Blob(['<!DOCTYPE html>\\n'+c.outerHTML],{type:'text/html;charset=utf-8'}),a=document.createElement('a');",
  "a.href=URL.createObjectURL(blob);a.download=bg.getAttribute('data-nombre')||'informe.html';document.body.appendChild(a);a.click();",
  "setTimeout(function(){URL.revokeObjectURL(a.href);a.parentNode.removeChild(a);},1000);});",
  "})();"
].join("\n");

/**
 * @param {Object} modelo   ModeloGesper (modeloRepository.cargarModelo)
 * @param {{fecha?: Date, logo?: string, version?: string}} [opciones]
 *   logo: data URI de la imagen de la marca; sin logo se escribe el nombre.
 * @returns {string} documento HTML completo
 */
export function generarInformeHtml(modelo, opciones) {
  var ctx = contexto(modelo || {}, opciones || {});
  var cuerpo = portada(ctx) + '<main>' + indice() + seccionResumen(ctx) + seccionEstado(ctx) + seccionComparacion(ctx) +
    seccionIndicadores(ctx) + seccionProgramas(ctx) + seccionPlan(ctx) + seccionFinanciero(ctx) + seccionNotas(ctx) + seccionAnexo(ctx) + '</main>';
  var titulo = "Informe de resultados GESPER" + (ctx.estado.empresa ? " · " + ctx.estado.empresa : "");
  return '<!DOCTYPE html>\n<html lang="es">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n' +
    '<meta name="generator" content="Panel GESPER ' + esc(ctx.version) + '">\n' +
    '<title>' + esc(titulo) + '</title>\n<style>\n' + CSS + '\n</style>\n</head>\n<body>\n' +
    '<div class="acciones"><button type="button" id="btnGuardarHtml" data-nombre="' + esc(ctx.nombreArchivo) + '">Guardar HTML</button>' +
    '<button type="button" class="pri" id="btnImprimir">Imprimir o guardar como PDF</button></div>\n' +
    '<article class="hoja"><div class="marca"></div>' + cuerpo +
    '<footer class="pie"><span>Panel GESPER ' + esc(ctx.version) + ' · Dual Ingeniería Hidráulica</span><span>Generado el ' + esc(fechaLarga(ctx.fecha)) + ', ' + esc(hora(ctx.fecha)) + '</span></footer></article>\n' +
    '<div id="tip" role="status" aria-live="polite"></div>\n' +
    '<script type="application/json" id="datos-graficas">' + jsonSeguro(ctx.graficas) + '</script>\n' +
    '<script>\n' + SCRIPT + '\n</script>\n</body>\n</html>\n';
}
