/**
 * Renderizado de cada pantalla del panel. Todas las funciones aquí leen del
 * estado compartido (state.D, state.hojasLibro, ...) y escriben HTML en el
 * DOM; ninguna llama a Office.js directamente — eso vive en services/excelService.js.
 */

import { H, BLOQUES, ANCLAS, ANEXOS } from "../config/cellMap.js";
import { esc, num, fmt, fmtMM } from "../utils/format.js";
import { dependeDecision } from "../domain/anexos/salidaGesper.js";
import { state } from "./state.js";

function $(id) { return document.getElementById(id); }

export function errBox(e) {
  return '<div class="sig c"><span class="lt"></span><div class="tx"><div class="tt">No pude pintar esta sección</div>' +
    '<div class="ds">' + esc(e && e.message ? e.message : e) + '</div></div></div>';
}

function sig(cls, tit, desc, accion) {
  var b = "";
  if (accion) {
    if (accion.indexOf("go:") === 0) b = '<button class="act" data-go="' + accion.slice(3) + '">Ver</button>';
    else if (accion.indexOf("ir:") === 0) {
      var p = accion.slice(3).split("|");
      b = '<button class="act" data-hoja="' + esc(p[0]) + '" data-addr="' + esc(p[1] || "") + '">Ir a la celda</button>';
    }
  }
  return '<div class="sig ' + cls + '"><span class="lt"></span><div class="tx"><div class="tt">' + tit +
    '</div><div class="ds">' + desc + '</div>' + b + '</div></div>';
}

function card(t, v, p) {
  return '<div class="card"><h4>' + t + '</h4><div class="v">' + v + '</div>' + (p ? '<p>' + p + '</p>' : '') + '</div>';
}

/* ---------- PROGRAMAS ---------- */
export function programas() {
  var D = state.D;
  var out = [];
  if (!D.prog) return out;
  for (var i = 0; i < D.prog.v.length; i++) {
    var f = D.prog.v[i];
    var nom = String(f[0] || "").trim();
    if (!nom) continue;
    var n = num(f[1]);
    if (n === null) continue;
    out.push({
      n: n, nombre: nom, ing: num(f[2]), egr: num(f[3]), bc: num(f[4]),
      nota: (typeof f[2] === "string" && f[2]) ? String(f[2]) : ""
    });
  }
  return out;
}

/* ---------- PANTALLA: ESTADO ---------- */
/**
 * Consume el EstadoModelo ya mapeado (domain/mappers/estadoMapper.js), no el
 * raw de Excel -- ver docs/ARQUITECTURA_MVP.md §2/§4. `state.totalProgramas`
 * viaja por fuera del contrato EstadoModelo porque este no incluye el total
 * de programas, solo `programasSinDatos` (ver informe de la Fase 2C:
 * "requiere ajuste del contrato").
 */
export function pintaEstado() {
  var estado = state.estado;
  if (!estado) { $("s-estado").innerHTML = errBox("Todavía no se ha leído el libro."); return; }

  var h = "";
  h += '<div class="idc"><div class="co">' + esc(estado.empresa || "Caso sin identificar") + '</div><div class="mt">';
  if (estado.horizonte) {
    h += "Horizonte " + esc(estado.horizonte.inicio) + " – " + esc(estado.horizonte.fin) + " · " +
      estado.horizonte.cantidadAnios + " años<br>";
  }
  if (estado.metodoProyeccionUsuarios) h += "Proyección de usuarios: " + esc(estado.metodoProyeccionUsuarios) + "<br>";
  if (estado.metaIPUF !== null) h += "Meta IPUF de la empresa: " + fmt(estado.metaIPUF, 1) + " m³/susc·mes<br>";
  h += "Máximo por norma: 6,0 (Res. 330/2017 art. 9)";
  h += "</div></div>";

  h += '<div class="sec">Estado del modelo</div>';

  // 1. Integridad
  var ctTexto = esc(estado.controlesIntegridad.join(" · "));
  if (estado.integridad === "ok") {
    h += sig("k", "Integridad del modelo", "Los tres controles internos en verde. <code>Alternativas!X26:Y29</code>");
  } else if (estado.integridad === "motor_sin_ejecutar") {
    h += sig("w", "El motor no se ha ejecutado", ctTexto + " <code>Y29</code>");
  } else {
    h += sig("w", "Revisar los controles de integridad", ctTexto + " <code>X26:Y29</code>");
  }

  // 2. WACC
  if (estado.wacc.real === null || estado.wacc.real === 0) {
    h += sig("c", "Falta el WACC de la empresa",
      "Con <code>PARÁMETROS!E8</code> vacía el modelo descuenta a la tasa de inflación, no al WACC real, y ninguna fórmula lo advierte.",
      "ir:" + H.PAR + "|E8");
  } else {
    h += sig("k", "Tasa de descuento diligenciada",
      "WACC de la empresa <b>" + fmt(estado.wacc.real * 100, 2) + " %</b> → el modelo descuenta al <b>" +
      (estado.wacc.corriente !== null ? fmt(estado.wacc.corriente * 100, 2) : "—") + " %</b> en pesos corrientes. <code>E8</code>");
  }

  // 3. Alternativa que gobierna
  h += sig("w", "El consolidado sigue a la Alternativa 2",
    "Los resultados financieros —las 19 evaluaciones, el bloque tarifario y el flujo de caja— se calculan sobre la fila 49, que es la Alternativa 2. " +
    (estado.alternativaDesfasada
      ? "La Alternativa 1, la que propuso el motor, incluye programas distintos y <b>no afecta</b> a estos resultados."
      : "Hoy coincide con la selección de la Alternativa 1."),
    "go:alt");

  // 4. Programas sin datos
  var sinDatos = estado.programasSinDatos;
  if (sinDatos.length) {
    h += sig("w", sinDatos.length + " de " + state.totalProgramas + " programas sin datos",
      esc(sinDatos.map(function (p) { return "P." + p.numero; }).join(", ")) +
      " están en 0 de beneficio y 0 de costo. No se distingue si no aplican o si falta diligenciarlos.",
      "go:datos");
  } else {
    h += sig("k", "Todos los programas tienen datos", "Ningún programa quedó en 0 beneficio y 0 costo.");
  }

  // 5. IPC
  if (estado.ipc.desactualizado) {
    h += sig("c", "Series de referencia desactualizadas",
      "El modelo proyecta con un IPC de <b>" + fmt(estado.ipc.valor * 100, 2) + " %</b>. El DANE cerró 2025 en <b>5,10 %</b>. " +
      "Todo el horizonte se está indexando por debajo de la inflación real. <code>C9</code> ← <code>BANREPÚBLICA!V35</code>",
      "ir:" + H.PAR + "|C9");
  }

  h += '<div class="sec">Configuración del motor</div>';
  h += card("Presupuesto disponible", esc(estado.configuracionMotor.presupuesto) || "—");
  h += card("Relación B/C mínima exigida", esc(estado.configuracionMotor.bcMinimo) || "—");
  h += card("Combinación elegida para la Alternativa 1", esc(estado.configuracionMotor.combinacionA1) || "sin elegir");

  $("s-estado").innerHTML = h;
}

/* ---------- PANTALLA: NAVEGAR ---------- */
export function pintaNavegar() {
  var h = '<div class="card accent"><h4>Seis bloques, no 49 hojas</h4>' +
    '<p>El libro tiene ' + state.hojasLibro.length + ' hojas y la mayoría están ocultas. El panel las agrupa por lo que hacen, las muestra si hace falta y las abre en el punto correcto.</p></div>';
  h += '<div class="sec">Ir a</div>';
  BLOQUES.forEach(function (b, i) {
    h += '<button class="navb" data-hoja="' + esc(b.ir[0]) + '" data-addr="' + esc(b.ir[1]) + '">' +
      '<span class="ic">' + (i + 1) + '</span><span><span class="nm">' + esc(b.n) + '</span>' +
      '<span class="sb">' + esc(b.s) + '</span></span></button>';
  });
  $("s-navegar").innerHTML = h;
}

/* ---------- PANTALLA: DATOS ---------- */
export function pintaDatos() {
  if (state.vistaDatos === "p9") return pintaFormP9();
  var P = programas();
  var h = '<div class="sec">Los ' + P.length + ' programas</div>';
  P.forEach(function (p) { h += filaPrograma(p); });
  h += '<div class="flag">P.8 aparece sin cifras a propósito: su evaluación financiera se calcula dentro de P.19. No es un dato faltante.</div>';
  $("s-datos").innerHTML = h;
}

export function filaPrograma(p) {
  var cls = p.bc === null ? "n" : (p.bc >= 1 ? "g" : "r");
  var txt = p.bc === null ? "—" : fmt(p.bc, 2);
  var sub = "";
  if (p.nota) sub = '<span class="sub text-muted">' + esc(p.nota) + '</span>';
  else if (p.bc === 0) sub = '<span class="sub text-warning">Sin datos — 0 beneficio, 0 costo</span>';
  return '<button class="prow" data-prog="' + p.n + '"><span class="pn">P.' + p.n + '</span>' +
    '<span class="pl">' + esc(p.nombre.replace(/^P\.\s*\d+\s*/, "")) + '</span>' +
    '<span class="pb ' + cls + '">' + txt + '</span>' + sub + '</button>';
}

export function pintaFormP9() {
  var D = state.D;
  var a = D.p9a ? D.p9a.v[0][0] : "";
  var b = D.p9b ? D.p9b.v[0][0] : "";
  var c = D.p9c ? D.p9c.v[0][0] : "";
  var rep = D.p9rep ? D.p9rep.v[0] : [];
  var nue = D.p9nue ? D.p9nue.v[0] : [];
  var an = D.p9anios ? D.p9anios.t[0] : [];

  var h = '<button class="act back" data-volver="1">← Volver a los programas</button>';
  h += '<div class="sec">P.9 · Renovación de medidores</div>';

  h += fld("f_a", "Edad actual del parque de medidores", a, "años",
    "Promedio ponderado del parque instalado. <code>EV PROG CCIALES!C181</code>");
  h += fld("f_b", "Vida útil del parque de medidores", b, "años",
    "Determina la tasa de reposición anual. <code>C200</code>");
  h += fld("f_c", "Meses de recuperación de volumen en el año", c, "meses",
    "Entre 1 y 12. <code>C246</code>");

  h += '<div class="sec">Medidores repuestos por año <code>F197:J197</code></div>';
  h += filaAnios("r", an, rep);
  h += '<div class="sec">Medidores para usuarios nuevos <code>F198:J198</code></div>';
  h += filaAnios("n", an, nue);

  h += '<div class="btn-row">' +
    '<button class="act primary lg" id="btnGuardar" style="flex:1;font-size:12px">Guardar en el libro</button>' +
    '<button class="act ghost lg" id="btnDescartar">Descartar</button></div>';
  h += '<div class="flag">En el archivo estos campos están repartidos en cuatro tramos separados entre las filas 181 y 246 de una hoja de 458 filas. El panel nunca escribe sobre una celda que contenga una fórmula.</div>';

  $("s-datos").innerHTML = h;
}

function fld(id, lab, val, uni, hp) {
  return '<div class="fld"><label for="' + id + '">' + lab + '</label><div class="in">' +
    '<input id="' + id + '" type="text" value="' + esc(val === null || val === undefined ? "" : val) + '">' +
    '<span class="un">' + uni + '</span></div><div class="hp">' + hp + '</div></div>';
}

function filaAnios(pref, anios, vals) {
  var n = Math.max(5, vals.length);
  var h = '<div class="yr"><span class="yl"></span>';
  for (var i = 0; i < 5; i++) {
    var a = anios[i] !== undefined && String(anios[i]).trim() ? String(anios[i]) : ("col " + (i + 1));
    if (/^\d+$/.test(a) && Number(a) < 100) a = "Año " + a;
    h += '<span class="yh">' + esc(a) + '</span>';
  }
  h += '</div><div class="yr"><span class="yl">unid.</span>';
  for (var j = 0; j < 5; j++) {
    var v = vals[j] !== undefined && vals[j] !== null ? vals[j] : "";
    h += '<input id="' + pref + j + '" type="text" value="' + esc(v) + '">';
  }
  return h + '</div>';
}

/* ---------- PANTALLA: RESULTADOS ---------- */
/**
 * Los sub-tabs Hidráulica y Financiero consumen ResultadosHidraulicos /
 * ResultadosFinancieros ya mapeados (domain/mappers/resultadosMapper.js), no
 * state.D -- ver docs/ARQUITECTURA_MVP.md §2/§4. El sub-tab Comercial es la
 * excepción deliberada: sigue usando el helper compartido `programas()`
 * (state.D.prog), el mismo que usa la pantalla Datos -- migrarlo aquí
 * significaría migrar (o duplicar) esa lógica antes de tiempo, y Datos no
 * entra en esta fase. Ver informe de la Fase 2D.
 */
export function pintaResultados() {
  var hid = state.resultadosHidraulicos, fin = state.resultadosFinancieros;
  if (!hid || !fin) { $("s-result").innerHTML = errBox("Todavía no se ha leído el libro."); return; }

  var h = '<div class="subnav" id="subR">' +
    '<button data-r="hid" class="' + (state.subResult === "hid" ? "on" : "") + '">Hidráulica</button>' +
    '<button data-r="com" class="' + (state.subResult === "com" ? "on" : "") + '">Comercial</button>' +
    '<button data-r="fin" class="' + (state.subResult === "fin" ? "on" : "") + '">Financiero</button></div>';

  if (state.subResult === "hid") {
    h += '<div class="kpi"><div class="kl">IPUF promedio del horizonte · Alternativa 2</div><div class="kv">' +
      (hid.ipufPromedio !== null ? fmt(hid.ipufPromedio, 2) : "—") + '<span class="ku">m³/susc·mes</span></div>' +
      '<div class="ks">Alternativas!M62:V62 · meta ' + (hid.metaIPUF !== null ? fmt(hid.metaIPUF, 1) : "—") + '</div>' +
      (hid.cumpleMeta !== null ? '<div class="nt ' + (hid.cumpleMeta ? "text-ok" : "text-critical") + '">' +
        (hid.cumpleMeta ? "Cumple la meta de la empresa." : "No alcanza la meta de la empresa.") + '</div>' : "") + '</div>';
    h += sparkIPUF(hid);
    if (hid.ianc.inicio !== null) {
      h += '<div class="kpi"><div class="kl">IANC al inicio del horizonte</div><div class="kv">' + fmt(hid.ianc.inicio, 2) + '<span class="ku">%</span></div><div class="ks">Alternativas!J59</div></div>';
      h += '<div class="kpi"><div class="kl">IANC al final del horizonte</div><div class="kv">' + fmt(hid.ianc.fin, 2) + '<span class="ku">%</span></div><div class="ks">Alternativas!V59</div></div>';
    }
  }

  if (state.subResult === "com") {
    h += '<div class="sec">B/C de los programas comerciales</div>';
    programas().filter(function (p) { return p.n >= 9; }).forEach(function (p) { h += filaPrograma(p); });
    h += '<div class="flag">Un B/C en cero suele significar cantidades sin diligenciar o costo unitario faltante. Es el error más frecuente documentado del modelo.</div>';
  }

  if (state.subResult === "fin") {
    if (fin.flujoCajaEncontrado) {
      h += '<div class="kpi"><div class="kl">VNA del proyecto</div><div class="kv">' + esc(fin.vna) + '<span class="ku">millones $</span></div><div class="ks">FCAJA PROYECTO!C67</div></div>';
      h += '<div class="kpi"><div class="kl">TIR</div><div class="kv">' + esc(fin.tir) + '</div><div class="ks">FCAJA PROYECTO!C66</div>' +
        (!fin.tirDefinida ? '<div class="nt">No es un error: el flujo de caja no cambia de signo dentro del horizonte, así que la TIR no está definida.</div>' : '') + '</div>';
      h += '<div class="kpi"><div class="kl">TIR modificada</div><div class="kv">' +
        (fin.tirModificada !== null ? fmt(fin.tirModificada * 100, 1) + '<span class="ku">%</span>' : esc(fin.tirModificadaTexto)) +
        '</div><div class="ks">FCAJA PROYECTO!D66</div></div>';
      h += '<div class="kpi"><div class="kl">Período de recuperación</div><div class="kv">' + esc(fin.payback) + '<span class="ku">años</span></div><div class="ks">FCAJA PROYECTO!C65</div></div>';
    } else {
      h += '<div class="card"><p>No encuentro la hoja <code>' + esc(H.FCJ) + '</code> en este libro.</p></div>';
    }
    h += '<div class="sec">Parámetros que gobiernan estos números</div>';
    h += card("WACC en pesos corrientes", esc(fin.parametros.waccCorriente), "PARÁMETROS!C8");
    h += card("Cargo fijo acueducto", esc(fin.parametros.cargoFijoAcueducto), "C15");
    h += card("Cargo variable acueducto", esc(fin.parametros.cargoVariableAcueducto), "C16");
    if (fin.alcantarilladoEnCero)
      h += '<div class="flag">Los cargos de alcantarillado están en cero. Confirmar si la empresa no presta ese servicio o si falta el dato. <code>C17</code> <code>C18</code></div>';
    if (fin.parametros.gradiente !== null) h += card("Gradiente de perpetuidad", esc(fin.parametros.gradiente), "PARÁMETROS!C33");
  }
  $("s-result").innerHTML = h;
}

function sparkIPUF(hid) {
  var v = hid.serieIPUF.map(function (s) { return s.valor; });
  var an = hid.serieIPUF.map(function (s) { return s.anio; });
  var meta = hid.metaIPUF;
  var pts = [];
  for (var i = 0; i < v.length; i++) { if (v[i] !== null) pts.push([i, v[i]]); }
  if (pts.length < 2) return "";
  var ys = pts.map(function (p) { return p[1]; }).concat(meta !== null ? [meta] : []);
  var mn = Math.min.apply(null, ys), mx = Math.max.apply(null, ys);
  if (mx - mn < 0.2) { mx += 0.2; mn -= 0.2; }
  var W = 290, Hh = 88, PL = 26, PR = 8, PT = 10, PB = 18;
  function X(i) { return PL + (W - PL - PR) * (i / (v.length - 1)); }
  function Y(y) { return PT + (Hh - PT - PB) * (1 - (y - mn) / (mx - mn)); }
  var poly = pts.map(function (p) { return X(p[0]).toFixed(1) + "," + Y(p[1]).toFixed(1); }).join(" ");
  var s = '<div class="card"><h4>Trayectoria de IPUF</h4><svg viewBox="0 0 ' + W + ' ' + Hh + '" style="width:100%;height:auto" role="img" aria-label="IPUF por año">';
  if (meta !== null) {
    s += '<line x1="' + PL + '" y1="' + Y(meta).toFixed(1) + '" x2="' + (W - PR) + '" y2="' + Y(meta).toFixed(1) +
      '" stroke="var(--cr)" stroke-width="1" stroke-dasharray="3 3"/>';
    s += '<text x="' + (PL + 2) + '" y="' + (Y(meta) - 3).toFixed(1) + '" font-size="8" fill="var(--cr)" font-family="Consolas,monospace">meta ' + fmt(meta, 1) + '</text>';
  }
  s += '<polyline fill="none" stroke="var(--gs)" stroke-width="2" points="' + poly + '"/>';
  var last = pts[pts.length - 1];
  s += '<circle cx="' + X(last[0]).toFixed(1) + '" cy="' + Y(last[1]).toFixed(1) + '" r="2.5" fill="var(--gs)"/>';
  s += '<text x="2" y="' + (Y(mx) + 3).toFixed(1) + '" font-size="8" fill="var(--fnt)" font-family="Consolas,monospace">' + fmt(mx, 1) + '</text>';
  s += '<text x="2" y="' + (Y(mn) + 3).toFixed(1) + '" font-size="8" fill="var(--fnt)" font-family="Consolas,monospace">' + fmt(mn, 1) + '</text>';
  if (an.length) {
    s += '<text x="' + PL + '" y="' + (Hh - 4) + '" font-size="8" fill="var(--fnt)" font-family="Consolas,monospace">' + esc(an[0]) + '</text>';
    s += '<text x="' + (W - PR - 26) + '" y="' + (Hh - 4) + '" font-size="8" fill="var(--fnt)" font-family="Consolas,monospace">' + esc(an[an.length - 1]) + '</text>';
  }
  return s + '</svg></div>';
}

/* ---------- PANTALLA: ALTERNATIVAS ---------- */
export function pintaAlt() {
  var D = state.D;
  function bloque(bcRange) {
    var ben = bcRange ? num(bcRange.v[0][0]) : null;
    var cos = bcRange ? num(bcRange.v[1][0]) : null;
    return { ben: ben, cos: cos, bc: (ben !== null && cos !== null && cos !== 0) ? ben / cos : null };
  }
  var A1 = bloque(D.a1bc), A2 = bloque(D.a2bc), A3 = bloque(D.a3bc);
  function promIPUF(k) {
    if (!D[k]) return null;
    var v = D[k].v[0].map(num).filter(function (x) { return x !== null; });
    if (v.length < 4) return null;
    var s = v.slice(3);
    return s.reduce(function (a, b) { return a + b; }, 0) / s.length;
  }
  var i1 = promIPUF("ipuf1"), i2 = promIPUF("ipuf2");
  function cumple(vr) {
    if (!vr) return null; var t = String(vr.t[0][0] || "");
    if (t.indexOf("NO CUMPLE") >= 0) return false; if (t.indexOf("CUMPLE") >= 0) return true; return null;
  }
  var c1 = cumple(D.ver1), c2 = cumple(D.ver2), c3 = cumple(D.ver3);
  function vd(c) { return c === null ? "—" : '<span class="vd ' + (c ? "y" : "n") + '">' + (c ? "SÍ" : "NO") + '</span>'; }
  function bcCel(x) {
    if (x === null) return "—";
    return '<span class="' + (x >= 1 ? "text-ok" : "text-critical") + ' value-emphasis">' + fmt(x, 2) + '</span>';
  }

  var h = '<div class="sec">Comparación</div><table class="cmp"><thead><tr><th></th>' +
    '<th class="c">Alt. 1</th><th class="c">Alt. 2</th><th class="c">Alt. 3</th></tr></thead><tbody>';
  h += '<tr><td class="k">Beneficio</td><td class="c">' + fmtMM(A1.ben) + '</td><td class="c">' + fmtMM(A2.ben) + '</td><td class="c">' + fmtMM(A3.ben) + '</td></tr>';
  h += '<tr><td class="k">Costo</td><td class="c">' + fmtMM(A1.cos) + '</td><td class="c">' + fmtMM(A2.cos) + '</td><td class="c">' + fmtMM(A3.cos) + '</td></tr>';
  h += '<tr><td class="k">Relación B/C</td><td class="c">' + bcCel(A1.bc) + '</td><td class="c">' + bcCel(A2.bc) + '</td><td class="c">' + bcCel(A3.bc) + '</td></tr>';
  h += '<tr><td class="k">IPUF promedio</td><td class="c">' + (i1 !== null ? fmt(i1, 2) : "—") + '</td><td class="c">' + (i2 !== null ? fmt(i2, 2) : "—") + '</td><td class="c">—</td></tr>';
  h += '<tr><td class="k">Meta IPUF</td><td class="c">' + vd(c1) + '</td><td class="c">' + vd(c2) + '</td><td class="c">' + vd(c3) + '</td></tr>';
  h += '</tbody></table><p class="note">cifras en millones de pesos corrientes</p>';

  h += '<div class="flag gs"><b>La Alternativa 2 es la que alimenta el consolidado financiero.</b> Es la fila 49 de la hoja Alternativas la que gobierna todo lo que viene después.</div>';

  h += '<div class="sec">Alternativa 1 · propuesta por el motor</div>';
  h += '<div class="card"><div class="v desc">' +
    (D.selA1 ? (esc(D.selA1.t[0][0]) || "sin elegir") : "—") + '</div></div>';
  h += '<div class="sec">Alternativa 2 · compuesta a mano</div>';
  h += '<div class="card"><div class="v desc">' +
    (D.a2txt ? (esc(D.a2txt.t[0][0]) || "sin componer") : "—") + '</div></div>';
  h += '<div class="sec">Alternativa 3</div>';
  h += '<div class="card"><div class="v desc">' +
    (D.a3txt ? (esc(D.a3txt.t[0][0]) || "sin componer") : "—") + '</div></div>';

  h += '<button class="act lg block" style="margin-top:10px" data-hoja="' + esc(H.ALT) + '" data-addr="A4">Abrir la hoja Alternativas</button>';
  $("s-alt").innerHTML = h;
}

/* ---------- PANTALLA: CARGAR ANEXOS ---------- */
/**
 * Consume state.anexos (evaluación del archivo y plan de carga ya calculados
 * por domain/repositories/anexosRepository.js). Las reglas de qué se carga y
 * qué se omite viven en domain/anexos/salidaGesper.js, no aquí.
 */
function idCorto(id) { return esc(String(id || "").replace(/^A\d+\./, "")); }
function celdaTxt(v) { return v === null || v === undefined || v === "" ? '<span class="text-muted">vacía</span>' : esc(v); }

function tablaEscrituras(lista) {
  var h = '<div class="tscroll"><table class="cmp"><thead><tr><th>Dato</th><th>Celda</th><th class="c">Actual</th><th class="c">Nuevo</th></tr></thead><tbody>';
  lista.forEach(function (e) {
    h += '<tr' + (e.cambia ? ' class="chg"' : '') + ' title="' + esc(e.fila.CONCEPTO) + '"><td class="id">' + idCorto(e.fila.ID) +
      (dependeDecision(e.fila) ? ' <span class="text-warning" title="Depende de una decisión del profesional">•</span>' : '') +
      '</td><td>' + esc(e.celda) + '</td><td class="c">' + celdaTxt(e.actual) + '</td><td class="c">' + esc(e.valor) + '</td></tr>';
  });
  return h + '</tbody></table></div>';
}

/** Cada valor que cambia en GESPER, con su contexto: qué es, de qué grupo, en qué celda y año, y actual → nuevo. */
function listaCambios(cambian) {
  var h = '<div class="chg-list">';
  cambian.forEach(function (e) {
    var anio = e.anio === null ? "" : " · año " + e.anio + (e.anioCalendario ? " (" + esc(e.anioCalendario) + ")" : "");
    var vacia = e.actual === null || e.actual === undefined || e.actual === "";
    h += '<div class="anx-chg"><div class="tt">' + esc(e.fila.CONCEPTO) + '</div>' +
      '<div class="mt">' + esc(e.fila.SITIO || "") + ' · <code>' + esc(e.celda) + '</code>' + anio + '</div>' +
      '<div class="vv"><span class="' + (vacia ? "text-muted" : "old") + '">' + celdaTxt(e.actual) + '</span><span class="arr">→</span>' +
      '<span class="new">' + esc(e.valor) + '</span></div>' +
      (e.valor === 0 && typeof e.actual === "number" && e.actual !== 0
        ? '<div class="nt text-warning">El anexo trae 0: reemplaza el ' + esc(e.actual) + ' que hoy tiene GESPER.</div>' : '') +
      '</div>';
  });
  return h + '</div>';
}

function listaOmitidos(omitidos) {
  var h = '';
  omitidos.forEach(function (o) {
    h += '<div class="anx-it"><code>' + idCorto(o.fila.ID) + '</code>' +
      (o.fila.DESTINO_CELDA ? ' → <code>' + esc(o.fila.DESTINO_CELDA) + '</code>' : '') + '<div>' + esc(o.motivo) + '</div></div>';
  });
  return h;
}

function listaPendientes(ev) {
  if (!ev.noDisponibles.length) return '';
  var h = '<details class="anx-det"><summary>' + ev.noDisponibles.length + ' campos que este anexo no entrega (pendientes)</summary>';
  ev.noDisponibles.forEach(function (f) {
    h += '<div class="anx-it"><b>' + esc(f.CONCEPTO) + '</b><div>' + esc(f.DESTINO_ETIQUETA || f.NOTA || "") + '</div></div>';
  });
  return h + '</details>';
}

export function pintaAnexos() {
  var A = state.anexos;
  var anexo = ANEXOS.filter(function (a) { return a.id === A.anexoId; })[0];
  var h = '<div class="sec">Cargar anexos</div>';
  ANEXOS.forEach(function (a, i) {
    h += '<button class="navb' + (a.id === A.anexoId ? ' on' : '') + '" data-anexo="' + esc(a.id) + '">' +
      '<span class="ic">' + (i + 1) + '</span><span><span class="nm">' + esc(a.nombre) + '</span>' +
      '<span class="sb">SALIDA_GESPER → ' + esc(a.hojaDestino) + '</span></span></button>';
  });
  var input = '<input type="file" id="anexoArchivo" accept=".xlsx,.xlsm" hidden>';
  var elegir = function (txt, cls) { return '<button class="act ' + cls + '" id="btnElegirAnexo">' + txt + '</button>'; };

  if (A.fase === "inicio") {
    h += '<div class="card"><p>El panel lee la hoja <code>SALIDA_GESPER</code> del archivo del anexo, valida cada destino en GESPER y muestra una vista previa antes de escribir. Solo escribe en celdas de entrada sin fórmula.</p></div>';
    h += '<div class="flag">Antes de elegirlo: abre el anexo en Excel, usa <b>Datos › Actualizar todo</b> y guárdalo. El panel lee los valores guardados en el archivo, no los recalcula.</div>';
    h += '<div class="btn-row">' + elegir("Seleccionar archivo del anexo…", "primary lg block") + '</div>';
    $("s-anexos").innerHTML = h + input;
    return;
  }

  if (A.fase === "leyendo" || A.fase === "cargando") {
    h += '<div class="load"><div class="spin"></div>' + (A.fase === "leyendo" ? "Leyendo el anexo y validando destinos…" : "Escribiendo en GESPER…") + '</div>';
    $("s-anexos").innerHTML = h + input;
    return;
  }

  if (A.fase === "error") {
    h += sig("c", "No se pudo leer el anexo", (A.archivo ? "<b>" + esc(A.archivo.nombre) + "</b> · " : "") + esc(A.error));
    h += '<div class="btn-row">' + elegir("Elegir otro archivo", "lg block") + '</div>';
    $("s-anexos").innerHTML = h + input;
    return;
  }

  // vista | resumen
  var ev = A.evaluacion, plan = A.plan, meta = ev.meta;
  var cambian = plan.escrituras.filter(function (e) { return e.cambia; });

  h += '<div class="idc"><div class="co">' + esc(A.archivo.nombre) + '</div><div class="mt">' +
    'Anexo ' + esc(meta.ANEXO_ID) + ' · programas ' + esc(meta.PROGRAMAS || "—") + ' · especificación ' + esc(meta.SPEC_VERSION) + '<br>' +
    (A.archivo.modificado ? 'Guardado ' + esc(A.archivo.modificado.toLocaleString("es-CO")) + ' · ' : '') +
    (A.archivo.tamano >= 1048576 ? fmt(A.archivo.tamano / 1048576, 1) + ' MB' : Math.max(1, Math.round(A.archivo.tamano / 1024)) + ' KB') + '</div></div>';

  if (A.fase === "resumen") {
    var r = A.resultado;
    h += sig("k", "Carga completada", "Se escribieron <b>" + r.escritos + "</b> valores en <code>" + esc(anexo.hojaDestino) + "</code>" +
      " (" + r.cambiados + (r.cambiados === 1 ? " cambió" : " cambiaron") + ") y GESPER recalculó una vez." +
      (plan.omitidos.length ? " Se omitieron <b>" + plan.omitidos.length + "</b>." : ""));
    if (cambian.length) {
      h += '<details class="anx-det"><summary>' + (cambian.length === 1 ? 'Ver el valor que cambió' : 'Ver los ' + cambian.length + ' valores que cambiaron') +
        '</summary>' + listaCambios(cambian) + '</details>';
    }
    if (plan.omitidos.length) {
      h += '<details class="anx-det"><summary>Ver los ' + plan.omitidos.length + ' datos omitidos y por qué</summary>' + listaOmitidos(plan.omitidos) + '</details>';
    }
    h += listaPendientes(ev);
    if (r.cambiados) {
      h += '<div class="flag">Cambiaron cantidades de ' + esc(meta.PROGRAMAS || "") + '. Si las alternativas ya se habían optimizado, vuelve a ejecutar el motor para que el consolidado refleje los nuevos B/C.</div>';
    }
    h += '<div class="btn-row">' +
      '<button class="act lg" data-hoja="' + esc(anexo.hojaDestino) + '" data-addr="' + esc(plan.escrituras.length ? plan.escrituras[0].celda : "A1") + '">Ir a la hoja</button>' +
      elegir("Cargar otro archivo", "ghost lg") + '</div>';
    $("s-anexos").innerHTML = h + input;
    return;
  }

  // Estado de lectura
  if (plan.bloqueo) h += sig("c", "Carga bloqueada", esc(plan.bloqueo));
  else if (ev.estadoGlobal === "OK") h += sig("k", "Anexo leído y verificado", "Estado global del anexo: <b>OK</b>.");
  else h += sig("w", "Anexo con advertencias", "Estado global del anexo: <b>" + esc(meta.ESTADO_GLOBAL) + "</b>. Revisa las advertencias antes de cargar.");
  if (A.error) h += sig("c", "La carga no se completó", esc(A.error));

  h += '<div class="kgrid">' +
    '<div class="kpi"><div class="kl">Datos en el anexo</div><div class="kv">' + ev.datos.length + '</div><div class="ks">de ' + ev.totalFilas + ' filas</div></div>' +
    '<div class="kpi"><div class="kl">Listos para cargar</div><div class="kv">' + plan.escrituras.length + '</div><div class="ks">' + esc(anexo.hojaDestino) + '</div></div>' +
    '<div class="kpi"><div class="kl">Cambian en GESPER</div><div class="kv">' + cambian.length + '</div><div class="ks">el resto ya coincide</div></div>' +
    '<div class="kpi"><div class="kl">Omitidos</div><div class="kv' + (plan.omitidos.length ? ' text-critical' : '') + '">' + plan.omitidos.length + '</div><div class="ks">no se escriben</div></div>' +
    '</div>';

  // Advertencias y problemas
  if (plan.avisos.length || ev.controlesRevisar.length || plan.omitidos.length || cambian.length) h += '<div class="sec">Advertencias y problemas</div>';
  if (cambian.length) {
    h += sig("w", (cambian.length === 1 ? "1 valor cambia" : cambian.length + " valores cambian") + " en GESPER",
      "El anexo reemplaza lo que hoy tiene el libro. Revisa cada diferencia en la vista previa antes de confirmar.");
  }
  plan.avisos.forEach(function (a) {
    h += sig("w", (a.tipo === "HORIZONTE" ? "Año fuera del horizonte · " : "Advertencia · ") + idCorto(a.fila.ID),
      esc(a.fila.CONCEPTO) + ' <code>' + esc(a.fila.DESTINO_CELDA) + '</code><br>' + esc(a.texto));
  });
  ev.controlesRevisar.forEach(function (c) {
    h += sig("c", "Control en REVISAR · " + idCorto(c.ID), esc(c.CONCEPTO) + " · valor " + esc(c.VALOR) + " · control " + esc(c.VALOR_CONTROL));
  });
  if (plan.omitidos.length) {
    h += sig("c", plan.omitidos.length + " datos se omitirán", "No se escriben en GESPER. El panel no corrige destinos: revisa el anexo o la plantilla.");
    h += '<details class="anx-det"><summary>Ver detalle</summary>' + listaOmitidos(plan.omitidos) + '</details>';
  }
  var nDec = plan.escrituras.filter(function (e) { return dependeDecision(e.fila); }).length;
  if (nDec) h += '<div class="flag">' + nDec + ' de los datos a cargar dependen de una decisión del profesional (columna <code>DEPENDE_DECISION</code>). Se marcan con • en la vista previa.</div>';

  // Vista previa
  h += '<div class="sec">Vista previa · ' + plan.escrituras.length + ' datos</div>';
  if (cambian.length) {
    h += '<p class="note">Valores que cambian en GESPER (actual → nuevo):</p>' + listaCambios(cambian);
    if (plan.escrituras.length > cambian.length) {
      h += '<p class="note">Los otros ' + (plan.escrituras.length - cambian.length) + ' valores coinciden con los que hoy tiene GESPER.</p>';
    }
  }
  else if (plan.escrituras.length) h += '<div class="card"><p>Los ' + plan.escrituras.length + ' valores coinciden con los que hoy tiene GESPER.</p></div>';
  if (plan.escrituras.length) {
    h += '<details class="anx-det"><summary>Ver los ' + plan.escrituras.length + ' datos a cargar</summary>' + tablaEscrituras(plan.escrituras) + '</details>';
  }
  h += listaPendientes(ev);

  // Confirmación
  var puede = !plan.bloqueo && plan.escrituras.length > 0;
  if (puede && plan.requiereConfirmacion) {
    var revise = [];
    if (cambian.length) revise.push(cambian.length === 1 ? "el valor que cambia" : "los " + cambian.length + " valores que cambian");
    if (plan.avisos.length || ev.estadoGlobal === "OK CON ADVERTENCIAS") revise.push("las advertencias");
    h += '<label class="chk"><input type="checkbox" id="anxConfirma"' + (A.confirmado ? " checked" : "") + '> Revisé ' +
      revise.join(" y ") + ' y quiero cargar estos datos</label>';
  }
  var habilitado = puede && (!plan.requiereConfirmacion || A.confirmado);
  h += '<div class="btn-row">' +
    '<button class="act primary lg grow" id="btnCargarAnexo"' + (habilitado ? "" : " disabled") + '>' +
    (plan.bloqueo ? "Carga bloqueada" : (plan.escrituras.length ? "Cargar " + plan.escrituras.length + " datos en GESPER" : "No hay datos para cargar")) + '</button>' +
    '<button class="act ghost lg" id="btnAnexoCancelar">Cancelar</button></div>';
  h += '<p class="note">Un solo lote de escritura, un recálculo. Nunca sobre celdas con fórmula.</p>';
  $("s-anexos").innerHTML = h + input;
}

/* ---------- PANTALLA: DIAGNÓSTICO ---------- */
export function pintaDiag() {
  var h = '<div class="sec">Diagnóstico de conexión</div>';
  h += '<div class="card"><p>Comprueba que el panel puede leer y escribir en este libro, usando una hoja de trabajo propia y desechable (<code>GESPER_DIAG_TEMP</code>). No toca ninguna hoja del modelo.</p></div>';
  h += '<button class="act primary lg block" id="btnDiag" style="margin-bottom:8px">Ejecutar prueba de conexión</button>';
  h += '<div id="diagOut"></div>';
  $("s-diag").innerHTML = h;
}

export function pintaResultadoDiag(r) {
  var out = $("diagOut");
  if (!out) return;
  var h = '<div class="kpi"><div class="kl">Hojas detectadas en el libro</div><div class="kv">' + r.hojas + '</div></div>';
  h += sig(r.escrituraOk ? "k" : "c", "Escritura", r.escrituraOk ? "Se escribió un valor de prueba en la hoja temporal." : "No se pudo escribir.");
  h += sig(r.lecturaOk ? "k" : "c", "Lectura de vuelta", r.lecturaOk ? "El valor leído coincide con el escrito: ida y vuelta correcta." : "El valor leído no coincide.");
  h += sig(r.limpiado ? "k" : "w", "Limpieza", r.limpiado ? "La hoja temporal de diagnóstico fue eliminada." : "La hoja temporal quedó en el libro para inspección.");
  out.innerHTML = h;
}

/* ---------- ORQUESTACIÓN ---------- */
export function pintaTodo() {
  try { pintaEstado(); } catch (e) { $("s-estado").innerHTML = errBox(e); }
  try { pintaNavegar(); } catch (e) { $("s-navegar").innerHTML = errBox(e); }
  try { pintaDatos(); } catch (e) { $("s-datos").innerHTML = errBox(e); }
  try { pintaAnexos(); } catch (e) { $("s-anexos").innerHTML = errBox(e); }
  try { pintaResultados(); } catch (e) { $("s-result").innerHTML = errBox(e); }
  try { pintaAlt(); } catch (e) { $("s-alt").innerHTML = errBox(e); }
  try { pintaDiag(); } catch (e) { $("s-diag").innerHTML = errBox(e); }
}
