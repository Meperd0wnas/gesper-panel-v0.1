/**
 * Orquestación del panel: arranque, eventos de la UI y coordinación entre la
 * capa de Excel (services/excelService.js) y el renderizado (render.js).
 */

import { H, ANCLAS } from "../config/cellMap.js";
import { parseNum } from "../utils/format.js";
import { state } from "./state.js";
import {
  pintaTodo, pintaDatos, pintaFormP9, pintaResultados, pintaResultadoDiag, errBox
} from "./render.js";
import {
  escribirRango, irA as excelIrA, buscarFilaPorEtiqueta, ejecutarDiagnostico
} from "../services/excelService.js";
import { cargarModelo } from "../domain/repositories/modeloRepository.js";

function $(id) { return document.getElementById(id); }
function foot(msg, cls) { var f = $("foot"); f.textContent = msg; f.className = cls || ""; }

async function refrescar() {
  var btn = $("btnRefresh");
  if (btn) btn.classList.add("spin");
  try {
    foot("Leyendo el libro…");
    var modelo = await cargarModelo();
    state.hojasLibro = modelo.hojasLibro;
    state.estado = modelo.estado;
    state.totalProgramas = modelo.programas.length;
    state.D = modelo.raw; // puente temporal -- ver informe de la Fase 2C
    pintaTodo();
    foot("Leído · " + state.hojasLibro.length + " hojas · " + new Date().toLocaleTimeString("es-CO"), "ok");
  } catch (e) {
    foot("Error leyendo el libro: " + (e && e.message ? e.message : e), "err");
    $("s-estado").innerHTML = errBox(e);
  } finally {
    if (btn) btn.classList.remove("spin");
  }
}

async function irA(hoja, addr) {
  try {
    await excelIrA(hoja, addr, state.hojasLibro);
    foot(hoja + (addr ? ("!" + addr) : ""), "ok");
  } catch (e) {
    foot("No pude abrir " + hoja + ": " + (e && e.message ? e.message : e), "err");
  }
}

async function irAEtiqueta(hoja, texto) {
  try {
    var fila = await buscarFilaPorEtiqueta(hoja, texto, state.hojasLibro);
    if (fila) await irA(hoja, "B" + fila);
    else { await irA(hoja); foot("No encontré «" + texto + "»; abrí la hoja", "err"); }
  } catch (e) {
    foot("Error buscando: " + (e && e.message ? e.message : e), "err");
  }
}

async function guardarP9() {
  try {
    foot("Guardando…");
    var pa = parseNum($("f_a").value), pb = parseNum($("f_b").value), pc = parseNum($("f_c").value);
    if (pa === null || pb === null || pc === null) throw new Error("Hay un campo que no es un número válido.");
    if (pc < 1 || pc > 12) foot("Aviso: los meses de recuperación suelen estar entre 1 y 12.", "err");

    var rep = [], nue = [];
    for (var i = 0; i < 5; i++) {
      var r = parseNum($("r" + i).value), n = parseNum($("n" + i).value);
      if (r === null || n === null) throw new Error("Hay un valor de año que no es un número válido.");
      rep.push(r); nue.push(n);
    }
    await escribirRango(H.EVC, "C181", [[pa]]);
    await escribirRango(H.EVC, "C200", [[pb]]);
    await escribirRango(H.EVC, "C246", [[pc]]);
    await escribirRango(H.EVC, "F197:J197", [rep]);
    await escribirRango(H.EVC, "F198:J198", [nue]);
    await refrescar();
    foot("Guardado y recalculado.", "ok");
  } catch (e) { foot(e.message, "err"); }
}

function mostrar(id) {
  state.pantalla = id;
  var s = document.querySelectorAll(".scr");
  for (var i = 0; i < s.length; i++) s[i].className = "scr" + (s[i].id === "s-" + id ? " on" : "");
  var b = document.querySelectorAll("#nav button");
  for (var j = 0; j < b.length; j++) b[j].className = (b[j].getAttribute("data-s") === id ? "on" : "");
}

/* ---------- EVENTOS ---------- */
document.addEventListener("click", function (ev) {
  var t = ev.target;

  var nb = t.closest ? t.closest("#nav button") : null;
  if (nb) { mostrar(nb.getAttribute("data-s")); return; }

  var sr = t.closest ? t.closest("#subR button") : null;
  if (sr) { state.subResult = sr.getAttribute("data-r"); pintaResultados(); return; }

  var go = t.closest ? t.closest("[data-go]") : null;
  if (go) { mostrar(go.getAttribute("data-go")); return; }

  var ir = t.closest ? t.closest("[data-hoja]") : null;
  if (ir) { irA(ir.getAttribute("data-hoja"), ir.getAttribute("data-addr")); return; }

  var vb = t.closest ? t.closest("[data-volver]") : null;
  if (vb) { state.vistaDatos = "lista"; pintaDatos(); return; }

  var pr = t.closest ? t.closest("[data-prog]") : null;
  if (pr) {
    var n = parseInt(pr.getAttribute("data-prog"), 10);
    if (n === 9) { state.vistaDatos = "p9"; mostrar("datos"); pintaDatos(); irA(H.EVC, "C181"); }
    else if (ANCLAS[n]) { irA(ANCLAS[n][0], ANCLAS[n][1]); }
    else { irAEtiqueta(H.EVT, String(n) + "."); }
    return;
  }

  if (t.id === "btnRefresh") { refrescar(); return; }
  if (t.id === "btnGuardar") { guardarP9(); return; }
  if (t.id === "btnDescartar") { pintaFormP9(); foot("Cambios descartados."); return; }

  if (t.id === "btnDiag") {
    t.disabled = true;
    foot("Ejecutando diagnóstico…");
    ejecutarDiagnostico(true).then(function (r) {
      pintaResultadoDiag(r);
      foot("Diagnóstico completo.", "ok");
    }).catch(function (e) {
      foot("Error en el diagnóstico: " + (e && e.message ? e.message : e), "err");
    }).finally(function () { t.disabled = false; });
    return;
  }
});

document.addEventListener("input", function (ev) {
  var t = ev.target;
  if (t.tagName === "INPUT") {
    if (t.closest(".fld")) t.closest(".fld").className = "fld dirty";
    else t.className = "dirty";
  }
});

/* ---------- ARRANQUE ---------- */
Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    document.body.innerHTML = '<div class="load">Este panel solo funciona dentro de Excel.</div>';
    return;
  }
  refrescar();
});
