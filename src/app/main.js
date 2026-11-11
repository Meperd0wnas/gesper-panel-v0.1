/**
 * Orquestación del panel: arranque, eventos de la UI y coordinación entre la
 * capa de Excel (services/excelService.js) y el renderizado (render.js).
 */

import { H, ANCLAS } from "../config/cellMap.js";
import { parseNum } from "../utils/format.js";
import { state } from "./state.js";
import {
  pintaTodo, pintaDatos, pintaFormP9, pintaResultados, pintaResultadoDiag, pintaAnexos, pintaBotonCargaAnexo, errBox
} from "./render.js";
import {
  escribirRango, irA as excelIrA, buscarFilaPorEtiqueta, ejecutarDiagnostico
} from "../services/excelService.js";
import { cargarModelo } from "../domain/repositories/modeloRepository.js";
import {
  anexoPorId, leerArchivoAnexo, prepararCarga, ejecutarCarga
} from "../domain/repositories/anexosRepository.js";
import { seleccionInicial } from "../domain/anexos/salidaGesper.js";

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
    state.resultadosHidraulicos = modelo.resultadosHidraulicos;
    state.resultadosFinancieros = modelo.resultadosFinancieros;
    state.D = modelo.raw; // puente temporal -- ver informe de la Fase 2C/2D
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

/* ---------- CARGAR ANEXOS ---------- */
function reiniciarAnexo(anexoId) {
  var A = state.anexos;
  if (anexoId) A.anexoId = anexoId;
  A.fase = "inicio";
  A.archivo = A.evaluacion = A.plan = A.resultado = A.error = null;
  A.confirmado = false;
  A.grupos = {};
}

/** Nuevo plan en pantalla: la confirmación y los grupos marcados se vuelven a pedir. */
function ponerPlan(plan) {
  var A = state.anexos;
  A.plan = plan;
  A.confirmado = false;
  A.grupos = seleccionInicial(plan);
}

async function leerAnexo(file) {
  var A = state.anexos;
  var anexo = anexoPorId(A.anexoId);
  reiniciarAnexo();
  A.fase = "leyendo";
  A.archivo = { nombre: file.name, tamano: file.size, modificado: file.lastModified ? new Date(file.lastModified) : null };
  pintaAnexos();
  foot("Leyendo " + file.name + "…");
  try {
    A.evaluacion = await leerArchivoAnexo(await file.arrayBuffer(), anexo);
    ponerPlan(await prepararCarga(A.evaluacion, anexo));
    A.fase = "vista";
    foot("Anexo leído · " + A.plan.escrituras.length + " datos listos para cargar" +
      (A.plan.omitidos.length ? " · " + A.plan.omitidos.length + " omitidos" : ""), A.plan.bloqueo ? "err" : "ok");
  } catch (e) {
    A.fase = "error";
    A.error = e && e.message ? e.message : String(e);
    foot("No pude leer el anexo: " + A.error, "err");
  }
  pintaAnexos();
}

async function cargarAnexo() {
  var A = state.anexos;
  var anexo = anexoPorId(A.anexoId);
  if (!A.plan) return;
  // Carga por grupo (A03): marcar un grupo es confirmarlo; se escriben solo los marcados.
  var elegidos = Object.keys(A.grupos).filter(function (g) { return A.grupos[g]; });
  if (anexo.cargaPorGrupo ? !elegidos.length : (A.plan.requiereConfirmacion && !A.confirmado)) return;
  A.fase = "cargando";
  A.error = null;
  pintaAnexos();
  foot("Escribiendo en GESPER…");
  try {
    // La confirmación queda atada al plan que el usuario vio: si GESPER cambió
    // desde la vista previa, ejecutarCarga no escribe y devuelve el plan nuevo.
    A.resultado = anexo.cargaPorGrupo
      ? await ejecutarCarga(A.evaluacion, anexo, A.plan, elegidos)
      : await ejecutarCarga(A.evaluacion, anexo, A.confirmado ? A.plan : null);
    A.plan = A.resultado.plan;
    A.fase = "resumen";
    await refrescar();
    foot("Anexo cargado · " + A.resultado.escritos + " valores escritos · " + A.resultado.cambiados + " cambiaron · recalculado", "ok");
  } catch (e) {
    A.fase = "vista";
    A.error = e && e.message ? e.message : String(e);
    if (e && e.planActual) ponerPlan(e.planActual);
    foot("No se cargó el anexo: " + A.error, "err");
  }
  pintaAnexos();
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

  var ax = t.closest ? t.closest("[data-anexo]") : null;
  if (ax) { reiniciarAnexo(ax.getAttribute("data-anexo")); pintaAnexos(); return; }
  if (t.id === "btnAnexoCancelar") { reiniciarAnexo(); pintaAnexos(); foot("Carga cancelada. No se escribió nada."); return; }
  if (t.id === "btnCargarAnexo") { cargarAnexo(); return; }

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
  if (t.tagName === "INPUT" && t.type === "text") {
    if (t.closest(".fld")) t.closest(".fld").className = "fld dirty";
    else t.className = "dirty";
  }
});

document.addEventListener("change", function (ev) {
  var t = ev.target;
  if (t.id === "anexoArchivo" && t.files && t.files[0]) { var f = t.files[0]; t.value = ""; leerAnexo(f); return; }
  if (t.id === "anxConfirma") {
    state.anexos.confirmado = t.checked;
    pintaBotonCargaAnexo();
    return;
  }
  var grupo = t.getAttribute ? t.getAttribute("data-grupo") : null;
  if (grupo) {
    state.anexos.grupos[grupo] = t.checked;
    var tarjeta = t.closest(".anx-grp");
    if (tarjeta) tarjeta.classList.toggle("on", t.checked);
    pintaBotonCargaAnexo(); // sin repintar: no se cierran los detalles abiertos
  }
});

/* Arrastrar el archivo del anexo al panel: segunda vía que no depende del
   explorador de archivos. Se acepta solo cuando la pantalla muestra el
   selector (#anexoArchivo). preventDefault en todo el documento evita que el
   panel navegue al archivo soltado y se pierda. */
function hayDestinoDrop() { return state.pantalla === "anexos" && !!$("anexoArchivo"); }
function marcaDrop(on) { var z = $("anxDrop"); if (z) z.classList.toggle("over", on); }
document.addEventListener("dragover", function (ev) {
  ev.preventDefault();
  if (ev.dataTransfer) ev.dataTransfer.dropEffect = hayDestinoDrop() ? "copy" : "none";
  marcaDrop(hayDestinoDrop());
});
document.addEventListener("dragleave", function (ev) { if (!ev.relatedTarget) marcaDrop(false); });
document.addEventListener("drop", function (ev) {
  ev.preventDefault();
  marcaDrop(false);
  var f = ev.dataTransfer && ev.dataTransfer.files && ev.dataTransfer.files[0];
  if (f && hayDestinoDrop()) leerAnexo(f);
});

/* ---------- ARRANQUE ---------- */
Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    document.body.innerHTML = '<div class="load">Este panel solo funciona dentro de Excel.</div>';
    return;
  }
  pintaAnexos(); // no depende de la lectura del modelo: queda disponible aunque esta falle
  refrescar();
});
