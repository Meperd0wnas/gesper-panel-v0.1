/**
 * Mapea la tabla de programas (raw.prog, Alternativas!A6:E24) al contrato de
 * dominio ProgramaResumen (ver docs/ARQUITECTURA_MVP.md §4). Es la misma
 * lógica que hoy vive en app/render.js:programas(), separada para poder
 * probarla sin DOM ni Excel y para que estadoMapper.js la reutilice.
 */
import { num } from "../../utils/format.js";

var NUMERO_MAX_TECNICO = 8; // P.1-P.8 técnicos, P.9-P.19 comerciales (CLAUDE.md, docs/01)

function tipoDePrograma(numero) {
  return numero <= NUMERO_MAX_TECNICO ? "tecnico" : "comercial";
}

/**
 * bc === 0 sin nota es una ambigüedad real del modelo, no algo que este
 * mapper pueda resolver: no distingue si el programa no aplica o si falta
 * diligenciarlo (ver docs/ARQUITECTURA_MVP.md §8 y la fila P0.2 de §0).
 *
 * Clasificarlo como 'sin_datos' es comportamiento HEREDADO -- se conserva
 * porque es lo que ya hacía el panel (app/render.js:pintaEstado, filtro
 * `sinDatos`) -- y NO una afirmación de negocio de que BC=0 siempre
 * signifique "sin datos" (podría ser, igual de válido, un programa que
 * legítimamente no aplica). Distinguir un caso del otro requiere validación
 * funcional con el dueño del modelo antes de cambiar esta clasificación.
 */
function estadoDePrograma(bc, nota) {
  if (nota) return "no_aplica";
  if (bc === null) return "sin_datos";
  if (bc === 0) return "sin_datos"; // heredado, no concluyente -- requiere validación funcional, ver comentario arriba
  return bc >= 1 ? "viable" : "no_viable";
}

/**
 * @param {Object} raw - lo que devuelve excelService.leerLibro(LECTURAS).datos
 * @returns {Array<Object>} ProgramaResumen[]
 */
export function mapProgramas(raw) {
  if (!raw.prog) return [];
  var out = [];
  for (var i = 0; i < raw.prog.v.length; i++) {
    var fila = raw.prog.v[i];
    var nombre = String(fila[0] || "").trim();
    if (!nombre) continue;
    var numero = num(fila[1]);
    if (numero === null) continue;

    var beneficioVPN = num(fila[2]);
    var costoVPN = num(fila[3]);
    var bc = num(fila[4]);
    var nota = (typeof fila[2] === "string" && fila[2]) ? String(fila[2]) : "";

    out.push({
      numero: numero,
      nombre: nombre,
      tipo: tipoDePrograma(numero),
      beneficioVPN: beneficioVPN,
      costoVPN: costoVPN,
      bc: bc,
      estado: estadoDePrograma(bc, nota),
      nota: nota
    });
  }
  return out;
}
