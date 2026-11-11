/**
 * Carga de anexos técnicos a GESPER: compone el lector de archivos
 * (utils/xlsx.js), el contrato SALIDA_GESPER (domain/anexos/salidaGesper.js)
 * y excelService (lectura de validación + escritura por lotes). Igual que
 * modeloRepository, solo compone: las reglas viven en salidaGesper.js.
 *
 * Flujo: leerArchivoAnexo (solo el archivo, sin tocar GESPER) ->
 * prepararCarga (lee GESPER para validar destinos; vista previa) ->
 * ejecutarCarga (vuelve a validar contra el libro actual y escribe).
 */
import { ANEXOS } from "../../config/cellMap.js";
import { leerXlsx } from "../../utils/xlsx.js";
import { leerCeldas, escribirCampos } from "../../services/excelService.js";
import {
  HOJA_SALIDA, interpretarSalida, evaluarAnexo, direccionesAInspeccionar, planificarCarga, firmaPlan
} from "../anexos/salidaGesper.js";

export function anexoPorId(id) {
  return ANEXOS.filter(function (a) { return a.id === id; })[0] || null;
}

/**
 * Lee el archivo del anexo (bytes elegidos por el usuario) y clasifica sus
 * filas. No lee ni escribe el libro GESPER abierto.
 */
export async function leerArchivoAnexo(buffer, anexo) {
  var libro = await leerXlsx(buffer, { hojas: [HOJA_SALIDA] });
  return evaluarAnexo(interpretarSalida(libro), anexo);
}

/** Valida cada destino contra el libro GESPER abierto, en una sola lectura. */
export async function prepararCarga(evaluacion, anexo) {
  var inspeccion = await leerCeldas(anexo.hojaDestino, direccionesAInspeccionar(evaluacion, anexo));
  return planificarCarga(evaluacion, inspeccion, anexo);
}

/**
 * Vuelve a validar (el usuario pudo editar GESPER desde la vista previa) y
 * escribe todos los valores en un solo lote: una ejecución de Excel.run, un
 * recálculo, y la guarda anti-fórmula de escribirCampos como segunda barrera.
 *
 * @param {Object|null} planConfirmado el plan que el usuario vio y confirmó
 *   (null si no confirmó). Si la carga requiere confirmación y no la hay, o
 *   si lo que hay hoy en GESPER ya no es lo que el usuario vio, no se escribe
 *   nada: se lanza un Error con `planActual` para volver a mostrar la vista previa.
 */
export async function ejecutarCarga(evaluacion, anexo, planConfirmado) {
  var plan = await prepararCarga(evaluacion, anexo);
  if (plan.bloqueo) throw new Error(plan.bloqueo);
  if (plan.requiereConfirmacion && !(planConfirmado && firmaPlan(planConfirmado) === firmaPlan(plan))) {
    var e = new Error(planConfirmado
      ? "GESPER cambió desde la vista previa. No se escribió nada: revisa de nuevo los valores y confirma."
      : "Esta carga necesita tu confirmación. No se escribió nada.");
    e.planActual = plan;
    throw e;
  }
  if (plan.escrituras.length) {
    await escribirCampos(plan.escrituras.map(function (e) {
      return { hoja: e.hoja, direccion: e.celda, valores: [[e.valor]] };
    }));
  }
  return {
    plan: plan,
    escritos: plan.escrituras.length,
    cambiados: plan.escrituras.filter(function (e) { return e.cambia; }).length
  };
}
