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
 * @param {string[]} [gruposElegidos] solo con `anexo.cargaPorGrupo`: los
 *   grupos que el usuario marcó en esa vista previa. Se escriben solo esos;
 *   marcar un grupo que pide revisión es su confirmación.
 */
export async function ejecutarCarga(evaluacion, anexo, planConfirmado, gruposElegidos) {
  var plan = await prepararCarga(evaluacion, anexo);
  if (plan.bloqueo) throw new Error(plan.bloqueo);
  if (plan.requiereConfirmacion && !(planConfirmado && firmaPlan(planConfirmado) === firmaPlan(plan))) {
    var e = new Error(planConfirmado
      ? "GESPER cambió desde la vista previa. No se escribió nada: revisa de nuevo los valores y confirma."
      : "Esta carga necesita tu confirmación. No se escribió nada.");
    e.planActual = plan;
    throw e;
  }
  var escrituras = plan.escrituras;
  if (anexo.cargaPorGrupo) {
    var elegidos = gruposElegidos || [];
    escrituras = escrituras.filter(function (x) { return elegidos.indexOf(x.grupo) >= 0; });
    if (!escrituras.length) throw new Error("No marcaste ningún grupo para cargar. No se escribió nada.");
  }
  if (escrituras.length) {
    await escribirCampos(escrituras.map(function (x) {
      return { hoja: x.hoja, direccion: x.celda, valores: [[x.valor]] };
    }));
  }
  // Las referencias (CONTROL con destino, que en GESPER son fórmulas) se
  // vuelven a leer ya recalculadas, para compararlas con lo que se cargó. Si
  // esa lectura falla, la carga ya está hecha: se informa sin el dato.
  var referencias = plan.referencias;
  if (escrituras.length && referencias.length) {
    try {
      var despues = await leerCeldas(anexo.hojaDestino, referencias.map(function (r) { return r.celda; }));
      referencias = referencias.map(function (r) {
        return Object.assign({}, r, { despues: (despues.celdas[r.celda] || {}).valor });
      });
    } catch (e) { /* se muestran los valores de la vista previa */ }
  }
  return {
    plan: plan,
    escrituras: escrituras,
    grupos: anexo.cargaPorGrupo ? (gruposElegidos || []).slice() : null,
    referencias: referencias,
    escritos: escrituras.length,
    cambiados: escrituras.filter(function (x) { return x.cambia; }).length
  };
}
