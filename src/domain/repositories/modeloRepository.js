/**
 * Repositorio de lectura del modelo: compone excelService.leerLibro con los
 * mappers de src/domain/mappers para producir un ModeloGesper ya tipado (ver
 * docs/ARQUITECTURA_MVP.md §3.4, §4). Junto con los mappers, es la única
 * pieza que conoce la forma cruda que devuelve excelService (`{v, t}` por
 * clave de LECTURAS) — las pantallas migradas no deberían necesitar leerla.
 *
 * Deliberadamente delgado: solo compone, sin excepciones por programa (esas
 * viven en programasRepository.js cuando exista, ver §3.4).
 */
import { LECTURAS } from "../../config/cellMap.js";
import { leerLibro } from "../../services/excelService.js";
import { mapEstado } from "../mappers/estadoMapper.js";
import { mapProgramas } from "../mappers/programasMapper.js";
import { mapResultadosHidraulicos, mapResultadosFinancieros } from "../mappers/resultadosMapper.js";
import { mapAlternativas, mapDetalleAlternativas, mapPlanGestion } from "../mappers/alternativasMapper.js";

/**
 * Lee el libro una sola vez y devuelve el modelo completo ya mapeado.
 * @returns {Promise<Object>} ModeloGesper
 */
export async function cargarModelo() {
  var lectura = await leerLibro(LECTURAS);
  var hojasLibro = lectura.hojasLibro;
  var datos = lectura.datos;

  return {
    estado: mapEstado(datos),
    programas: mapProgramas(datos),
    resultadosHidraulicos: mapResultadosHidraulicos(datos),
    resultadosFinancieros: mapResultadosFinancieros(datos),
    alternativas: mapAlternativas(datos),
    // Para el informe de resultados (app/informe.js): series por año, programas
    // marcados y veredicto de cada alternativa, y el plan de gestión de pérdidas.
    detalleAlternativas: mapDetalleAlternativas(datos),
    planGestion: mapPlanGestion(datos),
    hojasLibro: hojasLibro,
    leidoEn: new Date(),
    // Puente temporal, fuera del contrato ModeloGesper documentado en
    // docs/ARQUITECTURA_MVP.md §4: las pantallas aún no migradas (Datos,
    // Resultados, Alternativas, Diagnóstico) siguen leyendo raw.v/raw.t
    // directamente en app/render.js. Se retira cuando esas pantallas migren
    // también a mappers -- ver "requiere ajuste del contrato" en el informe
    // de la Fase 2C.
    raw: datos
  };
}
