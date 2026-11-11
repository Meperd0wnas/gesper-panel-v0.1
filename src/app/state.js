/**
 * Estado de la sesión del panel, en memoria. Se pierde al recargar la
 * pestaña de tareas; la fuente de verdad sigue siendo el libro de Excel.
 */
export const state = {
  D: {},              // raw leído del libro -- en uso solo por las pantallas aún no migradas (Datos, Resultados, Alternativas, Diagnóstico)
  estado: null,        // EstadoModelo ya mapeado (domain/mappers/estadoMapper.js), consumido por la pantalla Estado (Fase 2C)
  totalProgramas: 0,   // ModeloGesper.programas.length -- ver "requiere ajuste del contrato" en el informe de la Fase 2C
  resultadosHidraulicos: null,  // ResultadosHidraulicos ya mapeado, consumido por la pantalla Resultados (Fase 2D)
  resultadosFinancieros: null,  // ResultadosFinancieros ya mapeado, consumido por la pantalla Resultados (Fase 2D)
  hojasLibro: [],      // nombres de hojas presentes en el libro
  pantalla: "estado",  // pantalla activa: estado | navegar | datos | anexos | result | alt | diag
  subResult: "hid",    // subpestaña activa dentro de Resultados
  vistaDatos: "lista", // lista | p9 (formulario de un programa)
  anexos: {            // pantalla Cargar anexos (domain/repositories/anexosRepository.js)
    anexoId: "A01",    // id de ANEXOS (config/cellMap.js) elegido
    fase: "inicio",    // inicio | leyendo | vista | cargando | resumen | error
    archivo: null,     // { nombre, tamano, modificado } del archivo elegido
    evaluacion: null,  // evaluarAnexo(...): filas clasificadas del anexo
    plan: null,        // planificarCarga(...): escrituras, omitidos, avisos
    resultado: null,   // ejecutarCarga(...): resumen de la carga
    error: null,       // mensaje del último error de lectura o carga
    confirmado: false, // el usuario marcó que revisó las advertencias
    grupos: {}         // carga por grupo (A03): { "ERP": true, ... } grupos marcados para cargar
  }
};
