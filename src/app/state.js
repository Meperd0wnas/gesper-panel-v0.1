/**
 * Estado de la sesión del panel, en memoria. Se pierde al recargar la
 * pestaña de tareas; la fuente de verdad sigue siendo el libro de Excel.
 */
export const state = {
  D: {},              // raw leído del libro -- en uso solo por las pantallas aún no migradas (Datos, Resultados, Alternativas, Diagnóstico)
  estado: null,        // EstadoModelo ya mapeado (domain/mappers/estadoMapper.js), consumido por la pantalla Estado (Fase 2C)
  totalProgramas: 0,   // ModeloGesper.programas.length -- ver "requiere ajuste del contrato" en el informe de la Fase 2C
  hojasLibro: [],      // nombres de hojas presentes en el libro
  pantalla: "estado",  // pantalla activa: estado | navegar | datos | result | alt | diag
  subResult: "hid",    // subpestaña activa dentro de Resultados
  vistaDatos: "lista"  // lista | p9 (formulario de un programa)
};
