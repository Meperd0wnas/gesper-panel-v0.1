/**
 * Estado de la sesión del panel, en memoria. Se pierde al recargar la
 * pestaña de tareas; la fuente de verdad sigue siendo el libro de Excel.
 */
export const state = {
  D: {},              // datos leídos del libro, indexados por clave de LECTURAS
  hojasLibro: [],      // nombres de hojas presentes en el libro
  pantalla: "estado",  // pantalla activa: estado | navegar | datos | result | alt | diag
  subResult: "hid",    // subpestaña activa dentro de Resultados
  vistaDatos: "lista"  // lista | p9 (formulario de un programa)
};
