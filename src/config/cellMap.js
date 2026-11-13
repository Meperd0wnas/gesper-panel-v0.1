/**
 * MAPA DE CELDAS — única fuente de verdad del panel para direcciones del libro.
 * Cambiar de versión de plantilla = cambiar este archivo, nada más.
 * Todas las direcciones fueron verificadas contra CASO_Bucaramanga_plantilla_v2.5.xlsm.
 *
 * Este archivo es "business mapping": traduce nombres funcionales (qué es cada dato)
 * a direcciones físicas del libro (dónde vive ese dato hoy). No contiene lógica de
 * cálculo ni llamadas a Office.js.
 */

export const H = {
  ALT: "Alternativas",
  PAR: "PARÁMETROS",
  EVT: "EV PROG TÉCNICOS",
  EVC: "EV PROG CCIALES",
  FCJ: "FCAJA PROYECTO TASA 16,28",
  RFI: "RESUMEN FINANCIERO",
  CON: "Consolidado Programas"
};

export const LECTURAS = [
  ["prog", H.ALT, "A6:E24"],   // tabla PROGRAMAS: nombre, n, VNA ing, VNA egr, B/C
  ["ct005", H.ALT, "Y27:Y29"],  // controles de integridad
  ["monto", H.ALT, "B4"],
  ["minbc", H.ALT, "E4"],
  ["selA1", H.ALT, "I4"],
  ["bcA1", H.ALT, "I5"],
  ["meta", H.ALT, "S3"],
  ["flagsV", H.ALT, "I7:AA7"],   // banderas de vista (Alternativa 1)
  ["flags2", H.ALT, "I49:AA49"], // banderas que alimentan el consolidado (Alternativa 2)
  ["flags3", H.ALT, "I91:AA91"],
  ["a1bc", H.ALT, "K4:K5"],
  ["a2txt", H.ALT, "I46:I47"],
  ["a2bc", H.ALT, "K46:K47"],
  ["a3txt", H.ALT, "I88:I89"],
  ["a3bc", H.ALT, "K88:K89"],
  ["ver1", H.ALT, "T22"],
  ["ver2", H.ALT, "T63"],
  ["ver3", H.ALT, "T105"],
  ["anios", H.ALT, "J9:V9"],
  ["ianc1", H.ALT, "J17:V17"],
  ["ipuf1", H.ALT, "J20:V20"],
  ["ianc2", H.ALT, "J59:V59"],
  ["ipuf2", H.ALT, "J62:V62"],
  ["par", H.PAR, "C4:C18"],
  ["parE", H.PAR, "E8:E9"],
  ["grad", H.PAR, "C33"],
  ["fcaja", H.FCJ, "C65:C67"],
  ["tirm", H.FCJ, "D66"],
  ["empresa", H.EVC, "F2"],
  ["metodo", H.EVC, "C34"],
  ["p9a", H.EVC, "C181"],
  ["p9b", H.EVC, "C200"],
  ["p9c", H.EVC, "C246"],
  ["p9rep", H.EVC, "F197:J197"],
  ["p9nue", H.EVC, "F198:J198"],
  ["p9anios", H.EVC, "F190:J190"]
];

/** Bloques de navegación */
export const BLOQUES = [
  { n: "Configuración global", s: "PARÁMETROS · INDICES · BANREPÚBLICA · PROYEC SUSC", ir: [H.PAR, "B3"] },
  { n: "Catálogos y costos unitarios", s: "C UNIT PROG TÉCNICOS · C UNIT PROG CCIALES", ir: ["C UNIT PROG TÉCNICOS", "B8"] },
  { n: "Componente técnico", s: "EV PROG TÉCNICOS · PROG TÉCNICOS", ir: [H.EVT, "D108"] },
  { n: "Componente comercial", s: "EV PROG CCIALES · PROG CCIALES", ir: [H.EVC, "C23"] },
  { n: "Consolidación financiera", s: "RESUMEN FINANCIERO · 19 hojas B/C", ir: [H.RFI, "B2"] },
  { n: "Alternativas y resultados", s: "Alternativas · Consolidado · FCAJA · EVOLUCION BH", ir: [H.ALT, "A4"] }
];

/**
 * Anclas verificadas de tramos de entrada por programa: las 19, para que
 * ningún programa dependa de buscar su rótulo. En EV PROG TÉCNICOS el número
 * del programa está en la columna A y el nombre en la B (fila 492: "3 |
 * GESTIÓN DEL PLANO DE PRESIONES"), así que buscar "3." en la B no lo
 * encuentra. Filas de P.2–P.8, P.16 y P.19 verificadas en V8_030726 y en
 * CASO_Bucaramanga_plantilla_v2.5 (iguales en ambos).
 */
export const ANCLAS = {
  1: [H.EVT, "D108:D109", "Macromedición · captación"],
  2: [H.EVT, "B358", "Sectorización"],
  3: [H.EVT, "B492", "Gestión del plano de presiones"],
  4: [H.EVT, "B634", "Impermeabilización de tanques"],
  5: [H.EVT, "B652", "Atención de daños"],
  6: [H.EVT, "B671", "Búsqueda sistemática de fugas"],
  7: [H.EVT, "B790", "Reposición de redes y acometidas"],
  8: [H.EVT, "B826", "Vinculación individual · parte técnica"],
  9: [H.EVC, "C181", "Renovación de medidores"],
  10: [H.EVC, "B275", "Gestión de fraudes"],
  11: [H.EVC, "B293", "Cobro de consumos por obras"],
  12: [H.EVC, "B310", "Seguimiento consumo 0"],
  13: [H.EVC, "B328", "Seguimiento de anomalías"],
  14: [H.EVC, "B346", "Medidores grandes consumidores"],
  15: [H.EVC, "B364", "Catastro de usuarios"],
  16: [H.EVC, "B382", "Vinculación de comunidades ilegales"],
  17: [H.EVC, "B401", "Telemetría grandes consumidores"],
  18: [H.EVC, "B417", "Software de inteligencia"],
  19: [H.EVC, "B430", "Vinculación individual · parte comercial"]
};

/**
 * Anexos técnicos que el panel sabe cargar desde su hoja SALIDA_GESPER
 * (contrato: Especificacion_SALIDA_GESPER_v1, ver domain/anexos/salidaGesper.js).
 * Las celdas destino las trae cada fila del anexo (DESTINO_CELDA); aquí solo
 * vive lo que el panel necesita saber de GESPER para validarlas:
 *  - filaAnios: fila de la hoja destino con el índice relativo de año (0..10)
 *    de cada columna, que se compara con ANIO_IDX;
 *  - filaAniosCalendario: fila con el año calendario; solo se muestra en la
 *    vista previa, nunca se usa para validar ni se fija en el código;
 *  - alcanceControles: qué invalida un CONTROL en REVISAR — "concepto" (spec
 *    A01 §6) o "grupo" (spec A02 §6);
 *  - filasBandera: filas «¿contemplado?», que nunca se escriben desde un anexo;
 *  - celdasProtegidas: otras celdas que el anexo nunca escribe (horizontes);
 *  - banderasHorizonte: para cada bloque de filas destino, su fila de bandera.
 *    Un año en "N/A" queda fuera del horizonte y GESPER descarta en silencio
 *    la cantidad: el panel lo avisa (spec §6 paso 7, §7).
 * Verificado contra MODELO GESTIÓN PÉRDIDAS _V8_030726.xlsm y
 * CASO_Bucaramanga_plantilla_v2.5.xlsm (mismas filas en ambos).
 */
export const ANEXOS = [
  {
    id: "A01",
    nombre: "Anexo 1 — Macromedición (P1)",
    programas: ["P1"],
    hojaDestino: H.EVT,
    filaAnios: 2,
    filaAniosCalendario: 3,
    alcanceControles: "concepto",
    filasBandera: [116, 130, 142, 154, 166, 182, 201, 215, 227, 239, 251, 267, 286, 300, 311, 322, 333, 349],
    banderasHorizonte: [
      { desde: 113, hasta: 115, fila: 116 }, // captación · instalación
      { desde: 177, hasta: 181, fila: 182 }, // captación · reemplazo
      { desde: 198, hasta: 200, fila: 201 }, // entrada PTAP · instalación
      { desde: 262, hasta: 266, fila: 267 }, // entrada PTAP · reemplazo
      { desde: 283, hasta: 285, fila: 286 }, // salida PTAP · instalación
      { desde: 344, hasta: 348, fila: 349 }  // salida PTAP · reemplazo
    ]
  },
  {
    // Especificacion_SALIDA_GESPER_A02_v1. SITIO es el subprograma (2.1.1,
    // 2.2, ...): el panel lo usa solo como texto de agrupación.
    id: "A02",
    nombre: "Anexo 2 — Sectorización (P2)",
    programas: ["P2"],
    hojaDestino: H.EVT,
    filaAnios: 2,
    filaAniosCalendario: 3,
    alcanceControles: "grupo",
    filasBandera: [369, 387, 404, 415, 426, 437, 452, 467, 483],
    celdasProtegidas: ["D447", "D478"], // tiempo estimado de reemplazo y de aislamiento (spec A02 §8)
    banderasHorizonte: [
      { desde: 363, hasta: 368, fila: 369 }, // 2.1.1 sectores
      { desde: 380, hasta: 386, fila: 387 }, // 2.1.2 tanques
      { desde: 398, hasta: 403, fila: 404 }, // 2.1.3 telemetría
      { desde: 448, hasta: 451, fila: 452 }, // 2.1.7 reemplazo
      { desde: 463, hasta: 466, fila: 467 }, // 2.2 válvulas
      { desde: 477, hasta: 482, fila: 483 }  // 2.3 aislamiento
    ]
  },
  {
    // Especificacion_SALIDA_GESPER_A03 v1.1: tres programas y diez grupos
    // independientes (segundo segmento del ID: ERP, AUT, FIJ, ROT, MOD, DAN,
    // DET, RMEN, RMAT, ACOM). cargaPorGrupo: el usuario elige qué grupos
    // cargar y confirma cada uno que traiga cambios, advertencias o datos en
    // REVISAR (spec A03 §6 pasos 2-3). Los bloques de horizonte van de la fila
    // del título del subprograma a su B/C parcial: en P3, P5 y P6 la bandera
    // queda en medio del bloque (los volúmenes recuperados van debajo de ella).
    id: "A03",
    nombre: "Anexo 3 — Presiones, daños y fugas (P3, P5, P6)",
    programas: ["P3", "P5", "P6"],
    hojaDestino: H.EVT,
    filaAnios: 2,
    filaAniosCalendario: 3,
    alcanceControles: "concepto",
    cargaPorGrupo: true,
    filasBandera: [503, 517, 535, 555, 570, 581, 592, 609, 625, 661, 678, 692, 705, 718],
    // Horizonte y periodos que el anexo no entrega (spec A03 §8, A03.GEN.HORIZONTE),
    // más los «Tiempo estimado para inversión total» de los demás subprogramas.
    celdasProtegidas: ["D494", "D514", "D528", "D545", "D565", "D602", "D619", "D620", "D654", "D689", "D702", "D715"],
    banderasHorizonte: [
      { desde: 493, hasta: 511, fila: 503 }, // 3.1 ERP's
      { desde: 513, hasta: 524, fila: 517 }, // 3.2 gestión automática
      { desde: 527, hasta: 542, fila: 535 }, // 3.3.1 equipos fijos
      { desde: 544, hasta: 562, fila: 555 }, // 3.3.2 equipos rotativos
      { desde: 618, hasta: 632, fila: 625 }, // 3.4 modelo hidráulico
      { desde: 653, hasta: 669, fila: 661 }, // 5.1 atención de daños
      { desde: 672, hasta: 685, fila: 678 }, // 6.1 detección
      { desde: 688, hasta: 699, fila: 692 }, // 6.2.1 redes menores
      { desde: 701, hasta: 712, fila: 705 }, // 6.2.2 redes matrices
      { desde: 714, hasta: 725, fila: 718 }  // 6.2.3 acometidas
    ]
  },
  {
    // Especificacion_SALIDA_GESPER_A04 v1.1 (mismo contrato que el A03): cuatro
    // grupos — IMP (4.1), REB (6.3.1), SEN (6.3.5) y GEN, el total de tanques
    // D638, que alimenta TQTOTAL y con él D729/D730 y la fila 640. Como en el
    // A03, la bandera queda en medio del bloque y el volumen recuperado debajo.
    id: "A04",
    nombre: "Anexo 4 — Impermeabilización y reboses de tanques (P4, P6)",
    programas: ["P4", "P6"],
    hojaDestino: H.EVT,
    filaAnios: 2,
    filaAniosCalendario: 3,
    alcanceControles: "concepto",
    cargaPorGrupo: true,
    filasBandera: [643, 734, 746, 757, 768, 781],
    // Frecuencia de impermeabilización y horizontes que el anexo no entrega
    // (spec A04 §8, A04.GEN.FRECUENCIA y A04.GEN.HORIZONTE).
    celdasProtegidas: ["D637", "D732", "D778"],
    banderasHorizonte: [
      { desde: 635, hasta: 650, fila: 643 }, // 4.1 impermeabilización
      { desde: 728, hasta: 741, fila: 734 }, // 6.3.1 control de nivel
      { desde: 777, hasta: 788, fila: 781 }  // 6.3.5 reemplazo de sensor
    ]
  },
  {
    // Especificacion_SALIDA_GESPER_A07 v1.1 (mismo contrato que el A03): grupos
    // RED (filas 795, 798-800 y 807-809) y ACO (816 y 822). GEN solo trae la
    // referencia D792 y filas NO_DISPONIBLE, así que no tiene casilla.
    id: "A07",
    nombre: "Anexo 7 — Reposición de redes y acometidas (P7)",
    programas: ["P7"],
    hojaDestino: H.EVT,
    filaAnios: 2,
    filaAniosCalendario: 3,
    alcanceControles: "concepto",
    cargaPorGrupo: true,
    filasBandera: [802, 817],
    // Horizontes que el anexo no entrega (spec A07 §8, A07.GEN.HORIZONTE). D796
    // y D815 son fórmulas (COUNT) y L:U de esas filas también; K796 y K815 son
    // el año 0 escrito como valor.
    celdasProtegidas: ["D796", "K796", "D815", "K815"],
    banderasHorizonte: [
      { desde: 792, hasta: 811, fila: 802 }, // reposición de redes: costos 803-806 y beneficio 810
      { desde: 813, hasta: 824, fila: 817 }  // reposición de acometidas: costos 818-821 y beneficio 823
    ]
  }
];
