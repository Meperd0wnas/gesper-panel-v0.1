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
 * Anclas verificadas de tramos de entrada por programa.
 * Donde no hay ancla verificada, el panel busca la fila por su rótulo.
 */
export const ANCLAS = {
  1: [H.EVT, "D108:D109", "Macromedición · captación"],
  9: [H.EVC, "C181", "Renovación de medidores"],
  10: [H.EVC, "B275", "Gestión de fraudes"],
  11: [H.EVC, "B293", "Cobro de consumos por obras"],
  12: [H.EVC, "B310", "Seguimiento consumo 0"],
  13: [H.EVC, "B328", "Seguimiento de anomalías"],
  14: [H.EVC, "B346", "Medidores grandes consumidores"],
  15: [H.EVC, "B364", "Catastro de usuarios"],
  17: [H.EVC, "B401", "Telemetría grandes consumidores"],
  18: [H.EVC, "B417", "Software de inteligencia"]
};
