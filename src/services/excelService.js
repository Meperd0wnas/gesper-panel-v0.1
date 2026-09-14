/**
 * Capa de integración con Excel (Office.js).
 *
 * Toda llamada a Excel.run del panel vive aquí. La UI (src/app) nunca llama a
 * Office.js directamente: pide datos o pide una acción a este módulo, y este
 * módulo es el único que conoce Excel.run, Excel.run(...).sync(), etc.
 *
 * Motivos (ver docs/DECISIONES.md):
 *  - Testeamos la lógica de negocio (formato, mapeo de celdas) sin un Excel real.
 *  - Si cambia la API de Office.js, el cambio queda contenido en un solo archivo.
 *  - Ninguna escritura ocurre sobre una celda con fórmula (ver hasFormula).
 */

import { hasFormula } from "../utils/format.js";

/**
 * Lee el nombre de todas las hojas del libro.
 * @returns {Promise<string[]>}
 */
export async function listarHojas() {
  let nombres = [];
  await Excel.run(async (ctx) => {
    const wss = ctx.workbook.worksheets;
    wss.load("items/name");
    await ctx.sync();
    nombres = wss.items.map((s) => s.name);
  });
  return nombres;
}

/**
 * Lee un conjunto de rangos en una sola pasada (una única llamada a ctx.sync()
 * de escritura de la petición y otra de lectura), a partir de una lista de
 * lecturas [clave, hoja, direccion]. Las hojas que no existen en el libro se
 * omiten en silencio: quien llama decide qué hacer con una clave ausente.
 *
 * @param {Array<[string,string,string]>} lecturas
 * @returns {Promise<{hojasLibro: string[], datos: Object}>}
 */
export async function leerLibro(lecturas) {
  let hojasLibro = [];
  const datos = {};
  await Excel.run(async (ctx) => {
    const wss = ctx.workbook.worksheets;
    wss.load("items/name");
    await ctx.sync();
    hojasLibro = wss.items.map((s) => s.name);

    const pendientes = [];
    lecturas.forEach(([clave, hoja, addr]) => {
      if (hojasLibro.indexOf(hoja) < 0) return;
      const rg = wss.getItem(hoja).getRange(addr);
      rg.load("values,text");
      pendientes.push([clave, rg]);
    });
    await ctx.sync();
    pendientes.forEach(([clave, rg]) => {
      datos[clave] = { v: rg.values, t: rg.text };
    });
  });
  return { hojasLibro, datos };
}

/**
 * Escribe valores en un rango, pero nunca sobre una celda que contenga una
 * fórmula. Es la única función del panel con permiso de escritura sobre el
 * modelo. Delega en escribirCampos (una sola entrada) para no duplicar la
 * guarda anti-fórmula ni el recálculo en dos sitios distintos.
 */
export async function escribirRango(hoja, addr, valores) {
  await escribirCampos([{ hoja, direccion: addr, valores }]);
}

/**
 * Escribe varios campos (posiblemente en distintas hojas) en una sola
 * ejecución de Excel.run, con un único recálculo completo al final en vez de
 * uno por campo.
 *
 * Por qué existe (ver docs/ARQUITECTURA_MVP.md §3.3): escribir N campos con
 * escribirRango dispara N recálculos completos de un libro con más de
 * 105.000 fórmulas. escribirCampos agrupa las escrituras de un mismo
 * guardado (p. ej. un formulario de programa) en un solo recálculo.
 *
 * Garantías que conserva de escribirRango:
 *  - ninguna celda con fórmula se sobrescribe (misma guarda hasFormula);
 *  - si CUALQUIER entrada del lote apunta a una celda con fórmula, no se
 *    escribe NINGUNA del lote (todo o nada): las fórmulas de todos los
 *    rangos se leen y se validan antes de asignar ningún valor.
 *
 * @param {Array<{hoja: string, direccion: string, valores: any[][]}>} entries
 *   `valores` sigue la misma forma que espera Range.values (matriz 2D),
 *   igual que el parámetro `valores` de escribirRango — así una entrada
 *   puede ser tanto una celda suelta ([[12.5]]) como una fila de una serie
 *   anual ([[21377, 22576, 15784, 25736, 31539]]).
 */
export async function escribirCampos(entries) {
  if (!Array.isArray(entries) || entries.length === 0) {
    throw new Error("escribirCampos necesita al menos una entrada.");
  }
  entries.forEach((entry, i) => {
    if (!entry || typeof entry.hoja !== "string" || !entry.hoja) {
      throw new Error("La entrada " + i + " de escribirCampos no tiene una hoja válida.");
    }
    if (typeof entry.direccion !== "string" || !entry.direccion) {
      throw new Error("La entrada " + i + " de escribirCampos (" + entry.hoja + ") no tiene una dirección válida.");
    }
    if (!Array.isArray(entry.valores)) {
      throw new Error(
        "La entrada " + i + " de escribirCampos (" + entry.hoja + "!" + entry.direccion +
        ") debe traer 'valores' como una matriz, igual que Range.values."
      );
    }
  });

  await Excel.run(async (ctx) => {
    const rangos = entries.map((entry) => {
      const rg = ctx.workbook.worksheets.getItem(entry.hoja).getRange(entry.direccion);
      rg.load("formulas,address");
      return rg;
    });
    await ctx.sync();

    // Se valida el lote completo antes de escribir nada: si una sola entrada
    // tiene fórmula, ninguna de las demás se escribe tampoco.
    rangos.forEach((rg) => {
      if (hasFormula(rg.formulas)) {
        throw new Error("La celda " + rg.address + " contiene una fórmula. El panel no la sobrescribe.");
      }
    });

    rangos.forEach((rg, i) => { rg.values = entries[i].valores; });
    ctx.workbook.application.calculate(Excel.CalculationType.full);
    await ctx.sync();
  });
}

/**
 * Activa una hoja (mostrándola si estaba oculta) y, opcionalmente, selecciona
 * una celda o rango dentro de ella.
 */
export async function irA(hoja, addr, hojasLibro) {
  if (hojasLibro.indexOf(hoja) < 0) throw new Error("No existe la hoja " + hoja);
  await Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getItem(hoja);
    sh.visibility = Excel.SheetVisibility.visible;
    sh.activate();
    if (addr) sh.getRange(addr).select();
    await ctx.sync();
  });
}

/**
 * Busca en la columna B de una hoja la primera fila cuyo texto contiene
 * `texto` (sin distinguir mayúsculas) y devuelve su número de fila, o 0 si
 * no la encuentra.
 */
export async function buscarFilaPorEtiqueta(hoja, texto, hojasLibro) {
  if (hojasLibro.indexOf(hoja) < 0) throw new Error("No existe la hoja " + hoja);
  let fila = 0;
  await Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getItem(hoja);
    const col = sh.getRange("B1:B900");
    col.load("text");
    await ctx.sync();
    const t = texto.toUpperCase();
    for (let i = 0; i < col.text.length; i++) {
      const c = String(col.text[i][0] || "").toUpperCase();
      if (c.indexOf(t) >= 0) { fila = i + 1; break; }
    }
  });
  return fila;
}

/**
 * Nombre de la hoja de trabajo desechable que usa el diagnóstico de conexión.
 * Nunca se lee ni se escribe en ninguna otra parte del panel.
 */
export const HOJA_DIAGNOSTICO = "GESPER_DIAG_TEMP";

/**
 * Prueba mínima de integración Add-in -> Office.js -> Excel, para verificar el
 * entorno de desarrollo sin tocar ningún dato real del modelo:
 *   1. confirma que hay un libro accesible y cuenta sus hojas;
 *   2. crea (si no existe) una hoja de trabajo propia y oculta, ajena al modelo;
 *   3. escribe un valor de prueba con marca de tiempo en A1;
 *   4. vuelve a leer esa misma celda y confirma que el valor coincide (ida y vuelta);
 *   5. dependiendo de `limpiar`, borra la hoja de prueba o la deja para inspección.
 *
 * No lee ni modifica ninguna hoja del modelo GESPER.
 */
export async function ejecutarDiagnostico(limpiar = true) {
  const token = "GESPER-PING-" + Date.now();
  const resultado = { hojas: 0, escrituraOk: false, lecturaOk: false, valorLeido: null, limpiado: false };

  await Excel.run(async (ctx) => {
    const wss = ctx.workbook.worksheets;
    wss.load("items/name");
    await ctx.sync();
    resultado.hojas = wss.items.length;

    let hoja = wss.getItemOrNullObject(HOJA_DIAGNOSTICO);
    hoja.load("name");
    await ctx.sync();
    if (hoja.isNullObject) {
      hoja = wss.add(HOJA_DIAGNOSTICO);
      hoja.visibility = Excel.SheetVisibility.hidden;
    }

    const celda = hoja.getRange("A1");
    celda.values = [[token]];
    resultado.escrituraOk = true;
    await ctx.sync();

    celda.load("values");
    await ctx.sync();
    resultado.valorLeido = celda.values[0][0];
    resultado.lecturaOk = resultado.valorLeido === token;

    if (limpiar) {
      hoja.delete();
      await ctx.sync();
      resultado.limpiado = true;
    }
  });

  return resultado;
}
