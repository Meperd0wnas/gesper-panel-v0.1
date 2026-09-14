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
 * fórmula: antes de escribir, lee Range.formulas y aborta si encuentra alguna.
 * Es la única función del panel con permiso de escritura sobre el modelo.
 */
export async function escribirRango(hoja, addr, valores) {
  await Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getItem(hoja);
    const rg = sh.getRange(addr);
    rg.load("formulas,address");
    await ctx.sync();

    if (hasFormula(rg.formulas)) {
      throw new Error("La celda " + rg.address + " contiene una fórmula. El panel no la sobrescribe.");
    }
    rg.values = valores;
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
