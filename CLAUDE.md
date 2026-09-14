# CLAUDE.md

Guía para trabajar en este repositorio con Claude Code. Ver `README.md` para la visión general del proyecto y `docs/DECISIONES.md` para el porqué de las decisiones técnicas.

## Qué es esto

Un Office Add-in (Task Pane) de Excel: la capa de UX/UI para **GESPER**, un modelo de gestión de pérdidas de agua construido íntegramente en Excel + VBA. Excel es y debe seguir siendo la fuente de verdad y el motor de cálculo. El Add-in solo presenta, navega, valida entradas y escribe en celdas sin fórmula — nunca reimplementa un cálculo del modelo.

## Reglas que no se negocian

- **No reescribir el VBA ni migrar cálculos de Excel a JavaScript.** Si una pantalla necesita un número que el modelo ya calcula, se lee esa celda; no se recalcula en JS.
- **Nunca escribir sobre una celda con fórmula.** La guarda vive en `src/utils/format.js` (`hasFormula`) y se aplica en `src/services/excelService.js` (`escribirRango`) antes de cualquier escritura. Si agregas una nueva función de escritura, pasa por ahí, no por `Excel.run` directo.
- **No tocar el `.xlsm` original** salvo que sea estrictamente necesario y quede explícitamente justificado (y aprobado por el usuario).
- **No auditar de nuevo el modelo.** Ya pasó por descubrimiento, documentación, limpieza, revisión y pruebas — está documentado en `docs/`. Si algo del modelo parece raro, primero revisa si ya está explicado ahí antes de asumir que es un error.

## Dónde va cada cosa

| Si necesitas... | Edita... |
|---|---|
| Leer o escribir una celda nueva | `src/config/cellMap.js` (la dirección) + la función correspondiente en `src/services/excelService.js` si el patrón de acceso es nuevo |
| Cambiar cómo se ve una pantalla | `src/app/render.js` |
| Cambiar el flujo de eventos / arranque | `src/app/main.js` |
| Una función de formato/parseo/validación reutilizable | `src/utils/format.js` |
| Actualizar direcciones tras cambiar de versión de plantilla | Solo `src/config/cellMap.js` |

La UI (`src/app/*`) nunca debe importar `Excel` ni llamar `Excel.run` directamente — siempre a través de `src/services/excelService.js`. Es lo que permite probar la lógica sin abrir Excel.

## Domino del modelo (dónde leer antes de asumir algo)

- `docs/01 - Especificación Técnica Oficial.pdf` — qué hace cada módulo funcional, reglas de negocio, qué es dato vs. estructura vs. resultado.
- `docs/04 - Documento de Arquitectura.pdf` — ficha por hoja: qué recibe, qué produce, qué pasa si se modifica mal. Consúltalo antes de mapear una celda nueva en `cellMap.js`.
- `docs/03 - Diagrama de Arquitectura.html` — diagrama de dependencias entre las 47 hojas y los 5 ciclos de retroalimentación (ábrelo en un navegador).

Puntos del modelo que generan confusión si no se conocen:
- El consolidado financiero (`RESUMEN FINANCIERO`) solo refleja la alternativa vigente **después** de que el motor de optimización (VBA) se haya ejecutado — cambiar la selección en la hoja `Alternativas` sin volver a correr el motor deja el consolidado desactualizado, sin ningún aviso de Excel. El panel ya expone esto como una alerta en la pantalla "Estado".
- Los programas técnicos (P1–P8) y comerciales (P9–P19) usan **dos fórmulas de B/C distintas**, no una sola parametrizada — ver `docs/01...pdf` §4.8.
- El programa 8 técnico no tiene evaluación B/C propia: se fusiona con el programa 19 comercial.

## Comandos

```
npm install                 # instalar dependencias
npx office-addin-dev-certs install   # una sola vez, certificado HTTPS local
npm start                   # servir el panel (deja la terminal abierta)
npm run instalar-en-excel   # abrir Excel con el complemento cargado
npm run quitar-de-excel     # desinstalar el sideload
npm test                    # node:test sobre /test
```

No hay paso de build ni linter configurado todavía. No introduzcas un bundler (webpack/vite) sin que el usuario lo pida explícitamente — es una decisión de arquitectura tomada a propósito (ver `docs/DECISIONES.md` §1).

## Al agregar código

- Sigue el estilo `var`/funciones simples ya presente en `render.js`/`main.js` en vez de introducir clases o un framework de componentes — es deliberadamente minimalista para un equipo de ~5 usuarios.
- Toda celda nueva que el panel necesite leer o escribir se agrega a `LECTURAS` en `cellMap.js`, con la clave que se usará en `state.D`.
- Si agregas lógica que no depende de Office.js ni del DOM, ponla en `src/utils/` y agrégale una prueba en `/test` — es la parte del código más barata de probar y la que más se beneficia de tener cobertura.
