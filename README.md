# Panel GESPER — Office Add-in para Excel

Complemento (Task Pane Add-in) de Excel para **GESPER**, el modelo de gestión de pérdidas de agua usado por operadores de acueducto y alcantarillado en Colombia. GESPER evalúa 19 programas de reducción de pérdidas (8 técnicos, 11 comerciales), calcula su relación beneficio/costo y su impacto tarifario según la Resolución CRA 688/2014, y usa un motor de optimización combinatoria (VBA) para proponer hasta tres alternativas de portafolio de inversión.

**El modelo Excel es y sigue siendo la fuente de verdad.** Este Add-in no recalcula nada: lee el libro, presenta lo que ya calculó de forma más navegable, valida entradas antes de escribirlas, y escribe de vuelta solo en las celdas de entrada que no tienen fórmula.

## Qué problema resuelve el Add-in

El libro tiene ~47 hojas, la mayoría ocultas, con un flujo de diligenciamiento secuencial estricto y un ciclo de retroalimentación no trivial entre el motor de optimización y el consolidado financiero (ver `docs/04 - Documento de Arquitectura.pdf`, sección 5). El Add-in existe para:

- mostrar el estado del modelo sin recorrer las 47 hojas a mano (WACC diligenciado, alternativa vigente, programas sin datos, IPC desactualizado);
- navegar directamente al bloque correcto en vez de buscar entre hojas ocultas;
- diligenciar programas puntuales con validación antes de escribir en el libro;
- presentar resultados (indicadores hídricos, comparación de alternativas, flujo de caja) con la celda de origen siempre visible.

Quiénes lo usan: personal interno de la empresa prestadora que opera el modelo (no un framework de uso público) — del orden de 5 usuarios.

## Arquitectura

```
                    USUARIO
                       │
                       ▼
              ┌─────────────────┐
              │   OFFICE ADD-IN │   src/app       (UI, orquestación)
              │                 │   src/config    (mapeo de celdas)
              └────────┬────────┘
                       │
                    Office.js         src/services/excelService.js
                       │              (única capa que llama a Excel.run)
                       ▼
              ┌─────────────────┐
              │     EXCEL       │   CASO_*.xlsm — datos, fórmulas, VBA
              └─────────────────┘
```

- **`src/services/excelService.js`** — toda llamada a `Excel.run` vive aquí. Lee rangos, escribe con protección contra sobrescribir fórmulas, navega a hojas/celdas, y expone un diagnóstico de conexión.
- **`src/config/cellMap.js`** — mapeo entre "qué es cada dato" y "en qué celda vive hoy". Es el único archivo que hay que tocar si cambia la versión de la plantilla.
- **`src/app/`** — estado en memoria (`state.js`), renderizado de cada pantalla (`render.js`) y orquestación/eventos (`main.js`). Ninguno de estos archivos llama a Office.js directamente.
- **`src/utils/format.js`** — funciones puras (formato de números, escape de HTML, parseo de entrada, la guarda `hasFormula`), sin dependencias de Office.js ni del DOM — por eso se pueden probar con Node normal.

Ver `docs/DECISIONES.md` para el razonamiento detrás de estas decisiones (por qué sin framework, por qué este mapeo de estado, qué se probó y qué no).

## Requisitos

- **Windows** con **Excel de escritorio** (Microsoft 365 o Excel 2019+). El manifiesto apunta a `Workbook` y usa `ExcelApi 1.1`.
- **Node.js 20 o superior** (el repo se desarrolló y probó con Node 22).
- Una copia del libro GESPER (`CASO_*.xlsm` o `docs/GESPER_CANONICO_PLANTILLA_v2.5.xlsm`) abierta en Excel.

## Instalación

```
npm install
npx office-addin-dev-certs install
```

El segundo comando instala y confía en el certificado HTTPS local de desarrollo de Office — Excel exige HTTPS para cargar el complemento, incluso en desarrollo local. Windows va a pedir confirmar que confías en el certificado.

## Ejecución local

**Terminal 1 — servidor:**

```
npm start
```

Sirve el panel en `https://localhost:3000/src/taskpane.html`. Déjala abierta.

**Terminal 2 — cargar el complemento en Excel:**

```
npm run instalar-en-excel
```

Abre Excel con el complemento ya registrado (sideloading vía `office-addin-debugging`). Abre tu `.xlsm` y busca el botón **Panel GESPER** en la pestaña Inicio de la cinta.

Para detener: `Ctrl+C` en la terminal del servidor, y `npm run quitar-de-excel`.

Si el sideloading automático falla (pasa según la versión de Office), `INSTRUCCIONES.md` documenta el método manual vía catálogo de complementos compartido — pensado para un usuario no técnico, léelo si vas a instalarlo en la máquina de alguien del equipo.

> **Antes de nada: trabajá sobre una copia del `.xlsm`.** El panel escribe en celdas reales cuando se pulsa "Guardar". No pruebes sobre el archivo canónico ni sobre datos de un cliente real.

## Cómo construir

No hay paso de build: es HTML + CSS + JavaScript (módulos ES nativos) servido directamente. `npm start` es, a la vez, "build" y "run" — no hay artefacto intermedio que generar.

## Pruebas

```
npm test
```

Corre la suite de `node:test` (incorporada en Node, sin dependencias nuevas) sobre `/test`: funciones de formato, integridad del mapeo de celdas, y la capa de Excel usando un doble de prueba de `Excel.run` (sin necesidad de tener Excel abierto). Ver `docs/DECISIONES.md` §7 para qué queda fuera de esta suite y por qué.

Para probar la comunicación real Add-in ↔ Excel (que no se puede automatizar sin Excel abierto), usa la pestaña **Diagnóstico** del panel: lee el conteo de hojas del libro y hace una escritura + lectura de ida y vuelta sobre una hoja temporal propia (`GESPER_DIAG_TEMP`), sin tocar ninguna hoja del modelo.

## Estructura del proyecto

```
gesper-panel-v0.1/
  manifest.xml              # lo que Excel lee para saber que el complemento existe
  server.js                 # servidor local HTTPS
  package.json
  src/
    taskpane.html            # shell: HTML + CSS + <script type="module">
    config/cellMap.js        # mapeo de negocio: direcciones de celdas
    services/excelService.js # única capa que llama a Office.js
    app/                     # estado, renderizado, orquestación
    utils/format.js          # funciones puras
  test/                      # pruebas con node:test
  assets/                    # iconos del complemento
  docs/                      # documentación del modelo GESPER y decisiones técnicas
  CASO_Bucaramanga_plantilla_v2.5.xlsm   # caso de referencia (ver nota de privacidad abajo)
  INSTRUCCIONES.md           # guía de instalación para un usuario no técnico
  CLAUDE.md                  # convenciones del repo para trabajar con Claude Code
```

## Dónde consultar la documentación

- **`INSTRUCCIONES.md`** — cómo instalar y usar el panel, en lenguaje no técnico. Es lo que le compartirías a alguien del equipo que solo necesita usarlo.
- **`docs/01 - Especificación Técnica Oficial.pdf`** — qué hace GESPER, sus 9 módulos funcionales, reglas de negocio, riesgos técnicos conocidos y recomendaciones de una eventual reimplementación. Léelo antes de tocar cualquier mapeo de celdas.
- **`docs/04 - Documento de Arquitectura.pdf`** — ficha hoja por hoja del libro: qué recibe, qué produce, qué pasa si se modifica mal. Es la referencia para saber si una celda es dato de entrada, parámetro, resultado o estructura fija.
- **`docs/03 - Diagrama de Arquitectura.html`** — diagrama interactivo (Mermaid) de las 47 hojas, sus dependencias y los 5 ciclos de retroalimentación del modelo. Ábrelo en un navegador.
- **`docs/GESPER_Inventario_Entradas.xlsx`** — inventario de celdas de entrada del modelo.
- **`docs/DECISIONES.md`** — decisiones técnicas de este Add-in y su porqué.

Si encuentras una contradicción entre estos documentos, la Especificación Técnica es v2.0 y el Documento de Arquitectura es v3.0 (agosto 2026) — el segundo es la fuente más reciente.

## Cómo empezar a desarrollar

1. Lee `docs/01 - Especificación Técnica Oficial.pdf` (secciones 1 a 4) para entender qué hace el modelo y qué produce.
2. Corre el panel siguiendo "Ejecución local" arriba y usa la pestaña Diagnóstico para confirmar que tu entorno está bien configurado.
3. Si vas a agregar una pantalla o un dato nuevo: la dirección de la celda va en `src/config/cellMap.js`; la lectura/escritura la hace `src/services/excelService.js`; la pantalla la pinta una función nueva en `src/app/render.js`, listada en `pintaTodo()`.
4. Si vas a cambiar de versión de plantilla: `src/config/cellMap.js` es el único archivo que debería necesitar cambios de direcciones.
