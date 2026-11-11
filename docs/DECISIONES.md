# Decisiones técnicas — Panel GESPER (Office Add-in)

Registro de las decisiones tomadas al preparar el repositorio para desarrollo incremental del Add-in. No repite lo obvio ni lo ya evidente en el código; solo lo que hubiera sido razonable resolver de otra forma.

## 1. Framework: ninguno (se mantiene la decisión previa)

El proyecto ya traía la decisión, documentada en `INSTRUCCIONES.md`, de no usar React/Vue/webpack ni paso de compilación: HTML + CSS + JavaScript nativo, servido tal cual por un servidor HTTPS propio (`server.js`). Se mantiene esa decisión — no se encontró un problema técnico real que la justifique cambiar, y el propio motivo original ("que lo pueda mantener una sola persona sin herramientas adicionales") sigue vigente para un equipo de ~5 usuarios internos.

Lo que sí cambia es **cómo se organiza** ese JavaScript: antes vivía entero dentro de un único `<script>` en `taskpane.html` (~640 líneas mezclando UI, mapeo de celdas y llamadas a Office.js). Ahora se reparte en módulos ES nativos del navegador (`<script type="module">`), sin ningún bundler. Excel de escritorio (WebView2 en Windows, Safari en Mac) soporta módulos ES nativos, así que esto no reintroduce un paso de compilación.

## 2. Estructura de carpetas

Se adaptó la referencia conceptual del enunciado (components/services/hooks/...) a un proyecto sin framework y sin componentes reutilizables tipo React:

```
src/
  taskpane.html        # shell: HTML + CSS + <script type="module">
  config/cellMap.js     # mapeo de negocio: direcciones de celdas del libro
  services/excelService.js  # única capa que llama a Office.js
  app/
    state.js            # estado de la sesión del panel (en memoria)
    render.js           # construcción del HTML de cada pantalla
    main.js             # arranque, eventos, orquestación
  utils/format.js        # funciones puras (formato, parseo, validación)
```

No se crearon `hooks/`, `pages/` ni `components/` porque no hay un framework de componentes: cada "pantalla" es una función de `render.js` que escribe HTML en un `<div>` fijo del shell. Si en el futuro el panel crece lo suficiente como para justificar React u otro framework, esta carpeta `app/` es el punto natural de reemplazo — `services/` y `config/` no deberían necesitar cambios.

## 3. Capa de integración con Excel

Toda llamada a `Excel.run` vive en `src/services/excelService.js`. La UI (`app/render.js`, `app/main.js`) nunca llama a Office.js directamente: pide una lectura, una escritura o una navegación, y esa capa es la única que sabe que existe `Excel.run`, `ctx.sync()`, `Range.load()`, etc.

Motivo concreto, no solo estilístico: la lógica de negocio de más riesgo — "nunca escribir sobre una celda con fórmula" — es una función pura (`hasFormula` en `utils/format.js`) que se puede probar con Node normal, sin un Excel real. Si esa comprobación hubiera quedado inline dentro de la llamada a `Excel.run`, no habría forma de probarla sin abrir Excel.

## 4. Manejo de errores

Cada pantalla se pinta dentro de su propio `try/catch` (`pintaTodo()` en `render.js`): si una pantalla falla al pintarse (por ejemplo, porque el libro no tiene una hoja que esa pantalla espera), las demás pantallas se siguen pintando con normalidad. Esto ya existía en el mockup original y se conservó tal cual al modularizar.

Los errores de comunicación con Excel (lectura, escritura, navegación) se capturan en `app/main.js` y se muestran en el pie de página (`foot()`), nunca como una excepción no controlada.

## 5. Manejo de estado

Un objeto mutable simple (`app/state.js`), sin librería de estado. Es proporcional al tamaño real del proyecto (un panel de tareas, ~5 usuarios internos, sin necesidad de sincronizar estado entre pestañas o componentes independientes). La fuente de verdad sigue siendo el libro de Excel; el estado en memoria es una caché de la última lectura.

## 6. Mapeo de datos (business mapping)

`src/config/cellMap.js` es la única fuente de verdad de direcciones del libro (`H`, `LECTURAS`, `BLOQUES`, `ANCLAS`). Cambiar de versión de plantilla — algo que ya ocurrió una vez según el propio código (comentario "verificadas contra la v2.5") — implica editar un solo archivo, no buscar direcciones dispersas por el código de la UI.

## 7. Estrategia de testing

Se usa el ejecutor de pruebas incorporado en Node (`node:test`, disponible sin dependencias adicionales desde Node 18+) en vez de introducir Jest/Vitest: evita añadir una dependencia nueva a un proyecto que deliberadamente tiene muy pocas, y alcanza para el tamaño actual del código.

La cobertura inicial (`/test`) se limitó a lo que aporta valor real ahora mismo:

- **Funciones puras** (`format.js`): formato de números, escape de HTML, parseo de entrada del usuario, y la guarda `hasFormula`.
- **Mapeo de celdas** (`cellMap.js`): que cada entrada de `LECTURAS`/`ANCLAS`/`BLOQUES` tenga la forma esperada y que las claves no se dupliquen — para detectar un error de copia/pegado al agregar una celda nueva.
- **Capa de Excel** (`excelService.js`): usando un doble de prueba (fake) mínimo de `Excel.run`, sin Office.js real, para verificar que la protección contra escribir sobre fórmulas efectivamente bloquea la escritura, y que la lectura por lotes respeta qué hojas existen.

Lo que **no** se probó con este runner, porque requiere Excel real y es más rápido de verificar a mano durante el desarrollo: la navegación entre hojas (`irA`), la búsqueda de fila por etiqueta, y el diagnóstico de conexión — ver punto 8.

## 8. Prueba mínima de integración Add-in → Office.js → Excel

Se agregó una pestaña **Diagnóstico** al panel (`ejecutarDiagnostico` en `excelService.js`) que:

1. cuenta las hojas del libro (prueba de lectura);
2. crea una hoja propia y oculta, `GESPER_DIAG_TEMP`, ajena al modelo;
3. escribe un valor con marca de tiempo en `A1` y lo vuelve a leer para confirmar la ida y vuelta;
4. borra esa hoja al terminar.

Es deliberadamente independiente del resto del panel (no lee ni escribe ninguna hoja del modelo GESPER) para poder usarse como prueba de humo del entorno de desarrollo sin arriesgar datos reales, tal como pide la preparación del repositorio. El botón está en la propia UI ("Ejecutar prueba de conexión"), no es un script aparte, porque la forma más simple de probar "el Add-in puede hablar con Excel" es hacerlo desde dentro del propio Add-in.

## 9. ES Modules nativos también en `server.js`

Al introducir `import`/`export` en el código del panel, se unificó todo el proyecto (incluido `server.js`, antes en CommonJS) a `"type": "module"` en `package.json`. Alternativa descartada: dejar `server.js` en CommonJS y solo modularizar el panel — se prefirió un único sistema de módulos en todo el repositorio para no obligar a quien lo mantenga a recordar cuál archivo usa `require` y cuál `import`.

## 10. Pendiente de decisión del usuario, no resuelto aquí

- El repositorio tiene commiteado un caso real (`CASO_Bucaramanga_plantilla_v2.5.xlsm`, ~2.4 MB) en la raíz. No se tocó ni se movió — no es parte de esta preparación de infraestructura — pero conviene que el usuario decida si un caso de un cliente real debe vivir en el control de versiones o si debería tratarse como dato de prueba/ejemplo y excluirse.

## 11. Carga de anexos: lectura del archivo sin librería

La pantalla **Anexos** lee un archivo que el usuario elige (hoy los anexos 1, 2 y 3, hoja `SALIDA_GESPER`, contrato v1.0 en `Especificacion_SALIDA_GESPER_v1` y `Especificacion_SALIDA_GESPER_A02_v1`, y v1.1 en `Especificacion_SALIDA_GESPER_A03_v1`) y escribe sus valores en GESPER. Office.js solo da acceso al libro abierto: `Excel.createWorkbook` abre otra ventana fuera del alcance del panel, y `insertWorksheetsFromBase64` metería hojas del anexo dentro de GESPER (y sus fórmulas, que apuntan a otras hojas del anexo, se romperían). Por eso el archivo se lee en el navegador.

Alternativa descartada: SheetJS u otra librería. El proyecto no tiene ninguna dependencia de ejecución y no tiene bundler (§1); habría que versionar un archivo de ~1 MB o cargarlo de un CDN. Lo que se necesita es poco: los valores guardados de una hoja y los nombres definidos. `src/utils/xlsx.js` (~250 líneas) lee el ZIP con el `DecompressionStream` nativo (WebView2, Safari 16.4+, Node 18+) y el XML con expresiones acotadas al formato que escribe Excel. Límites aceptados: no evalúa fórmulas (lee el último resultado guardado, por eso el panel pide "Actualizar todo" y guardar el anexo antes), no soporta `.xls` binario, archivos con contraseña ni ZIP64, y no funciona en Excel 2019 perpetuo con EdgeHTML (no tiene `DecompressionStream`); en esos casos el panel muestra un mensaje claro.

Capas, igual que la lectura del modelo: `utils/xlsx.js` (archivo → celdas) → `domain/anexos/salidaGesper.js` (contrato: estructura, estados, validación de destinos; puro y probado) → `domain/repositories/anexosRepository.js` (compone con `excelService.leerCeldas` para validar y `escribirCampos` para escribir: un lote, un recálculo, guarda anti-fórmula). Lo propio de cada anexo (hoja destino, filas de banderas) vive en `ANEXOS` de `config/cellMap.js`; agregar el A02 debería ser agregar una entrada ahí.

**Anexo 2 (2026-11-09).** Entró con una entrada nueva en `ANEXOS` y tres opciones genéricas, sin código propio del A02: `alcanceControles` (en el A01 un control en REVISAR invalida su concepto; en el A02, todo su grupo), `celdasProtegidas` (`D447`/`D478`, horizontes que nunca se cargan) y `filaAniosCalendario` (solo para mostrar el año en la vista previa). También se agregaron dos reglas para todos los anexos:

- **Confirmar los valores que cambian.** Si un valor nuevo difiere del que hoy tiene GESPER, la carga exige confirmación, aunque el anexo no traiga advertencias. Motivo: un 0 del anexo borra en silencio lo que había (caso `L465`: 10 → 0, hallazgo H-03 del A02, pendiente de validar). No es una excepción por celda.
- **Confirmación atada a lo que se vio.** `ejecutarCarga` recibe el plan confirmado y lo compara con una nueva validación justo antes de escribir. Si GESPER cambió entre la vista previa y el clic, no escribe nada y devuelve la vista previa actualizada. Así la regla "sin confirmación no se escribe" queda en el dominio y tiene pruebas, no solo en la interfaz.

Con el A01 de referencia no cambia nada visible: ya pedía confirmación (1 advertencia) y no tiene diferencias con GESPER.

**Anexo 3 (2026-11-11).** P3, P5 y P6 en un solo archivo de 79 MB (535 MB sin comprimir), contrato 1.1 (`Especificacion_SALIDA_GESPER_A03_v1`). Entró con su entrada en `ANEXOS` y estos cambios, todos genéricos:

- **REVISAR parcial (contrato 1.1).** Con `SPEC_VERSION` 1.1 o mayor, un `ESTADO_GLOBAL` en REVISAR ya no bloquea todo: se omiten las filas en REVISAR y lo que invalidan sus controles, y se avisa qué grupos quedan incompletos (`GRUPOS_REVISAR` más los que se ven en las filas). Con 1.0, o si no se puede saber qué grupos están en REVISAR, sigue bloqueando todo, como en el A01 y el A02.
- **Carga por grupo (`cargaPorGrupo`, solo A03).** Los diez grupos (segundo segmento del ID) son independientes, y varios traen diferencias pendientes de validar (H-03, H-04, H-05, H-07). Una sola casilla obligaría a aceptarlas todas o a no cargar nada. Por eso cada grupo tiene su casilla. Vienen marcados los grupos sin observaciones y sin marcar los que traen cambios, advertencias, controles con observaciones o datos que no se cargan; marcar uno es confirmarlo. La confirmación sigue atada a la vista previa: si GESPER cambió, no se escribe nada. El A01 y el A02 conservan su casilla única ("si se cancela, no se escribe nada").
- **Controles por serie.** `alcanceControles: "concepto"` ahora reconoce también los controles de cobertura (`COBERT_`), además de los `CTRL_`. En el A03, `DAN.CTRL_TAD` en REVISAR invalida la serie `DAN.TAD.*`, pero no `DAN.CAPAC.*`.
- **Referencias.** Un CONTROL con destino en GESPER (`A03.ROT.CAJAS_FALT` → `D548`, `A02.AIS.VERIF` → `L479:Q479`) solo se lee: se muestra lo que calcula GESPER junto a lo que trae el anexo, y se vuelve a leer después de la carga. Nunca se escribe.
- **Decimales.** El ruido de coma flotante (92653,6 contra 92653,59999999999) no cuenta como cambio. Una diferencia menor que 0,005 se escribe sin pedir confirmación y se muestra como "ajuste de decimales": GESPER guarda varias celdas con 2 decimales y el anexo no redondea. Con ese umbral, la vista previa da las mismas 6 diferencias que la spec A03 §6; por ejemplo, la longitud de redes (0,0056 m, H-07) sí cuenta como cambio.
- **Bandera de horizonte en medio del bloque.** En P3, P5 y P6 los volúmenes recuperados van debajo de la fila «¿contemplado?». Por eso los bloques de `banderasHorizonte` del A03 abarcan el subprograma completo.
- **Cadenas compartidas bajo demanda.** `xlsx.js` solo lee `sharedStrings.xml` si la hoja pedida lo usa. `SALIDA_GESPER` no lo usa, así que el A03 se lee en ~60 ms en lugar de ~570 ms.

**Cómo llega el archivo al panel (2026-11-10).** Desde la actualización automática de WebView2 Runtime a la 157.0.4341.42 (2026-11-08), Excel (Office 2021, build 16.0.14334) dejó de abrir el explorador de archivos desde el panel. Pasa tanto con `input.click()` por código como con un clic directo sobre el `<input type="file">`, aunque el mismo código sí lo abre en Edge. Arrastrar el archivo desde el Explorador de Windows sí funciona. Por eso la vía principal es una zona de arrastre (eventos `dragover`/`drop` en `main.js`; `preventDefault` en todo el documento para que el panel no navegue al archivo soltado), y el selector queda como alternativa secundaria, transparente sobre su botón, por si en otro equipo o versión vuelve a funcionar.
