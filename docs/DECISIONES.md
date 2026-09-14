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
