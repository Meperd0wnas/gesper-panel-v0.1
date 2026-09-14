# Arquitectura del MVP — Panel GESPER

Diseño de arquitectura, no implementación. Ningún archivo de `src/` cambia con este documento; es la base para planear el trabajo de implementación por venir.

> **Revisión 2.** La versión original de este documento se apoyaba solo en `docs/01` y `docs/04`. Esta revisión incorpora una verificación directa contra `CASO_Bucaramanga_plantilla_v2.5.xlsm` (valores y fórmulas reales, no solo lo documentado) y ajusta el diseño de `programFormSchemas.js` y del repositorio de dominio en función de lo que esa verificación encontró. La arquitectura de capas y la regla de dependencia **no cambiaron** — ver §8 para el detalle de qué se verificó y qué cambió como consecuencia.

## 0. Alcance P0 propuesto

No existía un roadmap UX formal en el repositorio. Este es un alcance P0 propuesto a partir del dominio (docs/01 y docs/04) y del mockup ya construido, a validar antes de implementar.

| # | Funcionalidad | Por qué es P0 | Estado hoy |
|---|---|---|---|
| P0.1 | **Estado del modelo** — alertas de salud al abrir (integridad, WACC, alternativa vigente, programas sin datos, IPC desactualizado) | Es la única pantalla que responde "¿puedo confiar en lo que este libro dice hoy?" — y el propio doc de arquitectura marca el desfase motor↔consolidado como *el* modo de falla silenciosa más importante del sistema (docs/04 §5). Sin esto, el usuario no tiene forma de saberlo desde la UI. | Ya construida en el mockup, sobre datos reales |
| P0.2 | **Datos de programas — formularios por esquema, no una plantilla única** | Es el mayor dolor de UX documentado: la entrada humana más extensa del modelo vive repartida en 40+ tramos no contiguos a lo largo de 859+458 filas (docs/04 §"Procesamiento técnico y comercial"). Hoy solo P.9 tiene formulario; ese flujo (leer → validar → escribir con guarda anti-fórmula → recalcular) es el único ya probado de punta a punta. **Corrección de la Revisión 2**: verificado contra el archivo real, los programas restantes *no* comparten la forma de P.9 (ver §8) — "generalizar" significa un esquema con varios tipos de campo más una vía de lógica propia para los que no encajan, no una plantilla única aplicada 10-15 veces. | Solo P.9 implementado, hardcodeado. De los 18 restantes: 4 (Grupo A) encajan limpio en un esquema declarativo, 6 (Grupo B) encajan con un ajuste puntual cada uno, el resto (Grupo C, incluido el propio P.9 completo) necesita lógica específica — ver §3.2 |
| P0.3 | **Resultados** — indicadores hídricos, comerciales y financieros, de solo lectura | Es lo que el usuario compara para decidir. Bajo riesgo (solo lectura) y alto valor. | Ya construida, sobre datos reales |
| P0.4 | **Alternativas** — comparación de las 3 alternativas, de solo lectura | Es el resultado final del modelo (docs/01 §7). Bajo riesgo, alto valor. | Ya construida, sobre datos reales |

**Fuera de P0** (candidatas a P1, con su razón):
- **Navegar** (saltar a bloques de hojas): Excel ya resuelve esto nativamente razonablemente bien; es la pantalla de menor valor agregado del mockup actual.
- **Ejecutar el motor de optimización desde el Add-in**: el motor sigue viviendo en VBA por diseño de esta fase ("no reescribir el VBA"); el Add-in no lo dispara ni lo reemplaza, solo lee su resultado. Automatizar su ejecución desde Office.js (`Application.Run`) es una decisión de mayor alcance que amerita su propia discusión de producto y de riesgo.
- **Editar Parámetros globales**: se diligencian una sola vez por proyecto y ya son cómodos de editar a mano; el riesgo de un formulario mal validado ahí es alto (30 hojas dependen de esa hoja) para un beneficio de UX bajo.
- **Diagnóstico de conexión**: es infraestructura de desarrollo, no una funcionalidad de negocio — se queda, pero no cuenta como ítem de roadmap.
- **Reemplazar el motor de optimización o cualquier otro cálculo VBA**: fuera de alcance de todo este repositorio, no solo del MVP (ver CLAUDE.md). Ningún componente de esta arquitectura ejecuta ni reimplementa lógica del motor.
- **Un motor de reglas genérico que infiera el "tipo de campo" a partir del formato de la celda**: se descarta explícitamente en la Revisión 2 (§8) — los tipos de campo se declaran a mano, verificados contra el Excel real, nunca se adivinan.
- **Generalizar los 8 programas técnicos (P.1-P.8)** antes de inspeccionarlos uno por uno con el mismo método que se usó para los comerciales (§8) — docs/04 ya advierte que cada uno tiene su propia estructura, y P.1 (el único inspeccionado) confirma una heterogeneidad mayor que la de los comerciales.

Todo lo que sigue asume este alcance. Si el alcance real difiere, la arquitectura no cambia demasiado — el punto central (aislar el acoplamiento a Excel en una capa angosta) es válido para cualquier P0 razonable sobre este modelo.

## 1. El problema de acoplamiento a resolver

Hoy `render.js` construye las 5 pantallas leyendo directamente la forma cruda que devuelve `excelService.leerLibro()`: objetos `{v, t}` (values/text de `Range`) indexados por la clave de `LECTURAS`, y dentro de ellos por **posición** — `D.par.v[4][0]` es el WACC porque es la 5.ª celda del rango `C4:C18`, no porque el código diga `wacc`. Eso tiene tres costos concretos:

1. **Fragilidad silenciosa**: si cambia el orden de filas de `PARÁMETROS!C4:C18` en una versión nueva de la plantilla, `D.par.v[4][0]` sigue "funcionando" (no tira error) pero devuelve el número equivocado. cellMap.js ya es la fuente de verdad de *direcciones*; falta una fuente de verdad de *qué significa cada posición dentro de un rango*.
2. **No es testeable como lógica de negocio**: `pintaEstado()` mezcla "¿qué significa este dato?" (dominio) con "¿cómo se ve en HTML?" (presentación) con "¿de qué celda salió?" (Excel). Las pruebas de `format.js` cubren funciones puras aisladas, pero ninguna prueba hoy cubre "si el WACC viene vacío, el panel debe mostrar la alerta roja correcta" — porque esa regla vive enredada dentro de una función que también genera HTML y depende de `state.D`.
3. **Cada P0 nuevo repite el problema**: generalizar P.9 a los demás programas (P0.2) sin resolver esto significa escribir 10-15 variantes de `pintaFormP9`/`guardarP9`, cada una con sus propios índices mágicos.

La arquitectura de abajo introduce **una sola capa nueva** para resolver los tres puntos, sin tocar `excelService.js` (que ya está bien aislado y no necesita cambiar) ni introducir un framework.

## 2. Capas propuestas

```
┌─────────────────────────────────────────────────────────┐
│  PANTALLAS (src/app/screens/*.js)                        │
│  Estado · Datos · Resultados · Alternativas               │
│  Reciben solo contratos de dominio. No conocen Excel.     │
└───────────────────────┬───────────────────────────────────┘
                         │ usa
┌───────────────────────▼───────────────────────────────────┐
│  COMPONENTES DE UI (src/ui/*.js)                          │
│  KpiCard, AlertRow, ProgramRow, ComparisonTable,           │
│  FormField, YearSeriesGrid — funciones puras: dato → HTML  │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────┐  ┌───────────────────────────────┐
│  modeloRepository.js         │  │  programasRepository.js         │
│  (src/domain/)                │  │  (src/domain/)                   │
│  cargarModelo()               │  │  guardarPrograma(n, valores)      │
│  Lectura + mapeo, delgado,    │  │  Escritura: esquema declarativo   │
│  sin excepciones por programa │  │  (Grupo A/B) o función propia     │
│                                │  │  (Grupo C). Ver §3.2 y §8.        │
└───────────┬───────────────────┘  └─────────────┬────────────────────┘
            │ usa                                 │ usa
┌───────────▼───────────────┐    ┌─────────────────▼─────────────┐
│  MAPPERS (src/domain/      │    │  ESQUEMAS DE FORMULARIO         │
│  mappers/*.js)             │    │  (src/config/                    │
│  raw {v,t} → objeto de     │    │  programFormSchemas.js)          │
│  dominio tipado. Puras,    │    │  Por tipo de campo: escalar /     │
│  testeables sin Excel.     │    │  serieAnualCalendario / relativa  │
└───────────┬───────────────┘    └─────────────┬─────────────────┘
            │ usa                              │ usa
┌───────────▼──────────────────────────────────▼─────────────┐
│  src/services/excelService.js                                │
│  leerLibro · escribirRango · escribirCampos(lote) ·           │
│  irA · ejecutarDiagnostico                                     │
└───────────────────────┬───────────────────────────────────┘
                         │ usa
┌───────────────────────▼───────────────────────────────────┐
│  src/config/cellMap.js  (sin cambios estructurales)        │
└─────────────────────────────────────────────────────────────┘
                         │
                      Office.js → Excel
```

Regla de dependencia (la que hace valer todo el diseño, sin cambios en la Revisión 2): **una pantalla o un componente de UI nunca importa `excelService.js` ni `cellMap.js` directamente.** Solo los repositorios (`modeloRepository.js`, `programasRepository.js`) y los mappers conocen la forma cruda de Excel. Es la misma regla que ya aplicamos en la preparación de infraestructura (UI → servicio Excel, nunca UI → Office.js), extendida un nivel más: ahora es UI → repositorio de dominio → (mappers + esquemas + servicio Excel).

## 3. Componentes nuevos, uno por uno

### 3.1 `src/domain/mappers/*.js` — la pieza central

Un archivo por área, cada uno con una única función pura `mapX(raw) -> ObjetoDeDominio`, donde `raw` es exactamente lo que hoy devuelve `excelService.leerLibro()` (`state.D` actual). Ejemplo de responsabilidad (no de código):

- `estadoMapper.js` — de `raw.par`, `raw.parE`, `raw.ct005`, `raw.flagsV`, `raw.flags2`, etc. arma un `EstadoModelo` (contrato en §4). Aquí vive hoy la lógica de "¿el IPC está por debajo de 4.5%?" o "¿coinciden las banderas de vista y consolidado?" — ya existe en `render.js`, solo se traslada y se separa de la generación de HTML.
- `programasMapper.js` — de `raw.prog` arma `ProgramaResumen[]`.
- `resultadosMapper.js` — de `raw.ipuf2`, `raw.ianc2`, `raw.fcaja`, `raw.tirm`, `raw.par`, `raw.grad` arma `ResultadosHidraulicos` y `ResultadosFinancieros`.
- `alternativasMapper.js` — de `raw.a1bc/a2bc/a3bc`, `raw.ipuf1/ipuf2`, `raw.ver1/2/3` arma `Alternativa[]`.

Por qué esto reduce el acoplamiento y no solo lo mueve de lugar: porque **es lo único que sabe leer `v`/`t`/posición**. Un cambio de plantilla que reordene `PARÁMETROS!C4:C18` se corrige en un mapper (una función pura, con una prueba que falla si algo no cuadra), no rastreando qué pantallas usan `D.par.v[4][0]` a mano.

**Verificado en la Revisión 2** contra `CASO_Bucaramanga_plantilla_v2.5.xlsm`: los índices posicionales que hoy usa `render.js` (`v[4][0]`=WACC=C8, `v[5][0]`=IPC=C9, `v[11..14][0]`=cargos C15-C18) son correctos contra datos reales, y la clasificación de `ProgramaResumen.estado` coincide con los 19 valores reales de `Alternativas!A6:E24` (incluido el caso de P.8, que trae una nota de texto en vez de B/C, y P.11/P.18/P.19, que traen ceros reales, no vacíos). Esto confirma que el contrato de `estadoMapper`/`programasMapper` es correcto tal como estaba diseñado — el ajuste de esta revisión es únicamente en `programFormSchemas.js` (§3.2), no en los mappers de lectura.

### 3.2 `src/config/programFormSchemas.js` — generalizar P0.2, solo donde realmente aplica

Hoy `pintaFormP9`/`guardarP9` están escritos a mano para un único programa. La versión original de este documento asumía que los ~18 programas restantes compartían la forma de P.9 (unos pocos campos escalares + un par de series de 5 años) y que un solo esquema declarativo bastaría para todos. **La Revisión 2 corrige esto**: verificado contra el archivo real (detalle completo en §8), los programas *no* comparten esa forma. La generalización sigue siendo válida, pero acotada a lo que el Excel real realmente permite generalizar — no se fuerzan abstracciones para que programas heterogéneos entren en el mismo molde.

**Tipos de campo declarados explícitamente** (en vez de uno solo):

- `escalar` — una celda con un valor y una unidad (el caso más simple; ej. "% de inspecciones").
- `serieAnualCalendario` — una serie de valores en columnas de año calendario (el patrón real de P.9: `F197:J197`, años 2021-2025).
- `serieAnualRelativa` — una serie de valores en columnas de año relativo al proyecto (0..4), el patrón que aparece en P.16/P.18/P.19 — **no es lo mismo** que `serieAnualCalendario` aunque visualmente sea "una fila de 5 celdas": la columna J no es el mismo año en un programa que en otro, y varias de esas columnas suelen ser fórmulas, no entradas.

**Vía de escape declarada, no accidental**: un programa que no se represente limpiamente con estos tipos (empieza en `logicaEspecial`, una función dedicada en vez de una entrada de esquema) es una salida de diseño válida, no un parche. `EsquemaFormularioPrograma` (§4) refleja esto con una unión de tipos, no con campos opcionales que se puedan olvidar.

**Clasificación resultante** (evidencia completa en §8; P.2-P.8 técnicos aún sin inspeccionar):

| Grupo | Programas | Tratamiento |
|---|---|---|
| **A — mismo patrón, generalizable** | P.13, P.14, P.15, P.17 | Esquema declarativo, solo campos `escalar` |
| **B — generalizable con una diferencia puntual cada uno** | P.10, P.11, P.12, P.16, P.18, P.19 | Esquema declarativo + una nota/override puntual documentada por programa (ver §8 para cuál en cada caso) |
| **C — requiere lógica propia, no esquema** | P.9 (completo, incluida la curva de deterioro pendiente — §9), P.1-P.8 (técnicos) | Función dedicada por programa, sin forzar el esquema declarativo |

Este archivo vive en `config/` junto a `cellMap.js` porque, para los Grupos A y B, sigue siendo la misma clase de cosa: mapeo de negocio, no lógica. Los programas del Grupo C no viven aquí — tienen su propio módulo (ver §3.4).

### 3.3 `src/services/excelService.js` — una función nueva: `escribirCampos` (escritura por lotes)

**Incorporado en la Revisión 2 como requisito de arquitectura, no como optimización opcional.** `guardarP9` hoy escribe cada campo con una llamada separada a `escribirRango`, y `escribirRango` dispara `application.calculate(Excel.CalculationType.full)` en cada llamada — guardar el formulario de P.9 son 5 recálculos completos de un libro con más de 105.000 fórmulas (docs/03). Generalizar el patrón actual a programas con más campos (el Grupo A/B tiene programas con 3-5 campos; el Grupo C, como P.1, tiene bloques con más de 10 celdas editables) multiplicaría ese costo en cada guardado.

`escribirCampos(entries)`:
- recibe una lista de `{hoja, direccion, valor}`;
- valida la guarda anti-fórmula (`hasFormula`) en **cada** celda antes de escribir ninguna — si una sola falla, no se escribe nada;
- ejecuta todas las escrituras dentro de un único `Excel.run`;
- llama `calculate(full)` **una sola vez**, al final, no una vez por campo.

`escribirRango` no desaparece (sigue siendo útil para una escritura suelta, como en el diagnóstico de conexión); `escribirCampos` es la que deben usar `programasRepository.guardarPrograma()` y cualquier flujo del Grupo C con más de un campo.

### 3.4 El repositorio se separa en dos: `modeloRepository.js` y `programasRepository.js`

La versión original de este documento proponía un único `gesperRepository.js` con `cargarModelo()` y `guardarPrograma()`. **La Revisión 2 lo separa en dos módulos** porque, a la luz de §3.2 y §8, `guardarPrograma()` va a acumular una excepción por cada programa del Grupo B y una función completa por cada programa del Grupo C — si comparte archivo con la lectura (que es y debe seguir siendo delgada y sin excepciones), ese archivo termina siendo el monolito que esta arquitectura busca evitar.

- **`modeloRepository.js`** — sin cambios respecto al diseño original, solo renombrado: `cargarModelo(): Promise<ModeloGesper>` llama `excelService.leerLibro(LECTURAS)` una vez y pasa el resultado por los 4 mappers. Reemplaza al actual `refrescar()` en la parte de datos (la parte de UI — spinner, pie de página — se queda en `main.js`).
- **`programasRepository.js`** — `guardarPrograma(numero, valores): Promise<ResultadoEscritura>` resuelve el programa contra `programFormSchemas.js`: si es Grupo A/B, arma la lista de `{hoja, direccion, valor}` a partir del esquema y llama `excelService.escribirCampos`; si es Grupo C, delega a la función dedicada de ese programa. En ningún caso vuelve a leer el libro — deja que quien llama decida si refrescar.

Ambos siguen siendo capas delgadas a propósito: no agregan lógica de negocio propia más allá de conectar mappers/esquemas/funciones especiales con `excelService`.

### 3.5 `src/ui/*.js` — extraer los componentes que ya existen implícitos

`render.js` ya tiene funciones puras de presentación mezcladas con las pantallas: `sig()`, `card()`, `filaPrograma()`, `fld()`, `filaAnios()`. Se separan a `src/ui/` como "componentes" reales (funciones `dato → HTML`, sin acceso a `state` ni a Excel): `AlertRow` (ex-`sig`), `KpiCard` (ex-`card`), `ProgramRow` (ex-`filaPrograma`), `FormField` (ex-`fld`), más uno nuevo, `ComparisonTable`, extraído de la tabla de `pintaAlt()`. Cada uno queda testeable con datos de ejemplo, sin DOM real (devuelven strings de HTML, como hoy).

**Ajuste de la Revisión 2**: `filaAnios()` se generaliza a `YearSeriesGrid`, pero según §3.2 recibe explícitamente qué variante está pintando (`serieAnualCalendario` o `serieAnualRelativa`) — no es la misma grilla con distinta data, porque el significado de las columnas difiere entre ambas.

### 3.6 `src/app/screens/*.js` — reemplazan las funciones "pinta*" de `render.js`

Un archivo por pantalla P0 (`estado.js`, `datos.js`, `resultados.js`, `alternativas.js`). Cada uno recibe el `ModeloGesper` ya mapeado (o la porción que le toca) y arma la pantalla combinando componentes de `src/ui/`. No conocen `cellMap`, no conocen `excelService`, no hacen `D.algo.v[i][j]`.

`render.js` como archivo único desaparece, repartido entre `src/ui/` (componentes) y `src/app/screens/` (pantallas). `state.js` se mantiene pero guarda el `ModeloGesper` ya mapeado, no el `raw` crudo.

## 4. Contratos de datos (la frontera real)

Esta es la frontera de acoplamiento: todo lo que está a la izquierda de estos contratos puede cambiar (incluida la plantilla de Excel) sin tocar una pantalla; todo lo que está a la derecha puede cambiar (una pantalla nueva, un rediseño visual) sin tocar `cellMap.js`. Se documentan como JSDoc typedefs — no se introduce TypeScript (ver docs/DECISIONES.md §1, sigue sin haber paso de compilación), pero si en algún momento se decide adoptar TS o JSDoc con `checkJs`, estos son literalmente los tipos a declarar.

```js
/**
 * @typedef {Object} EstadoModelo
 * @property {string} empresa
 * @property {{inicio: string, fin: string, cantidadAnios: number}|null} horizonte
 * @property {string} metodoProyeccionUsuarios
 * @property {number|null} metaIPUF
 * @property {'ok'|'motor_sin_ejecutar'|'revisar'} integridad
 * @property {string[]} controlesIntegridad     // texto crudo de los 3 controles, para mostrar
 * @property {{real: number|null, corriente: number|null}} wacc
 * @property {boolean} alternativaDesfasada     // vista (fila 7) != consolidado (fila 49)
 * @property {ProgramaResumen[]} programasSinDatos
 * @property {{valor: number|null, desactualizado: boolean}} ipc
 * @property {{presupuesto: string, bcMinimo: string, combinacionA1: string}} configuracionMotor
 */

/**
 * @typedef {Object} ProgramaResumen
 * @property {number} numero
 * @property {string} nombre
 * @property {'tecnico'|'comercial'} tipo
 * @property {number|null} beneficioVPN
 * @property {number|null} costoVPN
 * @property {number|null} bc
 * @property {'viable'|'no_viable'|'sin_datos'|'no_aplica'} estado
 * @property {string} nota
 */

/**
 * Campo de tipo "escalar": una celda, un valor. La forma que sí comparten
 * los programas del Grupo A y la mayoría del Grupo B (ver §3.2, §8).
 * @typedef {Object} CampoEscalar
 * @property {string} id
 * @property {'escalar'} tipo
 * @property {string} etiqueta
 * @property {string} unidad
 * @property {string} ayuda
 * @property {{hoja: string, direccion: string}} celda
 * @property {{tipo: 'entero'|'decimal', min?: number, max?: number}} validacion
 */

/**
 * Serie en columnas de AÑO CALENDARIO (patrón real de P.9: F197:J197 = 2021-2025).
 * @typedef {Object} SerieAnualCalendario
 * @property {string} id
 * @property {'serieAnualCalendario'} tipo
 * @property {string} etiqueta
 * @property {{hoja: string, direccion: string}} celdas       // rango horizontal, una col. por año
 * @property {{hoja: string, direccion: string}} celdasAnios  // rango con las etiquetas de año calendario
 */

/**
 * Serie en columnas de AÑO RELATIVO al proyecto (0..4) — patrón encontrado en
 * P.16/P.18/P.19 (§8). NO intercambiable con SerieAnualCalendario: el
 * significado de cada columna es distinto y varias suelen ser fórmulas, no
 * entradas — el mapeo debe verificarse celda por celda contra el Excel real
 * antes de declararla (ver §8, caso P.10).
 * @typedef {Object} SerieAnualRelativa
 * @property {string} id
 * @property {'serieAnualRelativa'} tipo
 * @property {string} etiqueta
 * @property {{hoja: string, direccion: string}} celdas       // columnas año 0..4, solo las que son entrada real
 * @property {number[]} aniosRelativos                         // p.ej. [0,1,2,3,4], para etiquetar sin ambigüedad
 */

/**
 * @typedef {CampoEscalar | SerieAnualCalendario | SerieAnualRelativa} CampoFormulario
 */

/**
 * Esquema declarativo — solo para programas de Grupo A/B (§3.2). Los
 * programas de Grupo C no tienen un EsquemaFormularioPrograma: tienen una
 * función propia en `programasRepository.js` (§3.4).
 * @typedef {Object} EsquemaFormularioPrograma
 * @property {number} numero
 * @property {string} titulo
 * @property {'A'|'B'} grupo
 * @property {CampoFormulario[]} campos
 * @property {string} [notaVerificacion]  // para Grupo B: qué ajuste puntual tiene y por qué (ver §8)
 */

/**
 * @typedef {Object} ResultadosHidraulicos
 * @property {number|null} ipufPromedio
 * @property {number|null} metaIPUF
 * @property {boolean|null} cumpleMeta
 * @property {{inicio: number|null, fin: number|null}} ianc
 * @property {{anio: string, valor: number|null}[]} serieIPUF
 */

/**
 * @typedef {Object} ResultadosFinancieros
 * @property {string|null} vna
 * @property {string|null} tir
 * @property {boolean} tirDefinida
 * @property {number|null} tirModificada
 * @property {string|null} payback
 * @property {{waccCorriente: string, cargoFijoAcueducto: string, cargoVariableAcueducto: string, gradiente: string|null}} parametros
 * @property {boolean} alcantarilladoEnCero
 */

/**
 * @typedef {Object} Alternativa
 * @property {1|2|3} id
 * @property {number|null} beneficio
 * @property {number|null} costo
 * @property {number|null} bc
 * @property {number|null} ipufPromedio
 * @property {boolean|null} cumpleMeta
 * @property {string} composicionTexto
 */

/**
 * @typedef {Object} ModeloGesper
 * @property {EstadoModelo} estado
 * @property {ProgramaResumen[]} programas
 * @property {ResultadosHidraulicos} resultadosHidraulicos
 * @property {ResultadosFinancieros} resultadosFinancieros
 * @property {Alternativa[]} alternativas
 * @property {string[]} hojasLibro
 * @property {Date} leidoEn
 */

/**
 * @typedef {Object} ResultadoEscritura
 * @property {boolean} ok
 * @property {string} mensaje
 */
```

Contratos de función (la API que sí ven las pantallas — repartida en dos módulos, §3.4):

```
// modeloRepository.js
cargarModelo(): Promise<ModeloGesper>

// programasRepository.js
guardarPrograma(numero: number, valores: Record<string, number>): Promise<ResultadoEscritura>
listarEsquemasDisponibles(): EsquemaFormularioPrograma[]   // solo Grupo A/B; Grupo C se lista aparte, con su propio título, sin esquema

// excelService.js — nuevo en la Revisión 2
escribirCampos(entries: Array<{hoja: string, direccion: string, valor: number}>): Promise<void>
```

## 5. Qué NO cambia (y por qué es una decisión, no un olvido)

- **`excelService.js` se queda intacto en su rol.** Gana una función (`escribirCampos`, §3.3), pero sigue siendo la única capa que llama a `Excel.run`; no incorpora lógica de dominio.
- **`cellMap.js` se queda intacto** en su rol de mapeo de direcciones; se le suma `programFormSchemas.js` como un mapeo hermano (campos de formulario), no un reemplazo.
- **No se introduce un framework ni gestor de estado.** `state.js` sigue siendo un objeto mutable simple; solo cambia qué guarda (`ModeloGesper` en vez de `raw`). Para 4 pantallas y ~5 usuarios, un mapper puro + un objeto de estado alcanza; Redux/Zustand/etc. serían sobre-ingeniería (ver docs/DECISIONES.md §5).
- **El motor de optimización sigue en VBA**, disparado desde Excel, no desde el Add-in. Ningún componente nuevo lo reemplaza ni lo re-implementa.
- **No se introduce un motor de reglas que infiera el tipo de campo** (escalar / serie calendario / serie relativa) a partir del formato de la celda. Se declara a mano, programa por programa, verificado contra el Excel real — la Revisión 2 (§8) mostró que inferirlo del rótulo o de la posición produce mapeos incorrectos (P.10, P.12/14/15).
- **No se edita `PARÁMETROS` (parámetros globales) en este MVP.** Confirmado en la Revisión 2: 41 de las 48 hojas del libro están ocultas por diseño y, específicamente en `PARÁMETROS`, el color amarillo *no* marca las celdas de entrada (docs/04) — un formulario ahí necesita más precisión de la que esta fase puede darle, para una hoja de la que dependen 30 hojas más.
- **No se generalizan los programas técnicos (P.1-P.8)** sin inspeccionarlos antes uno por uno con el método de §8. P.1 (el único revisado) ya muestra un patrón más heterogéneo que el de los comerciales (valores de texto tipo "SI"/"NO"/"I/O" mezclados con números).
- **`EV PROG CCIALES!C166`/`D166`** (ponderación de curvas de deterioro de medidores) no se incorpora silenciosamente a la migración de P.9 — queda como decisión funcional aparte (§9).

## 6. Estrategia de pruebas para este diseño

Sigue el mismo criterio que ya se aplicó en `/test` (docs/DECISIONES.md §7): probar lo que es barato y valioso probar.

- **Mappers** (`src/domain/mappers/*.js`): la mayor superficie de pruebas nuevas. Son funciones puras `raw -> objeto`; se prueban con fixtures de `raw` construidos a mano (sin Excel), verificando casos límite ya conocidos del dominio (WACC vacío, TIR no definida, banderas desfasadas, programa con nota en vez de B/C).
- **`programFormSchemas.js`**: mismo tipo de prueba de integridad que ya existe para `cellMap.js` (`test/cellMap.test.js`) — que cada campo tenga celda, tipo y validación, que no haya ids duplicados dentro de un esquema, y **que un esquema declarado como Grupo A no incluya ningún campo de tipo serie** (si lo necesita, no es Grupo A — la prueba debe forzar la reclasificación, no dejarla pasar).
- **`excelService.escribirCampos`**: nuevo caso de prueba en `excelService.test.js`, con el mismo doble de `Excel.run` ya usado hoy — verificar que si una celda del lote tiene fórmula, no se escribe ninguna del lote, y que `calculate(full)` se llama exactamente una vez por lote, no una vez por campo.
- **`modeloRepository.js`**: se prueba igual que `excelService.test.js` hoy — con el mismo doble de prueba de `Excel.run`, verificando que `cargarModelo()` compone bien los 4 mappers.
- **`programasRepository.js`**: mismo enfoque, verificando que `guardarPrograma()` rechaza un valor fuera de la validación del esquema antes de intentar escribir, y que resuelve correctamente Grupo A/B (vía esquema) vs. Grupo C (vía función dedicada).
- **Componentes de `src/ui/`**: pruebas de humo simples (dato de ejemplo → contiene el texto/clase esperada), igual de baratas que las de `format.js` hoy.
- **Pantallas (`src/app/screens/*.js`)**: no se prueban unitariamente (requieren DOM real); se verifican manualmente dentro de Excel, como hoy.

## 7. Orden de implementación sugerido (para cuando se empiece a programar)

Revisado en la Revisión 2 para reflejar que la generalización es incremental, no un solo paso.

1. **`excelService.escribirCampos`** + su prueba — resuelve el riesgo de recálculo repetido (§3.3) antes de que exista código nuevo que lo repita.
2. Contratos + mappers de **Estado** (ya tiene toda la lógica escrita en `render.js`, es mover y separar, no crear de cero) — valida el patrón con el menor riesgo.
3. `modeloRepository.cargarModelo()` conectando los 4 mappers, y migrar `main.js`/pantallas existentes a consumirlo — sin agregar ninguna funcionalidad nueva todavía, solo probar que el refactor no rompe nada (paridad con el mockup actual).
4. `programFormSchemas.js` (con los tres tipos de campo) + `programasRepository.js`, migrando **P.9** al nuevo mecanismo como caso Grupo C con función dedicada (debe verse y comportarse igual que hoy, curva de deterioro aparte — §9).
5. Implementar los **4 programas del Grupo A** (P.13, P.14, P.15, P.17) con el esquema declarativo — mayor retorno, menor riesgo.
6. Implementar los **6 programas del Grupo B** (P.10, P.11, P.12, P.16, P.18, P.19), uno por uno, cada uno con su ajuste puntual documentado (§8).
7. Inspeccionar **P.2-P.8** con el mismo método usado en §8 antes de decidir cuánto del Grupo C técnico entra al MVP o queda para después — no antes.

## 8. Hallazgos de la verificación contra el archivo real (Revisión 2)

Verificación hecha leyendo valores y fórmulas reales de `CASO_Bucaramanga_plantilla_v2.5.xlsm` (no solo la documentación) contra las anclas de `cellMap.js` y el entorno de cada programa. Estos hallazgos son la evidencia detrás de los ajustes de §3.2, §3.3 y §3.4 — se listan aquí como justificación, no como una lista de parches a aplicar sin más:

- **P.10 (Gestión de fraudes)**: la celda que por posición uno esperaría como entrada (`C280`) es en realidad una fórmula; el valor literal que parece editarse de verdad vive en `K280` (columna de año 0). Un esquema que asumiera "columna C = entrada" habría mapeado mal el campo — y es exactamente el tipo de error que `hasFormula` detecta en tiempo de ejecución pero que conviene no cometer en el diseño.
- **P.12, P.14, P.15**: filas cuyo rótulo dice "Número de..." (`B313`, `B349`, `B367`) contienen en realidad *porcentajes* (0.2, 0.05...) en la columna C, no conteos. El nombre de la fila no es garantía del tipo de dato que contiene.
- **P.16, P.18, P.19**: usan series en columnas de **año relativo al proyecto** (`J:N` = año 0..4), no de año calendario como P.9 (`F:J` = 2021-2025). Es la razón concreta de separar `serieAnualCalendario` de `serieAnualRelativa` en vez de tratarlas como una sola "serie de 5 columnas".
- **P.18** además mezcla, dentro del mismo programa, una fila de serie literal (`I420:N420`) con una sola celda de tasa suelta (`C422`) — dos formas distintas en un solo programa, motivo por el que quedó en Grupo B y no en A.
- **P.1 (técnico, ya anclado)**: mezcla celdas numéricas con texto tipo `"SI"/"NO"/"I/O"/"N/A"` y placeholders literales (`"Definir valor anual"`) en celdas que en otros años sí son numéricas. Ningún tipo de campo de §3.2 lo cubre limpiamente — confirma por qué P.1-P.8 quedan en Grupo C hasta inspeccionarlos uno por uno.
- **P.16 y P.19 no están cubiertos por `ANCLAS`** en `cellMap.js` hoy, aunque ambos tienen datos reales en el caso de Bucaramanga. Antes de prometerlos en un esquema hay que agregarles ancla o decidir explícitamente dejarlos fuera.
- **P.11 (Cobro de consumos por obras)**: su celda de entrada principal (`EV PROG CCIALES!C296`) no devuelve un valor normal en este caso real (posible error o particularidad del caso). Antes de fijar su esquema hay que abrirlo en Excel y confirmar qué es esa celda — no asumirlo desde un volcado de datos.

Ninguno de estos hallazgos exige cambiar la arquitectura de capas; todos exigen que `programFormSchemas.js` soporte varios tipos de campo y una vía de lógica propia (§3.2), en vez de un solo molde.

## 9. Pendiente separado: la ponderación de curvas de deterioro de medidores

`EV PROG CCIALES!C166` (peso de la curva general) y `D166` (peso de la curva segmentada, `=1-C166`) son, según docs/01 §4.4, un parámetro que "debe ajustarse en cada proyecto" y que hoy **no está expuesto** en ningún `LECTURAS` ni en el formulario actual de P.9. Es un hueco funcional real, anterior a esta arquitectura. Se documenta aquí para que, cuando se migre P.9 al nuevo mecanismo (§7, paso 4), quede una decisión explícita — incorporarlo a esa migración o tratarlo como un ítem de producto aparte — en vez de que se cuele calladamente en el primer PR que toque P.9.

## 10. Documentación histórica vs. comportamiento comprobado del archivo real

El propio libro trae una hoja **`RESUMEN EJECUTIVO`** (visible, sin ocultar, no contemplada por ningún mapper) que registra qué se corrigió en esta versión del modelo. Contradice en al menos un punto a `docs/01` (v2.0): dice que el motor de optimización "ya no descarta en silencio la opción de tamaño máximo", mientras que la Especificación Técnica todavía lista ese comportamiento como pendiente de corregir. También menciona una hoja `FIX PLANTILLA` con "decisiones abiertas y defectos de código pendientes" que no existe en este archivo (puede estar en la plantilla canónica u otro caso).

**Regla práctica para esta y futuras revisiones**: cuando la documentación histórica (`docs/01`, `docs/04`) y el comportamiento comprobado de un archivo real difieran, prevalece el archivo real, y se anota la discrepancia en vez de propagar una afirmación de `docs/` que ya no es cierta. Antes de cerrar el alcance final de P0/P1, alguien con conocimiento del modelo debería revisar `RESUMEN EJECUTIVO` y buscar `FIX PLANTILLA` en `docs/GESPER_CANONICO_PLANTILLA_v2.5.xlsm`.
