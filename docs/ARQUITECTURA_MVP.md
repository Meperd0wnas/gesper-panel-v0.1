# Arquitectura del MVP — Panel GESPER

Diseño de arquitectura, no implementación. Ningún archivo de `src/` cambia con este documento; es la base para planear el trabajo de implementación por venir.

## 0. Alcance P0 propuesto

No existía un roadmap UX formal en el repositorio. Este es un alcance P0 propuesto a partir del dominio (docs/01 y docs/04) y del mockup ya construido, a validar antes de implementar.

| # | Funcionalidad | Por qué es P0 | Estado hoy |
|---|---|---|---|
| P0.1 | **Estado del modelo** — alertas de salud al abrir (integridad, WACC, alternativa vigente, programas sin datos, IPC desactualizado) | Es la única pantalla que responde "¿puedo confiar en lo que este libro dice hoy?" — y el propio doc de arquitectura marca el desfase motor↔consolidado como *el* modo de falla silenciosa más importante del sistema (docs/04 §5). Sin esto, el usuario no tiene forma de saberlo desde la UI. | Ya construida en el mockup, sobre datos reales |
| P0.2 | **Datos de programas — formularios generalizados** | Es el mayor dolor de UX documentado: la entrada humana más extensa del modelo vive repartida en 40+ tramos no contiguos a lo largo de 859+458 filas (docs/04 §"Procesamiento técnico y comercial"). Hoy solo P.9 tiene formulario; generalizarlo es el único flujo de escritura ya validado de punta a punta (leer → validar → escribir con guarda anti-fórmula → recalcular). | Solo P.9 implementado, hardcodeado |
| P0.3 | **Resultados** — indicadores hídricos, comerciales y financieros, de solo lectura | Es lo que el usuario compara para decidir. Bajo riesgo (solo lectura) y alto valor. | Ya construida, sobre datos reales |
| P0.4 | **Alternativas** — comparación de las 3 alternativas, de solo lectura | Es el resultado final del modelo (docs/01 §7). Bajo riesgo, alto valor. | Ya construida, sobre datos reales |

**Fuera de P0** (candidatas a P1, con su razón):
- **Navegar** (saltar a bloques de hojas): Excel ya resuelve esto nativamente razonablemente bien; es la pantalla de menor valor agregado del mockup actual.
- **Ejecutar el motor de optimización desde el Add-in**: el motor sigue viviendo en VBA por diseño de esta fase ("no reescribir el VBA"); el Add-in no lo dispara ni lo reemplaza, solo lee su resultado. Automatizar su ejecución desde Office.js (`Application.Run`) es una decisión de mayor alcance que amerita su propia discusión de producto y de riesgo.
- **Editar Parámetros globales**: se diligencian una sola vez por proyecto y ya son cómodos de editar a mano; el riesgo de un formulario mal validado ahí es alto (30 hojas dependen de esa hoja) para un beneficio de UX bajo.
- **Diagnóstico de conexión**: es infraestructura de desarrollo, no una funcionalidad de negocio — se queda, pero no cuenta como ítem de roadmap.

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

┌─────────────────────────────────────────────────────────┐
│  REPOSITORIO DE DOMINIO (src/domain/gesperRepository.js)  │
│  cargarModelo() · guardarPrograma(n, valores)             │
│  Único punto que combina lectura cruda + mapeo + escritura │
└───────────┬─────────────────────────────────┬───────────┘
            │ usa                              │ usa
┌───────────▼───────────────┐    ┌─────────────▼─────────────┐
│  MAPPERS (src/domain/      │    │  ESQUEMAS DE FORMULARIO    │
│  mappers/*.js)             │    │  (src/config/               │
│  raw {v,t} → objeto de     │    │  programFormSchemas.js)     │
│  dominio tipado. Puras,    │    │  Qué campos tiene cada       │
│  testeables sin Excel.     │    │  programa y a qué celda va    │
└───────────┬───────────────┘    └─────────────┬─────────────┘
            │ usa                              │ usa
┌───────────▼──────────────────────────────────▼─────────────┐
│  src/services/excelService.js  (sin cambios)                │
│  leerLibro · escribirRango · irA · ejecutarDiagnostico       │
└───────────────────────┬───────────────────────────────────┘
                         │ usa
┌───────────────────────▼───────────────────────────────────┐
│  src/config/cellMap.js  (sin cambios estructurales)        │
└─────────────────────────────────────────────────────────────┘
                         │
                      Office.js → Excel
```

Regla de dependencia (la que hace valer todo el diseño): **una pantalla o un componente de UI nunca importa `excelService.js` ni `cellMap.js` directamente.** Solo el repositorio y los mappers conocen la forma cruda de Excel. Es la misma regla que ya aplicamos en la preparación de infraestructura (UI → servicio Excel, nunca UI → Office.js), extendida un nivel más: ahora es UI → repositorio de dominio → (mappers + servicio Excel).

## 3. Componentes nuevos, uno por uno

### 3.1 `src/domain/mappers/*.js` — la pieza central

Un archivo por área, cada uno con una única función pura `mapX(raw) -> ObjetoDeDominio`, donde `raw` es exactamente lo que hoy devuelve `excelService.leerLibro()` (`state.D` actual). Ejemplo de responsabilidad (no de código):

- `estadoMapper.js` — de `raw.par`, `raw.parE`, `raw.ct005`, `raw.flagsV`, `raw.flags2`, etc. arma un `EstadoModelo` (contrato en §4). Aquí vive hoy la lógica de "¿el IPC está por debajo de 4.5%?" o "¿coinciden las banderas de vista y consolidado?" — ya existe en `render.js`, solo se traslada y se separa de la generación de HTML.
- `programasMapper.js` — de `raw.prog` arma `ProgramaResumen[]`.
- `resultadosMapper.js` — de `raw.ipuf2`, `raw.ianc2`, `raw.fcaja`, `raw.tirm`, `raw.par`, `raw.grad` arma `ResultadosHidraulicos` y `ResultadosFinancieros`.
- `alternativasMapper.js` — de `raw.a1bc/a2bc/a3bc`, `raw.ipuf1/ipuf2`, `raw.ver1/2/3` arma `Alternativa[]`.

Por qué esto reduce el acoplamiento y no solo lo mueve de lugar: porque **es lo único que sabe leer `v`/`t`/posición**. Un cambio de plantilla que reordene `PARÁMETROS!C4:C18` se corrige en un mapper (una función pura, con una prueba que falla si algo no cuadra), no rastreando qué pantallas usan `D.par.v[4][0]` a mano.

### 3.2 `src/config/programFormSchemas.js` — generalizar P0.2

Hoy `pintaFormP9`/`guardarP9` están escritos a mano para un único programa. Para los ~10-15 tramos de captura restantes sin reescribir el mismo patrón cada vez, se necesita un **esquema declarativo por programa**: qué campos tiene, con qué etiqueta/unidad/ayuda, a qué celda mapea cada uno, y qué series anuales tiene. Con eso, una única función genérica `renderFormularioPrograma(schema, valores)` (en `src/ui/`) y una única `guardarFormulario(schema, valoresDelForm)` (en el repositorio) cubren cualquier programa nuevo con solo agregar su entrada al esquema — igual que `cellMap.js` es hoy el único lugar para agregar una celda nueva.

Este archivo vive en `config/` junto a `cellMap.js` porque es la misma clase de cosa: mapeo de negocio, no lógica.

### 3.3 `src/domain/gesperRepository.js` — el único punto de orquestación

Reemplaza las llamadas dispersas a `excelService` que hoy hace `main.js`. Expone una API de dominio, no de Excel:

- `cargarModelo(): Promise<ModeloGesper>` — internamente llama `excelService.leerLibro(LECTURAS)` una vez y pasa el resultado por los 4 mappers. Reemplaza al actual `refrescar()` en la parte de datos (la parte de UI de `refrescar()` — spinner, pie de página — se queda en `main.js`, que sigue siendo el orquestador de eventos).
- `guardarPrograma(numero, valores): Promise<ResultadoEscritura>` — busca el esquema del programa en `programFormSchemas.js`, valida (reusa `parseNum`/reglas del esquema), escribe cada campo con `excelService.escribirRango` (que conserva su guarda anti-fórmula intacta), y no vuelve a leer el libro él mismo — deja que quien llama decida si refrescar.

Es una capa delgada a propósito: no agrega lógica de negocio propia, solo conecta mappers + esquemas + `excelService`. Sin esto, cada pantalla tendría que saber invocar 3-4 mappers por separado.

### 3.4 `src/ui/*.js` — extraer los componentes que ya existen implícitos

`render.js` ya tiene funciones puras de presentación mezcladas con las pantallas: `sig()`, `card()`, `filaPrograma()`, `fld()`, `filaAnios()`. Se separan a `src/ui/` como "componentes" reales (funciones `dato → HTML`, sin acceso a `state` ni a Excel): `AlertRow` (ex-`sig`), `KpiCard` (ex-`card`), `ProgramRow` (ex-`filaPrograma`), `FormField` (ex-`fld`), `YearSeriesGrid` (ex-`filaAnios`), más uno nuevo, `ComparisonTable`, extraído de la tabla de `pintaAlt()`. Cada uno queda testeable con datos de ejemplo, sin DOM real (devuelven strings de HTML, como hoy).

### 3.5 `src/app/screens/*.js` — reemplazan las funciones "pinta*" de `render.js`

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
 * @typedef {Object} CampoFormulario
 * @property {string} id
 * @property {string} etiqueta
 * @property {string} unidad
 * @property {string} ayuda
 * @property {{hoja: string, direccion: string}} celda
 * @property {{tipo: 'entero'|'decimal', min?: number, max?: number}} validacion
 */

/**
 * @typedef {Object} SerieAnualFormulario
 * @property {string} id
 * @property {string} etiqueta
 * @property {{hoja: string, direccion: string}} celdas   // rango horizontal, una col. por año
 * @property {{hoja: string, direccion: string}} celdasAnios // rango con las etiquetas de año
 */

/**
 * @typedef {Object} EsquemaFormularioPrograma
 * @property {number} numero
 * @property {string} titulo
 * @property {CampoFormulario[]} campos
 * @property {SerieAnualFormulario[]} series
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

Contratos de función del repositorio (la API que sí ven las pantallas):

```
cargarModelo(): Promise<ModeloGesper>
guardarPrograma(numero: number, valores: Record<string, number>): Promise<ResultadoEscritura>
listarEsquemasDisponibles(): EsquemaFormularioPrograma[]   // para que "Datos" sepa qué programas tienen formulario
```

## 5. Qué NO cambia (y por qué es una decisión, no un olvido)

- **`excelService.js` se queda intacto.** Ya cumple exactamente el rol que le corresponde (única capa que llama a `Excel.run`, guarda anti-fórmula) y no tiene lógica de dominio que extraer.
- **`cellMap.js` se queda intacto** en su rol de mapeo de direcciones; se le suma `programFormSchemas.js` como un mapeo hermano (campos de formulario), no un reemplazo.
- **No se introduce un framework ni gestor de estado.** `state.js` sigue siendo un objeto mutable simple; solo cambia qué guarda (`ModeloGesper` en vez de `raw`). Para 4 pantallas y ~5 usuarios, un mapper puro + un objeto de estado alcanza; Redux/Zustand/etc. serían sobre-ingeniería (ver docs/DECISIONES.md §5).
- **El motor de optimización sigue en VBA**, disparado desde Excel, no desde el Add-in. Ningún componente nuevo lo reemplaza ni lo re-implementa.

## 6. Estrategia de pruebas para este diseño

Sigue el mismo criterio que ya se aplicó en `/test` (docs/DECISIONES.md §7): probar lo que es barato y valioso probar.

- **Mappers** (`src/domain/mappers/*.js`): la mayor superficie de pruebas nuevas. Son funciones puras `raw -> objeto`; se prueban con fixtures de `raw` construidos a mano (sin Excel), verificando casos límite ya conocidos del dominio (WACC vacío, TIR no definida, banderas desfasadas, programa con nota en vez de B/C).
- **`programFormSchemas.js`**: mismo tipo de prueba de integridad que ya existe para `cellMap.js` (`test/cellMap.test.js`) — que cada campo tenga celda y validación, que no haya ids duplicados dentro de un esquema.
- **`gesperRepository.js`**: se prueba igual que `excelService.test.js` hoy — con el mismo doble de prueba de `Excel.run`, verificando que `cargarModelo()` compone bien los 4 mappers y que `guardarPrograma()` rechaza un valor fuera de la validación del esquema antes de intentar escribir.
- **Componentes de `src/ui/`**: pruebas de humo simples (dato de ejemplo → contiene el texto/clase esperada), igual de baratas que las de `format.js` hoy.
- **Pantallas (`src/app/screens/*.js`)**: no se prueban unitariamente (requieren DOM real); se verifican manualmente dentro de Excel, como hoy.

## 7. Orden de implementación sugerido (para cuando se empiece a programar)

1. Contratos + mappers de **Estado** (ya tiene toda la lógica escrita en `render.js`, es mover y separar, no crear de cero) — valida el patrón con el menor riesgo.
2. `gesperRepository.cargarModelo()` conectando los 4 mappers, y migrar `main.js`/pantallas existentes a consumirlo — sin agregar ninguna funcionalidad nueva todavía, solo probar que el refactor no rompe nada (paridad con el mockup actual).
3. `programFormSchemas.js` + formulario genérico, migrando P.9 al nuevo mecanismo (debe verse y comportarse igual que hoy).
4. Agregar el resto de los ~10-15 esquemas de programas restantes — este paso ya no debería tocar `src/ui/` ni `src/app/screens/datos.js`, solo el archivo de esquemas.
