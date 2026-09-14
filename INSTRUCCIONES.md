# Panel GESPER v0.1 — cómo levantarlo

Esto es un complemento de Excel de verdad: lee y escribe tu libro. No es una maqueta.

> **Antes de nada: trabajá sobre una copia del `.xlsm`.** El panel escribe en celdas reales cuando pulsás Guardar. Copiá `CASO_Bucaramanga_plantilla_v2.5.xlsm` y probá sobre la copia.

---

## Lo que necesitás una sola vez

**1. Node.js.** Descargalo de [nodejs.org](https://nodejs.org) — la versión **LTS**. Siguiente, siguiente, terminar.

**2. Descomprimí esta carpeta** en algún lugar fácil, por ejemplo `C:\gesper-panel`.

**3. Abrí PowerShell en esa carpeta.** En el Explorador de Windows, entrá a la carpeta, hacé clic derecho en un espacio vacío y elegí *Abrir en Terminal* (o escribí `powershell` en la barra de direcciones).

**4. Instalá las dependencias:**

```
npm install
```

**5. Instalá el certificado de seguridad:**

```
npx office-addin-dev-certs install
```

Windows te va a preguntar si confiás en un certificado. **Decí que sí.** Sin esto Excel se niega a cargar el panel, porque exige HTTPS.

---

## Cada vez que quieras usarlo

**Ventana 1 — el servidor.** En la carpeta del proyecto:

```
node server.js
```

Vas a ver `Panel GESPER — servidor local activo`. **Dejá esa ventana abierta.** Si la cerrás, el panel se queda en blanco.

**Ventana 2 — abrir Excel con el panel.** Abrí otra PowerShell en la misma carpeta:

```
npm run instalar-en-excel
```

Esto abre Excel ya con el complemento registrado. Abrí tu archivo `.xlsm` y buscá el botón **Panel GESPER** en la pestaña **Inicio** de la cinta.

Para cerrar todo: `Ctrl + C` en la ventana del servidor, y `npm run quitar-de-excel`.

---

## Si el paso 2 no funciona

A veces `office-addin-debugging` falla según la versión de Office. Método manual, que siempre funciona:

1. En el Explorador, clic derecho sobre la carpeta `gesper-panel` → **Propiedades** → pestaña **Compartir** → **Compartir** → agregate a vos mismo con permiso de lectura → **Compartir**. Anotá la ruta que te muestra, algo como `\\TU-PC\gesper-panel`.
2. Abrí Excel → **Archivo** → **Opciones** → **Centro de confianza** → **Configuración del Centro de confianza** → **Catálogos de complementos**.
3. Pegá esa ruta en *Dirección URL del catálogo*, pulsá **Agregar catálogo**, marcá **Mostrar en el menú**, aceptá.
4. Cerrá Excel y volvé a abrirlo.
5. **Insertar** → **Mis complementos** → pestaña **CARPETA COMPARTIDA** → elegí *Panel GESPER* → **Agregar**.

El servidor (`node server.js`) tiene que estar corriendo igual.

---

## Qué hace el panel hoy

| Pestaña | Qué hace |
|---|---|
| **Estado** | Lee la identidad del caso y levanta cinco alertas reales: controles de integridad, si falta el WACC, cuál alternativa gobierna el consolidado, cuántos programas están sin datos, y si el IPC del modelo quedó por debajo de la inflación real. |
| **Navegar** | Seis bloques en vez de 49 hojas. Muestra la hoja si está oculta y salta a la celda correcta. |
| **Datos** | Los 19 programas con su B/C leído del libro, con semáforo. P.9 abre un formulario que **escribe de verdad** en el libro. |
| **Resultados** | Tres vistas —hidráulica, comercial, financiera— con las cifras leídas de sus celdas. Cada tarjeta dice de qué celda sale. |
| **Alternativas** | Las tres, lado a lado, con beneficio, costo, B/C, IPUF y veredicto. |

**Lo que el panel nunca hace:** escribir sobre una celda que contenga una fórmula. Antes de cada escritura la revisa y, si encuentra una fórmula, cancela y te avisa. Probalo si querés: cambiá en el mapa de celdas una dirección por una que tenga fórmula y vas a ver el rechazo.

---

## Si algo falla

**El panel sale en blanco o dice que no carga**
El servidor no está corriendo, o el certificado no quedó instalado. Volvé a `npx office-addin-dev-certs install`.

**Dice «No pude pintar esta sección»**
El panel esperaba una hoja o un rango que este libro no tiene. El mensaje dice cuál. Las direcciones están todas juntas al principio de `src/taskpane.html`, en el bloque `LECTURAS` — se corrigen ahí, sin tocar nada más.

**Los números no coinciden con el Excel**
Pulsá **Actualizar** arriba a la derecha. El panel lee una vez al abrirse; no se entera solo de los cambios que hagas en las celdas.

**Quiero apuntarlo a otra versión de la plantilla**
Todo el mapa de celdas está en un solo bloque al inicio del `<script>` de `src/taskpane.html`, marcado como *MAPA DE CELDAS*. Es el único sitio que hay que tocar cuando cambie la plantilla.

---

## Estructura de archivos

```
gesper-panel/
  manifest.xml        ← lo que Excel lee para saber que existe el complemento
  server.js           ← servidor local HTTPS
  package.json
  src/
    taskpane.html     ← TODO el panel: estructura, estilos y lógica en un solo archivo
  assets/             ← iconos
  INSTRUCCIONES.md
```

Un solo archivo para todo el panel, sin framework ni paso de compilación. Es a propósito: para que lo pueda mantener una sola persona sin herramientas adicionales.
