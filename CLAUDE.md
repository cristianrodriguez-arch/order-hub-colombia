# Order Hub Colombia — Apps Script

## Qué es este proyecto

Portal web construido en **Google Apps Script** (no Node/React, aunque el repo tenga `package.json` solo para `@google/clasp`) para el equipo de **facturación** de ISDIN Colombia. Automatiza el ciclo de una orden de compra por cliente:

1. Leer el archivo que el cliente deposita en su carpeta de Drive (PDF, Excel, JSON según el cliente).
2. Volcar la información a la hoja `CONSOLIDADO` en Sheets.
3. Generar un CSV con el formato fijo de carga a SAP.
4. Guardar el CSV en Drive y mover + renombrar el archivo original a la carpeta de "procesados".

Cada cliente cumple condiciones distintas (formato de archivo, cómo llega la orden, cómo se identifica la bodega/destino), así que cada uno tiene su propio **motor** (`<cliente>_motor.js`) y su propia **interfaz** (`<cliente>_ui.html`), pero el pipeline de negocio (arriba) es el mismo para todos.

Todo el código vive en un único proyecto Apps Script "plano" (sin carpetas reales; `.clasp.json` con `rootDir: ""`). No hay módulos: todo comparte el mismo scope global. El único "namespacing" es un prefijo de texto por cliente (`medipiel_`, `cmx_`).

## Archivos

| Archivo | Rol |
|---|---|
| `main.js` | `CONFIG` global (única fuente de configuración por cliente), inyección de vistas (`doGet`, `include`), diccionarios compartidos (`obtenerDiccionarioDestinatarios`, `obtenerMapaBodegaGeneral_`), contador de pendientes (`contarArchivosPendientes`) |
| `index.html` | Dashboard (tarjetas de cliente) + utilidades globales (`verificarPendientes`, `abrirModulo`, `volverDashboard`, `toggleLoading`, `mostrarError`, `mostrarAlertaExito`, `descargarArchivoCSV`). Incluye el wizard unificado vía `<?!= include('wizard_ui'); ?>`. `abrirModulo` consulta el registro `CLIENTES_UI` (si el cliente está registrado y no tiene `flujoCustom` → `wizardInit(cliente)`; si no → fallback a `window.initModulo<CLIENTE>`; si tampoco existe → "Módulo no disponible") |
| `wizard_ui.html` | **Wizard unificado (Opción B, implementado).** Un solo bloque HTML (ids `wizard-*` / `wizard_*` sin prefijo de cliente) + registro `CLIENTES_UI` (títulos, textos, `colorAccent`, `fnPreAnalizar`/`fnGuardar`/`fnGuardarDestinatario`) + un solo set de funciones JS `wizard*` que invocan el backend dinámicamente vía `google.script.run[...]`. Estado en un único objeto `wizardState`, reseteado completo en cada `wizardInit`. Incluye override manual de Material SAP (`wizard_override_mat_*`, en asignación automática y en multi-lote), checkbox "Notificar a la KAM" (`wizard-check-notificar-kam` → `paquete.notificarKAM`) y cálculo de back orders en el cliente (`backordersSeleccionados`, `stock < CantTotal`). **Ya no maneja el protocolo `MISSING_BODEGA`** (ver deuda técnica 11) |
| `farmatodo_ui.html` | Solo un stub visual, sin motor detrás. **No está incluido** desde `index.html` (su `initModuloFARMATODO` no existe en runtime; el botón del dashboard cae en "Módulo no disponible") |
| `medipiel_motor.js` | Motor completo de MEDIPIEL (lee PDF vía OCR/regex, escribe Consolidado, genera CSV, mueve/renombra, envía correo al KAM) |
| `cmx_motor.js` | Motor completo de CMX (lee Excel vía heurística de celdas, valoriza con la lista de precios del cliente, escribe Consolidado, genera CSV de 34 columnas, mueve/renombra, envía correo al KAM). **Reescrito por completo en el Cambio 6** |
| `cen_motor.js` | Prototipo exploratorio, **no forma parte del pipeline real** (ver estado por cliente) |
| `test.js` | Script de diagnóstico manual de permisos de Drive (`testMoverArchivo`), no es una suite de pruebas automatizada |

## Estado real por cliente

| Cliente | Motor | UI | Escribe en Consolidado | CSV real | Mueve/renombra | En `CONFIG.CLIENTES` | En el dashboard |
|---|---|---|---|---|---|---|---|
| **MEDIPIEL** | Completo | Wizard unificado (`CLIENTES_UI.MEDIPIEL` en `wizard_ui.html`) | Sí | Sí | Sí | Sí | Sí |
| **CMX** | Completo (reescrito, ver Cambio 6) | Wizard unificado (`CLIENTES_UI.CMX` en `wizard_ui.html`) | Sí | Sí (34 cols, va a `<procesados>/CSV`, no a `FOLDER_CSV_ID`) | Sí | Sí | Sí |
| **CEN** | Solo un volcado de JSON crudo a una hoja de pruebas (`procesarArchivosCEN_Prueba`), no sigue el contrato `preAnalizar/guardarDefinitivo` | No existe | No | No | No | **No** (se sacó de `CONFIG.CLIENTES`) | No |
| **BELLA_PIEL** | No existe | No existe | — | — | — | **No** (se sacó de `CONFIG.CLIENTES`) | No |
| **FARMATODO** | No existe | Solo stub visual | — | — | — | No (nunca tuvo entrada) | Sí (el botón no crashea gracias al guard en `contarArchivosPendientes`) |

## Cambios ya aplicados (histórico de esta conversación)

### 1. `CONFIG.CLIENTES` data-driven
- Se eliminaron las entradas `CEN` y `BELLA_PIEL` de `main.js` (placeholders sin implementación real, `HEADER_BODEGA` era config muerta que no se leía en ningún lado).
- Se añadió por cliente: `NOMBRE_CLIENTE` (antes hardcodeado como `"MEDIPIEL S.A.S."` / `"CMX S.A.S."` en el código), y `SAP: { CLASE_PEDIDO, ORG_VENTAS, CANAL, SECTOR }` (antes literales `"ZDIR"`, `"Z011"`, `"10"`, `"00"` dentro de `guardarDefinitivoMEDIPIEL`).
- `CMX.SAP` se asumió **igual al de MEDIPIEL** (misma sociedad/país en SAP). El `TODO` que lo marcaba ya no está en `main.js`, pero **sigue sin confirmarse con el equipo SAP** — hay que hacerlo antes de cargar pedidos reales de CMX, sobre todo `CLASE_PEDIDO`.
- Se añadió `MIME_TYPE` (MEDIPIEL = PDF) y `FILE_QUERY` (CMX = búsqueda que cubre `.xls`/`.xlsx`) por cliente.

### 2. `contarArchivosPendientes` (main.js) ahora es data-driven
- Antes tenía un `if/else` por nombre de cliente para decidir el mimeType (y ese `if` para CMX solo cubría `.xls`, no `.xlsx` — inconsistente con lo que realmente lee `preAnalizarCMX`).
- Ahora usa el helper compartido `obtenerArchivosCliente_(folder, configCli)` (nuevo, en `main.js`), que decide `searchFiles(FILE_QUERY)` vs `getFilesByType(MIME_TYPE)` según la config del cliente.
- Ese mismo helper se usa también dentro de `preAnalizarMEDIPIEL` y `preAnalizarCMX` — antes cada motor tenía su propia lógica de "qué archivos me corresponden", ahora hay una sola fuente de verdad.
- `contarArchivosPendientes` devuelve `0` si el cliente no existe en `CONFIG` (evita que reviente si algún botón del dashboard —p.ej. Farmatodo, que nunca tuvo entrada en `CONFIG.CLIENTES`— llama a esta función).

### 3. `guardarDefinitivoCMX` implementada (antes era un stub que devolvía un CSV falso)
Replica el pipeline completo de `guardarDefinitivoMEDIPIEL`, adaptado a los datos de CMX:
- Busca el código de agente (KAM) y la ciudad de destino en las hojas compartidas `Asignacion_KAM` / `Listado clientes` / `Iniciales de agente`, ahora consultadas por `configCli.SOLICITANTE` (antes ese lookup en MEDIPIEL usaba el literal `"11026712"`).
- Genera el mismo esquema global de `Id` consecutivo (`<Agente>-<Año><Consecutivo>`), escaneando **todo** `CONSOLIDADO` (no solo las filas de CMX), para no chocar con los IDs de MEDIPIEL.
- Arma las filas del CSV con `configCli.SAP.*` / `configCli.SOLICITANTE` (nada hardcodeado).
- Escribe el resumen en `CONSOLIDADO`, guarda el CSV en `CONFIG.FOLDER_CSV_ID`, mueve el archivo original a `PROCESSED_FOLDER_ID` y lo renombra a `Pedido_<Id>.<ext>` (con el mismo fallback Drive API avanzada → `DriveApp.moveTo` que ya usaba MEDIPIEL).
- Diferencias reales vs. MEDIPIEL (no duplicación gratuita, sino adaptación al dato):
  - `cmx_obtenerExtension_`: conserva la extensión original del archivo (Excel), en vez de forzar `.pdf`.
  - `cmx_parseDateFlexible_`: la fecha de pedido de CMX puede llegar como texto `dd/mm/yyyy` **o** como ISO `yyyy-mm-dd` (cuando la celda de Excel es una fecha real), soporta ambos.
  - **No envía correo al KAM** (MEDIPIEL sí). No se pidió y no está confirmado que CMX deba notificar al mismo destinatario.
- Helpers nuevos en `cmx_motor.js`: `cmx_createRowArray_`, `cmx_getEstructuraCSVHeaders_`, `cmx_obtenerExtension_`, `cmx_parseDateFlexible_` (mismo patrón que sus equivalentes `medipiel_*`).

### 4. UI unificada (Opción B) — IMPLEMENTADA
Se materializó la decisión arquitectónica de UI en un nuevo `wizard_ui.html`; se **eliminaron** `medipiel_ui.html` y `cmx_ui.html` (~1.170 líneas duplicadas). Estructura final:

- **Un solo bloque HTML del wizard** (contenedor `ui-wizard`, clase `modulo-cliente`): ids estáticos `wizard-start-step`, `wizard-loading`/`wizard-loading-text`, `wizard-flow-container`, `wizard-seccion-{revision,fechas,auto,materiales,guardar,descarga}`, `wizard-tabla-{faltantes,auto,materiales}`, `wizard-lista-pedidos`, `wizard-resumen-texto`, `wizard-btn-descargar-csv`, `wizard-contenedor-resumen`, `wizard-detalle-resumen`; ids dinámicos `wizard_fecha_${id}`, `wizard_ean_${cod}`, `wizard_card_faltante_${cod}`, `wizard_check_ean_${cod}`, `wizard_manual_mat_${cod}`, `wizard_mat_${cod}`, `wizard_check_mat_${cod}`, `wizard-swal-destino-{id,nombre}`.
- **Registro `CLIENTES_UI`** (en `wizard_ui.html`): por cliente `{ titulo, tituloInicio, textoInicio, textoBotonInicio, textoCargando, tituloSinResultados, textoSinResultados, etiquetaCodigo, textoOmitir, textoGuardando, tituloExito, prefijoError, textoPopupBodega (con token {cruce}), etiquetaCruce, placeholderDestino*, colorAccent, fnPreAnalizar, fnGuardar, fnGuardarDestinatario }`. Backend invocado dinámicamente: `google.script.run.withSuccessHandler(...).withFailureHandler(...)[config.fn](...)`. Escape hatch: `flujoCustom: true` → `abrirModulo` delega a `window.initModulo<CLIENTE>` propio.
- **Un solo set de funciones JS**: `wizardInit`, `wizardResetFlujo`, `wizardIniciarFlujo`, `wizardProcesarRespuesta`, `wizardCheckEansFaltantes`, `wizardEliminarItemFaltante`, `wizardContinuarEan`, `wizardVerOrdenesAsociadas`, `wizardRenderEtapaPrincipal`, `wizardGuardarTodo`, `wizardGenerarResumenOC`, `wizardVerDetalleResumenOC`, `wizardMostrarPopupBodega`. Base canónica: la versión MEDIPIEL (la más completa, incluye el Resumen por OC — CMX lo gana con esta unificación, no lo tenía en su versión inline).
- **Estado en un único objeto `wizardState`** `{ cliente, config, payload, materialesPreseleccionados, productosAgrupados, ultimoPaquete }`, reseteado COMPLETO en cada `wizardInit` (corrige el bug de estado residual de la versión inline de CMX, que no reseteaba `cmxMaterialesPreseleccionados` ni `cmxProductosAgrupadosGlobal`).
- **`abrirModulo` (index.html) consulta `CLIENTES_UI` directamente** (variante elegida en vez de wrappers `initModuloMEDIPIEL`/`initModuloCMX`, que ya no existen): registro sin `flujoCustom` → `wizardInit(cliente)`; con `flujoCustom` o sin registro → fallback `window.initModulo<CLIENTE>`; si nada existe (p.ej. FARMATODO) → "Módulo no disponible" + volver al dashboard.
- **`colorAccent`** se aplica en runtime al botón de inicio, al spinner del wizard y al `confirmButtonColor` del popup MISSING_BODEGA. MEDIPIEL `#E41F33`; CMX `#9A64CD` (el morado que pretendía la UI dedicada de CMX — **pendiente de validar con el usuario**, la versión inline usaba el negro/rojo genérico).
- Textos por cliente preservados de las versiones que estaban vivas (medipiel_ui.html y la CMX inline de index.html).

### 5. Lector de Excel de CMX tolerante a múltiples plantillas — ⚠️ REEMPLAZADO POR EL CAMBIO 6

> **Ojo**: el análisis de las plantillas (tabla de abajo) sigue siendo válido y verificado, pero **la implementación descrita en esta sección ya no existe en el código**. Todos los helpers que se nombran (`cmx_elegirHojaOrden_`, `CMX_COLUMNAS`, `cmx_detectarColumnas_`, `CMX_ETIQUETAS`, `cmx_buscarEtiqueta_`, `cmx_celdaANumeroTexto_`, `cmx_esFilaCierre_`, `cmx_leerCelda_`, `cmx_obtenerEanValido_`, `cmx_getProductosMap_`) fueron borrados en el Cambio 6 y sustituidos por una heurística más corta. Se conserva la sección porque documenta *por qué* existían: cada viñeta es un fallo real que las plantillas provocan, y el Cambio 6 lo resuelve o lo re-abre según el caso (comparativa en el Cambio 6).

**Problema real**: CMX no tiene una sola plantilla. Cada comprador de Línea Estética (NIT 900.816.838-2) usa la suya. Se analizaron dos órdenes reales (`Examples CMX/`):

| | `OC_ISDIN…SC 911.xlsx` (John Arcila) | `ISDIN #644 (1).xlsm` (Jonathan Rendón) |
|---|---|---|
| Hojas | 1: `Imprimir` | 3: `BASE PROVEEDORES`+`BASE SKU` **ocultas**, luego `FORMATO OC ` (con espacio final) |
| N° OC | `D5="N° OC:"` → `E5="SC 911"` | `A6="N° O.C"` → `B6=644` |
| Fila encabezado | 15 | 13 |
| Rótulos | Codigo · Descripción · Código de Barras · Bonificado · Unidades · A Facturar | SKU · Descripcion · CODIGOS DE BARRAS · **BONICADOS** (typo) · CANTIDAD · CANTIDAD FACTURADA |
| Cola de la tabla | limpia | **128 filas de relleno** con `0` |
| Celda del EAN | **numérica** → notación científica | texto |

El `.xlsm` **no se procesaba en absoluto** (4 fallos: hoja 0 equivocada, `"N° O.C"` no matcheaba `"OC"`, `"CANTIDAD FACTURADA"` pisaba el índice de cantidad, y las filas de relleno entraban como ítems fantasma).

Estrategia elegida (híbrida): **encabezado por rótulo con respaldo posicional**. Lo que se agregó en `cmx_motor.js`:

- **`cmx_elegirHojaOrden_`**: puntúa las hojas y elige la de la orden. Ya no se asume `getSheets()[0]`. La señal decisiva es tener un encabezado con descripción **y** EAN **y** cantidad — exigir las tres es lo que descarta `BASE SKU` (tiene CODIGO/BARRAS/PRODUCTO pero ninguna columna de cantidad).
- **`CMX_COLUMNAS` / `cmx_detectarColumnas_`**: sinónimos por campo con prioridad y listas de exclusión, comparados contra el rótulo normalizado (tolera acentos y typos). Una columna solo puede ser un campo: los campos se resuelven por especificidad y reservan su columna, así `"Código de Barras"` no termina también como código de cliente ni `"CANTIDAD FACTURADA"` pisa `"CANTIDAD"`.
- **`CMX_COLUMNAS_FALLBACK`**: si no hay rótulos legibles, se asume el orden común a las plantillas conocidas (`A=cod, B=desc, C=ean, D=bonif, E=cant`). Si tampoco hay filas con esa forma, lanza un error explícito en vez de leer basura.
- **`CMX_ETIQUETAS` / `cmx_buscarEtiqueta_`**: rótulos de cabecera comparados por **igualdad** contra el texto normalizado — `"N° OC:"` y `"N° O.C"` colapsan ambos a `"NOC"`, que es justo lo que hacía falta. `cmx_esRotulo_` impide que un título arrastre como valor el rótulo de otro campo. `cmx_buscarOCEnLinea_` es el último recurso cuando rótulo y número comparten celda.
- **`cmx_celdaANumeroTexto_`**: cuando el cliente pega varios EAN sin separador, Excel lo guarda como **número** y el redondeo a ~15 dígitos significativos destruye todo lo que sigue al primer código (los 13 primeros sobreviven intactos). `String(n)` devolvía `"8.42942022626584e+25"` y arrastraba el exponente entre los dígitos; ahora se reconstruye la mantisa. Antes funcionaba por accidente.
- **Fin de tabla**: `cmx_esFilaCierre_` (TOTAL/OBSERVACIONES/AUTORIZA en cualquier columna) + `CMX_MAX_FILAS_VACIAS`, y una fila cuenta solo si identifica producto **y** pide unidades > 0. Esto corta las 128 filas de relleno sin depender de que exista un TOTAL.
- **`cmx_leerCelda_` / `cmx_leerNumero_` / `cmx_esFecha_`**: lectura de celda tipada; `cmx_esFecha_` usa duck typing en vez de `instanceof Date`.

**Resolución del EAN en dos rondas completas** (`cmx_obtenerEanValido_`, orden definido por el usuario). Cuando la celda trae varios códigos, se prueban **todos** contra una hoja antes de pasar a la siguiente — las rondas no se intercalan:

1. todos los candidatos contra **`Productos`** (catálogo maestro, fuente preferida: tiene el material de todos los productos, incluidos los sin stock);
2. si ninguno apareció, todos los candidatos contra **`Bodega`** (stock SAP), que aporta 194 EAN que no están en `Productos` (304 vs 119);
3. si tampoco, se devuelve el primero para que la UI lo pida a mano.

Antes la cascada era Bodega→Productos y se cortaba en el primer candidato. Caso real que lo motivó: la celda `"8429420248083,8429420113374"` (cod. CMX 2544, K OX EYES) — el 1.º no existe en ninguna hoja y el 2.º solo en `Bodega` (3 lotes).

Bodega **sigue además** siendo la fuente de `Materiales_Disponibles` (stock y fechas de caducidad que el wizard muestra para elegir lote); quitarla de ahí habría dejado todo en stock 999999 sin vencimientos.

**Columna `Valor`**: ninguna plantilla de CMX trae precio unitario (`"A Facturar"` / `"CANTIDAD FACTURADA"` vienen vacías y no son un precio — se las detectaba como tal por error). Ahora `cmx_getProductosMap_` devuelve `{materiales, pvp}` en una sola pasada y `Valor_Unitario` se llena con el **PVP Final** de `Productos`. Si alguna plantilla llegara a traer una columna de precio real, esa manda.

Ítems nuevos en el payload a la UI: `Cod_Cliente` (el código interno de CMX, útil cuando el EAN no resuelve) y `Bonificado`.

**Verificación**: se portaron las funciones puras a un arnés Node con los valores exactos de los dos Excel (emulando `getValues()`: números, textos, `Date`, `""`). Resultado: `SC 911` → 77 ítems / 8.946 uds / fecha 2026-05-06; `#644` → 34 ítems / 2.321 uds / fecha 2026-07-16, cero ítems fantasma en ambos. Además 11 casos de borde pasan: fallback posicional (encabezado borrado), columna insertada, columnas reordenadas, fila TOTAL, hoja sin tabla → error claro, y la resolución multi-EAN.

### 6. Lista de precios por cliente + reescritura completa del motor CMX

Cambio hecho por el usuario (`cmx_motor.js` reescrito de punta a punta + `LISTA_PRECIOS` en `main.js`), evaluado y verificado en esta sesión.

#### 6.1 `Productos` ya no es un diccionario EAN→material: es una matriz de precios

Estructura real de la hoja (verificada contra el spreadsheet): `A Product ID · B Material ID · C BU · D BRAND · E EAN · F Producto · G PVP Final · H Lista Cliente · I Costo`. Hay **una fila por (producto × lista de precios)** — ~20 listas por producto, ~2.921 filas.

- **`Costo` es heterogéneo**: para las listas de precio reales es el **precio neto** al cliente (CMX = 91.000 donde el PVP Final es 166.600); para las listas de *descuento* (`AXA Y PHARMAPLUS 20%`, `CATEGORÍA A 14.3%`…) es un **porcentaje** (24, 12). Cualquier cliente futuro que se cuelgue de esta columna tiene que usar su lista de precio, nunca una de descuento.
- La lista de cada cliente ya está en `Asignacion_KAM` (columna `LISTA DE PRECIOS`, indexada por `SAP ID`). `CONFIG.CLIENTES[x].LISTA_PRECIOS` **duplica ese dato**: MEDIPIEL `"MEDIPIEL BEAUTYCALIA"`, CMX `"CMX\nCADA PIEL\nThe Beauty club"` (la celda es literalmente multilínea, tanto en `Asignacion_KAM` como en `Productos`). El match aguanta el `\n` porque `cmx_normalizeKey_` borra todo lo que no sea alfanumérico en los dos lados. Se podría derivar de la hoja en vez de hardcodearlo.
- `MEDIPIEL.LISTA_PRECIOS` está en `CONFIG` pero **no se usa**: `medipiel_getProductosMap_` no filtra por lista y el precio de MEDIPIEL sigue saliendo del PDF (columna COSTO UN de la orden). Es config preparada, no activa.

#### 6.2 Qué hace el motor nuevo

- **`cmx_getProductosPreciosMap_()`**: filtra `Productos` por `LISTA_PRECIOS` del cliente e indexa **por EAN y por Material ID** al mismo objeto `{precio: Costo, descripcion, materialOficial, eanOficial}`. Sin colisión posible (material = 9 dígitos, EAN = 13).
- **`Valor_Unitario` = `Costo` de la lista del cliente** (antes: `PVP Final`, que sobrevaloraba todo el CONSOLIDADO). Es la mejora de fondo de este cambio. Si el EAN no resuelve se intenta por el material que devolvió Bodega.
- **`cmx_getBodegaStockMap_()`**: mapa local de stock por EAN, ya no usa `obtenerMapaBodegaGeneral_` de `main.js` (dos lecturas distintas de la misma hoja; ver deuda 12).
- **Correo al KAM para CMX** (`cmx_enviarNotificacionKam_`), valorizado en $ COP y con badges OK / Back Order, con asunto homologado al de MEDIPIEL. Se envía **por defecto**: solo se salta con `paquete.notificarKAM === false` (el checkbox del wizard).
- **CSV de 34 columnas** (`Contador … Salesforce ID`) con `Precio unitario (PR00)` = `Math.round(Valor_Unitario)`. Bloque de cabecera SAP (clase pedido, org ventas, canal, sector, solicitante, destinatario, nº pedido cliente, fecha) **solo en la primera posición de cada pedido**, en blanco en las siguientes. MEDIPIEL genera 16 columnas y repite la cabecera en cada fila → **los dos clientes ya no producen el mismo CSV; hay que confirmar con SAP/facturación cuál acepta el cargador**.
- `Item_ID` determinista (`<OC>_<n>`) en vez de `Math.random()`.
- Ciudad / Id_Destino / Destinatario salen de `Destinatarios` vía `obtenerDiccionarioDestinatarios()` (la fila de CMX tiene `Cruce_cliente = BODEGA_CMX_PRINCIPAL`, `Id_destino = 11033482`, `Ciudad = LA ESTRELLA`), con esos mismos valores repetidos como literales de fallback dentro del motor.

#### 6.3 Verificado (arnés Node con los dos Excel reales)

Se volvieron a portar las funciones puras a Node alimentándolas con el volcado exacto de las celdas de `Examples CMX/` (openpyxl → números / textos / `Date` / `""`, emulando `getValues()`):

| | `SC 911.xlsx` | `#644.xlsm` |
|---|---|---|
| Hoja elegida | `Imprimir` | `FORMATO OC ` (ignora las 2 ocultas) ✅ |
| OC | `SC 911` ✅ | `644` ✅ |
| Fecha pedido | 06/05/2026 ✅ | 16/07/2026 ✅ |
| Fila encabezado / columnas | 15 · `{cod:0, desc:1, ean:2, cant:4}` ✅ | 13 · idem ✅ |
| Ítems / unidades | **77 / 8.946** ✅ | **34 / 2.321** ✅ |
| Ítems fantasma | 0 | 0 (las 128 filas de relleno se cortan por `cantidad <= 0`) |

**Paridad de lectura con el motor anterior**: la heurística corta acierta lo mismo que la versión con `CMX_COLUMNAS`/`CMX_ETIQUETAS` en las dos plantillas reales. `"CANTIDAD FACTURADA"` sigue excluida correctamente, y `"N° O.C"` se resuelve. Lo que se perdió no es la lectura de la tabla, sino la robustez ante plantillas nuevas y la recuperación del EAN numérico (deudas 13 y 14).

#### 6.4 Regresiones confirmadas al reescribir (detalle en Deuda técnica)

Verificadas contra el spreadsheet real y/o el arnés, en orden de gravedad: antiduplicados de CMX inoperante (deuda 13), EAN numérico perdido por completo — 21 de 77 ítems de `SC 911` (deuda 14), fechas de vencimiento de los lotes siempre "Sin fecha" (deuda 15), iniciales y correo del KAM siempre en el default hardcodeado (deuda 16), un solo `Id` para todo el lote + `Num pedido cliente` = Id interno (deuda 17), archivos movidos aunque no se inyecte nada (deuda 18).

## Decisión arquitectónica: UI unificada (Opción B) — ✅ implementada

Tras comparar "un archivo `.html` por cliente con utilidades centralizadas" (Opción A) vs. "wizard unificado dinámico" (Opción B), se decidió e implementó la **Opción B** (ver "Cambio 4" arriba para la estructura final). Racional que sigue vigente para réplicas:

- El wizard de MEDIPIEL y el de CMX eran el mismo componente (funciones idénticas salvo prefijo, textos, colores y nombres de funciones backend). El contrato de datos ya estaba unificado.
- `<?!= include() ?>` se evalúa una sola vez en el servidor: los archivos separados no aportan caché ni lazy-loading, solo organización — y produjeron divergencia real (`cmx_ui.html` huérfano) y un bug de estado residual.
- Prerrequisito para entrar al wizard: todo motor debe hablar el contrato `{pedidosExtraidos, itemsExtraidos}` / `MISSING_BODEGA|...` (MEDIPIEL y CMX lo cumplen). Un cliente con flujo genuinamente distinto (p.ej. Farmatodo B2B) usa `flujoCustom: true` + su propio `initModulo<CLIENTE>`.

## Subagente del proyecto

Existe `.claude/agents/orderhub-replicator.md` (modelo Fable): especialista en mantener este Order Hub y **replicarlo en otros países** con otros clientes. Conoce los contratos de arquitectura, el playbook de "agregar cliente" y el de "replicar país". Usarlo para cualquier tarea sobre motores, UIs, `CONFIG.CLIENTES`, CSV SAP o CONSOLIDADO.

## Deuda técnica conocida (detectada en el análisis, pendiente de decidir)

1. **Lógica duplicada entre `medipiel_motor.js` y `cmx_motor.js`** (actualizado tras el Cambio 6, que cambió *qué* está duplicado): `normalizeKey_` (idéntica salvo el orden de los `replace`), la lectura de `Productos`, la lectura de `Bodega`, `guardarNuevoDestinatario`, la búsqueda del KAM en `Asignacion_KAM` + `Iniciales de agente`, y **el HTML del correo al KAM** (~90 líneas casi iguales, con los mismos estilos inline y el mismo formateo de $ COP). Candidatos claros a un `utils.js` compartido. **No se ha hecho todavía.** Ojo: hoy las dos copias de cada cosa *no son equivalentes* — la de MEDIPIEL acierta y la de CMX no (deudas 13, 15, 16), así que unificar es también la vía de arreglo.
2. ~~**UI duplicada**: `cmx_ui.html` y `medipiel_ui.html` comparten ~700 líneas casi idénticas.~~ **RESUELTO** en el Cambio 4 (wizard unificado; ambos archivos eliminados).
3. ~~**`cmx_ui.html` es código muerto** + copia inline divergente en `index.html`.~~ **RESUELTO** en el Cambio 4 (se borró `cmx_ui.html` y se eliminó todo el bloque inline de CMX de `index.html`).
4. **`cen_motor.js` quedó con una referencia rota**: `procesarArchivosCEN_Prueba` sigue leyendo `CONFIG.CLIENTES["CEN"]`, que ya no existe (se quitó de `main.js` a pedido del usuario). No está enlazada a ningún botón del dashboard, así que no afecta el flujo real, pero lanzará error si alguien la ejecuta manualmente desde el editor.
5. **`cabecera.fechaEntrega` de CMX siempre llega vacía** — y no es arreglable leyendo mejor el archivo: **ninguna de las dos plantillas reales de CMX trae fecha de entrega**, el dato no existe en el origen (habría que pedírselo al cliente). La UI aplica su default de "+7 días". El Cambio 6 además volvió a quitar la búsqueda del rótulo: `cmx_extraerDatosHeuristico_` declara `fechaEntrega` y no la asigna nunca.
6. **Farmatodo** no tiene motor ni entrada en `CONFIG.CLIENTES`; el botón del dashboard existe pero no hace nada real todavía.
7. **CSV**: solo escapa `;` y `"`; no neutraliza fórmulas (`=`, `+`, `-`, `@`) si el archivo se llegara a abrir en Excel.
8. **IDs de item no colisión-seguros en MEDIPIEL**: `Item_ID: Math.random().toString(36)...` en vez de `Utilities.getUuid()`. CMX ya no lo tiene (usa `<OC>_<n>`, determinista).
9. **La hoja `Listado clientes` no existe** en el spreadsheet (`1L5bxc9IX…`). Sus hojas reales son: CONSOLIDADO, Tabla dinámica 6, ZSD_168, CUTIS, BELLA PIEL, CIRUDERMA, Iniciales de agente, Destinatarios, Asignacion_KAM, Bodega, Productos, CSV_APP, Prueba1, PRUEBAS_CEN_RAW, prompt csv, CLIENTES CEN. `guardarDefinitivoMEDIPIEL` sigue haciendo `getSheetByName("Listado clientes")` → `null` → `diccCiudades` queda vacío, **pero la columna `Ciudad` no sale vacía** (corrección de lo que decía antes esta nota): el fallback `p.Ciudad` viene de `Destinatarios` y sí funciona (las filas reales del CONSOLIDADO tienen MEDELLÍN / BOGOTÁ / SABANETA). Es una lectura muerta, no un bug de dato. En CMX el equivalente muerto es **`cmx_getCiudadParaCliente_`**, que además no se llama desde ningún lado.
10. **Datos maestros incompletos para CMX**: `Productos` tiene ~146 productos (una fila por lista de precios) y `Bodega` 304 EAN; CMX pide ~132 SKU ISDIN. El conteo de "faltantes" de esta nota (12 de 77 y 5 de 34) se midió con el motor del Cambio 5 y **ya no aplica**: con el motor del Cambio 6 el piso son 21 de 77 en `SC 911` (ver deuda 14).
    - Prueba incontestable de que el origen está roto: el cod. 2544 (K OX EYES) **se resuelve en `#644` y no en `SC 911`** — mismo producto y mismo cliente, pero en `#644` la celda es texto y conserva los dos EAN, mientras en `SC 911` es numérica.
    - Vías para cerrarlo, ninguna aplicada: pedirle a Línea Estética que formatee "Códigos de barras" **como texto** (arregla el origen, no lo controlamos), o una tabla `CODIGO_CMX → Material` (el código de la columna A sí es limpio y estable: los 111 códigos de ambas órdenes existen en la `BASE SKU` del cliente, y los 20 que aparecen en las dos plantillas apuntan al mismo producto).
    - Efecto colateral: cuando facturación teclea a mano el material de uno de estos ítems, el `Valor_Unitario` sigue en 0 (el precio se resolvió antes, por EAN), así que **esos renglones quedan sin valor** en el CONSOLIDADO y en el CSV. El override del wizard no recalcula el precio.

### Regresiones introducidas por el Cambio 6 (verificadas, sin corregir)

11. **El protocolo `MISSING_BODEGA` quedó huérfano** (afecta a **MEDIPIEL**, no a CMX). `medipiel_parsePedidoCabecera_` sigue lanzando `MISSING_BODEGA|<bodega>|<solicitante>|<cliente>` cuando la bodega del PDF no está en `Destinatarios`, pero el `wizard_ui.html` actual **no tiene `wizardMostrarPopupBodega`**: el error cae en `mostrarError` y facturación ve el string crudo `MISSING_BODEGA|1500|11026712|…` y vuelve al dashboard, sin poder registrar el destinatario. Los `fnGuardarDestinatario` del registro `CLIENTES_UI` **nunca se invocan**. Además `cmx_guardarNuevoDestinatario` (código muerto hoy) **appendea en el orden equivocado**: escribe `[idOrigen, nombreOrigen, cruceCliente, idDestino, nombreDestino]` sobre una hoja cuyas columnas son `Id_origen · Nombre_origen · Id_destino · Nombre_destino · Cruce_cliente · Ciudad`, y no escribe `Ciudad`. La versión de MEDIPIEL sí está bien.
12. **Tres lecturas distintas de `Bodega` y dos de `Productos`** conviviendo: `obtenerMapaBodegaGeneral_` (main.js, la usa MEDIPIEL, agrupa stock por material+fecha y busca la columna de fecha por nombre con fallback posicional), `cmx_getBodegaStockMap_` (una entrada por fila, `colEan` hardcodeado en 9) y `medipiel_getProductosMap_` vs `cmx_getProductosPreciosMap_`. Esto es la deuda 1 empeorada, y es la causa raíz de la deuda 15.
13. **El antiduplicados de CMX no funciona** (verificado contra el spreadsheet). `preAnalizarCMX` lee los encabezados del CONSOLIDADO como `dataCons[0]`, pero **la fila 1 es el título `"PEDIDOS SIN FACTURAR"`; los encabezados están en la fila 3** (`frozenRowCount: 3`; MEDIPIEL lo lee bien con `medipiel_getHeaderMap_(sh, 3)`). Resultado: `colOc === -1`, `pedidosProcesados` queda vacío, **nunca se detecta una OC repetida** y `yaProcesados` siempre llega vacío a la UI (el mensaje "ya se encuentran registradas" es inalcanzable). Lo único que hoy evita el doble cargue es que el archivo se mueve a "procesados"; si el cliente vuelve a subir el mismo archivo, entra otra vez. El consecutivo `Id` sí lee bien porque escanea la columna A completa sin depender del encabezado.
14. **El EAN numérico se pierde por completo** (era el bug que el Cambio 5 había cerrado; ahora es peor que antes). `cmx_limpiarEanDeVerdad_` hace `Number(s).toFixed(0)`, pero **JS devuelve notación exponencial en `toFixed` para valores ≥ 1e21** (spec: si x ≥ 10²¹ → `ToString(x)`), así que `"8.42942022626584e+25"` → `"8.42942022626584e+25"` → `split(".")[0]` → `"8"` → descartado por `length < 7`. Medido en el arnés: **21 de los 77 ítems de `SC 911` se quedan sin ningún candidato de EAN** (antes se recuperaban los 13 primeros dígitos reconstruyendo la mantisa) → van todos a "faltantes" para teclear a mano, con `Producto_Cod = "8.42942022626584e+25"` y `Valor_Unitario = 0`. Ítems concretos afectados: FUSION WATER MAGIC LIGHT/MEDIUM, GEL CREMA SPF50+, HYALURONIC CONCENTRATE, K OX EYES, MELACLEAR ADVANCED, +15 más. Arreglo mínimo: reconstruir la mantisa (`"8.4294…e+25"` → dígitos) antes del `replace(/\D/g,"")`.
15. **Los lotes salen todos "Sin fecha"** (verificado contra los encabezados reales de `Bodega`). `cmx_getBodegaStockMap_` busca la columna de vencimiento con `h.includes("VENCIMIENTO"|"CADUCIDAD"|"FECHA"|"VENCE")`, pero el encabezado real es **`FeCaduc/FePreferCons`** → normaliza a `FECADUCFEPREFERCONS`, que no contiene ninguno de esos cuatro tokens → `colFecha = -1`. Consecuencia: el paso "Múltiples Lotes" del wizard muestra `Vence: N/A` en todas las opciones e `isExpired` es siempre `false`, o sea **facturación elige lote a ciegas**, que es justo el propósito de ese paso. `obtenerMapaBodegaGeneral_` (main.js) sí acierta porque busca el nombre exacto y cae al índice 7. Además `cmx_getBodegaStockMap_` no agrupa por lote, así que aparecen entradas duplicadas indistinguibles.
16. **Las iniciales y el correo del KAM siempre caen al default hardcodeado** (verificado contra `Iniciales de agente`, cuyas columnas son `TEAM · KAM ENCARGADO · Código · Correo · Seguimientos`). En `cmx_getKamInitialsAndName_` y `cmx_enviarNotificacionKam_` la fila se busca comparando el nombre del KAM contra **`dataAg[i][0]`, que es la columna `TEAM` (`"KAM"`/`"VM"`)** — nunca coincide. Y `cInic` se resuelve con `h.includes("KAM")`, que apunta a `KAM ENCARGADO` (el nombre) en vez de a `Código` (las iniciales). Hoy no se nota porque los defaults son `"SARA CARDONA"` / `"SC"` / `sara.cardona@isdin.com` y **el KAM de CMX (SAP ID 11033482) es efectivamente Sara Cardona**: funciona por casualidad. El día que cambie el KAM, el consecutivo seguirá siendo `SC-26xxxx` y el correo seguirá yendo a Sara. MEDIPIEL lo hace bien (busca por `CODIGO` y `CORREO` con `medipiel_normalizeKey_`). Corolario: `SARA CARDONA`, `SC`, `sara.cardona@isdin.com`, `11033482` y `LA ESTRELLA` están hardcodeados en `cmx_motor.js` violando la convención de "toda constante de negocio va en `CONFIG`".
17. **Un solo `Id` para todo el lote, y el CSV pierde la OC del cliente**. `guardarDefinitivoCMX` genera **un** `nuevoIdStr` y lo usa para todos los pedidos del paquete; MEDIPIEL genera un consecutivo **por pedido**. Si la carpeta trae dos órdenes, ambas filas del CONSOLIDADO comparten `Id`, y los dos archivos se renombran `Pedido_<mismo Id>.<ext>` (nombres duplicados en Drive). Peor: en el CSV, `Num pedido cliente` (col 9) se llena con `nuevoIdStr` en vez de con la OC del cliente (`p.Pedido_ID`) — MEDIPIEL manda la OC. **SAP dejaría de recibir el número de orden del cliente.**
18. **Se mueven archivos que no se inyectaron**. El bucle de mover/renombrar recorre `paquete.pedidos` completo, mientras que el CONSOLIDADO omite los pedidos con `totalUnidades === 0`. Un pedido sin ninguna posición válida **desaparece de la carpeta de pendientes sin quedar registrado en ninguna parte**. MEDIPIEL se protege con el set `pedidosValidos`. Agravantes: el `moveTo` no tiene el fallback a `Drive.Files.patch` que sí tiene MEDIPIEL (necesario en unidades compartidas) y los errores se tragan con `catch (ignore) {}`.
19. **Divergencias menores de CMX vs MEDIPIEL** (decidir si son intencionales): el CSV se guarda en una subcarpeta `CSV` creada dentro de `PROCESSED_FOLDER_ID` en vez de en `CONFIG.FOLDER_CSV_ID` (que queda sin uso para CMX); `Fecha Pedido` y `Fecha entrega` se escriben como **texto** (`"06/05/2026"`, `"2026-08-11"`) mientras MEDIPIEL escribe objetos `Date` — en una columna que hoy tiene números de serie; la columna `KAM` del CONSOLIDADO se deja vacía aunque el nombre ya se conoce (las filas reales la tienen llena); `errores` se devuelve al wizard pero el wizard no lo muestra (si 1 de 3 archivos falla, silencio); el `Contador` del CSV se incrementa también para pedidos descartados (deja huecos).
20. **`preAnalizarCMX` ya no usa `obtenerArchivosCliente_`** (rompe la convención del Cambio 2): filtra por extensión (`.xlsm`/`.xlsx`/`.xls`) sobre `folder.getFiles()`, mientras `contarArchivosPendientes` sigue usando `FILE_QUERY` de `CONFIG`. Son dos definiciones distintas de "pendiente": el badge del dashboard cuenta cualquier `mimeType contains 'spreadsheet'` (incluidos Google Sheets nativos) que el motor después ignora. Y como `cmx_convertirExcelASheet_` crea el Sheet temporal **dentro de la carpeta de pendientes** (`file.getParents().next()`), si el script muere antes del `setTrashed` queda un Sheet fantasma que el badge cuenta para siempre.

## Convenciones a respetar si se agregan clientes nuevos

- Un cliente "completo" implementa el par `preAnalizar<CLIENTE>()` / `guardarDefinitivo<CLIENTE>()` en su propio `<cliente>_motor.js`, y escribe en `CONSOLIDADO` usando `CONSOLIDADO_HEADERS` (definido en `main.js`, única fuente de verdad de esas columnas).
- Toda constante de negocio (SAP: clase de pedido/org ventas/canal/sector, solicitante, nombre del cliente, lista de precios, cómo se detectan sus archivos) va en `CONFIG.CLIENTES[cliente]` — **no hardcodeada en el motor**. Incluye el KAM, su correo y la ciudad de destino: si el motor los necesita, se leen de `Asignacion_KAM` / `Iniciales de agente` / `Destinatarios`, nunca como literal de fallback (ver deuda 16).
- **`LISTA_PRECIOS`** debe coincidir con el valor de la columna `Lista Cliente` de `Productos` (idéntico al de `LISTA DE PRECIOS` en `Asignacion_KAM` para ese `SAP ID`, celdas multilínea incluidas). El precio del cliente es la columna **`Costo`** de esa fila, no `PVP Final` — y `Costo` es un porcentaje, no un precio, en las listas de descuento.
- El descubrimiento de archivos pendientes de un cliente pasa siempre por `obtenerArchivosCliente_(folder, configCli)` (en `main.js`), nunca por un `if/else` de nombre de cliente.
- La UI de un cliente nuevo se agrega **registrándolo en `CLIENTES_UI`** (`wizard_ui.html`) con sus textos, `colorAccent` y las tres funciones backend — no se crean archivos `<cliente>_ui.html` con copias del wizard. Solo un flujo genuinamente distinto justifica `flujoCustom: true` + su propio `initModulo<CLIENTE>` (incluido vía `include()`, nunca pegado inline en `index.html`).
- Antes de dar un cliente por "funcional", debe tener tarjeta en el dashboard (`index.html`) con `verificarPendientes` + `abrirModulo`, y su entrada en `CLIENTES_UI` (o su flujo custom realmente incluido).
