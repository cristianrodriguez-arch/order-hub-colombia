# Order Hub — ISDIN Colombia

Portal web interno del equipo de **facturación** de ISDIN Colombia. Automatiza el ciclo de una orden de compra: lee el archivo que el cliente deja en Drive, lo cruza contra los datos maestros en Sheets, y genera el CSV de carga para SAP.

> ⚠️ **Repositorio privado.** Contiene IDs de Drive/Sheets, SAP IDs de clientes reales y correos internos. No hacer público.

---

## Qué hace

Por cada cliente, el mismo pipeline de 4 pasos:

1. **Leer** el archivo que el cliente deposita en su carpeta de Drive (PDF, Excel o consolidado del portal, según el cliente).
2. **Validar** contra los datos maestros en Sheets — el usuario confirma fechas, resuelve EAN faltantes y elige lote cuando hay varios.
3. **Escribir** el resumen en la hoja `CONSOLIDADO` y generar el CSV con el formato de carga a SAP.
4. **Archivar**: guardar el CSV en Drive, mover el archivo original a la carpeta de procesados y notificar al KAM por correo.

El CSV **no se descarga desde el navegador** — queda guardado en la carpeta de Drive y el usuario lo toma de ahí.

## Clientes soportados

Cada uno es un módulo independiente en el dashboard, con su propio motor:

| Módulo | Cliente / razón social | SAP ID | Formato de entrada | Particularidad |
|---|---|---|---|---|
| **MEDIPIEL** | MEDIPIEL S.A.S. | `11026712` | PDF (OCR) | El precio sale del propio PDF (columna COSTO UN) |
| **CMX** | CMX S.A.S. (Línea Estética) | `11033482` | Excel (`.xls`/`.xlsx`/`.xlsm`) | Cada comprador usa su propia plantilla → lectura heurística. Precio desde la lista de precios del cliente |
| **CEN** | *varios* (ver abajo) | *varios* | Excel consolidado del portal | **Multicliente**: un solo archivo trae órdenes de varios clientes finales |
| FARMATODO | FARMATODO COLOMBIA S.A. | `11049529` ⁽¹⁾ | — | Solo tarjeta en el dashboard, sin motor todavía |

⁽¹⁾ Farmatodo aún **no tiene entrada en `CONFIG.CLIENTES`**; ese SAP ID sale de la hoja `Asignacion_KAM`, no del código.

### Clientes finales dentro de CEN (portal Carvajal)

El archivo que baja del portal mezcla órdenes de varios clientes. Cada fila se homologa contra `CONFIG.CLIENTES.CEN.MAPEO_CLIENTES` (`main.js`) usando la razón social del comprador, y el wizard muestra una pantalla extra para tramitarlos uno por uno.

| Comprador en el portal | Cliente interno ISDIN | SAP ID | Lista de precios |
|---|---|---|---|
| Bella Piel S A S. | BELLA PIEL | `11026727` | Bella Piel |
| Colsubsidio | CAJA COLOMBIANA DE SUBSIDIO FAMILIA | `11026688` | Distribuidor |
| SUPERTIENDAS Y DROGUERIAS OLIMPICAS S.A. | SUPERTIENDAS Y DROGUERIAS OLIMPICA | `11059838` | Distribuidor |
| Copservir Ltda - Drogas La Rebaja | COPERATIVA MULTIACTIVA DE SERVICIOS | `11072880` | Distribuidor |
| Cafam | CAJA DE COMPENSACION FAMILIAR CAFAM | `11048830` | Distribuidor |
| Coopidrogas | COOPERATIVA NACIONAL DE DROGUISTAS | `11026732` | Copidrogas |

> **Pendientes de habilitar**: `Comfandi` (`11033303`) y `Distribuciones Axa S.A.` (`11045975`) aparecen como compradores en archivos reales del portal pero todavía no están en `MAPEO_CLIENTES`. Sus órdenes se reportan como error y se saltan hasta que se agreguen (rama `fix/cen-agregar-comfandi-axa`, en validación).

Un comprador que no esté en la tabla **no se procesa**: se reporta en `errores` para que alguien lo agregue. Es deliberado — antes se le inventaba un SAP ID y se cargaba con la lista de precios de otro cliente.

## Stack

**Google Apps Script** (runtime V8, web app con HtmlService). No hay Node, ni build, ni tests, ni linter — `package.json` existe solo para traer `@google/clasp`. El frontend es HTML + Tailwind (CDN) + SweetAlert2, sin framework.

Todo el código vive en un único proyecto Apps Script plano: **un solo scope global**, sin módulos. El "namespacing" es un prefijo de texto por cliente (`medipiel_`, `cmx_`, `cen_`).

Requiere el servicio avanzado **Drive API v2** habilitado (ya declarado en `appsscript.json`) — lo usan el OCR de PDFs y la conversión de Excel a Sheets.

## Estructura

```
main.js              CONFIG global (única fuente de config por cliente) + helpers compartidos
index.html           Dashboard con las tarjetas de cliente
wizard_ui.html       Wizard unificado: un solo componente para todos los clientes
<cliente>_motor.js   Motor por cliente: preAnalizar<CLIENTE>() / guardarDefinitivo<CLIENTE>()
test.js              Diagnóstico manual de permisos de Drive (no es una suite de tests)
```

**Contrato de cada motor** — dos etapas, invocadas desde el wizard vía `google.script.run`:

- `preAnalizar<CLIENTE>()` → lee Drive y devuelve `{ pedidosExtraidos, itemsExtraidos, errores, yaProcesados }`
- `guardarDefinitivo<CLIENTE>(paquete)` → escribe CONSOLIDADO + CSV, archiva y notifica

La UI de un cliente **no** es un archivo aparte: se registra en `CLIENTES_UI` dentro de `wizard_ui.html` (textos, color, y los nombres de sus funciones backend).

## Hojas de cálculo que consume

Todo vive en un mismo Spreadsheet (`CONFIG.SPREADSHEET_ID`):

| Hoja | Rol |
|---|---|
| `CONSOLIDADO` | Log de pedidos procesados. **Encabezados en la fila 3** (la fila 1 es el título) |
| `Productos` | Catálogo: una fila por (producto × lista de precios). El precio del cliente es la columna `Costo`, **no** `PVP Final` |
| `Bodega` | Stock SAP por EAN, con lotes y fechas de caducidad |
| `Destinatarios` | Mapeo destino cliente → destinatario SAP + ciudad |
| `Asignacion_KAM` | **Fuente única del KAM**: por SAP ID trae nombre, correo, iniciales y lista de precios |

> La hoja `Iniciales de agente` **ya no se usa** — su contenido duplicaba lo que ya está en `Asignacion_KAM`.

## Desarrollo

No hay forma de ejecutar el código localmente. Se prueba desplegando y abriendo la web app.

```bash
npm install          # solo la primera vez (trae clasp)
clasp push           # sube los cambios locales al proyecto Apps Script
clasp pull           # baja lo que esté en el editor online
clasp open-script    # abre el editor en el navegador
```

**`clasp push` publica en vivo.** No hay ambiente de staging: lo que subes es lo que ve facturación. Mergear a `main` en GitHub no despliega nada — el despliegue es siempre manual:

```bash
git pull origin main
clasp push
```

## Agregar un cliente nuevo

1. Entrada en `CONFIG.CLIENTES[<CLIENTE>]` (`main.js`) con sus IDs de carpeta, `SOLICITANTE`, `NOMBRE_CLIENTE`, `LISTA_PRECIOS` y bloque `SAP`. **Ninguna constante de negocio va hardcodeada en el motor.**
2. `<cliente>_motor.js` con el par `preAnalizar` / `guardarDefinitivo`, escribiendo en CONSOLIDADO vía `CONSOLIDADO_HEADERS`.
3. Registro en `CLIENTES_UI` (`wizard_ui.html`).
4. Tarjeta en el dashboard (`index.html`) con `verificarPendientes` + `abrirModulo`.

El detalle completo de convenciones, decisiones de arquitectura y deuda técnica conocida está en [CLAUDE.md](CLAUDE.md).

## Deuda técnica

`CLAUDE.md` lleva el inventario detallado. Los puntos abiertos más relevantes:

- **Lógica duplicada** entre motores (lectura de `Productos`/`Bodega`, HTML del correo al KAM) — candidata a un `utils.js` compartido.
- **CMX**: el EAN llega como número en algunas plantillas y se pierde por notación exponencial; las fechas de vencimiento de lote salen siempre "Sin fecha"; el antiduplicados no detecta OCs repetidas.
- **Zona horaria inconsistente**: `appsscript.json` declara `Europe/Madrid` mientras el código usa `"GMT-5"` hardcodeado en varios puntos.
- **`MISSING_BODEGA`**: MEDIPIEL todavía lanza ese protocolo pero el wizard ya no lo maneja — el usuario ve el string crudo.
- **CSV**: escapa `;` y `"` pero no neutraliza fórmulas (`=`, `+`, `-`, `@`).

---

Uso interno de ISDIN. No distribuir.
