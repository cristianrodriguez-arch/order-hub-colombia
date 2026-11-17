/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * ARCHIVO: falabella_motor.js
 * MÓDULO: Integración de órdenes de compra (OC) de FALABELLA DE COLOMBIA a partir de los Excel del B2B
 * =========================================================================================
 * CÓMO SON LAS ÓRDENES DE FALABELLA
 *   El B2B de Falabella entrega dos formatos de Excel. El motor reconoce las columnas por su nombre,
 *   así que la hoja puede estar en cualquier posición del libro (el KAM a veces agrega tablas dinámicas).
 *   1. "distributions-details" (Cross-Dock): una fila por producto y LOCAL. Columnas que se usan:
 *      Número OC, UPC (= EAN), Modelo, Unidades compradas, Fecha de emisión, Fecha fin recepción,
 *      ID lugar de entrega / Descripción lugar de entrega (a dónde se despacha, p. ej. 8030 CD Future WMOS),
 *      Id de local / Local (la tienda de Falabella).
 *   2. "products-details" (Almacenaje): una fila por producto, SIN local ni lugar de entrega.
 *
 * REGLA DE NEGOCIO
 *   En SAP se sube la misma OC una vez por cada local. Por eso cada combinación OC + local es un pedido:
 *   tiene su propio consecutivo, su propia fila en CONSOLIDADO y su propio bloque en el CSV.
 *   - CONSOLIDADO, columna "Orden De Compra": solo la OC. El Order Hub NO escribe el local: lo pone
 *     facturación al prefacturar, en la segunda columna "NO ENTREGA FACTURACIÓN" (M), p. ej. "36 SAN DIEGO".
 *   - CSV, "Num pedido cliente": solo la OC (la misma para todos los locales).
 *   - Duplicados: cliente (SAP ID) + OC + local, leyendo el local de esa columna M. Las filas "ANULADO"
 *     (columna L) no cuentan. Si la OC tiene filas sin local todavía, toda la OC se toma como registrada
 *     (ver falabella_motivoRegistrado_).
 *   - Producto cruzado: solo si su EAN (UPC) está en la hoja Productos. Si solo está en Bodega, va a
 *     "no cruzados": el cliente solo acepta el EAN exacto de la OC.
 *   - Fecha de entrega: se propone la "Fecha fin recepción" (fecha límite de Falabella); se puede cambiar
 *     en el asistente.
 *   - Destino: el "ID lugar de entrega" se busca en la hoja Destinatarios, en la columna Cruce_cliente de
 *     las filas de Falabella (Id_origen = SOLICITANTE). Si no está, el asistente pide el ID de destino.
 *   - Precio: columna Costo de Productos en la lista de Falabella (CONFIG.CLIENTES.FALABELLA.LISTA_PRECIOS).
 *
 * QUIÉN LO LLAMA
 *   - wizard_ui.html (CLIENTES_UI.FALABELLA): preAnalizarFALABELLA, guardarDefinitivoFALABELLA y
 *     falabella_resolverDestinoPorId.
 *   - Usa de main.js: CONFIG, leerOcsRegistradasConsolidado_, obtenerMapaBodegaGeneral_,
 *     obtenerInfoKamPorSolicitante_, construirCsvSap_, nombreArchivoSap_, enviarCorreoKam_, conBloqueoScript_
 *     y obtenerArchivosCliente_.
 * =========================================================================================
 */

/** Con más posiciones que esto, el correo a la KAM muestra un resumen por OC (Gmail limita el tamaño del cuerpo). */
const FALABELLA_MAX_POSICIONES_DETALLE_CORREO = 150;

/**
 * Columnas del Excel de Falabella. Cada campo se busca por el nombre normalizado del encabezado
 * (sin tildes, espacios ni signos). Los obligatorios son OC, UPC y Unidades compradas.
 */
const FALABELLA_COLUMNAS = {
  oc:            ["NUMEROOC"],
  upc:           ["UPC"],
  descripcion:   ["MODELO", "DESCRIPCION"],
  unidades:      ["UNIDADESCOMPRADAS"],
  canceladas:    ["UNIDADESCANCELADAS"],
  fechaEmision:  ["FECHADEEMISION"],
  fechaFin:      ["FECHAFINRECEPCION"],
  lugarId:       ["IDLUGARDEENTREGA"],
  lugarNombre:   ["DESCRIPCIONLUGARDEENTREGA"],
  localId:       ["IDDELOCAL"],
  localNombre:   ["LOCAL"],
  estado:        ["ESTADOOC"],
  tipo:          ["TIPODEORDEN"]
};

/**
 * Lee los Excel de Falabella y devuelve los pedidos (uno por OC + local) y sus productos, para que el
 * asistente los muestre antes de guardar. Solo lee: no escribe en CONSOLIDADO ni mueve archivos.
 * Crea copias temporales de los Excel en Google Sheets y las manda a la papelera al terminar.
 *
 * @return {{pedidosExtraidos: Array<Object>, itemsExtraidos: Array<Object>, errores: Array<String>,
 *           yaProcesados: Array<String>}}
 * @throws Si falta la configuración de FALABELLA o no se puede revisar CONSOLIDADO.
 */
function preAnalizarFALABELLA() {
  const configCli = CONFIG.CLIENTES["FALABELLA"];
  if (!configCli) throw new Error("No se encontró la configuración del cliente FALABELLA en CONFIG de main.js.");
  if (!configCli.FOLDER_ID) throw new Error("Falta la carpeta de Falabella (CONFIG.CLIENTES.FALABELLA.FOLDER_ID en main.js).");

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const registrados = falabella_leerRegistrados_(ss, configCli);
  const precios = falabella_getPreciosMap_(ss, configCli);
  const bodega = obtenerMapaBodegaGeneral_(ss.getSheetByName(CONFIG.SHEET_BODEGA || "Bodega"));
  const destinos = falabella_leerDestinos_(ss, configCli.SOLICITANTE);

  const archivos = falabella_listarArchivos_(configCli);
  const tablas = [];
  const errores = [];

  archivos.forEach(file => {
    let tempId = null;
    try {
      tempId = falabella_convertirExcelASheet_(file);
      const libro = SpreadsheetApp.openById(tempId);
      const tabla = falabella_buscarTabla_(libro.getSheets().map(h => h.getDataRange().getValues()));
      if (!tabla) throw new Error("no tiene la tabla de la OC (columnas Número OC, UPC y Unidades compradas).");
      tablas.push({ archivoId: file.getId(), archivoNombre: file.getName(), filas: tabla });
    } catch (e) {
      errores.push(`${file.getName()}: ${e.message}`);
    } finally {
      if (tempId) {
        try { DriveApp.getFileById(tempId).setTrashed(true); } catch (ignore) {}
      }
    }
  });

  const resultado = falabella_armarPedidos_(tablas, { registrados, precios, bodega, destinos, configCli });
  return {
    pedidosExtraidos: resultado.pedidosExtraidos,
    itemsExtraidos: resultado.itemsExtraidos,
    errores: errores.concat(resultado.errores),
    yaProcesados: resultado.yaProcesados
  };
}

/**
 * Lista los Excel (.xlsx, .xlsm, .xls) de la carpeta de Falabella, sin los temporales ("_temp").
 * @param {Object} configCli CONFIG.CLIENTES.FALABELLA.
 * @return {Array<DriveApp.File>} Hasta CONFIG.LIMIT_FILES archivos (200 si es 0).
 */
function falabella_listarArchivos_(configCli) {
  const it = obtenerArchivosCliente_(DriveApp.getFolderById(configCli.FOLDER_ID), configCli);
  const max = CONFIG.LIMIT_FILES > 0 ? CONFIG.LIMIT_FILES : 200;
  const archivos = [];
  while (it.hasNext() && archivos.length < max) {
    const f = it.next();
    const nombre = f.getName().toLowerCase();
    if (nombre.includes("_temp")) continue;
    if (nombre.endsWith(".xlsx") || nombre.endsWith(".xlsm") || nombre.endsWith(".xls")) archivos.push(f);
  }
  return archivos;
}

/**
 * Crea una copia del Excel como Google Sheets para leerlo. La copia queda en la raíz del Drive de quien
 * ejecuta (no en la carpeta de pendientes), así el contador del tablero no la cuenta si algo falla.
 * @param {DriveApp.File} file Excel original.
 * @return {String} Id de la copia (quien llama la manda a la papelera).
 */
function falabella_convertirExcelASheet_(file) {
  const recurso = { title: file.getName().replace(/\.[^/.]+$/, "") + "_temp", mimeType: MimeType.GOOGLE_SHEETS };
  return Drive.Files.insert(recurso, file.getBlob()).id;
}

/**
 * Busca, entre las hojas de un libro, la tabla de la OC: la primera hoja que en sus primeras 10 filas tenga
 * una fila de encabezados con Número OC, UPC y Unidades compradas. Así se ignoran las tablas dinámicas
 * que el KAM agrega al archivo.
 *
 * @param {Array<Array<Array<*>>>} hojas Valores de cada hoja (getValues()), en orden.
 * @return {Array<Object>|null} Una entrada por fila de datos { oc, upc, descripcion, unidades, fechaEmision,
 *   fechaFin, lugarId, lugarNombre, localId, localNombre, estado, tipo }, o null si ninguna hoja la tiene.
 */
function falabella_buscarTabla_(hojas) {
  for (const data of hojas) {
    for (let r = 0; r < Math.min(data.length, 10); r++) {
      const cols = falabella_detectarColumnas_(data[r]);
      if (cols.oc === -1 || cols.upc === -1 || cols.unidades === -1) continue;

      const filas = [];
      const leer = (fila, campo) => cols[campo] === -1 ? "" : fila[cols[campo]];
      for (let i = r + 1; i < data.length; i++) {
        const fila = data[i];
        const oc = falabella_limpiarCodigo_(leer(fila, "oc"));
        if (!oc) continue;
        const compradas = Number(leer(fila, "unidades")) || 0;
        const canceladas = Number(leer(fila, "canceladas")) || 0;
        filas.push({
          oc: oc,
          upc: falabella_limpiarCodigo_(leer(fila, "upc")),
          descripcion: String(leer(fila, "descripcion") || "").trim(),
          unidades: Math.max(0, compradas - canceladas),
          fechaEmision: falabella_fechaIso_(leer(fila, "fechaEmision")),
          fechaFin: falabella_fechaIso_(leer(fila, "fechaFin")),
          lugarId: falabella_limpiarCodigo_(leer(fila, "lugarId")),
          lugarNombre: String(leer(fila, "lugarNombre") || "").trim(),
          localId: falabella_limpiarCodigo_(leer(fila, "localId")),
          localNombre: String(leer(fila, "localNombre") || "").trim(),
          estado: String(leer(fila, "estado") || "").trim(),
          tipo: String(leer(fila, "tipo") || "").trim()
        });
      }
      return filas;
    }
  }
  return null;
}

/**
 * Ubica cada campo de FALABELLA_COLUMNAS en una fila de encabezados. Compara el nombre completo
 * normalizado (no por contenido), para que "Local" no se confunda con "Id de local".
 * @param {Array<*>} encabezados Fila candidata a encabezados.
 * @return {Object} { campo: índice base 0, o -1 si no está }.
 */
function falabella_detectarColumnas_(encabezados) {
  const norm = encabezados.map(falabella_normalizar_);
  const cols = {};
  Object.keys(FALABELLA_COLUMNAS).forEach(campo => {
    cols[campo] = -1;
    for (const nombre of FALABELLA_COLUMNAS[campo]) {
      const idx = norm.indexOf(nombre);
      if (idx !== -1) { cols[campo] = idx; break; }
    }
  });
  return cols;
}

/**
 * Agrupa las filas de todos los archivos en pedidos (OC + local) y arma sus productos.
 * Función sin acceso a Google: recibe todo lo que necesita, para poder probarla aparte.
 *
 * @param {Array<{archivoId, archivoNombre, filas}>} tablas Tablas leídas (falabella_buscarTabla_).
 * @param {Object} ctx { registrados: Set de claves ya en CONSOLIDADO, precios: mapa de falabella_getPreciosMap_,
 *   bodega: mapa de obtenerMapaBodegaGeneral_, destinos: lista de falabella_leerDestinos_, configCli }.
 * @return {{pedidosExtraidos, itemsExtraidos, errores, yaProcesados}}
 */
function falabella_armarPedidos_(tablas, ctx) {
  const pedidosExtraidos = [];
  const itemsExtraidos = [];
  const errores = [];
  const yaProcesados = [];
  const pedidosPorClave = {};
  const archivoDeClave = {};

  tablas.forEach(tabla => {
    // Una misma OC + local no puede venir en dos archivos (p. ej. el mismo Excel descargado dos veces).
    const repetidas = new Set();
    tabla.filas.forEach(f => {
      const clave = falabella_clavePedido_(f);
      if (archivoDeClave[clave] && archivoDeClave[clave] !== tabla.archivoId) repetidas.add(f.oc);
    });
    if (repetidas.size > 0) {
      errores.push(`${tabla.archivoNombre}: la OC ${[...repetidas].join(", ")} ya viene en otro archivo de la carpeta. Deje solo uno.`);
      return;
    }

    tabla.filas.forEach(f => {
      const estadoNorm = falabella_normalizar_(f.estado);
      if (estadoNorm.includes("ANUL") || estadoNorm.includes("CANCEL")) {
        const msg = `${tabla.archivoNombre}: la OC ${f.oc} está en estado "${f.estado}" y no se procesa.`;
        if (!errores.includes(msg)) errores.push(msg);
        return;
      }

      const clave = falabella_clavePedido_(f);
      const etiqueta = f.localNombre ? `${f.oc} · ${f.localId ? f.localId + " " : ""}${f.localNombre}` : f.oc;
      const motivo = falabella_motivoRegistrado_(ctx.registrados, f.oc, f.localId, f.localNombre);
      if (motivo) {
        // Si toda la OC está bloqueada por filas sin local, se informa una sola vez por OC
        const texto = motivo.startsWith("la OC") ? `${f.oc} (${motivo})` : etiqueta;
        if (!yaProcesados.includes(texto)) yaProcesados.push(texto);
        return;
      }

      let pedido = pedidosPorClave[clave];
      if (!pedido) {
        const destino = falabella_resolverDestino_(f.lugarId, ctx.destinos);
        const pedidoId = f.localId ? `${f.oc}-${f.localId}` : (f.localNombre ? `${f.oc}-${falabella_normalizar_(f.localNombre)}` : f.oc);
        pedido = {
          Pedido_ID: pedidoId,
          Etiqueta: f.localNombre ? etiqueta : `${f.oc}${f.tipo ? " · " + f.tipo : ""}`,
          OC: f.oc,
          Local_Id: f.localId,
          Local_Nombre: f.localNombre,
          Tipo_Orden: f.tipo,
          Fecha_Pedido: falabella_isoADdMmYyyy_(f.fechaEmision),
          Fecha_Entrega: f.fechaFin,
          Lugar_Entrega_Id: f.lugarId,
          Lugar_Entrega_Nombre: f.lugarNombre,
          Cruce_Destino: f.lugarId || "SIN_LUGAR",
          Id_Destino: destino ? destino.idDestino : "",
          Destinatario: destino ? destino.nombreDestino : (f.lugarNombre || ""),
          Ciudad: destino ? destino.ciudad : "",
          Destino_Pendiente: !destino,
          Archivo_Nombre: tabla.archivoNombre,
          Archivo_Id: tabla.archivoId
        };
        pedidosPorClave[clave] = pedido;
        archivoDeClave[clave] = tabla.archivoId;
        pedidosExtraidos.push(pedido);
      }

      if (f.unidades <= 0 || !f.upc) return;

      // Si el mismo UPC aparece dos veces para el mismo local, se suman las unidades.
      const existente = itemsExtraidos.find(i => i.Pedido_ID === pedido.Pedido_ID && i.Producto_Cod === f.upc);
      if (existente) {
        existente.Cant += f.unidades;
        return;
      }
      const prod = falabella_resolverProducto_(f.upc, ctx.precios, ctx.bodega);
      itemsExtraidos.push({
        Item_ID: `${pedido.Pedido_ID}_${itemsExtraidos.filter(i => i.Pedido_ID === pedido.Pedido_ID).length + 1}`,
        Pedido_ID: pedido.Pedido_ID,
        Producto_Cod: f.upc,
        Descripcion: prod.descripcion || f.descripcion || "Producto Falabella",
        Cant: f.unidades,
        Valor_Unitario: prod.precio,
        EAN_Mapeado: prod.eanMapeado,
        Materiales_Disponibles: prod.materiales
      });
    });
  });

  // Un pedido sin ninguna posición válida no se muestra (no habría nada que registrar).
  const conItems = new Set(itemsExtraidos.map(i => i.Pedido_ID));
  return {
    pedidosExtraidos: pedidosExtraidos.filter(p => conItems.has(p.Pedido_ID)),
    itemsExtraidos: itemsExtraidos,
    errores: errores,
    yaProcesados: yaProcesados
  };
}

/**
 * Material, lotes y precio de un UPC.
 * - El producto CRUZA solo si su EAN está en la hoja Productos. Si no está (aunque esté en Bodega), queda
 *   sin cruzar: el asistente lo muestra en "no cruzados" y va al CSV sin material, porque Falabella solo
 *   acepta el EAN exacto de la OC.
 * - Material oficial: el de Productos. Precio: Costo de la lista de Falabella (0 si el EAN está en Productos
 *   pero no en esa lista).
 * - Lotes: hoja Bodega. Sin lotes, va el material con stock 0 (quedará como back order).
 * @return {{eanMapeado: String, precio: Number, descripcion: String, materiales: Array<Object>}}
 */
function falabella_resolverProducto_(upc, productos, bodega) {
  const prod = productos[upc];
  if (!prod) return { eanMapeado: "", precio: 0, descripcion: "", materiales: [] };

  const lotes = bodega[upc] || [];
  const materiales = lotes.length > 0
    ? lotes.map(l => ({ material: prod.material, stock: l.stock, fechaDisplay: l.fechaDisplay }))
    : [{ material: prod.material, stock: 0, fechaDisplay: "Sin stock" }];
  return {
    eanMapeado: upc,
    precio: prod.precio,
    descripcion: prod.descripcion,
    materiales: materiales
  };
}

/** Clave de un pedido dentro de la lectura de archivos: OC + Id de local (o nombre, o nada si es Almacenaje). */
function falabella_clavePedido_(f) {
  return `${f.oc}|${f.localId || falabella_normalizar_(f.localNombre)}`;
}

/**
 * Lee de CONSOLIDADO lo que ya está registrado de Falabella, para revisar duplicados por OC + local.
 * El local NO lo escribe el Order Hub: lo pone facturación al prefacturar, en la SEGUNDA columna
 * "NO ENTREGA FACTURACIÓN" (columna M), con el formato "<Id de local> <Local>", p. ej. "36 SAN DIEGO".
 * Las filas con "ANULADO" en la primera "NO ENTREGA FACTURACIÓN" (columna L) no cuentan: ese pedido
 * se puede volver a registrar.
 *
 * @return {{locales: Set<String>, ocs: Object}} Ver falabella_registrosDesdeFilas_.
 * @throws Si CONSOLIDADO no tiene las columnas para revisar duplicados.
 */
function falabella_leerRegistrados_(ss, configCli) {
  const sh = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  if (!sh || sh.getLastRow() < 4) return falabella_registrosDesdeFilas_([], [], configCli.SOLICITANTE);
  const data = sh.getDataRange().getValues();
  return falabella_registrosDesdeFilas_(data.slice(0, 5), data, configCli.SOLICITANTE);
}

/**
 * Arma el índice de registros de Falabella a partir de los valores de CONSOLIDADO (función sin acceso a
 * Google, para poder probarla aparte).
 * @param {Array<Array<*>>} primeras Primeras filas de la hoja (para encontrar la fila de encabezados).
 * @param {Array<Array<*>>} data Todas las filas de la hoja.
 * @param {String} solicitante SAP ID de Falabella.
 * @return {{locales: Set<String>, ocs: Object}}
 *   locales: claves "<OC>|ID:<id de local>" y "<OC>|N:<nombre normalizado>" de los locales ya registrados.
 *   ocs: { OC: { filas: N, sinLocal: N } } filas vigentes (no anuladas) por OC y cuántas aún no tienen local.
 * @throws Si faltan las columnas Orden De Compra, Sap Id o las dos "NO ENTREGA FACTURACIÓN".
 */
function falabella_registrosDesdeFilas_(primeras, data, solicitante) {
  const resultado = { locales: new Set(), ocs: {} };
  if (data.length === 0) return resultado;
  const norm = h => String(h).trim().toUpperCase();
  const filaEnc = primeras.findIndex(f => f.some(h => norm(h) === "ORDEN DE COMPRA"));
  if (filaEnc === -1) throw new Error("No se encontró el encabezado 'Orden De Compra' en CONSOLIDADO; no se puede revisar duplicados.");
  const enc = data[filaEnc].map(h => falabella_normalizar_(h));
  const colOc = enc.indexOf("ORDENDECOMPRA");
  const colSap = enc.indexOf("SAPID");
  const colsNoEntrega = enc.map((h, i) => h === "NOENTREGAFACTURACION" ? i : -1).filter(i => i !== -1);
  if (colSap === -1 || colsNoEntrega.length < 2) {
    throw new Error("CONSOLIDADO debe tener 'Sap Id' y dos columnas 'NO ENTREGA FACTURACIÓN' (la segunda es el local de Falabella); no se puede revisar duplicados.");
  }
  const [colEstado, colLocal] = colsNoEntrega;
  const sapCli = falabella_normalizar_(solicitante);

  for (let i = filaEnc + 1; i < data.length; i++) {
    const fila = data[i];
    if (falabella_normalizar_(fila[colSap]) !== sapCli) continue;
    const oc = falabella_limpiarCodigo_(fila[colOc]);
    if (!oc) continue;
    if (falabella_normalizar_(fila[colEstado]).includes("ANULAD")) continue;

    const info = resultado.ocs[oc] || (resultado.ocs[oc] = { filas: 0, sinLocal: 0 });
    info.filas++;
    const local = String(fila[colLocal] || "").trim();
    if (!local) { info.sinLocal++; continue; }
    const m = local.match(/^(\d+)\s*[-–.]?\s*(.*)$/);
    if (m) {
      resultado.locales.add(`${oc}|ID:${m[1]}`);
      if (m[2]) resultado.locales.add(`${oc}|N:${falabella_normalizar_(m[2])}`);
    } else {
      resultado.locales.add(`${oc}|N:${falabella_normalizar_(local)}`);
    }
  }
  return resultado;
}

/**
 * ¿Este pedido (OC + local) ya está en CONSOLIDADO? Devuelve el motivo, o "" si es nuevo.
 * - Con local: está si facturación ya escribió ese local (por Id de local o por nombre) en una fila vigente
 *   de la OC. Si la OC tiene filas vigentes SIN local escrito todavía, no se puede saber qué locales son,
 *   así que toda la OC se considera registrada hasta que facturación complete la columna.
 * - Sin local (Almacenaje): está si la OC tiene cualquier fila vigente.
 */
function falabella_motivoRegistrado_(reg, oc, localId, localNombre) {
  const info = reg.ocs[oc];
  if (!info) return "";
  if (!localId && !localNombre) return "ya registrada";
  if (info.sinLocal > 0) return `la OC tiene ${info.sinLocal} pedido(s) registrados sin local en "NO ENTREGA FACTURACIÓN"`;
  if ((localId && reg.locales.has(`${oc}|ID:${localId}`)) || (localNombre && reg.locales.has(`${oc}|N:${falabella_normalizar_(localNombre)}`))) {
    return "local ya registrado";
  }
  return "";
}

/**
 * Catálogo de la hoja Productos por EAN: { ean: { material, precio, descripcion, enLista } }.
 * Están TODOS los EAN de la hoja (cualquier lista), porque estar en Productos es lo que hace que un
 * producto cruce. El precio solo sale de las filas cuya "Lista Cliente" es exactamente la de Falabella
 * (nunca el precio de otra lista); si el EAN no está en esa lista, precio 0 y enLista false.
 */
function falabella_getPreciosMap_(ss, configCli) {
  const sh = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  const mapa = {};
  if (!sh) return mapa;
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return mapa;
  const head = data[0].map(falabella_normalizar_);
  const col = nombre => head.indexOf(nombre);
  const cEan = col("EAN"), cMat = col("MATERIALID"), cLista = col("LISTACLIENTE"), cCosto = col("COSTO"), cProd = col("PRODUCTO");
  if (cEan === -1 || cMat === -1 || cLista === -1 || cCosto === -1) {
    throw new Error("La hoja Productos no tiene las columnas EAN, Material ID, Lista Cliente y Costo.");
  }
  const lista = falabella_normalizar_(configCli.LISTA_PRECIOS);
  for (let i = 1; i < data.length; i++) {
    const ean = falabella_limpiarCodigo_(data[i][cEan]);
    const material = String(data[i][cMat]).trim();
    if (!ean || !material) continue;
    const esLista = falabella_normalizar_(data[i][cLista]) === lista;
    if (mapa[ean] && mapa[ean].enLista) continue; // ya tiene la fila de la lista de Falabella
    if (mapa[ean] && !esLista) continue;          // primera fila encontrada de otra lista
    mapa[ean] = {
      material: material,
      precio: esLista ? falabella_numero_(data[i][cCosto]) : 0,
      descripcion: cProd !== -1 ? String(data[i][cProd]).trim() : "",
      enLista: esLista
    };
  }
  return mapa;
}

/**
 * Destinos de Falabella en la hoja Destinatarios (filas con Id_origen = solicitante).
 * @return {Array<{idDestino, nombreDestino, cruce, ciudad}>}
 */
function falabella_leerDestinos_(ss, solicitante) {
  const sh = ss.getSheetByName(CONFIG.SHEET_DESTINATARIOS || "Destinatarios");
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const head = data[0].map(falabella_normalizar_);
  const c = { origen: head.indexOf("IDORIGEN"), id: head.indexOf("IDDESTINO"), nombre: head.indexOf("NOMBREDESTINO"), cruce: head.indexOf("CRUCECLIENTE"), ciudad: head.indexOf("CIUDAD") };
  if (c.origen === -1 || c.id === -1) return [];
  const destinos = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][c.origen]).trim() !== String(solicitante).trim()) continue;
    destinos.push({
      idDestino: String(data[i][c.id]).trim(),
      nombreDestino: c.nombre !== -1 ? String(data[i][c.nombre]).trim() : "",
      cruce: c.cruce !== -1 ? falabella_limpiarCodigo_(data[i][c.cruce]) : "",
      ciudad: c.ciudad !== -1 ? String(data[i][c.ciudad]).trim().toUpperCase() : ""
    });
  }
  return destinos;
}

/**
 * Destino de un pedido: la fila de Falabella cuyo Cruce_cliente es el "ID lugar de entrega".
 * Sin lugar de entrega (órdenes de Almacenaje): si Falabella tiene un solo destino, ese.
 * Nunca se inventa ciudad: un destino sin ciudad cuenta como no encontrado.
 * @return {Object|null} Destino de falabella_leerDestinos_, o null (el asistente lo pide).
 */
function falabella_resolverDestino_(lugarId, destinos) {
  let destino = null;
  if (lugarId) destino = destinos.find(d => d.cruce === lugarId) || null;
  else if (destinos.length === 1) destino = destinos[0];
  return destino && destino.ciudad ? destino : null;
}

/**
 * Llamada desde el asistente cuando un lugar de entrega no está homologado: el usuario escribe el
 * Id_destino y se valida que exista entre los destinos de Falabella en Destinatarios y tenga ciudad.
 * No escribe en la hoja: si el destino no existe, hay que agregarlo en Destinatarios (con Cruce_cliente
 * = ID lugar de entrega, para que la próxima vez se reconozca solo).
 * @param {String} idDestino Id_destino escrito por el usuario.
 * @return {{idDestino, nombreDestino, ciudad}}
 * @throws Si el ID no está entre los destinos de Falabella o no tiene ciudad.
 */
function falabella_resolverDestinoPorId(idDestino) {
  const id = String(idDestino || "").trim();
  const configCli = CONFIG.CLIENTES["FALABELLA"];
  const destino = falabella_leerDestinos_(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID), configCli.SOLICITANTE).find(d => d.idDestino === id);
  if (!destino) {
    throw new Error(`El ID ${id} no está en Destinatarios para Falabella (Id_origen ${configCli.SOLICITANTE}). Agréguelo en la hoja, con el ID lugar de entrega en Cruce_cliente, y vuelva a intentar.`);
  }
  if (!destino.ciudad) throw new Error(`El ID ${id} no tiene ciudad en Destinatarios. Complétela y vuelva a intentar.`);
  return { idDestino: destino.idDestino, nombreDestino: destino.nombreDestino, ciudad: destino.ciudad };
}

/**
 * Guarda los pedidos de Falabella confirmados en el asistente. Va dentro del bloqueo del script para que
 * dos guardados a la vez no repitan consecutivos.
 * @param {Object} paquete Lo que envía el asistente (pedidos, items, fechasEntregas, materialesSeleccionados,
 *   noActualizadosSeleccionados, backordersSeleccionados, notificarKAM).
 * @return {{success, mensaje, csvData, fileName}}
 */
function guardarDefinitivoFALABELLA(paquete) {
  return conBloqueoScript_(() => falabella_guardarSinBloqueo_(paquete));
}

/**
 * Lógica del guardado (dentro del bloqueo):
 * 1. Rechaza pedidos sin destino y vuelve a revisar duplicados en CONSOLIDADO (si alguno ya está, no guarda nada).
 * 2. Asigna un consecutivo por pedido (OC + local) con las iniciales de la KAM.
 * 3. Guarda el CSV (si falla, no escribe nada), luego escribe CONSOLIDADO (si falla, manda el CSV a la papelera).
 * 4. Mueve cada Excel a procesados y lo renombra con sus consecutivos.
 * 5. Envía el correo a la KAM si se pidió (un fallo del correo no deshace el guardado).
 */
function falabella_guardarSinBloqueo_(paquete) {
  const configCli = CONFIG.CLIENTES["FALABELLA"];
  if (!configCli) throw new Error("No se encontró la configuración para FALABELLA.");
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  if (!sh) throw new Error("No se encontró la hoja CONSOLIDADO.");

  const pedidos = paquete.pedidos || [];
  const sinDestino = pedidos.filter(p => p.Destino_Pendiente || !p.Id_Destino || !p.Ciudad);
  if (sinDestino.length > 0) {
    throw new Error("Hay pedidos sin destino confirmado: " + sinDestino.map(p => p.Etiqueta || p.Pedido_ID).join(", ") + ".");
  }

  const registrados = falabella_leerRegistrados_(ss, configCli);
  const repetidos = pedidos.filter(p => falabella_motivoRegistrado_(registrados, p.OC, p.Local_Id, p.Local_Nombre));
  if (repetidos.length > 0) {
    throw new Error("Ya están registrados en CONSOLIDADO, por eso no se guardó nada: " + repetidos.map(p => p.Etiqueta || p.Pedido_ID).join(", ") + ".");
  }

  const kam = obtenerInfoKamPorSolicitante_(configCli.SOLICITANTE) || { nombre: "SIN KAM", iniciales: "MC", correo: "" };
  const prefijo = kam.iniciales + "-" + String(new Date().getFullYear()).slice(-2);
  let consecutivo = falabella_maxConsecutivo_(sh, prefijo);

  const lineasPorPedido = falabella_lineasPorPedido_(paquete);
  const filas = [];
  const registradosAhora = [];
  const idPorPedido = {};

  pedidos.forEach(p => {
    const lineas = lineasPorPedido[p.Pedido_ID] || [];
    const unidades = lineas.reduce((s, l) => s + l.cantidad, 0);
    if (unidades === 0) return;
    const valor = lineas.reduce((s, l) => s + (l.material ? l.cantidad * l.precio : 0), 0);

    consecutivo++;
    const id = prefijo + String(consecutivo).padStart(4, "0");
    idPorPedido[p.Pedido_ID] = id;
    registradosAhora.push(p);
    const fechaEnt = paquete.fechasEntregas && paquete.fechasEntregas[p.Pedido_ID] ? paquete.fechasEntregas[p.Pedido_ID].fecha : "";
    filas.push([
      id,
      p.OC, // Solo la OC: el local lo escribe facturación en "NO ENTREGA FACTURACIÓN" al prefacturar
      configCli.SOLICITANTE,
      configCli.NOMBRE_CLIENTE,
      p.Fecha_Pedido || Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy"),
      p.Id_Destino,
      p.Destinatario,
      p.Ciudad,
      falabella_isoADdMmYyyy_(fechaEnt),
      unidades,
      valor
    ]); // Solo A:K, igual que CMX: desde la L hay datos y fórmulas que no se tocan
  });

  if (filas.length === 0) throw new Error("No hay posiciones válidas para registrar.");

  const paqueteRegistrado = Object.assign({}, paquete, { pedidos: registradosAhora });
  const csv = construirCsvSap_(registradosAhora.map(p => ({
    sap: configCli.SAP,
    solicitante: configCli.SOLICITANTE,
    destinatario: p.Id_Destino,
    numPedidoCliente: p.OC,
    fechaEntrega: paquete.fechasEntregas && paquete.fechasEntregas[p.Pedido_ID] ? paquete.fechasEntregas[p.Pedido_ID].fecha : "",
    lineas: lineasPorPedido[p.Pedido_ID].map(l => ({ material: l.material, cantidad: l.cantidad }))
  })));

  const ids = registradosAhora.map(p => idPorPedido[p.Pedido_ID]);
  const nombreCsv = nombreArchivoSap_(ids.length === 1 ? `Pedido_${ids[0]}.csv` : `Pedido_${ids[0]}_a_${ids[ids.length - 1]}.csv`);
  let archivoCsv;
  try {
    archivoCsv = DriveApp.getFolderById(CONFIG.FOLDER_CSV_ID).createFile(nombreCsv, csv, MimeType.PLAIN_TEXT);
  } catch (e) {
    throw new Error("No se pudo guardar el CSV en Drive (no se registró nada en CONSOLIDADO): " + e.message);
  }

  try {
    const primera = sh.getLastRow() + 1;
    sh.getRange(primera, 1, filas.length, filas[0].length).setValues(filas);
    SpreadsheetApp.flush();
  } catch (e) {
    try { archivoCsv.setTrashed(true); } catch (ignore) {}
    throw new Error("Fallo al escribir en CONSOLIDADO; el CSV se envió a la papelera: " + e.message);
  }

  const avisos = falabella_moverArchivos_(registradosAhora, idPorPedido, configCli);

  if (paquete.notificarKAM !== false) {
    try {
      falabella_enviarCorreoKam_(paqueteRegistrado, lineasPorPedido, idPorPedido, kam, configCli);
    } catch (e) {
      console.error("Error enviando correo a la KAM de Falabella: " + e.message);
      avisos.push("no se pudo enviar el correo a la KAM (" + e.message + ")");
    }
  }

  return {
    success: true,
    mensaje: `Se registraron ${filas.length} pedidos de Falabella (${ids[0]}${ids.length > 1 ? " a " + ids[ids.length - 1] : ""}).` +
      (avisos.length > 0 ? " AVISO: " + avisos.join("; ") : ""),
    csvData: Utilities.base64Encode(csv, Utilities.Charset.UTF_8).replace(/[\r\n]/g, ""),
    fileName: nombreCsv
  };
}

/**
 * Líneas finales de cada pedido con el material elegido en el asistente. Un producto "no actualizado"
 * va con material vacío (SAP lo deja como back order) y no suma valor; mismo criterio que CMX y CEN.
 * @return {Object} { Pedido_ID: [{ itemId, ean, descripcion, material, cantidad, precio, backOrder }] }
 */
function falabella_lineasPorPedido_(paquete) {
  const mapa = {};
  (paquete.items || []).forEach(item => {
    const mat = paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID] ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";
    const noActualizado = (paquete.noActualizadosSeleccionados && paquete.noActualizadosSeleccionados[item.Item_ID]) ||
      !mat || mat.includes("SIN_") || mat === item.Producto_Cod || (mat.length === 13 && mat.startsWith("84"));
    if (!mapa[item.Pedido_ID]) mapa[item.Pedido_ID] = [];
    mapa[item.Pedido_ID].push({
      itemId: item.Item_ID,
      ean: item.Producto_Cod,
      descripcion: item.Descripcion,
      material: noActualizado ? "" : mat,
      cantidad: parseInt(item.Cant, 10) || 0,
      precio: Number(item.Valor_Unitario) || 0,
      backOrder: !noActualizado && !!(paquete.backordersSeleccionados && paquete.backordersSeleccionados[item.Item_ID])
    });
  });
  return mapa;
}

/** Consecutivo más alto con ese prefijo en la columna A de CONSOLIDADO (0 si no hay). */
function falabella_maxConsecutivo_(sh, prefijo) {
  const ultima = sh.getLastRow();
  if (ultima < 2) return 0;
  let max = 0;
  sh.getRange(1, 1, ultima, 1).getValues().forEach(f => {
    const id = String(f[0]).trim();
    if (id.startsWith(prefijo)) max = Math.max(max, parseInt(id.slice(prefijo.length), 10) || 0);
  });
  return max;
}

/**
 * Mueve cada Excel a la carpeta de procesados y lo renombra con sus consecutivos
 * ("Pedido_<primero>_a_<último>.xlsx" si el archivo tenía varios locales).
 * @return {Array<String>} Avisos de los archivos que no se pudieron mover o renombrar.
 */
function falabella_moverArchivos_(pedidos, idPorPedido, configCli) {
  const avisos = [];
  const porArchivo = {};
  pedidos.forEach(p => {
    if (!p.Archivo_Id) return;
    if (!porArchivo[p.Archivo_Id]) porArchivo[p.Archivo_Id] = [];
    porArchivo[p.Archivo_Id].push(idPorPedido[p.Pedido_ID]);
  });
  const carpeta = DriveApp.getFolderById(configCli.PROCESSED_FOLDER_ID);
  Object.keys(porArchivo).forEach(archivoId => {
    const ids = porArchivo[archivoId];
    try {
      const file = DriveApp.getFileById(archivoId);
      const original = file.getName();
      const punto = original.lastIndexOf(".");
      const ext = punto !== -1 ? original.substring(punto) : ".xlsx";
      const nuevoNombre = nombreArchivoSap_(ids.length === 1 ? `Pedido_${ids[0]}${ext}` : `Pedido_${ids[0]}_a_${ids[ids.length - 1]}${ext}`);
      try {
        Drive.Files.patch({ title: nuevoNombre }, archivoId, { addParents: configCli.PROCESSED_FOLDER_ID, removeParents: file.getParents().next().getId(), supportsAllDrives: true });
      } catch (eApi) {
        file.moveTo(carpeta);
        file.setName(nuevoNombre);
      }
    } catch (e) {
      avisos.push(`archivo ${archivoId}: ${e.message}`);
    }
  });
  return avisos;
}

/**
 * Correo a la KAM con el resumen por OC: cuántos locales, unidades, valor a facturar y back order, y el
 * detalle por producto sumando todos los locales. Si el envío tiene más de
 * FALABELLA_MAX_POSICIONES_DETALLE_CORREO productos en total, cada OC muestra solo su resumen.
 * Los productos no actualizados se listan siempre.
 */
function falabella_enviarCorreoKam_(paquete, lineasPorPedido, idPorPedido, kam, configCli) {
  if (!kam.correo) throw new Error("La KAM de Falabella no tiene correo en Asignacion_KAM.");
  const pesos = n => "$" + Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");

  // Agrupar los pedidos (locales) por OC
  const porOc = {};
  paquete.pedidos.forEach(p => {
    if (!porOc[p.OC]) porOc[p.OC] = { pedidos: [], productos: {} };
    porOc[p.OC].pedidos.push(p);
    (lineasPorPedido[p.Pedido_ID] || []).forEach(l => {
      const prod = porOc[p.OC].productos[l.ean] || (porOc[p.OC].productos[l.ean] = { descripcion: l.descripcion, material: l.material, cantidad: 0, valor: 0, backOrder: false });
      prod.cantidad += l.cantidad;
      if (l.material) prod.valor += l.cantidad * l.precio;
      prod.backOrder = prod.backOrder || l.backOrder;
    });
  });
  const totalProductos = Object.values(porOc).reduce((s, o) => s + Object.keys(o.productos).length, 0);
  const conDetalle = totalProductos <= FALABELLA_MAX_POSICIONES_DETALLE_CORREO;

  let html = `
    <div style="font-family: 'Inter', sans-serif; color: #2B2830; max-width: 640px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px;">
      <div style="text-align: center; border-bottom: 2px solid #AAD500; padding-bottom: 15px; margin-bottom: 20px;">
        <h1 style="color: #2B2830; font-size: 24px; margin: 0; font-weight: 900; letter-spacing: 2px;">ISDIN</h1>
        <p style="color: #5F7A00; font-size: 11px; font-weight: 700; margin: 5px 0 0; text-transform: uppercase; letter-spacing: 1px;">Notificación Automática de Órdenes B2B - Falabella</p>
      </div>
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Hola <strong>${kam.nombre}</strong>,</p>
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Se registraron los pedidos de <strong>${configCli.NOMBRE_CLIENTE}</strong>, uno por cada local.</p>`;

  Object.keys(porOc).forEach(oc => {
    const grupo = porOc[oc];
    const productos = Object.keys(grupo.productos).map(ean => Object.assign({ ean: ean }, grupo.productos[ean]));
    const noAct = productos.filter(x => !x.material);
    const ok = productos.filter(x => x.material && !x.backOrder).reduce((s, x) => s + x.valor, 0);
    const bo = productos.filter(x => x.material && x.backOrder).reduce((s, x) => s + x.valor, 0);
    const unidades = productos.reduce((s, x) => s + x.cantidad, 0);
    const ids = grupo.pedidos.map(p => idPorPedido[p.Pedido_ID]);
    const locales = grupo.pedidos.map(p => p.Local_Nombre || p.Tipo_Orden || "Sin local").join(", ");

    html += `
      <div style="background-color: #FAFAFA; border: 1px solid #F3F4F6; border-radius: 8px; padding: 15px; margin: 15px 0;">
        <div style="font-weight: bold; font-size: 14px; color: #111827; border-bottom: 1px solid #E5E7EB; padding-bottom: 8px; margin-bottom: 8px;">
          OC: ${oc} · ${grupo.pedidos.length} pedido(s) · Consecutivos ${ids[0]}${ids.length > 1 ? " a " + ids[ids.length - 1] : ""}
        </div>
        <p style="font-size: 12px; color: #4B5563; margin: 0 0 8px;"><strong>Locales:</strong> ${locales}</p>`;

    if (noAct.length > 0) {
      html += `<div style="background-color: #FEF2F2; border: 1px solid #F87171; border-radius: 8px; padding: 10px; margin-bottom: 10px; font-size: 11px; color: #7F1D1D;">
        <strong>⚠️ Productos no actualizados en la base de datos (van sin material al CSV):</strong>
        ${noAct.map(x => `<div>${x.descripcion} (EAN: ${x.ean}) · ${x.cantidad} ud.</div>`).join("")}
      </div>`;
    }

    if (conDetalle) {
      productos.filter(x => x.material).forEach(x => {
        const badge = x.backOrder
          ? '<span style="background-color: #FEF3C7; color: #D97706; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">Back Order</span>'
          : '<span style="background-color: #D1FAE5; color: #059669; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">OK</span>';
        html += `<div style="padding: 6px 0; border-bottom: 1px dashed #F3F4F6; font-size: 12px;">
          <strong>${x.descripcion}</strong> ${badge}<br>
          <span style="color: #6B7280; font-family: monospace; font-size: 11px;">EAN: ${x.ean} | Material SAP: <b>${x.material}</b> | ${x.cantidad} ud. | ${pesos(x.valor)}</span>
        </div>`;
      });
    } else {
      html += `<p style="font-size: 12px; color: #4B5563;">${productos.length} productos, ${unidades} unidades en total. El detalle por posición está en el CSV guardado en Drive.</p>`;
    }

    html += `<div style="text-align: right; font-size: 13px; font-weight: bold; margin-top: 10px;">
          ${unidades} ud. · Valor estimado a facturar (OK): ${pesos(ok)}${bo > 0 ? " · En back order: " + pesos(bo) : ""}
        </div>
      </div>`;
  });

  html += `<p style="font-size: 12px; color: #9CA3AF; margin-top: 20px; text-align: center;">Correo automático del Order Hub de ISDIN Colombia. Por favor no responder.</p></div>`;

  enviarCorreoKam_({
    to: kam.correo,
    subject: `📦 Resumen Pedidos FALABELLA - ${Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")}`,
    htmlBody: html
  });
}

/* ------------------------------- utilidades de texto y fecha ------------------------------- */

/** Mayúsculas sin tildes, espacios ni signos. "Número OC" -> "NUMEROOC". */
function falabella_normalizar_(texto) {
  if (texto === null || texto === undefined) return "";
  return String(texto).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-zA-Z0-9]/g, "").toUpperCase();
}

/** Código numérico como texto (OC, UPC, ID de local o lugar): quita apóstrofe, decimales ".0" y notación científica. */
function falabella_limpiarCodigo_(valor) {
  if (valor === null || valor === undefined) return "";
  if (typeof valor === "number") return Number.isInteger(valor) ? String(valor) : String(Math.round(valor));
  let s = String(valor).trim().replace(/^'/, "");
  if (/^\d+(\.\d+)?e\+?\d+$/i.test(s)) s = Number(s).toFixed(0);
  if (/^\d+\.0+$/.test(s)) s = s.split(".")[0];
  return s;
}

/** Número de una celda que puede venir como texto con separadores de miles ("86.800" o "86800"). */
function falabella_numero_(valor) {
  if (typeof valor === "number") return valor;
  const s = String(valor || "").trim().replace(/\./g, "").replace(",", ".");
  return Number(s) || 0;
}

/**
 * Fecha del B2B a "yyyy-MM-dd" (formato del campo de fecha del asistente). Acepta objetos Date (si la
 * copia en Sheets convirtió la celda) y textos "dd-MM-yyyy" o "dd/MM/yyyy". Devuelve "" si no la reconoce.
 */
function falabella_fechaIso_(valor) {
  if (valor instanceof Date && !isNaN(valor)) return Utilities.formatDate(valor, "GMT-5", "yyyy-MM-dd");
  const m = String(valor || "").trim().match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (!m) return "";
  return `${m[3]}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`;
}

/** "yyyy-MM-dd" -> "dd/MM/yyyy" (formato de fechas de CONSOLIDADO). Devuelve el texto tal cual si no es ISO. */
function falabella_isoADdMmYyyy_(iso) {
  const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? `${m[3]}/${m[2]}/${m[1]}` : String(iso || "");
}
