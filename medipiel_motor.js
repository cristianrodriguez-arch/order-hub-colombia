/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: INTEGRACIÓN DE ÓRDENES DE COMPRA - MEDIPIEL S.A.S. (PDF)
 * ARCHIVO: medipiel_motor.js
 * =========================================================================================
 * QUÉ CONTIENE:
 * Todo el proceso de las órdenes de compra (OC) de MEDIPIEL que llegan como PDF a la carpeta
 * de entrada de Google Drive. El proceso tiene dos pasos, que el asistente ejecuta en orden:
 *
 *   PASO 1 - Análisis (solo lectura): preAnalizarMEDIPIEL()
 *     Lee los PDF, saca la OC, la bodega, las fechas y los productos, y omite las OC que ya
 *     existen en la hoja CONSOLIDADO. No escribe nada.
 *
 *   PASO 2 - Guardado: guardarDefinitivoMEDIPIEL(datosFormulario)
 *     Con lo que confirmó el usuario:
 *       a) asigna el consecutivo de cada OC (formato <agente>-<año><número de 4 dígitos>);
 *       b) genera el CSV de carga a SAP y lo guarda en la carpeta de CSV de Drive;
 *       c) escribe una fila por OC en la hoja CONSOLIDADO (columnas A a K);
 *       d) mueve los PDF a la carpeta de procesados y les cambia el nombre;
 *       e) envía un correo de resumen a la KAM responsable del cliente.
 *
 * QUIÉN LO LLAMA:
 *   - wizard_ui.html: el asistente llama a preAnalizarMEDIPIEL y a guardarDefinitivoMEDIPIEL
 *     por su nombre (google.script.run), según la configuración MEDIPIEL de ese archivo.
 *   - main.js: NO llama a este archivo. Es al revés: main.js aporta lo que este archivo usa:
 *     CONFIG (con CONFIG.CLIENTES.MEDIPIEL), CONSOLIDADO_HEADERS, obtenerInfoKamPorSolicitante_,
 *     obtenerArchivosCliente_, obtenerDiccionarioDestinatarios, obtenerMapaBodegaGeneral_,
 *     construirCsvSap_, nombreArchivoSap_ y enviarCorreoKam_.
 *
 * MODO PRUEBA: si CONFIG.PRUEBA.ACTIVO es true (main.js), los nombres de archivo llevan el prefijo
 * PRUEBA_ y los correos van a CONFIG.PRUEBA.CORREO_DESTINO (o a la KAM, con [PRUEBA] en el asunto).
 *
 * HOJAS Y CARPETAS QUE USA:
 *   Hojas: CONSOLIDADO (escribe), Bodega, Productos, Destinatarios y Listado clientes (lee),
 *   y Asignacion_KAM (lee, a través de main.js).
 *   Carpetas de Drive: entrada de PDF de MEDIPIEL, procesados de MEDIPIEL y CSV (FOLDER_CSV_ID).
 * =========================================================================================
 */

/**
 * PASO 1 DEL PROCESO (solo lectura). Busca los PDF de órdenes de MEDIPIEL en la carpeta de
 * entrada, descarta las OC que ya están en CONSOLIDADO y devuelve lo encontrado para que el
 * usuario lo revise en el asistente. No escribe en hojas, carpetas ni correos.
 *
 * Lee: hojas Bodega, Productos, Destinatarios y CONSOLIDADO; carpeta de entrada de MEDIPIEL en Drive.
 * Escribe: nada.
 *
 * @return {Object} { pedidosExtraidos, itemsExtraidos, errores, yaProcesados } (ver medipiel_motorLectura).
 *   Si no hay PDF pendientes devuelve solo { pedidosExtraidos, itemsExtraidos, yaProcesados }, sin la clave "errores".
 * @throws {Error} Si MEDIPIEL no tiene FOLDER_ID en CONFIG.CLIENTES (main.js), o si la lectura
 *   encuentra una bodega no registrada (MISSING_BODEGA) o un error fatal de bodega (FATAL_BODEGA).
 */
function preAnalizarMEDIPIEL() {
  // 1. Leer la configuración de MEDIPIEL desde CONFIG (main.js) y verificar que tenga carpeta de entrada.
  const cliente = 'MEDIPIEL';
  const configCli = CONFIG.CLIENTES[cliente];
  if (!configCli || !configCli.FOLDER_ID) {
    throw new Error('La configuración de MEDIPIEL no está definida en CONFIG de main.js');
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 2. Cargar los catálogos: bodega (EAN a material), productos (EAN a material) y destinatarios.
  const shBodega = ss.getSheetByName(CONFIG.SHEET_BODEGA || "Bodega");
  const diccBodega = shBodega ? obtenerMapaBodegaGeneral_(shBodega) : {};

  const shProductos = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  const diccProductos = shProductos ? medipiel_getProductosMap_(shProductos) : {};

  const diccDestinatarios = obtenerDiccionarioDestinatarios();

  // 3. Antiduplicados: preparar la lista de OC ya registradas en CONSOLIDADO (clave OC_SAP).
  // Antiduplicados: Cargar OCs procesadas previamente en el Consolidado con clave compuesta
  let pedidosProcesados = new Set();
  const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  
  if (shConsolidado && shConsolidado.getLastRow() >= 3) {
    const mapConsolidado = medipiel_getHeaderMap_(shConsolidado, 3);
    const colOC = mapConsolidado[medipiel_normalizeKey_("Orden De Compra")];
    const colSap = mapConsolidado[medipiel_normalizeKey_("Sap Id")];
    
    if (colOC) {
      const dataRange = shConsolidado.getRange(3, 1, shConsolidado.getLastRow() - 2, Math.max(shConsolidado.getLastColumn(), 26)).getValues();
      dataRange.forEach(r => {
        const oc = medipiel_limpiarOc_(r[colOC - 1]);
        const sap = colSap ? medipiel_normalizeKey_(r[colSap - 1]) : medipiel_normalizeKey_(configCli.SOLICITANTE);
        if (oc && sap) {
          pedidosProcesados.add(`${oc}_${sap}`);
        }
      });
    }
  }

  // 4. Listar los PDF de la carpeta de entrada de MEDIPIEL, hasta el límite configurado (200 si no hay límite).
  const folder = DriveApp.getFolderById(configCli.FOLDER_ID);
  const fileIterator = obtenerArchivosCliente_(folder, configCli);
  const archivos = [];
  const maxLimit = CONFIG.LIMIT_FILES > 0 ? CONFIG.LIMIT_FILES : 200;

  while (fileIterator.hasNext() && archivos.length < maxLimit) {
    const file = fileIterator.next();
    archivos.push({
      id: file.getId(),
      name: file.getName(),
      title: file.getName()
    });
  }

  // 5. Si no hay PDF pendientes, devolver listas vacías (sin error).
  if (!archivos.length) {
    return { pedidosExtraidos: [], itemsExtraidos: [], yaProcesados: [] };
  }

  // 6. Leer y analizar cada PDF (ver medipiel_motorLectura).
  return medipiel_motorLectura(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios);
}

/**
 * PASO 2 DEL PROCESO (escritura). Se ejecuta cuando el usuario confirma en el asistente.
 * Asigna el consecutivo a cada OC, genera el CSV de carga a SAP, guarda el CSV en Drive,
 * escribe las filas en CONSOLIDADO, mueve los PDF a procesados y envía el correo a la KAM.
 *
 * Lee: hoja CONSOLIDADO (para calcular el siguiente consecutivo), hoja "Listado clientes"
 *   (ciudades) y Asignacion_KAM (a través de obtenerInfoKamPorSolicitante_).
 * Escribe:
 *   - Carpeta CONFIG.FOLDER_CSV_ID de Drive: un archivo CSV nuevo.
 *   - Hoja CONSOLIDADO: filas nuevas al final, solo columnas A a K. La columna KAM (N) no se escribe aquí.
 *   - Carpeta de procesados de MEDIPIEL: mueve y renombra los PDF de las OC validas.
 *   - Correo: un resumen a la KAM (solo si la KAM tiene un correo con "@").
 *
 * @param {Object} datosFormulario Lo que arma el asistente, con esta forma:
 *   - pedidos: lista de OC { Pedido_ID, Destinatario_Merc, Destinatario, Pedido_Fecha, Ciudad, Archivo_ID }.
 *   - items: lista de productos { Item_ID, Pedido_ID, Producto_Cod, Descripcion, Cant, Valor_Unitario }.
 *   - fechasEntregas: { [Pedido_ID]: { fecha: "yyyy-MM-dd" } }, la fecha de entrega elegida.
 *   - materialesSeleccionados: { [Item_ID]: codigo de material SAP elegido }.
 *   - noActualizadosSeleccionados: { [Item_ID]: true } si el usuario marco el producto como no actualizado.
 *   - backordersSeleccionados: { [Item_ID]: true } si el usuario marco el producto como back order.
 * @return {Object} { mensaje: String, csvData: String (contenido del CSV en base64), fileName: String (nombre del CSV) }.
 * @throws {Error} En estos casos:
 *   - "No hay productos válidos seleccionados. Proceso abortado.": no hay ninguna linea que enviar. No escribe nada.
 *   - "Falta configurar FOLDER_CSV_ID..." o "Fallo al guardar archivo CSV...": el CSV no se pudo guardar. No escribe en CONSOLIDADO.
 *   - "Fallo al escribir en la Base de Datos...": CONSOLIDADO fallo. El CSV recien creado se manda a la papelera.
 *   Los errores al mover los PDF y al enviar el correo NO detienen el proceso: solo quedan en el registro (console.error).
 */
/**
 * Guarda las órdenes de MEDIPIEL confirmadas en el asistente. Va dentro del bloqueo del script y
 * revisa antes de escribir que ninguna OC ya esté en CONSOLIDADO. (La lógica está en la función de abajo.)
 *
 * @param {Object} datosFormulario Lo que envía el asistente (pedidos, items, fechas, materiales, back orders).
 * @return {Object} { mensaje, csvData, fileName } para el asistente.
 * @throws Si alguna OC ya está registrada, o si falla la escritura: en ese caso no se registra nada.
 */
function guardarDefinitivoMEDIPIEL(datosFormulario) {
  return conBloqueoScript_(() => medipiel_guardarDefinitivoSinBloqueo_(datosFormulario));
}

/**
 * Revisa que las OC de la confirmación no estén ya en CONSOLIDADO para este cliente.
 * Lee la hoja (no escribe). Si alguna ya está, lanza error con la lista.
 *
 * @param {SpreadsheetApp.Spreadsheet} ss Libro de cálculo.
 * @param {Array<String>} ocs Números de OC que se van a guardar.
 * @param {String|Number} solicitante SAP ID del cliente (MEDIPIEL).
 */
function medipiel_verificarNoRegistradas_(ss, ocs, solicitante) {
  const sh = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  if (!sh || sh.getLastRow() < 4) return;
  const map = medipiel_getHeaderMap_(sh, 3);
  const colOC = map[medipiel_normalizeKey_("Orden De Compra")];
  const colSap = map[medipiel_normalizeKey_("Sap Id")];
  if (!colOC || !colSap) return;
  const filas = sh.getRange(4, 1, sh.getLastRow() - 3, Math.max(colOC, colSap)).getValues();
  const registradas = new Set();
  filas.forEach(f => {
    const oc = medipiel_limpiarOc_(f[colOC - 1]);
    if (oc) registradas.add(oc + "_" + String(f[colSap - 1]).trim());
  });
  const repetidas = ocs.filter(oc => registradas.has(medipiel_limpiarOc_(oc) + "_" + String(solicitante).trim()));
  if (repetidas.length > 0) {
    throw new Error("Ya están registradas en CONSOLIDADO, por eso no se guardó nada: " + repetidas.join(", ") + ". Quítelas de la confirmación.");
  }
}

/**
 * Lógica real de guardado de MEDIPIEL (se ejecuta dentro del bloqueo). Antes de escribir revisa duplicados.
 * @param {Object} datosFormulario Lo que envía el asistente.
 * @return {Object} { mensaje, csvData, fileName }.
 */
function medipiel_guardarDefinitivoSinBloqueo_(datosFormulario) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const configCli = CONFIG.CLIENTES["MEDIPIEL"];
  // Revisión final de duplicados antes de escribir nada
  medipiel_verificarNoRegistradas_(ss, datosFormulario.pedidos.map(p => p.Pedido_ID), configCli.SOLICITANTE);

  // 1. Obtención 100% DESACOPLADA de la información del KAM desde main.js
  const infoKam = typeof obtenerInfoKamPorSolicitante_ === "function" 
    ? obtenerInfoKamPorSolicitante_(configCli.SOLICITANTE) 
    : null;

  let codigoAgente = (infoKam && infoKam.iniciales) ? infoKam.iniciales : "SC";
  let kamEncargado = (infoKam && infoKam.nombre) ? infoKam.nombre : "KAM ENCARGADO";
  let correoKam = (infoKam && infoKam.correo) ? infoKam.correo : "";

  // 3. Preparar la hoja CONSOLIDADO
  let shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  if (!shConsolidado) {
    shConsolidado = ss.insertSheet(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    shConsolidado.getRange(3, 1, 1, CONSOLIDADO_HEADERS.length).setValues([CONSOLIDADO_HEADERS]);
  } else if (shConsolidado.getLastRow() < 3) {
    shConsolidado.getRange(3, 1, 1, CONSOLIDADO_HEADERS.length).setValues([CONSOLIDADO_HEADERS]);
  }
  
  const headerMapConsolidado = medipiel_getHeaderMap_(shConsolidado, 3);
  // Solo se escriben las columnas A:K. Desde la L en adelante hay fórmulas y formatos condicionales.
  const numColsConsolidado = 11;

  // 4. Consecutivo dinámico secuencial por agente (<CódigoAgente>-<Año><Consecutivo>)
  const currentYearSuffix = String(new Date().getFullYear()).slice(-2);
  let maxConsecutivo = 0;
  
  const lastRowConsolidado = shConsolidado.getLastRow();
  if (lastRowConsolidado >= 3) {
    const dataRange = shConsolidado.getRange(3, 1, lastRowConsolidado - 2, numColsConsolidado).getValues();
    const colIdIndex = headerMapConsolidado[medipiel_normalizeKey_("Id")] ? headerMapConsolidado[medipiel_normalizeKey_("Id")] - 1 : 0;
    
    const regexAgenteAno = new RegExp(`^${codigoAgente}-${currentYearSuffix}(\\d+)$`, 'i');
    
    for (let i = 0; i < dataRange.length; i++) {
      const idVal = String(dataRange[i][colIdIndex] || "").trim();
      const match = idVal.match(regexAgenteAno);
      
      if (match) {
        const numConsecutivo = parseInt(match[1], 10);
        if (numConsecutivo > maxConsecutivo) {
          maxConsecutivo = numConsecutivo;
        }
      }
    }
  }

  // 5. Mapeo de datos para CSV y CONSOLIDADO
  const pedidosCsv = [];
  let filasConsolidado = [];
  const pedidosValidos = new Set();
  let primerConsecutivoGenerado = null; 
  const mapaNombresArchivos = {}; 

  // Paso 5 (continuación): recorrer cada OC del formulario y calcular sus totales y su línea de CSV.
  datosFormulario.pedidos.forEach(p => {
    let sumaUnidadesPedido = 0;
    let sumaValorPedido = 0;

    let fechaUI = datosFormulario.fechasEntregas[p.Pedido_ID] ? datosFormulario.fechasEntregas[p.Pedido_ID].fecha : "";
    let fechaEntregaNormalObj = "";

    if (fechaUI) {
      let parts = fechaUI.split("-");
      if (parts.length === 3) {
        fechaEntregaNormalObj = new Date(parts[0], parts[1] - 1, parts[2]);
      }
    }

    const pedidoCsv = {
      sap: configCli.SAP,
      solicitante: configCli.SOLICITANTE,
      destinatario: p.Destinatario_Merc,
      numPedidoCliente: p.Pedido_ID,
      fechaEntrega: fechaUI,
      lineas: []
    };
    pedidosCsv.push(pedidoCsv);

    const items = datosFormulario.items.filter(i => i.Pedido_ID === p.Pedido_ID);

    items.forEach(i => {
      let materialFinal = (datosFormulario.materialesSeleccionados && datosFormulario.materialesSeleccionados[i.Item_ID]) 
                            ? String(datosFormulario.materialesSeleccionados[i.Item_ID]).trim() : "";
      
      // CORRECCIÓN: Permite códigos de hasta 12 dígitos oficiales (descarta vacíos, SIN_ o EANs de 13 dígitos que empiezan por 84)
      const isNoActualizado = (datosFormulario.noActualizadosSeleccionados && datosFormulario.noActualizadosSeleccionados[i.Item_ID]) || 
                              (!materialFinal || materialFinal.includes("SIN_") || materialFinal === i.Producto_Cod || (materialFinal.length === 13 && materialFinal.startsWith("84")));

      let matCsv = isNoActualizado ? "" : materialFinal;

      pedidosValidos.add(p.Pedido_ID);
      sumaUnidadesPedido += parseInt(i.Cant) || 0;
      sumaValorPedido += (Number(i.Valor_Unitario) || 0) * (Number(i.Cant) || 0);

      // Lleva el Material SAP (incluyendo códigos de 12 dígitos) o vacío si no cruzó
      pedidoCsv.lineas.push({ material: matCsv, cantidad: i.Cant });
    });

    // Solo las OC con al menos una línea reciben consecutivo y fila en CONSOLIDADO.
    if (pedidosValidos.has(p.Pedido_ID)) {
      maxConsecutivo++;
      const consecutivoStr = String(maxConsecutivo).padStart(4, '0');
      const idGenerado = `${codigoAgente}-${currentYearSuffix}${consecutivoStr}`;
      
      const ciudadDestino = p.Ciudad; // siempre viene con valor: si no, el pedido no se leyó

      if (!primerConsecutivoGenerado) {
        primerConsecutivoGenerado = idGenerado;
      }

      if (p.Archivo_ID) {
        mapaNombresArchivos[p.Archivo_ID] = nombreArchivoSap_(`Pedido_${idGenerado}.pdf`);
      }

      filasConsolidado.push(medipiel_createRowArray_(headerMapConsolidado, {
        "Id": idGenerado,
        "Orden De Compra": p.Pedido_ID,
        "Sap Id": configCli.SOLICITANTE,
        "Solicitante": configCli.NOMBRE_CLIENTE,
        "Fecha Pedido": medipiel_parseDateEs_(p.Pedido_Fecha) || "",
        "Id Destino": p.Destinatario_Merc,
        "Destinatario": p.Destinatario,
        "Ciudad": ciudadDestino,
        "Fecha entrega": fechaEntregaNormalObj,
        "Unidades": sumaUnidadesPedido,
        "Valor": sumaValorPedido
      }, numColsConsolidado));
    }
  });

  // Verificación: si ninguna OC tiene líneas, no se genera nada.
  if (!pedidosCsv.some(pc => pc.lineas.length > 0)) {
     throw new Error("No hay productos válidos seleccionados. Proceso abortado.");
  }

  // Nombre del CSV: Pedido_<primer consecutivo>.csv (con prefijo PRUEBA_ en modo prueba).
  let csvFileName = nombreArchivoSap_(`Pedido_${primerConsecutivoGenerado || Utilities.formatDate(new Date(), "GMT-5", "yyyyMMdd_HHmm")}.csv`);
  const csvString = construirCsvSap_(pedidosCsv);

  // 6. Guardar el archivo CSV en Google Drive ANTES de CONSOLIDADO: si falla, la OC no queda
  //    marcada como procesada y se puede reintentar.
  let archivoCsv = null;
  try {
    if (CONFIG.FOLDER_CSV_ID) {
      const folderCSV = DriveApp.getFolderById(CONFIG.FOLDER_CSV_ID);
      archivoCsv = folderCSV.createFile(csvFileName, csvString, MimeType.CSV);
    } else {
      throw new Error("Falta configurar FOLDER_CSV_ID en CONFIG.");
    }
  } catch (errCSV) {
    throw new Error("Fallo al guardar archivo CSV en Drive (no se registró nada en CONSOLIDADO): " + errCSV.message);
  }

  // 7. Escribir en la hoja CONSOLIDADO. Si falla, se retira el CSV para no dejar un archivo huérfano.
  try {
    if (filasConsolidado.length > 0) {
      shConsolidado.getRange(shConsolidado.getLastRow() + 1, 1, filasConsolidado.length, numColsConsolidado).setValues(filasConsolidado);
      SpreadsheetApp.flush();
    }
  } catch (errSheets) {
    try { archivoCsv.setTrashed(true); } catch (ignore) {}
    throw new Error("Fallo al escribir en la Base de Datos (Consolidado); el CSV se envió a la papelera: " + errSheets.message);
  }

  // 8. Traslado y renombrado de PDFs procesados en Google Drive
  const folderDestinoId = configCli.PROCESSED_FOLDER_ID;
  const folderOrigenId = configCli.FOLDER_ID;
  const archivosAMover = new Set();
  
  datosFormulario.pedidos.forEach(p => { 
    if (pedidosValidos.has(p.Pedido_ID) && p.Archivo_ID) {
      archivosAMover.add(p.Archivo_ID); 
    }
  });

  // Por cada PDF válido: moverlo a procesados y renombrarlo.
  archivosAMover.forEach(fileId => {
    const nuevoNombre = mapaNombresArchivos[fileId];
    try {
      // Intento 1: API avanzada de Drive (cambia nombre y carpeta en una sola llamada).
      Drive.Files.patch(nuevoNombre ? { title: nuevoNombre } : {}, fileId, { addParents: folderDestinoId, removeParents: folderOrigenId, supportsAllDrives: true });
    } catch (e1) {
      try {
        const file = DriveApp.getFileById(fileId);
        if (nuevoNombre) file.setName(nuevoNombre);
        file.moveTo(DriveApp.getFolderById(folderDestinoId));
      } catch (e2) {
        console.error(`Error definitivo al mover el archivo ${fileId}: ${e2.message}`);
      }
    }
  });

  // 9. Enviar notificación por correo a la KAM encargada
  if (correoKam && correoKam.includes("@")) {
    try {
      let correoHTML = `
        <div style="font-family: 'Inter', sans-serif; color: #2B2830; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
          <div style="text-align: center; border-bottom: 2px solid #E41F33; padding-bottom: 15px; margin-bottom: 20px;">
            <h1 style="color: #2B2830; font-size: 24px; margin: 0; font-weight: 900; letter-spacing: 2px;">ISDIN</h1>
            <p style="color: #E41F33; font-size: 11px; font-weight: 700; margin: 5px 0 0; text-transform: uppercase; letter-spacing: 1px;">Notificación Automática de Órdenes B2B - MEDIPIEL</p>
          </div>
          
          <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Estimada <strong>${kamEncargado}</strong>,</p>
          <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Te informamos que se han consolidado de manera exitosa los nuevos pedidos de <strong>MEDIPIEL S.A.S.</strong>. El consecutivo asignado es el <span style="font-family: monospace; font-weight: bold; background-color: #F3F4F6; padding: 2px 6px; border-radius: 4px; border: 1px solid #E5E7EB;">Pedido_${primerConsecutivoGenerado}</span>.</p>
          
          <h2 style="font-size: 16px; font-weight: 700; border-bottom: 1px solid #F3F4F6; padding-bottom: 8px; margin-top: 25px; color: #2B2830;">Detalle de Facturación por Orden de Compra</h2>
          <div style="margin-top: 15px;">
      `;
      
      // Armar el HTML del correo: por cada OC, los productos válidos, los back orders y los no actualizados.
      datosFormulario.pedidos.forEach(p => {
        const itemsPedido = datosFormulario.items.filter(i => i.Pedido_ID === p.Pedido_ID);
        let itemsNoActualizados = [];
        let itemsValidos = [];
        let totalValorOk = 0;
        let totalValorBO = 0;
        let hasBackOrder = false;

        itemsPedido.forEach(item => {
          let mat = (datosFormulario.materialesSeleccionados && datosFormulario.materialesSeleccionados[item.Item_ID]) ? String(datosFormulario.materialesSeleccionados[item.Item_ID]).trim() : "";
          
          // CORRECCIÓN: Permite materiales de 12 dígitos sin enviarlos erróneamente a No Actualizados
          let isNoActualizado = (datosFormulario.noActualizadosSeleccionados && datosFormulario.noActualizadosSeleccionados[item.Item_ID]) || 
                                (!mat || mat.includes("SIN_") || mat === item.Producto_Cod || (mat.length === 13 && mat.startsWith("84")));
          let isBO = !isNoActualizado && (datosFormulario.backordersSeleccionados && datosFormulario.backordersSeleccionados[item.Item_ID]);

          const cant = parseInt(item.Cant) || 0;
          const price = parseFloat(item.Valor_Unitario) || 0;
          const sub = cant * price;

          if (isNoActualizado) {
            itemsNoActualizados.push({
              desc: item.Descripcion || "Referencia no cruzada",
              ean: item.Producto_Cod,
              cant: cant
            });
          } else {
            if (isBO) {
              hasBackOrder = true;
              totalValorBO += sub;
            } else {
              totalValorOk += sub;
            }
            itemsValidos.push({
              desc: item.Descripcion,
              ean: item.Producto_Cod,
              cant: cant,
              subtotal: sub,
              isBO: isBO,
              material: mat
            });
          }
        });

        let alertaNoActualizadosHtml = "";
        if (itemsNoActualizados.length > 0) {
          alertaNoActualizadosHtml = `
            <div style="background-color: #FEF2F2; border: 1px solid #F87171; border-radius: 8px; padding: 12px 15px; margin-bottom: 12px;">
              <div style="color: #991B1B; font-weight: bold; font-size: 12px; margin-bottom: 4px;">
                ⚠️ Estos productos no se encuentran actualizados en la base de datos:
              </div>
              <p style="font-size: 11px; color: #7F1D1D; margin: 0 0 8px 0; line-height: 1.4;">
                El cliente solicitó las siguientes referencias en la OC pero su código EAN no cruzó con el portafolio activo
              </p>
              <div style="border-top: 1px dashed #FCA5A5; padding-top: 6px;">
          `;

          itemsNoActualizados.forEach(it => {
            alertaNoActualizadosHtml += `
              <div style="padding: 4px 0; display: flex; justify-content: space-between; font-size: 11px; color: #7F1D1D;">
                <div>
                  <strong>${it.desc}</strong> <span style="font-family: monospace;">(EAN: ${it.ean})</span>
                </div>
                <div style="font-weight: bold;">${it.cant} ud.</div>
              </div>
            `;
          });

          alertaNoActualizadosHtml += `</div></div>`;
        }

        let listHtml = "";
        itemsValidos.forEach(item => {
          let badge = item.isBO
            ? '<span style="background-color: #FEF3C7; color: #D97706; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">⚠️ Back Order</span>'
            : '<span style="background-color: #D1FAE5; color: #059669; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">✅ OK</span>';

          const formattedSub = "$" + Math.round(item.subtotal).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
          const subtotalText = item.subtotal > 0 ? ` | Subtotal: <strong>${formattedSub}</strong>` : "";

          listHtml += `
            <div style="padding: 10px 0; border-bottom: 1px dashed #F3F4F6; display: flex; justify-content: space-between; font-size: 12px;">
              <div style="flex: 1; min-width: 0; padding-right: 10px;">
                <span style="font-weight: bold; color: #111827; display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${item.desc}</span>
                <span style="color: #6B7280; font-family: monospace; font-size: 11px;">EAN: ${item.ean} | Material SAP: <b>${item.material}</b> | ${item.cant} ud.${subtotalText}</span>
              </div>
              <div style="text-align: right; white-space: nowrap;">
                ${badge}
              </div>
            </div>
          `;
        });

        let statusLabel = '';
        if (itemsNoActualizados.length > 0) {
          statusLabel = '<span style="background-color: #FEE2E2; color: #B91C1C; border: 1px solid #FCA5A5; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">🚨 EANs No Actualizados</span>';
        } else if (hasBackOrder) {
          statusLabel = '<span style="background-color: #FEF3C7; color: #D97706; border: 1px solid #FCD34D; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">⚠️ Contiene Back Orders</span>';
        } else {
          statusLabel = '<span style="background-color: #D1FAE5; color: #059669; border: 1px solid #6EE7B7; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">✅ Completada con Éxito</span>';
        }

        let valueSummaryText = `Valor Estimado Facturar (OK): $${Math.round(totalValorOk).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
        if (totalValorBO > 0) {
          valueSummaryText += ` | Valor en Back Order: $${Math.round(totalValorBO).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
        }
        if (itemsNoActualizados.length > 0) {
          valueSummaryText += ` <br><span style="color:#B91C1C; font-size:11px; margin-top:4px; display:block;"><strong>Alerta Comercial:</strong> Hay ${itemsNoActualizados.length} ítem(s) que requieren actualización de catálogo.</span>`;
        }

        correoHTML += `
          <div style="background-color: #FAFAFA; border: 1px solid #F3F4F6; border-radius: 8px; padding: 15px; margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #E5E7EB; padding-bottom: 10px; margin-bottom: 10px;">
              <span style="font-weight: bold; font-size: 14px; color: #111827;">Orden: ${p.Pedido_ID}</span>
              ${statusLabel}
            </div>
            ${alertaNoActualizadosHtml}
            ${listHtml}
            <div style="text-align: right; font-size: 13px; font-weight: bold; margin-top: 15px; color: #111827;">
              ${valueSummaryText}
            </div>
          </div>
        `;
      });

      correoHTML += `
          </div>
          <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">El archivo CSV (<strong>${csvFileName}</strong>) ha sido guardado automáticamente.</p>
          <p style="font-size: 11px; color: #9CA3AF; margin-top: 25px; border-top: 1px solid #F3F4F6; padding-top: 15px; text-align: center;">Este es un correo automático generado por el Order Hub de ISDIN Colombia. Por favor no responder.</p>
        </div>
      `;

      enviarCorreoKam_({
        to: correoKam,
        subject: `📦 Resumen Pedidos MEDIPIEL - ${Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")}`,
        htmlBody: correoHTML
      });
      
    } catch (eCorreo) {
      console.error("Fallo al enviar correo de notificación: " + eCorreo.message);
    }
  }

  return {
    mensaje: `¡Éxito! Se consolidaron los datos, se guardó el CSV y se notificó al KAM.`,
    csvData: Utilities.base64Encode(csvString).replace(/[\r\n]/g, ""), 
    fileName: csvFileName
  };
}

/**
 * Lee una lista de PDF de MEDIPIEL y, de cada uno, saca la OC, la cabecera y los productos.
 * Separa los resultados en: pedidos listos, productos, PDF con error y OC ya procesadas.
 * Solo lee: no escribe en hojas, carpetas ni correos.
 *
 * @param {Array<Object>} archivos Lista de archivos { id, name, title }: id es el ID del PDF en Drive.
 * @param {Object} diccBodega Mapa EAN -> lista de lotes { material, stock, fechaDisplay } (de obtenerMapaBodegaGeneral_).
 * @param {Object} diccProductos Mapa EAN -> codigo de material (hoja Productos).
 * @param {Object} configCli Configuracion de MEDIPIEL en CONFIG.CLIENTES (main.js).
 * @param {Set<String>} pedidosProcesados Claves "OC_SAP" de las OC que ya estan en CONSOLIDADO.
 * @param {Object} diccDestinatarios Mapa de destinatarios por codigo de bodega (de obtenerDiccionarioDestinatarios).
 * @return {Object} { pedidosExtraidos: Array, itemsExtraidos: Array, errores: Array<String>, yaProcesados: Array<String> }.
 *   errores: un texto "nombre del PDF: motivo" por cada PDF que no se pudo leer.
 *   yaProcesados: las OC que se omitieron porque ya existen (sin repetir).
 * @throws {Error} Relanza los errores de bodega (MISSING_BODEGA o FATAL_BODEGA) y detiene todo el análisis.
 *   Cualquier otro error se anota en "errores" y el proceso sigue con el siguiente PDF.
 */
function medipiel_motorLectura(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios) {
  const pedidosExtraidos = [];
  const itemsExtraidos = [];
  const errores = [];
  const yaProcesados = [];
  const targetSap = medipiel_normalizeKey_(configCli.SOLICITANTE || "11026712");

  archivos.forEach((archivo) => {
    const fileId = archivo.id;
    const fileName = archivo.title || archivo.name || 'archivo.pdf';

    try {
      // 1. Leer el PDF y convertirlo a texto.
      const blob = DriveApp.getFileById(fileId).getBlob();
      const texto = medipiel_pdfToTextPreferNative_(blob, fileName);
      // 2. Sacar la cabecera de la OC (número, bodega, fechas).
      const pedido = medipiel_parsePedidoCabecera_(texto, fileId, fileName, configCli, diccDestinatarios);
      
      // 3. Revisar si esta OC ya está registrada en CONSOLIDADO.
      const ocTrim = medipiel_limpiarOc_(pedido.Pedido_ID);
      const comboKey = `${ocTrim}_${targetSap}`;

      // ESCUDO ANTI-DUPLICADOS: Si la OC ya existe para MEDIPIEL, se omite
      if (pedidosProcesados.has(comboKey)) {
        if (!yaProcesados.includes(ocTrim)) yaProcesados.push(ocTrim);
        return;
      }

      // 4. Sacar la tabla de productos. Si no hay ninguno, el PDF se trata como error.
      const items = medipiel_parseItems_(texto, pedido.Pedido_ID, diccBodega, diccProductos);
      if (!items.length) throw new Error('No se detectaron productos en la tabla del PDF.');

      // 5. Guardar el resultado para el asistente.
      pedidosExtraidos.push(pedido);
      items.forEach(it => itemsExtraidos.push(it));

    } catch (err) {
      // Errores de bodega: se relanzan para detener el análisis y que el asistente los muestre.
      if (err.message.includes('MISSING_BODEGA|')) throw err; 
      if (err.message.includes('FATAL_BODEGA:')) throw new Error(err.message.replace('FATAL_BODEGA:', '')); 
      // Otro error (PDF ilegible, sin productos, etc.): se anota y se sigue con el siguiente PDF.
      errores.push(fileName + ": " + err.message);
    }
  });

  // Resultado final para el asistente.
  return { pedidosExtraidos, itemsExtraidos, errores, yaProcesados };
}

/**
 * Convierte un PDF en texto usando el convertidor de Google Docs (sin servicio externo de OCR).
 * Crea un Google Doc temporal con el contenido del PDF, toma su texto y borra el temporal.
 * Después limpia el texto: cambia espacios especiales por espacios normales, los retornos de
 * carro por saltos de linea y quita espacios repetidos.
 *
 * Escribe: crea y borra un archivo temporal llamado TEMP_MEDIPIEL en Drive.
 *
 * @param {Blob} blob Contenido del PDF (de DriveApp.getFileById(id).getBlob()).
 * @param {String} fileName Nombre del PDF. No se usa dentro de la funcion; se recibe por uniformidad.
 * @return {String} Texto del PDF ya limpio.
 * @throws {Error} Si la API avanzada de Drive no esta activada o la conversion falla. Si falla
 *   despues de crear el temporal, el temporal NO se borra y queda en Drive.
 */
function medipiel_pdfToTextPreferNative_(blob, fileName) {
  // 1. Crear el Google Doc temporal a partir del PDF (conversión de Google).
  const doc = Drive.Files.insert({ title: 'TEMP_MEDIPIEL', mimeType: 'application/vnd.google-apps.document' }, blob, { convert: true });
  // 2. Leer el texto del documento temporal.
  const text = DocumentApp.openById(doc.id).getBody().getText();
  // 3. Borrar el documento temporal.
  Drive.Files.remove(doc.id);
  // 4. Limpiar espacios especiales, saltos de línea y espacios repetidos.
  return String(text || '').replace(/\u00A0/g, ' ').replace(/\r/g, '\n').replace(/[ ]{2,}/g, ' ').trim();
}

/**
 * Saca de la cabecera del PDF los datos de la OC: numero, fechas, bodega de entrega y ciudad.
 * La bodega se cruza con la hoja Destinatarios para obtener el ID de destino y el nombre.
 * El numero de OC se busca primero en el texto del PDF y, si no aparece, en el nombre del archivo.
 * Solo lee: no escribe nada.
 *
 * @param {String} texto Texto del PDF (de medipiel_pdfToTextPreferNative_).
 * @param {String} fileId ID del PDF en Drive.
 * @param {String} fileName Nombre del PDF.
 * @param {Object} configCli Configuracion de MEDIPIEL en CONFIG.CLIENTES (main.js).
 * @param {Object} diccDestinatarios Mapa de destinatarios por codigo de bodega (de obtenerDiccionarioDestinatarios).
 * @return {Object} { Pedido_ID, Destinatario_Merc, Destino_Original, Destinatario, Pedido_Fecha, Fecha_Entrega,
 *   Archivo, Archivo_ID, Archivo_URL, Ciudad }. Las fechas van como texto, tal como aparecen en el PDF.
 * @throws {Error} En estos casos:
 *   - "OC no encontrada en el PDF ni en el nombre del archivo.": no hay numero de OC.
 *   - "FATAL_BODEGA:...": el PDF no trae bodega. Detiene el analisis.
 *   - "MISSING_BODEGA|bodega|SAP|cliente": la bodega no esta en Destinatarios. Detiene el analisis
 *     y el asistente pide registrarla.
 */
function medipiel_parsePedidoCabecera_(texto, fileId, fileName, configCli, diccDestinatarios) {
  const t = texto;
  
  // 1. Buscar el número de OC en el texto del PDF (con varias formas de escribirlo).
  let ordenCompra = medipiel_matchRegex_(t, /(?:No\.?|N°|Nro\.?|Orden(?:\s+de\s+Compra)?(?:\s+No\.?)?|Pedido(?:\s+No\.?)?)\s*[:.]?\s*(\d{4,12})/i);
  
  // 2. Si no apareció en el texto, buscar un número de 4 a 12 dígitos en el nombre del archivo.
  if (!ordenCompra && fileName) {
    const matchFile = fileName.match(/\d{4,12}/);
    if (matchFile) ordenCompra = matchFile[0];
  }

  if (!ordenCompra) {
    ordenCompra = medipiel_matchRegex_(t, /No\s*(\d{4,12})/i);
  }

  // 3. Leer las fechas de la OC y de entrega (se devuelven como texto).
  const fechaPedido = medipiel_matchRegex_(t, /Fecha\s*(?:de\s*)?(?:Orden|Pedido)?\s*:\s*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i);
  const fechaEntrega = medipiel_matchRegex_(t, /Fecha\s*(?:de\s*)?Entrega\s*:\s*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i);
  
  // 4. Leer la bodega de entrega. Si no aparece con el formato normal, intentar otro formato.
  let idDestinoRaw = medipiel_matchRegex_(t, /Bodega\s*(?:de\s*Entrega)?\s*:\s*([A-Z0-9_-]+)/i);
  if (!idDestinoRaw) {
    idDestinoRaw = medipiel_matchRegex_(t, /Bodega\s*:\s*([^\n\r,]+)/i);
    if (idDestinoRaw) idDestinoRaw = idDestinoRaw.trim();
  }

  // 5. Validar que se leyeron la OC y la bodega.
  if (!ordenCompra) throw new Error(`OC no encontrada en el PDF ni en el nombre del archivo.`);
  if (!idDestinoRaw) throw new Error(`FATAL_BODEGA:No se pudo leer la Bodega en la OC ${ordenCompra}.`);

  // 6. Cruzar la bodega del PDF con la hoja Destinatarios (ID de destino, nombre y ciudad).
  const destinoMap = diccDestinatarios[idDestinoRaw];
  
  // Si la bodega no está registrada, se lanza MISSING_BODEGA para que el asistente pida registrarla.
  if (!destinoMap) {
    throw new Error(`MISSING_BODEGA|${idDestinoRaw}|${configCli.SOLICITANTE}|${configCli.NOMBRE_CLIENTE}`);
  }

  // Sin ciudad en Destinatarios no se asigna una por defecto: se avisa para completar la hoja.
  if (!destinoMap.ciudad) {
    throw new Error(`El destino ${idDestinoRaw} no tiene ciudad en Destinatarios. Complétela y vuelva a intentar.`);
  }

  return {
    Pedido_ID: ordenCompra,
    Destinatario_Merc: destinoMap.idDestino,      
    Destino_Original: idDestinoRaw,                
    Destinatario: destinoMap.nombreDestino,        
    Pedido_Fecha: fechaPedido || '',
    Fecha_Entrega: fechaEntrega || '',
    Archivo: fileName,
    Archivo_ID: fileId,
    Archivo_URL: `https://drive.google.com/file/d/${fileId}/view`,
    Ciudad: destinoMap.ciudad
  };
}

/**
 * Saca la tabla de productos del texto del PDF. Cada fila debe tener un EAN de 13 digitos, un codigo
 * de producto, la descripcion, la unidad, la cantidad y el valor unitario.
 * El material SAP se busca en la bodega por EAN. Si el EAN esta en Productos pero no en Bodega,
 * se deja con stock 0, para que salga como back order y nunca como OK.
 * Solo lee: no escribe nada.
 *
 * @param {String} texto Texto completo del PDF.
 * @param {String} pedidoId Numero de OC al que pertenecen los productos.
 * @param {Object} diccBodega Mapa EAN -> lista de lotes { material, stock, fechaDisplay } (de obtenerMapaBodegaGeneral_).
 * @param {Object} diccProductos Mapa EAN -> codigo de material (hoja Productos).
 * @return {Array<Object>} Productos { Item_ID (codigo aleatorio), Pedido_ID, Producto_Cod (EAN o "SIN_EAN"),
 *   Descripcion, Cant, Valor_Unitario, EAN_Mapeado, EAN_Cliente, Materiales_Disponibles }.
 *   Lista vacia si no se encuentra la tabla.
 */
function medipiel_parseItems_(texto, pedidoId, diccBodega, diccProductos) {
  // 1. Recortar el bloque de la tabla: desde "CÓDIGO BARRAS" hasta "Valor Subtotal" o "Valor Neto".
  const m = texto.match(/C[ÓO]DIGO\s+BARRAS[\s\S]*?COSTO\s+UN([\s\S]*?)(?:Valor\s+Subtotal\s*:|Valor\s+Neto\s*:|$)/i);
  const bloqueTabla = m && m[1] ? m[1].trim() : '';
  if (!bloqueTabla) return [];

  const items = [];
  
  // 2. Expresión que reconoce cada fila: EAN de 13 dígitos, código, descripción, unidad, cantidad y valor.
  // Expresión regular ESTRICTA para EAN de 13 dígitos (\d{13}) tal como lo requieres
  const re = /(\d{13})\s+(\d{3,})\s+([\s\S]+?)\s*(?:Und|UND|Un|UN|und|U\.M\.)\s+(\d+(?:[.,]\d+)?)\s+([$\d.,]+)/gmi;
  let match;

  // 3. Recorrer cada fila encontrada en la tabla.
  while ((match = re.exec(bloqueTabla)) !== null) {
    const eanRaw = String(match[1]).trim();
    // 3a. Buscar el material en Bodega por EAN.
    let mats = eanRaw ? (diccBodega[eanRaw] || []) : [];
    
    // Existe en Productos pero no en Bodega: stock 0 para que salga como Back Order, nunca como OK
    if (mats.length === 0 && eanRaw && diccProductos[eanRaw]) {
        mats = [{ material: diccProductos[eanRaw], stock: 0, fechaDisplay: "Sin stock" }];
    }
    
    // 3b. Guardar el producto con sus datos.
    items.push({
      Item_ID: Math.random().toString(36).substring(2, 10).toUpperCase(),
      Pedido_ID: pedidoId,
      Producto_Cod: eanRaw || "SIN_EAN",
      Descripcion: String(match[3]).replace(/\s+/g, ' ').trim(),
      Cant: medipiel_parseNumber_(match[4]),
      Valor_Unitario: medipiel_parseNumber_(String(match[5]).replace(/[^\d.,-]/g, '')),
      EAN_Mapeado: eanRaw,
      EAN_Cliente: eanRaw,
      Materiales_Disponibles: mats 
    });
  }
  return items;
}

/**
 * Deja el numero de OC listo para compararlo: quita espacios y el apostrofe inicial, convierte la
 * notacion cientifica (ej. 4.5E+9) a numero entero y elimina la parte decimal.
 *
 * @param {*} val Valor de la celda o del PDF (texto o numero).
 * @return {String} OC limpia. Cadena vacia si el valor es nulo o vacio.
 */
function medipiel_limpiarOc_(val) {
  if (val === null || val === undefined) return "";
  // 1. Quitar espacios y el apóstrofe inicial.
  let s = String(val).trim().replace(/^'/, "");
  // 2. Si viene en notación científica (ej. 4.5E+9), convertirla a número entero.
  if (/e/i.test(s)) {
    let num = Number(s);
    if (!isNaN(num)) s = num.toFixed(0);
  }
  // 3. Quitar la parte decimal, si la hay.
  if (s.includes(".")) s = s.split(".")[0];
  return s.trim();
}

/**
 * Convierte un texto en una clave comparable: sin tildes, en mayusculas y solo letras y numeros.
 * Ejemplo: "Orden De Compra" -> "ORDENDECOMPRA". Sirve para comparar encabezados y codigos SAP.
 *
 * @param {*} str Texto o numero a normalizar.
 * @return {String} Clave normalizada. Cadena vacia si el valor es vacio.
 */
function medipiel_normalizeKey_(str) {
  if (!str) return "";
  return String(str)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

/**
 * Lee la fila de encabezados de una hoja y devuelve un mapa de "nombre normalizado" a numero de columna.
 * La columna A es 1, la B es 2, y asi sucesivamente. Solo lee: no escribe nada.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} s Hoja a leer.
 * @param {Number} rowNum Fila donde estan los encabezados (en CONSOLIDADO es la 3).
 * @return {Object} { NOMBRE_NORMALIZADO: numeroDeColumna }. Objeto vacio si la hoja no llega a esa fila.
 */
function medipiel_getHeaderMap_(s, rowNum) {
  if (s.getLastRow() < rowNum) return {};
  // Leer al menos 26 columnas (A a Z), aunque la hoja tenga menos encabezados.
  const numCols = Math.max(s.getLastColumn(), 26);
  const h = s.getRange(rowNum, 1, 1, numCols).getValues()[0];
  const m = {};
  h.forEach((x, i) => { 
    if (x) m[medipiel_normalizeKey_(x)] = i + 1; 
  });
  return m;
}

/**
 * Arma una fila para escribir en la hoja. Recibe los datos por nombre de columna y los pone en la
 * posicion que indica el mapa de encabezados. Las columnas sin dato quedan con cadena vacia.
 * Solo arma el arreglo en memoria: no escribe en la hoja.
 *
 * @param {Object} map Mapa de encabezados (ver medipiel_getHeaderMap_).
 * @param {Object} dataObj Datos por nombre de columna, por ejemplo { "Orden De Compra": "4500123" }.
 *   Los nombres que no existen en el mapa se ignoran.
 * @param {Number} numCols Cantidad de columnas de la fila (en CONSOLIDADO: 11, de A a K).
 * @return {Array} Fila de longitud numCols, lista para escribirse.
 */
function medipiel_createRowArray_(map, dataObj, numCols) {
  const row = new Array(numCols).fill("");
  // Por cada dato recibido, buscar su columna por nombre normalizado y ponerlo en esa posición.
  for (const rawKey in dataObj) {
    const normKey = medipiel_normalizeKey_(rawKey);
    if (map[normKey]) {
      row[map[normKey] - 1] = dataObj[rawKey];
    }
  }
  return row;
}

/**
 * Lee la hoja Productos y arma un mapa de EAN a codigo de material SAP.
 * Busca las columnas "EAN" y "MATERIAL ID" por nombre. Si no las encuentra, usa la columna 5 (E)
 * para el EAN y la columna 2 (B) para el material. Solo lee: no escribe nada.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Hoja Productos.
 * @return {Object} { EAN: codigoMaterial }. Objeto vacio si la hoja tiene menos de 2 filas.
 */
function medipiel_getProductosMap_(sheet) {
  const map = {}; 
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  const h = data[0].map(x => medipiel_normalizeKey_(x));
  
  // 1. Buscar las columnas EAN y MATERIAL ID por nombre. Si no existen, usar posiciones fijas.
  let cE = h.indexOf(medipiel_normalizeKey_("EAN"));
  let cM = h.indexOf(medipiel_normalizeKey_("MATERIAL ID"));
  
  if (cE === -1) cE = 4; 
  if (cM === -1) cM = 1; 
  
  // 2. Recorrer las filas (desde la 2, sin el encabezado) y guardar EAN a material.
  for (let i = 1; i < data.length; i++) {
    const ean = data[i][cE] ? String(data[i][cE]).trim() : "";
    if (ean) map[ean] = String(data[i][cM]).trim();
  }
  return map;
}

/**
 * Busca un patron en un texto y devuelve lo que capturo el primer parentesis, sin espacios a los lados.
 *
 * @param {String} text Texto donde buscar.
 * @param {RegExp} regex Expresion regular. El valor buscado debe ir en el primer parentesis.
 * @return {String|null} Texto encontrado, o null si no hay coincidencia.
 */
function medipiel_matchRegex_(text, regex) { const m = text.match(regex); return m ? m[1].trim() : null; }
/**
 * Convierte un numero escrito en formato colombiano (punto para miles, coma para decimales) a numero.
 * Ejemplos: "1.234,5" -> 1234.5 y "12" -> 12. Si no puede convertirlo, devuelve 0.
 * Atencion: un decimal escrito con punto ("2.5") se toma como separador de miles y da 25.
 *
 * @param {String|Number} v Texto del numero.
 * @return {Number} Numero convertido. 0 si el texto no es valido.
 */
function medipiel_parseNumber_(v) { return Number(String(v).trim().replace(/\./g, '').replace(',', '.')) || 0; }
/**
 * Convierte una fecha escrita como dd/mm/aaaa (o dd-mm-aaaa) en un objeto Date.
 * Si el año tiene dos digitos, le suma 2000. Si el texto no tiene el formato esperado,
 * lo devuelve sin cambios.
 *
 * @param {String} d Fecha en texto, por ejemplo "05/03/26".
 * @return {Date|String} Date si se pudo convertir; "" si viene vacia; el mismo texto si no coincide con el formato.
 */
function medipiel_parseDateEs_(d) {
  if (!d) return "";
  // 1. Separar día, mes y año.
  const p = d.split(/[-/]/);
  if (p.length === 3) {
    let y = parseInt(p[2], 10);
    // 2. Si el año tiene dos dígitos, asumir 20xx.
    if (y < 100) y += 2000;
    return new Date(y, p[1] - 1, p[0]);
  }
  return d;
}
