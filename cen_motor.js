/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: INTEGRACIÓN Y PROCESAMIENTO MULTICLIENTE - CEN (CENTRO DE NEGOCIOS CARVAJAL)
 * ARCHIVO: cen_logica.gs
 * ESTADO: PRODUCCIÓN - Anti-Duplicado Compuesto, Asignación Dinámica KAM y Limpieza Base64
 * =========================================================================================
 * 📝 DESCRIPCIÓN TÉCNICA:
 * Este módulo gestiona la ingesta, lectura, validación y consolidación de órdenes de compra
 * provenientes del Centro Electrónico de Negocios (CEN / Carvajal).
 * 
 * 🛠️ FUNCIONALIDADES Y REGLAS CLAVE:
 * 1. ESCUDO ANTI-DUPLICADOS COMPUESTO (OC + CLIENTE): La verificación evalúa la combinación
 *    `${OrdenDeCompra}_${SolicitanteSAP}`. Evita que OCs con números iguales pertenecientes
 *    a clientes distintos (ej: OC 6 de Bella Piel vs OC 6 de Cruz Verde) se bloqueen.
 * 2. ASIGNACIÓN DINÁMICA DE KAM (COLUMNAS G Y H): Lee el correo de la Columna G y las iniciales
 *    de la Columna H en la hoja `Asignacion_KAM` para asociar los consecutivos (ej: MC-26XXXX).
 * 3. TRASLADO Y RENOMBRADO EN DRIVE: Mueve los archivos procesados a la carpeta de históricos
 *    y los renombra como `${NombreCliente}_CEN_${FechaPedido}.xlsx`.
 * 4. SANITIZACIÓN BASE64: Limpia saltos de línea en la cadena Base64 generada por Apps Script
 *    para prevenir descargas de archivos CSV de 0 bytes (vacíos).
 * =========================================================================================
 */

/**
 * Punto de entrada principal para la pre-analítica de los archivos CEN.
 * Escanea la carpeta de Google Drive, filtra OCs procesadas previamente por cliente
 * y extrae las órdenes y posiciones asociadas.
 * 
 * @return {Object} Objeto con clientesDetectados, pedidosExtraidos, itemsExtraidos y errores.
 */
function preAnalizarCEN() {
  try {
    const configCli = CONFIG.CLIENTES["CEN"];
    if (!configCli || !configCli.FOLDER_ID) {
      throw new Error("La configuración de CEN no está definida en CONFIG de main.gs.");
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    const pedidosProcesados = new Set();
    
    // ESCUDO 1: Carga en memoria de la combinación (OC + SAP ID) para antiduplicado estricto por cliente
    if (shConsolidado && shConsolidado.getLastRow() >= 3) {
      const dataCons = shConsolidado.getDataRange().getValues();
      const headCons = dataCons[0].map(h => cen_normalizeKey_(h));
      
      let colOcCons = 1;  // Orden de Compra (Columna B por defecto)
      let colSapCons = 2; // Sap Id / Solicitante (Columna C por defecto)

      const foundOcIdx = headCons.indexOf(cen_normalizeKey_("Orden de Compra"));
      if (foundOcIdx !== -1) colOcCons = foundOcIdx;

      const foundSapIdx = headCons.findIndex(h => h.includes("SAP") || h.includes("SOLICITANTE"));
      if (foundSapIdx !== -1) colSapCons = foundSapIdx;

      for (let i = 1; i < dataCons.length; i++) {
        const ocVal = cen_limpiarOc_(dataCons[i][colOcCons]);
        const sapVal = cen_normalizeKey_(dataCons[i][colSapCons]);
        if (ocVal && sapVal) {
          pedidosProcesados.add(`${ocVal}_${sapVal}`);
        }
      }
    }

    // Carga de diccionarios auxiliares para geolocalización e inventario
    const diccDestinatarios = obtenerDiccionarioDestinatarios();
    const diccBodega = obtenerMapaBodegaGeneral_(ss.getSheetByName(CONFIG.SHEET_BODEGA || "Bodega"));

    const folder = DriveApp.getFolderById(configCli.FOLDER_ID);
    const files = obtenerArchivosCliente_(folder, configCli);
    const archivos = [];
    const maxLimit = CONFIG.LIMIT_FILES > 0 ? CONFIG.LIMIT_FILES : 200;

    while (files.hasNext() && archivos.length < maxLimit) {
      const file = files.next();
      const name = file.getName().toLowerCase();
      if (name.endsWith('.xlsm') || name.endsWith('.xlsx') || name.endsWith('.xls')) {
        archivos.push(file);
      }
    }

    if (!archivos.length) {
      return { clientesDetectados: [], pedidosExtraidos: [], itemsExtraidos: [], errores: [], yaProcesados: [] };
    }

    return cen_motorLecturaExcel(archivos, diccBodega, configCli, pedidosProcesados, diccDestinatarios, ss);

  } catch (err) {
    throw new Error("Error en preAnalizarCEN: " + err.message);
  }
}

/**
 * Convierte temporalmente cada archivo Excel en Google Sheets para realizar la extracción
 * de la cabecera del comprador, número de OC e ítems de venta.
 * 
 * @param {Array<DriveApp.File>} archivos Lista de archivos a procesar.
 * @param {Object} diccBodega Mapa de inventario de lotes disponible.
 * @param {Object} configCli Objeto de configuración general de CEN.
 * @param {Set} pedidosProcesados Conjunto de OCs históricas indexadas por combinación `${OC}_${SAPID}`.
 * @param {Object} diccDestinatarios Diccionario de puntos de entrega.
 * @param {SpreadsheetApp.Spreadsheet} ss Referencia al libro de cálculo activo.
 * @return {Object} Payload consolidado con pedidos e ítems extraídos.
 */
function cen_motorLecturaExcel(archivos, diccBodega, configCli, pedidosProcesados, diccDestinatarios, ss) {
  const pedidosExtraidos = [];
  const itemsExtraidos = [];
  const errores = [];
  const yaProcesados = [];
  const mapaClientesCount = {};

  const shProd = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  const diccProductosGlobal = cen_getCatalogoProductosGlobal_(shProd);

  archivos.forEach((file) => {
    let tempFileId = null;
    try {
      tempFileId = cen_convertirExcelASheet_(file);
      const tempSs = SpreadsheetApp.openById(tempFileId);
      const sheet = tempSs.getSheets()[0];
      const data = sheet.getDataRange().getValues();

      if (data.length < 2) return;

      const head = data[0].map(h => cen_normalizeKey_(h));
      
      const colCompradorEan = head.findIndex(h => h.includes("EANEMPRESACOMPRADORA") || h.includes("EANCOMPRADOR"));
      const colCompradorNombre = head.findIndex(h => h.includes("RAZONSOCIALEMPRESACOMPRADORA") || h.includes("NOMBRECOMPRADOR"));
      
      // DETECCIÓN EXACTA DE COLUMNA {Numero de la Orden de compra}
      let colOc = head.findIndex(h => (h.includes("NUMERODELAORDEN") || h.includes("NUMEROORDEN") || h.includes("NUMEROOC")) && !h.includes("TIPO") && !h.includes("LINEA") && !h.includes("POSICION"));
      if (colOc === -1) {
        colOc = head.findIndex(h => h.includes("ORDENDECOMPRA") && !h.includes("TIPO") && !h.includes("FECHA") && !h.includes("LINEA") && !h.includes("POSICION"));
      }

      const colFechaDoc = head.findIndex(h => (h.includes("FDOCUMENTO") || h.includes("FECHAORDEN") || h.includes("FECHADOCUMENTO")) && !h.includes("ENTREGA"));
      const colFechaEnt = head.findIndex(h => h.includes("FECHAMAXIMA") || h.includes("FECHAENTREGA"));
      const colEanItem = head.findIndex(h => (h.includes("EANDELITEM") || h.includes("EANITEM")) && !h.includes("LUGAR") && !h.includes("EMPRESA"));
      const colDescItem = head.findIndex(h => h.includes("DESCRIPCIONDELITEM") || h.includes("DESCRIPCIONITEM"));
      const colCant = head.findIndex(h => h.includes("CANTIDADTOTAL") || h === "CANTIDAD");
      const colLugarEntregaEan = head.findIndex(h => h.includes("EANLUGARENTREGA") || h.includes("EANPUNTODEVENTA") || h.includes("PUNTODEVENTA"));
      const colLugarEntregaNombre = head.findIndex(h => h.includes("NOMBRELUGARENTREGA") || h.includes("NOMBREPUNTODEVENTA"));

      if (colOc === -1 || colEanItem === -1 || colCant === -1) {
        throw new Error(`El archivo ${file.getName()} no contiene la estructura esperada de columnas CEN.`);
      }

      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        if (!row.join("").trim()) continue;

        const ocRaw = cen_limpiarOc_(row[colOc]);
        if (!ocRaw) continue;

        const eanComprador = colCompradorEan !== -1 ? String(row[colCompradorEan]).trim() : "";
        const nombreComprador = colCompradorNombre !== -1 ? String(row[colCompradorNombre]).trim() : "CLIENTE_CEN";
        
        const clienteInternoObj = cen_identificarClienteInterno_(eanComprador, nombreComprador, configCli);
        const clienteKey = clienteInternoObj.key;
        const solicitanteSap = clienteInternoObj.solicitante;

        // Clave combinada única: OC + SAP ID del Cliente
        const comboKey = `${ocRaw}_${cen_normalizeKey_(solicitanteSap)}`;

        // ESCUDO 2: Descarte únicamente si esta OC YA FUE REGISTRADA PARA ESTE MISMO CLIENTE
        if (pedidosProcesados.has(comboKey)) {
          if (!yaProcesados.includes(ocRaw)) yaProcesados.push(ocRaw);
          continue;
        }

        let pedidoObj = pedidosExtraidos.find(p => p.Pedido_ID === ocRaw);
        if (!pedidoObj) {
          const eanEntrega = colLugarEntregaEan !== -1 ? String(row[colLugarEntregaEan]).trim() : "";
          const nombreEntrega = colLugarEntregaNombre !== -1 ? String(row[colLugarEntregaNombre]).trim() : "";

          const destinoMap = diccDestinatarios[eanEntrega] || diccDestinatarios[clienteInternoObj.solicitante] || {
            idDestino: clienteInternoObj.solicitante,
            nombreDestino: nombreEntrega || clienteInternoObj.nombreDisplay,
            ciudad: "COTA"
          };

          let fechaPedidoStr = colFechaDoc !== -1 && row[colFechaDoc] ? cen_formatFechaStr_(row[colFechaDoc]) : Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
          let fechaEntregaStr = colFechaEnt !== -1 && row[colFechaEnt] ? cen_formatFechaStr_(row[colFechaEnt]) : "";

          pedidoObj = {
            Pedido_ID: ocRaw,
            Cliente_Key: clienteKey,
            Cliente_Nombre: clienteInternoObj.nombreDisplay,
            Solicitante_SAP: clienteInternoObj.solicitante,
            Lista_Precios: clienteInternoObj.listaPrecios,
            Fecha_Pedido: fechaPedidoStr,
            Fecha_Entrega: fechaEntregaStr,
            Id_Destino: destinoMap.idDestino,
            Destinatario: destinoMap.nombreDestino,
            Ciudad: destinoMap.ciudad || "COTA",
            Archivo_Nombre: file.getName(),
            Archivo_Id: file.getId()
          };
          pedidosExtraidos.push(pedidoObj);

          if (!mapaClientesCount[clienteKey]) {
            mapaClientesCount[clienteKey] = { key: clienteKey, nombreDisplay: clienteInternoObj.nombreDisplay, count: 0 };
          }
          mapaClientesCount[clienteKey].count++;
        }

        const eanItemRaw = cen_limpiarEan_(row[colEanItem]);
        const descItemExcel = colDescItem !== -1 ? String(row[colDescItem]).trim() : "Producto CEN";
        const cantItem = parseInt(row[colCant]) || 0;

        if (cantItem <= 0 || !eanItemRaw) continue;

        const keyTarifa = cen_normalizeKey_(clienteInternoObj.listaPrecios);
        const productoEncontrado = diccProductosGlobal[keyTarifa] ? diccProductosGlobal[keyTarifa][eanItemRaw] : null;

        let eanMapeado = "";
        let valorUnitario = 0;
        let materialesDisponibles = [];
        let descripcionOficial = productoEncontrado && productoEncontrado.nombre ? productoEncontrado.nombre : descItemExcel;

        if (productoEncontrado) {
          eanMapeado = eanItemRaw;
          valorUnitario = productoEncontrado.precio;
          const oficialMat = productoEncontrado.materialOficial;

          const stockLotes = diccBodega[eanItemRaw] || [];
          if (stockLotes.length > 0) {
            materialesDisponibles = stockLotes.map(s => ({
              material: s.material || oficialMat,
              stock: s.stock,
              fechaDisplay: s.fechaDisplay
            }));
          } else {
            materialesDisponibles = [{
              material: oficialMat,
              stock: 0,
              fechaDisplay: "Sin fecha"
            }];
          }
        } else {
          eanMapeado = "";
          valorUnitario = 0;
          materialesDisponibles = [];
        }

        itemsExtraidos.push({
          Item_ID: `${ocRaw}_${itemsExtraidos.length + 1}`,
          Pedido_ID: ocRaw,
          Cliente_Key: clienteKey,
          Producto_Cod: eanItemRaw,
          Descripcion: descripcionOficial,
          Cant: cantItem,
          Valor_Unitario: valorUnitario,
          EAN_Mapeado: eanMapeado,
          Materiales_Disponibles: materialesDisponibles
        });
      }

    } catch (e) {
      errores.push(`${file.getName()}: ${e.message}`);
    } finally {
      if (tempFileId) {
        try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (ignore) {}
      }
    }
  });

  const clientesDetectados = Object.values(mapaClientesCount);

  return {
    clientesDetectados: clientesDetectados,
    pedidosExtraidos: pedidosExtraidos,
    itemsExtraidos: itemsExtraidos,
    errores: errores,
    yaProcesados: yaProcesados
  };
}

/**
 * Remueve apóstrofes iniciales, notación científica y decimales residuales de los números de OC.
 * 
 * @param {*} val Valor original extraído de la celda de Excel.
 * @return {String} Cadena limpia del número de orden.
 */
function cen_limpiarOc_(val) {
  if (val === null || val === undefined) return "";
  let s = String(val).trim();
  s = s.replace(/^'/, ""); 
  if (/e/i.test(s)) {
    let num = Number(s);
    if (!isNaN(num)) s = num.toFixed(0);
  }
  if (s.includes(".")) s = s.split(".")[0];
  return s.trim();
}

/**
 * Identifica la razón social interna y credenciales SAP de ISDIN basadas en la tabla MAPEO_CLIENTES.
 * 
 * @param {String} eanComprador EAN de la empresa compradora.
 * @param {String} nombreComprador Nombre o razón social expuesta en el portal CEN.
 * @param {Object} configCli Configuración general del cliente CEN.
 * @return {Object} Mapeo con key, nombreDisplay, solicitante y listaPrecios.
 */
function cen_identificarClienteInterno_(eanComprador, nombreComprador, configCli) {
  const mapeo = configCli.MAPEO_CLIENTES || {};
  const normComprador = cen_normalizeKey_(nombreComprador);
  const normEan = cen_normalizeKey_(eanComprador);

  for (const key in mapeo) {
    const cli = mapeo[key];
    const matchEan = cli.eans && cli.eans.includes(normEan);
    const matchNombre = cli.aliases && cli.aliases.some(alias => normComprador.includes(cen_normalizeKey_(alias)));
    
    if (matchEan || matchNombre) {
      return {
        key: key,
        nombreDisplay: cli.NOMBRE_CLIENTE,
        solicitante: cli.SOLICITANTE,
        listaPrecios: cli.LISTA_PRECIOS
      };
    }
  }

  return {
    key: "CLIENTE_CEN_GENERICO",
    nombreDisplay: nombreComprador || "CLIENTE CEN",
    solicitante: configCli.SOLICITANTE || "11000000",
    listaPrecios: "MEDIPIEL BEAUTYCALIA"
  };
}

/**
 * Carga el catálogo global de productos filtrado por lista de precios B2B para obtener
 * descripciones oficiales y precios unitarios.
 * 
 * @param {SpreadsheetApp.Sheet} shProd Hoja de Productos.
 * @return {Object} Diccionario indexado por [listaPrecios][ean].
 */
function cen_getCatalogoProductosGlobal_(shProd) {
  const dicc = {};
  if (!shProd) return dicc;

  const data = shProd.getDataRange().getValues();
  if (data.length < 2) return dicc;

  const head = data[0].map(h => cen_normalizeKey_(h));
  const colEan = head.findIndex(h => h.includes("EAN") || h.includes("BARRA"));
  const colMat = head.findIndex(h => h.includes("MATERIAL") || h === "SAP");
  const colNombre = head.findIndex(h => h.includes("DESCRIPCION") || h.includes("NOMBRE") || h.includes("PRODUCTO"));
  const colLista = head.findIndex(h => h.includes("LISTA") || h.includes("CLIENTE"));
  const colCosto = head.findIndex(h => h.includes("COSTO") || h.includes("PRECIO") || h.includes("VALOR"));

  if (colEan === -1 || colMat === -1 || colLista === -1 || colCosto === -1) return dicc;

  for (let i = 1; i < data.length; i++) {
    const listaKey = cen_normalizeKey_(data[i][colLista]);
    const ean = cen_limpiarEan_(data[i][colEan]);
    const mat = String(data[i][colMat]).trim();
    const nombre = colNombre !== -1 ? String(data[i][colNombre]).trim() : "";
    const precio = parseFloat(data[i][colCosto]) || 0;

    if (!listaKey || !ean) continue;

    if (!dicc[listaKey]) dicc[listaKey] = {};
    dicc[listaKey][ean] = { materialOficial: mat, nombre: nombre, precio: precio };
  }
  return dicc;
}

/**
 * Persiste los pedidos confirmados en la hoja CONSOLIDADO, genera la estructura del CSV para SAP,
 * traslada y renombra los archivos en Google Drive y envía notificación vía correo a la KAM.
 * 
 * @param {Object} paquete Datos completos procesados en la interfaz UI.
 * @return {Object} Respuesta con estado de éxito, Base64 del CSV y nombre del archivo generado.
 */
function guardarDefinitivoCEN(paquete) {
  try {
    const configCli = CONFIG.CLIENTES["CEN"];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");

    if (!shConsolidado) throw new Error("No se encontró la hoja CONSOLIDADO.");

    const currentYearSuffix = String(new Date().getFullYear()).slice(-2);
    const shAgentes = ss.getSheetByName("Iniciales de agente");
    const shAsignacion = ss.getSheetByName("Asignacion_KAM");

    const filasConsolidado = [];
    const filasCSV = [];
    let contadorPedidoCSV = 1;
    let primerConsecutivo = "";
    const pedidosValidos = new Set(); // Pedido_ID que sí quedaron inyectados en CONSOLIDADO

    const pedidosPorCliente = {};
    paquete.pedidos.forEach(p => {
      if (!pedidosPorCliente[p.Cliente_Key]) pedidosPorCliente[p.Cliente_Key] = [];
      pedidosPorCliente[p.Cliente_Key].push(p);
    });

    for (const cliKey in pedidosPorCliente) {
      const pedidosCli = pedidosPorCliente[cliKey];
      const primerPedido = pedidosCli[0];
      const solicitanteSap = primerPedido.Solicitante_SAP;

      // RESOLUCIÓN DE KAM E INICIALES DESDE Asignacion_KAM (Columna G = Correo, Columna H = Iniciales)
      const kamInfo = cen_getKamInfo_(solicitanteSap, shAsignacion, shAgentes);
      const prefix = kamInfo.initials + "-" + currentYearSuffix;

      let maxConsecutive = cen_obtenerMaxConsecutivo_(shConsolidado, prefix);

      pedidosCli.forEach(p => {
        const itemsPedido = paquete.items.filter(i => i.Pedido_ID === p.Pedido_ID);
        let totalUnidades = 0;
        let totalValor = 0;
        let posicionItem = 10;

        itemsPedido.forEach(item => {
          const matSeleccionado = paquete.materialesSeleccionados[item.Item_ID];
          if (matSeleccionado && !matSeleccionado.includes("SIN_")) {
            const cant = parseInt(item.Cant) || 0;
            const valUnit = parseFloat(item.Valor_Unitario) || 0;
            totalUnidades += cant;
            totalValor += (cant * valUnit);

            const fechaEntUI = paquete.fechasEntregas[p.Pedido_ID] ? paquete.fechasEntregas[p.Pedido_ID].fecha : "";
            let fechaEntFormatted = "";
            if (fechaEntUI) {
              const parts = fechaEntUI.split("-");
              if (parts.length === 3) fechaEntFormatted = `${parts[2]}/${parts[1]}/${parts[0]}`;
            }

            filasCSV.push([
              contadorPedidoCSV,
              "",
              configCli.SAP.CLASE_PEDIDO,
              configCli.SAP.ORG_VENTAS,
              configCli.SAP.CANAL,
              configCli.SAP.SECTOR,
              solicitanteSap,
              p.Id_Destino,
              p.Pedido_ID,
              fechaEntFormatted,
              "", "",
              posicionItem,
              matSeleccionado,
              "",
              cant
            ]);
            posicionItem += 10;
          }
        });

        if (totalUnidades > 0) {
          pedidosValidos.add(p.Pedido_ID);
          maxConsecutive++;
          const consecutivoStr = String(maxConsecutive).padStart(4, "0");
          const idGenerado = prefix + consecutivoStr;
          if (!primerConsecutivo) primerConsecutivo = idGenerado;

          filasConsolidado.push([
            idGenerado,
            p.Pedido_ID,
            solicitanteSap,
            p.Cliente_Nombre,
            p.Fecha_Pedido,
            p.Id_Destino,
            p.Destinatario,
            p.Ciudad,
            paquete.fechasEntregas[p.Pedido_ID] ? paquete.fechasEntregas[p.Pedido_ID].fecha : "",
            totalUnidades,
            totalValor,
            "", "",
            kamInfo.name // Registra el nombre del KAM en la Columna N
          ]);
          contadorPedidoCSV++;
        }
      });

      if (paquete.notificarKAM !== false && kamInfo.email) {
        try {
          cen_enviarCorreoKam_(pedidosCli, paquete, kamInfo, primerPedido.Cliente_Nombre);
        } catch (eMailErr) {
          console.error("Error notificando KAM: " + eMailErr.message);
        }
      }
    }

    if (filasConsolidado.length > 0) {
      const lastRow = shConsolidado.getLastRow();
      shConsolidado.getRange(lastRow + 1, 1, filasConsolidado.length, filasConsolidado[0].length).setValues(filasConsolidado);
    }

    // ESCUDO 3: Traslado y renombrado de archivos Excel procesados a carpeta de históricos
    if (configCli.PROCESSED_FOLDER_ID) {
      const folderDestino = DriveApp.getFolderById(configCli.PROCESSED_FOLDER_ID);
      const archivosProcesadosMap = {};

      paquete.pedidos.forEach(p => {
        // Solo se mueve el archivo si el pedido efectivamente quedó registrado en CONSOLIDADO
        // (evita que una OC sin ítems válidos desaparezca de la carpeta de pendientes sin rastro).
        if (p.Archivo_Id && pedidosValidos.has(p.Pedido_ID)) {
          archivosProcesadosMap[p.Archivo_Id] = {
            clienteNombre: p.Cliente_Nombre || "CLIENTE_CEN",
            fechaPedido: p.Fecha_Pedido || Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")
          };
        }
      });

      for (const fileId in archivosProcesadosMap) {
        try {
          const file = DriveApp.getFileById(fileId);
          file.moveTo(folderDestino);

          const info = archivosProcesadosMap[fileId];
          const nameClean = info.clienteNombre.replace(/[^a-zA-Z0-9_ -]/g, "").trim();
          const fechaClean = info.fechaPedido.replace(/[\/]/g, "-");
          
          const originalName = file.getName();
          const extIdx = originalName.lastIndexOf(".");
          const ext = extIdx !== -1 ? originalName.substring(extIdx) : ".xlsx";

          const nuevoNombre = `${nameClean}_CEN_${fechaClean}${ext}`;
          file.setName(nuevoNombre);
        } catch (eMove) {
          console.error("Error trasladando archivo CEN procesado: " + eMove.message);
        }
      }
    }

    const csvHeader = [ "Contador", "Nº Pedido SAP", "Clase pedido", "Org ventas", "Canal", "Sector", "Solicitante", "Destinatario Merc", "Num pedido cliente", "Fecha preferente Entrega", "Descuento ZTK1 (%)", "Motivo pedido", "Posición pedido", "Material", "EAN", "Cantidad" ];
    filasCSV.unshift(csvHeader);
    const csvContent = filasCSV.map(r => r.join(";")).join("\n");
    const fileName = `Pedido_CEN_${primerConsecutivo || Utilities.formatDate(new Date(), "GMT-5", "yyyyMMdd")}.csv`;

    try {
      const folderCSV = DriveApp.getFolderById(CONFIG.FOLDER_CSV_ID);
      folderCSV.createFile(fileName, csvContent, MimeType.PLAIN_TEXT);
    } catch (eCsv) {
      console.error("Error guardando CSV en Drive: " + eCsv.message);
    }

    // SANITIZACIÓN BASE64: Elimina saltos de línea para prevenir descargas vacías (0 bytes)
    const base64Limpio = Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8).replace(/[\r\n]/g, "");

    return {
      success: true,
      mensaje: `Se han procesado exitosamente ${filasConsolidado.length} órdenes de compra de CEN.`,
      csvData: base64Limpio,
      fileName: fileName
    };

  } catch (err) {
    throw new Error(err.message);
  }
}

function cen_convertirExcelASheet_(file) {
  const resource = {
    title: file.getName().replace(/\.[^/.]+$/, "") + "_TEMP",
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: file.getParents().next().getId() }]
  };
  const tempFile = Drive.Files.insert(resource, file.getBlob());
  return tempFile.id;
}

function cen_limpiarEan_(val) {
  if (!val) return "";
  let s = String(val).trim();
  if (/e/i.test(s)) {
    let num = Number(s);
    if (!isNaN(num)) s = num.toFixed(0);
  }
  if (s.includes(".")) s = s.split(".")[0];
  s = s.replace(/\D/g, "");
  if (s.length > 13) s = s.substring(0, 13);
  return s;
}

function cen_formatFechaStr_(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(val).trim();
}

function cen_normalizeKey_(text) {
  if (!text) return "";
  return String(text).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]/g, "").toUpperCase().trim();
}

function cen_obtenerMaxConsecutivo_(shConsolidado, prefix) {
  let max = 0;
  const data = shConsolidado.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const idVal = String(data[i][0]).trim();
    if (idVal.startsWith(prefix)) {
      const numPart = parseInt(idVal.replace(prefix, ""), 10) || 0;
      if (numPart > max) max = numPart;
    }
  }
  return max;
}

/**
 * Obtiene el Nombre, Iniciales (Columna H) y Correo electrónico (Columna G) del KAM asignado.
 * Escanea dinámicamente las cabeceras de las hojas Asignacion_KAM e Iniciales de agente.
 * 
 * @param {String} solicitanteSap Código SAP ID del cliente solicitante (ej: 11026727).
 * @param {SpreadsheetApp.Sheet} shAsignacion Hoja Asignacion_KAM.
 * @param {SpreadsheetApp.Sheet} shAgentes Hoja Iniciales de agente.
 * @return {Object} Objeto con name, initials y email.
 */
function cen_getKamInfo_(solicitanteSap, shAsignacion, shAgentes) {
  let name = "SARA CARDONA";
  let initials = "SC";
  let email = "sara.cardona@isdin.com";

  if (shAsignacion) {
    const dataAsig = shAsignacion.getDataRange().getValues();
    if (dataAsig.length >= 2) {
      const headAsig = dataAsig[0].map(h => cen_normalizeKey_(h));
      
      const colSap = headAsig.findIndex(h => h.includes("SAP") || h.includes("ID"));
      const colKam = headAsig.findIndex(h => h.includes("KAMENCARGADO") || h === "KAM" || h.includes("ENCARGADO"));
      const colMail = headAsig.findIndex(h => h.includes("CORREO") || h.includes("EMAIL") || h.includes("MAIL"));
      const colInic = headAsig.findIndex(h => h.includes("INICIAL") || h.includes("AGENTE") || h.includes("CODIGO"));

      const normSolicitante = cen_normalizeKey_(solicitanteSap);

      for (let i = 1; i < dataAsig.length; i++) {
        const sapVal = colSap !== -1 ? cen_normalizeKey_(dataAsig[i][colSap]) : cen_normalizeKey_(dataAsig[i][0]);
        if (sapVal === normSolicitante) {
          if (colKam !== -1 && dataAsig[i][colKam]) name = String(dataAsig[i][colKam]).trim();
          if (colMail !== -1 && dataAsig[i][colMail]) email = String(dataAsig[i][colMail]).trim();
          if (colInic !== -1 && dataAsig[i][colInic]) initials = String(dataAsig[i][colInic]).trim().toUpperCase();
          break;
        }
      }
    }
  }

  // Fallback a Iniciales de agente si no estuviese en Asignacion_KAM
  if (shAgentes && (!initials || initials === "SC" || !email || email === "sara.cardona@isdin.com")) {
    const dataAg = shAgentes.getDataRange().getValues();
    if (dataAg.length >= 2) {
      const headAg = dataAg[0].map(h => cen_normalizeKey_(h));
      const colNameAg = headAg.findIndex(h => h.includes("NOMBRE") || h.includes("AGENTE") || h.includes("KAM")) !== -1 
                        ? headAg.findIndex(h => h.includes("NOMBRE") || h.includes("AGENTE") || h.includes("KAM")) 
                        : 0;
      const colInicAg = headAg.findIndex(h => h.includes("INICIAL") || h.includes("CODIGO") || h.includes("AGENTE"));
      const colMailAg = headAg.findIndex(h => h.includes("CORREO") || h.includes("EMAIL") || h.includes("MAIL"));

      const normName = cen_normalizeKey_(name);

      for (let i = 1; i < dataAg.length; i++) {
        const nameAg = cen_normalizeKey_(dataAg[i][colNameAg]);
        if (nameAg && (normName.includes(nameAg) || nameAg.includes(normName))) {
          if (colInicAg !== -1 && dataAg[i][colInicAg]) initials = String(dataAg[i][colInicAg]).trim().toUpperCase();
          if (colMailAg !== -1 && dataAg[i][colMailAg]) email = String(dataAg[i][colMailAg]).trim();
          break;
        }
      }
    }
  }

  return { name: name, initials: initials, email: email };
}

/**
 * Construye el cuerpo del correo HTML formateado y notifica a la KAM encargada
 * con el desglose de productos, subtotales e indicadores de Back Order.
 * 
 * @param {Array} pedidosCli Lista de pedidos pertenecientes al cliente procesado.
 * @param {Object} paquete Datos de formulario recolectados en el cliente.
 * @param {Object} kamInfo Información del KAM (name, initials, email).
 * @param {String} clienteNombre Nombre para mostrar del cliente.
 */
function cen_enviarCorreoKam_(pedidosCli, paquete, kamInfo, clienteNombre) {
  let htmlBody = `
    <div style="font-family: 'Inter', sans-serif; color: #2B2830; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
      <div style="text-align: center; border-bottom: 2px solid #E41F33; padding-bottom: 15px; margin-bottom: 20px;">
        <h1 style="color: #2B2830; font-size: 24px; margin: 0; font-weight: 900; letter-spacing: 2px;">ISDIN</h1>
        <p style="color: #E41F33; font-size: 11px; font-weight: 700; margin: 5px 0 0; text-transform: uppercase; letter-spacing: 1px;">Notificación Automática de Órdenes B2B - CEN</p>
      </div>
      
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Estimado/a <strong>${kamInfo.name}</strong>,</p>
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Te informamos que se han consolidado de manera exitosa los nuevos pedidos de <strong>${clienteNombre}</strong> ingresados a través del portal CEN.</p>
      
      <h2 style="font-size: 16px; font-weight: 700; border-bottom: 1px solid #F3F4F6; padding-bottom: 8px; margin-top: 25px; color: #2B2830;">Detalle de Facturación por Orden de Compra</h2>
      <div style="margin-top: 15px;">
  `;

  pedidosCli.forEach(p => {
    const itemsPedido = paquete.items.filter(i => i.Pedido_ID === p.Pedido_ID);
    let hasBackOrder = false;
    let listHtml = "";
    let totalValorOk = 0;
    let totalValorBO = 0;

    itemsPedido.forEach(item => {
      const matSeleccionado = paquete.materialesSeleccionados[item.Item_ID];
      if (matSeleccionado && !matSeleccionado.includes("SIN_")) {
        const isBO = paquete.backordersSeleccionados[item.Item_ID] || false;
        if (isBO) hasBackOrder = true;

        const cant = parseInt(item.Cant) || 0;
        const price = parseFloat(item.Valor_Unitario) || 0;
        const sub = cant * price;

        if (isBO) {
          totalValorBO += sub;
        } else {
          totalValorOk += sub;
        }

        const formattedSub = "$" + Math.round(sub).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
        const subtotalText = sub > 0 ? ` | Subtotal: <strong>${formattedSub}</strong>` : "";

        listHtml += `
          <div style="padding: 10px 0; border-bottom: 1px dashed #F3F4F6; display: flex; justify-content: space-between; font-size: 12px;">
            <div style="flex: 1; min-width: 0; padding-right: 10px;">
              <span style="font-weight: bold; color: #111827; display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${item.Descripcion}</span>
              <span style="color: #6B7280; font-family: monospace; font-size: 11px;">EAN: ${item.Producto_Cod} | ${cant} ud.${subtotalText}</span>
            </div>
            <div style="text-align: right; white-space: nowrap;">
              ${isBO ? '<span style="background-color: #FEF3C7; color: #D97706; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">Back Order</span>' : '<span style="background-color: #D1FAE5; color: #059669; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">OK</span>'}
            </div>
          </div>
        `;
      }
    });

    if (listHtml === "") return;

    const statusLabel = hasBackOrder 
      ? '<span style="background-color: #FEF3C7; color: #D97706; border: 1px solid #FCD34D; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">Contiene Back Orders</span>' 
      : '<span style="background-color: #D1FAE5; color: #059669; border: 1px solid #6EE7B7; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">Completada</span>';

    let valueSummaryText = `Valor Estimado Facturar (OK): $${Math.round(totalValorOk).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
    if (totalValorBO > 0) {
      valueSummaryText += ` | Valor en Back Order: $${Math.round(totalValorBO).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
    }

    htmlBody += `
      <div style="background-color: #FAFAFA; border: 1px solid #F3F4F6; border-radius: 8px; padding: 15px; margin-bottom: 15px;">
        <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #E5E7EB; padding-bottom: 10px; margin-bottom: 10px;">
          <span style="font-weight: bold; font-size: 14px; color: #111827;">OC: ${p.Pedido_ID}</span>
          ${statusLabel}
        </div>
        ${listHtml}
        <div style="text-align: right; font-size: 13px; font-weight: bold; margin-top: 10px; color: #111827;">
          ${valueSummaryText}
        </div>
      </div>
    `;
  });

  htmlBody += `
      </div>
      <p style="font-size: 13px; color: #9CA3AF; margin-top: 25px; border-top: 1px solid #F3F4F6; padding-top: 15px; text-align: center;">Este es un correo automático generado por el Order Hub de ISDIN Colombia. Por favor no responder a este mensaje.</p>
    </div>
  `;

  MailApp.sendEmail({
    to: kamInfo.email,
    subject: `📦 Resumen Pedidos CEN (${clienteNombre}) - ${Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")}`,
    htmlBody: htmlBody
  });
}