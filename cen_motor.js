/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: INTEGRACIÓN Y PROCESAMIENTO MULTICLIENTE - CEN (CENTRO DE NEGOCIOS CARVAJAL)
 * ARCHIVO: cen_motor.gs
 * ESTADO: PRODUCCIÓN - ENVÍO DE CORREO KAM AUTÓNOMO Y RESILIENTE + SAP 12 DÍGITOS
 * =========================================================================================
 * 📝 NOVEDADES Y CORRECCIONES IMPLEMENTADAS:
 * 1. RESOLUCIÓN DE CORREO KAM RESILIENTE (DOBLE CAPA): `cen_getKamInfo_` ahora busca primero
 *    en `obtenerInfoKamPorSolicitante_()` y, si no existe o falla, busca de forma AUTÓNOMA
 *    directamente en la pestaña `Asignacion_KAM` de Google Sheets. El correo NUNCA queda vacío.
 * 2. COMPATIBILIDAD CON MATERIALES SAP DE 12 DÍGITOS: Elimina la restricción que descartaba
 *    códigos de Material SAP con más de 10 dígitos (ej: 690033624480 y 690032151480).
 * 3. DOBLE CATÁLOGO CON FALLBACK GLOBAL: Si un EAN existe en la hoja Productos (bajo cualquier
 *    tarifa), recupera su Material SAP oficial. Si no tiene saldo en Bodega, pasa a Back Order.
 * 4. ESCUDO ANTI-DUPLICADOS COMPUESTO: Evalúa la combinación única `${OC}_${SAPID}`.
 * 5. NOMBRE DEL CSV UNIFICADO POR FECHA: Estándar `Pedido_CEN_YYYYMMDD.csv`.
 * 6. FÓRMULA DE KAM PRESERVADA: Inyecta celda vacía ("") en CONSOLIDADO para proteger fórmulas.
 * 7. SANITIZACIÓN BASE64: Evita descargas vacías (0 bytes) eliminando saltos de línea (\r\n).
 * =========================================================================================
 */

/**
 * Punto de entrada para la pre-analítica de los archivos CEN en Google Drive.
 * 
 * @return {Object} Contenedor con clientesDetectados, pedidosExtraidos, itemsExtraidos y errores.
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
    
    if (shConsolidado && shConsolidado.getLastRow() >= 2) {
      const dataCons = shConsolidado.getDataRange().getValues();
      let colOcCons = 1;  // Columna B por defecto
      let colSapCons = 2; // Columna C por defecto
      let startRow = 1;

      // Escaneo de filas 1 a 5 para detectar la fila real de cabeceras
      for (let r = 0; r < Math.min(dataCons.length, 5); r++) {
        const rowKeys = dataCons[r].map(h => cen_normalizeKey_(h));
        const foundOc = rowKeys.findIndex(h => h.includes("ORDEN") || h.includes("OC") || h.includes("PEDIDO"));
        if (foundOc !== -1) {
          colOcCons = foundOc;
          const foundSap = rowKeys.findIndex(h => h.includes("SAP") || h.includes("SOLICITANTE"));
          if (foundSap !== -1) colSapCons = foundSap;
          startRow = r + 1;
          break;
        }
      }

      for (let i = startRow; i < dataCons.length; i++) {
        const ocVal = cen_limpiarOc_(dataCons[i][colOcCons]);
        const sapVal = cen_normalizeKey_(dataCons[i][colSapCons]);
        if (ocVal && sapVal) {
          pedidosProcesados.add(`${ocVal}_${sapVal}`);
        }
      }
    }

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
 * Convierte temporalmente los libros Excel de CEN y extrae cabeceras y posiciones de venta.
 */
function cen_motorLecturaExcel(archivos, diccBodega, configCli, pedidosProcesados, diccDestinatarios, ss) {
  const pedidosExtraidos = [];
  const itemsExtraidos = [];
  const errores = [];
  const yaProcesados = [];
  const mapaClientesCount = {};

  const shProd = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  const { catalogoPorTarifa, catalogoGlobal } = cen_getCatalogoProductosGlobal_(shProd);

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

        // Clave combinada: OC + SAP ID del Cliente
        const comboKey = `${ocRaw}_${cen_normalizeKey_(solicitanteSap)}`;

        if (pedidosProcesados.has(comboKey)) {
          if (!yaProcesados.includes(ocRaw)) yaProcesados.push(ocRaw);
          continue;
        }

        let pedidoObj = pedidosExtraidos.find(p => p.Pedido_ID === ocRaw && p.Cliente_Key === clienteKey);
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
        
        // BÚSQUEDA 1: En la tarifa asignada al cliente
        let productoEncontrado = catalogoPorTarifa[keyTarifa] ? catalogoPorTarifa[keyTarifa][eanItemRaw] : null;

        // BÚSQUEDA 2 (RED DE SEGURIDAD): Si no está en su tarifa pero existe en el catálogo maestro
        if (!productoEncontrado && catalogoGlobal[eanItemRaw]) {
          productoEncontrado = catalogoGlobal[eanItemRaw];
        }

        let eanMapeado = "";
        let valorUnitario = 0;
        let materialesDisponibles = [];
        let descripcionOficial = productoEncontrado && productoEncontrado.nombre ? productoEncontrado.nombre : descItemExcel;

        if (productoEncontrado && productoEncontrado.materialOficial) {
          eanMapeado = eanItemRaw;
          valorUnitario = productoEncontrado.precio || 0;
          const oficialMat = String(productoEncontrado.materialOficial).trim();

          const stockLotes = diccBodega[eanItemRaw] || [];
          if (stockLotes.length > 0) {
            materialesDisponibles = stockLotes.map(s => ({
              material: s.material || oficialMat,
              stock: s.stock,
              fechaDisplay: s.fechaDisplay
            }));
          } else {
            // BACK ORDER GARANTIZADO: Si no hay saldo en bodega, conserva su Material SAP oficial con stock 0
            materialesDisponibles = [{
              material: oficialMat,
              stock: 0,
              fechaDisplay: "Sin stock"
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

function cen_identificarClienteInterno_(eanComprador, nombreComprador, configCli) {
  const mapeo = configCli.MAPEO_CLIENTES || {};
  const normComprador = cen_normalizeKey_(nombreComprador);
  const normEan = cen_normalizeKey_(eanComprador);

  // 1. Búsqueda en MAPEO_CLIENTES
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

  // 2. Búsqueda dinámica en Asignacion_KAM
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shAsig = ss.getSheetByName("Asignacion_KAM");
    if (shAsig && shAsig.getLastRow() >= 2) {
      const dataAsig = shAsig.getDataRange().getValues();
      const headAsig = dataAsig[0].map(h => cen_normalizeKey_(h));
      const colSap = headAsig.findIndex(h => h.includes("SAP") || h.includes("ID"));
      const colNom = headAsig.findIndex(h => h.includes("CLIENTE") || h.includes("NOMBRE") || h.includes("SOLICITANTE"));

      if (colSap !== -1 && colNom !== -1) {
        for (let i = 1; i < dataAsig.length; i++) {
          const nomClienteFila = cen_normalizeKey_(dataAsig[i][colNom]);
          if (nomClienteFila && (normComprador.includes(nomClienteFila) || nomClienteFila.includes(normComprador))) {
            const sapId = String(dataAsig[i][colSap]).trim();
            return {
              key: cen_normalizeKey_(dataAsig[i][colNom]),
              nombreDisplay: String(dataAsig[i][colNom]).trim(),
              solicitante: sapId,
              listaPrecios: String(dataAsig[i][colNom]).trim()
            };
          }
        }
      }
    }
  } catch (eScan) {
    console.warn("Búsqueda en Asignacion_KAM falló: " + eScan.message);
  }

  // 3. Fallback de contingencia
  const cleanKey = normComprador.substring(0, 15) || "CLIENTE_CEN";
  return {
    key: cleanKey,
    nombreDisplay: nombreComprador || "CLIENTE CEN",
    solicitante: configCli.SOLICITANTE || "11000000",
    listaPrecios: nombreComprador || "Distribuidor"
  };
}

/**
 * Carga el catálogo en dos niveles:
 * 1. catalogoPorTarifa[tarifa][ean] -> Para precios por cliente.
 * 2. catalogoGlobal[ean] -> Red de seguridad para asegurar el Material SAP (hasta 12 dígitos).
 */
function cen_getCatalogoProductosGlobal_(shProd) {
  const catalogoPorTarifa = {};
  const catalogoGlobal = {};
  if (!shProd) return { catalogoPorTarifa, catalogoGlobal };

  const data = shProd.getDataRange().getValues();
  if (data.length < 2) return { catalogoPorTarifa, catalogoGlobal };

  const head = data[0].map(h => cen_normalizeKey_(h));
  const colEan = head.findIndex(h => h.includes("EAN") || h.includes("BARRA"));
  const colMat = head.findIndex(h => h.includes("MATERIAL") || h === "SAP");
  const colNombre = head.findIndex(h => h.includes("DESCRIPCION") || h.includes("NOMBRE") || h.includes("PRODUCTO"));
  const colLista = head.findIndex(h => h.includes("LISTA") || h.includes("CLIENTE"));
  const colCosto = head.findIndex(h => h.includes("COSTO") || h.includes("PRECIO") || h.includes("VALOR"));

  if (colEan === -1 || colMat === -1) return { catalogoPorTarifa, catalogoGlobal };

  for (let i = 1; i < data.length; i++) {
    const listaKey = colLista !== -1 ? cen_normalizeKey_(data[i][colLista]) : "GENERAL";
    const ean = cen_limpiarEan_(data[i][colEan]);
    const mat = String(data[i][colMat]).trim();
    const nombre = colNombre !== -1 ? String(data[i][colNombre]).trim() : "";
    const precio = colCosto !== -1 ? (parseFloat(data[i][colCosto]) || 0) : 0;

    if (!ean || !mat) continue;

    const prodInfo = { materialOficial: mat, nombre: nombre, precio: precio };

    // 1. Guardar en catálogo específico por tarifa
    if (listaKey) {
      if (!catalogoPorTarifa[listaKey]) catalogoPorTarifa[listaKey] = {};
      catalogoPorTarifa[listaKey][ean] = prodInfo;
    }

    // 2. Guardar en catálogo global como red de seguridad
    if (!catalogoGlobal[ean] || (precio > 0 && catalogoGlobal[ean].precio === 0)) {
      catalogoGlobal[ean] = prodInfo;
    }
  }

  return { catalogoPorTarifa, catalogoGlobal };
}

/**
 * Guarda las órdenes en el CONSOLIDADO, genera el CSV SAP unificado por fecha,
 * traslada archivos procesados y notifica a cada KAM.
 */
function guardarDefinitivoCEN(paquete) {
  try {
    const configCli = CONFIG.CLIENTES["CEN"];
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");

    if (!shConsolidado) throw new Error("No se encontró la hoja CONSOLIDADO.");

    const currentYearSuffix = String(new Date().getFullYear()).slice(-2);

    const filasConsolidado = [];
    const filasCSV = [];
    let contadorPedidoCSV = 1;

    const pedidosPorCliente = {};
    paquete.pedidos.forEach(p => {
      if (!pedidosPorCliente[p.Cliente_Key]) pedidosPorCliente[p.Cliente_Key] = [];
      pedidosPorCliente[p.Cliente_Key].push(p);
    });

    for (const cliKey in pedidosPorCliente) {
      const pedidosCli = pedidosPorCliente[cliKey];
      const primerPedido = pedidosCli[0];
      const solicitanteSap = primerPedido.Solicitante_SAP;

      // Consumo con doble capa de resolución (main.gs o búsqueda directa en hoja)
      const kamInfo = cen_getKamInfo_(solicitanteSap, primerPedido.Cliente_Nombre);
      const prefix = (kamInfo.initials || "MC") + "-" + currentYearSuffix;

      let maxConsecutive = cen_obtenerMaxConsecutivo_(shConsolidado, prefix);

      pedidosCli.forEach(p => {
        const itemsPedido = paquete.items.filter(i => i.Pedido_ID === p.Pedido_ID);
        let totalUnidades = 0;
        let totalValor = 0;
        let posicionItem = 10;

        itemsPedido.forEach(item => {
          let matSeleccionado = (paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID]) ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";
          
          // Validación flexible: Admite códigos SAP de 12 dígitos, descarta cadenas vacías, SIN_ y EANs de 13 dígitos
          const isNoActualizado = (paquete.noActualizadosSeleccionados && paquete.noActualizadosSeleccionados[item.Item_ID]) || 
                                  (!matSeleccionado || matSeleccionado.includes("SIN_") || matSeleccionado === item.Producto_Cod || (matSeleccionado.length === 13 && matSeleccionado.startsWith("84")));

          let matCsv = isNoActualizado ? "" : matSeleccionado;

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
            matCsv,
            "",
            cant
          ]);
          posicionItem += 10;
        });

        if (totalUnidades > 0) {
          maxConsecutive++;
          const consecutivoStr = String(maxConsecutive).padStart(4, "0");
          const idGenerado = prefix + consecutivoStr;

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
            "" // DEJAR VACÍO: Para que actúe la fórmula nativa de Google Sheets
          ]);
          contadorPedidoCSV++;
        }
      });

      // Disparo de correo al KAM con validación estricta y logs de trazabilidad
      if (paquete.notificarKAM !== false) {
        if (kamInfo.email && kamInfo.email.includes("@")) {
          try {
            cen_enviarCorreoKam_(pedidosCli, paquete, kamInfo, primerPedido.Cliente_Nombre);
            console.log(`Correo enviado con éxito a ${kamInfo.name} (${kamInfo.email}) para cliente ${primerPedido.Cliente_Nombre}`);
          } catch (eMailErr) {
            console.error(`Error enviando correo a ${kamInfo.email}: ${eMailErr.message}`);
          }
        } else {
          console.warn(`No se envió notificación de ${primerPedido.Cliente_Nombre} porque no se encontró un correo registrado para el Solicitante SAP ${solicitanteSap} en Asignacion_KAM.`);
        }
      }
    }

    if (filasConsolidado.length > 0) {
      const lastRow = shConsolidado.getLastRow();
      shConsolidado.getRange(lastRow + 1, 1, filasConsolidado.length, filasConsolidado[0].length).setValues(filasConsolidado);
    }

    if (configCli.PROCESSED_FOLDER_ID) {
      const folderDestino = DriveApp.getFolderById(configCli.PROCESSED_FOLDER_ID);
      const archivosProcesadosMap = {};

      paquete.pedidos.forEach(p => {
        if (p.Archivo_Id) {
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

    // Sin fila de encabezado: el cargador SAP no la espera.
    const csvContent = filasCSV.map(r => r.join(";")).join("\n");

    const fechaProceso = Utilities.formatDate(new Date(), "GMT-5", "yyyyMMdd");
    const fileName = `Pedido_CEN_${fechaProceso}.csv`;

    try {
      const folderCSV = DriveApp.getFolderById(CONFIG.FOLDER_CSV_ID);
      folderCSV.createFile(fileName, csvContent, MimeType.PLAIN_TEXT);
    } catch (eCsv) {
      console.error("Error guardando CSV en Drive: " + eCsv.message);
    }

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
 * Consulta la información del KAM asignado con DOBLE CAPA de resolución:
 * Capa 1: Invoca la función centralizada obtenerInfoKamPorSolicitante_() de main.gs.
 * Capa 2: Si no existe o falla, busca de forma autónoma en la pestaña Asignacion_KAM.
 * 
 * @param {String|Number} solicitanteSap ID de SAP del cliente.
 * @param {String} [clienteNombre] Nombre display del cliente para búsqueda por alias.
 * @return {Object} Objeto garantizado con { name, initials, email }.
 */
function cen_getKamInfo_(solicitanteSap, clienteNombre) {
  // Capa 1: Consumo de main.gs
  if (typeof obtenerInfoKamPorSolicitante_ === "function") {
    try {
      const infoKam = obtenerInfoKamPorSolicitante_(solicitanteSap);
      if (infoKam && infoKam.correo) {
        return {
          name: infoKam.nombre || "KAM ISDIN",
          initials: infoKam.iniciales || "MC",
          email: infoKam.correo
        };
      }
    } catch (eCentral) {
      console.warn("Llamada a obtenerInfoKamPorSolicitante_ en main.gs falló: " + eCentral.message);
    }
  }

  // Capa 2: Búsqueda autónoma directa en Asignacion_KAM
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shAsig = ss.getSheetByName("Asignacion_KAM");
    if (shAsig && shAsig.getLastRow() >= 2) {
      const dataAsig = shAsig.getDataRange().getValues();
      const headAsig = dataAsig[0].map(h => cen_normalizeKey_(h));
      
      const colSap = headAsig.findIndex(h => h.includes("SAP") || h.includes("ID") || h.includes("SOLICITANTE"));
      const colKam = headAsig.findIndex(h => h.includes("KAM") || h.includes("ENCARGADO") || h.includes("NOMBRE") || h.includes("RESPONSABLE"));
      const colMail = headAsig.findIndex(h => h.includes("CORREO") || h.includes("EMAIL") || h.includes("MAIL"));
      const colInic = headAsig.findIndex(h => h.includes("INICIAL") || h.includes("AGENTE") || h.includes("CODIGO"));
      const colNomCli = headAsig.findIndex(h => h.includes("CLIENTE") || h.includes("CUENTA"));

      const normTargetSap = cen_normalizeKey_(solicitanteSap);
      const normTargetNom = clienteNombre ? cen_normalizeKey_(clienteNombre) : "";

      for (let i = 1; i < dataAsig.length; i++) {
        const sapVal = colSap !== -1 ? cen_normalizeKey_(dataAsig[i][colSap]) : "";
        const nomVal = colNomCli !== -1 ? cen_normalizeKey_(dataAsig[i][colNomCli]) : "";

        const matchSap = sapVal && sapVal === normTargetSap;
        const matchNom = normTargetNom && nomVal && (normTargetNom.includes(nomVal) || nomVal.includes(normTargetNom));

        if (matchSap || matchNom) {
          const kamMail = colMail !== -1 ? String(dataAsig[i][colMail]).trim() : "";
          if (kamMail) {
            return {
              name: colKam !== -1 ? String(dataAsig[i][colKam]).trim() : "KAM ISDIN",
              initials: colInic !== -1 ? String(dataAsig[i][colInic]).trim().toUpperCase() : "MC",
              email: kamMail
            };
          }
        }
      }
    }
  } catch (eDirect) {
    console.warn("Búsqueda autónoma en Asignacion_KAM falló: " + eDirect.message);
  }

  // Capa 3: Fallback seguro
  return {
    name: "KAM ISDIN",
    initials: "MC",
    email: ""
  };
}

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
    let itemsNoActualizados = [];
    let itemsValidos = [];
    let totalValorOk = 0;
    let totalValorBO = 0;
    let hasBackOrder = false;

    itemsPedido.forEach(item => {
      let mat = (paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID]) ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";
      
      // Validación sin rechazar materiales SAP de 12 dígitos
      let isNoActualizado = (paquete.noActualizadosSeleccionados && paquete.noActualizadosSeleccionados[item.Item_ID]) || 
                            (!mat || mat.includes("SIN_") || mat === item.Producto_Cod || (mat.length === 13 && mat.startsWith("84")));
      let isBO = !isNoActualizado && (paquete.backordersSeleccionados && paquete.backordersSeleccionados[item.Item_ID]);

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
      statusLabel = '<span style="background-color: #D1FAE5; color: #059669; border: 1px solid #6EE7B7; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">✅ Completada</span>';
    }

    let valueSummaryText = `Valor Estimado Facturar (OK): $${Math.round(totalValorOk).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
    if (totalValorBO > 0) {
      valueSummaryText += ` | Valor en Back Order: $${Math.round(totalValorBO).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
    }
    if (itemsNoActualizados.length > 0) {
      valueSummaryText += ` <br><span style="color:#B91C1C; font-size:11px; margin-top:4px; display:block;"><strong>Alerta Comercial:</strong> Hay ${itemsNoActualizados.length} ítem(s) que requieren actualización de catálogo.</span>`;
    }

    htmlBody += `
      <div style="background-color: #FAFAFA; border: 1px solid #F3F4F6; border-radius: 8px; padding: 15px; margin-bottom: 15px;">
        <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #E5E7EB; padding-bottom: 10px; margin-bottom: 10px;">
          <span style="font-weight: bold; font-size: 14px; color: #111827;">OC: ${p.Pedido_ID}</span>
          ${statusLabel}
        </div>
        ${alertaNoActualizadosHtml}
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