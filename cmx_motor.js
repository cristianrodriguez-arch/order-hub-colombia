/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: INTEGRACIÓN Y PROCESAMIENTO DE ÓRDENES DE COMPRA - CMX S.A.S. (EXCEL)
 * ARCHIVO: cmx_motor.gs
 * ESTADO: PRODUCCIÓN - SOPORTE MÚLTIPLES LOTES Y ASUNTO HOMOLOGADO
 * =========================================================================================
 * 📝 DESCRIPCIÓN DETALLADA:
 * Este módulo procesa las órdenes de compra de CMX S.A.S. en formato Excel (.xlsx, .xlsm).
 * Incluye:
 * 1. Preservación de múltiples lotes/materiales de la hoja Bodega para la UI.
 * 2. Extracción de ciudad destino nativa ("LA ESTRELLA") desde Destinatarios.
 * 3. Asunto de correo homologado: 📦 Resumen Pedidos CMX - DD/MM/YYYY.
 * 4. Valorización en $ COP por ítem y orden.
 * =========================================================================================
 */

function preAnalizarCMX() {
  try {
    const configCli = CONFIG.CLIENTES["CMX"];
    if (!configCli) {
      throw new Error("No se encontró la configuración de CMX en CONFIG de main.gs.");
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    const pedidosProcesados = new Set();
    
    if (shConsolidado) {
      const dataCons = shConsolidado.getDataRange().getValues();
      const headCons = dataCons[0].map(h => cmx_normalizeKey_(h));
      const colOc = headCons.indexOf(cmx_normalizeKey_("Orden de Compra"));
      if (colOc !== -1) {
        for (let i = 1; i < dataCons.length; i++) {
          const ocVal = String(dataCons[i][colOc]).trim();
          if (ocVal) pedidosProcesados.add(ocVal);
        }
      }
    }

    // Carga de Destinatarios utilizando la función unificada de main.gs
    const diccDestinatarios = obtenerDiccionarioDestinatarios();
    const diccBodega = {};
    for (const key in diccDestinatarios) {
      diccBodega[key] = diccDestinatarios[key].idDestino;
    }

    const diccProductos = cmx_getProductosPreciosMap_();

    const folder = DriveApp.getFolderById(configCli.FOLDER_ID);
    const files = folder.getFiles();
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
      return { pedidosExtraidos: [], itemsExtraidos: [], errores: [], yaProcesados: [] };
    }

    const resultado = cmx_motorLecturaExcel(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios);

    if (resultado.pedidosExtraidos.length === 0 && resultado.errores.length > 0) {
      throw new Error("Fallo en la lectura de CMX. Detalle: " + resultado.errores.join(' | '));
    }

    return {
      pedidosExtraidos: resultado.pedidosExtraidos,
      itemsExtraidos: resultado.itemsExtraidos,
      errores: resultado.errores,
      yaProcesados: resultado.yaProcesados
    };

  } catch (err) {
    throw new Error(err.message);
  }
}

function cmx_motorLecturaExcel(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios) {
  const pedidosExtraidos = [];
  const itemsExtraidos = [];
  const errores = [];
  const yaProcesados = [];

  archivos.forEach((file) => {
    let tempFileId = null;
    try {
      tempFileId = cmx_convertirExcelASheet_(file);
      const tempSs = SpreadsheetApp.openById(tempFileId);
      
      let targetSheet = tempSs.getSheets()[0];
      const sheets = tempSs.getSheets();
      for (let s = 0; s < sheets.length; s++) {
        const sName = sheets[s].getName().toUpperCase().trim();
        if (sName === "FORMATO OC") {
          targetSheet = sheets[s];
          break;
        } else if (sName.includes("OC") || sName.includes("ORDEN") || sName.includes("FORMATO")) {
          targetSheet = sheets[s];
        }
      }

      const data = targetSheet.getDataRange().getValues();

      const { cabecera, items } = cmx_extraerDatosHeuristico_(data, diccBodega, diccProductos, configCli, file.getName(), diccDestinatarios);

      if (!cabecera.ordenCompra) {
        throw new Error(`No se detectó número de Orden de Compra en el archivo ${file.getName()}`);
      }

      const ocTrim = String(cabecera.ordenCompra).trim();

      if (pedidosProcesados.has(ocTrim)) {
        yaProcesados.push(ocTrim);
        return;
      }

      pedidosExtraidos.push({
        Pedido_ID: ocTrim,
        Fecha_Pedido: cabecera.fechaPedido || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        Fecha_Entrega: cabecera.fechaEntrega || "",
        Id_Destino: cabecera.idDestino || configCli.SOLICITANTE,
        Destinatario: cabecera.destinatario || configCli.NOMBRE_CLIENTE || "CMX S.A.S.",
        Ciudad: cabecera.ciudad || "COTA",
        Archivo_Nombre: file.getName(),
        Archivo_Id: file.getId()
      });

      items.forEach((item, index) => {
        itemsExtraidos.push({
          Item_ID: `${ocTrim}_${index + 1}`,
          Pedido_ID: ocTrim,
          Producto_Cod: item.codigo,
          Descripcion: item.descripcion,
          Cant: item.cant,
          Valor_Unitario: item.valorUnitario || 0,
          EAN_Mapeado: item.eanMapeado || "",
          Materiales_Disponibles: item.materialesDisponibles || []
        });
      });

    } catch (e) {
      errores.push(`${file.getName()}: ${e.message}`);
    } finally {
      if (tempFileId) {
        try {
          DriveApp.getFileById(tempFileId).setTrashed(true);
        } catch (ignore) {}
      }
    }
  });

  return { pedidosExtraidos, itemsExtraidos, errores, yaProcesados };
}

function cmx_convertirExcelASheet_(file) {
  const fileId = file.getId();
  const folderId = file.getParents().next().getId();
  
  const resource = {
    title: file.getName().replace(/\.[^/.]+$/, ""),
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }]
  };
  
  const tempFile = Drive.Files.insert(resource, file.getBlob());
  return tempFile.id;
}

function cmx_extraerDatosHeuristico_(data, diccBodega, diccProductos, configCli, fileName, diccDestinatarios) {
  let cabecera = {
    ordenCompra: "",
    fechaPedido: "",
    fechaEntrega: "",
    idDestino: configCli.SOLICITANTE || "11033482", 
    destinatario: configCli.NOMBRE_CLIENTE || "CMX S.A.S.",
    ciudad: "COTA"
  };

  if (diccDestinatarios) {
    const destinoBySolicitante = diccDestinatarios[configCli.SOLICITANTE] || diccDestinatarios["11033482"] || diccDestinatarios["BODEGA_CMX_PRINCIPAL"];
    if (destinoBySolicitante && destinoBySolicitante.ciudad) {
      cabecera.ciudad = destinoBySolicitante.ciudad;
      cabecera.idDestino = destinoBySolicitante.idDestino || cabecera.idDestino;
      cabecera.destinatario = destinoBySolicitante.nombreDestino || cabecera.destinatario;
    }
  }
  
  let items = [];
  let filaInicioTabla = -1;
  let colIndices = { sku: -1, desc: -1, cant: -1, ean: -1 };
  let maxTableScore = 0;

  for (let r = 0; r < Math.min(data.length, 30); r++) {
    const filaTexto = data[r].map(c => String(c).toUpperCase().trim());
    const filaUnida = filaTexto.join(" ");

    if (!cabecera.ordenCompra && (filaUnida.includes("ORDEN") || filaUnida.includes("PEDIDO") || filaUnida.includes("O.C") || filaUnida.includes("PO") || filaUnida.includes("N°"))) {
      let rawOc = cmx_buscarValorAlLado_O_Abajo_(data, r, ["OC", "O.C", "COMPRA", "PEDIDO", "ORDER", "PURCHASE", "PO", "N°", "NRO", "NOC", "NOOC"]);
      if (rawOc) {
        cabecera.ordenCompra = rawOc.replace(/^(OC|O\.C|PEDIDO|N°|NRO)\s*/i, "").trim();
      }
    }
    if (!cabecera.fechaPedido && (filaUnida.includes("FECHA") || filaUnida.includes("DATE"))) {
      cabecera.fechaPedido = cmx_buscarValorAlLado_O_Abajo_(data, r, ["FECHA", "DATE"]);
    }

    let score = 0;
    let tempCols = { sku: -1, desc: -1, cant: -1, ean: -1 };

    for (let c = 0; c < data[r].length; c++) {
      const celda = cmx_normalizeKey_(data[r][c]);
      if (!celda) continue;

      if (celda.includes("BARRAS") || celda.includes("EAN") || celda.includes("BARCODE")) {
        score += 100;
        tempCols.ean = c;
      } else if (celda.includes("CODIGO") || celda.includes("CODIGOS") || celda.includes("MATERIAL") || celda.includes("SKU")) {
        score += 50;
        if (tempCols.ean === -1) tempCols.ean = c;
        tempCols.sku = c;
      } else if (celda.includes("DESCRIPCION") || celda.includes("PRODUCTO") || celda.includes("REFERENCIA") || celda.includes("ITEM")) {
        score += 80;
        tempCols.desc = c;
      } else if (celda.includes("CANTIDAD") || celda.includes("CANT") || celda.includes("UNIDADES") || celda.includes("QTY")) {
        if (!celda.includes("FACTURADA") && !celda.includes("DESPACHADA")) {
          score += 80;
          tempCols.cant = c;
        }
      }
    }

    if (score > maxTableScore && tempCols.cant !== -1 && (tempCols.ean !== -1 || tempCols.sku !== -1)) {
      maxTableScore = score;
      filaInicioTabla = r;
      colIndices = tempCols;
    }
  }

  if (!cabecera.ordenCompra && fileName) {
    const matches = fileName.match(/\d+/);
    if (matches) cabecera.ordenCompra = matches[0];
  }

  if (filaInicioTabla === -1) {
    throw new Error("No se pudo detectar la tabla de productos.");
  }

  const stockBodega = cmx_getBodegaStockMap_();

  for (let r = filaInicioTabla + 1; r < data.length; r++) {
    if (!data[r].join("").trim()) continue; 

    let rawEan = colIndices.ean !== -1 ? String(data[r][colIndices.ean]).trim() : "";
    let descripcion = colIndices.desc !== -1 ? String(data[r][colIndices.desc]).trim() : "Producto Sin Descripción";
    let cantidad = colIndices.cant !== -1 ? parseInt(data[r][colIndices.cant]) || 0 : 0;

    if (cantidad <= 0) continue;
    if (descripcion.toUpperCase().includes("TOTAL") || rawEan.toUpperCase().includes("TOTAL")) continue;

    const resultEan = cmx_procesarYBuscarEan_(rawEan, stockBodega);

    let precioUnitario = 0;
    let oficialMaterial = "";
    const normEanMapeado = cmx_normalizeKey_(resultEan.eanMapeado);
    
    if (normEanMapeado && diccProductos[normEanMapeado]) {
      precioUnitario = diccProductos[normEanMapeado].precio;
      oficialMaterial = diccProductos[normEanMapeado].materialOficial;
    } else if (resultEan.materialesDisponibles.length > 0) {
      const materialSap = cmx_normalizeKey_(resultEan.materialesDisponibles[0].material);
      if (diccProductos[materialSap]) {
        precioUnitario = diccProductos[materialSap].precio;
        oficialMaterial = diccProductos[materialSap].materialOficial;
      }
    }

    // PRESERVACIÓN DE LOTES Y MATERIALES DE BODEGA:
    // Si la bodega tiene registros para este EAN, los conserva intactos con sus respectivos Material IDs y fechas
    let materialesFiltrados = resultEan.materialesDisponibles || [];
    if (materialesFiltrados.length > 0) {
      materialesFiltrados = materialesFiltrados.map(m => {
        return {
          material: m.material || oficialMaterial, // Prioriza el Material ID propio del lote en Bodega
          stock: m.stock,
          fechaDisplay: m.fechaDisplay,
          isExpired: m.isExpired
        };
      });
    } else if (oficialMaterial) {
      materialesFiltrados = [{
        material: oficialMaterial,
        stock: 0,
        fechaDisplay: "Sin fecha",
        isExpired: false
      }];
    }

    items.push({
      codigo: resultEan.eanMapeado || rawEan || "SIN_EAN",
      descripcion: descripcion,
      cant: cantidad,
      valorUnitario: precioUnitario,
      eanMapeado: resultEan.eanMapeado || "",
      materialesDisponibles: materialesFiltrados
    });
  }

  return { cabecera, items };
}

function cmx_procesarYBuscarEan_(rawVal, stockBodega) {
  let candidatos = [];
  if (!rawVal) return { eanMapeado: "", materialesDisponibles: [] };

  let partes = String(rawVal).split(/[\s,;]+/);
  
  partes.forEach(p => {
    let clean = cmx_limpiarEanDeVerdad_(p);
    if (clean && clean.length >= 7) {
      candidatos.push(clean);
    }
  });

  if (candidatos.length === 0) {
    return { eanMapeado: "", materialesDisponibles: [] };
  }

  for (let i = 0; i < candidatos.length; i++) {
    const cand = candidatos[i];
    const normCand = cmx_normalizeKey_(cand);
    const stock = stockBodega[normCand];
    if (stock && stock.length > 0) {
      return { eanMapeado: cand, materialesDisponibles: stock };
    }
  }

  const primerCandidato = candidatos[0];
  const normFirst = cmx_normalizeKey_(primerCandidato);
  return {
    eanMapeado: primerCandidato,
    materialesDisponibles: stockBodega[normFirst] || []
  };
}

function cmx_limpiarEanDeVerdad_(val) {
  if (val === null || val === undefined) return "";
  let s = String(val).trim();
  
  if (/e/i.test(s)) {
    let num = Number(s);
    if (!isNaN(num)) {
      s = num.toFixed(0);
    }
  }
  
  if (s.includes(".")) {
    s = s.split(".")[0];
  }
  
  s = s.replace(/\D/g, "");
  
  if (s.length > 13) {
    s = s.substring(0, 13);
  }
  
  return s;
}

function cmx_buscarValorAlLado_O_Abajo_(data, row, keywords) {
  for (let c = 0; c < data[row].length; c++) {
    const cellVal = cmx_normalizeKey_(data[row][c]);
    if (!cellVal) continue;

    const match = keywords.some(kw => cellVal === cmx_normalizeKey_(kw) || cellVal.includes(cmx_normalizeKey_(kw)));
    if (match) {
      for (let i = c + 1; i < data[row].length; i++) {
        let valDer = data[row][i];
        if (valDer !== "" && valDer !== null && valDer !== undefined) {
          if (valDer instanceof Date) {
            return Utilities.formatDate(valDer, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          return String(valDer).trim();
        }
      }
      if (row + 1 < data.length) {
        let valAbajo = data[row + 1][c];
        if (valAbajo !== "" && valAbajo !== null && valAbajo !== undefined) {
          if (valAbajo instanceof Date) {
            return Utilities.formatDate(valAbajo, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
          return String(valAbajo).trim();
        }
      }
    }
  }
  return "";
}

function cmx_normalizeKey_(text) {
  if (text === null || text === undefined) return "";
  return String(text)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]/g, "")
    .toUpperCase()
    .trim();
}

function cmx_getProductosPreciosMap_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const shProd = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  const dicc = {};
  
  if (!shProd) return dicc;
  
  const data = shProd.getDataRange().getValues();
  if (data.length < 2) return dicc;
  
  const head = data[0].map(h => cmx_normalizeKey_(h));
  
  let colEan = head.indexOf(cmx_normalizeKey_("EAN"));
  if (colEan === -1) {
    colEan = head.findIndex(h => h.includes("EAN") || h.includes("BARRA") || h.includes("CODE"));
  }
  
  let colMat = head.indexOf(cmx_normalizeKey_("Material Id"));
  if (colMat === -1) {
    colMat = head.findIndex(h => h.includes("MATERIAL") || h === "SAP" || h === "CODIGO");
  }
  
  let colLista = head.indexOf(cmx_normalizeKey_("Lista Cliente"));
  if (colLista === -1) {
    colLista = head.findIndex(h => h.includes("LISTA") || h.includes("CLIENTE"));
  }
  
  let colCosto = head.indexOf(cmx_normalizeKey_("Costo"));
  if (colCosto === -1) {
    colCosto = head.findIndex(h => h.includes("COSTO") || h.includes("PRECIO") || h.includes("VALOR"));
  }
  
  let colDesc = head.indexOf(cmx_normalizeKey_("Producto"));
  if (colDesc === -1) {
    colDesc = head.findIndex(h => h.includes("PRODUCTO") || h.includes("DESC") || h.includes("REF"));
  }
  
  if (colEan === -1 || colCosto === -1 || colMat === -1) return dicc;
  
  const targetListaKey = cmx_normalizeKey_(CONFIG.CLIENTES["CMX"].LISTA_PRECIOS);

  for (let i = 1; i < data.length; i++) {
    const listaVal = cmx_normalizeKey_(data[i][colLista]);
    if (listaVal !== targetListaKey) continue;
    
    const ean = String(data[i][colEan]).trim();
    const material = String(data[i][colMat]).trim();
    const costo = parseFloat(data[i][colCosto]) || 0;
    const desc = colDesc !== -1 ? String(data[i][colDesc]).trim() : "";
    
    const normEan = cmx_normalizeKey_(ean);
    const normMat = cmx_normalizeKey_(material);
    
    const productData = { 
      precio: costo, 
      descripcion: desc, 
      materialOficial: material, 
      eanOficial: ean 
    };
    
    if (normEan) {
      dicc[normEan] = productData;
    }
    if (normMat) {
      dicc[normMat] = productData;
    }
  }
  return dicc;
}

function cmx_getBodegaStockMap_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const shBodega = ss.getSheetByName(CONFIG.SHEET_BODEGA || "Bodega");
  const stockMap = {};
  
  if (!shBodega) return stockMap;
  
  const data = shBodega.getDataRange().getValues();
  if (data.length < 2) return stockMap;
  
  const head = data[0].map(h => cmx_normalizeKey_(h));
  
  let colEan = 9; 

  let colMat = head.indexOf("MATERIAL");
  if (colMat === -1) colMat = head.indexOf("SAP");
  if (colMat === -1) {
    colMat = head.findIndex(h => h.includes("MATERIAL") || h.includes("SAP") || h === "CODIGO" || h === "REF");
  }
  if (colMat === -1) {
    colMat = 0; 
  }

  let colStock = head.findIndex(h => {
    return h.includes("STOCK") || h.includes("CANTIDAD") || h.includes("SALDO") || h.includes("DISP") || (h.includes("LIBRE") && h.includes("UTIL"));
  });
  if (colStock === -1) {
    colStock = 5; 
  }

  const colFecha = head.findIndex(h => h.includes("VENCIMIENTO") || h.includes("CADUCIDAD") || h.includes("FECHA") || h.includes("VENCE"));
  
  const hoy = new Date();
  
  for (let i = 1; i < data.length; i++) {
    let eanRaw = String(data[i][colEan]).trim();
    if (!eanRaw) continue;

    let eanLimpio = cmx_limpiarEanDeVerdad_(eanRaw);
    if (!eanLimpio) continue;
    
    const material = String(data[i][colMat]).trim();
    const stock = parseInt(data[i][colStock]) || 0;
    
    let fechaDisplay = "Sin fecha";
    let isExpired = false;
    
    if (colFecha !== -1 && data[i][colFecha]) {
      const fVal = data[i][colFecha];
      if (fVal instanceof Date) {
        isExpired = fVal < hoy;
        fechaDisplay = Utilities.formatDate(fVal, Session.getScriptTimeZone(), "yyyy-MM");
      } else {
        fechaDisplay = String(fVal).trim();
      }
    }
    
    const normEan = cmx_normalizeKey_(eanLimpio);
    if (!stockMap[normEan]) stockMap[normEan] = [];
    stockMap[normEan].push({
      material: material,
      stock: stock,
      fechaDisplay: fechaDisplay,
      isExpired: isExpired
    });
  }
  return stockMap;
}

function cmx_getCiudadParaCliente_(solicitanteId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const shClientes = ss.getSheetByName("Clientes") || ss.getSheetByName("Listado clientes") || ss.getSheetByName("Listado_clientes") || ss.getSheetByName("Listado Clientes");
  if (!shClientes) return "COTA";
  
  const data = shClientes.getDataRange().getValues();
  const head = data[0].map(h => cmx_normalizeKey_(h));
  
  const colPob = head.findIndex(h => h.includes("POBLACION") || h.includes("CIUDAD") || h.includes("MUNICIPIO") || h.includes("PBL"));
  if (colPob === -1) return "COTA";

  const searchKey = cmx_normalizeKey_(solicitanteId);

  for (let i = 1; i < data.length; i++) {
    let filaCompletaTexto = data[i].map(c => cmx_normalizeKey_(c));
    if (filaCompletaTexto.includes(searchKey)) {
      return String(data[i][colPob]).trim().toUpperCase() || "COTA";
    }
  }
  return "COTA";
}

function cmx_getKamInitialsAndName_(solicitanteId) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const shAsignacion = ss.getSheetByName("Asignacion_KAM");
  const shAgentes = ss.getSheetByName("Iniciales de agente");
  
  let kamName = "SARA CARDONA"; 
  let kamInitials = "SC"; 
  
  const normSolicitante = cmx_normalizeKey_(solicitanteId);

  if (shAsignacion) {
    const dataAsig = shAsignacion.getDataRange().getValues();
    const headAsig = dataAsig[0].map(h => cmx_normalizeKey_(h));
    const cSap = headAsig.indexOf(cmx_normalizeKey_("SAP ID"));
    const cKam = headAsig.findIndex(h => h.includes("KAMENCARGADO") || h === "KAM" || h.includes("ENCARGADO"));
    
    if (cSap !== -1 && cKam !== -1) {
      for (let i = 1; i < dataAsig.length; i++) {
        if (cmx_normalizeKey_(dataAsig[i][cSap]) === normSolicitante) {
          kamName = String(dataAsig[i][cKam]).trim();
          break;
        }
      }
    }
  }

  if (shAgentes && kamName) {
    const dataAg = shAgentes.getDataRange().getValues();
    const headAg = dataAg[0].map(h => cmx_normalizeKey_(h));
    const cInic = headAg.findIndex(h => h.includes("INICIAL") || h.includes("AGENTE") || h.includes("KAM") || h.includes("NOMBRE"));
    
    const normKamName = cmx_normalizeKey_(kamName);
    if (cInic !== -1) {
      for (let i = 1; i < dataAg.length; i++) {
        const nameVal = cmx_normalizeKey_(dataAg[i][0]); 
        const inicVal = String(dataAg[i][cInic]).trim().toUpperCase();
        if (nameVal && (normKamName.includes(nameVal) || nameVal.includes(normKamName))) {
          kamInitials = inicVal;
          break;
        }
      }
    }
  }
  
  return { name: kamName, initials: kamInitials };
}

function guardarDefinitivoCMX(paquete) {
  try {
    const configCli = CONFIG.CLIENTES["CMX"];
    if (!configCli) throw new Error("No se encontró la configuración para CMX.");

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    if (!shConsolidado) throw new Error("No se encontró la hoja CONSOLIDADO.");

    const kamInfo = cmx_getKamInitialsAndName_(configCli.SOLICITANTE);
    const yy = String(new Date().getFullYear()).slice(-2); 
    const prefix = kamInfo.initials + "-" + yy; 

    let maxConsecutive = 0; 
    const dataCons = shConsolidado.getDataRange().getValues();
    for (let i = 1; i < dataCons.length; i++) {
      const idVal = String(dataCons[i][0]).trim();
      if (idVal.startsWith(prefix)) {
        const numPart = parseInt(idVal.replace(prefix, ""), 10) || 0;
        if (numPart > maxConsecutive) {
          maxConsecutive = numPart;
        }
      }
    }
    
    let nuevoIdNumero = maxConsecutive + 1;
    const nuevoIdStr = prefix + String(nuevoIdNumero).padStart(4, "0");

    const filasNuevas = [];
    
    paquete.pedidos.forEach(p => {
      const itemsPedido = paquete.items.filter(i => i.Pedido_ID === p.Pedido_ID);
      
      let totalUnidades = 0;
      let totalValor = 0;
      
      itemsPedido.forEach(item => {
        const matSeleccionado = paquete.materialesSeleccionados[item.Item_ID];
        if (matSeleccionado && !matSeleccionado.includes("SIN_")) {
          const cant = parseInt(item.Cant) || 0;
          const valUnit = parseFloat(item.Valor_Unitario) || 0;
          totalUnidades += cant;
          totalValor += (cant * valUnit);
        }
      });

      if (totalUnidades === 0) return;

      const fechaEnt = paquete.fechasEntregas[p.Pedido_ID] ? paquete.fechasEntregas[p.Pedido_ID].fecha : "";

      filasNuevas.push([
        nuevoIdStr,
        p.Pedido_ID,
        configCli.SOLICITANTE,
        configCli.NOMBRE_CLIENTE || "CMX S.A.S.",
        p.Fecha_Pedido || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        p.Id_Destino || "11033482", 
        p.Destinatario || "CMX S.A.S.",
        p.Ciudad || "LA ESTRELLA",
        fechaEnt,
        totalUnidades,
        totalValor,
        "", 
        "", 
        "" 
      ]);
    });

    if (filasNuevas.length > 0) {
      const lastRow = shConsolidado.getLastRow();
      shConsolidado.getRange(lastRow + 1, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);
    } else {
      throw new Error("No hay posiciones para inyectar al consolidado.");
    }

    const folderDestino = DriveApp.getFolderById(configCli.PROCESSED_FOLDER_ID);
    paquete.pedidos.forEach(p => {
      if (p.Archivo_Id) {
        try {
          const file = DriveApp.getFileById(p.Archivo_Id);
          file.moveTo(folderDestino);
          
          const originalName = file.getName();
          const extIdx = originalName.lastIndexOf(".");
          const ext = extIdx !== -1 ? originalName.substring(extIdx) : ".xlsm";
          file.setName("Pedido_" + nuevoIdStr + ext);
        } catch (ignore) {}
      }
    });

    const csvData = cmx_generarCsvSap_(paquete, nuevoIdStr);
    const fileName = `Pedido_${nuevoIdStr}.csv`;

    try {
      let csvFolder = null;
      const parentFolder = DriveApp.getFolderById(configCli.PROCESSED_FOLDER_ID);
      const subFolders = parentFolder.getFoldersByName("CSV");
      if (subFolders.hasNext()) {
        csvFolder = subFolders.next();
      } else {
        csvFolder = parentFolder.createFolder("CSV");
      }
      csvFolder.createFile(fileName, csvData, MimeType.PLAIN_TEXT);
    } catch (csvDriveErr) {
      try {
        const parentFolder = DriveApp.getFolderById(configCli.PROCESSED_FOLDER_ID);
        parentFolder.createFile(fileName, csvData, MimeType.PLAIN_TEXT);
      } catch (e) {
        console.log("Error al respaldar el CSV en Drive: " + e.message);
      }
    }

    if (paquete.notificarKAM) {
      try {
        cmx_enviarNotificacionKam_(paquete, nuevoIdStr, paquete.pedidos[0].Ciudad || "LA ESTRELLA");
      } catch (eMailError) {
        console.log("Fallo el envío del correo: " + eMailError.message);
      }
    }

    return {
      success: true,
      mensaje: `Se han procesado ${filasNuevas.length} pedidos. Código de consolidación: ${nuevoIdStr}`,
      csvData: Utilities.base64Encode(csvData, Utilities.Charset.UTF_8),
      fileName: fileName
    };

  } catch (err) {
    throw new Error(err.message);
  }
}

function cmx_generarCsvSap_(paquete, nuevoIdStr) {
  const configCli = CONFIG.CLIENTES["CMX"];
  const sap = configCli.SAP || { CLASE_PEDIDO: "ZDIR", ORG_VENTAS: "Z011", CANAL: "10", SECTOR: "00" };
  
  // Filas de 16 columnas (Columna A a Columna P), sin fila de encabezado: el cargador SAP no la espera.
  const rows = [];
  let orderIndex = 0;

  paquete.pedidos.forEach((pedido) => {
    orderIndex++;
    const itemsPedido = paquete.items.filter(i => i.Pedido_ID === pedido.Pedido_ID);
    let posicionContador = 10;
    let esPrimeraPosicion = true;

    itemsPedido.forEach((item) => {
      const matSeleccionado = paquete.materialesSeleccionados[item.Item_ID];
      if (matSeleccionado && !matSeleccionado.includes("SIN_")) {
        const fechaEnt = paquete.fechasEntregas[item.Pedido_ID] ? paquete.fechasEntregas[item.Pedido_ID].fecha : "";
        
        let fechaEntFormatted = "";
        if (fechaEnt) {
          const dateParts = fechaEnt.split("-");
          if (dateParts.length === 3) {
            fechaEntFormatted = `${dateParts[2]}/${dateParts[1]}/${dateParts[0]}`;
          } else {
            fechaEntFormatted = fechaEnt;
          }
        }

        // Fila limitada exactamente a 16 posiciones (índices 0 al 15)
        const rowData = new Array(16).fill("");

        rowData[0] = orderIndex; // Col A: Contador
        rowData[1] = "";         // Col B: Nº Pedido SAP

        if (esPrimeraPosicion) {
          rowData[2] = sap.CLASE_PEDIDO;  // Col C: Clase pedido
          rowData[3] = sap.ORG_VENTAS;    // Col D: Org ventas
          rowData[4] = sap.CANAL;         // Col E: Canal
          rowData[5] = sap.SECTOR;        // Col F: Sector
          rowData[6] = configCli.SOLICITANTE; // Col G: Solicitante
          rowData[7] = pedido.Id_Destino || "11033482"; // Col H: Destinatario Merc
          rowData[8] = nuevoIdStr;        // Col I: Num pedido cliente
          rowData[9] = fechaEntFormatted; // Col J: Fecha preferente Entrega
          esPrimeraPosicion = false;
        } else {
          for (let c = 2; c <= 9; c++) {
            rowData[c] = "";
          }
        }

        rowData[10] = "";               // Col K: Descuento ZTK1 (%)
        rowData[11] = "";               // Col L: Motivo pedido
        rowData[12] = posicionContador; // Col M: Posición pedido
        rowData[13] = matSeleccionado;  // Col N: Material
        rowData[14] = "";               // Col O: EAN
        rowData[15] = item.Cant;        // Col P: Cantidad

        rows.push(rowData);
        posicionContador += 10;
      }
    });
  });

  return rows.map(r => r.join(";")).join("\n");
}

function cmx_enviarNotificacionKam_(paquete, nuevoIdStr, ciudadDestino) {
  const kamInfo = cmx_getKamInitialsAndName_(CONFIG.CLIENTES["CMX"].SOLICITANTE);
  let kamEmail = "sara.cardona@isdin.com"; 

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const shAgentes = ss.getSheetByName("Iniciales de agente");
  if (shAgentes) {
    const dataAg = shAgentes.getDataRange().getValues();
    const headAg = dataAg[0].map(h => cmx_normalizeKey_(h));
    const cMail = headAg.findIndex(h => h.includes("MAIL") || h.includes("CORREO") || h.includes("EMAIL"));
    
    if (cMail !== -1) {
      const normKamName = cmx_normalizeKey_(kamInfo.name);
      for (let i = 1; i < dataAg.length; i++) {
        const nameVal = cmx_normalizeKey_(dataAg[i][0]);
        if (nameVal && (normKamName.includes(nameVal) || nameVal.includes(normKamName))) {
          kamEmail = String(dataAg[i][cMail]).trim();
          break;
        }
      }
    }
  }

  let htmlBody = `
    <div style="font-family: 'Inter', sans-serif; color: #2B2830; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
      <div style="text-align: center; border-bottom: 2px solid #E41F33; padding-bottom: 15px; margin-bottom: 20px;">
        <h1 style="color: #2B2830; font-size: 24px; margin: 0; font-weight: 900; letter-spacing: 2px;">ISDIN</h1>
        <p style="color: #E41F33; font-size: 11px; font-weight: 700; margin: 5px 0 0; text-transform: uppercase; letter-spacing: 1px;">Notificación Automática de Órdenes B2B</p>
      </div>
      
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Estimada <strong>${kamInfo.name}</strong>,</p>
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Te informamos que se han consolidado de manera exitosa los nuevos pedidos de <strong>CMX S.A.S.</strong> para la ciudad de <strong>${ciudadDestino}</strong>. El consecutivo asignado es el <span style="font-family: monospace; font-weight: bold; background-color: #F3F4F6; padding: 2px 6px; border-radius: 4px; border: 1px solid #E5E7EB;">${nuevoIdStr}</span>.</p>
      
      <h2 style="font-size: 16px; font-weight: 700; border-bottom: 1px solid #F3F4F6; padding-bottom: 8px; margin-top: 25px; color: #2B2830;">Detalle de Facturación por Orden de Compra</h2>
      <div style="margin-top: 15px;">
  `;

  paquete.pedidos.forEach(p => {
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

  // Asunto exactamente homologado al estándar de Medipiel con icono de caja y fecha actual
  MailApp.sendEmail({
    to: kamEmail,
    subject: `📦 Resumen Pedidos CMX - ${Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")}`,
    htmlBody: htmlBody
  });
}

function cmx_guardarNuevoDestinatario(dest) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shDest = ss.getSheetByName("Destinatarios");
    if (!shDest) throw new Error("No se encontró la hoja Destinatarios.");
    
    shDest.appendRow([
      dest.idOrigen,
      dest.nombreOrigen,
      dest.cruceCliente,
      dest.idDestino,
      dest.nombreDestino
    ]);
    
    return { success: true };
  } catch (err) {
    throw new Error("Fallo al guardar destinatario: " + err.message);
  }
}