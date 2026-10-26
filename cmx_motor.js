/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: INTEGRACIÓN Y PROCESAMIENTO DE ÓRDENES DE COMPRA - CMX S.A.S. (EXCEL)
 * ARCHIVO: cmx_motor.gs
 * ESTADO: PRODUCCIÓN - SOPORTE MATERIALES SAP DE 12 DÍGITOS Y KAM DESACOPLADO
 * =========================================================================================
 * 📝 CORRECCIONES APLICADAS:
 * 1. ELIMINACIÓN DE LA RESTRICCIÓN DE 10 DÍGITOS: Tanto en `cmx_generarCsvSap_` como en
 *    `cmx_enviarNotificacionKam_`, se eliminó la condición `length > 10` que borraba
 *    los materiales de 12 dígitos (ej: 690032151480 / Ureadin Cream10).
 * 2. ESTRUCTURA CSV HOMOLOGADA: 16 columnas para SAP oficial (idéntico a Medipiel y CEN).
 * 3. KAM 100% DESACOPLADO: Utiliza `obtenerInfoKamPorSolicitante_()` de main.gs.
 * 4. ESCUDO ANTI-DUPLICADOS COMPUESTO: Evalúa la combinación `${OC}_${SAPID}`.
 * 5. DOBLE CATÁLOGO CON RED DE SEGURIDAD: Precios CMX con fallback al catálogo maestro.
 * =========================================================================================
 */

function preAnalizarCMX() {
  try {
    const configCli = CONFIG.CLIENTES["CMX"];
    if (!configCli) {
      throw new Error("No se encontró la configuración del cliente CMX en CONFIG de main.gs.");
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    
    // -------------------------------------------------------------------------------------
    // PASO 1: DETECTOR DE DUPLICADOS EN EL CONSOLIDADO (CLAVE COMPUESTA OC + SAP ID)
    // -------------------------------------------------------------------------------------
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    const pedidosProcesados = new Set();
    
    if (shConsolidado && shConsolidado.getLastRow() >= 2) {
      const dataCons = shConsolidado.getDataRange().getValues();
      let colOc = 1;  // Columna B por defecto
      let colSap = 2; // Columna C por defecto
      let startRow = 1;

      for (let r = 0; r < Math.min(dataCons.length, 5); r++) {
        const rowKeys = dataCons[r].map(h => cmx_normalizeKey_(h));
        const foundOc = rowKeys.findIndex(h => h.includes("ORDEN") || h.includes("OC") || h.includes("PEDIDO"));
        if (foundOc !== -1) {
          colOc = foundOc;
          const foundSap = rowKeys.findIndex(h => h.includes("SAP") || h.includes("SOLICITANTE"));
          if (foundSap !== -1) colSap = foundSap;
          startRow = r + 1;
          break;
        }
      }

      for (let i = startRow; i < dataCons.length; i++) {
        const ocVal = cmx_limpiarOc_(dataCons[i][colOc]);
        const sapVal = cmx_normalizeKey_(dataCons[i][colSap]);
        if (ocVal && sapVal) {
          pedidosProcesados.add(`${ocVal}_${sapVal}`);
        }
      }
    }

    // -------------------------------------------------------------------------------------
    // PASO 2: MAPEO DE DESTINATARIOS PARA DIRECCIONAMIENTO GEOGRÁFICO DE ENTREGAS
    // -------------------------------------------------------------------------------------
    const shDest = ss.getSheetByName(CONFIG.SHEET_DESTINATARIOS || "Destinatarios");
    const diccDestinatarios = {};
    const diccBodega = {}; 
    
    if (shDest) {
      const dataDest = shDest.getDataRange().getValues();
      const headDest = dataDest[0].map(h => cmx_normalizeKey_(h));
      const colCruce = headDest.indexOf(cmx_normalizeKey_("Cruce_cliente"));
      const colIdDest = headDest.indexOf(cmx_normalizeKey_("Id_destino"));
      const colNomDest = headDest.indexOf(cmx_normalizeKey_("Nombre_destino"));
      const colCiudad = headDest.findIndex(h => h === "CIUDAD" || h.includes("POBLACION") || h.includes("MUNICIPIO"));
      
      if (colCruce !== -1 && colIdDest !== -1) {
        for (let i = 1; i < dataDest.length; i++) {
          const cruceVal = String(dataDest[i][colCruce]).trim();
          if (cruceVal) {
            diccDestinatarios[cruceVal] = {
              idDestino: String(dataDest[i][colIdDest]).trim(),
              nombreDestino: colNomDest !== -1 ? String(dataDest[i][colNomDest]).trim() : "",
              ciudad: colCiudad !== -1 ? String(dataDest[i][colCiudad]).trim().toUpperCase() : "COTA"
            };
            diccBodega[cruceVal] = String(dataDest[i][colIdDest]).trim();
          }
        }
      }
    }

    // -------------------------------------------------------------------------------------
    // PASO 3: CARGA DE PRECIOS B2B DE CMX DESDE EL MAIN (100% DESACOPLADO)
    // -------------------------------------------------------------------------------------
    const diccProductos = cmx_getProductosPreciosMap_();

    // -------------------------------------------------------------------------------------
    // PASO 4: RECOLECCIÓN DE ARCHIVOS EXCEL DESDE GOOGLE DRIVE
    // -------------------------------------------------------------------------------------
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

    // -------------------------------------------------------------------------------------
    // PASO 5: MOTOR DE PROCESAMIENTO EN CALIENTE DE LIBROS EXCEL
    // -------------------------------------------------------------------------------------
    const resultado = cmx_motorLecturaExcel(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios);

    if (resultado.pedidosExtraidos.length === 0 && resultado.errores.length > 0) {
      throw new Error("Fallo en la lectura del archivo Excel CMX. Detalle técnico: " + resultado.errores.join(' | '));
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

      const ocTrim = cmx_limpiarOc_(cabecera.ordenCompra);
      const targetSap = cmx_normalizeKey_(configCli.SOLICITANTE || "11033482");
      const comboKey = `${ocTrim}_${targetSap}`;

      // ESCUDO ANTI-DUPLICADOS: Solo bloquea si coincide OC Y Solicitante de CMX
      if (pedidosProcesados.has(comboKey)) {
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

function cmx_limpiarOc_(val) {
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
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  
  let cabecera = {
    ordenCompra: "",
    fechaPedido: "",
    fechaEntrega: "",
    idDestino: configCli.SOLICITANTE || "11033482", 
    destinatario: configCli.NOMBRE_CLIENTE || "CMX S.A.S.",
    ciudad: "COTA"
  };
  
  let items = [];
  let filaInicioTabla = -1;
  let colIndices = { sku: -1, desc: -1, cant: -1, ean: -1 };
  let maxTableScore = 0;

  // 1. ESCANEO GEOGRÁFICO DE CABECERA (Filas 1 a 12 - Todas las columnas)
  let textoDireccion = "";
  for (let r = 0; r < Math.min(data.length, 12); r++) {
    for (let c = 0; c < data[r].length; c++) {
      const cellVal = String(data[r][c]).trim();
      if (cellVal) {
        textoDireccion += " " + cellVal;
      }
    }
  }

  let normTexto = cmx_normalizeKey_(textoDireccion);
  let ciudadMapeadaDestinatarios = "";
  if (diccDestinatarios) {
    for (const key in diccDestinatarios) {
      let normKey = cmx_normalizeKey_(key);
      if (normKey && normTexto.includes(normKey)) {
        cabecera.idDestino = diccDestinatarios[key].idDestino;
        cabecera.destinatario = diccDestinatarios[key].nombreDestino;
        ciudadMapeadaDestinatarios = diccDestinatarios[key].ciudad;
        break;
      }
    }
  }
  
  let ciudadDetectada = "";
  const shClientes = ss.getSheetByName("Clientes") || ss.getSheetByName("Listado clientes") || ss.getSheetByName("Listado_clientes") || ss.getSheetByName("Listado Clientes");
  if (shClientes) {
    const dataCli = shClientes.getDataRange().getValues();
    const headCli = dataCli[0].map(h => cmx_normalizeKey_(h));
    const colPob = headCli.findIndex(h => h.includes("POBLACION") || h.includes("CIUDAD") || h.includes("MUNICIPIO") || h.includes("PBL"));
    if (colPob !== -1) {
      let normB2 = cmx_normalizeKey_(textoDireccion);
      for (let i = 1; i < dataCli.length; i++) {
        let pobVal = String(dataCli[i][colPob]).trim().toUpperCase();
        if (pobVal && pobVal.length > 3) {
          let normPob = cmx_normalizeKey_(pobVal);
          if (normB2.includes(normPob)) {
            ciudadDetectada = pobVal;
            break;
          }
        }
      }
    }
  }

  cabecera.ciudad = ciudadMapeadaDestinatarios || ciudadDetectada || cmx_getCiudadParaCliente_(cabecera.idDestino) || "COTA";

  // 2. Escanear fila por fila para ubicar cabecera e inicio de la tabla
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
    throw new Error("No se pudo detectar la tabla de productos (columna de Código de barras y Cantidad) en el archivo.");
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

    let materialesFiltrados = resultEan.materialesDisponibles || [];
    if (oficialMaterial) {
      if (materialesFiltrados.length > 0) {
        materialesFiltrados = materialesFiltrados.map(m => {
          return {
            material: oficialMaterial,
            stock: m.stock,
            fechaDisplay: m.fechaDisplay,
            isExpired: m.isExpired
          };
        });
      } else {
        materialesFiltrados = [{
          material: oficialMaterial,
          stock: 0,
          fechaDisplay: "Sin fecha",
          isExpired: false
        }];
      }
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
    if (stock && stock.length > 0 && stock.some(s => s.stock > 0)) {
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

/**
 * Carga el catálogo en 2 fases:
 * Fase 1: Carga el catálogo global de Productos como red de seguridad.
 * Fase 2: Sobreescribe con máxima prioridad los precios de la lista de CMX.
 */
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
  
  const targetListaKey = cmx_normalizeKey_(CONFIG.CLIENTES["CMX"].LISTA_PRECIOS || "CMX\nCADA PIEL\nThe Beauty club");

  // Fase 1: Red de seguridad con catálogo global
  for (let i = 1; i < data.length; i++) {
    const ean = String(data[i][colEan]).trim();
    const material = String(data[i][colMat]).trim();
    const costo = parseFloat(data[i][colCosto]) || 0;
    const desc = colDesc !== -1 ? String(data[i][colDesc]).trim() : "";
    
    const normEan = cmx_normalizeKey_(ean);
    const normMat = cmx_normalizeKey_(material);
    
    if (normEan && !dicc[normEan]) {
      dicc[normEan] = { precio: costo, descripcion: desc, materialOficial: material, eanOficial: ean };
    }
    if (normMat && !dicc[normMat]) {
      dicc[normMat] = { precio: costo, descripcion: desc, materialOficial: material, eanOficial: ean };
    }
  }

  // Fase 2: Sobreescribir con la tarifa específica de CMX
  for (let i = 1; i < data.length; i++) {
    const listaVal = cmx_normalizeKey_(data[i][colLista]);
    if (listaVal !== targetListaKey && !targetListaKey.includes(listaVal)) continue;
    
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
    
    if (normEan) dicc[normEan] = productData;
    if (normMat) dicc[normMat] = productData;
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
    const normMat = cmx_normalizeKey_(material);

    const itemObj = {
      material: material,
      stock: stock,
      fechaDisplay: fechaDisplay,
      isExpired: isExpired
    };

    if (!stockMap[normEan]) stockMap[normEan] = [];
    stockMap[normEan].push(itemObj);

    if (normMat && normMat !== normEan) {
      if (!stockMap[normMat]) stockMap[normMat] = [];
      stockMap[normMat].push(itemObj);
    }
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

/**
 * Consulta la información del KAM asignado consumiendo directamente
 * la función centralizada obtenerInfoKamPorSolicitante_() de main.gs.
 * 
 * @param {String|Number} solicitanteId ID de SAP del solicitante.
 * @return {Object} Objeto con name, initials y email.
 */
function cmx_getKamInitialsAndName_(solicitanteId) {
  if (typeof obtenerInfoKamPorSolicitante_ === "function") {
    try {
      const infoKam = obtenerInfoKamPorSolicitante_(solicitanteId);
      if (infoKam) {
        return {
          name: infoKam.nombre,
          initials: infoKam.iniciales,
          email: infoKam.correo
        };
      }
    } catch (e) {
      console.warn("Llamada a obtenerInfoKamPorSolicitante_ falló: " + e.message);
    }
  }

  return {
    name: "KAM ISDIN",
    initials: "MC",
    email: ""
  };
}

function guardarDefinitivoCMX(paquete) {
  try {
    const configCli = CONFIG.CLIENTES["CMX"];
    if (!configCli) throw new Error("No se encontró la configuración para CMX.");

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    if (!shConsolidado) throw new Error("No se encontró la hoja CONSOLIDADO.");

    // ESCUDO EN TIEMPO REAL: Re-verifica antes de escribir en CONSOLIDADO
    const dataConsCheck = shConsolidado.getDataRange().getValues();
    let colOcCheck = 1;
    let colSapCheck = 2;

    for (let r = 0; r < Math.min(dataConsCheck.length, 5); r++) {
      const rKeys = dataConsCheck[r].map(h => cmx_normalizeKey_(h));
      const fOc = rKeys.findIndex(h => h.includes("ORDEN") || h.includes("OC") || h.includes("PEDIDO"));
      if (fOc !== -1) {
        colOcCheck = fOc;
        const fSap = rKeys.findIndex(h => h.includes("SAP") || h.includes("SOLICITANTE"));
        if (fSap !== -1) colSapCheck = fSap;
        break;
      }
    }

    const ocsExistentes = new Set();
    const targetSapCheck = cmx_normalizeKey_(configCli.SOLICITANTE || "11033482");

    for (let i = 1; i < dataConsCheck.length; i++) {
      const ocVal = cmx_limpiarOc_(dataConsCheck[i][colOcCheck]);
      const sapVal = cmx_normalizeKey_(dataConsCheck[i][colSapCheck]);
      if (ocVal && sapVal) {
        ocsExistentes.add(`${ocVal}_${sapVal}`);
      }
    }

    const pedidosFiltrados = paquete.pedidos.filter(p => {
      const ocTrim = cmx_limpiarOc_(p.Pedido_ID);
      return !ocsExistentes.has(`${ocTrim}_${targetSapCheck}`);
    });

    if (pedidosFiltrados.length === 0) {
      throw new Error("La Orden de Compra elegida ya se encuentra registrada para CMX S.A.S. en el Consolidado.");
    }

    paquete.pedidos = pedidosFiltrados;

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
        let matSeleccionado = (paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID]) ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";
        
        // Validación segura sin rechazar materiales SAP de 12 dígitos
        const isNoActualizado = (paquete.noActualizadosSeleccionados && paquete.noActualizadosSeleccionados[item.Item_ID]) || 
                                (!matSeleccionado || matSeleccionado.includes("SIN_") || matSeleccionado === item.Producto_Cod || (matSeleccionado.length === 13 && matSeleccionado.startsWith("84")));

        if (!isNoActualizado) {
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
        "" // Dejado en blanco para no sobreescribir la fórmula
      ]);
    });

    if (filasNuevas.length > 0) {
      const lastRow = shConsolidado.getLastRow();
      shConsolidado.getRange(lastRow + 1, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);
    } else {
      throw new Error("No hay posiciones válidas para inyectar al consolidado.");
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
      csvData: Utilities.base64Encode(csvData, Utilities.Charset.UTF_8).replace(/[\r\n]/g, ""),
      fileName: fileName
    };

  } catch (err) {
    throw new Error(err.message);
  }
}

/**
 * Genera el archivo plano oficial de 16 columnas para SAP.
 * Acepta códigos SAP oficiales de hasta 12 dígitos (ej: 690032151480).
 */
function cmx_generarCsvSap_(paquete, nuevoIdStr) {
  const configCli = CONFIG.CLIENTES["CMX"];
  const sap = configCli.SAP || { CLASE_PEDIDO: "ZDIR", ORG_VENTAS: "Z011", CANAL: "10", SECTOR: "00" };
  
  const headers = [
    "Contador", "Nº Pedido SAP", "Clase pedido", "Org ventas", "Canal", "Sector", "Solicitante", "Destinatario Merc",
    "Num pedido cliente", "Fecha preferente Entrega", "Descuento ZTK1 (%)", "Motivo pedido", "Posición pedido",
    "Material", "EAN", "Cantidad"
  ];

  const rows = [headers];
  let orderIndex = 0;

  paquete.pedidos.forEach((pedido) => {
    orderIndex++;
    const itemsPedido = paquete.items.filter(i => i.Pedido_ID === pedido.Pedido_ID);
    let posicionContador = 10;
    let esPrimeraPosicion = true;

    itemsPedido.forEach((item) => {
      let matSeleccionado = (paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID]) ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";
      
      // CORRECCIÓN: Permite materiales de 12 dígitos, descarta EANs (13 dígitos que empiezan por 84) o vacíos
      const isNoActualizado = (paquete.noActualizadosSeleccionados && paquete.noActualizadosSeleccionados[item.Item_ID]) || 
                              (!matSeleccionado || matSeleccionado.includes("SIN_") || matSeleccionado === item.Producto_Cod || (matSeleccionado.length === 13 && matSeleccionado.startsWith("84")));

      const matCsv = isNoActualizado ? "" : matSeleccionado;

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

      const rowData = new Array(16).fill("");

      rowData[0] = orderIndex; // Col 1: Contador
      rowData[1] = "";         // Col 2: Nº Pedido SAP

      // Datos de cabecera solo en la primera posición del pedido
      if (esPrimeraPosicion) {
        rowData[2] = sap.CLASE_PEDIDO;      // Col 3: Clase pedido
        rowData[3] = sap.ORG_VENTAS;        // Col 4: Org ventas
        rowData[4] = sap.CANAL;             // Col 5: Canal
        rowData[5] = sap.SECTOR;            // Col 6: Sector
        rowData[6] = configCli.SOLICITANTE; // Col 7: Solicitante
        rowData[7] = pedido.Id_Destino || "11033482"; // Col 8: Destinatario Merc
        rowData[8] = nuevoIdStr;            // Col 9: Num pedido cliente
        rowData[9] = fechaEntFormatted;     // Col 10: Fecha preferente Entrega
        esPrimeraPosicion = false;
      }

      rowData[10] = "";                // Col 11: Descuento ZTK1 (%)
      rowData[11] = "";                // Col 12: Motivo pedido
      rowData[12] = posicionContador;  // Col 13: Posición pedido (10, 20, 30...)
      rowData[13] = matCsv;            // Col 14: Material SAP (ahora incluye códigos de 12 dígitos)
      rowData[14] = "";                // Col 15: EAN
      rowData[15] = item.Cant;         // Col 16: Cantidad

      rows.push(rowData);
      posicionContador += 10;
    });
  });

  return rows.map(r => r.join(";")).join("\n");
}

function cmx_enviarNotificacionKam_(paquete, nuevoIdStr, ciudadDestino) {
  const kamInfo = cmx_getKamInitialsAndName_(CONFIG.CLIENTES["CMX"].SOLICITANTE);
  const kamEmail = kamInfo.email;

  let htmlBody = `
    <div style="font-family: 'Inter', sans-serif; color: #2B2830; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
      <div style="text-align: center; border-bottom: 2px solid #E41F33; padding-bottom: 15px; margin-bottom: 20px;">
        <h1 style="color: #2B2830; font-size: 24px; margin: 0; font-weight: 900; letter-spacing: 2px;">ISDIN</h1>
        <p style="color: #E41F33; font-size: 11px; font-weight: 700; margin: 5px 0 0; text-transform: uppercase; letter-spacing: 1px;">Notificación Automática de Órdenes B2B - CMX</p>
      </div>
      
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Estimada <strong>${kamInfo.name}</strong>,</p>
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Te informamos que se han consolidado de manera exitosa los nuevos pedidos de <strong>CMX S.A.S.</strong> para la ciudad de <strong>${ciudadDestino}</strong>. El consecutivo asignado es el <span style="font-family: monospace; font-weight: bold; background-color: #F3F4F6; padding: 2px 6px; border-radius: 4px; border: 1px solid #E5E7EB;">${nuevoIdStr}</span>.</p>
      
      <h2 style="font-size: 16px; font-weight: 700; border-bottom: 1px solid #F3F4F6; padding-bottom: 8px; margin-top: 25px; color: #2B2830;">Detalle de Facturación por Orden de Compra</h2>
      <div style="margin-top: 15px;">
  `;

  paquete.pedidos.forEach(p => {
    const itemsPedido = paquete.items.filter(i => i.Pedido_ID === p.Pedido_ID);
    let itemsNoActualizados = [];
    let itemsValidos = [];
    let totalValorOk = 0;
    let totalValorBO = 0;
    let hasBackOrder = false;

    itemsPedido.forEach(item => {
      let mat = (paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID]) ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";
      
      // CORRECCIÓN: Admite materiales de 12 dígitos sin mandarlos erróneamente a No Actualizados
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
        ? '<span style="background-color: #FEF3C7; color: #D97706; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">⚠️️ Back Order</span>'
        : '<span style="background-color: #D1FAE5; color: #059669; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">✅ OK</span>';

      const formattedSub = "$" + Math.round(item.subtotal).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
      const subtotalText = item.subtotal > 0 ? ` | Subtotal: <strong>${formattedSub}</strong>` : "";

      listHtml += `
        <div style="padding: 10px 0; border-bottom: 1px dashed #F3F4F6; display: flex; justify-content: space-between; font-size: 12px;">
          <div style="flex: 1; min-width: 0; padding-right: 10px;">
            <span style="font-weight: bold; color: #111827; display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${item.desc}</span>
            <span style="color: #6B7280; font-family: monospace; font-size: 11px;">EAN: ${item.Producto_Cod} | Material SAP: <b>${item.material}</b> | ${item.cant} ud.${subtotalText}</span>
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