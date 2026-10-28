/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * ARCHIVO: cmx_motor.js
 * MÓDULO: Integración de órdenes de compra (OC) de CMX S.A.S. a partir de archivos Excel
 * =========================================================================================
 * QUÉ CONTIENE
 *   Toda la lógica que procesa los pedidos de CMX:
 *   1. Lee cada Excel de la carpeta de CMX (CONFIG.CLIENTES.CMX.FOLDER_ID) y busca el número
 *      de OC, la fecha, el destino y la tabla de productos.
 *   2. Asigna a cada producto su material SAP y su precio, usando las hojas Productos y Bodega.
 *   3. Decide el destino de cada OC con la columna Texto_OC de la hoja Destinatarios.
 *   4. Al confirmar, escribe una fila por OC en la hoja CONSOLIDADO, genera el CSV de carga a SAP
 *      en la carpeta CONFIG.FOLDER_CSV_ID, mueve cada Excel a la carpeta de procesados y envía
 *      un correo a la KAM asignada.
 *
 * QUIÉN LO LLAMA
 *   - wizard_ui.html (asistente de carga). La configuración CMX de esa pantalla indica los nombres
 *     de las funciones (fnPreAnalizar = preAnalizarCMX, fnGuardar = guardarDefinitivoCMX). Además,
 *     cuando hay OC sin destino, llama directamente a cmx_resolverDestinoPorId().
 *   - main.js NO llama a este archivo. Lo que este archivo usa de main.js es: CONFIG,
 *     nombreArchivoSap_(), enviarCorreoKam_(), obtenerInfoKamPorSolicitante_() y construirCsvSap_().
 * =========================================================================================
 */

/**
 * Lee los Excel de CMX y devuelve las OC y los productos encontrados, para que el asistente
 * los muestre antes de guardar. NO escribe en CONSOLIDADO ni en carpetas; solo lee.
 * Llamada desde wizard_ui.html (fnPreAnalizar de CMX).
 *
 * Pasos:
 * PASO 1. Lee las OC que ya existen en CONSOLIDADO (clave OC_SAP) para no volver a cargarlas.
 * PASO 2. Lee la hoja Destinatarios (cruce y bodega). Ojo: el destino final lo decide
 *         cmx_resolverDestino_, no este paso.
 * PASO 3. Carga precios y materiales de CMX (cmx_getProductosPreciosMap_).
 * PASO 4. Lista los Excel (.xls, .xlsx, .xlsm) de la carpeta de CMX.
 * PASO 5. Lee cada Excel (cmx_motorLecturaExcel).
 *
 * Efectos: crea y borra (manda a la papelera) copias temporales de los Excel en Google Sheets.
 * No envía correos.
 *
 * @return {{pedidosExtraidos: Array<Object>, itemsExtraidos: Array<Object>,
 *           errores: Array<String>, yaProcesados: Array<String>}}
 *   pedidosExtraidos: una entrada por OC (Pedido_ID, Fecha_Pedido, Id_Destino, Destinatario,
 *                     Ciudad, Destino_Pendiente, Archivo_Nombre, Archivo_Id, etc.).
 *   itemsExtraidos: una entrada por producto (Item_ID, Pedido_ID, Producto_Cod, Descripcion,
 *                   Cant, Valor_Unitario, EAN_Mapeado, Materiales_Disponibles).
 *   errores: textos "archivo: motivo" de los Excel que no se pudieron leer.
 *   yaProcesados: OC que ya estaban en CONSOLIDADO y se omiten.
 * @throws {Error} Si no existe la configuración de CMX en main.js, o si ningún Excel se pudo
 *                 leer y hubo errores (el mensaje incluye el detalle técnico).
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
    // PASO 2: MAPEO DE DESTINATARIOS (cruce por texto, Id de destino y bodega)
    // NOTA: diccDestinatarios y diccBodega se llenan aquí y se pasan a los siguientes pasos,
    // pero la decisión de destino de cada OC la toma cmx_resolverDestino_ (no usa estos mapas).
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

/**
 * Recorre los Excel de CMX uno por uno: crea una copia temporal en Google Sheets, lee la hoja
 * de la OC, extrae cabecera y productos, y clasifica cada OC como nueva o ya procesada.
 * Si un archivo falla, se registra en "errores" y se sigue con el siguiente.
 *
 * Efectos: crea una copia temporal del Excel (Drive) y la manda a la papelera al terminar.
 * No escribe en CONSOLIDADO ni envía correos.
 *
 * @param {Array<GoogleAppsScript.Drive.File>} archivos Excel de Drive a procesar.
 * @param {Object} diccBodega Mapa cruce de destino -> Id de destino. Se reenvía a
 *                            cmx_extraerDatosHeuristico_, donde no se usa.
 * @param {Object} diccProductos Precios y materiales por EAN o material (de cmx_getProductosPreciosMap_).
 * @param {Object} configCli Configuración de CMX (CONFIG.CLIENTES.CMX de main.js).
 * @param {Set<String>} pedidosProcesados Claves "OC_SAP" que ya están en CONSOLIDADO.
 * @param {Object} diccDestinatarios Destinos por cruce. Se reenvía; no se usa dentro.
 * @return {{pedidosExtraidos: Array<Object>, itemsExtraidos: Array<Object>,
 *           errores: Array<String>, yaProcesados: Array<String>}}
 *   yaProcesados contiene las OC que se omitieron por estar ya registradas.
 * @throws No lanza error: los fallos de cada archivo quedan en "errores".
 */
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
        Id_Destino: cabecera.idDestino || "",
        Destinatario: cabecera.destinatario || "",
        Ciudad: cabecera.ciudad || "",
        Destino_Pendiente: !!cabecera.destinoPendiente,
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

/**
 * Limpia un número de OC para compararlo y guardarlo: quita el apóstrofe inicial, convierte
 * notación científica (ej. "4.5E+09") a número entero y elimina la parte decimal.
 * Ejemplo: "'4500123.0" devuelve "4500123".
 *
 * @param {*} val Valor de la celda o del texto de la OC (puede ser número, texto o vacío).
 * @return {String} OC limpia; cadena vacía si el valor era null, undefined o vacío.
 * @throws No lanza errores.
 */
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

/**
 * Crea una copia del Excel como Google Sheets en la misma carpeta de Drive, para poder leerla
 * con SpreadsheetApp. La copia se llama igual que el Excel, sin extensión.
 * Usa el servicio avanzado "Drive" (debe estar activado en el proyecto de Apps Script).
 *
 * Efectos: crea un archivo nuevo en Drive (la función que llama lo borra al terminar).
 *
 * @param {GoogleAppsScript.Drive.File} file Excel original en Drive.
 * @return {String} Id de la copia en Google Sheets.
 * @throws {Error} Si Drive no puede crear la copia (lo captura cmx_motorLecturaExcel y lo registra en errores).
 */
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

/**
 * Interpreta una hoja de OC de CMX sin plantilla fija: busca el número de OC, la fecha, el destino
 * y la fila donde empieza la tabla de productos (por los encabezados), y arma los productos con
 * su precio y su material SAP.
 * No escribe en ninguna hoja ni carpeta. Abre la hoja de configuración (SpreadsheetApp), pero
 * esa variable no se usa.
 *
 * Pasos:
 * 1. Junta el texto de las filas 1 a 12 (todas las columnas) para buscar el destino.
 * 1b. Decide el destino con cmx_resolverDestino_ (texto de la OC contra Texto_OC).
 * 2. Recorre las filas 1 a 30: toma la OC y la fecha del pedido, y puntúa cada fila como
 *    posible encabezado de la tabla (EAN/código, descripción y cantidad).
 * 3. Si no encontró OC en el contenido, toma el primer número del nombre del archivo.
 * 4. Si no encontró encabezados de tabla, lanza error.
 * 5. Carga el stock de bodega (cmx_getBodegaStockMap_).
 * 6. Recorre las filas de la tabla: omite filas vacías, cantidades <= 0 y filas TOTAL; para cada
 *    producto busca el EAN con stock y el precio en diccProductos.
 *
 * @param {Array<Array<*>>} data Valores de la hoja del Excel (getValues()).
 * @param {Object} diccBodega No se usa en esta función (se conserva en la firma).
 * @param {Object} diccProductos Precios y materiales por EAN o material (de cmx_getProductosPreciosMap_).
 * @param {Object} configCli Configuración de CMX (CONFIG.CLIENTES.CMX).
 * @param {String} fileName Nombre del Excel (sirve de respaldo para la OC).
 * @param {Object} diccDestinatarios No se usa; el destino se lee con cmx_leerDestinosCliente_.
 * @return {{cabecera: Object, items: Array<Object>}}
 *   cabecera: { ordenCompra, fechaPedido, fechaEntrega (siempre vacía aquí), idDestino,
 *               destinatario, ciudad, destinoPendiente }.
 *   items: una entrada por producto { codigo, descripcion, cant, valorUnitario, eanMapeado,
 *          materialesDisponibles }.
 * @throws {Error} Si no se detecta la tabla de productos (columna de código de barras y cantidad).
 */
function cmx_extraerDatosHeuristico_(data, diccBodega, diccProductos, configCli, fileName, diccDestinatarios) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  let cabecera = {
    ordenCompra: "",
    fechaPedido: "",
    fechaEntrega: "",
    idDestino: "",
    destinatario: "",
    ciudad: "",
    destinoPendiente: false
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

  // 1b. DESTINO: se compara el texto de la OC con la columna Texto_OC de Destinatarios (solo destinos de CMX).
  //     Si no coincide: con un solo destino se usa ese; con varios, la orden queda PENDIENTE y el asistente
  //     pide el ID de destino. Nunca se asigna una ciudad por defecto.
  const resDestino = cmx_resolverDestino_(textoDireccion, cmx_leerDestinosCliente_(configCli.SOLICITANTE));
  cabecera.destinoPendiente = resDestino.estado === "pendiente";
  if (resDestino.destino) {
    cabecera.idDestino = resDestino.destino.idDestino;
    cabecera.destinatario = resDestino.destino.nombreDestino;
    cabecera.ciudad = resDestino.destino.ciudad;
  }

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

  // 3. Respaldo: si no se encontró la OC en el contenido, se toma el primer número del nombre del archivo.
  if (!cabecera.ordenCompra && fileName) {
    const matches = fileName.match(/\d+/);
    if (matches) cabecera.ordenCompra = matches[0];
  }

  // 4. Si no se encontró ninguna fila con encabezados de tabla, el archivo no se puede leer.
  if (filaInicioTabla === -1) {
    throw new Error("No se pudo detectar la tabla de productos (columna de Código de barras y Cantidad) en el archivo.");
  }

  // 5. Stock de bodega por EAN y por material (se carga una vez por archivo).
  const stockBodega = cmx_getBodegaStockMap_();

  // 6. Recorre los productos, debajo del encabezado de la tabla.
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

/**
 * Elige el código EAN de una fila de producto. Una celda puede traer varios códigos separados
 * por espacio, coma o punto y coma. Se prefiere el primero que tenga stock; si ninguno tiene,
 * se usa el primero.
 *
 * Pasos:
 * 1. Separa la celda en códigos y limpia cada uno (solo dígitos, máximo 13).
 * 2. Descarta los códigos de menos de 7 dígitos.
 * 3. Devuelve el primer código con stock mayor a cero y sus materiales.
 * 4. Si ninguno tiene stock, devuelve el primer código y los materiales que tenga en bodega.
 *
 * @param {*} rawVal Texto de la celda de código de barras (puede estar vacía).
 * @param {Object} stockBodega Mapa de cmx_getBodegaStockMap_.
 * @return {{eanMapeado: String, materialesDisponibles: Array<Object>}}
 *   eanMapeado: código elegido, o "" si no hay código válido.
 *   materialesDisponibles: { material, stock, fechaDisplay, isExpired } de bodega, o [].
 * @throws No lanza errores.
 */
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

/**
 * Deja solo los dígitos de un código EAN. Si el valor llega en notación científica
 * (ej. "7.7E+12") lo convierte a número completo. Si tiene más de 13 dígitos, conserva
 * solo los primeros 13 (sin avisar).
 *
 * @param {*} val Valor de la celda (número o texto).
 * @return {String} Solo dígitos, máximo 13; cadena vacía si no quedó ningún dígito.
 * @throws No lanza errores.
 */
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

/**
 * Busca en una fila una celda que contenga alguna de las palabras clave y devuelve el valor
 * que está a su derecha (primera celda no vacía). Si no hay nada a la derecha, usa la celda
 * de la fila siguiente, en la misma columna. Las fechas se devuelven como dd/MM/yyyy.
 *
 * La coincidencia es por contenido: la celda sirve si contiene la palabra clave (sin tildes
 * ni signos). Ojo: palabras cortas (ej. "PO", "N") coinciden con muchas palabras.
 *
 * @param {Array<Array<*>>} data Valores de la hoja.
 * @param {Number} row Índice (base 0) de la fila donde se busca.
 * @param {Array<String>} keywords Palabras clave a buscar (ej. ["OC", "PEDIDO"]).
 * @return {String} Valor encontrado, o cadena vacía si no hay coincidencia.
 * @throws No lanza errores.
 */
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

/**
 * Convierte un texto en una clave para comparar: quita tildes, espacios, puntos y signos,
 * y lo pasa a mayúsculas. Ejemplo: "Material Id" devuelve "MATERIALID"; "N°" devuelve "N".
 * Se usa en todas las comparaciones de encabezados y códigos de este archivo.
 *
 * @param {*} text Texto, número o celda vacía.
 * @return {String} Clave normalizada; cadena vacía si el valor era null o undefined.
 * @throws No lanza errores.
 */
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
 * Arma el mapa de precios y materiales de la hoja Productos, en dos fases:
 * Fase 1: carga todos los productos (catálogo general) como respaldo; si una clave ya existe, no se cambia.
 * Fase 2: sobrescribe con la tarifa de CMX: solo las filas cuya columna "Lista Cliente" coincide
 *         con la lista de CMX (CONFIG.CLIENTES.CMX.LISTA_PRECIOS).
 * Cada clave es el EAN o el material, normalizados con cmx_normalizeKey_.
 *
 * Lee la hoja Productos (no escribe). Si la hoja no existe, o no tiene las columnas EAN, Costo
 * o Material, devuelve un mapa vacío sin avisar.
 *
 * @return {Object} Mapa clave -> { precio: Number, descripcion: String,
 *                  materialOficial: String, eanOficial: String }.
 * @throws No lanza errores.
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
    // Se acepta la fila si su Lista Cliente es la de CMX. Ojo: una Lista Cliente vacía, o una
    // parte del nombre de la lista, también cumple esta condición.
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

/**
 * Arma el mapa de stock de la hoja Bodega, indexado por EAN y por material (normalizados).
 * Un mismo producto puede tener varias filas (lotes con distinta fecha de vencimiento);
 * todas se guardan en la misma lista.
 *
 * Columnas: el EAN se lee SIEMPRE de la columna J (posición 9, fija). El material se busca
 * por encabezado (MATERIAL o SAP) y, si no aparece, se toma la columna A. El stock se busca
 * por encabezado y, si no aparece, se toma la columna F. La fecha de vencimiento se toma
 * del primer encabezado que diga VENCIMIENTO, CADUCIDAD, FECHA o VENCE.
 *
 * Lee la hoja Bodega (no escribe). Si la hoja no existe o tiene menos de 2 filas, devuelve {}.
 *
 * @return {Object} Mapa clave -> Array<{ material: String, stock: Number,
 *                  fechaDisplay: String, isExpired: Boolean }>.
 *                  fechaDisplay es "yyyy-MM", el texto original, o "Sin fecha".
 *                  isExpired es true si la fecha ya pasó.
 * @throws No lanza errores.
 */
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

/**
 * SIN USO, PENDIENTE DE ELIMINAR. Ninguna función del proyecto la llama desde que el destino
 * de CMX se decide con cmx_resolverDestino_. No se borra todavía; se deja documentada.
 * Buscaba la ciudad de un cliente en la hoja Clientes (o "Listado clientes"). Si no la encontraba,
 * devolvía "COTA", algo que hoy contradice la regla de no asignar ciudad por defecto.
 *
 * @param {String|Number} solicitanteId ID SAP del cliente a buscar.
 * @return {String} Ciudad en mayúsculas, o "COTA" si no se encuentra.
 * @throws No lanza errores controlados (si la hoja no existe devuelve "COTA").
 */
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
 * Devuelve el nombre, las iniciales y el correo de la KAM asignada al cliente.
 * Usa obtenerInfoKamPorSolicitante_() (main.js). Si esa función no existe, falla o no encuentra
 * la KAM, devuelve un valor de respaldo: nombre "KAM ISDIN", iniciales "MC" y correo vacío.
 * Con correo vacío el envío de la notificación falla (lo captura guardarDefinitivoCMX).
 * No escribe nada.
 *
 * @param {String|Number} solicitanteId ID de SAP del solicitante (CMX: 11033482).
 * @return {{name: String, initials: String, email: String}} Nombre, iniciales y correo de la KAM.
 * @throws No lanza errores (los errores de la consulta se registran en el log).
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

/**
 * Guarda en firme las OC de CMX que el usuario confirmó en el asistente. Es la única función que
 * escribe datos. Llamada desde wizard_ui.html (fnGuardar de CMX).
 *
 * Escribe en: hoja CONSOLIDADO (filas nuevas en columnas A a K), carpeta CONFIG.FOLDER_CSV_ID
 * (un CSV), carpeta CONFIG.CLIENTES.CMX.PROCESSED_FOLDER_ID (mueve y renombra cada Excel),
 * y envía un correo a la KAM si paquete.notificarKAM es true.
 *
 * Pasos:
 * 1. Revisa que exista la configuración de CMX y la hoja CONSOLIDADO.
 * 2. Rechaza el guardado si alguna OC no tiene destino confirmado.
 * 3. Escudo anti-duplicados: quita las OC que ya están en CONSOLIDADO (clave OC_SAP).
 * 4. Calcula el prefijo del consecutivo (iniciales KAM + año, por ejemplo "MC-26") y el último número usado.
 * 5. Arma una fila por OC (total de unidades y valor; las OC sin unidades se omiten).
 * 6. Crea el CSV en la carpeta de CSV. Si falla, no se escribe nada en CONSOLIDADO.
 * 7. Escribe las filas en CONSOLIDADO. Si falla, manda el CSV a la papelera.
 * 8. Mueve cada Excel a la carpeta de procesados y lo renombra con su consecutivo (errores ignorados).
 * 9. Si se pidió, envía el correo a la KAM (si falla, solo queda en el log y el guardado sigue).
 *
 * @param {Object} paquete Datos confirmados en el asistente:
 *   pedidos: Array de OC de cmx_motorLecturaExcel (Pedido_ID, Fecha_Pedido, Id_Destino,
 *            Destinatario, Ciudad, Destino_Pendiente, Archivo_Id, etc.).
 *   items: Array de productos (Item_ID, Pedido_ID, Producto_Cod, Cant, Valor_Unitario, Descripcion).
 *   materialesSeleccionados: { Item_ID: material SAP elegido }.
 *   noActualizadosSeleccionados: { Item_ID: true } si el producto se marca como no actualizado.
 *   backordersSeleccionados: { Item_ID: true } si el producto va como back order (solo en el correo).
 *   fechasEntregas: { Pedido_ID: { fecha: String } } fecha de entrega que escribió el usuario.
 *   notificarKAM: Boolean, si se envía el correo a la KAM.
 * @return {{success: Boolean, mensaje: String, csvData: String, fileName: String}}
 *   mensaje: resumen con los consecutivos asignados.
 *   csvData: el CSV en base64 (para descargarlo desde el navegador).
 *   fileName: nombre del CSV creado en Drive.
 * @throws {Error} Con mensaje para el usuario si: falta la configuración o la hoja CONSOLIDADO;
 *                 hay OC sin destino; todas las OC ya estaban registradas; no hay posiciones válidas;
 *                 no se pudo guardar el CSV (no se registra nada) o no se pudo escribir CONSOLIDADO.
 */
function guardarDefinitivoCMX(paquete) {
  try {
    const configCli = CONFIG.CLIENTES["CMX"];
    if (!configCli) throw new Error("No se encontró la configuración para CMX.");

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    if (!shConsolidado) throw new Error("No se encontró la hoja CONSOLIDADO.");

    // 2. No se guarda ninguna orden sin destino confirmado
    const sinDestino = (paquete.pedidos || []).filter(p => p.Destino_Pendiente || !p.Id_Destino || !p.Ciudad);
    if (sinDestino.length > 0) {
      throw new Error("Hay órdenes sin destino confirmado: " + sinDestino.map(p => p.Pedido_ID).join(", ") + ". Indique el ID de destino antes de guardar.");
    }

    // 3. ESCUDO EN TIEMPO REAL: vuelve a revisar CONSOLIDADO antes de escribir, por si alguien ya guardó esas OC
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

    // 4. Prefijo del consecutivo (iniciales KAM + año) y búsqueda del último número usado en CONSOLIDADO.
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
    
    // 5. Arma una fila por OC. Cada OC recibe su propio consecutivo (ej. MC-260001, MC-260002).
    const filasNuevas = [];
    const pedidosRegistrados = [];
    const idPorPedido = {};

    paquete.pedidos.forEach(p => {
      const itemsPedido = paquete.items.filter(i => i.Pedido_ID === p.Pedido_ID);
      let totalUnidades = 0;
      let totalValor = 0;
      
      itemsPedido.forEach(item => {
        let matSeleccionado = (paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID]) ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";
        
        // Se considera "no actualizado" si el usuario lo marcó, si no hay material, si el material
        // contiene "SIN_", si es igual al EAN, o si es un EAN de 13 dígitos que empieza por 84.
        // Los materiales SAP de 12 dígitos sí se aceptan.
        const isNoActualizado = (paquete.noActualizadosSeleccionados && paquete.noActualizadosSeleccionados[item.Item_ID]) || 
                                (!matSeleccionado || matSeleccionado.includes("SIN_") || matSeleccionado === item.Producto_Cod || (matSeleccionado.length === 13 && matSeleccionado.startsWith("84")));

        // Las líneas sin Material SAP también cuentan: van al CSV con Material vacío para que
        // SAP las deje como back order (igual que MEDIPIEL y CEN). Solo no suman valor.
        const cant = parseInt(item.Cant) || 0;
        totalUnidades += cant;
        if (!isNoActualizado) {
          const valUnit = parseFloat(item.Valor_Unitario) || 0;
          totalValor += (cant * valUnit);
        }
      });

      if (totalUnidades === 0) return;

      maxConsecutive++;
      const idPedido = prefix + String(maxConsecutive).padStart(4, "0");
      idPorPedido[p.Pedido_ID] = idPedido;
      pedidosRegistrados.push(p);

      const fechaEnt = paquete.fechasEntregas[p.Pedido_ID] ? paquete.fechasEntregas[p.Pedido_ID].fecha : "";

      filasNuevas.push([
        idPedido,
        p.Pedido_ID,
        configCli.SOLICITANTE,
        configCli.NOMBRE_CLIENTE || "CMX S.A.S.",
        p.Fecha_Pedido || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        p.Id_Destino || "11033482", 
        p.Destinatario || "CMX S.A.S.",
        p.Ciudad,
        fechaEnt,
        totalUnidades,
        totalValor
      ]); // Solo A:K: desde la L en adelante hay fórmulas y formatos condicionales
    });

    if (filasNuevas.length === 0) {
      throw new Error("No hay posiciones válidas para inyectar al consolidado.");
    }

    const idsGenerados = pedidosRegistrados.map(p => idPorPedido[p.Pedido_ID]);
    const paqueteRegistrado = Object.assign({}, paquete, { pedidos: pedidosRegistrados });

    // 6. Guardar el CSV ANTES de CONSOLIDADO: si falla, la OC no queda marcada como procesada.
    const csvData = cmx_generarCsvSap_(paqueteRegistrado);
    const fileName = nombreArchivoSap_(`Pedido_${idsGenerados[0]}.csv`);
    let archivoCsv = null;

    // Igual que CEN y MEDIPIEL: el CSV va a CONFIG.FOLDER_CSV_ID
    try {
      archivoCsv = DriveApp.getFolderById(CONFIG.FOLDER_CSV_ID).createFile(fileName, csvData, MimeType.PLAIN_TEXT);
    } catch (csvDriveErr) {
      throw new Error("No se pudo guardar el CSV en Drive (no se registró nada en CONSOLIDADO): " + csvDriveErr.message);
    }

    // 7. Escribir CONSOLIDADO. Si falla, se retira el CSV para no dejar un archivo huérfano.
    try {
      const lastRow = shConsolidado.getLastRow();
      shConsolidado.getRange(lastRow + 1, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);
      SpreadsheetApp.flush();
    } catch (errSheets) {
      try { archivoCsv.setTrashed(true); } catch (ignore) {}
      throw new Error("Fallo al escribir en CONSOLIDADO; el CSV se envió a la papelera: " + errSheets.message);
    }

    // 8. Mover y renombrar cada Excel con el consecutivo de su propia OC. Si un archivo falla, se ignora.
    const folderDestino = DriveApp.getFolderById(configCli.PROCESSED_FOLDER_ID);
    pedidosRegistrados.forEach(p => {
      if (p.Archivo_Id) {
        try {
          const file = DriveApp.getFileById(p.Archivo_Id);
          file.moveTo(folderDestino);

          const originalName = file.getName();
          const extIdx = originalName.lastIndexOf(".");
          const ext = extIdx !== -1 ? originalName.substring(extIdx) : ".xlsm";
          file.setName(nombreArchivoSap_("Pedido_" + idPorPedido[p.Pedido_ID] + ext));
        } catch (ignore) {}
      }
    });

    // 9. Correo a la KAM (opcional). Si falla, el guardado ya quedó hecho y solo se registra en el log.
    if (paquete.notificarKAM) {
      try {
        cmx_enviarNotificacionKam_(paqueteRegistrado, idsGenerados.join(", "), pedidosRegistrados[0].Ciudad);
      } catch (eMailError) {
        console.log("Fallo el envío del correo: " + eMailError.message);
      }
    }

    return {
      success: true,
      mensaje: `Se han procesado ${filasNuevas.length} pedidos. Consecutivos asignados: ${idsGenerados.join(", ")}`,
      csvData: Utilities.base64Encode(csvData, Utilities.Charset.UTF_8).replace(/[\r\n]/g, ""),
      fileName: fileName
    };

  } catch (err) {
    throw new Error(err.message);
  }
}

/**
 * Arma el contenido del CSV de carga a SAP para las OC de CMX. Cada línea toma el material elegido
 * por el usuario; si el producto es "no actualizado", el material va vacío (SAP lo deja como back order).
 * El formato del CSV lo define construirCsvSap_ (main.js). Acepta materiales SAP de 12 dígitos
 * (ej. 690032151480). No escribe en Drive; quien lo llama crea el archivo.
 *
 * @param {Object} paquete Datos de guardarDefinitivoCMX (pedidos ya filtrados, items,
 *                         materialesSeleccionados, noActualizadosSeleccionados, fechasEntregas).
 * @return {String} Texto del CSV.
 * @throws No lanza errores controlados; un dato mal formado puede provocar error de JavaScript.
 */
function cmx_generarCsvSap_(paquete) {
  const configCli = CONFIG.CLIENTES["CMX"];

  const pedidosCsv = paquete.pedidos.map((pedido) => {
    const itemsPedido = paquete.items.filter(i => i.Pedido_ID === pedido.Pedido_ID);

    const lineas = itemsPedido.map((item) => {
      let matSeleccionado = (paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID]) ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";

      // Material "no actualizado" = marcado por el usuario, vacío, con "SIN_", igual al EAN, o EAN de 13 dígitos que empieza por 84.
      // Los materiales de 12 dígitos sí se aceptan.
      const isNoActualizado = (paquete.noActualizadosSeleccionados && paquete.noActualizadosSeleccionados[item.Item_ID]) ||
                              (!matSeleccionado || matSeleccionado.includes("SIN_") || matSeleccionado === item.Producto_Cod || (matSeleccionado.length === 13 && matSeleccionado.startsWith("84")));

      return { material: isNoActualizado ? "" : matSeleccionado, cantidad: item.Cant };
    });

    return {
      sap: configCli.SAP,
      solicitante: configCli.SOLICITANTE,
      destinatario: pedido.Id_Destino || "11033482",
      numPedidoCliente: pedido.Pedido_ID,
      fechaEntrega: paquete.fechasEntregas[pedido.Pedido_ID] ? paquete.fechasEntregas[pedido.Pedido_ID].fecha : "",
      lineas: lineas
    };
  });

  return construirCsvSap_(pedidosCsv);
}

/**
 * Envía a la KAM un correo HTML con el resumen de las OC registradas: por cada OC, los productos,
 * el material SAP, el valor estimado a facturar (OK), el valor en back order y los productos
 * "no actualizados" (con alerta). Lo llama guardarDefinitivoCMX solo si se pidió notificar.
 *
 * Escribe: envía un correo (MailApp, vía enviarCorreoKam_ de main.js). En modo PRUEBA, el correo
 * sale con el prefijo [PRUEBA] y puede ir a otra dirección. Con correo de KAM vacío, el envío falla.
 *
 * @param {Object} paquete Datos de la OC registrada: pedidos (los que quedaron registrados),
 *                         items, materialesSeleccionados, noActualizadosSeleccionados y
 *                         backordersSeleccionados (ver guardarDefinitivoCMX).
 * @param {String} idsAsignados Consecutivos asignados, separados por coma (ej. "MC-260001, MC-260002").
 * @param {String} ciudadDestino Ciudad que aparece en el correo. Se toma solo de la primera OC y
 *                               se usa para todas.
 * @return {void}
 * @throws {Error} Si MailApp no puede enviar el correo (quien llama lo captura).
 */
function cmx_enviarNotificacionKam_(paquete, idsAsignados, ciudadDestino) {
  const kamInfo = cmx_getKamInitialsAndName_(CONFIG.CLIENTES["CMX"].SOLICITANTE);
  const kamEmail = kamInfo.email;

  let htmlBody = `
    <div style="font-family: 'Inter', sans-serif; color: #2B2830; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
      <div style="text-align: center; border-bottom: 2px solid #E41F33; padding-bottom: 15px; margin-bottom: 20px;">
        <h1 style="color: #2B2830; font-size: 24px; margin: 0; font-weight: 900; letter-spacing: 2px;">ISDIN</h1>
        <p style="color: #E41F33; font-size: 11px; font-weight: 700; margin: 5px 0 0; text-transform: uppercase; letter-spacing: 1px;">Notificación Automática de Órdenes B2B - CMX</p>
      </div>
      
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Estimada <strong>${kamInfo.name}</strong>,</p>
      <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Te informamos que se han consolidado de manera exitosa los nuevos pedidos de <strong>CMX S.A.S.</strong> para la ciudad de <strong>${ciudadDestino}</strong>. Consecutivo(s) asignado(s): <span style="font-family: monospace; font-weight: bold; background-color: #F3F4F6; padding: 2px 6px; border-radius: 4px; border: 1px solid #E5E7EB;">${idsAsignados}</span>.</p>
      
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
      
      // Mismo criterio de "no actualizado" que en el CSV (los materiales de 12 dígitos no cuentan como no actualizados)
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

  enviarCorreoKam_({
    to: kamEmail,
    subject: `📦 Resumen Pedidos CMX - ${Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")}`,
    htmlBody: htmlBody
  });
}

/**
 * =========================================================================================
 * DESTINO DE CMX
 * -----------------------------------------------------------------------------------------
 * Destinatarios tiene una columna "Texto_OC" con la dirección de la bodega tal como sale en la OC.
 * Si un destino tiene varias versiones del texto, se separan con " | ".
 * El destino se decide así:
 *   1. El texto de la OC contiene algún Texto_OC de un destino de CMX  -> ese destino.
 *   2. CMX tiene un solo destino en Destinatarios                       -> ese destino.
 *   3. Si no, la orden queda PENDIENTE: el asistente pide el ID de destino.
 * Nunca se asigna una ciudad por defecto.
 * =========================================================================================
 */

/**
 * Busca en los encabezados de la hoja Destinatarios la posición de las columnas que usa el destino
 * (Id_origen, Id_destino, Nombre_destino, Ciudad, Texto_OC). La búsqueda no distingue mayúsculas,
 * tildes ni espacios (usa cmx_normalizeKey_). Así, el orden de las columnas no importa.
 * No escribe nada.
 *
 * @param {Array<*>} encabezados Primera fila de la hoja Destinatarios.
 * @return {{origen: Number, id: Number, nombre: Number, ciudad: Number, texto: Number}}
 *   Posición (base 0) de cada columna; -1 si la columna no existe.
 * @throws No lanza errores.
 */
function cmx_colsDestinatarios_(encabezados) {
  const h = encabezados.map(x => cmx_normalizeKey_(x));
  return {
    origen: h.indexOf(cmx_normalizeKey_("Id_origen")),
    id: h.indexOf(cmx_normalizeKey_("Id_destino")),
    nombre: h.indexOf(cmx_normalizeKey_("Nombre_destino")),
    ciudad: h.indexOf(cmx_normalizeKey_("Ciudad")),
    texto: h.indexOf(cmx_normalizeKey_("Texto_OC"))
  };
}

/**
 * Convierte la celda Texto_OC en una lista de textos. Los textos van separados por " | ".
 * Ejemplo: "CALLE 1 # 2 | BODEGA NORTE" devuelve ["CALLE 1 # 2", "BODEGA NORTE"].
 * Quita espacios y descarta los textos vacíos. No escribe nada.
 *
 * @param {*} valor Contenido de la celda Texto_OC (puede estar vacía).
 * @return {Array<String>} Lista de textos; [] si la celda está vacía.
 * @throws No lanza errores.
 */
function cmx_separarTextosOC_(valor) {
  return String(valor || "").split("|").map(t => t.trim()).filter(t => t);
}

/**
 * Lee de la hoja Destinatarios todos los destinos que pertenecen al cliente, es decir, las filas
 * cuyo Id_origen es igual al solicitante. Solo lee; no escribe.
 * Si la hoja no existe, o no tiene las columnas Id_origen o Id_destino, devuelve [] sin avisar.
 *
 * @param {String|Number} solicitante SAP ID del cliente (CMX: 11033482).
 * @return {Array<{idDestino: String, nombreDestino: String, ciudad: String,
 *                 textosOC: Array<String>, fila: Number}>}
 *   idDestino: Id_destino de la fila.
 *   nombreDestino: Nombre_destino ("" si no existe la columna).
 *   ciudad: Ciudad en mayúsculas ("" si no existe la columna o la celda está vacía).
 *   textosOC: textos de Texto_OC (ver cmx_separarTextosOC_).
 *   fila: número de fila en la hoja (base 1), para poder actualizarla después.
 * @throws No lanza errores controlados.
 */
function cmx_leerDestinosCliente_(solicitante) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_DESTINATARIOS || "Destinatarios");
  if (!sh) return [];
  const data = sh.getDataRange().getValues();
  const cols = cmx_colsDestinatarios_(data[0]);
  if (cols.origen === -1 || cols.id === -1) return [];

  const destinos = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][cols.origen]).trim() !== String(solicitante).trim()) continue;
    destinos.push({
      idDestino: String(data[i][cols.id]).trim(),
      nombreDestino: cols.nombre !== -1 ? String(data[i][cols.nombre]).trim() : "",
      ciudad: cols.ciudad !== -1 ? String(data[i][cols.ciudad]).trim().toUpperCase() : "",
      textosOC: cols.texto !== -1 ? cmx_separarTextosOC_(data[i][cols.texto]) : [],
      fila: i + 1
    });
  }
  return destinos;
}

/**
 * Decide el destino de una OC de CMX, con las reglas del encabezado de este bloque:
 * 1. Si el texto de la OC contiene algún Texto_OC de un destino de CMX, ese destino.
 * 2. Si no, y CMX tiene un solo destino, ese destino (aunque el texto no coincida).
 * 3. Si no, la OC queda "pendiente" y no tiene destino; el asistente pide el ID.
 * La comparación ignora mayúsculas, tildes y signos (cmx_normalizeKey_). No escribe nada.
 *
 * @param {String} textoOC Texto de las filas 1 a 12 de la OC, unido en una sola cadena.
 * @param {Array<Object>} destinos Resultado de cmx_leerDestinosCliente_.
 * @return {{estado: String, destino: Object|null}}
 *   estado: "ok" (destino asignado) o "pendiente" (sin destino).
 *   destino: el objeto de destino de cmx_leerDestinosCliente_, o null si está pendiente.
 * @throws No lanza errores.
 */
function cmx_resolverDestino_(textoOC, destinos) {
  const textoNorm = cmx_normalizeKey_(textoOC);
  const coincide = destinos.find(d => d.textosOC.some(t => {
    const norm = cmx_normalizeKey_(t);
    return norm && textoNorm.includes(norm);
  }));
  if (coincide) return { estado: "ok", destino: coincide };
  if (destinos.length === 1) return { estado: "ok", destino: destinos[0] };
  return { estado: "pendiente", destino: null };
}

/**
 * Llamada desde wizard_ui.html (cmx_resolverDestinoPorId, con google.script.run) cuando el usuario
 * escribe el ID de destino de una OC pendiente. Busca ese ID entre los destinos de CMX.
 * Si además llega la dirección de la OC y no estaba guardada, la agrega a Texto_OC de ese destino
 * (en la hoja Destinatarios, columna Texto_OC, separada por " | "), para que la próxima OC con esa
 * dirección se reconozca sola ("aprendido").
 *
 * Escribe en: hoja Destinatarios, celda Texto_OC de la fila del destino (solo si aprende).
 *
 * Pasos:
 * 1. Busca el ID entre los destinos de CMX. 2. Verifica que tenga ciudad.
 * 3. Si hay dirección nueva, la guarda en Texto_OC.
 *
 * @param {String} idDestino ID escrito por el usuario (Id_destino de Destinatarios).
 * @param {String} [textoOC] Dirección tal como aparece en la OC (opcional).
 * @return {{idDestino: String, nombreDestino: String, ciudad: String, aprendido: Boolean}}
 *   aprendido: true si se guardó la dirección en Texto_OC.
 * @throws {Error} Si el ID no está en Destinatarios para CMX; si ese destino no tiene ciudad;
 *                 o si la hoja Destinatarios no tiene la columna Texto_OC (solo cuando hay que guardar).
 */
function cmx_resolverDestinoPorId(idDestino, textoOC) {
  const id = String(idDestino || "").trim();
  const destino = cmx_leerDestinosCliente_(CONFIG.CLIENTES["CMX"].SOLICITANTE).find(d => d.idDestino === id);
  if (!destino) {
    throw new Error(`El ID ${id} no está en Destinatarios para CMX. Agréguelo a la hoja y vuelva a intentar.`);
  }
  if (!destino.ciudad) {
    throw new Error(`El ID ${id} no tiene ciudad en Destinatarios. Complétela y vuelva a intentar.`);
  }

  let aprendido = false;
  const texto = String(textoOC || "").trim();
  if (texto && !destino.textosOC.some(t => cmx_normalizeKey_(t) === cmx_normalizeKey_(texto))) {
    const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_DESTINATARIOS || "Destinatarios");
    const cols = cmx_colsDestinatarios_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    if (cols.texto === -1) throw new Error("La hoja Destinatarios no tiene la columna Texto_OC.");
    sh.getRange(destino.fila, cols.texto + 1).setValue(destino.textosOC.concat([texto]).join(" | "));
    aprendido = true;
  }
  return { idDestino: destino.idDestino, nombreDestino: destino.nombreDestino, ciudad: destino.ciudad, aprendido: aprendido };
}
