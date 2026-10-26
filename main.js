/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: CONFIGURACIÓN CENTRALIZADA Y CONTROLADOR DE VISTAS (COLOMBIA)
 * ARCHIVO: main.gs
 * ESTADO: PRODUCCIÓN - CON ASIGNACIÓN DINÁMICA DE KAM CENTRALIZADA (SINGLE SOURCE OF TRUTH)
 * =========================================================================================
 * 📝 DESCRIPCIÓN TÉCNICA:
 * Centraliza configuraciones, parámetros de SAP, rutas de Google Drive y utilidades globales.
 * Incluye la función central `obtenerInfoKamPorSolicitante_()` que lee dinámicamente
 * la pestaña `Asignacion_KAM` para proveer nombre, iniciales y correo electrónico a todos
 * los módulos (Medipiel, CMX, CEN).
 * =========================================================================================
 */

const CONSOLIDADO_HEADERS = [
  "Id", 
  "Orden De Compra", 
  "Sap Id", 
  "Solicitante", 
  "Fecha Pedido", 
  "Id Destino", 
  "Destinatario", 
  "Ciudad", 
  "Fecha entrega", 
  "Unidades", 
  "Valor", 
  "NO ENTREGA FACTURACIÓN", 
  "Observaciones",
  "KAM"
];

const CONFIG = {
  SPREADSHEET_ID: "1L5bxc9IXtBNUxYG54rBTESC3vkBO5HXNxV2L7vz02u8",
  SHEET_PEDIDOS: "CONSOLIDADO",
  SHEET_PRODUCTOS: "Productos",
  SHEET_BODEGA: "Bodega", 
  SHEET_DESTINATARIOS: "Destinatarios",
  OCR_LANGUAGE: "es",
  TRASH_TEMP_OCR_DOC: true, 
  LIMIT_FILES: 0,
  FOLDER_CSV_ID: "1NrHwmCYC866bej3MdUsqAWPnbSxRDs3d",
  
  CLIENTES: {
    "MEDIPIEL": {
      FOLDER_ID: "10k7ltzA6CclAY8z9wDhEACrGl5-3bGnx",
      PROCESSED_FOLDER_ID: "1dMqDRf92i5qgjUnEwvgM7qO1931YUu6z",
      SOLICITANTE: "11026712",
      NOMBRE_CLIENTE: "MEDIPIEL S.A.S.",
      MIME_TYPE: MimeType.PDF,
      LISTA_PRECIOS: "MEDIPIEL BEAUTYCALIA",
      SAP: {
        CLASE_PEDIDO: "ZDIR",
        ORG_VENTAS: "Z011",
        CANAL: "10",
        SECTOR: "00"
      }
    },
    "CMX": {
      FOLDER_ID: "1JiT2US9uOHgILK-ekaxyQdzgpG4zHwpv",
      PROCESSED_FOLDER_ID: "1dMqDRf92i5qgjUnEwvgM7qO1931YUu6z",
      SOLICITANTE: "11033482",
      NOMBRE_CLIENTE: "CMX S.A.S.",
      FILE_QUERY: "mimeType contains 'spreadsheet' or mimeType contains 'excel'",
      LISTA_PRECIOS: "CMX\nCADA PIEL\nThe Beauty club",
      SAP: {
        CLASE_PEDIDO: "ZDIR",
        ORG_VENTAS: "Z011",
        CANAL: "10",
        SECTOR: "00"
      }
    },
    "CEN": {
      FOLDER_ID: "1hCjs2bI3YvvmcqIOh-NufG9CCh2Zx3RF", // ID real de la carpeta CEN_sin_procesar
      PROCESSED_FOLDER_ID: "1dMqDRf92i5qgjUnEwvgM7qO1931YUu6z",
      FILE_QUERY: "mimeType contains 'spreadsheet' or mimeType contains 'excel'",
      SAP: {
        CLASE_PEDIDO: "ZDIR",
        ORG_VENTAS: "Z011",
        CANAL: "10",
        SECTOR: "00"
      },
      // TABLA DE HOMOLOGACIÓN: Mapea Razón Social / EAN del Comprador en CEN -> Cliente Interno ISDIN
      MAPEO_CLIENTES: {
        "BELLA PIEL": {
          NOMBRE_CLIENTE: "BELLA PIEL",
          SOLICITANTE: "11026727",
          LISTA_PRECIOS: "Bella Piel",
          aliases: ["Bella Piel S A S."]
        },
        "COLSUBSIDIO": {
          NOMBRE_CLIENTE: "CAJA COLOMBIANA DE SUBSIDIO FAMILIA",
          SOLICITANTE: "11026688",
          LISTA_PRECIOS: "Distribuidor",
          aliases: ["Colsubsidio"]
        },
        "OLIMPICA": {
          NOMBRE_CLIENTE: "SUPERTIENDAS Y DROGUERIAS OLIMPICA",
          SOLICITANTE: "11059838",
          LISTA_PRECIOS: "Distribuidor",
          aliases: ["SUPERTIENDAS Y DROGUERIAS OLIMPICAS S.A."]
        },
        "COPSERVIR": {
          NOMBRE_CLIENTE: "COPERATIVA MULTIACTIVA DE SERVICIOS",
          SOLICITANTE: "11072880",
          LISTA_PRECIOS: "Distribuidor",
          aliases: ["Copservir Ltda - Drogas La Rebaja"]
        },
        "CAFAM": {
          NOMBRE_CLIENTE: "CAJA DE COMPENSACION FAMILIAR CAFAM",
          SOLICITANTE: "11048830",
          LISTA_PRECIOS: "Distribuidor",
          aliases: ["Cafam"]
        },
        "COOPIDROGAS": {
          NOMBRE_CLIENTE: "COOPERATIVA NACIONAL DE DROGUISTAS",
          SOLICITANTE: "11026732",
          LISTA_PRECIOS: "Copidrogas",
          aliases: ["Coopidrogas"]
        },
        "COMFANDI":{
         NOMBRE_CLIENTE: "CAJA COMP FAMILIAR COMFANDI",
          SOLICITANTE: "11033303",
          LISTA_PRECIOS: "Distribuidor",
          aliases: ["Comfandi"] 
        },
        "LOCATEL":{
         NOMBRE_CLIENTE: "BRECCIA SALUD S A S",
          SOLICITANTE: "11027358",
          LISTA_PRECIOS: "Distribuidor",
          aliases: ["Breccia Salud S A S"] 
        }
      }
    }
  }
};

/**
 * FORMATO ÚNICO DEL CSV DE CARGA A SAP (para todos los clientes).
 * Para cambiar el formato se edita SOLO este objeto; construirCsvSap_() lo aplica.
 *
 * Cada columna define:
 *   - nombre: rótulo de referencia (el CSV NO lleva fila de encabezado: el cargador SAP no la espera).
 *   - valor(ctx): de dónde sale el dato. ctx = { contador, pedido, linea, posicion }.
 *   - soloPrimeraLinea: true si el dato va solo en la primera línea de cada pedido
 *     y en blanco en las siguientes.
 *
 * Estructura validada con "Ejemplo estructura CSV para SAP.xlsx" (pedido SC-260412).
 */
const CSV_SAP = {
  SEPARADOR: ";",
  FIN_LINEA: "\r\n",
  FORMATO_FECHA: "dd.MM.yyyy",
  PASO_POSICION: 10,
  COLUMNAS: [
    { nombre: "Contador",                 valor: c => c.contador,                  soloPrimeraLinea: true },  // A
    { nombre: "Nº Pedido SAP",            valor: c => "" },                                                   // B
    { nombre: "Clase pedido",             valor: c => c.pedido.sap.CLASE_PEDIDO,   soloPrimeraLinea: true },  // C
    { nombre: "Org ventas",               valor: c => c.pedido.sap.ORG_VENTAS,     soloPrimeraLinea: true },  // D
    { nombre: "Canal",                    valor: c => c.pedido.sap.CANAL,          soloPrimeraLinea: true },  // E
    { nombre: "Sector",                   valor: c => c.pedido.sap.SECTOR,         soloPrimeraLinea: true },  // F
    { nombre: "Solicitante",              valor: c => c.pedido.solicitante },                                 // G
    { nombre: "Destinatario Merc",        valor: c => c.pedido.destinatario },                                // H
    { nombre: "Num pedido cliente",       valor: c => c.pedido.numPedidoCliente },                            // I
    { nombre: "Fecha preferente Entrega", valor: c => formatearFechaSap_(c.pedido.fechaEntrega) },            // J
    { nombre: "Descuento ZTK1 (%)",       valor: c => "" },                                                   // K
    { nombre: "Motivo pedido",            valor: c => "" },                                                   // L
    { nombre: "Posición pedido",          valor: c => c.posicion },                                           // M
    { nombre: "Material",                 valor: c => c.linea.material },                                     // N
    { nombre: "EAN",                      valor: c => "" },                                                   // O
    { nombre: "Cantidad",                 valor: c => c.linea.cantidad }                                      // P
  ]
};

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Order Hub - Colombia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function abrirInterfaz() {
  const html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Order Hub Colombia')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Order Hub Colombia');
}

/**
 * Consulta la información del KAM asignado a un solicitante SAP en la hoja Asignacion_KAM.
 * Fuente única de verdad para iniciales, nombre y correo electrónico corporativo.
 *
 * @param {String|Number} solicitanteSap ID de SAP del cliente.
 * @return {Object|null} Objeto con { nombre, iniciales, correo } o null si no se encuentra.
 */
function obtenerInfoKamPorSolicitante_(solicitanteSap) {
  if (!solicitanteSap) return null;
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shAsig = ss.getSheetByName("Asignacion_KAM");
    if (!shAsig || shAsig.getLastRow() < 2) return null;

    const data = shAsig.getDataRange().getValues();
    const headers = data[0].map(h => String(h).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]/g, "").toUpperCase().trim());
    
    const colSap = headers.findIndex(h => h.includes("SAP") || h.includes("ID") || h.includes("SOLICITANTE"));
    const colKam = headers.findIndex(h => h.includes("KAM") || h.includes("ENCARGADO") || h.includes("NOMBRE") || h.includes("RESPONSABLE"));
    const colMail = headers.findIndex(h => h.includes("CORREO") || h.includes("EMAIL") || h.includes("MAIL"));
    const colInic = headers.findIndex(h => h.includes("INICIAL") || h.includes("AGENTE") || h.includes("CODIGO"));

    const targetSap = String(solicitanteSap).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]/g, "").toUpperCase().trim();

    for (let i = 1; i < data.length; i++) {
      const sapVal = colSap !== -1 ? String(data[i][colSap]).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]/g, "").toUpperCase().trim() : "";
      if (sapVal === targetSap) {
        return {
          nombre: colKam !== -1 ? String(data[i][colKam]).trim() : "KAM ISDIN",
          iniciales: colInic !== -1 ? String(data[i][colInic]).trim().toUpperCase() : "MC",
          correo: colMail !== -1 ? String(data[i][colMail]).trim() : ""
        };
      }
    }
  } catch (err) {
    console.error("Error en obtenerInfoKamPorSolicitante_: " + err.message);
  }
  return null;
}

function contarArchivosPendientes(cliente) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "pendientes_" + cliente;

  const configCli = CONFIG.CLIENTES[cliente];
  if (!configCli || !configCli.FOLDER_ID) return 0;

  try {
    const folder = DriveApp.getFolderById(configCli.FOLDER_ID);
    const files = obtenerArchivosCliente_(folder, configCli);

    let count = 0;
    while (files.hasNext()) {
      const file = files.next();
      
      // 1. Descartar archivos que estén en la papelera
      if (file.isTrashed()) continue;

      const name = file.getName().toLowerCase();

      // 2. Descartar archivos temporales generados por pruebas previas
      if (name.includes("_temp")) continue;

      // 3. Para CEN y CMX, validar estrictamente que sea un Excel válido
      if (cliente === "CEN" || cliente === "CMX") {
        if (!name.endsWith(".xlsx") && !name.endsWith(".xlsm") && !name.endsWith(".xls")) {
          continue;
        }
      }

      count++;
    }

    cache.put(cacheKey, count.toString(), 60); // Refresco dinámico cada 60s
    return count;
  } catch (err) {
    console.error("Error contando pendientes de " + cliente + ": " + err.message);
    return 0;
  }
}

function obtenerArchivosCliente_(folder, configCli) {
  if (configCli.FILE_QUERY) {
    // Filtro Poka-Yoke: Asegurar que nunca busque archivos en la papelera
    const query = "(" + configCli.FILE_QUERY + ") and trashed = false";
    return folder.searchFiles(query);
  }
  return folder.getFilesByType(configCli.MIME_TYPE || MimeType.PDF);
}

function obtenerDiccionarioDestinatarios() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sh = ss.getSheetByName(CONFIG.SHEET_DESTINATARIOS || "Destinatarios");
  if (!sh) return {};
  
  const data = sh.getDataRange().getValues();
  const dicc = {};
  
  if (data.length < 2) return dicc;

  const headers = data[0].map(h => String(h).trim().toLowerCase());
  const colCruce = headers.indexOf('cruce_cliente');
  const colIdDestino = headers.indexOf('id_destino');
  const colNombreDestino = headers.indexOf('nombre_destino');
  const colCiudad = headers.findIndex(h => h.includes('ciudad') || h.includes('poblacion') || h.includes('municipio') || h === 'f');

  for (let i = 1; i < data.length; i++) {
    const cruce = String(data[i][colCruce]).trim();
    const idDestino = colIdDestino !== -1 ? String(data[i][colIdDestino]).trim() : "";
    const nombreDestino = colNombreDestino !== -1 ? String(data[i][colNombreDestino]).trim() : "";
    
    let ciudadVal = "COTA";
    if (colCiudad !== -1 && data[i][colCiudad]) {
      ciudadVal = String(data[i][colCiudad]).trim().toUpperCase();
    }

    const objInfo = {
      idDestino: idDestino,
      nombreDestino: nombreDestino,
      ciudad: ciudadVal
    };

    if (cruce) dicc[cruce] = objInfo;
    if (idDestino) dicc[idDestino] = objInfo;
  }
  return dicc;
}

function obtenerMapaBodegaGeneral_(sheet) {
  const map = {}; 
  if (!sheet) return map;
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;

  const h = data[0].map(x => String(x).normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().trim());
  
  let cE = h.indexOf("EAN");
  let cM = h.indexOf("MATERIAL");
  let cS = h.indexOf("LIBRE UTILIZACION");
  let cF = h.indexOf("FECADUC/FEPREFERCONS");

  if (cE === -1) cE = 9;  
  if (cM === -1) cM = 0;  
  if (cS === -1) cS = 5;  
  if (cF === -1) cF = 7; 

  for (let i = 1; i < data.length; i++) {
    const ean = data[i][cE] ? String(data[i][cE]).trim() : "";
    if (!ean) continue;

    const mat = data[i][cM] ? String(data[i][cM]).trim() : "";
    const stock = Number(data[i][cS]) || 0;
    
    let fechaRaw = data[i][cF];
    let fechaFormat = "Sin fecha";
    
    if (fechaRaw) {
      if (fechaRaw instanceof Date) {
        let y = fechaRaw.getFullYear();
        let m = String(fechaRaw.getMonth() + 1).padStart(2, '0');
        fechaFormat = `${y}-${m}`;
      } else {
        fechaFormat = String(fechaRaw).trim();
      }
    }

    if (!map[ean]) map[ean] = [];
    
    const existente = map[ean].find(m => m.material === mat && m.fechaDisplay === fechaFormat);
    
    if (existente) {
      existente.stock += stock;
    } else {
      map[ean].push({ material: mat, stock: stock, fechaDisplay: fechaFormat });
    }
  }
  return map;
}

/**
 * Alta de un destinatario en la hoja Destinatarios. Única función para todos los clientes.
 * Escribe cada dato en la columna de su encabezado (fila 1), así que no depende del
 * orden de las columnas. Si falta una columna obligatoria, no escribe nada.
 *
 * @param {Object} datos { idOrigen, nombreOrigen, idDestino, nombreDestino, cruceCliente, ciudad }
 * @return {Object} { success: true }
 */
function guardarNuevoDestinatario(datos) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const nombreHoja = CONFIG.SHEET_DESTINATARIOS || "Destinatarios";
  let sh = ss.getSheetByName(nombreHoja);
  if (!sh) {
    sh = ss.insertSheet(nombreHoja);
    sh.appendRow(["Id_origen", "Nombre_origen", "Id_destino", "Nombre_destino", "Cruce_cliente", "Ciudad"]);
  }

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  const columnas = {
    idOrigen: headers.indexOf('id_origen'),
    nombreOrigen: headers.indexOf('nombre_origen'),
    idDestino: headers.indexOf('id_destino'),
    nombreDestino: headers.indexOf('nombre_destino'),
    cruceCliente: headers.indexOf('cruce_cliente'),
    // Misma regla que obtenerDiccionarioDestinatarios para ubicar la ciudad
    ciudad: headers.findIndex(h => h.includes('ciudad') || h.includes('poblacion') || h.includes('municipio'))
  };

  const faltantes = Object.keys(columnas).filter(k => columnas[k] === -1);
  if (faltantes.length > 0) {
    throw new Error(`La hoja ${nombreHoja} no tiene las columnas: ${faltantes.join(", ")}. No se guardó el destinatario.`);
  }

  const fila = new Array(headers.length).fill("");
  for (const campo in columnas) {
    fila[columnas[campo]] = datos[campo] !== undefined && datos[campo] !== null ? String(datos[campo]).trim() : "";
  }
  sh.appendRow(fila);
  return { success: true };
}

/**
 * Catálogo de Materiales SAP válidos con su stock total en Bodega.
 * El wizard lo usa para validar los materiales digitados a mano: solo se acepta un
 * material que exista en Productos o en Bodega, y su stock es el real (nunca infinito).
 *
 * @return {Object} { "<material>": { stock: Number, descripcion: String } }
 */
function obtenerCatalogoMaterialesSap() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const catalogo = {};
  const norm = x => String(x).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-zA-Z0-9]/g, "").toUpperCase();

  const shProd = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  if (!shProd) throw new Error("No se encontró la hoja Productos para validar materiales.");
  const dataProd = shProd.getDataRange().getValues();
  const headProd = dataProd[0].map(norm);
  const colMat = headProd.findIndex(h => h.includes("MATERIAL"));
  const colDesc = headProd.findIndex(h => h.includes("DESCRIPCION") || h.includes("NOMBRE") || h.includes("PRODUCTO"));
  if (colMat === -1) throw new Error("La hoja Productos no tiene columna de Material.");

  for (let i = 1; i < dataProd.length; i++) {
    const mat = String(dataProd[i][colMat]).trim();
    if (!mat || catalogo[mat]) continue;
    catalogo[mat] = { stock: 0, descripcion: colDesc !== -1 ? String(dataProd[i][colDesc]).trim() : "" };
  }

  const mapaBodega = obtenerMapaBodegaGeneral_(ss.getSheetByName(CONFIG.SHEET_BODEGA || "Bodega"));
  for (const ean in mapaBodega) {
    mapaBodega[ean].forEach(lote => {
      if (!lote.material) return;
      if (!catalogo[lote.material]) catalogo[lote.material] = { stock: 0, descripcion: "" };
      catalogo[lote.material].stock += Number(lote.stock) || 0;
    });
  }
  return catalogo;
}

/**
 * Genera el contenido del CSV de carga a SAP aplicando el formato único CSV_SAP.
 * Lo usan los tres motores (MEDIPIEL, CMX, CEN): ninguno arma filas del CSV por su cuenta.
 *
 * @param {Array<Object>} pedidos Lista en el orden en que deben salir en el CSV:
 *   {
 *     sap: { CLASE_PEDIDO, ORG_VENTAS, CANAL, SECTOR },  // configCli.SAP
 *     solicitante: String,       // SAP ID del cliente
 *     destinatario: String,      // Id destino de la mercancía
 *     numPedidoCliente: String,  // OC del cliente
 *     fechaEntrega: String,      // "yyyy-MM-dd" (valor del <input type="date"> del wizard)
 *     lineas: [{ material: String, cantidad: Number }]
 *   }
 *   Los pedidos sin líneas se omiten y no consumen contador.
 * @return {String} Contenido del CSV, sin fila de encabezado.
 */
function construirCsvSap_(pedidos) {
  const filas = [];
  let contador = 0;

  pedidos.forEach(pedido => {
    if (!pedido.lineas || pedido.lineas.length === 0) return;
    contador++;

    pedido.lineas.forEach((linea, idx) => {
      const ctx = {
        contador: contador,
        pedido: pedido,
        linea: linea,
        posicion: (idx + 1) * CSV_SAP.PASO_POSICION
      };
      filas.push(CSV_SAP.COLUMNAS.map(col => (col.soloPrimeraLinea && idx > 0) ? "" : col.valor(ctx)));
    });
  });

  return filas.map(fila => fila.map(csvEscaparCampo_).join(CSV_SAP.SEPARADOR)).join(CSV_SAP.FIN_LINEA);
}

/**
 * Convierte la fecha del wizard ("yyyy-MM-dd") al formato CSV_SAP.FORMATO_FECHA.
 * Si llega vacía o en otro formato, se devuelve tal cual.
 */
function formatearFechaSap_(fechaIso) {
  const partes = String(fechaIso || "").trim().split("-");
  if (partes.length !== 3) return String(fechaIso || "").trim();
  return CSV_SAP.FORMATO_FECHA
    .replace("yyyy", partes[0])
    .replace("MM", partes[1])
    .replace("dd", partes[2]);
}

/**
 * Encierra en comillas los campos que contienen el separador o comillas.
 */
function csvEscaparCampo_(val) {
  const s = (val === null || val === undefined) ? "" : String(val);
  return (s.includes(CSV_SAP.SEPARADOR) || s.includes('"')) ? `"${s.replace(/"/g, '""')}"` : s;
}