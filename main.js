/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: CONFIGURACIÓN CENTRALIZADA Y CONTROLADOR DE VISTAS (COLOMBIA)
 * ARCHIVO: main.gs
 * ESTADO: PRODUCCIÓN - CON DOBLE INDEXACIÓN DE DESTINATARIOS
 * =========================================================================================
 * 📝 DESCRIPCIÓN TÉCNICA:
 * Centraliza configuraciones, parámetros de SAP y utilidades globales.
 * Incluye la doble indexación (por Cruce_cliente y por Id_destino) en Destinatarios
 * para garantizar la correcta extracción de la ciudad (ej: LA ESTRELLA).
 * =========================================================================================
 */

/* STREAMING_CHUNK: Definiendo la cabecera oficial del Consolidado... */
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

/* STREAMING_CHUNK: Inicializando el objeto de configuración global CONFIG... */
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
      FOLDER_ID: "1hCjs2bI3YvvmcqIOh-NufG9CCh2Zx3RF", 
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
          aliases: ["Distribuidor"]
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
      }
    }
  }
};

/* STREAMING_CHUNK: Métodos para renderizado de la aplicación web... */
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

function contarArchivosPendientes(cliente) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "pendientes_" + cliente;
  const cachedCount = cache.get(cacheKey);

  if (cachedCount !== null) return parseInt(cachedCount);

  const configCli = CONFIG.CLIENTES[cliente];
  if (!configCli) return 0;

  const folder = DriveApp.getFolderById(configCli.FOLDER_ID);
  const files = obtenerArchivosCliente_(folder, configCli);

  let count = 0;
  while (files.hasNext()) {
    files.next();
    count++;
  }

  cache.put(cacheKey, count.toString(), 600);
  return count;
}

function obtenerArchivosCliente_(folder, configCli) {
  if (configCli.FILE_QUERY) {
    return folder.searchFiles(configCli.FILE_QUERY);
  }
  return folder.getFilesByType(configCli.MIME_TYPE || MimeType.PDF);
}

/* STREAMING_CHUNK: Creando el diccionario de Destinatarios con doble indexación (Cruce_cliente e Id_destino)... */
/**
 * Lee la hoja de Destinatarios e indexa por Cruce_cliente e Id_destino.
 * De esta manera, si la búsqueda se hace por "BODEGA_CMX_PRINCIPAL" o por "11033482",
 * retornará de inmediato la ciudad correcta (ej: LA ESTRELLA).
 * 
 * @return {Object} Diccionario mapeado por Cruce_cliente e Id_destino.
 */
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

    // Doble indexación para garantizar la coincidencia sin importar cuál clave use el motor
    if (cruce) dicc[cruce] = objInfo;
    if (idDestino) dicc[idDestino] = objInfo;
  }
  return dicc;
}

/* STREAMING_CHUNK: Creando el mapa de existencias e inventario de Bodega... */
function obtenerMapaBodegaGeneral_(sheet) {
  const map = {}; 
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