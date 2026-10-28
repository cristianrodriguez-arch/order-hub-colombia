/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * ARCHIVO: main.js
 * =========================================================================================
 * QUÉ CONTIENE ESTE ARCHIVO:
 * Es el archivo central del proyecto. Reúne en un solo lugar:
 *   - La configuración general (CONFIG): IDs de la hoja de cálculo y de las carpetas de Drive,
 *     nombres de las hojas, modo de prueba y los datos de cada cliente (MEDIPIEL, CMX, CEN).
 *   - El formato del archivo CSV que se carga en SAP (CSV_SAP).
 *   - Los títulos de columnas de la hoja CONSOLIDADO (CONSOLIDADO_HEADERS).
 *   - Funciones compartidas: leer la hoja de KAM, leer destinatarios, leer el stock de bodega,
 *     validar materiales, enviar correos a las KAM y armar el CSV.
 *   - Las funciones de entrada de la aplicación web (doGet y abrirInterfaz).
 *
 * QUIÉN LO USA:
 *   - cmx_motor.js: procesa los pedidos de CMX.
 *   - cen_motor.js: procesa los pedidos de CEN.
 *   - medipiel_motor.js: procesa los pedidos de MEDIPIEL.
 *   - wizard_ui.html: asistente de la pantalla; llama a guardarNuevoDestinatario y a obtenerCatalogoMaterialesSap.
 *   - index.html: página principal; llama a contarArchivosPendientes y se carga desde doGet.
 *
 * REGLA PARA EDITAR: los datos que cambian con frecuencia (IDs de carpetas, correos, clientes,
 * nombres de hojas) están en CONFIG. Cambiarlos ahí no obliga a tocar la lógica de las funciones.
 * =========================================================================================
 */

/**
 * CONSOLIDADO_HEADERS: títulos de las columnas de la hoja CONSOLIDADO, en el orden exacto en que se escriben.
 * Se edita solo si se agrega o quita una columna de la hoja CONSOLIDADO. Si cambia la cantidad o el orden,
 * también hay que revisar medipiel_motor.js, que escribe estos títulos en la fila 3.
 * La última columna (posición 14) es "KAM".
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

/**
 * CONFIG: configuración general del sistema. Es el único lugar donde están los IDs de la hoja de cálculo
 * y de las carpetas de Google Drive, los nombres de las hojas, el modo de prueba y los datos de cada cliente.
 *
 * Qué se edita en cada caso:
 *   - Cambiar el libro o el nombre de una hoja: SPREADSHEET_ID, SHEET_PEDIDOS, SHEET_PRODUCTOS, SHEET_BODEGA, SHEET_DESTINATARIOS.
 *   - Cambiar la carpeta donde llegan los archivos de un cliente: FOLDER_ID del cliente.
 *   - Cambiar la carpeta de archivos procesados: PROCESSED_FOLDER_ID.
 *   - Cambiar el ID de cliente en SAP o la lista de precios: SOLICITANTE y LISTA_PRECIOS del cliente.
 *   - Cambiar datos fijos del pedido en SAP (clase, organización de ventas, canal, sector): bloque SAP de cada cliente.
 *   - Agregar un cliente de la cadena CEN: una entrada nueva en CLIENTES.CEN.MAPEO_CLIENTES.
 *   - Encender o apagar el modo prueba: PRUEBA.ACTIVO (y el correo de prueba en PRUEBA.CORREO_DESTINO).
 *   - Máximo de archivos por ejecución: LIMIT_FILES (0 significa usar 200).
 *   - OCR_LANGUAGE y TRASH_TEMP_OCR_DOC: ninguna función del proyecto los usa; no afectan el proceso.
 */
const CONFIG = {
  SPREADSHEET_ID: "1TAFu_ThSTccBj7uZRU3KFGWFE3mauEvqJfVfyjwInGU",
  SHEET_PEDIDOS: "CONSOLIDADO",
  SHEET_PRODUCTOS: "Productos",
  SHEET_BODEGA: "Bodega", 
  SHEET_DESTINATARIOS: "Destinatarios",
  OCR_LANGUAGE: "es",
  TRASH_TEMP_OCR_DOC: true, 
  LIMIT_FILES: 0,
  FOLDER_CSV_ID: "1aOk4DI5hEpNJwxI9JD0f-MrR6g7AWA6E",
  // MODO PRUEBA. En producción debe estar apagado (ACTIVO en false). Con ACTIVO en true, los archivos y CSV
  // generados llevan el PREFIJO_ARCHIVOS, y los correos a las KAM van a CORREO_DESTINO.
  // NOTA DE REVISIÓN: si CORREO_DESTINO está vacío, el código NO deja de enviar: manda el correo a la KAM real
  // (ver enviarCorreoKam_). Valor actual en el código: ACTIVO = true.
  PRUEBA: {
    ACTIVO: true,
    PREFIJO_ARCHIVOS: "PRUEBA_",
    CORREO_DESTINO: ""
  },

  // CLIENTES: un bloque por cliente. FOLDER_ID = carpeta de Drive donde llegan los archivos;
  // PROCESSED_FOLDER_ID = carpeta de archivos ya procesados; SOLICITANTE = ID del cliente en SAP;
  // LISTA_PRECIOS = lista de precios del pedido; SAP = datos fijos del pedido (ver CSV_SAP);
  // MAPEO_CLIENTES (solo CEN) = tabla que convierte el nombre o EAN del comprador de CEN en un cliente interno.
  CLIENTES: {
    "MEDIPIEL": {
      FOLDER_ID: "12gYDfn3yRixa92jYQ_T08FWZGLvr5kfc",
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
      FOLDER_ID: "1iNqZ-ni0PUAt_IqTpiJ0pTdEx2KwRZ_D",
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
      FOLDER_ID: "1aG4LW-EDpF3jaIM6Q19EzEG_tQhEE4u7", // carpeta de prueba CEN_sin_procesar PRUEBAS
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
 * Devuelve el nombre con el que se guarda un archivo generado (PDF o Excel procesado, o CSV).
 * En modo prueba (CONFIG.PRUEBA.ACTIVO = true) le antepone PREFIJO_ARCHIVOS, por ejemplo "PRUEBA_".
 *
 * @param {String} nombre Nombre original del archivo, por ejemplo "Pedido_123.csv".
 * @return {String} Nombre final del archivo.
 * Escribe: no escribe nada. No lanza errores.
 */
function nombreArchivoSap_(nombre) {
  return (CONFIG.PRUEBA.ACTIVO ? CONFIG.PRUEBA.PREFIJO_ARCHIVOS : "") + nombre;
}

/**
 * Envía un correo a la KAM (la persona comercial que atiende al cliente) con MailApp.
 * - Producción (CONFIG.PRUEBA.ACTIVO = false): el correo va a la dirección de opts.to.
 * - Prueba (CONFIG.PRUEBA.ACTIVO = true): el asunto lleva el prefijo "[PRUEBA] ". El destinatario es
 *   CONFIG.PRUEBA.CORREO_DESTINO si tiene un correo. Si está vacío, el código usa opts.to, es decir,
 *   la KAM real. NOTA DE REVISIÓN: el comentario de CONFIG decía que sin CORREO_DESTINO no se envía nada;
 *   el código no hace eso. Pendiente de decisión.
 *
 * @param {Object} opts Datos del correo:
 *   - to {String}: correo de la KAM (sale de la hoja Asignacion_KAM).
 *   - subject {String}: asunto del correo.
 *   - htmlBody {String}: cuerpo del correo en HTML.
 * @return {void} No devuelve nada.
 * Escribe: envía un correo electrónico. No escribe en hojas ni carpetas.
 * Lanza error si MailApp no puede enviar (por ejemplo, si se agotó la cuota diaria de correos).
 */
function enviarCorreoKam_(opts) {
  if (!CONFIG.PRUEBA.ACTIVO) {
    MailApp.sendEmail({ to: opts.to, subject: opts.subject, htmlBody: opts.htmlBody });
    return;
  }
  const destino = CONFIG.PRUEBA.CORREO_DESTINO || opts.to;
  MailApp.sendEmail({
    to: destino,
    subject: `[PRUEBA] ${opts.subject}`,
    htmlBody: opts.htmlBody
  });
}

/**
 * CSV_SAP: formato único del archivo CSV de carga a SAP, para los tres clientes (MEDIPIEL, CMX y CEN).
 * Para cambiar el formato se edita SOLO este objeto; construirCsvSap_() lo aplica.
 *
 * Qué se edita en cada caso:
 *   - SEPARADOR: carácter que separa las columnas (hoy ";").
 *   - FIN_LINEA: cómo termina cada línea (hoy el salto de línea de Windows, "\r\n").
 *   - FORMATO_FECHA: cómo se escribe la fecha de entrega (hoy dd.MM.yyyy, por ejemplo 31.12.2026).
 *   - PASO_POSICION: de cuánto en cuánto se numeran las posiciones de cada pedido (hoy 10, 20, 30...).
 *   - COLUMNAS: una entrada por columna del CSV, en orden de izquierda a derecha (columnas A a P).
 *     Cada entrada tiene:
 *       - nombre: rótulo de referencia. El CSV NO lleva fila de encabezado, porque el cargador SAP no la espera.
 *       - valor(ctx): de dónde sale el dato. ctx = { contador, pedido, linea, posicion }.
 *         Si devuelve "", la columna queda vacía.
 *       - soloPrimeraLinea: true si el dato va solo en la primera línea de cada pedido y en blanco en las siguientes.
 *
 * Columnas que hoy quedan vacías a propósito: B (Nº Pedido SAP), K (Descuento ZTK1), L (Motivo), O (EAN).
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

/**
 * Lee un archivo HTML del proyecto y devuelve su contenido como texto. Se usa en index.html con
 * <?!= include('wizard_ui'); ?> para pegar el asistente dentro de la página principal.
 *
 * @param {String} filename Nombre del archivo HTML sin la extensión .html (por ejemplo "wizard_ui").
 * @return {String} Contenido HTML del archivo.
 * Escribe: no escribe en hojas, carpetas ni correos. Lanza error si el archivo no existe.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Punto de entrada web. Cuando alguien abre la URL de la aplicación publicada como Web App,
 * esta función devuelve la página principal (index.html) con el título "Order Hub - Colombia".
 *
 * @param {Object} e Objeto de evento que Apps Script envía con la petición (parámetros de la URL). No se usa en el código.
 * @return {HtmlOutput} Página HTML lista para mostrar.
 * Escribe: no escribe en hojas, carpetas ni correos. Lanza error si index.html no existe.
 * Nota: ALLOWALL permite que cualquier sitio muestre esta página dentro de un marco (iframe).
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Order Hub - Colombia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Abre la aplicación como ventana emergente dentro de la hoja de cálculo (desde el menú de Google Sheets).
 * Usa el mismo index.html que doGet, con tamaño 900 x 700.
 *
 * @return {void} No devuelve nada.
 * Escribe: no escribe datos. Lanza error si el proyecto no está ligado a una hoja de cálculo,
 * porque SpreadsheetApp.getUi() no está disponible en ese caso.
 */
function abrirInterfaz() {
  const html = HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Order Hub Colombia')
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, 'Order Hub Colombia');
}

/**
 * Busca el KAM (persona comercial) asignado a un cliente SAP en la hoja Asignacion_KAM del libro principal.
 * Es la fuente única para nombre, iniciales y correo del KAM; la usan MEDIPIEL, CMX y CEN.
 * Compara el ID sin importar tildes, mayúsculas, espacios ni símbolos. Las columnas se reconocen por su nombre:
 *   - ID: encabezado que contiene SAP, ID o SOLICITANTE.
 *   - Nombre: encabezado que contiene KAM, ENCARGADO, NOMBRE o RESPONSABLE. Si no existe, se usa "KAM ISDIN".
 *   - Correo: encabezado que contiene CORREO, EMAIL o MAIL.
 *   - Iniciales: encabezado que contiene INICIAL, AGENTE o CODIGO. Si no existe, se usa "MC".
 *
 * @param {String|Number} solicitanteSap ID de SAP del cliente (por ejemplo "11026712").
 * @return {Object|null} { nombre: String, iniciales: String en mayúsculas, correo: String } de la primera fila
 *   que coincide. Devuelve null si el ID está vacío, si la hoja no existe o tiene menos de 2 filas,
 *   si no hay coincidencia o si ocurre un error.
 * Escribe: no escribe nada. Los errores no se lanzan: se registran en el log (console.error) y se devuelve null.
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

/**
 * Cuenta cuántos archivos están pendientes de procesar para un cliente (los que están en su carpeta de entrada).
 * Ignora los archivos en la papelera, los que tienen "_temp" en el nombre y, para CMX y CEN, todo lo que no sea
 * Excel (.xlsx, .xlsm o .xls). El resultado se guarda en la caché del script durante 60 segundos.
 *
 * @param {String} cliente Clave del cliente tal como aparece en CONFIG.CLIENTES: "MEDIPIEL", "CMX" o "CEN".
 * @return {Number} Cantidad de archivos pendientes. Devuelve 0 si el cliente no existe, si no tiene carpeta
 *   configurada o si ocurre un error (el error se registra en el log).
 * Escribe: guarda el conteo en CacheService con la clave "pendientes_<cliente>". No escribe en hojas ni correos.
 */
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

/**
 * Devuelve el listado de archivos de la carpeta de un cliente, para que los motores los procesen.
 * Si el cliente tiene FILE_QUERY, busca con esa consulta de Drive y excluye la papelera.
 * Si no la tiene, filtra por tipo de archivo (MIME_TYPE; por defecto PDF).
 *
 * @param {Folder} folder Carpeta de Google Drive del cliente (CONFIG.CLIENTES[cliente].FOLDER_ID).
 * @param {Object} configCli Bloque de configuración del cliente en CONFIG.CLIENTES.
 * @return {FileIterator} Iterador de archivos de Drive (se recorre con hasNext() y next()).
 * Escribe: no escribe nada. Lanza error si Drive no puede leer la carpeta.
 */
function obtenerArchivosCliente_(folder, configCli) {
  if (configCli.FILE_QUERY) {
    // Filtro Poka-Yoke: Asegurar que nunca busque archivos en la papelera
    const query = "(" + configCli.FILE_QUERY + ") and trashed = false";
    return folder.searchFiles(query);
  }
  return folder.getFilesByType(configCli.MIME_TYPE || MimeType.PDF);
}

/**
 * Lee la hoja Destinatarios y crea un diccionario para buscar un destino por clave.
 * Cada fila se guarda por el valor de cruce_cliente (si tiene) y por id_destino (si tiene).
 * La ciudad sale de la columna que contiene "ciudad", "poblacion" o "municipio"; si la celda está vacía, se usa "COTA".
 *
 * @return {Object} Diccionario { "<clave>": { idDestino: String, nombreDestino: String, ciudad: String } }.
 *   Devuelve {} si no existe la hoja o si solo tiene encabezados.
 * Escribe: no escribe nada. Lanza error si no puede abrir el libro CONFIG.SPREADSHEET_ID.
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

    if (cruce) dicc[cruce] = objInfo;
    if (idDestino) dicc[idDestino] = objInfo;
  }
  return dicc;
}

/**
 * Resume el stock de la hoja Bodega por código EAN. Cada EAN tiene una lista de lotes; si dos filas tienen
 * el mismo material y la misma fecha, sus cantidades se suman. Las filas sin EAN se ignoran.
 * Las columnas se buscan por nombre de encabezado; si no se encuentran, se usan posiciones fijas
 * (EAN = columna J, Material = columna A, Libre utilización = columna F, Fecha de caducidad = columna H).
 *
 * @param {Sheet} sheet Hoja Bodega. Si es null o no existe, devuelve {}.
 * @return {Object} Mapa { "<EAN>": [ { material: String, stock: Number, fechaDisplay: String } ] }.
 *   fechaDisplay es "yyyy-MM" si la celda es fecha, el texto original si es texto, y "Sin fecha" si está vacía.
 * Escribe: no escribe nada. Lanza error si no puede leer la hoja.
 */
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
 * @param {Object} datos Datos del destinatario; todos son texto:
 *   - idOrigen: ID del cliente de origen (SAP).
 *   - nombreOrigen: nombre del cliente de origen.
 *   - idDestino: ID del destino de la mercancía.
 *   - nombreDestino: nombre del destino.
 *   - cruceCliente: clave que usa el archivo del cliente para encontrar este destino (cruce_cliente).
 *   - ciudad: ciudad o municipio del destino.
 * @return {Object} { success: true } si se guardó.
 * Escribe: agrega una fila nueva al final de la hoja Destinatarios. Si la hoja no existe, la crea con sus encabezados.
 * Lanza error si a la hoja le faltan columnas; en ese caso no escribe nada.
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
 * @return {Object} Catálogo { "<material>": { stock: Number, descripcion: String } }. El stock es la suma
 *   de la hoja Bodega; los materiales que solo están en Productos quedan con stock 0.
 * Escribe: no escribe nada. Lanza error si no existe la hoja Productos o si no tiene columna de Material.
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
 * @return {String} Contenido del CSV, sin fila de encabezado. Las líneas van separadas por CSV_SAP.FIN_LINEA
 *   y el texto no termina con salto de línea.
 * Escribe: no escribe en hojas ni carpetas; solo arma el texto. Lanza error (TypeError) si un pedido no trae
 *   el bloque SAP de su cliente.
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
 * Convierte la fecha del wizard ("yyyy-MM-dd") al formato de CSV_SAP.FORMATO_FECHA (hoy dd.MM.yyyy).
 * Ejemplo: "2026-12-31" pasa a "31.12.2026".
 * Si llega vacía o en otro formato, se devuelve tal cual, sin validar.
 *
 * @param {String} fechaIso Fecha en formato "yyyy-MM-dd".
 * @return {String} Fecha en el formato del CSV, o el texto original si no se pudo convertir.
 * Escribe: no escribe nada. No lanza errores.
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
 * Prepara un valor para una celda del CSV: si contiene el separador (;) o comillas dobles, lo encierra
 * entre comillas y duplica las comillas internas. Si el valor es vacío, null o undefined, devuelve "".
 *
 * @param {*} val Valor a escribir (texto, número o vacío).
 * @return {String} Texto listo para la celda del CSV.
 * Escribe: no escribe nada. No lanza errores.
 * Nota: no encierra en comillas los valores con saltos de línea; un salto dentro de un valor partiría la fila.
 */
function csvEscaparCampo_(val) {
  const s = (val === null || val === undefined) ? "" : String(val);
  return (s.includes(CSV_SAP.SEPARADOR) || s.includes('"')) ? `"${s.replace(/"/g, '""')}"` : s;
}