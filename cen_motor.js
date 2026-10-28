/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * ARCHIVO: cen_motor.js
 * MÓDULO: Integración de pedidos CEN (Centro de Negocios de Carvajal), multicliente
 * =========================================================================================
 * QUÉ CONTIENE
 * El portal CEN de Carvajal entrega un Excel con órdenes de compra (OC) de varios clientes de
 * ISDIN. Este archivo:
 *  1. Lee los Excel de la carpeta de entrada de CEN y arma la lista de pedidos y productos.
 *  2. Registra cada cliente como PROVISIONAL en la hoja CONSOLIDADO (columna V).
 *  3. Confirma todo al generar UN solo CSV de carga a SAP, y pasa las filas a definitivas.
 *  4. Mueve los Excel terminados a la carpeta de procesados y envía los correos a las KAM.
 *
 * QUIÉN LO LLAMA
 *  - wizard_ui.html (asistente de carga): llama a preAnalizarCEN, cenGuardarProvisional y
 *    guardarDefinitivoCEN con google.script.run.
 *  - main.js: no llama a estas funciones; es al revés. Este archivo usa de main.js la
 *    configuración (CONFIG), los encabezados de CONSOLIDADO y las funciones comunes
 *    (construirCsvSap_, enviarCorreoKam_, nombreArchivoSap_, obtenerInfoKamPorSolicitante_,
 *    obtenerDiccionarioDestinatarios, obtenerMapaBodegaGeneral_, obtenerArchivosCliente_).
 *  - medipiel_motor.js: aporta medipiel_getHeaderMap_ y medipiel_createRowArray_.
 * =========================================================================================
 */

/**
 * Punto de entrada del paso 1 de CEN: lee los Excel de la carpeta de entrada de CEN en Drive y
 * devuelve los pedidos y productos que el asistente muestra.
 * Antes de leer, borra las filas PROVISIONALES de CONSOLIDADO de TODAS las sesiones (quedaron de un
 * análisis anterior, por ejemplo si se recargó la página), para que los consecutivos no se cuenten dos veces.
 *
 * ESCRIBE: borra filas provisionales en CONSOLIDADO. Los archivos temporales de conversión se crean
 * y borran dentro de cen_motorLecturaExcel. No envía correos.
 * LEE: CONSOLIDADO, Bodega, Productos, Destinatarios, Asignacion_KAM (indirectamente) y la carpeta de CEN.
 *
 * @return {Object} { clientesDetectados, pedidosExtraidos, itemsExtraidos, errores, yaProcesados, estadoArchivos }.
 *   Si la carpeta no tiene Excel, devuelve listas vacías (sin estadoArchivos).
 * @throws {Error} Si falta la configuración de CEN (CONFIG.CLIENTES.CEN o su FOLDER_ID) o si falla cualquier
 *   paso. El mensaje empieza con "Error en preAnalizarCEN: ".
 */
function preAnalizarCEN() {
  try {
    const configCli = CONFIG.CLIENTES["CEN"];
    if (!configCli || !configCli.FOLDER_ID) {
      throw new Error("La configuración de CEN no está definida en CONFIG de main.gs.");
    }

    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");

    // Un análisis nuevo borra las provisionales que dejó una sesión anterior (p. ej. si se recargó la página)
    cen_conBloqueo_(() => cen_borrarProvisionales_(shConsolidado, null));

    // 2. Leer las OC + SAP ya registradas en CONSOLIDADO, para no volver a cargarlas.
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

    // 3. Cargar destinatarios y bodega.
    const diccDestinatarios = obtenerDiccionarioDestinatarios();
    const diccBodega = obtenerMapaBodegaGeneral_(ss.getSheetByName(CONFIG.SHEET_BODEGA || "Bodega"));

    // 4. Listar los Excel de la carpeta de CEN (.xls, .xlsx, .xlsm), hasta el límite configurado.
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

    // 5. Leer los Excel y armar pedidos y productos.
    return cen_motorLecturaExcel(archivos, diccBodega, configCli, pedidosProcesados, diccDestinatarios, ss);

  } catch (err) {
    throw new Error("Error en preAnalizarCEN: " + err.message);
  }
}

/**
 * Lee cada Excel de CEN y arma la lista de pedidos (OC) y de productos (items) que muestra el asistente.
 * Para cada archivo: lo convierte a una hoja temporal, lee sus filas y borra la copia temporal al final,
 * haya o no error.
 *
 * Reglas principales:
 *  - Si el comprador de una OC no está en MAPEO_CLIENTES, el archivo queda BLOQUEADO (se agrega un
 *    mensaje a errores y el archivo no se mueve a procesados).
 *  - Si la combinación OC + Solicitante SAP ya está en CONSOLIDADO, la OC se omite y se anota en yaProcesados.
 *  - Cada producto se busca primero en la tarifa del cliente y, si no está, en el catálogo general.
 *    Si el EAN existe en catálogo, se toma su Material SAP. Si no tiene saldo en Bodega, queda como
 *    back order (stock 0).
 *  - Las líneas con cantidad 0 o sin EAN se ignoran.
 *  - Un error en un archivo no detiene a los demás: se registra en errores y ese archivo queda bloqueado.
 *
 * ESCRIBE: crea y borra archivos temporales en la carpeta de cada Excel. No escribe en CONSOLIDADO.
 *
 * @param {Array<DriveApp.File>} archivos Excel de CEN a leer.
 * @param {Object} diccBodega Mapa EAN -> lotes con stock (de obtenerMapaBodegaGeneral_).
 * @param {Object} configCli Configuración de CEN (CONFIG.CLIENTES.CEN): MAPEO_CLIENTES, SAP, carpetas, etc.
 * @param {Set<String>} pedidosProcesados Combinaciones "OC_SAP" ya registradas en CONSOLIDADO.
 * @param {Object} diccDestinatarios Destinatarios indexados por EAN de entrega o por Solicitante SAP.
 * @param {Spreadsheet} ss Libro principal (CONFIG.SPREADSHEET_ID), para leer la hoja Productos.
 * @return {Object} { clientesDetectados, pedidosExtraidos, itemsExtraidos, errores, yaProcesados, estadoArchivos }.
 *   estadoArchivos: por id de archivo, { ocsPendientes, bloqueado }. guardarDefinitivoCEN lo usa para
 *   decidir si el Excel se mueve a procesados.
 * @throws No lanza error por un archivo con problemas: los registra en errores y sigue con los demás.
 */
function cen_motorLecturaExcel(archivos, diccBodega, configCli, pedidosProcesados, diccDestinatarios, ss) {
  const pedidosExtraidos = [];
  const itemsExtraidos = [];
  const errores = [];
  const yaProcesados = [];
  const mapaClientesCount = {};
  // Por archivo: OCs nuevas que contiene y si tiene algo que impide darlo por terminado.
  // guardarDefinitivoCEN solo mueve el Excel a Procesados si todas sus OCs quedaron registradas.
  const estadoArchivos = {};

  // Catálogo de Productos: se lee una sola vez para todos los archivos.
  const shProd = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  const { catalogoPorTarifa, catalogoGlobal } = cen_getCatalogoProductosGlobal_(shProd);

  archivos.forEach((file) => {
    let tempFileId = null;
    const estado = { ocsPendientes: [], bloqueado: false };
    estadoArchivos[file.getId()] = estado;
    try {
      // 1. Convertir el Excel a hoja temporal y leer sus filas.
      tempFileId = cen_convertirExcelASheet_(file);
      const tempSs = SpreadsheetApp.openById(tempFileId);
      const sheet = tempSs.getSheets()[0];
      const data = sheet.getDataRange().getValues();

      // Archivo vacío o solo con encabezado: no hay nada que leer.
      if (data.length < 2) return;

      const head = data[0].map(h => cen_normalizeKey_(h));
      
      // 2. Ubicar las columnas por el nombre del encabezado, no por su posición.
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

      // 3. Si faltan columnas esenciales (OC, EAN o cantidad), el archivo queda bloqueado.
      if (colOc === -1 || colEanItem === -1 || colCant === -1) {
        throw new Error(`El archivo ${file.getName()} no contiene la estructura esperada de columnas CEN.`);
      }

      // 4. Recorrer las filas de datos (la fila 1 son los encabezados).
      for (let r = 1; r < data.length; r++) {
        const row = data[r];
        if (!row.join("").trim()) continue;

        const ocRaw = cen_limpiarOc_(row[colOc]);
        if (!ocRaw) continue;

        const eanComprador = colCompradorEan !== -1 ? String(row[colCompradorEan]).trim() : "";
        const nombreComprador = colCompradorNombre !== -1 ? String(row[colCompradorNombre]).trim() : "CLIENTE_CEN";
        
        // a) Homologar el comprador. Si no está en MAPEO_CLIENTES, el archivo queda bloqueado.
        const clienteInternoObj = cen_identificarClienteInterno_(eanComprador, nombreComprador, configCli);
        if (!clienteInternoObj) {
          // No se inventa un Solicitante SAP: la OC se bloquea hasta registrar el comprador en
          // CONFIG.CLIENTES.CEN.MAPEO_CLIENTES, y el Excel se queda en la carpeta de entrada.
          estado.bloqueado = true;
          const msgError = `Cliente no homologado: "${nombreComprador}" (EAN ${eanComprador || "N/A"}), OC ${ocRaw}, archivo ${file.getName()}. Agréguelo a MAPEO_CLIENTES en main.gs.`;
          if (!errores.includes(msgError)) errores.push(msgError);
          continue;
        }
        const clienteKey = clienteInternoObj.key;
        const solicitanteSap = clienteInternoObj.solicitante;

        // Clave combinada: OC + SAP ID del Cliente
        const comboKey = `${ocRaw}_${cen_normalizeKey_(solicitanteSap)}`;

        // b) Si esta OC + SAP ya están en CONSOLIDADO, se omite (ya fue registrada).
        if (pedidosProcesados.has(comboKey)) {
          if (!yaProcesados.includes(ocRaw)) yaProcesados.push(ocRaw);
          continue;
        }

        // c) Crear la cabecera del pedido solo la primera vez que aparece la OC.
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

          if (!estado.ocsPendientes.includes(comboKey)) estado.ocsPendientes.push(comboKey);

          pedidoObj = {
            Pedido_ID: ocRaw,
            Combo_Key: comboKey,
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

        // d) Datos de la línea: EAN, descripción y cantidad.
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

        // e) Si el producto existe en catálogo, se toma su Material SAP oficial y su precio.
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
      // Error en este archivo: queda bloqueado y se sigue con el siguiente.
      estado.bloqueado = true;
      errores.push(`${file.getName()}: ${e.message}`);
    } finally {
      // Borrar siempre la copia temporal, haya o no error.
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
    yaProcesados: yaProcesados,
    estadoArchivos: estadoArchivos
  };
}

/**
 * Limpia el número de orden de compra (OC) para que siempre sea el mismo texto, sin importar cómo lo guardó el Excel.
 * Quita espacios y la comilla inicial. Si viene en notación científica (por ejemplo 4.5E+09) lo convierte a
 * número entero. Quita la parte decimal (".0").
 *
 * @param {*} val Valor de la celda (texto, número o vacío).
 * @return {String} OC limpia, o "" si la celda estaba vacía.
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
 * Homologa el comprador CEN contra CONFIG.CLIENTES.CEN.MAPEO_CLIENTES. Busca primero por EAN del comprador
 * y, si no coincide, por alias contenido en el nombre del comprador. Devuelve el primer cliente que coincida.
 * No escribe nada ni lanza errores.
 *
 * @param {String} eanComprador EAN de la empresa compradora (puede venir vacío).
 * @param {String} nombreComprador Razón social del comprador, tal como viene en el Excel.
 * @param {Object} configCli Configuración de CEN (CONFIG.CLIENTES.CEN).
 * @return {Object|null} { key, nombreDisplay, solicitante, listaPrecios }, o null si el comprador no está en
 *   MAPEO_CLIENTES. Nunca se inventa un solicitante ni una lista de precios.
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

  return null;
}

/**
 * Lee la hoja Productos y arma dos catálogos para buscar productos por EAN:
 *  1. catalogoPorTarifa[tarifa][ean]: material y precio por lista de precios del cliente
 *     (si un EAN se repite dentro de la misma tarifa, gana la última fila).
 *  2. catalogoGlobal[ean]: red de seguridad sin importar la tarifa. Si un EAN aparece varias veces, se queda
 *     con la primera fila, salvo que esa no tenga precio y una posterior sí.
 * Las columnas se detectan por nombre de encabezado: EAN o BARRA; MATERIAL o SAP; DESCRIPCION, NOMBRE o
 * PRODUCTO; LISTA o CLIENTE; COSTO, PRECIO o VALOR. Si falta la hoja o las columnas de EAN o Material,
 * devuelve catálogos vacíos. No escribe ni lanza errores.
 *
 * @param {SpreadsheetApp.Sheet|null} shProd Hoja Productos (puede ser null).
 * @return {Object} { catalogoPorTarifa, catalogoGlobal }.
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
 * =========================================================================================
 * REGISTRO PROVISIONAL DE CEN
 * -----------------------------------------------------------------------------------------
 * Cada cliente CEN, al pulsar "Guardar ... y pasar al siguiente", se registra en CONSOLIDADO
 * como PROVISIONAL: su columna V lleva la marca "PROVISIONAL|<sesión>".
 * - Así, el siguiente cliente con las mismas iniciales continúa el consecutivo.
 * - Las filas solo quedan definitivas (se quita la marca) cuando se genera el CSV unificado.
 * - Las provisionales se borran de CONSOLIDADO si no se puede guardar el CSV, si falla la confirmación,
 *   o al empezar un análisis CEN nuevo (de cualquier sesión, por ejemplo tras recargar la página).
 *   Al confirmar, también se borran las provisionales de la sesión que no están en el paquete.
 * - Las filas de datos de CONSOLIDADO empiezan en la fila 4 (la 3 son los encabezados).
 * =========================================================================================
 */

const CEN_PROVISIONAL = "PROVISIONAL";

/**
 * Arma el texto que se escribe en la columna V de CONSOLIDADO para una orden provisional de una sesión.
 * No escribe nada; solo devuelve el texto.
 *
 * @param {String} sessionId Identificador de la sesión del asistente.
 * @return {String} Por ejemplo "PROVISIONAL|cen-lx2k3-ab12cd".
 */
function cen_marcaProvisional_(sessionId) {
  return `${CEN_PROVISIONAL}|${sessionId}`;
}

/**
 * Indica si el contenido de la columna V es una marca provisional.
 * Sin sessionId, cuenta cualquier marca PROVISIONAL de cualquier sesión. Con sessionId, solo cuenta la
 * marca exacta de esa sesión. No escribe ni lanza errores.
 *
 * @param {*} valor Contenido de la celda.
 * @param {String} [sessionId] Opcional. Si se indica, solo cuenta la marca de esa sesión.
 * @return {Boolean} true si la celda tiene una marca provisional.
 */
function cen_esProvisional_(valor, sessionId) {
  const texto = String(valor || "");
  return sessionId ? texto === cen_marcaProvisional_(sessionId) : texto.indexOf(CEN_PROVISIONAL + "|") === 0;
}

/**
 * CEN_COL_MARCA: columna V de CONSOLIDADO (número base 1) donde va la marca PROVISIONAL.
 * Está fuera de los encabezados de CONSOLIDADO (CONSOLIDADO_HEADERS de main.js). Solo guarda la marca
 * PROVISIONAL; no escribe datos del pedido.
 */
const CEN_COL_MARCA = 22;

/**
 * Ejecuta una función con el bloqueo del script de Apps Script. Así dos clics o dos usuarios no escriben
 * en CONSOLIDADO al mismo tiempo, lo que daría dos consecutivos iguales.
 * Espera hasta 30 segundos el bloqueo. Si no lo obtiene, la función no se ejecuta. El bloqueo se libera
 * siempre al final, aunque fn falle.
 *
 * @param {Function} fn Función a ejecutar dentro del bloqueo (sin parámetros).
 * @return {*} Lo que devuelve fn.
 * @throws Si no se obtiene el bloqueo en 30 segundos, o si fn lanza un error (se pasa tal cual).
 */
function cen_conBloqueo_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Busca en CONSOLIDADO las filas que tienen marca provisional en la columna V. Las filas de datos empiezan
 * en la fila 4 (la 3 son los encabezados). De cada fila lee la OC (columna B) y el Sap Id (columna C).
 * No escribe nada.
 *
 * @param {SpreadsheetApp.Sheet} sh Hoja CONSOLIDADO.
 * @param {String} [sessionId] Opcional. Si se omite, lista las provisionales de todas las sesiones.
 * @return {Array<{fila: Number, oc: String, sap: String, combo: String}>}
 *   fila: número de fila en la hoja (base 1). oc y sap: valores limpios. combo: "<OC>_<SAP normalizado>",
 *   igual que Combo_Key. Lista vacía si no hay filas.
 */
function cen_listarProvisionales_(sh, sessionId) {
  const ultima = sh.getLastRow();
  if (ultima < 4) return [];
  const marcas = sh.getRange(4, CEN_COL_MARCA, ultima - 3, 1).getValues();
  const datos = sh.getRange(4, 1, ultima - 3, 3).getValues();
  const resultado = [];
  marcas.forEach((marca, i) => {
    if (!cen_esProvisional_(marca[0], sessionId)) return;
    const oc = cen_limpiarOc_(datos[i][1]);
    const sap = cen_normalizeKey_(datos[i][2]);
    resultado.push({ fila: i + 4, oc: oc, sap: sap, combo: `${oc}_${sap}` });
  });
  return resultado;
}

/**
 * Borra de CONSOLIDADO las filas provisionales de una sesión (o de todas, si no se indica sesión).
 * Se borra de abajo hacia arriba para que los números de fila no cambien mientras se borra.
 *
 * ESCRIBE: borra filas en CONSOLIDADO.
 *
 * @param {SpreadsheetApp.Sheet} sh Hoja CONSOLIDADO. Si es null, no hace nada.
 * @param {String} [sessionId] Opcional. Si se omite, borra las de todas las sesiones.
 * @return {Number} Cantidad de filas borradas.
 * @throws Si Google Sheets falla al borrar alguna fila.
 */
function cen_borrarProvisionales_(sh, sessionId) {
  if (!sh) return 0;
  const provisionales = cen_listarProvisionales_(sh, sessionId);
  for (let i = provisionales.length - 1; i >= 0; i--) {
    sh.deleteRow(provisionales[i].fila);
  }
  if (provisionales.length > 0) SpreadsheetApp.flush();
  return provisionales.length;
}

/**
 * Registra en CONSOLIDADO, como PROVISIONALES, las órdenes de UN cliente CEN. Lo llama el asistente al pulsar
 * "Guardar ... y pasar al siguiente cliente" (y al último cliente, antes de generar el CSV).
 *
 * El consecutivo se calcula leyendo CONSOLIDADO dentro del bloqueo, así que cuenta también las provisionales
 * de los clientes anteriores de la misma sesión. El formato es INICIALES-AA + 4 dígitos, por ejemplo
 * MC-260001 (AA = últimos dos dígitos del año). Un pedido sin unidades no se registra y no gasta consecutivo.
 *
 * ESCRIBE: agrega filas nuevas en CONSOLIDADO (columnas A a K) y pone la marca PROVISIONAL en la columna V
 * de cada una. No envía correos ni crea CSV.
 * LEE: CONSOLIDADO y Asignacion_KAM (para las iniciales del KAM, mediante cen_getKamInfo_).
 *
 * @param {String} sessionId Identificador de la sesión del asistente. Es obligatorio.
 * @param {Object} conf Datos del cliente:
 *   - pedidos: pedidos de ese cliente (Pedido_ID, Cliente_Key, Solicitante_SAP, Cliente_Nombre,
 *     Fecha_Pedido, Id_Destino, Destinatario, Ciudad).
 *   - items: posiciones de esos pedidos (Pedido_ID, Cliente_Key, Cant, Valor_Unitario).
 *   - fechasEntregas: objeto { <Pedido_ID>: { fecha: ... } }. La fecha se copia tal cual. Si falta, la celda queda vacía.
 * @return {Object} { ids: [{ Id, Pedido_ID, Solicitante_SAP }] }: consecutivos asignados, uno por pedido registrado.
 * @throws {Error} "Falta el identificador de sesión CEN.", "No se encontró la hoja CONSOLIDADO." o errores de
 *   Google Sheets o de datos incompletos en conf.
 */
function cenGuardarProvisional(sessionId, conf) {
  return cen_conBloqueo_(() => {
    if (!sessionId) throw new Error("Falta el identificador de sesión CEN.");
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    if (!sh) throw new Error("No se encontró la hoja CONSOLIDADO.");

    // 1. Preparar: año de dos dígitos, columnas que se escriben (A a K) y mapa de encabezados (fila 3).
    const sufijoAno = String(new Date().getFullYear()).slice(-2);
    // Solo se escriben las columnas A:K. Desde la L en adelante hay fórmulas y formatos condicionales.
    const numCols = 11;
    const headerMap = medipiel_getHeaderMap_(sh, 3);
    const maxPorPrefijo = {};
    const filas = [];
    const ids = [];

    // 2. Por cada pedido: KAM, consecutivo, unidades y valor, y fila nueva.
    conf.pedidos.forEach(p => {
      const kamInfo = cen_getKamInfo_(p.Solicitante_SAP, p.Cliente_Nombre);
      const prefix = (kamInfo.initials || "MC") + "-" + sufijoAno;
      // Se lee CONSOLIDADO una vez por prefijo; después se cuenta en memoria
      if (maxPorPrefijo[prefix] === undefined) maxPorPrefijo[prefix] = cen_obtenerMaxConsecutivo_(sh, prefix);

      // 3. Sumar unidades y valor de las posiciones del pedido. Sin unidades, no se registra.
      let totalUnidades = 0;
      let totalValor = 0;
      conf.items.filter(i => i.Pedido_ID === p.Pedido_ID && i.Cliente_Key === p.Cliente_Key).forEach(item => {
        const cant = parseInt(item.Cant) || 0;
        totalUnidades += cant;
        totalValor += cant * (parseFloat(item.Valor_Unitario) || 0);
      });
      if (totalUnidades === 0) return;

      // 4. Siguiente consecutivo del prefijo y armado de la fila.
      maxPorPrefijo[prefix]++;
      const id = prefix + String(maxPorPrefijo[prefix]).padStart(4, "0");
      const fechaEntrega = conf.fechasEntregas[p.Pedido_ID] ? conf.fechasEntregas[p.Pedido_ID].fecha : "";

      filas.push(medipiel_createRowArray_(headerMap, {
        "Id": id,
        "Orden De Compra": p.Pedido_ID,
        "Sap Id": p.Solicitante_SAP,
        "Solicitante": p.Cliente_Nombre,
        "Fecha Pedido": p.Fecha_Pedido,
        "Id Destino": p.Id_Destino,
        "Destinatario": p.Destinatario,
        "Ciudad": p.Ciudad,
        "Fecha entrega": fechaEntrega,
        "Unidades": totalUnidades,
        "Valor": totalValor
      }, numCols));
      ids.push({ Id: id, Pedido_ID: p.Pedido_ID, Solicitante_SAP: p.Solicitante_SAP });
    });

    // 5. Escribir las filas nuevas y la marca PROVISIONAL en la columna V.
    if (filas.length > 0) {
      const primera = sh.getLastRow() + 1;
      sh.getRange(primera, 1, filas.length, numCols).setValues(filas);
      // Marca PROVISIONAL en la columna V de cada fila nueva (fuera de A:K, no toca las fórmulas)
      filas.forEach((_, i) => sh.getRange(primera + i, CEN_COL_MARCA).setValue(cen_marcaProvisional_(sessionId)));
      SpreadsheetApp.flush();
    }
    return { ids: ids };
  });
}

/**
 * Confirma el lote CEN de todos los clientes de la sesión. Es el paso final: genera el CSV unificado para SAP,
 * confirma las órdenes provisionales y envía los correos a las KAM.
 *
 * Orden de los pasos:
 *  1. Se guarda el CSV en Drive. Si falla, las provisionales se borran y no queda nada registrado.
 *  2. Las provisionales de la sesión pasan a definitivas (se quita la marca de la columna V) y se
 *     borran las que sobran.
 *  3. Los Excel completos se mueven a la carpeta de procesados y se renombran. Un Excel con alguna OC
 *     sin registrar, o con un comprador sin homologar, se deja en la carpeta de entrada.
 *  4. Se envían los correos a las KAM, solo por las órdenes confirmadas.
 *
 * Reglas del CSV: un pedido sin unidades no entra. Una línea sin Material SAP válido (vacío, con "SIN_",
 * igual al EAN, o un EAN de 13 dígitos que empieza por 84) va con material vacío, que SAP trata como back order.
 *
 * ESCRIBE: CSV en CONFIG.FOLDER_CSV_ID; CONSOLIDADO (quita marcas o borra filas provisionales); mueve y
 * renombra Excel en la carpeta de procesados; envía correos a las KAM.
 *
 * @param {Object} paquete Todo lo que arma el asistente:
 *   - sessionId: identificador de la sesión (obligatorio).
 *   - pedidos: pedidos de todos los clientes (Cliente_Key, Pedido_ID, Combo_Key, Solicitante_SAP,
 *     Id_Destino, Archivo_Id, Archivo_Nombre, Fecha_Pedido, Cliente_Nombre).
 *   - items: posiciones (Item_ID, Pedido_ID, Cliente_Key, Producto_Cod, Descripcion, Cant, Valor_Unitario).
 *   - fechasEntregas: fecha de entrega por Pedido_ID.
 *   - materialesSeleccionados: { Item_ID: Material SAP elegido }.
 *   - noActualizadosSeleccionados: { Item_ID: true } marca líneas sin cruce de EAN.
 *   - backordersSeleccionados: { Item_ID: true } marca líneas en back order (para el correo).
 *   - estadoArchivos: estado de cada Excel devuelto por preAnalizarCEN.
 *   - notificarKAM: si es false, no se envían correos (por defecto sí se envían).
 * @return {Object} { success: true, mensaje, csvData (CSV en base64, sin saltos de línea), fileName }.
 *   fileName tiene la forma Pedido_CEN_AAAAMMDD.csv (con el prefijo de prueba si aplica).
 * @throws {Error} Si falta sessionId, si no existe CONSOLIDADO, si no hay provisionales de la sesión para
 *   confirmar, si el CSV no se puede guardar, o si falla la confirmación en CONSOLIDADO. Los errores de
 *   correo NO detienen el proceso: se registran en el log.
 */
function guardarDefinitivoCEN(paquete) {
  try {
    return cen_conBloqueo_(() => {
      if (!paquete.sessionId) throw new Error("Falta el identificador de sesión CEN.");
      const configCli = CONFIG.CLIENTES["CEN"];
      const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
      const sh = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
      if (!sh) throw new Error("No se encontró la hoja CONSOLIDADO.");

      // Preparación: líneas del CSV por pedido y lista de notificaciones por cliente.
      const pedidosCsv = [];
      const notificaciones = [];
      const combosPaquete = new Set();

      // Líneas del CSV por cliente y orden. Las órdenes sin unidades no se registran.
      const pedidosPorCliente = {};
      paquete.pedidos.forEach(p => {
        if (!pedidosPorCliente[p.Cliente_Key]) pedidosPorCliente[p.Cliente_Key] = [];
        pedidosPorCliente[p.Cliente_Key].push(p);
      });

      for (const cliKey in pedidosPorCliente) {
        const pedidosCli = pedidosPorCliente[cliKey];
        const primerPedido = pedidosCli[0];
        const solicitanteSap = primerPedido.Solicitante_SAP;
        const kamInfo = cen_getKamInfo_(solicitanteSap, primerPedido.Cliente_Nombre);

        pedidosCli.forEach(p => {
          // Por OC y cliente: dos clientes pueden traer el mismo número de OC en el mismo lote
          const itemsPedido = paquete.items.filter(i => i.Pedido_ID === p.Pedido_ID && i.Cliente_Key === p.Cliente_Key);
          let totalUnidades = 0;

          const pedidoCsv = {
            sap: configCli.SAP,
            solicitante: solicitanteSap,
            destinatario: p.Id_Destino,
            numPedidoCliente: p.Pedido_ID,
            fechaEntrega: paquete.fechasEntregas[p.Pedido_ID] ? paquete.fechasEntregas[p.Pedido_ID].fecha : "",
            lineas: []
          };
          pedidosCsv.push(pedidoCsv);

          itemsPedido.forEach(item => {
            const matSeleccionado = (paquete.materialesSeleccionados && paquete.materialesSeleccionados[item.Item_ID]) ? String(paquete.materialesSeleccionados[item.Item_ID]).trim() : "";
            // Sin Material SAP (vacío, SIN_ o EAN de 13 dígitos): la línea va al CSV con material vacío => back order en SAP
            const isNoActualizado = (paquete.noActualizadosSeleccionados && paquete.noActualizadosSeleccionados[item.Item_ID]) ||
                                    (!matSeleccionado || matSeleccionado.includes("SIN_") || matSeleccionado === item.Producto_Cod || (matSeleccionado.length === 13 && matSeleccionado.startsWith("84")));
            const cant = parseInt(item.Cant) || 0;
            totalUnidades += cant;
            pedidoCsv.lineas.push({ material: isNoActualizado ? "" : matSeleccionado, cantidad: cant });
          });

          if (totalUnidades > 0) combosPaquete.add(p.Combo_Key);
        });

        notificaciones.push({ pedidosCli: pedidosCli, kamInfo: kamInfo, clienteNombre: primerPedido.Cliente_Nombre, solicitanteSap: solicitanteSap });
      }

      // Filas provisionales de esta sesión: se confirman las que siguen en el paquete; el resto se borra.
      const provisionales = cen_listarProvisionales_(sh, paquete.sessionId);
      const aConfirmar = provisionales.filter(r => combosPaquete.has(r.combo));
      if (aConfirmar.length === 0) {
        throw new Error("No hay órdenes provisionales de esta sesión para confirmar. Vuelva a registrar los clientes.");
      }
      const combosRegistrados = new Set(aConfirmar.map(r => r.combo));

      // 1. CSV primero: si no se puede guardar, no queda nada registrado.
      const csvContent = construirCsvSap_(pedidosCsv);
      const fechaProceso = Utilities.formatDate(new Date(), "GMT-5", "yyyyMMdd");
      const fileName = nombreArchivoSap_(`Pedido_CEN_${fechaProceso}.csv`);
      let archivoCsv = null;
      try {
        archivoCsv = DriveApp.getFolderById(CONFIG.FOLDER_CSV_ID).createFile(fileName, csvContent, MimeType.PLAIN_TEXT);
      } catch (eCsv) {
        cen_borrarProvisionales_(sh, paquete.sessionId);
        throw new Error("No se pudo guardar el CSV en Drive; las órdenes provisionales se borraron de CONSOLIDADO: " + eCsv.message);
      }

      // 2. Confirmar: se quita la marca de las filas de esta sesión y se borran las que sobran.
      try {
        aConfirmar.forEach(r => sh.getRange(r.fila, CEN_COL_MARCA).setValue(""));
        SpreadsheetApp.flush();
        cen_borrarProvisionales_(sh, paquete.sessionId);
      } catch (errSheets) {
        try { archivoCsv.setTrashed(true); } catch (ignore) {}
        cen_borrarProvisionales_(sh, paquete.sessionId);
        throw new Error("Fallo al confirmar en CONSOLIDADO; el CSV se envió a la papelera y las provisionales se borraron: " + errSheets.message);
      }

      // 3. Traslado de los Excel. Un Excel CEN trae OCs de varios clientes: solo se mueve si TODAS
      //    sus OCs quedaron registradas y no tiene compradores sin homologar.
      const archivosRetenidos = [];
      if (configCli.PROCESSED_FOLDER_ID) {
        const folderDestino = DriveApp.getFolderById(configCli.PROCESSED_FOLDER_ID);
        const estadoArchivos = paquete.estadoArchivos || {};
        const archivosProcesadosMap = {};

        paquete.pedidos.forEach(p => {
          if (p.Archivo_Id && !archivosProcesadosMap[p.Archivo_Id]) {
            archivosProcesadosMap[p.Archivo_Id] = {
              nombreOriginal: p.Archivo_Nombre || p.Archivo_Id,
              clienteNombre: p.Cliente_Nombre || "CLIENTE_CEN",
              fechaPedido: p.Fecha_Pedido || Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")
            };
          }
        });

        for (const fileId in archivosProcesadosMap) {
          const estado = estadoArchivos[fileId];
          const completo = estado && !estado.bloqueado && estado.ocsPendientes.every(k => combosRegistrados.has(k));
          if (!completo) {
            archivosRetenidos.push(archivosProcesadosMap[fileId].nombreOriginal);
            continue;
          }
          try {
            const file = DriveApp.getFileById(fileId);
            file.moveTo(folderDestino);

            const info = archivosProcesadosMap[fileId];
            const nameClean = info.clienteNombre.replace(/[^a-zA-Z0-9_ -]/g, "").trim();
            const fechaClean = info.fechaPedido.replace(/[\/]/g, "-");
            const originalName = file.getName();
            const extIdx = originalName.lastIndexOf(".");
            const ext = extIdx !== -1 ? originalName.substring(extIdx) : ".xlsx";
            file.setName(nombreArchivoSap_(`${nameClean}_CEN_${fechaClean}${ext}`));
          } catch (eMove) {
            console.error("Error trasladando archivo CEN procesado: " + eMove.message);
          }
        }
      }

      // 4. Correos a las KAM, solo por las órdenes confirmadas
      if (paquete.notificarKAM !== false) {
        notificaciones.forEach(n => {
          const pedidosRegistradosCli = n.pedidosCli.filter(p => combosRegistrados.has(p.Combo_Key));
          if (pedidosRegistradosCli.length === 0) return;
          if (n.kamInfo.email && n.kamInfo.email.includes("@")) {
            try {
              cen_enviarCorreoKam_(pedidosRegistradosCli, paquete, n.kamInfo, n.clienteNombre);
            } catch (eMailErr) {
              console.error(`Error enviando correo a ${n.kamInfo.email}: ${eMailErr.message}`);
            }
          } else {
            console.warn(`No se envió notificación de ${n.clienteNombre}: no hay correo para el Solicitante SAP ${n.solicitanteSap} en Asignacion_KAM.`);
          }
        });
      }

      // Mensaje que verá el asistente.
      let mensaje = `Se confirmaron ${aConfirmar.length} órdenes de compra de CEN.`;
      if (archivosRetenidos.length > 0) {
        mensaje += ` ${archivosRetenidos.length} archivo(s) se dejaron en la carpeta de entrada porque aún tienen órdenes sin registrar: ${archivosRetenidos.join(", ")}.`;
      }

      return {
        success: true,
        mensaje: mensaje,
        csvData: Utilities.base64Encode(csvContent, Utilities.Charset.UTF_8).replace(/[\r\n]/g, ""),
        fileName: fileName
      };
    });
  } catch (err) {
    throw new Error(err.message);
  }
}

/**
 * Convierte un Excel (.xls, .xlsx o .xlsm) en una hoja de Google Sheets temporal, guardada en la misma carpeta
 * del Excel, para poder leerla. La conversión usa el servicio avanzado de Drive (Drive.Files.insert), que
 * debe estar activado en el proyecto.
 *
 * ESCRIBE: crea un archivo temporal en Drive con nombre "<nombre del Excel>_TEMP". Quien lo llama es
 * responsable de borrarlo.
 *
 * @param {DriveApp.File} file Excel de CEN a convertir.
 * @return {String} Id del archivo temporal creado.
 * @throws Si Drive no puede convertir el archivo, o si el servicio avanzado de Drive no está activo.
 */
function cen_convertirExcelASheet_(file) {
  const resource = {
    title: file.getName().replace(/\.[^/.]+$/, "") + "_TEMP",
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: file.getParents().next().getId() }]
  };
  const tempFile = Drive.Files.insert(resource, file.getBlob());
  return tempFile.id;
}

/**
 * Limpia un código EAN (código de barras del producto o de la empresa). Deja solo dígitos, con máximo 13.
 * Quita espacios, la parte decimal y la notación científica (4.5E+12 se convierte en número entero).
 * Un código de más de 13 dígitos se recorta a sus primeros 13.
 * No escribe ni lanza errores.
 *
 * @param {*} val Valor de la celda.
 * @return {String} EAN limpio, o "" si estaba vacío.
 */
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

/**
 * Convierte una fecha a texto dd/MM/yyyy (formato de CONSOLIDADO). Si el valor es una fecha, la formatea con la
 * zona horaria del script. Si es texto, solo quita los espacios de los extremos. No escribe ni lanza errores.
 *
 * @param {*} val Fecha (objeto Date) o texto.
 * @return {String} Fecha en texto.
 */
function cen_formatFechaStr_(val) {
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  return String(val).trim();
}

/**
 * Prepara un texto para compararlo: quita tildes, espacios, puntos, guiones y otros símbolos, y lo pasa a
 * mayúsculas. Por ejemplo "Razón Social" queda "RAZONSOCIAL". Así "SAP ID", "sap-id" y "Sap Id" se
 * consideran iguales al buscar columnas o clientes. No escribe ni lanza errores.
 *
 * @param {*} text Texto a preparar.
 * @return {String} Texto normalizado, o "" si estaba vacío.
 */
function cen_normalizeKey_(text) {
  if (!text) return "";
  return String(text).normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-zA-Z0-9]/g, "").toUpperCase().trim();
}

/**
 * Busca en la columna A de CONSOLIDADO el consecutivo más alto que ya usa un prefijo (por ejemplo "MC-26").
 * El siguiente número se calcula sumando uno a ese máximo. Revisa las filas desde la 2. Los IDs que no
 * empiezan por el prefijo se ignoran. No escribe ni lanza errores.
 *
 * @param {SpreadsheetApp.Sheet} shConsolidado Hoja CONSOLIDADO.
 * @param {String} prefix Prefijo, por ejemplo "MC-26".
 * @return {Number} Consecutivo más alto encontrado; 0 si no hay ninguno.
 */
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
 * Consulta el KAM (ejecutivo de cuenta) asignado a un cliente. Usa tres niveles para que el correo casi nunca
 * quede vacío:
 *  - Capa 1: obtenerInfoKamPorSolicitante_ (main.js), por Solicitante SAP.
 *  - Capa 2: búsqueda directa en la hoja Asignacion_KAM, por Solicitante SAP o, si no coincide, por nombre
 *    del cliente (coincidencia parcial).
 *  - Capa 3: valores por defecto (KAM ISDIN, iniciales MC, correo vacío). Con correo vacío no se envía
 *    el correo al KAM.
 * No escribe. Los fallos de las capas 1 y 2 se registran en el log (console.warn) y se pasa a la siguiente capa.
 *
 * @param {String|Number} solicitanteSap ID de SAP del cliente (Solicitante).
 * @param {String} [clienteNombre] Nombre del cliente, para buscar por nombre si el SAP no coincide.
 * @return {Object} { name, initials, email }: nombre del KAM, iniciales (para el consecutivo) y correo (puede ser "").
 */
function cen_getKamInfo_(solicitanteSap, clienteNombre) {
  // Capa 1: Consumo de main.js
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

/**
 * Arma el correo HTML con el resumen de las órdenes confirmadas de un cliente y lo envía a la KAM mediante
 * enviarCorreoKam_ (main.js). Por cada OC muestra:
 *  - productos OK y en back order, con su Material SAP, cantidad y subtotal;
 *  - productos sin Material válido ("EANs no actualizados"), en un bloque de alerta;
 *  - valor estimado OK y valor en back order.
 * Asunto: "Resumen Pedidos CEN (<cliente>) - <fecha>".
 *
 * ESCRIBE: envía un correo (MailApp, vía enviarCorreoKam_). En modo PRUEBA el destino lo decide
 * enviarCorreoKam_. No escribe en hojas.
 *
 * @param {Array<Object>} pedidosCli Pedidos del cliente que sí quedaron registrados.
 * @param {Object} paquete Paquete completo del asistente (items, materialesSeleccionados,
 *   noActualizadosSeleccionados, backordersSeleccionados).
 * @param {Object} kamInfo { name, initials, email } devuelto por cen_getKamInfo_.
 * @param {String} clienteNombre Nombre del cliente que aparece en el correo.
 * @return {void}
 * @throws Si MailApp no puede enviar el correo. guardarDefinitivoCEN captura el error y lo registra en el log.
 */
function cen_enviarCorreoKam_(pedidosCli, paquete, kamInfo, clienteNombre) {
  // 1. Encabezado del correo.
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

  // 2. Un bloque por OC: separar posiciones OK, back order y sin material, y calcular valores.
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

  // 3. Cerrar el HTML.
  htmlBody += `
      </div>
      <p style="font-size: 13px; color: #9CA3AF; margin-top: 25px; border-top: 1px solid #F3F4F6; padding-top: 15px; text-align: center;">Este es un correo automático generado por el Order Hub de ISDIN Colombia. Por favor no responder a este mensaje.</p>
    </div>
  `;

  // 4. Enviar el correo a la KAM.
  enviarCorreoKam_({
    to: kamInfo.email,
    subject: `📦 Resumen Pedidos CEN (${clienteNombre}) - ${Utilities.formatDate(new Date(), "GMT-5", "dd/MM/yyyy")}`,
    htmlBody: htmlBody
  });
}