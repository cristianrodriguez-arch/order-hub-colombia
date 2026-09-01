/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: INTEGRACIÓN Y PROCESAMIENTO DE ÓRDENES DE COMPRA - MEDIPIEL S.A.S. (PDF)
 * ARCHIVO: medipiel_logica.gs
 * ESTADO: VERSION_CONTROL_MEDIPIEL_V15_FULLY_VALORIZED_PROD
 * =========================================================================================
 * 📝 RESUMEN EJECUTIVO Y FUNCIONALIDAD:
 * Este módulo gestiona de forma desacoplada la extracción de texto mediante OCR/conversión
 * nativa de los pedidos en PDF de Medipiel S.A.S. en Google Drive.
 * 
 * 🛠️ CAPACIDADES CLAVE IMPLEMENTADAS:
 * 1. DESACOPLAMIENTO DE CONFIGURACIÓN: Consume IDs de libros, carpetas y parámetros SAP
 *    directamente desde el objeto CONFIG global definido en main.gs.
 * 2. GEOLOCALIZACIÓN NATIVA (COLUMNA F): Extrae la ciudad real de entrega directamente desde
 *    la Columna F ("Ciudad") de la pestaña "Destinatarios" para registrarla en el Consolidado.
 * 3. VALORIZACIÓN FINANCIERA TOTAL (KAM): Calcula y formatea los subtotales en $ COP
 *    tanto para referencias en estado OK como para ítems en Back Order en el correo notificador.
 * 4. CONSECUTIVO DINÁMICO SANANIZADO: Mantiene la secuencia por agente (ej. SC-260001) y
 *    deja limpia ("") la columna de KAM en el Consolidado.
 * =========================================================================================
 */

/**
 * FUNCIÓN 1: PRE-ANALIZAR (Punto de entrada llamado desde la UI para escanear PDFs)
 * Escanea la carpeta de entrada de Medipiel, valida órdenes existentes para evitar
 * duplicados y extrae la cabecera y tabla de productos.
 * 
 * @return {Object} Objeto con pedidosExtraidos e itemsExtraidos.
 */
function preAnalizarMEDIPIEL() {
  const cliente = 'MEDIPIEL';
  const configCli = CONFIG.CLIENTES[cliente];
  if (!configCli || !configCli.FOLDER_ID) {
    throw new Error('La configuración de MEDIPIEL no está definida en CONFIG de main.gs');
  }

  // 1. Conexión centralizada al libro de cálculo
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // 2. Cargar diccionarios de soporte
  const shBodega = ss.getSheetByName(CONFIG.SHEET_BODEGA || "Bodega");
  const diccBodega = shBodega ? obtenerMapaBodegaGeneral_(shBodega) : {};

  const shProductos = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  const diccProductos = shProductos ? medipiel_getProductosMap_(shProductos) : {};

  // Obtener diccionario de destinatarios con soporte para Columna F (Ciudad)
  const diccDestinatarios = obtenerDiccionarioDestinatarios();

  // 3. Antiduplicados: Cargar OCs procesadas previamente en el Consolidado
  let pedidosProcesados = new Set();
  const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  if (shConsolidado && shConsolidado.getLastRow() >= 3) {
    const mapConsolidado = medipiel_getHeaderMap_(shConsolidado, 3);
    const colOC = mapConsolidado[medipiel_normalizeKey_("Orden De Compra")];
    if (colOC) {
      const dataRange = shConsolidado.getRange(3, 1, shConsolidado.getLastRow() - 2, Math.max(shConsolidado.getLastColumn(), 26)).getValues();
      dataRange.forEach(r => {
        const oc = r[colOC - 1];
        if (oc) pedidosProcesados.add(String(oc).trim()); 
      });
    }
  }

  // 4. Buscar archivos PDF en la carpeta de entrada de Drive
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

  if (!archivos.length) {
    return { pedidosExtraidos: [], itemsExtraidos: [] };
  }

  // 5. Iniciar motor de lectura de PDFs
  return medipiel_motorLectura(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios);
}

/**
 * FUNCIÓN 2: GUARDAR DEFINITIVO (Invocada tras la confirmación en la UI)
 * Genera el consecutivo dinámico, inyecta los registros en CONSOLIDADO con la ciudad
 * real, compila el CSV para SAP, traslada los PDFs e informa a la KAM con valorización.
 * 
 * @param {Object} datosFormulario Payload enviado por la interfaz.
 * @return {Object} Respuesta con estado, datos en Base64 para el CSV y nombre del archivo.
 */
function guardarDefinitivoMEDIPIEL(datosFormulario) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const configCli = CONFIG.CLIENTES["MEDIPIEL"];

  // 1. Obtención de agente y correo electrónico de la KAM
  const shAsignacion = ss.getSheetByName("Asignacion_KAM");
  const shClientes = ss.getSheetByName("Listado clientes");
  const shAgentes = ss.getSheetByName("Iniciales de agente");
  
  let diccCiudades = {};
  let codigoAgente = "XX"; 
  let kamEncargado = "";
  let correoKam = ""; 

  if (shAsignacion) {
    const dataAsig = shAsignacion.getDataRange().getValues();
    const headAsig = dataAsig[0].map(h => medipiel_normalizeKey_(h));
    const cSap = headAsig.indexOf(medipiel_normalizeKey_("SAP ID"));
    const cKam = headAsig.findIndex(h => h.includes("KAMENCARGADO") || h === "KAM");
    if (cSap !== -1 && cKam !== -1) {
      for (let i = 1; i < dataAsig.length; i++) {
        if (String(dataAsig[i][cSap]).trim() === configCli.SOLICITANTE) {
          kamEncargado = String(dataAsig[i][cKam]).trim(); 
          break;
        }
      }
    }
  }

  if (shClientes) {
    const dataCli = shClientes.getDataRange().getValues();
    const headCli = dataCli[0].map(h => medipiel_normalizeKey_(h));
    const cSub = headCli.findIndex(h => h.includes("SUBSIDIARIO"));
    const cPob = headCli.findIndex(h => h.includes("POBLACION"));
    if (cSub !== -1 && cPob !== -1) {
      for (let i = 1; i < dataCli.length; i++) {
        const subsidiario = String(dataCli[i][cSub]).trim();
        if (subsidiario) diccCiudades[subsidiario] = String(dataCli[i][cPob]).trim();
      }
    }
  }
    
  if (shAgentes && kamEncargado) {
    const dataAg = shAgentes.getDataRange().getValues();
    const headAg = dataAg[0].map(h => medipiel_normalizeKey_(h));
    const cKamAg = headAg.findIndex(h => h.includes("KAMENCARGADO") || h === "KAM");
    const cCodAg = headAg.findIndex(h => h.includes("CODIGO"));
    const cCorreoAg = headAg.findIndex(h => h.includes("CORREO")); 
    
    if (cKamAg !== -1 && cCodAg !== -1) {
      for (let i = 1; i < dataAg.length; i++) {
        if (String(dataAg[i][cKamAg]).trim().toUpperCase() === kamEncargado.toUpperCase()) {
          codigoAgente = String(dataAg[i][cCodAg]).trim(); 
          let idxCorreo = cCorreoAg !== -1 ? cCorreoAg : 3;
          correoKam = String(dataAg[i][idxCorreo]).trim();
          break;
        }
      }
    }
  }

  // 2. Preparar la hoja CONSOLIDADO
  let shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  if (!shConsolidado) {
    shConsolidado = ss.insertSheet(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    shConsolidado.getRange(3, 1, 1, CONSOLIDADO_HEADERS.length).setValues([CONSOLIDADO_HEADERS]);
  } else if (shConsolidado.getLastRow() < 3) {
    shConsolidado.getRange(3, 1, 1, CONSOLIDADO_HEADERS.length).setValues([CONSOLIDADO_HEADERS]);
  }
  
  const headerMapConsolidado = medipiel_getHeaderMap_(shConsolidado, 3);
  const numColsConsolidado = Math.max(shConsolidado.getLastColumn() || 13, CONSOLIDADO_HEADERS.length);

  // 3. Generar consecutivo dinámico secuencial por agente (<CódigoAgente>-<Año><Consecutivo>)
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

  // 4. Mapeo de datos para CSV y CONSOLIDADO
  let filasCSV = [];
  let filasConsolidado = [];
  let contadorPedido = 1;
  const pedidosValidos = new Set(); 
  let primerConsecutivoGenerado = null; 
  const mapaNombresArchivos = {}; 

  datosFormulario.pedidos.forEach(p => {
    let posicionItem = 10;
    let sumaUnidadesPedido = 0;
    let sumaValorPedido = 0;

    let fechaUI = datosFormulario.fechasEntregas[p.Pedido_ID] ? datosFormulario.fechasEntregas[p.Pedido_ID].fecha : "";
    let fechaEntregaSAP = "";
    let fechaEntregaNormalObj = "";
    
    if (fechaUI) {
      let parts = fechaUI.split("-");
      if (parts.length === 3) {
        fechaEntregaSAP = `${parts[2]}.${parts[1]}.${parts[0]}`;
        fechaEntregaNormalObj = new Date(parts[0], parts[1] - 1, parts[2]);
      }
    }

    const items = datosFormulario.items.filter(i => i.Pedido_ID === p.Pedido_ID);

    items.forEach(i => {
      let materialFinal = (datosFormulario.materialesSeleccionados && datosFormulario.materialesSeleccionados[i.Item_ID]) 
                            ? datosFormulario.materialesSeleccionados[i.Item_ID] : "";

      if (materialFinal !== "" && !materialFinal.includes("SIN_")) {
        pedidosValidos.add(p.Pedido_ID);
        sumaUnidadesPedido += parseInt(i.Cant);
        sumaValorPedido += (Number(i.Valor_Unitario) || 0) * (Number(i.Cant) || 0);

        filasCSV.push([
          contadorPedido,
          "",
          configCli.SAP.CLASE_PEDIDO,
          configCli.SAP.ORG_VENTAS,
          configCli.SAP.CANAL,
          configCli.SAP.SECTOR,
          configCli.SOLICITANTE,
          p.Destinatario_Merc,
          p.Pedido_ID,
          fechaEntregaSAP,
          "", 
          "",
          posicionItem,
          materialFinal,
          "",
          i.Cant
        ]);
        
        posicionItem += 10;
      }
    });

    if (pedidosValidos.has(p.Pedido_ID)) {
      maxConsecutivo++;
      const consecutivoStr = String(maxConsecutivo).padStart(4, '0');
      const idGenerado = `${codigoAgente}-${currentYearSuffix}${consecutivoStr}`;
      
      // Tomar la ciudad proveniente de la Columna F de Destinatarios, con fallback a diccCiudades
      const ciudadDestino = p.Ciudad || diccCiudades[p.Destinatario_Merc] || "COTA";

      if (!primerConsecutivoGenerado) {
        primerConsecutivoGenerado = idGenerado;
      }

      if (p.Archivo_ID) {
        mapaNombresArchivos[p.Archivo_ID] = `Pedido_${idGenerado}.pdf`;
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
        "Valor": sumaValorPedido, 
        "NO ENTREGA FACTURACIÓN": "",
        "Observaciones": "",
        "KAM": "" // Dejado en blanco de forma explícita
      }, numColsConsolidado));
      
      contadorPedido++;
    }
  });

  if (filasCSV.length === 0) {
     throw new Error("No hay productos válidos seleccionados. Proceso abortado.");
  }

  // 5. Generar archivo CSV en memoria
  let csvFileName = `Pedido_${primerConsecutivoGenerado || Utilities.formatDate(new Date(), "GMT-5", "yyyyMMdd_HHmm")}.csv`;
  let csvString = "";
  
  filasCSV.unshift(medipiel_getEstructuraCSVHeaders_());
  csvString = filasCSV.map(row => row.map(val => {
    let s = String(val);
    return (s.includes(';') || s.includes('"')) ? `"${s.replace(/"/g, '""')}"` : s;
  }).join(';')).join('\r\n');

  // 6. Escribir en la hoja CONSOLIDADO
  try {
    if (filasConsolidado.length > 0) {
      shConsolidado.getRange(shConsolidado.getLastRow() + 1, 1, filasConsolidado.length, numColsConsolidado).setValues(filasConsolidado);
      SpreadsheetApp.flush(); 
    }
  } catch (errSheets) {
    throw new Error("Fallo al escribir en la Base de Datos (Consolidado): " + errSheets.message);
  }

  // 7. Guardar el archivo CSV en Google Drive
  try {
    if (CONFIG.FOLDER_CSV_ID) {
      const folderCSV = DriveApp.getFolderById(CONFIG.FOLDER_CSV_ID);
      folderCSV.createFile(csvFileName, csvString, MimeType.CSV);
    } else {
      throw new Error("Falta configurar FOLDER_CSV_ID en CONFIG.");
    }
  } catch (errCSV) { 
    throw new Error("Fallo al guardar archivo CSV en Drive: " + errCSV.message); 
  }

  // 8. Mover y renombrar archivos PDF procesados en Drive
  const folderDestinoId = configCli.PROCESSED_FOLDER_ID;
  const folderOrigenId = configCli.FOLDER_ID;
  const archivosAMover = new Set();
  
  datosFormulario.pedidos.forEach(p => { 
    if (pedidosValidos.has(p.Pedido_ID) && p.Archivo_ID) {
      archivosAMover.add(p.Archivo_ID); 
    }
  });

  archivosAMover.forEach(fileId => {
    const nuevoNombre = mapaNombresArchivos[fileId];
    try {
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

  // 9. ENVÍO DE CORREO AUTOMÁTICO VALORIZADO AL KAM
  if (correoKam && correoKam.includes("@")) {
    try {
      let correoHTML = `
        <div style="font-family: 'Inter', sans-serif; color: #2B2830; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
          <div style="text-align: center; border-bottom: 2px solid #E41F33; padding-bottom: 15px; margin-bottom: 20px;">
            <h1 style="color: #2B2830; font-size: 24px; margin: 0; font-weight: 900; letter-spacing: 2px;">ISDIN</h1>
            <p style="color: #E41F33; font-size: 11px; font-weight: 700; margin: 5px 0 0; text-transform: uppercase; letter-spacing: 1px;">Notificación Automática de Órdenes B2B</p>
          </div>
          
          <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Estimada <strong>${kamEncargado}</strong>,</p>
          <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Te informamos que se han consolidado de manera exitosa los nuevos pedidos de <strong>MEDIPIEL S.A.S.</strong>. El consecutivo asignado es el <span style="font-family: monospace; font-weight: bold; background-color: #F3F4F6; padding: 2px 6px; border-radius: 4px; border: 1px solid #E5E7EB;">Pedido_${primerConsecutivoGenerado}</span>.</p>
          
          <h2 style="font-size: 16px; font-weight: 700; border-bottom: 1px solid #F3F4F6; padding-bottom: 8px; margin-top: 25px; color: #2B2830;">Detalle de Facturación por Orden de Compra</h2>
          <div style="margin-top: 15px;">
      `;
      
      const resumenOC = {};
      datosFormulario.items.forEach(item => {
         let oc = item.Pedido_ID;
         if (!resumenOC[oc]) {
           resumenOC[oc] = {
             inyectados: 0,
             itemsProcesados: [],
             contieneBackOrder: false,
             totalValorOk: 0,
             totalValorBO: 0
           };
         }
         
         let mat = datosFormulario.materialesSeleccionados[item.Item_ID] || "";
         let inyectado = (mat !== "" && !mat.includes("SIN_"));
         let isBackOrder = datosFormulario.backordersSeleccionados[item.Item_ID] || false;
         
         if (inyectado) {
             resumenOC[oc].inyectados++;
             if (isBackOrder) resumenOC[oc].contieneBackOrder = true;
             
             const cant = parseInt(item.Cant) || 0;
             const price = parseFloat(item.Valor_Unitario) || 0;
             const sub = cant * price;
             
             if (isBackOrder) {
               resumenOC[oc].totalValorBO += sub;
             } else {
               resumenOC[oc].totalValorOk += sub;
             }
             
             let prodExistente = resumenOC[oc].itemsProcesados.find(p => p.cod === item.Producto_Cod);
             if (prodExistente) {
                 prodExistente.cant += cant;
                 prodExistente.subtotal += sub;
                 if (isBackOrder) prodExistente.isBackOrder = true; 
             } else {
                 resumenOC[oc].itemsProcesados.push({
                     desc: item.Descripcion, 
                     cant: cant, 
                     cod: item.Producto_Cod,
                     subtotal: sub,
                     isBackOrder: isBackOrder
                 });
             }
         }
      });

      for (let oc in resumenOC) {
         if (resumenOC[oc].inyectados === 0) continue;

         let statusOrden = resumenOC[oc].contieneBackOrder 
            ? `<span style="background-color: #FEF3C7; color: #D97706; border: 1px solid #FCD34D; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">⚠️ Contiene Back Orders</span>`
            : `<span style="background-color: #D1FAE5; color: #059669; border: 1px solid #6EE7B7; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">✅ Completada con Éxito</span>`;

         let valueSummaryText = `Valor Estimado Facturar (OK): $${Math.round(resumenOC[oc].totalValorOk).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
         if (resumenOC[oc].totalValorBO > 0) {
           valueSummaryText += ` | Valor en Back Order: $${Math.round(resumenOC[oc].totalValorBO).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
         }

         correoHTML += `
           <div style="background-color: #FAFAFA; border: 1px solid #F3F4F6; border-radius: 8px; padding: 15px; margin-bottom: 15px;">
             <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #E5E7EB; padding-bottom: 10px; margin-bottom: 10px;">
               <span style="font-weight: bold; font-size: 14px; color: #111827;">Orden: ${oc}</span>
               ${statusOrden}
             </div>
             <div style="margin-top: 10px;">
         `;
         
         resumenOC[oc].itemsProcesados.forEach(prod => {
             let badge = prod.isBackOrder
                ? '<span style="background-color: #FEF3C7; color: #D97706; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">Back Order</span>'
                : '<span style="background-color: #D1FAE5; color: #059669; padding: 2px 8px; border-radius: 9999px; font-size: 10px; font-weight: bold;">OK</span>';

             const formattedSub = "$" + Math.round(prod.subtotal).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
             // INYECCIÓN DE SUBTOTAL FINANCIERO EN CADA PRODUCTO
             const subtotalText = prod.subtotal > 0 ? ` | Subtotal: <strong>${formattedSub}</strong>` : "";

             correoHTML += `
               <div style="padding: 10px 0; border-bottom: 1px dashed #F3F4F6; display: flex; justify-content: space-between; font-size: 12px;">
                 <div style="flex: 1; min-width: 0; padding-right: 10px;">
                   <span style="font-weight: bold; color: #111827; display: block; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${prod.desc}</span>
                   <span style="color: #6B7280; font-family: monospace; font-size: 11px;">EAN: ${prod.cod} | ${prod.cant} ud.${subtotalText}</span>
                 </div>
                 <div style="text-align: right; white-space: nowrap;">
                   ${badge}
                 </div>
               </div>
             `;
         });
         
         correoHTML += `
             </div>
             <div style="text-align: right; font-size: 13px; font-weight: bold; margin-top: 15px; color: #111827;">
               ${valueSummaryText}
             </div>
           </div>
         `;
      }

      correoHTML += `
          </div>
          <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">El archivo CSV (<strong>${csvFileName}</strong>) ha sido guardado automáticamente.</p>
          <p style="font-size: 11px; color: #9CA3AF; margin-top: 25px; border-top: 1px solid #F3F4F6; padding-top: 15px; text-align: center;">Este es un correo automático generado por el Order Hub de ISDIN Colombia. Por favor no responder.</p>
        </div>
      `;

      MailApp.sendEmail({
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
    csvData: Utilities.base64Encode(csvString), 
    fileName: csvFileName
  };
}

function medipiel_motorLectura(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios) {
  const pedidosExtraidos = [];
  const itemsExtraidos = [];
  const errores = [];

  archivos.forEach((archivo) => {
    const fileId = archivo.id;
    const fileName = archivo.title || archivo.name || 'archivo.pdf';

    try {
      const blob = DriveApp.getFileById(fileId).getBlob();
      const texto = medipiel_pdfToTextPreferNative_(blob, fileName);
      const pedido = medipiel_parsePedidoCabecera_(texto, fileId, fileName, configCli, diccDestinatarios);
      
      const items = medipiel_parseItems_(texto, pedido.Pedido_ID, diccBodega, diccProductos);
      if (!items.length) throw new Error('No se detectaron productos.');

      pedidosExtraidos.push(pedido);
      items.forEach(it => itemsExtraidos.push(it));

    } catch (err) {
      if (err.message.includes('MISSING_BODEGA|')) throw err; 
      if (err.message.includes('FATAL_BODEGA:')) throw new Error(err.message.replace('FATAL_BODEGA:', '')); 
      errores.push(fileName + ": " + err.message);
    }
  });

  return { pedidosExtraidos, itemsExtraidos, errores };
}

function medipiel_pdfToTextPreferNative_(blob, fileName) {
  const doc = Drive.Files.insert({ title: 'TEMP_MEDIPIEL', mimeType: 'application/vnd.google-apps.document' }, blob, { convert: true });
  const text = DocumentApp.openById(doc.id).getBody().getText();
  Drive.Files.remove(doc.id);
  return String(text || '').replace(/\u00A0/g, ' ').replace(/\r/g, '\n').replace(/[ ]{2,}/g, ' ').trim();
}

function medipiel_parsePedidoCabecera_(texto, fileId, fileName, configCli, diccDestinatarios) {
  const t = texto;
  const ordenCompra = medipiel_matchRegex_(t, /No\.\s*(\d{6,})/i);
  const fechaPedido = medipiel_matchRegex_(t, /Fecha\s*Orden\s*:\s*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i);
  const fechaEntrega = medipiel_matchRegex_(t, /Fecha\s*Entrega\s*:\s*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i);
  const idDestinoRaw = medipiel_matchRegex_(t, /Bodega\s*:\s*([A-Z0-9_-]+)/i);

  if (!ordenCompra) throw new Error('OC no encontrada.');
  if (!idDestinoRaw) throw new Error(`FATAL_BODEGA:No se pudo leer la Bodega en la OC ${ordenCompra}.`);

  const destinoMap = diccDestinatarios[idDestinoRaw];
  
  if (!destinoMap) {
    throw new Error(`MISSING_BODEGA|${idDestinoRaw}|${configCli.SOLICITANTE}|${configCli.NOMBRE_CLIENTE}`);
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
    Ciudad: destinoMap.ciudad || 'COTA' // Extraída nativamente de la Columna F
  };
}

function medipiel_parseItems_(texto, pedidoId, diccBodega, diccProductos) {
  const m = texto.match(/C[ÓO]DIGO\s+BARRAS[\s\S]*?COSTO\s+UN([\s\S]*?)(?:Valor\s+Subtotal\s*:|Valor\s+Neto\s*:|$)/i);
  const bloqueTabla = m && m[1] ? m[1].trim() : '';
  if (!bloqueTabla) return [];

  const items = [];
  const re = /(\d{8,14})\s+(\d{3,})\s+([\s\S]+?)\s*(?:Und|UND|Un|UN|und|U\.M\.)\s+(\d+(?:[.,]\d+)?)\s+([$\d.,]+)/gmi;
  let match;

  while ((match = re.exec(bloqueTabla)) !== null) {
    const eanRaw = String(match[1]).trim();
    let mats = eanRaw ? (diccBodega[eanRaw] || []) : [];
    
    if (mats.length === 0 && eanRaw && diccProductos[eanRaw]) {
        mats = [{ material: diccProductos[eanRaw], stock: 999999 }];
    }
    
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

function medipiel_getEstructuraCSVHeaders_() { 
  return [ "Contador", "N Pedido SAP", "Clase pedido", "Org ventas", "Canal", "Sector", "Solicitante", "Destinatario Merc", "Num pedido cliente", "Fecha preferente Entrega", "Descuento ZTK1 (%)", "Motivo pedido", "Posicion pedido", "Material", "EAN", "Cantidad" ]; 
}

function medipiel_normalizeKey_(str) {
  if (!str) return "";
  return String(str)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

function medipiel_getHeaderMap_(s, rowNum) {
  if (s.getLastRow() < rowNum) return {};
  const numCols = Math.max(s.getLastColumn(), 26);
  const h = s.getRange(rowNum, 1, 1, numCols).getValues()[0];
  const m = {};
  h.forEach((x, i) => { 
    if (x) m[medipiel_normalizeKey_(x)] = i + 1; 
  });
  return m;
}

function medipiel_createRowArray_(map, dataObj, numCols) {
  const row = new Array(numCols).fill("");
  for (const rawKey in dataObj) {
    const normKey = medipiel_normalizeKey_(rawKey);
    if (map[normKey]) {
      row[map[normKey] - 1] = dataObj[rawKey];
    }
  }
  return row;
}

function medipiel_getProductosMap_(sheet) {
  const map = {}; 
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return map;
  const h = data[0].map(x => medipiel_normalizeKey_(x));
  
  let cE = h.indexOf(medipiel_normalizeKey_("EAN"));
  let cM = h.indexOf(medipiel_normalizeKey_("MATERIAL ID"));
  
  if (cE === -1) cE = 4; 
  if (cM === -1) cM = 1; 
  
  for (let i = 1; i < data.length; i++) {
    const ean = data[i][cE] ? String(data[i][cE]).trim() : "";
    if (ean) map[ean] = String(data[i][cM]).trim();
  }
  return map;
}

function medipiel_matchRegex_(text, regex) { const m = text.match(regex); return m ? m[1].trim() : null; }
function medipiel_parseNumber_(v) { return Number(String(v).trim().replace(/\./g, '').replace(',', '.')) || 0; }
function medipiel_parseDateEs_(d) {
  if (!d) return "";
  const p = d.split(/[-/]/);
  if (p.length === 3) {
    let y = parseInt(p[2], 10);
    if (y < 100) y += 2000;
    return new Date(y, p[1] - 1, p[0]);
  }
  return d;
}

function medipiel_guardarNuevoDestinatario(datos) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh = ss.getSheetByName(CONFIG.SHEET_DESTINATARIOS || "Destinatarios");
  
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_DESTINATARIOS || "Destinatarios");
    sh.appendRow(["Id_origen", "Nombre_origen", "Id_destino", "Nombre_destino", "Cruce_cliente", "Ciudad"]);
  }
  
  sh.appendRow([datos.idOrigen, datos.nombreOrigen, datos.idDestino, datos.nombreDestino, datos.cruceCliente, datos.ciudad || "COTA"]);
  return true;
}