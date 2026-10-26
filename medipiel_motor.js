/**
 * =========================================================================================
 * PROYECTO: ORDER HUB - ISDIN COLOMBIA
 * MÓDULO: INTEGRACIÓN Y PROCESAMIENTO DE ÓRDENES DE COMPRA - MEDIPIEL S.A.S. (PDF)
 * ARCHIVO: medipiel_logica.gs
 * ESTADO: PRODUCCIÓN - SOPORTE MATERIALES SAP 12 DÍGITOS, KAM DESACOPLADO Y EAN 13 DÍGITOS
 * =========================================================================================
 * 📝 REGLAS Y CORRECCIONES APLICADAS:
 * 1. COMPATIBILIDAD CON MATERIALES SAP DE 12 DÍGITOS: Se eliminó la condición `length > 10`
 *    tanto en la generación del CSV como en el correo a la KAM, permitiendo códigos oficiales
 *    de hasta 12 dígitos (ej: 690032151480 / Ureadin Cream10).
 * 2. EAN ESTRICTO (13 DÍGITOS): En `medipiel_parseItems_` se mantiene estrictamente la
 *    captura de códigos de barras de 13 dígitos (\d{13}).
 * 3. KAM 100% DESACOPLADO: Invoca `obtenerInfoKamPorSolicitante_()` de main.gs sin datos fijos.
 * 4. ESCUDO ANTI-DUPLICADOS COMPUESTO: Evalúa la combinación única `${OC}_${SAPID}`.
 * 5. FÓRMULA DE KAM PRESERVADA: En CONSOLIDADO, la columna KAM se deja vacía ("") para
 *    respetar la fórmula nativa de Google Sheets.
 * =========================================================================================
 */

function preAnalizarMEDIPIEL() {
  const cliente = 'MEDIPIEL';
  const configCli = CONFIG.CLIENTES[cliente];
  if (!configCli || !configCli.FOLDER_ID) {
    throw new Error('La configuración de MEDIPIEL no está definida en CONFIG de main.gs');
  }

  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  const shBodega = ss.getSheetByName(CONFIG.SHEET_BODEGA || "Bodega");
  const diccBodega = shBodega ? obtenerMapaBodegaGeneral_(shBodega) : {};

  const shProductos = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS || "Productos");
  const diccProductos = shProductos ? medipiel_getProductosMap_(shProductos) : {};

  const diccDestinatarios = obtenerDiccionarioDestinatarios();

  // Antiduplicados: Cargar OCs procesadas previamente en el Consolidado con clave compuesta
  let pedidosProcesados = new Set();
  const shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  
  if (shConsolidado && shConsolidado.getLastRow() >= 3) {
    const mapConsolidado = medipiel_getHeaderMap_(shConsolidado, 3);
    const colOC = mapConsolidado[medipiel_normalizeKey_("Orden De Compra")];
    const colSap = mapConsolidado[medipiel_normalizeKey_("Sap Id")];
    
    if (colOC) {
      const dataRange = shConsolidado.getRange(3, 1, shConsolidado.getLastRow() - 2, Math.max(shConsolidado.getLastColumn(), 26)).getValues();
      dataRange.forEach(r => {
        const oc = medipiel_limpiarOc_(r[colOC - 1]);
        const sap = colSap ? medipiel_normalizeKey_(r[colSap - 1]) : medipiel_normalizeKey_(configCli.SOLICITANTE);
        if (oc && sap) {
          pedidosProcesados.add(`${oc}_${sap}`);
        }
      });
    }
  }

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
    return { pedidosExtraidos: [], itemsExtraidos: [], yaProcesados: [] };
  }

  return medipiel_motorLectura(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios);
}

function guardarDefinitivoMEDIPIEL(datosFormulario) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const configCli = CONFIG.CLIENTES["MEDIPIEL"];

  // 1. Obtención 100% DESACOPLADA de la información del KAM desde main.gs
  const infoKam = typeof obtenerInfoKamPorSolicitante_ === "function" 
    ? obtenerInfoKamPorSolicitante_(configCli.SOLICITANTE) 
    : null;

  let codigoAgente = (infoKam && infoKam.iniciales) ? infoKam.iniciales : "SC";
  let kamEncargado = (infoKam && infoKam.nombre) ? infoKam.nombre : "KAM ENCARGADO";
  let correoKam = (infoKam && infoKam.correo) ? infoKam.correo : "";

  // 2. Mapeo complementario de ciudades por subsidiario
  const shClientes = ss.getSheetByName("Listado clientes");
  let diccCiudades = {};

  if (shClientes && shClientes.getLastRow() >= 2) {
    const dataCli = shClientes.getDataRange().getValues();
    const headCli = dataCli[0].map(h => medipiel_normalizeKey_(h));
    const cSub = headCli.findIndex(h => h.includes("SUBSIDIARIO"));
    const cPob = headCli.findIndex(h => h.includes("POBLACION") || h.includes("CIUDAD"));
    if (cSub !== -1 && cPob !== -1) {
      for (let i = 1; i < dataCli.length; i++) {
        const subsidiario = String(dataCli[i][cSub]).trim();
        if (subsidiario) diccCiudades[subsidiario] = String(dataCli[i][cPob]).trim();
      }
    }
  }

  // 3. Preparar la hoja CONSOLIDADO
  let shConsolidado = ss.getSheetByName(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
  if (!shConsolidado) {
    shConsolidado = ss.insertSheet(CONFIG.SHEET_PEDIDOS || "CONSOLIDADO");
    shConsolidado.getRange(3, 1, 1, CONSOLIDADO_HEADERS.length).setValues([CONSOLIDADO_HEADERS]);
  } else if (shConsolidado.getLastRow() < 3) {
    shConsolidado.getRange(3, 1, 1, CONSOLIDADO_HEADERS.length).setValues([CONSOLIDADO_HEADERS]);
  }
  
  const headerMapConsolidado = medipiel_getHeaderMap_(shConsolidado, 3);
  const numColsConsolidado = Math.max(shConsolidado.getLastColumn() || 13, CONSOLIDADO_HEADERS.length);

  // 4. Consecutivo dinámico secuencial por agente (<CódigoAgente>-<Año><Consecutivo>)
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

  // 5. Mapeo de datos para CSV y CONSOLIDADO
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
                            ? String(datosFormulario.materialesSeleccionados[i.Item_ID]).trim() : "";
      
      // CORRECCIÓN: Permite códigos de hasta 12 dígitos oficiales (descarta vacíos, SIN_ o EANs de 13 dígitos que empiezan por 84)
      const isNoActualizado = (datosFormulario.noActualizadosSeleccionados && datosFormulario.noActualizadosSeleccionados[i.Item_ID]) || 
                              (!materialFinal || materialFinal.includes("SIN_") || materialFinal === i.Producto_Cod || (materialFinal.length === 13 && materialFinal.startsWith("84")));

      let matCsv = isNoActualizado ? "" : materialFinal;

      pedidosValidos.add(p.Pedido_ID);
      sumaUnidadesPedido += parseInt(i.Cant) || 0;
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
        matCsv, // Lleva el Material SAP (incluyendo códigos de 12 dígitos) o vacío si no cruzó
        "",
        i.Cant
      ]);
      
      posicionItem += 10;
    });

    if (pedidosValidos.has(p.Pedido_ID)) {
      maxConsecutivo++;
      const consecutivoStr = String(maxConsecutivo).padStart(4, '0');
      const idGenerado = `${codigoAgente}-${currentYearSuffix}${consecutivoStr}`;
      
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
        "KAM": "" // Se deja en blanco para preservar la fórmula nativa de Google Sheets
      }, numColsConsolidado));
      
      contadorPedido++;
    }
  });

  if (filasCSV.length === 0) {
     throw new Error("No hay productos válidos seleccionados. Proceso abortado.");
  }

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

  // 8. Traslado y renombrado de PDFs procesados en Google Drive
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

  // 9. Enviar notificación por correo a la KAM encargada
  if (correoKam && correoKam.includes("@")) {
    try {
      let correoHTML = `
        <div style="font-family: 'Inter', sans-serif; color: #2B2830; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #E5E7EB; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
          <div style="text-align: center; border-bottom: 2px solid #E41F33; padding-bottom: 15px; margin-bottom: 20px;">
            <h1 style="color: #2B2830; font-size: 24px; margin: 0; font-weight: 900; letter-spacing: 2px;">ISDIN</h1>
            <p style="color: #E41F33; font-size: 11px; font-weight: 700; margin: 5px 0 0; text-transform: uppercase; letter-spacing: 1px;">Notificación Automática de Órdenes B2B - MEDIPIEL</p>
          </div>
          
          <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Estimada <strong>${kamEncargado}</strong>,</p>
          <p style="font-size: 14px; line-height: 1.6; color: #4B5563;">Te informamos que se han consolidado de manera exitosa los nuevos pedidos de <strong>MEDIPIEL S.A.S.</strong>. El consecutivo asignado es el <span style="font-family: monospace; font-weight: bold; background-color: #F3F4F6; padding: 2px 6px; border-radius: 4px; border: 1px solid #E5E7EB;">Pedido_${primerConsecutivoGenerado}</span>.</p>
          
          <h2 style="font-size: 16px; font-weight: 700; border-bottom: 1px solid #F3F4F6; padding-bottom: 8px; margin-top: 25px; color: #2B2830;">Detalle de Facturación por Orden de Compra</h2>
          <div style="margin-top: 15px;">
      `;
      
      datosFormulario.pedidos.forEach(p => {
        const itemsPedido = datosFormulario.items.filter(i => i.Pedido_ID === p.Pedido_ID);
        let itemsNoActualizados = [];
        let itemsValidos = [];
        let totalValorOk = 0;
        let totalValorBO = 0;
        let hasBackOrder = false;

        itemsPedido.forEach(item => {
          let mat = (datosFormulario.materialesSeleccionados && datosFormulario.materialesSeleccionados[item.Item_ID]) ? String(datosFormulario.materialesSeleccionados[item.Item_ID]).trim() : "";
          
          // CORRECCIÓN: Permite materiales de 12 dígitos sin enviarlos erróneamente a No Actualizados
          let isNoActualizado = (datosFormulario.noActualizadosSeleccionados && datosFormulario.noActualizadosSeleccionados[item.Item_ID]) || 
                                (!mat || mat.includes("SIN_") || mat === item.Producto_Cod || (mat.length === 13 && mat.startsWith("84")));
          let isBO = !isNoActualizado && (datosFormulario.backordersSeleccionados && datosFormulario.backordersSeleccionados[item.Item_ID]);

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
          statusLabel = '<span style="background-color: #D1FAE5; color: #059669; border: 1px solid #6EE7B7; padding: 3px 10px; border-radius: 9999px; font-size: 11px; font-weight: bold;">✅ Completada con Éxito</span>';
        }

        let valueSummaryText = `Valor Estimado Facturar (OK): $${Math.round(totalValorOk).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
        if (totalValorBO > 0) {
          valueSummaryText += ` | Valor en Back Order: $${Math.round(totalValorBO).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".")}`;
        }
        if (itemsNoActualizados.length > 0) {
          valueSummaryText += ` <br><span style="color:#B91C1C; font-size:11px; margin-top:4px; display:block;"><strong>Alerta Comercial:</strong> Hay ${itemsNoActualizados.length} ítem(s) que requieren actualización de catálogo.</span>`;
        }

        correoHTML += `
          <div style="background-color: #FAFAFA; border: 1px solid #F3F4F6; border-radius: 8px; padding: 15px; margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #E5E7EB; padding-bottom: 10px; margin-bottom: 10px;">
              <span style="font-weight: bold; font-size: 14px; color: #111827;">Orden: ${p.Pedido_ID}</span>
              ${statusLabel}
            </div>
            ${alertaNoActualizadosHtml}
            ${listHtml}
            <div style="text-align: right; font-size: 13px; font-weight: bold; margin-top: 15px; color: #111827;">
              ${valueSummaryText}
            </div>
          </div>
        `;
      });

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
    csvData: Utilities.base64Encode(csvString).replace(/[\r\n]/g, ""), 
    fileName: csvFileName
  };
}

function medipiel_motorLectura(archivos, diccBodega, diccProductos, configCli, pedidosProcesados, diccDestinatarios) {
  const pedidosExtraidos = [];
  const itemsExtraidos = [];
  const errores = [];
  const yaProcesados = [];
  const targetSap = medipiel_normalizeKey_(configCli.SOLICITANTE || "11026712");

  archivos.forEach((archivo) => {
    const fileId = archivo.id;
    const fileName = archivo.title || archivo.name || 'archivo.pdf';

    try {
      const blob = DriveApp.getFileById(fileId).getBlob();
      const texto = medipiel_pdfToTextPreferNative_(blob, fileName);
      const pedido = medipiel_parsePedidoCabecera_(texto, fileId, fileName, configCli, diccDestinatarios);
      
      const ocTrim = medipiel_limpiarOc_(pedido.Pedido_ID);
      const comboKey = `${ocTrim}_${targetSap}`;

      // ESCUDO ANTI-DUPLICADOS: Si la OC ya existe para MEDIPIEL, se omite
      if (pedidosProcesados.has(comboKey)) {
        if (!yaProcesados.includes(ocTrim)) yaProcesados.push(ocTrim);
        return;
      }

      const items = medipiel_parseItems_(texto, pedido.Pedido_ID, diccBodega, diccProductos);
      if (!items.length) throw new Error('No se detectaron productos en la tabla del PDF.');

      pedidosExtraidos.push(pedido);
      items.forEach(it => itemsExtraidos.push(it));

    } catch (err) {
      if (err.message.includes('MISSING_BODEGA|')) throw err; 
      if (err.message.includes('FATAL_BODEGA:')) throw new Error(err.message.replace('FATAL_BODEGA:', '')); 
      errores.push(fileName + ": " + err.message);
    }
  });

  return { pedidosExtraidos, itemsExtraidos, errores, yaProcesados };
}

function medipiel_pdfToTextPreferNative_(blob, fileName) {
  const doc = Drive.Files.insert({ title: 'TEMP_MEDIPIEL', mimeType: 'application/vnd.google-apps.document' }, blob, { convert: true });
  const text = DocumentApp.openById(doc.id).getBody().getText();
  Drive.Files.remove(doc.id);
  return String(text || '').replace(/\u00A0/g, ' ').replace(/\r/g, '\n').replace(/[ ]{2,}/g, ' ').trim();
}

function medipiel_parsePedidoCabecera_(texto, fileId, fileName, configCli, diccDestinatarios) {
  const t = texto;
  
  let ordenCompra = medipiel_matchRegex_(t, /(?:No\.?|N°|Nro\.?|Orden(?:\s+de\s+Compra)?(?:\s+No\.?)?|Pedido(?:\s+No\.?)?)\s*[:.]?\s*(\d{4,12})/i);
  
  if (!ordenCompra && fileName) {
    const matchFile = fileName.match(/\d{4,12}/);
    if (matchFile) ordenCompra = matchFile[0];
  }

  if (!ordenCompra) {
    ordenCompra = medipiel_matchRegex_(t, /No\s*(\d{4,12})/i);
  }

  const fechaPedido = medipiel_matchRegex_(t, /Fecha\s*(?:de\s*)?(?:Orden|Pedido)?\s*:\s*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i);
  const fechaEntrega = medipiel_matchRegex_(t, /Fecha\s*(?:de\s*)?Entrega\s*:\s*([0-9]{1,2}[\/-][0-9]{1,2}[\/-][0-9]{2,4})/i);
  
  let idDestinoRaw = medipiel_matchRegex_(t, /Bodega\s*(?:de\s*Entrega)?\s*:\s*([A-Z0-9_-]+)/i);
  if (!idDestinoRaw) {
    idDestinoRaw = medipiel_matchRegex_(t, /Bodega\s*:\s*([^\n\r,]+)/i);
    if (idDestinoRaw) idDestinoRaw = idDestinoRaw.trim();
  }

  if (!ordenCompra) throw new Error(`OC no encontrada en el PDF ni en el nombre del archivo.`);
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
    Ciudad: destinoMap.ciudad || 'COTA'
  };
}

function medipiel_parseItems_(texto, pedidoId, diccBodega, diccProductos) {
  const m = texto.match(/C[ÓO]DIGO\s+BARRAS[\s\S]*?COSTO\s+UN([\s\S]*?)(?:Valor\s+Subtotal\s*:|Valor\s+Neto\s*:|$)/i);
  const bloqueTabla = m && m[1] ? m[1].trim() : '';
  if (!bloqueTabla) return [];

  const items = [];
  
  // Expresión regular ESTRICTA para EAN de 13 dígitos (\d{13}) tal como lo requieres
  const re = /(\d{13})\s+(\d{3,})\s+([\s\S]+?)\s*(?:Und|UND|Un|UN|und|U\.M\.)\s+(\d+(?:[.,]\d+)?)\s+([$\d.,]+)/gmi;
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

function medipiel_limpiarOc_(val) {
  if (val === null || val === undefined) return "";
  let s = String(val).trim().replace(/^'/, "");
  if (/e/i.test(s)) {
    let num = Number(s);
    if (!isNaN(num)) s = num.toFixed(0);
  }
  if (s.includes(".")) s = s.split(".")[0];
  return s.trim();
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