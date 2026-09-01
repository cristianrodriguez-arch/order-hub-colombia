/**
 * Función de prueba para diagnosticar problemas de permisos al mover archivos en Drive.
 * Selecciona esta función en el menú superior y haz clic en "Ejecutar".
 * Luego revisa el "Registro de ejecución" en la parte inferior.
 */
function testMoverArchivo() {
  // 1. REEMPLAZA ESTOS VALORES CON TUS IDs REALES PARA LA PRUEBA
  const fileId = "ID_DEL_ARCHIVO_PDF_AQUI"; // Ej: "1-JL1bKuNVZrAEMIl2NTINFMtXyKEIbEp"
  const folderOrigenId = "10k7ltzA6CclAY8z9wDhEACrGl5-3bGnx"; // Carpeta Medipiel
  const folderDestinoId = "1dMqDRf92i5qgjUnEwvgM7qO1931YUu6z"; // Carpeta Procesados

  Logger.log("Iniciando prueba de movimiento de archivo...");
  Logger.log("Archivo ID: " + fileId);
  Logger.log("Origen ID: " + folderOrigenId);
  Logger.log("Destino ID: " + folderDestinoId);

  // INTENTO 1: API Avanzada de Drive (Recomendada para Unidades Compartidas)
  try {
    Logger.log("Intentando mover con Drive API Avanzada (Drive.Files.patch)...");
    
    // NOTA: Requiere habilitar el servicio "Drive API" en el panel izquierdo (+)
    Drive.Files.patch({}, fileId, {
      addParents: folderDestinoId,
      removeParents: folderOrigenId,
      supportsAllDrives: true // Crucial para Shared Drives
    });
    
    Logger.log("✅ ¡ÉXITO! Archivo movido correctamente con API Avanzada.");
    return; // Si funciona, terminamos aquí
    
  } catch (e1) {
    Logger.log("❌ Fallo con API Avanzada: " + e1.message);

    // INTENTO 2: DriveApp Nativo (El que fallaba en tu captura)
    try {
      Logger.log("Intentando mover con DriveApp Nativo (moveTo)...");
      const file = DriveApp.getFileById(fileId);
      const folderDestino = DriveApp.getFolderById(folderDestinoId);
      file.moveTo(folderDestino);
      Logger.log("✅ ¡ÉXITO! Archivo movido correctamente con DriveApp.");
      
    } catch (e2) {
      Logger.log("❌ Fallo con DriveApp Nativo: " + e2.message);
      Logger.log("======================================================");
      Logger.log("CONCLUSIÓN DEL DIAGNÓSTICO:");
      Logger.log("Definitivamente tienes un bloqueo de permisos por parte de Google Workspace.");
      Logger.log("Causas más comunes:");
      Logger.log("1. Si es una Unidad Compartida (Shared Drive), necesitas el rol de 'Gestor de Contenido' o 'Administrador', el rol de 'Editor' no puede mover archivos.");
      Logger.log("2. El archivo PDF fue subido/creado por un usuario externo a tu organización, por lo que tú no eres el Propietario del archivo y Google no te deja sacarlo de su carpeta original.");
      Logger.log("======================================================");
    }
  }
}