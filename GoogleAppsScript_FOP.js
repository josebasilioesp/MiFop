/**
 * ==============================================================
 * GOOGLE APPS SCRIPT PARA FIBRA ÓPTICA DEL PERÚ (Mi FOP)
 * ==============================================================
 * Este script permite que tu App:
 * 1. Actualice la Deuda a 0.00 y el Estado a "AL DIA" en tu Excel automáticamente.
 * 2. Reciba la foto o captura del voucher de Yape/Plin.
 * 3. Guarde la foto automáticamente en tu Google Drive (en la carpeta "Comprobantes FOP").
 * 4. Guarde el registro en una nueva pestaña "Historial_Pagos" con el enlace a la foto.
 * 
 * INSTRUCCIONES DE INSTALACIÓN (Toma menos de 2 minutos):
 * 1. Abre tu Google Sheet: https://docs.google.com/spreadsheets/d/1tgLlfBpNQ5JzuJ--XM01bgo4UABnAd8pmWrOJxjLUxU/edit
 * 2. En el menú superior haz clic en: Extensiones > Apps Script.
 * 3. Borra todo lo que esté en la pantalla y pega este código completo.
 * 4. Haz clic en el botón azul superior: "Implementar" (o Deploy) > "Nueva implementación".
 * 5. En el engranaje tipo selecciona: "Aplicación web".
 * 6. Configura:
 *    - Descripción: "API Pagos FOP"
 *    - Ejecutar como: "Yo" (tu cuenta de Google)
 *    - Quién tiene acceso: "Cualquier persona" (Anyone) -> ¡Muy importante para que la app pueda enviar los pagos!
 * 7. Haz clic en "Implementar", autoriza los permisos de Google Drive y copia la URL generada.
 *    (Ejemplo de URL: https://script.google.com/macros/s/AKfycb.../exec)
 */

function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);
    
    var dni = String(data.dni || '').trim();
    var monto = parseFloat(data.monto || 0);
    var canal = data.canal || 'Plin';
    var numOperacion = data.numOperacion || 'S/N';
    var fechaStr = data.fecha || Utilities.formatDate(new Date(), "America/Lima", "dd/MM/yyyy");
    var horaStr = data.hora || Utilities.formatDate(new Date(), "America/Lima", "hh:mm a");
    var voucherBase64 = data.voucherBase64 || null;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetClientes = ss.getSheets()[0]; // Primera hoja (Clientes)
    var rows = sheetClientes.getDataRange().getValues();
    
    var clienteNombre = "Abonado FOP";
    var planNombre = "Plan FOP";
    var filaEncontrada = -1;
    
    // 1. Buscar la fila del cliente por DNI (Columna 1)
    for (var i = 1; i < rows.length; i++) {
      var rowDni = String(rows[i][0] || '').trim().replace(/\D/g, '');
      var dniLimpio = dni.replace(/\D/g, '');
      if (rowDni === dniLimpio) {
        filaEncontrada = i + 1; // 1-based index
        clienteNombre = rows[i][1] || clienteNombre;
        
        // Auto-detectar si las columnas están desplazadas por comas en la dirección
        var c4Str = String(rows[i][4] || '').trim();
        var isShifted = /básico|estándar|premium/i.test(c4Str) || (!isNaN(parseFloat(rows[i][5])) && rows[i][5] !== '');
        
        var colDeuda = isShifted ? 8 : 7;
        var colEstado = isShifted ? 9 : 8;
        planNombre = isShifted ? rows[i][4] : rows[i][3];
        
        // ACTUALIZAR A DEUDA 0.00 Y ESTADO "AL DIA"
        sheetClientes.getRange(filaEncontrada, colDeuda).setValue(0.00);
        sheetClientes.getRange(filaEncontrada, colEstado).setValue("AL DIA");
        break;
      }
    }
    
    // 2. Almacenar la foto del Voucher en Google Drive (si se adjuntó)
    var voucherUrl = "Sin comprobante adjunto";
    if (voucherBase64) {
      try {
        var folderName = "Comprobantes FOP - Medios de Pago";
        var folders = DriveApp.getFoldersByName(folderName);
        var targetFolder;
        if (folders.hasNext()) {
          targetFolder = folders.next();
        } else {
          targetFolder = DriveApp.createFolder(folderName);
        }
        
        // Limpiar encabezado data:image/...;base64,
        var base64Clean = voucherBase64.replace(/^data:image\/\w+;base64,/, "");
        var decodedBytes = Utilities.base64Decode(base64Clean);
        var fileName = "VOUCHER_" + dni + "_" + Utilities.formatDate(new Date(), "America/Lima", "yyyyMMdd_HHmmss") + ".jpg";
        var blob = Utilities.newBlob(decodedBytes, "image/jpeg", fileName);
        
        var file = targetFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        voucherUrl = file.getUrl();
      } catch (errDrive) {
        voucherUrl = "Error al guardar en Drive: " + errDrive.toString();
      }
    }
    
    // 3. Registrar en la pestaña "Historial_Pagos"
    var sheetPagos = ss.getSheetByName("Historial_Pagos");
    if (!sheetPagos) {
      sheetPagos = ss.insertSheet("Historial_Pagos");
      sheetPagos.appendRow([
        "Fecha", "Hora", "DNI", "Cliente", "Plan", "Monto Pagado", "Método", "N° Operación", "Comprobante (Voucher)", "Estado"
      ]);
      sheetPagos.getRange(1, 1, 1, 10).setFontWeight("bold").setBackground("#001b68").setFontColor("#ffffff");
    }
    
    sheetPagos.appendRow([
      fechaStr,
      horaStr,
      dni,
      clienteNombre,
      planNombre,
      monto,
      canal,
      numOperacion,
      voucherUrl,
      "CONFIRMADO"
    ]);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Pago registrado con éxito. Deuda actualizada a 0.00 en Excel.",
      voucherUrl: voucherUrl,
      cliente: clienteNombre
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "online",
    message: "Servicio Webhook de Pagos FOP activo."
  })).setMimeType(ContentService.MimeType.JSON);
}
