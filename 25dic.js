function testExtraerTabla() {
  try {
    // 1) ID del Documento de Google (NO el PDF)
    //    Ajusta a tu docFile.id resultante de la conversión
    var docId = "1oa2UIah3NAWQ9IUpqTGgBqSHCXExewXxpxUsSmWuLCU";

    // 2) Abre el Doc y obtiene el texto
    var doc = DocumentApp.openById(docId);
    var textoCompleto = doc.getBody().getText();

    // 3) Parsear con la función anterior
    var datos = parsearLineasSeparadas(textoCompleto);

    // 4) Revisar en logs
    Logger.log("Item: " + datos.item);
    Logger.log("Servicio: " + datos.servicio);
    Logger.log("Descripción: " + datos.descripcion);
    Logger.log("Fecha entrega: " + datos.fechaEntrega);
    Logger.log("Oferta No.: " + datos.ofertaNo);
    Logger.log("CTD: " + datos.ctd);
    Logger.log("UM: " + datos.um);
    Logger.log("% Iva: " + datos.porcIva);
    Logger.log("Val DTO: " + datos.valDto);
    Logger.log("Valor unitario: " + datos.valorUnit);
    Logger.log("Valor total: " + datos.valorTotal);

    Logger.log("Subtotal: " + datos.subtotal);
    Logger.log("Subtotal con descuento: " + datos.subtotalConDesc);
    Logger.log("IVA: " + datos.iva);
    Logger.log("Total con IVA: " + datos.totalConIva);

    // 5) (Opcional) Guardar en una hoja "Procesados" de un Spreadsheet
    //    Ajusta sheetId a tu hoja y crea encabezados si quieres.
    /*
    var sheetId = "TU_SPREADSHEET_ID";
    var ss = SpreadsheetApp.openById(sheetId);
    var hoja = ss.getSheetByName("Procesados") || ss.insertSheet("Procesados");

    // Si quieres, crea encabezados si están vacíos
    var headers = ["Item","Servicio","Descripción","Fecha entrega","Oferta No.","CTD","UM","% Iva","Val DTO","Valor unit","Valor total","Subtotal","Subtotal c/ desc","IVA","Total c/ IVA"];
    if (hoja.getLastRow() === 0) {
      hoja.appendRow(headers);
    }

    // Prepara la fila
    var fila = [
      datos.item,
      datos.servicio,
      datos.descripcion,
      datos.fechaEntrega,
      datos.ofertaNo,
      datos.ctd,
      datos.um,
      datos.porcIva,
      datos.valDto,
      datos.valorUnit,
      datos.valorTotal,
      datos.subtotal,
      datos.subtotalConDesc,
      datos.iva,
      datos.totalConIva
    ];
    hoja.appendRow(fila);
    */
  } catch (err) {
    Logger.log("Error en testExtraerTabla: " + err);
  }
}
