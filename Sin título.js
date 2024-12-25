/**
 * Lee PDFs de la carpeta '1clf-Bt8LTjEuSH5pX8KXkcCF7CGAkzRo',
 * extrae campos específicos y los guarda en la hoja "Procesados"
 * del Spreadsheet 'RUXFgyGT2nFNRfU41-W9IVaGFnTw8VpfJDO2FI0Cslo'.
 */
function extraerDatosDePdfsEnCarpeta() {
  try {
    // 1. IDs de carpeta y Spreadsheet
    var folderId = "1clf-Bt8LTjEuSH5pX8KXkcCF7CGAkzRo"; 
    var sheetId  = "RUXFgyGT2nFNRfU41-W9IVaGFnTw8VpfJDO2FI0Cslo";

    // Acceder a la carpeta
    var folder = DriveApp.getFolderById(folderId);
    // Obtener PDFs
    var pdfFiles = folder.getFilesByType(MimeType.PDF);

    // Abrir la hoja "Procesados"
    var ss = SpreadsheetApp.openById(sheetId);
    var hoja = ss.getSheetByName("Procesados");
    if (!hoja) {
      throw new Error('No existe la hoja "Procesados" en el spreadsheet.');
    }

    // 2. Recorrer cada PDF y procesarlo
    while (pdfFiles.hasNext()) {
      var pdfFile = pdfFiles.next();

      // 2.1. Convertir PDF a Google Docs (con OCR)
      var resource = {
        title: "Temporal_" + pdfFile.getName(),
        mimeType: MimeType.GOOGLE_DOCS
      };
      var docFile = Drive.Files.copy(
        resource,
        pdfFile.getId(),
        {
          ocr: true,
          convert: true
        }
      );

      // 2.2. Extraer texto del Google Doc
      var doc = DocumentApp.openById(docFile.id);
      var texto = doc.getBody().getText();

      // 2.3. Parsear SOLO los campos solicitados
      //     Ajusta las claves y el delimitador "\n" según tu PDF real

      var pedidoServicios        = extraerValor(texto, "Pedido de Servicios", "\n");
      var fecha                  = extraerValor(texto, "Fecha:", "\n");
      var comprador              = extraerValor(texto, "Comprador:", "\n");
      var condicionPago          = extraerValor(texto, "Condición de pago:", "\n");
      var solicitudPedido        = extraerValor(texto, "Solicitud de pedido:", "\n");
      var solicitadoPor          = extraerValor(texto, "Solicitado por:", "\n");
      var almacen                = extraerValor(texto, "Almacén", "\n");
      var direccion              = extraerValor(texto, "Dirección:", "\n");

      var item                   = extraerValor(texto, "Item", "\n");
      var servicio               = extraerValor(texto, "Servicio", "\n");
      var descripcion            = extraerValor(texto, "Descripción", "\n");
      var fechaEntrega           = extraerValor(texto, "Fecha entrega", "\n");
      var ctd                    = extraerValor(texto, "CTD", "\n");
      var porcentajeIva          = extraerValor(texto, "% Iva", "\n");
      var valorUnitario          = extraerValor(texto, "Valor unitario", "\n");
      var valorTotal             = extraerValor(texto, "Valor total", "\n");

      var subtotal               = extraerValor(texto, "Subtotal:", "\n");
      var subtotalConDescuento   = extraerValor(texto, "Subtotal con descuento:", "\n");
      var iva                    = extraerValor(texto, "IVA:", "\n");
      var totalConIva            = extraerValor(texto, "Total con IVA:", "\n");
      var notas                  = extraerValor(texto, "NOTAS:", "\n");

      // 2.4. Guardar una nueva fila en la hoja
      var ultimaFila = hoja.getLastRow() + 1;
      hoja.getRange(ultimaFila, 1).setValue(new Date());        // Fecha de procesamiento
      hoja.getRange(ultimaFila, 2).setValue(pdfFile.getName()); // Nombre del PDF

      hoja.getRange(ultimaFila, 3).setValue(pedidoServicios);
      hoja.getRange(ultimaFila, 4).setValue(fecha);
      hoja.getRange(ultimaFila, 5).setValue(comprador);
      hoja.getRange(ultimaFila, 6).setValue(condicionPago);
      hoja.getRange(ultimaFila, 7).setValue(solicitudPedido);
      hoja.getRange(ultimaFila, 8).setValue(solicitadoPor);
      hoja.getRange(ultimaFila, 9).setValue(almacen);
      hoja.getRange(ultimaFila, 10).setValue(direccion);

      hoja.getRange(ultimaFila, 11).setValue(item);
      hoja.getRange(ultimaFila, 12).setValue(servicio);
      hoja.getRange(ultimaFila, 13).setValue(descripcion);
      hoja.getRange(ultimaFila, 14).setValue(fechaEntrega);
      hoja.getRange(ultimaFila, 15).setValue(ctd);
      hoja.getRange(ultimaFila, 16).setValue(porcentajeIva);
      hoja.getRange(ultimaFila, 17).setValue(valorUnitario);
      hoja.getRange(ultimaFila, 18).setValue(valorTotal);

      hoja.getRange(ultimaFila, 19).setValue(subtotal);
      hoja.getRange(ultimaFila, 20).setValue(subtotalConDescuento);
      hoja.getRange(ultimaFila, 21).setValue(iva);
      hoja.getRange(ultimaFila, 22).setValue(totalConIva);
      hoja.getRange(ultimaFila, 23).setValue(notas);

      // 2.5. (Opcional) Eliminar Doc temporal
      DriveApp.getFileById(docFile.id).setTrashed(true);
    }

    Logger.log("¡Proceso completado con éxito!");
  } catch (e) {
    Logger.log("Error: " + e.message);
  }
}

/**
 * Función auxiliar para extraer texto entre la claveInicio y el primer salto de línea
 * que encuentre. Por ejemplo, si en la línea del PDF dice:
 * "Comprador: YTLAYTON"
 * y llamas extraerValor(texto, "Comprador:", "\n"), devolverá "YTLAYTON".
 *
 * @param {string} texto        Texto completo extraído del PDF
 * @param {string} claveInicio  Palabra clave (p.e. "Comprador:")
 * @param {string} delimFin     Normalmente "\n"
 */
function extraerValor(texto, claveInicio, delimFin) {
  var i = texto.indexOf(claveInicio);
  if (i === -1) return "";
  i += claveInicio.length;

  var j = texto.indexOf(delimFin, i);
  if (j === -1) {
    j = texto.length;
  }

  return texto.substring(i, j).trim();
}
