/**
 * Convierte PDFs en una carpeta a Docs, extrae campos específicos
 * y los guarda en una hoja "Procesados" de un Spreadsheet.
 * Si la hoja o los encabezados no existen, los crea.
 */
function extraerDatosDePdfsEnCarpeta() {
  try {
    // REEMPLAZA estos IDs con los tuyos:
    var folderId = "1clf-Bt8LTjEuSH5pX8KXkcCF7CGAkzRo";  // Carpeta con PDFs
    var sheetId  = "1RUXFgyGT2nFNRfU41-W9IVaGFnTw8VpfJDO2FI0Cslo"; // Spreadsheet
    
    // Hoja donde pondremos los datos
    var nombreHoja = "Procesados";

    // 1) Acceder a la carpeta y a los PDFs
    var folder = DriveApp.getFolderById(folderId);
    var pdfFiles = folder.getFilesByType(MimeType.PDF);

    // 2) Abrir el spreadsheet y obtener/crear la hoja "Procesados"
    var ss = SpreadsheetApp.openById(sheetId);
    var hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) {
      hoja = ss.insertSheet(nombreHoja);
      Logger.log('Se creó la hoja "Procesados".');
    }

    // 3) Verificar si la fila 1 está vacía. Si es así, creamos los encabezados.
    var headers = [
      "Razón Social",
      "NIT",
      "Dirección",
      "Teléfono",
      "Ciudad/Depto",
      "Proveedor Razón Social",
      "Proveedor NIT",
      "Proveedor Dirección",
      "Proveedor Email",
      "Número Pedido",
      "Fecha Pedido",
      "Comprador",
      "Moneda",
      "Condición de Pago",
      "Datos Entrega (Almacén/Dirección)",
      "Item",
      "Servicio",
      "Descripción",
      "Fecha Entrega",
      "Cantidad",
      "Unidad de Medida (UM)",
      "IVA (%)",
      "Valor Unitario",
      "Descuento",
      "Valor Total",
      "Subtotal",
      "IVA",
      "Total con IVA"
    ];

    var headerRange = hoja.getRange(1, 1, 1, headers.length);
    var firstRow = headerRange.getValues()[0];
    var fila1Vacia = firstRow.every(function(cell) { return cell === ""; });
    if (fila1Vacia) {
      headerRange.setValues([headers]);
      Logger.log("Encabezados creados en la fila 1.");
    }

    // 4) Recorrer cada PDF
    while (pdfFiles.hasNext()) {
      var pdfFile = pdfFiles.next();

      // 4.1) Convertir PDF a Google Doc (OCR si escaneado)
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

      // 4.2) Abrir y extraer texto
      var doc = DocumentApp.openById(docFile.id);
      var texto = doc.getBody().getText();

      // 4.3) Extraer los campos (ajusta las claves a tu PDF real)
      //      Ej.: Si en el PDF aparece "Razón Social D1 S.A.S" sin dos puntos, 
      //      cambia la claveInicio a "Razón Social ".
      var razonSocial          = extraerValor(texto, "Razón Social:", "\n");
      var nit                  = extraerValor(texto, "NIT:", "\n");
      var direccion            = extraerValor(texto, "Dirección:", "\n");
      var telefono             = extraerValor(texto, "Teléfono:", "\n");
      var ciudadDepto          = extraerValor(texto, "Ciudad/Depto:", "\n");

      var provRazonSocial      = extraerValor(texto, "Proveedor Razón Social:", "\n");
      var provNIT              = extraerValor(texto, "Proveedor NIT:", "\n");
      var provDireccion        = extraerValor(texto, "Proveedor Dirección:", "\n");
      var provEmail            = extraerValor(texto, "Proveedor Email:", "\n");

      var numeroPedido         = extraerValor(texto, "Número Pedido:", "\n");
      var fechaPedido          = extraerValor(texto, "Fecha Pedido:", "\n");
      var comprador            = extraerValor(texto, "Comprador:", "\n");
      var moneda               = extraerValor(texto, "Moneda:", "\n");
      var condicionPago        = extraerValor(texto, "Condición de Pago:", "\n");
      var datosEntrega         = extraerValor(texto, "Datos Entrega (Almacén/Dirección):", "\n");

      var item                 = extraerValor(texto, "Item:", "\n");
      var servicio             = extraerValor(texto, "Servicio:", "\n");
      var descripcion          = extraerValor(texto, "Descripción:", "\n");
      var fechaEntrega         = extraerValor(texto, "Fecha Entrega:", "\n");
      var cantidad             = extraerValor(texto, "Cantidad:", "\n");
      var unidadMedida         = extraerValor(texto, "Unidad de Medida (UM):", "\n");
      var ivaPorc              = extraerValor(texto, "IVA (%):", "\n");
      var valorUnitario        = extraerValor(texto, "Valor Unitario:", "\n");
      var descuento            = extraerValor(texto, "Descuento:", "\n");
      var valorTotal           = extraerValor(texto, "Valor Total:", "\n");
      var subtotal             = extraerValor(texto, "Subtotal:", "\n");
      var iva                  = extraerValor(texto, "IVA:", "\n");
      var totalConIva          = extraerValor(texto, "Total con IVA:", "\n");

      // 4.4) Insertar los datos en la hoja
      var ultimaFila = hoja.getLastRow() + 1;
      var filaDatos = [
        razonSocial,
        nit,
        direccion,
        telefono,
        ciudadDepto,
        provRazonSocial,
        provNIT,
        provDireccion,
        provEmail,
        numeroPedido,
        fechaPedido,
        comprador,
        moneda,
        condicionPago,
        datosEntrega,
        item,
        servicio,
        descripcion,
        fechaEntrega,
        cantidad,
        unidadMedida,
        ivaPorc,
        valorUnitario,
        descuento,
        valorTotal,
        subtotal,
        iva,
        totalConIva
      ];

      hoja.getRange(ultimaFila, 1, 1, filaDatos.length).setValues([filaDatos]);

      // 4.5) (Opcional) Eliminar el Doc temporal
      DriveApp.getFileById(docFile.id).setTrashed(true);
    }

    Logger.log("¡Proceso completado con éxito!");

  } catch (error) {
    Logger.log("Error: " + error);
  }
}

/**
 * Extrae el texto que sigue a 'claveInicio' hasta que encuentra 'delimFin'.
 * Por ejemplo, si en tu PDF convertido aparece "Razón Social: D1 S.A.S\n"
 *   extraerValor(texto, "Razón Social:", "\n") => "D1 S.A.S"
 */
function extraerValor(texto, claveInicio, delimFin) {
  var start = texto.indexOf(claveInicio);
  if (start === -1) return "";
  start += claveInicio.length;

  var end = texto.indexOf(delimFin, start);
  if (end === -1) {
    end = texto.length;
  }

  return texto.substring(start, end).trim();
}
