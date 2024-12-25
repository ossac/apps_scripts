function buscarYGuardarLiberaciones() {
  // --------------------------------------------------
  // 1) CONFIGURACIONES
  // --------------------------------------------------
  var SPREADSHEET_ID = "1RUXFgyGT2nFNRfU41-W9IVaGFnTw8VpfJDO2FI0Cslo";  // Reemplaza con tu ID de Spreadsheet
  var SHEET_NAME = "Liberaciones";           // Nombre de la hoja/pestaña
  var FOLDER_ID = "1clf-Bt8LTjEuSH5pX8KXkcCF7CGAkzRo"; // ID de la carpeta en Drive

  // --------------------------------------------------
  // 2) OBTENER OBJETOS Y VALIDAR
  // --------------------------------------------------
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(SHEET_NAME);
    var folder = DriveApp.getFolderById(FOLDER_ID);
    
    if (!sheet || !folder) {
      throw new Error("No se pudo acceder a la hoja o la carpeta especificada");
    }
  } catch (e) {
    Logger.log("Error de inicialización: " + e.toString());
    return;
  }

  // --------------------------------------------------
  // 3) DEFINIR LA BÚSQUEDA EN GMAIL (RANGO DE FECHAS)
  // --------------------------------------------------
  var AFTER_DATE  = "2024/01/01";
  var BEFORE_DATE = "2024/12/25";

  var query = 'has:attachment filename:pdf subject:"Liberación Pedido"' 
            + ' after:' + AFTER_DATE 
            + ' before:' + BEFORE_DATE;

  // Obtenemos los hilos que coinciden con la búsqueda
  var threads = GmailApp.search(query);
  var totalThreads = threads.length;
  
  if (totalThreads >= 500) {
    Logger.log("!!! ADVERTENCIA: Se alcanzó el límite de 500 resultados !!!");
  }

  // --------------------------------------------------
  // 4) RECORRER LOS CORREOS Y SUS ADJUNTOS
  // --------------------------------------------------
  var contador = 0;
  var errores = 0;
  
  threads.forEach(function(thread) {
    try {
      var messages = thread.getMessages();
      
      messages.forEach(function(message) {
        var attachments = message.getAttachments();
        
        attachments.forEach(function(attachment) {
          var fileName = attachment.getName();
          
          if (fileName.startsWith("Liberación Pedido") && fileName.toLowerCase().endsWith(".pdf")) {
            procesarArchivo(attachment, fileName, message, folder, sheet);
            contador++;
          }
        });
      });
    } catch (e) {
      Logger.log("Error procesando hilo: " + e.toString());
      errores++;
    }
  });
  
  // --------------------------------------------------
  // 8) LOG FINAL
  // --------------------------------------------------
  Logger.log("Procesamiento completado:");
  Logger.log("- Total de PDFs procesados: " + contador);
  Logger.log("- Total de errores: " + errores);
}

function procesarArchivo(attachment, fileName, message, folder, sheet) {
  try {
    // --------------------------------------------------
    // 5) REVISAR SI EL ARCHIVO YA EXISTE EN DRIVE
    // --------------------------------------------------
    var resultado = verificarDuplicado(folder, fileName, attachment);
    
    // --------------------------------------------------
    // 6) OBTENER LOS DATOS DEL CORREO
    // --------------------------------------------------
    var asunto = message.getSubject();
    var fecha = Utilities.formatDate(message.getDate(), "GMT-5", "yyyy-MM-dd HH:mm");
    var remitente = message.getFrom();
    var cuerpo = obtenerCuerpoFormateado(message);

    // --------------------------------------------------
    // 7) GUARDAR EN LA HOJA "Liberaciones"
    // --------------------------------------------------
    sheet.appendRow([
      asunto,           // Col A
      fecha,            // Col B
      remitente,        // Col C
      cuerpo,           // Col D
      fileName,         // Col E
      resultado.url     // Col F
    ]);

    if (resultado.status === "nuevo") {
      Logger.log("Archivo nuevo procesado: " + fileName);
    } else {
      Logger.log(resultado.message);
    }

  } catch (e) {
    Logger.log("Error procesando archivo " + fileName + ": " + e.toString());
  }
}

function verificarDuplicado(folder, fileName, attachment) {
  try {
    var existingFiles = folder.searchFiles('title = "' + fileName + '"');
    
    if (existingFiles.hasNext()) {
      var existingFile = existingFiles.next();
      return {
        status: "duplicado",
        message: "El archivo ya existe en Drive: " + fileName,
        url: existingFile.getUrl()
      };
    } else {
      var newFile = folder.createFile(attachment);
      return {
        status: "nuevo",
        message: "Archivo creado exitosamente",
        url: newFile.getUrl()
      };
    }
  } catch (error) {
    Logger.log("Error al verificar duplicado: " + error.toString());
    return {
      status: "error",
      message: "Error al procesar el archivo: " + error.toString(),
      url: "ERROR: No se pudo procesar el archivo"
    };
  }
}

function obtenerCuerpoFormateado(message) {
  try {
    return message.getPlainBody()
      .replace(/[\r\n]+/g, ' ') // Reemplaza saltos de línea con espacios
      .replace(/\s+/g, ' ')     // Normaliza espacios múltiples
      .trim()                   // Elimina espacios al inicio y final
      .substring(0, 200)        // Limita a 200 caracteres
      + (message.getPlainBody().length > 200 ? "..." : "");
  } catch (e) {
    Logger.log("Error al obtener el cuerpo del mensaje: " + e.toString());
    return "No se pudo obtener el contenido del mensaje";
  }
}