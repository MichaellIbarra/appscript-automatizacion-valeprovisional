
function getRowDataJSON() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheets()[0];
    var data = sheet.getDataRange().getValues();
  
    var solicitantesSheet = spreadsheet.getSheetByName('Solicitantes');
    var solicitantesData = solicitantesSheet.getDataRange().getValues();
    var datosSolicitantes = {};
    var hojaDatosCentro = spreadsheet.getSheetByName('Centro de Costos');
    var datosCentroRange = hojaDatosCentro.getDataRange().getValues();
    var datosCentro = {};
    for (var i = 1; i < solicitantesData.length; i++) {
      var [solicitante, telefono, dni, firmaS] = solicitantesData[i];
      datosSolicitantes[solicitante] = { telefono, dni, firmaS };
      //   { 'Michaell Ibarra Martinez': 
      //  { telefono: 979333493,
      //    dni: 72036366,
      //    firmaS: 'https://drive.google.com/open?id=17pQRsAt4j_hJFytexmqtVL0YoRNLtgQw' },
    }
    //
    for (var i = 1; i < datosCentroRange.length; i++) {
      var [clave, digitos, nombre, firma] = datosCentroRange[i];
      datosCentro[clave] = { digitos, nombre, firma };
      //   {'080101 - DONACIONES': 
      //  { digitos: '080101',
      //    nombre: 'DONACIONES',
      //    firma: 'https://drive.google.com/open?id=13RsmIhyVzGjaQ3eDyWuhq_BhZd_War7F' }
    }
  
    var jsonData = [];
    for (var i = 1; i < data.length; i++) {
      var [marcaTemporal, correo, importe, fechaEntrega, fechaJustificacion, centroCosto, area, actividad, solicitante, telefonoSolicitante, dni, firmaS, estado] = data[i];
      var rowData = {
        'N°': i,
        'Marca temporal': marcaTemporal,
        'Dirección de correo electrónico': correo,
        'Importe S/': importe,
        'Fecha de Entrega de dinero': fechaEntrega,
        'Fecha de Justificación': fechaJustificacion,
        'Centro de Costo': centroCosto,
        'Área': area,
        'Actividad': actividad,
        'Solicitante': solicitante,
        'Teléfono Del Solicitante': telefonoSolicitante,
        'Dni': dni,
        'Firma': firmaS,
        'Estado': estado
      };
  
      if (solicitante in datosSolicitantes) {
        const { telefono, dni, firmaS } = datosSolicitantes[solicitante];
        rowData['Teléfono Del Solicitante'] = telefono;
        rowData['Dni'] = dni;
        rowData['Firma'] = firmaS;
        sheet.getRange(i + 1, 10).setValue(telefono);
        sheet.getRange(i + 1, 11).setValue(dni);
        sheet.getRange(i + 1, 12).setValue(firmaS);
      }
  
      if (centroCosto in datosCentro) {
        const { digitos, nombre, firma } = datosCentro[centroCosto];
        rowData['Centro de Costo'] = digitos;
        rowData['Actividad'] = nombre;
        sheet.getRange(i + 1, 6).setNumberFormat('@');
        sheet.getRange(i + 1, 6).setValue(digitos);
        sheet.getRange(i + 1, 7).setValue(nombre);
        sheet.getRange(i + 1, 14).setValue(firma);
      }
  
      jsonData.push(rowData);
  
      // [ { 'N°': 1,
      // 'Marca temporal': Sun Jul 02 2023 23:16:02 GMT-0500 (Peru Standard Time),
      // 'Dirección de correo electrónico': 'michaell.ibarra@vallegrande.edu.pe',
      // 'Importe S/': 100,
      // 'Fecha de Entrega de dinero': Thu Jun 01 2023 00:00:00 GMT-0500 (Peru Standard Time),
      // 'Fecha de Justificación': Sun Jan 01 2023 00:00:00 GMT-0500 (Peru Standard Time),
      // 'Centro de Costo': 30204,
      // 'Área': 'MANTENIMIENTO',
      // Actividad: 'mantenimiento de computadoras',
      // Solicitante: 'Michaell Ibarra Martinez',
      // 'Teléfono Del Solicitante': 979333493,
      // Dni: 72036366,
      // Firma: 'https://drive.google.com/open?id=17pQRsAt4j_hJFytexmqtVL0YoRNLtgQw',
      // Estado: 'Aprobado' } ]
  
    }
  
    return JSON.stringify(jsonData);
  
  }

  function updateSheetValue(value, index) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    sheet.getRange(index + 2, 13).setValue(value);
    if (value === 'Aprobado' || value === 'Rechazado') {
      generarDocumentos();
    }
  }
  
  
  function generarDocumentos() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = spreadsheet.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var plantillaDoc = '1VMI7MiJdstM_43fpry4ZC17ysm2DTl3UWYCIACv2Bxg';
    var carpetaID = '11DRSpMCg6oWyP8VD4uHFgFR3tFtgyNZ1';
  
    var folder = DriveApp.getFolderById(carpetaID);
  
    var processedRows = PropertiesService.getScriptProperties().getProperty("processedRows") || "";
  
    for (var i = 1; i < data.length; i++) {
      var rowData = data[i];
      var emailAddress = rowData[1];
  
      var status = rowData[12]; // Valor en la columna 12 (M)
      // La línea de código "if (!status)" significa "si el estado no existe" o "si el estado es falso". En JavaScript, el signo de exclamación (!)
  
      if (!status) {
        sheet.getRange(i + 1, 13).setValue("En Proceso");
        status = "En Proceso";
      }
      if (processedRows.includes(String(i))) {
        continue;
      }
      // Verifica si el estado es "En Proceso"
      if (status === "En Proceso") {
        // No realiza ninguna acción, pasa a la siguiente iteración
        continue;
      }
      ///
      var docCopy = DriveApp.getFileById(plantillaDoc).makeCopy(folder);
      var doc = DocumentApp.openById(docCopy.getId());
      var body = doc.getBody();
      ///
      var fechaEntrega = Utilities.formatDate(rowData[3], "GMT", "dd/MM/yyyy");
      var fechaJustificacion = Utilities.formatDate(rowData[4], "GMT", "dd/MM/yyyy");
      var fecha = Utilities.formatDate(rowData[0], "GMT", "dd/MM/yyyy");
  
      body.replaceText('{{Soli}}', rowData[8])
        .replaceText('{{imp}}', rowData[2])
        .replaceText('{{Fh_dinero}}', fechaEntrega)
        .replaceText('{{Fh_justi}}', fechaJustificacion)
        .replaceText('{{CCosto}}', rowData[5])
        .replaceText('{{Área}}', rowData[6])
        .replaceText('{{Actividad}}', rowData[7])
        .replaceText('{{Cel}}', rowData[9])
        .replaceText('{{Dni}}', rowData[10])
        .replaceText('{{fecha}}', fecha)
  
      ///
      var elemento = body.findText('{{es}}').getElement();
      var textoEncontrado = elemento.asText();
      if (rowData[12] === 'Aprobado') {
        textoEncontrado.setForegroundColor('#00FF00'); // Verde
      } else {
        textoEncontrado.setForegroundColor('#FF0000'); // Rojo
      }
      elemento.replaceText('{{es}}', rowData[12]);
  
      ///
      var yearActual = new Date().getFullYear();
      var orden = String(i).padStart(6, '0') + '-' + yearActual;
      body.replaceText('{{orden}}', orden);
  
      ///
      var valueToEncrypt = fecha + rowData[8] + rowData[10] + orden;
      var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, valueToEncrypt);
      var encryptedValue = Utilities.base64Encode(rawHash);
      body.replaceText('{{enc}}', encryptedValue);
  
      ///
      var response = UrlFetchApp.fetch('https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=' + encodeURIComponent("Fue " + rowData[12] + " Vale Provisional N°" + orden + "\n" + "Nombre del Solicitante: " + rowData[8] + "\n" + "Dni: " + rowData[10] + "\n" + "Firma:" + encryptedValue));
      var imageBlob = response.getBlob();
  
      var lugarqr = body.findText('{{qr}}').getElement().asText().setText('');
      var imgqr = lugarqr.getParent().asParagraph().insertInlineImage(0, imageBlob);
      imgqr.setHeight(140);
      imgqr.setWidth(140);
  
      ///
      var imageURL = rowData[13];
      var imagenID = imageURL.split("open?id=")[1];
      var imagen = DriveApp.getFileById(imagenID);
      var lugarTexto = body.findText('{{Firma_a}}').getElement().asText().setText('');
      var img = lugarTexto.getParent().asParagraph().insertInlineImage(0, imagen);
      img.setHeight(85);
      img.setWidth(155);
  
      ///
      var imgSurl = rowData[11];
      var imgSID = imgSurl.split("open?id=")[1];
      var imgS = DriveApp.getFileById(imgSID);
      var lugarText = body.findText('{{Firma_s}}').getElement().asText().setText('');
      var img = lugarText.getParent().asParagraph().insertInlineImage(0, imgS);
      img.setHeight(85);
      img.setWidth(155);
  
      doc.saveAndClose();
  
      var archivoPdf = folder.createFile(DriveApp.getFileById(docCopy.getId()).getAs(MimeType.PDF)).setName("Vale Provisional Prosip, Solicitante " + rowData[8] + ", N° " + orden);
  
      var subject, message;
      if (status === "Aprobado") {
        subject = 'Fue Aprobado La Solicitud del Vale Provisional N°' + orden;
        message = '<p style="font-weight: bold;">Estimado/a ' + rowData[8] + ',</p>\n\n<p style="font-size: 16px;">¡Tu solicitud ha sido <span style="color: green;">aprobada</span> con éxito!</p>\n\n<p style="font-family: Arial, sans-serif;">Aquí tienes el PDF adjunto.</p>\n\n<p style="font-style: italic;">Cualquier consulta o duda con respecto, puedes visitar nuestra página web y rellenar un formulario de reporte: <a href="https://sites.google.com/vallegrande.edu.pe/valexpress/contacto">Contactar</a> </p>';
      } else if (status === "Rechazado") {
        subject = 'Fue Rechazado la Solicitud del Vale Provisional N°' + orden;
        message = '<p style="font-weight: bold;">Estimado/a ' + rowData[8] + ',</p>\n\n<p style="font-size: 16px;">Lamentamos informarte que tu solicitud ha sido <span style="color: red;">rechazada</span>.</p>\n\n<p style="font-style: italic;">Cualquier consulta o duda con respecto, puedes visitar nuestra página web y rellenar un formulario de reporte: <a href="https://sites.google.com/vallegrande.edu.pe/valexpress/contacto">Contactar</a></p>';
      }
  
      MailApp.sendEmail(emailAddress, subject, message, { attachments: [archivoPdf], htmlBody: message });
  
      processedRows += i + ",";
      PropertiesService.getScriptProperties().setProperty("processedRows", processedRows);
  
      docCopy.setTrashed(true);
    }
  }
  
  // doGet() es una función de Google Apps Script que se ejecuta cuando se accede a la aplicación web correspondiente. Verifica si el usuario actual está autorizado y devuelve una página HTML diferente según el resultado.
  function doGet() {
    var usuariosAutorizados = ["michaell.ibarra@vallegrande.edu.pe", "jhon.melchor@vallegrande.edu.pe",];
    var usuarioActual = Session.getActiveUser().getEmail();
    if (usuariosAutorizados.includes(usuarioActual)) {
      return HtmlService.createHtmlOutputFromFile('inicio');
    } else {
      return HtmlService.createHtmlOutputFromFile('denegado');
    }
  }
  
  // enProceso() es una función que se llama cuando ocurre un evento específico en la hoja de cálculo y realiza ciertas acciones.
  function enProceso() {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var ultimaFila = hoja.getLastRow();
    var datosFila = hoja.getRange(ultimaFila, 1, 1, hoja.getLastColumn()).getValues()[0];
  
    var filasProcesadas = PropertiesService.getDocumentProperties().getProperty("filasProcesadas")?.split(",") || [];
  
    if (datosFila.some(celda => celda !== "") && !filasProcesadas.includes(String(ultimaFila))) {
      var direccionCorreo = datosFila[1];
      var asunto = "Estimado/a " + datosFila[8] + ". Su vale está en Proceso";
      var mensaje = `<p style="font-weight: bold;">Estimado/a ${datosFila[8]},</p>\n\n<p style="font-size: 16px;">Tu vale provisional está en proceso de evaluación.</p>\n\n<p style="font-size: 14px;">Este proceso puede tardar hasta 24 horas. Recibirás un correo de notificación con la decisión final, ya sea que tu vale haya sido <span style="color: green;">aprobado</span> o <span style="color: red;">rechazado</span>.</p>\n\n<p style="font-style: italic;">Cualquier consulta o duda con respecto, puedes visitar nuestra página web y rellenar un formulario de reporte: <a href="https://sites.google.com/vallegrande.edu.pe/valexpress/contacto">Contactar</a></p>`;
      MailApp.sendEmail({ to: direccionCorreo, subject: asunto, htmlBody: mensaje });
  
      filasProcesadas.push(String(ultimaFila));
      PropertiesService.getDocumentProperties().setProperty("filasProcesadas", filasProcesadas.join(","));
    }
  }
  