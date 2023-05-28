function enviarMail(){
  var sheet = SpreadsheetApp.getActiveSheet(); //Obtiene la hoja de calculo
  var startRow = 1730;  // primer fila para procesar
    
  var sourceRange = sheet.getRange ("B2:B");
  var allrowKws = sourceRange.getValues();
  var Alast = allrowKws.filter(String).length; // Obtiene la ultima fila con valores y si se pone muy lento probablemente se deberia hacer un respaldo cada 2 meses para que no haga un scan de 200 filas
  
  var dataRange = sheet.getRange(startRow, 1, Alast, 20);
  // Obtengo los valores de las filas
  var data = dataRange.getValues();
    for (var i = 0; i < data.length; ++i) {
       var row = data[i];
       if (row[19] == "Enviado") {
       }
       else {
         var numNC = row[1];
         var nroCliente = row[3];
         var cliente = row[4];
         var sF = row[10];
         var fFC = row[12];
         var fNC = row[13];
         var motivo = row[15];
         var vendedor = row[17];
         var monto = row[18];
         var observacion = row[16];
         if (observacion == "") {
           var observacion = "No tiene comentarios";
         }         
          if (Object.prototype.toString.call(fFC) !== '[object Date]') {
          Logger.log('The value "' + String(mes) + '" is not a date.');
          continue;
        }
        if (Object.prototype.toString.call(fNC) !== '[object Date]') {
          Logger.log('The value "' + String(mes) + '" is not a date.');
          continue;
        }

        var mails = 'anabella.carballo@elauditor.com.ar, marcos.torres@elauditor.com.ar, luciano.guzman@elauditor.com.ar, federico.sabatino@elauditor.com.ar, emmanuel.palomeque@elauditor.com.ar'
        var hoy = new Date(); 
        var mes = Utilities.formatDate(hoy,Session.getTimeZone(), "MM");
        var dia = Utilities.formatDate(hoy,Session.getTimeZone(), "dd");
        var fecha = dia+"/"+mes;
        var fechaFCmes = Utilities.formatDate(fFC,Session.getTimeZone(), "MM");
        var fechaFCdia = Utilities.formatDate(fFC,Session.getTimeZone(), "dd");
        var fechaFC = fechaFCdia + "/" + fechaFCmes;
        var fechaNCdia = Utilities.formatDate(fNC,Session.getTimeZone(),'dd');
        var fechaNCmes = Utilities.formatDate(fNC,Session.getTimeZone(),"MM");
        var fechaNC = fechaNCdia + "/" + fechaNCmes;
        var asunto = fecha + ' Entrega de SNC: ' + ' Logistica: ' + ' El Auditor '; 

        Logger.log(fechaFC)

        MailApp.sendEmail({
        to: mails,
        subject: asunto,
        body: "SNC: " + numNC + "\n" + "Numero de cliente: " + nroCliente + "\n" + "Nombre del cliente: " + cliente + "\n"  + "SF: " + sF + "\n" + "Fecha FC: " + fechaFC + "\n" + "Fecha NC: " + fechaNC + "\n" + "Motivo: " + motivo + "\n" + "Vendedor: " + vendedor + "\n" + "Monto: $" + monto + "\n" + "Observacion: " + observacion + "\n",
      });
      sheet.getRange(startRow + i, 20).setValue("Enviado");
    }   
  }
}

function onEdit(event){
  
  var ColK = 4;  // Número de la Columna "K"

  var changedRange = event.source.getActiveRange();
  if (changedRange.getColumn() == ColK) {
    // Una celda de la Columna K ha sido editada
    var state = changedRange.getValue();
    var adjacentFch = event.source.getActiveSheet().getRange(changedRange.getRow(),ColK+10);
    var timestamp = new Date(); // Obtener la fecha actual
    // Dependiendo del valor de la celda será lo que se hará
    adjacentFch.setValue(timestamp); 
  }
}

function onOpen() {
   var ui = SpreadsheetApp.getUi();
   ui.createMenu("Procesar")
  .addItem("Enviar mail", "enviarMail")
  .addToUi();
  
}