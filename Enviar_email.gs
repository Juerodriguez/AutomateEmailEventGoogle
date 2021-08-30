// Instrucciones:
// - Fila 16: Nombre de hoja de calculo
// - Fila 40: Id del calendario que contiene los eventos
// - Fila 42: Primera fecha en la que comienzan los eventos a consultar
// - Fila 43: Ultima fecha en la que terminan los eventos a consultar

function onOpen() {
  const spreadsheet = SpreadsheetApp.getActive();
  const menu = [{name: 'Invitar por calendario', functionName: 'enviarCorreos_main'}];
  spreadsheet.addMenu('Eventos', menu);
}


function enviarCorreos_main() {
 const libro = SpreadsheetApp.getActiveSpreadsheet();
 libro.setActiveSheet(libro.getSheetByName("Prueba"));  // Nombre de hoja de calculo
 const hoja = SpreadsheetApp.getActiveSheet();
 const filas = hoja.getRange("A2:E500").getValues();
 var cont = 0; 
  
 for (indiceFila in filas) {
   var candidato = crearCandidato(filas[indiceFila]);
   cont = cont + 1;
   enviarCorreo(candidato, cont);
   enviarCalendario(candidato)   
  }
}


function crearCandidato(datosFila) {
  const candidato = {
    nick: datosFila[0],
    nombre: datosFila[2],
    email: datosFila[4],
    emailEnviado: datosFila[14],
  };
  return candidato;
}


function enviarCalendario(candidato){
 let calendarId = 'jijr6ck2epov3ia9j0ndg6juj4@group.calendar.google.com'; // Id del calendario que contiene los eventos
 var calendar = CalendarApp.getCalendarById(calendarId);
 let startDate = new Date("2021-08-01"); // Primera fecha en la que comienzan los eventos a consultar
 let endDate = new Date("2021-10-30"); // Ultima fecha en la que terminan los eventos a consultar
 let calEvents = calendar.getEvents(startDate, endDate);
 for (var i = 0; i < calEvents.length; i++) {
   let event = calEvents[i];
   event.addGuest(candidato.email);
  }
}

function enviarCorreo(candidato, cont) {
 if (candidato.email == "") {
   return;}
 else if (candidato.emailEnviado == "true") return;
 const plantilla = HtmlService.createTemplateFromFile('Plantilla_mensaje');  // Llenar con el nombre de la plantilla 
 plantilla.candidato = candidato;
 const mensaje = plantilla.evaluate().getContent();
  
 MailApp.sendEmail({
   to: candidato.email,
   subject: "Space Apps Challenge 2021",
   htmlBody: mensaje 
  });
 var campo = "O" + cont;
 hoja.getRange(campo)..setValue("true");
}
