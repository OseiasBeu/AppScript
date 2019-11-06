function sendEmails() {

var sheet = SpreadsheetApp.getActiveSheet();
var data = sheet.getDataRange().getValues();
Logger.log(sheet.getLastRow())
 for (var i = 0; i < data.length; ++i) {
 
  var emailAddress = sheet.getRange(i+4,2).getValue();  
  var message = sheet.getRange(i+4,5).getValue();     
  var subject = sheet.getRange(i+4,1).getValue();
  var copyto = sheet.getRange(i+4,3).getValue();
  var bccto = sheet.getRange(i+4,4).getValue();
    

    //document.getElementById(message).innerHTML='<object type="text/html" data="home.html" ></object>';
    
 try
  {
    MailApp.sendEmail(emailAddress, subject, "", {htmlBody: message, cc: copyto, bcc: bccto});
  
    sheet.getRange(i+2, 6).setValue("Email enviado");
    SpreadsheetApp.flush();
  }
  catch(e)
  {
    sheet.getRange(i+4, 6).setValue("Falha de Envio");
  }

}
return Browser.msgBox("EMAILS ENVIADOS COM SUCESSO!");
//function map() {
//var sheet = SpreadsheetApp.getActiveSheet();
//var i=0
//
//  
//  // Log the name of every file in the user's Drive.
//var files = DriveApp.getFiles();
//  while (files.hasNext()) {
//var file = files.next();
//   sheet.getRange(i+4, 12).setValue(file.getName());
//   i++
//  }
//}
  
 
}

