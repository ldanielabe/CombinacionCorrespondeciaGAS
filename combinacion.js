function onOpen() {
  var ui = DocumentApp.getUi();
  //Puede usar DocumentApp, SpreadsheetApp o FormApp.
  var topMenu = ui.createMenu('Combinacion Correspondencia');
  topMenu.addItem('Enviar Correspondencia', 'enviar');
  topMenu.addSeparator();
  topMenu.addToUi();
}



function enviar() {
  

  var url = 'https://docs.google.com/spreadsheets/d/12pJKRj69NJa77FMCXrq3UhmzbL1mpZf7cvgjH0TEbsM/edit?usp=sharing';
    
  var spreadsheet= SpreadsheetApp.openByUrl(url);
  
  
  var sheet = spreadsheet.getActiveSheet();
  var rows = sheet.getDataRange();
  var numCols = rows.getNumColumns();
  var numRows = rows.getNumRows();
  var string ="";
  var urldoc='https://docs.google.com/document/d/1KytFxHw21ivUsVKG4zYPHaXRNw7RtjHHsNQ0lAKU9nA/edit?usp=sharing';
  var doc=DocumentApp.openByUrl(urldoc);
  var body = doc.getBody();

  
  
  for (var i = 0 ; i < numRows-1; ++i ){

    var document = DocumentApp.create('CXC Daniela Buitrago'+ i);
    var bodynew = document.getBody();
       
   // bodynew.setAttributes(doc.getBody().getAttributes());
    bodynew.appendParagraph(body.getText());
  


    var id=  document.getId();
 
   
    document.getActiveSection(); 
    
  var client = {
    
    fecha: sheet.getRange(i+2, 1).getValue(),
    ciudad: sheet.getRange(i+2, 2).getValue(),
    empresa: sheet.getRange(i+2, 3).getValue(),
    nit: sheet.getRange(i+2, 4).getValue(),
    valorDebe: sheet.getRange(i+2, 5).getValue(),
    valor: sheet.getRange(i+2, 6).getValue(),
    motivo: sheet.getRange(i+2, 7).getValue(),
    correo: sheet.getRange(i+2, 8).getValue()
  };

  bodynew.replaceText('{{ciudad}}', client.ciudad);
  bodynew.replaceText('{{fecha}}', client.fecha);
  bodynew.replaceText('{{empresa}}', client.empresa);
  bodynew.replaceText('{{nit}}', client.nit);
  bodynew.replaceText('{{valorDebe}}', client.valorDebe);
  bodynew.replaceText('{{valor}}', client.valor);
  bodynew.replaceText('{{motivo}}', client.motivo);
  
    //Formato
// bodynew.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    var style = {};
    style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
    style[DocumentApp.Attribute.FONT_FAMILY] = 'Arial';
    style[DocumentApp.Attribute.FONT_SIZE] = 12;
    style[DocumentApp.Attribute.BOLD] = false;
    

// Apply the custom style.
bodynew.setAttributes(style);
bodynew.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    
    
   document.saveAndClose();
  //Enviar correo
  var fileDoc = DriveApp.getFileById(id);
  var pdfDoc = DriveApp.createFile(fileDoc.getAs('application/pdf'));
  pdfDoc.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
  
    GmailApp.sendEmail(client.correo, 'Combinacion de correspondencia.', 'URL PDF No.' +i+'- URL: '+ pdfDoc.getUrl());
  
  
  }
 
}
