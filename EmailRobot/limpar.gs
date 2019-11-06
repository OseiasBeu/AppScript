function limpar() {
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange("A4:F400").clear();
  sheet.setActiveSelection("A4")
}

