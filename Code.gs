// Add menu entries
function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "Factures", functionName: "generateInvoiceFromCurrentSelection"} ];
  ss.addMenu("ADM", menuEntries);
}

function generateInvoiceFromCurrentSelection() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var historicSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
  Logger.log(sheet.getName());
  
  if(sheet.getIndex() != 1)
  {
    Logger.log("bad sheet : " + sheet.getName());
    return;
  }
  
  // Checking current year:
  var year = sheet.getName();
  // Checking current month:
  var currentSelection = sheet.getActiveRange();
  var currentColumn = currentSelection.getColumn();
  var month = sheet.getRange(1, currentColumn).getValue();
  var monthNumber = currentColumn - 5;

  var invoiceNumberPrefix = year.toString();
  if(monthNumber < 10)   // TODO : find a format function to create the invoice number
    invoiceNumberPrefix += "0";
  invoiceNumberPrefix += monthNumber.toString();
  
  Logger.log(monthNumber);
  if(currentColumn < 5)
  {
    Logger.log("bad cell : " + currentColumn);
    return;
  }
  
  var startIndex = currentSelection.getRowIndex();
  var endIndex = startIndex + currentSelection.getHeight();

  Logger.log(month);
  Logger.log(startIndex + " => " + endIndex);
  
  for(var i = startIndex;i<endIndex;i++)
  {
    var row = sheet.getRange(i, 1, 1, 26).getValues()[0];
    var id = row[2];
    if(id != "")
    {
      Logger.log(id)
      var price = row[currentColumn-1];
        Logger.log(price);
      if(price > 0)
      {
        var cell = sheet.getRange(i, currentColumn);
        cell.setBackground("red");
        
        var number = invoiceNumberPrefix.concat(id);           
        var name = row[1] + " " + row[0];
        var email = row[25];
        var company = row[21];
        var address = row[22] + "\n" + row[23] + " " + row[24];
        var coworker = new Coworker(id, name, email, company, address);
        var prestation = row[3];
        var invoice = new Invoice(number, month, coworker, prestation, price);
        invoice.generate();
        invoice.store(historicSheet);
        Logger.log(number + " - " + id + " (" + row[0] + ") : " + price + "€ for " + prestation);
        
        invoice.send(email);
        cell.setBackground("green");
      }
    }
  }
}

function testGetNextInvoiceNumber() {
  var historicSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1];
  var number = getNextInvoiceNumber(historicSheet);
  Logger.log(number);
}

function getNextInvoiceNumber(historicSheet) {
  var data = historicSheet.getDataRange();
  var values = data.getValues();
  var rowNumber = data.getHeight();
  var lastRow = values[rowNumber-1];
  var lastNumber = lastRow[0];
  
  return lastNumber + 1;
}

