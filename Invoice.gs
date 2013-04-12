var Invoice = function(number, month, coworker, prestation, priceati) {
  var _number = number;
  var _month = month;
  var _coworker = coworker;
  var _prestation = prestation;
  var _priceati = priceati;
  var _doc;
    
  this.getPriceATI = function() {
    return _priceati.toFixed(2);
  }
  
  this.getPriceET = function() {
    var price = Math.ceil(100 * this.getPriceATI() / 1.196) / 100;
    return price.toFixed(2);
  }
  
  this.getVAT = function() {
    var vat = this.getPriceATI() - this.getPriceET();
    return vat.toFixed(2);
  }
  
  this.generate = function() {
    Logger.log(coworker.getName());
    // Creating a copy of the template file
    var file = DocsList.getFileById("17dE95TWcy96zyUP8vXCvb4d1-4TYC2hQH-sbRicBc8s");
    
    file = file.makeCopy("current");

    // Put it in the "Factures" folder
    file.addToFolder(DocsList.getFolderById("0B_otmye6qyn3SUs2S3R4cHBYb3c"));
    
    // Edit the document to add the description:
    _doc = DocumentApp.openById(file.getId());
    _doc.replaceText("#NUMBER#", _number);
    _doc.replaceText("#COMPANY#", _coworker.getCompany());
    _doc.replaceText("#NAME#", _coworker.getName());
    _doc.replaceText("#ADDRESS#", _coworker.getAddress());
    _doc.replaceText("#PRESTATION#", _prestation);
    
    var currentTime = new Date();
    var currentDateString = Utilities.formatDate(currentTime, "GMT+1", "dd/MM/yyyy")
    Logger.log(currentDateString);
    _doc.replaceText("#DATE#", currentDateString);
    
    _doc.replaceText("#PRICEET#", this.getPriceET());  
    _doc.replaceText("#VAT#", this.getVAT());  
    _doc.replaceText("#PRICEATI#", this.getPriceATI());  
    
    _doc.saveAndClose();
    
    file.rename("Facture " + coworker.getName() + " - " + month);
  }
  
  this.store = function(historicSheet) {
    historicSheet.appendRow([_number, coworker.getName(), month, prestation, this.getPriceET(), this.getVAT(), this.getPriceATI()]);
  }
  
  this.send = function(email) {
    if(_doc != undefined) {
      var name = _doc.getName();
      var file = DocsList.getFileById(_doc.getId());
      var pdf = file.getAs('application/pdf').getBytes();
      var attach = {fileName: name+".pdf",content:pdf, mimeType:'application/pdf'};
      var content = 'Bonjour,\n\nEt voici la facture pour le mois de ' + _month + '.';
      MailApp.sendEmail(email, name, content, {attachments:[attach]});
    }
  }
}

function testInvoice() {
  generateInvoiceFromCurrentSelection();
}
