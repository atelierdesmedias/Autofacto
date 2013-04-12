var Coworker = function(id, name, email, company, address) {
  var _id = id;
  var _name = name;
  var _email = email;
  var _company = company;
  var _address = address;
  
  this.getName = function() {
    return _name;
  }
  
  this.getEmail = function() {
    return _email;
  }

  this.getCompany = function() {
    return _company;
  }

  this.getAddress = function() {
    return _address;
  }
}

function getContactById(id) {
  return new Coworker("md", "Martin Delille", "martin@dubware.net", "Dubware", "31 rue des Capucins\n69001 Lyon");
}

function test2() {
  var coworker = getContactById("md");
  Logger.log(coworker.getName());
}
