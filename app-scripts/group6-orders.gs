// Activate sheet by Name
function activateSheetByName(sheetName) {
  var sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  sheet.activate();
  return sheet;
}

// Write Orders Data to Firestore 

function writeOrdersDataToFirebase() {
  var email = "** Paste email here **";
  var key = "** Paste private key here **";
  var projectId = "** Paste project id here **";
  var firestore = FirestoreApp.getFirestore(email, key, projectId);
  // var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  // Orders
  // OrderID	CustomerID	EmployeeID	OrderDate	RequiredDate	ShippedDate	ShipVia	Freight	ShipName	ShipAddress	ShipCity	ShipRegion	ShipPostalCode	ShipCountry
  
  var sheetName = 'orders';
  var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = activateSheetByName(sheetName);
  Logger.log(sheetName);
  
  var data = sheet.getDataRange().getValues();
  var dataToImport = {};
  
  // Use this loop code if you want all rows in the sheet: 
  for(var i = 1; i < data.length; i++) {
    var OrderID = data[i][0];
    var CustomerID = data[i][1];
    Logger.log(OrderID + '-' + CustomerID);
    dataToImport[OrderID + '-' + CustomerID] = {
      OrderID:data[i][0],
      CustomerID:data[i][1],
      EmployeeID:data[i][2],
      OrderDate:data[i][3],
      RequiredDate:data[i][4],
      ShippedDate:data[i][5],
      ShipVia:data[i][6],
      Freight:data[i][7],
      ShipName:data[i][8],
      ShipAddress:data[i][9],
      ShipCity:data[i][10],
      ShipRegion:data[i][11],
      ShipPostalCode:data[i][12],
      ShipCountry:data[i][13]
    };
    firestore.createDocument("Orders/", dataToImport[OrderID + '-' + CustomerID]);
  }
}

// ------------------ Write Orders Details Data to Firestore -----------

function writeOrderDetailsDataToFirebase() {
  var email = "** Paste email here **";
  var key = "** Paste private key here **";
  var projectId = "** Paste project id here **";
  var firestore = FirestoreApp.getFirestore(email, key, projectId);
  // var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  
  // Order Details
  // ID	OrderID	ProductID	UnitPrice	Quantity	Discount		
  
  var sheetName = 'order_details';
  var url = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var ss = SpreadsheetApp.openByUrl(url);
  var sheet = activateSheetByName(sheetName);
  Logger.log(sheetName);
  
  var data = sheet.getDataRange().getValues();
  var dataToImport = {};
  
  // Logger.log(data.length); // the number of rows in the sheet
  
  // Use this loop code if you want all rows in the sheet: 
  for(var i = 1; i < data.length; i++) {
    var ID = data[i][0];
    var ProductID = data[i][1];
    Logger.log(ID + '-' + ProductID);
    dataToImport[ID + '-' + ProductID] = {
      ID: data[i][0],
      OrderID: data[i][1],
      ProductID: data[i][2],
      UnitPrice: data[i][3],
      Quantity: data[i][4],
      Discount: data[i][5],
    };
    firestore.createDocument("OrderDetails/", dataToImport[ID + '-' + ProductID]);
    // Logger.log(dataToImport);
  }
}
