/* Route
 * All Request with Method GET will be process here
 */
//var email = Session.getActiveUser().getEmail();
//var ldap = email.split("@")[0];
var ldap = "ldapTest",
    email = "emailTest";
function getEmail (){
  return Session.getActiveUser().getEmail();
}

function doGet(req) {
   var action = req.parameter.action;

//   var db = SpreadsheetApp.openById("16dquyA-xIKRa7i0NN68Iuf5fW0HwYq-Fs273cuWW1B0");
    var db = SpreadsheetApp.openById("1nLHf8k_NhRIp_NP16oQTqf59YxGgZ2d-kwpy4iXdMJ8");
  
  
   // Don't forget to change your Sheet Name by default is 'Sheet1'
   var sheetUsers = db.getSheetByName("All Cases");
   
   switch(action) {
       case "read":
           return doRead(req, sheetUsers);
           break;
       case "insert":
           return doInsert(req, sheetUsers);
           break;
       case "update":
           return doUpdate(req, sheetUsers);
           break;
       case "delete":
           return doDelete(req, sheetUsers);
           break;
       default:
           return response().json({
              status: false,
              message: 'silent!'
           });
   }
}

/* Read
 * request for all Data
 *
 * @request-parameter | action<string>
 * @example-request | ?action=read
 */
function doRead(request, sheetObject) {
   var data = {};
   
   data.records = _readData(sheetObject);

   return response().json(data);

}

/* Insert
 *
 */
function doInsert(req, sheet) {
   var id = req.parameter.id;
   
   // all data your needed

   var flag = 1;
   var Row = sheet.getLastRow();
//   for (var i = 1; i <= Row; i++) {
//      /* getRange(i, 2) 
//       * i | is a row index
//       * 1 | is a id column index ('id')
//       */
//      var idTemp = sheet.getRange(i, 1).getValue();
//      if (idTemp == id) {
//         flag = 0;
//         var result = "Sorry bratha, id already exist";
//      }
//   }
   
   // add new row with recieved parameter from client
   if (flag == 1) {
      var timestamp = Date.now();
      var currentTime = new Date().toLocaleString(); // Full Datetime
      var rowData = sheet.appendRow([
         ldap,
         email,
         id,
         timestamp,
         currentTime
      ]);
      var result = "Insertion successful";
   }

   return response().json({
      result: result
   });
}

/* Update
 * request for Update
 *
 * @request-parameter | id<string>, data<JSON>, action<string>
 * @example-request | ?id=1&action=update&data={"email":"ryandevstudio@gmail.com", "username":"nyancodeid"}
 */
function doUpdate(req, sheet) 
{
   var id = req.parameter.id;
   var updates = JSON.parse(req.parameter.data);
  
   var lr = sheet.getLastRow();

   var headers = _getHeaderRow(sheet);
   var updatesHeader = Object.keys(updates);
   
   // Looping for row
   for (var row = 1; row <= lr; row++) {
      // Looping for available header / column
      for (var i = 0; i <= (headers.length - 1); i++) {
         var header = headers[i];
         // Looping for column need to updated
         for (var update in updatesHeader) {
            if (updatesHeader[update] == header) {
               // Get ID for every row
               var rid = sheet.getRange(row, 1).getValue();

               if (rid == id) {
                  // Lets Update
                  sheet.getRange(row, i + 1).setValue(updates[updatesHeader[update]]);
               }
            }
         }
      }
   }

   
   // Output
   return response().json({
      status: true,
      message: "Update successfully"
   });
}


/* Delete
 *
 */
function doDelete(req, sheet) {
   var id = req.parameter.id;
   var flag = 0;

   var Row = sheet.getLastRow();
   for (var i = 1; i <= Row; i++) {
      var idTemp = sheet.getRange(i, 1).getValue();
      if (idTemp == id) {
         sheet.deleteRow(i);
         
         var result = "deleted successfully";
         flag = 1;
      }

   }

   if (flag == 0) {
      return response().json({
         status: false,
         message: "ID not found"
      });
   }

   return response().json({
      status: true,
      message: result
   });
}


/* Service
 */
function _readData(sheetObject, properties) {

   if (typeof properties == "undefined") {
      properties = _getHeaderRow(sheetObject);
      properties = properties.map(function (p) {
//         return p.replace(/\s+/g, '_');
        return p;
      });
   }

   var rows = _getDataRows(sheetObject),
      data = [];

   for (var r = 0, l = rows.length; r < l; r++) {
      var row = rows[r],
          record = {};

      for (var p in properties) {
         record[properties[p]] = row[p];
      }

      data.push(record);
   }
   
   return data;
}
function _getDataRows(sheetObject) {
   var sh = sheetObject;

   return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
}
function _getHeaderRow(sheetObject) {
   var sh = sheetObject;

   return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}
function response() {
   return {
      json: function(data) {
         return ContentService
            .createTextOutput(JSON.stringify(data))
            .setMimeType(ContentService.MimeType.JSON);
      }
   }
}
