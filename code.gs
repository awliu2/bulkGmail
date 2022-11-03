// function that opens a window on a google spreadsheet to get id
// ID is also found in url, between "/d/" and "/edit/"

function getID() {
  Browser.msgBox('Spreadsheet key: ' + SpreadsheetApp.getActiveSpreadsheet().getId());
}

// parses data in spreadsheet into json
function getData(sheet){
  var out = {};
  var dataArray = [];

  var rows = sheet.getRange(2,1,sheet.getLastRow()-1, sheet.getLastColumn()).getValues();
  for (var i = 0, l = rows.length; i < l; i++)
  {
    var dataRow = rows[i];
    var record = {};
    record['id'] = dataRow[0];
    record['name'] = dataRow[1];
    record['email'] = dataRow[2];
    record['school'] = dataRow[3];
    dataArray.push(record);
  }
  out = dataArray;
  var result = JSON.stringify(out);
  return out;
}

// adds a custom menu to the spreadsheet
function onOpen(e)
 {
spreadsheet.SpreadsheetApp.getUi().createMenu('Custom Menu').addItem('Send Emails', 'sendEmail').addToUi();
}

// main function that sends an email for each column
function sendMail() {
  // getID();
  id = "1csenpFp1hW22FQy4Z9UwxAl_qi-6eaQPsUoGsgiEMpk";
  console.log(id.length);
  var ss = SpreadsheetApp.openById(id);
  // sheet name is the tab in the bottom right, defaults to "Sheet1"
  var sheet = ss.getSheetByName("Sheet1");
  // console.log(sheet);
  var json_data = getData(sheet);
  for(var j = 0; j < json_data.length; j++)
  {
    console.log(json_data[j]);

    // var attachment = [DriveApp.getFileById(json_data[j].fileid).getBlob()];
    var message="Dear "+ json_data[j].name + ",<br><br>Hope you are doing well at " + json_data[j].school + ". Please find the attachment below.<br><br>";
    
    message = message + "Sincerely,<br>" + "Andi<br><br>";
    MailApp.sendEmail({to: json_data[j].email,
                       cc: '',
                       subject: "TEST EMAIL",
                       htmlBody: message} );
    sheet.getRange(j+2,sheet.getLastColumn() ).setValue("Mail sent");}
}



