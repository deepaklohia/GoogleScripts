function sendBulkEmail() {

var ss = SpreadsheetApp.getActive();
var ui = SpreadsheetApp.getUi() ;
var lastRow = ss.getLastRow() ; 

var strTo	 = "" ; 
var	strCc = "" ; 
var strBcc = "" ; 
var strSub = "" ; 
var strBody = "" ; 

 for(var i = 2 ; i <=lastRow ; i++){

  strTo	 = ss.getRange('A' + i).getValue()  ; 
  strCc = ss.getRange('B' + i).getValue()  ; 
  strBcc = ss.getRange('C' + i).getValue()  ; 
  strSub =ss.getRange('D' + i).getValue()  ; 
  strBody = ss.getRange('E' + i).getValue()  ; 

   if (strTo == "" || strCc == "" || strBcc == "" || strSub == "" || strBody == "" ){
      ui.alert("some values missing")
   }
   else{
      sendEmailMessage (strTo, strCc , strBcc , strSub,strBody) ;
   }
  }

  ui.alert("Done") ;
}

function sendEmailMessage(strTo, strCc , strBcc , strSub,strBody) {
   var message = {
    to: strTo,
    subject: strSub,
    body: strBody,
    cc: strCc,
    bcc: strBcc,
    replyTo: "help@test.com"
  }
  MailApp.sendEmail(message);
}

function getLastDataRow(sh) {
  var lastRow = sh.getLastRow();
  var range = sh.getRange("A" + lastRow);
  if (range.getValue() !== "") {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }              
}
