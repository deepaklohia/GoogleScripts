function sendBulk() {

var ss = SpreadsheetApp.getActive();
var ui = SpreadsheetApp.getUi() ;
var lastRow = ss.getLastRow() ; 

var strDomain	 = "" ; 
var strClient_id = "" ; 
var strInstance_id = "" ; 
var	strDelay = "" ; 
var strMax_message = "" ; 

 for(var i = 2 ; i <=lastRow ; i++){

  strDomain	 = ss.getRange('A' + i).getValue()  ; 
  strClient_id = ss.getRange('B' + i).getValue()  ; 
  strInstance_id = ss.getRange('C' + i).getValue()  ; 
  strDelay =ss.getRange('D' + i).getValue()  ; 
  strMax_message = ss.getRange('E' + i).getValue()  ; 

   if (strDomain != "" && strClient_id != "" && strInstance_id != "" && strDelay != "" && strMax_message != ""){
      ui.alert( + "value1:"  + strDomain + "  value2:"  + strClient_id  + " value3:"  + strInstance_id   + "  value4:"  +  strDelay)
   }
  }
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

function SendSMS(numbers,message) {
  
  var apiKey = "apikey=" + "cyN/Ep8VeBQ-xxx";
  message = "&message=" + message;
  var sender = "&sender=" + "TXTLCL";
  numbers = "&numbers=" + numbers;
  dta = "https://api.textlocal.in/send/?" + apiKey + numbers + message + sender;
  var response = UrlFetchApp.fetch(dta);
  //Logger.log(response.getContentText());
  var status = (response.getContentText());
  return status
  //return (status.split(" ", 3);
  
};
