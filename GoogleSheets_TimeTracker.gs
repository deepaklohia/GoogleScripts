/** @OnlyCurrentDoc */

// Jai Guru Dev Jai Shiv Shankar 
// v1.1 Designed by Deepak Lohia at https://deepaklohia.com/
// Visit https://www.fiverr.com/deepaklohia for BUSINESS QUERIES  

function btn_clear() {
  var spreadsheet = SpreadsheetApp.getActive();
  var ui = SpreadsheetApp.getUi();
  var userChoice = ui.alert('You are about to clear data', ui.ButtonSet.OK_CANCEL);
  
  if (userChoice == ui.Button.OK) {
     spreadsheet.getRange('AA1').setValue(2);
    spreadsheet.getRange('A2:D' + spreadsheet.getDataRange().getNumRows() +5 ).clearContent();
   } 

};

//THIS BUTTON IS PRESSED WHEN WE START THE PROCESS
function btn_start() {
  var spreadsheet = SpreadsheetApp.getActive();
  
  if (spreadsheet.getRange('AA1').getValue() == '' ) {  
      SpreadsheetApp.getUi().alert('Click on Clear Button First');
  }
  else {
    var CurrentRow = spreadsheet.getRange('AA1').getValue();
    spreadsheet.getRange('B' + CurrentRow).setValue(new Date());
    spreadsheet.getRange('B' + CurrentRow).setNumberFormat('h:mm:ss');
  }      

};

//THIS BUTTON IS PRESSED WHEN WE END THE PROCESS
function btn_stop() {
  var spreadsheet = SpreadsheetApp.getActive();
  var CurrentRow = spreadsheet.getRange('AA1').getValue();
  
  if (CurrentRow == '' ) {  
    SpreadsheetApp.getUi().alert('Click on Clear Button First');
  }
  
  else if ( spreadsheet.getRange('B' + CurrentRow).getValue() == '' ){
      SpreadsheetApp.getUi().alert('Click on Start Button first');
    }
  
  else{
    if ( spreadsheet.getRange('B' + CurrentRow).getValue() != '' )  {
      spreadsheet.getRange('C' + CurrentRow).setValue(new Date());
      spreadsheet.getRange('C' + CurrentRow).setNumberFormat('h:mm:ss');
      spreadsheet.getRange('D' + CurrentRow).setValue('=text(C' + CurrentRow + '-B' + CurrentRow + ', "hh:mm:ss")');
      spreadsheet.getRange('AA1').setValue(CurrentRow + 1 ) ;
    }
  }
  
};

//THIS BUTTON IS USED TO RECORD ENDTIME AND START TIME IN CONTIUATION
function btn_split() {
  
  //first we stop the process
  var spreadsheet = SpreadsheetApp.getActive();
  var CurrentRow = spreadsheet.getRange('AA1').getValue();
    
  if (CurrentRow == '' ) {  
    SpreadsheetApp.getUi().alert('Click on Clear Button First');
  }
  else if ( spreadsheet.getRange('B' + CurrentRow).getValue() == '' ){
      SpreadsheetApp.getUi().alert('Click on Start Button first');
    }
   
    else{
      spreadsheet.getRange('C' + CurrentRow).setValue(new Date());
      spreadsheet.getRange('C' + CurrentRow).setNumberFormat('h:mm:ss');
      spreadsheet.getRange('D' + CurrentRow).setValue('=text(C' + CurrentRow + '-B' + CurrentRow + ', "hh:mm:ss")');
      CurrentRow ++;
      spreadsheet.getRange('AA1').setValue(CurrentRow) ;
      
      //Start Time
      spreadsheet.getRange('B' + CurrentRow).setValue(new Date());
      spreadsheet.getRange('B' + CurrentRow).setNumberFormat('h:mm:ss');
     
  } 
  
};
