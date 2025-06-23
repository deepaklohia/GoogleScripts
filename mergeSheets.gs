/** @OnlyCurrentDoc */

function mergeSheets() {


var desWorkbook = SpreadsheetApp.getActive();
var mainWorksheet = desWorkbook.getSheetByName("Main");
var desWorksheet = desWorkbook.getSheetByName("Data") ;
desLastrow = desWorksheet.getLastRow();
desLastcol = desWorksheet.getLastColumn();
var srcLastrow  ;
var srcLastcol  ;

var sheetInfo = '';
var sheetId ='' ;
var tempSrcWorksheetName = '';
var headerFlag = false;

mainWorksheet.getRange("B2:B" +  mainWorksheet.getLastRow() ).clearContent();

var ui = SpreadsheetApp.getUi();
var userChoice = ui.alert('You are about to clear data', ui.ButtonSet.OK_CANCEL);

if (desLastrow ==0){desLastrow = 1};
if (desLastcol ==0){desLastcol = 1};

if (userChoice == ui.Button.OK) {
   desWorksheet.getRange(1, 1, desLastrow, desLastcol).clearContent();
} else { headerFlag = true ; }
 
for(var i = 2 ; i <=  mainWorksheet.getLastRow() ; i++){
 
     sheetInfo = mainWorksheet.getRange('A' + i).getValue();
     sheetId = checkFile(sheetInfo) ;
            
        if (sheetId != "" ){
           
             var srcWorkbook = SpreadsheetApp.openById(sheetId);
             var srcWorksheet  = srcWorkbook.getSheets()[0];

             tempSrcWorksheetName = srcWorksheet.copyTo(desWorkbook).getName() ;
             var tempSrcWorksheet =  desWorkbook.getSheetByName(tempSrcWorksheetName) ;
             srcLastrow = tempSrcWorksheet.getLastRow();
             srcLastcol = tempSrcWorksheet.getLastColumn();
             desLastrow = desWorksheet.getLastRow();
             desLastcol = desWorksheet.getLastColumn();
                          
             if (desLastrow == 0)  {  desLastrow = 1 }
             if (desLastcol  == 0) {desLastcol = 1 }
             
             var srcLastrow = srcWorksheet.getLastRow();
             var srcLastcol = srcWorksheet.getLastColumn();
             
             if (headerFlag ==  false){
               tempSrcWorksheet.getRange(1, 1, srcLastrow ,srcLastcol ).copyTo( desWorksheet.getRange(desLastrow, 1) , SpreadsheetApp.CopyPasteType.PASTE_VALUES     );
               headerFlag = true;
             }
             else{
                  tempSrcWorksheet.getRange(2, 1, srcLastrow ,srcLastcol ).copyTo( desWorksheet.getRange(desLastrow+1, 1) , SpreadsheetApp.CopyPasteType.PASTE_VALUES     );
             }            
           
             mainWorksheet.getRange('B' + i).setValue('Done');
             desWorkbook.deleteSheet(tempSrcWorksheet);
             
             }
          
        else{
            mainWorksheet.getRange('B' + i).setValue('Not Found');
            } 
  }
 
};


function checkFile(fileName){
  var results;
  var files  = DriveApp.getFilesByName(fileName);

  if(files.hasNext() == false){   //Does not exist
      results = '' ;
  }
  else{   //Does exist
      results =  files.next().getId() ;
  }
  return results;
}
 
