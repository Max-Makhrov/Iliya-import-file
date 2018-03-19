var C_RANGE_EVAL = 'eval';



// Declare
var STR_DELIMEER1 // delim1
var STR_DELIMEER2 // delim2
var STR_IDS // files ids
var STR_SHEETS // sheet names
var STR_RANGES // first data rows
var STR_SHEET_TARGET // sheet where to
var STR_RANGE_TARGET // Start import from cell
var STR_SQL_TEXTS // SQL texts


// get settings from named range
function getSettings()
{
  var range = SpreadsheetApp.getActive().getRangeByName(C_RANGE_EVAL);
  
  /*
    sample data for this script looks like this:    
    [
      ";",
      "~",
      "1LC6QmhBU-0OhUWo7R_eKPjuSCkmdpl6tRHPu83Co3Hk;1389-68_t6yFVQb72P8YhBPaTEnyE7sxJ7Imd9tPNF08;14L2QMZBtwzWDz-IkALq9-EUjdjvKDPdJJ9EyodSidRs",
      "Sales Central;Sales West;Sales East",
      "A2:G2;A2:G2;A2:G2",
      "Sales Total;Sales Total;Sales Total",
      "A2"
    ]
    Note:
    The data is collected from a cell of named range called "eval"
  */
  
  var value = range.getValue();
  var data = JSON.parse(value);
  
  // Assign
  STR_DELIMEER1 = data[0];
  STR_DELIMEER2 = data[1];
  STR_IDS = data[2];
  STR_SHEETS = data[3];
  STR_RANGES = data[4];
  STR_SHEET_TARGET = data[5];
  STR_RANGE_TARGET = data[6];
  STR_SQL_TEXTS = data[7];
}




function onOpen()
{
  // Add a custom menu to the spreadsheet.
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Admin')
      .addItem('Update', 'writeDataFromSheets')
      .addItem('Set Trigger', 'triggerWriteDataFromSheets')
      .addSeparator()
      .addItem('Delete extra cells', 'deleteEmptyCells')
      .addToUi()
}







function writeDataFromSheets()
{
  getSettings();
  
  var sheetNamesTo = STR_SHEET_TARGET.split(STR_DELIMEER1);
  var allFileIds = STR_IDS.split(STR_DELIMEER1);
  var allSheetNames = STR_SHEETS.split(STR_DELIMEER1);
  var allRangeNames = STR_RANGES.split(STR_DELIMEER1);
  var allSqlTexts = STR_SQL_TEXTS.split(STR_DELIMEER1);

  // get tasks
  var tasks = {}; 
  var task = {};
  sheetNamesTo.forEach(
  function(elt, i)
  {
    if (!(elt in tasks)) 
    { 
      tasks[elt] = {}; 
      task = tasks[elt];
      task.fileIds = [allFileIds[i]];
      task.sheetNames = [allSheetNames[i]];
      task.rangeNames = [allRangeNames[i]];
      task.sqlTexts = [allSqlTexts[i]]
    }
    else
    {
      task = tasks[elt];
      task.fileIds.push(allFileIds[i]);
      task.sheetNames.push(allSheetNames[i]);
      task.rangeNames.push(allRangeNames[i]);  
      task.sqlTexts.push(allSqlTexts[i])
    }  
  }
  );
  
 
 // do tasks
 var file = SpreadsheetApp.getActive();
 for (var key in tasks) {
  task = tasks[key];
  var fileIds = task.fileIds;
  var sheetNames = task.sheetNames;
  var rangeNames = task.rangeNames;
  var sqlTexts = task.sqlTexts;
  var data = getDataFromSheets_(fileIds, sheetNames, rangeNames, sqlTexts);
  if (data.length > 0) 
  {
   // write the result into sheet
    var sheet = file.getSheetByName(key);
    var range = sheet.getRange(STR_RANGE_TARGET);
    sheet.clearContents(); // clear old values
    writeDataIntoSheet_(file, sheet, data, range.getRow(), range.getColumn());
  }
 }

}


function getDataFromSheets_(fileIds, sheetNames, rangeNames, sqlTexts)
{  
  // get arrays
  var arrays = [];
  var array = [];
  
  for (var i = 0, l = fileIds.length; i < l; i++)
  {
    array = getDataFromSheet_(fileIds[i], sheetNames[i], rangeNames[i], sqlTexts[i]);
    if('null' != array) { arrays.push(array); }     
  }
  return combine2DArrays_(arrays);
}



function getDataFromSheet_(fileId, sheetName, rangeName, sqlText)
{
  var file = SpreadsheetApp.openById(fileId);
  var sheet = file.getSheetByName(sheetName);
  var r1 = sheet.getRange(rangeName);
  var row1 = r1.getRow();
  var col1 = r1.getColumn();
  var col2 = r1.getLastColumn();
  
  var row2 = sheet.getLastRow(); // last row from sheet
  
  if (row2 < row1) return null;
  
  var range = sheet.getRange(row1, col1, row2-row1+1, col2-col1+1);

  var data = range.getValues();
  
  return getAlaSqlResult_(data, sqlText);
}


// combine 2d arrays of different sizes
function combine2DArrays_(arrays)
{
  
  // check 2D-arryas are not empty
  if (arrays.length === 1 && arrays[0].length === 0) { return arrays[0]; }
  
  // detect max L
  var l = 0;
  var row = [];
  var result = [];
  var elt = '';
  arrays.forEach(function(arr) { l = Math.max(l, arr[0].length); } );
  arrays.forEach(function(arr) {
    for (var i = 0, h = arr.length; i < h; i++)
    {
      var row = arr[i];
      // fill with empty value
      for (var ii = row.length; ii < l; ii++) { row.push(''); }
      result.push(row);
    }
  }
  );  
  return result;
}


/*
    use getSheetsInfo(ids)
    
    write the report into sheet:
    
    input:
      * file                       SpreadSheet
      * strSheet                   'Sheet1'
      * data                       [['Name', 'Age'], ['Max', 28], ['Lu', 26]]
                              
  If strSheet doesn't exist → creates new sheet
                                    
*/
function writeDataIntoSheet_(file, sheet, data, rowStart, colStart) {
  file = file || SpreadsheetApp.getActiveSpreadsheet();
  
  // get sheet as object
  switch(typeof sheet) {
    case 'object':
        break;
    case 'string':
        sheet = createSheetIfNotExists(file, sheet);
        break;
    default:
        return 'sheet is invalid';
  }
  
  // get dimansions and get range
  rowStart = rowStart || 1;
  colStart = colStart || 1;   
  var numRows = data.length;
  var numCols = data[0].length; 
  var range = sheet.getRange(rowStart, colStart, numRows, numCols);
  
  // clear old data if rowStart or colStart are not defined
  if(!rowStart && !colStart) { sheet.clearContents(); }

  
  // set values
  range.setValues(data);
  
  // report success
  return 'Wtite data to sheet -- ok!';

}







function triggerWriteDataFromSheets()
{
  var nameFunction = 'writeDataFromSheets';
  setTriggerOnHour(nameFunction)
}


function setTriggerOnHour(nameFunction)
{
  if (checkTriggerExists(nameFunction, 'SPREADSHEETS')) { return -1; } // trigger exists
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger(nameFunction)
      .timeBased()
      .everyHours(1)
      .create();

}


/*
  USAGE
  
  var exists = checkTriggerExists('test_getSets', 'SPREADSHEETS')
*/
function checkTriggerExists(nameFunction, triggerSourceType)
{
  var triggers = ScriptApp.getProjectTriggers();
  var trigger = {};

  
  for (var i = 0; i < triggers.length; i++) {
   trigger = triggers[i];
   if (trigger.getHandlerFunction() == nameFunction && trigger.getTriggerSource() == triggerSourceType) return true;
  }
  
  return false; 

}
