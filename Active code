var dropDownNames = []; //Global dictionary or the all the stored dropdown Names.

var savingDictionary = {}; // Global dictionary to store saving details

var selectedEmployeeKey = 'selectedEmployee';

const ss = SpreadsheetApp.getActiveSpreadsheet();
const formSheet = ss.getSheetByName('Form Responses 3');
const empSheet = ss.getSheetByName('Employee Details');
const filmSheet = ss.getSheetByName('Film Details');
const empTemplatesSheet = ss.getSheetByName('Employee Template');
const editedColour = '#99ff99'
const emptyColour = '#ec8787'
const loadedColour = '#c9daf8'
const whiteColour = '#ffffff'
const blackColour = '#000000'
const loadingColour = '#ff4242'
const loadingCellValue = 'F2'
const dropdownCellValue = 'C2'
const dropdownRangeValue = 'C2:E2'
const autoDetailsCell = 'K2'

function onOpen() {

  /**reset loading in case it errors */
  resetLoadingIndicator();

  Logger.log("Starting onOpen");

  loadingIndicator();

  resetEmployeeDetailsLayout()

  var dropdownCell = empSheet.getRange(dropdownCellValue); // Replace with your dropdown cell
  dropdownCell.setValue('<Create New Employee>');

  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();

  //Fee Section user Properties
  userProperties.setProperty('feeCodeSelected', '');
  userProperties.setProperty('empFeeRange', '');

  //Employment Section user Properties
  userProperties.setProperty('employmentCodeSelected', '');
  userProperties.setProperty('empEmploymentRange', '');


  var empTemplateHeaderCellsJSON = userProperties.getProperty('empTemplateHeaderCells');
  var empTemplateHeaderCells = JSON.parse(empTemplateHeaderCellsJSON);
  var selectedEmployeeKey = "selectedEmployee"; // replace with your actual key

  clearEmployeeDetails(empTemplateHeaderCells, selectedEmployeeKey);

  //AutoDetails set to ON.
  userProperties.setProperty('autoDetails', 'ON');
  var autoCell = empSheet.getRange(autoDetailsCell);
  autoCell.setValue('Auto Details: ON')

  processFormResponses()
  processFilmDetails()
  processEmployeeTemplateDetailsCells()
  //processEmployeeDetailsCells()

  loadingIndicator();

  Logger.log("Finished onOpen"); 

}


/** **************************************************************************************************** */

// This function triggers every time a user makes an edit in the spreadsheet
function employeeOnEdit(e) {

  Logger.log("Starting employeeOnEdit");

  loadingIndicator();

  var range = e.range;
  var sheet = range.getSheet();
  //var sheetEndRow = empSheet.getLastRow();
  //var feeSectionRange = ss.getRangeByName('WeeklyDailyFee');
  //var hoursSectionRange = ss.getRangeByName('hoursSection')

  if (sheet.getName() === 'Employee Details') { 

    var row = range.getRow();
    var column = range.getColumn()

    var value = range.getValue();
    Logger.log('value = ' + value)

    /** Is it a dropdown menu? */   
    var hasDropdown = checkIfCellIsDropdown(value, row, column);
    /** Is it a dropdown menu? YES*/
    if (hasDropdown) {

      /**  Is it the main Employee dropdown menu? YES*/
      if (range.getA1Notation() === dropdownCellValue) {

        resetEmployeeDetailsLayout()
        /**  Is the value <CREATE NEW EMPLOYEE>? YES*/
        if (value === '<Create New Employee>') {
          Logger.log ('Setting <Create New Employee>')
          clearEmployeeDetails(empTemplateHeaderCells, selectedEmployeeKey);
          Logger.log ('Finishing Setting <Create New Employee>')

         
        } else {
          /**  Is the value <CREATE NEW EMPLOYEE>? NO*/
          // Populate the employee details if an existing employee is selected
          Logger.log('value trying to populate ' + JSON.stringify(value))

          var result = matchSelection(value, row, column);

          Logger.log("employeeOnEdit - empNameMatch: " + result.empNameMatch);
          unhideNamedRange(value,row,column,result.empNameMatch)
          
        }
      }else {
        /**  Is it the main Employee dropdown menu? NO*/ 
        Logger.log('value trying to populate is in a small dropdown: ' + JSON.stringify(value))
        var result = matchSelection(value, row, column);

        Logger.log("employeeOnEdit - empNameMatch: " + result.empNameMatch, result.dropdownHeader);
        unhideNamedRange(value,row,column,result.empNameMatch,dropdownHeader)
      }
    }

    
    // Retrieve and parse the empTemplateHeaderCells dictionary from user properties
    var userProperties = PropertiesService.getUserProperties();
    var empTemplateHeaderCellsJSON = userProperties.getProperty('empTemplateHeaderCells');
    var empTemplateHeaderCells = JSON.parse(empTemplateHeaderCellsJSON);
    var matchingResponse = JSON.parse(userProperties.getProperty('matchingResponse'));
    var filmDetails = JSON.parse(userProperties.getProperty('filmDetails'));

    populateEmployeeDetails(matchingResponse,result.empNameMatch)
    
    if (result.empNameMatch) {
      populateEmployeeDetails(filmDetails.find(response => response.objectName === "Film Details"), result.empNameMatch, true);
    }

    // Check if the selectedEmployee dictionary is complete
    var selectedEmployeeJSON = userProperties.getProperty(selectedEmployeeKey);
    var selectedEmployee = JSON.parse(selectedEmployeeJSON);

    Logger.log('selectedEmployee = ' + JSON.stringify(selectedEmployee))
    //Logger.log('selectedEmployeeKey = ' + JSON.stringify(selectedEmployeeKey))
  

    employeeCalc(e, value, row, column);


    if (selectedEmployee && range.getA1Notation() !== dropdownCellValue) {
      // Iterate over each cell in the edited range and apply the backgroundColour function
      range.getValues().forEach((row, i) => {
        row.forEach((value, j) => {
          var cell = range.getCell(i + 1, j + 1);
          backgroundColour( cell, empTemplateHeaderCells, selectedEmployee);
        });
      });
    }
  }

  loadingIndicator();

  Logger.log("Finished employeeOnEdit");
}

/** **************************************************************************************************** */


/** Do not use below - just to grab code if needed for loadingSection */

function feeSection(feeCodeValue) {


  // 2. Populate data for the selected namedRange
  Logger.log('name trying to populate ' + JSON.stringify(selectedEmployee['Name']))
  //populateEmployeeDetails(selectedEmployee['Name'], selectedNamedRange);

  // 3. Hide other namedRanges
  namedRanges.forEach(namedRange => {
    if (namedRange.getName() !== feeCodeValue && namedRange.getName() !== 'mainDetails' && namedRange.getName() !== 'hoursSection' && namedRange.getName() !== 'employeeDetails') {
      var range = namedRange.getRange();
      var targetRange = empSheet.getRange(range.getRow(), range.getColumn(), range.getNumRows(), range.getNumColumns());
      empSheet.hideRow(targetRange);
    }
  });

  // 4. If there's any data in the other hidden namedRanges, make them blank
  var empTemplateHeaderCellsJSON = userProperties.getProperty('empTemplateHeaderCells');
  var empTemplateHeaderCells = JSON.parse(empTemplateHeaderCellsJSON);

  namedRanges.forEach(namedRange => {
    if (namedRange.getName() !== feeCodeValue && namedRange.getName() !== 'mainDetails' && namedRange.getName() !== 'hoursSection' && namedRange.getName() !== 'employeeDetails') {
      var range = namedRange.getRange();
      for (var key in empTemplateHeaderCells) {
        if (empTemplateHeaderCells.hasOwnProperty(key)) {
          var cellInfo = empTemplateHeaderCells[key];
          var row = cellInfo.row;
          var column = cellInfo.column;

          // Check if the cell is in the hidden namedRange
          if (range.getRow() <= row && row < range.getLastRow() &&
              range.getColumn() <= column && column < range.getLastColumn()) {
            var cell = empSheet.getRange(row, column);
            if (cell.getValue()) {
              cell.setValue(""); // Clear the cell value
              cell.setBackground('#99ff99'); // Set the background color to edited color
            }
          }
        }
      }
    }
  });

  Logger.log("Finished feeSection");
}


