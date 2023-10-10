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


